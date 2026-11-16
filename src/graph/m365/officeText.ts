import { inflateRawSync } from "node:zlib";

/**
 * Small, bounded text extractor for Office Open XML files (.docx, .xlsx,
 * .pptx and their macro/template variants), using only node:zlib.
 *
 * Office files are ZIP archives of XML parts. This reader:
 *
 * - parses the ZIP central directory (no ZIP64, no encryption, no
 *   multi-disk archives — those are rejected, not guessed at);
 * - only inflates the few parts it needs (document.xml, slides, sheets,
 *   shared strings, workbook/presentation and their relationships);
 * - caps every inflated part with zlib's `maxOutputLength` and the whole
 *   archive with a total budget, so a zip bomb fails fast with an error
 *   instead of exhausting memory (declared sizes are never trusted);
 * - stops building text once `maxChars` is reached.
 *
 * It returns plain text only: formatting, images, embedded objects,
 * comments, headers/footers and formulas are ignored (cells yield their
 * cached values).
 */

export interface OfficeTextLimits {
  /** Max bytes one inflated ZIP entry may produce. */
  maxEntryBytes: number;
  /** Max inflated bytes across all entries read from one archive. */
  maxTotalBytes: number;
  /** Max entries in the central directory. */
  maxEntries: number;
  /** Text is cut at this many characters. */
  maxChars: number;
}

export const DEFAULT_OFFICE_LIMITS: OfficeTextLimits = {
  maxEntryBytes: 32 * 1024 * 1024,
  maxTotalBytes: 64 * 1024 * 1024,
  maxEntries: 10_000,
  maxChars: 50_000,
};

export type OfficeKind = "word" | "excel" | "powerpoint";

const EXTENSION_KIND: Record<string, OfficeKind> = {
  docx: "word",
  docm: "word",
  dotx: "word",
  dotm: "word",
  xlsx: "excel",
  xlsm: "excel",
  xltx: "excel",
  xltm: "excel",
  pptx: "powerpoint",
  pptm: "powerpoint",
  ppsx: "powerpoint",
  ppsm: "powerpoint",
  potx: "powerpoint",
  potm: "powerpoint",
};

export function officeKindForExtension(ext: string): OfficeKind | undefined {
  return EXTENSION_KIND[ext.toLowerCase()];
}

export class OfficeTextError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "OfficeTextError";
  }
}

export interface OfficeTextResult {
  kind: OfficeKind;
  text: string;
  truncated: boolean;
  /** Pages of content found: slides for PowerPoint, sheets for Excel, 1 for Word. */
  parts: number;
}

// ---------------------------------------------------------------------------
// ZIP reading

interface ZipEntry {
  name: string;
  flags: number;
  method: number;
  compressedSize: number;
  localHeaderOffset: number;
}

const SIG_EOCD = 0x06054b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_LOCAL = 0x04034b50;

export class ZipReader {
  private readonly entries = new Map<string, ZipEntry>();
  private budget: number;

  constructor(
    private readonly buf: Buffer,
    private readonly limits: OfficeTextLimits = DEFAULT_OFFICE_LIMITS
  ) {
    this.budget = limits.maxTotalBytes;
    this.readCentralDirectory();
  }

  names(): string[] {
    return [...this.entries.keys()];
  }

  has(name: string): boolean {
    return this.entries.has(name);
  }

  private readCentralDirectory(): void {
    const buf = this.buf;
    if (buf.length < 22) throw new OfficeTextError("Not a ZIP archive (too small).");
    const stop = Math.max(0, buf.length - 22 - 0xffff);
    let eocd = -1;
    for (let i = buf.length - 22; i >= stop; i--) {
      if (buf.readUInt32LE(i) === SIG_EOCD) {
        eocd = i;
        break;
      }
    }
    if (eocd < 0) throw new OfficeTextError("Not a ZIP archive (no end of central directory); the file may be encrypted or a legacy binary format.");
    const disk = buf.readUInt16LE(eocd + 4);
    const cdDisk = buf.readUInt16LE(eocd + 6);
    const total = buf.readUInt16LE(eocd + 10);
    const cdSize = buf.readUInt32LE(eocd + 12);
    const cdOffset = buf.readUInt32LE(eocd + 16);
    if (disk !== 0 || cdDisk !== 0) throw new OfficeTextError("Multi-disk ZIP archives are not supported.");
    if (total === 0xffff || cdOffset === 0xffffffff || cdSize === 0xffffffff) throw new OfficeTextError("ZIP64 archives are not supported.");
    if (total > this.limits.maxEntries) throw new OfficeTextError(`ZIP archive has too many entries (${total}).`);
    if (cdOffset + cdSize > eocd) throw new OfficeTextError("Corrupt ZIP central directory.");

    let p = cdOffset;
    for (let n = 0; n < total; n++) {
      if (p + 46 > eocd || buf.readUInt32LE(p) !== SIG_CENTRAL) throw new OfficeTextError("Corrupt ZIP central directory entry.");
      const flags = buf.readUInt16LE(p + 8);
      const method = buf.readUInt16LE(p + 10);
      const compressedSize = buf.readUInt32LE(p + 20);
      const nameLen = buf.readUInt16LE(p + 28);
      const extraLen = buf.readUInt16LE(p + 30);
      const commentLen = buf.readUInt16LE(p + 32);
      const localHeaderOffset = buf.readUInt32LE(p + 42);
      const end = p + 46 + nameLen + extraLen + commentLen;
      if (end > eocd) throw new OfficeTextError("Corrupt ZIP central directory entry.");
      const name = buf.toString("utf8", p + 46, p + 46 + nameLen);
      if (!this.entries.has(name)) this.entries.set(name, { name, flags, method, compressedSize, localHeaderOffset });
      p = end;
    }
  }

  /** Inflates one entry within the per-entry and total budgets, or returns undefined if absent. */
  read(name: string): Buffer | undefined {
    const entry = this.entries.get(name);
    if (!entry) return undefined;
    if (entry.flags & 0x1) throw new OfficeTextError("Encrypted ZIP entries are not supported.");
    const buf = this.buf;
    const off = entry.localHeaderOffset;
    if (off + 30 > buf.length || buf.readUInt32LE(off) !== SIG_LOCAL) throw new OfficeTextError("Corrupt ZIP local header.");
    const start = off + 30 + buf.readUInt16LE(off + 26) + buf.readUInt16LE(off + 28);
    const end = start + entry.compressedSize;
    if (end > buf.length) throw new OfficeTextError("Corrupt ZIP entry (data past end of file).");
    const data = buf.subarray(start, end);
    const allowed = Math.min(this.limits.maxEntryBytes, this.budget);
    let out: Buffer;
    if (entry.method === 0) {
      if (data.length > allowed) throw new OfficeTextError("A part of the document is too large to extract.");
      out = Buffer.from(data);
    } else if (entry.method === 8) {
      try {
        out = inflateRawSync(data, { maxOutputLength: Math.max(1, allowed) });
      } catch (err) {
        if (err instanceof RangeError || (err as { code?: string }).code === "ERR_BUFFER_TOO_LARGE") {
          throw new OfficeTextError("A part of the document is too large to extract (decompression limit reached).");
        }
        throw new OfficeTextError("A part of the document could not be decompressed.");
      }
    } else {
      throw new OfficeTextError(`Unsupported ZIP compression method ${entry.method}.`);
    }
    this.budget -= out.length;
    return out;
  }

  readText(name: string): string | undefined {
    const data = this.read(name);
    return data === undefined ? undefined : data.toString("utf8");
  }
}

// ---------------------------------------------------------------------------
// XML helpers (regex based: OOXML parts are machine written and regular)

const XML_ENTITIES: Record<string, string> = { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" };

export function decodeXmlEntities(s: string): string {
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g, (m, e: string) => {
    if (e.startsWith("#")) {
      const code = e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : m;
    }
    return XML_ENTITIES[e] ?? m;
  });
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export function xmlAttr(attrs: string, name: string): string | undefined {
  const m = new RegExp(`(?:^|\\s)${escapeRegExp(name)}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(attrs);
  if (!m) return undefined;
  return decodeXmlEntities(m[1] ?? m[2] ?? "");
}

/** Relationship id -> archive path, for a part's .rels file. */
function readRelationships(zip: ZipReader, relsPath: string, baseDir: string): Map<string, string> {
  const map = new Map<string, string>();
  const xml = zip.readText(relsPath);
  if (!xml) return map;
  for (const m of xml.matchAll(/<Relationship\b([^>]*?)\/?>/g)) {
    const attrs = m[1] ?? "";
    const id = xmlAttr(attrs, "Id");
    const target = xmlAttr(attrs, "Target");
    if (!id || !target || xmlAttr(attrs, "TargetMode") === "External") continue;
    map.set(id, resolvePartPath(baseDir, target));
  }
  return map;
}

function resolvePartPath(baseDir: string, target: string): string {
  const parts = target.startsWith("/") ? [] : baseDir.split("/").filter(Boolean);
  for (const seg of target.split("/")) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") parts.pop();
    else parts.push(seg);
  }
  return parts.join("/");
}

/** Collects text up to a character cap, so huge documents stop early. */
class TextSink {
  private chunks: string[] = [];
  private length = 0;
  truncated = false;

  constructor(private readonly maxChars: number) {}

  get full(): boolean {
    return this.length >= this.maxChars;
  }

  push(s: string): void {
    if (this.full) {
      if (s.length > 0) this.truncated = true;
      return;
    }
    this.chunks.push(s);
    this.length += s.length;
  }

  result(): { text: string; truncated: boolean } {
    const raw = tidy(this.chunks.join(""));
    if (raw.length > this.maxChars) return { text: raw.slice(0, this.maxChars), truncated: true };
    return { text: raw, truncated: this.truncated };
  }
}

function tidy(text: string): string {
  return text
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

// ---------------------------------------------------------------------------
// Word

export function wordXmlToText(xml: string): string {
  const out: string[] = [];
  const re = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\s*\/>|<w:(?:br|cr)\b[^>]*\/>|<\/w:(p|tc|tr)>/g;
  for (const m of xml.matchAll(re)) {
    const tok = m[0];
    if (m[1] !== undefined) out.push(decodeXmlEntities(m[1]));
    else if (tok.startsWith("<w:tab")) out.push("\t");
    else if (tok.startsWith("<w:br") || tok.startsWith("<w:cr")) out.push("\n");
    else if (m[2] === "tc") out.push("\t");
    else out.push("\n");
  }
  return out.join("").replace(/\n\t/g, "\t").replace(/\t\n/g, "\n");
}

function extractWord(zip: ZipReader, sink: TextSink): number {
  const xml = zip.readText("word/document.xml");
  if (xml === undefined) throw new OfficeTextError("Not a Word document (word/document.xml missing).");
  sink.push(wordXmlToText(xml));
  return 1;
}

// ---------------------------------------------------------------------------
// PowerPoint

export function drawingXmlToText(xml: string): string {
  const out: string[] = [];
  for (const m of xml.matchAll(/<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>|<a:br\b[^>]*\/>|<\/a:p>/g)) {
    if (m[1] !== undefined) out.push(decodeXmlEntities(m[1]));
    else out.push("\n");
  }
  return out.join("");
}

function slideOrder(zip: ZipReader): string[] {
  const presentation = zip.readText("ppt/presentation.xml");
  if (presentation) {
    const rels = readRelationships(zip, "ppt/_rels/presentation.xml.rels", "ppt");
    const ordered: string[] = [];
    for (const m of presentation.matchAll(/<p:sldId\b([^>]*?)\/?>/g)) {
      const target = rels.get(xmlAttr(m[1] ?? "", "r:id") ?? "");
      if (target && zip.has(target)) ordered.push(target);
    }
    if (ordered.length > 0) return ordered;
  }
  return zip
    .names()
    .map((n) => ({ n, m: /^ppt\/slides\/slide(\d+)\.xml$/.exec(n) }))
    .filter((x): x is { n: string; m: RegExpExecArray } => x.m !== null)
    .sort((a, b) => Number(a.m[1]) - Number(b.m[1]))
    .map((x) => x.n);
}

function extractPowerPoint(zip: ZipReader, sink: TextSink): number {
  const slides = slideOrder(zip);
  if (slides.length === 0 && !zip.has("ppt/presentation.xml")) throw new OfficeTextError("Not a PowerPoint presentation.");
  slides.forEach((path, i) => {
    if (sink.full) {
      sink.push("…");
      return;
    }
    const xml = zip.readText(path) ?? "";
    sink.push(`--- Slide ${i + 1} ---\n${drawingXmlToText(xml).trim()}\n\n`);
  });
  return slides.length;
}

// ---------------------------------------------------------------------------
// Excel

const MAX_PADDED_COLUMNS = 200;

export function columnIndex(ref: string): number {
  const letters = /^[A-Za-z]+/.exec(ref)?.[0] ?? "";
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function runText(xml: string): string {
  const cleaned = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, "");
  let s = "";
  for (const m of cleaned.matchAll(/<t(?:\s[^>]*)?>([^<]*)<\/t>/g)) s += decodeXmlEntities(m[1] ?? "");
  return s;
}

export function parseSharedStrings(xml: string): string[] {
  const out: string[] = [];
  for (const m of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>|<si\b[^>]*\/>/g)) out.push(runText(m[1] ?? ""));
  return out;
}

export function sheetXmlToRows(xml: string, shared: string[], sink: TextSink): void {
  for (const row of xml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)) {
    if (sink.full) {
      sink.push("…");
      return;
    }
    const cells: string[] = [];
    for (const c of (row[1] ?? "").matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const attrs = c[1] ?? "";
      const body = c[2] ?? "";
      const type = xmlAttr(attrs, "t");
      const v = /<v>([^<]*)<\/v>/.exec(body)?.[1];
      let value: string;
      if (type === "s") value = shared[Number(v)] ?? "";
      else if (type === "inlineStr") value = runText(body);
      else if (type === "b") value = v === "1" ? "TRUE" : v === "0" ? "FALSE" : "";
      else value = v === undefined ? "" : decodeXmlEntities(v);
      const ref = xmlAttr(attrs, "r");
      const col = ref ? columnIndex(ref) : -1;
      if (col >= 0 && col < MAX_PADDED_COLUMNS) {
        while (cells.length < col) cells.push("");
      }
      cells.push(value.replace(/[\t\r\n]+/g, " "));
    }
    while (cells.length > 0 && cells[cells.length - 1] === "") cells.pop();
    if (cells.length > 0) sink.push(cells.join("\t") + "\n");
  }
}

function extractExcel(zip: ZipReader, sink: TextSink): number {
  const workbook = zip.readText("xl/workbook.xml");
  if (workbook === undefined) throw new OfficeTextError("Not an Excel workbook (xl/workbook.xml missing).");
  const rels = readRelationships(zip, "xl/_rels/workbook.xml.rels", "xl");
  const sharedXml = zip.readText("xl/sharedStrings.xml");
  const shared = sharedXml ? parseSharedStrings(sharedXml) : [];
  let sheets = 0;
  for (const m of workbook.matchAll(/<sheet\b([^>]*?)\/?>/g)) {
    const attrs = m[1] ?? "";
    const target = rels.get(xmlAttr(attrs, "r:id") ?? "");
    if (!target || !zip.has(target)) continue;
    sheets++;
    if (sink.full) {
      sink.push("…");
      continue;
    }
    sink.push(`--- Sheet: ${xmlAttr(attrs, "name") ?? sheets} ---\n`);
    sheetXmlToRows(zip.readText(target) ?? "", shared, sink);
    sink.push("\n");
  }
  return sheets;
}

// ---------------------------------------------------------------------------

/** Extracts plain text from an Office Open XML file. Throws OfficeTextError for anything it will not handle. */
export function extractOfficeText(data: Buffer, kind: OfficeKind, limits: Partial<OfficeTextLimits> = {}): OfficeTextResult {
  const effective: OfficeTextLimits = { ...DEFAULT_OFFICE_LIMITS, ...limits };
  const zip = new ZipReader(data, effective);
  const sink = new TextSink(effective.maxChars);
  const parts = kind === "word" ? extractWord(zip, sink) : kind === "excel" ? extractExcel(zip, sink) : extractPowerPoint(zip, sink);
  const { text, truncated } = sink.result();
  return { kind, text, truncated, parts };
}
