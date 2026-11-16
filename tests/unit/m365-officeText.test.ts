import { describe, expect, it } from "vitest";
import { deflateRawSync } from "node:zlib";
import {
  columnIndex,
  decodeXmlEntities,
  extractOfficeText,
  officeKindForExtension,
  OfficeTextError,
  wordXmlToText,
  ZipReader,
} from "../../src/graph/m365/officeText.js";

interface ZipInput {
  name: string;
  data: string | Buffer;
  method?: 0 | 8;
  flags?: number;
}

/** Builds a minimal ZIP archive (no CRCs: the reader does not check them). */
function buildZip(entries: ZipInput[]): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;
  for (const e of entries) {
    const raw = Buffer.isBuffer(e.data) ? e.data : Buffer.from(e.data, "utf8");
    const method = e.method ?? 8;
    const comp = method === 8 ? deflateRawSync(raw) : raw;
    const name = Buffer.from(e.name, "utf8");
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(e.flags ?? 0, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt32LE(comp.length, 18);
    local.writeUInt32LE(raw.length, 22);
    local.writeUInt16LE(name.length, 26);
    locals.push(local, name, comp);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(e.flags ?? 0, 8);
    central.writeUInt16LE(method, 10);
    central.writeUInt32LE(comp.length, 20);
    central.writeUInt32LE(raw.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, name);
    offset += 30 + name.length + comp.length;
  }
  const cd = Buffer.concat(centrals);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(cd.length, 12);
  eocd.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, cd, eocd]);
}

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function docx(body: string): Buffer {
  return buildZip([
    { name: "[Content_Types].xml", data: "<Types/>" },
    { name: "word/document.xml", data: `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>` },
  ]);
}

describe("officeKindForExtension", () => {
  it("maps OOXML extensions and ignores legacy binary formats", () => {
    expect(officeKindForExtension("DOCX")).toBe("word");
    expect(officeKindForExtension("xlsm")).toBe("excel");
    expect(officeKindForExtension("pptx")).toBe("powerpoint");
    expect(officeKindForExtension("doc")).toBeUndefined();
    expect(officeKindForExtension("xlsb")).toBeUndefined();
  });
});

describe("Word extraction", () => {
  it("keeps paragraphs, tabs and line breaks, decodes entities and skips deleted text and tab stops", () => {
    const body =
      '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr><w:r><w:t>Hej &amp; velkommen</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t xml:space="preserve">Navn: </w:t><w:tab/><w:t>Ø&#248;</w:t><w:br/><w:t>linje 2</w:t></w:r>' +
      "<w:del><w:r><w:delText>slettet</w:delText></w:r></w:del></w:p>";
    const result = extractOfficeText(docx(body), "word");
    expect(result.text).toBe("Hej & velkommen\nNavn: \tØø\nlinje 2");
    expect(result.truncated).toBe(false);
    expect(result.parts).toBe(1);
  });

  it("renders table rows as tab-separated lines", () => {
    const cell = (t: string) => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
    const xml = `<w:tbl><w:tr>${cell("A")}${cell("B")}</w:tr><w:tr>${cell("1")}${cell("2")}</w:tr></w:tbl>`;
    expect(wordXmlToText(xml).trim()).toBe("A\tB\n1\t2");
  });

  it("truncates at maxChars", () => {
    const body = Array.from({ length: 200 }, (_, i) => `<w:p><w:r><w:t>Afsnit nummer ${i}</w:t></w:r></w:p>`).join("");
    const result = extractOfficeText(docx(body), "word", { maxChars: 100 });
    expect(result.text.length).toBe(100);
    expect(result.truncated).toBe(true);
  });

  it("reads stored (uncompressed) entries too", () => {
    const zip = buildZip([{ name: "word/document.xml", data: `<w:document ${W}><w:p><w:r><w:t>Gemt</w:t></w:r></w:p></w:document>`, method: 0 }]);
    expect(extractOfficeText(zip, "word").text).toBe("Gemt");
  });
});

describe("Excel extraction", () => {
  it("resolves sheets through workbook relationships, shared and inline strings, booleans and empty columns", () => {
    const zip = buildZip([
      {
        name: "xl/workbook.xml",
        data: '<workbook><sheets><sheet name="Budget &amp; plan" sheetId="1" r:id="rId7"/><sheet name="Tom" sheetId="2" r:id="rId8"/></sheets></workbook>',
      },
      {
        name: "xl/_rels/workbook.xml.rels",
        data:
          '<Relationships><Relationship Id="rId7" Type="worksheet" Target="worksheets/sheetA.xml"/>' +
          '<Relationship Id="rId8" Type="worksheet" Target="/xl/worksheets/sheetB.xml"/></Relationships>',
      },
      {
        name: "xl/sharedStrings.xml",
        data: '<sst><si><t>Post</t></si><si><r><t>Be</t></r><r><t>løb</t></r><rPh><t>x</t></rPh></si></sst>',
      },
      {
        name: "xl/worksheets/sheetA.xml",
        data:
          '<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row>' +
          '<row r="2"><c r="A2" t="inlineStr"><is><t>Leje</t></is></c><c r="B2" t="b"><v>1</v></c><c r="C2"><v>1250.5</v></c><c r="D2" s="1"/></row>' +
          "</sheetData></worksheet>",
      },
      { name: "xl/worksheets/sheetB.xml", data: "<worksheet><sheetData/></worksheet>" },
    ]);
    const result = extractOfficeText(zip, "excel");
    expect(result.parts).toBe(2);
    expect(result.text).toBe("--- Sheet: Budget & plan ---\nPost\t\tBeløb\nLeje\tTRUE\t1250.5\n\n--- Sheet: Tom ---");
  });

  it("computes column indexes from cell references", () => {
    expect(columnIndex("A1")).toBe(0);
    expect(columnIndex("Z9")).toBe(25);
    expect(columnIndex("AA10")).toBe(26);
  });
});

describe("PowerPoint extraction", () => {
  it("orders slides by presentation.xml, not by file name", () => {
    const slide = (t: string) => `<p:sld><p:txBody><a:p><a:r><a:t>${t}</a:t></a:r></a:p><a:p><a:r><a:t>punkt</a:t></a:r></a:p></p:txBody></p:sld>`;
    const zip = buildZip([
      { name: "ppt/presentation.xml", data: '<p:presentation><p:sldIdLst><p:sldId id="256" r:id="rId3"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst></p:presentation>' },
      {
        name: "ppt/_rels/presentation.xml.rels",
        data: '<Relationships><Relationship Id="rId2" Target="slides/slide1.xml"/><Relationship Id="rId3" Target="slides/slide2.xml"/></Relationships>',
      },
      { name: "ppt/slides/slide1.xml", data: slide("Anden") },
      { name: "ppt/slides/slide2.xml", data: slide("Første") },
    ]);
    const result = extractOfficeText(zip, "powerpoint");
    expect(result.parts).toBe(2);
    expect(result.text).toBe("--- Slide 1 ---\nFørste\npunkt\n\n--- Slide 2 ---\nAnden\npunkt");
  });
});

describe("ZIP safety", () => {
  it("rejects data that is not a ZIP archive (e.g. an encrypted OLE file)", () => {
    expect(() => extractOfficeText(Buffer.from("\xd0\xcf\x11\xe0 not a zip at all, just bytes", "latin1"), "word")).toThrow(OfficeTextError);
  });

  it("stops a zip bomb at the per-entry decompression limit", () => {
    const bomb = buildZip([{ name: "word/document.xml", data: Buffer.alloc(5 * 1024 * 1024, 0x41) }]);
    expect(bomb.length).toBeLessThan(20_000);
    expect(() => extractOfficeText(bomb, "word", { maxEntryBytes: 1024 * 1024 })).toThrow(/too large/);
  });

  it("enforces the total budget across entries", () => {
    const part = `<worksheet><sheetData><row><c t="inlineStr"><is><t>${"x".repeat(600_000)}</t></is></c></row></sheetData></worksheet>`;
    const zip = buildZip([
      { name: "xl/workbook.xml", data: '<workbook><sheets><sheet name="a" r:id="r1"/><sheet name="b" r:id="r2"/></sheets></workbook>' },
      { name: "xl/_rels/workbook.xml.rels", data: '<Relationships><Relationship Id="r1" Target="s1.xml"/><Relationship Id="r2" Target="s2.xml"/></Relationships>' },
      { name: "xl/s1.xml", data: part },
      { name: "xl/s2.xml", data: part },
    ]);
    expect(() => extractOfficeText(zip, "excel", { maxTotalBytes: 1_000_000, maxChars: 10_000_000 })).toThrow(/too large/);
  });

  it("rejects encrypted entries and archives with too many entries", () => {
    const encrypted = buildZip([{ name: "word/document.xml", data: "<w:document/>", flags: 1 }]);
    expect(() => extractOfficeText(encrypted, "word")).toThrow(/Encrypted/);
    const many = buildZip(Array.from({ length: 20 }, (_, i) => ({ name: `f${i}.xml`, data: "x" })));
    expect(() => new ZipReader(many, { maxEntries: 10, maxEntryBytes: 1000, maxTotalBytes: 1000, maxChars: 10 })).toThrow(/too many entries/);
  });

  it("rejects a central directory that points past the archive", () => {
    const zip = buildZip([{ name: "word/document.xml", data: "<w:document/>" }]);
    zip.writeUInt32LE(zip.length, zip.length - 6); // central directory offset beyond EOCD
    expect(() => extractOfficeText(zip, "word")).toThrow(OfficeTextError);
  });

  it("reports a missing main part as an OfficeTextError", () => {
    expect(() => extractOfficeText(buildZip([{ name: "other.xml", data: "<x/>" }]), "word")).toThrow(/word\/document.xml missing/);
  });
});

describe("decodeXmlEntities", () => {
  it("decodes named and numeric entities and leaves invalid ones alone", () => {
    expect(decodeXmlEntities("&lt;a&gt; &amp;&quot;&apos; &#x41;&#66; &#0; &#x110000;")).toBe("<a> &\"' AB &#0; &#x110000;");
  });
});
