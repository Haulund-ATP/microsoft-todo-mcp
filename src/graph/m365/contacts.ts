import type { Client } from "@microsoft/microsoft-graph-client";
import { graphSearchValue, nextPageRequest, pageOf, truncateText } from "../../tools/m365/common.js";

/**
 * Personal Outlook contacts (/me/contacts, Contacts.ReadWrite) and the
 * relevance-ranked people list (/me/people, People.Read) on Graph v1.0.
 *
 * Graph limits on /me/contacts: `$filter` only supports
 * `emailAddresses/any(a:a/address eq '…')` and `$search` is not documented
 * for contacts, so name search is done client-side over a bounded scan.
 */

export const DEFAULT_PAGE_SIZE = 25;
export const MAX_PAGE_SIZE = 100;
/** Upper bound for the client-side name search (pages of 100). */
export const MAX_SEARCH_SCAN = 1000;

type GraphObject = Record<string, unknown>;
interface GraphEmail {
  name?: string;
  address?: string;
}
interface GraphAddress {
  street?: string;
  city?: string;
  state?: string;
  postalCode?: string;
  countryOrRegion?: string;
}

const enc = encodeURIComponent;

const CONTACT_SUMMARY_SELECT = [
  "id",
  "displayName",
  "givenName",
  "surname",
  "companyName",
  "jobTitle",
  "emailAddresses",
  "mobilePhone",
  "businessPhones",
].join(",");

const CONTACT_DETAIL_SELECT = [
  CONTACT_SUMMARY_SELECT,
  "middleName",
  "nickName",
  "department",
  "officeLocation",
  "homePhones",
  "businessAddress",
  "homeAddress",
  "birthday",
  "personalNotes",
  "categories",
  "businessHomePage",
  "createdDateTime",
  "lastModifiedDateTime",
].join(",");

function clampTop(top: number | undefined): number {
  return Math.min(Math.max(Math.trunc(top ?? DEFAULT_PAGE_SIZE), 1), MAX_PAGE_SIZE);
}

function nonEmpty<T>(list: T[] | undefined): T[] | undefined {
  const filtered = (list ?? []).filter((x) => x !== null && x !== undefined && x !== "");
  return filtered.length ? filtered : undefined;
}

function shapeAddress(a: GraphAddress | undefined): GraphAddress | undefined {
  if (!a) return undefined;
  const entries = Object.entries(a).filter(([, v]) => typeof v === "string" && v !== "");
  return entries.length ? (Object.fromEntries(entries) as GraphAddress) : undefined;
}

export function shapeContactSummary(c: GraphObject) {
  return {
    id: c.id as string,
    display_name: (c.displayName as string | undefined) || undefined,
    given_name: (c.givenName as string | undefined) || undefined,
    surname: (c.surname as string | undefined) || undefined,
    company: (c.companyName as string | undefined) || undefined,
    job_title: (c.jobTitle as string | undefined) || undefined,
    emails: nonEmpty(((c.emailAddresses as GraphEmail[] | undefined) ?? []).map((e) => e.address ?? "")),
    mobile_phone: (c.mobilePhone as string | undefined) || undefined,
    business_phones: nonEmpty(c.businessPhones as string[] | undefined),
  };
}

export function shapeContactDetail(c: GraphObject) {
  const notes = (c.personalNotes as string | undefined) ?? "";
  const { text, truncated } = truncateText(notes.trim(), 10_000);
  return {
    ...shapeContactSummary(c),
    middle_name: (c.middleName as string | undefined) || undefined,
    nickname: (c.nickName as string | undefined) || undefined,
    department: (c.department as string | undefined) || undefined,
    office_location: (c.officeLocation as string | undefined) || undefined,
    home_phones: nonEmpty(c.homePhones as string[] | undefined),
    business_address: shapeAddress(c.businessAddress as GraphAddress | undefined),
    home_address: shapeAddress(c.homeAddress as GraphAddress | undefined),
    birthday: typeof c.birthday === "string" ? c.birthday.slice(0, 10) : undefined,
    notes: text || undefined,
    notes_truncated: truncated || undefined,
    categories: nonEmpty(c.categories as string[] | undefined),
    website: (c.businessHomePage as string | undefined) || undefined,
    created: c.createdDateTime as string | undefined,
    last_modified: c.lastModifiedDateTime as string | undefined,
  };
}

// ---------------------------------------------------------------------------
// Reads
// ---------------------------------------------------------------------------

export async function listContacts(client: Client, opts: { top?: number; pageToken?: string } = {}) {
  const request = opts.pageToken
    ? nextPageRequest(client, opts.pageToken)
    : client.api("/me/contacts").select(CONTACT_SUMMARY_SELECT).orderby("displayName").top(clampTop(opts.top));
  const page = pageOf<GraphObject>(await request.get());
  return { contacts: page.items.map(shapeContactSummary), next_page_token: page.next_page_token };
}

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function matches(c: GraphObject, needle: string): boolean {
  const haystack = [
    c.displayName,
    c.givenName,
    c.surname,
    c.companyName,
    ...((c.emailAddresses as GraphEmail[] | undefined) ?? []).flatMap((e) => [e.address, e.name]),
  ]
    .filter((v): v is string => typeof v === "string")
    .join("\n")
    .toLowerCase();
  return needle.split(/\s+/).every((word) => haystack.includes(word));
}

/**
 * Finds personal contacts. A full email address uses Graph's supported
 * `$filter` on emailAddresses; anything else is matched (every word,
 * case-insensitive) against names, company and emails over at most
 * MAX_SEARCH_SCAN contacts.
 */
export async function searchContacts(client: Client, query: string, opts: { limit?: number } = {}) {
  const needle = query.trim().toLowerCase();
  if (!needle) throw new Error("query must not be empty.");
  const limit = clampTop(opts.limit);
  if (EMAIL_RE.test(needle)) {
    const response = await client
      .api("/me/contacts")
      .filter(`emailAddresses/any(a:a/address eq '${needle.replace(/'/g, "''")}')`)
      .select(CONTACT_SUMMARY_SELECT)
      .top(limit)
      .get();
    return { contacts: pageOf<GraphObject>(response).items.map(shapeContactSummary), method: "email_filter" as const };
  }
  const found: GraphObject[] = [];
  let scanned = 0;
  let response = (await client.api("/me/contacts").select(CONTACT_SUMMARY_SELECT).top(MAX_PAGE_SIZE).get()) as {
    value?: GraphObject[];
    "@odata.nextLink"?: string;
  };
  for (;;) {
    for (const c of response.value ?? []) {
      scanned++;
      if (matches(c, needle)) found.push(c);
    }
    const next = response["@odata.nextLink"];
    if (found.length >= limit || !next || scanned >= MAX_SEARCH_SCAN) {
      return {
        contacts: found.slice(0, limit).map(shapeContactSummary),
        method: "name_scan" as const,
        scanned,
        scan_incomplete: next && found.length < limit ? true : undefined,
      };
    }
    response = await nextPageRequest(client, next).get();
  }
}

export async function getContact(client: Client, contactId: string) {
  const contact = (await client.api(`/me/contacts/${enc(contactId)}`).select(CONTACT_DETAIL_SELECT).get()) as GraphObject;
  return shapeContactDetail(contact);
}

// ---------------------------------------------------------------------------
// Writes
// ---------------------------------------------------------------------------

export interface ContactAddressInput {
  street?: string;
  city?: string;
  state?: string;
  postal_code?: string;
  country?: string;
}

export interface ContactInput {
  givenName?: string;
  surname?: string;
  displayName?: string;
  emails?: Array<{ address: string; name?: string }>;
  mobilePhone?: string | null;
  businessPhones?: string[];
  homePhones?: string[];
  companyName?: string | null;
  jobTitle?: string | null;
  department?: string | null;
  businessAddress?: ContactAddressInput | null;
  homeAddress?: ContactAddressInput | null;
  /** YYYY-MM-DD, or null to clear. */
  birthday?: string | null;
  notes?: string | null;
}

function addressPayload(a: ContactAddressInput | null): GraphAddress {
  if (!a) return {};
  const out: GraphAddress = {};
  if (a.street !== undefined) out.street = a.street;
  if (a.city !== undefined) out.city = a.city;
  if (a.state !== undefined) out.state = a.state;
  if (a.postal_code !== undefined) out.postalCode = a.postal_code;
  if (a.country !== undefined) out.countryOrRegion = a.country;
  return out;
}

/** Graph contact JSON with only the given fields; null clears a field. */
export function buildContactPayload(input: ContactInput): GraphObject {
  const p: GraphObject = {};
  if (input.givenName !== undefined) p.givenName = input.givenName;
  if (input.surname !== undefined) p.surname = input.surname;
  if (input.displayName !== undefined) p.displayName = input.displayName;
  if (input.emails !== undefined) {
    if (input.emails.length > 3) throw new Error("Outlook contacts hold at most 3 email addresses.");
    p.emailAddresses = input.emails.map((e) => (e.name ? { address: e.address, name: e.name } : { address: e.address }));
  }
  if (input.mobilePhone !== undefined) p.mobilePhone = input.mobilePhone ?? "";
  if (input.businessPhones !== undefined) p.businessPhones = input.businessPhones;
  if (input.homePhones !== undefined) p.homePhones = input.homePhones;
  if (input.companyName !== undefined) p.companyName = input.companyName ?? "";
  if (input.jobTitle !== undefined) p.jobTitle = input.jobTitle ?? "";
  if (input.department !== undefined) p.department = input.department ?? "";
  if (input.businessAddress !== undefined) p.businessAddress = addressPayload(input.businessAddress);
  if (input.homeAddress !== undefined) p.homeAddress = addressPayload(input.homeAddress);
  if (input.birthday !== undefined) {
    if (input.birthday !== null && !/^\d{4}-\d{2}-\d{2}$/.test(input.birthday)) throw new Error("birthday must be YYYY-MM-DD.");
    // Midday UTC keeps the date stable whatever time zone a client shows it in.
    p.birthday = input.birthday === null ? null : `${input.birthday}T12:00:00Z`;
  }
  if (input.notes !== undefined) p.personalNotes = input.notes ?? "";
  return p;
}

export async function createContact(client: Client, input: ContactInput) {
  if (!input.givenName && !input.surname && !input.displayName && !input.companyName && !input.emails?.length) {
    throw new Error("Give the contact at least a name, a company or an email address.");
  }
  const created = (await client.api("/me/contacts").post(buildContactPayload(input))) as GraphObject;
  return shapeContactDetail(created);
}

export async function updateContact(client: Client, contactId: string, input: ContactInput) {
  const payload = buildContactPayload(input);
  if (Object.keys(payload).length === 0) throw new Error("Nothing to update: pass at least one field to change.");
  // Graph may regenerate displayName when other fields change; keep the
  // current one unless the caller changes the name.
  if (input.displayName === undefined && input.givenName === undefined && input.surname === undefined) {
    const current = (await client.api(`/me/contacts/${enc(contactId)}`).select("id,displayName").get()) as GraphObject;
    if (typeof current.displayName === "string" && current.displayName) payload.displayName = current.displayName;
  }
  const updated = (await client.api(`/me/contacts/${enc(contactId)}`).patch(payload)) as GraphObject | undefined;
  return updated ? shapeContactDetail(updated) : { id: contactId };
}

export async function deleteContact(client: Client, contactId: string) {
  await client.api(`/me/contacts/${enc(contactId)}`).delete();
  return { deleted: true, contact_id: contactId };
}

// ---------------------------------------------------------------------------
// People
// ---------------------------------------------------------------------------

/**
 * People relevant to the signed-in user (frequent contacts, colleagues),
 * ranked by Graph; with a query, matched on display name and email.
 */
export async function searchPeople(client: Client, opts: { query?: string; top?: number; pageToken?: string } = {}) {
  let request;
  if (opts.pageToken) {
    request = nextPageRequest(client, opts.pageToken);
  } else {
    request = client
      .api("/me/people")
      .select("id,displayName,givenName,surname,scoredEmailAddresses,jobTitle,companyName,department,personType")
      .top(clampTop(opts.top));
    const query = opts.query?.trim();
    if (query) request = request.search(graphSearchValue(query));
  }
  const page = pageOf<GraphObject>(await request.get());
  return {
    people: page.items.map((p) => ({
      id: p.id as string,
      display_name: (p.displayName as string | undefined) || undefined,
      emails: ((p.scoredEmailAddresses as Array<{ address?: string; relevanceScore?: number }> | undefined) ?? [])
        .filter((e) => e.address)
        .map((e) => ({ address: e.address as string, score: e.relevanceScore })),
      job_title: (p.jobTitle as string | undefined) || undefined,
      company: (p.companyName as string | undefined) || undefined,
      department: (p.department as string | undefined) || undefined,
      kind: (p.personType as { class?: string; subclass?: string } | undefined)?.subclass,
    })),
    next_page_token: page.next_page_token,
  };
}
