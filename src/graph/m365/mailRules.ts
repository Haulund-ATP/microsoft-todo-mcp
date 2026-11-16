import type { Client } from "@microsoft/microsoft-graph-client";
import { withThrottleRetry } from "../pagination.js";
import { nextPageRequest } from "../../tools/m365/common.js";

/**
 * Outlook inbox rules for the signed-in user (`/me/mailFolders/inbox/messageRules`,
 * delegated MailboxSettings.ReadWrite — module "mail_rules"). Rules are
 * exposed in the same snake_case shape the tools take as input, so a rule
 * read with listRules can be passed back to updateRule unchanged.
 *
 * Safeguards:
 * - Forward and redirect actions keep sending mail out long after the
 *   conversation ends, so they are only accepted when the caller says
 *   forwarding is allowed (the tool layer ties this to the profile's
 *   allowSend), and the tool layer audits every rule change.
 * - `permanentDelete` is never offered: a rule may only move mail to
 *   Deleted Items, in line with the server's no-hard-delete policy.
 * - Folders in move/copy actions are resolved through Graph first, so a
 *   well-known name works and a wrong id fails before the rule is saved.
 */

const RULES_PATH = "/me/mailFolders/inbox/messageRules";

// ---------------------------------------------------------------- Graph shapes

interface GraphRecipient {
  emailAddress?: { name?: string; address?: string };
}

interface GraphSizeRange {
  minimumSize?: number;
  maximumSize?: number;
}

type GraphPredicates = Record<string, unknown> & {
  fromAddresses?: GraphRecipient[];
  sentToAddresses?: GraphRecipient[];
  withinSizeRange?: GraphSizeRange;
};

type GraphActions = Record<string, unknown> & {
  forwardTo?: GraphRecipient[];
  forwardAsAttachmentTo?: GraphRecipient[];
  redirectTo?: GraphRecipient[];
  moveToFolder?: string;
  copyToFolder?: string;
};

interface GraphRule {
  id: string;
  displayName?: string;
  sequence?: number;
  isEnabled?: boolean;
  hasError?: boolean;
  isReadOnly?: boolean;
  conditions?: GraphPredicates | null;
  exceptions?: GraphPredicates | null;
  actions?: GraphActions | null;
}

// ---------------------------------------------------------------- tool-facing shapes

export type Importance = "low" | "normal" | "high";
export type Sensitivity = "normal" | "personal" | "private" | "confidential";

/** Conditions and exceptions (Graph messageRulePredicates), snake_case. */
export interface RulePredicates {
  subject_contains?: string[];
  body_contains?: string[];
  body_or_subject_contains?: string[];
  sender_contains?: string[];
  recipient_contains?: string[];
  header_contains?: string[];
  from_addresses?: string[];
  sent_to_addresses?: string[];
  categories?: string[];
  importance?: Importance;
  sensitivity?: Sensitivity;
  has_attachments?: boolean;
  sent_to_me?: boolean;
  sent_only_to_me?: boolean;
  sent_cc_me?: boolean;
  sent_to_or_cc_me?: boolean;
  not_sent_to_me?: boolean;
  is_automatic_reply?: boolean;
  is_automatic_forward?: boolean;
  is_meeting_request?: boolean;
  is_meeting_response?: boolean;
  is_non_delivery_report?: boolean;
  is_read_receipt?: boolean;
  within_size_range_kb?: { min?: number; max?: number };
}

/** Rule actions (Graph messageRuleActions minus permanentDelete), snake_case. */
export interface RuleActions {
  move_to_folder?: string;
  copy_to_folder?: string;
  delete?: boolean;
  mark_as_read?: boolean;
  mark_importance?: Importance;
  assign_categories?: string[];
  stop_processing_rules?: boolean;
  forward_to?: string[];
  forward_as_attachment_to?: string[];
  redirect_to?: string[];
}

export interface MailRule {
  id: string;
  display_name?: string;
  sequence?: number;
  is_enabled?: boolean;
  has_error?: boolean;
  is_read_only?: boolean;
  conditions?: RulePredicates;
  exceptions?: RulePredicates;
  actions?: RuleActions;
  /** Display names of the folders in move/copy actions, when they could be looked up. */
  folder_names?: Record<string, string>;
}

// ---------------------------------------------------------------- field maps

/** snake_case predicate keys whose values pass through unchanged. */
const PLAIN_PREDICATES: Record<string, string> = {
  subject_contains: "subjectContains",
  body_contains: "bodyContains",
  body_or_subject_contains: "bodyOrSubjectContains",
  sender_contains: "senderContains",
  recipient_contains: "recipientContains",
  header_contains: "headerContains",
  categories: "categories",
  importance: "importance",
  sensitivity: "sensitivity",
  has_attachments: "hasAttachments",
  sent_to_me: "sentToMe",
  sent_only_to_me: "sentOnlyToMe",
  sent_cc_me: "sentCcMe",
  sent_to_or_cc_me: "sentToOrCcMe",
  not_sent_to_me: "notSentToMe",
  is_automatic_reply: "isAutomaticReply",
  is_automatic_forward: "isAutomaticForward",
  is_meeting_request: "isMeetingRequest",
  is_meeting_response: "isMeetingResponse",
  is_non_delivery_report: "isNonDeliveryReport",
  is_read_receipt: "isReadReceipt",
};

const RECIPIENT_PREDICATES: Record<string, string> = {
  from_addresses: "fromAddresses",
  sent_to_addresses: "sentToAddresses",
};

const PLAIN_ACTIONS: Record<string, string> = {
  delete: "delete",
  mark_as_read: "markAsRead",
  mark_importance: "markImportance",
  assign_categories: "assignCategories",
  stop_processing_rules: "stopProcessingRules",
};

const FOLDER_ACTIONS: Record<string, string> = {
  move_to_folder: "moveToFolder",
  copy_to_folder: "copyToFolder",
};

/** Actions that send mail outside the mailbox. */
export const FORWARD_ACTIONS: Record<string, string> = {
  forward_to: "forwardTo",
  forward_as_attachment_to: "forwardAsAttachmentTo",
  redirect_to: "redirectTo",
};

function invert(map: Record<string, string>): Record<string, string> {
  return Object.fromEntries(Object.entries(map).map(([k, v]) => [v, k]));
}

const PLAIN_PREDICATES_BACK = invert(PLAIN_PREDICATES);
const RECIPIENT_PREDICATES_BACK = invert(RECIPIENT_PREDICATES);
const PLAIN_ACTIONS_BACK = invert(PLAIN_ACTIONS);
const FOLDER_ACTIONS_BACK = invert(FOLDER_ACTIONS);
const FORWARD_ACTIONS_BACK = invert(FORWARD_ACTIONS);

function toRecipients(addresses: string[]): GraphRecipient[] {
  return addresses.map((address) => ({ emailAddress: { address } }));
}

function fromRecipients(list: GraphRecipient[] | undefined): string[] {
  return (list ?? []).map((r) => r.emailAddress?.address).filter((a): a is string => typeof a === "string" && a !== "");
}

// ---------------------------------------------------------------- input → Graph

export function buildPredicates(input: RulePredicates): GraphPredicates {
  const out: GraphPredicates = {};
  for (const [key, value] of Object.entries(input)) {
    if (value === undefined) continue;
    if (key in PLAIN_PREDICATES) out[PLAIN_PREDICATES[key]!] = value;
    else if (key in RECIPIENT_PREDICATES) out[RECIPIENT_PREDICATES[key]!] = toRecipients(value as string[]);
    else if (key === "within_size_range_kb") {
      const range = value as { min?: number; max?: number };
      if (range.min === undefined && range.max === undefined) throw new Error("within_size_range_kb needs min, max or both.");
      if (range.min !== undefined && range.max !== undefined && range.min > range.max) {
        throw new Error("within_size_range_kb: min must not be larger than max.");
      }
      out.withinSizeRange = {
        ...(range.min !== undefined ? { minimumSize: range.min } : {}),
        ...(range.max !== undefined ? { maximumSize: range.max } : {}),
      };
    } else throw new Error(`Unknown rule condition: ${key}.`);
  }
  return out;
}

export function forwardRecipientCount(actions: RuleActions | undefined): number {
  if (!actions) return 0;
  const unique = new Set<string>();
  for (const key of Object.keys(FORWARD_ACTIONS)) {
    for (const a of (actions[key as keyof RuleActions] as string[] | undefined) ?? []) unique.add(a.toLowerCase());
  }
  return unique.size;
}

export interface BuildActionsOptions {
  /** Forward and redirect actions are refused unless this is true. */
  allowForward: boolean;
}

interface FolderInfo {
  id: string;
  displayName?: string;
}

/** Folder lookups made during one operation, keyed by the reference used and by id. */
export type FolderCache = Map<string, FolderInfo>;

/** Resolves a folder id or well-known name to the folder's id; fails if Graph cannot find it. */
async function resolveFolder(client: Client, ref: string, cache?: FolderCache): Promise<FolderInfo> {
  const known = cache?.get(ref);
  if (known) return known;
  try {
    const folder = (await withThrottleRetry(() =>
      client.api(`/me/mailFolders/${encodeURIComponent(ref)}`).select("id,displayName").get()
    )) as FolderInfo;
    cache?.set(ref, folder).set(folder.id, folder);
    return folder;
  } catch (err) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status === 404 || status === 400) {
      throw new Error(`Mail folder not found: ${ref}. Use a folder id from mail_list_folders or a well-known name such as archive.`);
    }
    throw err;
  }
}

export async function buildActions(client: Client, input: RuleActions, opts: BuildActionsOptions, cache?: FolderCache): Promise<GraphActions> {
  const out: GraphActions = {};
  for (const [key, value] of Object.entries(input)) {
    if (value === undefined) continue;
    if (key in PLAIN_ACTIONS) out[PLAIN_ACTIONS[key]!] = value;
    else if (key in FOLDER_ACTIONS) out[FOLDER_ACTIONS[key]!] = (await resolveFolder(client, value as string, cache)).id;
    else if (key in FORWARD_ACTIONS) {
      if (!opts.allowForward) {
        throw new Error(
          "Forward and redirect rules are only allowed on a profile that allows sending, because they keep sending mail out automatically."
        );
      }
      if ((value as string[]).length === 0) throw new Error(`${key} needs at least one address.`);
      out[FORWARD_ACTIONS[key]!] = toRecipients(value as string[]);
    } else if (key === "permanent_delete") {
      throw new Error("Rules cannot delete mail permanently. Use delete, which moves it to Deleted Items.");
    } else throw new Error(`Unknown rule action: ${key}.`);
  }
  if (Object.keys(out).length === 0) throw new Error("A rule needs at least one action.");
  if (out.delete === true && (out.moveToFolder !== undefined || out.copyToFolder !== undefined)) {
    throw new Error("Use either delete or move/copy to a folder in one rule, not both.");
  }
  return out;
}

// ---------------------------------------------------------------- Graph → output

export function toPredicates(p: GraphPredicates | null | undefined): RulePredicates | undefined {
  if (!p) return undefined;
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(p)) {
    if (value === null || value === undefined || key.startsWith("@odata")) continue;
    if (Array.isArray(value) && value.length === 0) continue;
    if (key in PLAIN_PREDICATES_BACK) out[PLAIN_PREDICATES_BACK[key]!] = value;
    else if (key in RECIPIENT_PREDICATES_BACK) out[RECIPIENT_PREDICATES_BACK[key]!] = fromRecipients(value as GraphRecipient[]);
    else if (key === "withinSizeRange") {
      const range = value as GraphSizeRange;
      out.within_size_range_kb = {
        ...(range.minimumSize ? { min: range.minimumSize } : {}),
        ...(range.maximumSize ? { max: range.maximumSize } : {}),
      };
    }
    // Predicates this server does not expose (isEncrypted, messageActionFlag …) are left out
    // only when false/empty; a true one is reported under its Graph name so nothing is hidden.
    else if (value !== false) out[key] = value;
  }
  return Object.keys(out).length ? (out as RulePredicates) : undefined;
}

export function toActions(a: GraphActions | null | undefined): RuleActions | undefined {
  if (!a) return undefined;
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(a)) {
    if (value === null || value === undefined || key.startsWith("@odata")) continue;
    if (Array.isArray(value) && value.length === 0) continue;
    if (key in PLAIN_ACTIONS_BACK) {
      if (value === false) continue;
      out[PLAIN_ACTIONS_BACK[key]!] = value;
    } else if (key in FOLDER_ACTIONS_BACK) out[FOLDER_ACTIONS_BACK[key]!] = value;
    else if (key in FORWARD_ACTIONS_BACK) out[FORWARD_ACTIONS_BACK[key]!] = fromRecipients(value as GraphRecipient[]);
    else if (value !== false) out[key] = value;
  }
  return Object.keys(out).length ? (out as RuleActions) : undefined;
}

export function toMailRule(r: GraphRule): MailRule {
  const conditions = toPredicates(r.conditions);
  const exceptions = toPredicates(r.exceptions);
  const actions = toActions(r.actions);
  return {
    id: r.id,
    display_name: r.displayName,
    sequence: r.sequence,
    is_enabled: r.isEnabled,
    ...(r.hasError ? { has_error: true } : {}),
    ...(r.isReadOnly ? { is_read_only: true } : {}),
    ...(conditions ? { conditions } : {}),
    ...(exceptions ? { exceptions } : {}),
    ...(actions ? { actions } : {}),
  };
}

/** Adds folder_names for move/copy targets; lookups that fail are skipped. */
async function addFolderNames(client: Client, rules: MailRule[], cache: FolderCache = new Map()): Promise<void> {
  const ids = new Set<string>();
  for (const r of rules) {
    if (r.actions?.move_to_folder) ids.add(r.actions.move_to_folder);
    if (r.actions?.copy_to_folder) ids.add(r.actions.copy_to_folder);
  }
  const names = new Map<string, string>();
  for (const folderId of ids) {
    try {
      const f = await resolveFolder(client, folderId, cache);
      if (f.displayName) names.set(folderId, f.displayName);
    } catch {
      // A rule can point at a folder that was deleted since; leave the name out.
    }
  }
  for (const r of rules) {
    const own: Record<string, string> = {};
    for (const folderId of [r.actions?.move_to_folder, r.actions?.copy_to_folder]) {
      if (folderId && names.has(folderId)) own[folderId] = names.get(folderId)!;
    }
    if (Object.keys(own).length) r.folder_names = own;
  }
}

function rulePath(ruleId: string): string {
  return `${RULES_PATH}/${encodeURIComponent(ruleId)}`;
}

// ---------------------------------------------------------------- operations

/** Every inbox rule, in the order Outlook runs them (by sequence). */
export async function listRules(client: Client): Promise<{ rules: MailRule[] }> {
  const graphRules: GraphRule[] = [];
  let response = (await withThrottleRetry(() => client.api(RULES_PATH).get())) as { value?: GraphRule[]; "@odata.nextLink"?: string };
  graphRules.push(...(response.value ?? []));
  // Mailboxes hold at most a few hundred rules; follow Graph's own paging links to the end.
  for (let page = 0; response["@odata.nextLink"] && page < 20; page++) {
    const request = nextPageRequest(client, response["@odata.nextLink"]);
    response = (await withThrottleRetry(() => request.get())) as typeof response;
    graphRules.push(...(response.value ?? []));
  }
  const rules = graphRules.map(toMailRule).sort((a, b) => (a.sequence ?? Number.MAX_SAFE_INTEGER) - (b.sequence ?? Number.MAX_SAFE_INTEGER));
  await addFolderNames(client, rules);
  return { rules };
}

export async function getRule(client: Client, ruleId: string): Promise<MailRule> {
  const rule = toMailRule((await withThrottleRetry(() => client.api(rulePath(ruleId)).get())) as GraphRule);
  await addFolderNames(client, [rule]);
  return rule;
}

export interface CreateRuleInput {
  displayName: string;
  sequence?: number;
  isEnabled?: boolean;
  conditions?: RulePredicates;
  exceptions?: RulePredicates;
  actions: RuleActions;
}

export async function createRule(client: Client, input: CreateRuleInput, opts: BuildActionsOptions): Promise<MailRule> {
  const folders: FolderCache = new Map();
  const actions = await buildActions(client, input.actions, opts, folders);
  let sequence = input.sequence;
  if (sequence === undefined) {
    // Graph requires a sequence; by default the new rule runs after the existing ones.
    const existing = (await withThrottleRetry(() => client.api(RULES_PATH).select("id,sequence").get())) as { value?: GraphRule[] };
    sequence = Math.max(0, ...(existing.value ?? []).map((r) => r.sequence ?? 0)) + 1;
  }
  const body: Record<string, unknown> = {
    displayName: input.displayName,
    sequence,
    isEnabled: input.isEnabled ?? true,
    actions,
  };
  if (input.conditions) body.conditions = buildPredicates(input.conditions);
  if (input.exceptions) body.exceptions = buildPredicates(input.exceptions);
  const created = (await withThrottleRetry(() => client.api(RULES_PATH).post(body))) as GraphRule;
  const rule = toMailRule(created);
  await addFolderNames(client, [rule], folders);
  return rule;
}

export interface UpdateRuleInput {
  displayName?: string;
  sequence?: number;
  isEnabled?: boolean;
  conditions?: RulePredicates;
  exceptions?: RulePredicates;
  actions?: RuleActions;
}

export async function updateRule(client: Client, ruleId: string, input: UpdateRuleInput, opts: BuildActionsOptions): Promise<MailRule> {
  const fields = [input.displayName, input.sequence, input.isEnabled, input.conditions, input.exceptions, input.actions];
  if (fields.every((f) => f === undefined)) throw new Error("Nothing to update: pass at least one field to change.");
  const current = (await withThrottleRetry(() => client.api(rulePath(ruleId)).get())) as GraphRule;
  if (current.isReadOnly) throw new Error("This rule is read-only (created by another client or an administrator) and cannot be changed here.");
  // Switching on an existing forward/redirect rule (without replacing its actions)
  // would start sending mail out just like creating one.
  if (!opts.allowForward && input.isEnabled === true && input.actions === undefined && forwardRecipientCount(toActions(current.actions)) > 0) {
    throw new Error("This rule forwards or redirects mail; it can only be switched on from a profile that allows sending.");
  }
  const body: Record<string, unknown> = {};
  if (input.displayName !== undefined) body.displayName = input.displayName;
  if (input.sequence !== undefined) body.sequence = input.sequence;
  if (input.isEnabled !== undefined) body.isEnabled = input.isEnabled;
  if (input.conditions !== undefined) body.conditions = buildPredicates(input.conditions);
  if (input.exceptions !== undefined) body.exceptions = buildPredicates(input.exceptions);
  const folders: FolderCache = new Map();
  if (input.actions !== undefined) body.actions = await buildActions(client, input.actions, opts, folders);
  const updated = (await withThrottleRetry(() => client.api(rulePath(ruleId)).patch(body))) as GraphRule | undefined;
  const rule = toMailRule(updated?.id ? updated : ({ ...current, ...body, id: current.id } as GraphRule));
  await addFolderNames(client, [rule], folders);
  return rule;
}

/**
 * Deletes a rule. Rules have no recycle bin, so the removed definition is
 * returned in the same shape mail_create_rule takes, for recreating it.
 */
export async function deleteRule(client: Client, ruleId: string): Promise<{ deleted: true; removed_rule: MailRule }> {
  const current = (await withThrottleRetry(() => client.api(rulePath(ruleId)).get())) as GraphRule;
  if (current.isReadOnly) throw new Error("This rule is read-only (created by another client or an administrator) and cannot be deleted here.");
  await withThrottleRetry(() => client.api(rulePath(ruleId)).delete());
  return { deleted: true, removed_rule: toMailRule(current) };
}
