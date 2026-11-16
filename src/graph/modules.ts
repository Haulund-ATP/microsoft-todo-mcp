import type { AccountType } from "../storage/types.js";

/**
 * Microsoft 365 capability modules. Each connection records which modules it
 * has consented to (incremental consent), and each profile records which
 * modules it exposes. Token acquisition always asks for BASE_SCOPES plus the
 * scopes of the one module a tool belongs to — never a global union — so
 * adding a module can never break silent refresh for connections that have
 * not consented to it (e.g. personal Microsoft accounts that only use To Do).
 */

export const MODULE_IDS = ["todo", "mail", "calendar", "contacts", "files", "sites", "teams", "teams_channels"] as const;
export type ModuleId = (typeof MODULE_IDS)[number];

export const BASE_SCOPES = ["User.Read", "openid", "profile", "offline_access"] as const;

export interface ModuleDefinition {
  id: ModuleId;
  label: string;
  scopes: readonly string[];
  /** Whether Microsoft supports the module's scopes for personal (MSA) accounts. */
  personalSupported: boolean;
  /** Whether the scopes always need tenant admin consent. */
  adminConsent: boolean;
}

export const MODULES: Record<ModuleId, ModuleDefinition> = {
  todo: { id: "todo", label: "To Do", scopes: ["Tasks.ReadWrite"], personalSupported: true, adminConsent: false },
  mail: { id: "mail", label: "Mail", scopes: ["Mail.ReadWrite", "Mail.Send"], personalSupported: true, adminConsent: false },
  calendar: { id: "calendar", label: "Calendar", scopes: ["Calendars.ReadWrite"], personalSupported: true, adminConsent: false },
  contacts: {
    id: "contacts",
    label: "Contacts & people",
    scopes: ["Contacts.ReadWrite", "People.Read"],
    personalSupported: true,
    adminConsent: false,
  },
  files: { id: "files", label: "OneDrive & Loop files", scopes: ["Files.ReadWrite.All"], personalSupported: true, adminConsent: false },
  sites: { id: "sites", label: "SharePoint sites", scopes: ["Sites.ReadWrite.All"], personalSupported: false, adminConsent: false },
  teams: {
    id: "teams",
    label: "Teams chats & channel posting",
    scopes: ["Chat.ReadWrite", "ChatMessage.Send", "Team.ReadBasic.All", "Channel.ReadBasic.All", "ChannelMessage.Send"],
    personalSupported: false,
    adminConsent: false,
  },
  teams_channels: {
    id: "teams_channels",
    label: "Teams channel messages (read)",
    scopes: ["ChannelMessage.Read.All"],
    personalSupported: false,
    adminConsent: true,
  },
};

/** Connections and profiles created before modules existed only ever had To Do. */
export const DEFAULT_MODULES: readonly ModuleId[] = ["todo"];

export function isModuleId(value: string): value is ModuleId {
  return (MODULE_IDS as readonly string[]).includes(value);
}

/** Scopes for one module's token request. */
export function scopesForModule(module: ModuleId): string[] {
  return [...BASE_SCOPES, ...MODULES[module].scopes];
}

/** Scopes for an interactive sign-in requesting several modules at once. */
export function scopesForModules(modules: readonly ModuleId[]): string[] {
  const set = new Set<string>(BASE_SCOPES);
  for (const m of modules) for (const s of MODULES[m].scopes) set.add(s);
  return [...set];
}

/** Drops modules Microsoft does not support for the account type. */
export function supportedModules(modules: readonly ModuleId[], accountType: AccountType): ModuleId[] {
  return modules.filter((m) => accountType === "organizational" || MODULES[m].personalSupported);
}

/**
 * Modules fully covered by a token response's granted scopes. Scope strings
 * may come back fully qualified ("https://graph.microsoft.com/Mail.Read"),
 * and are compared case-insensitively.
 */
export function modulesCoveredByScopes(grantedScopes: readonly string[]): ModuleId[] {
  const granted = new Set(grantedScopes.map((s) => s.replace(/^https:\/\/graph\.microsoft\.com\//i, "").toLowerCase()));
  return MODULE_IDS.filter((m) => MODULES[m].scopes.every((s) => granted.has(s.toLowerCase())));
}

/**
 * Table Storage cannot hold arrays, so module lists are stored as a
 * comma-separated string. Unknown ids are dropped; empty/missing means the
 * pre-modules default.
 */
export function parseModuleList(raw: string | undefined | null): ModuleId[] {
  if (raw === undefined || raw === null || raw.trim() === "") return [...DEFAULT_MODULES];
  const out: ModuleId[] = [];
  for (const part of raw.split(",")) {
    const id = part.trim();
    if (isModuleId(id) && !out.includes(id)) out.push(id);
  }
  return out;
}

export function serializeModuleList(modules: readonly ModuleId[]): string {
  return MODULE_IDS.filter((m) => modules.includes(m)).join(",");
}
