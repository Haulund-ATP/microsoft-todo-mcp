import type { Connection } from "../storage/types.js";
import type { ProfileBinding } from "../storage/types.js";

/**
 * Server-rendered admin page (no client-side framework — this is a small,
 * owner-only surface, so plain HTML keeps the attack surface and
 * dependency count minimal). All dynamic values are HTML-escaped.
 */

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);
}

function statusBadge(status: Connection["status"]): string {
  const colors: Record<Connection["status"], string> = {
    active: "#1a7f37",
    expired: "#9a6700",
    revoked: "#6e7781",
    consent_required: "#bc4c00",
    error: "#cf222e",
  };
  return `<span style="display:inline-block;padding:2px 8px;border-radius:10px;font-size:0.8rem;color:white;background:${colors[status]}">${escapeHtml(status)}</span>`;
}

export function renderAccountsPage(input: {
  connections: Connection[];
  profiles: ProfileBinding[];
  csrfToken: string;
}): string {
  const { connections, profiles, csrfToken } = input;
  const csrf = escapeHtml(csrfToken);

  const connectionsRows = connections
    .map((c) => {
      const boundProfiles = profiles.filter((p) => p.connectionId === c.connectionId).map((p) => p.profileAlias);
      return `<tr>
        <td>${escapeHtml(c.displayName)}<br><small>${escapeHtml(c.maskedEmail)}</small></td>
        <td>${escapeHtml(c.accountType)}</td>
        <td>${escapeHtml(c.alias)}</td>
        <td>${statusBadge(c.status)}</td>
        <td>${c.lastUsedAt ? escapeHtml(new Date(c.lastUsedAt).toLocaleString("en-DK")) : "never"}</td>
        <td>${boundProfiles.map(escapeHtml).join(", ") || "&mdash;"}</td>
        <td style="white-space:nowrap;">
          <a href="/admin/connect/start?reauth=${encodeURIComponent(c.connectionId)}">Reauth</a> |
          <form style="display:inline" method="post" action="/accounts/connections/${c.connectionId}/alias">
            <input type="hidden" name="_csrf" value="${csrf}">
            <input type="text" name="alias" placeholder="new alias" size="8" required pattern="[a-z0-9-]+">
            <button type="submit">Rename</button>
          </form>
          <form style="display:inline" method="post" action="/admin/connect/revoke/${c.connectionId}" onsubmit="return confirm('Revoke this connection?');">
            <input type="hidden" name="_csrf" value="${csrf}">
            <button type="submit">Revoke</button>
          </form>
          <form style="display:inline" method="post" action="/admin/connect/delete/${c.connectionId}" onsubmit="return confirm('Permanently delete this connection?');">
            <input type="hidden" name="_csrf" value="${csrf}">
            <button type="submit" style="color:#cf222e;">Delete</button>
          </form>
        </td>
      </tr>`;
    })
    .join("\n");

  const connectionOptions = connections
    .map((c) => `<option value="${c.connectionId}">${escapeHtml(c.alias)} (${escapeHtml(c.displayName)})</option>`)
    .join("\n");

  const profileRows = profiles
    .map(
      (p) => `<tr>
        <td><code>/mcp/${escapeHtml(p.profileAlias)}</code></td>
        <td>${escapeHtml(p.connectionId)}</td>
        <td>
          <form method="post" action="/accounts/profiles/unbind/${escapeHtml(p.profileAlias)}">
            <input type="hidden" name="_csrf" value="${csrf}">
            <button type="submit">Unbind</button>
          </form>
        </td>
      </tr>`
    )
    .join("\n");

  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><title>Microsoft To Do MCP — Accounts</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  body { font-family: system-ui, -apple-system, sans-serif; max-width: 960px; margin: 2rem auto; padding: 0 1rem; color: #1f2328; }
  table { border-collapse: collapse; width: 100%; margin-bottom: 2rem; }
  th, td { border: 1px solid #d0d7de; padding: 0.5rem; text-align: left; font-size: 0.9rem; vertical-align: top; }
  th { background: #f6f8fa; }
  h1, h2 { margin-top: 2rem; }
  form.inline { display: inline; }
  .actions a, .actions button { margin-right: 0.25rem; }
</style>
</head><body>
<h1>Microsoft To Do MCP — Accounts</h1>
<p><a href="/admin/connect/start">+ Connect a new Microsoft account</a></p>

<h2>Connections</h2>
<table>
  <thead><tr><th>Account</th><th>Type</th><th>Alias</th><th>Status</th><th>Last used</th><th>Bound profiles</th><th>Actions</th></tr></thead>
  <tbody>${connectionsRows || `<tr><td colspan="7">No connections yet.</td></tr>`}</tbody>
</table>

<h2>Profile bindings</h2>
<table>
  <thead><tr><th>Route</th><th>Connection id</th><th></th></tr></thead>
  <tbody>${profileRows || `<tr><td colspan="3">No profile bindings yet.</td></tr>`}</tbody>
</table>

<h3>Bind a profile</h3>
<form method="post" action="/accounts/profiles/bind">
  <input type="hidden" name="_csrf" value="${csrf}">
  <label>Profile alias: <input type="text" name="profile_alias" placeholder="personal" required pattern="[a-z0-9-]+"></label>
  <label>Connection:
    <select name="connection_id" required>
      <option value="" disabled selected>Choose a connection</option>
      ${connectionOptions}
    </select>
  </label>
  <button type="submit">Bind</button>
</form>

<h2>Session</h2>
<form method="post" action="/accounts/logout">
  <input type="hidden" name="_csrf" value="${csrf}">
  <button type="submit">Log out</button>
</form>
</body></html>`;
}
