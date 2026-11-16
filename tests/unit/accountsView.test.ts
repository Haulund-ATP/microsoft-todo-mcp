import { describe, expect, it } from "vitest";
import { renderAccountsPage } from "../../src/admin/view.js";
import type { Connection, ProfileBinding } from "../../src/storage/types.js";

const connection: Connection = {
  connectionId: "c1",
  alias: "work",
  accountType: "organizational",
  tenantId: "t",
  msalHomeAccountId: "h",
  objectId: "o",
  displayName: "Example User",
  maskedEmail: "e***@example.com",
  createdAt: "2026-11-16T00:00:00Z",
  lastUsedAt: null,
  status: "active",
  tokenCacheBlobName: "c1.cache",
  grantedModules: "todo,mail",
};
const profile: ProfileBinding = { profileAlias: "haa", connectionId: "c1", createdAt: "2026-11-16T00:00:00Z", allowSend: true };

describe("accounts page connector URLs", () => {
  const html = renderAccountsPage({ connections: [connection], profiles: [profile], csrfToken: "csrf", publicBaseUrl: "https://todo.example/" });

  it("shows both full endpoint URLs, ready to paste into a connector", () => {
    expect(html).toContain('value="https://todo.example/mcp/haa"');
    expect(html).toContain('value="https://todo.example/m365/haa"');
  });

  it("offers a copy button per URL wired by a same-origin script (CSP forbids inline scripts)", () => {
    expect(html).toContain('data-copy="https://todo.example/m365/haa"');
    expect(html).toContain('<script src="/accounts/copy.js" defer></script>');
    expect(html).not.toMatch(/<script>(?!<\/script>)/);
    expect(html).not.toMatch(/\son(click|focus)=/i);
  });
});
