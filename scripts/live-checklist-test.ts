/**
 * Opt-in, single-account live reproduction test for the shared/non-shared
 * checklist-item read-after-write behavior. NOT run by `npm test` or CI —
 * vitest's config only picks up `tests/unit/**`, and this script isn't a
 * vitest test at all, since it needs the same live Azure credentials as the
 * running server (Key Vault, Table Storage) to build a real Graph client.
 *
 * Requires exactly one real, already-connected Microsoft account — no second
 * account, no test-account provisioning. Run from an environment where
 * `loadEnv()` succeeds (same env as the deployed Container App, e.g. locally
 * against the same Key Vault, or via `az containerapp exec`).
 *
 * Usage:
 *   npx tsx scripts/live-checklist-test.ts --connection-id <connectionId>
 *
 * It creates one dedicated task list ("MCP Live Test <timestamp>"), never
 * touching any of the account's existing lists, and deletes that list at the
 * end regardless of pass/fail.
 */
import { createGraphClientForConnection } from "../src/graph/client.js";
import * as todo from "../src/graph/todoApi.js";
import { withGraphDiagnostics } from "../src/graph/diagnostics.js";
import {
  verifyChecklistItemAdded,
  verifyChecklistItemUpdated,
  withChecklistConsistencyCheck,
} from "../src/graph/verification.js";

interface StepResult {
  step: string;
  ok: boolean;
  detail: string;
}

async function main(): Promise<void> {
  const connectionIdArgIndex = process.argv.indexOf("--connection-id");
  const connectionId = connectionIdArgIndex >= 0 ? process.argv[connectionIdArgIndex + 1] : undefined;
  if (!connectionId) {
    console.error("Usage: npx tsx scripts/live-checklist-test.ts --connection-id <connectionId>");
    process.exit(1);
  }

  const client = createGraphClientForConnection(connectionId);
  const results: StepResult[] = [];
  let listId: string | undefined;

  const record = (step: string, ok: boolean, detail: string) => {
    results.push({ step, ok, detail });
    console.log(`${ok ? "PASS" : "FAIL"}  ${step} — ${detail}`);
  };

  try {
    const list = await todo.createTaskList(client, `MCP Live Test ${new Date().toISOString()}`);
    listId = list.id;
    record("create dedicated test list", true, `listId=${list.id}`);

    const task = await todo.createTask(client, list.id, { title: "Live test task" });
    record("create task", true, `taskId=${task.id}, etag=${task.etag ?? "none"}`);

    const items: string[] = [];
    for (const name of ["Item A", "Item B", "Item C"]) {
      const { result: item, diagnostics } = await withGraphDiagnostics(async () => {
        const added = await todo.addChecklistItem(client, list.id, task.id, name);
        const verification = await verifyChecklistItemAdded(client, list.id, task.id, added.id);
        return { added, verification };
      });
      items.push(item.added.id);
      record(
        `add checklist item "${name}"`,
        item.verification.verificationStatus !== "inconsistent",
        `verificationStatus=${item.verification.verificationStatus}, graphRequestId=${diagnostics.at(-1)?.graphRequestId ?? "n/a"}`
      );
    }

    const { result: renamed, check: titleCheck } = await withChecklistConsistencyCheck(client, list.id, task.id, () =>
      todo.updateTask(client, list.id, task.id, { title: "Live test task (renamed)" })
    );
    record(
      "update_task title, compare checklist before/after",
      titleCheck.consistent,
      `before=${titleCheck.checklistCountBefore}, after=${titleCheck.checklistCountAfter}, etag=${renamed.etag ?? "none"}`
    );

    const { check: completeCheck } = await withChecklistConsistencyCheck(client, list.id, task.id, () =>
      todo.completeTask(client, list.id, task.id)
    );
    record(
      "complete_task, compare checklist before/after",
      completeCheck.consistent,
      `before=${completeCheck.checklistCountBefore}, after=${completeCheck.checklistCountAfter}, missing=${JSON.stringify(completeCheck.missingItemIds)}`
    );

    const { check: reopenCheck } = await withChecklistConsistencyCheck(client, list.id, task.id, () =>
      todo.reopenTask(client, list.id, task.id)
    );
    record(
      "reopen_task, compare checklist before/after",
      reopenCheck.consistent,
      `before=${reopenCheck.checklistCountBefore}, after=${reopenCheck.checklistCountAfter}, missing=${JSON.stringify(reopenCheck.missingItemIds)}`
    );

    const updateVerification = await (async () => {
      await todo.updateChecklistItem(client, list.id, task.id, items[1], { isChecked: true });
      return verifyChecklistItemUpdated(client, list.id, task.id, items[1], { isChecked: true });
    })();
    record(
      "update checklist item (check item B)",
      updateVerification.verificationStatus !== "inconsistent",
      `verificationStatus=${updateVerification.verificationStatus}, checklistCount=${updateVerification.checklistCount}`
    );

    const finalChecklist = await todo.listChecklistItems(client, list.id, task.id);
    const allThreeStillPresent = items.every((id) => finalChecklist.some((i) => i.id === id));
    record(
      "all three items still present after the update above",
      allThreeStillPresent,
      `expected ${items.length}, found ${finalChecklist.length}`
    );
  } finally {
    if (listId) {
      await todo.deleteTaskList(client, listId).catch((err) => console.error("cleanup failed:", err));
      console.log(`Cleaned up test list ${listId}`);
    }
  }

  const failed = results.filter((r) => !r.ok);
  console.log(`\n${results.length - failed.length}/${results.length} steps passed.`);
  if (failed.length > 0) {
    console.log("Failed steps:", failed.map((f) => f.step).join(", "));
    process.exit(1);
  }
}

main().catch((err) => {
  console.error("Live test crashed:", err);
  process.exit(1);
});
