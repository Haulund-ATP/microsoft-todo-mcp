# Working with large lists

Tools for analysing, searching, cleaning up and moving many tasks at once.
They exist because a chat client working through a ~600-task list
otherwise makes dozens of slow calls: on 2026-10-13 ChatGPT made 15
`search_tasks` calls in ten seconds, re-read the whole list four times at
4.5–8.5 s each, and copied tasks to another profile one by one.

All of them use the usual profile routing: on a bound route
(`/mcp/{alias}`) they act on that profile's account; on the universal
`/mcp` route they need `connection_id`.

## Reading

- **`summarize_list`** — one call returns totals by status and importance,
  open/completed/overdue/recurring counts, tasks created and completed per
  month, the ten oldest open tasks, and groups of likely duplicates (titles
  equal after ignoring case, accents and punctuation; at most 50 groups,
  with `total_groups`). Reads up to 5,000 tasks; `truncated` is set if the
  list is larger.
- **`list_tasks`** — returns compact rows (no notes; use `get_task` for
  those). New optional inputs: `status` (`all`/`open`/`completed`),
  `created_after`, `created_before`, `completed_before`, and `sort`
  (`created_asc`, `created_desc`, `title`, `due`). With a date filter or a
  sort, the whole list is read and `matched` reports how many tasks pass,
  even beyond `max_items`.
- **`search_tasks`** — `queries` takes up to 50 title fragments in one
  call, and `list_id` may be omitted to search every list. Each list is read
  once and every query matched locally, ignoring case and accents. `query`
  (single text) still works.
- `list_completed_tasks`, `list_overdue_tasks` and `list_tasks_due_between`
  now also return compact rows.

Listing reads 100 tasks per Graph page with only the needed fields
(`$select`). Graph's To Do documentation doesn't state the page-size limit
or `$select` support for tasks, so if Graph answers 400 the server falls
back to its previous 50-per-page, all-fields reads for the rest of the
process and logs one warning.

## Writing

Both tools below preview by default (`dry_run: true`) and report each task
separately, in input order; partial success is possible.

### `copy_tasks`

Copies or moves up to 100 tasks from `list_id` to `target_list_id`. For a
list in another profile's Microsoft account, pass `target_profile` (e.g.
`"personal"`) or `target_connection_id`; without either, the target is in
the same account.

Copied: title, notes (formatting kept), importance, status, completion
date, due/start/reminder dates, recurrence, categories, and — unless
`include_checklist: false` — every checklist item with its checked state.
If Graph rejects that payload, the task is created with title, notes,
importance, status and due date only, and the result carries a warning.

A task whose title already exists in the target list is reported as
`exists` and skipped, so retrying a partly failed request doesn't create
duplicates; pass `allow_duplicates: true` to copy anyway. With
`mode: "move"`, the source task is deleted only after the copy has been
read back with the same title and checklist count.

Per-task statuses: `preview`, `exists`, `copied`, `moved`,
`inconsistent` (read-back didn't match; source untouched),
`copied_not_deleted` (move: the copy exists but the source delete failed),
`error`.

### `bulk_task_action`

Completes, reopens or deletes up to 200 tasks in one list. Every item gives
`task_id` and the `expected_title` from a fresh read; a different current
title is skipped as `conflict`. `delete` skips tasks that aren't completed
(`skipped_not_completed`) unless `only_completed: false`. Each write is
confirmed by a fresh read (a deleted task must return 404; complete/reopen
also check the checklist wasn't lost). Deletion cannot be undone.

Typical cleanup: `summarize_list` → `list_tasks` with
`status: "completed"`, `completed_before: "2025-01-01"` → `bulk_task_action`
with `action: "delete"` as a preview → the same with `dry_run: false`.

After deploying, refresh the connector's tool discovery in ChatGPT/Claude
if the new tools aren't listed.
