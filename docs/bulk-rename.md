# Bulk task renaming and alphabetical grouping

`bulk_rename_tasks` updates up to 100 **existing** task titles in a single
MCP call. It uses the existing account/profile routing and Graph retry
logic, with at most four tasks processed concurrently. It is not a Graph
JSON batch and is not transactional: some tasks can succeed while others
fail. The response keeps the input order and reports every task separately.

## Prepare titles for the To Do app

Microsoft Graph's documented To Do task/list resources do not expose a
manual ordering field or the app's sort setting. Sorting an API response
does not change the app's displayed order. To group a freezer inventory,
prefix the titles with the location, leaving the original notes intact:

- `Hylde – Andebryst med skind`
- `Skuffe 1 – Hakket svinekød – 300 g – pakke 1/17`
- `Skuffe 2 – Hakket oksekød – 300 g – pakke 1/27`

Then select **Sort by → Alphabetically** in the To Do app. The tool does
not set this preference. Decide how to handle tasks with no known location
before including them; the tool never infers a location from meat type.

References:
- https://learn.microsoft.com/en-us/graph/api/resources/todotask
- https://learn.microsoft.com/en-us/graph/api/resources/todotasklist
- https://support.microsoft.com/en-us/todo/sort-and-search-in-lists

## Preview and apply

Read the tasks freshly and provide their exact titles as `expected_title`.
`dry_run` defaults to `true`, which reads the tasks without writing:

```json
{
  "list_id": "<list ID>",
  "dry_run": true,
  "updates": [
    {
      "task_id": "<task ID>",
      "expected_title": "Andebryst med skind",
      "title": "Hylde – Andebryst med skind"
    }
  ]
}
```

Use the same explicit updates with `dry_run: false` to apply. On a universal
`/mcp` endpoint, also supply `connection_id`. A bound profile uses its own
connection and rejects a conflicting connection ID, as other tools do.

The tool rejects duplicate task IDs, empty/whitespace-only new titles and
batches above 100. Each PATCH contains only `title`; notes, dates, status,
importance and checklist content are not rewritten. Checklist IDs are
read before and after the write, and the title is read back from Graph.

Per-task statuses:
- `preview`: the expected title matched; no write occurred.
- `unchanged`: the desired title is already present; no write occurred.
- `conflict`: the current title differed from both expected and desired;
  skipped without writing.
- `verified`: a fresh read confirmed the desired title and checklist IDs.
- `inconsistent`: the write returned, but verification did not confirm the
  desired title or checklist preservation. Inspect the task.
- `error`: a read, write or verification failed. If a write was attempted,
  it may already have succeeded; read the task before retrying.

`expected_title` is a best-effort guard, not atomic optimistic concurrency:
another client can edit between the read and PATCH. Repeating the same
explicit rename is safe when the desired title is already present. Do not
regenerate another prefix from an already-prefixed title. There is no
automatic rollback, reconstruction or deletion.

After deploying, refresh the connector's tool discovery in ChatGPT/Claude
if `bulk_rename_tasks` is not listed yet.
