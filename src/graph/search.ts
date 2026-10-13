import type { Client } from "@microsoft/microsoft-graph-client";
import * as todo from "./todoApi.js";
import { matchesQuery } from "./taskQuery.js";

/**
 * Searches one or many lists for one or many title fragments with a single
 * fetch per list. A chat client looking up 15 items used to make 15
 * search_tasks calls (each its own paged Graph query); here each list is
 * read once as compact tasks and every query is matched locally,
 * case- and accent-insensitively.
 */

/** Per-list read cap; far above any list seen in practice (~600). */
const SCAN_CAP = 5000;

export interface SearchMatch extends todo.CompactTask {
  list_id: string;
  list_name: string;
}

export async function searchTasksMulti(
  client: Client,
  queries: string[],
  options: { listId?: string; includeCompleted: boolean; maxPerQuery: number }
) {
  const lists = options.listId
    ? [await todo.getTaskList(client, options.listId)]
    : await todo.listTaskLists(client);

  const results = queries.map((query) => ({ query, match_count: 0, matches: [] as SearchMatch[] }));
  const truncatedLists: string[] = [];
  for (const list of lists) {
    const { tasks, truncated } = await todo.listTasksWithMeta(client, list.id, {
      compact: true,
      maxItems: SCAN_CAP,
      filter: options.includeCompleted ? undefined : "status ne 'completed'",
    });
    if (truncated) truncatedLists.push(list.id);
    for (const task of tasks) {
      for (const r of results) {
        if (!matchesQuery(task.title, r.query)) continue;
        r.match_count++;
        if (r.matches.length < options.maxPerQuery) {
          r.matches.push({ ...todo.toCompactTask(task), list_id: list.id, list_name: list.displayName });
        }
      }
    }
  }
  return {
    lists_scanned: lists.length,
    truncated_lists: truncatedLists.length ? truncatedLists : undefined,
    results,
  };
}
