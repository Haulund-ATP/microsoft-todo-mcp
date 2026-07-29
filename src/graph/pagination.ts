import type { Client } from "@microsoft/microsoft-graph-client";
import { isThrottled, toGraphErrorInfo } from "./client.js";

/**
 * Follows `@odata.nextLink` to collect all pages, retrying on 429 with the
 * server-provided Retry-After (falling back to exponential backoff if
 * absent). Callers should still pass `$top` to bound page size, and most
 * MCP tools additionally cap total results (see tools/*) to keep responses
 * small and predictable for the calling LLM.
 */
export interface FetchAllPagesOptions {
  maxItems?: number;
  maxRetries?: number;
}

export async function fetchAllPages<T>(
  client: Client,
  initialRequest: { get(): Promise<{ value: T[]; "@odata.nextLink"?: string }> },
  options: FetchAllPagesOptions = {}
): Promise<T[]> {
  const maxItems = options.maxItems ?? 500;
  const maxRetries = options.maxRetries ?? 5;
  const results: T[] = [];

  let nextLink: string | undefined;
  let request = initialRequest;

  // eslint-disable-next-line no-constant-condition -- bounded by the break/return below
  while (true) {
    const page = await withThrottleRetry(() => request.get(), maxRetries);
    results.push(...page.value);
    nextLink = page["@odata.nextLink"];
    if (!nextLink || results.length >= maxItems) break;
    request = client.api(nextLink) as unknown as typeof initialRequest;
  }

  return results.slice(0, maxItems);
}

export async function withThrottleRetry<T>(fn: () => Promise<T>, maxRetries = 5): Promise<T> {
  let attempt = 0;
  // eslint-disable-next-line no-constant-condition -- bounded by the throw/return below
  while (true) {
    try {
      return await fn();
    } catch (err) {
      if (!isThrottled(err) || attempt >= maxRetries) throw err;
      const info = toGraphErrorInfo(err);
      const delaySeconds = info.retryAfterSeconds ?? Math.min(2 ** attempt, 30);
      await sleep(delaySeconds * 1000);
      attempt += 1;
    }
  }
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
