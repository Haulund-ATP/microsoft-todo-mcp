import { randomUUID } from "node:crypto";
import { AsyncLocalStorage } from "node:async_hooks";
import type { Context } from "@microsoft/microsoft-graph-client";

/**
 * Per-Graph-call diagnostics captured for observability. Never includes
 * tokens, headers, or response bodies — only identifiers and status codes.
 */
export interface GraphCallDiagnostics {
  method: string;
  path: string;
  clientRequestId: string;
  graphRequestId?: string;
  httpStatus?: number;
}

const storage = new AsyncLocalStorage<GraphCallDiagnostics[]>();

/**
 * Runs `fn` with a fresh diagnostics collection scope. Every Graph call made
 * (directly or transitively) inside `fn`, via a client built with
 * `graphDiagnosticsMiddleware`, appends one record here.
 */
export async function withGraphDiagnostics<T>(fn: () => Promise<T>): Promise<{ result: T; diagnostics: GraphCallDiagnostics[] }> {
  const collected: GraphCallDiagnostics[] = [];
  const result = await storage.run(collected, fn);
  return { result, diagnostics: collected };
}

/** For call sites that only need the calls made during a single Graph function, without restructuring return values. */
export function currentGraphDiagnostics(): GraphCallDiagnostics[] | undefined {
  return storage.getStore();
}

function pathFromRequest(request: Context["request"]): string {
  const url = typeof request === "string" ? request : request.url;
  try {
    return new URL(url).pathname;
  } catch {
    return url;
  }
}

/**
 * Graph SDK middleware: tags every outgoing request with a unique
 * `client-request-id` and `return-client-request-id: true`, then records the
 * method/path/status/Graph-assigned `request-id` for diagnostic logging. Does
 * not read or modify the response body — only headers already set by the
 * time this middleware regains control from the next handler in the chain.
 */
export class GraphDiagnosticsMiddleware {
  private nextMiddleware!: { execute(context: Context): Promise<void> };

  async execute(context: Context): Promise<void> {
    const clientRequestId = randomUUID();
    const options = context.options ?? (context.options = {});
    const headers = { ...(options.headers as Record<string, string> | undefined) };
    headers["client-request-id"] = clientRequestId;
    headers["return-client-request-id"] = "true";
    options.headers = headers;

    await this.nextMiddleware.execute(context);

    const store = storage.getStore();
    if (store) {
      store.push({
        method: (context.options?.method as string | undefined) ?? "GET",
        path: pathFromRequest(context.request),
        clientRequestId,
        graphRequestId: context.response?.headers.get("request-id") ?? undefined,
        httpStatus: context.response?.status,
      });
    }
  }

  setNext(next: { execute(context: Context): Promise<void> }): void {
    this.nextMiddleware = next;
  }
}
