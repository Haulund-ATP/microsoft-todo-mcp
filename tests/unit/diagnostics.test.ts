import { describe, it, expect } from "vitest";
import type { Context } from "@microsoft/microsoft-graph-client";
import { GraphDiagnosticsMiddleware, withGraphDiagnostics } from "../../src/graph/diagnostics.js";

function fakeResponse(status: number, requestId?: string): Response {
  return {
    status,
    headers: { get: (key: string) => (key === "request-id" && requestId ? requestId : null) },
  } as unknown as Response;
}

describe("GraphDiagnosticsMiddleware", () => {
  it("tags the request with a unique client-request-id and return-client-request-id header", async () => {
    const middleware = new GraphDiagnosticsMiddleware();
    middleware.setNext({
      execute: async (context: Context) => {
        context.response = fakeResponse(200, "graph-req-1");
      },
    });

    const context: Context = { request: "https://graph.microsoft.com/v1.0/me/todo/lists/L/tasks/T", options: { method: "PATCH" } };
    await withGraphDiagnostics(async () => {
      await middleware.execute(context);
    });

    const headers = context.options?.headers as Record<string, string>;
    expect(headers["client-request-id"]).toMatch(/^[0-9a-f-]{36}$/);
    expect(headers["return-client-request-id"]).toBe("true");
  });

  it("records method, path, client-request-id, Graph request-id, and status per call", async () => {
    const middleware = new GraphDiagnosticsMiddleware();
    middleware.setNext({
      execute: async (context: Context) => {
        context.response = fakeResponse(201, "graph-req-2");
      },
    });

    const context: Context = {
      request: "https://graph.microsoft.com/v1.0/me/todo/lists/L/tasks/T/checklistItems",
      options: { method: "POST" },
    };
    const { diagnostics } = await withGraphDiagnostics(async () => {
      await middleware.execute(context);
    });

    expect(diagnostics).toHaveLength(1);
    expect(diagnostics[0]).toMatchObject({
      method: "POST",
      path: "/v1.0/me/todo/lists/L/tasks/T/checklistItems",
      graphRequestId: "graph-req-2",
      httpStatus: 201,
    });
    expect(diagnostics[0].clientRequestId).toBeTruthy();
  });

  it("assigns a distinct client-request-id to each call within the same diagnostics scope", async () => {
    const middleware = new GraphDiagnosticsMiddleware();
    middleware.setNext({
      execute: async (context: Context) => {
        context.response = fakeResponse(200);
      },
    });

    const { diagnostics } = await withGraphDiagnostics(async () => {
      const contextA: Context = { request: "https://graph.microsoft.com/v1.0/a", options: {} };
      const contextB: Context = { request: "https://graph.microsoft.com/v1.0/b", options: {} };
      await middleware.execute(contextA);
      await middleware.execute(contextB);
    });

    expect(diagnostics).toHaveLength(2);
    expect(diagnostics[0].clientRequestId).not.toBe(diagnostics[1].clientRequestId);
  });

  it("does not record anything outside of a withGraphDiagnostics scope", async () => {
    const middleware = new GraphDiagnosticsMiddleware();
    middleware.setNext({
      execute: async (context: Context) => {
        context.response = fakeResponse(200);
      },
    });
    const context: Context = { request: "https://graph.microsoft.com/v1.0/x", options: {} };
    await expect(middleware.execute(context)).resolves.toBeUndefined();
  });
});
