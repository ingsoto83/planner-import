import { afterAll, afterEach, beforeAll, describe, expect, it } from "vitest";
import { http, HttpResponse } from "msw";
import { setupServer } from "msw/node";

import { graphFetch } from "../graph-client";
import { GraphApiError } from "../graph-errors";

const server = setupServer();

beforeAll(() => server.listen());
afterEach(() => server.resetHandlers());
afterAll(() => server.close());

describe("graphFetch", () => {
  it("retries 429 responses using Retry-After", async () => {
    let attempts = 0;
    server.use(
      http.get("https://graph.microsoft.com/v1.0/retry", () => {
        attempts += 1;
        if (attempts === 1) {
          return HttpResponse.json(
            { error: { code: "TooManyRequests", message: "Retry later" } },
            { status: 429, headers: { "Retry-After": "0" } },
          );
        }
        return HttpResponse.json({ ok: true });
      }),
    );

    await expect(graphFetch<{ ok: boolean }>({ path: "/retry", token: "token" })).resolves.toEqual({ ok: true });
    expect(attempts).toBe(2);
  });

  it("does not retry functional 403 errors", async () => {
    let attempts = 0;
    server.use(
      http.get("https://graph.microsoft.com/v1.0/forbidden", () => {
        attempts += 1;
        return HttpResponse.json({ error: { code: "Forbidden", message: "No access" } }, { status: 403 });
      }),
    );

    await expect(graphFetch({ path: "/forbidden", token: "token" })).rejects.toBeInstanceOf(GraphApiError);
    expect(attempts).toBe(1);
  });
});
