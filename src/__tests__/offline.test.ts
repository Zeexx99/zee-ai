import { describe, expect, it } from "vitest";
import { networkFailures } from "../qol/offline";

describe("offline queue", () => {
  it("queues only network failures from this session", () => {
    const msgs = [
      { id: "a", status: "error", errorKind: "network" },
      { id: "b", status: "error", errorKind: "network" },
      { id: "c", status: "error", errorKind: "auth" },
      { id: "d", status: "done" },
    ];
    expect(networkFailures(msgs, new Set(["a", "c", "d"]))).toEqual(["a"]);
  });
});
