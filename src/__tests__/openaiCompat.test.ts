import { describe, expect, it } from "vitest";
import { fromOpenAIResponse, modelIdsFromList, normalizeBaseUrl, OpenAIStreamDecoder, providerName, toJsonSchema, toOpenAIRequest } from "../services/openaiCompat";

describe("Other API key (OpenAI-compatible) adapter", () => {
  it("normalizes base URLs and names providers", () => {
    expect(normalizeBaseUrl(" https://api.groq.com/openai/v1/chat/completions/ ")).toBe("https://api.groq.com/openai/v1");
    expect(providerName("https://api.groq.com/openai/v1/")).toBe("Groq");
    expect(providerName("http://localhost:1234/v1")).toBe("localhost");
  });

  it("converts Gemini schemas to JSON Schema", () => {
    expect(toJsonSchema({ type: "OBJECT", properties: { a: { type: "STRING", nullable: true } }, required: ["a"] })).toEqual({
      type: "object",
      properties: { a: { type: "string" } },
      required: ["a"],
    });
  });

  it("builds chat messages, tools and tool results", () => {
    const req = toOpenAIRequest(
      {
        systemInstruction: { parts: [{ text: "SYS" }] },
        contents: [
          { role: "user", parts: [{ text: "hi" }] },
          { role: "model", parts: [{ text: "x", thought: true }, { functionCall: { name: "read", args: { a: 1 } } }] },
          { role: "user", parts: [{ functionResponse: { name: "read", response: { ok: true } } }] },
        ],
        generationConfig: { temperature: 0.4 },
        tools: [{ functionDeclarations: [{ name: "read", description: "Read", parameters: { type: "OBJECT", properties: {} } }] }, { googleSearch: {} }],
      },
      "openai/gpt-oss-120b",
      true
    );
    const msgs = req.messages as Array<Record<string, unknown>>;
    expect(msgs[0]).toEqual({ role: "system", content: "SYS" });
    expect(msgs[1]).toEqual({ role: "user", content: "hi" });
    const calls = msgs[2].tool_calls as Array<{ id: string; function: { arguments: string } }>;
    expect(msgs[2].content).toBeNull();
    expect(calls[0].function.arguments).toBe('{"a":1}');
    expect(msgs[3]).toEqual({ role: "tool", tool_call_id: calls[0].id, content: '{"ok":true}' });
    expect((req.tools as unknown[]).length).toBe(1);
    expect(req.temperature).toBe(0.4);
    expect(req.stream).toBe(true);
  });

  it("asks for JSON when a schema is requested", () => {
    const req = toOpenAIRequest({ contents: [{ role: "user", parts: [{ text: "rows" }] }], generationConfig: { responseMimeType: "application/json", responseSchema: { type: "ARRAY" } } }, "m", false);
    expect(req.response_format).toEqual({ type: "json_object" });
    expect(String((req.messages as Array<{ content: string }>)[0].content)).toContain('"type":"array"');
  });

  it("reads a plain answer with tool calls", () => {
    const r = fromOpenAIResponse({
      choices: [{ message: { content: "<think>hmm</think>Done", tool_calls: [{ id: "c1", function: { name: "write", arguments: '{"v":2}' } }] }, finish_reason: "tool_calls" }],
      usage: { prompt_tokens: 10, completion_tokens: 5, total_tokens: 15 },
    });
    expect(r.candidates?.[0].content?.parts).toEqual([{ text: "Done" }, { functionCall: { name: "write", args: { v: 2 }, id: "c1" } }]);
    expect(r.usageMetadata?.totalTokenCount).toBe(15);
  });

  it("decodes a stream, hiding <think> split across chunks and assembling tool calls", () => {
    const d = new OpenAIStreamDecoder();
    const ev = (o: unknown) => `data: ${JSON.stringify(o)}\n\n`;
    const texts: string[] = [];
    const collect = (cs: ReturnType<typeof d.push>) => cs.forEach((c) => c.candidates?.[0]?.content?.parts?.forEach((p) => p.text && texts.push(p.text)));
    collect(d.push(ev({ choices: [{ delta: { content: "<thi" } }] })));
    collect(d.push(ev({ choices: [{ delta: { content: "nk>secret</th" } }] })));
    const split = ev({ choices: [{ delta: { content: "ink>Hello " } }] });
    collect(d.push(split.slice(0, 20)));
    collect(d.push(split.slice(20)));
    collect(d.push(ev({ choices: [{ delta: { content: "world" } }] })));
    collect(d.push(ev({ choices: [{ delta: { tool_calls: [{ index: 0, id: "t1", function: { name: "read", arguments: '{"a"' } }] } }] })));
    collect(d.push(ev({ choices: [{ delta: { tool_calls: [{ index: 0, function: { arguments: ":1}" } }] }, finish_reason: "tool_calls" }], x_groq: { usage: { prompt_tokens: 3, completion_tokens: 4 } } })));
    collect(d.push("data: [DONE]\n\n"));
    const last = d.end();
    expect(texts.join("")).toBe("Hello world");
    const final = last[last.length - 1];
    expect(final.candidates?.[0].content?.parts).toEqual([{ functionCall: { name: "read", args: { a: 1 }, id: "t1" } }]);
    expect(final.candidates?.[0].finishReason).toBe("STOP");
    expect(final.usageMetadata?.totalTokenCount).toBe(7);
  });

  it("lists text models only", () => {
    expect(modelIdsFromList({ data: [{ id: "openai/gpt-oss-120b" }, { id: "whisper-large-v3" }, { id: "old", active: false }] })).toEqual(["openai/gpt-oss-120b"]);
  });
});
