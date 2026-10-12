import { describe, expect, it } from "vitest";
import { builtinTools, callsOf, groundingSources, mergeStreamedParts, partsToMarkdown, turnToContent, type GenerateResponse } from "../services/geminiService";

describe("Gemini response parts", () => {
  it("renders text, skips thoughts and shows code execution as code blocks", () => {
    const md = partsToMarkdown([
      { text: "Let me compute." },
      { text: "secret reasoning", thought: true },
      { executableCode: { language: "PYTHON", code: "print(1+1)\n" } },
      { codeExecutionResult: { outcome: "OUTCOME_OK", output: "2\n" } },
      { text: "The answer is 2." },
    ]);
    expect(md).toContain("Let me compute.");
    expect(md).not.toContain("secret");
    expect(md).toContain("```python\nprint(1+1)\n```");
    expect(md).toContain("```text\n2\n```");
    expect(md.trim().endsWith("The answer is 2.")).toBe(true);
  });

  it("shows failed code runs", () => {
    expect(partsToMarkdown([{ codeExecutionResult: { outcome: "OUTCOME_FAILED", output: "NameError" } }])).toContain("(OUTCOME_FAILED)\nNameError");
  });

  it("extracts deduplicated grounding sources", () => {
    const r: GenerateResponse = {
      candidates: [
        {
          groundingMetadata: {
            groundingChunks: [{ web: { uri: "https://a.example/x", title: "a.example" } }, { web: { uri: "https://a.example/x" } }, { web: { uri: "https://b.example/y" } }, {}],
          },
        },
      ],
    };
    expect(groundingSources(r)).toEqual([
      { uri: "https://a.example/x", title: "a.example" },
      { uri: "https://b.example/y", title: "b.example" },
    ]);
  });

  it("merges streamed text but keeps signatures and calls intact", () => {
    const merged = mergeStreamedParts([
      { text: "Hel" },
      { text: "lo" },
      { functionCall: { name: "read_range", args: { address: "A1" } }, thoughtSignature: "sig1" },
      { text: "", thoughtSignature: "sig2" },
      { text: "a" },
      { text: "b" },
    ]);
    expect(merged).toEqual([
      { text: "Hello" },
      { functionCall: { name: "read_range", args: { address: "A1" } }, thoughtSignature: "sig1" },
      { text: "", thoughtSignature: "sig2" },
      { text: "ab" },
    ]);
    expect(callsOf(merged)).toEqual([{ name: "read_range", args: { address: "A1" }, id: undefined }]);
  });

  it("puts attachments before the text", () => {
    const c = turnToContent({ role: "user", text: "Digitise this curve", attachments: [{ name: "c.png", mimeType: "image/png", data: "QUJD", size: 3 }] });
    expect(c.parts).toEqual([{ inlineData: { mimeType: "image/png", data: "QUJD" } }, { text: "Digitise this curve" }]);
    expect(turnToContent({ role: "model", text: "ok" }).parts).toEqual([{ text: "ok" }]);
  });

  it("built-in tools follow Settings", () => {
    expect(builtinTools({ webSearch: false, codeExecution: false })).toEqual([]);
    expect(builtinTools({ webSearch: true, codeExecution: true })).toEqual([{ googleSearch: {} }, { codeExecution: {} }]);
  });
});
