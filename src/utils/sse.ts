/**
 * Incremental parser for Server-Sent Events as returned by Gemini `?alt=sse` endpoints.
 * Feed it decoded text chunks; it returns the complete `data:` payloads seen so far.
 */
export class SseParser {
  private buffer = "";

  push(chunk: string): string[] {
    this.buffer += chunk;
    const out: string[] = [];
    let idx: number;
    while ((idx = this.buffer.indexOf("\n")) >= 0) {
      const line = this.buffer.slice(0, idx).replace(/\r$/, "");
      this.buffer = this.buffer.slice(idx + 1);
      if (line.startsWith("data:")) {
        const data = line.slice(5).trim();
        if (data && data !== "[DONE]") out.push(data);
      }
    }
    return out;
  }

  /** Flush any trailing line without a newline at end of stream. */
  flush(): string[] {
    const rest = this.buffer.trim();
    this.buffer = "";
    if (rest.startsWith("data:")) {
      const data = rest.slice(5).trim();
      return data && data !== "[DONE]" ? [data] : [];
    }
    return [];
  }
}
