import { describe, expect, it } from "vitest";
import { MAX_FILES, MAX_TOTAL_BYTES, mimeOf, pickFiles, stripDataUrl } from "../utils/files";

const file = (name: string, size = 1000, type = "") => ({ name, size, type });

describe("attachments", () => {
  it("detects MIME types from the extension when the browser gives none", () => {
    expect(mimeOf(file("curve.PNG"))).toBe("image/png");
    expect(mimeOf(file("report.pdf"))).toBe("application/pdf");
    expect(mimeOf(file("x.jpg", 1, "image/jpg"))).toBe("image/jpeg");
    expect(mimeOf(file("data.xlsx"))).toBe("");
  });

  it("accepts images and PDFs, rejects other types", () => {
    const r = pickFiles([file("a.png"), file("b.pdf"), file("c.docx")], []);
    expect(r.accepted.map((f) => f.name)).toEqual(["a.png", "b.pdf"]);
    expect(r.error).toMatch(/c\.docx/);
  });

  it("enforces the file count and total size", () => {
    const many = Array.from({ length: MAX_FILES + 2 }, (_, i) => file(`p${i}.png`));
    expect(pickFiles(many, []).accepted).toHaveLength(MAX_FILES);
    const big = pickFiles([file("big.pdf", MAX_TOTAL_BYTES - 10), file("more.png", 100)], []);
    expect(big.accepted.map((f) => f.name)).toEqual(["big.pdf"]);
    expect(big.error).toMatch(/MB/);
    expect(pickFiles([file("x.png")], [{ size: MAX_TOTAL_BYTES }]).accepted).toHaveLength(0);
  });

  it("strips data URLs", () => {
    expect(stripDataUrl("data:image/png;base64,QUJD")).toBe("QUJD");
    expect(stripDataUrl("QUJD")).toBe("QUJD");
  });
});
