// Prints the task pane URL from manifest.prod.xml (used by use-online.bat).
import { readFileSync } from "node:fs";
const xml = readFileSync(new URL("../manifest.prod.xml", import.meta.url), "utf8");
const m = xml.match(/id="Taskpane\.Url" DefaultValue="([^"]+)"/);
console.log(m ? m[1] : "");
