// pdf.js in headless Chrome (system font fallback works as in a real browser): node pdfjsb.mjs jobs.tsv [mode]
// jobs: pdf \t page \t png \t dpi. Own loopback server, temp profile. No dialogs.
import { createRequire } from "module";
import fs from "fs";
import http from "http";
import os from "os";
import path from "path";
const DL = process.env.TODOKU_PK_DIR.replace(/\/$/, "") + "/../pdfjsdl/node_modules/";
const require = createRequire(DL + "x.js");
const { chromium } = require("playwright-core");
const jobs = fs.readFileSync(process.argv[2], "utf8").split("\n").filter(Boolean).map((l) => l.split("\t"));
const server = http.createServer((req, res) => {
  const u = decodeURIComponent(new URL(req.url, "http://x").pathname);
  if (u === "/") { res.setHeader("content-type", "text/html"); res.end("<!doctype html><title>r</title>"); return; }
  if (u.startsWith("/pdf/")) { res.end(fs.readFileSync(jobs[Number(u.slice(5))][0])); return; }
  const f = DL + "pdfjs-dist/" + u.replace(/^\/+/, "");
  if (!fs.existsSync(f)) { res.statusCode = 404; res.end(); return; }
  if (f.endsWith(".mjs")) res.setHeader("content-type", "text/javascript");
  res.end(fs.readFileSync(f));
});
await new Promise((r) => server.listen(0, "127.0.0.1", r));
const port = server.address().port;
const browser = await chromium.launch({ channel: "chrome", headless: true, args: ["--disable-gpu"] });
const page = await browser.newPage();
page.on("pageerror", (e) => console.error("pageerror", e.message));
await page.goto(`http://127.0.0.1:${port}/`);
await page.evaluate(async () => {
  const pdfjs = await import("/legacy/build/pdf.mjs");
  pdfjs.GlobalWorkerOptions.workerSrc = "/legacy/build/pdf.worker.mjs";
  window.pdfjs = pdfjs;
});
for (let i = 0; i < jobs.length; i++) {
  const [, pg, png, dpi] = jobs[i];
  const url = await page.evaluate(async ({ i, pg, dpi }) => {
    const pdfjs = window.pdfjs;
    const doc = await pdfjs.getDocument({ url: `/pdf/${i}`, cMapUrl: "/cmaps/", cMapPacked: true, standardFontDataUrl: "/standard_fonts/", useSystemFonts: true, verbosity: 0 }).promise;
    const p = await doc.getPage(Number(pg));
    const vp = p.getViewport({ scale: Number(dpi) / 72 });
    const c = document.createElement("canvas");
    c.width = Math.round(vp.width); c.height = Math.round(vp.height);
    const ctx = c.getContext("2d");
    ctx.fillStyle = "white"; ctx.fillRect(0, 0, c.width, c.height);
    await p.render({ canvasContext: ctx, viewport: vp, annotationMode: pdfjs.AnnotationMode.ENABLE }).promise;
    const u = c.toDataURL("image/png");
    await doc.destroy();
    return u;
  }, { i, pg, dpi });
  fs.writeFileSync(png, Buffer.from(url.split(",")[1], "base64"));
}
await browser.close();
server.close();
