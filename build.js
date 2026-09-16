#!/usr/bin/env node
// Compiles app.jsx (JSX) into plain JS and injects it into index.html,
// between the BUILD:APP:START / BUILD:APP:END markers. Run this after any
// change to app.jsx: `npm run build`.
//
// Why: index.html used to ship raw JSX and transpile it in the browser via
// Babel Standalone on every page load/refresh, which is slow (measured
// ~900ms+ just for the transform on a fast machine, before React ever
// mounts). Precompiling once here removes that cost for every visitor.

const fs = require("fs");
const path = require("path");
const babel = require("@babel/core");

const root = __dirname;
const jsxPath = path.join(root, "app.jsx");
const htmlPath = path.join(root, "index.html");

const source = fs.readFileSync(jsxPath, "utf8");

const { code } = babel.transform(source, {
  filename: "app.jsx",
  presets: [["@babel/preset-react", { runtime: "classic" }]],
  comments: true,
  retainLines: false,
});

const html = fs.readFileSync(htmlPath, "utf8");

const startMarker = "<!-- BUILD:APP:START — generated from app.jsx by build.js. Do not edit inline; edit app.jsx and run `npm run build`. -->\n<script>\n";
const endMarker = "\n</script>\n<!-- BUILD:APP:END -->";

const startIdx = html.indexOf(startMarker);
const endIdx = html.indexOf(endMarker);

if (startIdx === -1 || endIdx === -1 || endIdx < startIdx) {
  throw new Error("Could not find BUILD:APP:START/END markers in index.html");
}

const before = html.slice(0, startIdx + startMarker.length);
const after = html.slice(endIdx);

fs.writeFileSync(htmlPath, before + code + after, "utf8");

console.log(`Built app.jsx (${source.length} bytes) -> index.html (${code.length} bytes of compiled JS)`);
