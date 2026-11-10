#!/usr/bin/env node
// Generates manifest.xml from manifest.template.xml for a given hosting base URL.
// Usage: node scripts/build-manifest.mjs https://<github-user>.github.io/respice-for-outlook
import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const arg = process.argv[2];
if (!arg) {
  console.error("Usage: node scripts/build-manifest.mjs <https base url>");
  process.exit(1);
}
const url = new URL(arg);
if (url.protocol !== "https:") {
  console.error("Outlook add-ins must be served over HTTPS.");
  process.exit(1);
}
const base = url.href.replace(/\/+$/, "");
const version = JSON.parse(readFileSync(join(root, "package.json"), "utf8")).version; // single source of truth
if (!/^\d+\.\d+\.\d+$/.test(version)) {
  console.error(`package.json version must be x.y.z, got "${version}"`);
  process.exit(1);
}
const xml = readFileSync(join(root, "manifest.template.xml"), "utf8")
  .replace(/^<!-- Template\..*-->\n/m, "")
  .replaceAll("{{BASE_URL}}", base)
  .replaceAll("{{ORIGIN}}", url.origin)
  .replaceAll("{{VERSION}}", version);
writeFileSync(join(root, "manifest.xml"), xml);
console.log(`manifest.xml v${version} written for ${base}`);
