// Classic Outlook on Windows before Version 2403 runs event handlers in an engine that
// only supports ECMAScript 2016 and fails to load scripts using the conditional (?:) operator
// or async/await. See Microsoft's "Troubleshoot event-based add-ins".
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const acorn = require("acorn");
const walk = require("acorn-walk");

const file = path.join(__dirname, "../src/launchevent.js");
const src = fs.readFileSync(file, "utf8");

test("launchevent.js parses as ECMAScript 2016 script", () => {
  assert.doesNotThrow(() => acorn.parse(src, { ecmaVersion: 2016, sourceType: "script", locations: true }));
});

test("launchevent.js has no conditional (?:) operators", () => {
  const ast = acorn.parse(src, { ecmaVersion: 2016, sourceType: "script", locations: true });
  const hits = [];
  walk.simple(ast, { ConditionalExpression: (n) => hits.push(n.loc.start.line) });
  assert.deepEqual(hits, [], "ternaries at lines: " + hits.join(", "));
});
