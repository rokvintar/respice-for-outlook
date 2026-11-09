const test = require("node:test");
const assert = require("node:assert/strict");
const r = require("../src/launchevent.js");

const me = "me@example.com";
const rcp = (emailAddress, recipientType = "user") => ({ emailAddress, displayName: emailAddress, recipientType });
const rules = (findings) => findings.map((f) => f.rule).sort();

test("internal-only new mail passes", () => {
  const f = r.respiceEvaluate({ to: [rcp("a@example.com")], cc: [rcp("b@sub.example.com")], composeType: "newMail", userEmail: me }, {});
  assert.deepEqual(f, []);
});

test("external recipient in To, Cc or Bcc is flagged", () => {
  for (const field of ["to", "cc", "bcc"]) {
    const input = { to: [], cc: [], bcc: [], composeType: "newMail", userEmail: me };
    input[field] = [rcp("x@customer.com")];
    assert.deepEqual(rules(r.respiceEvaluate(input, {})), ["external"], field);
  }
});

test("lookalike domain is external", () => {
  const f = r.respiceEvaluate({ to: [rcp("x@notexample.com")], userEmail: me }, {});
  assert.deepEqual(rules(f), ["external"]);
});

test("explicit internal domains override the user's domain", () => {
  const f = r.respiceEvaluate({ to: [rcp("a@example.com"), rcp("b@partner.org")], userEmail: me }, { internalDomains: "example.com\npartner.org" });
  assert.deepEqual(f, []);
});

test("mailing list pattern match in Bcc, with wildcards and case-insensitive", () => {
  const s = { listPatterns: "all@example.com\n*-team@example.com # comment" };
  assert.deepEqual(rules(r.respiceEvaluate({ bcc: [rcp("ALL@example.com")], userEmail: me }, s)), ["mailingList"]);
  assert.deepEqual(rules(r.respiceEvaluate({ cc: [rcp("rnd-team@example.com")], userEmail: me }, s)), ["mailingList"]);
  assert.deepEqual(r.respiceEvaluate({ cc: [rcp("rnd-team@example.com.evil")], userEmail: me }, { ...s, rules: { external: false } }), []);
});

test("Outlook distribution list is a list unless disabled", () => {
  const input = { to: [rcp("staff@example.com", "distributionList")], userEmail: me };
  assert.deepEqual(rules(r.respiceEvaluate(input, {})), ["mailingList"]);
  assert.deepEqual(r.respiceEvaluate(input, { treatOutlookDistributionListsAsLists: false }), []);
});

test("external + list both reported", () => {
  const f = r.respiceEvaluate({ to: [rcp("x@customer.com")], cc: [rcp("all@example.com")], userEmail: me }, { listPatterns: ["all@example.com"] });
  assert.deepEqual(rules(f), ["external", "mailingList"]);
});

test("too many recipients counts To+Cc+Bcc, deduplicated", () => {
  const many = Array.from({ length: 6 }, (_, i) => rcp(`u${i}@example.com`));
  const input = { to: many.slice(0, 3), cc: many.slice(3), bcc: [rcp("u0@example.com")], userEmail: me };
  assert.deepEqual(r.respiceEvaluate(input, { maxRecipients: 6 }), []);
  assert.deepEqual(rules(r.respiceEvaluate(input, { maxRecipients: 5 })), ["tooMany"]);
});

test("reply to one recipient passes, reply to several is flagged", () => {
  assert.deepEqual(r.respiceEvaluate({ to: [rcp("a@example.com")], composeType: "reply", userEmail: me }, {}), []);
  const f = r.respiceEvaluate({ to: [rcp("a@example.com")], cc: [rcp("b@example.com")], composeType: "reply", userEmail: me }, {});
  assert.deepEqual(rules(f), ["replyAll"]);
  assert.deepEqual(r.respiceEvaluate({ to: [rcp("a@example.com"), rcp("b@example.com")], composeType: "forward", userEmail: me }, {}), []);
});

test("each rule can be disabled", () => {
  const input = {
    to: [rcp("x@customer.com"), rcp("all@example.com")],
    cc: Array.from({ length: 20 }, (_, i) => rcp(`u${i}@example.com`)),
    composeType: "reply",
    userEmail: me
  };
  const base = { listPatterns: ["all@example.com"] };
  assert.deepEqual(rules(r.respiceEvaluate(input, base)), ["external", "mailingList", "replyAll", "tooMany"]);
  for (const k of ["external", "mailingList", "tooMany", "replyAll"]) {
    const f = r.respiceEvaluate(input, { ...base, rules: { [k]: false } });
    assert.ok(!rules(f).includes(k), k);
    assert.equal(f.length, 3, k);
  }
});

test("settings normalization tolerates junk", () => {
  const s = r.respiceNormalizeSettings({ maxRecipients: "abc", rules: { external: "yes" }, internalDomains: "@Example.COM\n*.foo.org" });
  assert.equal(s.maxRecipients, 15);
  assert.equal(s.rules.external, true);
  assert.deepEqual(s.internalDomains, ["example.com", "foo.org"]);
  assert.deepEqual(r.respiceNormalizeSettings(null), r.respiceNormalizeSettings(r.respiceNormalizeSettings(null)));
});

test("dialog message stays under 500 chars", () => {
  const ext = Array.from({ length: 50 }, (_, i) => rcp(`someone.with.a.long.name${i}@customer${i}.com`));
  const f = r.respiceEvaluate({ to: ext, composeType: "reply", userEmail: me }, {});
  const msg = r.respiceFormatMessage(f, false);
  assert.ok(msg.length <= 490, String(msg.length));
  assert.match(msg, /\+47 more/);
});

test("handler fails open and blocks only on findings", async () => {
  const run = (item, stored) =>
    new Promise((resolve) => {
      global.Office = {
        AsyncResultStatus: { Succeeded: "succeeded" },
        context: {
          roamingSettings: { get: () => stored },
          mailbox: { userProfile: { emailAddress: me }, item }
        }
      };
      r.onMessageSendHandler({ completed: resolve });
    });
  const field = (v) => ({ getAsync: (cb) => cb({ status: "succeeded", value: v }) });
  const item = (to) => ({ to: field(to), cc: field([]), bcc: field([]), getComposeTypeAsync: (cb) => cb({ status: "succeeded", value: { composeType: "newMail" } }) });

  assert.deepEqual(await run(item([rcp("a@example.com")]), null), { allowEvent: true });
  const blocked = await run(item([rcp("x@customer.com")]), null);
  assert.equal(blocked.allowEvent, false);
  assert.match(blocked.errorMessage, /External recipient: x@customer\.com/);
  const broken = { to: { getAsync: () => { throw new Error("boom"); } } };
  assert.deepEqual(await run(broken, null), { allowEvent: true });
  delete global.Office;
});

// --- Stall protection (v1.0.2) ---
function mockOffice(item, stored) {
  global.Office = {
    AsyncResultStatus: { Succeeded: "succeeded" },
    context: { roamingSettings: { get: () => stored }, mailbox: { userProfile: { emailAddress: me }, item } }
  };
}
const ok = (v) => ({ getAsync: (cb) => cb({ status: "succeeded", value: v }) });
const never = { getAsync: () => {} }; // Outlook never calls back
const composeNew = (cb) => cb({ status: "succeeded", value: { composeType: "newMail" } });

test("an Outlook call that never answers is timed out and the rest still evaluated", async () => {
  r.RESPICE_CONFIG.stepTimeoutMs = 50;
  r.RESPICE_CONFIG.watchdogMs = 1000;
  mockOffice({ to: ok([rcp("x@customer.com")]), cc: never, bcc: ok([]), getComposeTypeAsync: composeNew }, null);
  const res = await new Promise((resolve) => r.onMessageSendHandler({ completed: resolve }));
  assert.equal(res.allowEvent, false);
  assert.match(res.errorMessage, /External recipient/);
  delete global.Office;
});

test("watchdog completes the event if the handler gets stuck", async () => {
  r.RESPICE_CONFIG.stepTimeoutMs = 10000;
  r.RESPICE_CONFIG.watchdogMs = 50;
  mockOffice({ to: never, cc: never, bcc: never, getComposeTypeAsync: () => {} }, null);
  const t0 = Date.now();
  const res = await new Promise((resolve) => r.onMessageSendHandler({ completed: resolve }));
  assert.deepEqual(res, { allowEvent: true });
  assert.ok(Date.now() - t0 < 1000);
  delete global.Office;
});

test("if event.completed throws, it is retried with the simplest options", async () => {
  r.RESPICE_CONFIG.stepTimeoutMs = 50;
  r.RESPICE_CONFIG.watchdogMs = 1000;
  mockOffice({ to: ok([rcp("x@customer.com")]), cc: ok([]), bcc: ok([]), getComposeTypeAsync: composeNew }, null);
  const calls = [];
  await new Promise((resolve) =>
    r.onMessageSendHandler({
      completed: (o) => {
        calls.push(o);
        if (calls.length === 1) throw new Error("unsupported option");
        resolve();
      }
    })
  );
  assert.equal(calls.length, 2);
  assert.deepEqual(calls[1], { allowEvent: true });
  delete global.Office;
});

test("diagnostic mode always shows a dialog listing the steps", async () => {
  r.RESPICE_CONFIG.stepTimeoutMs = 50;
  r.RESPICE_CONFIG.watchdogMs = 1000;
  mockOffice({ to: ok([rcp("a@example.com")]), cc: never, bcc: ok([]), getComposeTypeAsync: composeNew }, { diagnostic: true });
  const res = await new Promise((resolve) => r.onMessageSendHandler({ completed: resolve }));
  assert.equal(res.allowEvent, false);
  assert.match(res.errorMessage, /Respice diagnostics/);
  assert.match(res.errorMessage, /cc TIMEOUT/);
  assert.match(res.errorMessage, /to ok/);
  assert.match(res.errorMessage, /Result: no warnings/);
  assert.ok(res.errorMessage.length <= 490);
  delete global.Office;
});

// --- Runtime bootstrap (v1.0.3) ---
const vm = require("node:vm");
const fs = require("node:fs");
const path = require("node:path");
const SRC = fs.readFileSync(path.join(__dirname, "../src/launchevent.js"), "utf8");

function loadLikeOutlook(extraGlobals) {
  const calls = { onReady: 0, associate: [] };
  const Office = {
    onReady: (cb) => { calls.onReady++; if (cb) cb({}); return Promise.resolve({}); },
    actions: { associate: (name, fn) => calls.associate.push([name, typeof fn]) }
  };
  vm.runInNewContext(SRC, Object.assign({ Office, setTimeout, clearTimeout, Promise }, extraGlobals));
  return calls;
}

test("event runtime page calls Office.onReady and registers the handler", () => {
  const calls = loadLikeOutlook({});
  assert.equal(calls.onReady, 1);
  assert.deepEqual(calls.associate, [["onMessageSendHandler", "function"]]);
});

test("task pane context neither registers the handler nor calls onReady from launchevent.js", () => {
  const calls = loadLikeOutlook({ RESPICE_SKIP_ASSOCIATE: true });
  assert.equal(calls.onReady, 0);
  assert.deepEqual(calls.associate, []);
});
