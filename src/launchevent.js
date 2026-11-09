/*
 * Respice for Outlook: send-time recipient check for Outlook (Smart Alerts).
 *
 * This file is loaded in three places:
 *   1. Classic Outlook on Windows: directly, in a JavaScript-only runtime
 *      (no DOM, no imports, Office.onReady is not called).
 *   2. Outlook on the web / Mac / new Outlook: via src/commands.html.
 *   3. The settings task pane (src/taskpane.html), for the shared helpers.
 * It is also required by the Node unit tests. Keep it self-contained.
 *
 * MIT License. See LICENSE.
 */

var RESPICE_SETTINGS_KEY = "respiceSettings";

var RESPICE_DEFAULTS = {
  version: 1,
  rules: {
    external: true,     // 1. any recipient outside the internal domains
    mailingList: true,  // 2. any recipient matching a list pattern (or an Outlook distribution list)
    tooMany: true,      // 3. more than maxRecipients recipients
    replyAll: true      // 4. a reply that goes to more than one recipient
  },
  internalDomains: [],  // empty = use the domain of the signed-in user's own address
  listPatterns: [],     // e.g. "all@example.com", "*-all@example.com", "dl-*@example.com"
  treatOutlookDistributionListsAsLists: true,
  maxRecipients: 15
};

var RESPICE_LABELS = {
  external: "External recipient",
  mailingList: "Mailing list",
  tooMany: "Too many recipients",
  replyAll: "Reply to more than the sender"
};

// ---------------------------------------------------------------------------
// Pure helpers (unit-tested in Node)
// ---------------------------------------------------------------------------

function respiceParseLines(text) {
  if (Array.isArray(text)) text = text.join("\n");
  if (typeof text !== "string") return [];
  var out = [];
  var seen = {};
  text.split(/[\r\n,;]+/).forEach(function (line) {
    var v = line.replace(/#.*$/, "").trim().toLowerCase();
    if (v && !seen[v]) {
      seen[v] = true;
      out.push(v);
    }
  });
  return out;
}

function respiceNormalizeSettings(raw) {
  var d = RESPICE_DEFAULTS;
  var s = raw && typeof raw === "object" ? raw : {};
  var rules = s.rules && typeof s.rules === "object" ? s.rules : {};
  var max = parseInt(s.maxRecipients, 10);
  return {
    version: d.version,
    rules: {
      external: typeof rules.external === "boolean" ? rules.external : d.rules.external,
      mailingList: typeof rules.mailingList === "boolean" ? rules.mailingList : d.rules.mailingList,
      tooMany: typeof rules.tooMany === "boolean" ? rules.tooMany : d.rules.tooMany,
      replyAll: typeof rules.replyAll === "boolean" ? rules.replyAll : d.rules.replyAll
    },
    internalDomains: respiceParseLines(s.internalDomains || []).map(function (x) {
      return x.replace(/^.*@/, "").replace(/^\*?\./, "");
    }),
    listPatterns: respiceParseLines(s.listPatterns || []),
    treatOutlookDistributionListsAsLists:
      typeof s.treatOutlookDistributionListsAsLists === "boolean"
        ? s.treatOutlookDistributionListsAsLists
        : d.treatOutlookDistributionListsAsLists,
    maxRecipients: isFinite(max) && max >= 1 ? max : d.maxRecipients,
    diagnostic: s.diagnostic === true
  };
}

function respiceDomainOf(address) {
  if (typeof address !== "string") return "";
  var at = address.lastIndexOf("@");
  if (at < 0) return "";
  return address.slice(at + 1).trim().toLowerCase();
}

function respiceIsInternal(address, internalDomains) {
  var dom = respiceDomainOf(address);
  if (!dom) return true; // no SMTP domain (unresolved entry): cannot judge, do not flag
  for (var i = 0; i < internalDomains.length; i++) {
    var d = internalDomains[i];
    if (dom === d || dom.slice(-(d.length + 1)) === "." + d) return true;
  }
  return false;
}

function respiceGlobToRegExp(pattern) {
  var escaped = pattern.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".");
  return new RegExp("^" + escaped + "$", "i");
}

function respiceMatchesList(recipient, patterns, treatDl) {
  if (treatDl && recipient.recipientType === "distributionList") return true;
  var addr = (recipient.emailAddress || "").toLowerCase();
  if (!addr) return false;
  for (var i = 0; i < patterns.length; i++) {
    if (respiceGlobToRegExp(patterns[i]).test(addr)) return true;
  }
  return false;
}

function respiceUniqueRecipients(input) {
  var all = [].concat(input.to || [], input.cc || [], input.bcc || []);
  var seen = {};
  var out = [];
  all.forEach(function (r) {
    if (!r) return;
    var key = (r.emailAddress || r.displayName || "").toLowerCase();
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push(r);
  });
  return out;
}

/**
 * input:    { to: [], cc: [], bcc: [], composeType: "newMail"|"reply"|"forward"|null, userEmail: "" }
 *           recipients: { emailAddress, displayName, recipientType }
 * settings: normalized settings
 * returns:  array of { rule, label, items: [addresses], detail }
 */
function respiceEvaluate(input, settings) {
  var s = respiceNormalizeSettings(settings);
  var recipients = respiceUniqueRecipients(input);
  var findings = [];

  var internal = s.internalDomains.length
    ? s.internalDomains
    : [respiceDomainOf(input.userEmail)].filter(Boolean);

  if (s.rules.external && internal.length) {
    var ext = recipients.filter(function (r) {
      return !respiceIsInternal(r.emailAddress, internal);
    });
    if (ext.length) {
      findings.push({
        rule: "external",
        label: RESPICE_LABELS.external,
        items: ext.map(respiceShowName),
        detail: "outside " + internal.join(", ")
      });
    }
  }

  if (s.rules.mailingList) {
    var lists = recipients.filter(function (r) {
      return respiceMatchesList(r, s.listPatterns, s.treatOutlookDistributionListsAsLists);
    });
    if (lists.length) {
      findings.push({ rule: "mailingList", label: RESPICE_LABELS.mailingList, items: lists.map(respiceShowName), detail: "" });
    }
  }

  if (s.rules.tooMany && recipients.length > s.maxRecipients) {
    findings.push({
      rule: "tooMany",
      label: RESPICE_LABELS.tooMany,
      items: [],
      detail: recipients.length + " recipients (limit " + s.maxRecipients + ")"
    });
  }

  if (s.rules.replyAll && respiceIsReply(input) && recipients.length > 1) {
    findings.push({
      rule: "replyAll",
      label: RESPICE_LABELS.replyAll,
      items: [],
      detail: "this reply goes to " + recipients.length + " recipients. Did you mean Reply all?"
    });
  }

  return findings;
}

// Outlook's compose type is unreliable at send time (Mac/web report "newMail" for replies),
// so a reply is also recognised by its subject prefix. Forwards are never treated as replies.
var RESPICE_REPLY_PREFIX = /^\s*(\[[^\]]*\]\s*)*(re|aw|sv|odg|vs|rif|antw|r|答复|回复)\s*(\[\d+\])?\s*[:：]/i;

function respiceIsReplySubject(subject) {
  return typeof subject === "string" && RESPICE_REPLY_PREFIX.test(subject);
}

var RESPICE_FORWARD_PREFIX = /^\s*(\[[^\]]*\]\s*)*(fw|fwd|wg|tr|rv|vb|doorst|posl|i|转发)\s*[:：]/i;

function respiceIsForwardSubject(subject) {
  return typeof subject === "string" && RESPICE_FORWARD_PREFIX.test(subject);
}

// Reply signals, strongest first:
//   1. compose type "reply" (reliable only while the window is open; Mac/web report "newMail" at send time)
//   2. inReplyTo: ID of the message being answered, stored on the draft itself (Mailbox 1.14)
//   3. reply prefix in the subject (RE:, AW:, Odg: ...)
// Forwards are excluded by compose type or by a forward prefix (FW:, WG:, TR: ...).
function respiceIsReply(input) {
  if (input.composeType === "reply") return true;
  if (input.composeType === "forward") return false;
  if (respiceIsForwardSubject(input.subject)) return false;
  if (input.inReplyTo) return true;
  return respiceIsReplySubject(input.subject);
}

function respiceShowName(r) {
  return r.emailAddress || r.displayName || "?";
}

function respiceSummarizeItems(items, max) {
  if (!items.length) return "";
  var shown = items.slice(0, max).join(", ");
  return items.length > max ? shown + " +" + (items.length - max) + " more" : shown;
}

// Smart Alerts truncates long dialog text, so keep it compact.
function respiceFormatMessage(findings, markdown) {
  var lines = findings.map(function (f) {
    var what = respiceSummarizeItems(f.items, 3);
    var parts = [what, f.detail].filter(Boolean).join(" · ");
    return markdown ? "- **" + f.label + ":** " + parts : "• " + f.label + ": " + parts;
  });
  var head = "Check before sending:";
  var text = head + "\n" + lines.join("\n");
  return text.length > 490 ? text.slice(0, 487) + "..." : text;
}

// ---------------------------------------------------------------------------
// Office glue
// ---------------------------------------------------------------------------

function respiceLoadSettings() {
  try {
    return respiceNormalizeSettings(Office.context.roamingSettings.get(RESPICE_SETTINGS_KEY));
  } catch (e) {
    return respiceNormalizeSettings(null);
  }
}

var RESPICE_CONFIG = {
  stepTimeoutMs: 2000, // per Outlook API call
  watchdogMs: 4000     // whole handler; Outlook shows "taking longer than expected" after 5 s
};

function respiceGetAsync(field) {
  return new Promise(function (resolve) {
    if (!field || typeof field.getAsync !== "function") return resolve([]);
    field.getAsync(function (res) {
      resolve(res.status === Office.AsyncResultStatus.Succeeded && res.value ? res.value : []);
    });
  });
}

// Resolves { type, raw } where raw describes exactly what Outlook returned (for diagnostics).
function respiceGetComposeType(item) {
  return new Promise(function (resolve) {
    if (!item || typeof item.getComposeTypeAsync !== "function") {
      return resolve({ type: null, raw: "getComposeTypeAsync not available" });
    }
    item.getComposeTypeAsync(function (res) {
      if (res.status === Office.AsyncResultStatus.Succeeded && res.value) {
        resolve({ type: res.value.composeType, raw: JSON.stringify(res.value) });
      } else {
        resolve({ type: null, raw: "failed: " + (res.error && res.error.message) });
      }
    });
  });
}

// Runs makePromise() with a timeout; never rejects. Records "label ok 12ms" / "label TIMEOUT" / "label ERROR ..." in diag.
function respiceTimed(label, makePromise, fallback, diag) {
  var t0 = Date.now();
  return new Promise(function (resolve) {
    var settled = false;
    function settle(value, note) {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      diag.push(label + " " + note);
      resolve(value);
    }
    var timer = setTimeout(function () {
      settle(fallback, "TIMEOUT");
    }, RESPICE_CONFIG.stepTimeoutMs);
    try {
      makePromise().then(
        function (v) {
          settle(v, "ok " + (Date.now() - t0) + "ms");
        },
        function (e) {
          settle(fallback, "ERROR " + (e && e.message));
        }
      );
    } catch (e) {
      settle(fallback, "ERROR " + (e && e.message));
    }
  });
}

// Synchronous property, Mailbox 1.14+. Empty string when unsupported or not a reply.
function respiceGetInReplyTo(item) {
  try {
    var v = item && item.inReplyTo;
    return typeof v === "string" ? v : "";
  } catch (e) {
    return "";
  }
}

function respiceGetSubject(item) {
  return new Promise(function (resolve) {
    var f = item && item.subject;
    if (!f || typeof f.getAsync !== "function") return resolve("");
    f.getAsync(function (res) {
      resolve(res.status === Office.AsyncResultStatus.Succeeded && typeof res.value === "string" ? res.value : "");
    });
  });
}

function respiceGatherInput(item, diag) {
  diag = diag || [];
  var noType = { type: null, raw: "unavailable" };
  return Promise.all([
    respiceTimed("to", function () { return respiceGetAsync(item.to); }, [], diag),
    respiceTimed("cc", function () { return respiceGetAsync(item.cc); }, [], diag),
    respiceTimed("bcc", function () { return respiceGetAsync(item.bcc); }, [], diag),
    respiceTimed("composeType", function () { return respiceGetComposeType(item); }, noType, diag),
    respiceTimed("subject", function () { return respiceGetSubject(item); }, "", diag)
  ]).then(function (v) {
    var userEmail = "";
    try {
      userEmail = Office.context.mailbox.userProfile.emailAddress || "";
    } catch (e) {}
    return { to: v[0], cc: v[1], bcc: v[2], composeType: v[3].type, composeTypeRaw: v[3].raw, subject: v[4], inReplyTo: respiceGetInReplyTo(item), userEmail: userEmail };
  });
}

// Diagnostics: a short trail of send-time runs, readable from the settings pane.
// Uses localStorage (shared by the event runtime and the task pane on web / Mac / new Outlook);
// silently does nothing where it is unavailable (classic Windows JS runtime).
var RESPICE_TRACE_KEY = "respiceTrace";

function respiceTrace(stage, data) {
  try {
    if (typeof localStorage === "undefined") return;
    var t = JSON.parse(localStorage.getItem(RESPICE_TRACE_KEY) || "[]");
    t.push({ t: new Date().toISOString(), stage: stage, data: data || null });
    localStorage.setItem(RESPICE_TRACE_KEY, JSON.stringify(t.slice(-30)));
  } catch (e) {}
}

function respiceDescribeInput(input) {
  return {
    composeType: input.composeTypeRaw,
    replyBySubject: respiceIsReplySubject(input.subject),
    inReplyTo: input.inReplyTo ? "set" : "empty",
    to: (input.to || []).length,
    cc: (input.cc || []).length,
    bcc: (input.bcc || []).length,
    types: [].concat(input.to || [], input.cc || [], input.bcc || []).map(function (r) {
      return r.recipientType;
    })
  };
}

function onMessageSendHandler(event) {
  var t0 = Date.now();
  var diag = ["handler-start"];
  var done = false;
  var settings = respiceNormalizeSettings(null);

  function diagMessage(result) {
    var text = "Respice diagnostics (" + (Date.now() - t0) + "ms):\n" + diag.join("\n") + "\nResult: " + result;
    return text.length > 490 ? text.slice(0, 487) + "..." : text;
  }

  function finish(opts, result) {
    if (done) return;
    done = true;
    clearTimeout(watchdog);
    if (settings.diagnostic) {
      opts = { allowEvent: false, errorMessage: diagMessage(result || (opts.allowEvent ? "allow" : "warn")) };
    }
    respiceTrace("complete", { allowEvent: opts.allowEvent, diag: diag });
    try {
      event.completed(opts);
    } catch (e) {
      // Never leave Outlook waiting: fall back to the simplest possible completion.
      try {
        event.completed({ allowEvent: true });
      } catch (e2) {}
    }
  }

  var watchdog = setTimeout(function () {
    diag.push("WATCHDOG fired");
    finish({ allowEvent: true }, "watchdog: sent without check");
  }, RESPICE_CONFIG.watchdogMs);

  respiceTrace("handler-start");
  try {
    settings = respiceLoadSettings();
    diag.push("settings ok");
    var r = settings.rules;
    if (!r.external && !r.mailingList && !r.tooMany && !r.replyAll) return finish({ allowEvent: true }, "all rules off");

    respiceGatherInput(Office.context.mailbox.item, diag)
      .then(function (input) {
        var d = respiceDescribeInput(input);
        diag.push("composeType " + d.composeType + "; inReplyTo " + d.inReplyTo + "; replyBySubject " + d.replyBySubject + "; to/cc/bcc " + d.to + "/" + d.cc + "/" + d.bcc);
        respiceTrace("gathered", d);
        var findings = respiceEvaluate(input, settings);
        if (!findings.length) return finish({ allowEvent: true }, "no warnings");
        finish(
          {
            allowEvent: false,
            errorMessage: respiceFormatMessage(findings, false),
            errorMessageMarkdown: respiceFormatMessage(findings, true)
          },
          findings.map(function (f) { return f.rule; }).join(", ")
        );
      })
      .catch(function (e) {
        diag.push("ERROR " + (e && e.message));
        respiceTrace("error", String(e && e.message));
        finish({ allowEvent: true }, "error"); // fail open: never block mail because of an add-in error
      });
  } catch (e) {
    diag.push("ERROR " + (e && e.message));
    respiceTrace("error", String(e && e.message));
    finish({ allowEvent: true }, "error");
  }
}

if (
  typeof Office !== "undefined" &&
  Office.actions &&
  typeof Office.actions.associate === "function" &&
  typeof RESPICE_SKIP_ASSOCIATE === "undefined"
) {
  respiceTrace("runtime-loaded");
  // Browser runtime (web, Mac, new Outlook on Windows): the page must call Office.onReady(),
  // otherwise Office.js never finishes its handshake and Outlook never delivers the event.
  // Classic Outlook on Windows ignores this.
  if (typeof Office.onReady === "function") {
    Office.onReady(function () {
      respiceTrace("office-ready");
    });
  }
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    RESPICE_DEFAULTS: RESPICE_DEFAULTS,
    respiceParseLines: respiceParseLines,
    respiceNormalizeSettings: respiceNormalizeSettings,
    respiceIsInternal: respiceIsInternal,
    respiceMatchesList: respiceMatchesList,
    respiceEvaluate: respiceEvaluate,
    respiceFormatMessage: respiceFormatMessage,
    respiceIsReplySubject: respiceIsReplySubject,
    respiceIsForwardSubject: respiceIsForwardSubject,
    RESPICE_CONFIG: RESPICE_CONFIG,
    onMessageSendHandler: onMessageSendHandler
  };
}
