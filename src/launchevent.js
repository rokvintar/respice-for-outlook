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
    maxRecipients: isFinite(max) && max >= 1 ? max : d.maxRecipients
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

  if (s.rules.replyAll && input.composeType === "reply" && recipients.length > 1) {
    findings.push({
      rule: "replyAll",
      label: RESPICE_LABELS.replyAll,
      items: [],
      detail: "this reply goes to " + recipients.length + " recipients. Did you mean Reply all?"
    });
  }

  return findings;
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

function respiceGetAsync(field) {
  return new Promise(function (resolve) {
    if (!field || typeof field.getAsync !== "function") return resolve([]);
    field.getAsync(function (res) {
      resolve(res.status === Office.AsyncResultStatus.Succeeded && res.value ? res.value : []);
    });
  });
}

function respiceGetComposeType(item) {
  return new Promise(function (resolve) {
    if (!item || typeof item.getComposeTypeAsync !== "function") return resolve(null);
    item.getComposeTypeAsync(function (res) {
      resolve(res.status === Office.AsyncResultStatus.Succeeded && res.value ? res.value.composeType : null);
    });
  });
}

function respiceGatherInput(item) {
  return Promise.all([
    respiceGetAsync(item.to),
    respiceGetAsync(item.cc),
    respiceGetAsync(item.bcc),
    respiceGetComposeType(item)
  ]).then(function (v) {
    var userEmail = "";
    try {
      userEmail = Office.context.mailbox.userProfile.emailAddress || "";
    } catch (e) {}
    return { to: v[0], cc: v[1], bcc: v[2], composeType: v[3], userEmail: userEmail };
  });
}

function onMessageSendHandler(event) {
  var done = false;
  function finish(opts) {
    if (done) return;
    done = true;
    event.completed(opts);
  }
  try {
    var settings = respiceLoadSettings();
    var r = settings.rules;
    if (!r.external && !r.mailingList && !r.tooMany && !r.replyAll) return finish({ allowEvent: true });

    respiceGatherInput(Office.context.mailbox.item)
      .then(function (input) {
        var findings = respiceEvaluate(input, settings);
        if (!findings.length) return finish({ allowEvent: true });
        finish({
          allowEvent: false,
          errorMessage: respiceFormatMessage(findings, false),
          errorMessageMarkdown: respiceFormatMessage(findings, true)
        });
      })
      .catch(function () {
        finish({ allowEvent: true }); // fail open: never block mail because of an add-in error
      });
  } catch (e) {
    finish({ allowEvent: true });
  }
}

if (
  typeof Office !== "undefined" &&
  Office.actions &&
  typeof Office.actions.associate === "function" &&
  typeof RESPICE_SKIP_ASSOCIATE === "undefined"
) {
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
    onMessageSendHandler: onMessageSendHandler
  };
}
