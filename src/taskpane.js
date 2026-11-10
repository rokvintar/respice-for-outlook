/* Respice for Outlook: settings task pane. Depends on launchevent.js (loaded first). */

(function () {
  function $(id) {
    return document.getElementById(id);
  }

  function setStatus(text, isError) {
    var el = $("status");
    el.textContent = text;
    el.className = isError ? "error" : "ok";
  }

  function userDomain() {
    try {
      return respiceDomainOf(Office.context.mailbox.userProfile.emailAddress || "");
    } catch (e) {
      return "";
    }
  }

  // UI id suffix -> settings key
  var RULE_IDS = { external: "external", list: "mailingList", toomany: "tooMany", replyall: "replyAll" };

  function modeOf(s, key) {
    if (!s.rules[key]) return "off";
    if (s.atSend[key] && s.inBar[key]) return "send";
    if (s.atSend[key]) return "sendonly";
    return "bar";
  }

  function fill(s) {
    Object.keys(RULE_IDS).forEach(function (id) {
      $("mode-" + id).value = modeOf(s, RULE_IDS[id]);
    });
    $("internal-domains").value = s.internalDomains.join("\n");
    $("list-patterns").value = s.listPatterns.join("\n");
    $("treat-dl").checked = s.treatOutlookDistributionListsAsLists;
    $("max-recipients").value = s.maxRecipients;
    var d = userDomain();
    $("internal-domains").placeholder = d || "example.com";
    $("internal-hint").textContent = d ? "Leave empty to use your own domain: " + d : "";
  }

  function read() {
    // mode -> [rule on, check at Send, show in bar]; "off" keeps the defaults for when it's re-enabled
    var MODES = { send: [true, true, true], bar: [true, false, true], sendonly: [true, true, false], off: [false, true, true] };
    var rules = {};
    var atSend = {};
    var inBar = {};
    Object.keys(RULE_IDS).forEach(function (id) {
      var m = MODES[$("mode-" + id).value] || MODES.send;
      rules[RULE_IDS[id]] = m[0];
      atSend[RULE_IDS[id]] = m[1];
      inBar[RULE_IDS[id]] = m[2];
    });
    return respiceNormalizeSettings({
      rules: rules,
      atSend: atSend,
      inBar: inBar,
      internalDomains: $("internal-domains").value,
      listPatterns: $("list-patterns").value,
      treatOutlookDistributionListsAsLists: $("treat-dl").checked,
      maxRecipients: $("max-recipients").value
    });
  }

  function save(settings) {
    var rs = Office.context.roamingSettings;
    rs.set(RESPICE_SETTINGS_KEY, settings);
    rs.saveAsync(function (res) {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        fill(settings);
        setStatus("Saved.", false);
      } else {
        setStatus("Save failed: " + (res.error && res.error.message), true);
      }
    });
  }

  function runTest() {
    var out = $("test-result");
    out.hidden = false;
    out.textContent = "Checking…";
    $("test-details-box").hidden = true;
    respiceGatherInput(Office.context.mailbox.item)
      .then(function (input) {
        var settings = respiceLoadSettings();
        var findings = respiceEvaluate(input, settings);
        var result = findings.length ? respiceFormatMessage(findings, false) : "No warnings. This message would send without a prompt.";
        var details = respiceDescribeInput(input);
        details.rules = settings.rules;
        details.internalDomains = settings.internalDomains.length ? settings.internalDomains : "(own domain) " + userDomain();
        out.textContent = result;
        $("test-details").textContent = JSON.stringify(details, null, 1);
        $("test-details-box").hidden = false;
      })
      .catch(function (e) {
        out.textContent = "Error: " + (e && e.message);
      });
  }

  function showLog() {
    var out = $("log");
    try {
      var t = JSON.parse(localStorage.getItem(RESPICE_TRACE_KEY) || "[]");
      out.textContent = t.length
        ? t
            .map(function (e) {
              return e.t.slice(11, 19) + "  " + e.stage + (e.data ? "  " + JSON.stringify(e.data) : "");
            })
            .join("\n")
        : "(empty: no send-time run recorded in this Outlook yet)";
    } catch (e) {
      out.textContent = "(log unavailable: " + e.message + ")";
    }
  }

  Office.onReady(function () {
    fill(respiceLoadSettings());
    $("save").onclick = function () {
      save(read());
    };
    $("reset").onclick = function () {
      save(respiceNormalizeSettings(null));
    };
    var item = Office.context.mailbox && Office.context.mailbox.item;
    if (item && item.to && typeof item.to.getAsync === "function") {
      $("test-section").hidden = false;
      $("test").onclick = runTest;
    }
    $("log-refresh").onclick = showLog;
    $("log-section").addEventListener("toggle", function () {
      if ($("log-section").open) showLog();
    });
    $("log-clear").onclick = function () {
      try {
        localStorage.removeItem(RESPICE_TRACE_KEY);
      } catch (e) {}
      showLog();
    };
    showLog();
    $("version").textContent = "Respice for Outlook 1.3.1";
  });
})();
