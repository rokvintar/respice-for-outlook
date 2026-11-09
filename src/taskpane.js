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

  function fill(s) {
    $("rule-external").checked = s.rules.external;
    $("rule-list").checked = s.rules.mailingList;
    $("rule-toomany").checked = s.rules.tooMany;
    $("rule-replyall").checked = s.rules.replyAll;
    $("internal-domains").value = s.internalDomains.join("\n");
    $("list-patterns").value = s.listPatterns.join("\n");
    $("treat-dl").checked = s.treatOutlookDistributionListsAsLists;
    $("max-recipients").value = s.maxRecipients;
    var d = userDomain();
    $("internal-domains").placeholder = d || "example.com";
    $("internal-hint").textContent = d ? "Leave empty to use your own domain: " + d : "";
  }

  function read() {
    return respiceNormalizeSettings({
      rules: {
        external: $("rule-external").checked,
        mailingList: $("rule-list").checked,
        tooMany: $("rule-toomany").checked,
        replyAll: $("rule-replyall").checked
      },
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
    out.textContent = "Checking…";
    respiceGatherInput(Office.context.mailbox.item)
      .then(function (input) {
        var findings = respiceEvaluate(input, respiceLoadSettings());
        out.textContent = findings.length ? respiceFormatMessage(findings, false) : "No warnings. This message would send without a prompt.";
      })
      .catch(function (e) {
        out.textContent = "Error: " + (e && e.message);
      });
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
    $("version").textContent = "Respice for Outlook 1.0.0";
  });
})();
