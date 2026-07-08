/*
 * Auto Contacts — sweep task pane.
 *
 * Scans sent mail for people you've emailed who aren't in Contacts, mines
 * each person's most recent inbox message for signature details, and shows a
 * checklist of proposed cards. Nothing is written until "Add selected".
 */
/* global Office, GraphData, SigParser, document */
(function () {
  "use strict";

  // Senders that shouldn't become contacts.
  var AUTOMATED_RE = /no-?reply|do-?not-?reply|donotreply|notifications?@|mailer|postmaster|newsletter|marketing@|updates@|alerts?@|automated|@e(mail)?\.|unsubscribe/i;

  var candidates = []; // { person, sig, checked, state }

  Office.onReady(function () {
    byId("scan").addEventListener("click", scan);
    byId("selectAll").addEventListener("click", selectAll);
    byId("addSelected").addEventListener("click", addSelected);
  });

  function byId(id) { return document.getElementById(id); }

  function setStatus(kind, text) {
    var el = byId("status");
    if (!text) { el.hidden = true; return; }
    el.hidden = false;
    el.className = "status " + kind;
    el.textContent = text;
  }

  async function scan() {
    var daysBack = Math.max(1, Math.min(90, parseInt(byId("daysBack").value, 10) || 30));
    var skipAutomated = byId("skipAutomated").checked;
    byId("scan").disabled = true;
    byId("results").hidden = true;
    candidates = [];
    try {
      setStatus("work", "Reading your sent mail…");
      var token = await GraphData.getToken();
      // Own address comes from Office.js — no Graph /me call (no User.Read scope).
      var myAddr = ((Office.context.mailbox.userProfile || {}).emailAddress || "").toLowerCase();

      var people = await GraphData.sentRecipients(token, daysBack);
      setStatus("work", "Checking " + people.length + " people against your contacts…");
      var contacts = await GraphData.loadContacts(token);

      var fresh = people.filter(function (p) {
        var addr = p.email.toLowerCase();
        if (addr === myAddr) { return false; }
        if (contacts.byEmail[addr]) { return false; }
        if (skipAutomated && AUTOMATED_RE.test(addr)) { return false; }
        return true;
      });

      if (!fresh.length) {
        setStatus("info", "Everyone you've emailed in the last " + daysBack + " days is already in your contacts.");
        return;
      }

      // Mine signatures one by one so progress is visible.
      for (var i = 0; i < fresh.length; i++) {
        setStatus("work", "Reading signatures… " + (i + 1) + " of " + fresh.length);
        var sig = null;
        try {
          var recent = await GraphData.latestMessageFrom(token, fresh[i].email);
          if (recent) {
            sig = SigParser.parse(recent.body && recent.body.content, fresh[i].name, fresh[i].email);
          }
        } catch (e) { /* no inbox mail from them — card will be name+email only */ }
        candidates.push({ person: fresh[i], sig: sig, checked: true, state: "" });
      }

      render();
      byId("results").hidden = false;
      setStatus("info", fresh.length + " people found who aren't in your contacts. Review and add.");
    } catch (e) {
      var msg = (e && e.message) || String(e);
      if (/REPLACE_WITH_ENTRA_CLIENT_ID/.test(GraphData._config.clientId)) {
        msg = "Set your Entra client ID in src/graph.js before running. (" + msg + ")";
      }
      setStatus("error", "Scan failed: " + msg);
    } finally {
      byId("scan").disabled = false;
    }
  }

  function render() {
    var host = byId("people");
    host.innerHTML = "";
    candidates.forEach(function (c, idx) {
      var div = document.createElement("div");
      div.className = "person" + (c.state === "added" ? " done" : "");

      var cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = c.checked && c.state !== "added";
      cb.disabled = c.state === "added";
      cb.addEventListener("change", function () { candidates[idx].checked = cb.checked; });

      var who = document.createElement("div");
      who.className = "who";
      var name = document.createElement("div");
      name.className = "name";
      name.textContent = c.person.name || c.person.email;
      var email = document.createElement("div");
      email.className = "email";
      email.textContent = c.person.email;
      var fields = document.createElement("div");
      fields.className = "fields";
      var found = [];
      if (c.sig) {
        if (c.sig.title) { found.push(c.sig.title); }
        if (c.sig.company) { found.push(c.sig.company); }
        if (c.sig.phones.business) { found.push("☎ " + c.sig.phones.business); }
        if (c.sig.phones.mobile) { found.push("📱 " + c.sig.phones.mobile); }
        if (c.sig.website) { found.push("🌐 site"); }
        if (c.sig.linkedin) { found.push("in/ LinkedIn"); }
      }
      if (found.length) {
        found.forEach(function (f) {
          var s = document.createElement("span");
          s.textContent = f;
          fields.appendChild(s);
        });
      } else {
        fields.innerHTML = "<span class='nosig'>no signature found — name + email only</span>";
      }
      who.appendChild(name);
      who.appendChild(email);
      who.appendChild(fields);

      var state = document.createElement("div");
      state.className = "state";
      state.textContent = c.state === "added" ? "✓ added" : "";

      div.appendChild(cb);
      div.appendChild(who);
      div.appendChild(state);
      host.appendChild(div);
    });
  }

  function selectAll() {
    var anyOff = candidates.some(function (c) { return !c.checked && c.state !== "added"; });
    candidates.forEach(function (c) { if (c.state !== "added") { c.checked = anyOff; } });
    render();
  }

  async function addSelected() {
    var picked = candidates.filter(function (c) { return c.checked && c.state !== "added"; });
    if (!picked.length) { setStatus("info", "Nothing selected."); return; }
    byId("addSelected").disabled = true;
    try {
      var token = await GraphData.getToken();
      for (var i = 0; i < picked.length; i++) {
        setStatus("work", "Adding contacts… " + (i + 1) + " of " + picked.length);
        var payload = GraphData.toContactPayload(picked[i].person, picked[i].sig);
        await GraphData.createContact(token, payload);
        picked[i].state = "added";
      }
      render();
      setStatus("info", picked.length + " contact(s) added.");
    } catch (e) {
      render();
      setStatus("error", "Add failed: " + ((e && e.message) || e));
    } finally {
      byId("addSelected").disabled = false;
    }
  }
})();
