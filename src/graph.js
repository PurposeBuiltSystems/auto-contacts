/*
 * Auto Contacts — Microsoft Graph data layer.
 *
 * AUTH: Nested App Authentication (NAA) via MSAL — no backend; identical
 * pattern to the Reply-All and Biweekly add-ins. Only ever touches the
 * SIGNED-IN user's own mailbox and contacts (delegated Contacts.ReadWrite +
 * Mail.Read).
 *
 * NOTE: this add-in uses its OWN Entra app registration — do not reuse the
 * Reply-All app (87764ff9-…): adding Contacts scopes there would force
 * re-consent on every existing Reply-All/Biweekly user. See SETUP.md.
 *
 * Exposes a global `GraphData` object.
 */
/* global msal */
(function (root) {
  "use strict";

  var CLIENT_ID = "1a218911-cb74-4e94-8774-444baf11a9a8"; // "Auto Contacts" Entra app (purposebuilt.systems tenant)
  var GRAPH = "https://graph.microsoft.com/v1.0";
  var SCOPES = ["Contacts.ReadWrite", "Mail.Read"];

  var pcaPromise = null;

  function getPca() {
    if (!pcaPromise) {
      pcaPromise = msal.createNestablePublicClientApplication({
        auth: {
          clientId: CLIENT_ID,
          authority: "https://login.microsoftonline.com/common",
        },
      });
    }
    return pcaPromise;
  }

  async function getToken() {
    var pca = await getPca();
    try {
      var silent = await pca.acquireTokenSilent({ scopes: SCOPES });
      return silent.accessToken;
    } catch (e) {
      var interactive = await pca.acquireTokenPopup({ scopes: SCOPES });
      return interactive.accessToken;
    }
  }

  async function graph(token, method, path, body) {
    var res = await fetch(GRAPH + path, {
      method: method,
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/json",
      },
      body: body ? JSON.stringify(body) : undefined,
    });
    if (!res.ok) {
      var text = await res.text();
      throw new Error("Graph " + method + " " + path + " -> " + res.status + " " + text);
    }
    return res.status === 204 ? null : res.json();
  }

  /** Page through a Graph collection following @odata.nextLink. */
  async function graphAll(token, path) {
    var items = [];
    var url = GRAPH + path;
    var guard = 0;
    while (url && guard++ < 30) {
      var res = await fetch(url, { headers: { Authorization: "Bearer " + token } });
      if (!res.ok) { throw new Error("Graph GET " + url + " -> " + res.status); }
      var page = await res.json();
      items = items.concat(page.value || []);
      url = page["@odata.nextLink"] || null;
    }
    return items;
  }

  // ---------- contacts ----------

  /** All contacts, with a lowercase email → contact index for fast matching. */
  async function loadContacts(token) {
    var contacts = await graphAll(
      token,
      "/me/contacts?$select=id,displayName,emailAddresses,businessPhones,mobilePhone,jobTitle,companyName&$top=500"
    );
    var byEmail = {};
    contacts.forEach(function (c) {
      (c.emailAddresses || []).forEach(function (e) {
        if (e && e.address) { byEmail[e.address.toLowerCase()] = c; }
      });
    });
    return { list: contacts, byEmail: byEmail };
  }

  /** Build a Graph contact payload from a person + parsed signature fields. */
  function toContactPayload(person, sig) {
    var names = splitName(person.name || person.email);
    var payload = {
      givenName: names.given,
      surname: names.surname,
      displayName: person.name || person.email,
      emailAddresses: [{ address: person.email, name: person.name || person.email }],
    };
    if (sig) {
      if (sig.phones && sig.phones.business) { payload.businessPhones = [sig.phones.business]; }
      if (sig.phones && sig.phones.mobile) { payload.mobilePhone = sig.phones.mobile; }
      if (sig.title) { payload.jobTitle = sig.title; }
      if (sig.company) { payload.companyName = sig.company; }
      if (sig.website) { payload.businessHomePage = sig.website; }
      var notes = [];
      if (sig.linkedin) { notes.push("LinkedIn: " + sig.linkedin); }
      if (sig.address) { notes.push("Address (from signature): " + sig.address); }
      if (notes.length) { payload.personalNotes = notes.join("\n"); }
    }
    return payload;
  }

  function splitName(name) {
    var clean = String(name || "").replace(/["<>]/g, "").trim();
    if (clean.indexOf("@") !== -1) { return { given: clean, surname: "" }; }
    // "Last, First" → "First Last"
    if (clean.indexOf(",") !== -1) {
      var lf = clean.split(",");
      clean = (lf[1] || "").trim() + " " + lf[0].trim();
    }
    var parts = clean.split(/\s+/);
    if (parts.length === 1) { return { given: parts[0], surname: "" }; }
    return { given: parts.slice(0, -1).join(" "), surname: parts[parts.length - 1] };
  }

  async function createContact(token, payload) {
    return graph(token, "POST", "/me/contacts", payload);
  }

  /** Fill BLANK fields on an existing contact; never overwrite user data. */
  async function enrichContact(token, existing, payload) {
    var patch = {};
    if (!(existing.businessPhones || []).length && payload.businessPhones) { patch.businessPhones = payload.businessPhones; }
    if (!existing.mobilePhone && payload.mobilePhone) { patch.mobilePhone = payload.mobilePhone; }
    if (!existing.jobTitle && payload.jobTitle) { patch.jobTitle = payload.jobTitle; }
    if (!existing.companyName && payload.companyName) { patch.companyName = payload.companyName; }
    if (Object.keys(patch).length === 0) { return { updated: false }; }
    await graph(token, "PATCH", "/me/contacts/" + existing.id, patch);
    return { updated: true, fields: Object.keys(patch) };
  }

  // ---------- mail ----------

  /** A message by Graph REST id, with body + parties. */
  async function getMessage(token, restId) {
    return graph(
      token, "GET",
      "/me/messages/" + restId + "?$select=from,toRecipients,ccRecipients,body,sentDateTime"
    );
  }

  /** The most recent inbox message FROM the given address (for signature mining). */
  async function latestMessageFrom(token, email) {
    var safe = email.replace(/'/g, "''");
    var res = await graph(
      token, "GET",
      "/me/mailFolders/inbox/messages?$search=%22from:" + encodeURIComponent(safe) +
        "%22&$select=from,body,receivedDateTime&$top=5"
    );
    var hits = (res.value || []).filter(function (m) {
      return m.from && m.from.emailAddress &&
        m.from.emailAddress.address.toLowerCase() === email.toLowerCase();
    });
    hits.sort(function (a, b) { return (b.receivedDateTime || "").localeCompare(a.receivedDateTime || ""); });
    return hits[0] || null;
  }

  /** Unique people you SENT mail to in the window (name+email, deduped). */
  async function sentRecipients(token, daysBack) {
    var since = new Date(Date.now() - daysBack * 864e5).toISOString();
    var msgs = await graphAll(
      token,
      "/me/mailFolders/sentitems/messages?$select=toRecipients,ccRecipients,sentDateTime" +
        "&$filter=sentDateTime ge " + since + "&$top=100"
    );
    var seen = {};
    var people = [];
    msgs.forEach(function (m) {
      (m.toRecipients || []).concat(m.ccRecipients || []).forEach(function (r) {
        var e = r && r.emailAddress;
        if (!e || !e.address) { return; }
        var key = e.address.toLowerCase();
        if (seen[key]) { return; }
        seen[key] = true;
        people.push({ email: e.address, name: e.name || e.address });
      });
    });
    return people;
  }

  root.GraphData = {
    getToken: getToken,
    loadContacts: loadContacts,
    toContactPayload: toContactPayload,
    createContact: createContact,
    enrichContact: enrichContact,
    getMessage: getMessage,
    latestMessageFrom: latestMessageFrom,
    sentRecipients: sentRecipients,
    _config: { clientId: CLIENT_ID },
  };
})(typeof self !== "undefined" ? self : this);
