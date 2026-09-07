/*
 * The mobile task pane, driven as a real screen.
 *
 * The pane exists because an ExecuteFunction button gives an add-in no UI of
 * its own on a phone - only a line in Outlook's notification bar, whose
 * padding and icon the host owns. So the thing worth testing is that the pane
 * actually SHOWS the right screen for each outcome, and that it still refuses
 * to sign in when there is nobody to add.
 *
 * jsdom isn't a dependency of this project, so the DOM below is a stub with
 * exactly the surface mobile.js touches. If mobile.js grows past it, these
 * tests fail loudly rather than silently passing.
 *
 * Run: npm test
 */
"use strict";
var fs = require("fs");
var path = require("path");

var SRC = path.join(__dirname, "..", "src");
var MOBILE = fs.readFileSync(path.join(SRC, "mobile", "mobile.js"), "utf8");
var AddFlow = require(path.join(SRC, "addflow.js"));

var failures = 0, passes = 0;
function check(label, actual, expected) {
  if (actual === expected) { passes++; return; }
  failures++;
  console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) +
    "\n  actual:   " + JSON.stringify(actual));
}
function has(label, hay, needle) {
  if (String(hay).indexOf(needle) >= 0) { passes++; return; }
  failures++;
  console.error("FAIL  " + label + "\n  expected to contain: " + JSON.stringify(needle) +
    "\n  got: " + JSON.stringify(String(hay)));
}

// --- the smallest DOM that mobile.js can run against ----------------------

function El(id) {
  this.id = id;
  this.hidden = false;
  this.textContent = "";
  this.children = [];
  this.className = "";
  this.listeners = {};
  var self = this;
  this.classList = {
    toggle: function (name, on) {
      var set = self.className.split(" ").filter(Boolean);
      var i = set.indexOf(name);
      if (on && i < 0) { set.push(name); }
      if (!on && i >= 0) { set.splice(i, 1); }
      self.className = set.join(" ");
    },
  };
}
El.prototype.appendChild = function (c) { this.children.push(c); return c; };
El.prototype.addEventListener = function (ev, fn) { this.listeners[ev] = fn; };
El.prototype.click = function () { if (this.listeners.click) { this.listeners.click(); } };

function makeDocument() {
  var ids = ["working", "done", "message", "workingText", "doneTitle",
             "doneList", "doneNote", "messageTitle", "messageText", "closeBtn"];
  var map = {};
  ids.forEach(function (id) { map[id] = new El(id); });
  return {
    _map: map,
    getElementById: function (id) {
      if (!map[id]) { throw new Error("mobile.js asked for an element the stub lacks: #" + id); }
      return map[id];
    },
    createElement: function (tag) { return new El(tag); },
  };
}

function addr(email, name) { return { emailAddress: email, displayName: name || email }; }

/** Load and run the real pane against one scenario. */
async function runPane(opts) {
  var log = { tokens: 0, closed: false };
  var document = makeDocument();
  var ready = null;

  var Office = {
    onReady: function (cb) { ready = cb; },
    MailboxEnums: { RestVersion: { v2_0: "v2.0" } },
    context: {
      ui: { closeContainer: function () { log.closed = true; } },
      mailbox: {
        userProfile: { emailAddress: opts.me },
        convertToRestId: function (id) { return id; },
        item: { itemId: "AAA", from: opts.from, to: opts.to || [] },
      },
    },
  };

  var GraphData = {
    getToken: function () {
      log.tokens++;
      if (opts.tokenError) { return Promise.reject(new Error(opts.tokenError)); }
      return Promise.resolve("token");
    },
    getMessage: function () { return Promise.resolve(opts.graphMessage); },
    loadContacts: function () { return Promise.resolve({ byEmail: opts.existing || {} }); },
    toContactPayload: function (p) { return p; },
    createContact: function () { return Promise.resolve({}); },
    enrichContact: function () { return Promise.resolve({ updated: true }); },
    latestMessageFrom: function () { return Promise.resolve(null); },
  };

  var factory = new Function("document", "window", "Office", "GraphData",
                             "SigParser", "AddFlow", MOBILE);
  factory(document, {}, Office, GraphData, { parse: function () { return null; } }, AddFlow);

  if (!ready) { throw new Error("mobile.js never registered Office.onReady"); }
  ready();
  // Let the async work settle.
  for (var i = 0; i < 50; i++) { await Promise.resolve(); }
  await new Promise(function (r) { setTimeout(r, 5); });

  return { log: log, d: document._map };
}

var ME = "matt@iowadot.us";
var JANE = { emailAddress: { address: "jane@contoso.com", name: "Jane Smith" } };

(async function () {
  // --- the reviewer's message: a note to self ----------------------------
  var note = await runPane({
    me: ME, from: addr(ME, "Matthew Miller"), to: [addr(ME)],
    graphMessage: null,
  });
  check("note to self: no sign-in", note.log.tokens, 0);
  check("note to self: shows the message screen", note.d.message.hidden, false);
  check("note to self: hides the spinner", note.d.working.hidden, true);
  check("note to self: hides the results", note.d.done.hidden, true);
  check("note to self: titled plainly", note.d.messageTitle.textContent, "No one to add");
  check("note to self: not styled as an error", note.d.message.className.indexOf("error") >= 0, false);

  // --- a real sender ------------------------------------------------------
  var ok = await runPane({
    me: ME, from: addr("jane@contoso.com", "Jane Smith"), to: [addr(ME)],
    graphMessage: { from: JANE, toRecipients: [], body: { content: "<p>hi</p>" } },
  });
  check("a real sender: shows the done screen", ok.d.done.hidden, false);
  check("a real sender: spinner is gone", ok.d.working.hidden, true);
  check("a real sender: headline names them", ok.d.doneTitle.textContent, "Added Jane Smith.");
  check("a real sender: one row in the list", ok.d.doneList.children.length, 1);
  has("a real sender: the row says what happened",
    ok.d.doneList.children[0].children.map(function (c) { return c.textContent; }).join(" "),
    "Added to your contacts");

  // --- an existing contact ------------------------------------------------
  var known = await runPane({
    me: ME, from: addr("jane@contoso.com", "Jane Smith"), to: [addr(ME)],
    graphMessage: { from: JANE, toRecipients: [], body: { content: "" } },
    existing: { "jane@contoso.com": { id: "c1" } },
  });
  check("a known contact: reported as updated", known.d.doneTitle.textContent, "Updated Jane Smith.");

  // --- sign-in fails, which is the mobile case that started all this ------
  //
  // The pane must show getToken's own message, because that text is already
  // written for the device in hand ("open it on your computer once").
  var failed = await runPane({
    me: ME, from: addr("jane@contoso.com", "Jane Smith"), to: [addr(ME)],
    tokenError: "Couldn't sign in on this device. Open Auto Contacts once in Outlook on your computer",
    graphMessage: null,
  });
  check("failed sign-in: shows the message screen", failed.d.message.hidden, false);
  check("failed sign-in: titled as a failure", failed.d.messageTitle.textContent, "Couldn't finish");
  has("failed sign-in: passes the advice through",
    failed.d.messageText.textContent, "on your computer");
  check("failed sign-in: styled as an error",
    failed.d.message.className.indexOf("error") >= 0, true);

  // --- closing returns you to the message --------------------------------
  ok.d.closeBtn.click();
  check("Close closes the pane", ok.log.closed, true);

  if (failures) {
    console.error("\n" + failures + " mobile pane test(s) failed.");
    process.exit(1);
  }
  console.log("All " + passes + " mobile pane checks passed.");
})();
