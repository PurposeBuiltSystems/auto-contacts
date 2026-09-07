/*
 * The shipped code, driven end to end.
 *
 * test/summary.test.js checks the pure helpers. This runs the REAL
 * addToContacts and the REAL getToken - the actual files that are deployed -
 * with only Office.js, MSAL and fetch stubbed. It exists to answer two
 * questions that unit tests on helpers cannot:
 *
 *   1. Does a note to self really avoid signing in? That is the fix for what
 *      the Store reviewer photographed, and it is only true if the ORDER of
 *      operations in addToContacts is right. A green pickTargets() test would
 *      not have caught the token call happening first.
 *   2. Does the sign-in path really treat a phone differently - shorter wait,
 *      different advice? The timeout is a number buried in a Promise.race, so
 *      the test captures what was actually asked of setTimeout.
 *
 * What this CANNOT test: whether NAA itself works on a real device. That
 * needs a phone and an account that has never consented.
 *
 * Run: npm test
 */
"use strict";
var fs = require("fs");
var path = require("path");

var SRC = path.join(__dirname, "..", "src");
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

// --------------------------------------------------------------------------
// Part 1 — addToContacts, the real function
// --------------------------------------------------------------------------

var COMMANDS = fs.readFileSync(path.join(SRC, "commands", "commands.js"), "utf8");

function addr(email, name) { return { emailAddress: email, displayName: name || email }; }

/**
 * Build a world, run the real addToContacts in it, and report what happened.
 */
async function runCommand(opts) {
  var log = { tokens: 0, created: [], enriched: [], notes: [], completed: false };

  var Office = {
    onReady: function () {},
    actions: null,
    MailboxEnums: {
      RestVersion: { v2_0: "v2.0" },
      ItemNotificationMessageType: {
        InformationalMessage: "info", ErrorMessage: "error",
      },
    },
    context: {
      mailbox: {
        userProfile: { emailAddress: opts.me },
        convertToRestId: function (id) { return id; },
        item: {
          itemId: "AAA",
          from: opts.from,
          to: opts.to || [],
          notificationMessages: {
            replaceAsync: function (key, msg) { log.notes.push(msg.message); },
          },
        },
      },
    },
  };

  var GraphData = {
    getToken: function () { log.tokens++; return Promise.resolve("token"); },
    getMessage: function () { return Promise.resolve(opts.graphMessage); },
    loadContacts: function () { return Promise.resolve({ byEmail: opts.existing || {} }); },
    toContactPayload: function (person) { return person; },
    createContact: function (t, p) { log.created.push(p.email); return Promise.resolve({}); },
    enrichContact: function (t, e, p) { log.enriched.push(p.email); return Promise.resolve({ updated: true }); },
    latestMessageFrom: function () { return Promise.resolve(null); },
  };

  var SigParser = { parse: function () { return null; } };

  var AddFlow = require(path.join(SRC, "addflow.js"));
  var factory = new Function("Office", "GraphData", "SigParser", "AddFlow", "module",
    COMMANDS + "\nreturn addToContacts;");
  var addToContacts = factory(Office, GraphData, SigParser, AddFlow, undefined);

  await addToContacts({ completed: function () { log.completed = true; } });
  return log;
}

var ME = "matt@iowadot.us";

(async function () {
  // --- the reviewer's exact test: a note to self -------------------------
  //
  // "Note to self / To You". The only address on the message is the user's
  // own. Nothing to add - and, critically, nothing to sign in for.
  var self1 = await runCommand({
    me: ME,
    from: addr(ME, "Matthew Miller"),
    to: [addr(ME, "Matthew Miller")],
    graphMessage: null,   // never reached; a throw here would prove otherwise
  });
  check("note to self: never signs in", self1.tokens, 0);
  check("note to self: says so", self1.notes[0], "No one to add from this message.");
  check("note to self: completes the event", self1.completed, true);

  // --- a real message from a real person ---------------------------------
  var jane = { emailAddress: { address: "jane@contoso.com", name: "Jane Smith" } };
  var normal = await runCommand({
    me: ME,
    from: addr("jane@contoso.com", "Jane Smith"),
    to: [addr(ME)],
    graphMessage: { from: jane, toRecipients: [], body: { content: "<p>hi</p>" } },
  });
  check("a real sender: signs in once", normal.tokens, 1);
  check("a real sender: is added", normal.created[0], "jane@contoso.com");
  check("a real sender: is reported by name", normal.notes[0], "Added Jane Smith.");
  check("a real sender: completes the event", normal.completed, true);

  // --- a message you sent to other people --------------------------------
  var sent = await runCommand({
    me: ME,
    from: addr(ME),
    to: [addr("a@x.com", "A"), addr("b@x.com", "B")],
    graphMessage: {
      from: { emailAddress: { address: ME, name: "Me" } },
      toRecipients: [
        { emailAddress: { address: "a@x.com", name: "A" } },
        { emailAddress: { address: "b@x.com", name: "B" } },
      ],
      body: { content: "" },
    },
  });
  check("a sent message: signs in", sent.tokens, 1);
  check("a sent message: adds both recipients", sent.created.length, 2);
  check("a sent message: counts them", sent.notes[0], "Added 2 contacts.");

  // --- an already-known contact ------------------------------------------
  var known = await runCommand({
    me: ME,
    from: addr("jane@contoso.com", "Jane Smith"),
    to: [addr(ME)],
    graphMessage: { from: jane, toRecipients: [], body: { content: "" } },
    existing: { "jane@contoso.com": { id: "c1" } },
  });
  check("a known contact is enriched, not duplicated", known.created.length, 0);
  check("and reported as updated", known.notes[0], "Updated Jane Smith.");

  // ------------------------------------------------------------------------
  // Part 2 — getToken, the real function, on each platform
  // ------------------------------------------------------------------------

  var GRAPHJS = fs.readFileSync(path.join(SRC, "graph.js"), "utf8");

  /** Load graph.js against a given platform, capturing every timeout asked for. */
  function loadGraph(platformValue, hostName) {
    var delays = [];
    var self = {};
    var Office = {
      PlatformType: { PC: "PC", Mac: "Mac", iOS: "iOS", Android: "Android", OfficeOnline: "OfficeOnline" },
      context: {
        platform: platformValue,
        roamingSettings: { get: function () { return false; }, set: function () {}, saveAsync: function (cb) { cb(); } },
        mailbox: { diagnostics: { hostName: hostName || "Outlook" } },
      },
    };
    // Silent acquisition fails (a first-time user), and no popup ever
    // arrives - which is precisely the mobile failure being guarded against.
    var msal = {
      createNestablePublicClientApplication: function () {
        return Promise.resolve({
          acquireTokenSilent: function () { return Promise.reject(new Error("interaction_required")); },
          acquireTokenPopup: function () { return new Promise(function () {}); },
          getAllAccounts: function () { return []; },
        });
      },
    };
    var fakeSetTimeout = function (fn, ms) { delays.push(ms); return { ms: ms }; };
    var factory = new Function("self", "Office", "msal", "fetch", "setTimeout", "clearTimeout",
      GRAPHJS + "\nreturn self.GraphData;");
    var api = factory(self, Office, msal, function () {}, fakeSetTimeout, function () {});
    return { api: api, delays: delays };
  }

  var phone = loadGraph("iOS", "OutlookIOS");
  var desk = loadGraph("Mac", "Outlook");

  check("a phone is detected as mobile", phone.api._platform(), "mobile");
  check("a Mac is not", desk.api._platform(), "desktop");

  has("mobile advice points at the computer", phone.api._signInHelp("finish"), "on your computer");
  has("desktop advice keeps the popup hint", desk.api._signInHelp("finish"), "Mission Control");
  has("mobile start advice drops Cmd+Q", phone.api._signInHelp("start"), "Close Outlook completely");
  has("desktop start advice keeps it", desk.api._signInHelp("start"), "Cmd+Q");

  // The point of the change: a phone must not wait two minutes for a popup
  // that is never coming, because the host's progress bar stays up the whole
  // time and the add-in looks hung.
  phone.api.getToken().catch(function () {});
  desk.api.getToken().catch(function () {});
  await new Promise(function (r) { setTimeout(r, 20); });

  check("mobile waits 45s for the popup", phone.delays.indexOf(45000) >= 0, true);
  check("mobile never waits 120s", phone.delays.indexOf(120000) >= 0, false);
  check("desktop still waits 120s", desk.delays.indexOf(120000) >= 0, true);

  if (failures) {
    console.error("\n" + failures + " flow test(s) failed.");
    process.exit(1);
  }
  console.log("All " + passes + " end-to-end flow checks passed.");
})();
