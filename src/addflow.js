/*
 * Auto Contacts — the "add these people" flow, with no UI attached.
 *
 * Two surfaces run this: the desktop ribbon command, which reports through
 * Outlook's notification bar, and the mobile task pane, which draws its own
 * screen. Both must behave identically — who gets added, when a sign-in is
 * needed, and what the result is called — so the decisions live here and the
 * surfaces only decide how to say it.
 *
 * Exposes a global `AddFlow`.
 */
(function (root) {
  "use strict";

  /** Office.js EmailAddressDetails -> the {address, name} shape Graph uses. */
  function officeAddr(a) {
    if (!a) { return null; }
    return { address: a.emailAddress || "", name: a.displayName || "" };
  }

  /**
   * Who to add from a message: the sender, or - on a message you sent - the
   * people you sent it to. Your own address never counts, and no address twice.
   */
  function pickTargets(sender, recipients, myAddr) {
    var senderIsMe = !!(sender && sender.address &&
      String(sender.address).toLowerCase() === myAddr);
    var list = senderIsMe ? (recipients || []) : (sender ? [sender] : []);
    var seen = {};
    return list.filter(function (t) {
      if (!t || !t.address) { return false; }
      var a = String(t.address).toLowerCase();
      if (a === myAddr || seen[a]) { return false; } // self + same address twice
      seen[a] = true;
      return true;
    });
  }

  /**
   * Who there is to add, from Office.js alone — no token, no network.
   *
   * Returns null when Office.js cannot tell (no sender on the item), in which
   * case the caller should sign in and let Graph decide, as before.
   *
   * This exists so a note to self — the one message where the answer is
   * "nobody" — never triggers a sign-in. On a phone, where an interactive
   * sign-in may not complete at all, checking first is the difference between
   * an instant answer and a dead end.
   */
  function targetsFromItem(item, myAddr) {
    if (!item || !item.from || !item.from.emailAddress) { return null; }
    return pickTargets(officeAddr(item.from), (item.to || []).map(officeAddr), myAddr);
  }

  /**
   * One short line describing the outcome.
   *
   * Kept to a single line because on mobile this is also the notification bar,
   * which is about 34 characters wide. Listing every name wrapped it into a
   * slab; a name when there is one and a count when there are several does not.
   */
  function summarize(added, enriched, skipped) {
    if (added.length && enriched.length) {
      return added.length + " added, " + enriched.length + " updated.";
    }
    var names = added.length ? added : enriched;
    if (!names.length) {
      return skipped ? "Already in your contacts." : "Nothing to add.";
    }
    var verb = added.length ? "Added" : "Updated";
    if (names.length === 1) { return verb + " " + names[0] + "."; }
    return verb + " " + names.length + " contacts.";
  }

  /**
   * Add (or enrich) everyone on the open message.
   *
   * deps: { restId, graph, parser, myAddr, onProgress }
   * onProgress(text) is optional and is called before each slow step, so a
   * surface with a screen can say what is happening instead of spinning
   * silently.
   *
   * Resolves to { added, enriched, skipped, empty, message }.
   */
  async function run(deps) {
    var graph = deps.graph, parser = deps.parser, myAddr = deps.myAddr;
    var say = deps.onProgress || function () {};

    say("Signing in…");
    var token = await graph.getToken();

    say("Reading the message…");
    var msg = await graph.getMessage(token, deps.restId);

    var sender = msg.from && msg.from.emailAddress;
    var targets = pickTargets(
      sender,
      (msg.toRecipients || []).map(function (r) { return r.emailAddress; }),
      myAddr
    );
    var senderIsMe = !!(sender && sender.address &&
      String(sender.address).toLowerCase() === myAddr);

    if (!targets.length) {
      return { added: [], enriched: [], skipped: 0, empty: true,
               message: "No one to add from this message." };
    }

    say("Checking your contacts…");
    var contacts = await graph.loadContacts(token);
    var added = [], enriched = [], skipped = 0;

    for (var i = 0; i < targets.length; i++) {
      var t = targets[i];
      var person = { email: t.address, name: t.name };
      say("Adding " + (person.name || person.email) + "…");

      // Mine the signature: from THIS message if they sent it, else from
      // their latest inbox message (best effort — blank card beats no card).
      var sig = null;
      if (!senderIsMe && sender && t.address === sender.address) {
        sig = parser.parse(msg.body && msg.body.content, person.name, person.email);
      } else {
        var recent = await graph.latestMessageFrom(token, t.address);
        if (recent) { sig = parser.parse(recent.body && recent.body.content, person.name, person.email); }
      }

      var payload = graph.toContactPayload(person, sig);
      var existing = contacts.byEmail[t.address.toLowerCase()];
      if (existing) {
        var r = await graph.enrichContact(token, existing, payload);
        if (r.updated) { enriched.push(person.name || person.email); } else { skipped++; }
      } else {
        await graph.createContact(token, payload);
        added.push(person.name || person.email);
      }
    }

    return {
      added: added, enriched: enriched, skipped: skipped, empty: false,
      message: summarize(added, enriched, skipped),
    };
  }

  var api = {
    officeAddr: officeAddr,
    pickTargets: pickTargets,
    targetsFromItem: targetsFromItem,
    summarize: summarize,
    run: run,
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.AddFlow = api; }
})(typeof self !== "undefined" ? self : this);
