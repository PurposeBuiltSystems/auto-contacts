/*
 * Auto Contacts — one-click "Add to Contacts" ribbon command.
 *
 * On a received message: adds the SENDER to your contacts, mining their
 * signature from the open message for phone / title / company / website.
 * On a message you sent: adds the To recipients instead (their signatures are
 * mined from their most recent inbox message, when one exists).
 *
 * If the person already exists in Contacts, blank fields are enriched from
 * the signature — existing data is never overwritten.
 */
/* global Office, GraphData, SigParser */
"use strict";

async function addToContacts(event) {
  try {
    var item = Office.context.mailbox.item;
    var restId = Office.context.mailbox.convertToRestId(
      item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );

    var token = await GraphData.getToken();
    var msg = await GraphData.getMessage(token, restId);
    var my = await GraphData.me(token);
    var myAddr = (my.mail || my.userPrincipalName || "").toLowerCase();

    // Received message → the sender. Sent message → the To recipients.
    var sender = msg.from && msg.from.emailAddress;
    var targets;
    var senderIsMe = sender && sender.address && sender.address.toLowerCase() === myAddr;
    if (senderIsMe) {
      targets = (msg.toRecipients || []).map(function (r) { return r.emailAddress; });
    } else {
      targets = sender ? [sender] : [];
    }
    var seenAddr = {};
    targets = targets.filter(function (t) {
      if (!t || !t.address) { return false; }
      var a = t.address.toLowerCase();
      if (a === myAddr || seenAddr[a]) { return false; } // self + same address twice
      seenAddr[a] = true;
      return true;
    });
    if (!targets.length) {
      notify("error", "No one to add from this message.");
      finish(event);
      return;
    }

    var contacts = await GraphData.loadContacts(token);
    var added = [], enriched = [], skipped = 0;

    for (var i = 0; i < targets.length; i++) {
      var t = targets[i];
      var person = { email: t.address, name: t.name };

      // Mine the signature: from THIS message if they sent it, else from
      // their latest inbox message (best effort — blank card beats no card).
      var sig = null;
      if (!senderIsMe && sender && t.address === sender.address) {
        sig = SigParser.parse(msg.body && msg.body.content, person.name, person.email);
      } else {
        var recent = await GraphData.latestMessageFrom(token, t.address);
        if (recent) { sig = SigParser.parse(recent.body && recent.body.content, person.name, person.email); }
      }

      var payload = GraphData.toContactPayload(person, sig);
      var existing = contacts.byEmail[t.address.toLowerCase()];
      if (existing) {
        var r = await GraphData.enrichContact(token, existing, payload);
        if (r.updated) { enriched.push(person.name || person.email); } else { skipped++; }
      } else {
        await GraphData.createContact(token, payload);
        added.push(person.name || person.email);
      }
    }

    var parts = [];
    if (added.length) { parts.push("Added: " + added.join(", ")); }
    if (enriched.length) { parts.push("Updated: " + enriched.join(", ")); }
    if (!parts.length) { parts.push(skipped ? "Already in your contacts - nothing new to add." : "Done."); }
    notify("info", parts.join("  "));
    finish(event);
  } catch (e) {
    notify("error", "Auto Contacts failed: " + ((e && e.message) || e));
    finish(event);
  }
}

function notify(kind, text) {
  try {
    var item = Office.context.mailbox.item;
    if (!item || !item.notificationMessages) { return; }
    item.notificationMessages.replaceAsync("autoContacts", {
      type:
        kind === "error"
          ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
          : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: String(text).substring(0, 150),
      icon: "Icon.16",
      persistent: false,
    });
  } catch (e) { /* ignore */ }
}

function finish(event) {
  if (event && typeof event.completed === "function") { event.completed(); }
}

Office.onReady(function () {});
if (Office.actions && Office.actions.associate) {
  Office.actions.associate("addToContacts", addToContacts);
}
