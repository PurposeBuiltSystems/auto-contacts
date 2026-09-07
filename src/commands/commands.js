/*
 * Auto Contacts — one-click "Add to Contacts" ribbon command (desktop/web).
 *
 * On a received message: adds the SENDER to your contacts, mining their
 * signature from the open message for phone / title / company / website.
 * On a message you sent: adds the To recipients instead (their signatures are
 * mined from their most recent inbox message, when one exists).
 *
 * If the person already exists in Contacts, blank fields are enriched from
 * the signature — existing data is never overwritten.
 *
 * The decisions live in AddFlow, shared with the mobile task pane. This file
 * only turns the result into an Outlook notification.
 */
/* global Office, GraphData, SigParser, AddFlow */
"use strict";

async function addToContacts(event) {
  try {
    var item = Office.context.mailbox.item;
    // Own address comes from Office.js — no Graph /me call (no User.Read scope).
    var myAddr = ((Office.context.mailbox.userProfile || {}).emailAddress || "").toLowerCase();

    // Decide whether there is anyone to add BEFORE signing in. A note to self
    // is answered here, with no token at all.
    var early = AddFlow.targetsFromItem(item, myAddr);
    if (early && !early.length) {
      notify("error", "No one to add from this message.");
      finish(event);
      return;
    }

    var restId = Office.context.mailbox.convertToRestId(
      item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );

    var res = await AddFlow.run({
      restId: restId,
      graph: GraphData,
      parser: SigParser,
      myAddr: myAddr,
    });

    notify(res.empty ? "error" : "info", res.message);
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

// Node (the tests) has no Office, and the bootstrap must not run there.
if (typeof Office !== "undefined") {
  Office.onReady(function () {});
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("addToContacts", addToContacts);
  }
}
