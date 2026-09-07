/*
 * Auto Contacts — mobile task pane.
 *
 * Why this exists at all: with an ExecuteFunction button, the only thing an
 * add-in can put on a phone screen is a line of text in Outlook's own
 * notification bar — whose padding, alignment and icon are fixed by the host.
 * A task pane is the only surface where the add-in controls its own loading
 * state and layout, which is what the mobile design guidelines ask for.
 *
 * It runs on open — one tap from the message, same as before — and shows what
 * is happening while it works.
 *
 * The decisions all live in AddFlow, shared with the desktop command, so the
 * two surfaces cannot drift apart.
 */
/* global Office, GraphData, SigParser, AddFlow */
"use strict";

(function () {
  function byId(id) { return document.getElementById(id); }

  function show(which) {
    ["working", "done", "message"].forEach(function (id) {
      byId(id).hidden = (id !== which);
    });
  }

  function say(text) {
    byId("workingText").textContent = text;
  }

  /** A result screen listing who was added and who was updated. */
  function renderDone(res) {
    byId("doneTitle").textContent = res.message;
    var list = byId("doneList");
    list.textContent = "";

    function row(name, what) {
      var li = document.createElement("li");
      var who = document.createElement("span");
      who.className = "who";
      who.textContent = name;
      var sub = document.createElement("span");
      sub.className = "what";
      sub.textContent = what;
      li.appendChild(who);
      li.appendChild(sub);
      list.appendChild(li);
    }

    res.added.forEach(function (n) { row(n, "Added to your contacts"); });
    res.enriched.forEach(function (n) { row(n, "Details filled in"); });

    var note = byId("doneNote");
    if (res.skipped) {
      note.textContent = res.skipped === 1
        ? "1 person was already complete."
        : res.skipped + " people were already complete.";
      note.hidden = false;
    } else {
      note.hidden = true;
    }
    show("done");
  }

  function renderMessage(title, text, isError) {
    byId("messageTitle").textContent = title;
    byId("messageText").textContent = text || "";
    byId("message").classList.toggle("error", !!isError);
    show("message");
  }

  /**
   * Close the pane and go back to the message.
   *
   * On mobile a task pane covers the whole screen, and the docs recommend
   * closing it when the scenario is complete rather than making the user do
   * it. closeContainer isn't on every host, so falling back to window.close
   * keeps the button honest.
   */
  function closePane() {
    try {
      if (Office.context && Office.context.ui && Office.context.ui.closeContainer) {
        Office.context.ui.closeContainer();
        return;
      }
    } catch (e) { /* fall through */ }
    try { window.close(); } catch (e2) { /* nothing else to try */ }
  }

  async function start() {
    try {
      var item = Office.context.mailbox.item;
      var myAddr = ((Office.context.mailbox.userProfile || {}).emailAddress || "").toLowerCase();

      // Answer "nobody to add" without signing in — this is the note-to-self
      // case, and on a phone a needless sign-in can stall indefinitely.
      var early = AddFlow.targetsFromItem(item, myAddr);
      if (early && !early.length) {
        renderMessage("No one to add",
          "This message is only from you, so there are no contacts to add.", false);
        return;
      }

      var restId = Office.context.mailbox.convertToRestId(
        item.itemId, Office.MailboxEnums.RestVersion.v2_0);

      var res = await AddFlow.run({
        restId: restId,
        graph: GraphData,
        parser: SigParser,
        myAddr: myAddr,
        onProgress: say,
      });

      if (res.empty) {
        renderMessage("No one to add",
          "This message is only from you, so there are no contacts to add.", false);
        return;
      }
      renderDone(res);
    } catch (e) {
      // getToken's messages are already written for the device in hand, so
      // they are shown as-is rather than wrapped in something generic.
      renderMessage("Couldn't finish", (e && e.message) || String(e), true);
    }
  }

  Office.onReady(function () {
    byId("closeBtn").addEventListener("click", closePane);
    start();
  });
})();
