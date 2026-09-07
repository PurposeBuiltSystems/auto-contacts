/*
 * The notification line, which on mobile is the ENTIRE user interface.
 *
 * Auto Contacts has no task pane on a phone - the manifest gives mobile an
 * ExecuteFunction button and nothing else - so this one string is everything
 * the user reads after tapping. The Store review flagged the notification as
 * too tall, and length is the only thing about that bar we control: Outlook
 * draws it, we supply the words.
 *
 * The mobile bar fits roughly 34 characters per line. These tests hold the
 * line to one.
 *
 * Run: npm test
 */
"use strict";
var C = require("../src/addflow.js");

var MOBILE_WIDTH = 34;
var failures = 0, passes = 0;

function check(label, actual, expected) {
  if (actual === expected) { passes++; return; }
  failures++;
  console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) +
    "\n  actual:   " + JSON.stringify(actual));
}

function fits(label, text) {
  if (text.length <= MOBILE_WIDTH) { passes++; return; }
  failures++;
  console.error("FAIL  " + label + "\n  " + text.length + " chars wraps the mobile bar (max " +
    MOBILE_WIDTH + ")\n  " + JSON.stringify(text));
}

// --- the wording ---

check("one added is named", C.summarize(["Jane Smith"], [], 0), "Added Jane Smith.");
check("one updated is named", C.summarize([], ["Jane Smith"], 0), "Updated Jane Smith.");
check("several added are counted", C.summarize(["A", "B", "C"], [], 0), "Added 3 contacts.");
// Both kinds at once drops to counts - naming them wraps the bar.
check("both at once", C.summarize(["A", "B"], ["C"], 0), "2 added, 1 updated.");
check("nothing new, but some skipped", C.summarize([], [], 3), "Already in your contacts.");
check("nothing at all", C.summarize([], [], 0), "Nothing to add.");

// --- the height, which is what the review was about ---

fits("one ordinary name", C.summarize(["Jane Smith"], [], 0));
fits("five added", C.summarize(["A", "B", "C", "D", "E"], [], 0));
fits("nothing new", C.summarize([], [], 2));

// The old code pasted every name in. This is the case that produced the slab.
var many = ["Robert Delacroix-Whitfield", "Ana Beltrán", "Jonathan Fitzgerald",
            "Priya Raghunathan", "Wei Zhang"];
fits("five long names", C.summarize(many, [], 0));
fits("long names on both sides", C.summarize(many, many, 0));

// A single very long name is still one line of text - it may wrap to two, but
// it is a name, and truncating a person's name to save a row is worse.
check("a long single name is kept whole",
  C.summarize(["Robert Delacroix-Whitfield"], [], 0),
  "Added Robert Delacroix-Whitfield.");

// --- and it never exceeds what replaceAsync accepts ---

var huge = [];
for (var i = 0; i < 500; i++) { huge.push("Person Number " + i); }
var out = C.summarize(huge, huge, 0);
if (out.length <= 150) { passes++; }
else {
  failures++;
  console.error("FAIL  500 contacts stays under the 150-char notification limit\n  got " +
    out.length + " chars");
}


// --- who gets added, decided before sign-in ---------------------------------
//
// The reviewer's screenshot was taken on a "Note to self" - the one message
// where the answer is "nobody". That answer is now reached without a token.

var me = "matt@iowadot.us";
function addr(a) { return { address: a, name: a }; }

check("a note to self has no one to add",
  C.pickTargets(addr(me), [addr(me)], me).length, 0);
check("a received message adds the sender",
  C.pickTargets(addr("jane@x.com"), [addr(me)], me)[0].address, "jane@x.com");
check("a sent message adds the recipients",
  C.pickTargets(addr(me), [addr("a@x.com"), addr("b@x.com")], me).length, 2);
check("my own address is never a target",
  C.pickTargets(addr(me), [addr("a@x.com"), addr(me)], me).length, 1);
check("the same address twice counts once",
  C.pickTargets(addr(me), [addr("a@x.com"), addr("A@X.com")], me).length, 1);
check("case does not hide me from myself",
  C.pickTargets(addr("MATT@IowaDot.us"), [addr(me)], me).length, 0);
check("no sender at all yields nobody",
  C.pickTargets(null, [], me).length, 0);
check("a recipient with no address is skipped",
  C.pickTargets(addr(me), [{ name: "no address" }, addr("a@x.com")], me).length, 1);

// The Office.js shape must map onto the same fields, or the pre-check would
// silently disagree with the Graph run.
check("Office.js addresses map across",
  C.officeAddr({ emailAddress: "jane@x.com", displayName: "Jane" }).address, "jane@x.com");
check("a missing Office.js address maps to null", C.officeAddr(null), null);
check("the two shapes agree on a note to self",
  C.pickTargets(C.officeAddr({ emailAddress: me }), [C.officeAddr({ emailAddress: me })], me).length, 0);

if (failures) {
  console.error("\n" + failures + " summary test(s) failed.");
  process.exit(1);
}
console.log("All " + passes + " notification-summary checks passed.");
