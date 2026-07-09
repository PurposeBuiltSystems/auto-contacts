/* Offline unit tests for the signature parser. Run: npm test */
"use strict";
var P = require("../src/parser.js");

var failures = 0;
function check(label, actual, expected) {
  var ok = actual === expected;
  if (!ok) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) + "\n  actual:   " + JSON.stringify(actual));
  }
}

// 1. Classic corporate HTML signature
var html1 = [
  "<div><p>Hi Matt,</p><p>Sounds good, see attached.</p><p>Thanks!</p>",
  "<p>Jane P. Doe<br>Senior Project Manager<br>Acme Systems LLC<br>",
  "O: 515-555-1234 | C: (515) 555-9876<br>",
  "1234 Grand Ave, Suite 200<br>Des Moines, IA 50309<br>",
  "www.acmesystems.com | linkedin.com/in/janedoe</p></div>",
].join("");
var r1 = P.parse(html1, "Jane Doe", "jane.doe@acmesystems.com");
check("html1 title", r1.title, "Senior Project Manager");
check("html1 company", r1.company, "Acme Systems LLC");
check("html1 business phone", r1.phones.business, "515-555-1234");
check("html1 mobile phone", r1.phones.mobile, "(515) 555-9876");
check("html1 website", r1.website, "https://www.acmesystems.com");
check("html1 linkedin", r1.linkedin, "https://linkedin.com/in/janedoe");
check("html1 address", r1.address, "1234 Grand Ave, Suite 200, Des Moines, IA 50309");

// 2. Plain-text signature, unlabeled phone, quoted thread below
var text2 = [
  "Sure, Friday works.",
  "",
  "Bob Smith",
  "Traffic Operations Engineer",
  "Iowa Department of Transportation",
  "515.555.4321",
  "",
  "From: Matt Miller <matt@example.com>",
  "Sent: Monday",
  "To: Bob Smith",
  "Subject: RE: Friday",
  "His old signature: 999-999-9999",
].join("\n");
var r2 = P.parse(text2, "Bob Smith", "bob.smith@iowadot.us");
check("text2 title", r2.title, "Traffic Operations Engineer");
check("text2 company", r2.company, "Iowa Department of Transportation");
check("text2 unlabeled phone -> business", r2.phones.business, "515.555.4321");
check("text2 quoted thread ignored", r2.phones.mobile, null);

// 3. Minimal signature — nothing to find, nothing invented
var r3 = P.parse("Thanks!\n\nSam", "Sam Jones", "sam@somewhere.org");
check("minimal no title", r3.title, null);
check("minimal no phone", r3.phones.business, null);
check("minimal no website", r3.website, null);

// 4. Email addresses are not mistaken for URLs; fax is skipped
var text4 = [
  "Best,",
  "Ann Lee",
  "Director of Operations",
  "Lee Consulting Group",
  "ann.lee@leeconsulting.com",
  "F: 800-555-0000",
  "M: +1 515-555-7777",
].join("\n");
var r4 = P.parse(text4, "Ann Lee", "ann.lee@leeconsulting.com");
check("text4 website not from email", r4.website, null);
check("text4 fax skipped", r4.phones.business, null);
check("text4 mobile", r4.phones.mobile, "+1 515-555-7777");
check("text4 company", r4.company, "Lee Consulting Group");

if (failures) {
  console.error("\n" + failures + " parser test(s) FAILED");
  process.exit(1);
}
console.log("All signature parser tests passed.");

// 5. Multiple phone numbers: two business + mobile, duplicate suppressed
var text5 = [
  "Regards,",
  "Carl Ortiz",
  "Field Operations Lead",
  "Ortiz Bridge Services LLC",
  "Office: 515-555-1000 | Direct: 515-555-2000 | Cell: 515-555-3000",
  "Tel: 515-555-1000",
].join("\n");
var r5 = P.parse(text5, "Carl Ortiz", "carl@ortizbridge.com");
check("text5 business #1", r5.phones.businessList[0], "515-555-1000");
check("text5 business #2", r5.phones.businessList[1], "515-555-2000");
check("text5 mobile", r5.phones.mobile, "515-555-3000");
check("text5 duplicate suppressed", r5.phones.businessList.length, 2);
