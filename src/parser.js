/*
 * Auto Contacts — signature parser (pure logic, no Office/Graph dependencies).
 *
 * Given an email body (HTML or plain text) and the sender's name/address,
 * extract contact fields from the signature block: phone numbers, job title,
 * company, website, LinkedIn, and a street address when present.
 *
 * Deliberately deterministic — no AI, no network. A partially-filled card is
 * the goal; anything ambiguous is left blank rather than guessed wrong.
 *
 * Works in the browser (global `SigParser`) and in Node (module.exports) so
 * the same file is unit-testable offline.
 */
(function (root) {
  "use strict";

  // ---------- HTML → text ----------

  function htmlToLines(html) {
    var text = String(html)
      .replace(/<style[\s\S]*?<\/style>/gi, " ")
      .replace(/<script[\s\S]*?<\/script>/gi, " ")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/(p|div|tr|li|h[1-6]|table)>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/&nbsp;/gi, " ")
      .replace(/&amp;/gi, "&")
      .replace(/&lt;/gi, "<")
      .replace(/&gt;/gi, ">")
      .replace(/&#39;|&apos;/gi, "'")
      .replace(/&quot;/gi, '"');
    return text.split("\n").map(function (l) { return l.replace(/\s+/g, " ").trim(); });
  }

  // Cut the body off at the first sign of a quoted earlier message, so we only
  // ever mine the sender's own (top) portion.
  var QUOTE_MARKERS = [
    /^from:\s/i,
    /^on .+ wrote:\s*$/i,
    /^-{3,}\s*original message\s*-{3,}/i,
    /^_{5,}\s*$/,
    /^>+\s/,
  ];

  function cutQuoted(lines) {
    for (var i = 0; i < lines.length; i++) {
      for (var m = 0; m < QUOTE_MARKERS.length; m++) {
        if (QUOTE_MARKERS[m].test(lines[i])) { return lines.slice(0, i); }
      }
    }
    return lines;
  }

  // ---------- field extractors ----------

  // Phone: 7+ digits once separators are removed, optionally labeled.
  var PHONE_RE = /(?:^|[\s:|•])((?:\+?\d{1,3}[\s.-]?)?(?:\(\d{2,4}\)[\s.-]?)?\d{2,4}[\s.-]\d{2,4}[\s.-]?\d{0,4})(?=$|[\s,;|•])/;
  var PHONE_LABELS = {
    mobile: /\b(m|mob|mobile|cell|c)\b[\s.:]/i,
    business: /\b(o|off|office|w|work|d|direct|desk|t|tel|ph|phone|p)\b[\s.:]/i,
    fax: /\b(f|fax)\b[\s.:]/i,
  };

  function digitCount(s) { return (s.match(/\d/g) || []).length; }

  function findPhones(lines) {
    var out = { mobile: null, business: null };
    lines.forEach(function (line) {
      // A line can hold several labeled numbers ("O: 515-555-1234 | C: 515-555-9876").
      var segments = line.split(/[|•·]+/);
      segments.forEach(function (seg) {
        var m = PHONE_RE.exec(seg);
        if (!m) { return; }
        var num = m[1].trim();
        if (digitCount(num) < 7 || digitCount(num) > 15) { return; }
        // Years/zips masquerade as phones; require a separator or "+".
        if (!/[\s().+-]/.test(num)) { return; }
        if (PHONE_LABELS.fax.test(seg)) { return; }
        if (PHONE_LABELS.mobile.test(seg)) {
          if (!out.mobile) { out.mobile = num; }
        } else if (PHONE_LABELS.business.test(seg)) {
          if (!out.business) { out.business = num; }
        } else if (!out.business) {
          out.business = num; // unlabeled → assume office line
        }
      });
    });
    return out;
  }

  var URL_RE = /\b((?:https?:\/\/)?(?:www\.)?[a-z0-9][a-z0-9-]*(?:\.[a-z0-9-]+)+(?:\/[^\s<>"']*)?)/i;

  function findUrls(lines, senderDomain) {
    var out = { website: null, linkedin: null };
    lines.forEach(function (line) {
      // Lines often carry several URLs ("www.acme.com | linkedin.com/in/…").
      line.split(/[\s|•·]+/).forEach(function (tok) {
        if (/@/.test(tok) && !/linkedin\.com/i.test(tok)) { return; } // email address, not a URL
        var m = URL_RE.exec(tok);
        if (!m) { return; }
        var url = m[1];
        if (/linkedin\.com/i.test(url)) {
          if (!out.linkedin) { out.linkedin = normalizeUrl(url); }
        } else if (!out.website) {
          out.website = normalizeUrl(url);
        } else if (senderDomain && url.toLowerCase().indexOf(senderDomain.toLowerCase()) !== -1 &&
                   out.website.toLowerCase().indexOf(senderDomain.toLowerCase()) === -1) {
          out.website = normalizeUrl(url); // upgrade to the sender's own domain
        }
      });
    });
    return out;
  }

  function normalizeUrl(url) {
    return /^https?:\/\//i.test(url) ? url : "https://" + url;
  }

  // Street address: "1234 Something St/Ave/Blvd..." optionally followed on the
  // same or next line by "City, ST 12345".
  var STREET_RE = /^\d{1,6}\s+[A-Za-z0-9 .'-]+\b(st|street|ave|avenue|blvd|boulevard|rd|road|dr|drive|ln|lane|way|ct|court|pkwy|parkway|hwy|highway|suite|ste)\b/i;
  var CITY_ST_ZIP_RE = /^[A-Za-z .'-]+,\s*[A-Z]{2}\s+\d{5}(-\d{4})?$/;

  function findAddress(lines) {
    for (var i = 0; i < lines.length; i++) {
      if (STREET_RE.test(lines[i])) {
        var street = lines[i];
        var rest = CITY_ST_ZIP_RE.test(lines[i + 1] || "") ? lines[i + 1] : null;
        return rest ? street + ", " + rest : street;
      }
    }
    return null;
  }

  // Title/company: the 1–3 non-field lines directly under the sender's name.
  var FIELD_LINE = /@|https?:|www\.|linkedin|\b(m|mob|mobile|cell|o|office|direct|tel|ph|phone|fax|f)\b[\s.:]|\d{3}[\s.-]\d{3,4}/i;
  var TITLE_WORDS = /\b(manager|director|engineer|officer|president|vp|vice|analyst|specialist|coordinator|lead|head|chief|architect|consultant|administrator|supervisor|planner|technician|developer|designer|owner|founder|principal|associate|assistant|attorney|counsel|agent|representative|ceo|cfo|cto|coo)\b/i;
  var COMPANY_WORDS = /\b(llc|inc|corp|co\.|ltd|company|group|systems|solutions|services|department|dept|agency|bureau|university|college|dot|division|associates|partners|technologies|consulting)\b/i;

  function findTitleCompany(lines, senderName) {
    var nameIdx = -1;
    if (senderName) {
      var nameLower = senderName.toLowerCase();
      var last = nameLower.split(/\s+/).pop();
      // Pass 1: the full display name. Keep the LAST occurrence — greeting
      // text ("Hi Jane,") can mention the name before the signature does.
      for (var i = 0; i < lines.length; i++) {
        var l = lines[i].toLowerCase();
        if (!l || FIELD_LINE.test(l)) { continue; } // emails/urls/phones can contain the name
        if (l.indexOf(nameLower) !== -1) { nameIdx = i; }
      }
      // Pass 2: surname only (signatures write "Jane P. Doe" for "Jane Doe") —
      // but not on lines that read like a title or company ("Lee Consulting Group").
      if (nameIdx === -1 && last && last.length > 2) {
        for (var i2 = 0; i2 < lines.length; i2++) {
          var l2 = lines[i2].toLowerCase();
          if (!l2 || FIELD_LINE.test(l2) || l2.length >= 60) { continue; }
          if (TITLE_WORDS.test(l2) || COMPANY_WORDS.test(l2)) { continue; }
          if (l2.indexOf(last) !== -1) { nameIdx = i2; }
        }
      }
    }
    var title = null, company = null;
    if (nameIdx === -1) { return { title: title, company: company }; }
    var taken = 0;
    for (var j = nameIdx + 1; j < lines.length && taken < 3; j++) {
      var line = lines[j];
      if (!line) { break; }               // blank line ends the block
      if (FIELD_LINE.test(line)) { break; } // phones/emails/urls end the block
      if (STREET_RE.test(line) || CITY_ST_ZIP_RE.test(line)) { break; }
      if (line.length > 70) { break; }     // prose, not a signature line
      taken++;
      if (!title && TITLE_WORDS.test(line) && !COMPANY_WORDS.test(line)) { title = line; continue; }
      if (!company && COMPANY_WORDS.test(line)) { company = line; continue; }
      if (!title && taken === 1) { title = line; }      // first line under name defaults to title
      else if (!company) { company = line; }            // next line defaults to company
    }
    return { title: title, company: company };
  }

  // ---------- main entry ----------

  /**
   * @param body        message body (HTML or plain text)
   * @param senderName  display name of the person whose signature we're mining
   * @param senderEmail their email address (for domain-preferring URL pick)
   * @returns {phones:{mobile,business}, title, company, website, linkedin, address}
   */
  function parse(body, senderName, senderEmail) {
    var isHtml = /<\s*(html|body|div|p|br|table)\b/i.test(String(body));
    var lines = isHtml ? htmlToLines(body) : String(body).split("\n").map(function (l) { return l.trim(); });
    lines = cutQuoted(lines);
    // The signature is the tail of the sender's own text: mine only the last
    // ~15 non-empty lines so body prose doesn't pollute the extraction.
    var nonEmpty = lines.filter(function (l) { return l; });
    var tail = nonEmpty.slice(-15);

    var senderDomain = null;
    if (senderEmail && senderEmail.indexOf("@") !== -1) {
      senderDomain = senderEmail.split("@")[1].replace(/^mail\./i, "");
    }

    var phones = findPhones(tail);
    var urls = findUrls(tail, senderDomain);
    var tc = findTitleCompany(tail, senderName);

    return {
      phones: phones,
      title: tc.title,
      company: tc.company,
      website: urls.website,
      linkedin: urls.linkedin,
      address: findAddress(tail),
    };
  }

  var api = { parse: parse, _internals: { htmlToLines: htmlToLines, cutQuoted: cutQuoted, findPhones: findPhones } };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.SigParser = api; }
})(typeof self !== "undefined" ? self : this);
