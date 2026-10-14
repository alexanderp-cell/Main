/**
 * Quote Desk — постоянный веб-интерфейс в Google Apps Script.
 * Ищет предложения по P/N в Gmail текущего аккаунта.
 */

var TUZ_HEADERS = [
  'Quote obtained date&time',
  'Initials',
  'Root Price per unit, $',
  'Root',
  'Supplier Price per unit, $',
  'Supplier',
  'Transit per unit, $',
  'Cond',
  'Lead Time (days)',
  'Qty Avail',
  'Cert Info',
  'PPWK',
  'Remarks'
];

function doGet() {
  // createHtmlOutputFromFile is more reliable than Template for a static page
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Quote Desk')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Quick self-check from the Apps Script editor: Run → checkDoGet_ */
function checkDoGet_() {
  var out = doGet();
  Logger.log('OK, content length: ' + out.getContent().length);
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Called from the browser UI. */
function searchQuotes(pnsText, limit) {
  limit = Math.max(5, Math.min(Number(limit) || 40, 80));
  var pns = parsePnText_(String(pnsText || ''));
  if (!pns.length) {
    throw new Error('Введите хотя бы один партийный номер.');
  }

  var results = [];
  var allRows = [];
  for (var i = 0; i < pns.length; i++) {
    var block = searchOnePn_(pns[i], limit);
    results.push(block);
    for (var j = 0; j < block.quotes.length; j++) {
      allRows.push(block.quotes[j].tuz_row);
    }
  }

  var tsv = TUZ_HEADERS.join('\t') + '\n';
  for (var r = 0; r < allRows.length; r++) {
    tsv += allRows[r].join('\t') + '\n';
  }

  return {
    ok: true,
    headers: TUZ_HEADERS,
    results: results,
    tsv: tsv.replace(/\n$/, ''),
    total_quotes: allRows.length
  };
}

function parsePnText_(text) {
  var seen = {};
  var out = [];
  var lines = String(text).split(/\r?\n/);
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].replace(/^\s+|\s+$/g, '');
    if (!line || line.charAt(0) === '#') continue;
    var parts = line.split(/[,;\s]+/);
    for (var j = 0; j < parts.length; j++) {
      var token = parts[j].replace(/^\s+|\s+$/g, '');
      if (!token) continue;
      var key = normalizePn_(token);
      if (seen[key]) continue;
      seen[key] = true;
      out.push(token);
    }
  }
  return out;
}

function normalizePn_(pn) {
  return String(pn).replace(/\s+/g, '').toUpperCase();
}

function softPn_(pn) {
  return normalizePn_(pn).replace(/[^A-Z0-9]/g, '');
}

function searchOnePn_(pn, limit) {
  var query = '"' + pn.replace(/"/g, '') + '"';
  var threads = GmailApp.search(query, 0, limit);
  var hits = [];
  var mailHits = 0;

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();
    // newest first inside thread
    for (var m = messages.length - 1; m >= 0; m--) {
      var msg = messages[m];
      mailHits++;
      var parsed = parseMessage_(pn, msg);
      for (var k = 0; k < parsed.length; k++) {
        hits.push(parsed[k]);
      }
      // Enough messages inspected for this PN
      if (mailHits >= limit) break;
    }
    if (mailHits >= limit) break;
  }

  hits = dedupeHits_(hits);
  return {
    pn: normalizePn_(pn),
    mail_hits: mailHits,
    quote_count: hits.length,
    quotes: hits
  };
}

function parseMessage_(pn, msg) {
  var subject = msg.getSubject() || '';
  var fromRaw = msg.getFrom() || '';
  var body = msg.getPlainBody() || '';
  if (!body) {
    body = stripHtml_(msg.getBody() || '');
  }
  var date = msg.getDate();
  var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  if (!messageAboutPn_(subject, body, pn)) return [];

  var head = topReply_(body);
  var blob = subject + '\n' + head;

  if (/\bno\s*quote\b|\bnq\b|unable to quote/i.test(blob) && !extractPrices_(head).length) {
    return [];
  }
  if (isLikelyInternal_(fromRaw) && !extractPrices_(head).length) {
    return [];
  }

  var supplier = supplierFromFrom_(fromRaw, head);
  var headConds = extractConds_(head);
  var headCert = extractCerts_(head);
  var headLt = extractLeadTime_(head);
  var headQty = extractQty_(head);
  var atts = [];
  try {
    var attachments = msg.getAttachments();
    for (var a = 0; a < attachments.length; a++) {
      atts.push(attachments[a].getName());
    }
  } catch (e) {}

  var lineOffers = structuredLineOffers_(head);
  var out = [];
  if (lineOffers.length) {
    for (var i = 0; i < lineOffers.length; i++) {
      var off = lineOffers[i];
      out.push(makeHit_(pn, dateStr, fromRaw, supplier, subject, off.price,
        off.cond || headConds[0] || '', off.qty || headQty, off.lead_time || headLt,
        off.cert || headCert, off.line, atts, head));
    }
    return out;
  }

  var prices = filterFeePrices_(head, extractPrices_(head));
  if (!prices.length) {
    if (atts.join(' ').match(/\.(pdf|xlsx?|csv)$/i)) {
      out.push(makeHit_(pn, dateStr, fromRaw, supplier, subject, null, '', '', '', '',
        'See attachment(s): ' + atts.join(', '), atts, head));
    }
    return out;
  }

  var cond = headConds[0] || '';
  if (!cond && /\brepaired\b/i.test(head)) cond = 'SV';
  out.push(makeHit_(pn, dateStr, fromRaw, supplier, subject, prices[0], cond,
    headQty || '1', headLt, headCert, head.replace(/\s+/g, ' ').slice(0, 220), atts, head));

  for (var p = 1; p < Math.min(prices.length, 3); p++) {
    if (prices[p] === prices[0] || prices[p] < 1000) continue;
    if (prices[0] && prices[p] < prices[0] * 0.05) continue;
    out.push(makeHit_(pn, dateStr, fromRaw, supplier, subject, prices[p],
      headConds[1] || cond, headQty, headLt, headCert,
      'additional price from same email', atts, head));
  }
  return out;
}

function makeHit_(pn, dateStr, fromRaw, supplier, subject, price, cond, qty, lt, cert, remarks, atts, head) {
  var priceS = (price == null || price === '') ? '' : String(price);
  var row = [dateStr, '', '', '', priceS, supplier, '', cond || '', lt || '', qty || '', cert || '', '', remarks || ''];
  return {
    pn: normalizePn_(pn),
    date: dateStr,
    from_raw: fromRaw,
    supplier: supplier,
    subject: subject,
    price: price,
    cond: cond || '',
    qty: qty || '',
    lead_time: lt || '',
    cert: cert || '',
    remarks: remarks || '',
    attachments: atts || [],
    snippet: String(head || '').replace(/\s+/g, ' ').slice(0, 300),
    tuz_row: row
  };
}

function messageAboutPn_(subject, body, pn) {
  var target = softPn_(pn);
  if (!target) return false;
  var head = topReply_(body);
  if (softPn_(subject).indexOf(target) !== -1) return true;
  if (softPn_(head).indexOf(target) !== -1) return true;
  if (/\b(quote|rfq|quotation|offer)\b/i.test(subject) &&
      softPn_(body).indexOf(target) !== -1 &&
      extractPrices_(head).length) {
    return true;
  }
  return false;
}

function topReply_(body) {
  var cut = body.length;
  var markers = [
    /\nOn .+ wrote:\n/i,
    /\nFrom:\s+.+\nSent:\s+/i,
    /\n-+Original Message-+/i,
    /\nPartsBase Quote Request/i,
    /\nRequest for Quote\n/i
  ];
  for (var i = 0; i < markers.length; i++) {
    var m = body.match(markers[i]);
    if (m && m.index != null) cut = Math.min(cut, m.index);
  }
  return body.slice(0, Math.min(cut, 1800)).replace(/^\s+|\s+$/g, '');
}

function stripHtml_(html) {
  return String(html)
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|tr)>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\u00a0/g, ' ');
}

function isLikelyInternal_(fromRaw) {
  return /@kt-mnt\.com>/i.test(fromRaw) || /@kt-mnt\.com$/i.test(fromRaw);
}

function supplierFromFrom_(fromRaw, head) {
  var name = '';
  var email = '';
  var m = String(fromRaw).match(/^(?:"?([^"]*)"?\s)?<?([^>]+@[^>]+)>?$/);
  if (m) {
    name = (m[1] || '').replace(/^\s+|\s+$/g, '').replace(/^'|'$/g, '');
    email = (m[2] || '').replace(/^\s+|\s+$/g, '');
  } else {
    name = String(fromRaw);
  }
  if (/ via /i.test(name)) name = name.split(/ via /i)[0].replace(/^\s+|\s+$/g, '');
  name = name.replace(/^['"]|['"]$/g, '');

  var co = head.match(/(?:quotation|quote|offer)\s+from\s+([A-Z][A-Za-z0-9 &.'\-\/]{2,60})/i);
  if (co && /^(sales|spares|info|offer|quotes|quote)$/i.test(name)) {
    return co[1].replace(/\.$/, '');
  }
  if (/^(sales|spares|info|offer|quotes|quote)$/i.test(name) && email) {
    var dom = email.split('@')[1] || '';
    var company = dom.split('.')[0] || name;
    return company.length <= 4 ? company.toUpperCase() : (company.charAt(0).toUpperCase() + company.slice(1));
  }
  return name || email || fromRaw;
}

function parseMoney_(token, suffix) {
  var v = parseFloat(String(token).replace(/,/g, ''));
  if (isNaN(v) || v <= 0 || v > 50000000) return null;
  if (suffix) {
    var s = String(suffix).toLowerCase();
    if (s === 'k') v *= 1000;
    if (s === 'm') v *= 1000000;
  }
  if (!suffix && v >= 1900 && v <= 2100) return null;
  if (v >= 10000000) return null;
  return v;
}

function extractPrices_(text) {
  var patterns = [
    /unit\s*price\s*[:\-]?\s*\$?\s*([\d,]+(?:\.\d+)?)\s*([kmb])?\s*(?:usd|\$)?/ig,
    /(?:price|offer(?:ed)?|cost|up)\s*[:\-]?\s*\$\s*([\d,]+(?:\.\d+)?)\s*([kmb])?/ig,
    /\$\s*([\d,]+(?:\.\d+)?)\s*([kmb])?\s*(?:each|ea|\/ea|usd|nett|net|\/1ea)?/ig,
    /\b([\d,]+(?:\.\d+)?)\s*([kmb])\s*(?:usd|\$|\/1ea|\/ea)?\b/ig,
    /\b([\d,]+(?:\.\d+)?)\s*(?:usd|\$)\b/ig,
    /(?:for|at)\s+\$\s*([\d,]+(?:\.\d+)?)\s*([kmb])?/ig
  ];
  var found = [];
  for (var i = 0; i < patterns.length; i++) {
    var re = patterns[i];
    var m;
    while ((m = re.exec(text)) !== null) {
      var suffix = m[2] && /[kmb]/i.test(m[2]) ? m[2] : null;
      var v = parseMoney_(m[1], suffix);
      if (v == null) continue;
      if (v < 20 && m[0].indexOf('$') === -1 && !/usd/i.test(m[0])) continue;
      if (found.indexOf(v) === -1) found.push(v);
    }
  }
  return found;
}

function extractConds_(text) {
  var re = /\b(SV\s*EX|IN\s*REP|AS\s*REMOVED|OVERHAULED|SERVICEABLE|REPAIRABLE|REPAIRED|REPAIR|NEW|NE|OH|SV|RP|AR|FN)\b/ig;
  var map = {
    NEW: 'NE', OVERHAULED: 'OH', SERVICEABLE: 'SV', 'AS REMOVED': 'AR',
    REPAIR: 'RP', REPAIRED: 'SV', REPAIRABLE: 'RP', 'SV EX': 'SV EX', 'IN REP': 'IN REP'
  };
  var out = [];
  var m;
  while ((m = re.exec(text)) !== null) {
    var c = m[1].replace(/\s+/g, ' ').toUpperCase();
    c = map[c] || c;
    if (out.indexOf(c) === -1) out.push(c);
  }
  return out;
}

function extractLeadTime_(text) {
  var m = text.match(/\b(\d+)\s*(?:day|days|d)\s*l\.?t\.?\b/i);
  if (m) return m[1];
  m = text.match(/(?:lead\s*time|delivery)\s*[:\-]?\s*([^\n|;]{1,40})/i);
  if (m) {
    var lt = m[1];
    var days = lt.match(/(\d+)\s*[-–]\s*(\d+)\s*(?:day|days|weeks?|w)\b/i) ||
               lt.match(/(\d+)\s*(?:day|days|d)\b/i) ||
               lt.match(/(\d+\s*[-–]\s*\d+|\d+)\s*weeks?\b/i);
    if (days) return days[0].replace(/\s+/g, '');
    if (/\bstock\b/i.test(lt)) return 'STOCK';
  }
  m = text.match(/\b(\d+\s*[-–]\s*\d+)\s*(?:business\s*)?days?\b/i);
  if (m) return m[1].replace(/\s+/g, '');
  if (/\b(in\s*stock|stock|ex\s*stock|ready\s*to\s*ship)\b/i.test(text)) return 'STOCK';
  return '';
}

function extractCerts_(text) {
  var re = /\b(8130(?:-3)?|EASA\s*Form[-\s]?1|Dual\s*Release|FAA|CoC|ARC|NIS|OEM\s*tag|trace(?:ability)?)\b/ig;
  var out = [];
  var m;
  while ((m = re.exec(text)) !== null) {
    var token = m[0].replace(/\s+/g, ' ');
    var key = token.toUpperCase();
    var exists = false;
    for (var i = 0; i < out.length; i++) {
      if (out[i].toUpperCase() === key) { exists = true; break; }
    }
    if (!exists) out.push(token);
    if (out.length >= 6) break;
  }
  return out.join(', ');
}

function extractQty_(text) {
  var m = text.match(/(?:qty|quantity|qnty)\s*[:\-]?\s*(\d+(?:\.\d+)?)\b/i);
  if (m) return m[1];
  m = text.match(/(?:^|[^0-9.,])(\d{1,4})\s*(?:ea|each|pcs|units?)\b/i);
  if (m) return m[1];
  m = text.match(/\/\s*(\d{1,4})\s*ea\b/i);
  if (m) return m[1];
  return '';
}

function structuredLineOffers_(text) {
  var lines = text.split(/\n/);
  var offers = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (!/(condition|unit price|\$\s*\d)/i.test(line)) continue;
    if (/(minimum\s*order|processing\s*\$|hazmat\s*fee|fee\s+is\s+added)/i.test(line)) continue;
    var prices = extractPrices_(line);
    if (!prices.length) continue;
    if (prices[0] < 500 && /(minimum\s*order)/i.test(text)) continue;
    var conds = extractConds_(line);
    offers.push({
      price: prices[0],
      cond: conds[0] || '',
      qty: extractQty_(line),
      lead_time: extractLeadTime_(line),
      cert: extractCerts_(line),
      line: line.replace(/^\s+|\s+$/g, '').slice(0, 240)
    });
  }
  return offers;
}

function filterFeePrices_(head, prices) {
  var feeRe = /(minimum\s*order|processing\s*\$|hazmat\s*fee|fee\s+is\s+added|under\s*\$\s*\d)/i;
  var out = [];
  for (var i = 0; i < prices.length; i++) {
    var p = prices[i];
    var keep = true;
    var sents = head.split(/[\n.]/);
    for (var s = 0; s < sents.length; s++) {
      var sent = sents[s];
      var sp = extractPrices_(sent);
      if (sp.indexOf(p) !== -1 && feeRe.test(sent)) { keep = false; break; }
    }
    if (keep) out.push(p);
  }
  return out.length ? out : prices.slice(0, 1);
}

function dedupeHits_(hits) {
  var seen = {};
  var out = [];
  for (var i = 0; i < hits.length; i++) {
    var h = hits[i];
    var key = [h.pn, String(h.supplier).toLowerCase(), h.price, h.cond, h.qty, (h.date || '').slice(0, 10)].join('|');
    if (seen[key]) continue;
    seen[key] = true;
    out.push(h);
  }
  out.sort(function (a, b) {
    if (a.date === b.date) return (b.price || 0) - (a.price || 0);
    return a.date < b.date ? 1 : -1;
  });
  return out;
}
