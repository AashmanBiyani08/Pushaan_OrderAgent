// ═══════════════════════════════════════════════════════════════
// PUSHAAN ORDER AGENT — WEB APP
// New standalone Apps Script project
// ═══════════════════════════════════════════════════════════════

var CONFIG = {
  SHEET_ID:        '1ii3L-HoC1zgCWJ37bpGQQ2vwhvCto-hqjU5iWQYKG1Y',
  NGROK_URL:       'https://dealer-armadillo-bauble.ngrok-free.dev',
  ORDER_LOG_TAB:   'Order Log',
  COUNTER_TAB:     'COUNTER',
  ORDER_SHEET_TAB: 'Order Sheet',
  LAST_ORDERS_TAB: 'LAST ORDERS',
  COMMENTS_TAB:    'SAVED COMMENTS',
};

// Cell references on the Order Sheet template
var CELL = {
  SPECIAL_COMMENT: 'B8',   // GM free-text special comment
  DISCOUNT:        'B9',   // discount value (was CHECK AND CONFIRM row)
  DUE:             'B53',  // OLD DUE / BALANCE
  TOTAL_BAGS:      'B44',  // TOTAL NUMBER OF BAGS
};

// Default quick-fill statements for Special Comment (GM can add more, stored in SAVED COMMENTS tab)
var DEFAULT_COMMENTS = [
  'Order jaldi nikalna hai',
  'Free sample jayenge iske saath',
  'Kal ki delivery hai',
];

var PACK_ROW_MAP = {
  'DIVINE UP':        { '5KG':14,'1KG':14,'500GR':15,'250GR':16 },
  'ROYAL RAJASTHANI': { '5KG':20,'1KG':20,'500GR':21,'250GR':22 },
  'AMCHI MUMBAI':     { '5KG':26,'1KG':26,'500GR':26,'250GR':27 },
  'SHRI KADAK':       { '5KG':31,'1KGP':32,'1KG BAG':33,'1KG':33,'500GR':34,'250GR':35,'50GR':36,'25GR':37 },
  'LOOSE':            { 'ROW1':41,'ROW2':42 },
};

var VARIANT_ALIAS = {
  'DUP':'DIVINE UP','RR':'ROYAL RAJASTHANI','AM':'AMCHI MUMBAI',
  'SK':'SHRI KADAK','NSK':'SHRI KADAK','SHRI':'SHRI KADAK','OF':'SHRI KADAK'
};

// ── WEB APP ENTRY ─────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Pushaan Order Agent')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── GET SAVED COMMENTS (default + GM-added) ───────────────────
function getSavedComments() {
  var ss  = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var tab = ss.getSheetByName(CONFIG.COMMENTS_TAB);
  var custom = [];
  if (tab) {
    var data = tab.getDataRange().getValues();
    custom = data.slice(1).map(function(r){ return String(r[0]).trim(); }).filter(Boolean);
  }
  // default first, then custom, de-duped
  var all = DEFAULT_COMMENTS.slice();
  custom.forEach(function(c){ if (all.indexOf(c) === -1) all.push(c); });
  return all;
}

// ── SAVE A NEW COMMENT TO THE PERMANENT LIST ──────────────────
function saveNewComment(text) {
  text = String(text || '').trim();
  if (!text) return { success:false, error:'Empty comment' };
  var ss  = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var tab = ss.getSheetByName(CONFIG.COMMENTS_TAB);
  if (!tab) {
    tab = ss.insertSheet(CONFIG.COMMENTS_TAB);
    tab.getRange(1,1).setValue('Saved Comment').setBackground('#1A237E').setFontColor('white').setFontWeight('bold');
  }
  // avoid duplicates (default or existing custom)
  var existing = getSavedComments();
  if (existing.indexOf(text) !== -1) return { success:true, comments:existing };
  tab.appendRow([text]);
  return { success:true, comments:getSavedComments() };
}

// ── PARSE ORDER MESSAGE ───────────────────────────────────────
function parseOrderMessage(message) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('CLAUDE_API_KEY');

  var prompt = 'Extract the tea order from the message below.\n' +
    'ALIASES (left = as written, right = correct variant name):\n' +
    'NSK = Shri Kadak, SHRI = Shri Kadak, DUP = Divine UP, RR = Royal Rajasthani, AM = Amchi Mumbai, LOOSE = Loose.\n' +
    'TYPE: "BP" means type BP. "OF" or "Dust" means type Dust. If not stated, type = UNKNOWN.\n' +
    'PACK SIZES: 1KG, 500GR, 250GR, 5KG, 50GR, 25GR.\n' +
    'IMPORTANT — DO NOT INVENT A PACK SIZE: only include a pack size if the customer actually wrote one (e.g. "1 KG", "500", "500G", "250", "5 KG", "50", "25"). If NO pack size is written in the line, leave the pack size BLANK in the quantity string — do NOT guess or default one. A downstream system will fill the pack size from the customer\'s previous order.\n' +
    'PACKING (Bag vs Pouch), applies ONLY when a 1 KG size is actually written: if the line contains "POUCH", "PACKET", or "PKT", the packing is Pouch. If the line contains "BAG" (Shri Kadak only), the packing is Bag. If a 1 KG Shri Kadak line has a size written but none of these packing words, default the packing to Pouch. Never add a packing word when no size was written.\n' +
    'QUANTITY FORMAT (single string):\n' +
    '- If a pack size IS written: "PACKSIZE x COUNT UNIT" for bags, or "PACKSIZE Pouch x COUNT UNIT" for pouches.\n' +
    '- If NO pack size is written: "x COUNT UNIT" (leave the pack blank before the x).\n' +
    'UNIT is "bags" by default. UNIT is "kg" ONLY if the message explicitly says kg. Examples:\n' +
    '- "RR BP 1 KG 3 BAGS" => quantity "1KG x 3 bags"\n' +
    '- "SHRI OF POUCH 1 KG 2 BAGS" => quantity "1KG Pouch x 2 bags"\n' +
    '- "1KG 50 kg" => quantity "1KG x 50 kg"\n' +
    '- "DUP PREMIUM 2 BAGS" (no size written) => quantity "x 2 bags"  (pack left blank)\n' +
    '- "NSK BP 4 BAGS" (no size written) => quantity "x 4 bags"  (pack left blank)\n' +
    'LOOSE TEA: "Loose" is a valid variant. It may be written as "LOOSE BP", "LOOSE DUST", "LOOSE DUP", "LOOSE TEA DUP" etc. Always set variant = "Loose". For "LOOSE DUP" or "LOOSE TEA DUP", set type = "BP". For "LOOSE DUST", set type = "Dust". For "LOOSE BP", set type = "BP". If a specific loose tea name is written (e.g. "Assam CTC"), put it in a "looseName" field. Quantity = "COUNT UNIT" where UNIT is "bags" if bags written or nothing written, and "kg" if kg is explicitly written. Example: "LOOSE BP 20 KG @300/-" => variant "Loose", type "BP", quantity "20 kg", rate "300".\n' +
    'RATE: extract only the numeric value of the rate, with no symbols. Remove any @ sign, currency symbol, and any trailing /-. Example: "@300/-" becomes "300". "@ 14" becomes "14". If no rate is written, rate = "UNKNOWN".\n' +
    'MULTIPLE LINES: each order line is a SEPARATE order entry. Never merge lines. Never drop a line. Include the original line text in a "rawText" field for each order.\n' +
    'NEVER GUESS THE VARIANT: if the customer does not write a brand/variant name (Shri Kadak, NSK, DUP, Divine, RR, Royal, Amchi, Loose) in a line, you MUST return variant = "UNKNOWN" for that line. Do not infer or assume a variant from pack size or quantity.\n' +
    'VAGUE: set "vague" to "YES" only if no brand/variant is named at all (not even Loose). If variant IS named, set vague="NO" regardless of whether rate is stated. RPO overrides vague — if message ends with RPO, set vague="RPO" regardless.\n' +
    'RPO: if the message ends with "RPO" (case insensitive), set vague="RPO", variant=UNKNOWN, type=UNKNOWN, rate=UNKNOWN, quantity=RPO.\n' +
    'Return EXACTLY this JSON structure and nothing else:\n' +
    '{"customer": "party name or UNKNOWN", "orders": [{"variant": "Shri Kadak or Divine UP or Royal Rajasthani or Amchi Mumbai or Loose or UNKNOWN", "type": "BP or Dust or UNKNOWN", "quantity": "as specified above", "rate": "exact rate text or UNKNOWN", "looseName": "loose tea name or empty", "rawText": "original line text"}], "vague": "YES or NO"}\n\n' +
    'MESSAGE: ' + message;

  try {
    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 1000,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    });

    var raw = JSON.parse(response.getContentText()).content[0].text
      .trim().replace(/```json|```/g,'').trim();
    // Extract the JSON object even if the model added preamble text
    var jsonStart = raw.indexOf('{');
    var jsonEnd   = raw.lastIndexOf('}');
    if (jsonStart !== -1 && jsonEnd !== -1) {
      raw = raw.substring(jsonStart, jsonEnd + 1);
    }
    var parsed = JSON.parse(raw);

    parsed.orders = (parsed.orders || []).map(function(o) {
      var vk = (o.variant || '').toUpperCase();
      if (VARIANT_ALIAS[vk]) o.variant = VARIANT_ALIAS[vk];
      var vmap = {
        'shri kadak':'SHRI KADAK','divine up':'DIVINE UP',
        'royal rajasthani':'ROYAL RAJASTHANI','amchi mumbai':'AMCHI MUMBAI','loose':'LOOSE'
      };
      var vl = (o.variant||'').toLowerCase();
      if (vmap[vl]) o.variant = vmap[vl];
      return o;
    });

    return { success: true, data: parsed };
  } catch(e) {
    return { success: false, error: String(e) };
  }
}

// ── ENRICH WITH TALLY ─────────────────────────────────────────
function enrichWithTally(party, orders) {
  var result = {
    resolvedName:  null,
    partyWarning:  '',
    dueText:       '',
    discountText:  '',
    discount:      null,
    orders:        orders
  };

  result.resolvedName = resolveLedgerName_(party);
  if (!result.resolvedName) {
    result.partyWarning = 'Party not found in Tally — check spelling';
  }
  var lookupName = result.resolvedName || party;

  try {
    var due = getDueBalanceForParty(lookupName);
    result.dueText = (due === null) ? 'LOOKUP_FAILED' : (due === 0) ? 'No dues' : 'Rs ' + due + ' due';
  } catch(e) { result.dueText = 'LOOKUP_FAILED'; }

  try {
    var disc = getDiscountForParty_(lookupName);
    if (disc === null)          { result.discountText = 'No discount found in last orders'; }
    else if (disc === 0)        { result.discountText = 'LOOKUP_FAILED'; }
    else if (disc.discount) {
      result.discountText = 'Rs ' + disc.discount + '/kg discount in previous order';
      result.discount     = disc.discount;
      result.discountDate = disc.date;
    }
  } catch(e) { result.discountText = 'LOOKUP_FAILED'; }
  var history = [];
  try { history = getLastOrdersForParty(lookupName); } catch(e) {}

  result.orders = orders.map(function(o) {
    var enriched = JSON.parse(JSON.stringify(o));
    // If the customer wrote no variant word, the parser's variant is a hallucination — reset to UNKNOWN
    if (!/shri|kadak|nsk|\bsk\b|dup|\bdp\b|divine|\brr\b|royal|rajasthani|\bam\b|amchi|mumbai|loose/i.test(o.rawText || '')) {
      enriched.variant = 'UNKNOWN';
    }
    
    enriched.rateSource   = '';
    enriched.packSource   = '';
    enriched.packInferred = false;
    enriched.rateInferred = false;

    // ── STEP 1: extract pack, count, and unit from the parser's quantity string ──
    // The parse returns quantity like "1KG x 10 kg", "500GR x 20 bags", "1KG Pouch x 2 bags",
    // or for loose "10 kg" / "5 bags". Pack is the leading token; count+unit follow the "x".
    var qstr = String(enriched.quantity || '').trim();
    var explicitPack = '';   // pack the CUSTOMER stated (wins over inference)
    var count = '';
    var unit  = '';

    // Match "PACK x COUNT UNIT"  (pack may include the word Pouch)
    var m = qstr.match(/^(.*?)\s*x\s*([\d.]+)\s*(kgs?|bags?)?/i);
    if (m) {
      explicitPack = (m[1] || '').trim();
      count = m[2];
      unit  = (m[3] || 'bags').toLowerCase();
    } else {
      // Loose form "COUNT UNIT" with no pack
      var m2 = qstr.match(/^([\d.]+)\s*(kgs?|bags?)?/i);
      if (m2) { count = m2[1]; unit = (m2[2] || 'bags').toLowerCase(); }
    }
    // Normalise unit to singular canonical 'kg' or 'bags'
    if (/^kg/i.test(unit)) unit = 'kg'; else if (/^bag/i.test(unit)) unit = 'bags';

    // Normalise the explicit pack the parser gave into our canonical pack labels
    if (explicitPack) {
      var ep = explicitPack.toUpperCase().replace(/\s+/g,' ').trim();
      var isShri = (o.variant || '').toUpperCase() === 'SHRI KADAK';
      // Priority: pouch/packet/pkt beats bag beats bare 1KG
      if (/POUCH|PACKET|PKT/.test(ep)) {
        enriched.pack = isShri ? '1KG POUCH' : '1KG';
      } else if (/500/.test(ep)) {
        enriched.pack = '500GR';
      } else if (/250/.test(ep)) {
        enriched.pack = '250GR';
      } else if (/\b50\b|50\s*GR|50GR/.test(ep)) {
        enriched.pack = '50GR';
      } else if (/\b25\b|25\s*GR|25GR/.test(ep)) {
        enriched.pack = '25GR';
      } else if (/5\s*KG/.test(ep)) {
        enriched.pack = '5KG';
      } else if (/BAG/.test(ep) && isShri) {
        enriched.pack = '1KG BAG';
      } else if (/1\s*KG/.test(ep)) {
        // Shri: bare 1KG with no packing word → pouch (default). Others: plain 1KG.
        enriched.pack = isShri ? '1KG POUCH' : '1KG';
      } else {
        enriched.pack = ep; // leave as-is for GM to correct
      }
    }

    enriched.qtyCount = count;
    enriched.unit     = unit || 'bags';

    // ── STEP 2: Tally history — rate always, pack ONLY if customer gave none ──
    if (history && history.length > 0) {
      var vl = (o.variant || '').toLowerCase();
      var vkw = {
        'divine up':['dup','divine',' dp '],'royal rajasthani':['royal','rajasthani',' rr '],
        'amchi mumbai':['amchi','mumbai',' am '],'shri kadak':['shri','kadak',' sk ',' nsk '],'loose':['loose']
      };
      var keys = vkw[vl] || [vl.split(' ')[0]];

      var foundRate = '', foundPack = '', foundType = '', foundDate = '';
      for (var h = 0; h < history.length; h++) {
        for (var i = 0; i < history[h].items.length; i++) {
          var itm = history[h].items[i];
          var il = itm.item.toLowerCase();
          if (!keys.some(function(k){ return il.indexOf(k) !== -1; })) continue;
          if (!foundRate && itm.rate) { foundRate = itm.rate; foundDate = history[h].rawDate; }
          if (!foundType) {
            if      (/\bdust\b/i.test(il)) foundType = 'DUST';
            else if (/\bbp\b/i.test(il))   foundType = 'BP';
          }
          if (!foundPack) {
            if      (il.indexOf('25 gm')!== -1||il.indexOf('25gm')!== -1)  foundPack='25GR';
            else if (il.indexOf('50 gm')!== -1||il.indexOf('50gm')!== -1)  foundPack='50GR';
            else if (il.indexOf('250 gm')!==-1||il.indexOf('250gm')!==-1)  foundPack='250GR';
            else if (il.indexOf('500 gm')!==-1||il.indexOf('500gm')!==-1)  foundPack='500GR';
            else if (il.indexOf('5 kg') !==-1||il.indexOf('5kg') !==-1)    foundPack='5KG';
            else if (il.indexOf('bag')  !==-1)                              foundPack='1KG BAG';
            else if (il.indexOf('pouch')!==-1)                              foundPack='1KG POUCH';
          }
        }
        if (foundRate && foundPack) break;
      }

      // Rate: fill only if the customer didn't state one
      if (foundRate && (!enriched.rate || enriched.rate === 'UNKNOWN')) {
        enriched.rate = foundRate;
        enriched.rateSource = 'Tally order ' + foundDate;
        enriched.rateInferred = true;
      }

      // Pack: fill ONLY when the customer gave none. Never override an explicit pack.
      if (!enriched.pack && foundPack) {
        enriched.pack = foundPack;
        enriched.packSource = 'Tally order ' + foundDate;
        enriched.packInferred = true;
      }

      // Type: fill ONLY when the customer gave none
      if ((!enriched.type || enriched.type === 'UNKNOWN') && foundType) {
        enriched.type = foundType;
        enriched.typeSource = 'Tally order ' + foundDate;
        enriched.typeInferred = true;
      }

      // Loose name: if variant is Loose and customer didn't name it, pull the item name from history
      if ((o.variant || '').toUpperCase() === 'LOOSE' && (!enriched.looseName || enriched.looseName === '')) {
        for (var lh = 0; lh < history.length; lh++) {
          for (var li = 0; li < history[lh].items.length; li++) {
            var lit = history[lh].items[li].item;
            if (lit.toLowerCase().indexOf('loose') !== -1) {
              enriched.looseName = lit.trim();
              enriched.looseNameInferred = true;
              enriched.looseNameSource = 'Tally order ' + history[lh].rawDate;
              break;
            }
          }
          if (enriched.looseName) break;
        }
      }
    }

    // If the customer's raw text has NO variant word, the parser's variant is a guess — ignore it and match by pack size
      var rawHasVariant = /shri|kadak|nsk|\bsk\b|dup|\bdp\b|divine|\brr\b|royal|rajasthani|\bam\b|amchi|mumbai|loose/i.test(o.rawText || '');
      if (!enriched.pack && o.rawText) {
        var rt = o.rawText.toLowerCase();
        if      (/250\s*gr|250\s*gm|250/.test(rt)) enriched.pack = '250GR';
        else if (/500\s*gr|500\s*gm|500/.test(rt)) enriched.pack = '500GR';
        else if (/25\s*gr|25\s*gm/.test(rt))       enriched.pack = '25GR';
        else if (/50\s*gr|50\s*gm/.test(rt))       enriched.pack = '50GR';
        else if (/5\s*kg/.test(rt))                enriched.pack = '5KG';
        else if (/1\s*kg|1kg/.test(rt))            enriched.pack = '1KG BAG';
      }
      
      Logger.log('DBG want=' + enriched.pack + ' rawHasVariant=' + rawHasVariant + ' items=' + JSON.stringify((history||[]).map(function(h){return h.items.map(function(i){return i.item;});})));

      Logger.log('GUARD rawHasVariant=' + rawHasVariant + ' pack=[' + enriched.pack + ']');
      if (!rawHasVariant && enriched.pack) {
        var wantPack = enriched.pack.toUpperCase();
        var matchedItem = '', matchedDate = '', matchedRate = '';
        for (var vh = 0; vh < history.length; vh++) {
          for (var vi = 0; vi < history[vh].items.length; vi++) {
            var vit = history[vh].items[vi];
            var vil = vit.item.toLowerCase();
            var itemPack = '';
            if      (vil.indexOf('250 gm')!==-1||vil.indexOf('250gm')!==-1) itemPack='250GR';
            else if (vil.indexOf('500 gm')!==-1||vil.indexOf('500gm')!==-1) itemPack='500GR';
            else if (vil.indexOf('25 gm')!==-1||vil.indexOf('25gm')!==-1)   itemPack='25GR';
            else if (vil.indexOf('50 gm')!==-1||vil.indexOf('50gm')!==-1)   itemPack='50GR';
            else if (vil.indexOf('5 kg')!==-1||vil.indexOf('5kg')!==-1)     itemPack='5KG';
            else if (vil.indexOf('bag')!==-1)                                itemPack='1KG BAG';
            else if (vil.indexOf('pouch')!==-1)                              itemPack='1KG POUCH';
            // compare loosely (1KG matches 1KG BAG/POUCH)
            var wp = wantPack.replace(' POUCH','').replace(' BAG','');
            var ip = itemPack.replace(' POUCH','').replace(' BAG','');
            Logger.log('CMP item=[' + vit.item + '] itemPack=[' + itemPack + '] wantPack=[' + wantPack + '] hit=' + (itemPack === wantPack || ip === wp));
            if (itemPack === wantPack || ip === wp) {
              matchedItem = vit.item; matchedDate = history[vh].rawDate; matchedRate = vit.rate;
              break;
            }
          }
          if (matchedItem) break;
        }
        if (matchedItem) {
          var ml = matchedItem.toLowerCase();
          var mv = '';
          if      (ml.indexOf('shri')!==-1||ml.indexOf('kadak')!==-1)       mv='SHRI KADAK';
          else if (ml.indexOf('divine')!==-1||/\bdup\b/.test(ml))            mv='DIVINE UP';
          else if (ml.indexOf('royal')!==-1||ml.indexOf('rajasthani')!==-1||/\brr\b/.test(ml))  mv='ROYAL RAJASTHANI';
          else if (ml.indexOf('amchi')!==-1||ml.indexOf('mumbai')!==-1)      mv='AMCHI MUMBAI';
          else if (ml.indexOf('loose')!==-1)                                  mv='LOOSE';
          if (mv) { enriched.variant = mv; enriched.variantInferred = true; enriched.variantSource = 'Tally order ' + matchedDate; }
          if (ml.indexOf('dust')!==-1)      { enriched.type = 'DUST'; enriched.typeInferred = true; }
          else if (/\bbp\b/.test(ml))        { enriched.type = 'BP';   enriched.typeInferred = true; }
          if (matchedRate && (!enriched.rate || enriched.rate === 'UNKNOWN')) {
            enriched.rate = matchedRate; enriched.rateInferred = true; enriched.rateSource = 'Tally order ' + matchedDate;
          }
        } else {
          enriched.variant = 'UNKNOWN';
          enriched.variantNotFound = true;
        }
      }

    return enriched;
  });

  return result;
}

// ── ENRICH PARTY ONLY (for manual mode party lookup, if ever needed) ──
function enrichPartyOnly(party) {
  return enrichWithTally(party, []);
}

// ── RPO ───────────────────────────────────────────────────────
function getAllParties() {
  var ss  = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var tab = ss.getSheetByName(CONFIG.LAST_ORDERS_TAB);
  if (!tab) return [];
  var data = tab.getDataRange().getValues();
  return data.slice(1).map(function(r){ return String(r[0]).trim(); }).filter(Boolean).sort();
}

function getLastOrderForRPO(party) {
  var ss  = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var tab = ss.getSheetByName(CONFIG.LAST_ORDERS_TAB);
  if (!tab) return null;
  var data = tab.getDataRange().getValues();
  var pl = party.trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === pl) {
      try { return { party:data[i][0], batchNo:data[i][1], lineItems:JSON.parse(data[i][3]) }; }
      catch(e) { return null; }
    }
  }
  return null;
}

function saveLastOrder_(party, batchNo, lineItems) {
  var ss  = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var tab = ss.getSheetByName(CONFIG.LAST_ORDERS_TAB);
  if (!tab) {
    tab = ss.insertSheet(CONFIG.LAST_ORDERS_TAB);
    tab.getRange(1,1,1,4).setValues([['Party','Batch No','Date','Line Items (JSON)']]);
    tab.getRange(1,1,1,4).setBackground('#1A237E').setFontColor('white').setFontWeight('bold');
  }
  var data = tab.getDataRange().getValues();
  var row  = [party, batchNo, new Date(), JSON.stringify(lineItems)];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === party.trim().toLowerCase()) {
      tab.getRange(i+1,1,1,4).setValues([row]); return;
    }
  }
  tab.appendRow(row);
}

// ── APPROVE ORDER ─────────────────────────────────────────────
function approveOrder(payload) {
  var ss      = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var counter = ss.getSheetByName(CONFIG.COUNTER_TAB);
  var log     = ss.getSheetByName(CONFIG.ORDER_LOG_TAB);
  var template= ss.getSheetByName(CONFIG.ORDER_SHEET_TAB);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { success:false, error:'Could not acquire lock — try again' };

  try {
    var lastNum = parseInt(String(counter.getRange('A1').getValue()).replace(/[^0-9]/g,''),10);
    if (isNaN(lastNum)) lastNum = 278;
    lastNum++;
    var batchNo = 'G' + lastNum;
    var party   = payload.resolvedName || payload.party;

    var monthNames = ['January','February','March','April','May','June','July',
                      'August','September','October','November','December'];
    var now = new Date();
    var mFolder = monthNames[now.getMonth()] + ' ' + now.getFullYear();
    var pf = DriveApp.getFoldersByName('Pushaan Order Agent');
    var parent = pf.hasNext() ? pf.next() : DriveApp.createFolder('Pushaan Order Agent');
    var mf = parent.getFoldersByName(mFolder);
    var folder = mf.hasNext() ? mf.next() : parent.createFolder(mFolder);

    var fn  = batchNo + ' - ' + party;
    var nSS = SpreadsheetApp.create(fn);
    var nF  = DriveApp.getFileById(nSS.getId());
    folder.addFile(nF); DriveApp.getRootFolder().removeFile(nF);
    var nSheet = template.copyTo(nSS);
    nSheet.setName(fn.substring(0,99));
    var ds = nSS.getSheets()[0];
    if (ds.getName() !== nSheet.getName()) nSS.deleteSheet(ds);

    // Header
    nSheet.getRange('B2').setValue(batchNo);
    nSheet.getRange('B4').setValue(now).setNumberFormat('DD-MMM-YYYY');
    nSheet.getRange('B6').setValue(party);
    if (payload.dueText)         nSheet.getRange(CELL.DUE).setValue(payload.dueText);
    if (payload.specialComment)  nSheet.getRange(CELL.SPECIAL_COMMENT).setValue(payload.specialComment);
    if (payload.discountText && payload.discountText !== 'No discount found in last orders')
      nSheet.getRange(CELL.DISCOUNT).setValue(payload.discountText);

    // Line items
    var looseRowUsed = 0; // 0 = none yet, then 41, then 42
    payload.orders.forEach(function(item) {
      var vk = (item.variant||'').toUpperCase();
      if (VARIANT_ALIAS[vk]) vk = VARIANT_ALIAS[vk];
      var packRows = PACK_ROW_MAP[vk];
      if (!packRows) return;

      var col = (item.type||'').toUpperCase()==='DUST' ? 3 : 2;
      var isLoose = (vk==='LOOSE');

      var qtyNum  = item.qtyCount || item.qty || '';
      var qtyUnit = (item.unit === 'kg') ? 'kg' : 'bags';

      if (isLoose) {
        // Loose goes to row 41 then 42; col B for BP, col C for DUST
        var lrow = (looseRowUsed === 0) ? 41 : 42;
        looseRowUsed++;
        var lname = (item.looseName || '').trim();
        var cellVal = (lname ? lname + ' — ' : '') + qtyNum + ' ' + qtyUnit;
        if (item.rate && item.rate!=='UNKNOWN') cellVal += ' @ ' + item.rate + '/-';
        nSheet.getRange(lrow, col).setValue(cellVal)
          .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
        return;
      }

      var pk = (item.pack||'').toUpperCase();
      // Map 1KG POUCH -> 1KGP row key for Shri Kadak
      if (pk === '1KG POUCH') pk = '1KGP';
      var rowNum = packRows[pk] || packRows['1KG BAG'] || packRows['1KG'] || null;
      if (!rowNum) return;

      // Unit stays as the customer wrote it — kg is kg, bags is bags, never converted
      var cellVal2 = (item.pack||'') + ' x ' + qtyNum + ' ' + qtyUnit;
      if (item.rate && item.rate!=='UNKNOWN') cellVal2 += ' @ ' + item.rate;
      nSheet.getRange(rowNum, col).setValue(cellVal2)
        .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
    });

    var totalBags = 0, totalKg = 0;
    payload.orders.forEach(function(item) {
      var n = parseFloat(item.qtyCount || item.qty || 0);
      if (isNaN(n)) return;
      if ((item.unit || 'bags') === 'kg') totalKg += n; else totalBags += n;
    });
    var totalText = '';
    if (totalKg > 0) totalText += totalKg + ' KG';
    if (totalBags > 0) { if (totalText) totalText += ' + '; totalText += totalBags + ' BAGS'; }
    if (totalText) nSheet.getRange(CELL.TOTAL_BAGS).setValue(totalText)
      .setHorizontalAlignment('center').setFontWeight('bold');


    // Order Log
    payload.orders.forEach(function(item) {
      var qStr = item.pack ? (item.pack+' x '+(item.qtyCount||item.qty||'')+' '+(item.unit||'bags')) : (item.quantity||'');
      log.appendRow([batchNo, now, party, item.variant||'', item.type||'',
                     qStr, item.rate||'', payload.dueText||'', 'OLD','NO','PENDING APPROVAL']);
    });

    counter.getRange('A1').setValue('G'+lastNum);
    saveLastOrder_(party, batchNo, payload.orders);

    return { success:true, batchNo:batchNo, url:nSS.getUrl() };

  } catch(e) {
    return { success:false, error:String(e) };
  } finally {
    lock.releaseLock();
  }
}

// ── TALLY HELPERS ─────────────────────────────────────────────
function normalizeTallyName_(n){
  return String(n||'').toLowerCase().replace(/[()\/.,-]/g,' ').replace(/\s+/g,' ').trim();
}
function pickLedgerFilterWord_(name){
  var stop={ji:1,shri:1,sri:1,store:1,kirana:1,traders:1,trader:1,trading:1,
            company:1,co:1,agency:1,agencies:1,general:1,provision:1,mart:1};
  var words=String(name||'').replace(/[()\/.,-]/g,' ').split(/\s+/)
    .filter(function(w){return w.length>=3&&!stop[w.toLowerCase()];});
  if(!words.length)return String(name||'').trim().split(/\s+/)[0]||'';
  words.sort(function(a,b){return b.length-a.length;}); return words[0];
}
function resolveLedgerName_(party){
  var sw=pickLedgerFilterWord_(party).replace(/"/g,'');
  if(!sw)return null;
  var xml='<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>PushLedList</ID></HEADER>'+
    '<BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>'+
    '<TDL><TDLMESSAGE><COLLECTION NAME="PushLedList" ISMODIFY="No" ISFIXED="No" ISINITIALIZE="Yes" ISOPTION="No" ISINTERNAL="No">'+
    '<TYPE>Ledger</TYPE><FETCH>Name</FETCH><BELONGSTO>Sundry Debtors</BELONGSTO><FILTER>NmFilter</FILTER></COLLECTION>'+
    '<SYSTEM TYPE="Formulae" NAME="NmFilter">$Name CONTAINS "'+sw+'"</SYSTEM>'+
    '</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>';
  try{
    var t=UrlFetchApp.fetch(CONFIG.NGROK_URL,{method:'POST',contentType:'text/xml',payload:xml,
      headers:{'ngrok-skip-browser-warning':'1'},muteHttpExceptions:true}).getContentText();
    var names=(t.match(/<LEDGER NAME="([^"]*)"/g)||[]).map(function(n){return n.replace(/<LEDGER NAME="/,'').replace(/"$/,'');});
    var full=String(party||'').toLowerCase().trim();
    var words=full.split(/\s+/).filter(function(x){return x.length>=2;});
    for(var i=0;i<names.length;i++){if(names[i].toLowerCase().trim()===full)return names[i];}
    var best=null,bs=-1;
    names.forEach(function(n){
      var lw=n.toLowerCase().replace(/[()]/g,' ').split(/\s+/).filter(function(x){return x.length>=2;});
      var sc=0; words.forEach(function(w){if(lw.some(function(l){return l.indexOf(w)!==-1||w.indexOf(l)!==-1;}))sc++;});
      if(sc>bs){bs=sc;best=n;}
    });
    return(bs>=words.length-1&&words.length>=2)?best:null;
  }catch(e){return null;}
}
function getDueBalanceForParty(party){
  var ln=resolveLedgerName_(party); if(!ln)return 0;
  var xml='<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>PushLedBal</ID></HEADER>'+
    '<BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>'+
    '<SVFROMDATE TYPE="Date">20250401</SVFROMDATE><SVTODATE TYPE="Date">20270331</SVTODATE>'+
    '</STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME="PushLedBal" ISMODIFY="No" ISFIXED="No" ISINITIALIZE="Yes" ISOPTION="No" ISINTERNAL="No">'+
    '<TYPE>Ledger</TYPE><FETCH>Name, ClosingBalance</FETCH><BELONGSTO>Sundry Debtors</BELONGSTO>'+
    '</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>';
  try{
    var t=UrlFetchApp.fetch(CONFIG.NGROK_URL,{method:'POST',contentType:'text/xml',payload:xml,
      headers:{'ngrok-skip-browser-warning':'1'},muteHttpExceptions:true}).getContentText();
    var ls=t.match(/<LEDGER [\s\S]*?<\/LEDGER>/g)||[];
    var tl=ln.toLowerCase().trim();
    for(var i=0;i<ls.length;i++){
      var nm=ls[i].match(/NAME="([^"]*)"/); if(!nm||nm[1].toLowerCase().trim()!==tl)continue;
      var bm=ls[i].match(/<CLOSINGBALANCE[^>]*>(.*?)<\/CLOSINGBALANCE>/); if(!bm)return 0;
      var amt=parseFloat(bm[1].trim().replace(/[^0-9.\-]/g,''));
      return isNaN(amt)?0:amt<0?Math.round(Math.abs(amt)):0;
    }
    return 0;
  }catch(e){return null;}
}
function getLastOrdersForParty(party){
  var ln=resolveLedgerName_(party)||party;
  var sw=pickLedgerFilterWord_(ln).replace(/"/g,''); if(!sw)return[];
  var xml='<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>PushaanSalesVch</ID></HEADER>'+
    '<BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>'+
    '<SVFROMDATE TYPE="Date">20250401</SVFROMDATE><SVTODATE TYPE="Date">20270331</SVTODATE>'+
    '</STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME="PushaanSalesVch" ISMODIFY="No" ISFIXED="No" ISINITIALIZE="Yes" ISOPTION="No" ISINTERNAL="No">'+
    '<TYPE>Voucher</TYPE><CHILDOF>$$VchTypeSales</CHILDOF><BELONGSTO>Yes</BELONGSTO>'+
    '<FETCH>Date, PartyLedgerName, VoucherTypeName, AllInventoryEntries</FETCH><FILTER>PartyFilter</FILTER></COLLECTION>'+
    '<SYSTEM TYPE="Formulae" NAME="PartyFilter">$PartyLedgerName CONTAINS "'+sw+'"</SYSTEM>'+
    '</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>';
  try{
    var t=UrlFetchApp.fetch(CONFIG.NGROK_URL,{method:'POST',contentType:'text/xml',payload:xml,
      headers:{'ngrok-skip-browser-warning':'1'},muteHttpExceptions:true}).getContentText();
    var vs=[],vs2=0;
    while(true){var si=t.indexOf('<VOUCHER ',vs2);if(si===-1)break;var ei=t.indexOf('</VOUCHER>',si);if(ei===-1)break;vs.push(t.substring(si,ei+10));vs2=ei+10;}
    var tn=normalizeTallyName_(ln),matched=[];
    vs.forEach(function(v){
      var vt=v.match(/VCHTYPE="([^"]*)"/); if(!vt||vt[1].indexOf('Sales')!==0)return;
      var pm=v.match(/<PARTYLEDGERNAME[^>]*>(.*?)<\/PARTYLEDGERNAME>/); if(!pm)return;
      var pn=normalizeTallyName_(pm[1]);
      if(pn!==tn&&pn.indexOf(tn)===-1&&tn.indexOf(pn)===-1)return;
      var dm=v.match(/<DATE[^>]*>(.*?)<\/DATE>/);
      var its=v.match(/<ALLINVENTORYENTRIES\.LIST>[\s\S]*?<\/ALLINVENTORYENTRIES\.LIST>/g)||[];
      var li=[];
      its.forEach(function(it){
        var nM=it.match(/<STOCKITEMNAME[^>]*>(.*?)<\/STOCKITEMNAME>/);
        var qM=it.match(/<ACTUALQTY[^>]*>(.*?)<\/ACTUALQTY>/)||it.match(/<BILLEDQTY[^>]*>(.*?)<\/BILLEDQTY>/);
        var rM=it.match(/<RATE[^>]*>(.*?)<\/RATE>/);
        if(!nM)return;
        var rd=''; if(rM){var rv=parseFloat(rM[1].replace(/[^0-9.]/g,''));rd=rv>0?(Math.round(rv*1.05*10)/10).toString():rM[1].trim();}
        li.push({item:nM[1].trim(),qty:qM?qM[1].trim():'',rate:rd});
      });
      if(li.length>0)matched.push({rawDate:dm?dm[1].trim():'',items:li});
    });
    matched.sort(function(a,b){return b.rawDate.localeCompare(a.rawDate);});
    Logger.log('DBG_HIST ' + JSON.stringify(matched.map(function(h){return h.items.map(function(i){return i.item;});})));
    return matched.slice(0,3);
  }catch(e){return[];}
}
function getDiscountForParty_(party){
  var ln=resolveLedgerName_(party)||party;
  var sw=pickLedgerFilterWord_(ln).replace(/"/g,'');
  var xml='<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>PushaanDiscVch</ID></HEADER>'+
    '<BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>'+
    '<SVFROMDATE TYPE="Date">20250401</SVFROMDATE><SVTODATE TYPE="Date">20270331</SVTODATE>'+
    '</STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME="PushaanDiscVch" ISMODIFY="No" ISFIXED="No" ISINITIALIZE="Yes" ISOPTION="No" ISINTERNAL="No">'+
    '<TYPE>Voucher</TYPE><CHILDOF>$$VchTypeSales</CHILDOF><BELONGSTO>Yes</BELONGSTO>'+
    '<FETCH>Date, PartyLedgerName, VoucherTypeName, AllInventoryEntries, LedgerEntries</FETCH><FILTER>PushaanDiscFilter</FILTER></COLLECTION>'+
    '<SYSTEM TYPE="Formulae" NAME="PushaanDiscFilter">$PartyLedgerName CONTAINS "'+sw+'"</SYSTEM>'+
    '</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>';
  try{
    var t=UrlFetchApp.fetch(CONFIG.NGROK_URL,{method:'POST',contentType:'text/xml',payload:xml,
      headers:{'ngrok-skip-browser-warning':'1'},muteHttpExceptions:true}).getContentText();
    var found=[];
    var tn=normalizeTallyName_(ln);
    var vlist=[],vp=0;
    while(true){var si=t.indexOf('<VOUCHER ',vp);if(si===-1)break;var ei=t.indexOf('</VOUCHER>',si);if(ei===-1)break;vlist.push(t.substring(si,ei+10));vp=ei+10;}
    vlist.forEach(function(vb){
      var pm=vb.match(/<PARTYLEDGERNAME[^>]*>(.*?)<\/PARTYLEDGERNAME>/); if(!pm)return;
      var pn=normalizeTallyName_(pm[1]);
      if(pn!==tn&&pn.indexOf(tn)===-1&&tn.indexOf(pn)===-1)return;
      if(vb.indexOf('Trade Discount on Sales')===-1)return;
      var at=vb.substring(vb.indexOf('Trade Discount on Sales'));
      var am=at.match(/<AMOUNT>(.*?)<\/AMOUNT>/); if(!am)return;
      var da=Math.abs(parseFloat(am[1])); if(da===0)return;
      var tk=0;
      (vb.match(/<ALLINVENTORYENTRIES\.LIST>[\s\S]*?<\/ALLINVENTORYENTRIES\.LIST>/g)||[]).forEach(function(ie){
        var qm=ie.match(/<ACTUALQTY[^>]*>\s*([\d.]+)\s*kgs/); if(qm)tk+=parseFloat(qm[1]);
      });
      if(tk===0)return;
      var dm=vb.match(/<DATE[^>]*>(.*?)<\/DATE>/);
      found.push({date:dm?dm[1].trim():'',discount:Math.round((da/tk)*100)/100});
    });
    if(!found.length)return null;
    found.sort(function(a,b){return b.date.localeCompare(a.date);});
    return { discount: found[0].discount, date: found[0].date };
  }catch(e){return 0;}
}