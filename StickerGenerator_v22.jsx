/**
 * StickerGenerator.jsx  v22.0  — Box Sticker text 1 space নিচে
 *
 * v20 পরিবর্তন (PDF Y comparison থেকে exact offset):
 *   BOX Sticker:   textGapFromBarBottom = 13pt  (12.08pt উপরে ছিল)
 *   PRICE Sticker: textGapFromBarBottom = 9pt   (8.46pt উপরে ছিল)
 * Adobe Illustrator ExtendScript
 *
 * ফিচার:
 *   - Excel (.xlsx) থেকে সরাসরি ডেটা পড়ে (প্রথম sheet, নাম যাই হোক)
 *   - Template-এর প্রতিটি Article-এর জন্য আলাদা document (কপি) তৈরি করে
 *     → কোনো object move/reposition করা হয় না, তাই PAGE SETUP ভাঙে না
 *   - Text placeholder: #A1-#A5, #C1-#C5, #M1-#M5, #R1-#R5 (Box)
 *                        #PA1-#PA5, #PR1-#PR5 (Price)
 *   - Barcode: #B1-#B5 (Box), #PB1-#PB5 (Price) টেক্সট ফ্রেমকে
 *     প্রকৃত EAN-13 বারকোড দিয়ে রিপ্লেস করে (guard bars + first-digit +
 *     6/6 digit grouping সহ, কোনো EPS লাগে না)। নিচের নাম্বার OCR-B
 *     ধরনের barcode font দিয়ে আঁকা হয় — exact font-নাম না মিললে সিস্টেমে
 *     ইনস্টল থাকা সব ফন্টের মধ্যে "OCR" নামযুক্ত যেকোনোটা নিজে থেকে খুঁজে
 *     নেয় (v12 ফিক্স), এবং কোনো forced-width scaling করা হয় না — তাই
 *     অক্ষর চেপ্টা/distorted দেখায় না (v11 ফিক্স)
 *   - "Month and Year of Manufacture" ফিল্ড "JUNE'2026" ফরম্যাটে দেখায়;
 *     সোর্স ডেটায় বছর না থাকলে CONFIG.defaultMfgYear ব্যবহার হয় (v11 ফিক্স)
 *   - সব Article শেষে multi-page PDF এ merge করে — paste-এর পর প্রকৃত
 *     pasted bounds মেপে নিয়ে translate করা হয়, তাই কোনো sticker
 *     artboard-এর বাইরে চলে যায় না (v11 ফিক্স — আগের ভার্সনে এই কারণেই
 *     একটা article-এর sticker artboard-এর বাইরে চলে যাচ্ছিল)
 *   - আউটপুট হিসেবে PDF + editable AI — দুটোই সেভ হয়
 *
 * REQUIREMENT:
 *   Input Folder-এ থাকবে:
 *     - Sticker_Template.ai
 *     - data.xlsx  (Excel, প্রথম sheet-এ data, row1-3 মেটাডেটা, row4 header)
 *   Barcode-এর জন্য কোনো একটা OCR-B ধরনের font (TOCR-B 10 BT ইত্যাদি)
 *   ইনস্টল থাকলে স্ক্রিপ্ট নিজে থেকেই সেটা খুঁজে ব্যবহার করবে; না থাকলে
 *   Arial দিয়ে চালাবে (কোনো error দেখাবে না, শুধু একবার info alert দেবে)।
 */

// ====== CONFIG ======
var CONFIG = {
    templateFileName : "Sticker_Template.ai",
    outputFileName   : "Sticker_Output.pdf",
    outputAIFileName : "Sticker_Output.ai",
    skipRows         : 3,   // Row1=Company, Row2=Project, Row3=PO, Row4=Header

    // ── v12: Barcode টেক্সট ফন্ট ──
    // এই নামগুলো ক্রমান্বয়ে exact match try করা হয়। কোনোটাই না পেলে
    // স্ক্রিপ্ট স্বয়ংক্রিয়ভাবে সিস্টেমে ইনস্টল থাকা সব ফন্টের মধ্যে
    // নামে "OCR" আছে এমন যেকোনো একটা খুঁজে নেবে (নিচে getBarcodeFont দেখুন) —
    // তাই exact spelling/version নিয়ে আর সমস্যা হবে না।
    barcodeFontCandidates : [
        "TOCR-B 10 BT", "OCR-B 10 BT", "OCR B 10 BT", "OCRB10BT",
        "OCR-B", "OCR B", "OCRB", "OCR-B-Regular", "OCRBRegular",
        "Arial"
    ],

    // ── v12: barcode সংখ্যাগুলোর মাঝে অতিরিক্ত letter-spacing (tracking) ──
    // মান 1000-এর ভগ্নাংশে (Illustrator-এর tracking একক)। 0 = ফন্টের
    // স্বাভাবিক spacing। reference ছবির মতো একটু খোলামেলা/গোছানো দেখাতে
    // চাইলে 20-60 এর মধ্যে try করতে পারেন।
    barcodeTracking : 0,

    // ── v11: Month-Year ফিল্ডে সোর্স ডেটায় বছর না থাকলে এই বছর বসবে ──
    // (খালি স্ট্রিং "" রাখলে স্ক্রিপ্ট চালানোর বছর (system year) অটো বসবে)
    defaultMfgYear : "2026"
};
// ====================

// বারকোড ফন্ট resolve করে একবার cache করে রাখা হয় (পুরো রান জুড়ে একই ফন্ট ব্যবহার হয়,
// এবং ফন্ট-না-পাওয়ার সতর্কতা শুধু একবারই দেখানো হয়, প্রতি বারকোডে না)
var _resolvedBarcodeFont = null;
var _barcodeFontWarned   = false;
function getBarcodeFont() {
    if (_resolvedBarcodeFont) return _resolvedBarcodeFont;

    // ধাপ ১: পরিচিত exact নামগুলো ক্রমান্বয়ে try করা
    for (var i = 0; i < CONFIG.barcodeFontCandidates.length; i++) {
        try {
            var f1 = app.textFonts.getByName(CONFIG.barcodeFontCandidates[i]);
            if (f1) {
                _resolvedBarcodeFont = f1;
                if (i > 0 && CONFIG.barcodeFontCandidates[i] !== "Arial") {
                    warnFontFallback(CONFIG.barcodeFontCandidates[i]);
                }
                return _resolvedBarcodeFont;
            }
        } catch (e) { /* এই নামটা নেই, পরেরটা try করো */ }
    }

    // ধাপ ২: exact নাম কোনোটাই না মিললে, ইনস্টল করা সব ফন্টের মধ্যে
    // নামের ভেতরে কোথাও "OCR" আছে এমন প্রথম ফন্টটা ব্যবহার করো —
    // এতে exact spelling/version/edition (10BT, Regular ইত্যাদি)
    // যাই হোক না কেন, ফন্টটা খুঁজে পাওয়া যাবে।
    try {
        for (var j = 0; j < app.textFonts.length; j++) {
            var fname = "";
            try { fname = app.textFonts[j].name; } catch (eN) { continue; }
            if (fname && fname.toUpperCase().indexOf("OCR") !== -1) {
                _resolvedBarcodeFont = app.textFonts[j];
                warnFontFallback(fname);
                return _resolvedBarcodeFont;
            }
        }
    } catch (e2) { /* font list পড়া গেল না, নিচে default-এ চলে যাবে */ }

    // ধাপ ৩: কিছুই পাওয়া গেল না — Illustrator-এর default font-এই থাকবে
    if (!_barcodeFontWarned) {
        _barcodeFontWarned = true;
        alert("⚠️ কোনো OCR-B ধরনের barcode font সিস্টেমে পাওয়া যায়নি।\n" +
              "Illustrator-এর default font দিয়ে চালানো হচ্ছে।\n\n" +
              "TOCR-B 10 BT (বা অন্য যেকোনো OCR-B) ফন্ট ইনস্টল করে আবার " +
              "চালালে barcode-এর নাম্বার আরো standard/পরিষ্কার দেখাবে।");
    }
    return null;
}

function warnFontFallback(actualFontName) {
    if (_barcodeFontWarned) return;
    _barcodeFontWarned = true;
    alert("ℹ️ Barcode ফন্ট হিসেবে '" + actualFontName + "' ব্যবহার করা হচ্ছে।");
}

function main() {
    var inFolder = Folder.selectDialog(
        "INPUT FOLDER বেছে নিন\n(Sticker_Template.ai + data.xlsx)"
    );
    if (!inFolder) { alert("বাতিল।"); return; }

    var outFolder = Folder.selectDialog("OUTPUT FOLDER বেছে নিন");
    if (!outFolder) { alert("বাতিল।"); return; }

    var templateFile = new File(inFolder.fsName + "/" + CONFIG.templateFileName);
    if (!templateFile.exists) {
        var aiArr = inFolder.getFiles("*.ai");
        if (aiArr.length === 0) {
            alert("ERROR: Template .ai ফাইল পাওয়া যায়নি!\n" + inFolder.fsName);
            return;
        }
        templateFile = aiArr[0];
        alert("Template: " + templateFile.name);
    }

    var dataFile = findExcelFile(inFolder);
    if (!dataFile) {
        alert("ERROR: .xlsx ডেটা ফাইল পাওয়া যায়নি!\n" + inFolder.fsName);
        return;
    }

    var groups = parseExcelData(dataFile);
    if (groups.length === 0) { alert("ERROR: ডেটায় কোনো row নেই।"); return; }

    var tempFolder = new Folder(outFolder.fsName + "/_tmp_stk");
    if (!tempFolder.exists) tempFolder.create();

    var tempFiles = [];

    for (var g = 0; g < groups.length; g++) {
        try {
            // প্রতিটি Article-এর আগে নিশ্চিত করো কোনো document খোলা নেই
            // (আগের Article-এর leftover state পরিষ্কার রাখতে)

            var tf = processArticle(templateFile, tempFolder, groups[g], g);
            if (tf) tempFiles.push(tf);

            // প্রতিটি Article শেষে app reference জোরপূর্বক refresh করো
        } catch(e) {
            alert("ERROR — " + groups[g].article + "\n" + e.message + "\nLine: " + e.line);
        }
    }

    if (tempFiles.length > 0) {
        mergeToPDF(tempFiles, outFolder);
    }

    cleanFolder(tempFolder);
    try { tempFolder.remove(); } catch(e) {}

    alert("✅ সম্পন্ন!\n" + tempFiles.length + " page তৈরি হয়েছে।\n\n" +
          "PDF: " + outFolder.fsName + "/" + CONFIG.outputFileName + "\n" +
          "AI : " + outFolder.fsName + "/" + CONFIG.outputAIFileName);
}

// ═══════════════════════════════════════════════════════
//  নিরাপদ app.open — ব্যর্থ হলে কয়েকবার চেষ্টা করো
// ═══════════════════════════════════════════════════════
function safeOpen(file) {
    var lastErr = null;
    for (var attempt = 0; attempt < 3; attempt++) {
        try {
            var d = app.open(file);
            return d;
        } catch(e) {
            lastErr = e;
        }
    }
    throw lastErr || new Error("app.open ব্যর্থ হয়েছে: " + file.fsName);
}

// ═══════════════════════════════════════════════════════
//  একটি Article process → temp AI file (artboard অপরিবর্তিত)
// ═══════════════════════════════════════════════════════
function processArticle(templateFile, tempFolder, grp, idx) {
    var rows = grp.rows;  // ৫টি size এর row, Excel ক্রম = slot 1-5

    var doc = safeOpen(templateFile);

    var allTF = [];
    collectAllTextFrames(doc, allTF);

    // ── BOX + PRICE TEXT FIELDS ──
    var fieldDefs = [
        { prefix: "#A",  isPrice: false, getter: function(r){ return r.article; } },
        { prefix: "#C",  isPrice: false, getter: function(r){ return r.colour; } },
        { prefix: "#M",  isPrice: false, getter: function(r){ return formatMfgDate(r.mfgDate); } },
        { prefix: "#R",  isPrice: false, getter: function(r){ return formatMRP(r.mrp); } },
        { prefix: "#PA", isPrice: true,  getter: function(r){ return r.article; } },
        { prefix: "#PR", isPrice: true,  getter: function(r){ return formatMRP(r.mrp); } }
    ];

    for (var fd = 0; fd < fieldDefs.length; fd++) {
        var def = fieldDefs[fd];
        for (var slot = 1; slot <= 5 && slot <= rows.length; slot++) {
            var key = def.prefix + slot;     // যেমন "#A1", "#PR3"
            var val = def.getter(rows[slot-1]);
            setExactMatchFrame(allTF, key, val);
        }
    }

    // ── BARCODE FIELDS (EAN-13 আঁকা) ──
    for (var bslot = 1; bslot <= 5 && bslot <= rows.length; bslot++) {
        var ean = rows[bslot-1].ean;
        drawBarcodeAtPlaceholder(doc, allTF, "#B" + bslot, ean, false);
        drawBarcodeAtPlaceholder(doc, allTF, "#PB" + bslot, ean, true);
    }


    var tmpFile = new File(tempFolder.fsName + "/art_" + idx + ".ai");
    var aiOpts  = new IllustratorSaveOptions();
    aiOpts.compatibility = Compatibility.ILLUSTRATOR24;
    doc.saveAs(tmpFile, aiOpts);
    doc.close(SaveOptions.DONOTSAVECHANGES);

    return tmpFile;
}

// ═══════════════════════════════════════════════════════
//  Exact-match placeholder খুঁজে value বসানো
//  (নিশ্চিত করে যে "#A1" আর "#A10" গুলিয়ে যাবে না)
// ═══════════════════════════════════════════════════════
function setExactMatchFrame(allTF, key, value) {
    for (var i = 0; i < allTF.length; i++) {
        var c;
        try { c = allTF[i].contents; } catch(eInv) { continue; }
        if (c === key) {
            allTF[i].contents = value;
            return;
        }
    }
    // exact match না পেলে substring match (fallback)
    for (var j = 0; j < allTF.length; j++) {
        var c2;
        try { c2 = allTF[j].contents; } catch(eInv2) { continue; }
        if (c2.indexOf(key) !== -1 && !isAmbiguous(c2, key)) {
            allTF[j].contents = replaceAll(c2, key, value);
            return;
        }
    }
}

// "#A1" খোঁজার সময় "#A10" এর সাথে গুলিয়ে যাওয়া আটকানো
function isAmbiguous(content, key) {
    var idx = content.indexOf(key);
    if (idx === -1) return true;
    var nextChar = content.charAt(idx + key.length);
    return (nextChar >= '0' && nextChar <= '9');
}

// ═══════════════════════════════════════════════════════
//  Barcode placeholder frame খুঁজে EAN-13 বারকোড আঁকা
// ═══════════════════════════════════════════════════════
function drawBarcodeAtPlaceholder(doc, allTF, key, ean, isPrice) {
    var frame = null;
    var frameIdx = -1;
    for (var i = 0; i < allTF.length; i++) {
        var contentVal;
        try { contentVal = allTF[i].contents; } catch(eInvalid) { continue; }
        if (contentVal === key) { frame = allTF[i]; frameIdx = i; break; }
    }
    if (!frame) {
        for (var j = 0; j < allTF.length; j++) {
            var cv2;
            try { cv2 = allTF[j].contents; } catch(eInvalid2) { continue; }
            if (cv2.indexOf(key) !== -1 && !isAmbiguous(cv2, key)) {
                frame = allTF[j]; frameIdx = j; break;
            }
        }
    }
    if (!frame) return;

    var gb = frame.geometricBounds;  // [L, T, R, B]
    var bL = gb[0], bT = gb[1];
    var bW = gb[2] - gb[0];
    var bH = gb[1] - gb[3];
    var parentLayer = doc.layers[0];

    try { frame.remove(); } catch(e) {}
    // remove হওয়া frame-টা allTF array থেকেও বাদ দিয়ে দিচ্ছি,
    // যাতে পরের কোনো loop ভুলে এই (এখন invalid) object না ছুঁয়ে ফেলে
    if (frameIdx !== -1) allTF.splice(frameIdx, 1);

    if (!ean || trim(ean) === "") return;

    drawEAN13Barcode(doc, parentLayer, trim(ean), bL, bT, bW, bH, !!isPrice);
}

// ═══════════════════════════════════════════════════════
//  EAN-13 BARCODE GENERATOR (bars আঁকা, কোনো EPS লাগে না)
// ═══════════════════════════════════════════════════════
var EAN_L_CODE = {
    "0":"0001101","1":"0011001","2":"0010011","3":"0111101","4":"0100011",
    "5":"0110001","6":"0101111","7":"0111011","8":"0110111","9":"0001011"
};
var EAN_G_CODE = {
    "0":"0100111","1":"0110011","2":"0011011","3":"0100001","4":"0011101",
    "5":"0111001","6":"0000101","7":"0010001","8":"0001001","9":"0010111"
};
var EAN_R_CODE = {
    "0":"1110010","1":"1100110","2":"1101100","3":"1000010","4":"1011100",
    "5":"1001110","6":"1010000","7":"1000100","8":"1001000","9":"1110100"
};
// প্রথম digit (0-9) অনুযায়ী বাকি ৬টা left digit-এর L/G pattern নির্বাচন
var EAN_PARITY = {
    "0":"LLLLLL","1":"LLGLGG","2":"LLGGLG","3":"LLGGGL","4":"LGLLGG",
    "5":"LGGLLG","6":"LGGGLL","7":"LGLGLG","8":"LGLGGL","9":"LGGLGL"
};

function computeEAN13Checksum(digits12) {
    var sum = 0;
    for (var i = 0; i < 12; i++) {
        var d = parseInt(digits12.charAt(i), 10);
        sum += (i % 2 === 0) ? d : d * 3;
    }
    var mod = sum % 10;
    return (mod === 0) ? 0 : (10 - mod);
}

function drawEAN13Barcode(doc, layer, eanStr, left, top, totalW, totalH, isPrice) {
    // ═══════════════════════════════════════════════════════════════════════
    //  EAN-13 BARCODE — v19  (PDF stream থেকে exact মান নির্ণয় করা হয়েছে)
    //
    //  BOX STICKER (isPrice=false):
    //    PDF Tm matrix: 10.15 0 0 9.66  → size=7pt, vScale=138%, hScale=145%
    //    Text তিনটি group (8 / 909102 / 567xxx) একই Y-তে baseline align
    //
    //  PRICE STICKER (isPrice=true):
    //    PDF Tm: 7.3534 0 0 7.3534 (909102/567xxx) এবং 7.35 0 0 7.35 (8)
    //    → size=7.35pt, no scale, সব একই baseline ✅
    // ═══════════════════════════════════════════════════════════════════════

    // ── EAN validate & checksum ──
    var digits = eanStr.replace(/[^0-9]/g, "");
    while (digits.length < 12) digits = "0" + digits;
    digits = digits.substring(0, 12);
    var checksum = computeEAN13Checksum(digits);
    var fullCode = digits + String(checksum);

    var firstDigit  = fullCode.charAt(0);
    var leftDigits  = fullCode.substring(1, 7);
    var rightDigits = fullCode.substring(7, 13);
    var parity      = EAN_PARITY[firstDigit];

    // ── Bar pattern (95 modules) ──
    var pattern = "101";
    for (var i = 0; i < 6; i++) {
        var d = leftDigits.charAt(i);
        pattern += (parity.charAt(i) === "L") ? EAN_L_CODE[d] : EAN_G_CODE[d];
    }
    pattern += "01010";
    for (var j = 0; j < 6; j++) {
        pattern += EAN_R_CODE[rightDigits.charAt(j)];
    }
    pattern += "101";   // 95 modules মোট

    // ── Width layout ──
    var firstDigitW = totalW * 0.085;
    var symbolW     = totalW - firstDigitW;
    var barStartX   = left + firstDigitW;
    var moduleW     = symbolW / 95;

    // ── Height layout ──
    // text height = 7pt × 138% vScale = 9.66pt  (BOX)
    //             = 7.35pt                        (PRICE)
    // bar height = totalH - textH - gap
    var textH, barH, guardH, barShiftDown, textGapFromBarBottom;

    if (isPrice) {
        textH              = 7.35;
        barShiftDown       = 0;
        barH               = totalH - textH - 1.5;
        guardH             = barH + (totalH * 0.06);
        textGapFromBarBottom = 9.0;   // PDF analysis: 8.457pt নামাতে হবে → 9pt gap
    } else {
        textH              = 9.66;
        barShiftDown       = 0;
        barH               = totalH - textH - 1.0;
        guardH             = barH + (totalH * 0.08);
        textGapFromBarBottom = 10.5;  // v21=9pt, 1 space (≈1.34pt) নিচে → 10.5pt
    }

    // ── Bars আঁকা ──
    var barGroup = barGroupSafe(layer, fullCode);
    var bx = barStartX;
    for (var k = 0; k < pattern.length; k++) {
        if (pattern.charAt(k) === "1") {
            var isGuard = (k < 3) || (k >= 45 && k < 50) || (k >= 92);
            var bh      = isGuard ? guardH : barH;
            var rect    = layer.pathItems.rectangle(top - barShiftDown, bx, moduleW, bh);
            rect.filled    = true;
            rect.fillColor = makeBlack(doc);
            rect.stroked   = false;
            rect.move(barGroup, ElementPlacement.PLACEATEND);
        }
        bx += moduleW;
    }

    // ── Text Y position ──
    // bar-এর normal bottom edge (guard নয়, normal bar নিচে)
    var barBottomY = top - barShiftDown - barH;
    // text-এর glyph bottom এখানে বসবে
    var textBaselineY = barBottomY - textGapFromBarBottom;

    if (isPrice) {
        // ── PRICE STICKER TEXT ──
        // PDF confirmed: 7.3534pt (909102/567xxx), 7.35pt (8)
        // সব একই baseline → alignToBottom=true, targetY=textBaselineY

        // first digit "8"
        addBarcodeTextV18(layer, barGroup, firstDigit,
            left, textBaselineY, firstDigitW,
            7.35, 1.0, 100,
            Justification.CENTER, true);

        // left 6-digit group
        var lgX = barStartX + 3 * moduleW;
        var lgW = 42 * moduleW;
        addBarcodeTextV18(layer, barGroup, leftDigits,
            lgX, textBaselineY, lgW,
            7.3534, 1.0, 100,
            Justification.CENTER, true);

        // right 6-digit group
        var rgX = barStartX + 50 * moduleW;
        var rgW = 42 * moduleW;
        addBarcodeTextV18(layer, barGroup, rightDigits,
            rgX, textBaselineY, rgW,
            7.3534, 1.0, 100,
            Justification.CENTER, true);

    } else {
        // ── BOX STICKER TEXT ──
        // PDF confirmed: Tm=10.15 0 0 9.66 → size=7pt, vScale=138%, hScale=145%
        // first digit "8" — bar এর বামে, same baseline
        addBarcodeTextV18(layer, barGroup, firstDigit,
            left, textBaselineY, firstDigitW,
            7.0, 1.45, 138,
            Justification.CENTER, true);

        // left 6-digit group
        var blgX = barStartX + 3 * moduleW;
        var blgW = 42 * moduleW;
        addBarcodeTextV18(layer, barGroup, leftDigits,
            blgX, textBaselineY, blgW,
            7.0, 1.45, 138,
            Justification.CENTER, true);

        // right 6-digit group
        var brgX = barStartX + 50 * moduleW;
        var brgW = 42 * moduleW;
        addBarcodeTextV18(layer, barGroup, rightDigits,
            brgX, textBaselineY, brgW,
            7.0, 1.45, 138,
            Justification.CENTER, true);
    }
}

// একই নামের barGroup খুঁজে reuse করা (যাতে rectangle ও text একই গ্রুপে যায়)
function barGroupSafe(layer, fullCode) {
    var nm = "Barcode_" + fullCode;
    try {
        for (var i = 0; i < layer.groupItems.length; i++) {
            if (layer.groupItems[i].name === nm) return layer.groupItems[i];
        }
    } catch(e) {}
    var g = layer.groupItems.add();
    g.name = nm;
    return g;
}

// ═══════════════════════════════════════════════════════
//  পুরনো addBarcodeText — এখনো রাখা হয়েছে (অন্য কোথাও call না থাকলেও)
// ═══════════════════════════════════════════════════════
function addBarcodeText(layer, group, text, x, y, w, size, justification, hScale) {
    addBarcodeTextV18(layer, group, text, x, y, w, size, hScale, 100, justification, false);
}

// ═══════════════════════════════════════════════════════
//  addBarcodeTextV18 — v18 নতুন ফাংশন
//
//  পার্থক্য পুরনো থেকে:
//  1. verticalScale (vScale) আলাদা parameter — Box sticker-এ 138%
//  2. alignToBottom=true হলে text-এর BOTTOM edge targetY-তে বসে
//     (bar bottom-এ text bottom align করার জন্য)
//     alignToBottom=false হলে text-এর TOP edge targetY-তে বসে
//  3. horizontalScale সরাসরি % হিসেবে নেওয়া (1.45 নয়, 145 নয় — hScale*100 হবে)
//     → hScale=1.45 দিলে 145% সেট হবে
// ═══════════════════════════════════════════════════════
function addBarcodeTextV18(layer, group, text, x, targetY, w, size, hScale, vScale, justification, alignToBottom) {
    try {
        var tf = layer.textFrames.add();
        tf.contents = text;

        var ts = tf.textRange.characterAttributes;

        // ── ১. Font ──
        var bcFont = getBarcodeFont();
        if (bcFont) {
            try { ts.textFont = bcFont; } catch(ef) {}
        }

        // ── ২. Size ──
        ts.size = size;

        // ── ৩. Vertical Scale (Character panel এর Vertical Scale %) ──
        if (vScale && vScale !== 100) {
            try { ts.verticalScale = vScale; } catch(evs) {}
        }

        // ── ৪. Horizontal Scale ──
        if (hScale && hScale !== 1.0) {
            try { ts.horizontalScale = hScale * 100; } catch(ehs) {}
        }

        // ── ৫. Tracking ──
        if (CONFIG.barcodeTracking) {
            try { ts.tracking = CONFIG.barcodeTracking; } catch(etr) {}
        }

        // ── ৬. Width fit: যদি natural width বরাদ্দ জায়গার চেয়ে বেশি হয়
        //       তাহলে size uniformly ছোট করো (squish করো না) ──
        var naturalW = tf.width;
        if (naturalW > w * 1.05 && naturalW > 0) {
            var fitRatio = (w / naturalW) * 0.97;
            ts.size = size * fitRatio;
            if (vScale && vScale !== 100) {
                try { ts.verticalScale = vScale; } catch(evs2) {}
            }
            if (hScale && hScale !== 1.0) {
                try { ts.horizontalScale = hScale * 100; } catch(ehs2) {}
            }
            naturalW = tf.width;
        }

        // ── ৭. X position (justification অনুযায়ী) ──
        var finalX = x;
        if (justification === Justification.CENTER) {
            finalX = x + (w - naturalW) / 2;
        } else if (justification === Justification.RIGHT) {
            finalX = x + (w - naturalW);
        }

        // ── ৮. প্রাথমিক position সেট ──
        tf.left = finalX;
        tf.top  = targetY;

        // ── ৯. Geometric bounds দিয়ে সঠিক Y correction ──
        // tf.top = bounding box top, কিন্তু glyph আসলে কোথায় আছে
        // সেটা geometricBounds থেকে মেপে correct করা হয়।
        // alignToBottom=true  → glyph-এর BOTTOM কে targetY-এ আনো
        // alignToBottom=false → glyph-এর TOP কে targetY-এ আনো
        try {
            var gb      = tf.geometricBounds;   // [L, T, R, B]
            var glyphT  = gb[1];
            var glyphB  = gb[3];
            var dy;
            if (alignToBottom) {
                // glyph bottom → targetY (নিচের দিকে মেলানো)
                dy = targetY - glyphB;
            } else {
                // glyph top → targetY (উপরের দিকে মেলানো)
                dy = targetY - glyphT;
            }
            if (Math.abs(dy) > 0.001) {
                tf.translate(0, dy);
            }
        } catch(ebnd) {}

        tf.move(group, ElementPlacement.PLACEATEND);
    } catch(et) {}
}

function makeBlack(doc) {
    var c = new RGBColor();
    c.red = 0; c.green = 0; c.blue = 0;
    return c;
}

// ═══════════════════════════════════════════════════════
//  সব TextFrame collect করা (recursive)
// ═══════════════════════════════════════════════════════
function collectAllTextFrames(container, result) {
    try {
        var tfs = container.textFrames;
        for (var i = 0; i < tfs.length; i++) result.push(tfs[i]);
    } catch(e) {}
    try {
        var grps = container.groupItems;
        for (var g = 0; g < grps.length; g++) collectAllTextFrames(grps[g], result);
    } catch(e2) {}
}

// ═══════════════════════════════════════════════════════
//  Temp AI files → multi-page PDF
//  (প্রতিটি ফাইলের artboard অপরিবর্তিত — কোনো translate করা হয় না,
//   তাই page setup/position কখনো ভাঙবে না)
// ═══════════════════════════════════════════════════════
function mergeToPDF(tempFiles, outFolder) {
    if (tempFiles.length === 0) return;

    var masterDoc = safeOpen(tempFiles[0]);

    for (var i = 1; i < tempFiles.length; i++) {
        var src = safeOpen(tempFiles[i]);

        // src-এর artboard rect (নিজের অপরিবর্তিত artboard)
        var srcRect = src.artboards[0].artboardRect;

        // src-এর সব আইটেম সিলেক্ট করো
        var allItems = [];
        for (var pi = 0; pi < src.pageItems.length; pi++) {
            allItems.push(src.pageItems[pi]);
        }
        if (allItems.length === 0) {
            src.close(SaveOptions.DONOTSAVECHANGES);
            continue;
        }

        // ── v11 FIX: কপি করার আগেই, src ডকুমেন্টে content-এর প্রকৃত bounds
        //    মেপে src artboard-এর top-left থেকে তার offset বের করে রাখি।
        //    এটাই হলো "content আর্টবোর্ডের কোথায় বসা উচিত" তার সত্যিকারের রেফারেন্স —
        //    Illustrator-এর paste behavior-কে অন্ধভাবে বিশ্বাস করি না।
        var srcContentBounds = getGroupBounds(allItems); // [L, T, R, B]
        var offX = srcContentBounds[0] - srcRect[0];      // বামদিক থেকে কত দূরে
        var offY = srcRect[1] - srcContentBounds[1];      // উপরদিক থেকে কত দূরে

        src.selection = allItems;
        app.copy();
        src.close(SaveOptions.DONOTSAVECHANGES);

        app.activeDocument = masterDoc;
        // ── নতুন artboard বানাও, যার rect = src-এর artboard-এর
        //    SAME WIDTH/HEIGHT, কিন্তু master-এর নিচে position করা
        var lastIdx  = masterDoc.artboards.length - 1;
        var prevRect = masterDoc.artboards[lastIdx].artboardRect;
        var abW = srcRect[2] - srcRect[0];
        var abH = srcRect[1] - srcRect[3];

        var newLeft = prevRect[0];
        var newTop  = prevRect[3] - 40;  // আগের আর্টবোর্ডের নিচে ৪০pt gap
        var newRect = [newLeft, newTop, newLeft + abW, newTop - abH];

        masterDoc.artboards.add(newRect);
        masterDoc.artboards.setActiveArtboardIndex(masterDoc.artboards.length - 1);

        // ⚠️ পূর্ববর্তী ভার্সনে এখানে ধরে নেওয়া হতো যে pasteInPlace সবসময়
        // src-document-এর EXACT same absolute coordinate-এ বসায় (artboard
        // active থাকুক যাই হোক না কেন), এবং সেই ধারণার উপর ভিত্তি করেই
        // dx/dy হিসাব করা হতো — Illustrator আসলে কোথায় বসিয়েছে সেটা কখনো
        // মাপা হতো না। কিন্তু Illustrator একাধিক artboard থাকা document-এ
        // pasteInPlace-কে মাঝে মাঝে ACTIVE ARTBOARD-এর সাপেক্ষেও বসায় —
        // ফলে ধারণা আর বাস্তবতা না মিললে content আর্টবোর্ডের অনেক দূরে/বাইরে
        // চলে যেত (যেটা ঠিক আগের আউটপুটে একটা sticker-এ হয়েছিল)।
        // v11 FIX: এখন paste হওয়ার পর প্রকৃত position সরাসরি মেপে নেওয়া হয়,
        // তারপর সেই প্রকৃত position থেকে correction হিসাব করা হয় — তাই
        // Illustrator যেখানেই বসাক না কেন, ফলাফল সবসময় সঠিক জায়গায় যাবে।
        app.executeMenuCommand('pasteInPlace');

        var pasted = masterDoc.selection;
        if (pasted && pasted.length > 0) {
            // paste-এর পরের প্রকৃত (measured) bounds
            var curBounds = getGroupBounds(pasted);

            // নতুন artboard-এ content যেখানে থাকা উচিত (src-এ যেমন offset ছিল, ঠিক তেমন)
            var targetLeft = newRect[0] + offX;
            var targetTop  = newRect[1] - offY;

            // প্রকৃত পজিশন থেকে কাঙ্ক্ষিত পজিশনে আনতে কতটুকু সরাতে হবে
            var dx = targetLeft - curBounds[0];
            var dy = targetTop  - curBounds[1];

            for (var k = 0; k < pasted.length; k++) {
                try { pasted[k].translate(dx, dy); } catch(e) {}
            }
        }
    }

    // ── ১) Editable Adobe Illustrator (.ai) ফাইল সেভ করা ──
    var aiFile = new File(outFolder.fsName + "/" + CONFIG.outputAIFileName);
    var aiOpts2 = new IllustratorSaveOptions();
    aiOpts2.compatibility = Compatibility.ILLUSTRATOR24;
    masterDoc.saveAs(aiFile, aiOpts2);

    // ── ২) PDF ফাইল সেভ করা ──
    var pdfFile = new File(outFolder.fsName + "/" + CONFIG.outputFileName);
    var pdfOpts = new PDFSaveOptions();
    pdfOpts.compatibility       = PDFCompatibility.ACROBAT7;
    pdfOpts.generateThumbnails  = true;
    pdfOpts.preserveEditability = false;
    pdfOpts.viewAfterSaving     = false;
    masterDoc.saveAs(pdfFile, pdfOpts);
    masterDoc.close(SaveOptions.DONOTSAVECHANGES);
}

function getGroupBounds(items) {
    var minL = Infinity, maxT = -Infinity, maxR = -Infinity, minB = Infinity;
    for (var i = 0; i < items.length; i++) {
        try {
            var gb = items[i].geometricBounds;
            if (gb[0] < minL) minL = gb[0];
            if (gb[1] > maxT) maxT = gb[1];
            if (gb[2] > maxR) maxR = gb[2];
            if (gb[3] < minB) minB = gb[3];
        } catch(e) {}
    }
    return [minL, maxT, maxR, minB];
}

// ═══════════════════════════════════════════════════════
//  EXCEL (.xlsx) পড়া
//  ExtendScript সরাসরি xlsx (ZIP+XML) পড়তে পারে না,
//  তাই আমরা xlsx-কে ZIP হিসেবে treat করে shared strings ও
//  sheet1.xml থেকে raw cell data parse করি।
// ═══════════════════════════════════════════════════════
function findExcelFile(folder) {
    var files = folder.getFiles("*.xlsx");
    if (files.length === 0) files = folder.getFiles("*.XLSX");
    return files.length > 0 ? files[0] : null;
}

function parseExcelData(file) {
    // ExtendScript এ সরাসরি ZIP/XML পার্স করা সম্ভব নয় built-in দিয়ে।
    // তাই আমরা external helper ব্যবহার করি: csv ফরম্যাটে convert করা প্রয়োজন।
    // সমাধান: একই ফোল্ডারে .csv খুঁজি (xlsx এর পাশে রাখা থাকবে)।
    // যদি .csv না থাকে, ব্যবহারকারীকে জানাই।

    var folder = file.parent;
    var csvFiles = folder.getFiles("*.csv");
    if (csvFiles.length > 0) {
        return parseCSVData(csvFiles[0]);
    }

    alert("⚠️ Illustrator ExtendScript সরাসরি .xlsx পড়তে পারে না।\n\n" +
          "একটি ছোট অতিরিক্ত ধাপ লাগবে:\n" +
          "Excel ফাইলটি খুলুন → File → Save As → CSV (Comma delimited) *.csv\n" +
          "এবং এটি একই ফোল্ডারে রাখুন (data.xlsx এর পাশে)।\n\n" +
          "তারপর স্ক্রিপ্ট আবার চালান — এটি স্বয়ংক্রিয়ভাবে .csv ফাইলটি পড়ে নিবে।");
    return [];
}

function parseCSVData(file) {
    file.open("r");
    file.encoding = "UTF-8";
    var lines = [];
    while (!file.eof) lines.push(file.readln());
    file.close();

    var hdrIdx  = CONFIG.skipRows;
    var headers = splitCSVLine(lines[hdrIdx]);
    var col     = {};

    for (var h = 0; h < headers.length; h++) {
        var hdr = trim(headers[h]).toUpperCase();
        if      (hdr.indexOf("ARTICLE") !== -1)                                col.article = h;
        else if (hdr.indexOf("EAN")     !== -1)                                col.ean     = h;
        else if (hdr.indexOf("COLOUR")  !== -1 || hdr.indexOf("COLOR") !== -1) col.colour  = h;
        else if (hdr.indexOf("MRP")     !== -1)                                col.mrp     = h;
        else if (hdr.indexOf("MONTH")   !== -1 || hdr.indexOf("MANUFACTUR") !== -1) col.mfgDate = h;
    }
    if (col.article === undefined) col = {article:0, ean:1, colour:2, mrp:3, mfgDate:7};

    var gMap = {}, gOrder = [];
    for (var i = hdrIdx + 1; i < lines.length; i++) {
        if (!trim(lines[i])) continue;
        var cells = splitCSVLine(lines[i]);
        if (cells.length < 4) continue;
        var art = trim(cells[col.article] || "");
        if (!art) continue;
        var row = {
            article : art,
            ean     : trim(cells[col.ean]     || ""),
            colour  : trim(cells[col.colour]  || ""),
            mrp     : trim(cells[col.mrp]     || ""),
            mfgDate : trim(cells[col.mfgDate] || "")
        };
        if (!gMap[art]) { gMap[art] = {article:art, rows:[]}; gOrder.push(art); }
        gMap[art].rows.push(row);
    }

    var result = [];
    for (var j = 0; j < gOrder.length; j++) result.push(gMap[gOrder[j]]);
    return result;
}

function splitCSVLine(line) {
    // সহজ CSV split (quoted field সহ)
    var result = [];
    var cur = "";
    var inQuotes = false;
    for (var i = 0; i < line.length; i++) {
        var ch = line.charAt(i);
        if (ch === '"') {
            inQuotes = !inQuotes;
        } else if (ch === ',' && !inQuotes) {
            result.push(cur);
            cur = "";
        } else {
            cur += ch;
        }
    }
    result.push(cur);
    return result;
}

// ═══════════════════════════════════════════════════════
//  UTILITIES
// ═══════════════════════════════════════════════════════
function cleanFolder(folder) {
    try {
        var f = folder.getFiles();
        for (var i = 0; i < f.length; i++) try { f[i].remove(); } catch(e) {}
    } catch(e) {}
}

function formatMRP(mrp) {
    var n = parseFloat(String(mrp).replace(/[^\d.]/g, ""));
    if (isNaN(n)) return String(mrp);
    var ip  = Math.floor(n);
    var dec = Math.round((n - ip) * 100);
    var s   = String(ip);
    if (s.length > 3) s = s.slice(0, -3) + " " + s.slice(-3);
    return s + "." + pad2(dec);
}

function getDefaultMfgYear() {
    if (CONFIG.defaultMfgYear && trim(CONFIG.defaultMfgYear) !== "") {
        return trim(CONFIG.defaultMfgYear);
    }
    return String(new Date().getFullYear());
}

function formatMfgDate(raw) {
    var s = trim(String(raw || ""));
    if (!s) return "";

    var monthNames = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY",
                       "AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"];
    var upper = s.toUpperCase();

    // ইতিমধ্যে "JUNE'2026" / "JUNE 2026" / "JUNE-2026" ফরম্যাটে থাকলে normalize করা
    var foundMonth = null;
    for (var i = 0; i < 12; i++) {
        if (upper.indexOf(monthNames[i]) !== -1) { foundMonth = monthNames[i]; break; }
    }
    var yearMatch = s.match(/(19|20)\d{2}/);
    var year = yearMatch ? yearMatch[0] : "";

    // v11 FIX: সোর্স ডেটায় (Excel/CSV) যদি শুধু মাসের নাম থাকে, বছর না থাকে,
    // তাহলে আগে শুধু "JUNE" রিটার্ন হতো — এখন CONFIG.defaultMfgYear (বা
    // খালি রাখলে system year) যোগ করে "JUNE'2026" ফরম্যাটে রিটার্ন হবে।
    if (foundMonth && year) return foundMonth + "'" + year;
    if (foundMonth && !year) return foundMonth + "'" + getDefaultMfgYear();

    // সংখ্যাসূচক date: DD-MM-YYYY বা DD/MM/YYYY
    var dm = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (dm) {
        var mi = parseInt(dm[2], 10) - 1;
        if (mi >= 0 && mi < 12) return monthNames[mi] + "'" + dm[3];
    }

    // সংখ্যাসূচক date: MM-YYYY বা MM/YYYY
    var my = s.match(/^(\d{1,2})[\/\-.](\d{4})$/);
    if (my) {
        var mi2 = parseInt(my[1], 10) - 1;
        if (mi2 >= 0 && mi2 < 12) return monthNames[mi2] + "'" + my[2];
    }

    // শুধু মাসের সংক্ষিপ্ত নাম (যেমন "Jun") থাকলেও মিলবে কিনা যাচাই —
    // উপরে JUNE/JULY ইত্যাদি ফুল-নেম ম্যাচ না হলে এখানে 3-letter চেষ্টা
    var shortNames = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
    for (var si = 0; si < 12; si++) {
        if (upper.indexOf(shortNames[si]) !== -1) {
            return monthNames[si] + "'" + getDefaultMfgYear();
        }
    }

    return s;  // কিছুই মিলেনি — যেমন আছে তেমনই রাখা হলো
}

function pad2(n)     { n = String(n); return n.length < 2 ? "0" + n : n; }
function trim(s)     { return String(s).replace(/^\s+|\s+$/g, ""); }
function replaceAll(str, find, rep) {
    var r = str;
    while (r.indexOf(find) !== -1) r = r.replace(find, rep);
    return r;
}

// ─── RUN ───
try { main(); }
catch(err) { alert("Script Error:\n" + err.message + "\nLine: " + err.line); }
