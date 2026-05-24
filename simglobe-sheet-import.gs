/**
 * ════════════════════════════════════════════════════════════════════
 *  SIMGLOBE — Nhập đơn từ App vào Sheet xuất hoá đơn
 *
 *  CƠ CHẾ MAPPING (v2):
 *    - Input là TÊN GÓI từ app (không phải tên bảng giá nữa)
 *    - NƯỚC, GÓI: map qua tab MAPPING (substring + regex pattern)
 *    - NGÀY: ưu tiên parse từ tên gói (\d+ Days/Day/ngày) → fallback cart.days → hỏi user
 *    - LOẠI SIM: fuzzy match từ cart.type + flag recharge (cũ)
 *
 *  HƯỚNG DẪN GHÉP VÀO SCRIPT HIỆN CÓ:
 *    1. Copy TOÀN BỘ block này dán vào script editor (cuối file).
 *    2. Sửa hàm onOpen() hiện có, thêm các dòng menu:
 *
 *       function onOpen() {
 *         SpreadsheetApp.getUi()
 *           .createMenu('⚙ Tools')
 *           .addItem('🧹 Dọn Conditional Formatting', 'cleanConditionalFormatting')
 *           .addSeparator()
 *           .addItem('🗑 Xóa dữ liệu giữ công thức', 'clearSelectedRowsKeepFormula')
 *           .addSeparator()
 *           .addItem('➕ Thêm hàng thông minh', 'addRowSmart')
 *           .addSeparator()
 *           .addItem('📥 Nhập đơn từ App', 'sgImportPayload')
 *           .addItem('🧠 Mở tab MAPPING',   'sgOpenMappingTab')
 *           .addItem('🔍 Mở tab _log (debug)', 'sgOpenLogTab')
 *           .addToUi();
 *       }
 *
 *    3. Save + reload spreadsheet.
 *    4. Lần đầu tiên:
 *       - Tab "MAPPING" auto setup (header xanh, frozen row, validation)
 *       - Mở file MAPPING_initial.csv → copy nội dung paste vào tab MAPPING từ dòng 2
 *         (119 dòng mapping pre-built: 64 NƯỚC + 20 GÓI special + 35 regex pattern)
 *       - Hoặc dùng menu "🧠 Mở tab MAPPING" để xem cách edit.
 *
 *  KIỂU DÒNG TRONG TAB MAPPING:
 *    - Loại: NƯỚC | GÓI (dropdown)
 *    - Text gốc (App): plaintext HOẶC bắt đầu "RE:" cho regex
 *      (vd "RE:(?i)^\s*1\s*gb\s*\/?\s*days?\s*$")
 *    - Map sang: option chính xác trong dropdown của sheet xuất hoá đơn
 *    - Lần dùng / Cập nhật cuối: tự động tăng
 * ════════════════════════════════════════════════════════════════════
 */


/* ══════════════════════════════════════════════════════════════
   CONFIG
   ══════════════════════════════════════════════════════════════ */
var SG_HEADER_ROW = 2;
var SG_HEADERS = {
  orderId    : 'MÃ ORDER ID',
  country    : 'NƯỚC',
  simType    : 'LOẠI SIM',
  days       : 'NGÀY',
  pkg        : 'GÓI',
  qty        : 'SỐ \nLƯỢNG',
  unitPrice  : 'ĐƠN GIÁ',
  discount   : 'TỔNG GIẢM GIÁ \nĐÃ PHÂN BỔ',
  orderCode  : 'MÃ ĐƠN HÀNG',
  link       : 'LINK ĐƠN HÀNG',
  retail     : 'XUẤT LẺ',
  mst        : 'MST',
  buyer      : 'TÊN NGƯỜI MUA, CTY',
  addr       : 'ĐỊA CHỈ',
  email      : 'EMAIL',
};
var SG_MAPPING_SHEET = 'MAPPING';
var SG_LOG_SHEET = '_log';


/* ══════════════════════════════════════════════════════════════
   MAIN — Nhập đơn từ App
   ══════════════════════════════════════════════════════════════ */
function sgImportPayload() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var active = sheet.getActiveRange();

  var headers = sgReadHeaders_(sheet);
  if (!headers.orderId) {
    ui.alert('❌ Không tìm thấy cột "MÃ ORDER ID" ở hàng ' + SG_HEADER_ROW + '.');
    return;
  }
  if (active.getColumn() !== headers.orderId.col) {
    ui.alert('⚠️ Hãy chọn 1 ô ở cột "MÃ ORDER ID" (cột ' + sgColLetter_(headers.orderId.col) + ') làm anchor.');
    return;
  }
  if (active.getRow() <= SG_HEADER_ROW) {
    ui.alert('⚠️ Anchor phải nằm dưới hàng tiêu đề (hàng > ' + SG_HEADER_ROW + ').');
    return;
  }

  var resp = ui.prompt('📥 Nhập đơn từ App', 'Paste payload (SIMGLOBE-SHEET:...):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var raw = resp.getResponseText().trim();
  if (!raw) { ui.alert('❌ Chưa paste payload.'); return; }

  var payload = sgDecodePayload_(raw);
  if (!payload) {
    sgAppendLog_(null, '❌ LỖI DECODE', active.getRow());
    ui.alert('❌ Payload không hợp lệ.');
    return;
  }
  if (!payload.items || !payload.items.length) {
    sgAppendLog_(payload, '❌ KHÔNG CÓ GÓI', active.getRow());
    ui.alert('❌ Payload không có gói nào.');
    return;
  }

  var sampleRow = Math.max(active.getRow(), SG_HEADER_ROW + 1);
  var dropdowns = {
    'NƯỚC'    : sgReadDropdown_(sheet, sampleRow, headers.country.col),
    'LOẠI SIM': sgReadDropdown_(sheet, sampleRow, headers.simType.col),
    'NGÀY'    : sgReadDropdown_(sheet, sampleRow, headers.days.col),
    'GÓI'     : sgReadDropdown_(sheet, sampleRow, headers.pkg.col),
  };

  var mappingState = sgLoadMapping_();

  // Resolve từng item
  var resolved = [];
  for (var i = 0; i < payload.items.length; i++) {
    var it = payload.items[i];
    var idxStr = '(gói ' + (i + 1) + '/' + payload.items.length + ')';
    var pkgText = (it.p || '').trim();   // INPUT chính cho NƯỚC + GÓI: tên gói

    // === NƯỚC: substring match từ tab MAPPING, fallback ASK user ===
    var rCountry = sgResolveCountry_(ui, pkgText, dropdowns['NƯỚC'], mappingState, sheet, headers.country.col, idxStr);
    if (rCountry.cancelled) { sgAppendLog_(payload, '⏹ HUỶ (NƯỚC) gói ' + (i+1), active.getRow()); ui.alert('⏹ Đã huỷ ở gói ' + (i+1) + '.'); return; }

    // === GÓI: exact mapping + regex pattern + fuzzy, fallback ASK user ===
    var rPkg = sgResolvePkg_(ui, pkgText, dropdowns['GÓI'], mappingState, sheet, headers.pkg.col, idxStr);
    if (rPkg.cancelled) { sgAppendLog_(payload, '⏹ HUỶ (GÓI) gói ' + (i+1), active.getRow()); ui.alert('⏹ Đã huỷ ở gói ' + (i+1) + '.'); return; }

    // === NGÀY: parse tên → fallback cart.days → ASK ===
    var daysInput = sgExtractDaysFromName_(pkgText) || (it.d || '').trim();
    var rDays = sgResolveSimple_(ui, 'NGÀY', daysInput, dropdowns['NGÀY'], sheet, headers.days.col, idxStr);
    if (rDays.cancelled) { sgAppendLog_(payload, '⏹ HUỶ (NGÀY) gói ' + (i+1), active.getRow()); ui.alert('⏹ Đã huỷ ở gói ' + (i+1) + '.'); return; }

    // === LOẠI SIM: fuzzy đơn giản (cart.type + recharge flag) ===
    var simTypeInput = it.r ? 'NẠP GÓI' : (it.t || '');
    var rSim = sgResolveSimple_(ui, 'LOẠI SIM', simTypeInput, dropdowns['LOẠI SIM'], sheet, headers.simType.col, idxStr);
    if (rSim.cancelled) { sgAppendLog_(payload, '⏹ HUỶ (LOẠI) gói ' + (i+1), active.getRow()); ui.alert('⏹ Đã huỷ ở gói ' + (i+1) + '.'); return; }

    resolved.push({
      country: rCountry.value,
      simType: rSim.value,
      days   : rDays.value,
      pkg    : rPkg.value,
      qty    : it.q,
      unitK  : it.u,
      discK  : it.g,
      note   : it.n || '',
    });
  }

  // Kiểm tra đủ hàng
  var startRow = active.getRow();
  var N = resolved.length;
  if (startRow + N - 1 > sheet.getLastRow() + 1) {
    sgAppendLog_(payload, '⚠️ THIẾU HÀNG', startRow);
    ui.alert('Sheet chỉ tới hàng ' + sheet.getLastRow() + ', cần ' + N + ' hàng từ ' + startRow + '. Dùng "➕ Thêm hàng thông minh" trước.');
    return;
  }

  // Fill
  for (var j = 0; j < N; j++) {
    var row = startRow + j;
    var r = resolved[j];
    sgSetCell_(sheet, row, headers.country.col,   r.country);
    sgSetCell_(sheet, row, headers.simType.col,   r.simType);
    sgSetCell_(sheet, row, headers.days.col,      r.days);
    sgSetCell_(sheet, row, headers.pkg.col,       r.pkg);
    sgSetCell_(sheet, row, headers.qty.col,       r.qty);
    sgSetCell_(sheet, row, headers.unitPrice.col, r.unitK * 1000);
    if (r.discK > 0) sgSetCell_(sheet, row, headers.discount.col, r.discK * 1000);
    if (headers.orderCode && payload.orderCode) sgSetCell_(sheet, row, headers.orderCode.col, payload.orderCode);
    if (headers.link && payload.link) sgSetCell_(sheet, row, headers.link.col, payload.link);
  }

  // VAT merge
  if (payload.vat) {
    var vatCols = [
      { h: headers.mst,   v: payload.vat.mst   || '', isMst: true },
      { h: headers.buyer, v: payload.vat.name  || '' },
      { h: headers.addr,  v: payload.vat.addr  || '' },
      { h: headers.email, v: payload.vat.email || '' },
    ];
    if (N > 1 && headers.retail) sgMergeRange_(sheet, startRow, headers.retail.col, N);
    for (var v = 0; v < vatCols.length; v++) {
      if (!vatCols[v].h) continue;
      if (N > 1) sgMergeRange_(sheet, startRow, vatCols[v].h.col, N);
      if (vatCols[v].isMst) {
        // MST: ép Plain Text format trước khi ghi để Sheets không strip leading 0
        sgSetMstCell_(sheet, startRow, vatCols[v].h.col, vatCols[v].v);
      } else {
        sgSetCell_(sheet, startRow, vatCols[v].h.col, vatCols[v].v);
      }
    }
  }

  // ─── ANCHOR (MÃ ORDER ID) ───
  // v10.7: 3 trường hợp:
  //   1) useCode=true + groupByType=true + có cả 2 loại → split 2 nhóm consecutive,
  //      mỗi nhóm anchor riêng với codes.esim / codes.physical.
  //   2) useCode=true + chỉ 1 loại → 1 anchor với code tương ứng.
  //   3) Còn lại (mặc định) → 1 anchor với tagPrefix + name (logic cũ).
  var useCode = payload.useCode === true;
  var codes = payload.codes || {};
  var groupByType = payload.groupByType === true;

  // Hàm helper: xác định nhóm của 1 item theo simType.
  // 'esim' / 'physical' / 'other'
  function _itemGroup(simType) {
    var s = String(simType || '').toUpperCase();
    if (s.indexOf('ESIM') !== -1) return 'esim';
    if (s.indexOf('SIM VẬT LÝ') !== -1 || s.indexOf('SIM VAT LY') !== -1 || s.indexOf('NẠP GÓI') !== -1 || s.indexOf('NAP GOI') !== -1) return 'physical';
    return 'other';
  }

  // Phân nhóm các row theo group
  var groupRuns = []; // array of { group, startRow, count }
  if (headers.orderId && N > 0) {
    var curGroup = _itemGroup(resolved[0].simType);
    var runStart = startRow;
    var runCount = 1;
    for (var k = 1; k < N; k++) {
      var g = _itemGroup(resolved[k].simType);
      if (g === curGroup) {
        runCount++;
      } else {
        groupRuns.push({ group: curGroup, startRow: runStart, count: runCount });
        curGroup = g;
        runStart = startRow + k;
        runCount = 1;
      }
    }
    groupRuns.push({ group: curGroup, startRow: runStart, count: runCount });
  }

  // Quyết định mode anchor
  var hasEsimCode = !!(codes.esim && String(codes.esim).trim());
  var hasPhysicalCode = !!(codes.physical && String(codes.physical).trim());
  var distinctGroups = {};
  for (var gg = 0; gg < groupRuns.length; gg++) distinctGroups[groupRuns[gg].group] = true;
  var hasEsimGroup = distinctGroups['esim'] === true;
  var hasPhysicalGroup = distinctGroups['physical'] === true;
  var isMixedConsecutive = groupRuns.length <= (hasEsimGroup && hasPhysicalGroup ? 2 : 1);

  // Apply anchors
  if (headers.orderId) {
    if (useCode && isMixedConsecutive && (hasEsimCode || hasPhysicalCode)) {
      // v10.7 MODE: mỗi run group có anchor riêng (dùng code tương ứng)
      for (var gr = 0; gr < groupRuns.length; gr++) {
        var run = groupRuns[gr];
        var anchorVal = '';
        if (run.group === 'esim' && hasEsimCode) anchorVal = String(codes.esim).trim();
        else if (run.group === 'physical' && hasPhysicalCode) anchorVal = String(codes.physical).trim();
        else continue; // không có code → skip nhóm này
        if (run.count > 1) sgMergeRange_(sheet, run.startRow, headers.orderId.col, run.count);
        sgSetCell_(sheet, run.startRow, headers.orderId.col, anchorVal);
      }
    } else if (payload.name) {
      // Mode mặc định: 1 anchor cho cả block với tagPrefix + name
      var tagPrefix = String(payload.tagPrefix || '');
      var displayName = tagPrefix + payload.name;
      if (N > 1) sgMergeRange_(sheet, startRow, headers.orderId.col, N);
      sgSetCell_(sheet, startRow, headers.orderId.col, displayName);
    }
  }

  // Tick TRUE vào cột XUẤT LẺ (anchor row nếu có merge VAT, từng hàng nếu đơn lẻ).
  if (headers.retail) {
    if (payload.vat) {
      // Đơn công ty: ô retail đã merge ở trên → set TRUE 1 phát vào anchor
      sgSetCell_(sheet, startRow, headers.retail.col, true);
    } else {
      // Đơn lẻ: tick TRUE từng hàng
      for (var rt = 0; rt < N; rt++) {
        sgSetCell_(sheet, startRow + rt, headers.retail.col, true);
      }
    }
  }

  // Gọi TRỰC TIẾP hàm xử lý của script highlight (cùng project) cho ĐÚNG vùng vừa fill.
  // Lý do: onChange của Apps Script không fire đáng tin cậy khi script khác setValue → tick
  // hiển thị nhưng highlight không chạy. Gọi tay đảm bảo highlight ngay, chỉ N hàng (nhanh).
  try {
    if (typeof getSheetConfig === 'function' && typeof processRows === 'function') {
      var hlConfig = getSheetConfig(sheet.getName());
      if (hlConfig) processRows(sheet, hlConfig, startRow, startRow + N - 1);
    }
  } catch (e) {
    Logger.log('highlight call skipped: ' + e);
  }

  var msg = '✅ Đã fill ' + N + ' gói từ hàng ' + startRow + '.\n';
  if (payload.orderCode) msg += 'Mã đơn: ' + payload.orderCode + '\n';
  if (useCode && (hasEsimCode || hasPhysicalCode)) {
    msg += 'Anchor: dùng MÃ\n';
    if (hasEsimCode) msg += '  • eSIM: ' + codes.esim + '\n';
    if (hasPhysicalCode) msg += '  • SIM VL: ' + codes.physical + '\n';
  } else if (payload.name) {
    var displayName2 = String(payload.tagPrefix || '') + payload.name;
    msg += 'Tên đơn: "' + displayName2 + '" (đã điền vào ô anchor)\n';
  }
  if (payload.vat && N > 1) msg += 'Đã merge VAT qua ' + N + ' hàng.\n';
  msg += '\n⏹ Bạn tự điền: NGÀY ORDER, NGÀY CUNG CẤP DV, ✓ ĐÃ XUẤT.';

  sgAppendLog_(payload, '✅ OK (' + N + ' gói)', startRow);
  ui.alert(msg);
}


/* ══════════════════════════════════════════════════════════════
   RESOLVE NƯỚC — substring match từ tên gói
   ══════════════════════════════════════════════════════════════ */
function sgResolveCountry_(ui, pkgText, dropdown, mappingState, sheet, colIdx, ctxStr) {
  if (!pkgText) return { value: '', cancelled: false };
  if (!dropdown.hasDropdown) return { value: '', cancelled: false };

  // Substring match: tìm trong tab MAPPING entry có raw nằm trọn trong pkgText
  // (token-aware: phải có word boundary)
  var hit = sgFindCountryInText_(pkgText, mappingState);
  if (hit && dropdown.options.indexOf(hit.mapped) >= 0) {
    sgIncrementMappingUse_(mappingState, hit.item);
    return { value: hit.mapped, cancelled: false };
  }

  // Không match → hỏi user
  return sgAskPickCountry_(ui, pkgText, dropdown, mappingState, sheet, colIdx, ctxStr);
}


function sgFindCountryInText_(pkgText, state) {
  if (!pkgText || !state || !state.items) return null;
  var nT = sgNorm_(pkgText);
  var padded = ' ' + nT + ' ';
  // Lọc items NƯỚC, bỏ regex entry (nếu có).
  var sorted = state.items.slice()
    .filter(function(it){ return it.type === 'NƯỚC' && it.raw.substring(0, 3) !== 'RE:'; })
    .sort(function(a, b){ return b.raw.length - a.raw.length; });
  for (var i = 0; i < sorted.length; i++) {
    var item = sorted[i];
    var nR = sgNorm_(item.raw);
    if (!nR || nR.length < 2) continue;
    if (padded.indexOf(' ' + nR + ' ') >= 0) {
      return { item: item, mapped: item.mapped, matchType: 'substring' };
    }
  }
  return null;
}


function sgAskPickCountry_(ui, pkgText, dropdown, mappingState, sheet, colIdx, ctxStr) {
  var options = dropdown.options;
  var list = options.map(function(o, i){ return (i+1) + '. ' + o; }).join('\n');
  var displayRaw = pkgText.replace(/[\r\n]+/g, ' / ').replace(/\s+/g, ' ').trim();
  if (displayRaw.length > 100) displayRaw = displayRaw.substring(0, 97) + '...';
  var prompt = '🌍 Chọn NƯỚC cho gói ' + ctxStr + ':\n  "' + displayRaw + '"\n\n'
             + 'Chọn số trong list:\n' + list
             + '\n\n────────────────\n'
             + 'Nhập:\n'
             + '  • Số (1-' + options.length + ') → mapping lưu vào tab MAPPING\n'
             + '  • "Text mới +" cuối → thêm option mới vào dropdown\n'
             + '  • (để trống) → skip, không fill';
  var resp = ui.prompt('🌍 NƯỚC ' + ctxStr, prompt, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return { value: '', cancelled: true };
  var input = resp.getResponseText().trim();
  if (!input) return { value: '', cancelled: false };

  // Số nguyên
  var idx = parseInt(input, 10);
  if (!isNaN(idx) && idx >= 1 && idx <= options.length && String(idx) === input) {
    var picked = options[idx - 1];
    // Lưu mapping với "key gốc" rút ngắn: extract phần đầu tên gói trước "X Days"
    var keyToSave = sgExtractCountryKey_(pkgText);
    if (keyToSave) sgUpsertMapping_(mappingState, 'NƯỚC', keyToSave, picked);
    return { value: picked, cancelled: false };
  }

  // Text + "+" → thêm option
  if (input.charAt(input.length - 1) === '+') {
    var newOpt = input.substring(0, input.length - 1).trim();
    if (!newOpt) return { value: '', cancelled: false };
    var added = sgAddToDropdown_(sheet, colIdx, newOpt, dropdown);
    if (!added.ok) {
      ui.alert('⚠️ Không tự thêm được: ' + added.reason + '\nHãy thêm tay trong sheet/range nguồn.');
      return { value: '', cancelled: false };
    }
    dropdown.options.push(newOpt);
    var k2 = sgExtractCountryKey_(pkgText);
    if (k2) sgUpsertMapping_(mappingState, 'NƯỚC', k2, newOpt);
    return { value: newOpt, cancelled: false };
  }

  // Text match option
  var inputNorm = sgNorm_(input);
  for (var k = 0; k < options.length; k++) {
    if (sgNorm_(options[k]) === inputNorm) {
      var k3 = sgExtractCountryKey_(pkgText);
      if (k3) sgUpsertMapping_(mappingState, 'NƯỚC', k3, options[k]);
      return { value: options[k], cancelled: false };
    }
  }

  // Không match — confirm fill thô
  var confirm = ui.alert('Text "' + input + '" không khớp option nào. Fill thô?', ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return { value: '', cancelled: false };
  return { value: input, cancelled: false };
}


// Extract "country key" để lưu mapping: lấy phần đầu tên gói trước "X Days" / số đầu tiên
//   "Mainland China A, 10 Days, 50GB" → "Mainland China A"
//   "TRUE Thailand - Thailand Tourist 449" → "TRUE Thailand"
//   "USA, 1 Day, 500MB" → "USA"
function sgExtractCountryKey_(text) {
  if (!text) return '';
  var s = String(text).split(/[\r\n]+/)[0] || '';
  // Cắt tại "<số> Days/Day/ngày"
  s = s.replace(/,\s*\d+\s*(days?|ngày).*/i, '');
  // Cắt tại "<số>GB" / "<số>MB" nếu vẫn còn
  s = s.replace(/,\s*\d+(GB|MB|TB)\b.*/i, '');
  // Cắt sau dấu "-" hoặc " - " (nhiều tên có format "Country - Variant")
  // → giữ phần trước vì đó là country
  // Nhưng "TRUE Thailand - Thailand Tourist" thì cắt sẽ mất "Thailand Tourist"
  // → không cắt, giữ nguyên
  return s.trim().replace(/[,\s]+$/, '');
}


/* ══════════════════════════════════════════════════════════════
   RESOLVE GÓI — substring + regex pattern
   ══════════════════════════════════════════════════════════════ */
function sgResolvePkg_(ui, pkgText, dropdown, mappingState, sheet, colIdx, ctxStr) {
  if (!pkgText) return { value: '', cancelled: false };
  if (!dropdown.hasDropdown) return { value: pkgText, cancelled: false };

  // 1. Extract pkg part từ tên (phần sau "X Days,")
  var pkgPart = sgExtractPkgPart_(pkgText);

  // 2. Try regex patterns trong MAPPING (entry có raw bắt đầu "RE:")
  //    CHIẾN LƯỢC: collect TẤT CẢ pattern match được, chọn match có chuỗi khớp DÀI NHẤT.
  //    Lý do: pattern "unlimited.*10mbps" cụ thể hơn pattern "^unlimited$" → match string dài hơn → ưu tiên.
  //    Match trên cả pkgPart lẫn full text.
  var regexCandidates = [];
  for (var i = 0; i < mappingState.items.length; i++) {
    var item = mappingState.items[i];
    if (item.type !== 'GÓI') continue;
    if (item.raw.substring(0, 3) !== 'RE:') continue;
    if (dropdown.options.indexOf(item.mapped) < 0) continue;
    try {
      var pat = item.raw.substring(3);
      var flags = '';
      if (pat.substring(0, 4) === '(?i)') {
        pat = pat.substring(4);
        flags += 'i';
      }
      var re = new RegExp(pat, flags);
      var m1 = pkgPart ? pkgPart.match(re) : null;
      var m2 = pkgText.match(re);
      // Lấy match dài nhất giữa pkgPart/fullText
      var matchLen = 0;
      if (m1 && m1[0]) matchLen = Math.max(matchLen, m1[0].length);
      if (m2 && m2[0]) matchLen = Math.max(matchLen, m2[0].length);
      if (matchLen > 0) {
        regexCandidates.push({ item: item, matchLen: matchLen });
      }
    } catch (e) { /* invalid regex in MAPPING — skip */ }
  }
  if (regexCandidates.length > 0) {
    regexCandidates.sort(function(a, b) { return b.matchLen - a.matchLen; });
    var best = regexCandidates[0];
    sgIncrementMappingUse_(mappingState, best.item);
    return { value: best.item.mapped, cancelled: false };
  }

  // 3. Try exact/substring match cho GÓI (entry không có "RE:")
  var hit = sgFindPkgInText_(pkgText, pkgPart, mappingState);
  if (hit && dropdown.options.indexOf(hit.mapped) >= 0) {
    sgIncrementMappingUse_(mappingState, hit.item);
    return { value: hit.mapped, cancelled: false };
  }

  // 4. Không match → hỏi user
  return sgAskPickPkg_(ui, pkgText, pkgPart, dropdown, mappingState, sheet, colIdx, ctxStr);
}


function sgFindPkgInText_(fullText, pkgPart, state) {
  var nFull = sgNorm_(fullText);
  var nPart = sgNorm_(pkgPart);
  var pFull = ' ' + nFull + ' ';
  var pPart = ' ' + nPart + ' ';
  var sorted = state.items.slice().filter(function(it){ return it.type === 'GÓI' && it.raw.substring(0,3) !== 'RE:'; }).sort(function(a,b){ return b.raw.length - a.raw.length; });
  for (var i = 0; i < sorted.length; i++) {
    var item = sorted[i];
    var nR = sgNorm_(item.raw);
    if (!nR || nR.length < 2) continue;
    // Check trong pkgPart trước
    if (nPart === nR) return { item: item, mapped: item.mapped, matchType: 'exact-part' };
    if (pPart.indexOf(' ' + nR + ' ') >= 0) return { item: item, mapped: item.mapped, matchType: 'substring-part' };
    // Check trong full text
    if (pFull.indexOf(' ' + nR + ' ') >= 0) return { item: item, mapped: item.mapped, matchType: 'substring-full' };
  }
  return null;
}


function sgAskPickPkg_(ui, pkgText, pkgPart, dropdown, mappingState, sheet, colIdx, ctxStr) {
  var options = dropdown.options;
  var list = options.map(function(o, i){ return (i+1) + '. ' + o; }).join('\n');
  var displayFull = pkgText.replace(/[\r\n]+/g, ' / ').replace(/\s+/g, ' ').trim();
  if (displayFull.length > 80) displayFull = displayFull.substring(0, 77) + '...';
  var prompt = '📦 Chọn GÓI ' + ctxStr + ':\n  Tên gói: "' + displayFull + '"\n  Phần pkg parse được: "' + pkgPart + '"\n\n'
             + list
             + '\n\n────────────────\n'
             + 'Nhập:\n'
             + '  • Số (1-' + options.length + ') → mapping lưu tab MAPPING\n'
             + '  • "Text mới +" cuối → thêm option mới vào dropdown\n'
             + '  • (để trống) → skip';
  var resp = ui.prompt('📦 GÓI ' + ctxStr, prompt, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return { value: '', cancelled: true };
  var input = resp.getResponseText().trim();
  if (!input) return { value: '', cancelled: false };

  var idx = parseInt(input, 10);
  if (!isNaN(idx) && idx >= 1 && idx <= options.length && String(idx) === input) {
    var picked = options[idx - 1];
    // Lưu mapping với key = pkgPart (gọn) nếu pkgPart != fullText
    var key = pkgPart && pkgPart.length >= 3 ? pkgPart : pkgText;
    sgUpsertMapping_(mappingState, 'GÓI', key, picked);
    return { value: picked, cancelled: false };
  }

  if (input.charAt(input.length - 1) === '+') {
    var newOpt = input.substring(0, input.length - 1).trim();
    if (!newOpt) return { value: '', cancelled: false };
    var added = sgAddToDropdown_(sheet, colIdx, newOpt, dropdown);
    if (!added.ok) {
      ui.alert('⚠️ Không thêm được: ' + added.reason);
      return { value: '', cancelled: false };
    }
    dropdown.options.push(newOpt);
    var key2 = pkgPart && pkgPart.length >= 3 ? pkgPart : pkgText;
    sgUpsertMapping_(mappingState, 'GÓI', key2, newOpt);
    return { value: newOpt, cancelled: false };
  }

  var inputNorm = sgNorm_(input);
  for (var k = 0; k < options.length; k++) {
    if (sgNorm_(options[k]) === inputNorm) {
      var key3 = pkgPart && pkgPart.length >= 3 ? pkgPart : pkgText;
      sgUpsertMapping_(mappingState, 'GÓI', key3, options[k]);
      return { value: options[k], cancelled: false };
    }
  }

  var confirm = ui.alert('Text "' + input + '" không khớp option. Fill thô?', ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return { value: '', cancelled: false };
  return { value: input, cancelled: false };
}


// Extract pkg part từ tên: phần sau "X Days," (loại bỏ ", 128kbps", ", EXP : ...")
//   "Korea, 1 Day, 3GB /day, 128kbps" → "3GB /day"
//   "Korea, 1GB /day, 128kbps"        → "1GB /day"   (strip country prefix)
//   "TRUE Thailand - Thailand Tourist 449" → "TRUE Thailand - Thailand Tourist 449"
function sgExtractPkgPart_(name) {
  if (!name) return '';
  var s = String(name).split(/[\r\n]+/)[0] || '';
  s = s.replace(/,\s*\d+\s*kbps[^,]*/gi, '');
  s = s.replace(/,\s*1\s*mbps[^,]*/gi, '');
  s = s.replace(/,\s*10\s*mbps[^,]*/gi, '');
  s = s.replace(/,\s*exp\s*:.*$/gi, '');
  s = s.trim().replace(/,\s*$/, '');
  // Có "X Days," → lấy phần sau
  var m = s.match(/,\s*\d+\s+days?\s*,\s*(.+)$/i);
  if (m) return m[1].trim();
  // Kết thúc bằng "X Days" — không có pkg riêng (vd "AT&T North America, 20 days")
  m = s.match(/,\s*\d+\s+days?\s*$/i);
  if (m) return '';
  // Không có "Days" — thử strip country prefix trước dấu phẩy đầu tiên
  // Nếu phần còn lại có pattern data (GB/MB/Unlimited/Premium...) → dùng nó
  var commaIdx = s.indexOf(',');
  if (commaIdx > 0) {
    var afterComma = s.substring(commaIdx + 1).trim();
    if (afterComma && /\b(\d+\s*(gb|mb|tb)|unlimited|premium|tổng)\b/i.test(afterComma)) {
      return afterComma;
    }
  }
  return s;
}


/* ══════════════════════════════════════════════════════════════
   RESOLVE NGÀY/LOẠI SIM — fuzzy đơn giản (không cache mapping)
   ══════════════════════════════════════════════════════════════ */
function sgResolveSimple_(ui, fieldLabel, rawText, dropdown, sheet, colIdx, ctxStr) {
  if (!rawText) return { value: '', cancelled: false };
  if (!dropdown.hasDropdown) return { value: rawText, cancelled: false };

  // Exact match (case-insensitive)
  var nT = sgNorm_(rawText);
  for (var i = 0; i < dropdown.options.length; i++) {
    if (sgNorm_(dropdown.options[i]) === nT) return { value: dropdown.options[i], cancelled: false };
  }

  // Fuzzy: token-aware substring
  var padded = ' ' + nT + ' ';
  var best = null, bestLen = 0;
  for (var j = 0; j < dropdown.options.length; j++) {
    var nO = sgNorm_(dropdown.options[j]);
    if (!nO || nO.length < 2) continue;
    if (padded.indexOf(' ' + nO + ' ') >= 0 && nO.length > bestLen) {
      best = dropdown.options[j];
      bestLen = nO.length;
    }
  }
  if (best) return { value: best, cancelled: false };

  // Không match → hỏi user
  var list = dropdown.options.map(function(o, k){ return (k+1) + '. ' + o; }).join('\n');
  var prompt = fieldLabel + ' không nhận diện: "' + rawText + '"\n\nChọn số:\n' + list;
  var resp = ui.prompt(fieldLabel + ' ' + ctxStr, prompt, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return { value: '', cancelled: true };
  var input = resp.getResponseText().trim();
  if (!input) return { value: '', cancelled: false };
  var idx = parseInt(input, 10);
  if (!isNaN(idx) && idx >= 1 && idx <= dropdown.options.length) {
    return { value: dropdown.options[idx-1], cancelled: false };
  }
  return { value: input, cancelled: false };
}


// Parse số ngày từ tên gói
function sgExtractDaysFromName_(name) {
  if (!name) return '';
  var m = String(name).match(/(\d+)\s*(days?|ngày|NGÀY)/i);
  if (m) return m[1] + ' ngày';
  return '';
}


/* ══════════════════════════════════════════════════════════════
   ADD TO DROPDOWN
   ══════════════════════════════════════════════════════════════ */
function sgAddToDropdown_(sheet, colIdx, newValue, dropdownInfo) {
  try {
    var sampleRow = SG_HEADER_ROW + 1;
    while (sampleRow <= sheet.getLastRow()) {
      if (sheet.getRange(sampleRow, colIdx).getDataValidation()) break;
      sampleRow++;
    }
    var dv = sheet.getRange(sampleRow, colIdx).getDataValidation();
    if (!dv) return { ok: false, reason: 'Không tìm thấy data validation.' };

    var crit = dv.getCriteriaType();
    var args = dv.getCriteriaValues();
    if (crit === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      var list = args[0].slice();
      list.push(newValue);
      var newDv = SpreadsheetApp.newDataValidation()
        .requireValueInList(list, true)
        .setAllowInvalid(dv.getAllowInvalid())
        .build();
      var firstDataRow = SG_HEADER_ROW + 1;
      var lastRow = Math.max(sheet.getMaxRows(), 1000);
      sheet.getRange(firstDataRow, colIdx, lastRow - firstDataRow + 1, 1).setDataValidation(newDv);
      return { ok: true };
    }
    if (crit === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      return { ok: false, reason: 'Dropdown lấy từ range (tab khác). Thêm tay vào tab nguồn.' };
    }
    return { ok: false, reason: 'Loại data validation không hỗ trợ thêm tự động.' };
  } catch (e) { return { ok: false, reason: e.message }; }
}


/* ══════════════════════════════════════════════════════════════
   TAB MAPPING — đọc/ghi
   ══════════════════════════════════════════════════════════════ */
function sgSetupMappingTab_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SG_MAPPING_SHEET);
  if (!sheet) sheet = ss.insertSheet(SG_MAPPING_SHEET);
  var h1 = sheet.getRange(1, 1).getValue();
  if (h1 !== 'Loại') {
    sheet.getRange(1, 1, 1, 5).setValues([['Loại', 'Text gốc (từ App)', 'Map sang (option dropdown)', 'Lần dùng', 'Cập nhật cuối']]);
    sheet.getRange(1, 1, 1, 5).setBackground('#4a90e2').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(1, 36);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 90);
    sheet.setColumnWidth(2, 350);
    sheet.setColumnWidth(3, 320);
    sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 150);
    sheet.getRange(2, 1, 999, 1).setHorizontalAlignment('center');
    sheet.getRange(2, 4, 999, 1).setHorizontalAlignment('center');
    sheet.getRange(2, 5, 999, 1).setHorizontalAlignment('center').setNumberFormat('yyyy-mm-dd hh:mm');
    var typeDv = SpreadsheetApp.newDataValidation().requireValueInList(['NƯỚC', 'GÓI'], true).build();
    sheet.getRange(2, 1, 999, 1).setDataValidation(typeDv);
    try { sheet.getRange(1, 1, 1000, 5).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY); } catch (e) {}
  }
  return sheet;
}

function sgLoadMapping_() {
  var sheet = sgSetupMappingTab_();
  var state = { items: [], sheet: sheet };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return state;
  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var type = String(row[0] || '').trim();
    var raw = String(row[1] || '').trim();
    var mapped = String(row[2] || '').trim();
    if (!type || !raw || !mapped) continue;
    state.items.push({
      type: type, raw: raw, mapped: mapped,
      count: Number(row[3]) || 0, lastUsed: row[4] || '',
      rowIdx: i + 2,
    });
  }
  return state;
}

function sgUpsertMapping_(state, type, raw, mapped) {
  if (!raw || !mapped) return;
  // Check existing by exact match (case sensitive)
  for (var i = 0; i < state.items.length; i++) {
    var it = state.items[i];
    if (it.type === type && it.raw === raw) {
      it.count += 1;
      it.lastUsed = new Date();
      if (it.mapped !== mapped) {
        it.mapped = mapped;
        state.sheet.getRange(it.rowIdx, 3).setValue(mapped);
      }
      state.sheet.getRange(it.rowIdx, 4).setValue(it.count);
      state.sheet.getRange(it.rowIdx, 5).setValue(it.lastUsed);
      return;
    }
  }
  // Append new
  var now = new Date();
  var newRow = state.sheet.getLastRow() + 1;
  if (newRow < 2) newRow = 2;
  state.sheet.getRange(newRow, 1, 1, 5).setValues([[type, raw, mapped, 1, now]]);
  state.items.push({ type: type, raw: raw, mapped: mapped, count: 1, lastUsed: now, rowIdx: newRow });
}

function sgIncrementMappingUse_(state, item) {
  item.count += 1;
  var now = new Date();
  item.lastUsed = now;
  state.sheet.getRange(item.rowIdx, 4).setValue(item.count);
  state.sheet.getRange(item.rowIdx, 5).setValue(now);
}

function sgOpenMappingTab() {
  var sheet = sgSetupMappingTab_();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert(
    '🧠 Đang ở tab MAPPING.\n\n'
  + 'Lần đầu setup: paste file MAPPING_initial.csv vào từ dòng 2 (~119 dòng pre-built).\n\n'
  + '• Cột Loại: NƯỚC / GÓI (dropdown).\n'
  + '• Cột "Text gốc (App)":\n'
  + '    - Plaintext: substring match (vd "Mainland China A")\n'
  + '    - "RE:..." prefix: regex match (vd "RE:(?i)^\\s*3\\s*gb\\s*/?\\s*days?\\s*$")\n'
  + '• Cột "Map sang": option chính xác trong dropdown sheet xuất hoá đơn.\n'
  + '• Có thể sửa tay cột 3, hoặc xoá hàng để reset 1 mapping.'
  );
}


/* ══════════════════════════════════════════════════════════════
   PAYLOAD DECODE
   ══════════════════════════════════════════════════════════════ */
function sgDecodePayload_(raw) {
  try {
    var s = String(raw).trim();
    var prefix = 'SIMGLOBE-SHEET:';
    if (s.indexOf(prefix) === 0) s = s.substring(prefix.length);
    s = s.replace(/-/g, '+').replace(/_/g, '/');
    while (s.length % 4) s += '=';
    var bytes = Utilities.base64Decode(s);
    var jsonStr = Utilities.newBlob(bytes).getDataAsString('UTF-8');
    var obj = JSON.parse(jsonStr);
    if (!obj || obj.kind !== 'simglobe_sheet_v1') return null;
    return obj;
  } catch (e) {
    Logger.log('Decode failed: ' + e);
    return null;
  }
}


/* ══════════════════════════════════════════════════════════════
   HEADER LOOKUP
   ══════════════════════════════════════════════════════════════ */
function sgReadHeaders_(sheet) {
  var headerRange = sheet.getRange(SG_HEADER_ROW, 1, 1, sheet.getLastColumn());
  var values = headerRange.getValues()[0];
  var result = {};
  for (var key in SG_HEADERS) {
    if (!SG_HEADERS.hasOwnProperty(key)) continue;
    var targetNorm = sgNormHeader_(SG_HEADERS[key]);
    for (var c = 0; c < values.length; c++) {
      if (sgNormHeader_(values[c]) === targetNorm) {
        result[key] = { col: c + 1, label: values[c] };
        break;
      }
    }
  }
  return result;
}

function sgNormHeader_(s) {
  return String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();
}


/* ══════════════════════════════════════════════════════════════
   DROPDOWN READ
   ══════════════════════════════════════════════════════════════ */
function sgReadDropdown_(sheet, row, col) {
  try {
    var dv = sheet.getRange(row, col).getDataValidation();
    if (!dv) return { options: [], hasDropdown: false };
    var crit = dv.getCriteriaType();
    var args = dv.getCriteriaValues();
    if (!args || !args.length) return { options: [], hasDropdown: false };
    if (crit === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      return { options: args[0].map(function(x){ return String(x); }), hasDropdown: true };
    }
    if (crit === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      var rng = args[0];
      var vals = rng.getValues();
      var flat = [];
      for (var r = 0; r < vals.length; r++) {
        for (var c = 0; c < vals[r].length; c++) {
          var v = String(vals[r][c] || '').trim();
          if (v) flat.push(v);
        }
      }
      var seen = {}, uniq = [];
      flat.forEach(function(x){ if (!seen[x]) { seen[x] = 1; uniq.push(x); } });
      return { options: uniq, hasDropdown: true };
    }
    return { options: [], hasDropdown: false };
  } catch (e) { return { options: [], hasDropdown: false }; }
}


/* ══════════════════════════════════════════════════════════════
   STRING UTILS
   ══════════════════════════════════════════════════════════════ */
function sgNorm_(s) {
  s = String(s || '').toLowerCase();
  s = s.replace(/[àáạảãâầấậẩẫăằắặẳẵ]/g, 'a')
       .replace(/[èéẹẻẽêềếệểễ]/g, 'e')
       .replace(/[ìíịỉĩ]/g, 'i')
       .replace(/[òóọỏõôồốộổỗơờớợởỡ]/g, 'o')
       .replace(/[ùúụủũưừứựửữ]/g, 'u')
       .replace(/[ỳýỵỷỹ]/g, 'y')
       .replace(/đ/g, 'd');
  s = s.replace(/[^\x20-\x7e]/g, ' ');
  s = s.replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim();
  return s;
}


/* ══════════════════════════════════════════════════════════════
   CELL HELPERS
   ══════════════════════════════════════════════════════════════ */
function sgSetCell_(sheet, row, col, value) {
  if (value === '' || value === null || value === undefined) return;
  var cell = sheet.getRange(row, col);
  if (cell.getFormula()) return;
  cell.setValue(value);
}

// Ghi MST: ép cell sang Plain Text (@) trước khi setValue.
// Lý do: Google Sheets mặc định convert chuỗi digit-only sang number → strip leading 0.
// MST VN nhiều khi bắt đầu bằng "0" (vd "0123456789"), nếu không ép text sẽ mất số 0.
// Strip "@" prefix nếu có (do client gửi từ cache của vat-sheet.gs).
function sgSetMstCell_(sheet, row, col, value) {
  if (value === '' || value === null || value === undefined) return;
  var s = String(value).trim();
  if (!s) return;
  // Bỏ prefix "@" / "'" nếu có (giá trị lưu từ vat-sheet.gs có prefix bảo vệ)
  if (s.charAt(0) === '@' || s.charAt(0) === "'") s = s.substring(1).trim();
  if (!s) return;
  var cell = sheet.getRange(row, col);
  if (cell.getFormula()) return;
  cell.setNumberFormat('@'); // Plain Text — KHÔNG strip leading 0
  cell.setValue(s);
}

function sgMergeRange_(sheet, startRow, col, numRows) {
  if (numRows <= 1) return;
  try {
    var range = sheet.getRange(startRow, col, numRows, 1);
    range.breakApart();
    range.mergeVertically();
  } catch (e) { Logger.log('mergeRange failed col ' + col + ': ' + e); }
}

function sgColLetter_(col) {
  var letter = '', t;
  while (col > 0) {
    t = (col - 1) % 26;
    letter = String.fromCharCode(65 + t) + letter;
    col = (col - 1 - t) / 26 | 0;
  }
  return letter;
}


/* ══════════════════════════════════════════════════════════════
   LOG TAB
   ══════════════════════════════════════════════════════════════ */
function sgSetupLogTab_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SG_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SG_LOG_SHEET);
    sheet.getRange(1, 1, 1, 7).setValues([['Thời gian', 'Mã đơn', 'Tên đơn', 'Số gói', 'Anchor row', 'Trạng thái', 'Payload decoded']]);
    sheet.getRange(1, 1, 1, 7).setBackground('#444444').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setRowHeight(1, 30);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 60);
    sheet.setColumnWidth(5, 90);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 600);
    sheet.getRange(2, 1, 999, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange(2, 4, 999, 2).setHorizontalAlignment('center');
    sheet.getRange(2, 7, 999, 1).setVerticalAlignment('top').setWrap(false);
    try { sheet.hideSheet(); } catch (e) {}
  }
  return sheet;
}

function sgAppendLog_(payload, status, anchorRow) {
  try {
    var sheet = sgSetupLogTab_();
    var now = new Date();
    var pretty = '';
    if (payload) {
      try {
        var expanded = {
          orderCode: payload.orderCode || '',
          name: payload.name || '',
          link: payload.link || '',
          vat: payload.vat || null,
          items: (payload.items || []).map(function(it) {
            return { package: it.p, type: it.t, days: it.d, qty: it.q, unitK: it.u, discountK: it.g, recharge: !!it.r, note: it.n };
          }),
        };
        pretty = JSON.stringify(expanded, null, 2);
      } catch (e) { pretty = '[Pretty failed: ' + e.message + ']'; }
    }
    sheet.appendRow([now, (payload && payload.orderCode) || '', (payload && payload.name) || '', (payload && payload.items) ? payload.items.length : 0, anchorRow || '', status, pretty]);
  } catch (e) { Logger.log('appendLog: ' + e); }
}

function sgOpenLogTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SG_LOG_SHEET);
  if (!sheet) sheet = sgSetupLogTab_();
  try { sheet.showSheet(); } catch (e) {}
  ss.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert('🔍 Tab _log: history các lần Import (✅/⏹/❌).');
}
