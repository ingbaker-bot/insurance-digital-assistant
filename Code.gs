/**
 * 產險數位特助系統 - 1150523A
 *
 * 本版新增：
 * 1. CONFIG 新增旅平卡兩張工作表名稱
 * 2. syncTravelCards()：同步旅平卡 JSON → 旅平卡主表 + 旅平卡被保人名冊
 * 3. searchTravelCards(input)：搜尋旅平卡（姓名/電話/身分證/卡片類別）
 * 4. getTravelCardDetails(certNo)：取得單筆旅平卡 + 完整被保人名冊
 * 5. onOpen() 新增旅平卡同步選單項目
 */

const CONFIG = {
  SPREADSHEET_ID:       '1PwMyCB00FOQLsyEDHEFfBv2M-I4SKUi24f73YOiVAxs',
  SOURCE_FOLDER_ID:     '1tDwHVj-daXEwRrs6Zkg16d5vYMpyGUtQ',
  TRAVEL_FOLDER_ID:     '1tDwHVj-daXEwRrs6Zkg16d5vYMpyGUtQ',
  MAIN_SHEET_NAME:      '保單主表',
  DETAIL_SHEET_NAME:    '險種明細表',
  SETTING_SHEET_NAME:   '系統設定',
  ARCHIVE_SHEET_NAME:   '原始資料',
  TRAVEL_MAIN_SHEET:    '旅平卡主表',
  TRAVEL_MEMBER_SHEET:  '旅平卡被保人名冊',
  COVERAGE_SHEET_NAME:  '比對旅平卡覆蓋表',
  NICKNAME_SHEET_NAME:  '客戶稱呼表'             // 新增：客戶稱呼管理
};

// ═══════════════════════════════════════════════════════
//  選單
// ═══════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚀 產險系統功能')
      .addItem('執行批次同步（產險）', 'runAutoSyncFixed')
      .addItem('🏖️ 執行旅平卡同步', 'syncTravelCards')
      .addItem('🔍 比對旅平卡覆蓋率', 'compareTravelCoverage')
      .addItem('🧹 清理過期保單（三個月前）', 'cleanupExpiredPolicies')
      .addItem('📅 手動執行本週續保提醒', 'checkAndSendWeeklyReminders')
      .addSeparator()
      .addItem('重新編排序號', 'reIndexAllSheets')
      .addToUi();
}

// ═══════════════════════════════════════════════════════
//  系統設定讀取
// ═══════════════════════════════════════════════════════
function getSystemConfig() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SETTING_SHEET_NAME);
    if (!sheet) return getDefaultConfig();
    const data  = sheet.getDataRange().getValues();
    const cfg   = getDefaultConfig();
    data.forEach(row => {
      const key = String(row[0] || "").trim();
      const val = String(row[1] || "").trim();
      if (key && val) cfg[key] = val;
    });
    return cfg;
  } catch(e) {
    Logger.log('getSystemConfig error: ' + e);
    return getDefaultConfig();
  }
}

function getDefaultConfig() {
  return {
    '顧問姓名': 'Pei-lin',
    '顧問電話': '',
    '顧問LINE':  '',
    '緊急天數':  '30',
    '警示天數':  '60'
  };
}

// ═══════════════════════════════════════════════════════
//  產險主執行函式
// ═══════════════════════════════════════════════════════
function runAutoSyncFixed() {
  const ss          = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const mainSheet   = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  const detailSheet = ss.getSheetByName(CONFIG.DETAIL_SHEET_NAME);

  // ── 步驟1：讀取主表現有資料，記住手動填入的 W/X/Y ──
  const lastRow  = mainSheet.getLastRow();
  const mainData = lastRow > 1 ? mainSheet.getRange(1, 1, lastRow, 25).getValues() : [];
  const manualMap = new Map();
  mainData.forEach((row, i) => {
    if (i > 0 && row[1]) {
      manualMap.set(String(row[1]), {
        w: String(row[22] || "").trim() || "無",
        x: String(row[23] || "").trim() || "無",
        y: String(row[24] || "").trim() || "無"
      });
    }
  });

  // ── 步驟2：清空主表與明細表 ──
  if (lastRow > 1) mainSheet.getRange(2, 1, lastRow - 1, 25).clear();
  const lastDetailRow = detailSheet.getLastRow();
  if (lastDetailRow > 1) detailSheet.getRange(2, 1, lastDetailRow - 1, 8).clear();

  // ── 步驟3：讀取 Drive 所有 JSON（跳過旅平卡格式）──
  const rawPolicyMap    = new Map();
  const detailBucket    = new Map();
  const detailDedupeSet = new Set();

  const files = DriveApp.getFolderById(CONFIG.SOURCE_FOLDER_ID).getFilesByType("application/json");
  while (files.hasNext()) {
    const file = files.next();
    let jsonContent;
    try {
      jsonContent = JSON.parse(file.getBlob().getDataAsString());
    } catch(e) {
      Logger.log('⚠️ JSON 解析失敗：' + file.getName());
      continue;
    }

    // 旅平卡 JSON 沒有「保單號碼」欄位，有「憑證號碼」→ 跳過
    if (!Array.isArray(jsonContent) || jsonContent.length === 0) continue;
    if (!jsonContent[0]["保單號碼"] && jsonContent[0]["憑證號碼"]) {
      Logger.log('略過旅平卡檔案：' + file.getName());
      continue;
    }

    jsonContent.forEach(policy => {
      const policyNo = String(policy["保單號碼"] || "").trim();
      if (!policyNo) return;

      if (rawPolicyMap.has(policyNo)) {
        const existing     = rawPolicyMap.get(policyNo);
        const existingDate = robustParseDate(existing["基本資訊"]?.["生效日"] || "") || new Date(0);
        const newDate      = robustParseDate(policy["基本資訊"]?.["生效日"]   || "") || new Date(0);
        if (newDate <= existingDate) return;
      }
      rawPolicyMap.set(policyNo, policy);

      if (policy["險種明細"]) {
        if (!detailBucket.has(policyNo)) detailBucket.set(policyNo, []);
        filterValidDetailsImproved(policy["險種明細"]).forEach(item => {
          const key = policyNo + "|" + (item["險種代號"]||"") + "|" + (item["保額"]||"") + "|" + (item["保費"]||"");
          if (!detailDedupeSet.has(key)) {
            detailDedupeSet.add(key);
            detailBucket.get(policyNo).push([
              0, policyNo,
              item["險種代號"] || "", item["險種名稱"] || "",
              item["自負額"]   || "無", item["保費"] || 0,
              item["保額"]     || "", item["備註"]  || ""
            ]);
          }
        });
      }
    });
  }

  const rawPolicies = Array.from(rawPolicyMap.values());
  Logger.log('去重後保單數：' + rawPolicies.length);

  // ── 步驟4：自動家庭分組 ──
  const familyMap = buildFamilyGroups(rawPolicies, manualMap);

  // ── 步驟5：組合主表資料 ──
  const finalMainRows = [];
  rawPolicies.forEach(policy => {
    const policyNo      = String(policy["保單號碼"] || "").trim();
    const manual        = manualMap.get(policyNo) || { w: "無", x: "無", y: "無" };
    const auto          = familyMap.get(policyNo) || { famId: "無", relation: "" };
    const finalFamId    = (manual.x && manual.x !== "無") ? manual.x : auto.famId;
    const finalRelation = (manual.y && manual.y !== "無") ? manual.y : auto.relation;

    finalMainRows.push([
      0, policyNo,
      policy["險種別"] || "",
      maskIdNumber(policy["被保人"]?.["身分證號"] || ""),
      policy["被保人"]?.["姓名"]       || "",
      policy["被保人"]?.["被保人電話"] || "無",
      policy["被保人"]?.["出生年月日"] || "無",
      maskIdNumber(policy["要保人"]?.["身分證號"] || ""),
      policy["要保人"]?.["姓名"]       || "",
      policy["要保人"]?.["要保人電話"] || "無",
      policy["保費資訊"]?.["總保費"]   || "",
      policy["基本資訊"]?.["繳費方式"] || "無",
      policy["基本資訊"]?.["繳費狀況"] || "無",
      policy["保費資訊"]?.["卡號"]     || "無",
      policy["保費資訊"]?.["有效期限"] || "無",
      policy["基本資訊"]?.["生效日"]   || "無",
      policy["基本資訊"]?.["到期日"]   || "無",
      policy["基本資訊"]?.["保單狀態"] || "無",
      policy["車牌"]                   || "無",
      policy["基本資訊"]?.["Email"]    || "無",
      policy["基本資訊"]?.["地址"]     || "無",
      new Date(),
      manual.w, finalFamId, finalRelation
    ]);
  });

  // ── 步驟6：收集所有明細 ──
  const allNewDetails = [];
  detailBucket.forEach(rows => rows.forEach(r => allNewDetails.push(r)));

  // ── 步驟7：寫入試算表 ──
  if (finalMainRows.length > 0) {
    mainSheet.getRange(2, 1, finalMainRows.length, 25).setValues(finalMainRows);
  }
  if (allNewDetails.length > 0) {
    detailSheet.getRange(2, 1, allNewDetails.length, 8).setValues(allNewDetails);
  }

  reIndexAllSheets();

  // ── 步驟8：同步到原始資料 ──
  archiveToRawData(ss, finalMainRows);

  // ── 步驟9：更新客戶稱呼表（新增人員，不覆蓋已有稱呼）──
  syncNicknameSheet(ss, finalMainRows);

  Logger.log('✅ 完成。主表：' + finalMainRows.length + ' 筆，明細：' + allNewDetails.length + ' 筆。');
}

// ═══════════════════════════════════════════════════════
//  ★ 旅平卡同步（新增）
// ═══════════════════════════════════════════════════════
/**
 * 讀取 Drive 資料夾中的旅平卡 JSON（識別方式：有「憑證號碼」欄位）
 * 寫入：旅平卡主表（一筆=一位要保人）、旅平卡被保人名冊（一筆=一位被保人）
 * 去重規則：憑證號碼已存在 → 略過（不覆蓋）
 */
function syncTravelCards() {
  const ss          = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const mainSheet   = ss.getSheetByName(CONFIG.TRAVEL_MAIN_SHEET);
  const memberSheet = ss.getSheetByName(CONFIG.TRAVEL_MEMBER_SHEET);

  if (!mainSheet || !memberSheet) {
    SpreadsheetApp.getUi().alert('⚠️ 找不到「旅平卡主表」或「旅平卡被保人名冊」工作表，請先建立。');
    return;
  }

  // 讀取已存在的憑證號碼（去重用）
  const mainLastRow  = mainSheet.getLastRow();
  const existingNos  = new Set();
  if (mainLastRow > 1) {
    mainSheet.getRange(2, 2, mainLastRow - 1, 1).getValues()
      .forEach(r => { if (r[0]) existingNos.add(String(r[0]).trim()); });
  }

  const newMainRows   = [];
  const newMemberRows = [];

  const files = DriveApp.getFolderById(CONFIG.SOURCE_FOLDER_ID).getFilesByType("application/json");
  while (files.hasNext()) {
    const file = files.next();
    let jsonContent;
    try {
      jsonContent = JSON.parse(file.getBlob().getDataAsString());
    } catch(e) { continue; }

    if (!Array.isArray(jsonContent) || jsonContent.length === 0) continue;

    // 只處理旅平卡格式（有憑證號碼、沒有保單號碼）
    if (!jsonContent[0]["憑證號碼"]) continue;

    jsonContent.forEach(card => {
      const certNo = String(card["憑證號碼"] || "").trim();
      if (!certNo || existingNos.has(certNo)) return;
      existingNos.add(certNo); // 防止同一批次重複

      const applicantData = card["要保人資料"] || {};
      const memberCount   = (card["親屬被保人名冊"] || []).length;

      // ── 問題1：電話補零 → 強制文字，加 ' 前綴避免 Sheets 去掉前導零 ──
      const mobile = formatPhone(applicantData["行動電話"] || "");
      const phone  = formatPhone(applicantData["聯絡電話"] || "");

      // ── 問題2：出生日期去掉「民國」兩字，加單引號強制文字避免 Sheets 自動解析 ──
      // 不加單引號：57/03/12 → Sheets 解析為西元1957年 → Apps Script 讀回 Date(1957,...)
      // 加單引號後：'57/03/12 → Sheets 存為文字 → Apps Script 讀回字串 "57/03/12"
      const birthdayRaw = (applicantData["出生日期"] || "").replace(/^民國/, "").trim();
      const birthday    = birthdayRaw ? "'" + birthdayRaw : "";

      // ── 問題3：生效起始日格式化為 yyyy/MM/dd ──
      const effDate = formatTravelDate(card["生效起始日"] || "");

      // 主表一筆
      newMainRows.push([
        0,                                          // A: 序號
        certNo,                                     // B: 憑證號碼
        card["卡片類別"]        || "",              // C: 卡片類別
        card["要保人ID"]        || "",              // D: 要保人ID（身分證）
        card["要保人姓名"]      || "",              // E: 要保人姓名
        birthday,                                   // F: 出生日期（文字格式，不被 Sheets 解析）
        mobile,                                     // G: 行動電話（補零）
        phone,                                      // H: 聯絡電話（補零）
        applicantData["通訊地址"]  || "",           // I: 通訊地址
        applicantData["E-Mail"]    || "",           // J: Email
        effDate,                                    // K: 生效起始日（格式化）
        memberCount,                                // L: 被保人數
        new Date(),                                 // M: 同步時間
        ""                                          // N: 業務員簡稱（手動填入 K 或 N）
      ]);

      // 被保人名冊（每位被保人一筆）
      (card["親屬被保人名冊"] || []).forEach(member => {
        // 出生日期：去年齡、去民國、加單引號強制文字
        const memberBirthdayRaw = (member["出生日期"] || "")
          .replace(/\s+\d+歲$/, "")
          .replace(/^民國/, "")
          .trim();
        const memberBirthday = memberBirthdayRaw ? "'" + memberBirthdayRaw : "";
        newMemberRows.push([
          0,                                        // A: 序號
          certNo,                                   // B: 憑證號碼（外鍵）
          member["被保險人姓名"]     || "",         // C: 姓名
          member["身分證字號"]       || "",         // D: 身分證字號
          memberBirthday,                           // E: 出生日期（文字格式）
          member["與要保人關係"]     || "",         // F: 與要保人關係
          member["受益人姓名"]       || "",         // G: 受益人姓名
          member["受益人關係"]       || "",         // H: 受益人關係
          member["受益人備註"]       || ""          // I: 受益人備註
        ]);
      });
    });
  }

  if (newMainRows.length === 0) {
    Logger.log('旅平卡：無新增資料。');
    SpreadsheetApp.getUi().alert('旅平卡同步完成，無新增資料。');
    return;
  }

  // 寫入主表
  const writeMainStart = mainLastRow < 1 ? 2 : mainLastRow + 1;
  mainSheet.getRange(writeMainStart, 1, newMainRows.length, 14).setValues(newMainRows);

  // 寫入被保人名冊
  const memberLastRow    = memberSheet.getLastRow();
  const writeMemberStart = memberLastRow < 1 ? 2 : memberLastRow + 1;
  if (newMemberRows.length > 0) {
    memberSheet.getRange(writeMemberStart, 1, newMemberRows.length, 9).setValues(newMemberRows);
  }

  // 重新編排序號
  reIndexRows(mainSheet);
  reIndexRows(memberSheet);

  const msg = '✅ 旅平卡同步完成。新增主約：' + newMainRows.length + ' 筆，被保人：' + newMemberRows.length + ' 筆。';
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}

// ═══════════════════════════════════════════════════════
//  ★ 旅平卡搜尋（新增）
// ═══════════════════════════════════════════════════════
/**
 * 搜尋旅平卡主表，同時也在被保人名冊搜尋（找到被保人→回傳對應的主約）
 * 支援：要保人姓名、要保人身分證、電話、卡片類別、憑證號碼
 *       被保人姓名、被保人身分證
 */
function searchTravelCards(rawInput) {
  try {
    const ss          = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const mainSheet   = ss.getSheetByName(CONFIG.TRAVEL_MAIN_SHEET);
    const memberSheet = ss.getSheetByName(CONFIG.TRAVEL_MEMBER_SHEET);
    if (!mainSheet) return [];

    const tokens = String(rawInput || "").trim().split(/\s+/).filter(t => t.length > 0);
    if (tokens.length === 0) return [];

    const mainData   = mainSheet.getDataRange().getValues();
    const memberData = memberSheet ? memberSheet.getDataRange().getValues() : [];

    // 先從被保人名冊找出符合的憑證號碼，同時記錄命中的成員
    const matchedCertNos  = new Set();
    const matchedMembersMap = new Map(); // certNo → [{name, relation}]

    for (let i = 1; i < memberData.length; i++) {
      const mRow     = memberData[i];
      const mName    = String(mRow[2] || "");
      const mId      = String(mRow[3] || "");
      const mRelation= String(mRow[5] || "");
      const certNo   = String(mRow[1] || "");

      let allMatch = true;
      for (const tk of tokens) {
        const upper = tk.toUpperCase().replace(/[-\s]/g, '');
        let hit = mName.includes(tk) || mId.toUpperCase().includes(upper) || mRelation.includes(tk);
        if (!hit) { allMatch = false; break; }
      }
      if (allMatch && certNo) {
        matchedCertNos.add(certNo);
        if (!matchedMembersMap.has(certNo)) matchedMembersMap.set(certNo, []);
        matchedMembersMap.get(certNo).push({ name: mName, relation: mRelation });
      }
    }

    // 搜尋主表
    const results  = [];
    const addedNos = new Set();

    for (let i = 1; i < mainData.length; i++) {
      const row     = mainData[i];
      const certNo  = String(row[1] || "");
      const cardType= String(row[2] || "");
      const appId   = String(row[3] || "");
      const appName = String(row[4] || "");
      const phone1  = String(row[6] || "").replace(/[-\s]/g, '');
      const phone2  = String(row[7] || "").replace(/[-\s]/g, '');
      const address = String(row[8] || "");
      const effDate = westernDateFromSheet(row[10]);
      const count   = String(row[11] || "");

      // 從被保人名冊命中的
      if (matchedCertNos.has(certNo) && !addedNos.has(certNo)) {
        addedNos.add(certNo);
        results.push({
          certNo, cardType, appId, appName,
          phone1: row[6], phone2: row[7], address, effDate,
          memberCount: count, agentCode: String(row[13] || ""),
          matchedMembers: matchedMembersMap.get(certNo) || []
        });
        continue;
      }

      // 主表直接比對
      let allMatch = true;
      for (const tk of tokens) {
        const upper = tk.toUpperCase().replace(/[-\s]/g, '');
        let hit =
          appName.includes(tk) ||
          appId.toUpperCase().includes(upper) ||
          phone1.includes(tk.replace(/[-\s]/g,'')) ||
          phone2.includes(tk.replace(/[-\s]/g,'')) ||
          cardType.includes(tk) ||
          certNo.includes(tk) ||
          address.includes(tk);
        if (!hit) { allMatch = false; break; }
      }
      if (allMatch && !addedNos.has(certNo)) {
        addedNos.add(certNo);
        results.push({
          certNo, cardType, appId, appName,
          phone1: row[6], phone2: row[7], address, effDate,
          memberCount: count, agentCode: String(row[13] || ""),
          matchedMembers: []
        });
      }
    }
    return results;
  } catch(e) { Logger.log(e); return []; }
}

// ═══════════════════════════════════════════════════════
//  ★ 旅平卡詳情（新增）
// ═══════════════════════════════════════════════════════
function getTravelCardDetails(certNo) {
  try {
    const ss          = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const mainSheet   = ss.getSheetByName(CONFIG.TRAVEL_MAIN_SHEET);
    const memberSheet = ss.getSheetByName(CONFIG.TRAVEL_MEMBER_SHEET);
    if (!mainSheet) return null;

    const mainData = mainSheet.getDataRange().getValues();
    let mainInfo   = null;
    for (let i = 1; i < mainData.length; i++) {
      if (String(mainData[i][1]).trim() === certNo) {
        mainInfo = {
          certNo:      String(mainData[i][1]),
          cardType:    String(mainData[i][2]),
          appId:       String(mainData[i][3]),
          appName:     String(mainData[i][4]),
          birthday:    rocDateFromSheet(mainData[i][5]),   // ← 民國年轉換
          mobile:      String(mainData[i][6]),
          phone:       String(mainData[i][7]),
          address:     String(mainData[i][8]),
          email:       String(mainData[i][9]),
          effDate:     westernDateFromSheet(mainData[i][10]), // ← 西元年格式化
          memberCount: String(mainData[i][11]),
          agentCode:   String(mainData[i][13] || "")
        };
        break;
      }
    }
    if (!mainInfo) return null;

    // 讀取被保人名冊
    const members = [];
    if (memberSheet) {
      const memberData = memberSheet.getDataRange().getValues();
      for (let i = 1; i < memberData.length; i++) {
        if (String(memberData[i][1]).trim() === certNo) {
          members.push({
            name:     String(memberData[i][2]),
            idNo:     String(memberData[i][3]),
            birthday: rocDateFromSheet(memberData[i][4]), // ← 民國年轉換
            relation: String(memberData[i][5]),
            beneNamee:String(memberData[i][6]),
            beneRel:  String(memberData[i][7]),
            beneNote: String(memberData[i][8])
          });
        }
      }
    }
    return { main: mainInfo, members };
  } catch(e) { Logger.log(e); return null; }
}

// ═══════════════════════════════════════════════════════
//  家庭自動分組（Union-Find）
// ═══════════════════════════════════════════════════════
function buildFamilyGroups(policies, manualMap) {
  const parent = new Map();
  function find(x) {
    if (!parent.has(x)) parent.set(x, x);
    if (parent.get(x) !== x) parent.set(x, find(parent.get(x)));
    return parent.get(x);
  }
  function union(a, b) {
    const ra = find(a), rb = find(b);
    if (ra !== rb) parent.set(rb, ra);
  }

  const idIndex   = new Map();
  const nameIndex = new Map();
  const addrIndex = new Map();

  policies.forEach(p => {
    const no    = String(p["保單號碼"] || "").trim();
    if (!no) return;
    parent.set(no, no);
    const idRaw = String(p["要保人"]?.["身分證號"] || "").trim();
    const name  = String(p["要保人"]?.["姓名"]     || "").trim();
    const addr  = normalizeAddress(p["基本資訊"]?.["地址"] || "");

    if (idRaw && idRaw !== "無") { if (!idIndex.has(idRaw))  idIndex.set(idRaw, []);  idIndex.get(idRaw).push(no); }
    if (name  && name  !== "無") { if (!nameIndex.has(name)) nameIndex.set(name, []); nameIndex.get(name).push(no); }
    if (addr  && addr  !== "無") { if (!addrIndex.has(addr)) addrIndex.set(addr, []); addrIndex.get(addr).push(no); }
  });

  idIndex.forEach(nos   => { for (let i = 1; i < nos.length; i++) union(nos[0], nos[i]); });
  nameIndex.forEach(nos => { for (let i = 1; i < nos.length; i++) union(nos[0], nos[i]); });
  addrIndex.forEach(nos => { for (let i = 1; i < nos.length; i++) union(nos[0], nos[i]); });

  const policyMap   = new Map();
  const groupMember = new Map();
  policies.forEach(p => { const no = String(p["保單號碼"] || "").trim(); if (no) policyMap.set(no, p); });
  policies.forEach(p => {
    const no = String(p["保單號碼"] || "").trim();
    if (!no) return;
    const root = find(no);
    if (!groupMember.has(root)) groupMember.set(root, []);
    groupMember.get(root).push(p);
  });

  const rootFamId       = new Map();
  const rootApplicantId = new Map();

  groupMember.forEach((group, root) => {
    group.sort((a, b) => {
      const da = robustParseDate(a["基本資訊"]?.["生效日"] || "") || new Date(9999,0);
      const db = robustParseDate(b["基本資訊"]?.["生效日"] || "") || new Date(9999,0);
      return da - db;
    });
    let famId = "無", repId = "";
    for (const p of group) {
      const idRaw = String(p["要保人"]?.["身分證號"] || "").trim();
      const name  = String(p["要保人"]?.["姓名"]     || "").trim();
      const no    = String(p["保單號碼"]             || "").trim();
      const fid   = generateFamId(idRaw, name, no);
      if (fid !== "無") { famId = fid; repId = idRaw; break; }
    }
    rootFamId.set(root, famId);
    rootApplicantId.set(root, repId);
  });

  const result = new Map();
  policies.forEach(p => {
    const no     = String(p["保單號碼"] || "").trim();
    if (!no) return;
    const root   = find(no);
    const famId  = rootFamId.get(root) || "無";
    const repId  = rootApplicantId.get(root) || "";
    const thisId = String(p["要保人"]?.["身分證號"] || "").trim();
    result.set(no, { famId, relation: (repId && thisId && repId === thisId) ? "本人" : "" });
  });
  return result;
}

// ═══════════════════════════════════════════════════════
//  FAM-ID 生成
// ═══════════════════════════════════════════════════════
function generateFamId(idRaw, applicantName, policyNo) {
  const id = (idRaw || "").trim();
  if (/^[A-Za-z][0-9]{9}$/.test(id)) return "FAM-" + id.substring(0, 2).toUpperCase() + id.substring(6);
  if (/^[0-9]{8}$/.test(id)) return "FAM-" + id;
  if (id.length >= 6) return "FAM-" + id.substring(0, 2).toUpperCase() + id.slice(-4);
  const name = (applicantName || "").replace(/\s/g, "").substring(0, 2);
  const tail = (policyNo || "").slice(-4);
  if (name) return "FAM-" + name + tail;
  return "無";
}

function normalizeAddress(addr) {
  if (!addr || addr === "無") return "無";
  return String(addr).replace(/[\s　]/g, "").replace(/[，。、]/g, "").trim();
}

// ═══════════════════════════════════════════════════════
//  產險智能搜尋
// ═══════════════════════════════════════════════════════
function searchPolicies(rawInput) {
  try {
    const ss   = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const data = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME).getDataRange().getValues();

    const tokens = String(rawInput || "").trim().split(/\s+/).filter(t => t.length > 0);
    if (tokens.length === 0) return [];

    function parseDateToken(t) {
      const clean = t.replace(/\D/g, '');
      if (clean.length < 4) return null;
      if (clean.length === 5) {
        const yy = parseInt(clean.substring(0, 3));
        const mm = String(parseInt(clean.substring(3)));
        return { roc: yy + '/' + mm, west: (yy + 1911) + '/' + mm };
      }
      if (/^\d{3}\/\d{1,2}$/.test(t)) {
        const parts = t.split('/');
        const yy = parseInt(parts[0]);
        const mm = String(parseInt(parts[1]));
        return { roc: yy + '/' + mm, west: (yy + 1911) + '/' + mm };
      }
      return null;
    }

    const results = [];
    for (let i = 1; i < data.length; i++) {
      const row      = data[i];
      const insured  = safeString(row[4]);
      const applicant= safeString(row[8]);
      const expiry   = safeString(row[16]);
      const plate    = safeString(row[18]).toUpperCase().replace(/[-\s]/g, '');
      const insType  = safeString(row[2]);
      const policyNo = safeString(row[1]).toUpperCase();
      const phone1   = safeString(row[5]).replace(/[-\s]/g,'');
      const phone2   = safeString(row[9]).replace(/[-\s]/g,'');
      const familyId = safeString(row[23]);
      const address  = safeString(row[20]);

      let allMatch = true;
      for (const tk of tokens) {
        const upper    = tk.toUpperCase().replace(/[-\s]/g, '');
        const dateInfo = parseDateToken(tk);
        let hit        = false;

        if (!hit) hit = insured.includes(tk) || applicant.includes(tk);
        if (!hit) hit = plate !== '無' && plate.includes(upper);
        if (!hit) hit = phone1.includes(tk.replace(/[-\s]/g,'')) || phone2.includes(tk.replace(/[-\s]/g,''));
        if (!hit) hit = insType.includes(tk);
        if (!hit) hit = policyNo.includes(upper);
        if (!hit) hit = familyId.toUpperCase().includes(upper);
        if (!hit && dateInfo) hit = expiry.includes(dateInfo.roc) || expiry.includes(dateInfo.west);

        if (!hit) { allMatch = false; break; }
      }

      if (allMatch) {
        results.push({
          policyNo:      safeString(row[1]),
          type:          safeString(row[2]),
          maskedId:      safeString(row[3]),
          name:          insured,
          birthday:      safeString(row[6], true),
          effectiveDate: safeString(row[15]),          // P欄：生效日（新增）
          expiry:        expiry,
          status:        safeString(row[17]),
          plateNo:       safeString(row[18]),
          premium:       safeString(row[10]),          // K欄：總保費（新增）
          agentCode:     safeString(row[22]),
          applicantName: applicant,
          applicantPhone:safeString(row[9]),
          address:       safeString(row[20])
        });
      }
    }
    return results;
  } catch(e) { Logger.log(e); return []; }
}

// ═══════════════════════════════════════════════════════
//  保單詳情
// ═══════════════════════════════════════════════════════
function getPolicyDetails(policyNo) {
  const ss       = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const mainData = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME).getDataRange().getValues();
  let mainInfo   = null;

  for (let i = 1; i < mainData.length; i++) {
    if (safeString(mainData[i][1]) === policyNo) {
      mainInfo = {
        policyNo:      safeString(mainData[i][1]),
        type:          safeString(mainData[i][2]),
        insuredName:   safeString(mainData[i][4]),
        maskedId:      safeString(mainData[i][3]),
        birthday:      safeString(mainData[i][6], true),
        premium:       safeString(mainData[i][10]),
        plateNo:       safeString(mainData[i][18]),
        expiryDate:    safeString(mainData[i][16]),
        effectiveDate: safeString(mainData[i][15]),
        agentCode:     safeString(mainData[i][22]),
        status:        safeString(mainData[i][17]),
        familyId:      safeString(mainData[i][23]),
        relation:      safeString(mainData[i][24])
      };
      break;
    }
  }
  if (!mainInfo) return null;

  const detailData = ss.getSheetByName(CONFIG.DETAIL_SHEET_NAME).getDataRange().getValues();
  const coverages  = detailData
    .filter(row => safeString(row[1]) === policyNo)
    .map(row => ({ name: safeString(row[3]), amount: safeString(row[6]), premium: safeString(row[5]) }));

  return { main: mainInfo, details: coverages };
}

// ═══════════════════════════════════════════════════════
//  ★ 旅平卡即時查詢（方案二：保單詳情頁即時提示）
// ═══════════════════════════════════════════════════════
/**
 * 依被保人身分證號查詢是否在旅平卡被保人名冊中
 * 供前端 renderModal 開啟保單詳情時呼叫
 *
 * 回傳：
 *   { hasTravelCard: true,  role: "主保人", cardType: "快樂旅平卡", certNo: "241800006718", appName: "陳麗敏" }
 *   { hasTravelCard: false }
 */
function checkTravelCoverage(insuredId) {
  try {
    if (!insuredId || insuredId === "無") return { hasTravelCard: false };

    const ss          = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const memberSheet = ss.getSheetByName(CONFIG.TRAVEL_MEMBER_SHEET);
    const mainSheet   = ss.getSheetByName(CONFIG.TRAVEL_MAIN_SHEET);
    if (!memberSheet || !mainSheet) return { hasTravelCard: false };

    // 建立旅平卡主表 certNo → { cardType, appName }
    const mainData   = mainSheet.getDataRange().getValues();
    const certInfoMap = new Map();
    for (let i = 1; i < mainData.length; i++) {
      const certNo   = String(mainData[i][1] || "").trim();
      const cardType = String(mainData[i][2] || "").trim();
      const appName  = String(mainData[i][4] || "").trim();
      if (certNo) certInfoMap.set(certNo, { cardType, appName });
    }

    // 搜尋被保人名冊
    const memberData = memberSheet.getDataRange().getValues();
    for (let i = 1; i < memberData.length; i++) {
      const idNo    = String(memberData[i][3] || "").trim(); // D欄：身分證
      if (idNo !== insuredId.trim()) continue;

      const certNo   = String(memberData[i][1] || "").trim(); // B欄：憑證號碼
      const relation = String(memberData[i][5] || "").trim(); // F欄：關係
      const certInfo = certInfoMap.get(certNo) || { cardType: "旅平卡", appName: "—" };

      return {
        hasTravelCard: true,
        role:     relation === "本人" ? "主保人" : "親屬（" + relation + "）",
        cardType: certInfo.cardType,
        certNo:   certNo,
        appName:  certInfo.appName
      };
    }

    return { hasTravelCard: false };
  } catch(e) {
    Logger.log('checkTravelCoverage error: ' + e);
    return { hasTravelCard: false };
  }
}

// ═══════════════════════════════════════════════════════
//  ★ 未申辦旅平卡名單（從比對旅平卡覆蓋表讀取）
// ═══════════════════════════════════════════════════════
/**
 * 讀取「比對旅平卡覆蓋表」中 H欄=「❌ 無」的資料
 * 回傳：[{ name, idNo, insType, expiry, applicant }, ...]
 */
function getUncoveredList() {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.COVERAGE_SHEET_NAME);
    if (!sheet) return { error: '請先執行「比對旅平卡覆蓋率」產出資料', list: [] };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { error: '比對旅平卡覆蓋表尚無資料，請先執行比對', list: [] };

    const data   = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const list   = [];
    const updateTime = sheet.getRange(2, 1).getValue(); // 用第一筆判斷是否有資料

    data.forEach(row => {
      const status = String(row[7] || "").trim(); // H欄：旅平卡狀態
      if (status === "❌ 無") {
        list.push({
          name:      String(row[1] || "").trim(),       // B欄：被保人姓名
          idNo:      String(row[2] || "").trim(),       // C欄：身分證
          insType:   String(row[3] || "").trim(),       // D欄：險種
          status:    String(row[4] || "").trim(),       // E欄：保單狀態
          expiry:    safeString(row[5]),                // F欄：到期日（用 safeString 轉換）
          applicant: String(row[6] || "").trim()        // G欄：要保人姓名
        });
      }
    });

    return { error: null, total: list.length, list };
  } catch(e) {
    Logger.log('getUncoveredList error: ' + e);
    return { error: '讀取失敗：' + e.message, list: [] };
  }
}

// ═══════════════════════════════════════════════════════
//  ★ 近期到期提醒名單（5 天內，Z 欄非已續保/不續保）
// ═══════════════════════════════════════════════════════
/**
 * 讀取保單主表，回傳 5 天（含）內到期且尚未完成的保單
 *
 * 過濾邏輯：
 * 1. Z 欄為「已續保」或「不續保」→ 跳過
 * 2. 同車牌 + 同險種，已有到期日 > 今天的有效保單 → 表示已續保，跳過
 * 3. 空白 / 已聯繫 → 顯示在名單
 *
 * 回傳依剩餘天數由少到多排序
 */
// ═══════════════════════════════════════════════════════
//  ★ 到期提醒名單（通用）
// ═══════════════════════════════════════════════════════
/**
 * 通用函式：取得指定天數範圍內到期的保單
 * startDay: 幾天後開始（含），0 = 今天
 * endDay:   幾天後結束（含）
 * 同時套用：Z 欄狀態過濾 + 車牌+險種自動續保識別
 */
function getExpiringListByRange(startDay, endDay) {
  try {
    const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet   = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const data       = sheet.getRange(1, 1, lastRow, 26).getValues();
    const today      = new Date(); today.setHours(0, 0, 0, 0);
    const rangeStart = new Date(today); rangeStart.setDate(today.getDate() + startDay);
    const rangeEnd   = new Date(today); rangeEnd.setDate(today.getDate() + endDay);

    // ── 第一步：建立「有效保單」集合（車牌+險種 → 最遠到期日）──
    const activePolicyMap = new Map();
    for (let i = 1; i < data.length; i++) {
      const row     = data[i];
      const plate   = String(row[18] || "").trim().toUpperCase().replace(/[-\s]/g, '');
      const insType = String(row[2]  || "").trim();
      const expDate = robustParseDate(safeString(row[16]));
      if (!plate || plate === '無' || !expDate || isNaN(expDate.getTime())) continue;
      if (expDate <= today) continue;
      const key = plate + '|' + insType;
      if (!activePolicyMap.has(key) || expDate > activePolicyMap.get(key)) {
        activePolicyMap.set(key, expDate);
      }
    }

    // ── 第二步：套用過濾邏輯 ──
    const results = [];
    for (let i = 1; i < data.length; i++) {
      const row      = data[i];
      const policyNo = String(row[1] || "").trim();
      if (!policyNo) continue;

      const renewStatus = String(row[25] || "").trim();
      if (renewStatus === '已續保' || renewStatus === '不續保' ||
          renewStatus === '已過戶' || renewStatus === '已報廢') continue;

      const expiryDate = robustParseDate(safeString(row[16]));
      if (!expiryDate || isNaN(expiryDate.getTime())) continue;
      if (expiryDate < rangeStart || expiryDate > rangeEnd) continue;

      const plate   = String(row[18] || "").trim().toUpperCase().replace(/[-\s]/g, '');
      const insType = String(row[2]  || "").trim();
      if (plate && plate !== '無') {
        const key          = plate + '|' + insType;
        const latestExpiry = activePolicyMap.get(key);
        if (latestExpiry && latestExpiry > expiryDate) continue;
      }

      const daysLeft = Math.ceil((expiryDate - today) / (1000 * 60 * 60 * 24));
      results.push({
        policyNo:      policyNo,
        name:          String(row[4]  || "").trim(),
        type:          insType,
        plateNo:       safeString(row[18]),
        expiry:        safeString(row[16]),
        effectiveDate: safeString(row[15]),
        premium:       safeString(row[10]),
        maskedId:      String(row[3]  || "").trim(),
        agentCode:     String(row[22] || "").trim(),
        renewStatus:   renewStatus,
        daysLeft:      daysLeft
      });
    }

    results.sort((a, b) => a.daysLeft - b.daysLeft);
    return results;
  } catch(e) {
    Logger.log('getExpiringListByRange error: ' + e);
    return [];
  }
}

// 5 天內到期（緊急提醒）
function getExpiringList() {
  return getExpiringListByRange(0, 5);
}

// 45～51 天後到期（提前預警）
function getEarlyWarningList() {
  return getExpiringListByRange(45, 51);
}

// 一次回傳兩個結果，只需一個後端請求
function getExpiringLists() {
  return {
    urgent:  getExpiringListByRange(0,  5),
    warning: getExpiringListByRange(45, 51)
  };
}

// ═══════════════════════════════════════════════════════
//  ★ 更新保單 Z 欄續保狀態
// ═══════════════════════════════════════════════════════
/**
 * 依保單號碼找到對應列，更新 Z 欄（第 26 欄）的續保狀態
 * status 可填：已聯繫、已續保、不續保、（空字串=清除）
 */
function updateRenewalStatus(policyNo, status) {
  try {
    const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet   = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, msg: '無資料' };

    const colB = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // B 欄：保單號碼
    for (let i = 0; i < colB.length; i++) {
      if (String(colB[i][0]).trim() === policyNo) {
        sheet.getRange(i + 2, 26).setValue(status); // Z 欄 = 第 26 欄
        return { success: true, policyNo, status };
      }
    }
    return { success: false, msg: '找不到保單號碼：' + policyNo };
  } catch(e) {
    Logger.log('updateRenewalStatus error: ' + e);
    return { success: false, msg: e.message };
  }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('保險特助系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ═══════════════════════════════════════════════════════
//  家庭彙整 API
// ═══════════════════════════════════════════════════════
function getFamilyReport(familyId) {
  if (!familyId || familyId === "無") return null;

  const ss       = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const mainData = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME).getDataRange().getValues();
  const cfg      = getSystemConfig();
  const today    = new Date(); today.setHours(0, 0, 0, 0);
  const urgentDays  = parseInt(cfg['緊急天數']  || '30');
  const warningDays = parseInt(cfg['警示天數'] || '60');

  const memberMap = new Map();

  for (let i = 1; i < mainData.length; i++) {
    if (safeString(mainData[i][23]) !== familyId) continue;

    const name      = safeString(mainData[i][4]);
    const maskedId  = safeString(mainData[i][3]); // D欄：被保人身分證（新增）
    const relation  = safeString(mainData[i][24]);
    const status    = safeString(mainData[i][17]);
    const expiryStr = safeString(mainData[i][16]);
    const policyNo  = safeString(mainData[i][1]);
    const type      = safeString(mainData[i][2]);
    const plate     = safeString(mainData[i][18]);
    const premium   = safeString(mainData[i][10]);

    const expiryDate = robustParseDate(expiryStr);
    let daysLeft = null, urgency = "expired";
    if (expiryDate && !isNaN(expiryDate.getTime())) {
      daysLeft = Math.ceil((expiryDate - today) / (1000 * 60 * 60 * 24));
      if (status === "退保" || daysLeft < 0) urgency = "expired";
      else if (daysLeft <= urgentDays)        urgency = "urgent";
      else if (daysLeft <= warningDays)       urgency = "warning";
      else                                    urgency = "ok";
    }

    const key = name + "|" + relation;
    if (!memberMap.has(key)) memberMap.set(key, { name, maskedId, relation, policies: [] });
    memberMap.get(key).policies.push({ policyNo, type, plate, expiry: expiryStr, premium, status, daysLeft, urgency });
  }

  if (memberMap.size === 0) return null;

  const relationOrder = ["本人","配偶","父母","父子","母子","父女","母女","兄弟","姊妹","兄妹","祖父母"];
  const memberGroups  = Array.from(memberMap.values()).sort((a, b) => {
    const ai = relationOrder.indexOf(a.relation), bi = relationOrder.indexOf(b.relation);
    if (ai === -1 && bi === -1) return 0;
    if (ai === -1) return 1; if (bi === -1) return -1;
    return ai - bi;
  });

  memberGroups.forEach(m => m.policies.sort((a, b) => {
    if (a.daysLeft === null) return 1;
    if (b.daysLeft === null) return -1;
    return a.daysLeft - b.daysLeft;
  }));

  const allPolicies = memberGroups.flatMap(m => m.policies);
  const summary = {
    total:   allPolicies.length,
    urgent:  allPolicies.filter(p => p.urgency === "urgent").length,
    warning: allPolicies.filter(p => p.urgency === "warning").length,
    ok:      allPolicies.filter(p => p.urgency === "ok").length,
    expired: allPolicies.filter(p => p.urgency === "expired").length
  };

  return {
    id:          familyId,
    memberGroups,
    summary,
    analysis:    generatePeiLinAnalysis(memberGroups, summary, cfg['顧問姓名'] || 'Pei-lin'),
    advisorName: cfg['顧問姓名'] || 'Pei-lin'
  };
}

function generatePeiLinAnalysis(memberGroups, summary, advisorName) {
  const name        = advisorName || 'Pei-lin';
  const allPolicies = memberGroups.flatMap(m => m.policies);
  const allTypes    = allPolicies.map(p => p.type);
  const hasFire     = allTypes.some(t => t.includes("火"));
  const hasHealth   = allTypes.some(t => t.includes("健康") || t.includes("傷害"));

  let comments = "🌸 " + name + " 的暖心分析：\n\n";
  comments += "這個家庭共有 " + memberGroups.length + " 位成員、" + summary.total + " 張有效保單，";

  if (summary.urgent > 0) {
    comments += "其中有 " + summary.urgent + " 張保單即將在 30 天內到期，請盡快安排續約！🔴\n";
  } else if (summary.warning > 0) {
    comments += "其中有 " + summary.warning + " 張保單將在 60 天內到期，建議提前與客戶確認續約意願。🟠\n";
  } else {
    comments += "目前所有保單狀況良好。\n";
  }
  if (!hasFire)   comments += "\n住宅保障目前尚未納入系統，家庭資產防護有缺口，住宅火險可防護住宅資產。☔";
  if (!hasHealth) comments += "\n家庭成員目前沒有健康傷害險，人身保障值得補強。💊";
  comments += "\n\n" + name + " 會持續為您全家守候！✨";
  return comments;
}

// ═══════════════════════════════════════════════════════
//  每週續保提醒
// ═══════════════════════════════════════════════════════
function checkAndSendWeeklyReminders() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const data      = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME).getDataRange().getValues();
  const today     = new Date();
  const startDate = new Date(); startDate.setDate(today.getDate() + 45); startDate.setHours(0,0,0,0);
  const endDate   = new Date(); endDate.setDate(today.getDate() + 51);   endDate.setHours(23,59,59,999);

  const reminderList = [];
  for (let i = 1; i < data.length; i++) {
    const expiryDate = robustParseDate(data[i][16]);
    if (expiryDate && expiryDate >= startDate && expiryDate <= endDate) {
      reminderList.push({
        name:   data[i][4],
        expiry: Utilities.formatDate(expiryDate, "GMT+8", "yyyy/MM/dd"),
        type:   data[i][2],
        plate:  data[i][18]
      });
    }
  }
  if (reminderList.length > 0) {
    const cfg  = getSystemConfig();
    const name = cfg['顧問姓名'] || 'Baker';
    let body   = name + " 您好，以下為本週續保追蹤清單 (45-51天後到期)：\n\n";
    reminderList.forEach((item, idx) => {
      body += (idx+1) + ". " + item.name + " (" + item.expiry + ") - " + item.type + " [" + item.plate + "]\n";
    });
    MailApp.sendEmail(Session.getActiveUser().getEmail(), "📅 本週續保提醒清單", body);
  }
}

// ═══════════════════════════════════════════════════════
//  清理過期保單
// ═══════════════════════════════════════════════════════
function cleanupExpiredPolicies() {
  const ui          = SpreadsheetApp.getUi();
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet   = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  const detailSheet = ss.getSheetByName(CONFIG.DETAIL_SHEET_NAME);
  const threshold   = new Date(); threshold.setMonth(threshold.getMonth() - 3);

  const mainData       = mainSheet.getDataRange().getValues();
  const activeMainRows = [mainData[0]];
  const deletedNos     = new Set();

  for (let i = 1; i < mainData.length; i++) {
    const exp = robustParseDate(mainData[i][16]);
    if (!exp || isNaN(exp.getTime()) || exp >= threshold) activeMainRows.push(mainData[i]);
    else deletedNos.add(String(mainData[i][1]));
  }

  if (deletedNos.size === 0) return ui.alert('📝 目前沒有符合清理條件的保單。');
  const confirm = ui.alert('❓ 確認清理', `預計刪除 ${deletedNos.size} 筆過期保單，確定？`, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const detailData       = detailSheet.getDataRange().getValues();
  const activeDetailRows = detailData.filter((row, idx) => idx === 0 || !deletedNos.has(String(row[1])));

  mainSheet.getRange(2, 1, mainSheet.getMaxRows() - 1, 25).clear();
  mainSheet.getRange(1, 1, activeMainRows.length, 25).setValues(activeMainRows);
  detailSheet.getRange(2, 1, detailSheet.getMaxRows() - 1, 8).clear();
  detailSheet.getRange(1, 1, activeDetailRows.length, 8).setValues(activeDetailRows);

  reIndexAllSheets();
  ui.alert('✅ 清理完成。');
}

// ═══════════════════════════════════════════════════════
//  序號編排
// ═══════════════════════════════════════════════════════
function reIndexAllSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  reIndexRows(ss.getSheetByName(CONFIG.MAIN_SHEET_NAME));
  reIndexRows(ss.getSheetByName(CONFIG.DETAIL_SHEET_NAME));
}

function reIndexRows(sheet) {
  if (!sheet) return;
  const last = sheet.getLastRow();
  if (last < 2) return;
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).clearContent();
  const bVals = sheet.getRange("B2:B" + last).getValues();
  const count = bVals.filter(r => r[0] !== "").length;
  if (count > 0) sheet.getRange(2, 1, count, 1).setValues(Array.from({length: count}, (_, i) => [i + 1]));
}

// ═══════════════════════════════════════════════════════
//  原始資料備份
// ═══════════════════════════════════════════════════════
function archiveToRawData(ss, finalMainRows) {
  const archiveSheet = ss.getSheetByName(CONFIG.ARCHIVE_SHEET_NAME);
  if (!archiveSheet) { Logger.log('⚠️ 找不到「原始資料」工作表，略過備份。'); return; }

  const lastRow     = archiveSheet.getLastRow();
  const existingNos = new Set();
  if (lastRow > 1) {
    archiveSheet.getRange(2, 2, lastRow - 1, 1).getValues()
      .forEach(row => { if (row[0]) existingNos.add(String(row[0]).trim()); });
  }

  const newRows = finalMainRows.filter(row => {
    const policyNo = String(row[1] || "").trim();
    return policyNo && !existingNos.has(policyNo);
  });

  if (newRows.length === 0) { Logger.log('原始資料：無新增保單。'); return; }

  archiveSheet.getRange(lastRow + 1, 1, newRows.length, 25).setValues(newRows);
  reIndexRows(archiveSheet);
  Logger.log('原始資料：新增 ' + newRows.length + ' 筆。');
}

// ═══════════════════════════════════════════════════════
//  ★ 客戶稱呼表管理
// ═══════════════════════════════════════════════════════
/**
 * 批次同步後自動更新客戶稱呼表
 * 規則：
 *   - 從保單主表抓取所有被保人、要保人（身分證號+姓名）
 *   - 去重（同一身分證只留一筆）
 *   - 已在稱呼表的不覆蓋（保留手動填入的稱呼）
 *   - 新人員追加到最後一列
 * 工作表格式：A=身分證號, B=姓名, C=稱呼（手動填入）
 */
function syncNicknameSheet(ss, finalMainRows) {
  let nicknameSheet = ss.getSheetByName(CONFIG.NICKNAME_SHEET_NAME);
  if (!nicknameSheet) {
    Logger.log('⚠️ 找不到「客戶稱呼表」工作表，略過更新。');
    return;
  }

  // 讀取已存在的身分證號
  const lastRow    = nicknameSheet.getLastRow();
  const existingIds = new Set();
  if (lastRow > 1) {
    nicknameSheet.getRange(2, 1, lastRow - 1, 1).getValues()
      .forEach(row => { if (row[0]) existingIds.add(String(row[0]).trim()); });
  }

  // 從保單主表收集所有人員（被保人 D+E欄，要保人 H+I欄）
  const personMap = new Map(); // idNo → name
  finalMainRows.forEach(row => {
    const insuredId   = String(row[3] || "").trim();  // D欄：被保人身分證
    const insuredName = String(row[4] || "").trim();  // E欄：被保人姓名
    const appId       = String(row[7] || "").trim();  // H欄：要保人身分證
    const appName     = String(row[8] || "").trim();  // I欄：要保人姓名

    if (insuredId && insuredId !== '無' && insuredName && insuredName !== '無') {
      if (!personMap.has(insuredId)) personMap.set(insuredId, insuredName);
    }
    if (appId && appId !== '無' && appName && appName !== '無') {
      if (!personMap.has(appId)) personMap.set(appId, appName);
    }
  });

  // 過濾出新人員
  const newRows = [];
  personMap.forEach((name, idNo) => {
    if (!existingIds.has(idNo)) {
      newRows.push([idNo, name, '']); // C欄稱呼留空，等手動填入
    }
  });

  if (newRows.length === 0) {
    Logger.log('客戶稱呼表：無新增人員。');
    return;
  }

  // 追加到最後
  const writeStart = lastRow < 1 ? 2 : lastRow + 1;
  nicknameSheet.getRange(writeStart, 1, newRows.length, 3).setValues(newRows);
  Logger.log('客戶稱呼表：新增 ' + newRows.length + ' 筆人員。');
}

/**
 * 依身分證號查詢稱呼
 * 回傳：{ nickname: '建明哥', name: '李建明' } 或 { nickname: '', name: '李建明' }
 */
function getCustomerNickname(insuredId) {
  try {
    if (!insuredId || insuredId === '無') return { nickname: '', name: '' };
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.NICKNAME_SHEET_NAME);
    if (!sheet) return { nickname: '', name: '' };
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === insuredId.trim()) {
        return {
          nickname: String(data[i][2] || "").trim(), // C欄：稱呼
          name:     String(data[i][1] || "").trim()  // B欄：姓名
        };
      }
    }
    return { nickname: '', name: '' };
  } catch(e) {
    Logger.log('getCustomerNickname error: ' + e);
    return { nickname: '', name: '' };
  }
}

/**
 * 儲存客戶稱呼（傳送時手動輸入後自動存回）
 * insuredId: 身分證號, nickname: 稱呼
 */
function saveCustomerNickname(insuredId, nickname) {
  try {
    if (!insuredId || !nickname) return { success: false };
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.NICKNAME_SHEET_NAME);
    if (!sheet) return { success: false };
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === insuredId.trim()) {
        sheet.getRange(i + 1, 3).setValue(nickname); // C欄寫入稱呼
        return { success: true };
      }
    }
    return { success: false, msg: '找不到該身分證號' };
  } catch(e) {
    Logger.log('saveCustomerNickname error: ' + e);
    return { success: false };
  }
}

// ═══════════════════════════════════════════════════════
//  輔助函式
// ═══════════════════════════════════════════════════════
function robustParseDate(input) {
  if (!input) return null;
  if (input instanceof Date) {
    return input.getFullYear() < 1911
      ? new Date(input.getFullYear() + 1911, input.getMonth(), input.getDate())
      : input;
  }
  const parts = String(input).split(/[\/\-\.]/);
  if (parts.length === 3) {
    let y = parseInt(parts[0]);
    if (y < 1911) y += 1911;
    return new Date(y, parseInt(parts[1]) - 1, parseInt(parts[2]));
  }
  return new Date(input);
}

function safeString(v, isBirthday = false) {
  if (!v || v === "") return "無";
  if (typeof v === 'string') return v;
  if (v instanceof Date) {
    let y = v.getFullYear(), m = v.getMonth() + 1, d = v.getDate();
    if (isBirthday) {
      if (y > 1911 && y < 2000) y += 11;
      else if (y <= 1911) y += 1911;
      return y + "/" + ("0"+m).slice(-2) + "/" + ("0"+d).slice(-2);
    }
    return ((y > 1911) ? (y - 1911) : y) + "/" + m + "/" + d;
  }
  return String(v);
}

function maskIdNumber(id) { if (!id) return "無"; return String(id); }

/**
 * 民國年日期轉換（被保人出生日期用）
 *
 * 測試結果確認（2026/05/16）：
 *   syncTravelCards 寫入「57/03/12」→ Sheets 自動解析為西元1957年
 *   Apps Script 讀回：getFullYear() = 1957
 *   正確民國年 = 1957 - 1900 = 57 ✅
 *
 *   同理：75/05/12 → 1975 → 1975-1900 = 75 ✅
 *         27/08/24 → 2027 → 2027-1900 = 127（民國127年）✅
 *         85/06/07 → 1985 → 1985-1900 = 85 ✅
 *
 * 結論：所有從 syncTravelCards 寫入的民國出生日期
 *        Sheets 都以「民國年+1900」的西元年儲存
 *        所以 getFullYear() - 1900 = 民國年
 */
function rocDateFromSheet(val) {
  if (!val && val !== 0) return "—";

  if (val instanceof Date) {
    const y = val.getFullYear();
    if (y < 1900) return String(val); // 異常值保護
    const rocYear = y - 1900;         // ✅ 正確換算
    const m = ("0" + (val.getMonth() + 1)).slice(-2);
    const d = ("0" + val.getDate()).slice(-2);
    return rocYear + "/" + m + "/" + d;
  }

  const s = String(val).trim();
  if (!s || s === "無") return "—";

  // 字串：Sheets 沒有自動轉換（文字格式儲存格）→ 直接回傳
  if (s.indexOf("/") !== -1) {
    const parts = s.split("/");
    if (parts.length === 3) {
      const y = parseInt(parts[0]);
      const m = ("0" + parseInt(parts[1])).slice(-2);
      const d = ("0" + parseInt(parts[2])).slice(-2);
      return y + "/" + m + "/" + d;
    }
  }
  return s;
}

/**
 * 西元年日期格式化（生效起始日用）
 * 輸入：Date 物件 或 "2011/09/01" 字串
 * 輸出："2011/09/01"
 */
function westernDateFromSheet(val) {
  if (!val && val !== 0) return "—";
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = ("0" + (val.getMonth() + 1)).slice(-2);
    const d = ("0" + val.getDate()).slice(-2);
    return y + "/" + m + "/" + d;
  }
  const s = String(val).trim();
  if (!s || s === "無") return "—";
  // 已是 yyyy/MM/dd 格式
  if (/^\d{4}\/\d{2}\/\d{2}$/.test(s)) return s;
  return s;
}

/**
 * 電話補零：確保以 0 開頭的電話號碼不被 Sheets 吃掉前導零
 * 寫入時在前面加單引號，強制 Sheets 視為文字
 */
function formatPhone(phone) {
  if (!phone) return "";
  const clean = String(phone).trim();
  if (!clean) return "";
  // 純數字且以 0 開頭 → 加單引號強制文字
  if (/^\d+$/.test(clean) && clean.startsWith('0')) return "'" + clean;
  return clean;
}

/**
 * 旅平卡生效日期格式化
 * 輸入：「2011年08月15日」→ 輸出：「2011/08/15」
 * 輸入：已是 Date 物件 → 輸出：「yyyy/MM/dd」
 */
function formatTravelDate(input) {
  if (!input) return "";
  // 處理「2011年08月15日」格式
  const match = String(input).match(/(\d{4})年(\d{2})月(\d{2})日/);
  if (match) return match[1] + "/" + match[2] + "/" + match[3];
  // 處理 Date 物件
  if (input instanceof Date) {
    const y = input.getFullYear();
    const m = ("0" + (input.getMonth() + 1)).slice(-2);
    const d = ("0" + input.getDate()).slice(-2);
    return y + "/" + m + "/" + d;
  }
  return String(input);
}

function filterValidDetailsImproved(details) {
  const garbage = ["繳費狀況","住址","繳費方式","卡號/帳號","富壽同意註記","生效日","到期日","出生年月日","同意註記"];
  return details.filter(item => {
    const code = String(item["險種代號"] || "");
    return code.length > 0 && code.length < 15 && !garbage.some(k => code.includes(k));
  });
}

// ═══════════════════════════════════════════════════════
//  ★ 旅平卡覆蓋率比對（方案一：Sheets 報告）
// ═══════════════════════════════════════════════════════
/**
 * 比對產險主表的所有被保人身分證號
 * 與旅平卡被保人名冊的身分證號
 * 結果寫入「比對旅平卡覆蓋表」工作表
 *
 * 報告欄位：
 *   A: 序號
 *   B: 被保人姓名
 *   C: 被保人身分證
 *   D: 險種
 *   E: 保單狀態
 *   F: 到期日
 *   G: 要保人姓名
 *   H: 旅平卡狀態（✅ 有 / ❌ 無）
 *   I: 旅平卡身份（主保人 / 親屬：關係）
 *   J: 旅平卡憑證號碼
 *   K: 旅平卡要保人
 */
function compareTravelCoverage() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // ── 讀取旅平卡被保人名冊，建立身分證 → 旅平卡資訊的 Map ──
  const memberSheet = ss.getSheetByName(CONFIG.TRAVEL_MEMBER_SHEET);
  const mainTravelSheet = ss.getSheetByName(CONFIG.TRAVEL_MAIN_SHEET);

  if (!memberSheet || !mainTravelSheet) {
    ui.alert('⚠️ 找不到「旅平卡被保人名冊」或「旅平卡主表」工作表，請先執行旅平卡同步。');
    return;
  }

  // 旅平卡主表：憑證號碼 → 要保人姓名
  const travelMainData  = mainTravelSheet.getDataRange().getValues();
  const certToApplicant = new Map(); // certNo → 要保人姓名
  for (let i = 1; i < travelMainData.length; i++) {
    const certNo  = String(travelMainData[i][1] || "").trim();
    const appName = String(travelMainData[i][4] || "").trim();
    if (certNo) certToApplicant.set(certNo, appName);
  }

  // 旅平卡被保人名冊：身分證號 → { certNo, relation, appName }
  // 同一人可能在多張旅平卡，取第一筆
  const memberData  = memberSheet.getDataRange().getValues();
  const travelIdMap = new Map(); // idNo → { certNo, relation, appName }
  for (let i = 1; i < memberData.length; i++) {
    const idNo    = String(memberData[i][3] || "").trim(); // D欄：身分證
    const certNo  = String(memberData[i][1] || "").trim(); // B欄：憑證號碼
    const relation= String(memberData[i][5] || "").trim(); // F欄：關係
    if (idNo && !travelIdMap.has(idNo)) {
      travelIdMap.set(idNo, {
        certNo,
        relation,
        appName: certToApplicant.get(certNo) || "—"
      });
    }
  }

  // ── 讀取產險主表，逐筆比對 ──
  const mainSheet = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
  const mainData  = mainSheet.getDataRange().getValues();

  // 去重：同一身分證號只保留一筆（避免同一人多張保單重複出現）
  const seenIds  = new Set();
  const reportRows = [];

  for (let i = 1; i < mainData.length; i++) {
    const insuredId   = String(mainData[i][3] || "").trim(); // D欄：被保人身分證
    const insuredName = String(mainData[i][4] || "").trim(); // E欄：被保人姓名
    const insType     = String(mainData[i][2] || "").trim(); // C欄：險種
    const status      = String(mainData[i][17]|| "").trim(); // R欄：保單狀態
    const expiry      = safeString(mainData[i][16]);          // Q欄：到期日
    const applicant   = String(mainData[i][8] || "").trim(); // I欄：要保人姓名

    if (!insuredId || insuredId === "無") continue;
    if (seenIds.has(insuredId)) continue;
    seenIds.add(insuredId);

    // 比對旅平卡
    const travelInfo = travelIdMap.get(insuredId);
    let travelStatus, travelRole, travelCertNo, travelAppName;

    if (travelInfo) {
      travelStatus  = "✅ 有";
      travelRole    = travelInfo.relation === "本人" ? "主保人" : "親屬：" + travelInfo.relation;
      travelCertNo  = travelInfo.certNo;
      travelAppName = travelInfo.appName;
    } else {
      travelStatus  = "❌ 無";
      travelRole    = "—";
      travelCertNo  = "—";
      travelAppName = "—";
    }

    reportRows.push([
      0,            // A: 序號（後面重新編）
      insuredName,  // B: 被保人姓名
      insuredId,    // C: 被保人身分證
      insType,      // D: 險種
      status,       // E: 保單狀態
      expiry,       // F: 到期日
      applicant,    // G: 要保人姓名
      travelStatus, // H: 旅平卡狀態
      travelRole,   // I: 旅平卡身份
      travelCertNo, // J: 旅平卡憑證號碼
      travelAppName // K: 旅平卡要保人
    ]);
  }

  // ── 寫入比對旅平卡覆蓋表 ──
  let coverageSheet = ss.getSheetByName(CONFIG.COVERAGE_SHEET_NAME);
  if (!coverageSheet) {
    coverageSheet = ss.insertSheet(CONFIG.COVERAGE_SHEET_NAME);
  }

  // 清空舊資料（保留第一列）
  const lastRow = coverageSheet.getLastRow();
  if (lastRow > 1) {
    coverageSheet.getRange(2, 1, lastRow - 1, 11).clear();
  }

  // 寫入標題（第一次執行才寫，或強制更新）
  coverageSheet.getRange(1, 1, 1, 11).setValues([[
    '序號', '被保人姓名', '被保人身分證', '險種', '保單狀態',
    '到期日', '要保人姓名', '旅平卡狀態', '旅平卡身份', '旅平卡憑證號碼', '旅平卡要保人'
  ]]);

  // 標題列格式
  const headerRange = coverageSheet.getRange(1, 1, 1, 11);
  headerRange.setBackground('#1565C0');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');

  if (reportRows.length === 0) {
    ui.alert('比對完成，無資料。');
    return;
  }

  // 補序號
  reportRows.forEach((row, idx) => { row[0] = idx + 1; });

  // 寫入資料
  coverageSheet.getRange(2, 1, reportRows.length, 11).setValues(reportRows);

  // 條件格式：H欄（旅平卡狀態）有旅平卡→綠色背景，無→紅色背景
  for (let i = 0; i < reportRows.length; i++) {
    const rowNum   = i + 2;
    const hasCard  = reportRows[i][7] === "✅ 有";
    const statusCell = coverageSheet.getRange(rowNum, 8); // H欄
    statusCell.setBackground(hasCard ? '#E8F5E9' : '#FFEBEE');
    statusCell.setFontColor(hasCard ? '#2E7D32' : '#C62828');
    statusCell.setFontWeight('bold');
  }

  // 自動調整欄寬
  coverageSheet.autoResizeColumns(1, 11);

  // 統計摘要
  const totalCount  = reportRows.length;
  const hasCount    = reportRows.filter(r => r[7] === "✅ 有").length;
  const noCount     = totalCount - hasCount;
  const coverageRate = Math.round(hasCount / totalCount * 100);

  const msg = '✅ 旅平卡覆蓋率比對完成！\n\n' +
    '總人數：' + totalCount + ' 人\n' +
    '已持旅平卡：' + hasCount + ' 人 ✅\n' +
    '未持旅平卡：' + noCount + ' 人 ❌\n' +
    '覆蓋率：' + coverageRate + '%\n\n' +
    '結果已寫入「比對旅平卡覆蓋表」工作表。';

  ui.alert(msg);
  Logger.log(msg);
}

function testSystemConfig() {
  const cfg = getSystemConfig();
  Logger.log('=== 系統設定讀取結果 ===');
  Logger.log('顧問姓名：' + cfg['顧問姓名']);
  Logger.log('顧問電話：' + cfg['顧問電話']);
  Logger.log('顧問LINE：' + cfg['顧問LINE']);
  Logger.log('緊急天數：' + cfg['緊急天數']);
  Logger.log('警示天數：' + cfg['警示天數']);
}