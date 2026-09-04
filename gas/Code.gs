// =========================================================================
// 諾思 NO01 2026 員工健康計劃 - 升級版 Google Apps Script (Code.gs)
// 高效能（5 分鐘 CacheService 快取加速）、安全鎖定與完整團隊/管理員權限邏輯
// =========================================================================

const scriptProperties = PropertiesService.getScriptProperties();
const SPREADSHEET_ID = scriptProperties.getProperty('SPREADSHEET_ID');
const INITIAL_SUPER_ADMIN = scriptProperties.getProperty('INITIAL_SUPER_ADMIN');

// 5 分鐘 快取機制 (300 秒)
const CACHE_TTL = 300;

function clearCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll(["CACHE_MAIN_DATA", "CACHE_CHAT_MESSAGES"]);
  } catch (e) {}
}

function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim()) {
    return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getAdminSheet(ss) {
  ss = ss || getSpreadsheet();
  let sheet = ss.getSheetByName("Admin_List");
  if (!sheet) {
    sheet = ss.insertSheet("Admin_List");
    sheet.appendRow(["Admin Email", "新增者", "新增時間"]);
    if (INITIAL_SUPER_ADMIN && INITIAL_SUPER_ADMIN.trim()) {
      sheet.appendRow([INITIAL_SUPER_ADMIN.trim().toLowerCase(), "System_Init", new Date()]);
    }
    sheet.appendRow(["zach.yang@north.com.tw", "System_Init", new Date()]);
  }
  return sheet;
}

function getUserProfilesSheet(ss) {
  ss = ss || getSpreadsheet();
  let sheet = ss.getSheetByName("User_Profiles");
  if (!sheet) {
    sheet = ss.insertSheet("User_Profiles");
    sheet.appendRow(["Email", "暱稱", "最後更新時間"]);
  }
  return sheet;
}

function getChatSheet(ss) {
  ss = ss || getSpreadsheet();
  let sheet = ss.getSheetByName("Chat_Messages");
  if (!sheet) {
    sheet = ss.insertSheet("Chat_Messages");
    sheet.appendRow(["時間", "Email", "暱稱", "訊息內容"]);
  }
  return sheet;
}

function getTeamsSheet(ss) {
  ss = ss || getSpreadsheet();
  let sheet = ss.getSheetByName("Teams");
  if (!sheet) {
    sheet = ss.insertSheet("Teams");
    sheet.appendRow(["TeamName", "CaptainEmail", "Member1", "Member2", "Member3", "ExtraPoints", "最後更新時間"]);
  }
  return sheet;
}

function getPacerPointsSheet(ss) {
  ss = ss || getSpreadsheet();
  let sheet = ss.getSheetByName("Pacer_Points");
  if (!sheet) {
    sheet = ss.insertSheet("Pacer_Points");
    sheet.appendRow(["Email", "PacerPoints", "最後更新時間"]);
  }
  return sheet;
}

function getReactionsSheet(ss) {
  ss = ss || getSpreadsheet();
  let sheet = ss.getSheetByName("Reactions");
  if (!sheet) {
    sheet = ss.insertSheet("Reactions");
    sheet.appendRow(["msgId", "type", "userEmail", "timestamp"]);
  }
  return sheet;
}

function getReactionsData(ss, userEmail) {
  const sheet = getReactionsSheet(ss);
  const data = sheet.getDataRange().getValues();
  const reactionsMap = {};
  const cleanUser = String(userEmail || "").trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const msgId = String(data[i][0] || "").trim();
    const type = String(data[i][1] || "").trim().toLowerCase();
    const email = String(data[i][2] || "").trim().toLowerCase();

    if (msgId && type) {
      if (!reactionsMap[msgId]) {
        reactionsMap[msgId] = {
          counts: { like: 0, cheer: 0, agree: 0, kneel: 0, boba: 0 },
          userReacted: []
        };
      }
      reactionsMap[msgId].counts[type] = (reactionsMap[msgId].counts[type] || 0) + 1;
      if (cleanUser && email === cleanUser && !reactionsMap[msgId].userReacted.includes(type)) {
        reactionsMap[msgId].userReacted.push(type);
      }
    }
  }
  return reactionsMap;
}

function toggleReactionInSheet(msgId, type, userEmail, ss) {
  ss = ss || getSpreadsheet();
  const sheet = getReactionsSheet(ss);
  const data = sheet.getDataRange().getValues();
  const cleanMsgId = String(msgId || "").trim();
  const cleanType = String(type || "").trim().toLowerCase();
  const cleanEmail = String(userEmail || "").trim().toLowerCase();

  if (!cleanMsgId || !cleanType || !cleanEmail) return "無效的訊息或使用者資訊";

  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    const rowMsgId = String(data[i][0] || "").trim();
    const rowType = String(data[i][1] || "").trim().toLowerCase();
    const rowEmail = String(data[i][2] || "").trim().toLowerCase();

    if (rowMsgId === cleanMsgId && rowType === cleanType && rowEmail === cleanEmail) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow > 0) {
    sheet.deleteRow(foundRow);
  } else {
    sheet.appendRow([cleanMsgId, cleanType, cleanEmail, new Date()]);
  }
  return null;
}

function getUserNicknameMap(ss) {
  const sheet = getUserProfilesSheet(ss);
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const email = row[0];
    const nick = row[1];
    if (email) {
      map[String(email).trim().toLowerCase()] = String(nick || "").trim();
    }
  }
  return map;
}

const FUN_NICKNAME_SUFFIXES = [
  "⚡ 獅子總裁", "🏃 雙子畫手", "🔋 水瓶速跑", "🚴 天蠍勁跑", "🏆 牡羊極限",
  "🍳 射手派對", "🌟 金牛慢活", "🎨 摩羯畫家", "💥 天秤運動", "🥳 巨蟹加油",
  "🦸 處女規律", "🔥 雙魚跑者", "🚀 ENFJ 隊長", "🎯 ISTP 獵人", "🌈 INFP 繪師",
  "💫 ESTJ 指揮", "🦁 ENTJ 領航", "🦊 ENFP 冒險", "🦉 INTP 演算", "🐆 ESTP 獵豹",
  "🦄 ISFP 健走", "🐻 ISFJ 神行", "🐬 ENFJ 發電", "🦅 INTJ 策劃", "🧋 珍奶去冰",
  "🍟 薯條加大", "☕ 雙倍濃縮", "🍕 起司熱量", "🥐 可頌甜心", "🍧 冰淇淋控",
  "🍔 雙層漢堡", "🍣 鮭魚大腹", "🍜 拉麵加麵", "🍩 甜甜圈控", "🥤 汽水加冰",
  "🍰 起司蛋糕", "🥑 酪梨吐司", "🍱 精致便當", "🥩 熟成牛排", "🌮 塔可狂人",
  "💻 跑步換腦", "📊 步數大師", "☕ 諾思咖啡", "🎯 達標八千", "🍳 諾思當家",
  "🏆 戰術大師", "⌨️ 零Bug跑者", "🚀 程式神手", "📈 點數飆升", "📂 檔案歸檔",
  "💾 隨時存檔", "🔍 邏輯天才", "📑 簡報神人", "💡 創意源泉", "⏱️ 效率之王",
  "🧠 超級大腦", "🐧 企鵝畫手", "🐕 柴犬衝鋒", "🐈 輕功大師", "🦥 達標專家",
  "🦩 紅鶴健走", "🐬 鯨豚長跑", "🦅 鷹眼路跑", "🐯 猛虎下山", "🐼 功夫熊貓",
  "🐰 閃電兔兔", "🐺 孤狼跑者", "🐎 飆風奔馬", "⚔️ 步數呼吸", "🦸 超級賽亞",
  "🥷 飆風忍者", "🔮 萬步魔法", "🏴‍☠️ 尋寶船長", "🤖 鋼鐵心臟", "🥊 拳擊熱血",
  "🎾 網球抽球", "🏀 灌籃高手", "⚽ 勁爆射門", "⛳ 一桿進洞", "🧗 攀岩達人"
];

function getRandomFunNickname(email) {
  let hash = 0;
  const str = String(email || Math.random()).toLowerCase();
  for (let i = 0; i < str.length; i++) {
    hash = (hash << 5) - hash + str.charCodeAt(i);
    hash |= 0;
  }
  const idx = Math.abs(hash) % FUN_NICKNAME_SUFFIXES.length;
  return `諾思 · ${FUN_NICKNAME_SUFFIXES[idx]}`;
}

function updateUserNickname(email, nickname, ss) {
  const cleanEmail = String(email || "").trim().toLowerCase();
  if (!cleanEmail) return "諾思夥伴";

  const sheet = getUserProfilesSheet(ss);
  const data = sheet.getDataRange().getValues();
  
  let foundRow = -1;
  let existingNick = "";
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === cleanEmail) {
      foundRow = i + 1;
      existingNick = String(data[i][1] || "").trim();
      break;
    }
  }

  let cleanName = String(nickname || "").trim();

  if (foundRow > 0) {
    // 既存同仁：若先前已有暱稱，除非傳入明確自訂的新暱稱 (非空且非預設)，否則一律保留既有暱稱！
    if (existingNick && existingNick !== "諾思夥伴") {
      if (!cleanName || cleanName === "諾思夥伴" || cleanName === cleanEmail || cleanName === existingNick) {
        return existingNick;
      }
    }
    // 既存同仁手動變更暱稱
    if (!cleanName || cleanName === "諾思夥伴") {
      cleanName = existingNick || getRandomFunNickname(cleanEmail);
    }
    sheet.getRange(foundRow, 2).setValue(cleanName);
    sheet.getRange(foundRow, 3).setValue(new Date());
    return cleanName;
  } else {
    // 首次登入新同仁：僅於第一次登入且無紀錄時演算產生專屬趣味暱稱
    if (!cleanName || cleanName === "諾思夥伴") {
      cleanName = getRandomFunNickname(cleanEmail);
    }
    sheet.appendRow([cleanEmail, cleanName, new Date()]);
    return cleanName;
  }
}

function getAdminList(ss) {
  const sheet = getAdminSheet(ss);
  const data = sheet.getDataRange().getValues();
  const admins = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      const email = String(data[i][0]).trim().toLowerCase();
      if (email) admins.push(email);
    }
  }
  if (INITIAL_SUPER_ADMIN && INITIAL_SUPER_ADMIN.trim()) {
    const initAdmin = INITIAL_SUPER_ADMIN.trim().toLowerCase();
    if (!admins.includes(initAdmin)) admins.push(initAdmin);
  }
  if (!admins.includes("zach.yang@north.com.tw")) admins.push("zach.yang@north.com.tw");
  return Array.from(new Set(admins));
}

// 取得團隊清單 (確保隊長及所有隊員皆完整歸屬並去除重複)
function getTeams(ss) {
  const sheet = getTeamsSheet(ss);
  const data = sheet.getDataRange().getValues();
  const teams = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      const captain = String(row[1] || "").trim();
      const m1 = String(row[2] || "").trim();
      const m2 = String(row[3] || "").trim();
      const m3 = String(row[4] || "").trim();
      
      const rawList = [captain, m1, m2, m3].filter(x => x && String(x).trim());
      const uniqueMembers = [];
      rawList.forEach(m => {
        const cleanM = String(m).trim();
        if (!uniqueMembers.some(x => x.toLowerCase() === cleanM.toLowerCase())) {
          uniqueMembers.push(cleanM);
        }
      });

      teams.push({
        teamName: String(row[0]).trim(),
        captainEmail: captain || (uniqueMembers[0] || ""),
        members: uniqueMembers,
        extraPoints: Number(row[5]) || 0
      });
    }
  }
  return teams;
}

function getPacerPointsMap(ss) {
  const sheet = getPacerPointsSheet(ss);
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      map[String(row[0]).trim().toLowerCase()] = Number(row[1]) || 0;
    }
  }
  return map;
}

// 高效能快取讀取
function getCachedMainData(ss) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("CACHE_MAIN_DATA");
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e){}
  }

  const nicknameMap = getUserNicknameMap(ss);
  const eventData = getEventData(ss, nicknameMap);
  const teams = getTeams(ss);
  const admins = getAdminList(ss);

  const mainData = {
    leaderboard: eventData.leaderboard,
    records: eventData.records,
    teams: teams,
    adminList: admins,
    nicknameMap: nicknameMap
  };

  try {
    const str = JSON.stringify(mainData);
    if (str.length < 95000) {
      cache.put("CACHE_MAIN_DATA", str, CACHE_TTL);
    }
  } catch(e){}

  return mainData;
}

function getCachedChatMessages(ss) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("CACHE_CHAT_MESSAGES");
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e){}
  }

  const chatMessages = getChatMessages(ss);
  try {
    const str = JSON.stringify(chatMessages);
    if (str.length < 95000) {
      cache.put("CACHE_CHAT_MESSAGES", str, CACHE_TTL);
    }
  } catch(e){}

  return chatMessages;
}

function doGet(e) {
  try {
    const ss = getSpreadsheet();
    const userEmail = (e && e.parameter && e.parameter.email ? e.parameter.email : "").trim().toLowerCase();
    
    const mainData = getCachedMainData(ss);
    const chatMessages = getCachedChatMessages(ss);

    const nicknameMap = mainData.nicknameMap || {};
    if (userEmail && userEmail.includes("@") && !nicknameMap[userEmail]) {
      const savedNick = updateUserNickname(userEmail, "諾思夥伴", ss);
      nicknameMap[userEmail] = savedNick;
      clearCache();
    }

    const admins = mainData.adminList || [];
    const isAdmin = admins.includes(userEmail);
    const userNickname = nicknameMap[userEmail] || "諾思夥伴";

    const reactions = getReactionsData(ss, userEmail);

    return createJsonResponse({
      isAdmin: isAdmin,
      userNickname: userNickname,
      leaderboard: mainData.leaderboard,
      records: mainData.records,
      teams: mainData.teams,
      chatMessages: chatMessages,
      nicknameMap: nicknameMap,
      reactions: reactions,
      pacerLogs: isAdmin ? getPacerLogs(ss, nicknameMap) : [],
      adminList: isAdmin ? admins : []
    });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.message });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const ss = getSpreadsheet();
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const operatorEmail = (payload.operatorEmail || payload.email || "").trim().toLowerCase();
    const admins = getAdminList(ss);
    const isOperatorAdmin = admins.includes(operatorEmail);

    // 1. 發布聊天訊息
    if (action === "sendChat" || action === "chat") {
      const chatSheet = getChatSheet(ss);
      chatSheet.appendRow([
        new Date(),
        payload.email,
        payload.name || "諾思夥伴",
        payload.message
      ]);
      clearCache();
      return createJsonResponse({ success: true, chatMessages: getChatMessages(ss) });
    }

    // 2. 修改與註冊暱稱 (registerUser / updateNickname / setNickname)
    if (action === "updateNickname" || action === "setNickname" || action === "registerUser") {
      const updated = updateUserNickname(operatorEmail || payload.email, payload.nickname || payload.name, ss);
      clearCache();
      return createJsonResponse({ success: true, nickname: updated });
    }

    // 3. 打卡回報
    if (action === "checkIn") {
      updateUserNickname(payload.email, payload.name, ss);
      submitCheckIn(payload, ss);
      clearCache();
      return createJsonResponse({ success: true, message: "回報成功" });
    }

    // 4. 管理員審核
    if (action === "review") {
      if (!isOperatorAdmin) return createJsonResponse({ success: false, message: `403 權限不足：帳號 (${operatorEmail}) 不在管理員名單中` });
      reviewRecord(payload.rowId, payload.status, payload.approvedSteps, ss);
      clearCache();
      return createJsonResponse({ success: true });
    }

    // 5. 管理員新增管理員
    if (action === "addAdmin") {
      if (!isOperatorAdmin) return createJsonResponse({ success: false, message: `403 權限不足：帳號 (${operatorEmail}) 不在管理員名單中` });
      const newAdmin = (payload.newAdminEmail || "").trim().toLowerCase();
      if (newAdmin && !admins.includes(newAdmin)) {
        getAdminSheet(ss).appendRow([newAdmin, operatorEmail, new Date()]);
      }
      clearCache();
      return createJsonResponse({ success: true, adminList: getAdminList(ss) });
    }

    // 6. 設定 3 人團隊 (setTeam)
    if (action === "setTeam") {
      const err = saveTeam(payload.teamName, payload.captainEmail || operatorEmail, payload.members, isOperatorAdmin, ss);
      if (err) return createJsonResponse({ success: false, message: err });
      clearCache();
      return createJsonResponse({ success: true, teams: getTeams(ss) });
    }

    // 7. 管理員修改隊名 (adminEditTeam)
    if (action === "adminEditTeam") {
      if (!isOperatorAdmin) return createJsonResponse({ success: false, message: `403 權限不足：帳號 (${operatorEmail}) 不在管理員名單中` });
      editTeamName(payload.oldTeamName, payload.newTeamName, ss);
      clearCache();
      return createJsonResponse({ success: true, teams: getTeams(ss) });
    }


    // 11. 管理員刪除同仁 (deleteUser)
    if (action === "deleteUser") {
      if (!isOperatorAdmin) return createJsonResponse({ success: false, message: `403 權限不足：帳號 (${operatorEmail}) 不在管理員名單中` });
      const targetEmail = (payload.targetEmail || "").trim().toLowerCase();
      if (targetEmail && operatorEmail && targetEmail === operatorEmail.trim().toLowerCase()) {
        return createJsonResponse({ success: false, message: "⚠️ 存取拒絕：管理員不可刪除您自身正在使用中的帳號！" });
      }
      const err = deleteUser(targetEmail, ss);
      if (err) return createJsonResponse({ success: false, message: err });
      clearCache();
      return createJsonResponse({ success: true, message: `已成功刪除同仁 (${targetEmail})` });
    }

    // 12. 管理員手動新增同仁 (addUser / adminAddUser)
    if (action === "addUser" || action === "adminAddUser") {
      if (!isOperatorAdmin) return createJsonResponse({ success: false, message: `403 權限不足：帳號 (${operatorEmail}) 不在管理員名單中` });
      const targetE = String(payload.targetEmail || payload.email || "").trim().toLowerCase();
      const nick = String(payload.nickname || payload.name || "諾思夥伴").trim();
      if (targetE) {
        updateUserNickname(targetE, nick, ss);
        clearCache();
        return createJsonResponse({ success: true, message: `已成功新增同仁 (${nick})` });
      }
      return createJsonResponse({ success: false, message: "無效的 Email" });
    }

    // 13. 表情貼圖按讚同步 (toggleReaction)
    if (action === "toggleReaction" || action === "reaction") {
      const msgId = payload.msgId;
      const type = payload.type;
      const email = payload.operatorEmail || payload.email || "";
      toggleReactionInSheet(msgId, type, email, ss);
      clearCache();
      return createJsonResponse({
        success: true,
        reactions: getReactionsData(ss, email)
      });
    }

    // 8. 管理員更新 Pacer 點數 (updatePacerPoints)
    if (action === "updatePacerPoints") {
      if (!isOperatorAdmin) return createJsonResponse({ success: false, message: `403 權限不足：帳號 (${operatorEmail}) 不在管理員名單中` });
      savePacerPoints(payload.targetEmail, payload.points, operatorEmail, ss);
      clearCache();
      return createJsonResponse({ success: true, pacerLogs: getPacerLogs(ss) });
    }

    // 9. 管理員強制清空快取 (clearCache)
    if (action === "clearCache") {
      if (!isOperatorAdmin) return createJsonResponse({ success: false, message: `403 權限不足：帳號 (${operatorEmail}) 不在管理員名單中` });
      clearCache();
      return createJsonResponse({ success: true, message: "全系統快取已成功強制清空！" });
    }

    // 10. 管理員 / 隊長刪除團隊 (deleteTeam)
    if (action === "deleteTeam") {
      const targetTeamName = String(payload.teamName || "").trim();
      if (!targetTeamName) return createJsonResponse({ success: false, message: "未提供刪除之團隊名稱" });
      const sheet = getTeamsSheet(ss);
      const data = sheet.getDataRange().getValues();
      let isCaptain = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim().toLowerCase() === targetTeamName.toLowerCase()) {
          const capt = String(data[i][1] || "").trim().toLowerCase();
          if (capt && capt === operatorEmail.trim().toLowerCase()) {
            isCaptain = true;
          }
        }
      }
      if (!isOperatorAdmin && !isCaptain) {
        return createJsonResponse({ success: false, message: `403 權限不足：您必須是團隊隊長或管理員才能刪除團隊` });
      }
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]).trim().toLowerCase() === targetTeamName.toLowerCase()) {
          sheet.deleteRow(i + 1);
        }
      }
      clearCache();
      return createJsonResponse({ success: true, message: `已成功刪除團隊「${targetTeamName}」`, teams: getTeams(ss) });
    }

    return createJsonResponse({ success: false, message: "未知動作" });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

function saveTeam(teamName, captainEmail, members, isAdmin, ss) {
  ss = ss || getSpreadsheet();
  const sheet = getTeamsSheet(ss);
  const data = sheet.getDataRange().getValues();
  const tName = String(teamName).trim();
  const captain = String(captainEmail).trim().toLowerCase();
  const rawMembers = (members || [captain, "", ""]).map(x => String(x).trim().toLowerCase()).filter(Boolean);
  
  if (captain && !rawMembers.includes(captain)) {
    rawMembers.unshift(captain);
  }

  const isCaptainOfExistingTeam = data.some((row, rIdx) => rIdx >= 1 && String(row[1]).trim().toLowerCase() === captain);

  // 1. 若隊長先前已建立過其他名稱的舊團隊，自動刪除舊團隊資料列，防止產生幽靈雙團隊
  for (let i = data.length - 1; i >= 1; i--) {
    const rowTName = String(data[i][0]).trim();
    if (rowTName.toLowerCase() !== tName.toLowerCase()) {
      const rowCapt = String(data[i][1]).trim().toLowerCase();
      if (rowCapt === captain) {
        sheet.deleteRow(i + 1);
      }
    }
  }

  // 重新獲取最新試算表資料
  const freshData = sheet.getDataRange().getValues();

  // 1.5 檢查是否為全新團隊建立 (僅限管理員權限或既有隊長修改隊名)
  const isExistingTeam = freshData.some((row, rIdx) => rIdx >= 1 && String(row[0]).trim().toLowerCase() === tName.toLowerCase());
  if (!isExistingTeam && !isCaptainOfExistingTeam && !isAdmin) {
    return "權限受限：團隊建立統一僅限管理員 (Admin) 配對與建置！如需新建團隊請聯繫管理員。";
  }

  // 2. 檢查同仁是否已加入其他團隊 (嚴格執行：一人只能參加一個團隊)
  for (let i = 1; i < freshData.length; i++) {
    const rowTName = String(freshData[i][0]).trim();
    if (rowTName.toLowerCase() !== tName.toLowerCase()) {
      const rowMembers = [
        String(freshData[i][1]).trim().toLowerCase(),
        String(freshData[i][2]).trim().toLowerCase(),
        String(freshData[i][3]).trim().toLowerCase(),
        String(freshData[i][4]).trim().toLowerCase()
      ].filter(Boolean);

      for (let m of rawMembers) {
        if (m && rowMembers.some(rm => rm && (rm === m || rm.includes(m) || m.includes(rm)))) {
          if (!isAdmin) {
            return "同仁 (" + m + ") 已加入團隊「" + rowTName + "」，每位同仁只能歸屬一個團隊！如需更換團隊請聯繫管理員。";
          } else {
            // 管理員操作：自動從原舊團隊將該同仁移除，徹底防止一人跨多隊
            for (let col = 1; col <= 4; col++) {
              const cellVal = String(data[i][col] || "").trim().toLowerCase();
              if (cellVal && (cellVal === m || cellVal.includes(m) || m.includes(cellVal))) {
                sheet.getRange(i + 1, col + 1).setValue("");
              }
            }
          }
        }
      }
    }
  }

  // 重新獲取最新的試算表列位
  const updatedData = sheet.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < updatedData.length; i++) {
    if (String(updatedData[i][0]).trim().toLowerCase() === tName.toLowerCase()) {
      foundRow = i + 1;
      break;
    }
  }

  const m1 = rawMembers[0] || captain;
  const m2 = rawMembers[1] || "";
  const m3 = rawMembers[2] || "";

  if (foundRow > 0) {
    sheet.getRange(foundRow, 2, 1, 4).setValues([[captain, m1, m2, m3]]);
    sheet.getRange(foundRow, 7).setValue(new Date());
  } else {
    sheet.appendRow([tName, captain, m1, m2, m3, 0, new Date()]);
  }
  return null;
}

function editTeamName(oldName, newName, ss) {
  ss = ss || getSpreadsheet();
  const sheet = getTeamsSheet(ss);
  const data = sheet.getDataRange().getValues();
  const oldT = String(oldName).trim();
  const newT = String(newName).trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === oldT) {
      sheet.getRange(i + 1, 1).setValue(newT);
      sheet.getRange(i + 1, 7).setValue(new Date());
      break;
    }
  }
}

function savePacerPoints(email, points, operatorEmail, ss) {
  ss = ss || getSpreadsheet();
  const sheet = getPacerPointsSheet(ss);
  const data = sheet.getDataRange().getValues();
  const cleanEmail = String(email).trim().toLowerCase();
  const cleanOp = String(operatorEmail || "").trim().toLowerCase();
  const pts = Number(points) || 0;

  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === cleanEmail) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow > 0) {
    sheet.getRange(foundRow, 2, 1, 3).setValues([[pts, new Date(), cleanOp]]);
  } else {
    sheet.appendRow([cleanEmail, pts, new Date(), cleanOp]);
  }

  // 寫入 Pacer 點數異動日誌 (Pacer_Logs)
  let logSheet = ss.getSheetByName("Pacer_Logs");
  if (!logSheet) {
    logSheet = ss.insertSheet("Pacer_Logs");
    logSheet.appendRow(["Timestamp", "TargetEmail", "PacerPoints", "OperatorEmail"]);
  }
  logSheet.appendRow([new Date(), cleanEmail, pts, cleanOp]);
}

function getPacerLogs(ss, nicknameMap) {
  ss = ss || getSpreadsheet();
  const sheet = ss.getSheetByName("Pacer_Logs");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (!nicknameMap) nicknameMap = getUserNicknameMap(ss);
  const list = [];
  const start = Math.max(1, data.length - 100);

  for (let i = data.length - 1; i >= start; i--) {
    const row = data[i];
    if (!row || !row[1]) continue;
    const targetEmail = String(row[1]).trim().toLowerCase();
    const opEmail = String(row[3] || "").trim().toLowerCase();

    list.push({
      timestamp: row[0] ? Utilities.formatDate(new Date(row[0]), "Asia/Taipei", "yyyy/MM/dd HH:mm") : "",
      targetEmail: targetEmail,
      targetName: nicknameMap[targetEmail] || targetEmail.split("@")[0],
      points: Number(row[2]) || 0,
      operatorEmail: opEmail,
      operatorName: nicknameMap[opEmail] || opEmail.split("@")[0] || "管理員"
    });
  }
  return list;
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function getChatMessages(ss) {
  ss = ss || getSpreadsheet();
  const sheet = getChatSheet(ss);
  const data = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row && row[1]) {
      list.push({
        id: "chat_row_" + (i + 1),
        time: row[0] ? Utilities.formatDate(new Date(row[0]), "Asia/Taipei", "MM/dd HH:mm") : "",
        email: String(row[1]).trim(),
        name: String(row[2] || "諾思夥伴").trim(),
        message: String(row[3] || "").trim()
      });
    }
  }
  return list.reverse();
}

function getEventData(ss, nicknameMap) {
  ss = ss || getSpreadsheet();
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  
  if (!nicknameMap) nicknameMap = getUserNicknameMap(ss);
  const pacerMap = getPacerPointsMap(ss);
  const records = [];
  const userStatsMap = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const email = row[1];
    if (!email) continue;
    const cleanEmail = String(email).trim().toLowerCase();

    const storedName = row[2];
    const currentName = nicknameMap[cleanEmail] || (storedName ? String(storedName).trim() : "") || cleanEmail.split('@')[0] || "諾思夥伴";
    const reportedSteps = Number(row[4]) || 0;
    const status = String(row[7] || "待審核").trim();
    const approvedSteps = (status === "通過") ? (Number(row[8]) || reportedSteps) : 0;

    records.push({
      rowId: i + 1,
      timestamp: row[0] ? Utilities.formatDate(new Date(row[0]), "Asia/Taipei", "yyyy/MM/dd HH:mm") : "",
      email: String(email).trim(),
      name: currentName,
      date: row[3] ? Utilities.formatDate(new Date(row[3]), "Asia/Taipei", "yyyy/MM/dd") : "",
      steps: reportedSteps,
      imageUrl: String(row[5] || "").trim(),
      message: String(row[6] || "").trim(),
      status: status,
      approvedSteps: approvedSteps
    });

    if (!userStatsMap[cleanEmail]) {
      userStatsMap[cleanEmail] = { email: email, name: currentName, validDays: 0, totalSteps: 0, sportCount: 0 };
    }
    if (status === "通過") {
      userStatsMap[cleanEmail].validDays += 1;
      userStatsMap[cleanEmail].totalSteps += approvedSteps;
      userStatsMap[cleanEmail].sportCount += 1;
    }
  }

  const leaderboard = Object.values(userStatsMap).map(user => {
    const cleanE = user.email.toLowerCase().trim();
    const pacerPts = pacerMap[cleanE];
    const calculatedPts = user.validDays + user.sportCount;
    const finalPoints = (pacerPts !== undefined) ? pacerPts : calculatedPts;

    let tier = "尚未達標", reward = 0;
    if (finalPoints >= 20) { tier = "Gold"; reward = 6000; }
    else if (finalPoints >= 15) { tier = "Silver"; reward = 4500; }
    else if (finalPoints >= 10) { tier = "Bronze"; reward = 3000; }

    return {
      email: user.email,
      name: user.name,
      validDays: user.validDays,
      sportCount: user.sportCount,
      totalSteps: user.totalSteps,
      points: finalPoints,
      tier: tier,
      reward: reward
    };
  });

  leaderboard.sort((a, b) => b.points - a.points);
  return { records: records.reverse(), leaderboard: leaderboard };
}

function submitCheckIn(payload, ss) {
  ss = ss || getSpreadsheet();
  const sheet = ss.getSheets()[0];
  let fileUrl = "";

  if (payload.base64Image) {
    try {
      const folderIterator = DriveApp.getFoldersByName("MOVE_ALL_2026_運動截圖");
      let folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.createFolder("MOVE_ALL_2026_運動截圖");
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      let base64Data = payload.base64Image;
      if (base64Data.includes(",")) base64Data = base64Data.split(",")[1];
      const decodedData = Utilities.base64Decode(base64Data);

      const cleanName = String(payload.name || "諾思夥伴").replace(/[^\w\u4e00-\u9fa5]/g, "");
      const emailPrefix = String(payload.email || "user").split("@")[0];
      const dateStr = String(payload.date || Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd")).replace(/\//g, "-");
      const timeStr = Utilities.formatDate(new Date(), "Asia/Taipei", "HHmmss");

      const fileName = `${dateStr}_${cleanName}_${emailPrefix}_${timeStr}.png`;

      const blob = Utilities.newBlob(decodedData, "image/png", fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
    } catch (err) {
      fileUrl = payload.imageUrl || ("上傳失敗: " + err.message);
    }
  }

  sheet.appendRow([
    new Date(),
    payload.email,
    payload.name || "諾思夥伴",
    payload.date,
    Number(payload.steps),
    fileUrl,
    payload.message,
    "待審核",
    0,
    ""
  ]);
}

function reviewRecord(rowId, status, approvedSteps, ss) {
  ss = ss || getSpreadsheet();
  const sheet = ss.getSheets()[0];
  sheet.getRange(rowId, 8).setValue(status);
  sheet.getRange(rowId, 9).setValue(status === "通過" ? approvedSteps : 0);
}

function deleteUser(targetEmail, ss) {
  const cleanEmail = String(targetEmail || "").trim().toLowerCase();
  if (!cleanEmail) return "無效的 Email 帳號";

  ss = ss || getSpreadsheet();

  // 1. 從 User_Profiles 試算表中刪除該同仁資料列
  const profileSheet = getUserProfilesSheet(ss);
  const profileData = profileSheet.getDataRange().getValues();
  for (let i = profileData.length - 1; i >= 1; i--) {
    if (String(profileData[i][0]).trim().toLowerCase() === cleanEmail) {
      profileSheet.deleteRow(i + 1);
    }
  }

  // 2. 從 Pacer_Points 試算表中刪除
  const pacerSheet = getPacerPointsSheet(ss);
  const pacerData = pacerSheet.getDataRange().getValues();
  for (let i = pacerData.length - 1; i >= 1; i--) {
    if (String(pacerData[i][0]).trim().toLowerCase() === cleanEmail) {
      pacerSheet.deleteRow(i + 1);
    }
  }

  // 3. 從所有團隊中移除該成員或重置隊伍
  const teamsSheet = getTeamsSheet(ss);
  const teamsData = teamsSheet.getDataRange().getValues();
  for (let i = teamsData.length - 1; i >= 1; i--) {
    const row = teamsData[i];
    const teamName = String(row[0] || "").trim();
    if (!teamName) continue;

    const capt = String(row[1] || "").trim();
    const m1 = String(row[2] || "").trim();
    const m2 = String(row[3] || "").trim();
    const m3 = String(row[4] || "").trim();

    const members = [m1, m2, m3].filter(m => m && String(m).trim().toLowerCase() !== cleanEmail);
    const isCaptMatched = capt.toLowerCase().trim() === cleanEmail;
    const isMemberMatched = [m1, m2, m3].some(m => String(m).trim().toLowerCase() === cleanEmail);

    if (isCaptMatched || isMemberMatched) {
      if (members.length === 0) {
        teamsSheet.deleteRow(i + 1);
      } else {
        const newCapt = isCaptMatched ? (members[0] || "") : capt;
        teamsSheet.getRange(i + 1, 2).setValue(newCapt);
        teamsSheet.getRange(i + 1, 3).setValue(members[0] || "");
        teamsSheet.getRange(i + 1, 4).setValue(members[1] || "");
        teamsSheet.getRange(i + 1, 5).setValue(members[2] || "");
        teamsSheet.getRange(i + 1, 7).setValue(new Date());
      }
    }
  }

  return null;
}
