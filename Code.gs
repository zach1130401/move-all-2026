// =========================================================================
// 諾思 NO01 2026 員工健康計劃 - 升級版 Google Apps Script (Code.gs)
// 100% 保留原有的 ScriptProperties、Admin_List、User_Profiles 與 Chat_Messages
// 新增：Teams 團隊工作表、Pacer_Points 每日點數工作表與對應 doPost 處理函式
// =========================================================================

// 從專案的指令碼屬性安全讀取 SPREADSHEET_ID 與 INITIAL_SUPER_ADMIN
const scriptProperties = PropertiesService.getScriptProperties();
const SPREADSHEET_ID = scriptProperties.getProperty('SPREADSHEET_ID');
const INITIAL_SUPER_ADMIN = scriptProperties.getProperty('INITIAL_SUPER_ADMIN');

function getAdminSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Admin_List");
  if (!sheet) {
    sheet = ss.insertSheet("Admin_List");
    sheet.appendRow(["Admin Email", "新增者", "新增時間"]);
    if (INITIAL_SUPER_ADMIN) {
      sheet.appendRow([INITIAL_SUPER_ADMIN, "System_Init", new Date()]);
    }
    sheet.appendRow(["zach.yang@north.com.tw", "System_Init", new Date()]);
  }
  return sheet;
}

function getUserProfilesSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("User_Profiles");
  if (!sheet) {
    sheet = ss.insertSheet("User_Profiles");
    sheet.appendRow(["Email", "暱稱", "最後更新時間"]);
  }
  return sheet;
}

function getChatSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Chat_Messages");
  if (!sheet) {
    sheet = ss.insertSheet("Chat_Messages");
    sheet.appendRow(["時間", "Email", "暱稱", "訊息內容"]);
  }
  return sheet;
}

// 新增：3 人團隊設定 Sheet
function getTeamsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Teams");
  if (!sheet) {
    sheet = ss.insertSheet("Teams");
    sheet.appendRow(["TeamName", "CaptainEmail", "Member1", "Member2", "Member3", "ExtraPoints", "最後更新時間"]);
  }
  return sheet;
}

// 新增：Pacer 累計點數 Sheet
function getPacerPointsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Pacer_Points");
  if (!sheet) {
    sheet = ss.insertSheet("Pacer_Points");
    sheet.appendRow(["Email", "PacerPoints", "最後更新時間"]);
  }
  return sheet;
}

function getUserNicknameMap() {
  const sheet = getUserProfilesSheet();
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

function updateUserNickname(email, nickname) {
  const cleanEmail = String(email).trim().toLowerCase();
  const cleanName = String(nickname).trim() || "諾思夥伴";
  const sheet = getUserProfilesSheet();
  const data = sheet.getDataRange().getValues();
  
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === cleanEmail) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow > 0) {
    sheet.getRange(foundRow, 2).setValue(cleanName);
    sheet.getRange(foundRow, 3).setValue(new Date());
  } else {
    sheet.appendRow([cleanEmail, cleanName, new Date()]);
  }
  return cleanName;
}

function getAdminList() {
  const sheet = getAdminSheet();
  const data = sheet.getDataRange().getValues();
  const admins = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) admins.push(String(data[i][0]).trim().toLowerCase());
  }
  if (!admins.includes("zach.yang@north.com.tw")) admins.push("zach.yang@north.com.tw");
  return admins;
}

// 取得團隊清單
function getTeams() {
  const sheet = getTeamsSheet();
  const data = sheet.getDataRange().getValues();
  const teams = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      teams.push({
        teamName: String(row[0]).trim(),
        captainEmail: String(row[1] || row[2] || "").trim(),
        members: [String(row[2] || "").trim(), String(row[3] || "").trim(), String(row[4] || "").trim()].filter(x => x),
        extraPoints: Number(row[5]) || 0
      });
    }
  }
  return teams;
}

// 取得 Pacer 點數對照表
function getPacerPointsMap() {
  const sheet = getPacerPointsSheet();
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

function doGet(e) {
  const userEmail = (e && e.parameter && e.parameter.email ? e.parameter.email : "").trim().toLowerCase();
  const admins = getAdminList();
  const isAdmin = admins.includes(userEmail);
  const nicknameMap = getUserNicknameMap();
  const userNickname = nicknameMap[userEmail] || "諾思夥伴";

  const eventData = getEventData(nicknameMap);
  const chatMessages = getChatMessages();
  const teams = getTeams();

  return createJsonResponse({
    isAdmin: isAdmin,
    userNickname: userNickname,
    leaderboard: eventData.leaderboard,
    records: eventData.records,
    teams: teams,
    chatMessages: chatMessages,
    adminList: isAdmin ? admins : []
  });
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const operatorEmail = (payload.operatorEmail || payload.email || "").trim().toLowerCase();
    const admins = getAdminList();
    const isOperatorAdmin = admins.includes(operatorEmail);

    // 1. 發布聊天訊息
    if (action === "sendChat" || action === "chat") {
      const chatSheet = getChatSheet();
      chatSheet.appendRow([
        new Date(),
        payload.email,
        payload.name || "諾思夥伴",
        payload.message
      ]);
      return createJsonResponse({ success: true, chatMessages: getChatMessages() });
    }

    // 2. 修改暱稱
    if (action === "updateNickname" || action === "setNickname") {
      const updated = updateUserNickname(operatorEmail, payload.nickname);
      return createJsonResponse({ success: true, nickname: updated });
    }

    // 3. 打卡回報
    if (action === "checkIn") {
      updateUserNickname(payload.email, payload.name);
      submitCheckIn(payload);
      return createJsonResponse({ success: true, message: "回報成功" });
    }

    // 4. 管理員審核
    if (action === "review") {
      if (!isOperatorAdmin) return createJsonResponse({ success: false, message: "403 權限不足" });
      reviewRecord(payload.rowId, payload.status, payload.approvedSteps);
      return createJsonResponse({ success: true });
    }

    // 5. 管理員新增管理員
    if (action === "addAdmin") {
      if (!isOperatorAdmin) return createJsonResponse({ success: false, message: "403 權限不足" });
      const newAdmin = (payload.newAdminEmail || "").trim().toLowerCase();
      if (!admins.includes(newAdmin)) {
        getAdminSheet().appendRow([newAdmin, operatorEmail, new Date()]);
      }
      return createJsonResponse({ success: true, adminList: getAdminList() });
    }

    // 6. 設定 3 人團隊 (setTeam)
    if (action === "setTeam") {
      saveTeam(payload.teamName, payload.captainEmail || operatorEmail, payload.members);
      return createJsonResponse({ success: true, teams: getTeams() });
    }

    // 7. 管理員修改隊名 (adminEditTeam)
    if (action === "adminEditTeam") {
      if (!isOperatorAdmin) return createJsonResponse({ success: false, message: "403 權限不足" });
      editTeamName(payload.oldTeamName, payload.newTeamName);
      return createJsonResponse({ success: true, teams: getTeams() });
    }

    // 8. 管理員更新 Pacer 點數 (updatePacerPoints)
    if (action === "updatePacerPoints") {
      if (!isOperatorAdmin) return createJsonResponse({ success: false, message: "403 權限不足" });
      savePacerPoints(payload.targetEmail, payload.points);
      return createJsonResponse({ success: true });
    }

    return createJsonResponse({ success: false, message: "未知動作" });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.message });
  }
}

function saveTeam(teamName, captainEmail, members) {
  const sheet = getTeamsSheet();
  const data = sheet.getDataRange().getValues();
  const tName = String(teamName).trim();
  const captain = String(captainEmail).trim().toLowerCase();
  const mList = (members || [captain, "", ""]).map(x => String(x).trim());

  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === tName) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow > 0) {
    sheet.getRange(foundRow, 2, 1, 4).setValues([[captain, mList[0] || "", mList[1] || "", mList[2] || ""]]);
    sheet.getRange(foundRow, 7).setValue(new Date());
  } else {
    sheet.appendRow([tName, captain, mList[0] || "", mList[1] || "", mList[2] || "", 0, new Date()]);
  }
}

function editTeamName(oldName, newName) {
  const sheet = getTeamsSheet();
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

function savePacerPoints(email, points) {
  const sheet = getPacerPointsSheet();
  const data = sheet.getDataRange().getValues();
  const cleanEmail = String(email).trim().toLowerCase();
  const pts = Number(points) || 0;

  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === cleanEmail) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow > 0) {
    sheet.getRange(foundRow, 2).setValue(pts);
    sheet.getRange(foundRow, 3).setValue(new Date());
  } else {
    sheet.appendRow([cleanEmail, pts, new Date()]);
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function getChatMessages() {
  const sheet = getChatSheet();
  const data = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row && row[1]) {
      list.push({
        time: row[0] ? Utilities.formatDate(new Date(row[0]), "Asia/Taipei", "MM/dd HH:mm") : "",
        email: String(row[1]).trim(),
        name: String(row[2] || "諾思夥伴").trim(),
        message: String(row[3] || "").trim()
      });
    }
  }
  return list.reverse();
}

function getEventData(nicknameMap) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  
  const pacerMap = getPacerPointsMap();
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

function submitCheckIn(payload) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets()[0];
  let fileUrl = "";

  if (payload.base64Image) {
    try {
      const folderIterator = DriveApp.getFoldersByName("MOVE_ALL_2026_運動截圖");
      let folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.createFolder("MOVE_ALL_2026_運動截圖");
      let base64Data = payload.base64Image;
      if (base64Data.includes(",")) base64Data = base64Data.split(",")[1];
      const decodedData = Utilities.base64Decode(base64Data);
      const blob = Utilities.newBlob(decodedData, "image/png", `${payload.email}_${payload.date}.png`);
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

function reviewRecord(rowId, status, approvedSteps) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets()[0];
  sheet.getRange(rowId, 8).setValue(status);
  sheet.getRange(rowId, 9).setValue(status === "通過" ? approvedSteps : 0);
}
