// 1. ฟังก์ชันรองรับการเช็กสถานะการเชื่อมต่อ (GET)
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: "API Online", timestamp: new Date() }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 2. ฟังก์ชันหลักสำหรับรับคำสั่งจาก GitHub Pages (POST)
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const payload = data.payload || {};
    let result = {};

    // Router: สั่งงานตาม action ที่ส่งมาจาก GitHub Pages
    switch (action) {
      case "getAgencies":
        result = getAgencies();
        break;
      case "getAllRiskCloudData":
        result = getAllRiskCloudData();
        break;
      case "saveAssessmentData":
        result = saveAssessmentData(payload);
        break;
      case "registerUser":
        result = registerUser(payload.agency, payload.email, payload.phone, payload.fullname);
        break;
      case "verifyUser":
        result = verifyUser(payload.agency, payload.email, payload.phone);
        break;
      case "verifyOTP":
        result = verifyOTP(payload.email, payload.userOtp);
        break;
      case "getIndexPage":
        result = getIndexPage();
        break;
      default:
        result = { success: false, message: "Unrecognized Action: " + action };
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// ฟังก์ชันเดิมของคุณคงไว้ทั้งหมด (ไม่ต้องแก้ไข)
// ==========================================

function getIndexPage() {
  try {
    return HtmlService.createHtmlOutputFromFile('Index').getContent();
  } catch (error) {
    return `<div style="text-align:center; padding: 50px;">
              <h2 style="color: red;">เกิดข้อผิดพลาดในการโหลดหน้าเว็บ</h2>
              <p>ไม่พบไฟล์ <b>index.html</b></p>
            </div>`;
  }
}

function getAgencies() {
  const cache = CacheService.getScriptCache();
  const cachedAgencies = cache.get("cache_agencies");
  if (cachedAgencies) return JSON.parse(cachedAgencies);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('risk_cloud');
    if (!sheet) return [];
    
    const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues(); 
    const uniqueAgencies = new Set();
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0]) uniqueAgencies.add(data[i][0].toString().trim());
    }
    
    const agencies = Array.from(uniqueAgencies).filter(Boolean);
    try { cache.put("cache_agencies", JSON.stringify(agencies), 1800); } catch (e) {}
    
    return agencies;
  } catch (error) {
    return [];
  }
}

function getAllRiskCloudData() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get("all_risk_cloud_data");
  if (cachedData) return JSON.parse(cachedData);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('risk_cloud');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const assets = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if(row[1]) {
        const hasSaved = (row[8] && row[8].toString().trim() !== "") || (row[10] && row[10].toString().trim() !== "");

        assets.push({
          id: row[0] ? row[0].toString() : '',
          agency: row[1] ? row[1].toString() : '',
          typeFilter: row[2] ? row[2].toString() : '', 
          name: row[3] ? row[3].toString() : '',       
          ip: row[4] ? row[4].toString() : '',          
          privateIp: row[5] ? row[5].toString() : '',   
          projectId: row[6] ? row[6].toString() : '',
          domain: row[7] ? row[7].toString() : '',      
          contact: row[8] ? row[8].toString() : '',
          note: row[9] ? row[9].toString() : '',         
          sysType: row[10] ? row[10].toString() : 'ระบบบริการ (Web Services)',
          pdpa: (row[11] === true || row[11] === 'TRUE' || row[11] === 'ใช่'),
          c: parseInt(row[12]) || 1,
          i: parseInt(row[13]) || 1,
          a: parseInt(row[14]) || 1,
          impact: parseInt(row[15]) || 1,
          isSaved: hasSaved
        });
      }
    }
    
    try { cache.put("all_risk_cloud_data", JSON.stringify(assets), 1800); } catch(e) {}
    return assets;
  } catch (error) {
    return [];
  }
}

function saveAssessmentData(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('risk_cloud');
    if (!sheet) return { success: false, message: "ไม่พบชีต risk_cloud" };

    const logSheet = ss.getSheetByName('Sheet1');
    if(logSheet) {
      logSheet.appendRow([new Date(), payload.agency, payload.assessor, "อัปเดตแบบประเมินความเสี่ยง", payload.assets.length + " รายการ"]);
    }

    const data = sheet.getDataRange().getValues();
    const updates = {};
    
    payload.assets.forEach(asset => {
      if(asset.id) updates[asset.id.toString()] = asset;
    });

    for (let i = 1; i < data.length; i++) {
      const rowId = data[i][0] ? data[i][0].toString() : null;
      
      if (rowId && updates[rowId]) {
        const update = updates[rowId];
        
        sheet.getRange(i + 1, 4).setValue(update.resourceName); 
        sheet.getRange(i + 1, 9).setValue(payload.assessor);    
        sheet.getRange(i + 1, 10).setValue(update.note);        
        sheet.getRange(i + 1, 11).setValue(update.sysType);     
        sheet.getRange(i + 1, 12).setValue(update.pdpa ? "TRUE" : "FALSE"); 
        sheet.getRange(i + 1, 13).setValue(update.c);          
        sheet.getRange(i + 1, 14).setValue(update.i);          
        sheet.getRange(i + 1, 15).setValue(update.a);          
        sheet.getRange(i + 1, 16).setValue(update.impact);     
      }
    }
    
    CacheService.getScriptCache().remove("all_risk_cloud_data");
    return { success: true, message: "บันทึกการประเมินลงฐานข้อมูลเรียบร้อยแล้ว!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}