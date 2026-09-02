// 1. ฟังก์ชันสมัครบัญชีใหม่
function registerUser(agency, email, phone, fullname) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('UserDB');
  if (!sheet) return { success: false, message: "ไม่พบฐานข้อมูล UserDB" };
  
  const data = sheet.getDataRange().getValues();
  
  // เช็กว่าอีเมลนี้มีอยู่ในระบบแล้วหรือไม่
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === email) {
      return { success: false, message: "อีเมลนี้มีอยู่ในระบบแล้ว" };
    }
  }
  
  // ลำดับคอลัมน์: A=Agency, B=Email, C=Phone, D=Fullname, E=Created_at, F=Active
  // ตรง phone ให้เติม ' เข้าไปข้างหน้าเพื่อให้ Google Sheets มองเป็นข้อความ (ไม่ลบ 0 นำหน้า)
  sheet.appendRow([agency, email, "'" + phone, fullname, new Date(), false]);
  
  return { success: true, message: "สมัครเสร็จบันทึกเรียบร้อย โปรดติดต่อเจ้าหน้าที่เพื่อนุมัติ" };
}

// 2. ฟังก์ชันตรวจสอบการเข้าสู่ระบบ (พร้อมระบบเก็บ Log และเช็ก Active)
function verifyUser(agency, email, phone) {
  const cache = CacheService.getScriptCache();
  const lockKey = "lock_" + email;
  const attemptKey = "attempt_" + email;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ฟังก์ชันย่อยสำหรับบันทึก Log
  function writeLog(statusMsg) {
    const logSheet = ss.getSheetByName('Log');
    if (logSheet) {
      logSheet.appendRow([new Date(), agency, email, phone, statusMsg]);
    }
  }
  
  // เช็กว่าอีเมลนี้ถูกระงับอยู่หรือไม่
  if (cache.get(lockKey)) {
    writeLog("LOCKED (บัญชีถูกระงับอยู่)");
    return { success: false, locked: true, message: "บัญชีนี้ถูกระงับชั่วคราวเป็นเวลา 3 นาที" };
  }

  const sheet = ss.getSheetByName('UserDB');
  if (!sheet) return { success: false, locked: false, message: "ไม่พบฐานข้อมูล UserDB" };
  
  const data = sheet.getDataRange().getValues();
  
  let emailFound = false;
  let isMatch = false;
  let fullname = "";
  let isActive = false;

  // ตรวจสอบข้อมูลผู้ใช้
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == email) {
      emailFound = true;
      // ลบเครื่องหมาย ' ออกจากเบอร์โทรใน Sheets เพื่อเปรียบเทียบ
      let dbPhone = data[i][2].toString().replace(/^'/, ''); 
      
      if (data[i][0] == agency && dbPhone == phone) {
        isMatch = true;
        fullname = data[i][3];
        // ตรวจสอบคอลัมน์ Active (คอลัมน์ F / Index 5) ว่าติ๊ก True หรือยัง
        isActive = (data[i][5] === true || data[i][5] === 'TRUE' || data[i][5] === 'true');
        break; 
      }
    }
  }
  
  // กรณี: ไม่พบอีเมลนี้ในระบบเลย
  if (!emailFound) {
    writeLog("FAILED (ไม่มีอีเมลในระบบ)");
    return { success: false, locked: false, message: "ไม่พบข้อมูลบัญชีนี้ในระบบ" };
  }

  // กรณี: กรอกข้อมูลผิด
  if (!isMatch) {
    let attempts = parseInt(cache.get(attemptKey) || "0");
    attempts++;
    
    if (attempts >= 5) {
      cache.put(lockKey, "true", 180);
      cache.remove(attemptKey);
      
      const otpSheet = ss.getSheetByName('OtpDB');
      if (otpSheet) otpSheet.appendRow([email, "LOCKED", new Date(), "locked_3mins"]);
      
      writeLog("LOCKED (กรอกผิดครบ 5 ครั้ง)");
      return { success: false, locked: true, message: "ข้อมูลไม่ถูกต้องครบ 5 ครั้ง บัญชีถูกระงับชั่วคราว (3 นาที)" };
    } else {
      cache.put(attemptKey, attempts.toString(), 3600);
      writeLog(`FAILED (ข้อมูลไม่ตรง - ครั้งที่ ${attempts})`);
      return { success: false, locked: false, message: `หน่วยงานหรือเบอร์โทรศัพท์ไม่ถูกต้อง (ผิดพลาด ${attempts}/5 ครั้ง)` };
    }
  }
  
  // กรณี: ข้อมูลถูกต้อง แต่ยังไม่ Active
  if (!isActive) {
    writeLog("FAILED (รออนุมัติ)");
    // inactive: true จะไปสั่งให้หน้าบ้านเปิด Pop-up "ติดต่อเจ้าหน้าที่" ทันที
    return { success: false, locked: false, inactive: true, message: "โปรดติดต่อเจ้าหน้าที่เพื่อนุมัติ" };
  }

  // กรณี: ข้อมูลถูกต้องและ Active แล้ว
  cache.remove(attemptKey);
  writeLog("SUCCESS (ขอ OTP สำเร็จ)");
  return generateAndSaveOTP(email, fullname); 
}

// 3. สร้าง OTP บันทึกลง OtpDB และส่งอีเมล
function generateAndSaveOTP(email, fullname) {
  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  const now = new Date();
  const expiresAt = new Date(now.getTime() + 5 * 60000); // อายุ 5 นาที
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OtpDB');
  if (sheet) sheet.appendRow([email, otp, expiresAt, "pending"]);
  
  const body = `สวัสดีคุณ ${fullname}\n\nรหัส OTP สำหรับเข้าสู่ระบบคลาวด์ของคุณคือ: ${otp}\nรหัสนี้จะหมดอายุในอีก 5 นาที\n\nหากคุณไม่ได้ทำรายการนี้ กรุณาเพิกเฉยต่ออีเมลฉบับนี้`;
  MailApp.sendEmail(email, "รหัส OTP ยืนยันการเข้าสู่ระบบ", body);
  
  return { success: true, message: "ระบบได้ส่งรหัส OTP ไปยังอีเมลของท่านแล้ว", fullname: fullname };
}

// 4. ตรวจสอบความถูกต้องของ OTP
function verifyOTP(email, userOtp) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OtpDB');
  if (!sheet) return { success: false, message: "ไม่พบฐานข้อมูล OtpDB" };
  
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] == email) {
      const dbOtp = data[i][1].toString();
      const expiresAt = new Date(data[i][2]);
      const status = data[i][3];
      
      if (status !== "pending") return { success: false, message: "รหัสนี้ถูกใช้งานไปแล้ว หรือถูกยกเลิก" };
      if (now > expiresAt) {
        sheet.getRange(i + 1, 4).setValue("expired");
        return { success: false, message: "รหัส OTP หมดอายุแล้ว กรุณาขอใหม่" };
      }
      if (dbOtp === userOtp.toString().trim()) {
        sheet.getRange(i + 1, 4).setValue("used");
        return { success: true, message: "ยืนยันตัวตนสำเร็จ" };
      } else {
        return { success: false, message: "รหัส OTP ไม่ถูกต้อง" };
      }
    }
  }
  return { success: false, message: "ไม่พบข้อมูลการขอ OTP" };
}