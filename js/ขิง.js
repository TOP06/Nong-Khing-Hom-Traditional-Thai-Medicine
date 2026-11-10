// ฟังก์ชันรับข้อมูลแบบ GET (ใช้ดึงข้อมูล และ ยกเลิกคิว)
function doGet(e) {
  if (!e || !e.parameter) {
    return ContentService.createTextOutput(JSON.stringify({error: "No Action Provided"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  var action = e.parameter.action;
  
  // ให้ GitHub ดึงข้อมูลการจอง
  if (action == 'getBookings') {
    var data = getBookings(); 
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
  }
  
  // ให้ GitHub สั่งยกเลิกคิว
  if (action == 'cancelBooking') {
    var refNo = e.parameter.refNo;
    var result = cancelBooking(refNo); 
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({error: "Invalid GET Action"})).setMimeType(ContentService.MimeType.JSON);
}

// ฟังก์ชันรับข้อมูลแบบ POST (ใช้บันทึกการจองจาก GitHub)
function doPost(e) {
  if (!e || !e.parameter) {
    return ContentService.createTextOutput(JSON.stringify({error: "No Action Provided"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  var action = e.parameter.action;
  
  if (action == 'saveBooking') {
    // รับค่า JSON ที่ส่งมาจาก GitHub fetch()
    var formData = JSON.parse(e.postData.contents);
    var result = saveBooking(formData); 
    
    // สำคัญ: ต้องคืนค่ากลับเพื่อให้หน้าเว็บรู้ว่าสำเร็จ
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({error: "Invalid POST Action"})).setMimeType(ContentService.MimeType.JSON);
}

// -------------------------------------------------------------------------
// ฟังก์ชันการทำงานหลังบ้าน (จากโค้ดเดิมของคุณ)
// -------------------------------------------------------------------------

// บันทึกข้อมูลการจองใหม่ลงใน Google Sheet (คอลัมน์ A - T)
function saveBooking(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
    
    // สร้าง Header ถ้าเป็นชีตใหม่
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "เลขอ้างอิง",              // A
        "เวลาที่บันทึก",            // B
        "รพ.สต. (สถานบริการ)",     // C
        "ผู้ให้บริการ",            // D
        "บริการ",                 // E
        "วันที่",                  // F
        "เวลาเริ่ม",               // G
        "เวลาสิ้นสุด",              // H
        "ชื่อผู้รับบริการ",           // I
        "เพศ",                   // J
        "อายุ",                   // K
        "อาชีพ",                  // L
        "เบอร์ติดต่อ",              // M
        "สิทธิการรักษา",            // N
        "จำนวนเงิน",               // O
        "สถานะ",                 // P
        "ไม่มีโรคประจำตัว",         // Q
        "โรคประจำตัว/ข้อควรระวัง",    // R
        "อาการเพิ่มเติม",           // S
        "ยินยอมให้เก็บข้อมูล"        // T
      ]);
    }
    
    var refNo = "KH-" + Math.floor(100000 + Math.random() * 900000);
    var timestamp = new Date();
    
    var rowData = [
      refNo,                                      // A
      timestamp,                                  // B
      e.hospital,                                 // C
      e.therapist,                                // D
      e.service,                                  // E
      e.date,                                     // F
      e.startTime,                                // G
      e.endTime || "16:00",                       // H
      e.fullName,                                 // I
      e.gender || "หญิง",                          // J
      e.age || "35",                              // K
      e.occupation || "เกษตรกรรม",                 // L
      e.phone,                                    // M
      e.rights,                                   // N
      e.price,                                    // O
      "รอรับบริการ",                               // P
      e.noMedicalCondition ? "ไม่มี" : "มีโรคประจำตัว", // Q
      e.medicalCondition || "-",                  // R
      e.symptoms || "-",                          // S
      e.consent ? "ยินยอม" : "ไม่ยินยอม"             // T
    ];
    
    sheet.appendRow(rowData);
    applyRowFormatting();
    
    return { success: true, refNo: refNo };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ดึงข้อมูลการจองทั้งหมดสำหรับแสดงผลและคำนวณ Dashboard
function getBookings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
  var rows = sheet.getDataRange().getValues();
  
  if (rows.length <= 1) return [];
  
  var data = [];
  
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    data.push({
      refNo: row[0],       // A
      createdTime: row[1], // B
      hospital: row[2],    // C
      therapist: row[3],   // D
      service: row[4],     // E
      date: row[5],        // F
      startTime: row[6],   // G
      endTime: row[7],     // H
      fullName: row[8],    // I
      gender: row[9] || "หญิง",     // J
      age: Number(row[10]) || 30,  // K
      occupation: row[11] || "อื่นๆ",// L
      phone: row[12],      // M
      rights: row[13],     // N
      price: Number(row[14]) || 0, // O
      status: row[15] || "รอรับบริการ", // P
      noCondition: row[16], // Q
      condition: row[17],   // R
      symptoms: row[18]    // S
    });
  }
  
  checkAndUpdatePastBookings();
  return data;
}

// ยกเลิกการจอง
function cancelBooking(refNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == refNo) {
        sheet.getRange(i + 1, 16).setValue("ยกเลิกแล้ว"); // คอลัมน์ P (Index 16)
        applyRowFormatting();
        return { success: true };
      }
    }
    return { success: false, message: "ไม่พบรหัสจองนี้ในระบบ" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// เช็กและอัปเดตคิวในอดีตเป็น "สำเร็จแล้ว"
function checkAndUpdatePastBookings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  var updated = false;
  for (var i = 1; i < data.length; i++) {
    var status = data[i][15]; // คอลัมน์ P (Index 15)
    var appointmentDateStr = data[i][5]; // คอลัมน์ F (Index 5)
    
    if (status === "รอรับบริการ" && appointmentDateStr) {
      var appDate = new Date(appointmentDateStr);
      appDate.setHours(0, 0, 0, 0);
      
      if (appDate < today) {
        sheet.getRange(i + 1, 16).setValue("สำเร็จแล้ว");
        updated = true;
      }
    }
  }
  
  if (updated) {
    applyRowFormatting();
  }
}

// ปรับสีตัวหนังสือทั้งแถวใน Google Sheet
function applyRowFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  if (lastRow <= 1) return;
  
  var statusRange = sheet.getRange(2, 16, lastRow - 1, 1);
  var statuses = statusRange.getValues();
  
  for (var i = 0; i < statuses.length; i++) {
    var status = statuses[i][0];
    var rowRange = sheet.getRange(i + 2, 1, 1, lastCol);
    
    rowRange.setFontColor("#000000");
    rowRange.setFontWeight("normal");
    rowRange.setFontStyle("normal");
    
    if (status === "รอรับบริการ") {
      rowRange.setFontColor("#2e7d32"); // สีเขียว
      rowRange.setFontWeight("bold");
    } else if (status === "ยกเลิกแล้ว") {
      rowRange.setFontColor("#c62828"); // สีแดง
      rowRange.setFontStyle("italic");
    } else if (status === "สำเร็จแล้ว") {
      rowRange.setFontColor("#0277bd"); // สีฟ้า
    }
  }
}
