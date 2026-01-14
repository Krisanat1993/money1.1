// ID ของ Spreadsheet และ Folder ที่คุณระบุ
const SPREADSHEET_ID = '1UlQqKGGrX6bS0jjI-SGcRjJWuL3ownITAjda5yCEkjc';
const FOLDER_ID = '1gjV9hOJygivAHMEFlsSlgQCH66wKt5vQ';
const SHEET_DATA_NAME = 'รวมข้อมูล'; // sheet1
const SHEET_SUMMARY_NAME = 'สรุป';     // sheet2

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบแจ้งโอนเงิน')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชันดึงข้อมูลสำหรับหน้าแรก (Page 1)
function getSummaryData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SUMMARY_NAME);
  const data = sheet.getDataRange().getValues();
  
  // ตัด Header ออก (สมมติว่าแถวแรกเป็น Header)
  const rows = data.slice(1);
  
  // ดึงข้อมูล Col B(1)=ยศ ชื่อ, C(2)=นามสกุล, F(5)=จ่ายแล้ว
  // หมายเหตุ: Array เริ่มนับที่ 0 ดังนั้น Col A=0, B=1, ...
  const result = rows.map(row => {
    return {
      rankName: row[1], // Col B
      lastName: row[2], // Col C
      paid: row[5]      // Col F
    };
  }).filter(item => item.rankName != ""); // กรองแถวว่างออก
  
  return result;
}

// ฟังก์ชันดึงรายชื่อสำหรับ Dropdown ในหน้าที่ 2
function getDropdownList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SUMMARY_NAME);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // ตัด Header
  
  // ส่งกลับเป็น list ของชื่อเต็มเพื่อใส่ใน dropdown
  return rows.map(row => {
    return {
      text: `${row[1]} ${row[2]}`, // แสดง "ยศ ชื่อ นามสกุล"
      val_rankName: row[1],
      val_lastName: row[2]
    };
  }).filter(item => item.val_rankName != "");
}

// ฟังก์ชันบันทึกข้อมูล (Page 2 Submit)
function saveTransferData(formObj) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_DATA_NAME);
    const folder = DriveApp.getFolderById(FOLDER_ID);
    
    let fileUrl = "";
    
    // จัดการไฟล์อัปโหลด
    if (formObj.fileData && formObj.fileName && formObj.mimeType) {
      const data = Utilities.base64Decode(formObj.fileData);
      const blob = Utilities.newBlob(data, formObj.mimeType, formObj.fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }
    
    // แยกชื่อและนามสกุลจาก Dropdown value (ส่งมาเป็น JSON string หรือจัดการที่ฝั่ง client)
    // ในที่นี้รับค่าแยกมาเลยเพื่อความชัวร์
    const rankName = formObj.rankName;
    const lastName = formObj.lastName;
    const amount = formObj.amount;
    const timestamp = new Date();
    
    // บันทึกลง Sheet: Timestamp, ยศ ชื่อ, นามสกุล, ยอดเงิน, หลักฐาน
    sheet.appendRow([timestamp, rankName, lastName, amount, fileUrl]);
    
    return { success: true };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
