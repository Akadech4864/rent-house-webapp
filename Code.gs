const SHEET_NAME = 'ระบบบ้านเช่า_บันทึกรับจ่าย'; 

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('ระบบจัดการค่าเช่าบ้าน')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0');
}

function getPendingPayment() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return { error: 'ไม่พบหน้า Sheet ชื่อ ' + SHEET_NAME };

    const data = sheet.getDataRange().getValues();
    let unpaidRows = [];
    let firstRecordDateDisplay = '';
    let firstRecordDateIso = '';

    for (let i = 1; i < data.length; i++) {
      const status = data[i][6]; 
      
      if (status === false || status === '' || status === 'FALSE') {
        let recordDate = data[i][0];
        unpaidRows.push(i + 1); 

        if (unpaidRows.length === 1 && recordDate instanceof Date) {
          firstRecordDateDisplay = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
          firstRecordDateIso = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
      }
    }

    if (unpaidRows.length > 0) {
      return {
        unpaidRows: unpaidRows,             
        monthsCount: unpaidRows.length,     
        recordDateDisplay: firstRecordDateDisplay,
        recordDateIso: firstRecordDateIso,
        baseRentPerMonth: 27800,
        error: null
      };
    }

    return { error: 'ไม่มีรอบบิลที่ค้างชำระในขณะนี้' };
  } catch (e) {
    return { error: e.toString() };
  }
}

function saveRecord(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, message: 'ไม่พบ Sheet' };
    
    let slipUrl = '';
    
    if (data.slipFile && data.slipFile.data) {
      try {
        const folderId = '1RnUiAJPajfxHDwFEWlUsrfWI11bD0wl-';
        const folder = DriveApp.getFolderById(folderId);
        const blob = Utilities.newBlob(Utilities.base64Decode(data.slipFile.data), data.slipFile.mimeType, data.slipFile.name);
        const file = folder.createFile(blob);
        slipUrl = file.getUrl();
      } catch (uploadError) {
        return { success: false, message: 'บันทึกไฟล์สลิปไม่สำเร็จ: ' + uploadError.toString() };
      }
    }
    
    const rows = data.unpaidRows;
    
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      sheet.getRange(r, 2).setValue(data.paymentDate); 
      sheet.getRange(r, 3).setValue(data.baseRentPerMonth); 
      
      if (i === 0) {
        sheet.getRange(r, 4).setValue(data.penalty); 
        sheet.getRange(r, 5).setValue(data.total); // E: ยอดรวมปกติ (ค่าเช่า + ค่าปรับ)
        sheet.getRange(r, 6).setValue(data.days); // F: วันล่าช้า
        sheet.getRange(r, 10).setValue(data.extraPayment || 0); // J: ยอดชำระนอกเหนือที่ตกลง
      } else {
        sheet.getRange(r, 4).setValue(0);
        sheet.getRange(r, 5).setValue(data.baseRentPerMonth);
        sheet.getRange(r, 6).setValue(0); 
        sheet.getRange(r, 10).setValue(0); 
      }
      
      let note = data.note;
      if (rows.length > 1) {
        note = note ? note + ` (ชำระรวม ${rows.length} เดือน)` : `ชำระรวม ${rows.length} เดือน`;
      }
      
      sheet.getRange(r, 7).setValue(true); 
      sheet.getRange(r, 8).setValue(slipUrl); 
      sheet.getRange(r, 9).setValue(note); 
    }
    
    return { success: true, message: 'บันทึกการชำระเงินเรียบร้อยแล้ว' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getPaymentHistory() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return { error: 'ไม่พบหน้า Sheet' };

    const data = sheet.getDataRange().getValues();
    const history = [];

    for (let i = data.length - 1; i >= 1; i--) {
      const status = data[i][6];
      if (status === true || status === 'TRUE') {
        let recordDate = data[i][0];
        let payDate = data[i][1];
        
        if (recordDate instanceof Date) recordDate = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
        if (payDate instanceof Date || typeof payDate === 'string') {
           const d = new Date(payDate);
           if(!isNaN(d)) payDate = Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
        }

        history.push({
          recordDate: recordDate || '-',
          payDate: payDate || '-',
          baseRent: data[i][2] || 0,
          penalty: data[i][3] || 0,
          total: data[i][4] || 0,
          days: data[i][5] || 0,
          slipUrl: data[i][7] || '',
          note: data[i][8] || '', 
          extraPayment: data[i][9] || 0 // J
        });
      }
    }
    return { success: true, data: history };
  } catch (e) {
    return { error: e.toString() };
  }
}

function getContractFiles() {
  try {
    const folderId = '1-IOLf01Wa4XumpNiSSlQlFDBacsJ_KDS';
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const fileList = [];

    while (files.hasNext()) {
      const file = files.next();
      fileList.push({
        name: file.getName(),
        url: file.getUrl() 
      });
    }
    return { success: true, data: fileList };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
