// sheet 有四個頁籤 (簽到頁, 報名名單, 補登人員, db)，詳見範例檔
// 適用讀取後輸出 16 進位外碼的讀卡機

function onEdit(e) {
  var ss = SpreadsheetApp.getActiveSheet();
  var r = ss.getActiveCell();

  if (r.getColumn() < 5 && ss.getName() == '簽到頁') {
    
    // sheet
    var enrolledsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('報名名單');
    var postenrollsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('補登人員');
    var dbsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('db');
    
    // input value
    var rid = ss.getActiveCell().getValue();
    var rrow = r.getRowIndex();
    
    // name of the id
    if (dbsheet.createTextFinder(rid).findNext())
      var name = dbsheet.getRange('A'+dbsheet.createTextFinder(rid).findNext().getRow()).getValue();
    else
      var name = '查無此人';
    
    // fill name and timestamp
    var tscell ='E' + rrow;
    var namecell = 'B' + rrow;
    var nowts = new Date();
    ss.getRange(tscell).setValue(nowts).setNumberFormat("yyyy/MM/dd hh:mm");
    ss.getRange(namecell).setValue(name);
    
    // fill timestamp to enrolledsheet
    if (name != '查無此人') {
      // enrolled
      if (enrolledsheet.createTextFinder(name).findNext()) {
        enrolledsheet.getRange('F'+enrolledsheet.createTextFinder(name).findNext().getRow()).setValue(nowts).setNumberFormat("yyyy/MM/dd hh:mm");
      } else { // post enroll
        var funcno = enrolledsheet.getRange('A2').getValue();
        var last = postenrollsheet.getLastRow() + 1;
        var pecell = 'D' + rrow;
        ss.getRange(pecell).setValue('TRUE');
        ss.getRange(SpreadsheetApp.getActiveRange().getRow(), 1, 1, ss.getLastColumn()).setBackground('#e6b8af');
        postenrollsheet.getRange('A' + last).setValue(funcno);
        postenrollsheet.getRange('B' + last).setValue('OO國小');
        postenrollsheet.getRange('C' + last).setValue(name);
        postenrollsheet.getRange('E' + last).setValue(nowts).setNumberFormat("yyyy/MM/dd hh:mm");
      }
    }
  }
};
