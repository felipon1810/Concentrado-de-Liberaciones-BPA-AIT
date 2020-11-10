function onOpen(e) {
  SpreadsheetApp.getUi()
     .createMenu('BEx QA & Production')
     .addItem('Open MultiSelect', 'showDialog')
     .addSeparator()
     .addItem('Send Checklist for Email (sheet!'+sheetName+')', 'mainSendEmail')
     .addSeparator()
     .addItem('Fix All Excel (sheet!'+sheetName+')', 'fixAllExcel')
     .addItem('Fix Only Duplicate (sheet!'+sheetName+')', 'fixOnlyDuplicate')
     .addItem('Fix Only Q in Progress (sheet!'+sheetName+')', 'fixOnlyQInProgress')
     .addItem('Fix Only Format Excel (sheet!'+sheetName+')', 'fixOnlyFormatExcel')
     .addItem('Fix Only Conditional Format Rule (sheet!'+sheetName+')', 'fixOnlyConditionalFormatRule')
     .addToUi();
  SpreadsheetApp.getActiveSpreadsheet().toast('Menú de BEx QA & Production en ejecución', 'AVISO', 5);
}