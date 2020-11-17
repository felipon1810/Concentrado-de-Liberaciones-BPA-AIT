function fixAllExcel() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  spreadsheet.setActiveSheet(sheet);
  
  setFunctionQInProgress(spreadsheet);
  setFormatTitle(spreadsheet);
  setFormatText(spreadsheet);
  setFormatDate(spreadsheet);
  setConditionalFormatRules(spreadsheet);
};

function fixOnlyQInProgress() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  spreadsheet.setActiveSheet(sheet);
  
  setFunctionQInProgress(spreadsheet);
};

function fixOnlyFormatExcel() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  spreadsheet.setActiveSheet(sheet);
  
  setFormatTitle(spreadsheet);
  setFormatText(spreadsheet);
  setFormatDate(spreadsheet);
};

function fixOnlyConditionalFormatRule() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  spreadsheet.setActiveSheet(sheet);
  
  setConditionalFormatRules(spreadsheet);
};

//-----------------------------------------------------------------------------------------

// Inserta y corrige la funcion para calcular el Q en curso
function setFunctionQInProgress(spreadsheet) {
  spreadsheet.getRangeList(rangeForQColumn).activate().setFormulaR1C1(
    '=IF(AND(R[0]C[2]>=DATEVALUE("2020-09-18"),R[0]C[2]<DATEVALUE("2020-12-15")),"Q4", \n' +
    '    IF(AND(R[0]C[2]>=DATEVALUE("2020-06-19"),R[0]C[2]<DATEVALUE("2020-09-18")),"Q3", \n' +
    '       IF(AND(R[0]C[2]>=DATEVALUE("2020-03-20"),R[0]C[2]<DATEVALUE("2020-06-19")),"Q2", \n' +
    '          IF(AND(R[0]C[2]>=DATEVALUE("2019-12-17"),R[0]C[2]<DATEVALUE("2020-03-20")),"Q1", \n' +
    '             "SIN DEFINIR" \n' + 
    '          ) \n' +
    '       ) \n' +
    '    ) \n' +
    ' )'
  );
};

// Cambia el formato del texto
function setFormatTitle(spreadsheet) {
  spreadsheet.getRangeList(rangeForTitleColumn).activate()
     .setFontFamily('Verdana')
     .setFontSize(9)
     .setFontWeight('bold')
     .setFontStyle(null)
     .setFontLine(null)
     .setFontColor('#ffffff')
     .setBackground('#073763')
//     .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle')
     .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
};

// Cambia el formato del texto
function setFormatText(spreadsheet) {
  spreadsheet.getRangeList(rangeForDataColumn).activate()
     .setFontFamily('Verdana')
     .setFontSize(7)
     .setFontWeight(null)
     .setFontStyle(null)
     .setFontLine(null)
     .setFontColor('#000000')
     .setBackground(null)
//     .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle')
     .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
};

// Cambia el formato de fechas para las diferentes columnas
function setFormatDate(spreadsheet) {
  spreadsheet.getRangeList(rangeForDateDolumn).activate()
     .setNumberFormat('yyyy"-"mm"-"dd" "hh":"mm');
};

// Crea de nuevo los formatos condicionales
function setConditionalFormatRules(spreadsheet) {
  var conditionalFormatRules;
  spreadsheet.getRangeList(rangeForDataColumn).activate();
  clearFormatting(spreadsheet);
 
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
     .setRanges([spreadsheet.getRange(rangeForDataColumn)])
     .whenFormulaSatisfied('=$'+colEstatusIni+'="VOBO-TOTAL"')
     .setBackground('#b7e1cd')
     .setFontColor('#000000')
     .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
     .setRanges([spreadsheet.getRange(rangeForDataColumn)])
     .whenFormulaSatisfied('=$'+colEstatusIni+'="VOBO-PARCIAL"')
     .setBackground('#b7e1cd')
     .setFontColor('#a61c00')
     .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
     .setRanges([spreadsheet.getRange(rangeForDataColumn)])
     .whenFormulaSatisfied('=$'+colEstatusIni+'="VOBO-QA"')
     .setBackground('#b7e1cd')
     .setFontColor('#6a1b9a')
     .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
     .setRanges([spreadsheet.getRange(rangeForDataColumn)])
     .whenFormulaSatisfied('=$'+colEstatusIni+'="ASIGNADO"')
     .setBackground('#FCE8B2')
     .setFontColor('#000000')
     .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);

  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
     .setRanges([spreadsheet.getRange(rangeForDataColumn)])
     .whenFormulaSatisfied('=$'+colEstatusIni+'="CANCELADO"')
     .setBackground('#c0b4de')
     .setFontColor('#000000')
     .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
     .setRanges([spreadsheet.getRange(rangeForDataColumn)])
     .whenFormulaSatisfied('=$'+colEstatusIni+'="RECHAZADO"')
     .setBackground('#dd7e6b')
     .setFontColor('#000000')
     .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
     .setRanges([spreadsheet.getRange(rangeForDataColumn)])
     .whenFormulaSatisfied('=$'+colEstatusIni+'="RETORNADO"')
     .setBackground('#a61c00')
     .setFontColor('#ffffff')
     .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
     .setRanges([spreadsheet.getRange(rangeForDataColumn)])
     .whenFormulaSatisfied('=$'+colEstatusIni+'="RETORNADO-PARCIAL"')
     .setBackground('#a61c00')
     .setFontColor('#ffd54f')
     .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
};

// Elimina todas los formatos condicionales
function clearFormatting (spreadsheet) {
  var ss = spreadsheet.getActiveSheet();
  ss.clearConditionalFormatRules();
}