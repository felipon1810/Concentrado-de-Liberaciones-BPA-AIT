

function NUEVOFOLIOQA_A() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRange().setFormulaR1C1('=CONCATENATE(DATEVALUE(R[0]C[1]),HOUR(R[0]C[1]),MINUTE(R[0]C[1]),SECOND(R[0]C[1]),"-",R[0]C[5])');
  spreadsheet.getCurrentCell().offset(0, -1).activate()
  .setFormulaR1C1('=IF(COUNTIF(C2:C2,R[0]C[1])>1,"SI","NO")');
  
  FormatoTexto()
  FormatoFecha()
  
  
   spreadsheet.getRange('B:B').activate()
  .setHorizontalAlignment('center');  
  
};



function NuevoFolioQAa() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getCurrentCell().setFormulaR1C1('=CONCATENAR(FILA()-2, "-", IZQUIERDA(SUSTITUIR(R[0]C[1],"-",""), 8))')
  .setFontWeight('bold');
  
  FormatoTexto()
  FormatoFecha()
  
   spreadsheet.getRange('B:B').activate()
  .setHorizontalAlignment('center');
};


/*Cambia el formato de fechas para las diferentes columnas*/
function FormatoFecha() {
  var spreadsheet = SpreadsheetApp.getActive(); 
  spreadsheet.getRangeList(['C:C', 'AD:AD', 'AE:AE']).activate()
  .setNumberFormat('yyyy"-"mm"-"dd" "h":"mm');  
  
  
  spreadsheet.getRangeList(['AL:AL', 'AK:AK']).activate()
  .setNumberFormat('yyyy"-"mm"-"dd');  
};

/*Cambia el formato de texto de toda la fila*/
function FormatoTexto() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(spreadsheet.getCurrentCell().getRow(), 1, 1, sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRangeList().setFontFamily('Arial')
  .setFontSize(8)
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  
  /*Cambia el aliniamiento de las columnas*/
  spreadsheet.getRangeList(['AJ:AJ', 'AR:AR']).activate()
   .setHorizontalAlignment('left')
   .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
};

/*

function SeleccionarColumnas() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRangeList(['B:B', 'F:F']).activate()
  .setNumberFormat('yyyy"-"mm"-"dd" "h":"mm');
};


function Formatosdetexto() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A9').activate();
  spreadsheet.getCurrentCell().setValue('asdasdasd');
  spreadsheet.getActiveRangeList().setFontFamily('Arial')
  .setFontSize(9)
  .setFontWeight('bold')
  .setFontStyle('italic')
  .setFontLine('line-through')
  .setFontColor('#ffffff')
  .setBackground('#1c4587')
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  .setVerticalText(true);
};

function Formatosdeparatexto() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B8').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('General')
  .setNumberFormat('@')
  .setNumberFormat('#,##0.00')
  .setNumberFormat('0.00%')
  .setNumberFormat('0.00E+00')
  .setNumberFormat('_-* #,##0.00\\ [$€-1]_-;\\-* #,##0.00\\ [$€-1]_-;_-* "-"??\\ [$€-1]_-;_-@')
  .setNumberFormat('#,##0.00;(#,##0.00)')
  .setNumberFormat('#,##0.00\\ [$€-1]')
  .setNumberFormat('#,##0\\ [$€-1]')
  .setNumberFormat('d/MM/yyyy')
  .setNumberFormat('H:mm:ss')
  .setNumberFormat('d/MM/yyyy H:mm:ss')
  .setNumberFormat('[h]:mm:ss')
  .setNumberFormat('yyyy"-"mm"-"dd')
  .setNumberFormat('yyyy"-"mm"-"dd" "h":"mm');
};


 spreadsheet.getCurrentCell().copyTo(spreadsheet.getCurrentCell().setFormulaR1C1('=CONCATENAR(FILA()-2, "-", IZQUIERDA(SUSTITUIR(R[0]C[1],"-",""), 8))'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

*/






function getDatosBacklogg() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D5').activate();
  spreadsheet.getCurrentCell().setFormula('=BUSCARV(A5,\'2020 Bakcup\'!B3:AR,2,FALSO)');
  spreadsheet.getRange('E5').activate();
  spreadsheet.getCurrentCell().setFormula('=BUSCARV(A5,\'2020 Bakcup\'!B3:AR,3,FALSO)');
};





function Fechasinformato() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D:D').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('@');
  spreadsheet.getRange('W:W').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('@');
  spreadsheet.getRange('X:X').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('@');
  spreadsheet.getRange('AE:AE').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('@');
  spreadsheet.getRange('AF:AF').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('@');
  spreadsheet.getRange('A4').activate();
};

