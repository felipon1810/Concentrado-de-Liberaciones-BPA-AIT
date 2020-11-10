function loadConfiguracion(){
  //delAllScriptProperties
  var scriptProperties = PropertiesService.getScriptProperties();  
  var data = getConfiguracionSite();
  Logger.log(data)
  data.forEach(function(r){
    if(r[0] !== ""){
      //Logger.log(r[0]+":"+r[1])
      scriptProperties.setProperty(r[0], r[1]);
     }
  });
}

function delAllScriptProperties(){
  // Delete all properties in the current script.
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteAllProperties();
}

//Obtener Hoja Configuración Site
function getConfiguracionSite(){
  var spSheet = SpreadsheetApp.openById("10tJo2FVV1c-mOypTlQUO7sOelQ6OsiNkd9Y_3jhKKCc");
  var sheet = spSheet.getSheetByName("Configuración");
  //sheet.getRange(row, column, numRows, numColumns)
  var data= sheet.getRange(2,2, sheet.getLastRow()-1,2).getValues();
  
  return data;
}

/*----------------------------*/
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function appendLeadingZeroes(n){
  if(n <= 9){
    return "0" + n;
  }
  return n
}

function showMesage(mesage){
  Browser.msgBox(mesage);
}

function getPossitionRegistrySelected(){
  var spSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spSheet.getSheetByName("2020");
  
  var pos = sheet.getActiveCell().getRowIndex();
  Logger.log("possiton: "+pos)
  
}
