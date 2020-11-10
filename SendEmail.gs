const  configProperties = PropertiesService.getScriptProperties();

function mainSendEmail(){
  var rData = getRowDataSelected("")
  
  if(rData.nameChecklist != "" && rData.statusChecklist == "EN PROCESO"){
    Browser.msgBox('El Folio QA: '+rData.folioqa + ', ya fue tomado por: '+rData.frenteQA);
    return
  }
  
  if(rData.nameChecklist != "" && rData.statusChecklist == "ENVIADO"){
    Browser.msgBox('El Checklist ya fue generado y enviado por email\n para el siguiente Folio QA: '+rData.folioqa);
    return
  }
  
  if(rData.nameChecklist != "" && rData.statusChecklist != "ENVIADO"){
    Browser.msgBox('El Checklist ya fue generado para el siguiente Folio QA: '+rData.folioqa);
    var btnPropertie = 'disabled'
    showSendEmailForm(rData,btnPropertie)
  }
  
  if(rData.nameChecklist == "" && rData.statusChecklist != 'ENVIADO'){
    var btnPropertie = ''
    showSendEmailForm(rData,btnPropertie)
  }
}

function showSendEmailForm(rData,btnPropertie){
  var template = HtmlService.createTemplateFromFile("sendEmailForm");
  
  var email_bexqa = configProperties.getProperty("email_bex-qa")
  
  var htmlEmailContact = getEmailsContact(rData)
  
  template.folioQA = rData.folioqa
  template.app     = rData.aplicacion
  template.emailTo = rData.emailto
  template.emailQA = email_bexqa
  template.emailsContacts = htmlEmailContact
  template.emailReplyTo = email_bexqa
  template.nameChecklist = rData.nameChecklist
  template.urlProcesoqa = rData.urlProcesoqa
  template.btnPropertie = btnPropertie
  
  var btnPropertieUrl = rData.urlProcesoqa.trim() != "" ?  btnPropertieUrl = 'disabled' : "" 
  var btnPropertieSend = rData.version > 1 ? 'disabled' : ""
  
  template.versionChecklist = rData.version  
  template.btnPropertieUrl = btnPropertieUrl
  template.idDoc = rData.idDoc
  template.btnPropertieSend = btnPropertieSend
    
  var html = template.evaluate();
  html.setTitle("BEx-QA").setWidth(850).setHeight(450);
  
  SpreadsheetApp.getUi().showModalDialog(html, "BEx-QA");
}

function testMethods(){
   /*var data = getRowDataSelected(124)
   Logger.log(data.objdesarrollo)
   Logger.log("rowFolio"+configProperties.getProperty("rowFolioQAChecklist"))
   Logger.log("hoja: "+configProperties.getProperty("hoja_Concentrado_Liberaciones"))
   Logger.log("url: "+data.objdesarrollo)*/
   

}
//configProperties.getProperty("hoja_Concentrado_Liberaciones")
function getRowDataSelected(row){
  var objData = {}
  var spSheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spSheet.getSheetByName("2020")
  Logger.log("method: getRowDataSelected ")
  var data = row == "" ? sheet.getRange(spSheet.getCurrentCell().getRow(), 1, 1, sheet.getMaxColumns()).getValues() : sheet.getRange(row, 1, 1, sheet.getMaxColumns()).getValues();
  //sheet.getRange(row, column, numRows, numColumns)
  objData.folioqa = data[0][configProperties.getProperty("posFolioQA")-1]
  objData.version = data[0][configProperties.getProperty("posVersion")-1]
  objData.frenteQA = data[0][configProperties.getProperty("posFrenteQA")-1]
  objData.emailto = data[0][configProperties.getProperty("posQuienRegistro")-1]
  objData.uuaa = data[0][configProperties.getProperty("posUUAA")-1]
  objData.aplicacion = data[0][configProperties.getProperty("posAPP")-1]
  objData.nameLider = data[0][configProperties.getProperty("posLider")-1]
  objData.nameFrente = data[0][configProperties.getProperty("posFrenteDev")-1]
  objData.nameTeamDevelop = data[0][configProperties.getProperty("posTeamDev")-1]
  objData.objdesarrollo = data[0][configProperties.getProperty("posObjetivo")-1]
  objData.statusChecklist = data[0][configProperties.getProperty("posEstatusCheck")-1]
  objData.urlProcesoqa = data[0][configProperties.getProperty("posUrlProcesoQA")-1]
  objData.nameChecklist = data[0][configProperties.getProperty("posNombreCheck")-1]
  
  objData.idDoc = objData.urlProcesoqa != "" ? objData.urlProcesoqa.slice(39, -18) : objData.idDoc = ""
    
  return objData
}

function getEmailsContact(objData){  
  var names = ""
  var emails = ""
  
  if(objData.nameLider.trim()){names = objData.nameLider+','}
  if(objData.nameFrente.trim()){names += objData.nameFrente}
  if(objData.nameTeamDevelop.trim()){names += ','+objData.nameTeamDevelop}
  
  if(names.trim()){
    var arrayNames = names.split(",");  
    for (var i=0; i < arrayNames.length; i++) {
       var c = arrayNames[i]
          .normalize('NFD')
          .replace(/([^n\u0300-\u036f]|n(?!\u0303(?![\u0300-\u036f])))[\u0300-\u036f]+/gi,"$1")
          .normalize(); 
       emails += searchMail(c)
       emails = i != arrayNames.length-1 ? emails+',': emails;
    }
    return emails
  }
  return emails = "Sin nombres para buscar emails"
}

// The code below retrieves a contact named "John Doe" and logs the email addresses
// associated with that contact
function searchMail(c){
  var contacts = ""
  contacts = ContactsApp.getContactsByName(c);
  
  if(contacts != ""){
    var emails = contacts[0].getEmails();
    for (var k in emails) {
      return emails[k].getAddress()
    }
  }
  return c
}

//Funcion para Crear Copia de Checklist
function copySheet(folioqa,app,version){
  Logger.log("Version: "+version)
    
  var nombrePlantilla = configProperties.getProperty("nombre_Plantilla");
  var nombreCopia = configProperties.getProperty("nombre_Copia_Plantilla");
  var idFolderDestino = configProperties.getProperty("idFolder_Repo_Destino");
  var idFolderOrigen = configProperties.getProperty("idFolder_Plantilla");
  
  var nombreArchivo = nombreCopia.replace("{FolioQA}",folioqa).replace("{APP}",app).replace("{Version}",version);
  var dApp= DriveApp;
  //Se obtiene el folder donde se encuentra la plantilla
  var folderOrigen = dApp.getFolderById(idFolderOrigen);
  var filesIter = folderOrigen.getFiles();
  var folderDestino = dApp.getFolderById(idFolderDestino);
  
  //Recorremos el folder y copiamos la ´plantilla
  while(filesIter.hasNext()){
    var file = filesIter.next();
    var filename= file.getName();
    
    if(filename == nombrePlantilla){
      file.makeCopy(nombreArchivo, folderDestino);
      return nombreArchivo;
    }
  }
  return "El nombre de la plantilla no fue encontrado.";
}

function copysheetExistente(folioqa,version,app){ 
  Logger.log("Version: "+version)
  
  var idFolderDestino = configProperties.getProperty("idFolder_Repo_Destino");
  var nombreCopia = configProperties.getProperty("nombre_Copia_Plantilla");
  
  var versionAntigua = version -1; 
  var nombreArchivo = nombreCopia.replace("{FolioQA}",folioqa).replace("{APP}",app).replace("{Version}",versionAntigua);
  Logger.log("nombreArchivo: "+nombreArchivo)
  try {
    var idFile = DriveApp.getFolderById(idFolderDestino).getFilesByName(nombreArchivo).next().getId();
    Logger.log("idFile: "+idFile)
  } catch (e) {
     // Logs an ERROR message.
     console.error('Error al encontrar Archivo, Iniciando diferente merodod de Busqueda: ' + e);
     var spSheet = SpreadsheetApp.getActive()
     var sheet = spSheet.getActiveSheet()
     var folioConcatenado = folioqa+"_V"+versionAntigua;
     console.error('Buscando Folio... : '+folioConcatenado );
     var row = getRowFolioQA(folioConcatenado)
     //var row = getRowFolioQAConcatenado(folioConcatenado)
     var column = configProperties.getProperty("posUrlProcesoQA") //URL PROCESO QA
     var url = sheet.getRange(row, column).getValue();
     url = url.substring(39, 83);
     Logger.log("URL: "+ url);
     idFile= url;
  }
  if(idFile != ""){
    var file =  DriveApp.getFileById(idFile)
    var folderDestino = DriveApp.getFolderById(idFolderDestino);
    nombreArchivo = nombreCopia.replace("{FolioQA}",folioqa).replace("{APP}",app).replace("{Version}",version);
    Logger.log("nombreArchivo: "+nombreArchivo)
    
    file.makeCopy(nombreArchivo, folderDestino);
    return nombreArchivo;
  }else {
    Logger.log("No se encontro un Folio QA Relacionado")
  }
}

//Funcion para obtener el id de la copia creada
function getDetailsChecklist(nombreChecklist){
  var idFolderDestino = configProperties.getProperty("idFolder_Repo_Destino")
  onSubmit()
  var folderIter = DriveApp.getFolderById(idFolderDestino)
 
  var idCopia = ""
  var url = ""
  
  var filesIter = folderIter.getFiles()  
  var data = {}
  
  if(nombreChecklist !== ""){
    //Recorremos el folder y buscamos el ID
    while(filesIter.hasNext()){
      var file = filesIter.next()
      if(file.getName() == nombreChecklist){
         
         data.nombreChecklist = nombreChecklist
         //Obtenemos el ID
         if(file.getId() != null){           
           data.idDoc = file.getId()
           data.urlDoc = file.getUrl()
           
           return data
         }else{
            data.nombreChecklist = ""
            data.idDoc = ""
            data.urlDoc = ""
            return data
         }
      }
    }
  }
  return data;
}

function setFoliQAchecklist(data){
  //Guarda el FolioQa dentro del checklist generado por medio del ID.
  var spSheet = SpreadsheetApp.openById(data.idDoc)
  var sheet   = spSheet.getSheetByName(configProperties.getProperty("hojaInfoChecklist"));
  
  var rowFolio = configProperties.getProperty("rowFolioQAChecklist")
  var rowVersion = configProperties.getProperty("rowVersionChecklist")
  var column =  configProperties.getProperty("colFolioVersionChecklist")
  
  sheet.getRange(rowFolio, column).setValue(data.folioqa)
  sheet.getRange(rowVersion, column).setValue(data.version)
}

function setUrlConcentrado(data){
  var spSheet = SpreadsheetApp.getActive()
  var sheet = spSheet.getActiveSheet()
  
  var row =  sheet.getActiveCell().getRowIndex()  
  var column = configProperties.getProperty("posUrlProcesoQA") //URL PROCESO QA
  
  sheet.getRange(row, column).setValue(data.urlDoc)
}

function setNameChecklistConcentrado(data){
  var spSheet = SpreadsheetApp.getActive()
  var sheet = spSheet.getActiveSheet()
  
  var row = sheet.getActiveCell().getRowIndex();              //Obtiene la fila de la celda seleccionada
  var column = configProperties.getProperty("posNombreCheck") //NOMBRE CHECKLIST
  
  sheet.getRange(row, column).setValue(data.nombreChecklist)
}

function setStatusChecklist(status){
  var spSheet = SpreadsheetApp.getActive()
  var sheet = spSheet.getActiveSheet()
  
  var row = sheet.getActiveCell().getRowIndex();    
  var column = configProperties.getProperty("posEstatusCheck") //ESTATUS Checklist
  
  sheet.getRange(row, column).setValue(status)
}

function setDataQA(folioqa,version){
  var spSheet = SpreadsheetApp.getActive()
  var sheet = spSheet.getActiveSheet()
  
  var row = sheet.getActiveCell().getRowIndex()   
  var current_datetime = new Date()
  
  var fecharegistro = getFechaRegistro(current_datetime);
  sheet.getRange(row, configProperties.getProperty("posEstatusCheck")).setValue('ENVIADO')//ESTATUS CHECKLIST 
  sheet.getRange(row, configProperties.getProperty("posFhRevisionIni")).setValue(fecharegistro) //FECHA DE INICIO DE REVISIÓN
  sheet.getRange(row, configProperties.getProperty("posEstatusFinal")).setValue('ASIGNADO') //ESTATUS FINAL DEL CAMBIO
}

//Funcion de Tiempo
function onSubmit() {
  // Se llema a la funcion Tiempo
  Logger.log("Entro en el Time");
  Utilities.sleep(5 * 1000);
  Logger.log("Salio en el Time")  
}

//Obtiene la el numero de fila donde se encuentra el folio QA
function getRowFolioQA(folio){
  var spSheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spSheet.getSheetByName("2020")
  
  var row = configProperties.getProperty("hoja_Concentrado_Liberaciones_row")
  
  var data = sheet.getRange(row, 1, sheet.getLastRow(), 1).getValues()
  var folioQAList = data.map(function(r){return r[0];})
  
  var position = folioQAList.indexOf(folio)
  if(position > -1){
    position = position + parseInt(row)
    return position
  } 
  Logger.log("Folio no encontrado: "+folio)
  return position
}
//Obtiene la el numero de fila donde se encuentra el folio QA_VERSION
function getRowFolioQAConcatenado(folioConcatenado){
  var spSheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spSheet.getSheetByName("2020")
  
  var row = configProperties.getProperty("hoja_Concentrado_Liberaciones_row")
  //Logger.log("Folio Concatenado: "+folioConcatenado);
  var data = sheet.getRange(row, 1, sheet.getLastRow(), 1).getValues()
  var folioQAList = data.map(function(r){return r[0];})
  //Logger.log("Folios: "+folioQAList);
  var position = folioQAList.indexOf(folioConcatenado)
  if(position > -1){
    position = position + parseInt(row)
    Logger.log("Posicion Folio: "+position)
    return position
  } 
  Logger.log("Folio no encontrado: "+folioConcatenado)
  return position
}

/*CREATE AND SEND DOCUMENT BY GMAIL*/
function createEmail(dataForm) {
  Logger.log("method: createEmail ")
  var email_bexqa = configProperties.getProperty("email_bex-qa")
  
  var subject ="[BEx QA & Production]: FolioQA ["+dataForm.folioqa+"] - ["+dataForm.aplicacion+"]";
  var spSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spSheet.getSheetByName(configProperties.getProperty("hoja_Concentrado_Liberaciones"));
  
  var pos = sheet.getActiveCell().getRowIndex();
  Logger.log("possiton: "+pos)
  
  var data = getRowDataSelected(pos) //getRowDataSelected(getRowFolioQA(dataForm.folioqa)) ;
   
  if(dataForm.urlDoc !== "" && email_bexqa !== "" && data!==""){
    Logger.log("createEmail...")
    var htmlTemplate = HtmlService.createTemplateFromFile("sendEmailTemplate")     
    
    htmlTemplate.fnfolioqa = data.folioqa;
    htmlTemplate.fnUUAA = data.uuaa;
    htmlTemplate.fnAplicacion = data.aplicacion;
    htmlTemplate.fnObjetivo = data.objdesarrollo;
    htmlTemplate.fnurl = dataForm.urlDoc;
    
    var body = ""
    var htmlBody = htmlTemplate.evaluate().getContent();
       
    GmailApp.sendEmail(data.emailto, subject, body,{htmlBody:htmlBody,cc:email_bexqa+','+dataForm.emailcc,replyTo:email_bexqa});
    setDataQA() 
    Logger.log("End createEmail...")
    setAddCommenter(dataForm)
  }else{
     Logger.log("No entro a createEmail...")
  }
}

function setAddCommenter(data){
  var emails = data.emailto+','+data.emailcc;
  var arrayEmails = emails.split(",");
  for(var i=0; i < arrayEmails.length; i++) {
      var mail = arrayEmails[i]
      if(mail != ""){
         Logger.log("enviando email: "+mail)
         DriveApp.getFileById(data.IdDoc).addCommenter(mail)
      }
  }
}

//Función para dar formato a la fecha de registro.
function getFechaRegistro(current_datetime){
  var formatted_date = current_datetime.getFullYear()+ "-" + appendLeadingZeroes(current_datetime.getMonth() + 1) + "-" + appendLeadingZeroes(current_datetime.getDate()) + " " + appendLeadingZeroes(current_datetime.getHours()) + ":" + appendLeadingZeroes(current_datetime.getMinutes()) + ":" + appendLeadingZeroes(current_datetime.getSeconds());
  return formatted_date;
}

function copyData(){
  var spSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spSheet.getSheetByName("2020 Bakcup");
  var objData = sheet.getRange(spSheet.getCurrentCell().getRow(), 1, 1, sheet.getMaxColumns()).activate().getValues();
  
  var folioQA = objData[0][1]
  var fecharegistro = objData[0][2]
  var email = objData[0][3]
  var folioQaRelacionado = objData[0][5]
  var uuaa = objData[0][6]
  var aplicacion = objData[0][7]
  var buildingblock = objData[0][8]
  var liderproyecto = objData[0][9]
  var desarrollador = objData[0][10]
  var modelosolucion = objData[0][11]
  var objdesarrollo = objData[0][12]
  var componenteApx = objData[0][14]
  var servicioGlobal = objData[0][15]
  
  /*16-23*/
  var componentes ="";
  if(objData[0][16]!==""){ componentes += "APX," }
  if(objData[0][17]!==""){ componentes += "ASO," }
  if(objData[0][18]!==""){ componentes += "BACKEND," }
  if(objData[0][19]!==""){ componentes += "BBDD," }
  if(objData[0][20]!==""){ componentes += "FRONTEND," }
  if(objData[0][21]!==""){ componentes += "HOST," }
  if(objData[0][22]!==""){ componentes += "WEBMETHODS," }
  if(objData[0][23]!==""){ componentes += "OTRO" }  
  
  var herramientaDespliegue = objData[0][24]
  var liderintegador = objData[0][25]  
  var folioCRQ = objData[0][27]
  var tipoCambio = objData[0][28]
  var instalacionInicio = objData[0][29]
  var instalacionFin = objData[0][30]  
  var infoadicional = objData[0][31]  
  var entornoActual = objData[0][33]
  var estatusCheck = objData[0][34]
  var urlProcesoQA = objData[0][35]  
  var fchInicioRev = objData[0][36]
  var fchFinRev = objData[0][37]
  var statusQA = objData[0][38]  
  var asignadoQA = objData[0][39]  
  var observacionQA = objData[0][40]
  var estatusFinalCambio = objData[0][41]
  var observacionFinal = objData[0][42]
  var evidenciaInstal = objData[0][43]
  
  var sdatool = ""
  var issuekey = ""
  var frentedev = ""
  var ultimoambiente ="PRODUCCIÓN"
  
  
  
  var sheetS = spSheet.getSheetByName("2020");
  
  sheetS.appendRow([folioQA,folioQaRelacionado,asignadoQA,fecharegistro,email,uuaa,aplicacion,servicioGlobal,
                   buildingblock,sdatool,issuekey,liderproyecto,frentedev,desarrollador,
                   modelosolucion,objdesarrollo,infoadicional,componenteApx,liderintegador,
                   folioCRQ,tipoCambio,' ',instalacionInicio,instalacionFin,
                   herramientaDespliegue,ultimoambiente,componentes,
                   entornoActual,estatusCheck,urlProcesoQA,fchInicioRev,fchFinRev,statusQA,
                  '','',estatusFinalCambio,'','','',observacionQA + ', '+observacionFinal]);  
}
