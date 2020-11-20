// Constant definitions
var SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEET = SS.getSheets()[0]
const START_ROW = 2
const START_COLUMN = 4
const LAST_COLUMN = 9
const LAST_ROW = SHEET.getLastRow()
const N_ROWS = LAST_ROW - START_ROW + 1
const N_COLUMNS = LAST_COLUMN - START_COLUMN + 1
const STATUS_COLUMN = "A"
const SUCCESS_COLOR = "#4BB543"
const FAILURE_COLOR = "#ff3333"

var sensitive_range = SHEET.getRange(START_ROW, START_COLUMN, N_ROWS, N_COLUMNS);
var protection = sensitive_range.protect().setDescription('Sample protected range');
if (protection.canDomainEdit()) {
  protection.setDomainEdit(false);
}



// custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Enrichissement INSEE')
      .addItem('Tout Enrichir/Raffraichir','enrichirTout')
      .addItem('Enrichir les manquants','enrichirManquants')
      .addItem('Enrichir les status non valides','enrichirStatusNonValides')
      .addToUi();
}

// Utility functions

function fillArray(value, len) {
  if (len == 0) return [];
  var a = [value];
  while (a.length * 2 <= len) a = a.concat(a);
  if (a.length < len) a = a.concat(a.slice(0, len - a.length));
  return a;
}

// fill error message:
function fillErrorMsg(error_msg, range, n_columns) {
  var error_msgs = [fillArray(error_msg, n_columns)]
  range.setValues(error_msgs);
}

function setStatus(range, status_bool) {
  var row_index = range.getRowIndex()
  SHEET.getRange(STATUS_COLUMN + row_index.toString()).setValue(status_bool)
  var color = parseInt(status_bool) ? SUCCESS_COLOR: FAILURE_COLOR
  range.setBackground(color)
  
}

function getStatus(row_index) {
  return SHEET.getRange(STATUS_COLUMN + row_index.toString()).getValue()
}

// End: utility functions


function getCompanyInfos(raison, codepostale)  {
  
//  Configurations
  const token ='' // set a free token from insee api;
  var champs = "&champs=siren,siret,complementAdresseEtablissement,numeroVoieEtablissement,indiceRepetitionEtablissement,typeVoieEtablissement,libelleVoieEtablissement,codePostalEtablissement,libelleCommuneEtablissement,denominationUniteLegale,dateCreationEtablissement,trancheEffectifsEtablissement,anneeEffectifsEtablissement,etatAdministratifUniteLegale,activitePrincipaleUniteLegale,nomenclatureActivitePrincipaleUniteLegale"
  var baseURL = 'https://api.insee.fr/entreprises/sirene/V3/siret'
  var query = '?q=denominationUniteLegale%3A%22' + raison + '%22%20AND%20codePostalEtablissement%3A' + codepostale
  var webAppUrl =  baseURL + query + '&nombre=1' + champs;
  
  const options = {
    'headers' :  { 'Authorization' : 'Bearer ' + token}
    , muteHttpExceptions : true
 };
  
//  Hitting the API
  const response = UrlFetchApp.fetch(webAppUrl, options);
  return response
  
}

function handleResponse(response, range, n_columns){
  
  const result = JSON.parse(response.getContentText())
  const response_code = response.getResponseCode()
  var status_bool = 0;
  
//  SpreadsheetApp.getUi().alert("response_code: "+response_code)
  
  //  Handling response
  switch(response_code) {
    // Success
    case 200:
      var company_infos = result['etablissements'][0]
      //    Retrieving pieces of data into variables
      var siret = company_infos["siret"]
      var siren = company_infos["siren"]
      var trancheEffectifsEtablissement = company_infos["trancheEffectifsEtablissement"]
      var anneeEffectifsEtablissement = company_infos["anneeEffectifsEtablissement"]
      var denominationUniteLegale = company_infos["uniteLegale"]["denominationUniteLegale"]
      //      DEBUT ADRESSE
      const sep = ', '
      const space = ' '
      const adresseInfos = company_infos["adresseEtablissement"]
      var complementAdresseEtablissement = adresseInfos["complementAdresseEtablissement"] ? adresseInfos["complementAdresseEtablissement"] : ""
      var numeroVoieEtablissement = adresseInfos["numeroVoieEtablissement"] ? adresseInfos["numeroVoieEtablissement"] : ""
      var indiceRepetitionEtablissement = adresseInfos["indiceRepetitionEtablissement"] ? adresseInfos["indiceRepetitionEtablissement"] : ""
      var typeVoieEtablissement = adresseInfos["typeVoieEtablissement"] ? adresseInfos["typeVoieEtablissement"] : ""
      var libelleVoieEtablissement = adresseInfos["libelleVoieEtablissement"] ? adresseInfos["libelleVoieEtablissement"] : ""
      var codePostalEtablissement = adresseInfos["codePostalEtablissement"] ? adresseInfos["codePostalEtablissement"] : ""
      var libelleCommuneEtablissement = adresseInfos["libelleCommuneEtablissement"] ? adresseInfos["libelleCommuneEtablissement"] : ""
      var adress = complementAdresseEtablissement + sep + numeroVoieEtablissement + space + indiceRepetitionEtablissement + space + typeVoieEtablissement + space + libelleVoieEtablissement + sep + codePostalEtablissement + sep + libelleCommuneEtablissement
      //      FIN ADRESSE
      //    Filling the google sheet cells with correct values
      var infos_enrichis = [[siren, siret, trancheEffectifsEtablissement, anneeEffectifsEtablissement, denominationUniteLegale, adress]]
      range.setValues(infos_enrichis);
      status_bool = 1
      break;
    //  Too many requests
    case 429:
      fillErrorMsg("Vous avez dépassé votre quota de requete.", range, n_columns)
      break;
    // other errors
    case 500:
    case 503:
      fillErrorMsg("Erreur du coté du serveur.", range, n_columns)
      setStatus(range, )
      break;
    case 403:
      fillErrorMsg("Clé d'accès à INSEE invalide.", range, n_columns)
      break;
    case 404:
      fillErrorMsg("Aucune entreprise trouvée sur INSEE", range, n_columns)
      break;
    case 400:
    case 401:
    case 406:
    case 414:
      fillErrorMsg("Erreur avec la requête.", range, n_columns)
      break;
    default:
      fillErrorMsg("Erreur inconnue.", range, n_columns)
      break;
  }
  setStatus(range, status_bool)
}

//function checkRequestInfos() {}

function shouldRequestForMissingValue(range) {
  if (range.isBlank()) {
    return true
  }
//  var start_row = range.getRowIndex()
  for (var c = 1; c <= N_COLUMNS; c++) {
//    SpreadsheetApp.getUi().alert("column : "+c)
    if (range.getCell(1, c).getValue() == "") {
      return true
    }
  }
  return false
}


function enrichirManquants() {
    // looping through the raws
  for (var i = START_ROW; i <= last_row; i++){
    var range = SHEET.getRange(i, START_COLUMN, 1, N_COLUMNS);
    
//  Request infos
    var company_name = SHEET.getRange(i, 2).getValue();
    var zip_code = SHEET.getRange(i, 3).getValue();
    
//  In case it is a blank row or one element is missing
    if (!company_name || !zip_code) {
//      range.clearContent()
      fillErrorMsg("Remplir tous les champs obligatoires", range, N_COLUMNS)
      continue;
    }
    if (!shouldRequestForMissingValue(range)) {
      continue
    }
    
    //    Making an api call to the INSEE API 
    var response = getCompanyInfos(company_name, zip_code)
    handleResponse(response, range, N_COLUMNS)
//    SpreadsheetApp.getUi().alert("updated row : "+i)
  }
  
  // Resizing columns
  SHEET.autoResizeColumns(START_COLUMN,LAST_COLUMN)
}


function enrichirStatusNonValides(){
  for (var i = START_ROW; i <= LAST_ROW; i++){
    var range = SHEET.getRange(i, START_COLUMN, 1, N_COLUMNS);
    
    if ( getStatus(i) == 1 ) {
        continue
      }
    
//  Request infos
    var company_name = SHEET.getRange(i, 2).getValue();
    var zip_code = SHEET.getRange(i, 3).getValue();
    
//  In case it is a blank row or one element is missing
    if (!company_name || !zip_code) {
      fillErrorMsg("Remplir tous les champs obligatoires", range, N_COLUMNS)
//      range.clearContent()
      continue;
    }
    
//    Making an api call to the INSEE API 
    var response = getCompanyInfos(company_name, zip_code)
    handleResponse(response, range, N_COLUMNS)
//    SpreadsheetApp.getUi().alert("updated row : "+i)
    
  }
  
  // Resizing columns
  SHEET.autoResizeColumns(START_COLUMN,LAST_COLUMN)
}


function enrichirTout(){
  
  
//  Getting the frame of our data in google sheet
  // looping through the rows of the sheet
  for (var i = START_ROW; i <= LAST_ROW; i++){
    var range = SHEET.getRange(i, START_COLUMN, 1, N_COLUMNS);
    
    
//  Request infos
    var company_name = SHEET.getRange(i, 2).getValue();
    var zip_code = SHEET.getRange(i, 3).getValue();
    
//  In case it is a blank row or one element is missing
    if (!company_name || !zip_code) {
      fillErrorMsg("Remplir tous les champs obligatoires", range, N_COLUMNS)
//      range.clearContent()
      continue;
    }
    
//    Making an api call to the INSEE API 
    var response = getCompanyInfos(company_name, zip_code)
    handleResponse(response, range, N_COLUMNS)
//    SpreadsheetApp.getUi().alert("updated row : "+i)
    
  }
  
  // Resizing columns
  SHEET.autoResizeColumns(START_COLUMN,LAST_COLUMN)
}

//function isrequestInfosValid(row_index) {
//
//}