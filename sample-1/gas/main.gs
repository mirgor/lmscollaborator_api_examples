var SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
var SHEET_MAIN = SPREADSHEET.getSheetByName("main");
var SHEET_SETTINGS = SPREADSHEET.getSheetByName("settings");
var rngJWT = SpreadsheetApp.getActive().getRangeByName('set_JWT');
var keyJWT = rngJWT.getValue();
var urlHome = SpreadsheetApp.getActive().getRangeByName('set_url').getValue();

function getAccess(){
  var set_url = SpreadsheetApp.getActive().getRangeByName('set_url').getValue();
  var set_login = SpreadsheetApp.getActive().getRangeByName('set_login').getValue();
  var set_pass = SpreadsheetApp.getActive().getRangeByName('set_pass').getValue();
  rngJWT.setValue(_getAccessToken(set_url, set_login, set_pass));
}

function checkJWT(){
  if(keyJWT < 20 ){
    getAccess();
  }
  keyJWT = rngJWT.getValue();
}

function main() {
  checkJWT();
  
  var records_limit = 100;
  var date_period_begin = SHEET_MAIN.getRange('D1').getValue();
  var date_period_end = SHEET_MAIN.getRange('D2').getValue();
  
  var requestURL = '/api/rest.php/auth/users?page=1&count=' + records_limit + 
    '&filter[last_activity]=' + dateToStr(date_period_begin) + 
    '&filter[last_activity]=' + dateToStr(date_period_end) + 
    '&sorting[last_activity]=desc';
  
  var res = _getResByCollaboratorAPI(keyJWT, urlHome + requestURL);
  if (JSON.stringify(res).substring(1, 6) == 'Error') {
      Logger.log(JSON.stringify(res));
      return 0;
  }
  
  SHEET_MAIN.getRange(1, 1, 1000, 2).clearContent();
  var rng = SHEET_MAIN.getRange(1, 1);
  var data_users = res.data;
  data_users.forEach( function(item, i, data_users) {
    rng.offset(i, 0).setValue(item.id);
    rng.offset(i, 1).setValue(item.fullname);
  });
  
}

function dateToStr(date_val){
  // return date value as <yyyy-mm-dd> format for URL request
  var dd = date_val.getDate(); 
  var mm = date_val.getMonth()+1; //January is 0!
  var yyyy = date_val.getFullYear();
  
  if(dd < 10) {
    dd = '0' + dd
  }
  if(mm < 10) {
    mm = '0' + mm
  } 
  return yyyy + '-' + mm + '-' + dd;
}