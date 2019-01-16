var SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
var SHEET_MAIN = SPREADSHEET.getSheetByName("main");
var SHEET_USERS = SPREADSHEET.getSheetByName("users");
var SHEET_SETTINGS = SPREADSHEET.getSheetByName("settings");
var rngJWT = SpreadsheetApp.getActive().getRangeByName('set_JWT');
var keyJWT = rngJWT.getValue();
var urlHome = SpreadsheetApp.getActive().getRangeByName('set_url').getValue();

var rngStartReport = SpreadsheetApp.getActive().getRangeByName('start_report');

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

function getUsersList() {
  checkJWT();
  
  var records_limit = 100;
  var date_period_begin = SHEET_MAIN.getRange('D3').getValue();
  var date_period_end = SHEET_MAIN.getRange('E3').getValue();
  var requestURL = '/api/rest.php/auth/users?page=1&count=' + records_limit + 
    '&filter[last_activity]=' + dateToStr(date_period_begin) + 
    '&filter[last_activity]=' + dateToStr(date_period_end) + 
    '&sorting[last_activity]=desc';
  Logger.log( urlHome + requestURL);
  var res = _getResByCollaboratorAPI(keyJWT, urlHome + requestURL);
  if (JSON.stringify(res).substring(1, 6) == 'Error') {
      Logger.log(JSON.stringify(res));
      return 0;
  }
  
  SHEET_USERS.getRange(1, 1, 1000, 2).clearContent();
  var rng = SHEET_USERS.getRange(1, 1);
  var data_users = res.data;
  data_users.forEach( function(item, i, data_users) {
    rng.offset(i, 1).setValue(item.id);
    rng.offset(i, 0).setValue(item.fullname);
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

function clearImages(){
  var ReportSheet = SpreadsheetApp.getActive();
  // Deletes all images (except image[0]) in a ReportSheet
  var images = ReportSheet.getImages();
  for (var i = 1; i < images.length; i++) {
    images[i].remove();
  }
}

function getImageFromURL(urlRequest){
  var ReportSheet = SpreadsheetApp.getActive();
  var sheet = SHEET_MAIN;
  var img = UrlFetchApp.fetch(urlRequest).getBlob();
  
  sheet.insertImage(img, 1, 5)
  var images = sheet.getImages();
  var img = images[images.length - 1];
  var imgH = img.getHeight();
  var imgW = img.getWidth();
  var maxW = 150;
  var newImgH = Math.round(imgH*(maxW/imgW));
  img.setWidth(maxW)
     .setHeight(newImgH)
     .setAltTextTitle('Изображение')
     .setAnchorCellYOffset(0)
     .setAnchorCellXOffset(5);
}


function getUserProfile(user_id){
  // /api/rest.php/auth/users/<user_id>
  var requestURL = '/api/rest.php/auth/users/' + user_id;
  var res = _getResByCollaboratorAPI(keyJWT, urlHome + requestURL);

  if (JSON.stringify(res).substring(1, 6) == 'Error') {
      Logger.log(JSON.stringify(res));
      return 0;
  } 
  // refill cells with user info
  SHEET_MAIN.getRange('A4').setValue(res.fullname);
  clearImages(); // clear photo
  if (res.photo_thumb){
    getImageFromURL(urlHome + res.photo_thumb); //get & insert photo_thumb
  }  
  SHEET_MAIN.getRange('B8').setValue(res.position);
  SHEET_MAIN.getRange('B9').setValue(res.department);
  SHEET_MAIN.getRange('B10').setValue(res.city); 
}

function getUserLearningHistoryForPeriod(user_id, date_period_begin, date_period_end, records_limit) {
  // get data of learning tasks for preset user who was assigned to this tasks in portal from <date_period_begin> to <date_period_end>
  // example api request:
  // api/rest.php/tasks?page=1&count=10&filter[date_assign]=2018-11-01&filter[date_assign]=2018-12-09&sorting[date_assign]=desc&user_id=3906&action=get-study-history-all-tasks
  records_limit = records_limit || 1000;
  
  var requestURL = '/api/rest.php/tasks?page=1&count=' + records_limit + 
    '&filter[date_assign]=' + dateToStr(date_period_begin) + '&filter[date_assign]=' + dateToStr(date_period_end) + 
    '&sorting[date_assign]=desc&user_id=' + user_id + '&action=get-study-history-all-tasks';
  var res = _getResByCollaboratorAPI(keyJWT, urlHome + requestURL);

  if (JSON.stringify(res).substring(1, 6) == 'Error') {
      Logger.log(JSON.stringify(res));
      return 0;
  } 
  
  // fill cells in sheet "users"
  SHEET_MAIN.getRange(rngStartReport.getRow(), 1, SHEET_MAIN.getLastRow(), 10).clearContent();
  var rng = rngStartReport;
  var data_tasks = res.data;
  var accessdates = [];
  data_tasks.forEach( function(item, i, data_tasks) {
    rng.offset(i, 0).setValue(item.title);
    rng.offset(i, 1).setValue(item.display_type);
    rng.offset(i, 2).setValue(item.status);
    rng.offset(i, 3).setValue(item.mark);
    rng.offset(i, 4).setValue(item.date_assign);
    //rng.offset(i, 5).setValue(item.date_start);
    rng.offset(i, 5).setValue(item.date_finish);
  });
}

function main(){
  var user_id = SpreadsheetApp.getActive().getRangeByName('user_id').getValue();
  var date_period_begin = SHEET_MAIN.getRange('D3').getValue();
  var date_period_end = SHEET_MAIN.getRange('E3').getValue();
  
  checkJWT(); //check and udate JWT-key if nessesary
  
  getUserProfile(user_id);
  getUserLearningHistoryForPeriod(user_id, date_period_begin, date_period_end);
}