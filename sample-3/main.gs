/**
* Global Setting for API Collaborator
*/
var SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
var SHEET_MAIN = SPREADSHEET.getSheetByName('data');
var SHEET_SETTINGS = SPREADSHEET.getSheetByName('settings');
var RNG_DATA_START = SHEET_MAIN.getRange('A2');
var CURRENT_ROW = 0;

var rngJWT = SpreadsheetApp.getActive().getRangeByName('set_JWT');
var keyJWT = rngJWT.getValue();
var urlHome = SpreadsheetApp.getActive().getRangeByName('set_url').getValue();
var date_JWT = SpreadsheetApp.getActive().getRangeByName('set_date_JWT').getValue();

/**
* SideBar
*/
/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
 */

var SIDEBAR_TITLE = 'LMS Collaborator. MBO Helper';

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Show MBO Helper', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle(SIDEBAR_TITLE)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
* Dublicate row
*/
function dublicateActiveRow() {
  var spreadsheet = SpreadsheetApp.getActive();
  Logger.log(spreadsheet.getName());
  if (spreadsheet.getActiveSheet().getName() === 'data') {
    var originalRow = spreadsheet.getActiveRange().getLastRow();
    spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveRange().getLastRow(), 1);
    //spreadsheet.getActiveRange().offset(spreadsheet.getActiveRange().getNumRows(), 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
    spreadsheet.getRange('A'+ (originalRow + 1)).activate();
    spreadsheet.getRange(originalRow +':' + originalRow).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    spreadsheet.getRange('C'+ (originalRow + 1)).activate();
    spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('E' + (originalRow + 1)).activate();
    setCheck(spreadsheet.getActiveRange().getLastRow());
  }
};

function setCheck(i_row){
  SHEET_MAIN.getRange(i_row, 1).setValue(true);
}

/****************************************************************
*
*/
function getAccess(){
  var set_url = SpreadsheetApp.getActive().getRangeByName('set_url').getValue();
  var set_login = SpreadsheetApp.getActive().getRangeByName('set_login').getValue();
  var set_pass = SpreadsheetApp.getActive().getRangeByName('set_pass').getValue();
  rngJWT.setValue(_getAccessToken(set_url, set_login, set_pass));
}

function checkJWT(){
  if(keyJWT < 20 || date_JWT < SHEET_SETTINGS.getRange('B7').getValue() - 0.1){
    getAccess();
    SpreadsheetApp.getActive().getRangeByName('set_date_JWT').setValue(timeConverter(Date.now()))
  }
  keyJWT = rngJWT.getValue();
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

function timeConverter(UNIX_timestamp){
  var a = new Date(UNIX_timestamp);
  var date = dateToStr(a)
  var hour = a.getHours();
  var min = a.getMinutes();
  var sec = a.getSeconds();
  var time = date + ' ' + hour + ':' + min + ':' + sec ;
  return time;
}

/**
* https://qa.davintoo.com/api/rest.php/mbo/structure
*
*
*/
function getMboTree(){
  checkJWT();
  
  var maxRows = SHEET_MAIN.getMaxRows();
  
  //Clear data
  SHEET_MAIN.deleteRows(4, maxRows-4);
  SHEET_MAIN.insertRowsAfter(SHEET_MAIN.getMaxRows(), 1000);
  
  var requestPlan = '/api/rest.php/mbo/structure';
  var resPlan = _getResByCollaboratorAPI(keyJWT, urlHome + requestPlan);
  var requestStructure = '/api/rest.php/mbo/structure/' + resPlan.data[0].id + '?status=all&content=&assigned_user_id=0';
  var res = _getResByCollaboratorAPI(keyJWT, urlHome + requestStructure);
  if (JSON.stringify(res).substring(1, 6) == 'Error') {
      Logger.log(JSON.stringify(res));
      return 0;
  }; 
  writeMboElement(res.id, res.children); 
}

function writeMboElement(id_parent, res_data){
  res_data.forEach( function(item, i, res_data) {
    CURRENT_ROW = CURRENT_ROW + 1; 
    SHEET_MAIN.getRange(1, 1).copyTo(RNG_DATA_START.offset(CURRENT_ROW, 0), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    RNG_DATA_START.offset(CURRENT_ROW, 1).setValue(id_parent);
    RNG_DATA_START.offset(CURRENT_ROW, 2).setValue(item.id);
    //depth	title	assigned_user_id	percent	status_id	self_mark	chief_mark	date_end	dete_end_fact
    RNG_DATA_START.offset(CURRENT_ROW, 3).setValue(item.depth);
    RNG_DATA_START.offset(CURRENT_ROW, 4).setValue(item.title);
    RNG_DATA_START.offset(CURRENT_ROW, 5).setValue(item.assigned_user_id);
    if(item.assigned_user_id > 0){
      RNG_DATA_START.offset(CURRENT_ROW, 6).setValue(item.assigned_user.fullname);
    }
    RNG_DATA_START.offset(CURRENT_ROW, 7).setValue(item.description);
    RNG_DATA_START.offset(CURRENT_ROW, 8).setValue(item.percent);
    RNG_DATA_START.offset(CURRENT_ROW, 9).setValue(item.status_id);
    RNG_DATA_START.offset(CURRENT_ROW, 10).setValue(item.status);
    RNG_DATA_START.offset(CURRENT_ROW, 11).setValue(item.self_mark);
    RNG_DATA_START.offset(CURRENT_ROW, 12).setValue(item.chief_mark);
    RNG_DATA_START.offset(CURRENT_ROW, 13).setValue(item.date_begin);
    RNG_DATA_START.offset(CURRENT_ROW, 14).setValue(item.dete_end_fact);
    RNG_DATA_START.offset(CURRENT_ROW, 15).setValue(item.is_published);
    
    var listKpi = getKpiForTaskId(item.id);
    listKpi.forEach( function(kpi, j, listKpi) {
      RNG_DATA_START.offset(CURRENT_ROW, 16 + j*3).setValue(kpi.id);
      RNG_DATA_START.offset(CURRENT_ROW, 17 + j*3).setValue(kpi.value);
      RNG_DATA_START.offset(CURRENT_ROW, 18 + j*3).setValue(kpi.name);
    });
    
    if(item.children.length > 0){
      writeMboElement(item.id, item.children);
    }
  });
}

function onGetLists() {
  checkJWT();
  SHEET_SETTINGS.getRange('C11:G').clearContent();
  getUsers();
  getStatuses();
}


function getUsers(){
  var requestActiveUsers = '/api/rest.php/auth/users?page=1&count=1000&filter[is_active]=1&filter[status]=1&sorting[fullname]=asc';
  var resActiveUsers = _getResByCollaboratorAPI(keyJWT, urlHome + requestActiveUsers).data;
  resActiveUsers.forEach( function(item, i, resActiveUsers) {
    SHEET_SETTINGS.getRange('C11').offset(i, 0).setValue(item.id);
    SHEET_SETTINGS.getRange('C11').offset(i, 1).setValue(item.fullname);
  });
}

function getStatuses(){
  var requestStatuses = '/api/rest.php/work-settings?action=get-statuses';
  var resStatuses = _getResByCollaboratorAPI(keyJWT, urlHome + requestStatuses).data;
  resStatuses.forEach( function(item, i, resStatuses) {
    SHEET_SETTINGS.getRange('F11').offset(i, 0).setValue(item.id);
    SHEET_SETTINGS.getRange('F11').offset(i, 1).setValue(item.title);
  });
}

function getKpiForTaskId(id){
  //var id = SHEET_MAIN.getRange(i_row, 3).getValue();
  var requestKpi = '/api/rest.php/mbo/kpi?structure_id=' + id;
  var resKpi = _getResByCollaboratorAPI(keyJWT, urlHome + requestKpi);
  Logger.log(resKpi.data);
  return resKpi.data;
}

function saveKpiForRow(i_row, i_kpi){
  //checkJWT();
  var name_kpi = SHEET_MAIN.getRange(i_row, 19 + i_kpi * 3).getValue();
  if (name_kpi){
    var id_kpi = SHEET_MAIN.getRange(i_row, 17 + i_kpi * 3).getValue();
    var value_kpi = SHEET_MAIN.getRange(i_row, 18 + i_kpi * 3).getValue();
    var structure_id = SHEET_MAIN.getRange(i_row, 3).getValue();
    
    if (id_kpi > 0){
      //save record
      var payload_obj = {
        structure_id: structure_id,
        '$edit':true,
        name: name_kpi,
        id:id_kpi,
        value: value_kpi
      };
      var payload = JSON.stringify(payload_obj);
      _saveKpi(keyJWT, urlHome, payload, id_kpi);
    } else {
      //create new record
      var payload_obj = {
        structure_id: structure_id,
        '$edit':true,
        name: name_kpi,
        value: value_kpi
      };
      var payload = JSON.stringify(payload_obj);
      var new_kpi = _insertKpiElement(keyJWT, urlHome, payload);
      Logger.log(new_kpi);
      SHEET_MAIN.getRange(i_row, 17 + i_kpi * 3).setValue(new_kpi.id);
    }
  }
}



/**
* UPDATE/CREATE ITEM
* Request URL: /api/rest.php/mbo/structure/__id__
* Request Method: PUT
*/
function saveMboElement(i_row){
  var id_parent = SHEET_MAIN.getRange(i_row, 2).getValue();
  var payload_obj = {
    parent_id: id_parent,
    title: SHEET_MAIN.getRange(i_row, 5).getValue(),
    assigned_user_id: SHEET_MAIN.getRange(i_row, 6).getValue(),
    description: SHEET_MAIN.getRange(i_row, 8).getValue(),
    percent: SHEET_MAIN.getRange(i_row, 9).getValue(),
    status_id: SHEET_MAIN.getRange(i_row, 10).getValue(),
    self_mark: SHEET_MAIN.getRange(i_row, 12).getValue(),
    chief_mark: SHEET_MAIN.getRange(i_row, 13).getValue(),
    date_begin: SHEET_MAIN.getRange(i_row, 14).getValue(),
    dete_end_fact: SHEET_MAIN.getRange(i_row, 15).getValue(),
    is_published: SHEET_MAIN.getRange(i_row, 16).getValue()
    };
  
  var payload = JSON.stringify(payload_obj);
  var this_rng = RNG_DATA_START.offset(CURRENT_ROW, i_row);
  var id = SHEET_MAIN.getRange(i_row, 3).getValue();
  
  checkJWT();
  
  if (id>0){
    // save record
    _saveMboElement(keyJWT, urlHome, payload, id);
  } else {
    // create new record
    var new_res = _insertMboElement(keyJWT, urlHome, payload, id_parent);
    SHEET_MAIN.getRange(i_row, 3).setValue(new_res.id);
  }
  
  // save or create kpi (to 5 items)
  for (var i_kpi = 0; i_kpi < 5; i_kpi++) {
    Logger.log(i_kpi);
    saveKpiForRow(i_row, i_kpi);
  }
  
  
}

function saveCurrentElement(){
  var i_row = SHEET_MAIN.getActiveRange().getRowIndex();
  Logger.log(i_row); 
  saveMboElement(i_row);
}

function saveCheckedElements(){
  var checkedRows = _getCheckedRows();
  if (checkedRows.length > 0){
    for (i in checkedRows){
      if (SHEET_MAIN.getRange(checkedRows[i], 2).getValue() > 0 ) {
        Logger.log(checkedRows[i]);
        saveMboElement(checkedRows[i]);
        SHEET_MAIN.getRange(checkedRows[i], 1).setValue(false);
      }
    }
  }
}


function onEdit(e){  
  
  if(e.range.getColumn() > 4 && e.range.getRow() > 2 && e.range.getSheet().getName() === 'data') {
    Logger.log('--0--');
    setCheck(e.range.getRow());
  }
  
  if(e.range.getColumn() === 7 && e.range.getRow() > 2 && e.range.getSheet().getName() === 'data') {
      var users_list = SHEET_SETTINGS.getRange('C11:D').getValues();
    var id_user = 0;
    for (i in users_list){
      id_user = users_list[i][0];
      if (users_list[i][1] == e.range.getValue()) break;
    }
    if(id_user != 0){
      SHEET_MAIN.getRange(e.range.getRow(), e.range.getColumn()-1).setValue(id_user);
    }
  };
  
  if(e.range.getColumn() === 11 && e.range.getRow() > 2 && e.range.getSheet().getName() === 'data') {
    var status_list = SHEET_SETTINGS.getRange('F11:G').getValues();
    Logger.log(status_list);
    var id_status = 0;
    for (i in status_list){
      id_status = status_list[i][0];
      if (status_list[i][1] == e.range.getValue()) break;
    }
    if(id_status != 0){
      SHEET_MAIN.getRange(e.range.getRow(), e.range.getColumn()-1).setValue(id_status);
    }
  }
}

/**
* DELETE ITEM
* Request URL: /api/rest.php/mbo/structure/--id--
* Request Method: DELETE
* response: true/false
*/
function deleteMboElement(i_row){
  var id = SHEET_MAIN.getRange(i_row, 3).getValue();
  if (id > 0){
    checkJWT();
    var res = _deleteMboElement(keyJWT, urlHome, id);
    Logger.log(res);
  }
  SHEET_MAIN.deleteRow(i_row);
}

function deleteCurrentElement(){
  var i_row = SHEET_MAIN.getActiveRange().getRowIndex();
  Logger.log(i_row); 
  deleteMboElement(i_row);
}

//function deleteCheckedElements(){
//  var checkedRows = _getCheckedRows();
//  if (checkedRows.length > 0){
//    Logger.log(checkedRows); 
//    for (i in checkedRows){
//      if (SHEET_MAIN.getRange(checkedRows[i], 3).getValue() > 0 ) {
//        deleteMboElement(checkedRows[i]);
//      }
//    }
//  }
//}

/**
* Group operation
*/
function _getCheckedRows(){
  var checkedRows=[];
  var i = 3;
  while (SHEET_MAIN.getRange(i, 2).getValue() > 0 ) {
    if (SHEET_MAIN.getRange(i, 1).getValue() === true) {
      checkedRows.push(i);
    } 
    i++;
  }
  return checkedRows;
}


/**
* Work with KPI
*/

function getKpiValue() {
  var i_row = SHEET_MAIN.getActiveRange().getRowIndex();
  var text = '';
  var kpi = [
    SHEET_MAIN.getRange(i_row, 19).getValue(),
    SHEET_MAIN.getRange(i_row, 22).getValue(),
    SHEET_MAIN.getRange(i_row, 25).getValue(),
    SHEET_MAIN.getRange(i_row, 28).getValue(),
    SHEET_MAIN.getRange(i_row, 31).getValue()
  ];
  for (i in kpi){
    if (kpi[i]) {
      text = text + kpi[i] + "\n";
    }
  }
  return text.slice(0, -1);
}


function setKpiValue(value) {
  Logger.log(value);
  if (value.length > 0) {
    var i_row = SHEET_MAIN.getActiveRange().getRowIndex();
    var kpi = value.split('\n');
    kpi.forEach( function(item, i, kpi) {
      SHEET_MAIN.getRange(i_row, 19 + i*3).setValue(item);
    });
    setCheck(i_row);
  }
}

function showAlert() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     'Please confirm',
     'Are you sure you want to continue?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert('Confirmation received.');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Permission denied.');
  }
}