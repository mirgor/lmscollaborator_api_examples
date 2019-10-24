function _getAccessToken(URL, login, pass){
  var keyJWT;
  var res;
  var options = {
    'method' : 'post',
    'contentType': 'application/json;charset=UTF-8',
    'payload' : '{"email":"' + login + '","password":"' + pass + '"}',
    'muteHttpExceptions' : true
  };
  var response = UrlFetchApp.fetch(URL + '/api/rest.php/auth/session', options);
  var responseCode = response.getResponseCode();
  switch(responseCode) {
    case 200:  
      res = JSON.parse(response.getContentText());
      keyJWT = res.access_token;
      Logger.log('User: ' + res.fullname + '\n' + keyJWT);
      break;
    case 400:  
      keyJWT = 'Error 400';
      Logger.log('Ошибка доступа: ' + responseCode + '\nПользователь не найден. Проверьте правильность ввода логина и пароля.');
      break;
    default:
      keyJWT = 'Error ' + responseCode;
      Logger.log('Ошибка: ' + response.getContentText());
      break;
  }
  return keyJWT;
}

function _getResByCollaboratorAPI(apiKey, urlRequest) {
  var response = UrlFetchApp.fetch(urlRequest, {
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type' : 'application/json;charset=UTF-8'
    },
    'method' : 'get',
    'muteHttpExceptions' : true    
  });
  var responseCode = response.getResponseCode();
  var res;
  switch(responseCode) {
    case 200:  
      res = JSON.parse(response.getContentText());
      //Logger.log(res);
      break;
    case 400:  
      res = 'Error 400';
      Logger.log('Ошибка доступа: ' + responseCode + '\nПроверьте правильность ввода логина и пароля.');
      break;
    default:
      res = 'Error ' + responseCode;
      Logger.log('Ошибка: ' + response.getContentText() + 
                '\nПроверьте, все ли параметры запроса правильные.\n' +
                '\n' + urlRequest);
      break;
  }
  return res;  
}

// /api/rest.php/mbo/kpi/73
//PUT
//{"structure_id":452,"$edit":true,"name":"New KPI r","id":73,"value":true}
function _saveKpi(keyJWT, URL, payload, id_kpi_element) {
  var res;
  var options = {
    headers: {
      'Authorization': 'Bearer ' + keyJWT,
      'Content-Type' : 'application/json;charset=UTF-8'
    },
    'method' : 'PUT',
    'contentType': 'application/json;charset=UTF-8',
    'payload' : payload,
    'muteHttpExceptions' : true
  };
  var response = UrlFetchApp.fetch(URL + '/api/rest.php/mbo/kpi/' + id_kpi_element, options);
  var responseCode = response.getResponseCode();
  switch(responseCode) {
    case 200:  
      res = JSON.parse(response.getContentText());
      break;
    case 400:  
      Logger.log('Ошибка доступа: ' + responseCode + '\nПользователь не найден. Проверьте правильность ввода логина и пароля.');
      break;
    default:
      Logger.log('Ошибка: ' + response.getContentText());
      break;
  }
  return res;
}

// /api/rest.php/mbo/kpi
//POST
//{"structure_id":452,"$edit":true,"name":"New KPI"}
function _insertKpiElement(keyJWT, URL, payload) {
  var res;
  var options = {
    headers: {
      'Authorization': 'Bearer ' + keyJWT,
      'Content-Type' : 'application/json;charset=UTF-8'
    },
    'method' : 'POST',
    'contentType': 'application/json;charset=UTF-8',
    'payload' : payload,
    'muteHttpExceptions' : true
  };
  var response = UrlFetchApp.fetch(URL + '/api/rest.php/mbo/kpi', options);
  var responseCode = response.getResponseCode();
  switch(responseCode) {
    case 200:  
      res = JSON.parse(response.getContentText());
      break;
    case 400:  
      Logger.log('Ошибка доступа: ' + responseCode + '\nПользователь не найден. Проверьте правильность ввода логина и пароля.');
      break;
    default:
      Logger.log('Ошибка: ' + response.getContentText());
      break;
  }
  return res;
}


function _saveMboElement(keyJWT, URL, payload, id_mbo_element) {
  var res;
  var options = {
    headers: {
      'Authorization': 'Bearer ' + keyJWT,
      'Content-Type' : 'application/json;charset=UTF-8'
    },
    'method' : 'PUT',
    'contentType': 'application/json;charset=UTF-8',
    'payload' : payload,
    'muteHttpExceptions' : true
  };
  var response = UrlFetchApp.fetch(URL + '/api/rest.php/mbo/structure/' + id_mbo_element, options);
  var responseCode = response.getResponseCode();
  switch(responseCode) {
    case 200:  
      res = JSON.parse(response.getContentText());
      break;
    case 400:  
      Logger.log('Ошибка доступа: ' + responseCode + '\nПользователь не найден. Проверьте правильность ввода логина и пароля.');
      break;
    default:
      Logger.log('Ошибка: ' + response.getContentText());
      break;
  }
  return res;
}


function _insertMboElement(keyJWT, URL, payload, id_parent) {
  var res;
  var options = {
    headers: {
      'Authorization': 'Bearer ' + keyJWT,
      'Content-Type' : 'application/json;charset=UTF-8'
    },
    'method' : 'POST',
    'contentType': 'application/json;charset=UTF-8',
    'payload' : payload,
    'muteHttpExceptions' : true
  };
  var response = UrlFetchApp.fetch(URL + '/api/rest.php/mbo/structure', options);
  var responseCode = response.getResponseCode();
  switch(responseCode) {
    case 200:  
      res = JSON.parse(response.getContentText());
      break;
    case 400:  
      Logger.log('Ошибка доступа: ' + responseCode + '\nПользователь не найден. Проверьте правильность ввода логина и пароля.');
      break;
    default:
      Logger.log('Ошибка: ' + response.getContentText());
      break;
  }
  return res;
}

function _deleteMboElement(keyJWT, URL, id) {
  var res;
  var options = {
    headers: {
      'Authorization': 'Bearer ' + keyJWT,
      'Content-Type' : 'application/json;charset=UTF-8'
    },
    'method' : 'DELETE',
    'contentType': 'application/json;charset=UTF-8',
    'payload' : "",
    'muteHttpExceptions' : true
  };
  var response = UrlFetchApp.fetch(URL + '/api/rest.php/mbo/structure/' + id, options);
  var responseCode = response.getResponseCode();
  switch(responseCode) {
    case 200:  
      res = JSON.parse(response.getContentText());
      break;
    case 400:  
      Logger.log('Ошибка доступа: ' + responseCode + '\nПользователь не найден. Проверьте правильность ввода логина и пароля.');
      break;
    default:
      Logger.log('Ошибка: ' + response.getContentText());
      break;
  }
  return res;
}
