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
      Logger.log('Ошибка доступа: ' + responseCode + '\nПроверьте, все ли параметры запроса правильные.');
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
      Logger.log(res);
      break;
    case 400:  
      res = 'Error 400';
      Logger.log('Ошибка доступа: ' + responseCode + '\nПроверьте правильность ввода логина и пароля.');
      break;
    default:
      res = 'Error ' + responseCode;
      Logger.log('Ошибка доступа: ' + responseCode + 
                '\nПроверьте, все ли параметры запроса правильные.\n' +
                '\n' + urlRequest);
      break;
  }
  return res;  
}



