var keyJWT;
var URL = 'https://<your-domen>';

function _getAccessToken(login, pass){
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
      keyJWT = '';
      Logger.log('Ошибка доступа: ' + responseCode + '\n Пользователь не найден. Проверьте правильность ввода логина и пароля.');
      break;
    default:
      keyJWT = '';
      Logger.log('Ошибка доступа: ' + responseCode + '\n Проверьте, все ли параметры запроса правильные.');
      break;
  }
}

function test(){
  _getAccessToken('<your-login>', '<your-pass>');
  //Печатаем токен в текст активного документа
  var body = DocumentApp.getActiveDocument().getBody();
  var text = body.editAsText();
  text.appendText('keyJWT:\n' + keyJWT);
}