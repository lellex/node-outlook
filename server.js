var express = require('express');
var bodyParser = require('body-parser');
var app = express();
var url = require('url');
var authHelper = require('./authHelper');
var fs = require('fs');
var outlook = require('node-outlook');

function getValueFromCookie(valueName, cookie) {
  if (cookie.indexOf(valueName) !== -1) {
    var start = cookie.indexOf(valueName) + valueName.length + 1;
    var end = cookie.indexOf(';', start);
    end = end === -1 ? cookie.length : end;
    return cookie.substring(start, end);
  }
}

function getUserEmail(token, callback) {
  // Set the API endpoint to use the v2.0 endpoint
  outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');

  // Set up oData parameters
  var queryParams = {
    '$select': 'DisplayName, EmailAddress',
  };

  outlook.base.getUser({token: token, odataParams: queryParams}, function(error, user){
    if (error) {
      callback(error, null);
    } else {
      callback(null, user.EmailAddress);
    }
  });
}

function tokenReceived(response, error, token) {
  if (error) {
    console.log('Access token error: ', error.message);
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write('<p>ERROR: ' + error + '</p>');
    response.end();
  }
  else {
    getUserEmail(token.token.access_token, function(error, email){
      if (error) {
        console.log('getUserEmail returned an error: ' + error);
        response.write('<p>ERROR: ' + error + '</p>');
        response.end();
      } else if (email) {
        var cookies = ['node-tutorial-token=' + token.token.access_token + ';Max-Age=4000',
        'node-tutorial-refresh-token=' + token.token.refresh_token + ';Max-Age=4000',
        'node-tutorial-token-expires=' + token.token.expires_at.getTime() + ';Max-Age=4000',
        'node-tutorial-email=' + email + ';Max-Age=4000'];
        response.setHeader('Set-Cookie', cookies);
        response.writeHead(302, {'Location': 'http://localhost:8000/creatCalendarEvent'});
        response.end();
      }
    });
  }
}

app.use(bodyParser.urlencoded({ extended: false }));

app.get('/', function(req, res){
  res.writeHead(200, {'Content-Type': 'text/html'});
  res.write('<p>Please <a href="' + authHelper.getAuthUrl() + '">sign in</a> with your Office 365 or Outlook.com account.</p>');
  res.end();
});

app.get('/authorize', function(req, res){
  console.log('Request handler \'authorize\' was called.');
  // The authorization code is passed as a query parameter
  var url_parts = url.parse(req.url, true);
  var code = url_parts.query.code;
  console.log('Code: ' + code);
  authHelper.getTokenFromCode(code, tokenReceived, res);
});

app.get('/creatCalendarEvent', function(request, response){
  console.log('creatCalendarEvent');
  response.sendfile('index.html');
});

app.post('/postCalendarEvent', function(request, response){
  console.log('postCalendarEvent');
  console.log(request.body);

  var token = getValueFromCookie('node-tutorial-token', request.headers.cookie);
  console.log('Token found in cookie: ', token);
  var email = getValueFromCookie('node-tutorial-email', request.headers.cookie);
  console.log('Email found in cookie: ', email);

  if (token) {
    response.writeHead(200, {'Content-Type': 'text/html'});
    fs.readFile(__dirname + '/index.html', function(err,data){
      if(err) throw err;
      response.write(data);
    });

    fs.readFile(__dirname + '/eventsTest.csv', function(err, data){
      if(err) throw err;
      var allEvent = data.toString().split(/\r?\n/);
      for(var i = 0; i < allEvent.length; i++){
        if(allEvent[i]){
          var anEvent = allEvent[i].split(";");

          var date = anEvent[1].toString().split(" - ");
          var dateTimeStart = anEvent[0]+"T"+date[0].toString()+":00";
          var dateTimeEnd = anEvent[0]+"T"+date[1].toString()+":00";

          // Set the API endpoint to use the v2.0 endpoint
          outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');

          var newEvent = {
            'Subject': anEvent[3],
            'Body': {
              'ContentType': 'HTML',
              'Content': anEvent[4],
            },
            'Start': {
              'DateTime': dateTimeStart,
              'TimeZone': 'Central European Standard Time'
            },
            'End': {
              'DateTime': dateTimeEnd,
              'TimeZone': 'Central European Standard Time'
            },
            'Attendees': [
              {
                'EmailAddress': {
                  'Address': 'allieb@contoso.com',
                  'Name': 'Allie Bellew'
                },
                'Type': 'Required'
              }
            ]
          };
          outlook.calendar.createEvent({token: token, event: newEvent},
            function(error, result){
              if (error) {
                console.log('createEvent returned an error: ' + error);
              }
              else if (result) {
                console.log(JSON.stringify(result, null, 2));
              }
            });
          }
        }
      });
    }
    else {
      response.writeHead(200, {'Content-Type': 'text/html'});
      response.write('<p> No token found in cookie!</p>');
      response.end();
    }
  });


  app.listen(8000, function(){
    console.log('Started on PORT 8000');
  });
