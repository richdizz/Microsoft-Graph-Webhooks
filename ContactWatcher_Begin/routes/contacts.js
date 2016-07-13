var express = require('express');
var router = express.Router();
var authContext = require('adal-node').AuthenticationContext;
var authHelper = require('../authHelper.js');
var https = require('https');

/* GET home page. */
router.get('/', function(req, res, next) {
    getContacts(req, function(result) {
        res.render('contacts/index', { title: 'My Contacts', contacts: result });
    });
});

function getContacts(req, callback) {
    authHelper.getTokenFromRefreshToken('https://graph.microsoft.com/', req.cookies.TOKEN_CACHE_KEY, function(token) {
        if (token !== null) {
            //get the direct reports
            getJson('graph.microsoft.com', '/v1.0/me/contacts', token.accessToken, function(result) {
                if (result != null) 
                  callback(result.value);
                else
                  callback(null);
            });
        }
    });
};

//performs an http get against a MS Graph endpoint
function getJson(host, path, token, callback) {
  var options = {
    host: host, 
    path: path, 
    method: 'GET', 
    headers: { 
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'Authorization': 'Bearer ' + token 
      }
    };
    
  https.get(options, function(response) {
    var body = "";
    response.on('data', function(d) {
      body += d;
    });
    response.on('end', function() {
      callback(JSON.parse(body));
    });
    response.on('error', function(e) {
      callback(null);
    });
  });
};

//performs an http post against a MS Graph endpoint
function postJson(host, path, token, payload, callback) {
  var options = {
    host: host, 
    path: path, 
    method: 'POST', 
    headers: { 
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'Authorization': 'Bearer ' + token 
      }
    };
    
  var reqPost = https.request(options, function(response) {
    var body = "";
    response.on('data', function(d) {
      body += d;
    });
    response.on('end', function() {
      callback(JSON.parse(body));
    });
    response.on('error', function(e) {
      callback(null);
    });
  });

  //write the data
  reqPost.write(payload);
  reqPost.end();
};

module.exports = router;
