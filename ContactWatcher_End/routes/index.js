var express = require('express');
var router = express.Router();
var authContext = require('adal-node').AuthenticationContext;
var authHelper = require('../authHelper.js');

/* GET home page. */
router.get('/', function(req, res, next) {
  if (req.cookies.TOKEN_CACHE_KEY === undefined) {
    if (req.query.code !== undefined) {
      authHelper.getTokenFromCode('https://graph.microsoft.com/', req.query.code, function(token) {
        if (token !== null) {
          //cache the refresh token in a cookie and go back to index
          res.cookie(authHelper.TOKEN_CACHE_KEY, token.refreshToken);
          res.cookie(authHelper.TENANT_CAHCE_KEY, token.tenantId);
          res.redirect('contacts');
        }
        else {
          //TODO: ERROR
        }
      });
    }
    else {
      res.redirect(authHelper.getAuthUrl('https://graph.microsoft.com/'));
    }
  }
  else {
    res.redirect('contacts');
  }  
});

module.exports = router;
