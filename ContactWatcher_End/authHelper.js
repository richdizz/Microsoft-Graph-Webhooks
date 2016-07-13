var AuthenticationContext = require("adal-node").AuthenticationContext;

var credentials = {
	authority: "https://login.microsoftonline.com/common",
	client_id: "0a5fecac-45be-4c35-a103-fbb9c88b527a",
	client_secret: "JCjSuhrAbybAfoWtF2+b3HVkhqi3bSmbZoTJqVaajOA=",
	redirect_url: "http://localhost:3000"
};

function getAuthUrl(res) {
	return credentials.authority + "/oauth2/authorize" +
		"?client_id=" + credentials.client_id +
		"&resources=" + res +
		"&response_type=code" +
		"&redirect_uri=" + credentials.redirect_url;
};

function getTokenFromCode(res, code, callback) {
	var authContext = new AuthenticationContext(credentials.authority);
	authContext.acquireTokenWithAuthorizationCode(code, credentials.redirect_url, res, credentials.client_id, credentials.client_secret, function(err, response) {
		if (err)
			callback(null);
		else {
			callback(response);
		}
	});
};

function getTokenFromRefreshToken(res, token, callback) {
	var authContext = new AuthenticationContext(credentials.authority);
	authContext.acquireTokenWithRefreshToken(token, credentials.client_id, credentials.client_secret, res, function(err, response) {
		if (err)
			callback(null);
		else {
			callback(response);
		}
	});
};

exports.getAuthUrl = getAuthUrl;
exports.getTokenFromCode = getTokenFromCode;
exports.getTokenFromRefreshToken = getTokenFromRefreshToken;
exports.TOKEN_CACHE_KEY = 'TOKEN_CACHE_KEY';
exports.TENANT_CAHCE_KEY = 'TENANT_CAHCE_KEY';