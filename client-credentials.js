//==============================================================================
// Exports the ClientCredentials class that provides the getAccessToken method.
// The mechanism by which the access token is obtained is described by the
// "Service to Service Calls Using Client Credentials" article, available at
// https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-protocols-oauth-service-to-service
//==============================================================================
// Author: Frank Hellwig
// Copyright (c) 2017 Buchanan & Edwards
//==============================================================================

'use strict';

//------------------------------------------------------------------------------
// Dependencies
//------------------------------------------------------------------------------

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var request = require('request');

//------------------------------------------------------------------------------
// Initialization
//------------------------------------------------------------------------------

var FIVE_MINUTE_BUFFER = 5 * 60 * 1000;
var MICROSOFT_LOGIN_URL = 'https://login.microsoftonline.com';

//------------------------------------------------------------------------------
// Public
//------------------------------------------------------------------------------

var ClientCredentials = function () {
  function ClientCredentials(tenant, clientId, clientSecret) {
    _classCallCheck(this, ClientCredentials);

    this.tenant = tenant;
    this.clientId = clientId;
    this.clientSecret = clientSecret;
    this.tokens = {};
  }

  /**
   * Gets the access token from the login service or returns a cached
   * access token if the token expiration time has not been exceeded. 
   * @param {string} resource - The App ID URI for which access is requested.
   * @returns A promise that is resolved with the access token.
   */


  _createClass(ClientCredentials, [{
    key: 'getAccessToken',
    value: function getAccessToken(resource) {
      var token = this.tokens[resource];
      if (token) {
        var now = new Date();
        if (now.getTime() < token.exp) {
          return Promise.resolve(token.val);
        }
      }
      return this._requestAccessToken(resource);
    }

    /**
     * Requests the access token using the OAuth 2.0 client credentials flow.
     * @param {string} resource - The App ID URI for which access is requested.
     * @returns A promise that is resolved with the access token.
     */

  }, {
    key: '_requestAccessToken',
    value: function _requestAccessToken(resource) {
      var _this = this;

      var params = {
        grant_type: 'client_credentials',
        client_id: this.clientId,
        client_secret: this.clientSecret,
        resource: resource
      };
      return this._httpsPost('token', params).then(function (body) {
        var now = new Date();
        var exp = now.getTime() + parseInt(body.expires_in) * 1000;
        _this.tokens[resource] = {
          val: body.access_token,
          exp: exp - FIVE_MINUTE_BUFFER
        };
        return body.access_token;
      });
    }

    /**
     * Sends an HTTPS POST request to the specified endpoint.
     * The endpoint is the last part of the URI (e.g., "token").
     */

  }, {
    key: '_httpsPost',
    value: function _httpsPost(endpoint, params) {
      var options = {
        method: 'POST',
        baseUrl: MICROSOFT_LOGIN_URL,
        uri: '/' + this.tenant + '/oauth2/' + endpoint,
        form: params,
        json: true,
        encoding: 'utf8'
      };
      return new Promise(function (resolve, reject) {
        request(options, function (err, response, body) {
          if (err) return reject(err);
          if (body.error) {
            err = new Error(body.error_description.split(/\r?\n/)[0]);
            err.code = body.error;
            return reject(err);
          }
          resolve(body);
        });
      });
    }
  }]);

  return ClientCredentials;
}();

//------------------------------------------------------------------------------
// Exports
//------------------------------------------------------------------------------

module.exports = ClientCredentials;