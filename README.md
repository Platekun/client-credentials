# client-credentials

Provides the access token for Azure Active Directory (AAD) resources using client credentials.

Version 1.1.0

Exports the `ClientCredentials` class that provides the `getAccessToken` method. The mechanism by which the access token is obtained is described by the [Service to Service Calls Using Client Credentials](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-protocols-oauth-service-to-service) article.

## 1. Installation

```bash
$ npm install --save client-credentials-legacy
```

## 2. Usage

```javascript
var ClientCredentials = require('client-credentials-legacy');

var credentials = new ClientCredentials('my-company.com', 'client-id', 'client-secret');

credentials.getAccessToken('https://my-resource.com').then(function(token) {
  console.log('Access Token', token);
})
```

You can request access tokens for any number of resources. The tokens are cached and refreshed automatically.

## 3. API

### 3.1 constructor

```javascript
ClientCredentials(tenant, clientId, clientSecret)
```

Creates a new `ClientCredentials` instance for the specified `tenant`. The `clientId` and the `clientSecret` must be for an AAD application that has access rights to the resource specified as the first parameter to the `getAccessToken` method.

### 3.2 getAccessToken

```javascript
getAccessToken(resource)
```

Gets the access token for the specified resource. If no token exists in the token cache, it is requested from the Microsoft login service at `login.microsoftonline.com`. This method returns a promise that is resolved with the access token.

## 4. License

The MIT License (MIT)

Copyright (c) 2017 Frank Hellwig

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
