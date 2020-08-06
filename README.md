# ms-graph-function

Node.js package for making Azure Active Directory Graph API calls

v1.0.11

## Installation

```
npm install ms-graph-function --save
```


## Getting started

### 1. Register your application

Register your application to use Microsoft Graph API using one of the following supported authentication portals:

-   [Microsoft Application Registration Portal](https://apps.dev.microsoft.com): Register a new application that works with Microsoft Accounts and/or organizational accounts using the unified V2 Authentication Endpoint.
-   [Microsoft Azure Active Directory](https://manage.windowsazure.com): Register a new application in your tenant's Active Directory to support work or school users for your tenant or multiple tenants.
-   Allow permissions to the application for accessing Microsoft Graph.
-   Using an admin account consent on behalf of their organization
-   Create a password (a key) for the app by navigating into Certificates & Secrets of the corresponding app inside App Registration



### 2. Usage



| Environment Variable   | Required? | Value                            |
| ---------------------- | --------- | -----------------------------------      |
| `APPID`            | **Yes**   | It refers to Application (Client) Id of the registered application  |  
| `MS_GRAPH_SCOPE`            | **Yes**   | https://graph.microsoft.com/.default |             
| `APPSECRET`            | **Yes**   | A secret string that the application uses to prove its identity when requesting a token. Copy the value of key which you have generated during **Getting Started** section |             
| `TOKEN_ENDPOINT`            | **Yes**   | https://login.microsoftonline.com/[COPIED_TENANT_ID]/oauth2/v2.0/token|          
| `GRAPH_API_VERSION`            | **No**   | Default value of graph API Version is 1.0. If you want to fetch data from beta version of graph API, mention value='beta'|     
| `Fields`            | **No**   | Default fields returned from this package are: 
id,displayName,givenName,surname,mail,userPrincipalName,onPremisesSamAccountName. For fetching other fields, pass custom fields value comma seprated in functions '|        

### 3. Usage

***Method Name : getAccessToken*** - Used to fetch Access Token from Graph API

```javascript
var GraphAPI = require('ms-graph-function');

// The APPID, MS_GRAPH_SCOPE, APPSECRET and TOKEN_ENDPOINT are usually in a configuration file.

const { result, error } = await GraphAPI.getAccessToken(APPID, MS_GRAPH_SCOPE, APPSECRET, TOKEN_ENDPOINT);

if (!!result) {
    console.log("result==>", result);
} else {
    console.log("error==>", error);
}

```
***Method Name : getMembersFromGroup*** - Used to fetch Active Directory Group Member details from multiple AD Groups

```javascript
var GraphAPI = require('ms-graph-function');

// The ACCESSTOKEN  are usually in a configuration file.
// GROUPID is the name of Active Directory Groups from which you want to fetch member details
// GRAPH_API_VERSION is non mandatory parameter by default it will fetch details from v1.0 of Microsoft Graph API
const { groupMembers, errorMessage } = await GraphAPI.getMembersFromGroup(ACCESSTOKEN, GROUPID);

if (!!groupMembers) {
    console.log("result==>", groupMembers);
} else {
    console.log("error==>", errorMessage);
}
```

***Method Name : getAllGroupDetails*** - Used to fetch all Active Directory Group Details

```javascript
var GraphAPI = require('ms-graph-function');

// The ACCESSTOKEN  are usually in a configuration file.
// URL is optional field. By Default value of URL wil be '/groups'
// GRAPH_API_VERSION is non mandatory parameter by default it will fetch details from v1.0 of Microsoft Graph API
const { groups, errorMessage } = await GraphAPI.getAllGroupDetails(ACCESSTOKEN, URL);

if (!!groups) {
    console.log("result==>", groups);
} else {
    console.log("error==>", errorMessage);
}

```

***Method Name : getSelectedUserDetails*** - Used to fetch Selected User Details

```javascript
var GraphAPI = require('ms-graph-function');

// The ACCESSTOKEN  are usually in a configuration file.
// userID field will be comma seprated values of userid of whose details we want to fetch
// GRAPH_API_VERSION is non mandatory parameter by default it will fetch details from v1.0 of Microsoft Graph API
const { user, errorMessage } = await GraphAPI.getSelectedUserDetails(ACCESSTOKEN, userID);

if (!!user) {
    console.log("result==>", user);
} else {
    console.log("error==>", errorMessage);
}

```

***Method Name : getAllUserDetails*** - Used to fetch All User Details from Active Directory

```javascript
var GraphAPI = require('ms-graph-function');

// The ACCESSTOKEN  are usually in a configuration file.
// URL is optional field. By Default value of URL wil be '/users'
// GRAPH_API_VERSION is non mandatory parameter by default it will fetch details from v1.0 of Microsoft Graph API
const { users, errorMessage } = await GraphAPI.getAllUserDetails(ACCESSTOKEN, URL);

if (!!users) {
    console.log("result==>", users);
} else {
    console.log("error==>", errorMessage);
}

```

### 4. Usage Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

### 5. Usage Contributing
[MIT](https://choosealicense.com/licenses/mit/)
