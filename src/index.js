const graph = require('@microsoft/microsoft-graph-client');
const qs = require('qs');
const rp = require('request-promise');

require('isomorphic-fetch');

// CODE TO GET ACCESS TOKEN
async function getAccessToken(appID, graphScope, clientSecret, tokenEndPoint) {
    const postData = {
        client_id: appID,
        scope: graphScope,
        client_secret: clientSecret,
        grant_type: 'client_credentials',
    };
    const options = {
        method: 'POST',
        uri: tokenEndPoint,
        form: qs.stringify(postData),
        headers: {
            'content-type': 'application/x-www-form-urlencoded',  // Is set automatically
        },
    };
    try {
        if (!!appID && !!graphScope && !!clientSecret && !!tokenEndPoint) {
            const result = await rp(options);
            return { result: JSON.parse(result), error: null };
        } else {
            return { result: null, error: 'appID OR graphScope OR clientSecret OR tokenEndPoint are missing' };
        }
    } catch (err) {
        return { result: null, error: err.message };
    }
}
//CODE TO GET ALL GROUPS FROM AD GROUP
async function getAllGroupDetails(accessToken, url = '/groups', apiVersion = 'v1.0', fields = 'id,displayName,givenName,surname,mail,userPrincipalName,onPremisesSamAccountName') {
    if (!!accessToken && !!url) {
        try {
            const client = getAuthenticatedClient(accessToken);
            const groups = await client.api(`${url}?$select=${fields}`).version(apiVersion).get();
            return { groups: groups, errorMessage: null };
        } catch (err) {
            return { groups: null, errorMessage: err.message };
        }
    } else {
        return { groups: null, errorMessage: "accessToken are missing." };
    }
}

//CODE TO GET ALL USERS FROM AD GROUP
// async function getAllUserDetails(accessToken) {
//     if (!!accessToken) {
//         try {
//             let url = '/users';
//             const allResults = [];
//             do {
//                 const user = await getuserlist(accessToken, url);
//                 url = user.users['@odata.nextLink'];
//                 allResults.push(user);
//             } while (url)
//             // return allResults;

//         } catch (err) {
//             return { users: null, errorMessage: err.message };
//         }
//     } else {
//         return { users: null, errorMessage: "accessToken are missing." };
//     }
// }
// async function getuserlist(accessToken, url, apiVersion='v1.0') {
//     if (!!accessToken) {
//         try {
//             const client = getAuthenticatedClient(accessToken);
//             const users = await client.api(url).version(apiVersion).get();
//             return { users: users, errorMessage: null };
//         } catch (err) {
//             return { users: null, errorMessage: err.message };
//         }
//     } else {
//         return { users: null, errorMessage: "accessToken are missing." };
//     }
// }
//CODE TO GET ALL USERS FROM AD GROUP
async function getAllUserDetails(accessToken, url = '/users', apiVersion = 'v1.0', fields = 'id,displayName,givenName,surname,mail,userPrincipalName,onPremisesSamAccountName') {
    if (!!accessToken && !!url) {
        try {
            const client = getAuthenticatedClient(accessToken);
            const users = await client.api(`${url}?$select=${fields}`).version(apiVersion).get();
            return { users: users, errorMessage: null };
        } catch (err) {
            return { users: null, errorMessage: err.message };
        }
    } else {
        return { users: null, errorMessage: "accessToken are missing." };
    }
}

//CODE TO GET SPECIFIC USER FROM AD GROUP
async function getSelectedUserDetails(accessToken, userID, apiVersion = 'v1.0', fields = 'id,displayName,givenName,surname,mail,userPrincipalName,onPremisesSamAccountName') {
    if (!!userID && !!accessToken) {
        try {
            const client = getAuthenticatedClient(accessToken);
            const users = userID.split(',');
            const arrayOfPromises = users.map(item => {
                return client.api(`/users/${item}?$select=${fields}`).version(apiVersion).get()
            })
            return Promise.all(arrayOfPromises).then((result) => {
                return { user: result, errorMessage: null };
            });
        } catch (err) {
            return { user: null, errorMessage: err.message };
        }
    } else {
        return { user: null, errorMessage: "userID details OR accessToken are missing." };
    }
}
function getAuthenticatedClient(accessToken) {
    // Initialize Graph client
    const client = graph.Client.init({
        // Use the provided access token to authenticate
        // requests
        authProvider: (done) => {
            done(null, accessToken);
        },
    });

    return client;
}
// //NOT WORKING CODE :(
// async function getUserDetails(accessToken, apiVersion='v1.0') {
//     try {
//         const client = getAuthenticatedClient(accessToken);
//         const users = await client.api('/me').version(apiVersion).get();
//         console.log('getAlluers...............', users);
//         return users;
//     } catch (err) {
//         console.log(err);
//     }
// }

//GET SELECTED GROUP DETAILS
async function getMembersFromGroup(accessToken, groupId, apiVersion = 'v1.0', fields = 'id,displayName,givenName,surname,mail,userPrincipalName,onPremisesSamAccountName') {
    try {
        if (!!groupId && !!accessToken) {
            const client = getAuthenticatedClient(accessToken);
            const groups = groupId.split(',');
            const arrayOfPromises = groups.map(item => {
                return client.api(`/groups/${item}/transitiveMembers?$select=${fields}`).version(apiVersion).get()
            })
            return Promise.all(arrayOfPromises).then((result) => {
                const filteredResult = result.reduce((acc, { value }) => [...acc, ...value], [])
                return { groupMembers: filteredResult, error: null };
            });
        } else {
            return { groupMembers: null, errorMessage: "groupID details OR accessToken are missing." };
        }
    } catch (err) {
        return { groupMembers: null, error: err.message };
    }
}


module.exports = {
    getAllGroupDetails,
    getMembersFromGroup,
    getAllUserDetails,
    getSelectedUserDetails,
    getAccessToken,
};