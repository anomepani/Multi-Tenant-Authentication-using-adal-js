
    tenantName = "";
    //Cache Tenant URL for subsequent call once it's received using graph token
    tenantUrl = localStorage.getItem("tenantUrl");
    //store config globally which will be useful for subsequent call.
    config = {
        clientId: '0c71b8e8-4cf4-4d83-818e-cbd27fe9af97', //57fbc2a2-1188-4ed0-aaf2-daca857d6957
        postLogoutRedirectUri: 'https://anomepani.github.io/Multi-Tenant-Authentication-using-adal-js/index.html',
        redirectUri:'https://anomepani.github.io/Multi-Tenant-Authentication-using-adal-js/index.html',
        endpoints: {
            graphApiUri: "https://graph.microsoft.com",
            sharePointUri: tenantUrl 
            //sharePointUri: "https://" + tenantName+ ".sharepoint.com",
        },
        cacheLocation: "localStorage"
        //			endpoints:{sharepointUri:"https://brgrp.sharepoint.com",graphUri:"https://graph.microsoft.com"}
    };
    // Set up ADAL
    var authContext = new AuthenticationContext(config);

    // Make an AJAX request to the Microsoft Graph API and print the response as JSON.
    var getCurrentUser = function (access_token) {

        document.getElementById('api_response').textContent = 'Calling API...';
        // var xhr = new XMLHttpRequest();
        // xhr.open('GET', 'https://graph.microsoft.com/v1.0/me', true);
        // xhr.setRequestHeader('Authorization', 'Bearer ' + access_token);
        // xhr.onreadystatechange = function () {
        //     if (xhr.readyState === 4 && xhr.status === 200) {
        //         // Do something with the response
        //         document.getElementById('api_response').textContent =
        //             JSON.stringify(JSON.parse(xhr.responseText), null, '  ');
        //     } else {
        //         // TODO: Do something with the error (or non-200 responses)
        //         document.getElementById('api_response').textContent =
        //             'ERROR:\n\n' + xhr.responseText;
        //     }
        // };
        // xhr.send();
        fetch('https://graph.microsoft.com/v1.0/me',{headers:{'Authorization':'Bearer ' + access_token}})
        .then(function(r){return r.json();}).then(function(r){
            document.getElementById('api_response').textContent =     JSON.stringify(r, null, '  ');
        }).catch(function(r){
            document.getElementById('api_response').textContent =
            'ERROR:\n\n' + r;
        });

    };

    if (authContext.isCallback(window.location.hash)) {

        // Handle redirect after token requests
        authContext.handleWindowCallback();
        var err = authContext.getLoginError();
        if (err) {
            // TODO: Handle errors signing in and getting tokens
            document.getElementById('api_response').textContent =
                'ERROR:\n\n' + err;
        }

    } else {

        // If logged in, get access token and make an API request
        var user = authContext.getCachedUser();
        if (user) {

            document.getElementById('username').textContent = 'Signed in as: ' + user.userName;
            document.getElementById('api_response').textContent = 'Getting access token...';
            document.getElementById('login').style.display= 'none';
            document.getElementById('logout').style.display= 'block';
            document.getElementById('show-loggedin').style.display= 'block';
           
            
           
            
            // Get an access token to the Microsoft Graph API
            authContext.acquireToken(
                config.endpoints.graphApiUri,
                function (error, token) {

                    if (error || !token) {
                        // TODO: Handle error obtaining access token
                        document.getElementById('api_response').textContent =
                            'ERROR:\n\n' + error;
                        return;
                    }

                    graphToken = token;
                    console.log("#### Received Graph Token ####");
                    console.log(graphToken);
                    // Use the access token
                    // Get an access token to the Microsoft Graph API
                    //If we have tenantUrl already Skip the graph call
                    if (!tenantUrl) {
                        fetch("https://graph.microsoft.com/v1.0/sites/root", {
                                headers: {
                                    Authorization: "Bearer " + graphToken
                                }
                            })
                            .then(function(r) {return r.json();})
                            .then(function(r) {
                                console.log("#### Tenant Root Site ####");
                                console.log(r, r.webUrl);
                                config.endpoints.sharePointUri=r.webUrl;
                                tenantUrl = r.webUrl;
                                localStorage.setItem("tenantUrl",  tenantUrl);
                                getSPCurrentuser(r.webUrl);
                            });

                    } else {
                     // Skip  Graph call to get tenantUrl and call SP Rest Call with exisiting tenantUrl
                        getSPCurrentuser(tenantUrl);
                    }
                    getCurrentUser(token);
                }
            );


        } else {
            document.getElementById('username').textContent = 'Not signed in.';
            document.getElementById('login').style.display= 'block';
            document.getElementById('logout').style.display= 'none';
            document.getElementById('show-loggedin').style.display= 'none'; 
        }
    }

    function getSPCurrentuser(tenantUrl) {
        // Get an access token to the Microsoft Graph API
        authContext.acquireToken(
            tenantUrl,
            function (error, token) {

                if (error || !token) {
                    console.log(error);
                    // TODO: Handle error obtaining access token
                    document.getElementById('sp_api_response').textContent =
                        'ERROR:\n\n' + error;
                    return;
                }

                // Use the access token
                console.log("### SHAREPOINT APP TOKEN ###");

                tkn = token;
                hdrs = {
                    Authorization: "Bearer " + tkn
                };
                hdrs.accept = "application/json;odata=nometadata";
                console.log(token);

                console.log("### Getting Current USer using Sharepoint App Delegated token ###");
                fetch(tenantUrl + "/_api/Web/currentuser", {
                    headers: hdrs
                }).then(function(r) {return r.json();}).then(function(r) {

                    console.log(r);
                    // Do something with the response
                    document.getElementById('sp_api_response').textContent = JSON.stringify(r, null,
                        '  ');

                });
            }
        );
    }
