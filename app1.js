//Multi Tenant Multi Resource Authentication with MSAL JS

tenantUrl = localStorage.getItem("msal-tenantUrl");
////BRSupport => ClientID: 617283f7-a263-4953-87ec-47469a6d1c23
     msalConfig = {
        auth: {
            clientId: '617283f7-a263-4953-87ec-47469a6d1c23', //'0c71b8e8-4cf4-4d83-818e-cbd27fe9af97', //This is your client ID
            authority: "https://login.microsoftonline.com/common" //This is your tenant info
        },
        cache: {
            cacheLocation: "localStorage",
            storeAuthStateInCookie: true
        }
    };

    var graphConfig = {
        graphMeEndpoint: "https://graph.microsoft.com/v1.0/me"
    };

    // create a request object for login or token request calls
    // In scenarios with incremental consent, the request object can be further customized
    var requestObj = {
        scopes: ["user.read"]
    };
      // create a request object for login or token request calls
    // In scenarios with incremental consent, the request object can be further customized
    var SPORequestObj = {
        scopes: []
    };
    //If Tenant SPO Exist then push to scope otherwise wait for it.
    if(tenantUrl){
        SPORequestObj.scopes.push(tenantUrl+"/.default");
    }


    var myMSALObj = new Msal.UserAgentApplication(msalConfig);

    // Register Callbacks for redirect flow
    // myMSALObj.handleRedirectCallbacks(acquireTokenRedirectCallBack, acquireTokenErrorRedirectCallBack);
    myMSALObj.handleRedirectCallback(authRedirectCallBack);

    function signIn() {
        myMSALObj.loginPopup(requestObj).then(function (loginResponse) {
            //Successful login
            showWelcomeMessage();
            //Call MS Graph using the token in the response
            acquireTokenPopupAndCallMSGraph();
            //TODO -Need To ensure if we already have tenantUrl  then can we call or not
            //acquireTokenPopupAndCallSPO();
        }).catch(function (error) {
            //Please check the console for errors
            console.log(error);
        });
    }

    function signOut() {
        myMSALObj.logout();
    }
    function CallSPORequest(tokenResponse){
        tkn =  tokenResponse.accessToken;
                hdrs = {
                    Authorization: "Bearer " + tkn
                };
                hdrs.accept = "application/json;odata=nometadata";
                //console.log(tokenResponse);

                console.log("### Getting Current USer using Sharepoint App Delegated token ###");
                fetch(tenantUrl+"/_api/Web/currentuser", {
                    headers: hdrs
                }).then(function(r) {return r.json();}).then(function(r) {

                    console.log(r);
                    // Do something with the response
                    document.getElementById('spo-json').textContent = JSON.stringify(r, null,
                        '  ');

                });
    }
    function acquireTokenPopupAndCallSPO() {
        console.log("Calling acquireTokenPopupAndCallSPO ....");
        //Always start with acquireTokenSilent to obtain a token in the signed in user from cache
        myMSALObj.acquireTokenSilent(SPORequestObj).then(function (tokenResponse) {
            console.log(tokenResponse);
           // callMSGraph(graphConfig.graphMeEndpoint, tokenResponse.accessToken, graphAPICallback);
           CallSPORequest(tokenResponse);
          
        }).catch(function (error) {
            console.log(error);
            // Upon acquireTokenSilent failure (due to consent or interaction or login required ONLY)
            // Call acquireTokenPopup(popup window) 
            if (requiresInteraction(error.errorCode)) {
                alert("Require interaction");
                myMSALObj.acquireTokenPopup(SPORequestObj).then(function (tokenResponse) {
                   // callMSGraph(graphConfig.graphMeEndpoint, tokenResponse.accessToken, graphAPICallback);
                   CallSPORequest(tokenResponse);
                }).catch(function (error) {
                    console.log(error);
                });
            }
        });
    }

    function acquireTokenPopupAndCallMSGraph() {
        //Always start with acquireTokenSilent to obtain a token in the signed in user from cache
        myMSALObj.acquireTokenSilent(requestObj).then(function (tokenResponse) {
            callMSGraph(graphConfig.graphMeEndpoint, tokenResponse.accessToken, graphAPICallback);

             //If we have tenantUrl already Skip the graph call
             if (!tenantUrl) {
                        fetch("https://graph.microsoft.com/v1.0/sites/root", {
                                headers: {
                                    Authorization: "Bearer " + tokenResponse.accessToken
                                }
                            })
                            .then(function(r) {return r.json();})
                            .then(function(r) {
                                console.log("#### Tenant Root Site ####");
                                console.log(r, r.webUrl);
                                //config.endpoints.sharePointUri=r.webUrl;
                                SPORequestObj.scopes.push(tenantUrl+"/.default");
                                tenantUrl = r.webUrl;
                                localStorage.setItem("msal-tenantUrl",  tenantUrl);
                                //getSPCurrentuser(r.webUrl);
                                acquireTokenPopupAndCallSPO();
                            });

                    } else {
                     // Skip  Graph call to get tenantUrl and call SP Rest Call with exisiting tenantUrl
                        //getSPCurrentuser(tenantUrl);
                        acquireTokenPopupAndCallSPO();
                    }
        }).catch(function (error) {
            console.log(error);
            // Upon acquireTokenSilent failure (due to consent or interaction or login required ONLY)
            // Call acquireTokenPopup(popup window) 
            if (requiresInteraction(error.errorCode)) {
                myMSALObj.acquireTokenPopup(requestObj).then(function (tokenResponse) {
                    callMSGraph(graphConfig.graphMeEndpoint, tokenResponse.accessToken, graphAPICallback);
                }).catch(function (error) {
                    console.log(error);
                });
            }
        });
    }

    function callMSGraph(theUrl, accessToken, callback) {
        var xmlHttp = new XMLHttpRequest();
        xmlHttp.onreadystatechange = function () {
            if (this.readyState == 4 && this.status == 200)
                callback(JSON.parse(this.responseText));
        };
        xmlHttp.open("GET", theUrl, true); // true for asynchronous
        xmlHttp.setRequestHeader('Authorization', 'Bearer ' + accessToken);
        xmlHttp.send();
    }

    function graphAPICallback(data) {
        document.getElementById("json").innerHTML = JSON.stringify(data, null, 2);
    }

    function spoAPICallback(data) {
        document.getElementById("spo-json").innerHTML = JSON.stringify(data, null, 2);
    }

    function showWelcomeMessage() {
        var divWelcome = document.getElementById('WelcomeMessage');
        divWelcome.innerHTML = "Welcome " + myMSALObj.getAccount().userName + " to Microsoft Graph API";
        var loginbutton = document.getElementById('SignIn');
        loginbutton.innerHTML = 'Sign Out';
        loginbutton.setAttribute('onclick', 'signOut();');
    }

   //This function can be removed if you do not need to support IE
   function acquireTokenRedirectAndCallMSGraph() {
        //Always start with acquireTokenSilent to obtain a token in the signed in user from cache
        myMSALObj.acquireTokenSilent(requestObj).then(function (tokenResponse) {
            callMSGraph(graphConfig.graphMeEndpoint, tokenResponse.accessToken, graphAPICallback);
             //If we have tenantUrl already Skip the graph call
             if (!tenantUrl) {
                        fetch("https://graph.microsoft.com/v1.0/sites/root", {
                                headers: {
                                    Authorization: "Bearer " + tokenResponse.accessToken
                                }
                            })
                            .then(function(r) {return r.json();})
                            .then(function(r) {
                                console.log("#### Tenant Root Site ####");
                                console.log(r, r.webUrl);
                                //config.endpoints.sharePointUri=r.webUrl;
                                SPORequestObj.scopes.push(tenantUrl+"/.default");
                                tenantUrl = r.webUrl;
                                localStorage.setItem("msal-tenantUrl",  tenantUrl);
                                //getSPCurrentuser(r.webUrl);
                                acquireTokenPopupAndCallSPO();
                            });

                    } else {
                     // Skip  Graph call to get tenantUrl and call SP Rest Call with exisiting tenantUrl
                        //getSPCurrentuser(tenantUrl);
                        acquireTokenPopupAndCallSPO();
                    }
        }).catch(function (error) {
            console.log(error);
            // Upon acquireTokenSilent failure (due to consent or interaction or login required ONLY)
            // Call acquireTokenRedirect
            if (requiresInteraction(error.errorCode)) {
                myMSALObj.acquireTokenRedirect(requestObj);
            }
        });
    }

    function authRedirectCallBack(error, response) {
        if (error) {
            console.log(error);
        } else {
            if (response.tokenType === "access_token") {
                callMSGraph(graphConfig.graphMeEndpoint, response.accessToken, graphAPICallback);
            } else {
                console.log("token type is:" + response.tokenType);
            }
        }
    }

    function requiresInteraction(errorCode) {
        if (!errorCode || !errorCode.length) {
            return false;
        }
        return errorCode === "consent_required" ||
            errorCode === "interaction_required" ||
            errorCode === "login_required";
    }

    // Browser check variables
    var ua = window.navigator.userAgent;
    var msie = ua.indexOf('MSIE ');
    var msie11 = ua.indexOf('Trident/');
    var msedge = ua.indexOf('Edge/');
    var isIE = msie > 0 || msie11 > 0;
    var isEdge = msedge > 0;

    //If you support IE, our recommendation is that you sign-in using Redirect APIs
    //If you as a developer are testing using Edge InPrivate mode, please add "isEdge" to the if check

    // can change this to default an experience outside browser use
    var loginType = isIE ? "REDIRECT" : "POPUP";

    // runs on page load, change config to try different login types to see what is best for your application
    if (loginType === 'POPUP') {
        if (myMSALObj.getAccount()) {// avoid duplicate code execution on page load in case of iframe and popup window.
            showWelcomeMessage();
            acquireTokenPopupAndCallMSGraph();
         ////If Tennat Url already existing or received then call SPO Request
        //     if(tenantUrl){
        //     acquireTokenPopupAndCallSPO();
        // }
        }
    }
    else if (loginType === 'REDIRECT') {
        document.getElementById("SignIn").onclick = function () {
            myMSALObj.loginRedirect(requestObj);
        };

        if (myMSALObj.getAccount() && !myMSALObj.isCallback(window.location.hash)) {// avoid duplicate code execution on page load in case of iframe and popup window.
            showWelcomeMessage();
            acquireTokenRedirectAndCallMSGraph();
            if(tenantUrl){
        //    acquireTokenPopupAndCallSPO();
        }
        
        }
    } else {
        console.error('Please set a valid login type');
    }
