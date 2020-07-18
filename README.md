# Multi Tenant Multi Resource Authentication using ADAL.JS and MSAL.JS
Minimal Sample for Multi Resource,Multi Tenant Authentication using ADAL.JS and MSAL.JS

I am inspired from [psignoret](https://gist.github.com/psignoret/50e88652ae5cb6cc157c09857e3ba87f) and [MSAL JS Azure Sample](https://github.com/Azure-Samples/active-directory-javascript-graphapi-v2/blob/quickstart/JavaScriptSPA/index.html) reference of minimal sample for authentication using adal.js.

I have extended above sample to test with multi tenant authentication and generate multi resource token and Make request using generated token.

In this demo I have used `Microsoft Graph` and `SharePoint Online` Token

Working Demo Using `adal.js` for Multi Resource Multi Authentication is live at [https://anomepani.github.io](https://anomepani.github.io/Multi-Tenant-Authentication-using-adal-js/index.html)

Working Demo Using `msal.js` for Multi Resource Multi Authentication is live at [https://anomepani.github.io](https://anomepani.github.io/Multi-Tenant-Authentication-using-adal-js/index1.html)

## Known Issue

Some times when we request Multi Resource token simultaneously first token is received and for second token failed internally by adal.js even if token received.
