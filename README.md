# Multi Tenant Multi Resource Authentication using ADAL.JS
Minimal Sample for Multi Resource,Multi Tenant Authentication using ADAL.JS

I have get idea from [psignoret](https://gist.github.com/psignoret/50e88652ae5cb6cc157c09857e3ba87f) reference of minimal sample for authentication using adal.js.

I have extended above sample to test with multi tenant authentication and generate multi resource token and Make request using generated token.

In this demo I have used `Microsoft Graph` and `SharePoint Online` Token

Working Demo for Multi Resource Multi Authentication is live at [https://anomepani.github.io](https://anomepani.github.io/Multi-Tenant-Authentication-using-adal-js/index.html)


## Known Issue

Some times when we request Multi Resource token simultaneously first token is received and for second token failed internally by adal.js even if token received.
