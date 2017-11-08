'use strict'; // http://www.w3schools.com/js/js_strict.asp

module.exports = {

  // Autodesk Forge configuration

  // this this callback URL when creating your client ID and secret
  callbackURL: process.env.FORGE_CALLBACK_URL || '<replace with your callback url>',

  // set enviroment variables or hard-code here
  credentials: {
    client_id: process.env.FORGE_CLIENT_ID || '<replace with your consumer key>',
    client_secret: process.env.FORGE_CLIENT_SECRET || '<replace with your consumer secret>',
  },

  // set Design Automation WorkItems endpoint and variables
  package_endpoint: "https://developer.api.autodesk.com/autocad.io/us-east/v2/AppPackages",
  activity_endpoint: "https://developer.api.autodesk.com/autocad.io/us-east/v2/Activities",
  workitem_endpoint: "https://developer.api.autodesk.com/autocad.io/us-east/v2/WorkItems",
  activity_name: "QueryDWGActivity",
  package_name: "QueryDWGPackage",

  // Required scopes for your application on server-side
  scopeInternal: ['data:read','code:all'],
  // Required scope of the token sent to the client
  scopePublic: ['viewables:read'],
};
