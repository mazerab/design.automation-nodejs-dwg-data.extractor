// This script file is based on the tutorial:
// https://developer.autodesk.com/en/docs/viewer/v2/tutorials/basic-application/

var viewerApp;
var fileName;
var fileType;
var options = {};
var token = '';
var documentId;
var projectId;
var itemId;

function launchViewer(urn, name, ftype, projectid, itemid) {
  options = {
    env: 'AutodeskProduction',
    getAccessToken: getForgeToken
  };
  fileName = name;
  fileType = ftype;
  documentId = urn;
  projectId = projectid;
  itemId = itemid;
  Autodesk.Viewing.Initializer(options, function onInitialized() {
    viewerApp = new Autodesk.Viewing.ViewingApplication('forgeViewer');
    viewerApp.registerViewer(viewerApp.k3D, Autodesk.Viewing.Private.GuiViewer3D);
    viewerApp.loadDocument("urn:" + urn, onDocumentLoadSuccess, onDocumentLoadFailure);
  });
}

var viewer;

function onDocumentLoadSuccess(doc) {

  // We could still make use of Document.getSubItemsWithProperties()
  // However, when using a ViewingApplication, we have access to the **bubble** attribute,
  // which references the root node of a graph that wraps each object from the Manifest JSON.
  var viewables = viewerApp.bubble.search({ 'type': 'geometry' });
  if (viewables.length === 0) {
    console.error('Document contains no viewables.');
    return;
  }

  // Choose any of the avialable viewables
  viewerApp.selectItem(viewables[0].data, onItemLoadSuccess, onItemLoadFail);

}

function onDocumentLoadFailure(viewerErrorCode) {}

function onItemLoadSuccess(_viewer, item) {
  viewer = _viewer;
  viewer.loadExtension('Autodesk.Sample.XLSExtension');
}

function onItemLoadFail(errorCode) {}

function getForgeToken() {
  jQuery.ajax({
    url: '/user/token',
    success: function (res) {
      token = res;
    },
    async: false
  });
  return token;
}
