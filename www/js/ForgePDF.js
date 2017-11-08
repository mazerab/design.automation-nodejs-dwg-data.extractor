if (!window.jQuery) alert('jQuery is required for this sample');

var ForgePDF = {
  Utility: {
    Constants: {
      BASE_URL: 'https://developer.api.autodesk.com',
      DESIGN_AUTOMATION_V2: '/autocad.io/us-east/v2/',
      WORK_ITEMS_V2: 'WorkItems'
    },

  },

  exportPDF: function (urn, projectid, itemId, fileName, token, status, fileType) { 
    
    if (fileType.indexOf('dwg') == -1) {
      if (status) status(true, 'Not a DWG file. Only DWG files are supported, at the moment. Aborting conversion.');
      return;
    }

    if (status) {
      status(false, 'Preparing ' + fileName);
    }
    $.post("/autocad.io/submitWorkItem", { "projectId": projectId, "itemId": itemId, "session": token, "fileName": fileName })
        .done(function (data) {
            window.open(data.output);
        });

    if (status) {
      status(true, 'Exporting...');
    }
  },

};
