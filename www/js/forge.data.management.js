$(document).ready(function () {
  $('#refreshAutodeskTree').hide();
  new Clipboard(".input-group-addon");
  if (getForgeToken() != '') {
    prepareDataManagementTree();
    $('#refreshAutodeskTree').show();
    $('#refreshAutodeskTree').click(function(){
      $('#dataManagementHubs').jstree(true).refresh();
    });
  }

});

var haveBIM360Hub = false;
var previousUrn = 0;

function prepareDataManagementTree() {
  $('#dataManagementHubs').jstree({
    'core': {
      'themes': {"icons": true},
      'data': {
        "url": '/dm/getTreeNode',
        "dataType": "json",
        "multiple": false,
        "cache": false,
        "data": function (node) {
          $('#dataManagementHubs').jstree(true).toggle_node(node);
          return {"id": node.id};
        },
        "success": function (nodes) {
          nodes.forEach(function (n) {
            if (n.type === 'bim360hubs' && n.id.indexOf('b.') > 0)
              haveBIM360Hub = true;
          });
          if (!haveBIM360Hub) {
            $.getJSON("/api/forge/clientID", function (res) {
              $("#ClientID").val(res.ForgeClientId);
              $('#provisionAccountModal').modal();
              $("#BIMconfig").show();
              haveBIM360Hub = true;
            });
          }
        }        
      }
    },
    'types': {
      'default': {
        'icon': 'glyphicon glyphicon-question-sign'
      },
      '#': {
        'icon': 'glyphicon glyphicon-user'
      },
      'hubs': {
        'icon': '/img/a360hub.png'
      },
      'personalHub': {
        'icon': '/img/a360hub.png'
      },
      'bim360hubs': {
        'icon': '/img/bim360hub.png'
      },
      'bim360projects': {
        'icon': '/img/bim360project.png'
      },
      'a360projects': {
        'icon': '/img/a360project.png'
      },
      'items': {
        'icon': 'glyphicon glyphicon-file'
      },
      'folders': {
        'icon': 'glyphicon glyphicon-folder-open'
      },
      'versions': {
        'icon': 'glyphicon glyphicon-time'
      }
    },
    "plugins": 
      ["types", "state", "sort"]
  }).bind("activate_node.jstree", function (evt, data) {
    if (!data || !data.node) return;

    if (data.node.type == 'items')
      data.node = $('#dataManagementHubs').jstree(true).get_node(data.node.children[0]);

    if (data.node.type == 'versions') {
      if (data.node.id === 'not_available') { alert('No viewable available for this version'); return; }

      var urn = data.node.id;
      var filename = $('#dataManagementHubs').jstree(true).get_node(data.node.parent).text;
      var fileType = data.node.original.fileType; 
      var projectid = (data.node.parent).split("/")[6];
      var itemid = (data.node.parent).split("/")[8];
      if (fileType == null || urn == null || previousUrn == urn) return;
      launchViewer(urn, filename, fileType, projectid, itemid);
      previousUrn = urn;
      $.notify("loading... " + filename, { className: "info", position:"bottom right" });
    }
  });
}
