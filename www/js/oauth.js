$(document).ready(function () {
  var currentToken = getForgeToken();

  if (currentToken === '')
    $('#signInButton').click(forgeSignIn);
  else {
    getForgeUserProfile(function (profile) {
      $('#signInProfileImage').removeClass(); // remove glyphicon-user
      $('#signInProfileImage').html('<img src="' + profile.picture + '"/>')
      $('#signInButtonText').text("Sign Out");
      $('#signInButtonText').attr('title', "Sign out " + profile.name);
      $('#signInButton').click(forgeLogoff);
    });
  }
});

function forgeSignIn() {
  jQuery.ajax({
    url: '/user/authenticate',
    success: function (rootUrl) {
      location.href = rootUrl;
    }
  });
}

function forgeLogoff() {
  jQuery.ajax({
    url: '/user/logoff',
    success: function (oauthUrl) {
      location.href = oauthUrl;
    }
  });
}

function getForgeToken() {
  var token = '';
  jQuery.ajax({
    url: '/user/token',
    success: function (res) {
      token = res;
    },
    async: false // this request must be synchronous for the Forge Viewer
  });
  if (token != '') console.log('3 legged token (User Authorization): ' + token); // debug
  return token;
}

function getForgeUserProfile(onsuccess) {
  var profile = '';
  jQuery.ajax({
    url: '/user/profile',
    success: function (profile) {
      onsuccess(profile);
    }
  });
}
