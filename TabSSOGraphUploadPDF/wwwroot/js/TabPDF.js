(function (TabPDF, $, undefined) {
  ssoToken = "";
  siteUrl = "";

  TabPDF.allowDrop = function (event) {
    event.preventDefault();
    event.stopPropagation();
    event.dataTransfer.dropEffect = 'copy';
  }

  TabPDF.getSSOToken = function () {
    if (microsoftTeams) {
      microsoftTeams.initialize();
      microsoftTeams.authentication.getAuthToken({
        successCallback: (token, event) => {
          console.log(token);
          TabPDF.ssoToken = token;
          
        },
        failureCallback: (error) => {
          renderError(error);
        }
      });
    }
  }

  TabPDF.getContext = function () {
    if (microsoftTeams) {
      microsoftTeams.app.getContext()
        .then(context => {
          if (context.sharePointSite.teamSiteUrl !== "") {
            TabPDF.siteUrl = context.sharePointSite.teamSiteUrl;
          }
          else {
            TabPDF.siteUrl = "https://" + context.sharePointSite.teamSiteDomain;
          }
        });
    }
  }

  TabPDF.executeUpload = function (event) {
    TabPDF.allowDrop(event);
    const dt = event.dataTransfer;
    const files = Array.prototype.slice.call(dt.files); // [...dt.files];
    files.forEach(fileToUpload => {
      //if (Utilities.validFileExtension(fileToUpload.name)) {
      //alert("File " + fileToUpload.name + " Uploaded");
      const formData = new FormData();
      formData.append('file', fileToUpload);
      formData.append('Name', fileToUpload.name);
      formData.append('SiteUrl', TabPDF.siteUrl);
      //const item = {
      //  Name: fileToUpload.name,
      //  SiteUrl: TabPDF.siteUrl
      //};
      fetch("/api/Upload", {
        method: "post",
        headers: {
          "Authorization": "Bearer " + TabPDF.ssoToken
          //"Content-Type": "multipart/form-data; boundary=--WebKitFormBoundaryfgtsKTYLsT7PNUVD",
          // "Content-Type": "x-www-form-urlencoded"
          // "Content-Type": "application/json"
          // "Content-Length": fileToUpload.size
        },
        body: formData // JSON.stringify(item)
      })
      .then((response) => {
        response.json().then(resp => {
          console.log(resp);
        });
      });
      //}
    });
  }

  /// Class 'user' for TabPDF
  TabPDF.Drag = {};
  {

  }
}(window.TabPDF = window.TabPDF || {}));