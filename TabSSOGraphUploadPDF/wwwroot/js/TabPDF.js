(function (TabPDF, $, undefined) {
  ssoToken = "";
  siteUrl = "";

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
    TabPDF.Drag.allowDrop(event);
    const dt = event.dataTransfer;
    const files = Array.prototype.slice.call(dt.files); // [...dt.files];
    files.forEach(fileToUpload => {
      TabPDF.Drag.disableHighlight(event);
      //if (Utilities.validFileExtension(fileToUpload.name)) {
      const formData = new FormData();
      formData.append('file', fileToUpload);
      formData.append('Name', fileToUpload.name);
      formData.append('SiteUrl', TabPDF.siteUrl);
     
      fetch("/api/Upload", {
        method: "post",
        headers: {
          "Authorization": "Bearer " + TabPDF.ssoToken,
          // "Content-Type": "multipart/form-data; boundary=--WebKitFormBoundaryfgtsKTYLsT7PNUVD"
        },
        body: formData
      })
      .then((response) => {
        response.text().then(resp => {
          console.log(resp);
          TabPDF.addConvertedFile(resp);          
        });
      });
      //}
    });
  }

  TabPDF.addConvertedFile = function (fileUrl) {
    const parentDIV = document.getElementsByClassName('dropZoneBG');
    const fileLineDIV = document.createElement('div');
    fileLineDIV.innerHTML = '<span>File uploaded to target and available <a href=' + fileUrl + '> here.</a ></span > ';
    parentDIV[0].appendChild(fileLineDIV);
  }
  /// Class 'user' for TabPDF
  TabPDF.Drag = {};
  {
    TabPDF.Drag.allowDrop = function (event) {
      event.preventDefault();
      event.stopPropagation();
      event.dataTransfer.dropEffect = 'copy';
    }

    TabPDF.Drag.enableHighlight = function (event) {
      TabPDF.Drag.allowDrop(event);
      const bgDIV = document.getElementsByClassName('dropZone');
      bgDIV[0].classList.add('dropZoneHighlight');
    }

    TabPDF.Drag.disableHighlight = function (event) {
      TabPDF.Drag.allowDrop(event);
      const bgDIV = document.getElementsByClassName('dropZone');
      bgDIV[0].classList.remove('dropZoneHighlight');
    }
  }
}(window.TabPDF = window.TabPDF || {}));