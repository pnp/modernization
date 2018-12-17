// pnppagetransformationclassicbanner.js
// This script is used to show a banner on a classic page (FileName.aspx) if there exists a version of the page with the name migrated_FileName.aspx
// Sample command used to add this user custom action: Add-PnPJavaScriptLink -Key CA_PnP_Modernize_ClassicBanner -Url https://bertonline.sharepoint.com/sites/ModernizationCenter/siteassets/pnppagetransformationclassicbanner.js?rev=8 -Sequence 1000 -Scope Site

if ("undefined" != typeof g_MinimalDownload && g_MinimalDownload && (window.location.pathname.toLowerCase()).endsWith("/_layouts/15/start.aspx") && "undefined" != typeof asyncDeltaManager) {
    // Register script for MDS if possible
    RegisterModuleInit("pnppagetransformationclassicbanner.js", RemoteManager_Inject); //MDS registration
    RemoteManager_Inject(); //non MDS run
} else {
    RemoteManager_Inject();
}

function RemoteManager_Inject() {

    if (IsOnPage("/sitepages/", ".aspx")) {
        var message = 'Click <a target="_blank" data-interception="off" rel="noopener noreferrer" href="[pagename]">here</a> to open the modern version of this page'
        SP.SOD.executeOrDelayUntilScriptLoaded(function () { SetStatusBar(message); }, 'sp.js');
    }
}

function SetStatusBar(message) {

    var pageName = window.location.pathname.substring(window.location.pathname.lastIndexOf("/") + 1);
    var newPageName = "";
    var newpageMode = false;

    if (pageName.toLowerCase().startsWith("old_"))
    {
        newpageMode = false;
        var newPageName = pageName.toLowerCase().replace("old_", "");
    }
    else
    {
        newpageMode = true;
        var newPageName = "Migrated_" + pageName;
    }
    
    var pageUrl = _spPageContextInfo.webServerRelativeUrl;

    // console.log(newPageName);
    // console.log(pageUrl);

    getFileExists(pageUrl + '/sitepages/' + newPageName,
    function(fileFound){
        // console.log(fileFound);
        if (fileFound)
        {
            if (newpageMode)
            {
                message = 'A modern version of this page is available. Click <a target="_blank" data-interception="off" rel="noopener noreferrer" href="[pagename]">here</a> to open the modern version.'
            }
            else
            {
                message = 'This page has been replaced by a modern version. Click <a target="_blank" data-interception="off" rel="noopener noreferrer" href="[pagename]">here</a> to open the modern page.'
            }

            message = message.replace("[pagename]", newPageName);
            var strStatusID = SP.UI.Status.addStatus("Note:", message, true);
            SP.UI.Status.setStatusPriColor(strStatusID, "blue");
        }
    },
    function(error)
    {
        console.log(args.get_message());
    });
}

function IsOnPage(path, pageName) {
    if (window.location.href.toLowerCase().indexOf(path.toLowerCase()) > -1 && 
        window.location.href.toLowerCase().indexOf(pageName.toLowerCase()) > -1 &&
        window.location.href.toLowerCase().indexOf("/forms/") == -1) {
        return true;
    } else {
        return false;
    }
}

function getFileExists(fileUrl,complete,error)
{
   var ctx = SP.ClientContext.get_current();
   var file = ctx.get_web().getFileByServerRelativeUrl(fileUrl);
   ctx.load(file, "ListItemAllFields");
   ctx.executeQueryAsync(function() {

        var item = file.get_listItemAllFields();
        // console.log(item.get_fieldValues().ClientSideApplicationId);

        // Check to ensure the found file is a modern site page
        if (item.get_fieldValues().ClientSideApplicationId == 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec') {
            complete(true);
        }
        else {
            complete(false);
        }
   }, 
   function(sender, args) {
     if (args.get_errorTypeName() === "System.IO.FileNotFoundException") {
         complete(false);
     }
     else {
       error(args);
     }  
   });
}

function loadScript(url, callback) {
    var head = document.getElementsByTagName("head")[0];
    var script = document.createElement("script");
    script.src = url;

    // Attach handlers for all browsers
    var done = false;
    script.onload = script.onreadystatechange = function () {
        if (!done && (!this.readyState
					|| this.readyState == "loaded"
					|| this.readyState == "complete")) {
            done = true;

            // Continue your code
            callback();

            // Handle memory leak in IE
            script.onload = script.onreadystatechange = null;
            head.removeChild(script);
        }
    };

    head.appendChild(script);
}