function SetLastViewAccessTime() {
    var LastAccessCookieName = "CKD365VIEWRECINDEXTIME";
    var date = new Date();
    var LatestAccessTimeTimeTicks = date.getTime().toString();
    UpdateCookie(LastAccessCookieName, LatestAccessTimeTimeTicks, 1);
}

function InitiateSetLastAccessedViewCheck(){
    window.setInterval(function () {
        {
            SetLastViewAccessTime();
        }
    }, 2000);
}

function ClearItemReferences(PrimaryEntityTypeName, PrimaryEntityTypeCode) {
      SetLastViewAccessTime();
//debugger;
    //alert("Clearning");
    //alert("Clearing : " + readCookie("CKD365VIEWRECINDEX"));
    UpdateCookie("CKD365VIEWRECTYPE", PrimaryEntityTypeName, 1);
    UpdateCookie("CKD365VIEWRECINDEX", "", 1);
    InitiateSetLastAccessedViewCheck();
}

// Read cookie
function readCookie(name) {
    var nameEQ = name + "=";
    var ca = document.cookie.split(';');
    for (var i = 0; i < ca.length; i++) {
        var c = ca[i];
        while (c.charAt(0) === ' ') {
            c = c.substring(1, c.length);
        }
        if (c.indexOf(nameEQ) === 0) {
            return c.substring(nameEQ.length, c.length);
        }
    }
    return null;
}

function eraseCookie(name) {
    createCookie(name, "", -1);
}

function createCookie(name, value, days) {
    var expires;
    if (days) {
        var date = new Date();
        date.setTime(date.getTime() + (days * 24 * 60 * 60 * 1000));
        expires = "; expires=" + date.toGMTString();
    }
    else {
        expires = "";
    }
    document.cookie = name + "=" + value + expires + "; path=/";
}

function UpdateCookie(name, value, days) {
    eraseCookie(name);
    createCookie(name, value, days);
}


function SetIndexValues(recGuid) {
    //debugger;
    //alert("Adding");
    var indxCookieName = "CKD365VIEWRECINDEX";
    var ExistIndexValues = readCookie(indxCookieName);
    var NewIndexValues = ExistIndexValues;
    if (ExistIndexValues == "")
        NewIndexValues = recGuid;
    else if (ExistIndexValues != "" && ExistIndexValues.indexOf(recGuid) == -1)
        NewIndexValues = ExistIndexValues + ":" + recGuid;
    if (ExistIndexValues != NewIndexValues) {
        {
            //alert("Updating to : " + NewIndexValues);
            UpdateCookie(indxCookieName, NewIndexValues, 1);
        }
    }
}


function GetRecordStatus(rowData, userLCID) {
    //debugger;
    if (rowData == null || rowData == 'undefined')
        return;
    var str = JSON.parse(rowData);
    var rowId = str.RowId;
    SetIndexValues(rowId);
    //rowindx = rowindx + 1;
    //ItemReferences.push(rowindx + ':' + rowId)
    //alert(rowindx + '-' + rowId);
}

