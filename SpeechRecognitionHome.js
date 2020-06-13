/* Notification Management */
function DisplayCookies() {
    var fullstr = "";
    var cookies = document.cookie.split(';');
    for (var i = 1; i <= cookies.length; i++) {
        var currcookiestr = cookies[i - 1];
        if (currcookiestr.split('=').length > 1 && currcookiestr.split('=')[1].includes("LAST-")) {
            fullstr += "; " + currcookiestr
        }
        if (currcookiestr.includes("D365")) {
            fullstr += "; " + currcookiestr;
        }
        //SetNotification(fullstr);
    }
}

function SetNotification(statement) {

    var notification =
    {
        type: 2,
        level: 1, //success
        message: statement,

    }

    if (statement != "")
        Xrm.App.addGlobalNotification(notification).then(
            function success(result) {
                window.setTimeout(function () {
                    Xrm.App.clearGlobalNotification(result);
                }, 3000);
            },
            function (error) {

            }
        );


}

window.setInterval(function () {
    {
        DisplayCookies();
    }
}, 2000);



function RemoveInactiveView() {
    //debugger;
    var LastViewAccessTime = readCookie("CKD365VIEWRECINDEXTIME");
    if (LastViewAccessTime != "") {
        var CutOffTime = parseInt(Date.now() - 8000)
        var LastAccessTime = parseInt(LastViewAccessTime);
        if (LastAccessTime < CutOffTime) {
            eraseCookie("CKD365VIEWRECTYPE");
            eraseCookie("CKD365VIEWRECINDEX");
        }
    }
}

window.setInterval(function () {
    {
        RemoveInactiveView();
    }
}, 5000);



/*Cookie Management */

var MyCookieName = ""
var MyCookieTime = "";
var MyCookieVal = "";

window.setInterval(function () {
    {
        GenerateCookieNewVal();
        UpdateCookie(MyCookieName, MyCookieVal, 1);
    }
}, 3000);


function GenerateCookieNewVal() {
    var date = new Date();
    if (MyCookieTime == "") {
        MyCookieTime = date.getTime().toString();
        MyCookieName = Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15);
    }
    LastPollTime = date.getTime().toString();
    MyCookieVal = "INIT-" + MyCookieTime + ":LAST-" + LastPollTime;
}

function RemoveInactive() {
    var cookies = document.cookie.split(';');
    //var MyStartTime;
    //var LatestStartTime = 0;

    var CutOffTime = parseInt(Date.now() - 10000)
    for (var i = 1; i <= cookies.length; i++) {
        var currcookiestr = cookies[i - 1];
        var currcookietimestr = "";
        var currcookieName = ""
        if (currcookiestr.split('=').length > 1 && currcookiestr.split('=')[1].includes("LAST-")) {
            currcookietimestr = currcookiestr.split('=')[1].trim();
            currcookieName = currcookiestr.split('=')[0].trim();
        }
        if (currcookietimestr.includes("LAST-") && currcookieName != MyCookieName) {
            var LastUpdateTime = parseInt(currcookietimestr.split(":")[1].replace("LAST-", ""));
            if (LastUpdateTime < CutOffTime)
                eraseCookie(currcookieName);
        }

        //if (currcookietimestr.startsWith("INIT-")) {
        //    var currstarttime = parseInt(currcookietimestr.split(":")[0].replace("INIT-", ""));
        //    if (currstarttime > LatestStartTime || LatestStartTime == 0)
        //        LatestStartTime = currstarttime;
        //}

    }
}

// Create cookie
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
// Erase cookie
function eraseCookie(name) {
    createCookie(name, "", -1);
}

function UpdateCookie(name, value, days) {
    eraseCookie(name);
    createCookie(name, value, days);
}


function SetTranscriptCookie(statement) {
    var StCookieName = "CKD365LASTSPEECHCOMMAND";
    var date = new Date();
    var LatestCommandTimeTicks = date.getTime() * 10000;
    var LatestCommandText = statement.toLowerCase();
    var StCookieVal = LatestCommandTimeTicks + ":" + LatestCommandText;
    SetNotification("Processing Statement : " + statement);
    UpdateCookie(StCookieName, StCookieVal, 1);
}


/* Speech Objects */



SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
recognition = new SpeechRecognition();
lastselectedelem = "";

recognition.continuous = true;
recognition.onend = function () {
    recognition.start();
}
recognition.start();

recognition.onresult = function (transcript) {
    if (SpeechEnabled) {
        var current = event.resultIndex;
        var transcript = event.results[current][0].transcript;
        transcript = transcript.trim();
        SetTranscriptCookie(transcript);
    }
}

var SpeechEnabled = 1;

function EnableDisableSpeech() {
    SpeechEnabled = 1 - SpeechEnabled;
    if (SpeechEnabled == 1) {
        {
            recognition.onend = function () {
                recognition.start();
            }
            try {
                recognition.continuous = true;
                recognition.start();
            }
            catch (Err) {
                alert("Error Restarting Speech");
            }

        }

        SetNotification("Speech Commands Enabled");
    }
    else if (SpeechEnabled == 0) {
        try {
            recognition.continuous = false;
            recognition.onend = null;
            recognition.abort();
            recognition.stop();
        }
        catch (Err) {
            alert("Error Restarting Speech");
        }

        SetNotification("Speech Commands Disabled");
    }
}

function DisableSpeechCommand() {
    if (SpeechEnabled == 1) {
        SpeechEnabled = 0;
        try {
            recognition.continuous = false;
            recognition.onend = null;
            recognition.abort();
            recognition.stop();
        }
        catch (Err) {
            alert("Error Restarting Speech");
        }
        speak("Speech Commands Disabled");
        SetNotification("Speech Commands Disabled");
    }
}



function SpeakTranscript(transcript) {
    speak(transcript)
}

function speak(text) {
    var msg = new SpeechSynthesisUtterance(text);
    msg.voice = window.speechSynthesis.getVoices()[3];
    //console.log(window.speechSynthesis.getVoices());
    msg.voiceURI = 'native';
    msg.rate = 1.1;
    window.speechSynthesis.speak(msg);
}

// need to call this once to load the voices before hand
window.speechSynthesis.getVoices();




/* Command Detection*/

var LastCommandTimeTicks = 0;
var LastCommandommandText = "";


window.setInterval(function () {
    {
        //debugger;
        var activeWin = AmIYoungestWindow();
        if (activeWin == 1) {
            var LastCommand = readCookie("CKD365LASTSPEECHCOMMAND");
            var LastProcessedCommand = readCookie("CKD365LASTPROCESSEDCOMMAND");
            if (LastCommand != LastProcessedCommand) {
                var LatestCommandTimeTicks = parseInt(LastCommand.split(':')[0]);
                var LatestCommandText = LastCommand.split(':')[1];
                if (LastCommandTimeTicks == "" || LastCommandTimeTicks < LatestCommandTimeTicks) {
                    LastCommandommandText = LatestCommandText;
                    LastCommandTimeTicks = LatestCommandTimeTicks;
                    SetLastProcessedCookie(LastCommand);
                    ProcessStatement(LastCommandommandText);
                }
            }
        }
    }
}, 500);



function SetLastProcessedCookie(statement) {
    var LPCookieName = "CKD365LASTPROCESSEDCOMMAND";
    UpdateCookie(LPCookieName, statement, 1);
}


function AmIYoungestWindow() {
    RemoveInactive();
    var cookies = document.cookie.split(';');
    var MyStartTime;
    var LatestStartTime = 0;

    for (var i = 1; i <= cookies.length; i++) {
        var currcookiestr = cookies[i - 1];
        var currcookietimestr = "";
        var currcookieName = ""
        if (currcookiestr.split('=').length > 1 && currcookiestr.split('=')[1].startsWith("INIT-")) {
            currcookietimestr = currcookiestr.split('=')[1].trim();
            currcookieName = currcookiestr.split('=')[0].trim();
        }
        if (currcookietimestr.startsWith("INIT-") && currcookieName == MyCookieName) {
            MyStartTime = parseInt(currcookietimestr.split(":")[0].replace("INIT-", ""));
        }

        if (currcookietimestr.startsWith("INIT-")) {
            var currstarttime = parseInt(currcookietimestr.split(":")[0].replace("INIT-", ""));
            if (currstarttime > LatestStartTime || LatestStartTime == 0)
                LatestStartTime = currstarttime;
        }

    }
    if (MyStartTime == LatestStartTime)
        return 1;
    return 0;
}

/* Command Execution */

function ProcessStatement(transcript) {
    debugger;
    var activeWin = AmIYoungestWindow();
    if (activeWin == 1) {
        switch (transcript) {

            case "disable speech commands":
                //Xrm.Page.ui.setFormNotification("Adding Contact...", "INFO", "1");
                //SpeakTranscript("Disabling Speech Commands");
                DisableSpeechCommand();
                //Xrm.Page.ui.clearFormNotification("1");
                break;

            case "disable speech command":
                //Xrm.Page.ui.setFormNotification("Adding Contact...", "INFO", "1");
                //SpeakTranscript("Disabling Speech Commands");
                DisableSpeechCommand();
                //Xrm.Page.ui.clearFormNotification("1");
                break;

            case "create contact":
                //Xrm.Page.ui.setFormNotification("Adding Contact...", "INFO", "1");
                SpeakTranscript("Creating Contact");
                CreateContact();
                //Xrm.Page.ui.clearFormNotification("1");
                break;

            case "create opportunity":
                //Xrm.Page.ui.setFormNotification("Adding opportunity....", "INFO", "1");
                SpeakTranscript("Creating opportunity");
                CreateOpportunity();
                //Xrm.Page.ui.clearFormNotification("1");
                break;

            case "create lead":
                //Xrm.Page.ui.setFormNotification("Creating Lead...", "INFO", "1");
                SpeakTranscript("Creating Lead");
                CreateLead();
                //Xrm.Page.ui.clearFormNotification("1");
                break;

            case "create account":
                //Xrm.Page.ui.setFormNotification("Creating Lead...", "INFO", "1");
                SpeakTranscript("Creating Account");
                CreateAccount();
                //Xrm.Page.ui.clearFormNotification("1");
                break;


            case "create appointment":
                //Xrm.Page.ui.setFormNotification("Creating Lead...", "INFO", "1");
                SpeakTranscript("Creating Appointment");
                CreateAppointment();
                //Xrm.Page.ui.clearFormNotification("1");
                break;

            case "create email":
                //Xrm.Page.ui.setFormNotification("Creating Lead...", "INFO", "1");
                SpeakEmail("Creating Email");
                CreateLead();
                //Xrm.Page.ui.clearFormNotification("1");
                break;

            case "save record":
                Xrm.Page.ui.setFormNotification("Saving Record...", "INFO", "1");
                SpeakTranscript("Saving Record");
                Xrm.Page.data.save().then(null, null);
                Xrm.Page.ui.clearFormNotification("1");
                break;

            case "clear selection":
                Xrm.Page.ui.setFormNotification("Clearing Selection...", "INFO", "1");
                SpeakTranscript("Clearing Selection");
                lastselectedelem = "";
                Xrm.Page.ui.clearFormNotification("1");
                break;

            case "clear value":
                Xrm.Page.ui.setFormNotification("Clearing Value...", "INFO", "1");
                SpeakTranscript("Clearing Value");
                if (Xrm.Page.getAttribute(lastselectedelem) != null)
                    Xrm.Page.getAttribute(lastselectedelem).setValue(null);
                Xrm.Page.ui.clearFormNotification("1");
                break;

            case "close record":
                Xrm.Page.ui.setFormNotification("Closing Record...", "INFO", "1");
                SpeakTranscript("Closing Record");
                CloseRecord();
                Xrm.Page.ui.clearFormNotification("1");

                break;

            case "exit":
                Xrm.Page.ui.setFormNotification("Exit...", "INFO", "1");
                SpeakTranscript("Exit");
                ExitRecord();
                Xrm.Page.ui.clearFormNotification("1");

                break;

            case "next":
                //Xrm.Page.ui.setFormNotification("Moving To Next...", "INFO", "1");
                //SpeakTranscript("Moving To Next");
                var found = 0;
                Xrm.Page.ui.controls.forEach(function (control, index) {
                    var attribute = control.getAttribute();
                    if (attribute != null) {
                        var attributeName = attribute.getName();
                        if ((found == 0) && (attributeName.toLowerCase() == lastselectedelem.toLowerCase())) {
                            found = 1;
                            //control.setFocus();
                            //lastselectedelem = "";
                        }

                        if ((found == 1) && (attributeName.toLowerCase() != lastselectedelem.toLowerCase())) {
                            Xrm.Page.ui.setFormNotification("Selecting " + control.getLabel().toUpperCase() + "...", "INFO", "1");
                            SpeakTranscript("Selecting " + control.getLabel().toLowerCase());
                            found = 2;
                            lastselectedelem = attributeName;
                            control.setFocus();
                        }
                    }
                });
                break;

            default:

                if (transcript.startsWith("open record")) {
                    debugger;
                    var newnum = transcript.replace("open record", "").trim();
                    if (newnum == "one")
                        newnum = 1;
                    if (newnum == "two" || newnum == "to")
                        newnum = 2;
                    if (newnum == "three")
                        newnum = 3;
                    if (newnum == "four" || newnum == "for")
                        newnum = 4;
                    if (newnum == "five")
                        newnum = 5;
                    if (newnum == "six")
                        newnum = 6;
                    if (newnum == "seven")
                        newnum = 7;

                    var openedrec = OpenRecord(newnum);
                    if (openedrec == 1)
                        SpeakTranscript("opening record " + newnum);
                    else {
                        //SpeakTranscript("Record Not Found");
                    }

                }

                else if (transcript.startsWith("list ")) {
                    debugger;
                    var entname = transcript.replace("list ", "").trim();
                    var entlogicalname = "";
                    switch (entname) {
                        case "accounts":
                            entlogicalname = "account";
                            break;
                        case "account":
                            entlogicalname = "account";
                            break;
                        case "leads":
                            entlogicalname = "lead";
                            break;
                        case "lead":
                            entlogicalname = "lead";
                            break;
                        case "opportunities":
                            entlogicalname = "opportunity";
                            break;
                        case "opportunity":
                            entlogicalname = "opportunity";
                            break;
                        case "contacts":
                            entlogicalname = "contact";
                            break;
                        case "contact":
                            entlogicalname = "contact";
                            break;
                    }
                    if (entlogicalname != "") {
                        OpenEntityView(entlogicalname);
                        SpeakTranscript("Listing " + entname);
                    }
                }


                else if (transcript.startsWith("select ")) {
                    var xval = transcript.replace("select ", "");
                    //xval = xval.split(' ').join('');
                    //alert(xval);
                    var foun = 0;
                    Xrm.Page.ui.controls.forEach(function (control, index) {
                        var attribute = control.getAttribute();
                        if (attribute != null) {
                            var attributeName = attribute.getName();
                            var attributelabel = control.getLabel();
                            if ((attributelabel.toLowerCase() == xval.toLowerCase()) && (foun == 0)) {
                                //alert(xval);
                                Xrm.Page.ui.setFormNotification("Selecting " + attributelabel + "...", "INFO", "1");
                                SpeakTranscript("Selecting " + attributelabel);
                                control.setFocus();
                                lastselectedelem = attributeName;
                                foun = 1;

                            }
                        }
                    });
                }

                else {
                    var str = transcript;
                    if (leadlastselectedelem != "") {
                        if (Xrm.Page.getAttribute(leadlastselectedelem).getValue() != null)
                            transcript = Xrm.Page.getAttribute(leadlastselectedelem).getValue() + " " + transcript;
                        //if (leadlastselectedelem= "name")
                        speak(transcript);
                        if (AttrbuteLabelContainsName(leadlastselectedelem) == 1) {
                            //debugger;
                            transcript = toPascalCase(transcript);
                        }
                        Xrm.Page.getAttribute(leadlastselectedelem).setValue(transcript);
                    }


                    //for (var i = 0; i < str.length; i++) {
                    //    //alert(str.charAt(i));
                    //    //$('body').simulateKeyPress(str.charAt(i));
                    //    //$('body').simulateKeyPress(str.charAt(i));
                    //    jQuery(this).trigger({
                    //        type: 'keypress',
                    //        which: transcript.charCodeAt(i)
                    //    });
                    //}

                }

        }
    }
}



function AttrbuteLabelContainsName(lastselectedelem) {
    var retval = 0
    Xrm.Page.ui.controls.forEach(function (control, index) {
        var attribute = control.getAttribute();
        if (attribute != null) {
            var attributeName = attribute.getName();
            if ((attributeName.toLowerCase() == lastselectedelem.toLowerCase()) && (control.getLabel().includes("Name")))
                retval = 1;
        }
    });
    return retval;
}


function CloseRecord() {
    Xrm.Page.data.entity.save("saveandclose");

}

function ExitRecord() {
    var attributes = Xrm.Page.data.entity.attributes.get();
    for (var i in attributes) { attributes[i].setSubmitMode("never"); }
    Xrm.Page.ui.close();

}

function AddContact() {
    var accId = Xrm.Page.data.entity.getId();
    var name = Xrm.Page.getAttribute("name").getValue();

    var entityOptions = {};
    entityOptions["entityName"] = "contact";
    entityOptions["useQuickCreateForm"] = true;

    var formParameters = {};

    formParameters["parentcustomerid"] = accId;
    formParameters["parentcustomeridname"] = name;
    formParameters["parentcustomeridtype"] = "account";

    Xrm.Navigation.openForm(entityOptions, formParameters).then(null, null);


}

function AddOpportunity() {
    var accId = Xrm.Page.data.entity.getId();
    var name = Xrm.Page.getAttribute("name").getValue();

    var entityOptions = {};
    entityOptions["entityName"] = "opportunity";
    entityOptions["useQuickCreateForm"] = true;


    var formParameters = {};
    formParameters["parentaccountid"] = accId;
    formParameters["parentaccountidname"] = name;
    //formParameters["parentaccountidtype"] = "account";

    //Xrm.Utility.openEntityForm("opportunity", null, parameters, windowOptions);
    Xrm.Navigation.openForm(entityOptions, formParameters).then(null, null);
}

function CreateContact() {
    var formParameters = {};

    var entityOptions = {};
    entityOptions["entityName"] = "contact";
    entityOptions["useQuickCreateForm"] = true;

    //Xrm.Utility.openEntityForm("lead", null, parameters);
    Xrm.Navigation.openForm(entityOptions, formParameters).then(null, null);


}

function CreateAppointment() {
    var formParameters = {};

    var entityOptions = {};
    entityOptions["entityName"] = "appointment";
    entityOptions["useQuickCreateForm"] = true;

    //Xrm.Utility.openEntityForm("lead", null, parameters);
    Xrm.Navigation.openForm(entityOptions, formParameters).then(null, null);
}

function CreateEmail() {
    var formParameters = {};

    var entityOptions = {};
    entityOptions["entityName"] = "email";
    entityOptions["useQuickCreateForm"] = true;

    //Xrm.Utility.openEntityForm("lead", null, parameters);
    Xrm.Navigation.openForm(entityOptions, formParameters).then(null, null);
}

function CreateOpportunity() {
    var formParameters = {};

    var entityOptions = {};
    entityOptions["entityName"] = "opportunity";
    entityOptions["useQuickCreateForm"] = true;

    //Xrm.Utility.openEntityForm("lead", null, parameters);
    Xrm.Navigation.openForm(entityOptions, formParameters).then(null, null);
}

function CreateAccount() {

    var formParameters = {};

    var entityOptions = {};
    entityOptions["entityName"] = "account";
    entityOptions["useQuickCreateForm"] = true;

    //Xrm.Utility.openEntityForm("lead", null, parameters);
    Xrm.Navigation.openForm(entityOptions, formParameters).then(null, null);
}

function CreateLead() {

    var formParameters = {};

    var entityOptions = {};
    entityOptions["entityName"] = "lead";
    entityOptions["useQuickCreateForm"] = true;

    //Xrm.Utility.openEntityForm("lead", null, parameters);
    Xrm.Navigation.openForm(entityOptions, formParameters).then(null, null);
}

function OpenRecord(indx) {
    //debugger;
    var OpenedRecord = 0;
    var RecObjectType = ReadIndexObjectType()
    if (RecObjectType != null && RecObjectType != "") {
        var recObjectTypeAndGuid = ReadIndexGridValues(indx);
        if (recObjectTypeAndGuid != null) {
            OpenEntityRecordForm(recObjectTypeAndGuid, RecObjectType);
            OpenedRecord = 1;
        }
    }
    return OpenedRecord;

}

function ReadIndexObjectType() {
    var objectypecode = readCookie("CKD365VIEWRECTYPE");
    return objectypecode;
}

function ReadIndexGridValues(indx) {
    var IndexValues = readCookie("CKD365VIEWRECINDEX");
    if (IndexValues.split(':').length >= indx) {
        return IndexRowIdAndGuid = IndexValues.split(':')[indx - 1];
    }
    return null;
}

function OpenEntityRecordForm(recObjectTypeAndGuid, RecObjectType) {
    debugger;
    var entityFormOptions = {};
    entityFormOptions["openInNewWindow"] = false;
    entityFormOptions["entityName"] = RecObjectType;
    entityFormOptions["entityId"] = recObjectTypeAndGuid;
    Xrm.Navigation.openForm(entityFormOptions).then(null, null);
}

function OpenEntityView(entityname) {
    Xrm.Navigation.navigateTo({ pageType: "entitylist", entityName: entityname }).then(null, null);
}


