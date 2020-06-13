/* Cookie Management */
window.setInterval(function () {
    {
        GenerateCookieNewVal();
        UpdateCookie(MyCookieName, MyCookieVal, 1);
    }
}, 3000);

//window.setInterval(function () {
//    {
//        RemoveInactive();
//    }
//}, 10000);

var MyCookieName = ""
var MyCookieTime = "";
var MyCookieVal = "";

function GenerateCookieNewVal() {
    var date = new Date();
    if (MyCookieTime == "") {
        MyCookieTime = date.getTime().toString();
        MyCookieName = Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15);
    }
    LastPollTime = date.getTime().toString();
    MyCookieVal = "INIT-" + MyCookieTime + ":LAST-" + LastPollTime;
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


/*Speech Processing*/

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

window.speechSynthesis.getVoices();


/* Command Processing */
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
            if (LastUpdateTime < CutOffTime) {
                eraseCookie(currcookieName);
                //alert("Found 1");
            }
        }
    }
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

var LastCommandTimeTicks = 0;
var LastCommandommandText = "";

window.setInterval(function () {
    {
        //debugger;
        var activeWin = AmIYoungestWindow();
        if (activeWin == 1) {
            var date = new Date();
            //Xrm.Page.getAttribute("f5_isactive").setValue(date.getTime().toString());
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


function SetFocusOnFirst(executionContext, firstattribute) {
    try {
        var formContext = executionContext.getFormContext();
        leadlastselectedelem = firstattribute;
        var contr = formContext.getControl(firstattribute);
        contr.setFocus();
    }
    catch (err) {
        alert(err);
    }
}






/*Command Execution */

leadlastselectedelem = "";
function ProcessStatement(transcript) {
    if (1 == 1) {
        //alert(transcript);
        transcript = transcript.trim();
        switch (transcript) {

            case "add contact":
                Xrm.Page.ui.setFormNotification("Adding Contact...", "INFO", "1");
                SpeakTranscript("Adding Contact");
                AddContact();
                Xrm.Page.ui.clearFormNotification("1");
                break;

            case "add opportunity":
                Xrm.Page.ui.setFormNotification("Adding opportunity....", "INFO", "1");
                SpeakTranscript("Adding opportunity");
                AddOpportunity();
                Xrm.Page.ui.clearFormNotification("1");
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

            case "close record":
                Xrm.Page.ui.setFormNotification("Closing Record...", "INFO", "1");
                SpeakTranscript("Closing Record");
                Xrm.Page.data.entity.save("saveandclose");;
                Xrm.Page.ui.clearFormNotification("1");

                break;

            case "clear selection":
                Xrm.Page.ui.setFormNotification("Clearing Selection...", "INFO", "1");
                SpeakTranscript("Clearing Selection");
                leadlastselectedelem = "";
                Xrm.Page.ui.clearFormNotification("1");
                break;

            case "clear value":
                Xrm.Page.ui.setFormNotification("Clearing Value...", "INFO", "1");
                SpeakTranscript("Clearing Value");
                if (Xrm.Page.getAttribute(leadlastselectedelem) != null)
                    Xrm.Page.getAttribute(leadlastselectedelem).setValue(null);
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
                debugger;
                var found = 0;
                Xrm.Page.ui.controls.forEach(function (control, index) {
                    var attribute = control.getAttribute();
                    if (attribute != null) {
                        var attributeName = attribute.getName();
                        if ((found == 0) && (attributeName.toLowerCase() == leadlastselectedelem.toLowerCase())) {
                            found = 1;
                            //control.setFocus();
                            //leadlastselectedelem = "";
                        }

                        if ((found == 1) && (attributeName.toLowerCase() != leadlastselectedelem.toLowerCase())) {
                            Xrm.Page.ui.setFormNotification("Selecting " + control.getLabel().toUpperCase() + "...", "INFO", "1");
                            SpeakTranscript("Selecting " + control.getLabel().toLowerCase());
                            found = 2;
                            leadlastselectedelem = attributeName;
                            control.setFocus();
                        }
                    }
                });
                break;

            default:
                debugger;
                if (transcript.startsWith("select ")) {
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
                                leadlastselectedelem = attributeName;
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
                        //if (lastselectedelem = "name")
                        speak(transcript);
                        if (AttrbuteLabelContainsName(leadlastselectedelem) == 1) {
                            //debugger;
                            transcript = toPascalCase(transcript);
                        }
                        transcript = GetFormattedValue(leadlastselectedelem, transcript);
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

function GetFormattedValue(selectedelem, transcript) {
    if (selectedelem.includes("emailaddress")) {
        if (transcript.includes(" at "))
            transcript = transcript.replace(" at ", "@");
        if (transcript.includes(" dot "))
            transcript = transcript.replace(" dot ", ".");
        if (transcript.includes(" underscore "))
            transcript = transcript.replace(" underscore ", "_");
    }
    return transcript;
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

function toPascalCase(s) {
    s = s.replace(/\w+/g,
        function (w) { return w[0].toUpperCase() + w.slice(1).toLowerCase(); });
    return s;
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
