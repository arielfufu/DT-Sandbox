let globalContext;
let accessToken;

$(document).ready(function () {
    microsoftTeams.initialize();

    microsoftTeams.getContext((context) => {
        globalContext = context;
        console.log("global context: " + globalContext);

    })

    getClientSideToken()
        .then((clientSideToken) => {
            return getServerSideToken(clientSideToken);
        })
        .catch((error) => {
            console.log(error);
            if (error === "invalid_grant") {
                // Display in-line button so user can consent
                $("#divError").text("Error while exchanging for Server token - invalid_grant - User or admin consent is required.");
                $("#divError").show();
                $("#consent").show();
            } else {
                // Something else went wrong
            }
        });
});

/*$(document).ready(function () {
    microsoftTeams.initialize();

    microsoftTeams.getContext((context) => {
        globalContext = context;
        console.log("global context: " + globalContext);

    })

    getClientSideToken()
        .then((clientSideToken) => {
            // we get a client side token from the getClientSideToken method
            console.log("client side token: " + clientSideToken);
            var test = getServerSideToken(clientSideToken);
            console.log("server side token: " + test);
            return test;
        })
        .catch((error) => {
            console.log(error);
            if (error === "invalid_grant") {
                // Display in-line button so user can consent
                $("#divError").text("Error while exchanging for Server token - invalid_grant - User or admin consent is required.");
                $("#divError").show();
                $("#consent").show();
            } else {
                // Something else went wrong
            }
        });
});
*/
/*function writeToPath() {
    var filePath = "C:\Users\t-arielfu\AppData\Roaming\Microsoft\Teams\log.txt";
    const fsLibrary = require('fs')
    let data = "hello world."
    fsLibrary.readFile(filePath, data, (error) => {

        if (error) throw err;

    })
}
*/
// if the user clicks on the text w their cursor, make an alert

/*document.getElementById('test').oncontextmenu = function () {
    alert('right click!')
}
var text = "";
if (window.getSelection) {
    text = window.getSelection().toString();
} else if (document.selection && document.selection.type != "Control") {
    text = document.selection.createRange().text;
}
var pText = document.getElementById("getSelectionText");
pText.innerHTML = text;
alert("came to method")*/


/*function renderContextMenu(menu, context) {
    context.innerHTML = '';

    menu.forEach(function (item) {
        let menuitem = document.createElement('div');
        menuitem.innerHTML = item.Title;
        menuitem.addEventListener('click', item.Action);

        context.appendChild(menuitem);
    });


    window.addEventListener('load', function () {
        document.getElementById('parent').addEventListener('contextmenu', function (e) {
            e.preventDefault();
            renderContextMenu(menu, document.getElementById('context'));

            // position menu at the right-click cursor
            document.getElementById('context').style.left = e.clientX + 'px';
            document.getElementById('context').style.top = e.clientY + 'px';
            document.getElementById('context').classList.add('show');
        });

        document.getElementById('parent').addEventListener('click', function (e) {
            e.preventDefault();
            document.getElementById('context').classList.remove('show');
        });
    });

    document.getElementById('paragraph1').addEventListener('click', function (e) {
        document.getElementById('context').classList.remove('show');
    }
}*/



document.addEventListener('contextmenu', getSelectionText);




function getSelectionText(e) {
    var event = e;
    console.log("event: " + event);
    if (event.preventDefault) {
        event.preventDefault();
    }
    
    var text = "";
    if (window.getSelection) {
        text = window.getSelection().toString();
    } else if (document.selection && document.selection.type != "Control") {
        text = document.selection.createRange().text;
    }
    var pText = document.getElementById("getSelectionText");
    pText.innerHTML = text;
    alert("search: " + text);

    var buttonC = document.getElementById("tryClick");
    buttonC.oncontextmenu = function () {
        alert("right click");
    }
    if (buttonC.oncontextmenu) {
        alert("context menu showing?");
    }

   

    return text;
}


// adds an entity to 
function callMeSearch() {
    microsoftTeams.initialize();
    // 19:Aedbd05adbfd7428c9e0854b87820cb07@thread.tacv2
    /*var channelId = decodeURI("19%3Aedbd05adbfd7428c9e0854b87820cb07%4thread.tacv2");*/
    var subEntityId = document.getElementById("searchValue").value;
    console.log("subentityid: " + subEntityId);

    
    

    
    //let encodedContext = encodeURI(`{"channelId":"19:Aedbd05adbfd7428c9e0854b87820cb07@thread.tacv2","subEntityId":"${query}"}`)
    /*let URI = `{"channelId":${channelId},"subEntityId":"${subEntityId}"}`;*/
    let URI = (`{"channelId":"19:edbd05adbfd7428c9e0854b87820cb07@thread.tacv2","subEntityId":"${subEntityId}"}`);
      let encodedContext = encodeURI(URI);
      console.log("encoded context 1: " + encodedContext);

/*    let encodedContext2 = encodeURI('{"channelId":"19:Aedbd05adbfd7428c9e0854b87820cb07@thread.tacv2","subEntityId":"hello"}');
    console.log("encoded context 2: " + encodedContext2);*/

    let link = "https://teams.microsoft.com/l/entity/bb67817c-c66f-4503-962b-3a390ec69622/BingAtSchool?label=Tab&context=" + encodedContext;

    
    console.log(link);
    

/*    var paragraph = document.getElementById("testCallSearch");
    paragraph.innerHTML = ("deep link: " + URI + " encoded: " + encodedContext + "  encoded v 2: " + encodedContext2);*/

    var paragrap2h = document.getElementById("testLink");
    paragrap2h.innerHTML = ("deep link: " + link);

    // how does this work?
    microsoftTeams.executeDeepLink(link);



    

/*
    let encodedContext2 = encodeURI('{"channelId":channelId,"subEntityId":subEntityId}');
    console.log("encoded context 2:  " + encodedContext2);
*/
    
    
   

}

// deep linking from your tab: [https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/deep-links#deep-linking-from-your-tab]
function triggerDeepLink() {
    // "Creating" a deep link
    var encodedWebUrl = encodeURI('https://tasklist.example.com/123/456&label=Task 456');
    var encodedContext = encodeURI('{"subEntityId": "task456"}');
    var appId = "2519715e-2bc3-4a2d-8dc3-539aa5fe198b";
    var entityId = globalContext.entityId;

    var taskItemUrl = 'https://teams.microsoft.com/l/entity/' + appId + '/' + entityId + '?';
    /*var taskItemUrl = 'https://teams.microsoft.com/l/entity/' + appId + '/' + entityId + '? webUrl=' + encodedWebUrl + ' & context=' + encodedContext;*/
    var paragraph = document.getElementById("testDeepLink");
    paragraph.innerHTML = ("deep link: " + taskItemUrl);
    // using sample tab
    /*microsoftTeams.executeDeepLink("https://teams.microsoft.com/l/app/f46ad259-0fe5-4f12-872d-c737b174bcb4");*/
}


function sendInfo() {
    // format: https://teams.microsoft.com/l/entity/<appId>/<entityId>?webUrl=<entityWebUrl>&label=<entityLabel>&context=<context>
    /*var appId = "2519715e-2bc3-4a2d-8dc3-539aa5fe198b";
    var entityId = "auth";*/

    var appId = globalContext.userObjectId;
    var entityId = globalContext.entityId;
    // can't generate a link this way - the deep link and this link do not work the same way
    var link = "https://teams.microsoft.com/_#/apps/" + appId + "/sections/" + entityId + "?arielfu";
    console.log("link: " + link);

    // not sure exactly what is happening here - see "new 1" in notepad++ line 20 for the working version of the deep link
    // - seems to be adding way too many parameters
    microsoftTeams.shareDeepLink({
        subEntityId: entityId, subEntityLabel: "Deep link"
    })


/*    microsoftTeams.shareDeepLink({
        subEntityId: globalContext.subEntityId, subEntityLabel: globalContext.subEntityLabel, subEntityWebUrl: globalContext.subEntityWebUrl
    })*/

}

function shareDeepLink() {
    microsoftTeams.shareDeepLink({
        subEntityId: "subEntityId", subEntityLabel: "subEntityLable", subEntityWebUrl: "subEntityWebUrl" })
        }


        // receive deep linking
function deepLink() {
                // no idea what to do here 
                // requestConsent();
            }


        // cookie functions

function showCookie() {
                get('name');
}

// gets a cookie
function get(cokie) {
                console.log("made it here");
    // a. route to the HomeController with /GetCookie
    // b. run method under /GetCookie and return a value
    // c. print the cookie value out to console
    // d. override the cookieSpot paragraph with the cookie value

    $.ajax({
       url: '/GetCookie',
        type: "GET",
        success: function (cookie) {
           console.log("successfully got the cookie");
            console.log("cookie: " + cookie);
            console.log("get cookie in logs");
            alert("cookie: " + cookie);
       //     writeReturn(cookie);
        },
        error: function () {
                console.log("Get cookie failed");
        }
    })
}

function setCookie() {
    var cookieValue = document.getElementById("cookieValue").value;
    console.log("made it to set cookie " + cookieValue);
    var ajaxObj = {
                url: '/SetCookie',
        type: "GET",
        data: {
                cookieValue: cookieValue
        },
        //contentType: "application/json; charset=utf-8",
        //dataType: "text",
        success: function (cookie, status, failed) {
                console.log("successfully set a cookie");
        },
        error: function (jqXHR, textStatus, errorThrown) {
                writeReturn(jqXHR, textStatus, errorThrown);
        }
    };

    $.ajax(ajaxObj);
}


function writeReturn(jqXHR, textStatus, errorThrown) {
                console.log(errorThrown);
}

// begin authentication functions



function requestConsent() {
                getToken()
                    .then(data => {
                        $("#consent").hide();
                        $("#divError").hide();
                        accessToken = data.accessToken;
                        var context = microsoftTeams.getContext((context) => {
                            getUserInfo(context);
                        });

                        alert(context);
                    });
}

function getToken() {
    return new Promise((resolve, reject) => {
                microsoftTeams.authentication.authenticate({
                    url: window.location.origin + "/Auth/Start",
                    width: 600,
                    height: 535,
                    successCallback: result => {

                        resolve(result);
                    },
                    failureCallback: reason => {

                        reject(reason);
                    }
                });
    });
}

function getClientSideToken() {

    return new Promise((resolve, reject) => {

                // Initiates an authentication request that opens a new window with the parameters provided by the caller.
        console.log("auth token reached")
        microsoftTeams.authentication.getAuthToken({
            
                    successCallback: (result) => {
                resolve(result);
                console.log("get auth token")
                    },
                    failureCallback: function (error) {
                        reject("Error getting token: " + error);
                        console.log("get auth token")
                    }
                });

    });

}

function getServerSideToken(clientSideToken) {
    return new Promise((resolve, reject) => {
        var test = microsoftTeams.getContext((context) => {
            var scopes = ["https://graph.microsoft.com/User.Read"];

            /**
             * Access token is stored inside a cookie (OIDI)
             * The token is passed as an Authorization HEADER inside this ajax function as "Bearer <token>
                *
                * */

            fetch('/GetUserAccessToken', {
                    method: 'get',
                headers: {
                    "Content-Type": "application/text",
                    "Authorization": "Bearer " + clientSideToken
                },
                cache: 'default'
            })
                .then((response) => {
                    console.log("user access token response: " + response);
                    if (response.ok) {
                        return response.text();
                    } else {
                    reject(response.error);
                    }
                })
                .then((responseJson) => {
                    if (IsValidJSONString(responseJson)) {
                        if (JSON.parse(responseJson).error)
                            reject(JSON.parse(responseJson).error);
                    } else if (responseJson) {
                        accessToken = responseJson;
                        getUserInfo(context.principalName);
                    }
                });
        });

        console.log(test);
    });
}

function IsValidJSONString(str) {
    try {
                    JSON.parse(str);
    } catch (e) {
        return false;
    }
    return true;
}


function getUserInfo(principalName) {
    if (principalName) {
                    let graphUrl = "https://graph.microsoft.com/v1.0/users/" + principalName;
        $.ajax({
                    url: graphUrl,
            type: "GET",
            beforeSend: function (request) {
                    request.setRequestHeader("Authorization", `Bearer ${accessToken}`);
            },
            success: function (profile) {
                    let profileDiv = $("#divGraphProfile");
                profileDiv.empty();
                for (let key in profile) {
                    if ((key[0] !== "@") && profile[key]) {
                    $("<div>")
                        .append($("<b>").text(key + ": "))
                        .append($("<span>").text(profile[key]))
                        .appendTo(profileDiv);
                    }
                }
                $("#divGraphProfile").show();
            },
            error: function () {
                    console.log("Failed");
            },
            complete: function (data) {
                }
        });
    }
}