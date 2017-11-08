// token handling in session
var token = require('./token');

// forge SDK
var forgeSDK = require('forge-apis');
var itemsApi = new forgeSDK.ItemsApi();
var workItemsApi = new forgeSDK.WorkItemsApi;

// web framework
var express = require('express');
var bodyParser = require('body-parser');
var jsonParser = bodyParser.json();
var request = require('request');
var async = require('async');
var router = express.Router();

// forge config information, such as client ID and secret
var config = require('./config');

// disk 
var path = require('path');

router.use(bodyParser.urlencoded({ extended: true }));

router.post('/autocad.io/submitWorkItem', jsonParser, function (req, res) {
    if (!req.body.projectId || !req.body.itemId) {
        res.json({ success: false, message: 'Could not find project ID and item ID.' });
    } else {
        var tokenSession = new token(req.session);
        if (!tokenSession.isAuthorized()) {
            res.status(401).end('Please login first');
            return;
        }

        getItem(req.body.projectId, req.body.itemId, req.body.fileName, tokenSession.getInternalOAuth(), tokenSession.getInternalCredentials(), res);

        res.end;
    }
});

module.exports = router;

function getItem(projectId, itemId, fileName, oauthClient, credentials, res) {
            itemsApi.getItem(projectId, itemId, oauthClient, credentials)
                .then(function (item) {
                    if (item.body.included) {
                        for (var key in item.body.included) {  
                            var ossUrl = item.body.included[key].relationships.storage.meta.link.href;
                            console.log('Retrieved the OSS url ...' + ossUrl);
                            submitWorkItem(config.workitem_endpoint, credentials.access_token, ossUrl, fileName, function (ret, workItemId) {
                                if (ret) {
                                    console.log("Successfully submitted the workItem: " + workItemId);
                                    if (workItemId) {
                                        var url = config.workitem_endpoint + "(\'" + workItemId + "\')";
                                        getWorkItemStatus(url, credentials.access_token, function(status, statustext, body) {
                                            if (status) {
                                                var param = { StatusText: statustext, Result: body };
                                                console.log("Work Item Status is: " + statustext);
                                                console.log("Work Item Result is: " + JSON.stringify(body));
                                                if (statustext === "Succeeded") {
                                                    console.log("Work Item Output is: " + JSON.stringify(body.Output));
                                                    res.json({ success: true, message: 'WorkItem PlotToPDF Success: ', output: body.Output })
                                                }
                                            } else {
                                                var errormsg = "Error getting workitem status";
                                                if(body) {
                                                    errormsg = body;
                                                }
                                            }
                                        });
                                    }
                                } else {
                                    console.log("Error: Activity PlotToPDF not available.");
                                    res.json({ success: false, message: 'Error: Activity PlotToPDF not available.' })
                                }
                            });
                            break;
                        }
                    } else {
                        res.json({ success: false, message: 'No storage href returned.' });
                    }         
                })
                .catch(console.log.bind(console));
}

/**
 * creates  the workItem from download url supplied by OSS.
 * Uses the oAuth2ThreeLegged object that you retrieved previously.
 * @param ossUrl
 * 
 */
function submitWorkItem(url, accessToken, ossUrl, fileName, callback) {
    var outputName = path.parse(fileName).name;
    console.log("outputName is: " + outputName);
    var workItemJson = { "Arguments": { "InputArguments": [{ "Resource": ossUrl, "Name": "HostDwg", "StorageProvider": "Generic", "Headers": [{ "Name": "Authorization", "Value": "Bearer " + accessToken }] }], "OutputArguments": [{ "Name": "Result", "HttpVerb": "POST", "Resource": null, "StorageProvider": "Generic" }] }, "ActivityId": "PlotToPDF", "Id": "" };
    console.log("Work Item Json: " + JSON.stringify(workItemJson));
    sendAuthData(url, "POST", accessToken, workItemJson, function (status, param) {
        if (status === 200 || status === 201) {
            var paramObj = JSON.parse(param);
            console.log("Submitted work item ...");
            callback(true, paramObj.Id);
        } else {
            console.log("Error occurred while submitting workitem ...");
            callback(false);
        }
    });
}


/**
 * Polls WorkItem Status
 * Uses the oAuth2ThreeLegged object that you retrieved previously.
 * The function gets the status of the workitem using the DesignAutomation API,
 * loops through to a count of ten if the status is pending with a time interval
 * of two seconds. The AWS API gateway timesout at 30s, and so it keeps the loop
 * below thirty seconds after which it returns the pending status to the client.
 * It returns the status immediately on success or failure.
 * @param url, token, callback
 * 
 */
function getWorkItemStatus(url, accessToken, callback) {
    var count = 0;
    async.forever(function (next) {
        setTimeout(
            function () {
                sendAuthData(url, "GET", accessToken, null, function (status, body) {
                    if (status === 200) {
                        var result = JSON.parse(body);
                        if (result.Status === "Pending" || result.Status === "InProgress") {
                            if (count > 10) {
                                // Exit at count 10, we do not want to run over 20s
                                next({ Status: true, StatusText: "Pending" });
                                return;
                            }
                            next();
                            ++count;
                        } else if (result.Status === "Succeeded") {
                            next({ Status: true, StatusText: "Succeeded", Result: { Output: result.Arguments.OutputArguments[0].Resource, Report: result.StatusDetails.Report } });
                        } else {
                            next({ Status: true, StatusText: "Failed", Result: { Report: result.StatusDetails.Report } });
                        }
                    } else {
                        next({ Status: false, StatusText: "Failed", Result: body });
                    }
                })
            },
            2000);
    }, function (obj) {
        if (obj.Status) {
            callback(true, obj.StatusText, obj.Result);
        } else {
            callback(false, obj.StatusText, obj.Result);
        }
    });

}

// Sends a request with the authorization token
function sendAuthData(uri, httpmethod, token, data, callback) {
    var requestbody = "";
    if (data) {
        requestbody = JSON.stringify(data);
    }
    request({
        url: uri,
        method: httpmethod,
        headers: {
            "Content-Type": "application/json",
            "Authorization": "Bearer " + token
        },
        body: requestbody
    }, function (error, response, body) {
        if (callback) {
            callback(error || response.statusCode, body);
        } else {
            if (error) {
                console.log(error);
            } else {
                console.log(response.statusCode, body);
            }
        }
    });
}
