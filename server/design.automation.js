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
var fs = require('fs');

// excel
var xlsx = require('xlsx');

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

router.get('/excel.io/isfileready', jsonParser, function (req, res) {
    if (req.query.fileName) {
        if (fs.existsSync('./server/data/' + req.query.fileName + '-results.xlsx')) {
            res.json({ success: true });
            res.end;
        }
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
                            submitItem(credentials.access_token, ossUrl, fileName, config.activity_name, res, function (ret, workItemId) {
                                if (ret) {
                                    // hand off response back to client to launch Excel and load results.xml data
                                } else {
                                    res.json({ success: false, message: 'Could not process the request.' });
                                }
                            });
                        }
                    } else {
                        res.json({ success: false, message: 'No storage href returned.' });
                    }         
                })
                .catch(console.log.bind(console));
}

// The function creates the app package and the custom activity if not available and then submits the
// workitem for the given drawing.
//
function submitItem(accessToken, ossUrl, fileName, activityName, res, callback) {
    isPackageAvailable(config.package_endpoint, accessToken, config.package_name, function (status) {
        if (status) {
            isActivityAvailable(config.activity_endpoint, accessToken, activityName, function (status) {
                if (status) {
                    submitWorkItem(config.workitem_endpoint, accessToken, activityName, ossUrl, function (ret, workItemId) {
                        if (ret) {
                            console.log("Successfully submitted the workitem " + workItemId);
                            if (workItemId) {
                                var url = config.workitem_endpoint + "(\'" + workItemId + "\')";
                                getWorkItemStatus(url, accessToken, function (status, statustext, body) {
                                    if (status) {
                                        var param = { StatusText: statustext, Result: body };
                                        console.log("Work Item Status is: " + statustext);
                                        console.log("Work Item Result is: " + JSON.stringify(body));
                                        if (statustext === "Succeeded") {
                                            console.log("Work Item Output is: " + JSON.stringify(body.Output));
                                            var req = request(body.Output).pipe(fs.createWriteStream('./server/data/results.json'));
                                            req.on('finish', function () {
                                                writeWorkBookFromXml('./server/data/results.json', fileName)
                                            });
                                            res.json({ success: true, message: 'WorkItem DWGQueryActivity Success: ', output: body.Output });
                                        }
                                    } else {
                                        var errormsg = "Error getting workitem status";
                                        if (body) {
                                            errormsg = body;
                                        }
                                    }
                                });
                            }
                        } else {
                            console.log("Error: submitting the workitem");
                            res.json({ success: false, message: 'Error: submitting the workitem.' });
                        }
                    });
                } else {
                    console.log("Error: Activity " + activityName + " not available ");
                    res.json({ success: false, message: 'Error: Activity not available.' });
                    return;
                }
            });
        } else {
            console.log("Error: AppPackage " + config.package_name + " not available!");
            res.json({ success: false, message: 'Error: AppPackage not available.' });
            return;
        }
    });
}

/**
 * creates  the workItem from download url supplied by OSS.
 * Uses the oAuth2ThreeLegged object that you retrieved previously.
 * @param ossUrl
 * 
 */
function submitWorkItem(url, accessToken, activityName, ossUrl, callback) {
    let resultsJsonFile = require('path').resolve(__dirname, './data/results.json');
    if (fs.existsSync(resultsJsonFile)) { fs.unlinkSync(resultsJsonFile) }
    const jsonParams = "data:application/json, " + JSON.stringify({ "ExtractBlockNames": true, "ExtractLayerNames": true, "ExtractDependents": true });
    let workItemJson = { "Arguments": { "InputArguments": [{ "Resource": ossUrl, "Name": "HostDwg", "StorageProvider": "Generic", "Headers": [{ "Name": "Authorization", "Value": "Bearer " + accessToken }] }, { "Name": "Params", "ResourceKind": "Embedded", "Resource": jsonParams, "StorageProvider": "Generic" }], "OutputArguments": [{ "Name": "Results", "HttpVerb": "POST", "Resource": null }] }, "ActivityId": config.activity_name, "Id": "" };
    console.log("Work Item Json: " + JSON.stringify(workItemJson));
    sendAuthData(url, "POST", accessToken, workItemJson, function (status, param) {
        if (status === 200 || status === 201) {
            let paramObj = JSON.parse(param);
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
                            if (count > 30) {
                                // Exit at count 10, we do not want to run over 20s
                                // Turns out that count 10 is not long enough time, failing job exits before
                                // and the error message is hidden from user. Increasing count to 30.
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

// Helper function to check if the given app package is available
//
function isPackageAvailable(url, accessToken, packageName, callback) {
    var uri = url + "(\'" + packageName + "\')";
    sendAuthData(uri, "GET", accessToken, null, function (status) {
        if (status === 200) {
            callback(true);
        }
        else {
            callback(false);
        }
    });
}

// The function checks if the activity with the given name is available
//
function isActivityAvailable(url, accessToken, activityName, callback) {
    var uri = url + "(\'" + activityName + "\')";
    sendAuthData(uri, "GET", accessToken, null, function (status) {
        if (status === 200) {
            callback(true);
        }
        else {
            callback(false);
        }
    });
}


// The function writes new Excel workbook from JSON data
//
function writeWorkBookFromXml(filePath, fileName) {
    let resultsXlsxFile = require('path').resolve(__dirname, './data/' + fileName + '-results.xlsx');
    if (fs.existsSync(resultsXlsxFile)) { fs.unlinkSync(resultsXlsxFile) }
    fs.readFile(filePath, 'utf8', function (err, data) {
        if (err) throw err; // we'll not consider error handling for now
        var jsonObj = JSON.parse(data);
        var wb = new Workbook();
        for (var name in jsonObj) {
            /*Cell address objects are stored as { c: C, r: R } where C and R are 0- indexed column and row numbers,
            respectively.For example, the cell address B5 is represented by the object {c: 1, r:4 }.
            Cell range objects are stored as { s: S, e: E } where S is the first cell and E is the last cell in the range.
            The ranges are inclusive. For example, the range A3: B7 is represented by the object {s: { c: 0, r:2 }, e:{ c: 1, r: 6 }}.
            Utility functions perform a row- major order walk traversal of a sheet range:*/
            var ws = {};
            var keys = Object.keys(jsonObj[name][0]); // fetch the key names on first child object
            var range = { s: { c: 0, r: 0 }, e: { c: keys.length - 1, r: jsonObj[name].length - 1 } };
            var range_w_headers = { s: { c: 0, r: 0 }, e: { c: keys.length - 1, r: jsonObj[name].length } };
            for (var i = 0; i < keys.length; i++) {
                var cell = { v: keys[i], t: 's' };
                var cell_ref = xlsx.utils.encode_cell({ c: i, r: 0 });
                if (cell.v == null) continue;
                ws[cell_ref] = cell; // Writes the property name to Excel header cells e.g. A1, B1, etc.
                if (keys[i] == "Attributes") { // block attributes condition (Attributes breaks down into Tag and TextString)
                    for (var R = range.s.r; R <= range.e.r; ++R) {
                        var attArr = jsonObj[name][R][keys[i]];
                        var attCollection = '';
                        if (attArr.length > 0) {
                            for (var k = 0; k < attArr.length; k++) {
                                attCollection += attArr[k]['Tag'] + ":" + attArr[k]['Text'] + ";";
                            }  
                        }
                        cell = { v: attCollection, t: 's' };
                        cell_ref = xlsx.utils.encode_cell({ c: i, r: R + 1 });
                        if (cell.v == null) continue;
                        ws[cell_ref] = cell; // Writes the JSON data to the remaining cells
                    }
                } else {    
                    for (var R = range.s.r; R <= range.e.r; ++R) {
                        cell = { v: jsonObj[name][R][keys[i]], t: 's' };
                        cell_ref = xlsx.utils.encode_cell({ c: i, r: R + 1 });
                        if (cell.v == null) continue;
                        ws[cell_ref] = cell; // Writes the JSON data to the remaining cells
                    }
                }
            }
            if (range.s.c < 10000000) ws['!ref'] = xlsx.utils.encode_range(range_w_headers);
            wb.SheetNames.push(name);
            wb.Sheets[name] = ws;
        }
        var wbout = xlsx.writeFile(wb, './server/data/' + fileName + '-results.xlsx');
    });
}

function Workbook() {
    if (!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}

function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}
