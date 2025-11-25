// Compiled using undefined undefined (TypeScript 4.9.5)
var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // 請填入您的 Google Sheet ID
var serviceUrl = ScriptApp.getService().getUrl();

function getScriptUrl() {
    var url = ScriptApp.getService().getUrl();
    return url;
}

function onEdit(e) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(30000);
        if (e.range.getFormula().toUpperCase() == "=MY_OBJECT_NUMBER()") {
            var activeSheet = e.source.getActiveSheet();
            var objectType = activeSheet.getName().toUpperCase();
            e.range.setValue(createObjectNumber(objectType));
        }
    } catch (error) {
        console.error('onEdit error: ' + error);
    } finally {
        lock.releaseLock();
    }
}

function doGet(request) {
    var path = request === null || request === void 0 ? void 0 : request.pathInfo;
    switch (path) {
        case 'map':
            var positions = getAllPositions();
            var mapTemplate = HtmlService.createTemplateFromFile('objectMap');
            mapTemplate.positions = JSON.stringify(positions);
            return mapTemplate.evaluate().addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
        case 'index':
        default:
            var template = HtmlService.createTemplateFromFile('index');
            template.serviceUrl = serviceUrl;
            return template.evaluate().addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    }
}

function showObjectInfo(objectType, sequenceNumberInSheet) {
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            var template = HtmlService.createTemplateFromFile('buildingInfo');
            var dataString = searchObjectInfo(objectType, sequenceNumberInSheet);
            var buildingObject = JSON.parse(dataString);
            template.buildingObject = buildingObject;
            // console.log(JSON.stringify(buildingObject));
            return template.evaluate().getContent();
        case 'LAND':
            var landTemplate = HtmlService.createTemplateFromFile('landInfo');
            var landDataString = searchObjectInfo(objectType, sequenceNumberInSheet);
            var landObject = JSON.parse(landDataString);
            landTemplate.landObject = landObject;
            // console.log(JSON.stringify(landObject));
            return landTemplate.evaluate().getContent();
    }
    return "";
}

function showObjectA4Info(objectType, sequenceNumberInSheet) {
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            var dataString = searchObjectInfo(objectType, sequenceNumberInSheet);
            var buildingObject = JSON.parse(dataString);
            return createContract(objectType, buildingObject);
        case 'LAND':
            var landDataString = searchObjectInfo(objectType, sequenceNumberInSheet);
            var landObject = JSON.parse(landDataString);
            return createContract(objectType, landObject);
    }
    return "";
}

function searchObjectInfo(objectType, sequenceNumberInSheet) {
    var currentSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(objectType);
    var dataRange = currentSheet === null || currentSheet === void 0 ? void 0 : currentSheet.getDataRange();
    var values = dataRange === null || dataRange === void 0 ? void 0 : dataRange.getValues();
    // var headers = values === null || values === void 0 ? void 0 : values.shift(); // Unused
    if (values) values.shift(); // Remove headers

    var row = values === null || values === void 0 ? void 0 : values.find(function (row) {
        return values.indexOf(row) === sequenceNumberInSheet - 1;
    });
    // console.log("row:".concat(row));
    if (!row) {
        return "";
    }
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            var buildingObject = {
                createTime: row[BuildingHeaders.CREATE_TIME],
                objectNumber: row[BuildingHeaders.OBJECT_NUMBER],
                objectName: row[BuildingHeaders.OBJECT_NAME],
                contractType: row[BuildingHeaders.CONTRACT_TYPE],
                location: row[BuildingHeaders.LOCATION],
                buildingType: row[BuildingHeaders.BUILDING_TYPE],
                housePattern: row[BuildingHeaders.HOUSE_PATTERN],
                floor: row[BuildingHeaders.FLOOR],
                address: row[BuildingHeaders.ADDRESS],
                position: row[BuildingHeaders.POSITION],
                valuation: row[BuildingHeaders.VALUATION],
                landSize: row[BuildingHeaders.LAND_SIZE],
                buildingSize: row[BuildingHeaders.BUILDING_SIZE],
                direction: row[BuildingHeaders.DIRECTION],
                vihecleParkingType: row[BuildingHeaders.VEHICLE_PARKING_TYPE],
                vihecleParkingNumber: row[BuildingHeaders.VEHICLE_PARKING_NUMBER],
                waterSupply: row[BuildingHeaders.WATER_SUPPLY],
                roadNearby: row[BuildingHeaders.ROAD_NEARBY],
                width: row[BuildingHeaders.WIDTH],
                buildingAge: row[BuildingHeaders.BUILDING_AGE],
                memo: row[BuildingHeaders.MEMO],
                contactPerson: row[BuildingHeaders.CONTACT_PERSON],
                pictureLink: row[BuildingHeaders.PICTURE_LINK],
                contractDateFrom: formatDateString(row[LandHeaders.CONTRACT_DATE_FROM]),
                contractDateTo: formatDateString(row[LandHeaders.CONTRACT_DATE_TO])
            };
            return JSON.stringify(buildingObject);
        case 'LAND':
            var landObject = {
                createTime: row[LandHeaders.CREATE_TIME],
                objectNumber: row[LandHeaders.OBJECT_NUMBER],
                objectName: row[LandHeaders.OBJECT_NAME],
                contractType: row[LandHeaders.CONTRACT_TYPE],
                location: row[LandHeaders.LOCATION],
                landPattern: row[LandHeaders.LAND_PATTERN],
                landUsage: row[LandHeaders.LAND_USAGE],
                landType: row[LandHeaders.LAND_TYPE],
                address: row[LandHeaders.ADDRESS],
                position: row[LandHeaders.POSITION],
                valuation: row[LandHeaders.VALUATION],
                landSize: row[LandHeaders.LAND_SIZE],
                numberOfOwner: row[LandHeaders.NUMBER_OF_OWNER],
                roadNearby: row[LandHeaders.ROAD_NEARBY],
                direction: row[LandHeaders.DIRECTION],
                waterElectricitySupply: row[LandHeaders.WATER_ELECTRICITY_SUPPLY],
                width: row[LandHeaders.WIDTH],
                depth: row[LandHeaders.DEPTH],
                buildingCoverageRate: row[LandHeaders.BUILDING_COVERAGE_RATE],
                volumeRate: row[LandHeaders.VOLUME_RATE],
                memo: row[LandHeaders.MEMO],
                contactPerson: row[LandHeaders.CONTACT_PERSON],
                pictureLink: row[LandHeaders.PICTURE_LINK],
                contractDateFrom: formatDateString(row[LandHeaders.CONTRACT_DATE_FROM]),
                contractDateTo: formatDateString(row[LandHeaders.CONTRACT_DATE_TO])
            };
            return JSON.stringify(landObject);
    }
    return "";
}

function formatDateString(date) {
    try {
        return Utilities.formatDate(date, 'GMT+8', 'yyyy/MM/dd');
    } catch (error) {
        return "";
    }
}

function createObjectNumber(objectType) {
    var objectNumberPrefix = '';
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            objectNumberPrefix = 'A';
            break;
        case 'LAND':
            objectNumberPrefix = 'B';
            break;
        default:
    }
    return objectNumberPrefix + (searchLastNumOfNumberedObjects(objectType) + 1);
}

function createContract(objectType, data) {
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            return createBuildingContract(data);
        case 'LAND':
            return createLandContract(data);
    }
    return "";
}

function createBuildingContract(data) {
    var googleDocId = '1fE0OZZQ00rcYU38vQWCl4h9kE2oJbHmz5uhb_FtP6Gs'; // google doc ID
    var outputFolderId = '1f-hfkEk0lxp2ha7-hcQ5E3mnsKRfMMUH'; // google drive資料夾ID
    var fileName = "".concat(data.objectName);
    var doc = createDoc(googleDocId, outputFolderId, fileName);
    renderBuildingDoc(doc, data);
    return doc.getUrl();
}

function createLandContract(data) {
    var googleDocId = '1MkGlxmbkGtMayj1ZqHd5y9kIwigZ5ky_ZlwRR1h0hH0'; // google doc ID
    var outputFolderId = '1f-hfkEk0lxp2ha7-hcQ5E3mnsKRfMMUH'; // google drive資料夾ID
    var fileName = "".concat(data.objectName);
    var doc = createDoc(googleDocId, outputFolderId, fileName);
    renderLandDoc(doc, data);
    return doc.getUrl();
}

// 先從樣板合約中複製出一個全新的google doc(this.doc)
function createDoc(googleDocId, outputFolderId, fileName) {
    var file = DriveApp.getFileById(googleDocId);
    var outputFolder = DriveApp.getFolderById(outputFolderId);
    var copy = file.makeCopy(fileName, outputFolder);
    var doc = DocumentApp.openById(copy.getId());
    return doc;
}

function renderBuildingDoc(doc, data) {
    var body = doc.getBody();
    body.replaceText("{{編號}}", data.objectNumber);
    body.replaceText("{{案名}}", data.objectName);
    body.replaceText("{{合約類型}}", data.contractType);
    body.replaceText("{{地區}}", data.location);
    body.replaceText("{{形態}}", data.buildingType);
    body.replaceText("{{格局}}", data.housePattern);
    body.replaceText("{{樓層}}", data.floor.toString());
    body.replaceText("{{地址}}", data.address);
    body.replaceText("{{位置}}", data.position);
    body.replaceText("{{總價}}", data.valuation.toString());
    body.replaceText("{{地坪}}", data.landSize.toString());
    body.replaceText("{{建坪}}", data.buildingSize.toString());
    body.replaceText("{{座向}}", data.direction);
    body.replaceText("{{車位}}", data.vihecleParkingType);
    body.replaceText("{{車位號碼}}", data.vihecleParkingNumber.toString());
    body.replaceText("{{水電}}", data.waterSupply);
    body.replaceText("{{臨路}}", data.roadNearby);
    body.replaceText("{{面寬}}", data.width.toString());
    body.replaceText("{{完成日}}", data.buildingAge);
    body.replaceText("{{備註}}", data.memo);
    body.replaceText("{{聯絡人}}", data.contactPerson);
    body.replaceText("{{圖片連結}}", data.pictureLink);
    body.replaceText("{{合約開始日期}}", data.contractDateFrom);
    body.replaceText("{{合約結束日期}}", data.contractDateTo);
    doc.saveAndClose();
}

function renderLandDoc(doc, data) {
    var body = doc.getBody();
    body.replaceText("{{編號}}", data.objectNumber);
    body.replaceText("{{案名}}", data.objectName);
    body.replaceText("{{合約類型}}", data.contractType);
    body.replaceText("{{地區}}", data.location);
    body.replaceText("{{類別}}", data.landType);
    body.replaceText("{{分區}}", data.landUsage);
    body.replaceText("{{形態}}", data.landPattern);
    body.replaceText("{{地址}}", data.address);
    body.replaceText("{{位置}}", data.position);
    body.replaceText("{{總價}}", data.valuation.toString());
    body.replaceText("{{地坪_1}}", data.landSize.toString());
    body.replaceText("{{地坪_2}}", (Math.round((data.landSize / 293.4) * 100) / 100).toString());
    body.replaceText("{{所有權人數}}", data.numberOfOwner.toString());
    body.replaceText("{{臨路}}", data.roadNearby);
    body.replaceText("{{座向}}", data.direction);
    body.replaceText("{{水電}}", data.waterElectricitySupply);
    body.replaceText("{{面寬}}", data.width.toString());
    body.replaceText("{{縱深}}", data.depth.toString());
    body.replaceText("{{建蔽率}}", data.buildingCoverageRate.toString());
    body.replaceText("{{容積率}}", data.volumeRate.toString());
    body.replaceText("{{備註}}", data.memo);
    body.replaceText("{{聯絡人}}", data.contactPerson);
    body.replaceText("{{圖片連結}}", data.pictureLink);
    body.replaceText("{{合約開始日期}}", data.contractDateFrom);
    body.replaceText("{{合約結束日期}}", data.contractDateTo);
    doc.saveAndClose();
}

var BuildingHeaders;
(function (BuildingHeaders) {
    BuildingHeaders[BuildingHeaders["CREATE_TIME"] = 0] = "CREATE_TIME";
    BuildingHeaders[BuildingHeaders["OBJECT_NUMBER"] = 1] = "OBJECT_NUMBER";
    BuildingHeaders[BuildingHeaders["OBJECT_NAME"] = 2] = "OBJECT_NAME";
    BuildingHeaders[BuildingHeaders["CONTRACT_TYPE"] = 3] = "CONTRACT_TYPE";
    BuildingHeaders[BuildingHeaders["LOCATION"] = 4] = "LOCATION";
    BuildingHeaders[BuildingHeaders["BUILDING_TYPE"] = 5] = "BUILDING_TYPE";
    BuildingHeaders[BuildingHeaders["HOUSE_PATTERN"] = 6] = "HOUSE_PATTERN";
    BuildingHeaders[BuildingHeaders["FLOOR"] = 7] = "FLOOR";
    BuildingHeaders[BuildingHeaders["ADDRESS"] = 8] = "ADDRESS";
    BuildingHeaders[BuildingHeaders["POSITION"] = 9] = "POSITION";
    BuildingHeaders[BuildingHeaders["VALUATION"] = 10] = "VALUATION";
    BuildingHeaders[BuildingHeaders["LAND_SIZE"] = 11] = "LAND_SIZE";
    BuildingHeaders[BuildingHeaders["BUILDING_SIZE"] = 12] = "BUILDING_SIZE";
    BuildingHeaders[BuildingHeaders["DIRECTION"] = 13] = "DIRECTION";
    BuildingHeaders[BuildingHeaders["VEHICLE_PARKING_TYPE"] = 14] = "VEHICLE_PARKING_TYPE";
    BuildingHeaders[BuildingHeaders["VEHICLE_PARKING_NUMBER"] = 15] = "VEHICLE_PARKING_NUMBER";
    BuildingHeaders[BuildingHeaders["WATER_SUPPLY"] = 16] = "WATER_SUPPLY";
    BuildingHeaders[BuildingHeaders["ROAD_NEARBY"] = 17] = "ROAD_NEARBY";
    BuildingHeaders[BuildingHeaders["WIDTH"] = 18] = "WIDTH";
    BuildingHeaders[BuildingHeaders["BUILDING_AGE"] = 19] = "BUILDING_AGE";
    BuildingHeaders[BuildingHeaders["MEMO"] = 20] = "MEMO";
    BuildingHeaders[BuildingHeaders["CONTACT_PERSON"] = 21] = "CONTACT_PERSON";
    BuildingHeaders[BuildingHeaders["PICTURE_LINK"] = 22] = "PICTURE_LINK";
    BuildingHeaders[BuildingHeaders["OBJECT_CREATE_DATE"] = 23] = "OBJECT_CREATE_DATE";
    BuildingHeaders[BuildingHeaders["CONTRACT_DATE_FROM"] = 24] = "CONTRACT_DATE_FROM";
    BuildingHeaders[BuildingHeaders["CONTRACT_DATE_TO"] = 25] = "CONTRACT_DATE_TO";
    BuildingHeaders[BuildingHeaders["OBJECT_UPDATE_DATE"] = 26] = "OBJECT_UPDATE_DATE";
})(BuildingHeaders || (BuildingHeaders = {}));

var LandHeaders;
(function (LandHeaders) {
    LandHeaders[LandHeaders["CREATE_TIME"] = 0] = "CREATE_TIME";
    LandHeaders[LandHeaders["OBJECT_NUMBER"] = 1] = "OBJECT_NUMBER";
    LandHeaders[LandHeaders["OBJECT_NAME"] = 2] = "OBJECT_NAME";
    LandHeaders[LandHeaders["CONTRACT_TYPE"] = 3] = "CONTRACT_TYPE";
    LandHeaders[LandHeaders["LOCATION"] = 4] = "LOCATION";
    LandHeaders[LandHeaders["LAND_PATTERN"] = 5] = "LAND_PATTERN";
    LandHeaders[LandHeaders["LAND_USAGE"] = 6] = "LAND_USAGE";
    LandHeaders[LandHeaders["LAND_TYPE"] = 7] = "LAND_TYPE";
    LandHeaders[LandHeaders["ADDRESS"] = 8] = "ADDRESS";
    LandHeaders[LandHeaders["POSITION"] = 9] = "POSITION";
    LandHeaders[LandHeaders["VALUATION"] = 10] = "VALUATION";
    LandHeaders[LandHeaders["LAND_SIZE"] = 11] = "LAND_SIZE";
    LandHeaders[LandHeaders["NUMBER_OF_OWNER"] = 12] = "NUMBER_OF_OWNER";
    LandHeaders[LandHeaders["ROAD_NEARBY"] = 13] = "ROAD_NEARBY";
    LandHeaders[LandHeaders["DIRECTION"] = 14] = "DIRECTION";
    LandHeaders[LandHeaders["WATER_ELECTRICITY_SUPPLY"] = 15] = "WATER_ELECTRICITY_SUPPLY";
    LandHeaders[LandHeaders["WIDTH"] = 16] = "WIDTH";
    LandHeaders[LandHeaders["DEPTH"] = 17] = "DEPTH";
    LandHeaders[LandHeaders["BUILDING_COVERAGE_RATE"] = 18] = "BUILDING_COVERAGE_RATE";
    LandHeaders[LandHeaders["VOLUME_RATE"] = 19] = "VOLUME_RATE";
    LandHeaders[LandHeaders["MEMO"] = 20] = "MEMO";
    LandHeaders[LandHeaders["CONTACT_PERSON"] = 21] = "CONTACT_PERSON";
    LandHeaders[LandHeaders["PICTURE_LINK"] = 22] = "PICTURE_LINK";
    LandHeaders[LandHeaders["OBJECT_CREATE_DATE"] = 23] = "OBJECT_CREATE_DATE";
    LandHeaders[LandHeaders["CONTRACT_DATE_FROM"] = 24] = "CONTRACT_DATE_FROM";
    LandHeaders[LandHeaders["CONTRACT_DATE_TO"] = 25] = "CONTRACT_DATE_TO";
    LandHeaders[LandHeaders["OBJECT_UPDATE_DATE"] = 26] = "OBJECT_UPDATE_DATE";
})(LandHeaders || (LandHeaders = {}));

function searchObjects(contractType, objectType, objectPattern, objectName, valuationFrom, valuationTo, landSizeFrom, landSizeTo, roadNearby, roomFrom, roomTo, isHasParkingSpace, buildingAgeFrom, buildingAgeTo, direction, objectWidthFrom, objectWidthTo, contactPerson) {

    // Cache Key Generation
    var cache = CacheService.getScriptCache();
    var cacheKey = "search_" + Utilities.base64Encode(JSON.stringify(arguments));
    var cachedResult = cache.get(cacheKey);

    if (cachedResult != null) {
        return cachedResult;
    }

    var listOfSheet = new Array();
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

    if (objectType.toUpperCase() === 'BUILDING' || objectType.toUpperCase() === 'LAND') {
        var currentSheet = spreadsheet.getSheetByName(objectType);
        if (currentSheet) {
            listOfSheet.push(currentSheet);
        }
    } else {
        listOfSheet = spreadsheet.getSheets();
    }

    var filteredValues = new Map();
    var _loop_1 = function (currentSheet) {
        var dataRange = currentSheet.getDataRange();
        var values = dataRange.getValues();
        // var headers = values.shift(); // Unused
        if (values) values.shift(); // Remove headers

        // console.log(objectPattern);
        currentfilteredValues = values
            .map(function (row) {
                var obj = {};
                obj = [values.indexOf(row) + 1, row];
                return obj;
            })
            .filter(function (row) {
                var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m;
                var andConditionList = new Array();
                var orConditionList = new Array();
                var roadNearbyRange = roadNearby.split('|');
                var objectNameKeywordList = objectName.split(' ');
                var sheetName = currentSheet.getName().toUpperCase();
                switch (sheetName) {
                    case 'BUILDING':
                        andConditionList.push(((_a = row[1][BuildingHeaders.CONTRACT_TYPE]) === null || _a === void 0 ? void 0 : _a.toString().indexOf(contractType)) > -1);

                        orConditionList.push(objectNameKeywordList.some(function (keywords) {
                            var _a;
                            return ((_a = row[1][LandHeaders.OBJECT_NUMBER]) === null || _a === void 0 ? void 0 : _a.toString().toLocaleUpperCase().indexOf(keywords.toLocaleUpperCase())) > -1;
                        }));
                        orConditionList.push(objectNameKeywordList.some(function (keywords) {
                            var _a;
                            return ((_a = row[1][BuildingHeaders.OBJECT_NAME]) === null || _a === void 0 ? void 0 : _a.toString().toLocaleUpperCase().indexOf(keywords.toLocaleUpperCase())) > -1;
                        }));
                        orConditionList.push(objectNameKeywordList.some(function (keywords) {
                            var _a;
                            return ((_a = row[1][BuildingHeaders.LOCATION]) === null || _a === void 0 ? void 0 : _a.toString().toLocaleUpperCase().indexOf(keywords.toLocaleUpperCase())) > -1;
                        }));
                        orConditionList.push(objectNameKeywordList.some(function (keywords) {
                            var _a;
                            return ((_a = row[1][BuildingHeaders.ADDRESS]) === null || _a === void 0 ? void 0 : _a.toString().toLocaleUpperCase().indexOf(keywords.toLocaleUpperCase())) > -1;
                        }));
                        var buildingUsageList = (_b = row[1][BuildingHeaders.BUILDING_TYPE]) === null || _b === void 0 ? void 0 : _b.toString().split(',');

                        andConditionList.push(objectPattern.some(function (pattern) {
                            return buildingUsageList.includes(pattern);
                        }));
                        if (roadNearbyRange && roadNearbyRange.length > 1) {
                            andConditionList.push(row[1][BuildingHeaders.ROAD_NEARBY] >= roadNearbyRange[0] && row[1][BuildingHeaders.ROAD_NEARBY] <= roadNearbyRange[1]);
                        }
                        if (valuationFrom > 0) {
                            andConditionList.push(row[1][BuildingHeaders.VALUATION] >= valuationFrom);
                        }
                        if (valuationTo > 0) {
                            andConditionList.push(row[1][BuildingHeaders.VALUATION] <= valuationTo);
                        }
                        if (landSizeFrom > 0) {
                            andConditionList.push(row[1][BuildingHeaders.LAND_SIZE] >= landSizeFrom);
                        }
                        if (landSizeTo > 0) {
                            andConditionList.push(row[1][BuildingHeaders.LAND_SIZE] <= landSizeTo);
                        }
                        var roomOfBuilding = row[1][BuildingHeaders.HOUSE_PATTERN].toString().split('/');
                        if (roomFrom > 0 && roomOfBuilding.length > 0) {
                            andConditionList.push(roomOfBuilding[0] >= roomFrom);
                        }
                        if (roomTo > 0 && roomOfBuilding.length > 0) {
                            andConditionList.push(roomOfBuilding[0] <= roomTo);
                        }

                        if (isHasParkingSpace !== '') {
                            var matchCondition = isHasParkingSpace === '1';
                            // console.log("matchCondition:".concat(matchCondition));
                            // console.log("VEHICLE_PARKING_TYPE:".concat((_c = row[1][BuildingHeaders.VEHICLE_PARKING_TYPE]) === null || _c === void 0 ? void 0 : _c.toString().trim()));
                            // console.log("VEHICLE_PARKING_TYPE:".concat(((_d = row[1][BuildingHeaders.VEHICLE_PARKING_TYPE]) === null || _d === void 0 ? void 0 : _d.toString().trim()) != '沒車位'));
                            andConditionList.push((((_e = row[1][BuildingHeaders.VEHICLE_PARKING_TYPE]) === null || _e === void 0 ? void 0 : _e.toString().trim()) != '沒車位') == matchCondition);
                        }

                        andConditionList.push(((_f = row[1][BuildingHeaders.DIRECTION]) === null || _f === void 0 ? void 0 : _f.toString().indexOf(direction)) > -1);
                        if (objectWidthFrom > 0) {
                            andConditionList.push(row[1][BuildingHeaders.WIDTH] >= objectWidthFrom);
                        }
                        if (objectWidthTo > 0) {
                            andConditionList.push(row[1][BuildingHeaders.WIDTH] <= objectWidthTo);
                        }
                        var buildingAge = ((_g = row[1][BuildingHeaders.BUILDING_AGE]) === null || _g === void 0 ? void 0 : _g.toString().split('/').pop()) || '0';
                        if (buildingAgeFrom > 0) {
                            andConditionList.push(buildingAge >= buildingAgeFrom);
                        }
                        if (buildingAgeTo > 0) {
                            andConditionList.push(buildingAge <= buildingAgeTo);
                        }
                        andConditionList.push(((_h = row[1][BuildingHeaders.CONTACT_PERSON]) === null || _h === void 0 ? void 0 : _h.toString().indexOf(contactPerson)) > -1);
                        break;
                    case 'LAND':
                        andConditionList.push(((_j = row[1][LandHeaders.CONTRACT_TYPE]) === null || _j === void 0 ? void 0 : _j.toString().indexOf(contractType)) > -1);

                        orConditionList.push(objectNameKeywordList.some(function (keywords) {
                            var _a;
                            return ((_a = row[1][LandHeaders.OBJECT_NUMBER]) === null || _a === void 0 ? void 0 : _a.toString().toLocaleUpperCase().indexOf(keywords.toLocaleUpperCase())) > -1;
                        }));
                        orConditionList.push(objectNameKeywordList.some(function (keywords) {
                            var _a;
                            return ((_a = row[1][LandHeaders.OBJECT_NAME]) === null || _a === void 0 ? void 0 : _a.toString().toLocaleUpperCase().indexOf(keywords.toLocaleUpperCase())) > -1;
                        }));
                        orConditionList.push(objectNameKeywordList.some(function (keywords) {
                            var _a;
                            return ((_a = row[1][LandHeaders.LOCATION]) === null || _a === void 0 ? void 0 : _a.toString().toLocaleUpperCase().indexOf(keywords.toLocaleUpperCase())) > -1;
                        }));
                        orConditionList.push(objectNameKeywordList.some(function (keywords) {
                            var _a;
                            return ((_a = row[1][LandHeaders.ADDRESS]) === null || _a === void 0 ? void 0 : _a.toString().toLocaleUpperCase().indexOf(keywords.toLocaleUpperCase())) > -1;
                        }));
                        var landUsageList = (_k = row[1][LandHeaders.LAND_USAGE]) === null || _k === void 0 ? void 0 : _k.toString().split(',');

                        andConditionList.push(objectPattern.some(function (pattern) {
                            return landUsageList.includes(pattern);
                        }));
                        if (roadNearbyRange && roadNearbyRange.length > 1) {
                            andConditionList.push(row[1][LandHeaders.ROAD_NEARBY] >= roadNearbyRange[0] && row[1][LandHeaders.ROAD_NEARBY] <= roadNearbyRange[1]);
                        }
                        if (valuationFrom > 0) {
                            andConditionList.push(row[1][LandHeaders.VALUATION] >= valuationFrom);
                        }
                        if (valuationTo > 0) {
                            andConditionList.push(row[1][LandHeaders.VALUATION] <= valuationTo);
                        }
                        if (landSizeFrom > 0) {
                            andConditionList.push(row[1][LandHeaders.LAND_SIZE] >= landSizeFrom);
                        }
                        if (landSizeTo > 0) {
                            andConditionList.push(row[1][LandHeaders.LAND_SIZE] <= landSizeTo);
                        }

                        andConditionList.push(((_l = row[1][LandHeaders.DIRECTION]) === null || _l === void 0 ? void 0 : _l.toString().indexOf(direction)) > -1);
                        if (objectWidthFrom > 0) {
                            andConditionList.push(row[1][LandHeaders.WIDTH] >= objectWidthFrom);
                        }
                        if (objectWidthTo > 0) {
                            andConditionList.push(row[1][LandHeaders.WIDTH] <= objectWidthTo);
                        }
                        andConditionList.push(((_m = row[1][LandHeaders.CONTACT_PERSON]) === null || _m === void 0 ? void 0 : _m.toString().indexOf(contactPerson)) > -1);
                        break;
                    default:
                }
                // andConditionList.forEach(function (value, index) {
                //     console.log("".concat(sheetName, ":").concat(index, " ").concat(value));
                // });
                var orCondition = orConditionList.some(Boolean);
                // console.log("orCondition:".concat(orCondition));
                return andConditionList.every(Boolean) && orCondition;
            });
        filteredValues = filteredValues.set(currentSheet.getName(), currentfilteredValues);
    };
    var currentfilteredValues;
    for (var _i = 0, listOfSheet_1 = listOfSheet; _i < listOfSheet_1.length; _i++) {
        var currentSheet = listOfSheet_1[_i];
        _loop_1(currentSheet);
    }
    // console.log("filteredValues.size:".concat(filteredValues.size));
    var extractedData = [];
    Array.from(filteredValues).map(function (_a) {
        var key = _a[0], filteredData = _a[1];
        // console.log("key:".concat(key, ", filteredData.length:").concat(filteredData.length));
        var temp = filteredData.map(function (row) {
            var data = {};
            switch (key.toUpperCase()) {
                case 'BUILDING':
                    data = {
                        objectType: key,
                        sequenceNumberInSheet: row[0],
                        objectNumber: row[1][BuildingHeaders.OBJECT_NUMBER],
                        objectName: row[1][BuildingHeaders.OBJECT_NAME],
                        valuation: row[1][BuildingHeaders.VALUATION],
                        landSize: row[1][BuildingHeaders.LAND_SIZE],
                        buildingSize: row[1][BuildingHeaders.BUILDING_SIZE],
                        housePattern: row[1][BuildingHeaders.HOUSE_PATTERN],
                        position: row[1][BuildingHeaders.POSITION],
                        location: row[1][BuildingHeaders.LOCATION],
                        address: row[1][BuildingHeaders.ADDRESS],
                        pictureLink: row[1][BuildingHeaders.PICTURE_LINK]
                    };
                    break;
                case 'LAND':
                    data = {
                        objectType: key,
                        sequenceNumberInSheet: row[0],
                        objectNumber: row[1][LandHeaders.OBJECT_NUMBER],
                        objectName: row[1][LandHeaders.OBJECT_NAME],
                        valuation: row[1][LandHeaders.VALUATION],
                        landSize: row[1][LandHeaders.LAND_SIZE],
                        buildingSize: 0,
                        housePattern: "",
                        position: row[1][LandHeaders.POSITION],
                        location: row[1][LandHeaders.LOCATION],
                        address: row[1][LandHeaders.ADDRESS],
                        pictureLink: row[1][LandHeaders.PICTURE_LINK]
                    };
                    break;
                default:
                    break;
            }
            return data;
        });
        extractedData = extractedData.concat(temp);
    });

    var result = JSON.stringify(extractedData);
    cache.put(cacheKey, result, 7200); // Cache for 2 hours (7200 seconds)
    return result;
}

function getAllPositions() {
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var buildingSheet = spreadsheet.getSheetByName('Building');
    var landSheet = spreadsheet.getSheetByName('Land');
    var buildingDataRange = buildingSheet === null || buildingSheet === void 0 ? void 0 : buildingSheet.getDataRange();
    var landDataRange = landSheet === null || landSheet === void 0 ? void 0 : landSheet.getDataRange();
    var buildingValues = buildingDataRange === null || buildingDataRange === void 0 ? void 0 : buildingDataRange.getValues();
    var landValues = landDataRange === null || landDataRange === void 0 ? void 0 : landDataRange.getValues();
    // var buildingHeaders = buildingValues === null || buildingValues === void 0 ? void 0 : buildingValues.shift(); // Unused
    // var landHeaders = landValues === null || landValues === void 0 ? void 0 : landValues.shift(); // Unused
    if (buildingValues) buildingValues.shift();
    if (landValues) landValues.shift();

    var positions = new Array();
    if (buildingValues) {
        positions = positions.concat(buildingValues
            .filter(function (row) {
                if (!row[BuildingHeaders.POSITION]) {
                    return false;
                }
                var value = row[BuildingHeaders.POSITION].split(' ')[0];
                return value !== '' && value !== null && value !== undefined && isNaN(Number(value));
            })
            .map(function (row) {
                var objectMapData = {
                    objectType: 'building',
                    objectNumber: row[BuildingHeaders.OBJECT_NUMBER],
                    objectName: row[BuildingHeaders.OBJECT_NAME],
                    contractType: row[BuildingHeaders.CONTRACT_TYPE],
                    location: row[BuildingHeaders.LOCATION],
                    position: row[BuildingHeaders.POSITION].split(' ')[0],
                    valuation: row[BuildingHeaders.VALUATION],
                    description: row[BuildingHeaders.OBJECT_NAME],
                    memo: row[BuildingHeaders.MEMO],
                    contractPerson: row[BuildingHeaders.CONTACT_PERSON]
                };
                return objectMapData;
            }));
    }
    if (landValues) {
        positions = positions.concat(landValues
            .filter(function (row) {
                if (!row[LandHeaders.POSITION]) {
                    return false;
                }
                var value = row[LandHeaders.POSITION].split(',');
                return value != null && value.length == 2 && !isNaN(value[0]) && !isNaN(value[1]);
            })
            .map(function (row) {
                var objectMapData = {
                    objectType: 'land',
                    objectNumber: row[LandHeaders.OBJECT_NUMBER],
                    objectName: row[LandHeaders.OBJECT_NAME],
                    contractType: row[LandHeaders.CONTRACT_TYPE],
                    location: row[LandHeaders.LOCATION],
                    position: row[LandHeaders.POSITION],
                    valuation: row[LandHeaders.VALUATION],
                    description: row[LandHeaders.OBJECT_NAME],
                    memo: row[LandHeaders.MEMO],
                    contractPerson: row[LandHeaders.CONTACT_PERSON]
                };
                return objectMapData;
            }));
    }
    return positions;
}

function searchLastNumOfNumberedObjects(objectType) {
    var listOfSheet = new Array();
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (objectType.toUpperCase() === 'BUILDING' || objectType.toUpperCase() === 'LAND') {
        var currentSheet = spreadsheet.getSheetByName(objectType);
        if (currentSheet) {
            listOfSheet.push(currentSheet);
        }
    } else {
        listOfSheet = spreadsheet.getSheets();
    }

    var objectNumberPrefix = '';
    var objectNumberColumn = 0;
    switch (objectType.toUpperCase()) {
        case 'BUILDING':
            objectNumberPrefix = 'A';
            objectNumberColumn = BuildingHeaders.OBJECT_NUMBER;
            break;
        case 'LAND':
            objectNumberPrefix = 'B';
            objectNumberColumn = LandHeaders.OBJECT_NUMBER;
            break;
        default:
    }
    var lastNumberOfObjectNumber = '';
    for (var _i = 0, listOfSheet_2 = listOfSheet; _i < listOfSheet_2.length; _i++) {
        var currentSheet = listOfSheet_2[_i];
        var dataRange = currentSheet.getDataRange();
        var values = dataRange.getValues();
        // var headers = values.shift(); // Unused
        if (values) values.shift();

        var objectNumbers = values.map(function (row) {
            return row[objectNumberColumn];
        });
        lastNumberOfObjectNumber = objectNumbers.reduce(function (prev, current) {
            var isHasPrefix = current.toString().startsWith(objectNumberPrefix);
            var currentNumberPart = Number(current.toString().substring(1));
            var prevNumberPart = Number(prev.toString().substring(1));
            var isCurrentANumber = !isNaN(currentNumberPart);
            var isPrevANumber = !isNaN(prevNumberPart);
            if (!isPrevANumber) {
                prevNumberPart = 0;
            }
            if (isHasPrefix && isCurrentANumber) {
                return currentNumberPart > prevNumberPart ? current : prev;
            }
            return prev;
        });
    }
    return Number(lastNumberOfObjectNumber.toString().substring(1));
}
