/* 
    This project is for Tableau Excel download
    Product developed by LTIMindtree
*/

'use strict';

window.onload = function () {
    if (!window.top.currentDashboardName) {
        let currentDashboardName = getDashboardNameFromDocumentReferrer();
        window.top.currentDashboardName = currentDashboardName;
        console.log('HREF INITIALZED');
    } else {
        let currentDashboardName = getDashboardNameFromDocumentReferrer();
        if (window.top.currentDashboardName !== currentDashboardName) {
            window.top.currentDashboardName = currentDashboardName;
            console.log('HREF INITIALZED 1');
        }
    }
    //Function Runs when user clicks add button
    document.getElementById("add").onclick = () => {
        
        //TABLEAU EXTENSION API CALL
        tableau.extensions.initializeAsync().then(function () {
            let dashboard = tableau.extensions.dashboardContent.dashboard;

            //Check if dashboard has already been added
            if (window.top?.dashboards) {
                let dashboards = window.top?.dashboards;
                let isDashboardAlreadyPresent = dashboards.some((existingDashboard) => existingDashboard === dashboard.name);
                if (isDashboardAlreadyPresent) {
                    alert(`${dashboard.name} has already been added!`);
                    return;
                }
            }
            //Open Pop-up and Show the loader
            window.top.openReportPopup();
            window.top.Spinner.show();
            window.top.Mask.show();

            //PROMISE
            Promise.all([processDashboard(dashboard)]).then((values) => {
                if (values != 'error') {
                    console.log('Excel Export with Sidebar - V2');
                    // Hide the spinner and mask
                    window.top.Mask.hide();
                    window.top.Spinner.hide();
                } else {
                    console.log('Error while fetching and adding data of dashboard');
                    alert('Error while fetching and adding data of dashboard');
                }
            });
        });
    }
}

// get the dashboard name from the document.referrer
function getDashboardNameFromDocumentReferrer() {
    let referrer = document.referrer;
    let tempUrl = referrer.substring(referrer.indexOf('views/') + 6);
    let currentDashboardName = tempUrl.substring(0, tempUrl.indexOf('/'));
    return currentDashboardName;
}

// get maximum character of each column
function fitToColumn(arrayOfArray) {
    return arrayOfArray[0].map((a, i) => ({ wch: Math.max(...arrayOfArray.map(a2 => a2[i] ? a2[i].toString().length : 0)) }));
}

//find whether object with specific value is present in array of objects
function getIndex(arr, name) {
    const { length } = arr;
    const id = length + 1;
    return arr.findIndex(el => el.fieldName === name);
}

//find whether object with specific value is present in array of objects
function getIndexUsingStartsWith(arr, name) {
    const { length } = arr;
    const id = length + 1;
    return arr.findIndex(el => el.fieldName.startsWith(name));
}

//returns an array of elements with includes given name
function getIncludedArr(arr, name) {
    return arr.filter(x => x.fieldName.includes(name)).map(x => x.fieldName);
}

//returns an array by removing duplicate elements
function removeDuplicates(arr) {
    return arr.filter((item,
        index) => arr.indexOf(item) === index);
}

//Extract the dashboard data, format it as per sheetJS standard, save the data to a window object
function processDashboard(dashboard) {
    //DECLARE REQUIRED OBJECTS FOR STYLEJS
    const DEF_Size14Vert = { font: { sz: 24 }, alignment: { vertical: 'center', horizontal: 'center' } };
    const DEF_FxSz14RgbVert = { font: { name: 'Calibri', sz: 11, color: { rgb: '000000' } }, alignment: { vertical: 'center', horizontal: 'center' } };
    let detailsWorksheet;

    return new Promise(async function (resolve, reject) {
        //Dashboard Worksheet Array
        let arr = dashboard.worksheets;

        //Declare required variables
        let checkCount = 0;
        let dashboardFilters = [];
        let dashboardParameters = [];
        let dashboardName = dashboard.name; //Dashboard Name
        let workbookName = '';
        let sheetName = '';
        let reportHeader = '';
        let reportRefreshTime = '';
        let reportFooter = '';
        let user = '';
        let groupsParams = '';
        let setsParams = '';
        let p = '';
        let f = '';
        let totalRowCount = 0;
        let columnLength = 0;
        let filtersCounter = 0;
        let parametersCounter = 0;
        let groupParametersCounter = 0;
        let setParametersCounter = 0;
        let parameters = [];

        let buildDataArr = [];
        let buildFilterDataArr = [];
        let buildParameterDataArr = [];
        let buildFilterParamsDataArr = [];
        let buildGroupParameterDataArr = [];
        let buildSetParameterDataArr = [];
        let buildHeaderDataArr = [];
        let buildFooterDataArr = [];

        let resultSet = [];

        //Identify the filters and parameters used in dashboard
        for (let object of dashboard.objects) {
            console.log(object);
            if (object.type === 'quick-filter')
                dashboardFilters.push(object.name);
            if (object.type === 'parameter-control')
                dashboardParameters.push(object.name);
        }

        //Identify the worksheets to be extracted from the dashboard
        let worksheetsToBeExtracted = arr.reduce((accumulator, obj) => {
            if (obj.name.includes('Report_Export_Details_D')) {
                return accumulator + 1;
            }
            return accumulator;
        }, 0);

        //Extract the parameter data
        await dashboard.getParametersAsync().then(async function (rawParameters) {
            for (let rawParameter of rawParameters) {
                if (dashboardParameters.includes(rawParameter.name)) {
                    parameters.push({
                        'parameterName': rawParameter.name,
                        'parameterValue': rawParameter.currentValue.formattedValue
                    });
                    parametersCounter++;
                }
            }

            /** Build Parameters Data */
            if (parameters.length > 0) {
                for (let parameter of parameters) {
                    let tt = [];
                    for (let i = 0; i < 2; i++) {
                        if (i == 0) {
                            tt.push({ v: parameter.parameterName, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });

                            buildFilterParamsDataArr.push({ v: parameter.parameterName, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                        } else if (i == 1) {
                            tt.push({ v: parameter.parameterValue, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });

                            buildFilterParamsDataArr.push({ v: parameter.parameterValue, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                        }
                            
                    }
                    buildParameterDataArr.push(tt);
                }
            }
        });

        /**
         * Logic for extracting the worksheets data using Tableau Extension API.
         * First - Identify the metadata sheet. Any sheet starting from 'Report_Export_Details_D' will be a metadata sheet. For Example - 'Report_Export_Details_D1'.
         * Get the actual sheet name from the metadata sheet.
         * Further extract the details of actual sheet.
         * Merge both the metadata sheet data and actual sheet data and form an Object.
         * Set the object to the window object's attributes.
         * Variable 'x' is used for storing the array of Objects which is formed during the above process.
         * Variable 'dashboards' has the list of all dashboards added
         */
        await dashboard.worksheets.forEach(async function (worksheet, key, arr) {
            // If it is a metadata sheet
            if (worksheet.name.includes('Report_Export_Details_D')) {
                detailsWorksheet = worksheet;

                //Extract the details of metadata sheet
                await detailsWorksheet.getSummaryDataAsync().then(async function (mydata) {
                    let dashboardData = mydata.data;           //Metadata sheet data
                    let dashboardColumns = mydata.columns;     //Metadata sheet columns

                    // Workbook name extraction
                    if (getIndexUsingStartsWith(dashboardColumns, 'WB Name') != -1) {
                        workbookName = dashboardData[0][getIndexUsingStartsWith(dashboardColumns, 'WB Name')].value;
                        window.top.workbookName = workbookName;
                    }

                    //Sheet name extraction
                    if (getIndexUsingStartsWith(dashboardColumns, 'Sheet name') != -1) {
                        sheetName = dashboardData[0][getIndexUsingStartsWith(dashboardColumns, 'Sheet name')].value;
                    }

                    //Report Header Extraction
                    if (getIndexUsingStartsWith(dashboardColumns, 'Report Header') != -1) {
                        reportHeader = dashboardData[0][getIndexUsingStartsWith(dashboardColumns, 'Report Header')].value;
                    }

                    // Report Refresh time extraction
                    if (getIndex(dashboardColumns, 'Report Refresh Time') != -1) {
                        reportRefreshTime = dashboardData[0][getIndex(dashboardColumns, 'Report Refresh Time')].value;
                    }

                    //Report Footer extraction
                    if (getIndexUsingStartsWith(dashboardColumns, 'Report Footer') != -1) {
                        reportFooter = dashboardData[0][getIndexUsingStartsWith(dashboardColumns, 'Report Footer')].value;
                    }

                    //User extraction
                    if (getIndex(dashboardColumns, 'User') != -1) {
                        user = dashboardData[0][getIndex(dashboardColumns, 'User')].value;
                    }

                    //Group parameters extraction
                    if (getIndex(dashboardColumns, 'Groups Parameter') != -1) {
                        groupsParams = dashboardData[0][getIndex(dashboardColumns, 'Groups Parameter')].value;
                    }

                    //Sets parameters extraction
                    if (getIndex(dashboardColumns, 'Sets Parameter') != -1) {
                        setsParams = dashboardData[0][getIndex(dashboardColumns, 'Sets Parameter')].value;
                    }

                    //Actual sheet data extraction
                    await dashboard.worksheets.forEach(async function (sheet) {
                        //If the sheet name matches the sheet name present in metadata sheet
                        if (sheet.name === sheetName) {
                            //Varible declarations for data building
                            let builData = {};
                            let tempResult = [];

                            let filters = [];

                            //Filter information extraction used in the sheet level & filtering it out based on whether it has been used in dashboard or not
                            await sheet.getFiltersAsync().then(async function (mydata) {
                                let rawFilters = mydata;

                                if (rawFilters.length > 0) {
                                    for (let rawFilter of rawFilters) {
                                        if (dashboardFilters.includes(rawFilter.fieldName)) {
                                            if (rawFilter.isAllSelected) {
                                                let tempObj = {
                                                    'fieldName': rawFilter.fieldName,
                                                    'filterValues': 'All'
                                                }
                                                filtersCounter++;
                                                filters.push(tempObj);
                                            } else {
                                                let appliedValues = rawFilter.appliedValues || [];
                                                let rawValues = [];
                                                if (appliedValues.length > 0) {
                                                    for (let appliedValue of appliedValues) {
                                                        rawValues.push(appliedValue.formattedValue);
                                                    }
                                                }
                                                let tempObj = {
                                                    'fieldName': rawFilter.fieldName,
                                                    'filterValues': rawValues
                                                }
                                                filtersCounter++;
                                                filters.push(tempObj);
                                            }
                                        }
                                    }

                                    /* Build Filters Data */
                                    if (filters.length > 0) {
                                        for (let filter of filters) {
                                            let tt = [];
                                            for (let i = 0; i < 2; i++) {
                                                if (i == 0) {
                                                    tt.push({ v: filter.fieldName, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                                
                                                    buildFilterParamsDataArr.push({ v: filter.fieldName, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                                }
                                                else if (i == 1) {
                                                    tt.push({ v: filter.filterValues, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });

                                                    buildFilterParamsDataArr.push({ v: filter.filterValues, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } })
                                                }
                                            }
                                            buildFilterDataArr.push(tt);
                                        }
                                    }
                                }
                            });

                            //Use Tableau Extension to extract the actual sheet data
                            await sheet.getSummaryDataAsync().then(async function (d) {
                                if (checkCount == 0) {
                                    let sheetData = d;    //Assign sheet data

                                    checkCount++;

                                    columnLength = sheetData.columns.length;
                                    let columns = sheetData.columns;
                                    let slNoIndex = -1;
                                    let hiddenColumns = [];
                                    let emptyColIndex = -1;

                                    /* Excel data type map */
                                    let definedExcelDataTypeMap = {
                                        'string': 's',
                                        'date': 'd',
                                        'int': 'n',
                                        'float': 'n',
                                        'date-time': 'd'
                                    };

                                    let columnDataTypeMap = {};

                                    /* Check whether column as Measure Names and Measure values field.
                                    If present, find the index */
                                    let measureNamesIndex = -1;
                                    let measureValuesIndex = -1;

                                    for (let i = 0; i < columnLength; i++) {
                                        let colEle = columns[i];
                                        if (colEle.fieldName === 'Measure Names') {
                                            measureNamesIndex = i;
                                        } else if (colEle.fieldName === 'Measure Values') {
                                            measureValuesIndex = i;
                                        }

                                        /* Get Sl_No index */
                                        if (colEle.fieldName === 'AGG(Sl_No)') {
                                            slNoIndex = i;
                                        }

                                        /* Get Index of Hidden Columns */
                                        if (colEle.fieldName.startsWith('Hidden_')) {
                                            hiddenColumns.push(i);
                                        }

                                        /* Get the empty column index */
                                        if (colEle.fieldName.trim() === "' '") {
                                            emptyColIndex = i;
                                        }

                                        /* Get the data type of each column and populate into map */
                                        columnDataTypeMap[i] = colEle.dataType;
                                    }

                                    /* If measure names are present, count how much measure names are present */
                                    let colData = sheetData.data;
                                    let measureNames = [];
                                    let mCount = 1;

                                    if (measureNamesIndex != -1) {
                                        // let mFlag = false;
                                        let mIndex = -1;
                                        for (let i = 0; i < colData.length; i++) {
                                            let arrEle = colData[i];

                                            if (mIndex == -1) {
                                                for (let j = 0; j < arrEle.length; j++) {
                                                    if (measureNamesIndex != j || measureValuesIndex != j) {
                                                        mIndex = j;
                                                        break;
                                                    }
                                                }
                                            }
                                            if (mIndex != -1) {
                                                if (colData[i]?.[mIndex].value === colData[i + 1]?.[mIndex].value) {
                                                    // mCount++;
                                                    measureNames.push(colData[i][measureNamesIndex].formattedValue);
                                                    measureNames.push(colData[i + 1][measureNamesIndex].formattedValue);
                                                } else {
                                                    break;
                                                }
                                            }

                                        }

                                        if (measureNames.length == 0) {
                                            measureNames.push(colData[0][measureNamesIndex].formattedValue);
                                        }
                                    }

                                    measureNames = removeDuplicates(measureNames);  //Measure Names
                                    mCount = measureNames.length;                   //Count of Measure Names

                                    //Array declaration to hold cell values of excel
                                    let tt = [];
                                    let rr = [];
                                    let empt = [];

                                    //Logic to find out actual column length
                                    let actualColumnLength = columnLength;
                                    columnLength = measureNames.length > 0 ? columnLength - 2 + mCount : columnLength;
                                    columnLength = slNoIndex == -1 ? columnLength : columnLength - 1;
                                    columnLength = emptyColIndex == -1 ? columnLength : columnLength - 1;
                                    columnLength = hiddenColumns.length === 0 ? columnLength : columnLength - hiddenColumns.length;

                                    //Build Report Header and Report Executed by row
                                    for (let i = 0; i < columnLength; i++) {
                                        if (i == 0) {
                                            tt.push({ v: reportHeader, t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '538DD5' } }, font: { sz: 14, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { horizontal: 'left', vertical: 'center' } } });
                                        } else {
                                            tt.push({ v: ' ', t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '538DD5' } }, font: { sz: 22, name: 'Calibri', color: { rgb: 'f1f1f1' } } } });
                                        }
                                        if (i == 0) {
                                            rr.push({ v: `Report executed by ${user} ${reportRefreshTime}`, t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '538DD5' } }, font: { sz: 11, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { horizontal: 'left' } } });
                                        } else {
                                            rr.push({ v: ' ', t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '538DD5' } }, font: { sz: 11, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { horizontal: 'right' } } });
                                        }
                                        empt.push(" ");
                                    }

                                    buildHeaderDataArr.push(tt);
                                    buildHeaderDataArr.push(empt);
                                    buildHeaderDataArr.push(rr);

                                    //Group paramters Data Building
                                    if (groupsParams != '') {
                                        tt = [];
                                        for (let i = 0; i < columnLength; i++) {
                                            if (i == columnLength - 2) {
                                                tt.push({ v: groupsParams, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                            } else {
                                                tt.push({ v: '', t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                            }
                                        }
                                        groupParametersCounter++;
                                        buildGroupParameterDataArr.push(tt);
                                    }

                                    //Set parameters Data Building
                                    if (setsParams != '') {
                                        tt = [];
                                        for (let i = 0; i < columnLength; i++) {
                                            if (i == columnLength - 2) {
                                                tt.push({ v: setsParams, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                            } else {
                                                tt.push({ v: '', t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                            }
                                        }
                                        setParametersCounter++;
                                        buildSetParameterDataArr.push(tt);
                                    }

                                    //Column Header Data Building
                                    tt = [];

                                    //If Columns involve measures then apply the below logic
                                    if (measureNames.length > 0) {
                                        for (let i = 0; i < actualColumnLength; i++) {
                                            if ((i != measureNamesIndex) && (i != measureValuesIndex) && (i != slNoIndex) && (i != emptyColIndex) && !(hiddenColumns.includes(i))) {
                                                let colEle = columns[i];

                                                tt.push({ v: ((colEle.fieldName.startsWith('SUM(') || colEle.fieldName.startsWith('AGG(') || colEle.fieldName.startsWith('ATTR(')) && colEle.fieldName.endsWith(')')) ? colEle.fieldName.substring(4, colEle.fieldName.length - 1) : (colEle.fieldName.startsWith('ATTR(') && colEle.fieldName.endsWith(')')) ? colEle.fieldName.substring(5, colEle.fieldName.length - 1) : colEle.fieldName, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'left', wrapText: '1' } } });
                                            }
                                        }
                                        for (let i = 0; i < measureNames.length; i++) {
                                            tt.push({ v: measureNames[i], t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'left', wrapText: '1' } } });
                                        }
                                    } else { // If columns does not contain the measure names apply the below logic
                                        for (let i = 0; i < actualColumnLength; i++) {
                                            let colEle = columns[i];

                                            if ((i != slNoIndex) && (i !== emptyColIndex) && !(hiddenColumns.includes(i))) {
                                                tt.push({ v: ((colEle.fieldName.startsWith('SUM(') || colEle.fieldName.startsWith('AGG(')) && colEle.fieldName.endsWith(')')) ? colEle.fieldName.substring(4, colEle.fieldName.length - 1) : (colEle.fieldName.startsWith('ATTR(') && colEle.fieldName.endsWith(')')) ? colEle.fieldName.substring(5, colEle.fieldName.length - 1) : colEle.fieldName, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'left', wrapText: '1' } } });
                                            }
                                        }
                                    }

                                    tempResult.push(tt);

                                    //Column data building
                                    //If columns contain the measure names apply the below logic
                                    if (measureNames.length > 0) {
                                        let lCount = mCount;
                                        let tempDict = {};
                                        let tempArr = [];
                                        for (let i = 0; i < colData.length; i++) {
                                            let arrEle = colData[i];

                                            if (lCount != 0) {
                                                for (let j = 0; j < arrEle.length; j++) {
                                                    if ((j != measureNamesIndex) && (j != measureValuesIndex) && (j != slNoIndex) && !(hiddenColumns.includes(j)) && (j != emptyColIndex) && (lCount == mCount)) {
                                                        tempArr.push({ v: arrEle[j].value == '%null%' ? '' : columnDataTypeMap[j] === 'date' || columnDataTypeMap[j] === 'date-time' ? arrEle[j].formattedValue.substring(0, arrEle[j].formattedValue.indexOf(" ") === -1 ? arrEle[j].formattedValue.length : arrEle[j].formattedValue.indexOf(" ")) : arrEle[j].value, t: arrEle[j].value == '%null%' ? 's' : definedExcelDataTypeMap?.[columnDataTypeMap[j]] ? definedExcelDataTypeMap?.[columnDataTypeMap[j]] : isNaN(arrEle[j].value) ? 's' : 'n', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, alignment: isNaN(arrEle[j].value) ? { horizontal: 'left' } : { horizontal: 'right' } }, z: arrEle[j].value == '%null%' ? '' : (definedExcelDataTypeMap?.[columnDataTypeMap[j]] && definedExcelDataTypeMap?.[columnDataTypeMap[j]] === 'd') ? 'dd-mm-yyy' : '#,##0.00##############;-#,##0.00##############;0;General' });
                                                    }
                                                }
                                                tempDict[arrEle[measureNamesIndex].formattedValue] = arrEle[measureValuesIndex].value;
                                                lCount--;
                                            }

                                            if (lCount == 0) {
                                                for (let j = 0; j < measureNames.length; j++) {
                                                    let tempData = tempDict[measureNames[j]];
                                                    tempArr.push({ v: tempData == '%null%' ? '' : tempData, t: isNaN(tempData) ? 's' : 'n', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, alignment: isNaN(tempData) ? { horizontal: 'left' } : { horizontal: 'right' } }, z: tempData == '%null%' ? '' : isNaN(tempData) ? '' : '#,##0.00##############;-#,##0.00##############;0;General' });
                                                }

                                                tempResult.push(tempArr);

                                                totalRowCount++;
                                                tempArr = [];
                                                tempDict = {};
                                                lCount = mCount;
                                            }

                                        }
                                    } else {  //If columns does not contain the measure names apply the below logic
                                        for (let i = 0; i < colData.length; i++) {
                                            let arrEle = colData[i];
                                            let tempArr = [];
                                            for (let j = 0; j < arrEle.length; j++) {
                                                if ((j != slNoIndex) && (j != emptyColIndex) && !(hiddenColumns.includes(j))) {
                                                    tempArr.push({ v: arrEle[j].value == '%null%' ? '' : arrEle[j].value, t: arrEle[j].value == '%null%' ? 's' : definedExcelDataTypeMap?.[columnDataTypeMap[j]] ? definedExcelDataTypeMap?.[columnDataTypeMap[j]] : isNaN(arrEle[j].value) ? 's' : 'n', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, alignment: isNaN(arrEle[j].value) ? { horizontal: 'left' } : { horizontal: 'right' } }, z: arrEle[j].value == '%null%' ? '' : (definedExcelDataTypeMap?.[columnDataTypeMap[j]] && definedExcelDataTypeMap?.[columnDataTypeMap[j]] === 'd') ? 'dd-mm-yyyy' : '#,##0.00##############;-#,##0.00##############;0;General' });
                                                }
                                            }

                                            tempResult.push(tempArr);

                                            totalRowCount++;
                                        }
                                    }

                                    builData.columnLength = columnLength;
                                    builData.sheetData = tempResult;

                                    buildDataArr.push(builData);
                                } else {
                                    let sheetData = d;

                                    checkCount++;
                                    
                                    let sheetColumnLength = sheetData.columns.length;
                                    let columns = sheetData.columns;
                                    let slNoIndex = -1;
                                    let hiddenColumns = [];
                                    let emptyColIndex = -1;

                                    /* Excel data type map */
                                    let definedExcelDataTypeMap = {
                                        'string': 's',
                                        'date': 'd',
                                        'int': 'n',
                                        'float': 'n',
                                        'date-time': 'd'
                                    };

                                    let columnDataTypeMap = {};

                                    /* Check whether column as Measure Names and Measure values field.
                                    If present, find the index */
                                    let measureNamesIndex = -1;
                                    let measureValuesIndex = -1;

                                    for (let i = 0; i < sheetColumnLength; i++) {
                                        let colEle = columns[i];
                                        if (colEle.fieldName === 'Measure Names') {
                                            measureNamesIndex = i;
                                        } else if (colEle.fieldName === 'Measure Values') {
                                            measureValuesIndex = i;
                                        }

                                        /* Get Sl_No index */
                                        if (colEle.fieldName === 'AGG(Sl_No)') {
                                            slNoIndex = i;
                                        }

                                        /* Get Index of Hidden Columns */
                                        if (colEle.fieldName.startsWith('Hidden_')) {
                                            hiddenColumns.push(i);
                                        }

                                        /* Get the empty column index */
                                        if (colEle.fieldName.trim() === "' '") {
                                            emptyColIndex = i;
                                        }

                                        /* Get the data type of each column and populate into map */
                                        columnDataTypeMap[i] = colEle.dataType;
                                    }

                                    /* If measure names are present, count how much measure names are present */
                                    let colData = sheetData.data;
                                    let measureNames = [];
                                    let mCount = 1;

                                    if (measureNamesIndex != -1) {
                                        // let mFlag = false;
                                        let mIndex = -1;
                                        for (let i = 0; i < colData.length; i++) {
                                            let arrEle = colData[i];

                                            if (mIndex == -1) {
                                                for (let j = 0; j < arrEle.length; j++) {
                                                    if (measureNamesIndex != j || measureValuesIndex != j) {
                                                        mIndex = j;
                                                        break;
                                                    }
                                                }
                                            }
                                            if (mIndex != -1) {
                                                if (colData[i]?.[mIndex].value === colData[i + 1]?.[mIndex].value) {
                                                    // mCount++;
                                                    measureNames.push(colData[i][measureNamesIndex].formattedValue);
                                                    measureNames.push(colData[i + 1][measureNamesIndex].formattedValue);
                                                } else {
                                                    break;
                                                }
                                            }

                                        }

                                        if (measureNames.length == 0) {
                                            measureNames.push(colData[0][measureNamesIndex].formattedValue);
                                        }
                                    }

                                    measureNames = removeDuplicates(measureNames);   //Measure Names
                                    mCount = measureNames.length;                    //Count of measure names

                                    //Declare and initialize empty array
                                    let empt = [];
                                    let tt = [];

                                    //Logic for finding actual column length
                                    let actualColumnLength = sheetColumnLength;
                                    sheetColumnLength = measureNames.length > 0 ? sheetColumnLength - 2 + mCount : sheetColumnLength;
                                    sheetColumnLength = slNoIndex == -1 ? sheetColumnLength : sheetColumnLength - 1;
                                    sheetColumnLength = hiddenColumns.length === 0 ? sheetColumnLength : sheetColumnLength - hiddenColumns.length;
                                    sheetColumnLength = emptyColIndex == -1 ? sheetColumnLength : sheetColumnLength - 1;

                                    //Column header data building
                                    if (measureNames.length > 0) {  //If measure names available
                                        for (let i = 0; i < actualColumnLength; i++) {
                                            if ((i != measureNamesIndex) && (i != measureValuesIndex) && (i != slNoIndex) && (i != emptyColIndex) && !(hiddenColumns.includes(i))) {
                                                let colEle = columns[i];
                                                tt.push({ v: ((colEle.fieldName.startsWith('SUM(') || colEle.fieldName.startsWith('AGG(') || colEle.fieldName.startsWith('ATTR(')) && colEle.fieldName.endsWith(')')) ? colEle.fieldName.substring(4, colEle.fieldName.length - 1) : (colEle.fieldName.startsWith('ATTR(') && colEle.fieldName.endsWith(')')) ? colEle.fieldName.substring(5, colEle.fieldName.length - 1) : colEle.fieldName, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'left', wrapText: '1' } } });
                                            }
                                        }
                                        for (let i = 0; i < measureNames.length; i++) {
                                            tt.push({ v: measureNames[i], t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'left', wrapText: '1' } } });
                                        }
                                    } else {  //If measure names not available
                                        for (let i = 0; i < actualColumnLength; i++) {
                                            let colEle = columns[i];

                                            if ((i !== slNoIndex) && (i !== emptyColIndex) && !(hiddenColumns.includes(i))) {
                                                tt.push({ v: ((colEle.fieldName.startsWith('SUM(') || colEle.fieldName.startsWith('AGG(')) && colEle.fieldName.endsWith(')')) ? colEle.fieldName.substring(4, colEle.fieldName.length - 1) : (colEle.fieldName.startsWith('ATTR(') && colEle.fieldName.endsWith(')')) ? colEle.fieldName.substring(5, colEle.fieldName.length - 1) : colEle.fieldName, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'left', wrapText: '1' } } });
                                            }
                                        }
                                    }

                                    tempResult.push(tt);

                                    //Column data building
                                    if (measureNames.length > 0) {  //If measure names available
                                        let lCount = mCount;
                                        let tempDict = {};
                                        let tempArr = [];
                                        for (let i = 0; i < colData.length; i++) {
                                            let arrEle = colData[i];

                                            if (lCount != 0) {
                                                for (let j = 0; j < arrEle.length; j++) {
                                                    if ((j != measureNamesIndex) && (j != measureValuesIndex) && (j != slNoIndex) && (j != emptyColIndex) && (lCount == mCount) && !(hiddenColumns.includes(j))) {
                                                        tempArr.push({ v: arrEle[j].value == '%null%' ? '' : columnDataTypeMap[j] === 'date' || columnDataTypeMap[j] === 'date-time' ? arrEle[j].formattedValue.substring(0, arrEle[j].formattedValue.indexOf(" ") === -1 ? arrEle[j].formattedValue.length : arrEle[j].formattedValue.indexOf(" ")) : arrEle[j].value, t: arrEle[j].value == '%null%' ? 's' : definedExcelDataTypeMap?.[columnDataTypeMap[j]] ? definedExcelDataTypeMap?.[columnDataTypeMap[j]] : isNaN(arrEle[j].value) ? 's' : 'n', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, alignment: isNaN(arrEle[j].value) ? { horizontal: 'left' } : { horizontal: 'right' } }, z: arrEle[j].value == '%null%' ? '' : (definedExcelDataTypeMap?.[columnDataTypeMap[j]] && definedExcelDataTypeMap?.[columnDataTypeMap[j]] === 'd') ? 'dd-mm-yyyy' : '#,##0.00##############;-#,##0.00##############;0;General' });
                                                    }
                                                }
                                                tempDict[arrEle[measureNamesIndex].formattedValue] = arrEle[measureValuesIndex].value;
                                                lCount--;
                                            }

                                            if (lCount == 0) {
                                                for (let j = 0; j < measureNames.length; j++) {
                                                    let tempData = tempDict[measureNames[j]];
                                                    tempArr.push({ v: tempData == '%null%' ? '' : tempData, t: isNaN(tempData) ? 's' : 'n', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, alignment: isNaN(tempData) ? { horizontal: 'left' } : { horizontal: 'right' } }, z: tempData == '%null%' ? '' : isNaN(tempData) ? '' : '#,##0.00##############;-#,##0.00##############;0;General' });
                                                }

                                                tempResult.push(tempArr);

                                                totalRowCount++;
                                                tempArr = [];
                                                tempDict = {};
                                                lCount = mCount;
                                            }

                                        }
                                    } else {  //If measure names not available
                                        for (let i = 0; i < colData.length; i++) {
                                            let arrEle = colData[i];
                                            let tempArr = [];
                                            for (let j = 0; j < arrEle.length; j++) {
                                                if ((j != slNoIndex) && (j != emptyColIndex) && !(hiddenColumns.includes(j))) {
                                                    tempArr.push({ v: arrEle[j].value == '%null%' ? '' : arrEle[j].value, t: arrEle[j].value == '%null%' ? 's' : definedExcelDataTypeMap?.[columnDataTypeMap[j]] ? definedExcelDataTypeMap?.[columnDataTypeMap[j]] : isNaN(arrEle[j].value) ? 's' : 'n', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, alignment: isNaN(arrEle[j].value) ? { horizontal: 'left' } : { horizontal: 'right' } }, z: arrEle[j].value == '%null%' ? '' : (definedExcelDataTypeMap?.[columnDataTypeMap[j]] && definedExcelDataTypeMap?.[columnDataTypeMap[j]] === 'd') ? 'dd-mm-yyyy' : '#,##0.00##############;-#,##0.00##############;0;General' });
                                                }
                                            }

                                            tempResult.push(tempArr);

                                            totalRowCount++;
                                        }
                                    }

                                    builData.columnLength = sheetColumnLength;
                                    builData.sheetData = tempResult;

                                    buildDataArr.push(builData);
                                }

                                if (checkCount == worksheetsToBeExtracted) {
                                    let tt = [];
                                    let empt = [];

                                    //Building Footer
                                    for (let i = 0; i < columnLength; i++) {
                                        if (i == 0) {
                                            tt.push({ v: reportFooter, t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '404040' } }, font: { sz: 11, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { vertical: 'center', horizontal: 'left' } } });
                                        } else {
                                            tt.push({ v: ' ', t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '404040' } }, font: { sz: 11, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { vertical: 'bottom', horizontal: 'center' } } });
                                        }
                                        empt.push(" ");
                                    }

                                    buildFooterDataArr.push(tt);

                                    buildDataArr.sort((a, b) => +b.columnLength - +a.columnLength);  //Sort the data array according to the columnLength

                                    let finalColumnLength = buildDataArr[0].columnLength;  //Find the report with highest number of columns

                                    //Add dummy cell to rows which has less cell to format properly
                                    if (buildHeaderDataArr[0].length < finalColumnLength) {
                                        let increaseIndex = buildHeaderDataArr[0].length;

                                        for (let i = increaseIndex; i < finalColumnLength; i++) {
                                            buildHeaderDataArr[0].push({ v: ' ', t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '538DD5' } }, font: { sz: 22, name: 'Calibri', color: { rgb: 'f1f1f1' } } } });
                                            buildHeaderDataArr[1].push({ v: ' ', t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '538DD5' } }, font: { sz: 22, name: 'Calibri', color: { rgb: 'f1f1f1' } } } })
                                            buildHeaderDataArr[2].push({ v: ' ', t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '538DD5' } }, font: { sz: 22, name: 'Calibri', color: { rgb: 'f1f1f1' } } } });
                                            buildFooterDataArr[0].push({ v: ' ', t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '404040' } }, font: { sz: 11, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { vertical: 'bottom', horizontal: 'center' } } });
                                        }
                                    }

                                    //Create final output
                                    resultSet.push([...[...buildHeaderDataArr]]);  //Header part

                                    let trackColumnLength = 0;
                                    let filterAndParamsResultSet = [];
                                    let filterParamsCounter = 0;

                                    for (let z = 0; z < buildFilterParamsDataArr.length; z+=2) {
                                        if ((trackColumnLength + 2) > finalColumnLength) {
                                            resultSet.push([...[filterAndParamsResultSet]]);
                                            filterParamsCounter++;
                                            filterAndParamsResultSet = [];
                                            trackColumnLength = 0;
                                            filterAndParamsResultSet.push(buildFilterParamsDataArr[z]);
                                            filterAndParamsResultSet.push(buildFilterParamsDataArr[z + 1]);
                                            trackColumnLength = trackColumnLength + 2;
                                        } else {
                                            filterAndParamsResultSet.push(buildFilterParamsDataArr[z]);
                                            filterAndParamsResultSet.push(buildFilterParamsDataArr[z + 1]);
                                            trackColumnLength = trackColumnLength + 2;
                                        }
                                    }

                                    if (filterAndParamsResultSet.length > 0) {
                                        resultSet.push([...[filterAndParamsResultSet]]);
                                        filterParamsCounter++;
                                    }

                                    //If filer or parameter present add them
                                    // buildFilterDataArr.length > 0 ? resultSet.push([...[...buildFilterDataArr]]) : null;
                                    // buildParameterDataArr.length > 0 ? resultSet.push([...[...buildParameterDataArr]]) : null;

                                    //If group parameters or set parameters present add them
                                    buildGroupParameterDataArr.length > 0 ? resultSet.push([...[...buildGroupParameterDataArr]]) : null;
                                    buildSetParameterDataArr.length > 0 ? resultSet.push([...[...buildSetParameterDataArr]]) : null;

                                    //Add empty row for better visual
                                    resultSet.push([['', '']]);
                                    resultSet.push([['', '']]);

                                    let alternateWorksheetCount = 0;  //All other that main sheet

                                    //Main data building
                                    buildDataArr.forEach((data, index) => {
                                        if (index === 0) {
                                            resultSet.push([...[...data.sheetData]]);
                                            resultSet.push([['', '']]);
                                        } else {
                                            data.sheetData.length > 2 ? resultSet.push([...[...data.sheetData]]) : resultSet.push([...[...data.sheetData.slice(1)]]);
                                            resultSet.push([['', '']]);
                                            alternateWorksheetCount++;
                                            data.sheetData.length > 2 && alternateWorksheetCount++;
                                        }
                                    });

                                    //If there is only one worksheet an extra row needs to be inserted
                                    if (worksheetsToBeExtracted === 1) {
                                        resultSet.push([['', '']]);
                                    }

                                    resultSet.push([...[...buildFooterDataArr]]); //Add footer data

                                    let finalResult = resultSet.flatMap(x => x);  //Convert to 1D Array for better mapping

                                    //CREATE WORKSHEET(S) AND ADD IT TO EXCEL FILE
                                    let worksheet = XLSX.utils.aoa_to_sheet(finalResult);

                                    //Calculate the footer starting position for merging of cell
                                    let rowFooterMergeStart = worksheetsToBeExtracted === 1 ? 
                                    8 + totalRowCount : 7 + alternateWorksheetCount + totalRowCount;

                                    // rowFooterMergeStart = filtersCounter !== 0 ? rowFooterMergeStart + filtersCounter : rowFooterMergeStart;
                                    rowFooterMergeStart = filterParamsCounter !== 0 ? rowFooterMergeStart + filterParamsCounter : rowFooterMergeStart;
                                    rowFooterMergeStart = groupParametersCounter !== 0 ? rowFooterMergeStart + groupParametersCounter : rowFooterMergeStart;
                                    rowFooterMergeStart = setParametersCounter !== 0 ? rowFooterMergeStart + setParametersCounter : rowFooterMergeStart;

                                    worksheet['!cols'] = fitToColumn(finalResult);   //Fit columns width
                                    worksheet['!rows'] = [{ 'hpt': 40 }];            //Set row height

                                    //Start cell merge
                                    worksheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 1, c: finalColumnLength - 1 } },
                                    { s: { r: rowFooterMergeStart, c: 0 }, e: { r: rowFooterMergeStart + 2, c: finalColumnLength - 1 } }
                                    ];

                                    worksheet["!merges"].push({ s: { r: 2, c: 0 }, e: { r: 2, c: finalColumnLength - 1 } });
                                    worksheet["!merges"] = groupsParams != '' ? [...worksheet["!merges"], { s: { r: 5, c: finalColumnLength - 2 }, e: { r: 5, c: finalColumnLength - 1 } }] : worksheet["!merges"];
                                    worksheet["!merges"] = setsParams != '' ? [...worksheet["!merges"], { s: { r: 6, c: finalColumnLength - 2 }, e: { r: 6, c: finalColumnLength - 1 } }] : worksheet["!merges"];

                                    //Creation of final object
                                    let obj = {
                                        name: dashboardName,
                                        worksheet: worksheet
                                    }

                                    //Add dashboard name to window object dashboard variable
                                    if (!window.top.dashboards) {
                                        window.top.dashboards = [dashboardName];
                                    } else {
                                        window.top.dashboards = [dashboardName, ...window.top.dashboards];
                                    }

                                    //Add Dashboard to side pane
                                    window.top.addDashboardToSidePane();

                                    //Add actual object to the window object x variable
                                    if (!window.top.x) {
                                        window.top.x = [obj];
                                    } else {
                                        window.top.x = [obj, ...window.top.x];
                                    }

                                    resolve();

                                }
                            });
                        }
                    });
                });
            }
        });

    });
}
