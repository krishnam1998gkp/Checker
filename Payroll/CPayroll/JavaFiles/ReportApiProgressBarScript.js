/*
 Created By       : Vishal Chauhan.
 Created Date     : 14 Jan 2025
 Purpose          : To show processbar on report apis.
 ===============================================================================================================================================
 SNo    Date            Name                 Purpose
 1.     29 Jan 2025     Vishal Chauhan       Progress bar added for Year to date salary register Excel report(Rpt_IndEmpSalaryRegister.aspx)
 2.     28 Apr 2025     Vishal Chauhan       YTD Report Builder and YTD Salary Register zip download falg changed
 2.     13 May 2025     Vishal Chauhan       No Record Found Message on UI
 2.     08 Jun 2025     Vishal Chauhan       Bonus Report Process type added
 ===============================================================================================================================================    
*/
var _AppDomain;
var _processType;
var _timeIntervalInExecution = 5000;
var _check;
var flag;
var countProcessedSalary;
var countTotalToProcess;
var countTotalExcelRecords;
var countTotalExcelRowsProcessed;
var step_message;
var ErrMsg;
var TimeSpan = 0;
var TotalRecords = 0;
var ExcelRowsProcessed = 0;
var ExcelRowsDividedBy = 0;
var ExcelRowsIterations = 0;
var hdfile;
var filepath = '';
var filename = '';
var filesize = '';
var RptName = 'YTD Report';
function AjaxJsonPost(postUrl, postData, successCallbackFunction, errorCallbackFunction) {
    $.ajax({
        type: "POST",
        url: postUrl,
        data: postData,
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        success: successCallbackFunction,
        error: errorCallbackFunction
    });
}
function OpenAttendanceProcessBar(_AppMain, _procType, reptName) {
    _AppDomain = _AppMain;
    _processType = _procType;
    RptName = reptName;
    setTimeout(ExcelProcessbarSelection, 200);
}

/*Modified by Vishal Chauhan to add processbar on YTD Sal Register Excel report*/
var ExcelProcessbarSelection = function () {
    $("#CommonProgressBarModelElement").show();
    $("#CommonProgressBarCloseBtn").hide();
    if (_processType == 'YTDSALREG') {
        setTimeout(YTDSalRegExcelProcessSummary, 1000);
    }
    else if (_processType == 'DYNAMICSALREG') {
        setTimeout(DynamicRegisterExcelProcessSummary, 1000);
    }
    else if (_processType == 'SALARYREGISTER') {
        setTimeout(DynamicRegisterExcelProcessSummary, 1000);
    }
    else if (_processType == 'BONUSREPORTCUSTOM') {
        setTimeout(BonusCustomReportExcelProcessSummary, 1000);
    }
    else {
        setTimeout(RptBuilderProcessSummary, 1000);
    }
}
var RptBuilderProcessSummary = function (thr) {
    var urlPost = "/" + _AppDomain + "/CPayroll/ScriptServices/ReportApiProgressBarService.asmx/GetCustomProcessStatus";
    var dataToPost = "{'_process':'" + _processType + "'}";
    statusOfProcess = 'START'
    var successFunction = function (data) {
        var responseObj = '';
        try {
            eval('responseObj =' + data["d"] + ';');
            countProcessedSalary = responseObj.totalProcessed;
            countTotalToProcess = responseObj.totalToProcess;
            statusOfProcess = responseObj.processStatus;
            step_message = responseObj.step_message;
            TimeSpan = responseObj.TimeSpan;
            filepath = responseObj.filepath;
            filename = responseObj.filename;
            filesize = responseObj.filesize;
            TotalRecords = responseObj.TotalRecords;
            ExcelRowsProcessed = responseObj.ExcelRowsProcessed;
            ExcelRowsDividedBy = responseObj.ExcelRowsDividedBy;
            ExcelRowsIterations = responseObj.ExcelRowsIterations;
            ErrMsg = responseObj.ErrMsg;
        }
        catch (e) {
            console.log(e)
            if (thr <= 50) {
                statusOfProcess = "START";
                countProcessedSalary = 2;
                countTotalToProcess = 10;
            }
            else {
                statusOfProcess = "ERROR";
            }
        }

        $("#CommonProgressBarBody").hide();
        if (statusOfProcess.toUpperCase() == "ERROR") {
            $("#CommonProgressBarStatusWrapper").hide();
            $("#CommonProgressBarBody").hide();
            $("#ErrorWrapper .progressbar-outer").addClass("border-error");
            $("#CommonProgressBarTitle").show();
            $("#CommonProgressBarTitle").html(RptName+' failed to generate excel');
            $("#ErrorWrapper").show();
            $("#CommonProgressBarCloseBtn").show();
            if (parseInt(TotalRecords) == 0 && ErrMsg.toLowerCase() == 'no record found!') {
                $("#spnerrmsg").html(ErrMsg);
            }
        }
        else if (statusOfProcess.toUpperCase() == 'DONE') {
            hideModal();
            hdfile = filepath + '~' + filename + '~' + 'N'
            //hdfile = filepath + '~' + filename + '~' + 'REGPROCESSBAR'
            OpenDownloadDiaog(hdfile);
        }
        else {
            $("#CommonProgressBarBody").html('')
            $("#CommonProgressBarBody").show();
            $("#CommonProgressBarCloseBtn").hide();
            $("#CommonProgressBarStatusWrapper .progressbar-outer").show();
            $("#totalProcessed").show();
            $("#CommonProgressBarTitle").show();
            $("#CommonProgressBarTitle").html(RptName + " Process is in Progress...");
            $('#statusWrapper').show();

            if (countProcessedSalary > 0 || (countProcessedSalary >= 0)) {
                var returnPct = UpdateProgressBar(countTotalToProcess, countProcessedSalary, TotalRecords, ExcelRowsProcessed, ExcelRowsDividedBy);
                if (returnPct < 50) {
                    if (TimeSpan > 0 && TimeSpan % 3 == 0) {
                        $("#CommonProgressBarBody").html('Processing speed may vary during this process....')
                    }
                    else if (TimeSpan > 0 && TimeSpan % 4 == 0) {
                        $("#CommonProgressBarBody").html('Process is still in progress. Data may take more time to complete...')
                    }
                    else if (TimeSpan > 0 && TimeSpan % 5 == 0) {
                        $("#CommonProgressBarBody").html('Please ensure a stable connection to avoid interruptions...')
                    }
                    else if (TimeSpan > 0 && TimeSpan % 6 == 0) {
                        $("#CommonProgressBarBody").html('Loading a large dataset. Progress might appear slow but will continue until complete.')
                    }
                    else if (TimeSpan > 0 && TimeSpan % 2 == 0) {
                        $("#CommonProgressBarBody").html('Please do not close this window while the data is being prepared...')
                    }
                    else {
                        $("#CommonProgressBarBody").html('Keep this page open until the process is complete to avoid issues...')
                    }
                    $('#totalProcessed').html(returnPct + "% " + step_message);
                }
                else {
                    if (ExcelRowsProcessed > 0)
                        $("#CommonProgressBarBody").html('Records Processed (' + ExcelRowsProcessed + "/" + TotalRecords + ")")
                    if (returnPct > 50 && returnPct <= 60)
                        $('#totalProcessed').html(returnPct + "% Generating excel file, please wait for the moment...");
                    else if (returnPct > 60 && returnPct <= 70)
                        $('#totalProcessed').html(returnPct + "% Keep this page open until the process is complete to avoid issues...");
                    else if (returnPct > 70 && returnPct <= 80)
                        $('#totalProcessed').html(returnPct + "% Please do not close this window while the excel is being generated...");
                    else if (returnPct > 80 && returnPct <= 88)
                        $('#totalProcessed').html(returnPct + "% Wrapping things up, Thankyou for being with us...");
                    else if (returnPct > 88 && returnPct <= 98)
                        $('#totalProcessed').html(returnPct + "% Almost there, Please hold for a moment...");
                    else if (returnPct > 98 && returnPct < 101) {
                        $("#progressBar").addClass("progressbarsuccess");
                        document.getElementById("progressBar").style.width = '99%';
                        $('#totalProcessed').html("99% Excel file is ready. Zipping and Downloading...");
                        //Show dynamic messages
                        if (TimeSpan > 0 && TimeSpan % 2 == 0) {
                            $("#CommonProgressBarBody").html('Please do not close this window, file is zipping and downloading...')
                        }
                        else {
                            $("#CommonProgressBarBody").html('Due to the large file size, zipping and downloading may take some time...')
                        }
                        $("#CommonProgressBarBody").show();
                    }
                    else
                        $('#totalProcessed').html(returnPct + "% " + step_message);
                }
            }
            setTimeout(RptBuilderProcessSummary, 3000);
        }
        
    };
    function UpdateProgressBar(total, current, totalrecords, rowsprocessed, divideby) {
        var pct = Math.round((current / total) * 100);
        pct = Math.round(pct / 2);
        if (rowsprocessed > 0 && totalrecords > 0 && divideby > 0) {
            pct = pct + Math.round(Math.round((rowsprocessed / totalrecords) * 100) / 2);
        }
        //pct = pct + '%';
        var oldWidth = document.getElementById("progressBar").style.width
        if ((pct + '%') != oldWidth) {
            document.getElementById("progressBar").style.width = pct + '%';
        }
        return pct;
    }
    var errorFunction = function (e, xhr) {
        alert(e.responseText);
    };
    AjaxJsonPost(urlPost, dataToPost, successFunction, errorFunction);
}
var ExcelProcessBarLockedSummary = function () {
    var _ServiceMsg = 0;
    var _ServiceStatus = 0;
    var _IsAbleToStart = 0;
    var _StartedByUserId = 0;
    var _QueryByUserId = 0;
    var _RptName = 0;

    var urlPost = "/" + _AppDomain + "/CPayroll/ScriptServices/ReportApiProgressBarService.asmx/CheckExcelProcessbarAlready";
    var dataToPost = "{'_process':'" + _processType + "'}";
    var successFunction = function (data) {
        var responseObj = '';
        try {
            eval('responseObj =' + data["d"] + ';');
        }
        catch (e) {
            alert('Error in ExcelProcessBarLockedSummary()');
            console.log('217: '+e);
        }
        _ServiceMsg = responseObj.ServiceMsg;
        _ServiceStatus = responseObj.ServiceStatus;
        _IsAbleToStart = responseObj.IsAbleToStart;
        _StartedByUserId = responseObj.StartedByUserId;
        _QueryByUserId = responseObj.QueryByUserId;
        _RptName = responseObj.RptName;
        if (_ServiceStatus == '1' && parseInt(_IsAbleToStart) == '0') {
            let strmsg = '' + _RptName + ' is already processing. Please wait till the completion.';
            $('#lblProcessStatusExcel').html(strmsg);
            $('#divSocialExcel').css('display', '');
            progressBarTimeoutId = setTimeout(ExcelProcessBarLockedSummary, 3000);
        }
        else {
            $('#divSocialExcel').css('display', 'none');
            $('#lblProcessStatusExcel').html('');
            CloseExcelLockedStatusInterval();
        }
    };
    var errorFunction = function (e, xhr) {
        alert(e.responseText);
    };
    AjaxJsonPost(urlPost, dataToPost, successFunction, errorFunction);
}
function RefreshExcelProcessbarLockedStatus(_AppMain, _procType) {
    _AppDomain = _AppMain;
    _processType = _procType;
    progressBarTimeoutId = setTimeout(ExcelProcessBarLockedSummary, 3000);
}
function CloseExcelLockedStatusInterval() {
    try {
        if (progressBarTimeoutId) {
            clearTimeout(progressBarTimeoutId);
            progressBarTimeoutId = null; // Reset the timeout ID
        }
    }
    catch (e) { console.log('Error in stopProgressBar() i.e. ' + e) }
}
function initiateProgressModal(title, message, timeIntervalInExecution = 2000) {
    $("#CommonProgressBarStatusWrapper .progressbar-outer").removeClass("border-error");
    $("#CommonProgressBarTitle").html(title)
    $("#CommonProgressBarBody").html(message)
    $("#CommonProgressBarBody").show();
    $("#CommonProgressBarTitle").hide();
    $("#CommonProgressBarStatusWrapper .progressbar-outer").hide();
    $("#totalProcessed").html("");
    $("#totalProcessed").removeClass("ErrorMessage");
    $("#CommonProgressBarCloseBtn").hide();
    $("#CommonProgressBarModelElement").show();
    $("#ErrorWrapper").hide();
    _timeIntervalInExecution = timeIntervalInExecution;
}
function CloseProgressModal() {
    $("#CommonProgressBarModelElement").hide();
}

/*Added by Vishal Chauhan to add processbar on YTD Sal Register Excel Report*/
function InitiateYTDSalRegExcelProcess(title, message, timeIntervalInExecution = 2000) {
    $("#CommonProgressBarStatusWrapper .progressbar-outer").removeClass("border-error");
    $("#CommonProgressBarTitle").html(title)
    $("#CommonProgressBarBody").html(message)
    $("#CommonProgressBarBody").show();
    $("#CommonProgressBarTitle").hide();
    $("#CommonProgressBarStatusWrapper .progressbar-outer").hide();
    $("#totalProcessedExcel").html("");
    $("#totalProcessedExcel").removeClass("ErrorMessage");
    $("#CommonProgressBarCloseBtn").hide();
    $("#CommonProgressBarModelElement").show();
    _timeIntervalInExecution = timeIntervalInExecution;
}

/* Addded by Vishal Chauhan for YTD Sal Register Excel report where _processType=YTDSALREG */
var YTDSalRegExcelProcessSummary = function (thr) {
    var urlPost = "/" + _AppDomain + "/CPayroll/ScriptServices/ReportApiProgressBarService.asmx/GetCustomProcessStatus";
    var dataToPost = "{'_process':'" + _processType + "'}";
    statusOfProcess = 'START'
    var successFunction = function (data) {
        var responseObj = '';
        try {
            eval('responseObj =' + data["d"] + ';');
            countProcessedSalary = responseObj.totalProcessed;
            countTotalToProcess = responseObj.totalToProcess;
            statusOfProcess = responseObj.processStatus;
            step_message = responseObj.step_message;
            TimeSpan = responseObj.TimeSpan;
            filepath = responseObj.filepath;
            filename = responseObj.filename;
            filesize = responseObj.filesize;
            TotalRecords = responseObj.TotalRecords;
            ExcelRowsProcessed = responseObj.ExcelRowsProcessed;
            ExcelRowsDividedBy = responseObj.ExcelRowsDividedBy;
            ExcelRowsIterations = responseObj.ExcelRowsIterations;
            ErrMsg = responseObj.ErrMsg;
        }
        catch (e) {
            console.log(e)
            statusOfProcess = "ERROR";
        }

        $("#CommonProgressBarBody").hide();
        //Process thrown error!
        if (statusOfProcess.toUpperCase() == "ERROR") {
            $("#CommonProgressBarStatusWrapper").hide();
            $("#CommonProgressBarBody").hide();
            $("#ErrorWrapper .progressbar-outer").addClass("border-error");
            $("#CommonProgressBarTitle").show();
            $("#CommonProgressBarTitle").html(RptName + ' failed to generate excel!');
            $("#ErrorWrapper").show();
            $("#CommonProgressBarCloseBtn").show();
            if (parseInt(TotalRecords) == 0 && ErrMsg.toLowerCase() == 'no record found!') {
                $("#spnerrmsg").html(ErrMsg);
            }
        }
        //Process Completed
        else if (statusOfProcess.toUpperCase() == 'DONE') {
            hideModal();
            //hdfile = filepath + '~' + filename + '~' + 'N'
            hdfile = filepath + '~' + filename + '~' + 'REGPROCESSBAR'
            OpenDownloadDiaog(hdfile);
        }
        //Process in working
        else {
            $("#CommonProgressBarBody").html('')
            $("#CommonProgressBarBody").show();
            $("#CommonProgressBarCloseBtn").hide();
            $("#CommonProgressBarStatusWrapper .progressbar-outer").show();
            $("#totalProcessedExcel").show();
            $("#CommonProgressBarTitle").show();
            $("#CommonProgressBarTitle").html(RptName + " Process is in Progress...");
            $('#statusWrapper').show();
            if (countProcessedSalary > 0 || (countProcessedSalary >= 0 && _processType == 'YTDSALREG')) {
                var returnPct = UpdateProgressBar(countTotalToProcess, countProcessedSalary, TotalRecords, ExcelRowsProcessed, ExcelRowsDividedBy);
                if (returnPct < 50) {
                    if (TimeSpan > 0 && TimeSpan % 3 == 0) {
                        $("#CommonProgressBarBody").html('Processing speed may vary during this process....')
                    }
                    else if (TimeSpan > 0 && TimeSpan % 4 == 0) {
                        $("#CommonProgressBarBody").html('Process is still in progress. Data may take more time to complete...')
                    }
                    else if (TimeSpan > 0 && TimeSpan % 5 == 0) {
                        $("#CommonProgressBarBody").html('Please ensure a stable connection to avoid interruptions...')
                    }
                    else if (TimeSpan > 0 && TimeSpan % 6 == 0) {
                        $("#CommonProgressBarBody").html('Loading a large dataset. Progress might appear slow but will continue until complete.')
                    }
                    else if (TimeSpan > 0 && TimeSpan % 2 == 0) {
                        $("#CommonProgressBarBody").html('Please do not close this window while the data is being prepared...')
                    }
                    else {
                        $("#CommonProgressBarBody").html('Keep this page open until the process is complete to avoid issues...')
                    }
                    //$("#CommonProgressBarBody").show();
                    $('#totalProcessedExcel').html(returnPct + "% " + step_message);
                }
                else {
                    if (ExcelRowsProcessed > 0)
                        $("#CommonProgressBarBody").html('Records Processed (' + ExcelRowsProcessed + "/" + TotalRecords+")")
                    if (returnPct > 50 && returnPct <= 60)
                        $('#totalProcessedExcel').html(returnPct + "% Generating excel file, please wait for the moment...");
                    else if (returnPct > 60 && returnPct <= 70)
                        $('#totalProcessedExcel').html(returnPct + "% Keep this page open until the process is complete to avoid issues...");
                    else if (returnPct > 70 && returnPct <= 80)
                        $('#totalProcessedExcel').html(returnPct + "% Please do not close this window while the excel is being generated...");
                    else if (returnPct > 80 && returnPct <= 88)
                        $('#totalProcessedExcel').html(returnPct + "% Wrapping things up, Thankyou for being with us...");
                    else if (returnPct > 88 && returnPct <= 98)
                        $('#totalProcessedExcel').html(returnPct + "% Almost there, Please hold for a moment...");
                    else if (returnPct > 98 && returnPct < 101) {
                        $("#progressBarExcel").addClass("progressbarsuccess");
                        document.getElementById("progressBar").style.width = '99%';
                        $('#totalProcessedExcel').html("99% Excel file is ready. Zipping and Downloading...");
                        //Show dynamic messages
                        if (TimeSpan > 0 && TimeSpan % 2 == 0) {
                            $("#CommonProgressBarBody").html('Please do not close this window, file is zipping and downloading...')
                        }
                        else {
                            $("#CommonProgressBarBody").html('Due to the large file size, zipping and downloading may take some time...')
                        }
                        $("#CommonProgressBarBody").show();
                    }
                    else
                        $('#totalProcessedExcel').html(returnPct + "% " + step_message);
                }
            }
            setTimeout(YTDSalRegExcelProcessSummary, 3000);
        }

    };
    function UpdateProgressBar(total, current, totalrecords, rowsprocessed, divideby) {
        var pct = Math.round((current / total) * 100);
        pct = Math.round(pct / 2);
        if (rowsprocessed > 0 && totalrecords > 0 && divideby > 0) {
            pct = pct + Math.round(Math.round((rowsprocessed / totalrecords) * 100) / 2);
        }
        //pct = pct + '%';
        var oldWidth = document.getElementById("progressBarExcel").style.width
        if ((pct + '%') != oldWidth) {
            document.getElementById("progressBarExcel").style.width = pct + '%';
        }
        return pct;
    }
    var errorFunction = function (e, xhr) {
        alert(e.responseText);
    };
    AjaxJsonPost(urlPost, dataToPost, successFunction, errorFunction);
}

/*Added by Vishal Chauhan to add processbar on Dynamic Register Excel Report*/
function InitiateDynamicRegisterExcelProcess(title, message, timeIntervalInExecution = 2000) {
    $("#CommonProgressBarStatusWrapper .progressbar-outer").removeClass("border-error");
    $("#CommonProgressBarTitle").html(title)
    $("#CommonProgressBarBody").html(message)
    $("#CommonProgressBarBody").show();
    $("#CommonProgressBarTitle").hide();
    $("#CommonProgressBarStatusWrapper .progressbar-outer").hide();
    $("#totalProcessedExcel").html("");
    $("#totalProcessedExcel").removeClass("ErrorMessage");
    $("#CommonProgressBarCloseBtn").hide();
    $("#CommonProgressBarModelElement").show();
    _timeIntervalInExecution = timeIntervalInExecution;
}
/* Addded by Vishal Chauhan for Dynamic Register Excel report */
var DynamicRegisterExcelProcessSummary = function (thr) {
    var urlPost = "/" + _AppDomain + "/CPayroll/ScriptServices/ReportApiProgressBarService.asmx/GetCustomProcessStatus";
    var dataToPost = "{'_process':'" + _processType + "'}";
    statusOfProcess = 'START'
    var successFunction = function (data) {
        var responseObj = '';
        try {
            eval('responseObj =' + data["d"] + ';');
            countProcessedSalary = responseObj.totalProcessed;
            countTotalToProcess = responseObj.totalToProcess;
            statusOfProcess = responseObj.processStatus;
            step_message = responseObj.step_message;
            TimeSpan = responseObj.TimeSpan;
            filepath = responseObj.filepath;
            filename = responseObj.filename;
            filesize = responseObj.filesize;
            TotalRecords = responseObj.TotalRecords;
            ExcelRowsProcessed = responseObj.ExcelRowsProcessed;
            ExcelRowsDividedBy = responseObj.ExcelRowsDividedBy;
            ExcelRowsIterations = responseObj.ExcelRowsIterations;
            ErrMsg = responseObj.ErrMsg;
        }
        catch (e) {
            console.log(e)
            statusOfProcess = "ERROR";
        }

        $("#CommonProgressBarBody").hide();
        //Process thrown error!
        if (statusOfProcess.toUpperCase() == "ERROR") {
            $("#CommonProgressBarStatusWrapper").hide();
            $("#CommonProgressBarBody").hide();
            $("#ErrorWrapper .progressbar-outer").addClass("border-error");
            $("#CommonProgressBarTitle").show();
            $("#CommonProgressBarTitle").html(RptName + ' failed to generate excel!');
            $("#ErrorWrapper").show();
            $("#CommonProgressBarCloseBtn").show();
            if (parseInt(TotalRecords) == 0 && ErrMsg.toLowerCase() == 'no record found!') {
                $("#spnerrmsg").html(ErrMsg);
            }
        }
        //Process Completed
        else if (statusOfProcess.toUpperCase() == 'DONE') {
            hideModal();
            if (_processType == "DYNAMICSALREG" || _processType == "SALARYREGISTER") {
                hdfile = filepath + '~' + filename + '~' + 'REGPROCESSBAR'
            }
            else {
                hdfile = filepath + '~' + filename + '~' + 'N'
            }
            OpenDownloadDiaog(hdfile);
        }
        //Process in working
        else {
            $("#CommonProgressBarBody").html('')
            $("#CommonProgressBarBody").show();
            $("#CommonProgressBarCloseBtn").hide();
            $("#CommonProgressBarStatusWrapper .progressbar-outer").show();
            $("#totalProcessedExcel").show();
            $("#CommonProgressBarTitle").show();
            $("#CommonProgressBarTitle").html(RptName + " Process is in Progress...");
            $('#statusWrapper').show();
            if (countProcessedSalary > 0 || (countProcessedSalary >= 0)) {
                var returnPct = UpdateProgressBar(countTotalToProcess, countProcessedSalary, TotalRecords, ExcelRowsProcessed, ExcelRowsDividedBy);
                if (returnPct < 50) {
                    if (TimeSpan > 0 && TimeSpan % 3 == 0) {
                        $("#CommonProgressBarBody").html('Processing speed may vary during this process....')
                    }
                    else if (TimeSpan > 0 && TimeSpan % 4 == 0) {
                        $("#CommonProgressBarBody").html('Process is still in progress. Data may take more time to complete...')
                    }
                    else if (TimeSpan > 0 && TimeSpan % 5 == 0) {
                        $("#CommonProgressBarBody").html('Please ensure a stable connection to avoid interruptions...')
                    }
                    else if (TimeSpan > 0 && TimeSpan % 6 == 0) {
                        $("#CommonProgressBarBody").html('Loading a large dataset. Progress might appear slow but will continue until complete.')
                    }
                    else if (TimeSpan > 0 && TimeSpan % 2 == 0) {
                        $("#CommonProgressBarBody").html('Please do not close this window while the data is being prepared...')
                    }
                    else {
                        $("#CommonProgressBarBody").html('Keep this page open until the process is complete to avoid issues...')
                    }
                    //$("#CommonProgressBarBody").show();
                    $('#totalProcessedExcel').html(returnPct + "% " + step_message);
                }
                else {
                    if (ExcelRowsProcessed > 0)
                        $("#CommonProgressBarBody").html('Records Processed (' + ExcelRowsProcessed + "/" + TotalRecords + ")")
                    if (returnPct > 50 && returnPct <= 60)
                        $('#totalProcessedExcel').html(returnPct + "% Generating excel file, please wait for the moment...");
                    else if (returnPct > 60 && returnPct <= 70)
                        $('#totalProcessedExcel').html(returnPct + "% Keep this page open until the process is complete to avoid issues...");
                    else if (returnPct > 70 && returnPct <= 80)
                        $('#totalProcessedExcel').html(returnPct + "% Please do not close this window while the excel is being generated...");
                    else if (returnPct > 80 && returnPct <= 88)
                        $('#totalProcessedExcel').html(returnPct + "% Wrapping things up, Thankyou for being with us...");
                    else if (returnPct > 88 && returnPct <= 98)
                        $('#totalProcessedExcel').html(returnPct + "% Almost there, Please hold for a moment...");
                    else if (returnPct > 98 && returnPct < 101) {
                        $("#progressBarExcel").addClass("progressbarsuccess");
                        document.getElementById("progressBar").style.width = '99%';
                        $('#totalProcessedExcel').html("99% Excel file is ready. Zipping and Downloading...");
                        //Show dynamic messages
                        if (TimeSpan > 0 && TimeSpan % 2 == 0) {
                            $("#CommonProgressBarBody").html('Please do not close this window, file is zipping and downloading...')
                        }
                        else {
                            $("#CommonProgressBarBody").html('Due to the large file size, zipping and downloading may take some time...')
                        }
                        $("#CommonProgressBarBody").show();
                    }
                    else
                        $('#totalProcessedExcel').html(returnPct + "% " + step_message);
                }
            }
            setTimeout(DynamicRegisterExcelProcessSummary, 3000);
        }

    };
    function UpdateProgressBar(total, current, totalrecords, rowsprocessed, divideby) {
        var pct = Math.round((current / total) * 100);
        pct = Math.round(pct / 2);
        if (rowsprocessed > 0 && totalrecords > 0 && divideby > 0) {
            pct = pct + Math.round(Math.round((rowsprocessed / totalrecords) * 100) / 2);
        }
        //pct = pct + '%';
        var oldWidth = document.getElementById("progressBarExcel").style.width
        if ((pct + '%') != oldWidth) {
            document.getElementById("progressBarExcel").style.width = pct + '%';
        }
        return pct;
    }
    var errorFunction = function (e, xhr) {
        alert(e.responseText);
    };
    AjaxJsonPost(urlPost, dataToPost, successFunction, errorFunction);
}
function hideModal() {
    $("#CommonProgressBarModelElement").hide();
}


/*Added by Vishal Chauhan to add processbar on Bonus Custom Report _processType=BONUSREPORTCUSTOM */
function InitiateBonusCustomReportExcelProcess(title, message, timeIntervalInExecution = 2000) {
    $("#CommonProgressBarStatusWrapper .progressbar-outer").removeClass("border-error");
    $("#CommonProgressBarTitle").html(title)
    $("#CommonProgressBarBody").html(message)
    $("#CommonProgressBarBody").show();
    $("#CommonProgressBarTitle").hide();
    $("#CommonProgressBarStatusWrapper .progressbar-outer").hide();
    $("#totalProcessedExcel").html("");
    $("#totalProcessedExcel").removeClass("ErrorMessage");
    $("#CommonProgressBarCloseBtn").hide();
    $("#CommonProgressBarModelElement").show();
    _timeIntervalInExecution = timeIntervalInExecution;
}

/* Addded by Vishal Chauhan for Bonus Custom Report Excel report where _processType=BONUSREPORTCUSTOM */
var BonusCustomReportExcelProcessSummary = function (thr) {
    var urlPost = "/" + _AppDomain + "/CPayroll/ScriptServices/ReportApiProgressBarService.asmx/GetCustomProcessStatus";
    var dataToPost = "{'_process':'" + _processType + "'}";
    statusOfProcess = 'START'
    var successFunction = function (data) {
        var responseObj = '';
        try {
            console.log('BonusCustomReportExcelProcessSummary...');
            eval('responseObj =' + data["d"] + ';');
            countProcessedSalary = responseObj.totalProcessed;
            countTotalToProcess = responseObj.totalToProcess;
            statusOfProcess = responseObj.processStatus;
            step_message = responseObj.step_message;
            TimeSpan = responseObj.TimeSpan;
            filepath = responseObj.filepath;
            filename = responseObj.filename;
            filesize = responseObj.filesize;
            TotalRecords = responseObj.TotalRecords;
            ExcelRowsProcessed = responseObj.ExcelRowsProcessed;
            ExcelRowsDividedBy = responseObj.ExcelRowsDividedBy;
            ExcelRowsIterations = responseObj.ExcelRowsIterations;
            ErrMsg = responseObj.ErrMsg;
        }
        catch (e) {
            console.log(e)
            statusOfProcess = "ERROR";
        }

        $("#CommonProgressBarBody").hide();
        //Process thrown error!
        if (statusOfProcess.toUpperCase() == "ERROR") {
            $("#CommonProgressBarStatusWrapper").hide();
            $("#CommonProgressBarBody").hide();
            $("#ErrorWrapper .progressbar-outer").addClass("border-error");
            $("#CommonProgressBarTitle").show();
            $("#CommonProgressBarTitle").html(RptName + ' failed to generate excel!');
            $("#ErrorWrapper").show();
            $("#CommonProgressBarCloseBtn").show();
            if (parseInt(TotalRecords) == 0 && ErrMsg.toLowerCase() == 'no record found!') {
                $("#spnerrmsg").html(ErrMsg);
            }
        }
        //Process Completed
        else if (statusOfProcess.toUpperCase() == 'DONE') {
            hideModal();
            hdfile = filepath + '~' + filename + '~' + 'N'
            //hdfile = filepath + '~' + filename + '~' + 'REGPROCESSBAR'
            OpenDownloadDiaog(hdfile);
        }
        //Process in working
        else {
            $("#CommonProgressBarBody").html('')
            $("#CommonProgressBarBody").show();
            $("#CommonProgressBarCloseBtn").hide();
            $("#CommonProgressBarStatusWrapper .progressbar-outer").show();
            $("#totalProcessedExcel").show();
            $("#CommonProgressBarTitle").show();
            $("#CommonProgressBarTitle").html(RptName + " Process is in Progress...");
            $('#statusWrapper').show();
            if (countProcessedSalary > 0 || (countProcessedSalary >= 0 && _processType == 'BONUSREPORTCUSTOM')) {
                var returnPct = UpdateProgressBar(countTotalToProcess, countProcessedSalary, TotalRecords, ExcelRowsProcessed, ExcelRowsDividedBy);
                if (returnPct < 50) {
                    if (TimeSpan > 0 && TimeSpan % 3 == 0) {
                        $("#CommonProgressBarBody").html('Processing speed may vary during this process....')
                    }
                    else if (TimeSpan > 0 && TimeSpan % 4 == 0) {
                        $("#CommonProgressBarBody").html('Process is still in progress. Data may take more time to complete...')
                    }
                    else if (TimeSpan > 0 && TimeSpan % 5 == 0) {
                        $("#CommonProgressBarBody").html('Please ensure a stable connection to avoid interruptions...')
                    }
                    else if (TimeSpan > 0 && TimeSpan % 6 == 0) {
                        $("#CommonProgressBarBody").html('Loading a large dataset. Progress might appear slow but will continue until complete.')
                    }
                    else if (TimeSpan > 0 && TimeSpan % 2 == 0) {
                        $("#CommonProgressBarBody").html('Please do not close this window while the data is being prepared...')
                    }
                    else {
                        $("#CommonProgressBarBody").html('Keep this page open until the process is complete to avoid issues...')
                    }
                    //$("#CommonProgressBarBody").show();
                    $('#totalProcessedExcel').html(returnPct + "% " + step_message);
                }
                else {
                    if (ExcelRowsProcessed > 0)
                        $("#CommonProgressBarBody").html('Records Processed (' + ExcelRowsProcessed + "/" + TotalRecords + ")")
                    if (returnPct > 50 && returnPct <= 60)
                        $('#totalProcessedExcel').html(returnPct + "% Generating excel file, please wait for the moment...");
                    else if (returnPct > 60 && returnPct <= 70)
                        $('#totalProcessedExcel').html(returnPct + "% Keep this page open until the process is complete to avoid issues...");
                    else if (returnPct > 70 && returnPct <= 80)
                        $('#totalProcessedExcel').html(returnPct + "% Please do not close this window while the excel is being generated...");
                    else if (returnPct > 80 && returnPct <= 88)
                        $('#totalProcessedExcel').html(returnPct + "% Wrapping things up, Thankyou for being with us...");
                    else if (returnPct > 88 && returnPct <= 98)
                        $('#totalProcessedExcel').html(returnPct + "% Almost there, Please hold for a moment...");
                    else if (returnPct > 98 && returnPct < 101) {
                        $("#progressBarExcel").addClass("progressbarsuccess");
                        document.getElementById("progressBar").style.width = '99%';
                        $('#totalProcessedExcel').html("99% Excel file is ready. Zipping and Downloading...");
                        //Show dynamic messages
                        if (TimeSpan > 0 && TimeSpan % 2 == 0) {
                            $("#CommonProgressBarBody").html('Please do not close this window, file is zipping and downloading...')
                        }
                        else {
                            $("#CommonProgressBarBody").html('Due to the large file size, zipping and downloading may take some time...')
                        }
                        $("#CommonProgressBarBody").show();
                    }
                    else
                        $('#totalProcessedExcel').html(returnPct + "% " + step_message);
                }
            }
            setTimeout(BonusCustomReportExcelProcessSummary, 3000);
        }

    };
    function UpdateProgressBar(total, current, totalrecords, rowsprocessed, divideby) {
        var pct = Math.round((current / total) * 100);
        pct = Math.round(pct / 2);
        if (rowsprocessed > 0 && totalrecords > 0 && divideby > 0) {
            pct = pct + Math.round(Math.round((rowsprocessed / totalrecords) * 100) / 2);
        }
        //pct = pct + '%';
        var oldWidth = document.getElementById("progressBarExcel").style.width
        if ((pct + '%') != oldWidth) {
            document.getElementById("progressBarExcel").style.width = pct + '%';
        }
        return pct;
    }
    var errorFunction = function (e, xhr) {
        alert(e.responseText);
    };
    AjaxJsonPost(urlPost, dataToPost, successFunction, errorFunction);
}