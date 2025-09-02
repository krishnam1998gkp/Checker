<%@ Page Language="vb" AutoEventWireup="false" Inherits="Payroll.Rpt_IndEmpSalaryRegister"
    CodeFile="Rpt_IndEmpSalaryRegister.aspx.vb" %>
    <%@ Register Src="~/CPayroll/UsearchGroup.ascx" TagName="Group" TagPrefix="mc1" %>
        <%@ Register Src="~/CPayroll/UsearchWithMultipleEmpCodeproxy.ascx" TagName="ucsearch" TagPrefix="uc1" %>
            <%@ Register Src="~/CPayroll/CompanyMenu.ascx" TagName="AdminMenu" TagPrefix="uc1" %>
                <html>

                <head runat="server">
                    <title>
                        <%=Session("Title")%>
                    </title>
                    <meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
                    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
                    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
                    <meta content="JavaScript" name="vs_defaultClientScript" />
                    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
                    <link href="Css/hospCSS.css" rel="stylesheet" />
                    <link href="Css/CommonProgressBarStyle.css?v=1.0.1" type="text/css" rel="stylesheet" />
                    <link href="Css/jquery-ui-1.8.1.custom.css" type="text/css" rel="stylesheet" />
                    <script language="javascript" src="JavaFiles/jquery-1.4.2.min.js" type="text/javascript"></script>
                    <script language="javascript" src="JavaFiles/jquery-ui-1.8.1.custom.min.js"
                        type="text/javascript"></script>
                    <script language="javascript" src="JavaFiles/jquery.bgiframe.js" type="text/javascript"></script>
                    <script language="javascript" src="JavaFiles/Script_utill.js" type="text/javascript"></script>
                    <script language="javascript" src="JavaFiles/JS_CommonUtill.js" type="text/javascript"></script>
                    <script language="javascript" src="JavaFiles/JSBalloon.js" type="text/javascript"></script>
                    <script language="javascript" src="JavaFiles/CommonDownload.js?v=1.0.1"
                        type="text/javascript"></script>
                    <script language="javascript" src="JavaFiles/ProcessBarYTD.js?v=1.0.3"
                        type="text/javascript"></script>
                    <script language="javascript" src="JavaFiles/ReportApiProgressBarScript.js?v=1.1.2"
                        type="text/javascript"></script>
                    <style>
                        #divSocial img {
                            position: relative;
                            border: none;
                        }

                        #divSocialExcel img {
                            position: relative;
                            border: none;
                        }

                        .UserStatusMsg {
                            font-family: arial;
                            color: #006699;
                            font-size: 14px !important;
                            list-style: none;
                            font-weight: bold;
                        }

                        .UserPrevRunMsg {
                            font-family: arial;
                            color: #006699;
                            font-size: 10px !important;
                            list-style: none;
                            font-weight: bold;
                        }
                    </style>
                    <script language="javascript" type="text/javascript">
                        $(document).ready(function () {
                            $("#divSocial img").hover(
                                function () {
                                    $(this).animate({ left: "-20" });
                                },
                                function () {
                                    $(this).animate({ left: "0" });
                                });
                            $("#divSocialExcel img").hover(
                                function () {
                                    $(this).animate({ left: "-20" });
                                },
                                function () {
                                    $(this).animate({ left: "0" });
                                });
                        });
                        function validatePswd() {
                            if ($('#ddlreportIn').val() == 'P') {
                                $('#btnPublishedPDF').css('display', '');
                                $("#btnPriview").attr('value', 'Publish YTD Slip');
                                $('#trpswd').css('display', '');
                            }
                            else {
                                $('#btnPublishedPDF').css('display', 'none');
                                $("#btnPriview").attr('value', 'Preview');
                                $('#trpswd').css('display', 'none');
                            }
                        }
                        //Added by Vishal Chauhan for process bar initial message
                        function StartProcessbar(reptName) {
                            CloseDialog();
                            console.log('reptName: ' + reptName);
                            if (reptName == '') {
                                reptName = 'YTD Salary Register Report';
                            }
                            InitiateYTDSalRegExcelProcess(reptName + " Process", "Preparing data to generate " + reptName + ", This will only take a moment. Please do not close the window.", 500);
                        }
                        //Sameer
                        function datecheck(_det, cnt) {
                            if ($('#txtpagevalue').val() == "") {
                                alert("No. of page should not be left blank !");
                                return false;
                            }
                            var Dep = $('#hdf_USearchMulti_DdlDept').val();
                            var Desig = $('#hdf_USearchMulti_ddldesignation').val();
                            var Grad = $('#hdf_USearchMulti_ddlGrade').val();
                            var Lable = $('#hdf_USearchMulti_ddllevel').val();
                            var CC = $('#hdf_USearchMulti_DdlCostCenter').val();
                            var Loc = $('#hdf_USearchMulti_ddllocation').val();
                            var unit = $('#hdf_USearchMulti_ddlunit').val();
                            var SalBase = $('#USearchMulti_ddlSalBasis').val();
                            var EmpCode = $('#hidEmpCode').val()
                            var Ename = $('#USearchMulti_txtEmpname').val();
                            var ES = 'F';
                            var Month;
                            var Month1;
                            var Reptype;
                            if ($("#rblextrapolate").find(':checked').val() == 'Y') {
                                Reptype = 'Y'
                                Month = $('#ddlASmonth').val();
                                Month1 = $('#ddlAEmonth').val();
                                //Added by Nisha(07 Feb 2013) Fin Year Change
                                var monthyear = $('#ddlASmonth :selected').text()
                                var arr = new Array;
                                arr = (monthyear.split(" - "));
                                var Year = arr[1];

                                var monthyear = $('#ddlAEmonth :selected').text()
                                var arr = new Array;
                                arr = (monthyear.split(" - "));
                                var Year1 = arr[1];

                            }
                            else {
                                Reptype = 'N'
                                Month = $('#ddlmonth').val();
                                Month1 = $('#ddlmonth1').val();
                                //Added by Nisha(07 Feb 2013) Fin Year Change
                                var monthyear = $('#ddlmonth :selected').text()
                                var arr = new Array;
                                arr = (monthyear.split(" - "));
                                var Year = arr[1];

                                var monthyear = $('#ddlmonth1 :selected').text()
                                var arr = new Array;
                                arr = (monthyear.split(" - "));
                                var Year1 = arr[1];
                            }
                            var PageNo = $('#txtpagevalue').val();
                            var height = window.screen.availHeight;
                            var width = window.screen.availWidth;
                            var EmpSts = $('#USearchMulti_ddlEmp').val();
                            var _checkPdf = 'H';
                            var password = '0';

                            var reptformate = $("#rbtnreportformate :checked").val();

                            if (reptformate == 'V') {
                                _checkPdf = $('#ddlreportIn').val();
                            }
                            if ($('#ddlreportIn').val() == 'P') {
                                password = $('#ddlEmpPass').val();
                            }
                            var ArrSmeMnth = $('#chkArrSmeMnth').val(); //$("#chkArrSmeMnth :checked").val();

                            if ($('#chkArrSmeMnth').is(":checked")) {
                                ArrSmeMnth = 'Y';
                            }
                            else {
                                ArrSmeMnth = 'N';
                            }

                            qryP = Dep + '~' + Desig + '~' + Grad + '~' + Lable + '~' + CC + '~' + Loc + '~' + unit + '~' + SalBase + '~' + EmpCode + '~' + Month + '~' + Month1 + '~' + Reptype + '~' + PageNo + '~' + ES + '~' + Ename + '~' + EmpSts + '~' + _det + '~' + Year + '~' + Year1 + '~' + _checkPdf + '~' + ArrSmeMnth + '~' + password + '~' + cnt;

                            if (reptformate == 'H') {
                                window.open('reports/preIndSalaryRegister.aspx?id=' + qryP, '', ' height=' + height + ' width=' + width + ' left=3' + ' menubar=1' + ' top=0' + ' scrollbars=1');
                            }
                            else {
                                if (_checkPdf == 'P') {
                                    ShowYTDDetails();
                                }
                                else {
                                    window.open('reports/preindsalaryregisterVertical.aspx?id=' + qryP, '', ' height=' + height + ' width=' + width + ' left=3' + ' menubar=1' + ' top=0' + ' scrollbars=1');
                                }

                            }
                        }
                        //added by Vishal Chauhan to show msg on progress before intialization
                        function LoadProcessDialog() {
                            if ($("#rbtnreportformate :checked").val() == 'H' && $('#USearchMulti_txtEmpCode').val() == '') {
                                alert('Please enter employee codes for the YTD Payslip in HTML format..');
                                return false;
                            }
                            if ($('#ddlreportIn').val() == 'H' && $('#USearchMulti_txtEmpCode').val() == '') {
                                alert('Please enter employee codes for the YTD Payslip in HTML format..');
                                return false;
                            }
                            if ($('#ddlreportIn').val() == 'P' && $('#rbtnreportformate').find(':checked').val() == "V") {
                                if ($('#<%= lblProcessStatus.ClientID %>').text() != '') {
                                    let str = 'Hi ' + $('#hdnusername').val() + ',\nYTD Payslip PDF is publishing, please wait till the completion.\nFor the latest update click on Refresh button on top right corner of Show/Hide search parameters.';
                                    alert(str);
                                    return false;
                                }
                                LoadSlipProgressYTD();
                            }

                        }
                        function CloseProcessbarDialog() {
                            UnLoadSlipProgressYTD();
                        }
                        function ShowYTDDetails() {
                            var qryPv = $("#HidPreVal").val();
                            var AppPath = $("#HidAppPath").val();
                            var path = $("#HidPath").val();
                            var page = 'reports/preIndSalaryRegisterVerticalNew.aspx?id=' + qryPv;
                            OpenSlipProgressYTD(AppPath, page, path, "YTDSLIPV");
                        }
                        //End

                        function CheckChange() {
                            if ($('#rbtnreportformate').find(':checked').val() == "V") {
                                $('#trRepIn').css('display', '');

                                if ($('#ddlreportIn').val() == 'P') {
                                    $('#btnPublishedPDF').css('display', '');
                                    $("#btnPriview").attr('value', 'Publish YTD Slip');
                                    if ($('#btnEdit').val() == 'EDIT') {
                                        $('#trpswd').css('display', 'none');
                                    } else {
                                        $('#trpswd').css('display', '');
                                    }
                                }
                                else {
                                    $("#btnPriview").attr('value', 'Preview');
                                    $('#btnPublishedPDF').css('display', 'none');
                                    $('#trpswd').css('display', 'none');
                                }
                            }
                            else {
                                $("#btnPriview").attr('value', 'Preview');
                                $('#btnPublishedPDF').css('display', 'none');
                                $('#trRepIn').css('display', 'none');
                                $('#trpswd').css('display', 'none');
                            }
                        }
                        //Start: Rohtas Singh on 02 Apr 2015
                        function ShowBalloon(obj, width, height, msg) {
                            var b2 = new JSBalloon({ width: width, height: String(height) });
                            b2.Show({
                                title: String("Instruction"), message: String(msg),
                                anchor: obj, top: document.body.scrollTop, left: document.body.scrollLeft, icon: 'Help'
                            });
                        }
                        //Start: Vishal Chauhan on 14 OCt 2024
                        function checkvals() {
                            var value = $('#txtpagevalue').val();
                            if (isNaN(value) || value < 1 || value > 99) {
                                $('#txtpagevalue').val("");
                                $('#txtpagevalue').focus();
                            }
                        }
                        function BtnHidenShow() {
                            if ($('#btnEdit').val() == 'EDIT') {
                                $('#btnEdit').val('DONE');
                                $('#TrSave').css('display', 'none');
                                $('#TrEmpDetails').css('display', '');
                                $('#TrFormat').css('display', '');
                                $('#TrExtrapolate').css('display', '');
                                $('#Trpublish').css('display', '');
                                $('#TrSvConfig').css('display', '');
                                CheckChange();
                            }
                            else {
                                $('#btnEdit').val('EDIT');
                                $('#TrSave').css('display', '');
                                $('#TrEmpDetails').css('display', 'none');
                                $('#TrFormat').css('display', 'none');
                                $('#TrExtrapolate').css('display', 'none');
                                $('#Trpublish').css('display', 'none');
                                $('#TrSvConfig').css('display', 'none');
                                $('#trpswd').css('display', 'none');

                            }
                            return false;
                        }

                        //Modified by Debargha on 17-July-2024 for 'Please Wait' clickbait on YTD Salary Register Report
                        function btnExcel_Click() {
                            if ($('#hdnCSV').val().toUpperCase() == 'Y') {
                                return confirm('Downloading the Excel file may take longer. If you prefer a faster option, instead you can use the EXPORT TO CSV button.');
                            }
                            return PleaseWaitWithDailog();
                        }
                        function OnClickDownloadCSV() {
                            PleaseWaitWithDailog();
                        }

                        var loopInstance = null;
                        var isLoopInstanceActive = false;
                        var _ReportType = "";
                        var _jobId = "";
                        var _appDir = '<%=GetAppPath() %>'
                        function handleLoopingOfDownloadStatus() {
                            if (!isLoopInstanceActive) {
                                return;
                            }
                            $.ajax({
                                url: "/" + _appDir + "/CPayroll/ScriptServices/ReportStatusService.asmx/CheckAndDownloadFile",
                                data: JSON.stringify({ RepType: _ReportType, JobId: _jobId }),
                                type: 'POST',
                                contentType: "application/json; charset=utf-8",
                                dataType: 'json',
                                async: false,
                                //  data: passingParams,
                                success: function (res) {
                                    //console.log(res);
                                    var response = JSON.parse(res.d)
                                    if (response.Status == 'Completed') {
                                        $('#hdfile').val(response.FilePath + '~' + response.FileName + '~' + 'N');
                                        ShowDownload($('#hdfile').val())
                                        CloseDialog();
                                        //console.log("Response ended");
                                        if (isLoopInstanceActive) {
                                            isLoopInstanceActive = false;
                                            clearInterval(loopInstance);
                                        }
                                    }
                                    else if (response.Status == 'NR') {
                                        CloseDialog();
                                        ShowErrorDialog('', 'R', '', 'No Record Found according to selection criteria!', '');
                                        if (isLoopInstanceActive) {
                                            isLoopInstanceActive = false;
                                            clearInterval(loopInstance);
                                        }
                                    }
                                    else if (response.Status == 'Failed') {
                                        CloseDialog();
                                        ShowErrorDialog('', 'R', '', 'API Execution Failed!', '');
                                        if (isLoopInstanceActive) {
                                            isLoopInstanceActive = false;
                                            clearInterval(loopInstance);
                                        }
                                    }

                                },
                                error: function (xhr, status, errorThrown) {
                                    CloseDialog();
                                    ShowErrorDialog('', 'R', '', errorThrown, '');
                                    if (isLoopInstanceActive) {
                                        isLoopInstanceActive = false;
                                        clearInterval(loopInstance);
                                    }
                                }
                            });
                        }
                        function ChkRptDownloadStatus(RepType, JobId) {
                            PleaseWaitWithDailog();
                            _ReportType = RepType;
                            _jobId = JobId;
                            isLoopInstanceActive = true;
                            setTimeout(() => {
                                handleLoopingOfDownloadStatus();
                            }, 4000);
                            loopInstance = setInterval(() => {
                                handleLoopingOfDownloadStatus();
                            }, 10000);
                        }
                    </script>
                </head>

                <body bottommargin="0" bgcolor="#f7fcff" leftmargin="0" topmargin="0" rightmargin="0" marginheight="0"
                    marginwidth="0">
                    <div id="dlg" style="font-size: 1.1em; padding: 20px; display: none" title="Progress">
                        <div id="statusMessage" style="font-weight: bold; padding: 5px 0px">
                        </div>
                        <div style="display: none;">
                            <iframe id="Slipframe" style="border: 0px; border-spacing: 0px; border-collapse: collapse;"
                                marginheight="0" marginwidth="0" frameborder="0"></iframe>
                        </div>
                        <div id="statusWrapper" style="display: block; line-height: 30px; padding-left: 10px;">
                            <div id="progressBar" style="width: 500px;">
                            </div>
                            <div>
                                <span id="totalProcessed" style="font-weight: bold"></span><span>&nbsp;of </span>
                                <span id="totalToProcess" style="font-weight: bold"></span><span>&nbsp;processed.
                                </span>
                            </div>
                            <div>
                                <span>Estimated time left: </span><span id="estimatedTimeLeft"
                                    style="font-weight: bold">
                                </span>
                            </div>
                        </div>
                        <div id="summaryWrapper" style="display: none; line-height: 30px; padding-left: 10px;">
                            <div>
                                <span>Total employees whose slip was published: </span><span id="totalProcessedSalary"
                                    style="font-weight: bold"></span>
                            </div>
                            <div>
                                <span>Total employees whose slip not published due to some error: </span><span
                                    id="totalUnprocessedSalary" style="font-weight: bold"></span>
                            </div>
                            <div style="display: none;">
                                <span>Total employees whose slip already published : </span><span
                                    id="TotAlreadyProcssed" style="font-weight: bold"></span>
                            </div>
                            <div>
                                <span>Processed in: </span><span id="totalTimeTaken" style="font-weight: bold"></span>
                            </div>
                            <div id="summaryError"
                                style="display: none; line-height: 15px; height: 80px; overflow: auto">
                                <span id="errorSummary" style="color: Red;"></span>
                            </div>
                        </div>
                        <%--<div id="summary" style="display: none; line-height: 30px; padding-left: 10px;">
                            <div><span id="spansummary" runat="Server" style="display:none"></span></div>--%>
                    </div>
                    <div id="CommonProgressBarModelElement" class="alt-module-true" style="display: none;">
                        <div class="alt-modal dialog-500">
                            <div class="alt-modal-title">
                                <span id="CommonProgressBarTitle">Progress</span>
                            </div>
                            <div class="alt-modal-body" id="CommonProgressBarBody" style="display: block;">
                            </div>
                            <div class="alt-modal-body" id="CommonProgressBarStatusWrapper" style="display: block;">
                                <div class="progressbar-outer">
                                    <div id="progressBarExcel" class="progressbar">
                                    </div>
                                </div>
                                <div class="clearfix">
                                    <span class="FL"><span class="blue-color" id="totalProcessedExcel"></span><span
                                            class="blue-color" id="totalToProcessExcel"></span></span><span class="FR">
                                        <span class="blue-color" id="estimatedTimeLeftExcel"></span></span>
                                </div>
                            </div>
                            <div class="alt-modal-body" id="ErrorWrapper" style="display: none;">
                                <div class="progressbar-outer">
                                    <div class="progressbarerr" style="width: 100%">
                                    </div>
                                </div>
                                <p class="p-lines">
                                    <%--<span id="spnerrmsg">Error in Processing the YTD Report, Please connect with App
                                        Support</span>--%>
                                        <span id="spnerrmsg">Error in Processing the YTD Report, Please refresh the page
                                            and try again. If the issue persists, connect with App Support.</span>
                                </p>
                            </div>
                            <div id="CommonProgressBarFooter" class="alt-modal-body">
                                <button id="CommonProgressBarCloseBtn" class="Btn"
                                    style="height: 20px; width: 55px; font-family: Verdana,Arial,sans-serif"
                                    onclick="hideModal();">Close</button>
                            </div>

                        </div>
                    </div>
                    <form id="Form1" method="post" runat="server">
                        <uc1:AdminMenu ID="AdminMenu1" runat="server" />
                        <table cellspacing="0" cellpadding="0" width="100%" border="0">
                            <tr>
                                <td>
                                    &nbsp;
                                    <asp:ScriptManager ID="scr1" runat="server">
                                    </asp:ScriptManager>
                                </td>
                            </tr>
                            <tr>
                                <td width="100%">
                                    <table cellspacing="0" cellpadding="0" width="97%" align="center" border="0">
                                        <tr>
                                            <td align="center">
                                                <fieldset class="fieldset">
                                                    <legend class="legend">Salary Register Year to Date</legend>
                                                    <asp:UpdatePanel ID="upStatus" runat="server">
                                                        <ContentTemplate>
                                                            <table cellspacing="0" cellpadding="0" width="98%"
                                                                align="center" border="0" id="divSocial" runat="server"
                                                                style="display: none;">
                                                                <tr>
                                                                    <td align="right">
                                                                        <img src="Images/icon3.png" id="Img1"
                                                                            style="vertical-align: middle;" />
                                                                        <%--<asp:Label ID="lblProcessStatus"
                                                                            runat="server" Text="" ForeColor="Red"
                                                                            Font-Bold="true"
                                                                            style="font-size: 12px;font-family: Verdana, 'Times New Roman', 'Courier New'">
                                                                            </asp:Label>--%>
                                                                            <asp:Label ID="lblProcessStatus"
                                                                                runat="server" Text=""
                                                                                CssClass="UserStatusMsg"></asp:Label>
                                                                            <asp:ImageButton ID="imgpdfprocess"
                                                                                runat="server" Visible="false"
                                                                                Style="vertical-align: middle;"
                                                                                ToolTip="Refresh" Width="25px"
                                                                                AlternateText="Refresh"
                                                                                ImageUrl="img/refreshtds.gif">
                                                                            </asp:ImageButton>
                                                                            <asp:HiddenField ID="hdnusername"
                                                                                runat="server" Value="User" />
                                                                            <asp:HiddenField ID="IsShowlnkIcon"
                                                                                runat="server" Value="N" />
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </ContentTemplate>
                                                    </asp:UpdatePanel>
                                                    <table cellspacing="0" cellpadding="0" width="98%" align="center"
                                                        border="0" id="divSocialExcel" runat="server" visible="false">
                                                        <tr>
                                                            <td colspan="3">
                                                                <table cellspacing="0" cellpadding="0" width="98%"
                                                                    align="center" border="0">
                                                                    <tr>
                                                                        <td align="right">
                                                                            <img src="Images/icon3.png" id="ImgExcel"
                                                                                style="vertical-align: middle;" />
                                                                            <asp:Label ID="lblProcessStatusExcel"
                                                                                runat="server" CssClass="UserStatusMsg"
                                                                                Text=""></asp:Label>
                                                                            <%--<asp:Button ID="btnProgressbarExcel"
                                                                                runat="server"
                                                                                Text="click here to check progress"
                                                                                CssClass="CustomBtn" Width="175px"
                                                                                CausesValidation="False"
                                                                                OnClientClick="return PleaseWaitWithDailog();" />--%>
                                                                            <asp:ImageButton ID="btnProgressbarExcel"
                                                                                runat="server"
                                                                                Style="vertical-align: middle;"
                                                                                ToolTip="click here to check progress"
                                                                                Width="25px" AlternateText="Refresh"
                                                                                ImageUrl="img/refreshtds.gif"
                                                                                OnClientClick="return PleaseWaitWithDailog();">
                                                                            </asp:ImageButton>
                                                                        </td>
                                                                    </tr>
                                                                </table>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <table cellspacing="0" cellpadding="0" width="98%" align="center"
                                                        border="0">
                                                        <tr>
                                                            <td>
                                                                <mc1:Group ID="USearchMulti" runat="server" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>
                                                                <asp:UpdatePanel ID="upd1" runat="server">
                                                                    <ContentTemplate>
                                                                        <table cellspacing="0" cellpadding="0"
                                                                            width="100%" align="center" border="0">
                                                                            <tr>
                                                                                <td class="TdCaption" colspan="3"
                                                                                    valign="middle">
                                                                                    <table border="0" cellpadding="0"
                                                                                        cellspacing="0" width="100%">
                                                                                        <tr class="trupbtn">
                                                                                            <td colspan="3"
                                                                                                class="trbtnup"></td>
                                                                                        </tr>
                                                                                        <tr>
                                                                                            <td class="TdCaption"
                                                                                                width="12%"
                                                                                                valign="middle"></td>
                                                                                            <td class="TDcolon"
                                                                                                valign="middle"></td>
                                                                                            <td colspan="3">
                                                                                                <asp:Button ID="btnEdit"
                                                                                                    runat="server"
                                                                                                    class="btn"
                                                                                                    Text="EDIT"
                                                                                                    OnClientClick="javascript:return BtnHidenShow();" />
                                                                                            </td>
                                                                                        </tr>
                                                                                        <tr>
                                                                                            <td class="TdCaption"
                                                                                                valign="middle"
                                                                                                width="12%">From date
                                                                                            </td>
                                                                                            <td class="TDcolon"
                                                                                                valign="middle">:</td>
                                                                                            <td>
                                                                                                <table cellspacing="0"
                                                                                                    cellpadding="0"
                                                                                                    width="100%"
                                                                                                    border="0">
                                                                                                    <tr id="PnlSMonth"
                                                                                                        runat="server"
                                                                                                        style="display: none">
                                                                                                        <td width="36.5%"
                                                                                                            valign="middle">
                                                                                                            <asp:DropDownList
                                                                                                                ID="ddlmonth"
                                                                                                                runat="server"
                                                                                                                CssClass="dropdownlist">
                                                                                                            </asp:DropDownList>
                                                                                                        </td>
                                                                                                        <td class="TdCaption"
                                                                                                            width="14%"
                                                                                                            valign="middle">
                                                                                                            To date</td>
                                                                                                        <td class="TDcolon"
                                                                                                            valign="middle">
                                                                                                            :</td>
                                                                                                        <td colspan="2"
                                                                                                            valign="middle">
                                                                                                            <asp:DropDownList
                                                                                                                ID="ddlmonth1"
                                                                                                                runat="server"
                                                                                                                CssClass="dropdownlist">
                                                                                                            </asp:DropDownList>
                                                                                                        </td>
                                                                                                    </tr>
                                                                                                    <tr id="PnlAfromdatatodate"
                                                                                                        runat="server"
                                                                                                        style="display: none">
                                                                                                        <td valign="middle"
                                                                                                            width="36.5%">
                                                                                                            <asp:DropDownList
                                                                                                                ID="ddlASmonth"
                                                                                                                runat="server"
                                                                                                                CssClass="dropdownlist">
                                                                                                            </asp:DropDownList>
                                                                                                        </td>
                                                                                                        <td class="TdCaption"
                                                                                                            valign="middle"
                                                                                                            width="14%">
                                                                                                            To date</td>
                                                                                                        <td class="TDcolon"
                                                                                                            valign="middle">
                                                                                                            :</td>
                                                                                                        <td colspan="2"
                                                                                                            valign="middle">
                                                                                                            <asp:DropDownList
                                                                                                                ID="ddlAEmonth"
                                                                                                                runat="server"
                                                                                                                CssClass="dropdownlist"
                                                                                                                DataMember="dropdownlist">
                                                                                                            </asp:DropDownList>
                                                                                                        </td>
                                                                                                    </tr>
                                                                                                </table>
                                                                                            </td>
                                                                                        </tr>
                                                                                        <tr id="TrFormat" runat="server"
                                                                                            style="display: none;">
                                                                                            <td class="TdCaption"
                                                                                                width="12%"
                                                                                                valign="middle">View
                                                                                                Format</td>
                                                                                            <td class="TDcolon"
                                                                                                valign="middle">:</td>
                                                                                            <td colspan="3">
                                                                                                <asp:RadioButtonList
                                                                                                    ID="rbtnreportformate"
                                                                                                    runat="server"
                                                                                                    CssClass="HospRadiobuttonlist"
                                                                                                    RepeatDirection="Horizontal"
                                                                                                    onchange="CheckChange();">
                                                                                                    <asp:ListItem
                                                                                                        Selected="True"
                                                                                                        Value="H">
                                                                                                        Horizontal
                                                                                                    </asp:ListItem>
                                                                                                    <asp:ListItem
                                                                                                        Value="V">
                                                                                                        Vertical
                                                                                                    </asp:ListItem>
                                                                                                </asp:RadioButtonList>
                                                                                            </td>
                                                                                        </tr>
                                                                                        <tr id="TrExtrapolate"
                                                                                            runat="server"
                                                                                            style="display: none;">
                                                                                            <td class="TdCaption"
                                                                                                width="12%"
                                                                                                valign="middle">
                                                                                                Extrapolate</td>
                                                                                            <td class="TDcolon"
                                                                                                valign="middle">:</td>
                                                                                            <td>
                                                                                                <table cellspacing="0"
                                                                                                    cellpadding="0"
                                                                                                    width="100%"
                                                                                                    border="0">
                                                                                                    <tr>
                                                                                                        <td width="10%"
                                                                                                            valign="middle">
                                                                                                            <asp:RadioButtonList
                                                                                                                ID="rblextrapolate"
                                                                                                                runat="server"
                                                                                                                AutoPostBack="True"
                                                                                                                CssClass="HospRadiobuttonlist"
                                                                                                                RepeatDirection="Horizontal"
                                                                                                                onchange="PleaseWaitWithDailog();">
                                                                                                                <asp:ListItem
                                                                                                                    Selected="True"
                                                                                                                    Value="N">
                                                                                                                    No
                                                                                                                </asp:ListItem>
                                                                                                                <asp:ListItem
                                                                                                                    Value="Y">
                                                                                                                    Yes
                                                                                                                </asp:ListItem>
                                                                                                            </asp:RadioButtonList>
                                                                                                        </td>
                                                                                                        <td width="26.5%"
                                                                                                            valign="middle">
                                                                                                            <img id="Instruction"
                                                                                                                style="cursor: hand"
                                                                                                                alt=""
                                                                                                                onclick="ShowBalloon(this,600,70,'<strong>1. Horizontal: </strong>Display months row wise<br /><strong>2</strong>. <strong>Vertical: </strong>Display months column wise<br /><strong>3. Extrapolate</strong>: Projected salary')"
                                                                                                                src="ImagesCA/new4-0985.gif" />
                                                                                                        </td>
                                                                                                        <td class="TdCaption"
                                                                                                            width="14%"
                                                                                                            valign="middle">
                                                                                                            Show Arrear
                                                                                                            in Same
                                                                                                            Month</td>
                                                                                                        <td class="TDcolon"
                                                                                                            valign="middle">
                                                                                                            :</td>
                                                                                                        <td
                                                                                                            valign="TDVd">
                                                                                                            <asp:CheckBox
                                                                                                                ID="chkArrSmeMnth"
                                                                                                                runat="server"
                                                                                                                CssClass="Checkbox" />
                                                                                                        </td>
                                                                                                        <td
                                                                                                            class="Message">
                                                                                                            <asp:Literal
                                                                                                                ID="Literal2"
                                                                                                                runat="server"
                                                                                                                Text="use for Excel,CSV.">
                                                                                                            </asp:Literal>
                                                                                                        </td>
                                                                                            </td>
                                                                                        </tr>

                                                                                    </table>
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                            </td>
                                                        </tr>

                                                        <tr id="Trpublish" runat="server" style="display: none;">
                                                            <td class="TdCaption" style="width: 12%">Publish per page
                                                            </td>
                                                            <td class="TDcolon">:</td>
                                                            <td>
                                                                <table cellspacing="0" cellpadding="0" width="100%"
                                                                    border="0">
                                                                    <tr>
                                                                        <td class="TDVd" style="width:36.5%">
                                                                            <asp:TextBox ID="txtpagevalue"
                                                                                runat="server" CssClass="textbox"
                                                                                oninput="checkvals()" MaxLength="2"
                                                                                Columns="3"></asp:TextBox>*
                                                                        </td>
                                                                        <td>
                                                                            <table cellspacing="0" cellpadding="0"
                                                                                width="100%" border="0">
                                                                                <tr id="trRepIn" style="display: none"
                                                                                    runat="server">
                                                                                    <td class="TdCaption"
                                                                                        style="width:22%">Report In</td>
                                                                                    <td class="TDcolon"
                                                                                        style="width:5%">:</td>
                                                                                    <td class="TDVd">
                                                                                        <asp:DropDownList
                                                                                            ID="ddlreportIn"
                                                                                            runat="server"
                                                                                            CssClass="DropdownList"
                                                                                            onchange="validatePswd();">
                                                                                            <asp:ListItem Text="PDF"
                                                                                                Value="P">
                                                                                            </asp:ListItem>
                                                                                            <asp:ListItem Text="HTML"
                                                                                                Value="H">
                                                                                            </asp:ListItem>
                                                                                        </asp:DropDownList>
                                                                                        (PDF: In PDF report will be
                                                                                        generated employee wise.)
                                                                                    </td>
                                                                                </tr>
                                                                            </table>
                                                                        </td>
                                                                    </tr>
                                                                </table>
                                                            </td>
                                                        </tr>
                                                        <%--Password for PDF--%>
                                                            <tr id="trpswd" style="display: none" runat="server">
                                                                <td class="TDCaption">PDF Password Type</td>
                                                                <td class="TdColon">:</td>
                                                                <td class="tdvd">
                                                                    <asp:DropDownList ID="ddlEmpPass" runat="server"
                                                                        CssClass="dropdownlist" Enabled="true">
                                                                        <asp:ListItem Text="No Password" Value="0">
                                                                        </asp:ListItem>
                                                                        <asp:ListItem Text="Emp. Code" Value="1">
                                                                        </asp:ListItem>
                                                                        <asp:ListItem Text="Pan No." Value="2">
                                                                        </asp:ListItem>
                                                                        <asp:ListItem Text="First Name & DOB" Value="3">
                                                                        </asp:ListItem>
                                                                        <asp:ListItem Text="Bank A/C No. & DOB"
                                                                            Value="4"></asp:ListItem>
                                                                        <asp:ListItem Text="First Name & DOB(DDMMYYYY)"
                                                                            Value="5"></asp:ListItem>
                                                                        <asp:ListItem Text="Pan No. & DOB(DDMMYY)"
                                                                            Value="6"></asp:ListItem>
                                                                        <asp:ListItem Text="Emp. Code & DOB(DDMMYY)"
                                                                            Value="7"></asp:ListItem>
                                                                    </asp:DropDownList>
                                                                </td>
                                                            </tr>
                                                            <tr id="TrEmpDetails" runat="server" style="display: none;">
                                                                <td class="TDCaption">With Employee Details</td>
                                                                <td class="TdColon">:</td>
                                                                <td class="tdvd">
                                                                    <table id="Table2" cellpadding="0" cellspacing="0"
                                                                        border="0" runat="server" width="100%">
                                                                        <tr>
                                                                            <td class="tdvd" width="60%">
                                                                                <div
                                                                                    style="border:1px;border-color:Blue;Width:600px;border-style:solid;height:50px">
                                                                                    <asp:CheckBoxList ID="chkEmpdet"
                                                                                        RepeatDirection="Horizontal"
                                                                                        RepeatColumns="5" runat="server"
                                                                                        CssClass="chkList"
                                                                                        Width="600px">
                                                                                    </asp:CheckBoxList>
                                                                                </div>
                                                                            </td>
                                                                            <td class="TDCaption"><img
                                                                                    src="ImagesCA/Note.gif"
                                                                                    border="0" /></td>
                                                                            <td class="Message" width="40%">
                                                                                <asp:Literal ID="Literal1"
                                                                                    runat="server"
                                                                                    Text="In excel report all the data (Costcenter,Location,Unit,Designation,Department, Level,Grade,DOJ,DOL) will populate always, only PAN is configurable.">
                                                                                </asp:Literal>
                                                                            </td>
                                                                        </tr>
                                                                    </table>

                                                                </td>
                                                            </tr>

                                                            <tr class="trupbtn">
                                                                <td colspan="3" class="trbtnup"></td>
                                                            </tr>
                                                            <tr id="TrSave" runat="server">
                                                                <td class="TDCaption" style="width: 12%"><input
                                                                        type="hidden" runat="server" id="hidEmpCode" />
                                                                </td>
                                                                <td class="TdColon"></td>
                                                                <td>
                                                                    <table id="Table1" cellspacing="0" cellpadding="0"
                                                                        border="0">
                                                                        <tr>
                                                                            <td class="tdbetweenbtn">
                                                                                <asp:Button ID="btnDownloadCSV"
                                                                                    runat="server" Visible="false"
                                                                                    CssClass="btn"
                                                                                    ToolTip="Download in CSV format"
                                                                                    Text="EXPORT TO CSV" Width="120px"
                                                                                    CausesValidation="False"
                                                                                    OnClientClick="javascript:return OnClickDownloadCSV();">
                                                                                </asp:Button>
                                                                            </td>
                                                                            <td class="tdbetweenbtn"></td>
                                                                            <td class="tdbetweenbtn">
                                                                                <asp:Button ID="btnexcel" runat="server"
                                                                                    CausesValidation="False"
                                                                                    CssClass="btn"
                                                                                    Text="EXPORT TO EXCEL" Width="120px"
                                                                                    OnClientClick="javascript:return btnExcel_Click();" />
                                                                            </td>
                                                                            <td class="tdbetweenbtn"></td>
                                                                            <td valign="tdbetweenbtn">
                                                                                <asp:Button ID="btnPriview"
                                                                                    runat="server" class="btn"
                                                                                    Text="Preview" Width="100px"
                                                                                    OnClientClick="return LoadProcessDialog();" />
                                                                                <asp:HiddenField
                                                                                    ID="hdf_USearchMulti_DdlCostCenter"
                                                                                    runat="server" Value="" />
                                                                                <asp:HiddenField
                                                                                    ID="hdf_USearchMulti_DdlDept"
                                                                                    runat="server" Value="" />
                                                                                <asp:HiddenField
                                                                                    ID="hdf_USearchMulti_ddllocation"
                                                                                    runat="server" Value="" />
                                                                                <asp:HiddenField
                                                                                    ID="hdf_USearchMulti_ddldesignation"
                                                                                    runat="server" Value="" />
                                                                                <asp:HiddenField
                                                                                    ID="hdf_USearchMulti_ddlunit"
                                                                                    runat="server" Value="" />
                                                                                <asp:HiddenField
                                                                                    ID="hdf_USearchMulti_ddlGrade"
                                                                                    runat="server" Value="" />
                                                                                <asp:HiddenField
                                                                                    ID="hdf_USearchMulti_ddllevel"
                                                                                    runat="server" Value="" />
                                                                            </td>
                                                                            <td class="tdbetweenbtn"></td>
                                                                            <td valign="tdbetweenbtn">
                                                                                <asp:Button ID="btnPublishedPDF"
                                                                                    runat="server" Visible="true"
                                                                                    CssClass="btn"
                                                                                    ToolTip="Download Already Published YTD Slips"
                                                                                    Text="Download Already Publish YTD Slips"
                                                                                    Width="250px"
                                                                                    CausesValidation="False"
                                                                                    OnClientClick="return PleaseWaitWithDailog();">
                                                                                </asp:Button>
                                                                            </td>
                                                                            <td class="tdbetweenbtn"></td>
                                                                            <td valign="tdbetweenbtn">
                                                                                <asp:Button ID="btnReset" runat="server"
                                                                                    CssClass="btn" Text="Reset"
                                                                                    CausesValidation="False"
                                                                                    OnClientClick="PleaseWaitWithDailog();">
                                                                                </asp:Button>
                                                                            </td>
                                                                            <td class="tdbetweenlbl"></td>

                                                                        </tr>
                                                                        <tr class="trupbtn">
                                                                            <td colspan="7" class="trbtnup">
                                                                                &nbsp;
                                                                            </td>
                                                                        </tr>

                                                                        <tr class="trupbtn">
                                                                            <td colspan="7" class="trbtnup">
                                                                                &nbsp;
                                                                            </td>
                                                                        </tr>

                                                                        <tr class="trupbtn">
                                                                            <td colspan="7" class="trbtnup">
                                                                                <asp:Label ID="lblmsg" runat="server"
                                                                                    CssClass="usermessage"></asp:Label>
                                                                            </td>
                                                                        </tr>
                                                                        <tr class="trupbtn">
                                                                            <td colspan="7" class="trbtnup"></td>
                                                                        </tr>
                                                                        <tr id="tdDwn" style="Display:none;"
                                                                            runat="server">
                                                                            <td class="TDVd" colspan="7">
                                                                                <span id="spnmsgtolink"
                                                                                    runat="server">Click to PDF icon to
                                                                                    download the .zip file.</span>
                                                                                <asp:LinkButton ID="LnkPDF"
                                                                                    runat="server"><img
                                                                                        src="Images/pdf.bmp" border="0"
                                                                                        height="18"></asp:LinkButton>
                                                                            </td>
                                                                        </tr>
                                                                        <tr class="trupbtn">
                                                                            <td colspan="7" class="trbtnup"></td>
                                                                        </tr>
                                                                    </table>
                                                                </td>
                                                            </tr>

                                                            <tr id="TrSvConfig" runat="server" style="Display:none;">
                                                                <td class="tdcaption" width="12%">
                                                                </td>
                                                                <td class="TdColon"></td>
                                                                <td colspan="3">
                                                                    <asp:Button ID="btnSaveConfig" CssClass="btn"
                                                                        CausesValidation="false" runat="server"
                                                                        Text="Save Configuration" Width="120px"
                                                                        OnClientClick="PleaseWaitWithDailog();">
                                                                    </asp:Button>

                                                                    <input type="hidden" runat="server"
                                                                        id="HidAppPath" />
                                                                    <input id="HidPreVal" type="hidden" runat="server"
                                                                        name="hidval" />
                                                                    <input type="hidden" runat="server" id="HidPath" />
                                                                    <input type="hidden" runat="server"
                                                                        id="HidEmpCodes4Pdf" />
                                                                </td>
                                                                <td></td>
                                                            </tr>

                                                            <tr id="trAlreadyReg" runat="server" visible="false">
                                                                <td class="TDCaption">Download Previous Run Report</td>
                                                                <td class="TdColon">:</td>
                                                                <td class="tdvd">
                                                                    <span id="spnAlreadyYTDGenerated" runat="server"
                                                                        class="UserPrevRunMsg"></span>
                                                                    <asp:LinkButton ID="lnkAlreadyYTDGenerated"
                                                                        runat="server"
                                                                        OnClientClick="return PleaseWaitWithDailog();"
                                                                        ToolTip="Click here to download already generated Year to Date Register">
                                                                        <img src="Images/Excel_img.jpg" border="0"
                                                                            height="18"></asp:LinkButton>
                                                                    <br />
                                                                    <span id="spnAlreadyYTDGeneratedEnd" runat="server"
                                                                        class="UserPrevRunMsg"></span>
                                                                    <input type="hidden" runat="server"
                                                                        id="hdnYtdRegPath" />
                                                                    <input type="hidden" runat="server"
                                                                        id="hdnYtdRegFileName" />
                                                                </td>
                                                            </tr>
                                                    </table>
                                                    </ContentTemplate>
                                                    <Triggers>
                                                        <asp:PostBackTrigger ControlID="btnDownloadCSV" />
                                                        <asp:PostBackTrigger ControlID="btnexcel" />
                                                        <asp:PostBackTrigger ControlID="LnkPDF" />
                                                        <asp:PostBackTrigger ControlID="btnPublishedPDF" />
                                                        <asp:PostBackTrigger ControlID="lnkAlreadyYTDGenerated" />
                                                    </Triggers>
                                                    </asp:UpdatePanel>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <input type="hidden" runat="server" id="hidPDFFileName" />
                                                <input type="hidden" runat="server" id="HidEmpPdfName" />
                                                <input id="hdfile" type="hidden" name="hdfile" runat="server" />
                                                <input type="hidden" id="JobUniqueId" name="JobUniqueId"
                                                    runat="server" />
                                                <input type="hidden" id="hdnBatchId" name="BatchId" runat="server"
                                                    value="" />
                                                <input type="hidden" id="hdnFileFormat" name="FileFormat" runat="server"
                                                    value="EXCEL" />
                                                <input type="hidden" id="hdnRptName" name="hdnRptName" runat="server"
                                                    value="YTD Salary Register Report" />
                                                <input type="hidden" id="hdnProcessType" name="ProcessType"
                                                    runat="server" value="YTDSALREG" />
                                                <input type="hidden" id="hdnCSV" name="ShowCSV" runat="server"
                                                    value="N" />
                                                <input type="hidden" id="report_service" name="report_service"
                                                    runat="server" value="N" />
                                                <input type="hidden" id="report_service_url" name="report_service_url"
                                                    runat="server" value="N" />
                                                <input type="hidden" id="report_service_file_name"
                                                    name="report_service_file_name" runat="server" value="" />
                                            </td>
                                        </tr>
                                    </table>
                                    </fieldset>
                                </td>
                            </tr>
                        </table>
                        <!--This is Internal Body Table-->
                        </td>
                        </tr>
                        <tr>
                            <td>
                            </td>
                            &nbsp;
                        </tr>
                        </table>
                    </form>
                </body>

                </html>