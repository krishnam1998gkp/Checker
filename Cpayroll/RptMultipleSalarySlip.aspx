<%@ Page Language="vb" AutoEventWireup="false" Debug="true" Inherits="Payroll.RptMultipleSalarySlip"
    CodeFile="RptMultipleSalarySlip.aspx.vb" ValidateRequest="false" EnableEventValidation="false" %>

    <%@ Register Src="~/CPayroll/UsearchWithMultipleEmpCodeproxy.ascx" TagName="UserControl" TagPrefix="UC" %>
        <%@ Register TagPrefix="cc1" Namespace="EasyWebEditorControl" Assembly="EasyWebEditorControl" %>
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
                    <meta http-equiv="Page-Enter" content="blendTrans(duration=1)" />
                    <link href="Css/hospCSS.css" type="text/css" rel="stylesheet" />
                    <link href="Css/ReportStylePayRolAdmin.css" type="text/css" rel="stylesheet" />
                    <link href="Css/CommonProgressBarStyle.css?v=1.0.1" type="text/css" rel="stylesheet" />
                    <script language="javascript" src="JavaFiles/JSBalloon.js" type="text/javascript"></script>
                    <link href="Css/jquery-ui-1.8.1.custom.css" type="text/css" rel="stylesheet" />
                    <script language="javascript" src="JavaFiles/jquery-1.4.2.min.js" type="text/javascript"></script>
                    <script language="javascript" src="JavaFiles/jquery-ui-1.8.1.custom.min.js"
                        type="text/javascript"></script>
                    <script language="javascript" src="JavaFiles/jquery.bgiframe.js" type="text/javascript"></script>
                    <script language="javascript" src="JavaFiles/Script_utill.js" type="text/javascript"></script>
                    <script language="javascript" src="JavaFiles/JS_CommonUtill.js" type="text/javascript"></script>
                    <script language="javascript" src="JavaFiles/JQ_PayslipConfig.js" type="text/javascript"></script>
                    <script language="javascript" src="JavaFiles/Js_ScrollGrid.js" type="text/javascript"></script>
                    <%--<script language="javascript" src="JavaFiles/ProcessBar.js" type="text/javascript"></script>--%>
                        <script language="javascript" src="JavaFiles/ProcessBarWithoutPasswd.js?v=1.0.2"
                            type="text/javascript"></script>
                        <script language="javascript" src="JavaFiles/PaySlipsProcessBar.js?v=1.0.28"
                            type="text/javascript"></script>
                        <script language="javascript" src="JavaFiles/CommonDownload.js?v=1.0.1"
                            type="text/javascript"></script>
                        <script language="javascript" src="JavaFiles/ReportApiProgressBarScript.js?v=1.0.10"
                            type="text/javascript"></script>
                        <style type="text/css">
                            .btnGreen {
                                cursor: pointer;
                                color: #FFFFFF;
                                text-decoration: none;
                                font-family: arial;
                                font-size: 9px;
                                height: 20px;
                                background-color: green;
                                text-transform: uppercase;
                                font-weight: bold;
                                filter: progid:DXImageTransform.Microsoft.Gradient(endColorstr='#036503', startColorstr='#64ab4d', gradientType='0');
                                background: -moz-linear-gradient(top, #036503, #64ab4d);
                                background: -webkit-gradient(linear, left top, left bottom, from(#036503), to(#64ab4d));
                                background: -o-linear-gradient(#036503, #64ab4d);
                                border: black 1px solid;
                            }

                            .btnGreen:hover {
                                background-image: linear-gradient(to bottom, #fff, #35aa35);
                            }

                            @keyframes blinkColor {
                                0% {
                                    color: blue;
                                }

                                35% {
                                    color: green;
                                }

                                70% {
                                    color: blue;
                                }

                                100% {
                                    color: green;
                                }
                            }

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
                                animation: blinkColor 1s infinite alternate;
                            }

                            table.grid {
                                font: 11px arial, helvetica, sans-serif;
                                border-collapse: collapse;
                                border: 1px solid #ccc;
                            }

                            table.grid th {
                                background: #DBDEE5;
                                font-size: 14px;
                                text-align: center;
                                color: #0066CC;
                                border-right: 1px solid silver;
                                position: relative;
                                cursor: default;
                                z-index: 10;
                            }

                            .tab {
                                cursor: hand;
                                color: #000000;
                                text-decoration: none;
                                font-family: Verdana;
                                font-size: 10px;
                                text-align: center;
                                width: 80px;
                                height: 20px;
                                border-bottom: #9BC324 2px solid;
                                border-left: #9BC324 2px solid;
                                border-top: #9BC324 2px solid;
                                border-right: #9BC324 2px solid;
                                background-color: #A4C0E8;
                                text-transform: uppercase;
                                font-weight: bold;
                                vertical-align: text-bottom;
                                filter: progid:DXImageTransform.Microsoft.Gradient(endColorstr='#f0f0f0', startColorstr='#72bfee', gradientType='0');
                                background-image: url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxMDAlIiBoZWlnaHQ9IjEwMCUiPjxsaW5lYXJHcmFkaWVudCBpZD0iZzEiIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIiB4MT0iMCUiIHkxPSIwJSIgeDI9IjAlIiB5Mj0iMTAwJSI+PHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjNzJiZmVlIi8+PHN0b3Agb2Zmc2V0PSIxIiBzdG9wLWNvbG9yPSIjZjBmMGYwIi8+PC9saW5lYXJHcmFkaWVudD48cmVjdCB4PSIwIiB5PSIwIiB3aWR0aD0iMTAwJSIgaGVpZ2h0PSIxMDAlIiBmaWxsPSJ1cmwoI2cxKSIgLz48L3N2Zz4=);
                                background-image: -webkit-gradient(linear, center top, center bottom, color-stop(0%, #72bfee), color-stop(100%, #f0f0f0));
                                background-image: -webkit-linear-gradient(top, #72bfee 0%, #f0f0f0 100%);
                                background-image: -moz-linear-gradient(top, #72bfee 0%, #f0f0f0 100%);
                                background-image: -ms-linear-gradient(top, #72bfee 0%, #f0f0f0 100%);
                                background-image: -o-linear-gradient(top, #72bfee 0%, #f0f0f0 100%);
                                background-image: linear-gradient(to bottom, #72bfee 0%, #f0f0f0 100%);
                            }

                            .btn_pdf {
                                border: 0;
                                cursor: pointer;
                            }
                        </style>
                </head>
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
                        let gcs_powered = document.getElementById("is_gcs_powered");
                        if (gcs_powered && gcs_powered.value == "1") {
                            $("LnkPDFWOPWD").hide();
                            $("LnkPDF").hide();
                        }

                    });
                    // fixing error
                    window.createPopup = function () {

                    };
                    //Added by Vishal Chauhan for process bar initial message
                    function StartProcessbar(reptName) {
                        if (reptName == '') {
                            reptName = 'Salary Register Report';
                        }
                        InitiateDynamicRegisterExcelProcess(reptName + " Process", "Preparing data to generate " + reptName + ", This will only take a moment. Please do not close the window.", 500);
                    }

                    function StartCSVProcessbar(reptName) {
                        if (reptName == '') {
                            reptName = 'Salary Register Report';
                        }
                        InitiateDynamicRegisterExcelCSVProcess(reptName + " Process", "Preparing data to generate " + reptName + ", This will only take a moment. Please do not close the window.", 500);
                    }

                    function isValidFileName() {
                        // Disallow /\:*?"<>|
                        var regex = /^[^\/\\:\*\?"<>\|]+$/;
                        if (!regex.test($("#txtrptName").val())) {
                            alert('Special characters / \\ : * ? \" < > | are not allowed.');
                            $("#txtrptName").val("");
                            $("#txtrptName").focus();
                            return false;
                        }
                    }

                    function Check4PublishedPayslip() {
                        /*Validation Added by Vishal Chauhan on 18 Dec 2024 to lock process userwise*/
                        if ($('#<%= lblProcessStatus.ClientID %>').text().trim() != '') {
                            let str = $('#<%= lblProcessStatus.ClientID %>').text();
                            alert(str);
                            return false;
                        }
                        if ($('#lblProcessStatus').text().trim().length > 1) {
                            let str = $('#lblProcessStatus').text();
                            alert(str);
                            return false;
                        }

        <% --if ($('#<%= lblProcessStatusExcel.ClientID %>').text().trim() != '') {
                            let str = $('#<%= lblProcessStatusExcel.ClientID %>').text();
                            alert(str);
                            return false;
                        } --%>
        return PleaseWaitWithDailog();
                    }
                    function ValidateEmails(txtCC, txtBCC) {
                        var filter = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
                        var emailCC = txtCC.val().split(',');
                        var emailBCCLen = txtBCC.val().split(',');
                        _varCheck = true;
                        if (txtCC.val() != '') {
                            if (emailCC.length > 10) {
                                alert("Please enter max 10 CC emails!");
                                return false;
                            }
                            else {
                                for (i = 0; i <= emailCC.length - 1; i++) {
                                    if (!filter.test(emailCC[i])) {
                                        alert("Please enter a valid email address in CC [" + emailCC[i] + "]");
                                        txtCC.css({ "background-color": "yellow" });
                                        txtCC.focus();
                                        _varCheck = false;
                                    }
                                    if (_varCheck == false) { break; }
                                }
                            }
                        }
                        if (txtBCC.val() != '') {
                            if (emailBCCLen.length > 10) {
                                alert("Please enter max 10 BCC emails!");
                                return false;
                            }
                            else {
                                for (i = 0; i <= emailBCCLen.length - 1; i++) {
                                    if (!filter.test(emailBCCLen[i])) {
                                        alert("Please enter a valid email address in BCC [" + emailBCCLen[i] + "]");
                                        txtBCC.css({ "background-color": "yellow" });
                                        txtBCC.focus();
                                        _varCheck = false;
                                    }
                                    if (_varCheck == false) { break; }
                                }
                            }
                        }
                        return _varCheck;
                    }

                    function empchecked(name, _objID) {
                        var ProcessName, COC_Check = $('#HidCocManCheck').val();
                        var str = _objID.id;

                        $('#lblMsgSlip,#lblMailMsg').text("")
                        if (name != 'btnSave' && ValidateEmails($("#txtccc"), $("#txtBCC")) == false) { return false; }
                        if ($('#RblNoSearch input:checked').val().trim() != 'P') {
                            /*Validation Added by Vishal Chauhan on 18 Dec 2024 to lock process userwise*/
                            if ($('#<%= lblProcessStatus.ClientID %>').text().trim() != '') {
                                let str = $('#<%= lblProcessStatus.ClientID %>').text();
                                alert(str);
                                return false;
                            }
            <% --if ($('#<%= lblProcessStatusExcel.ClientID %>').text().trim() != '') {
                                let str = $('#<%= lblProcessStatusExcel.ClientID %>').text();
                                alert(str);
                                return false;
                            } --%>
            if (mandatoryCheckBoxCheckOnGridWithPWait('DvEmp', 'chkEmpHold', 'Please select atleast one employee!', _objID) == true) {
                                if (name == 'BtnSend' && $("#ddlRepIn").val() != "H") {
                                    if ($("#chkmailformat").attr('checked') == true) {
                                        if (($("#txtheader").val() == "") || ($('#EasyWebMAilBodyHTMLCONTENT').val() == "") || ($("#Textfooter").val() == "")) {
                                            var StrMand = '';
                                            StrMand += $("#txtheader").val() == "" ? '<%=_objCommon.MandatoryMsg("Header")%>\n' : "";
                                            StrMand += $("#EasyWebMAilBodyHTMLCONTENT").val() == "" ? '<%=_objCommon.MandatoryMsg("Contents")%>\n' : "";
                                            StrMand += $("#Textfooter").val() == "" ? '<%=_objCommon.MandatoryMsg("Footer")%>\n' : "";
                                            if (StrMand != '') {
                                                alert(StrMand);
                                                $("#BtnSend").val("Send email to employees and CC/BCC Email ID�s");
                                                $('#' + name).css('width', '300px');
                                                return false;
                                            }
                                        }
                                        else {
                                            if (confirm("You have modified the mail format. Do you wish to continue?") == true) {
                                                $("#BtnSend").val("Please Wait..Sending Mails...");
                                                $("#hidSaveMail").val("Y");
                                                return true;
                                            }
                                            /*Added by Rohtas Singh on 27 Dec 2017 for save value if user click on cancel option*/
                                            else {
                                                $("#hidSaveMail").val("N");
                                            }
                                        }
                                    }
                                }

                                if (name == 'BtnSend') { $('#' + name).val("Please Wait..Sending Mails..."); $('#' + name).css('width', '210px'); ProcessName = "mailed"; }
                                else {
                                    if (_objID.id == 'btnWOPWD') { $('#btnWOPWD').val("Please Wait Pay Slips Publishing..."); $('#btnWOPWD').css('width', '230px'); ProcessName = "published"; }
                                    else if (name == 'btnSave') {
                                        $('#' + name).val("Please Wait Pay Slips Publishing..."); $('#' + name).css('width', '230px'); ProcessName = "published";
                                    }
                                    else {
                                        $('#' + name).val("Please Wait..Sending Mails...");
                                        $('#' + name).css('width', '200px');
                                        ProcessName = "published";
                                    }
                                }
                                if ($("#ddlRepIn").val() == "P") {
                                    var msg = "\nHi " + '<%=session("UName")%>';
                                    msg = msg + ",\n";
                                    var RptName = $("#DdlreportType option:selected").text();
                                    var Password = $("#ddlEmpPass option:selected").text();
                                    if (name == 'BtnPublishGrpBy') {
                                        msg = 'Do you want to publish salary slip(s) unit wise and email to unit authority?';
                                    }
                                    else if (str == 'btnWOPWD') {
                                        msg = msg + "All the " + RptName + " will be " + ProcessName + " without password.\n\nDo you want to continue?";
                                    }
                                    else {
                                        /*Added by Quadir on 14 OCT 2020 for specific alert message*/
                                        if ($("#ddlEmpPass").val() == '0') {
                                            if (($("#DdlreportType").val() == "R") && ($('#rbtSlipPubMode input:checked').val() == 'I')) { msg = msg + RptName + " will be " + ProcessName + " without password.\nOnce payslips are published they are sent to file server.\nIn Incremental Mode the system fetches the existing payslips for the month from file server and regenerates the one's which do not exist there to create the zip folder of payslips.\n\nDo you want to continue?"; }
                                            else if (($("#DdlreportType").val() == "R") && ($('#rbtSlipPubMode input:checked').val() == 'O')) { msg = msg + RptName + " will be " + ProcessName + " without password.\nOnce payslips are published they are sent to file server.\nIn Overwrite Mode the system generates payslips for the filtered employees and places all of them on file server, replacing any pre existing payslips for the filtered employees.\n\nDo you want to continue?"; }
                                            else {
                                                msg = msg + "All the " + RptName + " will be " + ProcessName + " without password.\n\nDo you want to continue?";
                                            }
                                        }

                                        else { msg = msg + "All the " + RptName + " will be " + ProcessName + " password protected. Password to open the PDF file is " + Password + ".\n\nDo you want to continue?"; }
                                    }

                                    if (ConfirmDelete(msg) == true) {
                                        if (name == 'BtnPublishGrpBy') {
                                            $("#LnkPDF").css("display", "none");
                                        }
                                        else if (name == 'btnSave') {
                                            $("#LnkPDF").css("display", "");
                                            if ($("#DdlreportType").val().toUpperCase() == 'T' || $("#DdlreportType").val().toUpperCase() == 'R' || $("#DdlreportType").val().toUpperCase() == 'S' || $("#DdlreportType").val().toUpperCase() == '57') {
                                                LoadPaySlipProgress(RptName);
                                            }
                                        }
                                        else {
                                            $("#LnkPDF").css("display", "");
                                        }
                                        // added show progress bar immediately.  added by Kangkan
                                        if (['BtnSendCCBCC', 'BtnSend', 'BtnPublishGrpBy'].includes(name) && $("#DdlreportType").val().toUpperCase() === 'T') {

                                            LoadPaySlipProgress(RptName);
                                        }


                                        return true;
                                    }
                                    else {
                                        $("#btnSave").val("Publish and generate Zip file");
                                        $("#BtnSend").val("Send email to employees and CC/BCC Email ID�s");
                                        $("#BtnSendCCBCC").val("Publish and email to CC / BCC");
                                        $("#BtnPublishGrpBy").val("Unit Wise Publish & Email");
                                        $("#btnWOPWD").val("Publish Payslip Without Password");
                                        return false;
                                    }
                                }

                                else
                                    return true;
                            }
                            else return false;
                        }
                        else if (name == 'btnSave') {
                            if ($("#USearch_ddlcostcenter").val() == "" && COC_Check == 'Y' && $("#USearch_ddlcostcenter").val() == "" && COC_Check == 'Y' && $("#DdlreportType").val() != "T" && $("#DdlreportType").val() != "I" && $("#DdlreportType").val() != "SH" && $("#DdlreportType").val() != "RS" && $("#DdlreportType").val() != "43") {
                                alert('<%=_objcommon.DisplayCaption("COC")%>' + " cannot be left blank!");
                                $("#USearch_ddlcostcenter").focus();
                                return false;
                            }
                            /*Validation Added by Vishal Chauhan on 18 Dec 2024 to lock process userwise*/
                            if ($('#<%= lblProcessStatus.ClientID %>').text().trim() != '') {
                                let str = $('#<%= lblProcessStatus.ClientID %>').text();
                                alert(str);
                                return false;
                            }
            <% --if ($('#<%= lblProcessStatusExcel.ClientID %>').text().trim() != '') {
                                let str = $('#<%= lblProcessStatusExcel.ClientID %>').text();
                                alert(str);
                                return false;
                            } --%>

            /*START: Added by Rajarshi on 01 Nov 2017*/
            var RepName = $("#DdlreportType option:selected").text();
                            if (RepName == "--Select Report--") {
                                alert("Please select Report Type!");
                                return false;
                            }

                            if ($("#DdlreportType").val().toUpperCase() === "SL") {
                                RepName = $("#DdlreportType option:selected").text();
                                LoadPaySlipProgress(RptName);
                            } else {

                            }
                            /*END: Added by Rajarshi on 01 Nov 2017*/
                            /*START: Added by Rohtas Singh on 15 Jul 2020*/
                            if ($("#DdlreportType").val() == 63) {
                                if ($("#USearch_ddlunit").val() == "") {
                                    alert('<%=_objcommon.DisplayCaption("UNT")%>' + " cannot be left blank!");
                                    $("#USearch_ddlunit").focus();
                                    return false;
                                }
                                if ($("#USearch_ddldesignation").val() == "") {
                                    alert('<%=_objcommon.DisplayCaption("DES")%>' + " cannot be left blank!");
                                    $("#USearch_ddldesignation").focus();
                                    return false;
                                }
                            }
                            /*END: Added by Rohtas Singh on 15 Jul 2020*/
                            if (str == 'btnWOPWD') {
                                $('#btnWOPWD').val("Please Wait Pay Slips Publishing..."); $('#btnWOPWD').css('width', '210px'); ProcessName = "published";

                            } else {
                                $('#' + name).val("Please Wait Pay Slips Publishing..."); $('#' + name).css('width', '210px'); ProcessName = "published";

                            }
                            var msg = "\nHi " + '<%=session("UName")%>';

                            msg = msg + ",\n";
                            var RptName = $("#DdlreportType option:selected").text();
                            var Password = $("#ddlEmpPass option:selected").text();

                            /*Added by Quadir on 14 OCT 2020 for specific alert message*/
                            if (str == 'btnWOPWD') {
                                msg = msg + RptName + " will be " + ProcessName + " without password.\n\nDo you want to continue?";
                            }
                            else if ($("#ddlEmpPass").val() == '0') {
                                if (($("#DdlreportType").val() == "R") && ($('#rbtSlipPubMode input:checked').val() == 'I')) { msg = msg + RptName + " will be " + ProcessName + " without password.\nOnce payslips are published they are sent to file server.\nIn Incremental Mode the system fetches the existing payslips for the month from file server and regenerates the one's which do not exist there to create the zip folder of payslips.\n\nDo you want to continue?"; }
                                else if (($("#DdlreportType").val() == "R") && ($('#rbtSlipPubMode input:checked').val() == 'O')) { msg = msg + RptName + " will be " + ProcessName + " without password.\nOnce payslips are published they are sent to file server.\nIn Overwrite Mode the system generates payslips for the filtered employees and places all of them on file server, replacing any pre existing payslips for the filtered employees.\n\nDo you want to continue?"; }
                                else {
                                    msg = msg + "All the " + RptName + " will be " + ProcessName + " without password.\n\nDo you want to continue?";
                                }
                            }

                            else { msg = msg + "All the " + RptName + " will be " + ProcessName + " password protected. Password to open the PDF file is " + Password + ".\n\nDo you want to continue?"; }

                            if (ConfirmDelete(msg) == true) {
                                $("#LnkPDF").css("display", "");
                                if ($("#DdlreportType").val().toUpperCase() == 'T' || $("#DdlreportType").val().toUpperCase() == 'R' || $("#DdlreportType").val().toUpperCase() == 'S' || $("#DdlreportType").val().toUpperCase() == '57') {
                                    LoadPaySlipProgress(RptName);
                                }
                                return true;
                            }
                            else {
                                $("#btnSave").val("Publish and generate Zip file");
                                $("#BtnSend").val("Send email to employees and CC/BCC Email ID�s");
                                $("#BtnSendCCBCC").val("Publish and email to CC / BCC");
                                $("#BtnPublishGrpBy").val("Unit Wise Publish & Email");
                                $("#btnWOPWD").val("Publish Payslip Without Password");
                                return false;
                            }

                        }
                        else if (name == 'BtnPublishGrpBy') {
                            if ($("#USearch_ddlcostcenter").val() == "" && COC_Check == 'Y') {
                                alert('<%=_objcommon.DisplayCaption("COC")%>' + " cannot be left blank!");
                                $("#USearch_ddlcostcenter").focus();
                                return false;
                            }

                            $('#' + name).val("Please Wait Pay Slips Publishing..."); $('#' + name).css('width', '200px'); ProcessName = "published";
                            var msg;
                            var RptName = $("#DdlreportType option:selected").text();
                            var Password = $("#ddlEmpPass option:selected").text();
                            msg = 'Do you want to publish salary slip(s) unit wise and email to unit authority?';
                            if (ConfirmDelete(msg) == true) {
                                $("#LnkPDF").css("display", "none");
                                return true;
                            }
                            else {
                                $("#btnSave").val("Publish and generate Zip file");
                                $("#BtnSend").val("Send email to employees and CC/BCC Email ID�s");
                                $("#BtnSendCCBCC").val("Publish and email to CC / BCC");
                                $("#BtnPublishGrpBy").val("Unit Wise Publish & Email");
                                $("#btnWOPWD").val("Publish Payslip Without Password");
                                return false;
                            }
                        }
                    }

                    function CloseSlipProgressbar() {
                        $('#dlg').dialog('destroy');
                    }

                    function DblClicked(form) {
                        var s = $('#lsthelp').val()
                        var temp1 = null;
                        var hiddenformail = null;
                        var hidden2 = $('#Hidden4').val();
                        if (hidden2 == "") {

                            form1.hiddenformail.value = s;
                            temp1 = form1.hiddenformail.createTextRange();
                            temp1.execCommand('copy');
                            cmdExec('Paste');

                            return false;
                        }
                        else {
                            form1.hiddenformail.value = s;
                            $('#' + hidden2).val($('#' + hidden2).val() + s);
                        }
                        return false;
                    }

                    function Clicked(ctrl, _flg) {
                        if (_flg == 'Y') {
                            $('#Hidden4').val(ctrl);
                            return true;
                        }
                        else { $('#Hidden4').val(''); }
                    }

                    function popluateoffcycledate() {

                        if ($("#DDLPaySlipType").val() == "62") {
                            $('#DDLPaySlipType').attr('selectedIndex', '0');
                            $('#troffcycle').css("display", "none");
                            __doPostBack('', '');
                            PleaseWaitWithDailog();

                        }

                        if ($("#DdlreportType").val() == "62") {
                            $('#DdlreportType').attr('selectedIndex', '0');
                            $('#troffcycledt').css("display", "none");
                            __doPostBack('', '');
                            $('#rbtnmail').attr('checked', true);
                            PleaseWaitWithDailog();
                        }


                    }

                    function btnPreview_Click() {
                        if ($('#<%= lblProcessStatusExcel.ClientID %>').text().trim() != '') {
                            let str = $('#<%= lblProcessStatusExcel.ClientID %>').text();
                            alert(str);
                            return false;
                        }
        <% --if ($('#<%= lblProcessStatus.ClientID %>').text().trim() != '') {
                            let str = $('#<%= lblProcessStatus.ClientID %>').text();
                            alert(str);
                            return false;
                        } --%>
        if ($("#DDLPaySlipType").val() == "") {
                            alert("Please select Report Type!");
                            $("#DDLPaySlipType").focus();
                            return false;
                        }
                        else if ($("#DDLPaySlipType").val() == "6" && $("#ddlformat").val() == "") {
                            alert('<%=_objCommon.MandatoryMsg("Report Format")%>');
                            $("#ddlformat").focus();
                            return false;
                        }
                        else if ($("#DDLPaySlipType").val() == "0" || $("#DDLPaySlipType").val() == "52" || $("#DDLPaySlipType").val() == "51" || $("#DDLPaySlipType").val() == "49" || $("#DDLPaySlipType").val() == "18" || $("#DDLPaySlipType").val() == "12" || $("#DDLPaySlipType").val() == "50" || $("#DDLPaySlipType").val() == "20" || $("#DDLPaySlipType").val() == "7" || $("#DDLPaySlipType").val() == "23" || $("#DDLPaySlipType").val() == "24" || $("#DDLPaySlipType").val() == "37" || $("#DDLPaySlipType").val() == "5" || $("#DDLPaySlipType").val() == "10" || $("#DDLPaySlipType").val() == "26" || $("#DDLPaySlipType").val() == "53" || $("#DDLPaySlipType").val() == "55" || $("#DDLPaySlipType").val() == "64" || $("#DDLPaySlipType").val() == "65" || $("#DDLPaySlipType").val() == "66" || $("#DDLPaySlipType").val() == "68" || $("#DDLPaySlipType").val() == "76" || $("#DDLPaySlipType").val() == "77") {
                            var COC_Check = $('#HidCocManCheck').val();
                            if ($("#USearch_ddlcostcenter").val() == "" && COC_Check == 'Y') {
                                alert('<%=_objcommon.DisplayCaption("COC")%>' + " cannot be left blank!");
                                $("#USearch_ddlcostcenter").focus();
                                return false;
                            }
                        }
                        else if ($("#DDLPaySlipType").val() == "62" && $("#ddlmonthyearS").val() == "") {
                            alert('<%=_objCommon.MandatoryMsg("Off-cycle date is Manadatory!")%>');
                            $("#ddlmonthyearS").focus();
                            return false;
                        }
                        //Added by Debargha on 17-May-2024 for 'Please Wait' clickbait in case of other payslip report_types
                        else if ($("#DDLPaySlipType").val() == "38") {
                            if (chkSFTP.checked == true) {
                                if (confirm("Monthly Salary Register In Excel[Dynamic] will also transfer to SFTP Server. Do you want to continue?")) {
                                    PleaseWaitWithDailog();
                                    return true;
                                }
                                else { return false; }
                            }
                            else {
                                PleaseWaitWithDailog();
                            }
                        }
                        else {
                            if ($('#tblpaycode').css('display') != 'none') {
                                var count = $("#AddPaycode input[type='checkbox'][id^='cbladd']:checked").length;
                                var count1 = $("#DedPaycode input[type='checkbox'][id^='cbldeduction']:checked").length;
                                if (count <= 0 && count1 <= 0) {
                                    alert('Please select atleast one addition or deduction type paycode!')
                                    return false;
                                }
                                else if (count <= 0 && count1 > 0) { return confirm('You have selected only deduction type paycode. If there is no any values according to the selection report will not appear. Do you want continue!') }
                                else if (count > 0 && count1 <= 0) { return confirm('You have selected only addition type paycode. If there is no any values according to the selection report will not appear. Do you want continue!') }
                                else return true;
                            }
                        }
                    }

                    function btnSearch_Click(ctrl) {
                        var COC_Check = $('#HidCocManCheck').val();
                        $('#lblMsgSlip,#lblMailMsg').text("")

                        /*Validation Added by Vishal Chauhan on 18 Dec 2024 to lock process userwise*/
                        if ($('#<%= lblProcessStatus.ClientID %>').text().trim() != '') {
                            let str = $('#<%= lblProcessStatus.ClientID %>').text();
                            alert(str);
                            return false;
                        }

        <% --if ($('#<%= lblProcessStatusExcel.ClientID %>').text().trim() != '') {
                            let str = $('#<%= lblProcessStatusExcel.ClientID %>').text();
                            alert(str);
                            return false;
                        } --%>
        if ($("#rbtnmail:checked").val().trim().toLowerCase() == 'rbtnmail' && $('#USearch_txtEmpCode').val().trim() == '' && $("#DdlreportType").val().toUpperCase() == 'T' && $("#ddlRepIn").val().toUpperCase() == 'H') {
                            alert("Please enter employee codes to print TDS Estimation Slip in HTML format.");
                            return false;
                        }
                        if ($("#DdlreportType").val() == '') {
                            alert("Please select Report Type!");
                            $("#DdlreportType").focus();
                            return false;
                        }
                        else if ($("#USearch_ddlcostcenter").val() == "" && COC_Check == 'Y' && $("#DdlreportType").val() != "T" && $("#DdlreportType").val() != "I" && $("#DdlreportType").val() != "SH" && $("#DdlreportType").val() != "RS" && $("#DdlreportType").val() != "43") {
                            alert('<%=_objCommon.DisplayCaption("COC")%>' + " cannot be left blank!");
                            $("#USearch_ddlcostcenter").focus();
                            return false;
                        }
                        else {
                            if ($("#USearch_ddlcostcenter").val() != '' && $("#DdlreportType").val() != "T" && $("#DdlreportType").val() != "43" && $("#DdlreportType").val() != "RS" && $("#DdlreportType").val() != "SH" && $("#DdlreportType").val() != "I" && $("#DdlreportType").val() != "55" && $("#DdlreportType").val() != "56") {
                                $('#USearch_ddlcostcenter').attr('disabled', 'disabled');
                            }
                            $('#ddlRepIn,#chkemailrepmanager,#DdlreportType,#ddlSalaryProcess,#ddlEMailExist,#ddlEMailSend,#ddlSalWithHeld,#ddlEmpPass,#ddlmonthyearS').attr('disabled', 'disabled');
                            $('#LnkPDF,#tblmail,#trrepformat').css("display", "none")
                            $('#hidstatus').val("1")
                            if ($('#ddlRepIn').val() == "P") { $('#btnSave,#BtnSendCCBCC,#BtnSend,#BtnLog').css("display", ""); }
                            else {
                                $('#btnSave,#BtnSendCCBCC,#BtnLog').css("display", "none");
                                $('#BtnSend').css("display", "");

                                if ($("#DdlreportType").val() == '43') {
                                    $('#trselall').css("display", "none");
                                }
                            }
                            PleaseWaitWithDailog();
                        }
                    }
                    // if you use jQuery, you can load them when dom is read.
                    $(document).ready(function () {
                        // this needed to capture the client side events of the update panel
                        var prm = Sys.WebForms.PageRequestManager.getInstance();
                        prm.add_endRequest(getDisabledEnc);
                    });

                    function getDisabledEnc() {
                        if ($('#hidstatus').val() == "1") {
                            $('#ddlRepIn,#chkemailrepmanager,#DdlreportType,#ddlSalaryProcess,#ddlEMailExist,#ddlEMailSend,#ddlSalWithHeld,#ddlEmpPass').attr('disabled', 'disabled');
                            $('#trrepformat').css("display", "none")
                            $('#hidstatus').val("1")
                            //$('#lblMsgSlip,#lblMailMsg').text("")
                            if ($('#ddlRepIn').val() == "P") { $('#btnSave,#BtnSendCCBCC,#BtnSend,#BtnLog').css("display", "") }
                            else {
                                $('#btnSave,#BtnSendCCBCC,#BtnLog').css("display", "none");
                                $('#BtnSend').css("display", "");
                            }
                        }
                    }

                    function ShowHide(obj) {
                        if (!obj.checked) { $("#tbl_QuestionTags_Software").css("display", "none"); }
                        else { $("#tbl_QuestionTags_Software").css("display", ""); }
                    }

                    function blankcheck1() {
                        if ($("#DDLPaySlipType").val() == 63) {
                            if ($("#USearch_ddlunit").val() == "") {
                                alert('<%=_objCommon.DisplayCaption("UNT")%>' + " cannot be left blank!");
                                $("#USearch_ddlunit").focus();
                                return false;
                            }
                            if ($("#USearch_ddldesignation").val() == "") {
                                alert('<%=_objCommon.DisplayCaption("DES")%>' + " cannot be left blank!");
                                $("#USearch_ddldesignation").focus();
                                return false;
                            }
                        }

                        if ($("#DDLPaySlipType").val() == "") {
                            alert('<%=_objCommon.MandatoryMsg("Report Type")%>');
                            $("#DDLPaySlipType").focus();
                            return false;
                        }
                        else { openlist('N'); }
                    }
                    function openlist(id) {
                        var qryP = $("#HidPreVal").val();
                        var Reptype;
                        var height = window.screen.availHeight;
                        var width = window.screen.availWidth;
                        if (id == 'N') { Reptype = $("#DDLPaySlipType").val(); }
                        else { Reptype = $("#ddlreporttype").val(); }
                        /*This statement is used for display "Monthly Pay Slip Include Details With Leave Bal"*/
                        if (Reptype == 0 || Reptype == 50 || Reptype == 56) { qryP = 'Pre_SalarySlip.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly Salary Register"*/
                        else if (Reptype == 1) { qryP = 'Pre_SalaryRegister.aspx?id=' + qryP; }
                        /*This statement is used for display "Salary Register Head Wise"*/
                        else if (Reptype == 2) { qryP = 'preSalaryRegisterUpdated.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly Arrear Register"*/
                        else if (Reptype == 3) { qryP = 'Pre_EmpArrearDetails.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly CTC Arrear Register"*/
                        else if (Reptype == 4) { qryP = 'Pre_ArrearsReport.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly Pay Slip With Investment"*/
                        else if (Reptype == 5) { qryP = ($("#ddlarrear").val() == "2" ? "Pre_SalSlipwithINVEST.aspx" : "PreSalSlipInvestment.aspx") + '?id=' + qryP; }
                        /*This statement is used for display "Department Wise Salary"*/
                        else if (Reptype == 6) { qryP = ($("#ddlformat").val() == "0" ? "Pre_DepartmentWiseSalarywithpaycode.aspx" : "Pre_DepartmentWiseSalarywithoutpaycode.aspx") + '?id=' + qryP; }
                        /*This statement is used for display "Monthly Pay Slip Exclude Details"*/
                        else if (Reptype == 7) { qryP = 'PreSalarySlip_Multiserv.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly A4Size Salary Register"*/
                        else if (Reptype == 8) { qryP = 'Pre_SalaryRegisterA4Size.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly A4Size(4Rows) Salary Register"*/
                        else if (Reptype == 9) { qryP = 'Pre_SalaryRegisterA4Size4Row.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly Pay Slip With PF and Loan details"*/
                        else if (Reptype == 10) { qryP = 'PreSalarySlip_PFSecurity.aspx?id=' + qryP; }
                        /*This statement is used for display "Annual Arrear Details"*/
                        else if (Reptype == 11) { qryP = 'pre_ArrearAnnualDetails.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly Pay Slip Include Details W/O Leave Bal"*/
                        else if (Reptype == 12 || Reptype == 67) { qryP = 'Pre_SalarySlipInclude.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly A4Size(4Rows) Salary Register Bold Caption"*/
                        else if (Reptype == 13) { qryP = 'pre_salaryregistera4size4rowBoldCaption.aspx?id=' + qryP; }
                        /*This statement is used for display "Register Of Payment Of Wages"*/
                        else if (Reptype == 14) { qryP = 'preSalaryRegisterWages.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly A4Size(5Rows) Salary Register (A2Z)"*/
                        else if (Reptype == 15) { qryP = 'Pre_SalaryRegisterA4Size5Row.aspx?id=' + qryP; }
                        /*For Display Salary Register A4 Size in 4Row with Arrear, This statement is used for display "Monthly A4Size(4Rows) Sal. Reg. With Arrear"*/
                        else if (Reptype == 16) { qryP = 'PreA4SalRegArr.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly Salary Slip With Tax Details"  or  for display "Tax computation Report"*/
                        else if (Reptype == 18) { qryP = 'PreTaxSlip.aspx?id=' + qryP; }
                        else if (Reptype == 52 || Reptype == 49 || Reptype == 51) { qryP = 'PreSalSlip.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly A4Size(5Rows) Salary Register with Arrear"*/
                        else if (Reptype == 19) { qryP = 'Pre_SalaryRegisterA4Size5RowArr.aspx?id=' + qryP; }
                        /*This statement is used for display "Year To Date Salary Slip"*/
                        else if (Reptype == 20) { qryP = 'PreSalSlipYTD.aspx?id=' + qryP; }
                        /*For Display "Monthly Salary Register New"*/
                        else if (Reptype == 22) { qryP = 'PreNewMnthSalReg.aspx?id=' + qryP; }
                        /*For Display "Monthly Salary slip (Nigeria)"*/
                        else if (Reptype == 23) { qryP = 'PreSalSlipNewInvestment.aspx?id=' + qryP; }
                        /*For Display "Year To Date Salary Slip(Nigeria)"*/
                        else if (Reptype == 24) { qryP = 'PreSalSlipYTD_Nis.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly salary register with seperately reimbursement "*/
                        else if (Reptype == 25) { qryP = 'PreSalaryRegister_Reimb.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly Salary Slip With Tax Details"*/
                        else if (Reptype == 26) { qryP = 'PreSalSlip_LoanDetail.aspx?id=' + qryP; }
                        //Added By geeta on 1 Jun 2012
                        /*This statement is used for display "Monthly Pay-Register"*/
                        else if (Reptype == 32) { qryP = 'prepayregister.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly Salary-Register Apprentice"*/
                        else if (Reptype == 33) { qryP = 'PreAppenticeReg.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly Salary-Register Casual"*/
                        else if (Reptype == 34) { qryP = 'precasualRegister.aspx?id=' + qryP; }
                        else if (Reptype == 37) { qryP = 'pre_SalarySlipKnya.aspx?id=' + qryP; }
                        else if (Reptype == 43) { qryP = 'Pre_WageSlipForEasySource.aspx?id=' + qryP; }
                        /*Add condition for "Monthly Final Salary" by Nisha on 31 Aug 2013*/
                        /*This statement is used for display "Monthly Final Salary"*/
                        else if (Reptype == 45) { qryP = 'preMonthlyFinalSalary.aspx?id=' + qryP; }
                        /*Add condition for "Monthly Salary Register For Luxur" by Rajesh on 04 oct 2013*/
                        else if (Reptype == 46) { qryP = 'Pre_SalaryRegister_Luxur.aspx?id=' + qryP; }
                        /*Add condition for "Monthly Salary Slip in Hindi" by Jay on 11 Mar 2014*/
                        else if (Reptype == 47 || Reptype == 48) { qryP = 'Pre_Salaryslipmiscellaneous.aspx?id=' + qryP; }
                        else if (Reptype == 53) { qryP = 'PreSalSlipPTC.aspx?id=' + qryP; }
                        /*This statement is used for display "Tax computation Report"*/
                        else if (Reptype == 55) { qryP = 'PreNewTdsEstimationSlip.aspx?id=' + qryP; }
                        /*This statement is used for display "Tax computation Report"*/
                        else if (Reptype == 57) { qryP = 'PreTaxSheetForeCast.aspx?id=' + qryP; }
                        else if (Reptype == 69) { qryP = 'PreNewTdsEstimationSlip.aspx?id=' + qryP; }
                        else if (Reptype == 58 || Reptype == 59 || Reptype == 60) { qryP = 'pre_SalarySlipTimecard.aspx?id=' + qryP; }
                        else if (Reptype == 62) { qryP = 'Pre_OffCycle.aspx?id=' + qryP; }
                        else if (Reptype == 63) { qryP = 'PreSalRegPDF.aspx?id=' + qryP; }
                        else if (Reptype == 64) { qryP = 'PreSalSlipPFA.aspx?id=' + qryP; }
                        else if (Reptype == 65) { qryP = 'PreSalSlipMNF.aspx?id=' + qryP; }
                        else if (Reptype == 66) { qryP = 'PreSalSlipMiddleEast.aspx?id=' + qryP; }
                        else if (Reptype == 68) { qryP = 'PreSalSlipTrainee.aspx?id=' + qryP; }
                        /*This statement is used for display "Monthly Pay Slip Attra"*/
                        else if (Reptype == 76 || Reptype == 77) { qryP = 'Pre_SalarySlipAttra.aspx?id=' + qryP; }


                        else if (Reptype == 74) { qryP = 'PreSlipArrDetails.aspx?id=' + qryP; }
                        window.open('reports/' + qryP, '', ' height=' + height + ' width=' + width + ' left=3' + ' menubar=1' + ' top=0' + ' scrollbars=1');
                        return false;
                    }

                    /*For generating the id of the check box*/
                    function generateIdString(_chkboxlistParentDivId, chkname) {
                        var sel = 0;
                        var names = [];
                        $('#dlist [id*=dls_]:checked').each(function () {
                            sel++
                            names.push($(this).val());
                        });
                        $('#hid' + chkname + 'id').val(names.join(','));
                        $('#hid' + chkname + 'selcount').val(sel);
                    }
                    function showReport() {
                        _repVal = $("#Hidden5").val();

                        if (_repVal == "63") {
                            if ($("#USearch_ddlunit").val() == "") {
                                alert('<%=_objCommon.DisplayCaption("UNT")%>' + " cannot be left blank!");
                                $("#USearch_ddlunit").focus();
                                return false;
                            }
                            if ($("#USearch_ddldesignation").val() == "") {
                                alert('<%=_objCommon.DisplayCaption("DES")%>' + " cannot be left blank!");
                                $("#USearch_ddldesignation").focus();
                                return false;
                            }
                        }

                        _qryP = $("#hidstring1").val();

                        if (_repVal == 'SL' || _repVal == '50' || _repVal == '56') { qryP = "Pre_SalarySlip.aspx?id=" + _qryP; }
                        else if (_repVal == "S") { qryP = "PreSlipWOLeaveDetails.aspx?id=" + _qryP; }
                        else if (_repVal == "67") { qryP = "Pre_SalarySlipInclude.aspx?id=" + _qryP; }
                        else if (_repVal == "T") { qryP = "PreNewTdsEstimationSlip.aspx?id=" + _qryP; }
                        else if (_repVal == "52" || _repVal == "49" || _repVal == "51") { qryP = "PreSalSlip.aspx?id=" + _qryP; }
                        else if (_repVal == "R") { qryP = "PreTaxSlip.aspx?id=" + _qryP; }
                        else if (_repVal == "74") { qryP = "PreSlipArrDetails.aspx?id=" + _qryP; }
                        else if (_repVal == "TL") { qryP = "PreSalSlip_LoanDetail.aspx?id=" + _qryP; }
                        else if (_repVal == "I") { qryP = "PreInvestmentDetails.aspx?id=" + _qryP; }
                        else if (_repVal == "SI") { qryP = "Pre_SalSlipwithINVEST.aspx?id=" + _qryP; }
                        else if (_repVal == "SH") { qryP = "Pre_MonthlySalarySlipHindiforpgl.aspx?id=" + _qryP; }
                        else if (_repVal == "RS") { qryP = "Pre_SalSlipWithReimb.aspx?id=" + _qryP; }
                        else if (_repVal == "YTD") { qryP = "PreSalSlipYtd.aspx?id=" + _qryP; }
                        else if (_repVal == "RN") { qryP = "Pre_SalSlipWithReimbNewFormat.aspx?id=" + _qryP; }
                        else if (_repVal == "43") { qryP = "Pre_WageSlipForEasySource.aspx?id=" + _qryP; }
                        else if (_repVal == "53") { qryP = "PreSalSlipPTC.aspx?id=" + _qryP; }
                        else if (_repVal == "55" || _repVal == "69") { qryP = "PreNewTdsEstimationSlip.aspx?id=" + _qryP; }
                        else if (_repVal == "57") { qryP = "PreTaxSheetForeCast.aspx?id=" + _qryP; }
                        else if (_repVal == "58" || _repVal == "59" || _repVal == "60") { qryP = "pre_SalarySlipTimecard.aspx?id=" + _qryP; }
                        else if (_repVal == "62") { qryP = "Pre_OffCycle.aspx?id=" + _qryP; }
                        else if (_repVal == "63") { qryP = 'PreSalRegPDF.aspx?id=' + _qryP; }
                        else if (_repVal == "64") { qryP = 'PreSalSlipPFA.aspx?id=' + _qryP; }
                        else if (_repVal == "65") { qryP = 'PreSalSlipMNF.aspx?id=' + _qryP; }
                        else if (_repVal == "66") { qryP = 'PreSalSlipMiddleEast.aspx?id=' + _qryP; }
                        else if (_repVal == "68") { qryP = 'PreSalSlipTrainee.aspx?id=' + _qryP; }
                        if (_repVal == "I") {
                            window.open('../EmpUser/' + qryP, '', 'width=1000, height=660,left=3,top=0,location=0, menubar=1, resizable=Yes, scrollbars=1');
                            return false;
                        }
                        else {
                            window.open('Reports/' + qryP, '', 'width=1000, height=660,left=3,top=0,location=0, menubar=1, resizable=Yes, scrollbars=1');
                            return false;
                        }
                    }
                    function CheckAllList(ctrl) {
                        checkUncheckCheckBoxlixtChkBox('dlist', ctrl, 'dls_');
                        var names = [];
                        $('#dlist [id*=dls_]:checked').each(function () {
                            names.push($(this).val());
                        });
                        $('#hiddls_id').val(names.join(','));
                    }
                    function rbtnChange(_ctrl) {

                        $('#USearch_ddlunit option').remove();
                        $("<option value='6'>-- Select Unit --</option>").appendTo("#USearch_ddlunit");
                        $('#TdSearch [id^=USearch]').val('');
                        $('#lblmsg,#lit,#lblmsg2').text("");
                        $('#ddllEncrType').attr('selectedIndex', '0');
                        $('#ddlrepformat').attr('selectedIndex', '0');
                        $('#txtrptName').val("");
                        $('#chkSFTP').attr('checked', false);
                        $('#trEncrType,#trSftpID,#trFileName,#trformat').css("display", "none");

                        if (_ctrl == 'R') {
                            $('#tblsp,#tblSh,#tblpwd,#TrNoSearch,#tremail,#tableshow,#tblrepin,#trsortbasis,#trGroupBY,#TRDIV,#tblothepaycode,#tblsection,#tblpaycode,#TblReimb,#trrepformat,#trGroupBY,#divEmail,#trselall,#trShowClr,#divSocial').css('display', 'none');
                            $('#DDLPaySlipType').attr('selectedIndex', '0');
                            $('#DdlreportType').attr('selectedIndex', '0');
                            $('#SlipRegPre,#divgenerate').css("display", "")
                            //stopProgressBarTDS();
                        }
                        else {
                            $('#SlipRegPre,#trrepformat,#trGroupBY,#tblsection,#tblpaycode,#TblReimb,#trview,#tableshow,#tblothepaycode,#trsortbasis,#trGroupBY,#TRDIV,#divgenerate,#LnkPDF,#BtnPreviewdivActive,#trselall,#trMerge,#trShowClr').css('display', 'none');
                            $('#tblsp,#tblSh,#tblpwd,#TrNoSearch,#tremail,#tblrepin,#divEmail').css('display', '');
                            $('#DdlreportType').attr('selectedIndex', '0');
                            $('#DDLPaySlipType').attr('selectedIndex', '0');
                            /*Added by Rohtas Singh on 14 Feb 2018*/
                            $('#chkMerge').attr('checked', false);
                        }
                        ShowPublish();
                        $("#btnPreview").val("Preview");
                        $('#ddlRepIn,#chkemailrepmanager,#DdlreportType,#ddlSalaryProcess,#ddlEMailExist,#ddlEMailSend,#ddlSalWithHeld,#ddlEmpPass,#USearch_ddlcostcenter').removeAttr('disabled')
                        $('#chkmailformat,#chkemailrepmanager').attr('checked', false);
                        $('#txtccc,#txtBCC,#hidstatus').val("")
                        $('#ddlMonthYear').attr('selectedIndex', $('#ddlMonthYear option').length - 1)
                        $('#DdlreportType,#ddlshowsal,#ddlSalaryProcess,#ddlEMailExist,#ddlEMailSend,#ddlEmpPass,#ddlRepIn').attr('selectedIndex', 0)
                        $('#troffcycledt').css("display", "none");
                    }
                    function ShowPublish() {
                        $('#trMergeMsg').css("display", "none");
                        if ($('#ddlRepIn').val() == "H") {
                            $('#TrNoSearch').css("display", "none");
                            $('#TdSearch1').css('display', '');
                            $('#TdSearch2').css('display', '');
                            $('#RblNoSearch input:checked').val('P');
                        }
                        else {
                            $('#TrNoSearch').css('display', '');
                            $('#TdSearch1').css('display', '');
                            $('#TdSearch2').css('display', '');
                            $('#RblNoSearch input:checked').val('P');
                        }
                    }
                    function ShowRepDet() {
                        /*Added by Quadir on 14 OCT 2020 */
                        $('#TrSlipPubMode').css("display", "none");

                        if ($('#DdlreportType').val() == "SL") {
                            if ($('#ddlRepIn').val() == "L") {
                                $('#trRepEmail').css("display", "none");
                                $('#chkemailrepmanager').attr('checked', false);
                            }
                            else { $('#trRepEmail').css("display", ""); }
                        }
                        if ($('#ddlRepIn').val() == "P") {

                            if ($('#DdlreportType').val() == "R") {
                                $('#tblpwd,#TrNoSearch,#trselall').css("display", "");
                                $('#TrSlipPubMode').css("display", "");
                            }
                            else {

                                $('#tblpwd,#TrNoSearch,#trselall').css("display", "");
                            }
                        }
                        else if ($('#ddlRepIn').val() == "") { $('#tblpwd,#TrNoSearch,#trselall,#TrSlipPubMode').css("display", "none"); }
                        else {
                            $('#tblpwd,#TrNoSearch').css("display", "none");
                            $('#trselall').css("display", $('#ddlRepIn').val() == "H" ? "none" : "");
                        }
                        $('#trrepformat').css("display", "none");
                        PleaseWaitWithDailog();
                    }
                    function ResetCtrl() {
                        $('#tableshow').css("display", "none")
                        $('#tblsp,#tblSh,#tblpwd,#TrNoSearch,#tremail,#tblrepin').css("display", "")
                        $('#USearch_ddlunit option').remove();
                        $("<option value='6'>-- Select Unit --</option>").appendTo("#USearch_ddlunit");
                        $('#TdSearch [id^=USearch]').val('');
                        $('#ddlRepIn,#chkemailrepmanager,#DdlreportType,#ddlSalaryProcess,#ddlEMailExist,#ddlEMailSend,#ddlSalWithHeld,#ddlEmpPass,#USearch_ddlcostcenter,#ddlmonthyearS').removeAttr('disabled')
                        $('#chkmailformat,#chkemailrepmanager').attr('checked', false);
                        $('#txtccc,#txtBCC,#hidstatus').val("")
                        $('#lblmsg').text("");
                        $('#ddlMonthYear').attr('selectedIndex', $('#ddlMonthYear option').length - 1)
                        $('#DdlreportType,#ddlshowsal,#ddlSalaryProcess,#ddlEMailExist,#ddlEMailSend,#ddlEmpPass,#ddlRepIn').attr('selectedIndex', 0)
                        $('#download_pdf1').hide();
                        $('#download_pdf2').hide();
                        $("#lblMsgSlip").html("")
                        $("#lblMailMsgWOPWD").html("")
                        PleaseWaitWithDailog();
                    }

                    function toggle() {
                        $('#trsetting').toggle();
                        if ($('#trsetting').css('display') == 'none') {
                            $('#trgap').css('display', '');
                        }
                        else {
                            $('#trgap').css('display', 'none');
                        }

                    }

                    function ShowTaxForcast() {
                        var height = window.screen.availHeight;
                        var width = window.screen.availWidth;
                        var qryP = $("#HidPreVal").val();
                        var AppPath = $("#HidAppPath").val();
                        var path = $("#HidPath").val();
                        var yyyy = $("#HidYear").val();
                        var mm = $('#ddlMonthYear').val();
                        var prefixpdfname;
                        if (mm.length > 1) {
                            prefixpdfname = yyyy + mm;
                        }
                        else {
                            prefixpdfname = yyyy + "0" + mm;
                        }
                        var page = 'reports/PreTaxSheetForeCast.aspx?id=' + qryP;
                        //OpenTaxForcastProgress(AppPath, page, path, "FORCAST",prefixpdfname);
                        OpenPaySlipProgressbar(AppPath, page, path, "FORCAST");

                    }

                    function ShowSlipWOLeave() {
                        var height = window.screen.availHeight;
                        var width = window.screen.availWidth;
                        var qryP = $("#HidPreVal").val();
                        var AppPath = $("#HidAppPath").val();
                        var path = $("#HidPath").val();
                        var page = 'reports/PreSlipWOLeaveDetails.aspx?id=' + qryP;
                        //OpenSlipWOLeaveProgress(AppPath, page, path, "SLIPWOLVE");
                        OpenPaySlipProgressbar(AppPath, page, path, "SLIPWOLVE");

                    }

                    function ShowTaxDetails() {
                        var height = window.screen.availHeight;
                        var width = window.screen.availWidth;
                        var qryP = $("#HidPreVal").val();
                        var AppPath = $("#HidAppPath").val();
                        var path = $("#HidPath").val();
                        var process_status_id = $("#process_status_id").val();
                        var page = 'reports/PreTaxSlip.aspx?id=' + qryP;
                        OpenPaySlipProgressbar(AppPath, page, path, "TAXSLIP", process_status_id);
                        //OpenSlipProgress(AppPath, page, path, "TAXSLIP");
                    }
                    function ShowTdsDetails() {
                        var qryPv = $("#HidPreVal").val();
                        var AppPath = $("#HidAppPath").val();
                        var path = $("#HidPath").val();
                        var page = 'reports/PreTdsEstimationSlip4Pdf.aspx?id=' + qryPv;
                        OpenPaySlipProgressbar(AppPath, page, path, "SLIPTDSV");
                    }
                    function ShowPayslipsLockSummaryDetails(processType) {
                        var AppPath = $("#HidAppPath").val();
                        RefreshProgressBarLockedStatus(AppPath, processType);
                    }
                    function ShowExcelLockSummaryDetails(processType) {
                        console.log('processType: ' + processType);
                        var AppPath = $("#HidAppPath").val();
                        RefreshExcelProcessbarLockedStatus(AppPath, processType);
                    }

                    function ShowTaxDetailsWOPWD() {
                        var height = window.screen.availHeight;
                        var width = window.screen.availWidth;
                        var qryP = $("#HidPreVal").val();
                        var AppPath = $("#HidAppPath").val();
                        var path = $("#HidPath").val();
                        var page = 'reports/PreTaxSlip.aspx?id=' + qryP;
                        var process_id = $("#process_status_id").val();
                        if (process_id) {
                            OpenPaySlipProgressbar(AppPath, page, path, "TAXSLIP");
                        } else {
                            OpenSlipProgressWD(AppPath, page, path, "TAXSLIP");
                        }


                    }
                    function ShowWithLeaveDetails() {
                        $("#LnkPDF").hide();
                        var qryPv = $("#HidPreVal").val();
                        var AppPath = $("#HidAppPath").val();
                        var path = $("#HidPath").val();
                        var page = 'reports/Pre_SalarySlip.aspx?id=' + qryPv;
                        OpenPaySlipProgressbar(AppPath, page, path, "SLIPWILVE");
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
                            success: function (res) {
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
                    function disablebtn45sec(button) {
                        button.disabled = true; // Disable the button
                        setTimeout(function () {
                            button.disabled = false; // Re-enable the button after 5 seconds
                        }, 5000);
                    }
                    // Added by Kangkan to download from Java service
                    function downloadPdf(e) {
                        e.preventDefault();

                        let id = document.getElementById("process_status_id").value;
                        let javaUrl = $("#java_url").val();
                        let companyCode = $("#companyCode").val();

                        const iframe = document.createElement("iframe");
                        iframe.name = "downloadIframe";
                        iframe.style.display = "none";
                        document.body.appendChild(iframe);

                        const form = document.createElement("form");
                        form.method = "POST";
                        form.action = javaUrl;
                        form.target = "downloadIframe";

                        const processID = document.createElement("input");
                        processID.type = "hidden";
                        processID.name = "processID";
                        processID.value = id && id.length > 0 ? id : process_status_id;
                        form.appendChild(processID);

                        const inputCompanyCode = document.createElement("input");
                        inputCompanyCode.type = "hidden";
                        inputCompanyCode.name = "companyCode";
                        inputCompanyCode.value = companyCode;
                        form.appendChild(inputCompanyCode);

                        document.body.appendChild(form);
                        form.submit();

                        setTimeout(() => {
                            document.body.removeChild(form);
                            document.body.removeChild(iframe);
                        }, 3000);
                    }


                    function openDownloadWindow(url, filePath, compCode, fileExt, slipType, mm, yyyy) {
                        const iframe = document.createElement("iframe");
                        iframe.name = "downloadIframe";
                        iframe.style.display = "none";
                        document.body.appendChild(iframe);

                        const form = document.createElement("form");
                        form.method = "POST";
                        form.action = url;
                        form.target = "downloadIframe";

                        const fields = {
                            empCode: filePath,
                            companyCode: compCode,
                            fileExt: fileExt,
                            slipType: slipType,
                            mm: mm,
                            yyyy: yyyy
                        };

                        for (const key in fields) {
                            const input = document.createElement("input");
                            input.type = "hidden";
                            input.name = key;
                            input.value = fields[key];
                            form.appendChild(input);
                        }

                        document.body.appendChild(form);
                        form.submit();

                        setTimeout(() => {
                            document.body.removeChild(form);
                            document.body.removeChild(iframe);
                        }, 3000);

                    }

                </script>

                <body bottommargin="0" bgcolor="#f7fcff" leftmargin="0" topmargin="0" rightmargin="0" marginheight="0"
                    marginwidth="0">
                    <%-- Payslips Progress bar div --%>
                        <div id="dlg" style="font-size: 1.1em; padding: 20px; display: none" title="Progress">
                            <div id="statusMessage" style="font-weight: bold; padding: 5px 0px">
                            </div>
                            <div style="display: none;">
                                <iframe id="Slipframe"
                                    style="border: 0px; border-spacing: 0px; border-collapse: collapse;"
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
                                        style="font-weight: bold"></span>
                                </div>
                            </div>
                            <div id="summaryWrapper" style="display: none; line-height: 30px; padding-left: 10px;">
                                <div style="display: none">
                                    <span>Total employees whose slip was published: </span><span id="TotaltobeProcssed"
                                        style="font-weight: bold"></span>
                                </div>
                                <div>
                                    <span>Total employees whose slips successfully published : </span><span
                                        id="totalProcessedSlips" style="font-weight: bold"></span>
                                </div>
                                <div>
                                    <span>Total employees whose slip not published due to some error: : </span><span
                                        id="totalUnprocessedSlips" style="font-weight: bold"></span>
                                </div>
                                <div>
                                    <span>Processed in : </span><span id="totalTimeTaken"
                                        style="font-weight: bold"></span>
                                </div>
                                <div id="summaryError"
                                    style="display: none; line-height: 15px; height: 80px; overflow: auto">
                                    <span id="errorSummary" style="color: Red;"></span>
                                </div>
                            </div>
                        </div>
                        <%-- Excel Progress bar div --%>
                            <div id="CommonProgressBarModelElement" class="alt-module-true" style="display: none;">
                                <div class="alt-modal dialog-500">
                                    <div class="alt-modal-title">
                                        <span id="CommonProgressBarTitle">Progress</span>
                                    </div>
                                    <div class="alt-modal-body" id="CommonProgressBarBody" style="display: block;">
                                    </div>
                                    <div class="alt-modal-body" id="CommonProgressBarStatusWrapper"
                                        style="display: block;">
                                        <div class="progressbar-outer">
                                            <div id="progressBarExcel" class="progressbar">
                                            </div>
                                        </div>
                                        <div class="clearfix">
                                            <span class="FL"><span class="blue-color"
                                                    id="totalProcessedExcel"></span><span class="blue-color"
                                                    id="totalToProcessExcel"></span></span><span class="FR">
                                                <span class="blue-color" id="estimatedTimeLeftExcel"></span></span>
                                        </div>
                                    </div>
                                    <div class="alt-modal-body" id="ErrorWrapper" style="display: none;">
                                        <div class="progressbar-outer">
                                            <div class="progressbarerr" style="width: 100%">
                                            </div>
                                        </div>
                                        <p class="p-lines">
                                            <span id="spnerrmsg">Error in Processing, Please connect with App Support
                                            </span>
                                        </p>
                                    </div>
                                    <div id="CommonProgressBarFooter" class="alt-modal-body">
                                        <button id="CommonProgressBarCloseBtn" class="Btn"
                                            style="height: 20px; width: 55px; font-family: Verdana,Arial,sans-serif"
                                            onclick="hideModal();">Close</button>
                                    </div>

                                </div>
                            </div>

                            <form id="form1" method="post" runat="server">
                                <uc1:AdminMenu ID="AdminMenu1" runat="server" />
                                <table cellspacing="0" cellpadding="0" width="100%" border="0">
                                    <tr>
                                        <td>
                                            <table cellspacing="0" cellpadding="0" width="100%" border="0">
                                                <tr>
                                                    <td>
                                                        &nbsp;
                                                    </td>
                                                    <!--<td background="ImagesCA/top1.jpg" width="100%">-->
                                                </tr>
                                                <tr>
                                                    <td width="100%">
                                                        <!--This is Internal Body Table-->
                                                        <table cellspacing="0" cellpadding="0" width="97%"
                                                            align="center" border="0">
                                                            <tr>
                                                                <td align="center">
                                                                    <asp:ScriptManager ID="ScriptManager2"
                                                                        runat="server">
                                                                    </asp:ScriptManager>
                                                                    <fieldset class="fieldset">
                                                                        <legend class="legend">Monthly Salary Slip
                                                                            Generation</legend>
                                                                        <table cellspacing="0" cellpadding="0"
                                                                            width="100%" border="0">
                                                                            <tr>
                                                                                <td align="left">
                                                                                    <table cellspacing="0"
                                                                                        cellpadding="0" width="96%"
                                                                                        align="center" border="0">
                                                                                        <tr>
                                                                                            <td class="tdcaption">
                                                                                                <asp:RadioButton
                                                                                                    ID="rbtnslip"
                                                                                                    runat="server"
                                                                                                    CssClass="radio"
                                                                                                    Text="Salary Slip Report"
                                                                                                    GroupName="check"
                                                                                                    Onclick="rbtnChange('R')"
                                                                                                    AutoPostBack="false" />
                                                                                                <asp:RadioButton
                                                                                                    ID="rbtnmail"
                                                                                                    runat="server"
                                                                                                    CssClass="radio"
                                                                                                    Text="E-Mail Payslip,TDS-Estimation Slip"
                                                                                                    Onclick="rbtnChange('M')"
                                                                                                    GroupName="check"
                                                                                                    AutoPostBack="false" />
                                                                                                <input id="hidstatus"
                                                                                                    type="hidden"
                                                                                                    name="hidstatus"
                                                                                                    runat="server" />
                                                                                            </td>
                                                                                            <td class="tdcaption"
                                                                                                colspan="2">
                                                                                                <%--<asp:UpdatePanel
                                                                                                    ID="upStatus"
                                                                                                    runat="server">
                                                                                                    <ContentTemplate>
                                                                                                        --%>
                                                                                                        <table
                                                                                                            cellspacing="0"
                                                                                                            cellpadding="0"
                                                                                                            width="98%"
                                                                                                            align="center"
                                                                                                            border="0"
                                                                                                            id="divSocial"
                                                                                                            runat="server"
                                                                                                            style="display: none;">
                                                                                                            <tr>
                                                                                                                <td
                                                                                                                    align="right">
                                                                                                                    <img src="Images/icon3.png"
                                                                                                                        id="Img1"
                                                                                                                        style="vertical-align: middle;" />
                                                                                                                    <%--<asp:Label
                                                                                                                        ID="lblProcessStatus"
                                                                                                                        runat="server"
                                                                                                                        CssClass="UserStatusMsg"
                                                                                                                        Text="">
                                                                                                                        </asp:Label>
                                                                                                                        --%>
                                                                                                                        <asp:Label
                                                                                                                            ID="lblProcessStatus"
                                                                                                                            runat="server"
                                                                                                                            CssClass="UserStatusMsg"
                                                                                                                            Text=""
                                                                                                                            style="color: red">
                                                                                                                        </asp:Label>
                                                                                                                        <asp:ImageButton
                                                                                                                            ID="imgpdfprocess"
                                                                                                                            runat="server"
                                                                                                                            Style="vertical-align: middle;"
                                                                                                                            Width="25px"
                                                                                                                            ToolTip="Refresh"
                                                                                                                            AlternateText="Refresh"
                                                                                                                            ImageUrl="img/refreshtds.gif">
                                                                                                                        </asp:ImageButton>
                                                                                                                        <%--<asp:ImageButton
                                                                                                                            ID="imgpdfprocess"
                                                                                                                            runat="server"
                                                                                                                            Visible="true"
                                                                                                                            Style="vertical-align: middle;"
                                                                                                                            ToolTip="Refresh"
                                                                                                                            AlternateText="Refresh"
                                                                                                                            ImageUrl="img/refresh.gif">
                                                                                                                            </asp:ImageButton>
                                                                                                                            --%>
                                                                                                                            <asp:HiddenField
                                                                                                                                ID="hdnusername"
                                                                                                                                runat="server"
                                                                                                                                Value="User" />
                                                                                                                            <asp:HiddenField
                                                                                                                                ID="IsShowlnkIcon"
                                                                                                                                runat="server"
                                                                                                                                Value="N" />
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                        </table>
                                                                                                        <%--
                                                                                                            </ContentTemplate>
                                                                                                            </asp:UpdatePanel>
                                                                                                            --%>
                                                                                            </td>
                                                                                        </tr>
                                                                                        <tr id="divSocialExcel"
                                                                                            runat="server"
                                                                                            visible="false">
                                                                                            <td colspan="3">
                                                                                                <table cellspacing="0"
                                                                                                    cellpadding="0"
                                                                                                    width="98%"
                                                                                                    align="center"
                                                                                                    border="0">
                                                                                                    <tr>
                                                                                                        <td
                                                                                                            align="right">
                                                                                                            <img src="Images/icon3.png"
                                                                                                                id="ImgExcel"
                                                                                                                style="vertical-align: middle;" />
                                                                                                            <asp:Label
                                                                                                                ID="lblProcessStatusExcel"
                                                                                                                runat="server"
                                                                                                                CssClass="UserStatusMsg"
                                                                                                                Text="">
                                                                                                            </asp:Label>
                                                                                                            <%--<asp:Button
                                                                                                                ID="btnProgressbarExcel"
                                                                                                                runat="server"
                                                                                                                Text="click here to check progress"
                                                                                                                CssClass="CustomBtn"
                                                                                                                Width="175px"
                                                                                                                CausesValidation="False"
                                                                                                                OnClientClick="return PleaseWaitWithDailog();" />--%>
                                                                                                            <asp:ImageButton
                                                                                                                ID="btnProgressbarExcel"
                                                                                                                runat="server"
                                                                                                                Style="vertical-align: middle;"
                                                                                                                Width="25px"
                                                                                                                ToolTip="click here to check progress"
                                                                                                                AlternateText="Refresh"
                                                                                                                ImageUrl="img/refreshtds.gif"
                                                                                                                OnClientClick="return PleaseWaitWithDailog();">
                                                                                                            </asp:ImageButton>
                                                                                                        </td>
                                                                                                    </tr>
                                                                                                </table>
                                                                                            </td>
                                                                                        </tr>
                                                                                        <tr runat="server" id="Tr1">
                                                                                            <td class="trupbtn"
                                                                                                colspan="3">
                                                                                            </td>
                                                                                        </tr>
                                                                                        <tr>
                                                                                            <td colspan="3"
                                                                                                id="TdSearch">
                                                                                                <UC:UserControl
                                                                                                    ID="USearch"
                                                                                                    runat="server" />
                                                                                            </td>
                                                                                        </tr>
                                                                                        <tr>
                                                                                            <td colspan="3">
                                                                                                <%-- <asp:UpdatePanel
                                                                                                    ID="UpdatePanel2"
                                                                                                    runat="server">
                                                                                                    <ContentTemplate>
                                                                                                        --%>
                                                                                                        <table
                                                                                                            border="0"
                                                                                                            cellpadding="0"
                                                                                                            cellspacing="0"
                                                                                                            width="100%">
                                                                                                            <tr>
                                                                                                                <td class="tdcaption"
                                                                                                                    width="12%">
                                                                                                                    Month
                                                                                                                    Year
                                                                                                                </td>
                                                                                                                <td
                                                                                                                    class="tdcolon">
                                                                                                                    :
                                                                                                                </td>
                                                                                                                <td
                                                                                                                    width="31%">
                                                                                                                    <asp:DropDownList
                                                                                                                        ID="ddlMonthYear"
                                                                                                                        runat="server"
                                                                                                                        CssClass="dropdownlist"
                                                                                                                        OnChange="popluateoffcycledate();">
                                                                                                                    </asp:DropDownList>
                                                                                                                </td>
                                                                                                                <td class="tdcaption"
                                                                                                                    width="10%">
                                                                                                                    Hold-Type
                                                                                                                </td>
                                                                                                                <td class="tdcolon"
                                                                                                                    align="center">
                                                                                                                    :
                                                                                                                </td>
                                                                                                                <td>
                                                                                                                    <asp:DropDownList
                                                                                                                        ID="ddlshowsal"
                                                                                                                        runat="server"
                                                                                                                        CssClass="Dropdownlist">
                                                                                                                        <asp:ListItem
                                                                                                                            Value="A">
                                                                                                                            All
                                                                                                                        </asp:ListItem>
                                                                                                                        <asp:ListItem
                                                                                                                            Value="H">
                                                                                                                            Only
                                                                                                                            Withheld
                                                                                                                        </asp:ListItem>
                                                                                                                        <asp:ListItem
                                                                                                                            Value="N">
                                                                                                                            Without
                                                                                                                            Withheld
                                                                                                                        </asp:ListItem>
                                                                                                                    </asp:DropDownList>
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                        </table>
                                                                                                        <%--
                                                                                                            </ContentTemplate>
                                                                                                            <Triggers>
                                                                                                                <asp:PostBackTrigger
                                                                                                                    ControlID="btnPreview" />
                                                                                                            </Triggers>
                                                                                                            </asp:UpdatePanel>
                                                                                                            --%>
                                                                                            </td>
                                                                                        </tr>
                                                                                        <tr id="SlipRegPre"
                                                                                            runat="server">
                                                                                            <td align="left"
                                                                                                colspan="3">
                                                                                                <asp:UpdatePanel
                                                                                                    ID="UpdatePanel3"
                                                                                                    runat="server">
                                                                                                    <ContentTemplate>
                                                                                                        <table
                                                                                                            cellspacing="0"
                                                                                                            cellpadding="0"
                                                                                                            width="100%"
                                                                                                            align="center"
                                                                                                            border="0">
                                                                                                            <tr>
                                                                                                                <td class="tdcaption"
                                                                                                                    width="12%">
                                                                                                                    Report
                                                                                                                    Type
                                                                                                                </td>
                                                                                                                <td
                                                                                                                    class="tdcolon">
                                                                                                                    :
                                                                                                                </td>
                                                                                                                <td
                                                                                                                    class="tdvd">
                                                                                                                    <asp:DropDownList
                                                                                                                        ID="DDLPaySlipType"
                                                                                                                        runat="server"
                                                                                                                        CssClass="dropdownlist"
                                                                                                                        AutoPostBack="true">
                                                                                                                    </asp:DropDownList>
                                                                                                                    *
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                            <tr id="trShowClr"
                                                                                                                runat="server"
                                                                                                                style="display: none">
                                                                                                                <td class="tdcaption"
                                                                                                                    width="12%">
                                                                                                                    Show
                                                                                                                    color
                                                                                                                    on
                                                                                                                    excel
                                                                                                                </td>
                                                                                                                <td
                                                                                                                    class="tdcolon">
                                                                                                                    :
                                                                                                                </td>
                                                                                                                <td class="TDVd"
                                                                                                                    valign="top">
                                                                                                                    <asp:RadioButtonList
                                                                                                                        ID="rbtshowclr"
                                                                                                                        runat="server"
                                                                                                                        CssClass="chkList"
                                                                                                                        RepeatColumns="2">
                                                                                                                        <asp:ListItem
                                                                                                                            Text="No"
                                                                                                                            Value="N"
                                                                                                                            Selected="True">
                                                                                                                        </asp:ListItem>
                                                                                                                        <asp:ListItem
                                                                                                                            Text="Yes"
                                                                                                                            Value="Y">
                                                                                                                        </asp:ListItem>
                                                                                                                    </asp:RadioButtonList>
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                            <tr id="troffcycle"
                                                                                                                runat="server"
                                                                                                                style="display: none;">
                                                                                                                <td class="tdcaption"
                                                                                                                    width="12%">
                                                                                                                    Select
                                                                                                                    Offcycle
                                                                                                                    Date
                                                                                                                </td>
                                                                                                                <td
                                                                                                                    class="tdcolon">
                                                                                                                    :
                                                                                                                </td>
                                                                                                                <td
                                                                                                                    class="tdvd">
                                                                                                                    <asp:DropDownList
                                                                                                                        ID="ddlmonthyearS"
                                                                                                                        runat="server"
                                                                                                                        CssClass="dropdownlist">
                                                                                                                    </asp:DropDownList>
                                                                                                                    *
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                            <tr id="trHelp"
                                                                                                                runat="server"
                                                                                                                style="display: none;">
                                                                                                                <td class="tdcaption"
                                                                                                                    width="12%">
                                                                                                                    Show
                                                                                                                    Pay
                                                                                                                    Head
                                                                                                                    Help
                                                                                                                </td>
                                                                                                                <td
                                                                                                                    class="tdcolon">
                                                                                                                    :
                                                                                                                </td>
                                                                                                                <td
                                                                                                                    class="tdvd">
                                                                                                                    <asp:CheckBox
                                                                                                                        runat="server"
                                                                                                                        ID="chkHelp"
                                                                                                                        CssClass="Checkbox" />
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                            <tr id="trrepformat"
                                                                                                                runat="server"
                                                                                                                style="display: none">
                                                                                                                <td
                                                                                                                    class="Tdcaption">
                                                                                                                    Report
                                                                                                                    Format
                                                                                                                </td>
                                                                                                                <td
                                                                                                                    class="TdColon">
                                                                                                                    :
                                                                                                                </td>
                                                                                                                <td class="tdvd"
                                                                                                                    colspan="5">
                                                                                                                    <asp:DropDownList
                                                                                                                        ID="ddlformat"
                                                                                                                        runat="server"
                                                                                                                        CssClass="dropdownlist">
                                                                                                                    </asp:DropDownList>
                                                                                                                    *
                                                                                                                    <asp:RadioButton
                                                                                                                        ID="rbHorizontal"
                                                                                                                        runat="server"
                                                                                                                        CssClass="Radio"
                                                                                                                        Checked="True"
                                                                                                                        GroupName="Group"
                                                                                                                        Text="Horizontal Format">
                                                                                                                    </asp:RadioButton>
                                                                                                                    <asp:RadioButton
                                                                                                                        ID="RbVertical"
                                                                                                                        runat="server"
                                                                                                                        CssClass="Radio"
                                                                                                                        GroupName="Group"
                                                                                                                        Text="Vertical Format">
                                                                                                                    </asp:RadioButton>
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                            <tr runat="server"
                                                                                                                id="trGroupBY"
                                                                                                                style="display: none">
                                                                                                                <td class="Tdcaption"
                                                                                                                    valign="top"
                                                                                                                    width="12%">
                                                                                                                    Group
                                                                                                                    By
                                                                                                                </td>
                                                                                                                <td class="TdColon"
                                                                                                                    valign="top">
                                                                                                                    :
                                                                                                                </td>
                                                                                                                <td
                                                                                                                    width="87%">
                                                                                                                    <table
                                                                                                                        border="0"
                                                                                                                        width="80%"
                                                                                                                        cellpadding="0"
                                                                                                                        cellspacing="0">
                                                                                                                        <tr>
                                                                                                                            <td>
                                                                                                                                <table
                                                                                                                                    border="0"
                                                                                                                                    width="33%"
                                                                                                                                    cellpadding="0"
                                                                                                                                    cellspacing="0"
                                                                                                                                    id="trsortbasis"
                                                                                                                                    runat="server"
                                                                                                                                    style="display: none">
                                                                                                                                    <tr>
                                                                                                                                        <td class="Tdcaption"
                                                                                                                                            valign="top"
                                                                                                                                            width="12%">
                                                                                                                                            Group
                                                                                                                                            By(1)
                                                                                                                                        </td>
                                                                                                                                        <td class="TdColon"
                                                                                                                                            valign="top"
                                                                                                                                            width="2%">
                                                                                                                                            :
                                                                                                                                        </td>
                                                                                                                                        <td class="tdvd"
                                                                                                                                            valign="top"
                                                                                                                                            colspan="5">
                                                                                                                                            <asp:DropDownList
                                                                                                                                                ID="ddlshortbasis"
                                                                                                                                                runat="server"
                                                                                                                                                CssClass="DropdownList"
                                                                                                                                                AutoPostBack="True">
                                                                                                                                            </asp:DropDownList>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr runat="server"
                                                                                                                                        id="TRDIV">
                                                                                                                                        <td class="Tdcaption"
                                                                                                                                            valign="top">
                                                                                                                                            List
                                                                                                                                        </td>
                                                                                                                                        <td class="TdColon"
                                                                                                                                            valign="top">
                                                                                                                                            :
                                                                                                                                        </td>
                                                                                                                                        <td class="tdcaption"
                                                                                                                                            valign="top">
                                                                                                                                            <div class="tdcolon"
                                                                                                                                                id="dlist"
                                                                                                                                                style="border-right: #003399 1px solid; border-top: #003399 1px solid;
                                                                                                                    overflow: auto; border-left: #003399 1px solid; width: 200px; border-bottom: #003399 1px solid;
                                                                                                                    height: 95px">
                                                                                                                                                <%=getdatalist()%>
                                                                                                                                            </div>
                                                                                                                                            <input
                                                                                                                                                id="hiddls_id"
                                                                                                                                                type="hidden"
                                                                                                                                                name="hiddls_id"
                                                                                                                                                runat="server" />
                                                                                                                                            <input
                                                                                                                                                id="hidShortVal"
                                                                                                                                                type="hidden"
                                                                                                                                                name="hidShortVal"
                                                                                                                                                runat="server" />
                                                                                                                                            <input
                                                                                                                                                id="hiddls_count"
                                                                                                                                                type="hidden"
                                                                                                                                                name="hiddls_count"
                                                                                                                                                runat="server" />
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                </table>
                                                                                                                            </td>
                                                                                                                            <td>
                                                                                                                                <table
                                                                                                                                    border="0"
                                                                                                                                    width="33%"
                                                                                                                                    cellpadding="0"
                                                                                                                                    cellspacing="0"
                                                                                                                                    id="trsortbasis2"
                                                                                                                                    runat="server"
                                                                                                                                    style="display: none">
                                                                                                                                    <tr>
                                                                                                                                        <td class="Tdcaption"
                                                                                                                                            valign="top"
                                                                                                                                            width="12%">
                                                                                                                                            Group
                                                                                                                                            By(2)
                                                                                                                                        </td>
                                                                                                                                        <td class="TdColon"
                                                                                                                                            valign="top"
                                                                                                                                            width="2%">
                                                                                                                                            :
                                                                                                                                        </td>
                                                                                                                                        <td class="tdvd"
                                                                                                                                            valign="top"
                                                                                                                                            colspan="5">
                                                                                                                                            <asp:DropDownList
                                                                                                                                                ID="ddlGroup2"
                                                                                                                                                runat="server"
                                                                                                                                                CssClass="DropdownList"
                                                                                                                                                AutoPostBack="True">
                                                                                                                                            </asp:DropDownList>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td class="Tdcaption"
                                                                                                                                            valign="top">
                                                                                                                                            List
                                                                                                                                        </td>
                                                                                                                                        <td class="TdColon"
                                                                                                                                            valign="top">
                                                                                                                                            :
                                                                                                                                        </td>
                                                                                                                                        <td colspan="5"
                                                                                                                                            class="tdcaption"
                                                                                                                                            valign="top">
                                                                                                                                            <div class="tdcolon"
                                                                                                                                                id="Div1"
                                                                                                                                                style="border-right: #003399 1px solid; border-top: #003399 1px solid;
                                                                                                                    overflow: auto; border-left: #003399 1px solid; width: 200px; border-bottom: #003399 1px solid;
                                                                                                                    height: 95px">
                                                                                                                                                <table
                                                                                                                                                    width="100%"
                                                                                                                                                    border="0"
                                                                                                                                                    cellpadding="0"
                                                                                                                                                    cellspacing="0"
                                                                                                                                                    id="Group2">
                                                                                                                                                    <tr>
                                                                                                                                                        <td align="left"
                                                                                                                                                            class="TDCaptionBold">
                                                                                                                                                            <asp:CheckBox
                                                                                                                                                                ID="chkAllGr1"
                                                                                                                                                                runat="server"
                                                                                                                                                                Checked="true"
                                                                                                                                                                onclick="checkUncheckCheckBoxlixtChkBox('Div1' ,this , 'chkListGroup1');"
                                                                                                                                                                CssClass="checkbox" />
                                                                                                                                                            Select
                                                                                                                                                            All
                                                                                                                                                        </td>
                                                                                                                                                    </tr>
                                                                                                                                                    <tr>
                                                                                                                                                        <td>
                                                                                                                                                            <asp:CheckBoxList
                                                                                                                                                                ID="chkListGroup1"
                                                                                                                                                                runat="server"
                                                                                                                                                                RepeatDirection="Horizontal"
                                                                                                                                                                RepeatColumns="1"
                                                                                                                                                                CssClass="chkList">
                                                                                                                                                            </asp:CheckBoxList>
                                                                                                                                                        </td>
                                                                                                                                                    </tr>
                                                                                                                                                </table>
                                                                                                                                            </div>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                </table>
                                                                                                                            </td>
                                                                                                                            <td>
                                                                                                                                <table
                                                                                                                                    border="0"
                                                                                                                                    width="34%"
                                                                                                                                    cellpadding="0"
                                                                                                                                    cellspacing="0"
                                                                                                                                    id="trsortbasis3"
                                                                                                                                    runat="server"
                                                                                                                                    style="display: none">
                                                                                                                                    <tr>
                                                                                                                                        <td class="Tdcaption"
                                                                                                                                            valign="top"
                                                                                                                                            width="12%">
                                                                                                                                            Group
                                                                                                                                            By(3)
                                                                                                                                        </td>
                                                                                                                                        <td class="TdColon"
                                                                                                                                            valign="top"
                                                                                                                                            width="2%">
                                                                                                                                            :
                                                                                                                                        </td>
                                                                                                                                        <td class="tdvd"
                                                                                                                                            valign="top"
                                                                                                                                            colspan="5">
                                                                                                                                            <asp:DropDownList
                                                                                                                                                ID="ddlGroup3"
                                                                                                                                                runat="server"
                                                                                                                                                CssClass="DropdownList"
                                                                                                                                                AutoPostBack="True">
                                                                                                                                            </asp:DropDownList>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td class="Tdcaption"
                                                                                                                                            valign="top">
                                                                                                                                            List
                                                                                                                                        </td>
                                                                                                                                        <td class="TdColon"
                                                                                                                                            valign="top">
                                                                                                                                            :
                                                                                                                                        </td>
                                                                                                                                        <td colspan="5"
                                                                                                                                            class="tdcaption"
                                                                                                                                            valign="top">
                                                                                                                                            <div class="tdcolon"
                                                                                                                                                id="Div2"
                                                                                                                                                style="border-right: #003399 1px solid; border-top: #003399 1px solid;
                                                                                                                    overflow: auto; border-left: #003399 1px solid; width: 200px; border-bottom: #003399 1px solid;
                                                                                                                    height: 95px">
                                                                                                                                                <table
                                                                                                                                                    width="100%"
                                                                                                                                                    border="0"
                                                                                                                                                    cellpadding="0"
                                                                                                                                                    cellspacing="0"
                                                                                                                                                    id="Group3">
                                                                                                                                                    <tr>
                                                                                                                                                        <td align="left"
                                                                                                                                                            class="TDCaptionBold">
                                                                                                                                                            <asp:CheckBox
                                                                                                                                                                ID="chkAllGr2"
                                                                                                                                                                runat="server"
                                                                                                                                                                Checked="true"
                                                                                                                                                                onclick="checkUncheckCheckBoxlixtChkBox('Div2' ,this , 'chkListGroup2');"
                                                                                                                                                                CssClass="checkbox" />
                                                                                                                                                            Select
                                                                                                                                                            All
                                                                                                                                                        </td>
                                                                                                                                                    </tr>
                                                                                                                                                    <tr>
                                                                                                                                                        <td>
                                                                                                                                                            <asp:CheckBoxList
                                                                                                                                                                ID="chkListGroup2"
                                                                                                                                                                runat="server"
                                                                                                                                                                RepeatDirection="Horizontal"
                                                                                                                                                                RepeatColumns="1"
                                                                                                                                                                CssClass="chkList">
                                                                                                                                                            </asp:CheckBoxList>
                                                                                                                                                            <input
                                                                                                                                                                id="hidGroup1Count"
                                                                                                                                                                type="hidden"
                                                                                                                                                                name="hidGroup1Count"
                                                                                                                                                                runat="server" />
                                                                                                                                                            <input
                                                                                                                                                                id="hidGroup2Count"
                                                                                                                                                                type="hidden"
                                                                                                                                                                name="hidGroup2Count"
                                                                                                                                                                runat="server" />
                                                                                                                                                        </td>
                                                                                                                                                    </tr>
                                                                                                                                                </table>
                                                                                                                                            </div>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                </table>
                                                                                                                            </td>
                                                                                                                        </tr>
                                                                                                                    </table>
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                            <%--START:
                                                                                                                Added
                                                                                                                row by
                                                                                                                Rohtas
                                                                                                                Singh on
                                                                                                                06 Dec
                                                                                                                2017 for
                                                                                                                Monthly
                                                                                                                Salary
                                                                                                                Slip
                                                                                                                (MAX
                                                                                                                Life--%>
                                                                                                                <tr id="trRptformat"
                                                                                                                    runat="server"
                                                                                                                    style="display: none">
                                                                                                                    <td
                                                                                                                        class="Tdcaption">
                                                                                                                        Report
                                                                                                                        Format
                                                                                                                    </td>
                                                                                                                    <td
                                                                                                                        class="TdColon">
                                                                                                                        :
                                                                                                                    </td>
                                                                                                                    <td class="tdvd"
                                                                                                                        colspan="5">
                                                                                                                        <asp:DropDownList
                                                                                                                            ID="ddlRptFormat"
                                                                                                                            runat="server"
                                                                                                                            CssClass="dropdownlist">
                                                                                                                            <asp:ListItem
                                                                                                                                Text="CSV"
                                                                                                                                Value="1">
                                                                                                                            </asp:ListItem>
                                                                                                                            <asp:ListItem
                                                                                                                                Text="Excel"
                                                                                                                                Value="2">
                                                                                                                            </asp:ListItem>
                                                                                                                        </asp:DropDownList>
                                                                                                                    </td>
                                                                                                                </tr>
                                                                                                                <%--END:
                                                                                                                    Added
                                                                                                                    row
                                                                                                                    by
                                                                                                                    Rohtas
                                                                                                                    Singh
                                                                                                                    on
                                                                                                                    06
                                                                                                                    Dec
                                                                                                                    2017
                                                                                                                    for
                                                                                                                    Monthly
                                                                                                                    Salary
                                                                                                                    Slip
                                                                                                                    (MAX
                                                                                                                    Life--%>


                                                                                                                    <tr id="trling"
                                                                                                                        runat="server"
                                                                                                                        style="display: none">
                                                                                                                        <td
                                                                                                                            class="Tdcaption">
                                                                                                                            Select
                                                                                                                            Multilingual
                                                                                                                        </td>
                                                                                                                        <td
                                                                                                                            class="TdColon">
                                                                                                                            :
                                                                                                                        </td>
                                                                                                                        <td class="tdvd"
                                                                                                                            colspan="5">
                                                                                                                            <asp:DropDownList
                                                                                                                                ID="ddllingual"
                                                                                                                                runat="server"
                                                                                                                                CssClass="dropdownlist">
                                                                                                                            </asp:DropDownList>
                                                                                                                        </td>
                                                                                                                    </tr>

                                                                                                                    <tr id="trformat"
                                                                                                                        runat="server"
                                                                                                                        style="display: none">
                                                                                                                        <td class="TDCaption"
                                                                                                                            width="12%">
                                                                                                                            Report
                                                                                                                            Format
                                                                                                                        </td>
                                                                                                                        <td
                                                                                                                            class="TdColon">
                                                                                                                            :
                                                                                                                        </td>
                                                                                                                        <td
                                                                                                                            class="TDVd">
                                                                                                                            <asp:DropDownList
                                                                                                                                CssClass="DropdownList"
                                                                                                                                ID="ddlrepformat"
                                                                                                                                runat="server"
                                                                                                                                AutoPostBack="true">
                                                                                                                                <asp:ListItem
                                                                                                                                    Text="XLS"
                                                                                                                                    Value="XLS"
                                                                                                                                    Selected="True">
                                                                                                                                </asp:ListItem>
                                                                                                                                <asp:ListItem
                                                                                                                                    Text="CSV"
                                                                                                                                    Value="CSV">
                                                                                                                                </asp:ListItem>
                                                                                                                            </asp:DropDownList>
                                                                                                                        </td>
                                                                                                                    </tr>

                                                                                                                    <tr id="trEncrType"
                                                                                                                        runat="server"
                                                                                                                        style="display: none">
                                                                                                                        <td class="TDCaption"
                                                                                                                            width="12%">
                                                                                                                            Encryption
                                                                                                                            Type
                                                                                                                        </td>
                                                                                                                        <td
                                                                                                                            class="TdColon">
                                                                                                                            :
                                                                                                                        </td>
                                                                                                                        <td
                                                                                                                            class="TDVd">
                                                                                                                            <asp:DropDownList
                                                                                                                                CssClass="DropdownList"
                                                                                                                                ID="ddllEncrType"
                                                                                                                                runat="server"
                                                                                                                                Width="265px">
                                                                                                                                <asp:ListItem
                                                                                                                                    Text="Without Encryption"
                                                                                                                                    Value="WE"
                                                                                                                                    Selected="True">
                                                                                                                                </asp:ListItem>
                                                                                                                                <asp:ListItem
                                                                                                                                    Text="With Encryption"
                                                                                                                                    Value="WP">
                                                                                                                                </asp:ListItem>
                                                                                                                            </asp:DropDownList>
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr id="trFileName"
                                                                                                                        runat="server"
                                                                                                                        style="display: none">
                                                                                                                        <td class="TDCaption"
                                                                                                                            width="12%">
                                                                                                                            Report
                                                                                                                            Name
                                                                                                                        </td>
                                                                                                                        <td
                                                                                                                            class="TdColon">
                                                                                                                            :
                                                                                                                        </td>
                                                                                                                        <td
                                                                                                                            class="tdvd">
                                                                                                                            <asp:TextBox
                                                                                                                                id="txtrptName"
                                                                                                                                runat="server"
                                                                                                                                CssClass="TextBox txtEmpname keypressPasteNoSplCharAlwSomeSplChr"
                                                                                                                                MaxLength="100"
                                                                                                                                onchange="isValidFileName();">
                                                                                                                            </asp:TextBox>

                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr id="trSftpID"
                                                                                                                        runat="server"
                                                                                                                        style="display: none">
                                                                                                                        <td class="TDCaption"
                                                                                                                            width="12%">
                                                                                                                            Transfer
                                                                                                                            File
                                                                                                                            to
                                                                                                                            SFTP
                                                                                                                            Server
                                                                                                                        </td>
                                                                                                                        <td
                                                                                                                            class="TdColon">
                                                                                                                            :
                                                                                                                        </td>
                                                                                                                        <td
                                                                                                                            class="tdvd">
                                                                                                                            <asp:CheckBox
                                                                                                                                ID="chkSFTP"
                                                                                                                                runat="server"
                                                                                                                                CssClass="Checkbox"
                                                                                                                                AutoPostBack="false" />
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr>
                                                                                                                        <td class="trupbtn"
                                                                                                                            colspan="3">
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr id="trview"
                                                                                                                        runat="server"
                                                                                                                        style="display: none">
                                                                                                                        <td
                                                                                                                            colspan="3">
                                                                                                                            <table
                                                                                                                                border="0"
                                                                                                                                cellpadding="0"
                                                                                                                                cellspacing="0"
                                                                                                                                width="100%">
                                                                                                                                <tr>
                                                                                                                                    <td>
                                                                                                                                        <div id="P1"
                                                                                                                                            class="tab"
                                                                                                                                            style="border-width: 1px; width: 100%;"
                                                                                                                                            runat="server"
                                                                                                                                            onclick="return toggle();">
                                                                                                                                            Show/Hide
                                                                                                                                            Salary
                                                                                                                                            slip
                                                                                                                                            register
                                                                                                                                            Configuration
                                                                                                                                        </div>
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                                <tr id="trsetting"
                                                                                                                                    runat="server"
                                                                                                                                    style="display: none;">
                                                                                                                                    <td>
                                                                                                                                        <table
                                                                                                                                            border="0"
                                                                                                                                            cellpadding="0"
                                                                                                                                            cellspacing="0"
                                                                                                                                            width="100%">
                                                                                                                                            <tr>
                                                                                                                                                <td align="left"
                                                                                                                                                    colspan="3">
                                                                                                                                                    <%=ViewState("SalaryData")%>
                                                                                                                                                </td>
                                                                                                                                            </tr>
                                                                                                                                            <tr>
                                                                                                                                                <td class="trupbtn"
                                                                                                                                                    colspan="3">
                                                                                                                                                </td>
                                                                                                                                            </tr>
                                                                                                                                            <tr>
                                                                                                                                                <td colspan="3"
                                                                                                                                                    align="center">
                                                                                                                                                    <asp:Button
                                                                                                                                                        ID="Btnclick"
                                                                                                                                                        runat="server"
                                                                                                                                                        CssClass="btn"
                                                                                                                                                        Text="Click Here to Change"
                                                                                                                                                        Width="140px" />
                                                                                                                                                </td>
                                                                                                                                            </tr>
                                                                                                                                            <tr>
                                                                                                                                                <td class="trupbtn"
                                                                                                                                                    colspan="3">
                                                                                                                                                </td>
                                                                                                                                            </tr>
                                                                                                                                        </table>
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                                <tr id="trgap"
                                                                                                                                    runat="server"
                                                                                                                                    class="trupbtn">
                                                                                                                                    <td>
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                            </table>
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr>
                                                                                                                        <td
                                                                                                                            colspan="7">
                                                                                                                            <table
                                                                                                                                id="tblothepaycode"
                                                                                                                                runat="server"
                                                                                                                                border="0"
                                                                                                                                cellpadding="0"
                                                                                                                                cellspacing="0"
                                                                                                                                width="100%"
                                                                                                                                style="display: none">
                                                                                                                                <tr>
                                                                                                                                    <td class="Tdcaption"
                                                                                                                                        width="12%">
                                                                                                                                        Display
                                                                                                                                        other
                                                                                                                                        type
                                                                                                                                        paycode
                                                                                                                                    </td>
                                                                                                                                    <td class="tdcolon"
                                                                                                                                        valign="top">
                                                                                                                                        :
                                                                                                                                    </td>
                                                                                                                                    <td colspan="2"
                                                                                                                                        valign="top">
                                                                                                                                        <asp:CheckBox
                                                                                                                                            ID="chkother"
                                                                                                                                            runat="server"
                                                                                                                                            onclick="ShowHide(this)"
                                                                                                                                            CssClass="radio" />
                                                                                                                                        &nbsp;
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                                <tr id="tbl_QuestionTags_Software"
                                                                                                                                    style="display: none">
                                                                                                                                    <td class="Tdcaption"
                                                                                                                                        width="12%">
                                                                                                                                    </td>
                                                                                                                                    <td class="tdcolon"
                                                                                                                                        valign="top">
                                                                                                                                    </td>
                                                                                                                                    <td width="12%"
                                                                                                                                        valign="top">
                                                                                                                                        <asp:CheckBox
                                                                                                                                            ID="chkAllot"
                                                                                                                                            runat="server"
                                                                                                                                            CssClass="cHKLIST"
                                                                                                                                            onclick="checkUncheckCheckBoxlixtChkBox('DivOther' ,this , 'Chklistothepaycode');"
                                                                                                                                            Text="Select All" />
                                                                                                                                    <td valign="top"
                                                                                                                                        align="left">
                                                                                                                                        <div id="DivOther"
                                                                                                                                            style="border-right: #003399 1px solid; border-top: #003399 1px solid;
                                                                                                            overflow: auto; border-left: #003399 1px solid; width: 100%; border-bottom: #003399 1px solid;
                                                                                                            height: 95px">
                                                                                                                                            <asp:CheckBoxList
                                                                                                                                                ID="Chklistothepaycode"
                                                                                                                                                runat="server"
                                                                                                                                                RepeatDirection="Horizontal"
                                                                                                                                                CssClass="chklist"
                                                                                                                                                CellPadding="0"
                                                                                                                                                CellSpacing="0"
                                                                                                                                                RepeatColumns="8">
                                                                                                                                            </asp:CheckBoxList>
                                                                                                                                        </div>
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                                <tr>
                                                                                                                                    <td class="TRUpBtn"
                                                                                                                                        colspan="4">
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                            </table>
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr>
                                                                                                                        <td
                                                                                                                            colspan="3">
                                                                                                                            <table
                                                                                                                                id="tblsection"
                                                                                                                                cellspacing="0"
                                                                                                                                cellpadding="0"
                                                                                                                                width="100%"
                                                                                                                                border="0"
                                                                                                                                runat="server">
                                                                                                                                <tr>
                                                                                                                                    <td class="tdcaption"
                                                                                                                                        colspan="6">
                                                                                                                                        View
                                                                                                                                        Format(Arrear
                                                                                                                                        OR
                                                                                                                                        W/O
                                                                                                                                        Arrear):
                                                                                                                                        <asp:DropDownList
                                                                                                                                            ID="ddlarrear"
                                                                                                                                            runat="server"
                                                                                                                                            CssClass="dropdownlist">
                                                                                                                                            <asp:ListItem
                                                                                                                                                Value="2">
                                                                                                                                                Format
                                                                                                                                                A
                                                                                                                                                (Arrear)
                                                                                                                                            </asp:ListItem>
                                                                                                                                            <asp:ListItem
                                                                                                                                                Value="1">
                                                                                                                                                Format
                                                                                                                                                B
                                                                                                                                                (W/O
                                                                                                                                                Arrear)
                                                                                                                                            </asp:ListItem>
                                                                                                                                        </asp:DropDownList>
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                                <tr>
                                                                                                                                    <td class="tdcaption"
                                                                                                                                        colspan="6">
                                                                                                                                        Would
                                                                                                                                        you
                                                                                                                                        like
                                                                                                                                        to
                                                                                                                                        see
                                                                                                                                        the
                                                                                                                                        following
                                                                                                                                        options
                                                                                                                                        on
                                                                                                                                        the
                                                                                                                                        Report:
                                                                                                                                        <asp:CheckBox
                                                                                                                                            ID="chkboxent"
                                                                                                                                            runat="server"
                                                                                                                                            Text="Entitlements">
                                                                                                                                        </asp:CheckBox>
                                                                                                                                        <asp:CheckBox
                                                                                                                                            ID="chkboxotherinc"
                                                                                                                                            runat="server"
                                                                                                                                            Text="Income from other sources">
                                                                                                                                        </asp:CheckBox>
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                                <tr>
                                                                                                                                    <td class="tdcaption"
                                                                                                                                        colspan="6">
                                                                                                                                        Select
                                                                                                                                        the
                                                                                                                                        following
                                                                                                                                        checkboxes
                                                                                                                                        under
                                                                                                                                        Chapter
                                                                                                                                        VI
                                                                                                                                        A
                                                                                                                                        as
                                                                                                                                        you
                                                                                                                                        would
                                                                                                                                        like
                                                                                                                                        to
                                                                                                                                        appear
                                                                                                                                        on
                                                                                                                                        Salary
                                                                                                                                        Slip:
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                                <tr>
                                                                                                                                    <td
                                                                                                                                        class="Tdcaption">
                                                                                                                                        <asp:CheckBoxList
                                                                                                                                            ID="CBLsection"
                                                                                                                                            runat="server"
                                                                                                                                            CssClass="chkList">
                                                                                                                                        </asp:CheckBoxList>
                                                                                                                                        <input
                                                                                                                                            id="hdquery"
                                                                                                                                            type="hidden"
                                                                                                                                            runat="server" />
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                            </table>
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <!--New message displayed by rohtas-->
                                                                                                                    <!-- For Show Addition & Deduction Paycode-->
                                                                                                                    <tr>
                                                                                                                        <td
                                                                                                                            colspan="3">
                                                                                                                            <table
                                                                                                                                id="tblpaycode"
                                                                                                                                cellspacing="0"
                                                                                                                                cellpadding="0"
                                                                                                                                width="100%"
                                                                                                                                runat="server">
                                                                                                                                <tbody>
                                                                                                                                    <tr>
                                                                                                                                        <td
                                                                                                                                            class="FormSubHeading">
                                                                                                                                            Addition
                                                                                                                                            Type
                                                                                                                                            Paycode<img
                                                                                                                                                alt=""
                                                                                                                                                title="Instruction"
                                                                                                                                                onclick="show('tb8');"
                                                                                                                                                src="ImagescA/new4-0985.gif"
                                                                                                                                                border="0" />
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td
                                                                                                                                            class="FormSubHeading">
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td
                                                                                                                                            class="FormSubHeading">
                                                                                                                                            <input
                                                                                                                                                type="checkbox"
                                                                                                                                                checked="checked"
                                                                                                                                                id="ChkAddAll"
                                                                                                                                                onclick="checkUncheckCheckBoxlixtChkBox('AddPaycode',this,'cbladd')"
                                                                                                                                                class="chkList"
                                                                                                                                                text="All Addition">
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td>
                                                                                                                                            <hr
                                                                                                                                                class="HRColor">
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td>
                                                                                                                                            <div id="AddPaycode"
                                                                                                                                                style="border-right: #003399 1px solid; border-top: #003399 1px solid;
                                                                                                            overflow: auto; border-left: #003399 1px solid; width: 100%; border-bottom: #003399 1px solid;
                                                                                                            height: 95px">
                                                                                                                                                <asp:CheckBoxList
                                                                                                                                                    ID="cbladd"
                                                                                                                                                    runat="server"
                                                                                                                                                    CssClass="chkList"
                                                                                                                                                    RepeatDirection="Horizontal"
                                                                                                                                                    RepeatColumns="9">
                                                                                                                                                </asp:CheckBoxList>
                                                                                                                                            </div>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td
                                                                                                                                            class="FormSubHeading">
                                                                                                                                            Deduction
                                                                                                                                            Type
                                                                                                                                            Paycode
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td
                                                                                                                                            class="FormSubHeading">
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td
                                                                                                                                            class="FormSubHeading">
                                                                                                                                            <input
                                                                                                                                                type="checkbox"
                                                                                                                                                id="ChkDedAll"
                                                                                                                                                checked="checked"
                                                                                                                                                onclick="checkUncheckCheckBoxlixtChkBox('DedPaycode',this,'cbldeduction')"
                                                                                                                                                class="chkList"
                                                                                                                                                text="All Deduction">
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td>
                                                                                                                                            <hr
                                                                                                                                                class="HRColor">
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td>
                                                                                                                                            <div id="DedPaycode"
                                                                                                                                                style="border-right: #003399 1px solid; border-top: #003399 1px solid;
                                                                                                            overflow: auto; border-left: #003399 1px solid; width: 100%; border-bottom: #003399 1px solid;
                                                                                                            height: 50px">
                                                                                                                                                <asp:CheckBoxList
                                                                                                                                                    ID="cbldeduction"
                                                                                                                                                    runat="server"
                                                                                                                                                    CssClass="chkList"
                                                                                                                                                    RepeatDirection="Horizontal"
                                                                                                                                                    RepeatColumns="9">
                                                                                                                                                </asp:CheckBoxList>
                                                                                                                                            </div>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td>
                                                                                                                                            <table
                                                                                                                                                id="Table5"
                                                                                                                                                cellspacing="0"
                                                                                                                                                cellpadding="0"
                                                                                                                                                width="20%"
                                                                                                                                                border="0">
                                                                                                                                                <tr>
                                                                                                                                                    <td align="left"
                                                                                                                                                        width="2%">
                                                                                                                                                    </td>
                                                                                                                                                    <td
                                                                                                                                                        valign="top">
                                                                                                                                                        <table
                                                                                                                                                            class="Show"
                                                                                                                                                            id="tb8"
                                                                                                                                                            style="display: none"
                                                                                                                                                            onclick="show('tb8');"
                                                                                                                                                            cellspacing="1"
                                                                                                                                                            cellpadding="1"
                                                                                                                                                            border="1">
                                                                                                                                                            <tr>
                                                                                                                                                                <td>
                                                                                                                                                                    <asp:Label
                                                                                                                                                                        ID="Label1"
                                                                                                                                                                        runat="server">
                                                                                                                                                                        At
                                                                                                                                                                        Lest
                                                                                                                                                                        One
                                                                                                                                                                        Addtion
                                                                                                                                                                        and
                                                                                                                                                                        Deduction
                                                                                                                                                                        Type
                                                                                                                                                                        Paycode
                                                                                                                                                                        Must
                                                                                                                                                                        Be
                                                                                                                                                                        Select
                                                                                                                                                                    </asp:Label>
                                                                                                                                                                </td>
                                                                                                                                                            </tr>
                                                                                                                                                        </table>
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                            </table>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                </tbody>
                                                                                                                            </table>
                                                                                                                            <table
                                                                                                                                id="tblhelp"
                                                                                                                                border="0"
                                                                                                                                cellpadding="0"
                                                                                                                                cellspacing="0"
                                                                                                                                width="100%"
                                                                                                                                runat="server">
                                                                                                                                <tr>
                                                                                                                                    <td
                                                                                                                                        class="TRGap">
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                                <tr
                                                                                                                                    class="Message">
                                                                                                                                    <td style="width: 2%;"
                                                                                                                                        valign="top">
                                                                                                                                        <img
                                                                                                                                            src="ImagesCA/helpbulb.gif" />
                                                                                                                                    </td>
                                                                                                                                    <td>
                                                                                                                                        <span
                                                                                                                                            style="color: Black; font-size: 11px; align: justify;"><b
                                                                                                                                                style="color: Red">Help:</b>
                                                                                                                                            Resister
                                                                                                                                            may
                                                                                                                                            not
                                                                                                                                            show
                                                                                                                                            correctly
                                                                                                                                            because
                                                                                                                                            pay
                                                                                                                                            components
                                                                                                                                            are
                                                                                                                                            hard
                                                                                                                                            coded.</span><span
                                                                                                                                            style="font-size: 7.5pt; color: Black; line-height: 115%; font-family: Verdana;
                                                                                                            mso-fareast-font-family: 'Times New Roman'; mso-bidi-font-family: 'Times New Roman';
                                                                                                            mso-ansi-language: EN-IN; mso-fareast-language: EN-US; mso-bidi-language: AR-SA">&nbsp;</span><br />
                                                                                                                                    </td>
                                                                                                                                </tr>
                                                                                                                            </table>
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <!-- End of For Show Addition & Deduction Paycode-->
                                                                                                                    <!--=================================================================================================================================================================================-->
                                                                                                                    <!-- For Show Reimbusement Paycode-->
                                                                                                                    <tr>
                                                                                                                        <td
                                                                                                                            colspan="3">
                                                                                                                            <table
                                                                                                                                id="TblReimb"
                                                                                                                                cellspacing="0"
                                                                                                                                cellpadding="0"
                                                                                                                                width="100%"
                                                                                                                                runat="server"
                                                                                                                                style="display: none">
                                                                                                                                <tbody>
                                                                                                                                    <tr>
                                                                                                                                        <td
                                                                                                                                            class="FormSubHeading">
                                                                                                                                            Show
                                                                                                                                            Selected
                                                                                                                                            Reimbursement
                                                                                                                                            Seperately
                                                                                                                                            <img alt=""
                                                                                                                                                title="Instruction"
                                                                                                                                                onclick="show('TblReimbMsg');"
                                                                                                                                                src="ImagescA/new4-0985.gif"
                                                                                                                                                border="0" />
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td
                                                                                                                                            class="FormSubHeading">
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td
                                                                                                                                            class="FormSubHeading">
                                                                                                                                            <input
                                                                                                                                                type="checkbox"
                                                                                                                                                id="ChkReimbAll"
                                                                                                                                                checked="checked"
                                                                                                                                                onclick="checkUncheckCheckBoxlixtChkBox('ReimbPaycode',this,'cblReimb')"
                                                                                                                                                class="chkList"
                                                                                                                                                text="All Reimbursement">
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td>
                                                                                                                                            <hr
                                                                                                                                                class="HRColor">
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td>
                                                                                                                                            <div id="ReimbPaycode"
                                                                                                                                                style="border-right: #003399 1px solid; border-top: #003399 1px solid;
                                                                                                            overflow: auto; border-left: #003399 1px solid; width: 100%; border-bottom: #003399 1px solid;
                                                                                                            height: 95px">
                                                                                                                                                <asp:CheckBoxList
                                                                                                                                                    ID="cblReimb"
                                                                                                                                                    runat="server"
                                                                                                                                                    CssClass="chkList"
                                                                                                                                                    RepeatDirection="Horizontal"
                                                                                                                                                    RepeatColumns="9">
                                                                                                                                                </asp:CheckBoxList>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td>
                                                                                                                                            <table
                                                                                                                                                id="TblReimb2"
                                                                                                                                                cellspacing="0"
                                                                                                                                                cellpadding="0"
                                                                                                                                                width="20%"
                                                                                                                                                border="0">
                                                                                                                                                <tr>
                                                                                                                                                    <td align="left"
                                                                                                                                                        width="2%">
                                                                                                                                                    </td>
                                                                                                                                                    <td
                                                                                                                                                        valign="top">
                                                                                                                                                        <table
                                                                                                                                                            class="Show"
                                                                                                                                                            id="TblReimbMsg"
                                                                                                                                                            style="display: none"
                                                                                                                                                            onclick="show('TblReimbMsg');"
                                                                                                                                                            cellspacing="1"
                                                                                                                                                            cellpadding="1"
                                                                                                                                                            border="1">
                                                                                                                                                            <tr>
                                                                                                                                                                <td>
                                                                                                                                                                    <asp:Label
                                                                                                                                                                        ID="lblReimb"
                                                                                                                                                                        runat="server">
                                                                                                                                                                        Select
                                                                                                                                                                        at
                                                                                                                                                                        least
                                                                                                                                                                        one
                                                                                                                                                                        reimbursement
                                                                                                                                                                        paycode
                                                                                                                                                                    </asp:Label>
                                                                                                                                                                </td>
                                                                                                                                                            </tr>
                                                                                                                                                        </table>
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                            </table>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                </tbody>
                                                                                                                            </table>
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                                    <tr>
                                                                                                                        <td colspan="3"
                                                                                                                            align="center">
                                                                                                                            <div id="divgenerate"
                                                                                                                                runat="server">
                                                                                                                                <table
                                                                                                                                    border="0"
                                                                                                                                    cellpadding="0"
                                                                                                                                    cellspacing="0"
                                                                                                                                    width="100%"
                                                                                                                                    align="center">
                                                                                                                                    <tr>
                                                                                                                                        <td class="Tdcaption"
                                                                                                                                            valign="top"
                                                                                                                                            width="12%">
                                                                                                                                            <input
                                                                                                                                                id="HidPreVal"
                                                                                                                                                type="hidden"
                                                                                                                                                runat="server"
                                                                                                                                                name="hidval" />
                                                                                                                                        </td>
                                                                                                                                        <td class="TdColon"
                                                                                                                                            valign="top">
                                                                                                                                        </td>
                                                                                                                                        <td>
                                                                                                                                            <table
                                                                                                                                                border="0"
                                                                                                                                                cellpadding="0"
                                                                                                                                                cellspacing="0">
                                                                                                                                                <tr>
                                                                                                                                                    <td
                                                                                                                                                        valign="top">
                                                                                                                                                        <asp:Button
                                                                                                                                                            ID="btnPreview"
                                                                                                                                                            runat="server"
                                                                                                                                                            CssClass="btn"
                                                                                                                                                            Text="Preview"
                                                                                                                                                            OnClientClick="javascript:return btnPreview_Click();" />
                                                                                                                                                        <asp:Button
                                                                                                                                                            ID="btnExport2CSV"
                                                                                                                                                            runat="server"
                                                                                                                                                            CssClass="btn"
                                                                                                                                                            Width="100px"
                                                                                                                                                            Visible="false"
                                                                                                                                                            Text="Export to CSV"
                                                                                                                                                            OnClientClick="javascript:return btnPreview_Click();" />

                                                                                                                                                        <%--Added
                                                                                                                                                            by
                                                                                                                                                            Debargha
                                                                                                                                                            on
                                                                                                                                                            17-May-2024--%>
                                                                                                                                                            <input
                                                                                                                                                                id="hdfile"
                                                                                                                                                                type="hidden"
                                                                                                                                                                name="hdfile"
                                                                                                                                                                runat="server" />
                                                                                                                                                            <input
                                                                                                                                                                id="hdtdsfile"
                                                                                                                                                                type="hidden"
                                                                                                                                                                name="hdtdsfile"
                                                                                                                                                                runat="server" />
                                                                                                                                                            <%--Added
                                                                                                                                                                by
                                                                                                                                                                Debargha
                                                                                                                                                                on
                                                                                                                                                                17-May-2024--%>
                                                                                                                                                    </td>
                                                                                                                                                    <td
                                                                                                                                                        class="tdbetweenbtn">
                                                                                                                                                    </td>
                                                                                                                                                    <td
                                                                                                                                                        valign="top">
                                                                                                                                                        <asp:Button
                                                                                                                                                            ID="btnresetsearch"
                                                                                                                                                            runat="server"
                                                                                                                                                            CssClass="btn"
                                                                                                                                                            Text="Reset" />
                                                                                                                                                    </td>
                                                                                                                                                    <td
                                                                                                                                                        class="tdbetweenlbl">
                                                                                                                                                    </td>
                                                                                                                                                    <td
                                                                                                                                                        valign="top">
                                                                                                                                                        <asp:Label
                                                                                                                                                            ID="lblmsg2"
                                                                                                                                                            runat="server"
                                                                                                                                                            CssClass="usermessage">
                                                                                                                                                        </asp:Label>
                                                                                                                                                    </td>
                                                                                                                                                </tr>

                                                                                                                                            </table>
                                                                                                                                            <asp:LinkButton
                                                                                                                                                ID="LinkButton2"
                                                                                                                                                runat="server">
                                                                                                                                            </asp:LinkButton>
                                                                                                                                            <input
                                                                                                                                                id="HidAdd"
                                                                                                                                                type="hidden"
                                                                                                                                                name="HidAdd"
                                                                                                                                                runat="server" />
                                                                                                                                            <input
                                                                                                                                                id="HidDed"
                                                                                                                                                type="hidden"
                                                                                                                                                name="HidDed"
                                                                                                                                                runat="server" />
                                                                                                                                            <input
                                                                                                                                                id="Hidothepaycode"
                                                                                                                                                type="hidden"
                                                                                                                                                name="HidPreVal"
                                                                                                                                                runat="server" />
                                                                                                                                            <input
                                                                                                                                                id="hidothrpaycode"
                                                                                                                                                type="hidden"
                                                                                                                                                name="HidPreVal"
                                                                                                                                                runat="server" />
                                                                                                                                            <input
                                                                                                                                                id="HidReimb"
                                                                                                                                                type="hidden"
                                                                                                                                                name="HidDed"
                                                                                                                                                runat="server" />
                                                                                                                                            <input
                                                                                                                                                id="hidGroup2Val"
                                                                                                                                                type="hidden"
                                                                                                                                                name="hidGroup2Val"
                                                                                                                                                runat="server" />
                                                                                                                                            <input
                                                                                                                                                id="hidGroup3Val"
                                                                                                                                                type="hidden"
                                                                                                                                                name="hidGroup3Val"
                                                                                                                                                runat="server" />
                                                                                                                                            <input
                                                                                                                                                type="hidden"
                                                                                                                                                id="JobUniqueId"
                                                                                                                                                name="JobUniqueId"
                                                                                                                                                runat="server" />
                                                                                                                                            <asp:Literal
                                                                                                                                                ID="lit"
                                                                                                                                                runat="server">
                                                                                                                                            </asp:Literal>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td
                                                                                                                                            colspan="3">
                                                                                                                                            &nbsp;
                                                                                                                                        </td>
                                                                                                                                    </tr>

                                                                                                                                </table>
                                                                                                                            </div>
                                                                                                                        </td>
                                                                                                                    </tr>
                                                                                                        </table>
                                                                                                    </ContentTemplate>
                                                                                                    <Triggers>
                                                                                                        <asp:PostBackTrigger
                                                                                                            ControlID="btnPreview" />
                                                                                                        <asp:PostBackTrigger
                                                                                                            ControlID="btnExport2CSV" />
                                                                                                        <asp:PostBackTrigger
                                                                                                            ControlID="btnProgressbarExcel" />
                                                                                                    </Triggers>
                                                                                                </asp:UpdatePanel>
                                                                                            </td>
                                                                                        </tr>
                                                                                        <tr>
                                                                                            <td class="tdcaption"
                                                                                                colspan="3"
                                                                                                align="center">
                                                                                                <div id="divEmail"
                                                                                                    style="border-right: #003399 0px solid; border-top: #003399 0px solid;
                                                                        border-left: #003399 0px solid; width: 100%; border-bottom: #003399 0px solid;
                                                                        height: auto; display: none" runat="server">
                                                                                                    <%--<asp:UpdatePanel
                                                                                                        ID="UpdatePanel1"
                                                                                                        runat="server">
                                                                                                        <ContentTemplate>
                                                                                                            --%>
                                                                                                            <table
                                                                                                                border="0"
                                                                                                                align="center"
                                                                                                                cellpadding="0"
                                                                                                                cellspacing="0"
                                                                                                                width="100%">
                                                                                                                <tr>
                                                                                                                    <td>
                                                                                                                        <asp:UpdatePanel
                                                                                                                            ID="UpdatePanel1"
                                                                                                                            runat="server">
                                                                                                                            <ContentTemplate>
                                                                                                                                <table
                                                                                                                                    border="0"
                                                                                                                                    align="center"
                                                                                                                                    cellpadding="0"
                                                                                                                                    cellspacing="0"
                                                                                                                                    width="100%">
                                                                                                                                    <tr>
                                                                                                                                        <td
                                                                                                                                            colspan="3">
                                                                                                                                            <table
                                                                                                                                                id="Table6"
                                                                                                                                                width="100%"
                                                                                                                                                cellpadding="0"
                                                                                                                                                cellspacing="0"
                                                                                                                                                runat="server"
                                                                                                                                                border="0">
                                                                                                                                                <tr>
                                                                                                                                                    <td class="tdcaption"
                                                                                                                                                        width="27%">
                                                                                                                                                        Report
                                                                                                                                                        Type
                                                                                                                                                    </td>
                                                                                                                                                    <td width="6%"
                                                                                                                                                        style="color: Black; font-family: verdana; font-size: 11; text-align: left;">
                                                                                                                                                        :
                                                                                                                                                    </td>
                                                                                                                                                    <td
                                                                                                                                                        class="TDVd">
                                                                                                                                                        <asp:DropDownList
                                                                                                                                                            ID="DdlreportType"
                                                                                                                                                            runat="server"
                                                                                                                                                            CssClass="dropdownlist"
                                                                                                                                                            Width="230px"
                                                                                                                                                            AutoPostBack="true">
                                                                                                                                                        </asp:DropDownList>
                                                                                                                                                        *&nbsp;&nbsp;
                                                                                                                                                        <asp:Button
                                                                                                                                                            ID="BtnPreviewdivActive"
                                                                                                                                                            Style="display: none"
                                                                                                                                                            runat="server"
                                                                                                                                                            BackColor="CadetBlue"
                                                                                                                                                            CssClass="btn"
                                                                                                                                                            Height="17px"
                                                                                                                                                            Text="You can change the format of this report by clicking here."
                                                                                                                                                            ToolTip="You can change the format of this report by clicking here."
                                                                                                                                                            Width="360px"
                                                                                                                                                            ForeColor="Blue" />
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                                <tr id="troffcycledt"
                                                                                                                                                    runat="server"
                                                                                                                                                    style="display: none;">
                                                                                                                                                    <td class="tdcaption"
                                                                                                                                                        width="12%">
                                                                                                                                                        Select
                                                                                                                                                        Offcycle
                                                                                                                                                        Date
                                                                                                                                                    </td>
                                                                                                                                                    <td
                                                                                                                                                        class="tdcolon">
                                                                                                                                                        :
                                                                                                                                                    </td>
                                                                                                                                                    <td
                                                                                                                                                        class="tdvd">
                                                                                                                                                        <asp:DropDownList
                                                                                                                                                            ID="ddloffcycledt"
                                                                                                                                                            runat="server"
                                                                                                                                                            CssClass="dropdownlist">
                                                                                                                                                        </asp:DropDownList>
                                                                                                                                                        *
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                                <tr id="trlingpdf"
                                                                                                                                                    runat="server"
                                                                                                                                                    style="display: none">
                                                                                                                                                    <td
                                                                                                                                                        class="Tdcaption">
                                                                                                                                                        Select
                                                                                                                                                        Multilingual
                                                                                                                                                    </td>
                                                                                                                                                    <td
                                                                                                                                                        class="TdColon">
                                                                                                                                                        :
                                                                                                                                                    </td>
                                                                                                                                                    <td class="tdvd"
                                                                                                                                                        colspan="5">
                                                                                                                                                        <asp:DropDownList
                                                                                                                                                            ID="ddlmultilingual"
                                                                                                                                                            runat="server"
                                                                                                                                                            CssClass="dropdownlist">
                                                                                                                                                        </asp:DropDownList>
                                                                                                                                                    </td>
                                                                                                                                                </tr>

                                                                                                                                            </table>
                                                                                                                                        </td>
                                                                                                                                        <td colspan="3"
                                                                                                                                            class="tdvd">
                                                                                                                                            <table
                                                                                                                                                width="100%"
                                                                                                                                                cellpadding="0"
                                                                                                                                                cellspacing="0">
                                                                                                                                                <tr>
                                                                                                                                                    <td>
                                                                                                                                                        &nbsp;
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                            </table>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr id="trhelp1"
                                                                                                                                        runat="server"
                                                                                                                                        style="display: none;">
                                                                                                                                        <td
                                                                                                                                            colspan="3">
                                                                                                                                            <table
                                                                                                                                                width="100%"
                                                                                                                                                cellpadding="0"
                                                                                                                                                cellspacing="0">
                                                                                                                                                <tr>
                                                                                                                                                    <td class="tdcaption"
                                                                                                                                                        width="27%">
                                                                                                                                                        Show
                                                                                                                                                        Pay
                                                                                                                                                        Head
                                                                                                                                                        Help
                                                                                                                                                    </td>
                                                                                                                                                    <td class="tdcolon"
                                                                                                                                                        style="width: 6%">
                                                                                                                                                        :
                                                                                                                                                    </td>
                                                                                                                                                    <td
                                                                                                                                                        class="tdvd">
                                                                                                                                                        <asp:CheckBox
                                                                                                                                                            runat="server"
                                                                                                                                                            ID="chkHelp1"
                                                                                                                                                            CssClass="Checkbox" />
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                            </table>
                                                                                                                                        </td>
                                                                                                                                        <td
                                                                                                                                            colspan="3">
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td
                                                                                                                                            colspan="3">
                                                                                                                                            <table
                                                                                                                                                id="tblrepin"
                                                                                                                                                width="100%"
                                                                                                                                                cellpadding="0"
                                                                                                                                                cellspacing="0"
                                                                                                                                                runat="server"
                                                                                                                                                border="0">
                                                                                                                                                <tr>
                                                                                                                                                    <td class="tdcaption"
                                                                                                                                                        width="27%">
                                                                                                                                                        Report
                                                                                                                                                        In
                                                                                                                                                    </td>
                                                                                                                                                    <td width="6%"
                                                                                                                                                        style="color: Black; font-family: verdana; font-size: 11; text-align: left;">
                                                                                                                                                        :
                                                                                                                                                    </td>
                                                                                                                                                    <td>
                                                                                                                                                        <asp:DropDownList
                                                                                                                                                            ID="ddlRepIn"
                                                                                                                                                            runat="server"
                                                                                                                                                            CssClass="dropdownlist"
                                                                                                                                                            AutoPostBack="true"
                                                                                                                                                            onchange="ShowRepDet();">
                                                                                                                                                        </asp:DropDownList>
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                            </table>
                                                                                                                                        </td>
                                                                                                                                        <td
                                                                                                                                            colspan="3">
                                                                                                                                            <table
                                                                                                                                                id="tblsp"
                                                                                                                                                width="100%"
                                                                                                                                                cellpadding="0"
                                                                                                                                                cellspacing="0"
                                                                                                                                                runat="server">
                                                                                                                                                <tr>
                                                                                                                                                    <td class="tdcaption"
                                                                                                                                                        width="20%">
                                                                                                                                                        &nbsp;&nbsp;&nbsp;Salary
                                                                                                                                                        Processed/Not
                                                                                                                                                    </td>
                                                                                                                                                    <td class="tdcolon"
                                                                                                                                                        style="width: 6%"
                                                                                                                                                        align="center">
                                                                                                                                                        :
                                                                                                                                                    </td>
                                                                                                                                                    <td>
                                                                                                                                                        <asp:DropDownList
                                                                                                                                                            ID="ddlSalaryProcess"
                                                                                                                                                            runat="server"
                                                                                                                                                            CssClass="dropdownlist">
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Selected="True"
                                                                                                                                                                Value="Y">
                                                                                                                                                                Processed
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Value="N">
                                                                                                                                                                Not
                                                                                                                                                                Processed
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                        </asp:DropDownList>
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                            </table>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr id="tremail"
                                                                                                                                        runat="server">
                                                                                                                                        <td
                                                                                                                                            colspan="3">
                                                                                                                                            <table
                                                                                                                                                width="100%"
                                                                                                                                                cellpadding="0"
                                                                                                                                                cellspacing="0">
                                                                                                                                                <tr>
                                                                                                                                                    <td class="tdcaption"
                                                                                                                                                        width="27%">
                                                                                                                                                        E-Mail
                                                                                                                                                        Id
                                                                                                                                                        Exist/Not
                                                                                                                                                    </td>
                                                                                                                                                    <td class="tdcolon"
                                                                                                                                                        style="width: 6%">
                                                                                                                                                        :
                                                                                                                                                    </td>
                                                                                                                                                    <td
                                                                                                                                                        class="tdvd">
                                                                                                                                                        <asp:DropDownList
                                                                                                                                                            ID="ddlEMailExist"
                                                                                                                                                            runat="server"
                                                                                                                                                            CssClass="dropdownlist">
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Selected="True"
                                                                                                                                                                Value="B">
                                                                                                                                                                Both
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Value="Y">
                                                                                                                                                                Exist
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Value="N">
                                                                                                                                                                Not
                                                                                                                                                                Exist
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                        </asp:DropDownList>
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                            </table>
                                                                                                                                        </td>
                                                                                                                                        <td colspan="3"
                                                                                                                                            width="55%">
                                                                                                                                            <table
                                                                                                                                                width="100%"
                                                                                                                                                cellpadding="0"
                                                                                                                                                cellspacing="0">
                                                                                                                                                <tr>
                                                                                                                                                    <td class="tdcaption"
                                                                                                                                                        width="20%">
                                                                                                                                                        &nbsp;&nbsp;&nbsp;E-Mail
                                                                                                                                                        Sent/Not
                                                                                                                                                    </td>
                                                                                                                                                    <td class="tdcolon"
                                                                                                                                                        style="width: 6%">
                                                                                                                                                        :
                                                                                                                                                    </td>
                                                                                                                                                    <td>
                                                                                                                                                        <asp:DropDownList
                                                                                                                                                            ID="ddlEMailSend"
                                                                                                                                                            runat="server"
                                                                                                                                                            CssClass="dropdownlist">
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Selected="True"
                                                                                                                                                                Value="B">
                                                                                                                                                                Both
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Value="Y">
                                                                                                                                                                Mail
                                                                                                                                                                Sent
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Value="N">
                                                                                                                                                                Mail
                                                                                                                                                                Not
                                                                                                                                                                Sent
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                        </asp:DropDownList>
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                            </table>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr>
                                                                                                                                        <td
                                                                                                                                            colspan="3">
                                                                                                                                            <table
                                                                                                                                                id="tblSh"
                                                                                                                                                width="100%"
                                                                                                                                                cellpadding="0"
                                                                                                                                                cellspacing="0"
                                                                                                                                                runat="server">
                                                                                                                                                <tr>
                                                                                                                                                    <td class="tdcaption"
                                                                                                                                                        width="27%">
                                                                                                                                                        Salary
                                                                                                                                                        On
                                                                                                                                                        Hold/Not
                                                                                                                                                    </td>
                                                                                                                                                    <td class="tdcolon"
                                                                                                                                                        style="width: 6%">
                                                                                                                                                        :
                                                                                                                                                    </td>
                                                                                                                                                    <td
                                                                                                                                                        class="tdvd">
                                                                                                                                                        <asp:DropDownList
                                                                                                                                                            ID="ddlSalWithHeld"
                                                                                                                                                            runat="server"
                                                                                                                                                            CssClass="dropdownlist">
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Value="Y">
                                                                                                                                                                Salary
                                                                                                                                                                Withheld
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Selected="True"
                                                                                                                                                                Value="N">
                                                                                                                                                                Salary
                                                                                                                                                                Not
                                                                                                                                                                Withheld
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                        </asp:DropDownList>
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                            </table>
                                                                                                                                        </td>
                                                                                                                                        <td colspan="3"
                                                                                                                                            width="55%">
                                                                                                                                            <table
                                                                                                                                                id="tblpwd"
                                                                                                                                                width="100%"
                                                                                                                                                cellpadding="0"
                                                                                                                                                cellspacing="0"
                                                                                                                                                runat="server">
                                                                                                                                                <tr>
                                                                                                                                                    <td class="tdcaption"
                                                                                                                                                        width="20%">
                                                                                                                                                        &nbsp;&nbsp;&nbsp;PDF
                                                                                                                                                        Password
                                                                                                                                                        Type
                                                                                                                                                    </td>
                                                                                                                                                    <td class="tdcolon"
                                                                                                                                                        style="width: 6%">
                                                                                                                                                        :
                                                                                                                                                    </td>
                                                                                                                                                    <td>
                                                                                                                                                        <asp:DropDownList
                                                                                                                                                            ID="ddlEmpPass"
                                                                                                                                                            runat="server"
                                                                                                                                                            CssClass="dropdownlist"
                                                                                                                                                            Enabled="true">
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Text="No Password"
                                                                                                                                                                Value="0">
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Text="Emp. Code"
                                                                                                                                                                Value="3">
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Text="Pan No."
                                                                                                                                                                Value="1">
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Text="First Name & DOB"
                                                                                                                                                                Value="2">
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Text="Bank A/C No. & DOB"
                                                                                                                                                                Value="4">
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                            <%-- <asp:ListItem
                                                                                                                                                                Text="First Name & DOB(DDMMYYYY)"
                                                                                                                                                                Value="5">
                                                                                                                                                                </asp:ListItem>
                                                                                                                                                                --%>
                                                                                                                                                        </asp:DropDownList>
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                            </table>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr runat="server"
                                                                                                                                        id="trRepEmail">
                                                                                                                                        <td
                                                                                                                                            colspan="3">
                                                                                                                                            <table
                                                                                                                                                width="100%"
                                                                                                                                                cellpadding="0"
                                                                                                                                                cellspacing="0">
                                                                                                                                                <tr>
                                                                                                                                                    <td class="tdcaption"
                                                                                                                                                        width="27%">
                                                                                                                                                        Send
                                                                                                                                                        pay
                                                                                                                                                        slips
                                                                                                                                                        in
                                                                                                                                                        email
                                                                                                                                                        to
                                                                                                                                                        reporting
                                                                                                                                                        manager
                                                                                                                                                    </td>
                                                                                                                                                    <td class="tdcolon"
                                                                                                                                                        style="width: 6%">
                                                                                                                                                        :
                                                                                                                                                    </td>
                                                                                                                                                    <td
                                                                                                                                                        class="tdvd">
                                                                                                                                                        <asp:CheckBox
                                                                                                                                                            runat="server"
                                                                                                                                                            ID="chkemailrepmanager">
                                                                                                                                                        </asp:CheckBox>
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                            </table>
                                                                                                                                        </td>
                                                                                                                                        <td colspan="3"
                                                                                                                                            align="right">
                                                                                                                                            <table
                                                                                                                                                class="Message"
                                                                                                                                                width="100%"
                                                                                                                                                cellpadding="0"
                                                                                                                                                cellspacing="0">
                                                                                                                                                <tr>
                                                                                                                                                    <td
                                                                                                                                                        class="tdcaption">
                                                                                                                                                        <img
                                                                                                                                                            src="ImagesCA/Note.gif" />
                                                                                                                                                    </td>
                                                                                                                                                    <td>
                                                                                                                                                        <span
                                                                                                                                                            style="color: red"><b>Note:</b></span><span
                                                                                                                                                            style="font-size: 7.5pt; color: Black;
                                                                                                                        line-height: 115%; font-family: Verdana; mso-fareast-font-family: 'Times New Roman';
                                                                                                                        mso-bidi-font-family: 'Times New Roman'; mso-ansi-language: EN-IN; mso-fareast-language: EN-US;
                                                                                                                        mso-bidi-language: AR-SA">&nbsp;</span>
                                                                                                                                                        Enable
                                                                                                                                                        the
                                                                                                                                                        check
                                                                                                                                                        box,
                                                                                                                                                        if
                                                                                                                                                        you
                                                                                                                                                        want
                                                                                                                                                        to
                                                                                                                                                        send
                                                                                                                                                        pay
                                                                                                                                                        slips
                                                                                                                                                        to
                                                                                                                                                        Reporting
                                                                                                                                                        Manager.
                                                                                                                                                        You
                                                                                                                                                        can
                                                                                                                                                        send
                                                                                                                                                        PDF
                                                                                                                                                        or
                                                                                                                                                        HTML
                                                                                                                                                        pay
                                                                                                                                                        slip.
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                            </table>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <tr id="TrNoSearch"
                                                                                                                                        runat="server"
                                                                                                                                        style="display: none;">
                                                                                                                                        <td
                                                                                                                                            colspan="3">
                                                                                                                                            <table
                                                                                                                                                width="100%"
                                                                                                                                                cellpadding="0"
                                                                                                                                                cellspacing="0">
                                                                                                                                                <tr>
                                                                                                                                                    <td class="tdcaption"
                                                                                                                                                        width="27%">
                                                                                                                                                        Search
                                                                                                                                                        Record(s)
                                                                                                                                                        /
                                                                                                                                                        Publish
                                                                                                                                                        Record(s)
                                                                                                                                                        Without
                                                                                                                                                        Search
                                                                                                                                                    </td>
                                                                                                                                                    <td class="tdcolon"
                                                                                                                                                        style="width: 6%">
                                                                                                                                                        :
                                                                                                                                                    </td>
                                                                                                                                                    <td
                                                                                                                                                        class="tdvd">
                                                                                                                                                        <asp:RadioButtonList
                                                                                                                                                            ID="RblNoSearch"
                                                                                                                                                            runat="server"
                                                                                                                                                            CssClass="Radio"
                                                                                                                                                            AutoPostBack="true"
                                                                                                                                                            RepeatDirection="Horizontal"
                                                                                                                                                            onchange="javascript:return PleaseWaitWithDailog();">
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Text="Publish Payslip(s) Without Search"
                                                                                                                                                                Value="P"
                                                                                                                                                                Selected="True">
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                            <asp:ListItem
                                                                                                                                                                Text="Search Record(s)"
                                                                                                                                                                Value="S">
                                                                                                                                                            </asp:ListItem>
                                                                                                                                                        </asp:RadioButtonList>
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                                <tr id="trMerge"
                                                                                                                                                    runat="server"
                                                                                                                                                    style="display: none">
                                                                                                                                                    <td class="tdcaption"
                                                                                                                                                        width="27%">
                                                                                                                                                    </td>
                                                                                                                                                    <td class="tdcolon"
                                                                                                                                                        style="width: 6%">
                                                                                                                                                    </td>
                                                                                                                                                    <td
                                                                                                                                                        class="tdvd">
                                                                                                                                                        <asp:CheckBox
                                                                                                                                                            ID="chkMerge"
                                                                                                                                                            runat="server"
                                                                                                                                                            Checked="false"
                                                                                                                                                            CssClass="tdcaption"
                                                                                                                                                            Text="<b>Merge Payslip</b>" />
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                            </table>
                                                                                                                                        </td>
                                                                                                                                        <td colspan="3"
                                                                                                                                            align="right">
                                                                                                                                            <table
                                                                                                                                                class="Message"
                                                                                                                                                width="100%"
                                                                                                                                                cellpadding="0"
                                                                                                                                                cellspacing="0">
                                                                                                                                                <tr>
                                                                                                                                                    <td
                                                                                                                                                        class="tdcaption">
                                                                                                                                                        <img
                                                                                                                                                            src="ImagesCA/Note.gif" />
                                                                                                                                                    </td>
                                                                                                                                                    <td>
                                                                                                                                                        <span
                                                                                                                                                            style="color: red"><b>Note:</b></span><span
                                                                                                                                                            style="font-size: 7.5pt; color: Black;
                                                                                                                        line-height: 115%; font-family: Verdana; mso-fareast-font-family: 'Times New Roman';
                                                                                                                        mso-bidi-font-family: 'Times New Roman'; mso-ansi-language: EN-IN; mso-fareast-language: EN-US;
                                                                                                                        mso-bidi-language: AR-SA">&nbsp;</span><b>Search
                                                                                                                                                            Record(s):</b>
                                                                                                                                                        You
                                                                                                                                                        can
                                                                                                                                                        search
                                                                                                                                                        record(s)
                                                                                                                                                        before
                                                                                                                                                        sending
                                                                                                                                                        mail,
                                                                                                                                                        publish
                                                                                                                                                        record(s)
                                                                                                                                                        as
                                                                                                                                                        zip
                                                                                                                                                        file
                                                                                                                                                        and
                                                                                                                                                        publish
                                                                                                                                                        and
                                                                                                                                                        email
                                                                                                                                                        to
                                                                                                                                                        CC
                                                                                                                                                        /
                                                                                                                                                        BCC.<br />
                                                                                                                                                        <b>Publish
                                                                                                                                                            Record(s)
                                                                                                                                                            Without
                                                                                                                                                            Search:</b>
                                                                                                                                                        You
                                                                                                                                                        can
                                                                                                                                                        directly
                                                                                                                                                        publish
                                                                                                                                                        record(s)
                                                                                                                                                        as
                                                                                                                                                        zip
                                                                                                                                                        file
                                                                                                                                                        in
                                                                                                                                                        pdf
                                                                                                                                                        format
                                                                                                                                                        without
                                                                                                                                                        searching.
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                                <tr id="trMergeMsg"
                                                                                                                                                    runat="server"
                                                                                                                                                    style="display: none">
                                                                                                                                                    <td
                                                                                                                                                        class="tdcaption">
                                                                                                                                                        <img
                                                                                                                                                            src="ImagesCA/Note.gif" />
                                                                                                                                                    </td>
                                                                                                                                                    <td>
                                                                                                                                                        <span
                                                                                                                                                            style="color: red"><b>Note:</b></span><span
                                                                                                                                                            style="font-size: 7.5pt; color: Black;
                                                                                                                        line-height: 115%; font-family: Verdana; mso-fareast-font-family: 'Times New Roman';
                                                                                                                        mso-bidi-font-family: 'Times New Roman'; mso-ansi-language: EN-IN; mso-fareast-language: EN-US;
                                                                                                                        mso-bidi-language: AR-SA">&nbsp;</span><b>Merge
                                                                                                                                                            Payslip:</b>
                                                                                                                                                        Select
                                                                                                                                                        if
                                                                                                                                                        you
                                                                                                                                                        want
                                                                                                                                                        to
                                                                                                                                                        publish
                                                                                                                                                        the
                                                                                                                                                        PDF
                                                                                                                                                        tax
                                                                                                                                                        report
                                                                                                                                                        merged
                                                                                                                                                        with
                                                                                                                                                        Pay
                                                                                                                                                        Slip
                                                                                                                                                        With
                                                                                                                                                        Leave
                                                                                                                                                        Details.
                                                                                                                                                    </td>
                                                                                                                                                </tr>
                                                                                                                                            </table>
                                                                                                                                        </td>
                                                                                                                                    </tr>
                                                                                                                                    <%--START:
                                                                                                                                        Added
                                                                                                                                        by
                                                                                                                                        Quadir
                                                                                                                                        on
                                                                                                                                        14
                                                                                                                                        OCT
                                                                                                                                        2020-
                                                                                                                                        Payslip
                                                                                                                                        Publish
                                                                                                                                        Mode
                                                                                                                                        (RadioButton)--%>
                                                                                                                                        <tr id="TrSlipPubMode"
                                                                                                                                            runat="server"
                                                                                                                                            style="display: none;">
                                                                                                                                            <td
                                                                                                                                                colspan="3">
                                                                                                                                                <table
                                                                                                                                                    width="100%"
                                                                                                                                                    cellpadding="0"
                                                                                                                                                    cellspacing="0">
                                                                                                                                                    <tr>
                                                                                                                                                        <td class="tdcaption"
                                                                                                                                                            width="27%">
                                                                                                                                                            Payslip
                                                                                                                                                            Publish
                                                                                                                                                            Mode
                                                                                                                                                        </td>
                                                                                                                                                        <td class="tdcolon"
                                                                                                                                                            style="width: 6%">
                                                                                                                                                            :
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="tdvd">
                                                                                                                                                            <asp:RadioButtonList
                                                                                                                                                                ID="rbtSlipPubMode"
                                                                                                                                                                runat="server"
                                                                                                                                                                CssClass="Radio"
                                                                                                                                                                AutoPostBack="true"
                                                                                                                                                                RepeatDirection="Horizontal"
                                                                                                                                                                onchange="javascript:return PleaseWaitWithDailog();">
                                                                                                                                                                <asp:ListItem
                                                                                                                                                                    Text="Incremental"
                                                                                                                                                                    Value="I"
                                                                                                                                                                    Selected="True">
                                                                                                                                                                </asp:ListItem>
                                                                                                                                                                <asp:ListItem
                                                                                                                                                                    Text="Overwrite"
                                                                                                                                                                    Value="O">
                                                                                                                                                                </asp:ListItem>
                                                                                                                                                            </asp:RadioButtonList>
                                                                                                                                                        </td>
                                                                                                                                                    </tr>
                                                                                                                                                </table>
                                                                                                                                            </td>
                                                                                                                                            <td colspan="3"
                                                                                                                                                align="right">
                                                                                                                                            </td>
                                                                                                                                        </tr>
                                                                                                                                        <%--END:
                                                                                                                                            Added
                                                                                                                                            by
                                                                                                                                            Quadir
                                                                                                                                            on
                                                                                                                                            14
                                                                                                                                            OCT
                                                                                                                                            2020-
                                                                                                                                            Payslip
                                                                                                                                            Publish
                                                                                                                                            Mode
                                                                                                                                            (RadioButton)--%>
                                                                                                                                            <tr id="TrGrpbyPublish"
                                                                                                                                                runat="server"
                                                                                                                                                style="display: none;">
                                                                                                                                                <td
                                                                                                                                                    colspan="3">
                                                                                                                                                    <table
                                                                                                                                                        width="100%"
                                                                                                                                                        cellpadding="0"
                                                                                                                                                        cellspacing="0">
                                                                                                                                                        <tr>
                                                                                                                                                            <td class="tdcaption"
                                                                                                                                                                width="27%">
                                                                                                                                                                Publish
                                                                                                                                                                Slip(s)
                                                                                                                                                                According
                                                                                                                                                                to
                                                                                                                                                            </td>
                                                                                                                                                            <td class="tdcolon"
                                                                                                                                                                style="width: 6%">
                                                                                                                                                                :
                                                                                                                                                            </td>
                                                                                                                                                            <td
                                                                                                                                                                class="tdvd">
                                                                                                                                                                <asp:RadioButtonList
                                                                                                                                                                    ID="RblGrpbyPublish"
                                                                                                                                                                    runat="server"
                                                                                                                                                                    CssClass="Radio"
                                                                                                                                                                    AutoPostBack="true"
                                                                                                                                                                    RepeatDirection="Horizontal"
                                                                                                                                                                    onchange="javascript:return PleaseWaitWithDailog();">
                                                                                                                                                                    <asp:ListItem
                                                                                                                                                                        Text="Employee Wise"
                                                                                                                                                                        Value=""
                                                                                                                                                                        Selected="True">
                                                                                                                                                                    </asp:ListItem>
                                                                                                                                                                    <asp:ListItem
                                                                                                                                                                        Text="Unit Wise"
                                                                                                                                                                        Value="UNT">
                                                                                                                                                                    </asp:ListItem>
                                                                                                                                                                </asp:RadioButtonList>
                                                                                                                                                            </td>
                                                                                                                                                        </tr>
                                                                                                                                                    </table>
                                                                                                                                                </td>
                                                                                                                                                <td colspan="3"
                                                                                                                                                    align="right">
                                                                                                                                                    <table
                                                                                                                                                        class="Message"
                                                                                                                                                        width="100%"
                                                                                                                                                        cellpadding="0"
                                                                                                                                                        cellspacing="0">
                                                                                                                                                        <tr>
                                                                                                                                                            <td
                                                                                                                                                                class="tdcaption">
                                                                                                                                                                <img
                                                                                                                                                                    src="ImagesCA/Note.gif" />
                                                                                                                                                            </td>
                                                                                                                                                            <td>
                                                                                                                                                                <span
                                                                                                                                                                    style="color: red"><b>Note:</b></span><span
                                                                                                                                                                    style="font-size: 7.5pt; color: Black;
                                                                                                                        line-height: 115%; font-family: Verdana; mso-fareast-font-family: 'Times New Roman';
                                                                                                                        mso-bidi-font-family: 'Times New Roman'; mso-ansi-language: EN-IN; mso-fareast-language: EN-US;
                                                                                                                        mso-bidi-language: AR-SA">&nbsp;</span><b>Employee
                                                                                                                                                                    Wise:</b>
                                                                                                                                                                Publish
                                                                                                                                                                salary
                                                                                                                                                                slip(s)
                                                                                                                                                                for
                                                                                                                                                                each
                                                                                                                                                                employee
                                                                                                                                                                separately.<br />
                                                                                                                                                                <b>Unit
                                                                                                                                                                    Wise:</b>
                                                                                                                                                                Publish
                                                                                                                                                                salary
                                                                                                                                                                slip(s)
                                                                                                                                                                unit
                                                                                                                                                                wise
                                                                                                                                                                as
                                                                                                                                                                single
                                                                                                                                                                zip
                                                                                                                                                                file
                                                                                                                                                                for
                                                                                                                                                                each
                                                                                                                                                                unit
                                                                                                                                                                and
                                                                                                                                                                email
                                                                                                                                                                to
                                                                                                                                                                unit
                                                                                                                                                                authority.
                                                                                                                                                                <b>"PDF
                                                                                                                                                                    Password
                                                                                                                                                                    Type"</b>
                                                                                                                                                                will
                                                                                                                                                                not
                                                                                                                                                                work
                                                                                                                                                                in
                                                                                                                                                                this
                                                                                                                                                                condition.
                                                                                                                                                            </td>
                                                                                                                                                        </tr>
                                                                                                                                                    </table>
                                                                                                                                                </td>
                                                                                                                                            </tr>
                                                                                                                                            <tr id="trselall"
                                                                                                                                                style="display: none"
                                                                                                                                                runat="server">
                                                                                                                                                <td
                                                                                                                                                    colspan="6">
                                                                                                                                                    <table
                                                                                                                                                        border="0"
                                                                                                                                                        cellpadding="0"
                                                                                                                                                        cellspacing="0"
                                                                                                                                                        width="100%">
                                                                                                                                                        <tr>
                                                                                                                                                            <td width="45%"
                                                                                                                                                                valign="top">
                                                                                                                                                                <asp:CheckBox
                                                                                                                                                                    ID="chkmailformat"
                                                                                                                                                                    runat="server"
                                                                                                                                                                    OnCheckedChanged="chkmailformat_CheckedChanged"
                                                                                                                                                                    Checked="false"
                                                                                                                                                                    AutoPostBack="True"
                                                                                                                                                                    CssClass="tdcaption"
                                                                                                                                                                    Text="Select if you want to create Email body." />
                                                                                                                                                            </td>
                                                                                                                                                            <td
                                                                                                                                                                width="70%">
                                                                                                                                                                <table
                                                                                                                                                                    class="Message"
                                                                                                                                                                    border="0"
                                                                                                                                                                    cellpadding="0"
                                                                                                                                                                    cellspacing="0"
                                                                                                                                                                    width="100%"
                                                                                                                                                                    align="right">
                                                                                                                                                                    <tr>
                                                                                                                                                                        <td
                                                                                                                                                                            valign="top">
                                                                                                                                                                            <img
                                                                                                                                                                                src="ImagesCA/Note.gif" />
                                                                                                                                                                        </td>
                                                                                                                                                                        <td>
                                                                                                                                                                            <span
                                                                                                                                                                                style="color: red"><b>Note:</b></span><span
                                                                                                                                                                                style="font-size: 7.5pt; color: Black;
                                                                                                                                    line-height: 115%; font-family: Verdana; mso-fareast-font-family: 'Times New Roman';
                                                                                                                                    mso-bidi-font-family: 'Times New Roman'; mso-ansi-language: EN-IN; mso-fareast-language: EN-US;
                                                                                                                                    mso-bidi-language: AR-SA">&nbsp;</span> Select
                                                                                                                                                                            if
                                                                                                                                                                            you
                                                                                                                                                                            want
                                                                                                                                                                            to
                                                                                                                                                                            create
                                                                                                                                                                            Email
                                                                                                                                                                            body.
                                                                                                                                                                            (<span
                                                                                                                                                                                style="color: red">
                                                                                                                                                                                <b>Custom
                                                                                                                                                                                    message
                                                                                                                                                                                    does
                                                                                                                                                                                    not
                                                                                                                                                                                    work</b></span>
                                                                                                                                                                            when
                                                                                                                                                                            you
                                                                                                                                                                            use
                                                                                                                                                                            the
                                                                                                                                                                            option
                                                                                                                                                                            Publish
                                                                                                                                                                            and
                                                                                                                                                                            Email
                                                                                                                                                                            to
                                                                                                                                                                            CC/BCC
                                                                                                                                                                            because
                                                                                                                                                                            we
                                                                                                                                                                            can�t
                                                                                                                                                                            pick
                                                                                                                                                                            the
                                                                                                                                                                            empcode
                                                                                                                                                                            and
                                                                                                                                                                            other
                                                                                                                                                                            fields
                                                                                                                                                                            which
                                                                                                                                                                            you
                                                                                                                                                                            use
                                                                                                                                                                            to
                                                                                                                                                                            define
                                                                                                                                                                            the
                                                                                                                                                                            custom
                                                                                                                                                                            message).
                                                                                                                                                                        </td>
                                                                                                                                                                    </tr>
                                                                                                                                                                </table>
                                                                                                                                                            </td>
                                                                                                                                                        </tr>
                                                                                                                                                        <tr id="tblmail"
                                                                                                                                                            runat="server"
                                                                                                                                                            style="display: none">
                                                                                                                                                            <td
                                                                                                                                                                colspan="2">
                                                                                                                                                                <table
                                                                                                                                                                    border="0"
                                                                                                                                                                    cellpadding="0"
                                                                                                                                                                    cellspacing="0"
                                                                                                                                                                    width="100%">
                                                                                                                                                                    <tr>
                                                                                                                                                                        <td colspan="8"
                                                                                                                                                                            valign="top"
                                                                                                                                                                            style="height: 36px">
                                                                                                                                                                            <table
                                                                                                                                                                                border="0"
                                                                                                                                                                                cellpadding="0"
                                                                                                                                                                                cellspacing="0"
                                                                                                                                                                                width="100%">
                                                                                                                                                                                <tr>
                                                                                                                                                                                    <td>
                                                                                                                                                                                        <%=_objCommon.makeHeading("Create
                                                                                                                                                                                            Custom
                                                                                                                                                                                            Email
                                                                                                                                                                                            Message")%>
                                                                                                                                                                                    </td>
                                                                                                                                                                                </tr>
                                                                                                                                                                            </table>
                                                                                                                                                                        </td>
                                                                                                                                                                    </tr>
                                                                                                                                                                    <tr>
                                                                                                                                                                        <td class="tdcaption"
                                                                                                                                                                            valign="top"
                                                                                                                                                                            width="12%">
                                                                                                                                                                            Header
                                                                                                                                                                        </td>
                                                                                                                                                                        <td class="tdcolon"
                                                                                                                                                                            valign="top">
                                                                                                                                                                            :
                                                                                                                                                                        </td>
                                                                                                                                                                        <td class="tdvd"
                                                                                                                                                                            colspan="5">
                                                                                                                                                                            <asp:TextBox
                                                                                                                                                                                ID="txtheader"
                                                                                                                                                                                CssClass="Textbox"
                                                                                                                                                                                runat="server"
                                                                                                                                                                                Width="519px"
                                                                                                                                                                                onblur="Javascript:return Clicked(this.id,'Y');">
                                                                                                                                                                            </asp:TextBox>
                                                                                                                                                                            *
                                                                                                                                                                        </td>
                                                                                                                                                                        <td
                                                                                                                                                                            class="tdvd">
                                                                                                                                                                        </td>
                                                                                                                                                                    </tr>
                                                                                                                                                                    <tr>
                                                                                                                                                                        <td class="trgap"
                                                                                                                                                                            colspan="8"
                                                                                                                                                                            valign="top">
                                                                                                                                                                        </td>
                                                                                                                                                                    </tr>
                                                                                                                                                                    <tr>
                                                                                                                                                                        <td class="tdcaption"
                                                                                                                                                                            valign="top"
                                                                                                                                                                            width="12%">
                                                                                                                                                                            Contents
                                                                                                                                                                        </td>
                                                                                                                                                                        <td class="tdcaption"
                                                                                                                                                                            valign="top">
                                                                                                                                                                            :
                                                                                                                                                                        </td>
                                                                                                                                                                        <td class="tdvd"
                                                                                                                                                                            colspan="5"
                                                                                                                                                                            valign="top">
                                                                                                                                                                            <div
                                                                                                                                                                                style="height: 225px">
                                                                                                                                                                                <cc1:EasyWebEdit
                                                                                                                                                                                    ID="EasyWebMAilBody"
                                                                                                                                                                                    runat="server"
                                                                                                                                                                                    AllowFileUpload="True"
                                                                                                                                                                                    FileManager="../"
                                                                                                                                                                                    FileUploadDir="/tempattachment/"
                                                                                                                                                                                    onclick="Clicked(this.id,'N');"
                                                                                                                                                                                    ImagesLocation="../SAPayroll/WebEditorImg"
                                                                                                                                                                                    Width="523px"
                                                                                                                                                                                    Height="131px">
                                                                                                                                                                                </cc1:EasyWebEdit>
                                                                                                                                                                            </div>
                                                                                                                                                                            *
                                                                                                                                                                        </td>
                                                                                                                                                                        <td class="tdcaption"
                                                                                                                                                                            valign="top">
                                                                                                                                                                            <asp:ListBox
                                                                                                                                                                                ID="lsthelp"
                                                                                                                                                                                runat="server"
                                                                                                                                                                                CssClass="dropdownlist"
                                                                                                                                                                                Height="100px"
                                                                                                                                                                                ondblclick="Javascript:return DblClicked(this.form);"
                                                                                                                                                                                Width="227px">
                                                                                                                                                                                <asp:ListItem
                                                                                                                                                                                    Text="Employee Code:[EMPCODE]"
                                                                                                                                                                                    Value="[EMPCODE]">
                                                                                                                                                                                </asp:ListItem>
                                                                                                                                                                                <asp:ListItem
                                                                                                                                                                                    Text="Employee Name:[EMPNAME]"
                                                                                                                                                                                    Value="[EMPNAME]">
                                                                                                                                                                                </asp:ListItem>
                                                                                                                                                                                <asp:ListItem
                                                                                                                                                                                    Text="Company Name:[COMPNAME]"
                                                                                                                                                                                    Value="[COMPNAME]">
                                                                                                                                                                                </asp:ListItem>
                                                                                                                                                                                <asp:ListItem
                                                                                                                                                                                    Text="Company Address :[COMPADDRESS]"
                                                                                                                                                                                    Value="[COMPADDRESS]">
                                                                                                                                                                                </asp:ListItem>
                                                                                                                                                                                <asp:ListItem
                                                                                                                                                                                    Text="Month :[MONTH]"
                                                                                                                                                                                    Value="[MONTH]">
                                                                                                                                                                                </asp:ListItem>
                                                                                                                                                                                <asp:ListItem
                                                                                                                                                                                    Text="Year :[YEAR]"
                                                                                                                                                                                    Value="[YEAR]">
                                                                                                                                                                                </asp:ListItem>
                                                                                                                                                                                <asp:ListItem
                                                                                                                                                                                    Text="Click hear to add Break[BR]"
                                                                                                                                                                                    Value="&lt;BR&gt;">
                                                                                                                                                                                </asp:ListItem>
                                                                                                                                                                            </asp:ListBox>
                                                                                                                                                                        </td>
                                                                                                                                                                    </tr>
                                                                                                                                                                    <tr>
                                                                                                                                                                        <td class="trgap"
                                                                                                                                                                            colspan="8"
                                                                                                                                                                            valign="top">
                                                                                                                                                                        </td>
                                                                                                                                                                    </tr>
                                                                                                                                                                    <tr>
                                                                                                                                                                        <td class="tdcaption"
                                                                                                                                                                            valign="top"
                                                                                                                                                                            width="12%">
                                                                                                                                                                            Footer
                                                                                                                                                                        </td>
                                                                                                                                                                        <td class="tdcolon"
                                                                                                                                                                            valign="top">
                                                                                                                                                                            :
                                                                                                                                                                        </td>
                                                                                                                                                                        <td class="tdvd"
                                                                                                                                                                            colspan="6"
                                                                                                                                                                            valign="top">
                                                                                                                                                                            <asp:TextBox
                                                                                                                                                                                ID="Textfooter"
                                                                                                                                                                                CssClass="Textbox"
                                                                                                                                                                                runat="server"
                                                                                                                                                                                Width="519px"
                                                                                                                                                                                onblur="Javascript:return Clicked(this.id,'Y');">
                                                                                                                                                                            </asp:TextBox>
                                                                                                                                                                            *
                                                                                                                                                                            <asp:Button
                                                                                                                                                                                ID="Btndelete"
                                                                                                                                                                                runat="server"
                                                                                                                                                                                CssClass="btn"
                                                                                                                                                                                Width="120px"
                                                                                                                                                                                Text="Delete Mail Content"
                                                                                                                                                                                OnClientClick="javascript:return ConfirmDelete('Would you like to delete Mail Format.');" />
                                                                                                                                                                        </td>
                                                                                                                                                                    </tr>
                                                                                                                                                                    <tr>
                                                                                                                                                                        <td class="trupbtn"
                                                                                                                                                                            colspan="8"
                                                                                                                                                                            valign="top">
                                                                                                                                                                        </td>
                                                                                                                                                                    </tr>
                                                                                                                                                                </table>
                                                                                                                                                            </td>
                                                                                                                                                        </tr>
                                                                                                                                                    </table>
                                                                                                                                                </td>
                                                                                                                                            </tr>
                                                                                                                                </table>
                                                                                                                            </ContentTemplate>
                                                                                                                        </asp:UpdatePanel>
                                                                                                                    </td>
                                                                                                                </tr>
                                                                                                                <tr
                                                                                                                    id="trColor">
                                                                                                                    <td class="tdcaption"
                                                                                                                        colspan="6">
                                                                                                                        <asp:UpdatePanel
                                                                                                                            ID="updMail"
                                                                                                                            runat="server">
                                                                                                                            <ContentTemplate>
                                                                                                                                <table
                                                                                                                                    border="0"
                                                                                                                                    cellpadding="0"
                                                                                                                                    cellspacing="0"
                                                                                                                                    width="100%">
                                                                                                                                    <%--<tr
                                                                                                                                        class="trupbtn">
                                                                                                                                        <td>
                                                                                                                                        </td>
                                                                                                                </tr>
                                                                                                                --%>
                                                                                                                <tr>
                                                                                                                    <td>
                                                                                                                        <table
                                                                                                                            border="0"
                                                                                                                            cellpadding="3"
                                                                                                                            cellspacing="0"
                                                                                                                            width="100%">
                                                                                                                            <tr>
                                                                                                                                <td class="Tdcaption"
                                                                                                                                    valign="top"
                                                                                                                                    width="12%">
                                                                                                                                </td>
                                                                                                                                <td class="TdColon"
                                                                                                                                    valign="top"
                                                                                                                                    width="3%">
                                                                                                                                </td>
                                                                                                                                <td valign="top"
                                                                                                                                    style="width: 30%">
                                                                                                                                    <table
                                                                                                                                        border="0"
                                                                                                                                        cellpadding="0"
                                                                                                                                        cellspacing="0"
                                                                                                                                        width="80%">
                                                                                                                                        <tr>
                                                                                                                                            <td id="TdSearch1"
                                                                                                                                                runat="server"
                                                                                                                                                valign="top"
                                                                                                                                                style="width: 27%">
                                                                                                                                                <asp:Button
                                                                                                                                                    ID="Btnsearch"
                                                                                                                                                    runat="server"
                                                                                                                                                    CssClass="btn"
                                                                                                                                                    Text="Search"
                                                                                                                                                    OnClientClick="javascript:return btnSearch_Click(this)" />
                                                                                                                                            </td>
                                                                                                                                            <td id="TdSearch2"
                                                                                                                                                runat="server"
                                                                                                                                                class="tdbetweenlbl"
                                                                                                                                                style="width: 1%">
                                                                                                                                            </td>
                                                                                                                                            <td
                                                                                                                                                valign="top">
                                                                                                                                                <asp:Button
                                                                                                                                                    ID="btnreset"
                                                                                                                                                    OnClientClick="ResetCtrl();"
                                                                                                                                                    class="btn"
                                                                                                                                                    Text="RESET"
                                                                                                                                                    runat="server" />
                                                                                                                                            </td>
                                                                                                                                        </tr>
                                                                                                                                    </table>
                                                                                                                                </td>
                                                                                                                                <td valign="top"
                                                                                                                                    align="right">
                                                                                                                                    <table
                                                                                                                                        class="Message"
                                                                                                                                        border="0"
                                                                                                                                        cellpadding="0"
                                                                                                                                        cellspacing="0"
                                                                                                                                        width="100%"
                                                                                                                                        align="right">
                                                                                                                                        <tr>
                                                                                                                                            <td
                                                                                                                                                valign="top">
                                                                                                                                                <img
                                                                                                                                                    src="ImagesCA/Note.gif" />
                                                                                                                                            </td>
                                                                                                                                            <td>
                                                                                                                                                <span
                                                                                                                                                    style="color: red"><b>Note:</b></span><span
                                                                                                                                                    style="font-size: 7.5pt; color: Black;
                                                                                                                                    line-height: 115%; font-family: Verdana; mso-fareast-font-family: 'Times New Roman';
                                                                                                                                    mso-bidi-font-family: 'Times New Roman'; mso-ansi-language: EN-IN; mso-fareast-language: EN-US;
                                                                                                                                    mso-bidi-language: AR-SA">&nbsp;</span>
                                                                                                                                                You
                                                                                                                                                can
                                                                                                                                                now
                                                                                                                                                send
                                                                                                                                                zip
                                                                                                                                                file
                                                                                                                                                of
                                                                                                                                                the
                                                                                                                                                payslips
                                                                                                                                                to
                                                                                                                                                the
                                                                                                                                                email
                                                                                                                                                id
                                                                                                                                                you
                                                                                                                                                mention
                                                                                                                                                in
                                                                                                                                                CC
                                                                                                                                                /
                                                                                                                                                BCC.
                                                                                                                                                You
                                                                                                                                                can
                                                                                                                                                ONLY
                                                                                                                                                send
                                                                                                                                                PDF
                                                                                                                                                payslip
                                                                                                                                                and
                                                                                                                                                not
                                                                                                                                                HTML.
                                                                                                                                                First
                                                                                                                                                select
                                                                                                                                                PDF
                                                                                                                                                and
                                                                                                                                                then
                                                                                                                                                choose
                                                                                                                                                the
                                                                                                                                                option
                                                                                                                                                �Publish
                                                                                                                                                and
                                                                                                                                                email
                                                                                                                                                to
                                                                                                                                                CC
                                                                                                                                                /
                                                                                                                                                BCC�.
                                                                                                                                            </td>
                                                                                                                                        </tr>
                                                                                                                                    </table>
                                                                                                                                </td>
                                                                                                                            </tr>
                                                                                                                            <tr>
                                                                                                                                <td class="Tdcaption"
                                                                                                                                    valign="top"
                                                                                                                                    width="12%">
                                                                                                                                </td>
                                                                                                                                <td class="TdColon"
                                                                                                                                    valign="top">
                                                                                                                                </td>
                                                                                                                                <td
                                                                                                                                    colspan="2">
                                                                                                                                    <div id="divlblMsg"
                                                                                                                                        style="overflow: auto; height: 40px; width: 100%">
                                                                                                                                        <asp:Label
                                                                                                                                            ID="lblmsg"
                                                                                                                                            runat="server"
                                                                                                                                            CssClass="usermessage">
                                                                                                                                        </asp:Label>
                                                                                                                                    </div>
                                                                                                                                </td>
                                                                                                                            </tr>
                                                                                                                        </table>
                                                                                                                    </td>
                                                                                                                </tr>
                                                                                                                <tr
                                                                                                                    style="display: none">
                                                                                                                    <td>
                                                                                                                        <input
                                                                                                                            id="HidYear"
                                                                                                                            runat="server"
                                                                                                                            type="hidden" />
                                                                                                                        <input
                                                                                                                            id="hid"
                                                                                                                            runat="server"
                                                                                                                            type="hidden" />
                                                                                                                        <input
                                                                                                                            id="Hidden1"
                                                                                                                            runat="server"
                                                                                                                            type="hidden" />
                                                                                                                        <input
                                                                                                                            id="HidRepId"
                                                                                                                            runat="server"
                                                                                                                            name="HidRepId"
                                                                                                                            type="hidden" />
                                                                                                                    </td>
                                                                                                                </tr>
                                                                                                                <tr id="tableshow"
                                                                                                                    runat="server"
                                                                                                                    style="display: none">
                                                                                                                    <td>
                                                                                                                        <table
                                                                                                                            border="0"
                                                                                                                            cellpadding="0"
                                                                                                                            cellspacing="0"
                                                                                                                            width="100%">
                                                                                                                            <tr id="TrDg"
                                                                                                                                runat="server"
                                                                                                                                style="display: none;">
                                                                                                                                <td
                                                                                                                                    colspan="4">
                                                                                                                                    <table
                                                                                                                                        cellpadding="0"
                                                                                                                                        cellspacing="0"
                                                                                                                                        border="0"
                                                                                                                                        width="100%">
                                                                                                                                        <tr>
                                                                                                                                            <td class="FormSubHeading"
                                                                                                                                                width="100%"
                                                                                                                                                colspan="4">
                                                                                                                                                <table
                                                                                                                                                    cellspacing="0"
                                                                                                                                                    cellpadding="0"
                                                                                                                                                    width="100%"
                                                                                                                                                    border="0">
                                                                                                                                                    <%=_objCommon.makeHeading("List
                                                                                                                                                        of
                                                                                                                                                        Employee(s)")%>
                                                                                                                                                </table>
                                                                                                                                            </td>
                                                                                                                                        </tr>
                                                                                                                                        <tr>
                                                                                                                                            <td>
                                                                                                                                                <table
                                                                                                                                                    cellpadding="0"
                                                                                                                                                    cellspacing="0"
                                                                                                                                                    border="0">
                                                                                                                                                    <tr>
                                                                                                                                                        <td>
                                                                                                                                                            <asp:Label
                                                                                                                                                                ID="LblL1"
                                                                                                                                                                runat="server"
                                                                                                                                                                Height="12px"
                                                                                                                                                                Width="12px"
                                                                                                                                                                Text="&nbsp;&nbsp;"
                                                                                                                                                                BackColor="#ffc0cb">
                                                                                                                                                            </asp:Label>
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="TDBetweenBtn">
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="TDCaption">
                                                                                                                                                            <font
                                                                                                                                                                color='#ffc0cb'>
                                                                                                                                                                Either
                                                                                                                                                                email
                                                                                                                                                                id
                                                                                                                                                                not
                                                                                                                                                                available
                                                                                                                                                                or
                                                                                                                                                                salary
                                                                                                                                                                withheld!
                                                                                                                                                            </font>
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="TDBetweenBtn">
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="TDBetweenBtn">
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="TDBetweenBtn">
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="TDBetweenBtn">
                                                                                                                                                        </td>
                                                                                                                                                        <td>
                                                                                                                                                            <asp:Label
                                                                                                                                                                ID="LblL2"
                                                                                                                                                                runat="server"
                                                                                                                                                                Height="12px"
                                                                                                                                                                Width="12px"
                                                                                                                                                                Text="&nbsp;&nbsp;"
                                                                                                                                                                BackColor="#87ceeb">
                                                                                                                                                            </asp:Label>
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="TDBetweenBtn">
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="TDCaption">
                                                                                                                                                            <font
                                                                                                                                                                color='#87ceeb'>
                                                                                                                                                                Salary
                                                                                                                                                                not
                                                                                                                                                                processed
                                                                                                                                                                or
                                                                                                                                                                email
                                                                                                                                                                already
                                                                                                                                                                sent!
                                                                                                                                                            </font>
                                                                                                                                                        </td>
                                                                                                                                                    </tr>
                                                                                                                                                </table>
                                                                                                                                            </td>
                                                                                                                                        </tr>
                                                                                                                                        <tr>
                                                                                                                                            <td>
                                                                                                                                            </td>
                                                                                                                                        </tr>
                                                                                                                                        <tr>
                                                                                                                                            <td>
                                                                                                                                                <div
                                                                                                                                                    id="DvEmp">
                                                                                                                                                    <asp:DataGrid
                                                                                                                                                        ID="DgPayslip"
                                                                                                                                                        runat="server"
                                                                                                                                                        AutoGenerateColumns="False"
                                                                                                                                                        CssClass="DataGridBody"
                                                                                                                                                        DataKeyField="fk_emp_code"
                                                                                                                                                        Width="100%">
                                                                                                                                                        <SelectedItemStyle
                                                                                                                                                            CssClass="DataGridSelectedItemStyle">
                                                                                                                                                        </SelectedItemStyle>
                                                                                                                                                        <ItemStyle
                                                                                                                                                            CssClass="DataGridItemStyle">
                                                                                                                                                        </ItemStyle>
                                                                                                                                                        <HeaderStyle
                                                                                                                                                            CssClass="DataGridHeaderStyle">
                                                                                                                                                        </HeaderStyle>
                                                                                                                                                        <FooterStyle
                                                                                                                                                            CssClass="DataGridFooterStyle">
                                                                                                                                                        </FooterStyle>
                                                                                                                                                        <AlternatingItemStyle
                                                                                                                                                            CssClass="AlternatingRowStyle" />
                                                                                                                                                        <Columns>
                                                                                                                                                            <asp:TemplateColumn
                                                                                                                                                                HeaderStyle-Width="3%">
                                                                                                                                                                <HeaderTemplate>
                                                                                                                                                                    <input
                                                                                                                                                                        id="Checkbox1"
                                                                                                                                                                        checked="checked"
                                                                                                                                                                        class="chk"
                                                                                                                                                                        onclick="checkUncheckGridCheckBox('DvEmp', this,'chkEmpHold');"
                                                                                                                                                                        type="checkbox">
                                                                                                                                                                </HeaderTemplate>
                                                                                                                                                                <ItemTemplate>
                                                                                                                                                                    <asp:CheckBox
                                                                                                                                                                        ID="chkEmpHold"
                                                                                                                                                                        runat="server"
                                                                                                                                                                        Checked="True"
                                                                                                                                                                        CssClass="chk"
                                                                                                                                                                        name="chkEmpHold"
                                                                                                                                                                        onclick="checkUncheckHeaderByRowChk('DvEmp', 'Checkbox1','chkEmpHold');" />
                                                                                                                                                                </ItemTemplate>
                                                                                                                                                            </asp:TemplateColumn>
                                                                                                                                                            <asp:BoundColumn
                                                                                                                                                                DataField="fk_emp_code"
                                                                                                                                                                HeaderText="Emp Code"
                                                                                                                                                                HeaderStyle-Width="7%" />
                                                                                                                                                            <asp:BoundColumn
                                                                                                                                                                DataField="EmpName"
                                                                                                                                                                HeaderText="Employee Name"
                                                                                                                                                                HeaderStyle-Width="15%" />
                                                                                                                                                            <asp:BoundColumn
                                                                                                                                                                DataField="Dept_desc"
                                                                                                                                                                HeaderText="Department"
                                                                                                                                                                HeaderStyle-Width="15%" />
                                                                                                                                                            <asp:BoundColumn
                                                                                                                                                                DataField="desig_desc"
                                                                                                                                                                HeaderText="Designation"
                                                                                                                                                                HeaderStyle-Width="15%" />
                                                                                                                                                            <asp:BoundColumn
                                                                                                                                                                DataField="Email"
                                                                                                                                                                HeaderText="EmailID"
                                                                                                                                                                HeaderStyle-Width="18%" />
                                                                                                                                                            <asp:BoundColumn
                                                                                                                                                                DataField="Status"
                                                                                                                                                                HeaderText="Status"
                                                                                                                                                                HeaderStyle-Width="10%" />
                                                                                                                                                            <asp:BoundColumn
                                                                                                                                                                DataField="Sent_Date"
                                                                                                                                                                HeaderText="Last Mail Sent"
                                                                                                                                                                HeaderStyle-Width="11%" />
                                                                                                                                                            <asp:TemplateColumn
                                                                                                                                                                HeaderText="Preview"
                                                                                                                                                                HeaderStyle-Width="6%">
                                                                                                                                                                <ItemTemplate>
                                                                                                                                                                    <asp:LinkButton
                                                                                                                                                                        ID="LinkButton1"
                                                                                                                                                                        runat="server"
                                                                                                                                                                        CommandName="Preview"
                                                                                                                                                                        CausesValidation="true">
                                                                                                                                                                        <img alt=""
                                                                                                                                                                            src="../ImagesCA/formview.ico"
                                                                                                                                                                            border="0" />
                                                                                                                                                                    </asp:LinkButton>
                                                                                                                                                                </ItemTemplate>
                                                                                                                                                            </asp:TemplateColumn>
                                                                                                                                                            <asp:BoundColumn
                                                                                                                                                                DataField="SalHold"
                                                                                                                                                                HeaderText="SalHold"
                                                                                                                                                                Visible="False" />
                                                                                                                                                            <asp:BoundColumn
                                                                                                                                                                DataField="EmpLastName"
                                                                                                                                                                HeaderText="EmpLastName"
                                                                                                                                                                Visible="False" />
                                                                                                                                                        </Columns>
                                                                                                                                                        <PagerStyle
                                                                                                                                                            CssClass="Datagridpaging"
                                                                                                                                                            HorizontalAlign="Right"
                                                                                                                                                            Mode="NumericPages" />
                                                                                                                                                    </asp:DataGrid>
                                                                                                                                                </div>
                                                                                                                                                <asp:Label
                                                                                                                                                    ID="lblmsg1"
                                                                                                                                                    runat="server"
                                                                                                                                                    CssClass="usermessage">
                                                                                                                                                </asp:Label>
                                                                                                                                            </td>
                                                                                                                                        </tr>
                                                                                                                                    </table>
                                                                                                                                </td>
                                                                                                                            </tr>
                                                                                                                            <tr>
                                                                                                                                <td>
                                                                                                                                    <table
                                                                                                                                        id="Table1"
                                                                                                                                        border="0"
                                                                                                                                        align="center"
                                                                                                                                        cellpadding="0"
                                                                                                                                        cellspacing="0"
                                                                                                                                        width="100%"
                                                                                                                                        runat="server">
                                                                                                                                        <tr id="Tr2"
                                                                                                                                            runat="server">
                                                                                                                                            <td
                                                                                                                                                colspan="2">
                                                                                                                                                <table
                                                                                                                                                    id="tbl1"
                                                                                                                                                    runat="server"
                                                                                                                                                    border="0"
                                                                                                                                                    cellpadding="3"
                                                                                                                                                    cellspacing="0"
                                                                                                                                                    width="100%">
                                                                                                                                                    <tr>
                                                                                                                                                        <td class="tdcaption"
                                                                                                                                                            width="12%">
                                                                                                                                                            CC
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="tdcolon">
                                                                                                                                                            :
                                                                                                                                                        </td>
                                                                                                                                                        <td>
                                                                                                                                                            <asp:TextBox
                                                                                                                                                                ID="txtccc"
                                                                                                                                                                runat="server"
                                                                                                                                                                CssClass="Textbox"
                                                                                                                                                                Width="519px">
                                                                                                                                                            </asp:TextBox>
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="tdvd">
                                                                                                                                                            <strong>Note
                                                                                                                                                                :<span
                                                                                                                                                                    lang="EN-US"
                                                                                                                                                                    style="font-size: 12pt; font-family: Wingdings;
                                                                                                                                                mso-ascii-font-family: 'Times New Roman'; mso-fareast-font-family: 'Times New Roman';
                                                                                                                                                mso-hansi-font-family: 'Times New Roman'; mso-bidi-font-family: 'Times New Roman';
                                                                                                                                                mso-ansi-language: EN-US; mso-fareast-language: EN-IN; mso-bidi-language: AR-SA;
                                                                                                                                                mso-char-type: symbol; mso-symbol-font-family: Wingdings"><span></span></span>
                                                                                                                                                            </strong><span
                                                                                                                                                                style="font-size: 7.5pt; color: black; line-height: 115%; font-family: Verdana;
                                                                                                                                                mso-fareast-font-family: 'Times New Roman'; mso-bidi-font-family: 'Times New Roman';
                                                                                                                                                mso-ansi-language: EN-IN; mso-fareast-language: EN-US; mso-bidi-language: AR-SA">
                                                                                                                                                                Max
                                                                                                                                                                10
                                                                                                                                                                email
                                                                                                                                                                id�s
                                                                                                                                                                and
                                                                                                                                                                use
                                                                                                                                                                comma
                                                                                                                                                                as
                                                                                                                                                                delimiter
                                                                                                                                                                (separator)</span>
                                                                                                                                                        </td>
                                                                                                                                                    </tr>
                                                                                                                                                    <tr>
                                                                                                                                                        <td class="tdcaption"
                                                                                                                                                            width="12%">
                                                                                                                                                            BCC
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="tdcolon">
                                                                                                                                                            :
                                                                                                                                                        </td>
                                                                                                                                                        <td>
                                                                                                                                                            <asp:TextBox
                                                                                                                                                                ID="txtBCC"
                                                                                                                                                                runat="server"
                                                                                                                                                                CssClass="Textbox"
                                                                                                                                                                Width="519px">
                                                                                                                                                            </asp:TextBox>
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="tdvd">
                                                                                                                                                            <strong>Note
                                                                                                                                                                :<span
                                                                                                                                                                    lang="EN-US"
                                                                                                                                                                    style="font-size: 12pt; font-family: Wingdings;
                                                                                                                                                mso-ascii-font-family: 'Times New Roman'; mso-fareast-font-family: 'Times New Roman';
                                                                                                                                                mso-hansi-font-family: 'Times New Roman'; mso-bidi-font-family: 'Times New Roman';
                                                                                                                                                mso-ansi-language: EN-US; mso-fareast-language: EN-IN; mso-bidi-language: AR-SA;
                                                                                                                                                mso-char-type: symbol; mso-symbol-font-family: Wingdings"><span></span></span>
                                                                                                                                                            </strong><span
                                                                                                                                                                style="font-size: 7.5pt; color: black; line-height: 115%; font-family: Verdana;
                                                                                                                                                mso-fareast-font-family: 'Times New Roman'; mso-bidi-font-family: 'Times New Roman';
                                                                                                                                                mso-ansi-language: EN-IN; mso-fareast-language: EN-US; mso-bidi-language: AR-SA">
                                                                                                                                                                Max
                                                                                                                                                                10
                                                                                                                                                                email
                                                                                                                                                                id�s
                                                                                                                                                                and
                                                                                                                                                                use
                                                                                                                                                                comma
                                                                                                                                                                as
                                                                                                                                                                delimiter
                                                                                                                                                                (separator)</span>
                                                                                                                                                        </td>
                                                                                                                                                    </tr>
                                                                                                                                                </table>
                                                                                                                                            </td>
                                                                                                                                        </tr>
                                                                                                                                        <%--<tr
                                                                                                                                            id="Tr3"
                                                                                                                                            runat="server">
                                                                                                                                            <td
                                                                                                                                                colspan="2">
                                                                                                                                            </td>
                                                                                                                            </tr>
                                                                                                                            --%>
                                                                                                                            <tr id="trbutton"
                                                                                                                                runat="server">
                                                                                                                                <td style="height: 91px"
                                                                                                                                    colspan="2"
                                                                                                                                    valign="top">
                                                                                                                                    <table
                                                                                                                                        id="Table2"
                                                                                                                                        border="0"
                                                                                                                                        cellpadding="0"
                                                                                                                                        cellspacing="0"
                                                                                                                                        runat="server"
                                                                                                                                        width="100%">
                                                                                                                                        <tr>
                                                                                                                                            <td
                                                                                                                                                width="15%">
                                                                                                                                            </td>
                                                                                                                                            <td>
                                                                                                                                                <table
                                                                                                                                                    id="Table3"
                                                                                                                                                    border="0"
                                                                                                                                                    cellpadding="0"
                                                                                                                                                    cellspacing="0"
                                                                                                                                                    runat="server">
                                                                                                                                                    <tr>
                                                                                                                                                        <td valign="tdbetweenbtn"
                                                                                                                                                            align="left"
                                                                                                                                                            colspan="8">
                                                                                                                                                            <asp:Button
                                                                                                                                                                ID="BtnSend"
                                                                                                                                                                runat="server"
                                                                                                                                                                CssClass="custombtn"
                                                                                                                                                                Text="Send email to employees and CC/BCC Email ID�s"
                                                                                                                                                                CausesValidation="True"
                                                                                                                                                                OnClientClick="javascript:return empchecked('BtnSend',this);"
                                                                                                                                                                Width="300px"
                                                                                                                                                                style="margin-top: 8px;">
                                                                                                                                                            </asp:Button>

                                                                                                                                                            <asp:Button
                                                                                                                                                                ID="btnSave"
                                                                                                                                                                runat="server"
                                                                                                                                                                CausesValidation="False"
                                                                                                                                                                CssClass="custombtn"
                                                                                                                                                                OnClientClick="javascript:return empchecked('btnSave',this);"
                                                                                                                                                                Text="Publish and generate Zip file"
                                                                                                                                                                ToolTip="This action also stores the PDF file of individual employee. When an employee views the payslip from ESS he/she has to click on the link. This is recommended action as it reduces the load."
                                                                                                                                                                Width="200px"
                                                                                                                                                                style="margin-top: 8px;" />
                                                                                                                                                            <asp:Button
                                                                                                                                                                ID="btnPublishedPDF"
                                                                                                                                                                runat="server"
                                                                                                                                                                Visible="false"
                                                                                                                                                                CssClass="btnGreen"
                                                                                                                                                                ToolTip="Download Already Published Pay Slips"
                                                                                                                                                                Text="Download Already Published"
                                                                                                                                                                Width="200px"
                                                                                                                                                                style="margin-top: 8px;"
                                                                                                                                                                CausesValidation="False"
                                                                                                                                                                OnClientClick="return Check4PublishedPayslip();">
                                                                                                                                                            </asp:Button>

                                                                                                                                                            <asp:Button
                                                                                                                                                                ID="BtnSendCCBCC"
                                                                                                                                                                runat="server"
                                                                                                                                                                CausesValidation="False"
                                                                                                                                                                CssClass="custombtn"
                                                                                                                                                                Text="Publish and email to CC / BCC"
                                                                                                                                                                Width="210px"
                                                                                                                                                                style="margin-top: 8px;"
                                                                                                                                                                OnClientClick="javascript:return empchecked('BtnSendCCBCC',this);" />

                                                                                                                                                            <asp:Button
                                                                                                                                                                ID="BtnPublishGrpBy"
                                                                                                                                                                runat="server"
                                                                                                                                                                CausesValidation="False"
                                                                                                                                                                CssClass="custombtn"
                                                                                                                                                                OnClientClick="javascript:return empchecked('BtnPublishGrpBy',this);"
                                                                                                                                                                Text="Unit Wise Publish & Email"
                                                                                                                                                                ToolTip="This action stores the zip file according to unit wise and send email!"
                                                                                                                                                                Width="170px"
                                                                                                                                                                style="margin-top: 8px;" />

                                                                                                                                                            <asp:Button
                                                                                                                                                                ID="BtnLog"
                                                                                                                                                                runat="server"
                                                                                                                                                                CausesValidation="False"
                                                                                                                                                                CssClass="custombtn"
                                                                                                                                                                Text="Download Log File"
                                                                                                                                                                ToolTip="Download text file which employee's already pdf has been generated on server."
                                                                                                                                                                Width="130px"
                                                                                                                                                                style="margin-top: 8px;" />

                                                                                                                                                            <asp:Button
                                                                                                                                                                ID="btnWOPWD"
                                                                                                                                                                runat="server"
                                                                                                                                                                CausesValidation="False"
                                                                                                                                                                CssClass="custombtn"
                                                                                                                                                                OnClientClick="javascript:return empchecked('btnSave',this);"
                                                                                                                                                                Text="Publish Payslip Without Password"
                                                                                                                                                                Width="200px"
                                                                                                                                                                style="margin-top: 8px;"
                                                                                                                                                                ToolTip="This button is used by the Admin to download the PDF without password. Salary slip generated will not have any impact of Email, Password and Payslip Publish Mode functionality. Note : Please do not select 'Email sent/Not' functionality while downloading PDF without password." />

                                                                                                                                                        </td>
                                                                                                                                                    </tr>
                                                                                                                                                    <tr>
                                                                                                                                                        <td
                                                                                                                                                            valign="top">
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="tdbetweenbtn">
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            valign="top">
                                                                                                                                                            <br />
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="tdbetweenBtn">
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            valign="top">
                                                                                                                                                            &nbsp;
                                                                                                                                                        </td>
                                                                                                                                                    </tr>
                                                                                                                                                    <tr>
                                                                                                                                                        <td valign="top"
                                                                                                                                                            colspan="8">
                                                                                                                                                            <asp:Label
                                                                                                                                                                ID="lblMsgSlipWOPWD"
                                                                                                                                                                runat="server"
                                                                                                                                                                CssClass="usermessage"
                                                                                                                                                                Font-Bold="True">
                                                                                                                                                            </asp:Label>
                                                                                                                                                            <asp:Label
                                                                                                                                                                ID="lblMailMsgWOPWD"
                                                                                                                                                                runat="server"
                                                                                                                                                                CssClass="usermessage"
                                                                                                                                                                Font-Bold="True">
                                                                                                                                                            </asp:Label>
                                                                                                                                                            <asp:LinkButton
                                                                                                                                                                ID="LnkPDFWOPWD"
                                                                                                                                                                runat="server"
                                                                                                                                                                Style="display: none">
                                                                                                                                                                <img src="Images/pdf.bmp"
                                                                                                                                                                    border="0"
                                                                                                                                                                    height="18">
                                                                                                                                                            </asp:LinkButton>
                                                                                                                                                            <button
                                                                                                                                                                id="download_pdf2"
                                                                                                                                                                runat="server"
                                                                                                                                                                url="#"
                                                                                                                                                                onclick="downloadPdf(event)"
                                                                                                                                                                style="display:none;border:0;cursor:pointer;"
                                                                                                                                                                class="btn_pdf">
                                                                                                                                                                <img src="Images/pdf.bmp"
                                                                                                                                                                    border="0"
                                                                                                                                                                    height="18">
                                                                                                                                                            </button>

                                                                                                                                                        </td>
                                                                                                                                                    </tr>
                                                                                                                                                    <tr>
                                                                                                                                                        <td
                                                                                                                                                            valign="top">
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="tdbetweenbtn">
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            valign="top">
                                                                                                                                                            <br />
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="tdbetweenBtn">
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            valign="top">
                                                                                                                                                            &nbsp;
                                                                                                                                                        </td>
                                                                                                                                                    </tr>
                                                                                                                                                    <tr>
                                                                                                                                                        <td valign="top"
                                                                                                                                                            colspan="8">
                                                                                                                                                            <asp:Label
                                                                                                                                                                ID="lblMsgSlip"
                                                                                                                                                                runat="server"
                                                                                                                                                                CssClass="usermessage"
                                                                                                                                                                Font-Bold="True">
                                                                                                                                                            </asp:Label>
                                                                                                                                                            <asp:Label
                                                                                                                                                                ID="lblMailMsg"
                                                                                                                                                                runat="server"
                                                                                                                                                                CssClass="usermessage"
                                                                                                                                                                Font-Bold="True">
                                                                                                                                                            </asp:Label>
                                                                                                                                                            <asp:LinkButton
                                                                                                                                                                ID="LnkPDF"
                                                                                                                                                                runat="server"
                                                                                                                                                                Style="display: none">
                                                                                                                                                                <img src="Images/pdf.bmp"
                                                                                                                                                                    border="0"
                                                                                                                                                                    height="18">
                                                                                                                                                            </asp:LinkButton>
                                                                                                                                                            <a id="download_pdf"
                                                                                                                                                                style="display:none"
                                                                                                                                                                href="#"
                                                                                                                                                                target="_blank">
                                                                                                                                                                <img src="Images/pdf.bmp"
                                                                                                                                                                    border="0"
                                                                                                                                                                    height="18">
                                                                                                                                                            </a>

                                                                                                                                                            <button
                                                                                                                                                                id="download_pdf1"
                                                                                                                                                                runat="server"
                                                                                                                                                                url="#"
                                                                                                                                                                onclick="downloadPdf(event)"
                                                                                                                                                                style="display:none;border:0;cursor:pointer;"
                                                                                                                                                                class="btn_pdf">

                                                                                                                                                                <img src="Images/pdf.bmp"
                                                                                                                                                                    border="0"
                                                                                                                                                                    height="18">
                                                                                                                                                            </button>
                                                                                                                                                        </td>
                                                                                                                                                    </tr>
                                                                                                                                                    <tr>
                                                                                                                                                        <td
                                                                                                                                                            valign="top">
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="tdbetweenbtn">
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            valign="top">
                                                                                                                                                            <br />
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            class="tdbetweenBtn">
                                                                                                                                                        </td>
                                                                                                                                                        <td
                                                                                                                                                            valign="top">
                                                                                                                                                            &nbsp;
                                                                                                                                                        </td>
                                                                                                                                                    </tr>

                                                                                                                                                    <tr>
                                                                                                                                                        <td valign="top"
                                                                                                                                                            colspan="7">
                                                                                                                                                            <div id="divProcessBarMsg"
                                                                                                                                                                style="overflow: auto; max-height: 100px;">
                                                                                                                                                                <asp:Label
                                                                                                                                                                    ID="lblProcessBarMsg"
                                                                                                                                                                    runat="server"
                                                                                                                                                                    CssClass="ErrorMessage"
                                                                                                                                                                    Font-Bold="True">
                                                                                                                                                                </asp:Label>
                                                                                                                                                            </div>
                                                                                                                                                        </td>
                                                                                                                                                    </tr>
                                                                                                                                                </table>
                                                                                                                                            </td>
                                                                                                                                        </tr>
                                                                                                                                    </table>
                                                                                                                                </td>
                                                                                                                            </tr>
                                                                                                                            <tr id="Tr4"
                                                                                                                                runat="server">
                                                                                                                                <td
                                                                                                                                    colspan="7">
                                                                                                                                    <asp:LinkButton
                                                                                                                                        ID="LinkButton3"
                                                                                                                                        runat="server">
                                                                                                                                    </asp:LinkButton>
                                                                                                                                    <input
                                                                                                                                        id="HidEmpCode"
                                                                                                                                        type="hidden"
                                                                                                                                        runat="server"
                                                                                                                                        name="HidEmpCode" />
                                                                                                                                    <input
                                                                                                                                        id="hidstring1"
                                                                                                                                        type="hidden"
                                                                                                                                        runat="server" />
                                                                                                                                    <input
                                                                                                                                        id="HidEmpCodeName"
                                                                                                                                        type="hidden"
                                                                                                                                        runat="server"
                                                                                                                                        name="HidEmpCodeName" />
                                                                                                                                    <input
                                                                                                                                        id="HidEmpMailCnt"
                                                                                                                                        type="hidden"
                                                                                                                                        runat="server"
                                                                                                                                        name="HidEmpMailCnt" />
                                                                                                                                    <input
                                                                                                                                        id="HidMailFaildEmpCode"
                                                                                                                                        type="hidden"
                                                                                                                                        runat="server"
                                                                                                                                        name="HidMailFaildEmpCode" />
                                                                                                                                    <input
                                                                                                                                        id="hidpaycode"
                                                                                                                                        type="hidden"
                                                                                                                                        name="hidpaycode"
                                                                                                                                        runat="server" />
                                                                                                                                    <input
                                                                                                                                        id="hiddls_selcount"
                                                                                                                                        type="hidden"
                                                                                                                                        name="hiddls_selcount"
                                                                                                                                        runat="server" />
                                                                                                                                    <input
                                                                                                                                        id="Hidden2"
                                                                                                                                        type="hidden"
                                                                                                                                        name="HidEmpcode"
                                                                                                                                        runat="server" />
                                                                                                                                    <input
                                                                                                                                        id="Hidden3"
                                                                                                                                        type="hidden"
                                                                                                                                        name="HidPreVal"
                                                                                                                                        runat="server" />
                                                                                                                                    <input
                                                                                                                                        id="Hidden4"
                                                                                                                                        type="hidden"
                                                                                                                                        name="HidPreVal"
                                                                                                                                        runat="server" />
                                                                                                                                    <input
                                                                                                                                        id="Hidden5"
                                                                                                                                        type="hidden"
                                                                                                                                        name="HidPreVal"
                                                                                                                                        runat="server" />
                                                                                                                                    <input
                                                                                                                                        id="hidmonthyear"
                                                                                                                                        type="hidden"
                                                                                                                                        name="HidPreVal"
                                                                                                                                        runat="server" />
                                                                                                                                    <input
                                                                                                                                        id="HidEmailCCBCC"
                                                                                                                                        type="hidden"
                                                                                                                                        name="HidPreVal"
                                                                                                                                        runat="server" />
                                                                                                                                    <input
                                                                                                                                        id="HidEmpPdf"
                                                                                                                                        type="hidden"
                                                                                                                                        runat="server"
                                                                                                                                        name="HidEmpPdf" />
                                                                                                                                    <input
                                                                                                                                        id="HidPdfName"
                                                                                                                                        type="hidden"
                                                                                                                                        runat="server"
                                                                                                                                        name="HidPdfName" />
                                                                                                                                    <input
                                                                                                                                        id="HidCocManCheck"
                                                                                                                                        runat="server"
                                                                                                                                        name="HidCocManCheck"
                                                                                                                                        type="hidden" />
                                                                                                                                    <input
                                                                                                                                        id="Hidforecast"
                                                                                                                                        type="hidden"
                                                                                                                                        name="Hidforecast"
                                                                                                                                        runat="server" />
                                                                                                                                    <asp:Literal
                                                                                                                                        ID="litJava"
                                                                                                                                        runat="server">
                                                                                                                                    </asp:Literal>
                                                                                                                                    <asp:Literal
                                                                                                                                        ID="ltrJavaS"
                                                                                                                                        runat="server">
                                                                                                                                    </asp:Literal>
                                                                                                                                    <asp:Literal
                                                                                                                                        ID="ltrjs"
                                                                                                                                        runat="server">
                                                                                                                                    </asp:Literal>
                                                                                                                                    <input
                                                                                                                                        type="hidden"
                                                                                                                                        runat="server"
                                                                                                                                        id="HidAppPath" />
                                                                                                                                    <input
                                                                                                                                        type="hidden"
                                                                                                                                        runat="server"
                                                                                                                                        id="HidPath" />
                                                                                                                                    <input
                                                                                                                                        type="hidden"
                                                                                                                                        runat="server"
                                                                                                                                        id="process_status_id" />
                                                                                                                                    <input
                                                                                                                                        type="hidden"
                                                                                                                                        runat="server"
                                                                                                                                        id="is_gcs_powered" />
                                                                                                                                    <input
                                                                                                                                        type="hidden"
                                                                                                                                        runat="server"
                                                                                                                                        name="companyCode"
                                                                                                                                        value=""
                                                                                                                                        id="companyCode" />
                                                                                                                                    <input
                                                                                                                                        type="hidden"
                                                                                                                                        runat="server"
                                                                                                                                        name="url"
                                                                                                                                        value=""
                                                                                                                                        id="java_url" />
                                                                                                                                </td>
                                                                                                                            </tr>
                                                                                                                        </table>
                                                                                                                    </td>
                                                                                                                </tr>
                                                                                                            </table>
                                                                                            </td>
                                                                                        </tr>
                                                                                    </table>
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                        </ContentTemplate>
                                                                        <Triggers>
                                                                            <asp:AsyncPostBackTrigger
                                                                                ControlID="BtnSend" EventName="Click" />
                                                                            <asp:AsyncPostBackTrigger
                                                                                ControlID="btnSave" EventName="Click" />
                                                                            <asp:PostBackTrigger ControlID="LnkPDF" />
                                                                            <asp:PostBackTrigger
                                                                                ControlID="btnPublishedPDF" />
                                                                            <asp:AsyncPostBackTrigger
                                                                                ControlID="BtnSendCCBCC"
                                                                                EventName="Click" />
                                                                            <asp:PostBackTrigger ControlID="BtnLog" />
                                                                            <asp:PostBackTrigger
                                                                                ControlID="LnkPDFWOPWD" />
                                                                        </Triggers>
                                                                        </asp:UpdatePanel>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                        <%--< /ContentTemplate>
                                                            <Triggers>
                                                                <asp:AsyncPostBackTrigger ControlID="BtnSend" />
                                                                <asp:AsyncPostBackTrigger ControlID="btnSave"
                                                                    EventName="Click" />
                                                                <asp:PostBackTrigger ControlID="LnkPDF" />
                                                                <asp:PostBackTrigger ControlID="BtnSendCCBCC" />
                                                            </Triggers>
                                                            </asp:UpdatePanel>--%>
                                                            </div>
                                                            <input id="hiddenformail" runat="server" type="hidden" />
                                                            <input id="hidSaveMail" runat="server" type="hidden" />
                                                            <input type="hidden" id="hdnBatchId" name="BatchId"
                                                                runat="server" value="" />
                                                            <input type="hidden" id="hdnFileFormat" name="FileFormat"
                                                                runat="server" value="EXCEL" />
                                                            <input type="hidden" id="hdnAlreadyRunRptId"
                                                                name="hdnAlreadyRunRptId" runat="server" value="" />
                                                            <input type="hidden" id="hdnAlreadyRunRptName"
                                                                name="hdnAlreadyRunRptName" runat="server" value="" />
                                                            <input type="hidden" id="hdnPGP" name="hdnPGP"
                                                                runat="server" value="" />
                                                            <%--<input type="hidden" id="hdnProcessType"
                                                                name="ProcessType" runat="server" value="" />--%>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                                </fieldset>
                                <!--<input type="image" src="ImagesCA/img_main_co.jpg">-->
                                </td>
                                </tr>
                                </table>
                                <!--This is Internal Body Table-->
                                </td>
                                </tr>
                                <tr>
                                    <td width="100%">
                                        <img height="37" src="ImagesCA/table_bottom.jpg" width="100%" />
                                    </td>
                                </tr>
                                </table>
                                </td>
                                </tr>
                                </table>
                            </form>


                            <script>
                                // set higher limit for emp codes
                                setTimeout(() => {
                                    let ele = document.getElementById("USearch_txtEmpCode");
                                    let is_gcs_powered = document.getElementById("is_gcs_powered");
                                    if (ele && is_gcs_powered) {
                                        if (is_gcs_powered.value === '1') {
                                            ele.setAttribute("size", "100000")
                                        }
                                    }

                                    let lblMsgSlip = document.getElementById("lblMsgSlip");
                                    let lblMailMsgWOPWD = document.getElementById("lblMailMsgWOPWD");
                                    let process_id = document.getElementById("process_status_id")
                                    let download_pdf1 = document.getElementById("download_pdf1")
                                    let download_pdf2 = document.getElementById("download_pdf2")

                                    if (lblMsgSlip && lblMsgSlip.innerText.length > 2 && process_id && process_id.value.length > 0) {
                                        if (download_pdf1) {
                                            download_pdf1.style = "border:0;cursor:pointer;"
                                        }
                                    } else {
                                        if (download_pdf1) {
                                            download_pdf1.style = "display:none;border:0;cursor:pointer;"
                                        }
                                    }


                                    if (lblMailMsgWOPWD && lblMailMsgWOPWD.innerText.length > 2 && process_id && process_id.value.length > 0) {
                                        if (download_pdf2) {
                                            download_pdf2.style = "border:0;cursor:pointer;"
                                        }
                                    } else {
                                        if (download_pdf2) {
                                            download_pdf2.style = "display:none;border:0;cursor:pointer;"
                                        }
                                    }

                                    if (download_pdf1 && download_pdf2) {
                                        if (download_pdf1.style.display !== 'none' && download_pdf2.style.display !== 'none') {
                                            download_pdf1.style.display = 'none'
                                        }
                                    }



                                }, 100)

                            </script>
                </body>

                </html>