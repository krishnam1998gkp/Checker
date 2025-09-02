'==============================================================================================================================================================
'CREATED BY            : PRAVEEN VERMA
'STARTED ON DATE       : 15 JAN 2011
'COMPLETION DATE       : 
'PURPOSE OF FORM       : TO GENERATE SALARY SLIP AND REGISTER REPORT.
'==============================================================================================================================================================
'FOR BUG FIX
'SL. NO. ===== NAME ======= DATE-TIME ============== PURPOSE ==================================================================================================
'99.    Rohtas Singh        12 May 2020     Add code for download "Leave Without Pay Sal Slip" from a folder, fix issue total two time dispaly in salary register dynamic.
'100.   Rohtas Singh        01 Jun 2020     Add if condition for resolved runtime error on "Salary Reg. in Excel (Form R)"
'101.   Rohtas Singh        19 Jul 2020     Add new report (Salary Register(PDF)) 
'102.   Rohtas Singh        20 Aug 2020     Add new Salary SLip PDF & Employee mail.
'103.	Ritu Malik          27 aug 2020     regarding add the two more option tocreate the password for payslip without leave details
'104.   Rajesh Yadav        09 Sep 2020     Add logic for Status bar on TAX Slip
'105.   Sushil              08 Oct 2020     Salary Slip with Tax and Miscellaneous Details REPID 49 - download issue fix 8 oct 2020
'106.   Quadir Nawaj        15 Oct 2020     Added code for Payslip Publish Mode (Radiobutton): Incremental & Overwrite Mode
'107.   Ritu Malik          03 nov 2020     regarding add the two more option tocreate the password for payslip with tax details
'108.   Quadir Nawaj        27 Nov 2020     Code Added for changing logic Payslip Publish Mode on Search for generating selected (From Datagrid) PDFs Time: [04:00 PM Evening]
'109.   Rohtas Singh        04 Mar 2021     Code add for a Salary Slip code - 65
'110.   Rohtas Singh        06 Apr 2021     Code add for a Salary Slip code - 66
'111.	Ritu Malik          5 may 2021     regarding add the two more option tocreate the password for payslip without leave details with (ddMM) format,13 may 21 - this version merge in 14 May 21 build release
'112. 	Ritu Share this form in build release 16 jun 2012 with new changes of multinlingual slips
'113.   Rohtas Singh        11 Jun 2021     Code add for a Salary Slip code - 68
'114.   Ritu Malik           8 jul 2021    Regarding  tax forcast format1
'115.  	Rohtas Singh        13 Jul 2021    Resolved an TAX slip overwrite issue
'116.   Rohtas Singh        19 Jul 2021    Correct searching logic on Hold employee and correct msg on increment option on TAX slip
'117   Ritu Malik            20 jul 2021   Regarding tax forecast ABFRL changes
'118.  Ritu Malik           10 jan 2022     Regarding REFNF Report
'119.  Ritu Malik            9 mar 2022    Regarding payslip without password
''120  Ritu Malik            24 jun 2022  regarding tata payslip
'121   Deepankar             2 mar 2023    To add code for attra payslip rep id = 76,77
'122   Huzefa               3 APR 2023    Zero Value Handle In Excel
'123   Pankaj Sachan        21 Jan 24     To imporve dyncamic register downloading 
'124.  Ritu MAlik           23 feb 2024   regarding apply progress bar  on payslip without leave details
'125.  Ritu Malik           29 mar 2024    regarding apply process bar on tax forecast
'126.  Debargha             17 May 2024    Added 'Please Wait' clickbait for Salary Register in Excel[Dynamic](ddlpayslipTyp val : 38) & Added CommonDownload.js script in aspx
'127.  Debargha             28 Oct 2024   Added Report API configuration for Salary Register
'128.  Vishal               28 Dec 2024   TDS estimation progress bar added
'129.  Vishal               03 Feb 2025   TDS estimation progress bar impact issues on other payslips progress bar resolved.
'130.  Vishal               02 Feb 2025   Process bar added for Salary Register & Dynamic Register in Excel.
'131.  Vishal               04 Apr 2025   Payslip Without password progress bar not shown fix added.
'132.  Kangkan Lahkar       18 May 2025   Pass file path to download from Cloud Bucket.
'133.  Kangkan Lahkar       16 Jun 2025   Download by post method from Java service
'134.  Vishal Chauhan       18 Jun 2025   To remove inline query dependency.
'135.  Kangkan Lahkar       10 July 2025  Call Java service only after enabled month
'136.  Kangkan Lahkar       20 July 2025  Slip with leave details added, progress bar added
'137.  Ritu Malik           30 jul 2025   regarding Dynamic register CSV
'138.  Vishal Chauhan       13 Aug 2025   Report API tag changed as "HeavyExcel" instead of "ApiRptExcel"
'139.  Vishal Chauhan       19 Aug 2025   Report Type onchange with E-Mail Payslip,TDS-Estimation Slip issue fixed
'==============================================================================================================================================================
Imports System
Imports System.Activities.Expressions
Imports System.Activities.Statements
Imports System.CodeDom.Compiler
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Reflection.Emit
Imports System.Runtime.Remoting.Metadata.W3cXsd2001
Imports System.Security.Cryptography
Imports System.Text
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls.Expressions
Imports System.Xml
Imports AltPayroll.Application.Services.PGP
Imports evointernal
Imports evointernal.dz
Imports Ionic.Zip
Imports itextsharp.text.pdf
Imports Microsoft.SqlServer.Management.Smo
Imports Newtonsoft.Json
Imports Org.BouncyCastle.Asn1.Utilities
Imports Org.BouncyCastle.Crypto.Modes
Imports WebPayUILink


Namespace Payroll
    Partial Class RptMultipleSalarySlip
        Inherits System.Web.UI.Page
        Private _ClsYTD As New WebPayUILink.YTDSalSlip, ClsTaxDetails As New WebPayUILink.ClsTaxDetails _
       , ClsConfInvest As New WebPayUILink.ClsConfInvestDeclaration, ClsNewTdsEstimationSlip As New WebPayUILink.ClsNewTdsEstimationSlip _
       , SalarySlipsWithLeave As New WebPayUILink.SalarySlipsWithLeave, _ClsPTC As New WebPayUILink.SalarySlipsPTC, ClsTaxDetails_Bangladesh As New WebPayUILink.ClsTaxDetails_Bangladesh _
       , _ExcelFilebyxml As New WriteExcelFileByXML, SalSlipsWithoutLeave As New WebPayUILink.SalarySlips, ClsAppraisal As New ClsAppraisal
        Dim arrCode() As String, PK_emp_code As String = "",
        filesToInclude As New System.Collections.Generic.List(Of [String])(),
        myCollection As New ArrayList, _msg As New List(Of PayrollUtility.UserMessage), oWrite As System.IO.StreamWriter, IsProcessBarStated As Boolean = True
        Private gcs_service_obj As New WebPayUILink.GcsService
#Region "Developer Generated code"
        Protected _objCommon As New PayrollUtility.common
        Private _ObjData As New PayrollUtility.Utilities, _objcommonExp As New PayrollUtility.ExceptionManager
        '', objreport As New  ClsReportNew
#End Region
        Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
            'to check the session

            _objCommon.sessionCheck(form1)
            Page.MaintainScrollPositionOnPostBack = True
            ScriptManager.RegisterStartupScript(Me.Page, GetType(Page), String.Format("StartupScript"), "<script type=""text/javascript""> GetAttributeName() </script>", False)
            ScriptManager.RegisterStartupScript(Page, GetType(Page), "Key", "<script>MakeStaticHeader('DgPayslip', '300', '1100' , 25 ,'false','Y'); </script>", False)
            ScriptManager.RegisterStartupScript(UpdatePanel3, UpdatePanel3.GetType(), Guid.NewGuid().ToString.Trim(), "CloseDialog();", True)
            ScriptManager.RegisterStartupScript(UpdatePanel1, UpdatePanel1.GetType(), Guid.NewGuid().ToString.Trim(), "CloseDialog();", True)
            ScriptManager.RegisterStartupScript(updMail, updMail.GetType(), Guid.NewGuid().ToString.Trim(), "CloseDialog();", True)



            companyCode.Value = Session("compCode")
            Dim java_url_loc As String = ""
            If Not System.Configuration.ConfigurationManager.AppSettings("JavaServiceDomain") Is Nothing Then
                If System.Configuration.ConfigurationManager.AppSettings("JavaServiceDomain").ToString <> "" Then
                    java_url_loc = System.Configuration.ConfigurationManager.AppSettings("JavaServiceDomain").ToString
                End If
            End If
            java_url.Value = java_url_loc & "/api/payslip/download-zip"


            'ScriptManager.RegisterStartupScript(UpdatePanel2, UpdatePanel2.GetType(), Guid.NewGuid().ToString.Trim(), "CloseDialog();", True
            If Not IsPostBack Then
                hdnusername.Value = Session("UId").ToString
                _ObjData.Bindddl(DDLPaySlipType, "PaySP_PaySlipConfigure_Sel", "ReportDesc", "RepID")
                _objCommon.PopulateDDL_SalaryProcMonthYr(ddlMonthYear, Session("Sfindate").ToString, Session("Efindate").ToString)
                hidstatus.Value = "0"
                tblsection.Style.Value = "display:none"
                'rbtsalaryEmail.SelectedValue = 0
                rbtnslip.Checked = True
                javascript()
                PopulateddlRepIn()
                populateDdlreportType()
                populatestrddl(ddlshortbasis)
                populatestrddl(ddlGroup2)
                populatestrddl(ddlGroup3)
                PopulateReportFormat()
                populate_multilingDDl()
                If Not IsNothing(Request.QueryString("id")) Then
                    getQuerryStringValue()
                End If
                TblReimb.Style.Value = "display:none"
                SlipRegPre.Style.Value = "display:"
                divgenerate.Style.Value = "display:"
                divEmail.Style.Value = "display:none"
                Btndelete.Style.Value = "display:none"
                litJava.Text = ""
                javascript()
                trvisualtrueandfalse()
                COC_MaindatoryCheck()
                CheckExcelProcessbarAlreadyProcessing()
                CheckProcessLocked()
                ShowPGP()
            End If
            If DdlreportType.SelectedValue = "R" Then
                'TdLog.Style.Value = "display:"
                BtnLog.Visible = True
            Else
                'TdLog.Style.Value = "display:none"
                BtnLog.Visible = False
            End If
            If rbtnmail.Checked Then
                divEmail.Style.Value = "display:"
                SlipRegPre.Style.Value = "display:none"
            End If
            tblhelp.Style.Value = "display:none"
            litJava.Text = ""
            lit.Text = ""
        End Sub
        'Added by praveen on 17 Aug 2011 to populate ddlreporttype.
        Private Sub populateDdlreportType()
            Try
                Dim dt As New DataTable
                dt = _ObjData.GetDataTableProc("paysp_bind_DDLPaySlipType")
                DdlreportType.DataValueField = "Repid"
                DdlreportType.DataTextField = "ReportDesc"
                DdlreportType.DataSource = dt
                DdlreportType.DataBind()
            Catch ex As Exception
                _objcommonExp.PublishError("populateDDLPaySlipType", ex)
            End Try
        End Sub
        Private Sub trvisualtrueandfalse()
            'to display the paycode table if selected report type is salary register
            lblmsg2.Text = ""
            If DDLPaySlipType.SelectedValue = "9" Or DDLPaySlipType.SelectedValue = "13" Or DDLPaySlipType.SelectedValue = "15" Or DDLPaySlipType.SelectedValue = "16" Or DDLPaySlipType.SelectedValue = "17" Or DDLPaySlipType.SelectedValue = "27" Or DDLPaySlipType.SelectedValue = "19" Or DDLPaySlipType.SelectedValue = "47" Or DDLPaySlipType.SelectedValue = "48" Then
                tblpaycode.Style.Value = "display:"
                trsortbasis.Style.Value = "display:"
                If DDLPaySlipType.SelectedValue = "47" Then
                    trGroupBY.Style.Value = "display:none"
                Else
                    trGroupBY.Style.Value = "display:"
                End If
                TRDIV.Style.Value = "display:"
            Else
                tblpaycode.Style.Value = "display:none"
                trsortbasis.Style.Value = "display:none"
                trGroupBY.Style.Value = "display:none"
                TRDIV.Style.Value = "display:none"
            End If

            If DDLPaySlipType.SelectedValue = "6" Then
                trrepformat.Style.Value = "display:"
            Else
                trrepformat.Style.Value = "display:none"
            End If
            If DDLPaySlipType.SelectedValue = "0" Or DDLPaySlipType.SelectedValue = "12" Or DDLPaySlipType.SelectedValue = "10" Or DDLPaySlipType.SelectedValue = "5" Or DDLPaySlipType.SelectedValue = "7" Or DDLPaySlipType.SelectedValue = "18" Or DDLPaySlipType.SelectedValue = "52" Or DDLPaySlipType.SelectedValue = "51" Or DDLPaySlipType.SelectedValue = "49" Or DDLPaySlipType.SelectedValue = "20" Or DDLPaySlipType.SelectedValue = "24" Or DDLPaySlipType.SelectedValue = "26" Then
                Dim Flag As String = ""
                If DDLPaySlipType.SelectedValue = "18" Or DDLPaySlipType.SelectedValue = "52" Or DDLPaySlipType.SelectedValue = "51" Or DDLPaySlipType.SelectedValue = "49" Then
                    Otherpaycode()
                    Flag = "A"
                Else
                    tblothepaycode.Style.Value = "display:none"
                    Flag = "B"
                End If
            Else
                tblothepaycode.Style.Value = "display:none"
            End If

            If DDLPaySlipType.SelectedValue = "8" Or DDLPaySlipType.SelectedValue = "9" Or DDLPaySlipType.SelectedValue = "13" Or DDLPaySlipType.SelectedValue = "15" Or DDLPaySlipType.SelectedValue = "16" Or DDLPaySlipType.SelectedValue = "17" Or DDLPaySlipType.SelectedValue = "27" Or DDLPaySlipType.SelectedValue = "19" Or DDLPaySlipType.SelectedValue = "47" Or DDLPaySlipType.SelectedValue = "48" Or DDLPaySlipType.SelectedValue = "22" Or DDLPaySlipType.SelectedValue = "29" Or DDLPaySlipType.SelectedValue = "30" Or DDLPaySlipType.SelectedValue = "31" Then
                tblpaycode.Style.Value = "display:"
            End If

            If DDLPaySlipType.SelectedValue = "8" Then
                tblpaycode.Style.Value = "display:"
            End If

            If DDLPaySlipType.SelectedValue = "25" Then
                TblReimb.Style.Value = "display:"
            Else
                TblReimb.Style.Value = "display:none"
            End If
            If DDLPaySlipType.SelectedValue = "31" Then
                lblmsg.Text = "This is a client specific report. In this report value of processed arrear and processed loan does not get publish, only selected paycode are published"
            End If
            If DdlreportType.SelectedValue <> "" Then
                TblReimb.Style.Value = "display:none"
                tblothepaycode.Style.Value = "display:none"
                tblpaycode.Style.Value = "display:none"
                trsortbasis.Style.Value = "display:none"
                trGroupBY.Style.Value = "display:none"
                TRDIV.Style.Value = "display:none"
            End If
            If rbtnslip.Checked = True Then
                divEmail.Style.Value = "display:none"
            End If
            'Added by Rajesh for Luxor Register on 18 oct 13
            If DDLPaySlipType.SelectedValue = "46" Then
                trsortbasis.Style.Value = "display:"
                trGroupBY.Style.Value = "display:"
                TRDIV.Style.Value = "display:"
                ddlshortbasis.Items.Clear()
                populatestrddl1(ddlshortbasis)
                ddlshortbasis.Enabled = False
            Else
                populatestrddl(ddlshortbasis)
                ddlshortbasis.Enabled = True
            End If
            'Added by Rohtas Singh on 06 Dec 2017
            If DDLPaySlipType.SelectedValue = "54" Then
                trRptformat.Style.Value = "display:"
                lblmsg2.Text = "This is a custom made report."
            Else
                trRptformat.Style.Value = "display:none"
            End If
            'Added by Anubha jain on 11 nov 2019
            If DDLPaySlipType.SelectedValue.ToString = "38" Then
                trShowClr.Style.Value = "display:"
                'rbtshowclr.SelectedValue = "Y"
                trFileName.Style.Value = "display:none"
                btnExport2CSV.Visible = False
            Else
                trShowClr.Style.Value = "display:none"
                trformat.Style.Value = "display:none"
                trEncrType.Style.Value = "display:none"
                trSftpID.Style.Value = "display:none"
                trFileName.Style.Value = "display:none"
                btnExport2CSV.Visible = False
            End If
        End Sub
        Private Sub Otherpaycode()
            tblothepaycode.Style.Value = "display:"
            Dim _dt As New DataTable, i As Integer = 0
            _dt = _ObjData.GetDataTableProc("Paysp_sel_OtherType_salreg")
            If Hidothepaycode.Value = "" Then
                If _dt.Rows.Count > 0 Then
                    Chklistothepaycode.DataSource = _dt
                    Chklistothepaycode.DataTextField = "Short_Desc"
                    Chklistothepaycode.DataValueField = "pk_pay_code"
                    Chklistothepaycode.DataBind()
                    Hidothepaycode.Value = _dt.Rows.Count
                End If
            End If
        End Sub
        Private Sub getQuerryStringValue()
            lblmsg2.Text = ""
            If Not IsNothing(Request.QueryString("id").ToString) Then
                Try
                    Dim dt As New DataTable, arrparam(0) As SqlClient.SqlParameter _
                    , code As String = Request.QueryString("id").ToString
                    arrCode = Split(code, "~")
                    DDLPaySlipType.SelectedValue = arrCode(0).ToString
                    trview.Style.Value = "display:"
                    arrparam(0) = New SqlClient.SqlParameter("@id", arrCode(0).ToString)
                    dt = _ObjData.GetDataTableProc("paysp_reportdetails", arrparam)
                    If dt.Rows.Count > 0 Then
                        ViewState("SalaryData") = Nothing
                        SalaryDataNew(dt)
                    End If
                    If DDLPaySlipType.SelectedValue = "5" Then
                        tblsection.Style.Value = "display:"
                        SectionSel()
                    Else
                        tblsection.Style.Value = "display:none"
                    End If
                    'here we checl the report type and according to the report type we show/Hide the group2 and group3 selection
                    If DDLPaySlipType.SelectedValue = "16" Then
                        populatestrddl(ddlGroup2)
                        populatestrddl(ddlGroup3)
                        trsortbasis2.Style.Value = "display:"
                        trsortbasis3.Style.Value = "display:"
                    Else
                        populatestrddl(ddlGroup2)
                        populatestrddl(ddlGroup3)
                        trsortbasis2.Style.Value = "display:none"
                        trsortbasis3.Style.Value = "display:none"
                    End If

                    'to populate the paycodes dropdownlist if report type is salary register
                    If DDLPaySlipType.SelectedValue = "8" Or DDLPaySlipType.SelectedValue = "9" Or DDLPaySlipType.SelectedValue = "13" Or DDLPaySlipType.SelectedValue = "15" Or DDLPaySlipType.SelectedValue = "16" Or DDLPaySlipType.SelectedValue = "17" Or DDLPaySlipType.SelectedValue = "27" Or DDLPaySlipType.SelectedValue = "19" Or DDLPaySlipType.SelectedValue = "47" Or DDLPaySlipType.SelectedValue = "48" Or DDLPaySlipType.SelectedValue = "22" Or DDLPaySlipType.SelectedValue = "29" Or DDLPaySlipType.SelectedValue = "30" Or DDLPaySlipType.SelectedValue = "31" Or DDLPaySlipType.SelectedValue = "46" Then
                        tblpaycode.Style.Value = "display:"
                        If DDLPaySlipType.SelectedValue = "22" Then
                            tblpaycode.Style.Value = "display:"
                        Else
                        End If
                        Bind_Paycode_Check()
                    Else
                        tblpaycode.Style.Value = "display:none"
                    End If
                    EnableUSearchDdl()
                    If DDLPaySlipType.SelectedValue = "25" Then
                        TblReimb.Style.Value = "display:"
                        Bind_Paycode_Check()
                    Else
                        TblReimb.Style.Value = "display:none"
                    End If
                    If DDLPaySlipType.SelectedValue = "31" Then
                        lblmsg2.Text = "This is a client specific report. In this report value of processed arrear and processed loan does not get publish, only selected paycode are published"
                    End If
                    'Added by Rohtas Singh on 06 Dec 2017
                    If DDLPaySlipType.SelectedValue = "54" Then
                        lblmsg2.Text = "This is a custom made report."
                    End If
                    trvisualtrueandfalse()
                Catch ex As Exception
                    _objcommonExp.PublishError("Salarydata())", ex)
                End Try
            End If
        End Sub

        Private Sub SectionSel()
            Try
                Dim dt As DataTable = Nothing
                dt = _ObjData.GetDataTableProc("sel_sections")
                If dt.Rows.Count > 0 Then
                    CBLsection.DataSource = dt
                    CBLsection.DataTextField = "sectionnumber"
                    CBLsection.DataValueField = "sectionnumber"
                    CBLsection.DataBind()
                    CBLsection.RepeatDirection() = RepeatDirection.Horizontal
                End If
            Catch ex As Exception
                _objcommonExp.PublishError("Error in SectionSel()", ex)
            End Try
        End Sub
        Private Sub PopulateddlRepIn()
            ddlRepIn.Items.Add(New ListItem("HTML", "H"))
            ddlRepIn.Items.Add(New ListItem("PDF", "P"))
            ddlRepIn.Items.Add(New ListItem("PaySlip Link", "L"))
        End Sub
        Private Sub javascript()
            'Btnsearch.Attributes.Add("onclick", "javascript:return ValidateReportType();")
            BtnSend.Attributes.Add("onclick", "javascript:return empchecked('BtnSend');")
        End Sub
        '---------Added By Geeta on 30 Aug 2012
        Private Sub ExportToExcelXML_GrandTotal(ByVal source As DataSet, ByRef _sw As StreamWriter)
            Dim _RowCount As Integer = 0
            Try
                Dim _ExcelXML As New StringBuilder()
                _sw.Write("<?xml version=""1.0""?><?mso-application progid=""Excel.Sheet""?>" & Chr(13) & "" & Chr(10) & "<Workbook ")
                _sw.Write("xmlns=""urn:schemas-microsoft-com:office:spreadsheet""" & Chr(13) & "" & Chr(10) & "")
                _sw.Write(" xmlns:o=""urn:schemas-microsoft-com:office:office""" & Chr(13) & "" & Chr(10) & " ")
                _sw.Write(" xmlns:x=""urn:schemas-microsoft-com:office:excel""" & Chr(13) & "" & Chr(10))
                _sw.Write(" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">" & Chr(13) & "" & Chr(10) & " <Styles>" & Chr(13) & "" & Chr(10) & " ")
                _sw.Write("<Style ss:ID=""Default"" ss:Name=""Normal"">" & Chr(13) & "" & Chr(10) & " ")
                '_sw.Write("<Alignment ss:Horizontal=\"Right\" ss:Vertical=\"Bottom\"/>\r\n <Borders/>"); 
                _sw.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>" & Chr(13) & "" & Chr(10) & " <Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders>")
                _sw.Write("" & Chr(13) & "" & Chr(10) & " <Font ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9""/>" & Chr(13) & "" & Chr(10) & " <Interior/>" & Chr(13) & "" & Chr(10) & " <NumberFormat/>")
                _sw.Write("" & Chr(13) & "" & Chr(10) & " <Protection/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")

                'Geta
                'Cumstomise format 
                _sw.Write("<Style ss:ID=""hd"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _sw.Write("ss:FontName=""Verdana"" ss:Bold=""1"" ss:Color=""#f2f2f2""  x:Family=""Swiss"" ss:Size=""10"" />" & Chr(13) & "" & Chr(10) & " <Interior ss:Color=""#4a452a"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & " ")

                _sw.Write("<Style ss:ID=""Cd"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _sw.Write("ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & " <Interior ss:Color=""#b6dde8"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & " ")

                _sw.Write("<Style ss:ID=""sh"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _sw.Write("ss:FontName=""Verdana"" ss:Color=""#ffffff"" x:Family=""Swiss"" ss:Size=""9"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & " <Interior ss:Color=""#337140"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & " ")

                _sw.Write("<Style ss:ID=""sr"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Left"" ss:Vertical=""Center""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _sw.Write("ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & " <Interior ss:Color=""#ffcc00"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & " ")

                _sw.Write("<Style ss:ID=""GT"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Right"" ss:Vertical=""Center""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _sw.Write("ss:FontName=""Verdana"" ss:Color=""#f2f2f2"" x:Family=""Swiss"" ss:Size=""11"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & " <Interior ss:Color=""#0f253f"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & " ")

                _sw.Write("<Style ss:ID=""Ed"">" & Chr(13) & "" & Chr(10) & "")
                _sw.Write("<Alignment ss:Horizontal=""Right"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""0""/>" & Chr(13) & "" & Chr(10) & "")
                _sw.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000"" ss:Bold=""0""/>" & Chr(13) & "" & Chr(10) & "")
                _sw.Write("<Interior ss:Color=""#CCFFFF"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                _sw.Write("<Style ss:ID=""Ge"">" & Chr(13) & "" & Chr(10) & "")
                _sw.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""0""/>" & Chr(13) & "" & Chr(10) & "")
                _sw.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000"" ss:Bold=""0""/>" & Chr(13) & "" & Chr(10) & "")
                _sw.Write("<Interior ss:Color=""#CCFFFF"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                '_sw.Write("\r\n <Style ss:ID=\"StyleBorder\"><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>"); 
                _sw.Write("" & Chr(13) & "" & Chr(10) & " <Style ss:ID=""BC"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _sw.Write("x:Family=""Swiss"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _sw.Write("<Style ss:ID=""Sl"">" & Chr(13) & "" & Chr(10) & " <NumberFormat")
                _sw.Write(" ss:Format=""@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _sw.Write("ss:ID=""Dc"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _sw.Write("ss:Format=""0.0000""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _sw.Write("<Style ss:ID=""IT"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _sw.Write("ss:Format=""0""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _sw.Write("ss:ID=""DL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _sw.Write("ss:Format=""mm/dd/yyyy;@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _sw.Write("<Style ss:ID=""s2"">" & Chr(13) & "" & Chr(10) & "")
                _sw.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Top"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _sw.Write("<Font x:CharSet=""1"" ss:Size=""9"" ss:Color=""#808080"" ss:Underline=""Single""/>" & Chr(13) & "" & Chr(10) & "")
                _sw.Write("<Interior ss:Color=""#FFFFFF"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                _sw.Write("</Styles>" & Chr(13) & "" & Chr(10) & " ")

                Const end_ExcelXML As String = "</Workbook>"
                Dim sheetCount As Integer = 1

                Dim _parm(1) As SqlClient.SqlParameter
                Dim dtbl As DataTable
                _parm(0) = New SqlParameter("@pk_id", DDLPaySlipType.SelectedValue.ToString)
                _parm(1) = New SqlClient.SqlParameter("@ReportFlg", SqlDbType.VarChar, 100)
                _parm(1).Direction = ParameterDirection.Output
                dtbl = _ObjData.GetDataTableProc("Paysp_Mstreporttype_excelHead", _parm)


                '_sw.Write(_sw.ToString())
                '_ExcelXML.Remove(0, _ExcelXML.Length - 1)

                _sw.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                _sw.Write("<Table>")

                Dim _dt As New DataTable, _Arrparam(0) As SqlClient.SqlParameter
                _Arrparam(0) = New SqlClient.SqlParameter("@CC_Code", USearch.UCddlcostcenter.ToString)
                _dt = _ObjData.GetDataTableProc("sp_sel_companydetails_forreport", _Arrparam)

                _sw.Write("<Row>")
                _sw.Write("<Cell ss:StyleID=""sh"" ss:MergeAcross=""" + source.Tables(3).Rows(0).Item("Counter1").ToString + """><Data ss:Type=""String"">")
                _sw.Write(_dt.Rows(0)("Comp_Name").ToString)
                _sw.Write("</Data></Cell>")
                _sw.Write("</Row>")

                _sw.Write("<Row>")
                _sw.Write("<Cell ss:StyleID=""sh"" ss:MergeAcross=""" + source.Tables(3).Rows(0).Item("Counter1").ToString + """><Data ss:Type=""String"">")
                _sw.Write(_dt.Rows(0)("Add1").ToString & " " & _dt.Rows(0)("Add2").ToString)
                _sw.Write("</Data></Cell>")
                _sw.Write("</Row>")

                _sw.Write("<Row>")
                _sw.Write("<Cell ss:StyleID=""sh"" ss:MergeAcross=""" + source.Tables(3).Rows(0).Item("Counter1").ToString + """><Data ss:Type=""String"">")
                _sw.Write(_dt.Rows(0)("Add3").ToString)
                _sw.Write("</Data></Cell>")
                _sw.Write("</Row>")

                _sw.Write("<Row>")
                _sw.Write("<Cell ss:MergeAcross=""" + source.Tables(3).Rows(0).Item("Counter1").ToString + """><Data ss:Type=""String"">")
                _sw.Write("")
                _sw.Write("</Data></Cell>")
                _sw.Write("</Row>")


                If _parm(1).Value.ToString.ToUpper = "EMPLOYEE MASTER DATA" Or _parm(1).Value.ToString.ToUpper = "EMPLOYEE HISTORY" Or _parm(1).Value.ToString.ToUpper = "INVESTMENT DETAIL" Or _parm(1).Value.ToString.ToUpper = "YTD REPORT" Then
                    _sw.Write("<Row>")
                    _sw.Write("<Cell ss:StyleID=""sr"" ss:MergeAcross=""" + source.Tables(3).Rows(0).Item("Counter1").ToString + """><Data ss:Type=""String"">")
                    _sw.Write(_parm(1).Value.ToString)
                    _sw.Write("</Data></Cell>")
                    _sw.Write("</Row>")
                Else
                    _sw.Write("<Row>")
                    _sw.Write("<Cell ss:StyleID=""sr"" ss:MergeAcross=""" + source.Tables(3).Rows(0).Item("Counter1").ToString + """><Data ss:Type=""String"">")
                    _sw.Write(_parm(1).Value.ToString & " for the month of " & ddlMonthYear.SelectedItem.ToString)
                    _sw.Write("</Data></Cell>")
                    _sw.Write("</Row>")

                    _sw.Write("<Row>")
                    _sw.Write("<Cell ss:StyleID=""sr"" ss:MergeAcross=""" + source.Tables(3).Rows(0).Item("Counter1").ToString + """><Data ss:Type=""String"">")
                    _sw.Write("Generated Date : " & DateTime.Now.ToString("dd MMMM yyyy HH:mm"))
                    _sw.Write("</Data></Cell>")
                    _sw.Write("</Row>")
                End If

                _sw.Write("<Row>")
                _sw.Write("")
                _sw.Write("</Row>")

                _sw.Write("<Row>")
                For x As Integer = 0 To source.Tables(0).Columns.Count - 1
                    If source.Tables(0).Columns(x).ColumnName.ToString.ToUpper <> "FK_COSTCENTER_CODE" And source.Tables(0).Columns(x).ColumnName.ToString.ToUpper <> "FK_GRADE_CODE" And source.Tables(0).Columns(x).ColumnName.ToString.ToUpper <> "FK_DEPT" And source.Tables(0).Columns(x).ColumnName.ToString.ToUpper <> "FK_DESIG_CODE" And source.Tables(0).Columns(x).ColumnName.ToString.ToUpper <> "FK_UNIT_CODE" And source.Tables(0).Columns(x).ColumnName.ToString.ToUpper <> "FK_LOC_CODE" And source.Tables(0).Columns(x).ColumnName.ToString.ToUpper <> "FK_SUBDEPT_CODE" And source.Tables(0).Columns(x).ColumnName.ToString.ToUpper <> "FK_LEVEL_CODE" And source.Tables(0).Columns(x).ColumnName.ToString.ToUpper <> "FK_DEPT_CODE" And source.Tables(0).Columns(x).ColumnName.ToString.ToUpper <> "FK_UNIT" And source.Tables(0).Columns(x).ColumnName.ToString.ToUpper <> "FK_SUBDEPT" Then
                        _sw.Write("<Cell ss:StyleID=""hd""><Data ss:Type=""String"">")
                        _sw.Write(source.Tables(0).Columns(x).ColumnName)
                        _sw.Write("</Data></Cell>")
                    End If
                Next
                _sw.Write("</Row>")

                For Each x As DataRow In source.Tables(0).Rows
                    _RowCount += 1
                    'If _RowCount = 63000 Then
                    '    _RowCount = 0
                    '    sheetCount += 1
                    '    _sw.Write("</Table>")
                    '    _sw.Write(" </Worksheet>")
                    '    _sw.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                    '    _sw.Write("<Table>")
                    '    _sw.Write("<Row>")
                    '    For xi As Integer = 0 To source.Tables(0).Columns.Count - 1
                    '        _sw.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    '        _sw.Write(source.Tables(0).Columns(xi).ColumnName)
                    '        _sw.Write("</Data></Cell>")
                    '    Next
                    '    _sw.Write("</Row>")
                    'End If
                    _sw.Write("<Row>")
                    For y As Integer = 0 To source.Tables(0).Columns.Count - 1
                        Dim XMLstring As String = x(y).ToString()
                        XMLstring = XMLstring.Trim()
                        If source.Tables(0).Columns(y).ColumnName.ToString.ToUpper <> "FK_COSTCENTER_CODE" And source.Tables(0).Columns(y).ColumnName.ToString.ToUpper <> "FK_GRADE_CODE" And source.Tables(0).Columns(y).ColumnName.ToString.ToUpper <> "FK_DEPT" And source.Tables(0).Columns(y).ColumnName.ToString.ToUpper <> "FK_DESIG_CODE" And source.Tables(0).Columns(y).ColumnName.ToString.ToUpper <> "FK_UNIT_CODE" And source.Tables(0).Columns(y).ColumnName.ToString.ToUpper <> "FK_LOC_CODE" And source.Tables(0).Columns(y).ColumnName.ToString.ToUpper <> "FK_SUBDEPT_CODE" And source.Tables(0).Columns(y).ColumnName.ToString.ToUpper <> "FK_LEVEL_CODE" And source.Tables(0).Columns(y).ColumnName.ToString.ToUpper <> "FK_DEPT_CODE" And source.Tables(0).Columns(y).ColumnName.ToString.ToUpper <> "FK_UNIT" And source.Tables(0).Columns(y).ColumnName.ToString.ToUpper <> "FK_SUBDEPT" Then
                            'Account Number print in String added by Geta on 18 May 10
                            If IsNumeric(XMLstring) And source.Tables(0).Columns(y).ColumnName.ToString.ToUpper <> "EMP_CODE" And source.Tables(0).Columns(y).ColumnName.ToString.ToUpper <> "EMPCODE" And source.Tables(0).Columns(y).ColumnName.ToString.ToUpper <> "ACCOUNT NUMBER" And source.Tables(0).Columns(y).ColumnName.ToString.ToUpper <> "NEWEMPCODE" Then
                                _sw.Write("<Cell ss:StyleID=""Ed"">" + "<Data ss:Type=""Number"">")
                            Else
                                _sw.Write("<Cell ss:StyleID=""Ge"">" + "<Data ss:Type=""String"">")
                            End If
                            _sw.Write(XMLstring)
                            _sw.Write("</Data></Cell>")
                        End If
                    Next
                    _sw.Write("</Row>")
                Next
                _sw.Write("<Row>")
                ' Loop excute according to the no. of Columns count
                For row As Integer = 0 To CType(source.Tables(1).Rows(0).Item("Counter").ToString, Integer) - 11
                    _sw.Write("<Cell ss:StyleID=""GT"">" + "<Data ss:Type=""String"">")
                    _sw.Write("")
                    _sw.Write("</Data></Cell>")
                Next
                _sw.Write("<Cell ss:StyleID=""GT"">" + "<Data ss:Type=""String"">")
                _sw.Write("Grand Total")
                _sw.Write("</Data></Cell>")

                _sw.Write("<Cell ss:StyleID=""GT"">" + "<Data ss:Type=""String"">")
                _sw.Write(":")
                _sw.Write("</Data></Cell>")
                ' To Show Grand Total
                For row1 As Integer = 0 To source.Tables(2).Columns.Count - 1
                    _sw.Write("<Cell ss:StyleID=""GT"">" + "<Data ss:Type=""Number"">")
                    _sw.Write(source.Tables(2).Rows(0).Item(source.Tables(2).Columns(row1).ColumnName))
                    _sw.Write("</Data></Cell>")
                Next
                _sw.Write("</Row>")
                _sw.Write("</Table>")
                _sw.Write(" </Worksheet>")
                _sw.Write(end_ExcelXML)
            Catch ex As Exception
                'for catch the error
                '_objExceptionMgr.PublishError("For gebnerate the excel file (fnExportTo_ExcelXML())", ex)
            End Try
        End Sub
        Private Sub btnpreview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPreview.Click
            Dim ds As New DataSet, arrparam(16) As SqlClient.SqlParameter, stringWrite As New System.IO.StringWriter, myHTMLTextWriter As New System.Web.UI.HtmlTextWriter(stringWrite) _
            , arrcode() As String = Nothing, countemp As Integer = 0, empcode As String = "", lstitem As ListItem, paycodeSel1 As String = "", lstitem1 As ListItem _
            , paycodeSel As String = "", Dep As String = "", Desig As String = "", Grad As String = "", Level As String = "", CC As String = "", Loc As String = "" _
            , unit As String = "", SalBase As String = "", EmpFName As String = "", EmpLName As String = "", EmpType As String = "", Month As String = "", Year As String = "" _
            , Sorttype As String = "", ShortType As String = "", PayCode As String = "", reptype As String = "", PayCodeAdd As String = "", Ids As String = "" _
            , srttp As String = "", Hold As String = "", doj As String = "", PF As String = "", Leave As String = "", Absent As String = "", ESI As String = "" _
            , BankAcc As String = "", PayCodeId As String = "", Salut As String = "", _strVal As String = Guid.NewGuid.ToString, _str As New System.Text.StringBuilder _
            , str As String = "", chkitem As ListItem, Comzero As String = "", ReptType As String = "", RepFormat As String = "", chkPF_Old As String = "" _
            , chkESI_Old As String = "", Comlogo As String = "", Loan As String = "", Advance As String = "", Ent As String = "", otherinc As String = "" _
            , StaffId As String = "", NegitiveSalFlg As String = "", arrp(20) As SqlClient.SqlParameter, arrpam(18) As SqlClient.SqlParameter _
            , arrpamwork(12) As SqlClient.SqlParameter, arrpamm(18) As SqlClient.SqlParameter, dtRepType As DataTable, ReportType As String = "", Fromdate As String = ""
            hdnFileFormat.Value = "EXCEL"
            CheckExcelProcessbarAlreadyProcessing()
            If (lblProcessStatusExcel.Text <> "") Then
                _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = lblProcessStatusExcel.Text})
                _objCommon.ShowMessage(_msg)
                Exit Sub
            End If
            'CheckProcessLocked()
            'If (hdnAlreadyRunRptName.Value.Trim().Length > 1) Then
            '    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = hdnAlreadyRunRptName.Value})
            '    _objCommon.ShowMessage(_msg)
            '    Exit Sub
            'End If            
            'Added by Debargha on 21 Oct 2024
            Dim APIConfigParam(2) As SqlClient.SqlParameter, IsNewUrl As String = "N"
            Dim AppPathStr As String = HttpRuntime.AppDomainAppVirtualPath.ToString, _array() As String
            _array = Split(AppPathStr, "/")
            AppPathStr = _array(_array.Length - 1)

            hidothrpaycode.Value = ""
            hdquery.Value = ""
            'Common variable start here
            Dep = USearch.UCddldept.ToString()
            Desig = USearch.UCddldesig.ToString()
            Grad = USearch.UCddlgrade.ToString()
            Level = USearch.UCddllevel.ToString()
            CC = USearch.UCddlcostcenter.ToString()
            Loc = USearch.UCddllocation.ToString()
            unit = USearch.UCddlunit.ToString()
            SalBase = USearch.UCddlsalbasis.ToString()
            empcode = USearch.UCTextcode.ToString
            EmpFName = IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), "")
            Month = _objCommon.nNz(ddlMonthYear.SelectedValue.ToString)
            Year = Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4)
            EmpLName = IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), "")
            EmpType = USearch.UCddlEmp.ToString()
            Hold = ddlshowsal.SelectedValue.ToString
            If DDLPaySlipType.SelectedValue = "21" Then
                'Added by Debargha on 21-Oct-2024
                APIConfigParam(0) = New SqlClient.SqlParameter("@SP_Name", "PaySP_SalaryRegInExcel")
                APIConfigParam(1) = New SqlClient.SqlParameter("@ReportName", "HRD Report")
                APIConfigParam(2) = New SqlClient.SqlParameter("@IsNewURL", SqlDbType.VarChar, 1)
                APIConfigParam(2).Direction = ParameterDirection.Output
                _ObjData.ExecuteStoredProcMsg("PaySP_ReportAPI_ConfigSel", APIConfigParam)
                IsNewUrl = APIConfigParam(2).Value.ToString
                If IsNewUrl = Nothing OrElse IsNewUrl = "N" Then
                    arrparam(0) = New SqlParameter("@fk_costcenter_code", USearch.UCddlcostcenter.ToString())
                    arrparam(1) = New SqlParameter("@fk_loc_code", USearch.UCddllocation.ToString())
                    arrparam(2) = New SqlParameter("@fk_unit", USearch.UCddlunit.ToString())
                    arrparam(3) = New SqlParameter("@salaried", USearch.UCddlsalbasis.ToString())
                    arrparam(4) = New SqlParameter("@pk_emp_code", USearch.UCTextcode.ToString)
                    arrparam(5) = New SqlParameter("@fk_dept_code", USearch.UCddldept.ToString())
                    arrparam(6) = New SqlParameter("@fk_desig_code", USearch.UCddldesig.ToString())
                    arrparam(7) = New SqlParameter("@fk_grade_code", USearch.UCddlgrade.ToString())
                    arrparam(8) = New SqlParameter("@fk_level_Code", USearch.UCddllevel.ToString())
                    arrparam(9) = New SqlClient.SqlParameter("@month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
                    arrparam(10) = New SqlClient.SqlParameter("@year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                    arrparam(11) = New SqlParameter("@first_name", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
                    arrparam(12) = New SqlParameter("@last_name", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                    arrparam(13) = New SqlParameter("@Hold", ddlshowsal.SelectedValue.ToString)
                    arrparam(14) = New SqlParameter("@UserGroup", Session("Ugroup").ToString)
                    arrparam(15) = New SqlParameter("@EmpStatus", USearch.UCddlEmp.ToString)
                    arrparam(16) = New SqlParameter("@RepID", DDLPaySlipType.SelectedValue.ToString)
                    ds = _ObjData.GetDataSetProc("PaySP_SalaryRegInExcel", arrparam)
                    If ds.Tables(0).Rows.Count > 0 Then
                        'Geeta2012
                        Dim filename As String = "", _sw As StreamWriter, complexID As Guid = Guid.NewGuid()

                        filename = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\Salary_Register" & "_" & Right(complexID.ToString, 6) & ".xls"
                        If System.IO.File.Exists(filename) Then
                            System.IO.File.Delete(filename)
                        End If
                        _sw = New StreamWriter(filename)
                        ExportToExcelXML(ds, _sw)
                        _sw.Close()
                        _sw.Dispose()

                        Response.Clear()
                        Response.BufferOutput = False

                        'Dim ReadmeText As [String] = "Hello!" & vbLf & vbLf & "This is a README..." & DateTime.Now.ToString("G")
                        Response.ContentType = "application/zip"
                        Response.AddHeader("content-disposition", "filename=Salary_Register.zip")
                        Using zip As New ZipFile()
                            zip.AddFile(filename, "Salary_Register")
                            zip.Save(Response.OutputStream)
                        End Using

                        If File.Exists(filename) Then
                            File.Delete(filename)
                        End If

                        lblmsg.Text = ""


                        Response.End()
                        'Response.Close()
                    Else
                        _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !")
                    End If
                    'New report for Monthly Excel sheet with dynamic column display,Added by praveen on 19 Feb 2013.
                Else
                    Dim arprm(7) As SqlClient.SqlParameter, apiurl As String
                    arprm(0) = New SqlClient.SqlParameter("@UserID", Session("UID").ToString)
                    arprm(1) = New SqlClient.SqlParameter("@Process_Type", GetApiProcessType(DDLPaySlipType.SelectedValue))
                    arprm(2) = New SqlClient.SqlParameter("@ActionType", "Init")
                    arprm(3) = New SqlClient.SqlParameter("@Sys_IP", "::1")
                    arprm(4) = New SqlClient.SqlParameter("@HostIP", ConfigurationManager.AppSettings("Hostip").ToString())
                    arprm(5) = New SqlClient.SqlParameter("@ProcName", "PaySP_SalaryRegInExcel")
                    arprm(6) = New SqlClient.SqlParameter("@DdlRptName", DDLPaySlipType.SelectedItem.Text.Replace("'", ""))
                    arprm(7) = New SqlClient.SqlParameter("@DdlRptId", DDLPaySlipType.SelectedValue)
                    Dim _dt As DataTable = _ObjData.GetDataTableProc("PaySP_ReportApi_ProcessBar", arprm)
                    If (_dt.Rows.Count > 0) Then
                        If (_dt.Rows(0)("IsAbleToStart").ToString = "1" AndAlso _dt.Rows(0)("BatchId").ToString <> "") Then
                            Dim scripttag As String = "StartProcessbar('" & DDLPaySlipType.SelectedItem.Text.Replace("'", "") & "');"
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpooce89sbar231", scripttag, True)
                            hdnBatchId.Value = _dt.Rows(0)("BatchId").ToString
                            btnProgressbarExcel.Visible = False
                            divSocialExcel.Visible = False
                            lblProcessStatusExcel.Text = ""
                            apiurl = _dt.Rows(0)("apiurl").ToString
                        Else
                            divSocialExcel.Visible = True
                            lblProcessStatusExcel.Text = DDLPaySlipType.SelectedItem.Text.Replace("'", "") & " is already processing. Please wait till the completion."
                            If (_dt.Rows(0)("UserId").ToString.ToUpper <> Session("UID").ToUpper.ToString) Then
                                btnProgressbarExcel.Visible = False
                            Else
                                btnProgressbarExcel.Visible = True
                            End If
                            _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = lblProcessStatusExcel.Text})
                            _objCommon.ShowMessage(_msg)
                            ScriptManager.RegisterStartupScript(Me.Page, GetType(Page), "refreshExeclProcessStatus99", "ShowExcelLockSummaryDetails('" & GetApiProcessType(DDLPaySlipType.SelectedValue) & "')", True)
                            Exit Sub
                        End If
                    End If
                    Dim keyValuePairs As New Dictionary(Of String, Object) From {
                        {"hostIp", ConfigurationManager.AppSettings("Hostip").ToString()},
                        {"userId", HttpContext.Current.Session("UID").ToString()},
                        {"moduleType", HttpContext.Current.Session("ModuleType").ToString()},
                        {"domainName", Session("CompCode").ToString},
                        {"showClr", rbtshowclr.SelectedValue.ToString.ToUpper},
                        {"fk_costcenter_code", USearch.UCddlcostcenter.ToString()},
                        {"fk_loc_code", USearch.UCddllocation.ToString()},
                        {"fk_unit", USearch.UCddlunit.ToString()},
                        {"salaried", USearch.UCddlsalbasis.ToString()},
                        {"pk_emp_code", USearch.UCTextcode.ToString.ToString},
                        {"fk_dept_code", USearch.UCddldept.ToString()},
                        {"fk_desig_code", USearch.UCddldesig.ToString()},
                        {"fk_grade_code", USearch.UCddlgrade.ToString()},
                        {"fk_level_code", USearch.UCddllevel.ToString()},
                        {"month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString)},
                        {"year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4)},
                        {"firstName", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), "")},
                        {"lastName", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), "")},
                        {"hold", ddlshowsal.SelectedValue.ToString},
                        {"userGroup", Session("Ugroup").ToString},
                        {"empStatus", USearch.UCddlEmp.ToString},
                        {"repId", DDLPaySlipType.SelectedValue.ToString},
                        {"BatchId", hdnBatchId.Value.ToString},
                        {"FileFormat", hdnFileFormat.Value.ToString}
                    }
                    Dim requestBody As String = JsonConvert.SerializeObject(keyValuePairs)
                    'Modified by Vishal Chauhan to call HeavyExcel API
                    CallReportAPIOnNewThreadHeavyExcel(requestBody, "StaticSalRegister", AppPathStr, apiurl)
                End If
            ElseIf DDLPaySlipType.SelectedValue.ToString = "41" Or DDLPaySlipType.SelectedValue.ToString = "42" Then
                arrparam(0) = New SqlParameter("@fk_costcenter_code", USearch.UCddlcostcenter.ToString())
                arrparam(1) = New SqlParameter("@fk_loc_code", USearch.UCddllocation.ToString())
                arrparam(2) = New SqlParameter("@fk_unit", USearch.UCddlunit.ToString())
                arrparam(3) = New SqlParameter("@salaried", USearch.UCddlsalbasis.ToString())
                arrparam(4) = New SqlParameter("@pk_emp_code", USearch.UCTextcode.ToString)
                arrparam(5) = New SqlParameter("@fk_dept_code", USearch.UCddldept.ToString())
                arrparam(6) = New SqlParameter("@fk_desig_code", USearch.UCddldesig.ToString())
                arrparam(7) = New SqlParameter("@fk_grade_code", USearch.UCddlgrade.ToString())
                arrparam(8) = New SqlParameter("@fk_level_Code", USearch.UCddllevel.ToString())
                arrparam(9) = New SqlClient.SqlParameter("@month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
                arrparam(10) = New SqlClient.SqlParameter("@year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                arrparam(11) = New SqlParameter("@first_name", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
                arrparam(12) = New SqlParameter("@last_name", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                arrparam(13) = New SqlParameter("@Hold", ddlshowsal.SelectedValue.ToString)
                arrparam(14) = New SqlParameter("@UserGroup", Session("Ugroup").ToString)
                arrparam(15) = New SqlParameter("@EmpStatus", USearch.UCddlEmp.ToString)
                arrparam(16) = New SqlParameter("@RepID", DDLPaySlipType.SelectedValue.ToString)
                If DDLPaySlipType.SelectedValue.ToString = "41" Then
                    ds = _ObjData.GetDataSetProc("PaySP_SalaryRegInExcel_FormR", arrparam)
                ElseIf DDLPaySlipType.SelectedValue.ToString = "42" Then
                    ds = _ObjData.GetDataSetProc("PaySp_MinWagesSalRegister_Eversendai", arrparam)
                Else
                    ds = _ObjData.GetDataSetProc("PaySP_SalaryRegInExcel", arrparam)
                End If
                If ds.Tables(0).Rows.Count > 0 Then
                    'Geeta2012
                    Dim filename As String = "", _sw As StreamWriter, complexID As Guid = Guid.NewGuid()

                    filename = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\Salary_Register" & "_" & Right(complexID.ToString, 6) & ".xls"
                    If System.IO.File.Exists(filename) Then
                        System.IO.File.Delete(filename)
                    End If
                    _sw = New StreamWriter(filename)
                    ExportToExcelXML(ds, _sw)
                    _sw.Close()
                    _sw.Dispose()

                    Response.Clear()
                    Response.BufferOutput = False

                    'Dim ReadmeText As [String] = "Hello!" & vbLf & vbLf & "This is a README..." & DateTime.Now.ToString("G")
                    Response.ContentType = "application/zip"
                    Response.AddHeader("content-disposition", "filename=Salary_Register.zip")
                    Using zip As New ZipFile()
                        zip.AddFile(filename, "Salary_Register")
                        zip.Save(Response.OutputStream)
                    End Using

                    If File.Exists(filename) Then
                        File.Delete(filename)
                    End If

                    lblmsg.Text = ""


                    Response.End()
                    'Response.Close()
                Else
                    _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !", False)
                End If
                'New report for Monthly Excel sheet with dynamic column display,Added by praveen on 19 Feb 2013.
            ElseIf DDLPaySlipType.SelectedValue = "38" Then
                If ddlrepformat.SelectedValue.Equals("XLS") Then
                    'Added by Debargha on 21-Oct-2024
                    Dim sParam(0) As SqlClient.SqlParameter
                    sParam(0) = New SqlClient.SqlParameter("@SP_Name", "PaySP_SalaryRegInExcel_Dynamic")
                    '_ObjData.ExecuteStoredProcMsg("PaySP_ReportAPI_ConfigSel", APIConfigParam)
                    Dim dt As DataTable = _ObjData.GetDataTableProc("PaySP_ReportAPI_ConfigSel", sParam)
                    If dt.Rows.Count = 0 OrElse dt.Rows(0)("isNewURL").ToString.ToUpper.Trim() <> "Y" Then
                        arrparam(0) = New SqlParameter("@fk_costcenter_code", USearch.UCddlcostcenter.ToString())
                        arrparam(1) = New SqlParameter("@fk_loc_code", USearch.UCddllocation.ToString())
                        arrparam(2) = New SqlParameter("@fk_unit", USearch.UCddlunit.ToString())
                        arrparam(3) = New SqlParameter("@salaried", USearch.UCddlsalbasis.ToString())
                        arrparam(4) = New SqlParameter("@pk_emp_code", USearch.UCTextcode.ToString)
                        arrparam(5) = New SqlParameter("@fk_dept_code", USearch.UCddldept.ToString())
                        arrparam(6) = New SqlParameter("@fk_desig_code", USearch.UCddldesig.ToString())
                        arrparam(7) = New SqlParameter("@fk_grade_code", USearch.UCddlgrade.ToString())
                        arrparam(8) = New SqlParameter("@fk_level_Code", USearch.UCddllevel.ToString())
                        arrparam(9) = New SqlClient.SqlParameter("@month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
                        arrparam(10) = New SqlClient.SqlParameter("@year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                        arrparam(11) = New SqlParameter("@first_name", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
                        arrparam(12) = New SqlParameter("@last_name", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                        arrparam(13) = New SqlParameter("@Hold", ddlshowsal.SelectedValue.ToString)
                        arrparam(14) = New SqlParameter("@UserGroup", Session("Ugroup").ToString)
                        arrparam(15) = New SqlParameter("@EmpStatus", USearch.UCddlEmp.ToString)
                        arrparam(16) = New SqlParameter("@RepID", DDLPaySlipType.SelectedValue.ToString)
                        ds = _ObjData.GetDataSetProc("PaySP_SalaryRegInExcel_Dynamic", arrparam)
                        If ds.Tables(0).Rows.Count > 0 Then
                            Dim filename As String = "", _sw As StreamWriter, complexID As Guid = Guid.NewGuid()
                            filename = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\Salary Register" & "_" & Right(complexID.ToString, 6) & ".xls"
                            MakeZipFolderbyXml(ds, filename, "Salary Register")
                        Else
                            _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !", False)
                        End If
                    End If
                    If dt.Rows.Count > 0 AndAlso dt.Rows(0)("isNewURL").ToString.ToUpper.Trim() = "Y" Then
                        If dt.Rows(0)("WithoutProcessbar").ToString.ToUpper.Trim() = "Y" Then
                            'Report API code(Debaragha) without Processbar
                            Dim keyValuePairs As New Dictionary(Of String, Object) From {
                            {"hostIp", ConfigurationManager.AppSettings("Hostip").ToString()},
                            {"userId", HttpContext.Current.Session("UID").ToString()},
                            {"moduleType", HttpContext.Current.Session("ModuleType").ToString()},
                            {"domainName", Session("CompCode").ToString},
                            {"showClr", rbtshowclr.SelectedValue.ToString.ToUpper},
                            {"fk_costcenter_code", USearch.UCddlcostcenter.ToString()},
                            {"fk_loc_code", USearch.UCddllocation.ToString()},
                            {"fk_unit", USearch.UCddlunit.ToString()},
                            {"salaried", USearch.UCddlsalbasis.ToString()},
                            {"pk_emp_code", USearch.UCTextcode.ToString.ToString},
                            {"fk_dept_code", USearch.UCddldept.ToString()},
                            {"fk_desig_code", USearch.UCddldesig.ToString()},
                            {"fk_grade_code", USearch.UCddlgrade.ToString()},
                            {"fk_level_code", USearch.UCddllevel.ToString()},
                            {"month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString)},
                            {"year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4)},
                            {"firstName", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), "")},
                            {"lastName", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), "")},
                            {"hold", ddlshowsal.SelectedValue.ToString},
                            {"userGroup", Session("Ugroup").ToString},
                            {"empStatus", USearch.UCddlEmp.ToString},
                            {"repId", DDLPaySlipType.SelectedValue.ToString},
                            {"FileName", txtrptName.Text.ToString}
                        }
                            Dim requestBody As String = JsonConvert.SerializeObject(keyValuePairs)
                            CallAPIReport(requestBody, "SalaryRegister")
                        Else
                            'New Report API(Vishal) code to show Processbar
                            Dim arprm(8) As SqlClient.SqlParameter, apiurl As String
                            arprm(0) = New SqlClient.SqlParameter("@UserID", Session("UID").ToString)
                            arprm(1) = New SqlClient.SqlParameter("@Process_Type", GetApiProcessType(DDLPaySlipType.SelectedValue))
                            arprm(2) = New SqlClient.SqlParameter("@ActionType", "Init")
                            arprm(3) = New SqlClient.SqlParameter("@Sys_IP", "::1")
                            arprm(4) = New SqlClient.SqlParameter("@HostIP", ConfigurationManager.AppSettings("Hostip").ToString())
                            arprm(5) = New SqlClient.SqlParameter("@ProcName", "PaySP_SalaryRegInExcel_Dynamic")
                            arprm(6) = New SqlClient.SqlParameter("@DdlRptName", DDLPaySlipType.SelectedItem.Text.Replace("'", ""))
                            arprm(7) = New SqlClient.SqlParameter("@DdlRptId", DDLPaySlipType.SelectedValue)
                            arprm(8) = New SqlClient.SqlParameter("@RPTFRMT", "XLS")
                            Dim _dt As DataTable = _ObjData.GetDataTableProc("PaySP_ReportApi_ProcessBar", arprm)
                            If (_dt.Rows.Count > 0) Then
                                If (_dt.Rows(0)("IsAbleToStart").ToString = "1" AndAlso _dt.Rows(0)("BatchId").ToString <> "") Then
                                    Dim scripttag As String = "StartProcessbar('" & DDLPaySlipType.SelectedItem.Text.Replace("'", "") & "');"
                                    'ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpooce89sbar231", scripttag, True)
                                    ScriptManager.RegisterStartupScript(Me, Me.GetType(), "openpooce89sbar231", scripttag, True)
                                    hdnBatchId.Value = _dt.Rows(0)("BatchId").ToString
                                    btnProgressbarExcel.Visible = False
                                    lblProcessStatusExcel.Text = ""
                                    divSocialExcel.Visible = False
                                    apiurl = _dt.Rows(0)("apiurl").ToString
                                Else
                                    divSocialExcel.Visible = True
                                    lblProcessStatusExcel.Text = DDLPaySlipType.SelectedItem.Text.Replace("'", "") & " is already processing. Please wait till the completion."
                                    If (_dt.Rows(0)("UserId").ToString.ToUpper <> Session("UID").ToUpper.ToString) Then
                                        btnProgressbarExcel.Visible = True
                                    Else
                                        btnProgressbarExcel.Visible = True
                                    End If
                                    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = lblProcessStatusExcel.Text})
                                    _objCommon.ShowMessage(_msg)
                                    ScriptManager.RegisterStartupScript(Me.Page, GetType(Page), "refreshExeclProcessStatus996", "ShowExcelLockSummaryDetails('" & _dt.Rows(0)("Process_Type").ToString.ToUpper & "')", True)
                                    Exit Sub
                                End If
                            End If
                            Dim keyValuePairs As New Dictionary(Of String, Object) From {
                            {"hostIp", ConfigurationManager.AppSettings("Hostip").ToString()},
                            {"userId", HttpContext.Current.Session("UID").ToString()},
                            {"moduleType", HttpContext.Current.Session("ModuleType").ToString()},
                            {"domainName", Session("CompCode").ToString},
                            {"showClr", rbtshowclr.SelectedValue.ToString.ToUpper},
                            {"fk_costcenter_code", USearch.UCddlcostcenter.ToString()},
                            {"fk_loc_code", USearch.UCddllocation.ToString()},
                            {"fk_unit", USearch.UCddlunit.ToString()},
                            {"salaried", USearch.UCddlsalbasis.ToString()},
                            {"pk_emp_code", USearch.UCTextcode.ToString.ToString},
                            {"fk_dept_code", USearch.UCddldept.ToString()},
                            {"fk_desig_code", USearch.UCddldesig.ToString()},
                            {"fk_grade_code", USearch.UCddlgrade.ToString()},
                            {"fk_level_code", USearch.UCddllevel.ToString()},
                            {"month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString)},
                            {"year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4)},
                            {"firstName", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), "")},
                            {"lastName", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), "")},
                            {"hold", ddlshowsal.SelectedValue.ToString},
                            {"userGroup", Session("Ugroup").ToString},
                            {"empStatus", USearch.UCddlEmp.ToString},
                            {"repId", DDLPaySlipType.SelectedValue.ToString},
                            {"BatchId", hdnBatchId.Value.ToString},
                            {"FileFormat", hdnFileFormat.Value.ToString},
                            {"FileName", txtrptName.Text.ToString}
                        }
                            Dim requestBody As String = JsonConvert.SerializeObject(keyValuePairs)
                            CallReportAPIOnNewThread(requestBody, "SalaryRegister", AppPathStr, apiurl)
                        End If

                    End If
                Else
                    If ddllEncrType.SelectedValue.ToUpper() = "WP" Then
                        Dim dt As New DataTable()
                        Dim param(3) As SqlClient.SqlParameter

                        param(0) = New SqlClient.SqlParameter("@ReportId", "38")
                        param(1) = New SqlClient.SqlParameter("@ReportFrmt", ddlrepformat.SelectedValue.ToUpper().Trim())
                        param(2) = New SqlClient.SqlParameter("@EncrType", ddllEncrType.SelectedValue.Trim())
                        param(3) = New SqlClient.SqlParameter("@Flag", "C")

                        dt = _ObjData.GetDataTableProc("Paysp_MstPGPEncryptionConfig_GetEncrKey", param)

                        If dt.Rows.Count = 0 Then
                            Dim _msg As New List(Of PayrollUtility.UserMessage)
                            _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "M", .MessageString = "Please configure the PGP encryption key or activate an existing key in Configuration → Company Miscellaneous Settings → PGP Encryption Configuration. Once completed, try the PGP encryption process again !"})
                            _objCommon.ShowMessage(_msg)
                            Exit Sub
                        End If
                    End If
                    ExportCSV()
                End If
                ' Add new payslip [Salary register group wise.], by praveen verma on 23 Aug 2013.
            ElseIf DDLPaySlipType.SelectedValue = "44" Then
                arrparam(0) = New SqlParameter("@fk_costcenter_code", USearch.UCddlcostcenter.ToString())
                arrparam(1) = New SqlParameter("@fk_loc_code", USearch.UCddllocation.ToString())
                arrparam(2) = New SqlParameter("@fk_unit", USearch.UCddlunit.ToString())
                arrparam(3) = New SqlParameter("@salaried", USearch.UCddlsalbasis.ToString())
                arrparam(4) = New SqlParameter("@pk_emp_code", USearch.UCTextcode.ToString)
                arrparam(5) = New SqlParameter("@fk_dept_code", USearch.UCddldept.ToString())
                arrparam(6) = New SqlParameter("@fk_desig_code", USearch.UCddldesig.ToString())
                arrparam(7) = New SqlParameter("@fk_grade_code", USearch.UCddlgrade.ToString())
                arrparam(8) = New SqlParameter("@fk_level_Code", USearch.UCddllevel.ToString())
                arrparam(9) = New SqlClient.SqlParameter("@month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
                arrparam(10) = New SqlClient.SqlParameter("@year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                arrparam(11) = New SqlParameter("@first_name", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
                arrparam(12) = New SqlParameter("@last_name", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                arrparam(13) = New SqlParameter("@Hold", ddlshowsal.SelectedValue.ToString)
                arrparam(14) = New SqlParameter("@UserGroup", Session("Ugroup").ToString)
                arrparam(15) = New SqlParameter("@EmpStatus", USearch.UCddlEmp.ToString)
                arrparam(16) = New SqlParameter("@RepID", DDLPaySlipType.SelectedValue.ToString)
                ds = _ObjData.GetDataSetProc("PaySP_SalaryRegInExcelGrpWise", arrparam)
                If ds.Tables(0).Rows.Count > 0 Then
                    Dim filename As String = "", _sw As StreamWriter, complexID As Guid = Guid.NewGuid()
                    filename = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\Salary Register GroupWise" & "_" & Right(complexID.ToString, 6) & ".xls"
                    If System.IO.File.Exists(filename) Then
                        System.IO.File.Delete(filename)
                    End If
                    _sw = New StreamWriter(filename)
                    lblmsg.Text = ""
                    ExportToExcelXMLGrpWise(ds, _sw)
                    _sw.Close()
                    _sw.Dispose()
                    Response.Clear()
                    Response.ContentType = "application/zip"
                    Response.AddHeader("content-disposition", "filename=Salary Register GroupWise.zip")
                    Using zip As New ZipFile()
                        zip.AddFile(filename, "Salary Register GroupWise")
                        zip.Save(Response.OutputStream)
                    End Using

                    If File.Exists(filename) Then
                        File.Delete(filename)
                    End If
                    Response.End()
                Else
                    _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !", False)
                End If

                '---------Added By Geeta on 30 Aug 2012
            ElseIf DDLPaySlipType.SelectedValue = "17" Then
                Dim _Param(27) As SqlClient.SqlParameter
                _Param(0) = New SqlClient.SqlParameter("@Report", _objCommon.nNz(DDLPaySlipType.SelectedValue.ToString))
                _Param(1) = New SqlClient.SqlParameter("@month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
                _Param(2) = New SqlClient.SqlParameter("@year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                _Param(3) = New SqlParameter("@pk_emp_code", USearch.UCTextcode.ToString)
                _Param(4) = New SqlParameter("@first_name", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
                _Param(5) = New SqlParameter("@last_name", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                _Param(6) = New SqlParameter("@fk_costcenter_code", USearch.UCddlcostcenter.ToString())
                _Param(7) = New SqlParameter("@fk_dept_code", USearch.UCddldept.ToString())
                _Param(8) = New SqlParameter("@fk_desig_code", USearch.UCddldesig.ToString())
                _Param(9) = New SqlParameter("@fk_grade_code", USearch.UCddlgrade.ToString())
                _Param(10) = New SqlParameter("@fk_loc_code", USearch.UCddllocation.ToString())
                _Param(11) = New SqlParameter("@fk_unit", USearch.UCddlunit.ToString())
                _Param(12) = New SqlParameter("@salaried", USearch.UCddlsalbasis.ToString())
                _Param(13) = New SqlParameter("@fk_level_Code", USearch.UCddllevel.ToString())
                _Param(14) = New SqlParameter("@EmpStatus", USearch.UCddlEmp.ToString)
                _Param(15) = New SqlParameter("@UserGroup", Session("Ugroup").ToString)
                _Param(16) = New SqlParameter("@Hold", Hold.ToString)
                _Param(17) = New SqlParameter("@SalFlg", "A")
                _Param(18) = New SqlParameter("@ArrFlg", "Y")
                'Added by Geeta on 24 Apr 10
                _Param(19) = New SqlParameter("@Fdate", "")
                _Param(20) = New SqlParameter("@Tdate", "")
                _Param(21) = New SqlParameter("@PayMOde", "")
                'Added By geeta on 21 Jul 10
                _Param(22) = New SqlParameter("@FromDol", "")
                _Param(23) = New SqlParameter("@ToDol", "")
                _Param(24) = New SqlParameter("@ENTflg", "N")
                _Param(25) = New SqlParameter("@LeaveFlag", "Y")
                _Param(26) = New SqlParameter("@servicePer", "12.36")
                _Param(27) = New SqlParameter("@RepDesc", DDLPaySlipType.SelectedItem.ToString)
                ds = _ObjData.GetDataSetProc("PaySp_MstReportDetails_CTCElixer", _Param)
                If ds.Tables(0).Rows.Count > 0 Then
                    Dim filename As String = "", _sw As StreamWriter, complexID As Guid = Guid.NewGuid()
                    filename = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\Salary Register" & "_" & Right(complexID.ToString, 6) & ".xls"
                    If System.IO.File.Exists(filename) Then
                        System.IO.File.Delete(filename)
                    End If
                    _sw = New StreamWriter(filename)
                    ExportToExcelXML_GrandTotal(ds, _sw)
                    _sw.Close()
                    _sw.Dispose()

                    Response.ContentType = "application/zip"
                    Response.AddHeader("content-disposition", "filename=Salary Register.zip")
                    Using zip As New ZipFile()
                        zip.AddFile(filename, "Salary Register")
                        zip.Save(Response.OutputStream)
                    End Using

                    If File.Exists(filename) Then
                        File.Delete(filename)
                    End If
                    Response.End()
                Else
                    _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !", False)
                End If
            ElseIf DDLPaySlipType.SelectedValue = "27" Then
                Dim _Param(27) As SqlClient.SqlParameter
                _Param(0) = New SqlClient.SqlParameter("@Report", "24")
                _Param(1) = New SqlClient.SqlParameter("@month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
                _Param(2) = New SqlClient.SqlParameter("@year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                _Param(3) = New SqlParameter("@pk_emp_code", USearch.UCTextcode.ToString)
                _Param(4) = New SqlParameter("@first_name", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
                _Param(5) = New SqlParameter("@last_name", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                _Param(6) = New SqlParameter("@fk_costcenter_code", USearch.UCddlcostcenter.ToString())
                _Param(7) = New SqlParameter("@fk_dept_code", USearch.UCddldept.ToString())
                _Param(8) = New SqlParameter("@fk_desig_code", USearch.UCddldesig.ToString())
                _Param(9) = New SqlParameter("@fk_grade_code", USearch.UCddlgrade.ToString())
                _Param(10) = New SqlParameter("@fk_loc_code", USearch.UCddllocation.ToString())
                _Param(11) = New SqlParameter("@fk_unit", USearch.UCddlunit.ToString())
                _Param(12) = New SqlParameter("@salaried", USearch.UCddlsalbasis.ToString())
                _Param(13) = New SqlParameter("@fk_level_Code", USearch.UCddllevel.ToString())
                _Param(14) = New SqlParameter("@EmpStatus", USearch.UCddlEmp.ToString)
                _Param(15) = New SqlParameter("@UserGroup", Session("Ugroup").ToString)
                _Param(16) = New SqlParameter("@Hold", Hold.ToString)
                _Param(17) = New SqlParameter("@SalFlg", "A")
                _Param(18) = New SqlParameter("@ArrFlg", "Y")
                'Added by Geeta on 24 Apr 10
                _Param(19) = New SqlParameter("@Fdate", "")
                _Param(20) = New SqlParameter("@Tdate", "")
                _Param(21) = New SqlParameter("@PayMOde", "")
                'Added By geeta on 21 Jul 10
                _Param(22) = New SqlParameter("@FromDol", "")
                _Param(23) = New SqlParameter("@ToDol", "")
                _Param(24) = New SqlParameter("@ENTflg", "N")
                _Param(25) = New SqlParameter("@LeaveFlag", "Y")
                _Param(26) = New SqlParameter("@servicePer", "12.36")
                _Param(27) = New SqlParameter("@RepDesc", DDLPaySlipType.SelectedItem.ToString)
                ds = _ObjData.GetDataSetProc("PaySp_MstReportDetails_CTCActual", _Param)
                If ds.Tables(0).Rows.Count > 0 Then
                    Dim filename As String = "", _sw As StreamWriter, complexID As Guid = Guid.NewGuid()
                    filename = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\Salary Register" & "_" & Right(complexID.ToString, 6) & ".xls"
                    If System.IO.File.Exists(filename) Then
                        System.IO.File.Delete(filename)
                    End If
                    _sw = New StreamWriter(filename)
                    ExportToExcelXML_GrandTotal(ds, _sw)
                    _sw.Close()
                    _sw.Dispose()

                    Response.ContentType = "application/zip"
                    Response.AddHeader("content-disposition", "filename=Salary Register.zip")
                    Using zip As New ZipFile()
                        zip.AddFile(filename, "Salary Register")
                        zip.Save(Response.OutputStream)
                    End Using

                    If File.Exists(filename) Then
                        File.Delete(filename)
                    End If
                    Response.End()
                Else
                    _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !", False)
                End If
                '---------
                'This statement is used for display "Monthly Pay Slip Include Details With Leave Bal"
            ElseIf DDLPaySlipType.SelectedValue = "0" Then

                HidPreVal.Value = "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString &
                "~" & Year.ToString & "~" & EmpLName.ToString & "~" &
                Comzero.ToString & "~" &
                Comlogo.ToString & "~" & Hold.ToString _
                & "~~" & Comzero.ToString & "~" & Comlogo.ToString & "~S" _
                & "~~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString _
                & "~~~~" & "N" 'Added by geeta on 11 sep 2012

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'Added by Rajesh on 01 Oct 2014 for bajaj salary slip
            ElseIf DDLPaySlipType.SelectedValue = "50" Then

                HidPreVal.Value = "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString &
                "~" & Year.ToString & "~" & EmpLName.ToString & "~" &
                Comzero.ToString & "~" &
                Comlogo.ToString & "~" & Hold.ToString _
                & "~~" & Comzero.ToString & "~" & Comlogo.ToString & "~S" _
                & "~~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString _
                & "~~~~" & "N"
                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'This statement is used for display "Monthly Salary Register"
            ElseIf DDLPaySlipType.SelectedValue = "1" Then
                reptype = DDLPaySlipType.SelectedValue.ToString
                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                 Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                 SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString _
                 & "~" & Year.ToString & "~" & EmpLName.ToString & "~" _
                 & EmpType.ToString & "~" & Hold & "~" & DDLPaySlipType.SelectedValue.ToString
                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'This statement is used for display "Salary Register Head Wise"
            ElseIf DDLPaySlipType.SelectedValue = "2" Then
                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                 Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                 SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                 Month.ToString & "~" & Year.ToString & "~" & EmpLName.ToString &
                 "~" & Hold.ToString & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString
                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)

                'This statement is used for display "Salary Register Head Wise" & "Monthly Final Salary"
                'Add condition [DDLPaySlipType.SelectedValue = "45"] by Nisha on 31 Aug 2013
            ElseIf DDLPaySlipType.SelectedValue = "45" Then
                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                 Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                 SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                 Month.ToString & "~" & Year.ToString & "~" & EmpLName.ToString &
                 "~" & Hold.ToString & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString
                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                'Added by Nisha on 25 Sep 2013
                dtRepType = _ObjData.ExecSQLQuery("Select IsNull(RepFormat,'E') RepFormat from payslipconfigure where pk_rep_id=" & DDLPaySlipType.SelectedValue.ToString)

                If dtRepType.Rows.Count > 0 Then
                    ReportType = dtRepType.Rows(0)("RepFormat").ToString
                End If

                If ReportType.ToString = "H" Then
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                Else
                    ExporttoExcelforFinalSalary()
                End If
                'This statement is used for display "Monthly Arrear Register"
            ElseIf DDLPaySlipType.SelectedValue = "3" Then
                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString _
                & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString
                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'This statement is used for display "Monthly CTC Arrear Register"
            ElseIf DDLPaySlipType.SelectedValue = "4" Then
                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                Month.ToString & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & DDLPaySlipType.SelectedValue.ToString & "~" & EmpType.ToString
                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'This statement is used for display "Monthly Pay Slip With Investment"
            ElseIf DDLPaySlipType.SelectedValue = "5" Then

                If chkboxent.Checked = True Then
                    Ent = 1
                Else
                    Ent = 0
                End If

                If chkboxotherinc.Checked = True Then
                    otherinc = 1
                Else
                    otherinc = 0
                End If

                For Each chkitem In CBLsection.Items
                    If chkitem.Selected = True Then
                        str = str + "'" + chkitem.Text + "',"
                    End If
                Next
                If str.ToString <> "" Then
                    str = Left(str, Len(str) - 1)
                    hdquery.Value = str 'hdquery.Value + "," + str
                End If
                PayCodeId = hdquery.Value.ToString

                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                Month.ToString & "~" & Year.ToString & "~" & EmpLName.ToString & "~" _
                & Ent.ToString & "~" & otherinc.ToString & "~" & Comzero.ToString & "~" & Hold.ToString _
                & "~" & "~" & "~" & "~" & "~" _
                & PayCodeId.ToString & "~" & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)

                'This statement is used for display "Department Wise Salary"
            ElseIf DDLPaySlipType.SelectedValue = "6" Then
                If rbHorizontal.Checked = True Then
                    ReptType = "H"
                Else
                    ReptType = "V"
                End If

                RepFormat = ddlformat.SelectedValue.ToString

                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                Month.ToString & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & ReptType.ToString & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If


                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'This statement is used for display "Monthly Pay Slip Include Details W/O Leave Bal"
            ElseIf DDLPaySlipType.SelectedValue = "12" Then

                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                Month.ToString & "~" & Year.ToString & "~" & EmpLName.ToString & "~" _
                & Comzero.ToString & "~" &
                 Hold.ToString & "~" _
                & "~S" & "~" & "~~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'This statement is used for display "Monthly Pay Slip Exclude Details"
            ElseIf DDLPaySlipType.SelectedValue = "7" Then

                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                Month.ToString & "~" & Year.ToString & "~" & EmpLName.ToString & "~" &
                Hold.ToString & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString
                '& IIf(ChkPF.Checked = True, "Y", "N") & "~" & IIf(ChkESI.Checked = True, "Y", "N")

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'This statement is used for display "Monthly A4Size Salary Register"
            ElseIf DDLPaySlipType.SelectedValue = "8" Then
                'here we store all selected list item in the paycodesel1 variable.
                For Each lstitem1 In cbldeduction.Items
                    If lstitem1.Selected = True Then
                        paycodeSel1 = paycodeSel1 + lstitem1.Value + ","
                    End If
                Next
                'here we store all selected list item in the paycodesel variable.
                For Each lstitem In cbladd.Items
                    If lstitem.Selected = True Then
                        paycodeSel = paycodeSel + lstitem.Value + ","
                    End If
                Next
                'here we add all paycode in hidden control
                PayCodeId = paycodeSel1 + paycodeSel

                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                Month.ToString & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & Hold.ToString & "~" _
                & PayCodeId.ToString & "~" & DDLPaySlipType.SelectedValue.ToString & "~" & EmpType.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'This statement is used for display "Monthly A4Size(4Rows) Salary Register"
            ElseIf DDLPaySlipType.SelectedValue = "9" Then
                Ids = hiddls_id.Value.ToString
                srttp = ddlshortbasis.SelectedValue.ToString
                If srttp.ToString <> "" Then
                    Left(srttp, Len(srttp) - 1)
                End If

                'here we store all selected list item in the paycodesel1 variable.
                For Each lstitem1 In cbldeduction.Items
                    If lstitem1.Selected = True Then
                        paycodeSel1 = paycodeSel1 + lstitem1.Value + ","
                    End If
                Next
                'here we store all selected list item in the paycodesel variable.
                For Each lstitem In cbladd.Items
                    If lstitem.Selected = True Then
                        paycodeSel = paycodeSel + lstitem.Value + ","
                    End If
                Next
                'here we add all paycode in hidden control
                PayCodeId = paycodeSel1 + paycodeSel

                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString _
                & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & Ids.ToString & "~" & srttp.ToString _
                & "~" & Hold.ToString & "~" & PayCodeId.ToString & "~" & DDLPaySlipType.SelectedValue.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'This statement is used for display "Monthly Pay Slip With PF and Loan details"
            ElseIf DDLPaySlipType.SelectedValue = "10" Then

                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                Month.ToString & "~" & Year.ToString & "~" & EmpLName.ToString & "~" _
                & Hold.ToString & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString & "~"    'Add UserGroup blank by Niraj on 11 Apr 2013

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)

                'This statement is used for display "Annual Arrear Details"
            ElseIf DDLPaySlipType.SelectedValue = "11" Then
                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString _
                & "~" & Year.ToString & "~" & EmpLName.ToString & "~" &
                EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)

                'This statement is used for display "Monthly Pay Slip Include Details W/O Leave Bal"
            ElseIf DdlreportType.SelectedValue = "12" Then

                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                Month.ToString & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & Salut.ToString & "~" _
                & chkPF_Old.ToString & "~" & chkESI_Old.ToString & "~" & Comzero.ToString & "~" &
                Comlogo.ToString & "~" & Hold.ToString & "~" & Loan.ToString & "~" _
                & Advance.ToString & "~" & "~S" & "~" & StaffId.ToString & "~~" & EmpType.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)

                'This statement is used for display "Monthly A4Size(4Rows) Salary Register Bold Caption"
            ElseIf DDLPaySlipType.SelectedValue = "13" Then
                Ids = hiddls_id.Value.ToString
                srttp = ddlshortbasis.SelectedValue.ToString
                If srttp.ToString <> "" Then
                    Left(srttp, Len(srttp) - 1)
                End If

                'here we store all selected list item in the paycodesel1 variable.
                For Each lstitem1 In cbldeduction.Items
                    If lstitem1.Selected = True Then
                        paycodeSel1 = paycodeSel1 + lstitem1.Value + ","
                    End If
                Next
                'here we store all selected list item in the paycodesel variable.
                For Each lstitem In cbladd.Items
                    If lstitem.Selected = True Then
                        paycodeSel = paycodeSel + lstitem.Value + ","
                    End If
                Next
                'here we add all paycode in hidden control
                PayCodeId = paycodeSel1 + paycodeSel

                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString _
                & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & Ids.ToString & "~" & srttp.ToString _
                & "~" & Hold.ToString & "~" & PayCodeId.ToString & "~" _
                & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'This statement is used for display "Register Of Payment Of Wages"
            ElseIf DDLPaySlipType.SelectedValue = "14" Then
                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                 Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                 SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                 Month.ToString & "~" & Year.ToString & "~" & EmpLName.ToString &
                 "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)

                'This statement is used for display "Monthly A4Size(5Rows) Salary Register (A2Z)"
            ElseIf DDLPaySlipType.SelectedValue = "15" Then
                Ids = hiddls_id.Value.ToString
                srttp = ddlshortbasis.SelectedValue.ToString
                If srttp.ToString <> "" Then
                    Left(srttp, Len(srttp) - 1)
                End If

                For Each lstitem1 In cbldeduction.Items
                    If lstitem1.Selected = True Then
                        paycodeSel1 = paycodeSel1 + lstitem1.Value + ","
                    End If
                Next

                For Each lstitem In cbladd.Items
                    If lstitem.Selected = True Then
                        paycodeSel = paycodeSel + lstitem.Value + ","
                    End If
                Next
                'here we add all paycode in hidden control
                PayCodeId = paycodeSel1 + paycodeSel

                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString _
                & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & Ids.ToString & "~" & srttp.ToString _
                & "~" & Hold.ToString & "~" & PayCodeId.ToString & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'added by Rajesh for add Monthly Salary Register for luxur on 04 oct 13
            ElseIf DDLPaySlipType.SelectedValue = "46" Then
                Ids = hiddls_id.Value.ToString
                srttp = ddlshortbasis.SelectedValue.ToString
                If srttp.ToString <> "" Then
                    Left(srttp, Len(srttp) - 1)
                End If

                For Each lstitem1 In cbldeduction.Items
                    If lstitem1.Selected = True Then
                        paycodeSel1 = paycodeSel1 + lstitem1.Value + ","
                    End If
                Next

                For Each lstitem In cbladd.Items
                    If lstitem.Selected = True Then
                        paycodeSel = paycodeSel + lstitem.Value + ","
                    End If
                Next
                'here we add all paycode in hidden control
                PayCodeId = paycodeSel1 + paycodeSel
                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString _
                & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & Ids.ToString & "~" & srttp.ToString _
                & "~" & Hold.ToString & "~" & PayCodeId.ToString & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString


                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)

                'This statement is used for display "Monthly A4Size(4Rows) Sal. Reg. With Arrear"
            ElseIf DDLPaySlipType.SelectedValue = "16" Then

                getGroupVal()
                For Each lstitem1 In cbldeduction.Items
                    If lstitem1.Selected = True Then
                        paycodeSel1 = paycodeSel1 + lstitem1.Value + ","
                    End If
                Next
                'here we store all selected list item in the paycodesel variable.
                For Each lstitem In cbladd.Items
                    If lstitem.Selected = True Then
                        paycodeSel = paycodeSel + lstitem.Value + ","
                    End If
                Next
                'here we add all paycode in hidden control
                PayCodeId = paycodeSel1 + paycodeSel

                Ids = hiddls_id.Value.ToString
                Ids = Ids + "!" + hidGroup2Val.Value.ToString
                Ids = Ids + "!" + hidGroup3Val.Value.ToString

                srttp = ddlshortbasis.SelectedValue.ToString
                If srttp.ToString <> "" Then
                    Left(srttp, Len(srttp) - 1)
                End If
                srttp = srttp + "!" + ddlGroup2.SelectedValue.ToString
                srttp = srttp + "!" + ddlGroup3.SelectedValue.ToString

                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString _
                & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & Ids.ToString & "~" & srttp.ToString _
                & "~" & Hold.ToString & "~" & PayCodeId.ToString _
                & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'This statement is used for display "Monthly Salary Slip With Tax Details"
            ElseIf DDLPaySlipType.SelectedValue = "18" Or DDLPaySlipType.SelectedValue = "52" Or DDLPaySlipType.SelectedValue = "51" Or DDLPaySlipType.SelectedValue = "49" Or DDLPaySlipType.SelectedValue.Trim.Equals("74") Then
                Dim LeaveBal As String = "", Flag As String = "S", lstitemOtg As ListItem
                For Each lstitemOtg In Chklistothepaycode.Items
                    If lstitemOtg.Selected = True Then
                        hidothrpaycode.Value = hidothrpaycode.Value.ToString + lstitemOtg.Value + ","
                    End If
                Next

                'Add on blank flag after "StaffID" for employee password
                HidPreVal.Value = "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString _
                & "~" & Year.ToString & "~" & EmpLName.ToString & "~" &
                 "~" & "~" & Comzero.ToString & "~" &
                Comlogo.ToString & "~" & Hold.ToString _
                & "~" & Flag.ToString & "~" & EmpType.ToString _
                & "~" & "0" & "~" & IIf(chkother.Checked.ToString, "Y", "N") & "~" & hidothrpaycode.Value.ToString & "~" & DDLPaySlipType.SelectedValue.ToString & "~~~~~" & IIf(chkHelp.Checked = True, "Y", "N").ToString


                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'lit.Text = _str.ToString

                'This statement is used for display "Monthly A4Size(5Rows) Salary Register with Arrear"
            ElseIf DDLPaySlipType.SelectedValue = "19" Then
                Ids = hiddls_id.Value.ToString
                srttp = ddlshortbasis.SelectedValue.ToString
                If srttp.ToString <> "" Then
                    Left(srttp, Len(srttp) - 1)
                End If
                'here we store all selected list item in the paycodesel1 variable.
                For Each lstitem1 In cbldeduction.Items
                    If lstitem1.Selected = True Then
                        paycodeSel1 = paycodeSel1 + lstitem1.Value + ","
                    End If
                Next
                'here we store all selected list item in the paycodesel variable.
                For Each lstitem In cbladd.Items
                    If lstitem.Selected = True Then
                        paycodeSel = paycodeSel + lstitem.Value + ","
                    End If
                Next
                'here we add all paycode in hidden control
                PayCodeId = paycodeSel1 + paycodeSel

                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString _
                & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & Ids.ToString & "~" & srttp.ToString _
                & "~" & Hold.ToString & "~" & PayCodeId.ToString _
                & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)

                'Added by Jay on 19 Mar 2014 for 
            ElseIf DDLPaySlipType.SelectedValue = "48" Or DDLPaySlipType.SelectedValue = "47" Then
                Ids = hiddls_id.Value.ToString
                srttp = ddlshortbasis.SelectedValue.ToString
                If srttp.ToString <> "" Then
                    Left(srttp, Len(srttp) - 1)
                End If
                For Each lstitem1 In cbldeduction.Items
                    If lstitem1.Selected = True Then
                        paycodeSel1 = paycodeSel1 + lstitem1.Value + ","
                    End If
                Next
                For Each lstitem In cbladd.Items
                    If lstitem.Selected = True Then
                        paycodeSel = paycodeSel + lstitem.Value + ","
                    End If
                Next
                PayCodeId = paycodeSel1 + paycodeSel

                If Right(Ids, 1) = "," Then
                    Ids = Left(Ids, Len(Ids) - 1).ToString
                End If
                If Right(PayCodeId, 1) = "," Then
                    PayCodeId = Left(PayCodeId, Len(PayCodeId) - 1).ToString
                End If

                HidPreVal.Value = "H~" & DDLPaySlipType.SelectedValue & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" & Level.ToString & "~" & CC.ToString _
                & "~" & Loc.ToString & "~" & unit.ToString & "~" & SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" _
                & Month.ToString & "~" & Year.ToString & "~" & Hold.ToString & "~" & EmpType.ToString & "~" & srttp.ToString & "~" & Ids.ToString & "~" & PayCodeId.ToString.Trim
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'Added by Jay on 11 Mar 2014 End


                'This statement is used for display "Year To Date Salary Slip"
            ElseIf DDLPaySlipType.SelectedValue = "20" Then
                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString _
                & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & Hold.ToString & "~" & EmpType.ToString & "~" _
                & DDLPaySlipType.SelectedValue.ToString & "~" & "" & "~~~~"

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'This code is used to display "Monthly Salary Register New" 

            ElseIf DDLPaySlipType.SelectedValue = "22" Then
                Dim RecNo As String = ""
                For Each lstitem1 In cbldeduction.Items
                    If lstitem1.Selected = True Then
                        paycodeSel1 = paycodeSel1 + lstitem1.Value + ","
                    End If
                Next
                'here we store all selected list item in the paycodesel variable.
                For Each lstitem In cbladd.Items
                    If lstitem.Selected = True Then
                        paycodeSel = paycodeSel + lstitem.Value + ","
                    End If
                Next
                'here we add all paycode in hidden control
                PayCodeId = paycodeSel1 + paycodeSel
                Hold = ddlshowsal.SelectedValue.ToString

                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
               Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
               SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
               EmpLName.ToString & "~" & Month.ToString & "~" & Year.ToString & "~" & Hold.ToString & "~" _
               & PayCodeId.ToString & "~" & DDLPaySlipType.SelectedValue.ToString & "~" & EmpType.ToString

                Session(_strVal) = HidPreVal.Value.ToString
                HidPreVal.Value = _strVal.ToString

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
            ElseIf DDLPaySlipType.SelectedValue.ToString = "23" Then
                Ent = 1
                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & EmpType.ToString & "~" & empcode.ToString & "~" &
                EmpFName.ToString & "~" & EmpLName.ToString & "~" & Month.ToString & "~" & Year.ToString &
                "~" & Hold.ToString & "~" & Ent.ToString & "~" & DDLPaySlipType.SelectedValue.ToString & "~~"
                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'This statement is used for display "Year To Date Salary Slip(Nigeria)"
            ElseIf DDLPaySlipType.SelectedValue = "24" Then

                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString _
                & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & Hold.ToString & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'This statement is used for display "Monthly salary register with seperately reimbursement"
            ElseIf DDLPaySlipType.SelectedValue = "25" Then
                For Each lstitem In cblReimb.Items
                    If lstitem.Selected = True Then
                        paycodeSel = paycodeSel + lstitem.Value + ","
                    End If
                Next
                'here we add all paycode in hidden control
                PayCodeId = paycodeSel
                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                 Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                 SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString _
                 & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & Hold.ToString & "~" &
                 EmpType.ToString & "~" & PayCodeId.ToString & "~" & DDLPaySlipType.SelectedValue.ToString
                'IIf(ChkStaffID.Checked = True, "Y", "N")

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                '-----------------------------------------------------------------------------------
                'This statement is used for display "Monthly Salary Slip With Tax Details"
            ElseIf DDLPaySlipType.SelectedValue = "26" Then
                Dim LeaveBal As String = "", Flag As String = "S" _
                , LoanBal As String = ""

                'Add on blank flag after "StaffID" for employee password
                HidPreVal.Value = "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString _
                & "~" & Year.ToString & "~" & EmpLName.ToString & "~" &
                 "~" & "~" & Comzero.ToString & "~" &
                Comlogo.ToString & "~" & Hold.ToString & "~" &
                Flag.ToString & "~" & EmpType.ToString _
                & "~" & "~" & DDLPaySlipType.SelectedValue.ToString
                '& ShowPDate.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
            ElseIf DDLPaySlipType.SelectedValue = "29" Then
                countemp = "0"
                'here we store all selected list item in the paycodesel1 variable.
                For Each lstitem1 In cbldeduction.Items
                    If lstitem1.Selected = True Then
                        paycodeSel1 = paycodeSel1 + lstitem1.Value + ","
                    End If
                Next
                'here we store all selected list item in the paycodesel variable.
                For Each lstitem In cbladd.Items
                    If lstitem.Selected = True Then
                        paycodeSel = paycodeSel + lstitem.Value + ","
                    End If
                Next
                'here we add all paycode in hidden control
                PayCodeId = paycodeSel1 + paycodeSel

                arrp(0) = New SqlParameter("@CC", USearch.UCddlcostcenter.ToString())
                arrp(1) = New SqlParameter("@Loc", USearch.UCddllocation.ToString())
                arrp(2) = New SqlParameter("@unit", USearch.UCddlunit.ToString())
                arrp(3) = New SqlParameter("@salbase", USearch.UCddlsalbasis.ToString())
                arrp(4) = New SqlParameter("@Empcode", USearch.UCTextcode.ToString)
                arrp(5) = New SqlParameter("@Dep", USearch.UCddldept.ToString())
                arrp(6) = New SqlParameter("@Desig", USearch.UCddldesig.ToString())
                arrp(7) = New SqlParameter("@Grad", USearch.UCddlgrade.ToString())
                arrp(8) = New SqlParameter("@Lable", USearch.UCddllevel.ToString())
                arrp(9) = New SqlClient.SqlParameter("@month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
                arrp(10) = New SqlClient.SqlParameter("@year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                arrp(11) = New SqlParameter("@EmpFName", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
                arrp(12) = New SqlParameter("@EmpLName", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                arrp(13) = New SqlParameter("@Hold", ddlshowsal.SelectedValue.ToString)
                arrp(14) = New SqlParameter("@UserGroup", Session("Ugroup").ToString)
                arrp(15) = New SqlParameter("@EmpStatus", USearch.UCddlEmp.ToString)
                arrp(16) = New SqlClient.SqlParameter("@Paycode", PayCodeId)
                arrp(17) = New SqlClient.SqlParameter("@Bas", SqlDbType.Char, 1)
                arrp(17).Direction = ParameterDirection.Output
                arrp(18) = New SqlClient.SqlParameter("@Hra", SqlDbType.Char, 1)
                arrp(19) = New SqlClient.SqlParameter("@RepID", DDLPaySlipType.SelectedValue.ToString)
                arrp(18).Direction = ParameterDirection.Output
                'Added by Nisha(04 Feb 2013) Fin Year Change
                arrp(20) = New SqlParameter("@userid", Session("uid").ToString)
                ds = _ObjData.GetDataSetProc("Sp_Rpt_Sel_FactoryRegister", arrp)
                If arrp(17).Value.ToString = "Y" And arrp(18).Value.ToString = "Y" Then
                    If ds.Tables(2).Rows.Count > 0 Then
                        'Geeta2012
                        Dim _sw1 As StreamWriter, filename1 As String = "", complexID As Guid = Guid.NewGuid()
                        filename1 = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\Salary Register" & "_" & Right(complexID.ToString, 6) & ".xls"
                        If System.IO.File.Exists(filename1) Then
                            System.IO.File.Delete(filename1)
                        End If
                        _sw1 = New StreamWriter(filename1)
                        ExportToExcel(ds, _sw1)
                        _sw1.Close()
                        _sw1.Dispose()

                        lblmsg.Text = ""
                        Response.Clear()

                        Response.ContentType = "application/zip"
                        Response.AddHeader("content-disposition", "filename=Salary Register.zip")
                        Using zip As New ZipFile()
                            zip.AddFile(filename1, "Salary Register")
                            zip.Save(Response.OutputStream)
                        End Using

                        If File.Exists(filename1) Then
                            File.Delete(filename1)
                        End If

                        Response.End()
                    Else
                        _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !", False)
                    End If
                Else
                    _objCommon.ShowMessage("M", lblmsg, "Basic Or HRA are not mapped !", False)
                End If
                ''For other format
            ElseIf (DDLPaySlipType.SelectedValue = "30" Or DDLPaySlipType.SelectedValue = "31") Then
                countemp = "0"
                'here we store all selected list item in the paycodesel1 variable.
                For Each lstitem1 In cbldeduction.Items
                    If lstitem1.Selected = True Then
                        paycodeSel1 = paycodeSel1 + lstitem1.Value + ","
                    End If
                Next
                'here we store all selected list item in the paycodesel variable.
                For Each lstitem In cbladd.Items
                    If lstitem.Selected = True Then
                        paycodeSel = paycodeSel + lstitem.Value + ","
                    End If
                Next
                'here we add all paycode in hidden control
                PayCodeId = paycodeSel1 + paycodeSel

                arrpam(0) = New SqlParameter("@CC", USearch.UCddlcostcenter.ToString())
                arrpam(1) = New SqlParameter("@Loc", USearch.UCddllocation.ToString())
                arrpam(2) = New SqlParameter("@unit", USearch.UCddlunit.ToString())
                arrpam(3) = New SqlParameter("@salbase", USearch.UCddlsalbasis.ToString())
                arrpam(4) = New SqlParameter("@Empcode", USearch.UCTextcode.ToString)
                arrpam(5) = New SqlParameter("@Dep", USearch.UCddldept.ToString())
                arrpam(6) = New SqlParameter("@Desig", USearch.UCddldesig.ToString())
                arrpam(7) = New SqlParameter("@Grad", USearch.UCddlgrade.ToString())
                arrpam(8) = New SqlParameter("@Lable", USearch.UCddllevel.ToString())
                arrpam(9) = New SqlClient.SqlParameter("@month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
                arrpam(10) = New SqlClient.SqlParameter("@year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                arrpam(11) = New SqlParameter("@EmpFName", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
                arrpam(12) = New SqlParameter("@EmpLName", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                arrpam(13) = New SqlParameter("@Hold", ddlshowsal.SelectedValue.ToString)
                arrpam(14) = New SqlParameter("@UserGroup", Session("Ugroup").ToString)
                arrpam(15) = New SqlParameter("@EmpStatus", USearch.UCddlEmp.ToString)
                arrpam(16) = New SqlClient.SqlParameter("@Paycode", PayCodeId)
                arrpam(17) = New SqlClient.SqlParameter("@RepID", DDLPaySlipType.SelectedValue.ToString)
                'Added by Nisha(04 Feb 2013) Fin Year Change
                arrpam(18) = New SqlClient.SqlParameter("@userid", Session("uid").ToString)
                If DDLPaySlipType.SelectedValue.ToString = "30" Then
                    ds = _ObjData.GetDataSetProc("Sp_Rpt_Sel_FactoryRegister_Staff", arrpam)
                    If ds.Tables(1).Rows.Count > 0 Then
                        'Geeta2012
                        Dim _sw1 As StreamWriter, filename1 As String = "", complexID As Guid = Guid.NewGuid()
                        filename1 = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\Salary Register" & "_" & Right(complexID.ToString, 6) & ".xls"
                        If System.IO.File.Exists(filename1) Then
                            System.IO.File.Delete(filename1)
                        End If
                        _sw1 = New StreamWriter(filename1)
                        ExportToExcelFoStaff(ds, _sw1)
                        _sw1.Close()
                        _sw1.Dispose()


                        lblmsg.Text = ""
                        Response.Clear()

                        Response.ContentType = "application/zip"
                        Response.AddHeader("content-disposition", "filename=Salary Register.zip")
                        Using zip As New ZipFile()
                            zip.AddFile(filename1, "Salary Register")
                            zip.Save(Response.OutputStream)
                        End Using

                        If File.Exists(filename1) Then
                            File.Delete(filename1)
                        End If
                        Response.End()
                    Else
                        _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !", False)
                    End If
                Else
                    ds = _ObjData.GetDataSetProc("Sp_Rpt_Sel_FactoryRegister_Worker", arrpam)
                    If ds.Tables.Count > 1 Then
                        If ds.Tables(1).Rows.Count > 0 Then
                            'Geeta2012
                            Dim _sw1 As StreamWriter, filename1 As String = "", complexID As Guid = Guid.NewGuid()
                            filename1 = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\Salary Register" & "_" & Right(complexID.ToString, 6) & ".xls"
                            If System.IO.File.Exists(filename1) Then
                                System.IO.File.Delete(filename1)
                            End If
                            _sw1 = New StreamWriter(filename1)
                            ExportToExcelFoWorker(ds, _sw1)
                            _sw1.Close()
                            _sw1.Dispose()


                            lblmsg.Text = ""
                            Response.Clear()

                            Response.ContentType = "application/zip"
                            Response.AddHeader("content-disposition", "filename=Salary Register.zip")
                            Using zip As New ZipFile()
                                zip.AddFile(filename1, "Salary Register")
                                zip.Save(Response.OutputStream)
                            End Using

                            If File.Exists(filename1) Then
                                File.Delete(filename1)
                            End If
                            Response.End()
                        Else
                            _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !", False)
                        End If
                    Else
                        _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !", False)
                    End If
                End If
                'Added By geeta on 1 Jun 2012
                'Added By geeta on 18 May 2012
            ElseIf DDLPaySlipType.SelectedValue = "32" Then
                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                Month.ToString & "~" & Year.ToString & "~" & EmpLName.ToString &
                "~" & Hold.ToString & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString
                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                '_str.Append("<script language=javascript>")
                '_str.Append("blankcheck1();")
                '_str.Append("</script>")
                'lit.Text = _str.ToString
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'Added by geeta 21 May
            ElseIf DDLPaySlipType.SelectedValue = "33" Then

                HidPreVal.Value = CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & Month.ToString & "~" & Year.ToString & "~" &
                EmpFName.ToString & "~" & EmpLName.ToString & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString
                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                '_str.Append("<script language=javascript>")
                '_str.Append("blankcheck1();")
                '_str.Append("</script>")
                'lit.Text = _str.ToString
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'Added by geeta 21 May
            ElseIf DDLPaySlipType.SelectedValue = "34" Then
                HidPreVal.Value = Month.ToString & "~" & Year.ToString & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
               Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
               SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                EmpLName.ToString & "~" & Session("UGroup") & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString
                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                '_str.Append("<script language=javascript>")
                '_str.Append("blankcheck1();")
                '_str.Append("</script>")
                'lit.Text = _str.ToString
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)

            ElseIf DDLPaySlipType.SelectedValue = "35" Then
                Dim _arrpam(16) As SqlClient.SqlParameter, _ds As New DataSet
                _arrpam(0) = New SqlParameter("@fk_costcenter_code", USearch.UCddlcostcenter.ToString())
                _arrpam(1) = New SqlParameter("@fk_loc_code", USearch.UCddllocation.ToString())
                _arrpam(2) = New SqlParameter("@Fk_unit", USearch.UCddlunit.ToString())
                _arrpam(3) = New SqlParameter("@salaried", USearch.UCddlsalbasis.ToString())
                _arrpam(4) = New SqlParameter("@pk_emp_code", USearch.UCTextcode.ToString)
                _arrpam(5) = New SqlParameter("@fk_dept_code", USearch.UCddldept.ToString())
                _arrpam(6) = New SqlParameter("@fk_desig_code", USearch.UCddldesig.ToString())
                _arrpam(7) = New SqlParameter("@fk_grade_code", USearch.UCddlgrade.ToString())
                _arrpam(8) = New SqlParameter("@fk_level_Code", USearch.UCddllevel.ToString())
                _arrpam(9) = New SqlClient.SqlParameter("@month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
                _arrpam(10) = New SqlClient.SqlParameter("@year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                _arrpam(11) = New SqlParameter("@first_name", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
                _arrpam(12) = New SqlParameter("@last_name", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                _arrpam(13) = New SqlParameter("@EmpStatus", USearch.UCddlEmp.ToString)
                _arrpam(14) = New SqlParameter("@UserGroup", Session("Ugroup").ToString)
                _arrpam(15) = New SqlParameter("@RepId", DDLPaySlipType.SelectedValue.ToString)
                'Added by Nisha(04 Feb 2013) Fin Year Change
                _arrpam(16) = New SqlParameter("@userid", Session("uid").ToString)
                _ds = _ObjData.GetDataSetProc("Paysp_Mstemployee_ApprenticesRegister", _arrpam)
                If _ds.Tables(0).Rows.Count > 1 Then
                    Response.Clear()
                    'Response.AddHeader("content-disposition", "attachment; filename=SalaryRegister.xls")
                    'Response.Charset = ""
                    'Response.ContentType = "application/vnd.xls"
                    Dim filepath As String = "", _FileName As String = "SalaryRegister_" & Left(Guid.NewGuid.ToString(), 5)
                    filepath = _ExcelFilebyxml.ExportToExcelXML_ByDataTable(_ds.Tables(0), "Apprentice Pay - Register for the month of " & ddlMonthYear.SelectedItem.ToString & "                Date : " & DateTime.Now.ToString("dd MMMM yyyy") & " Time : " & DateTime.Now.ToString("HH:mm"), , , _FileName)
                    'HttpContext.Current.Response.WriteFile(filepath.ToString)
                    'Response.Clear()
                    'Response.End()

                    Response.ContentType = "application/zip"
                    Response.AddHeader("content-disposition", "filename=" & _FileName & ".zip")
                    Using zip As New ZipFile()
                        zip.AddFile(filepath.ToString, _FileName)
                        zip.Save(Response.OutputStream)
                    End Using

                    If File.Exists(filepath.ToString) Then
                        File.Delete(filepath.ToString)
                    End If
                    Response.End()
                    Response.Clear()
                Else
                    _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !", False)
                End If
            ElseIf DDLPaySlipType.SelectedValue = "36" Then
                Dim _arrpams(16) As SqlClient.SqlParameter, _ds As New DataSet
                _arrpams(0) = New SqlClient.SqlParameter("@Month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
                _arrpams(1) = New SqlClient.SqlParameter("@Year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                _arrpams(2) = New SqlParameter("@Dep", USearch.UCddldept.ToString())
                _arrpams(3) = New SqlParameter("@Desig", USearch.UCddldesig.ToString())
                _arrpams(4) = New SqlParameter("@Grad", USearch.UCddlgrade.ToString())
                _arrpams(5) = New SqlParameter("@Lable", USearch.UCddllevel.ToString())
                _arrpams(6) = New SqlParameter("@CC", USearch.UCddlcostcenter.ToString())
                _arrpams(7) = New SqlParameter("@Loc", USearch.UCddllocation.ToString())
                _arrpams(8) = New SqlParameter("@unit", USearch.UCddlunit.ToString())
                _arrpams(9) = New SqlParameter("@SalBase", USearch.UCddlsalbasis.ToString())
                _arrpams(10) = New SqlParameter("@EmpCode", empcode.ToString())
                _arrpams(11) = New SqlParameter("@EmpFName", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
                _arrpams(12) = New SqlParameter("@EmpLName", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                _arrpams(13) = New SqlParameter("@userGroup", Session("Ugroup").ToString)
                _arrpams(14) = New SqlParameter("@EmpType", USearch.UCddlEmp.ToString)
                _arrpams(15) = New SqlParameter("@RepId", DDLPaySlipType.SelectedValue.ToString)
                'Added by Nisha(04 Feb 2013) Fin Year Change
                _arrpams(16) = New SqlParameter("@userid", Session("uid").ToString)
                _ds = _ObjData.GetDataSetProc("Paysp_CasualRegister", _arrpams)
                If _ds.Tables(0).Rows.Count > 1 Then
                    Response.Clear()
                    ''This is for add the line header on the excel sheet
                    'Response.AddHeader("content-disposition", "attachment; filename=SalaryRegister.xls")
                    'Response.Charset = ""
                    'Response.ContentType = "application/vnd.xls"
                    ''stringWrite = ExportToExcelFoWorker(ds)
                    ''Response.Write(stringWrite.ToString())
                    Dim filepath As String = "", _FileName As String = "SalaryRegister_" & Left(Guid.NewGuid.ToString(), 5)
                    filepath = _ExcelFilebyxml.ExportToExcelXML_ByDataTable(_ds.Tables(0), "Casual Pay - Register for the month of " & ddlMonthYear.SelectedItem.ToString & "                Date : " & DateTime.Now.ToString("dd MMMM yyyy") & " Time : " & DateTime.Now.ToString("HH:mm"), , , _FileName)
                    'HttpContext.Current.Response.WriteFile(filepath.ToString)
                    'Response.End()

                    Response.ContentType = "application/zip"
                    Response.AddHeader("content-disposition", "filename=" & _FileName & ".zip")
                    Using zip As New ZipFile()
                        zip.AddFile(filepath.ToString, _FileName)
                        zip.Save(Response.OutputStream)
                    End Using

                    If File.Exists(filepath.ToString) Then
                        File.Delete(filepath.ToString)
                    End If
                    Response.End()
                    Response.Clear()
                Else
                    _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !", False)
                End If
                ''Added by Sushil on 10 Oct 2012
            ElseIf DDLPaySlipType.SelectedValue = "37" Then
                HidPreVal.Value = "K~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                                Month.ToString & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & Salut.ToString & "~" _
                                & chkPF_Old.ToString & "~" & chkESI_Old.ToString & "~" & Comzero.ToString & "~" &
                                Comlogo.ToString & "~" & Hold.ToString & "~" & Loan.ToString & "~" _
                                & Advance.ToString & "~" & "~S" & "~" & StaffId.ToString & "~~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)

            ElseIf DDLPaySlipType.SelectedValue = "43" Then
                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                                Month.ToString & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue & "~H~0"

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If

                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)

                'Added by Mohan on 20 Feb 13
            ElseIf DDLPaySlipType.SelectedValue = "39" Then
                countemp = "0"
                'here we store all selected list item in the paycodesel1 variable.
                For Each lstitem1 In cbldeduction.Items
                    If lstitem1.Selected = True Then
                        paycodeSel1 = paycodeSel1 + lstitem1.Value + ","
                    End If
                Next
                'here we store all selected list item in the paycodesel variable.
                For Each lstitem In cbladd.Items
                    If lstitem.Selected = True Then
                        paycodeSel = paycodeSel + lstitem.Value + ","
                    End If
                Next
                'here we add all paycode in hidden control
                PayCodeId = paycodeSel1 + paycodeSel

                arrpamm(0) = New SqlParameter("@CC", USearch.UCddlcostcenter.ToString())
                arrpamm(1) = New SqlParameter("@Loc", USearch.UCddllocation.ToString())
                arrpamm(2) = New SqlParameter("@unit", USearch.UCddlunit.ToString())
                arrpamm(3) = New SqlParameter("@salbase", USearch.UCddlsalbasis.ToString())
                arrpamm(4) = New SqlParameter("@Empcode", USearch.UCTextcode.ToString)
                arrpamm(5) = New SqlParameter("@Dep", USearch.UCddldept.ToString())
                arrpamm(6) = New SqlParameter("@Desig", USearch.UCddldesig.ToString())
                arrpamm(7) = New SqlParameter("@Grad", USearch.UCddlgrade.ToString())
                arrpamm(8) = New SqlParameter("@Lable", USearch.UCddllevel.ToString())
                arrpamm(9) = New SqlClient.SqlParameter("@month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
                arrpamm(10) = New SqlClient.SqlParameter("@year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                arrpamm(11) = New SqlParameter("@EmpFName", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
                arrpamm(12) = New SqlParameter("@EmpLName", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                arrpamm(13) = New SqlParameter("@Hold", ddlshowsal.SelectedValue.ToString)
                arrpamm(14) = New SqlParameter("@UserGroup", Session("Ugroup").ToString)
                arrpamm(15) = New SqlParameter("@EmpStatus", USearch.UCddlEmp.ToString)
                arrpamm(16) = New SqlClient.SqlParameter("@Paycode", PayCodeId)
                arrpamm(17) = New SqlClient.SqlParameter("@RepID", DDLPaySlipType.SelectedValue.ToString)
                arrpamm(18) = New SqlClient.SqlParameter("@Flag", SqlDbType.Char, 1)
                arrpamm(18).Direction = ParameterDirection.Output
                ds = _ObjData.GetDataSetProc("PaySp_Rpt_Sel_FacReg_FORMXVII", arrpamm)
                If DDLPaySlipType.SelectedValue.ToString = "39" Then
                    If ds.Tables(1).Rows.Count > 0 Then
                        If arrpamm(18).Value.ToString.ToUpper = "N" Then
                            lblmsg2.Text = "Paycode named as Arrear already exists.Please change the name of Arrear paycode!"
                        Else
                            Dim _sw1 As StreamWriter, filename1 As String = "", complexID As Guid = Guid.NewGuid()
                            filename1 = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\MusterRollRegister" & "_" & Right(complexID.ToString, 6) & ".xls"
                            If System.IO.File.Exists(filename1) Then
                                System.IO.File.Delete(filename1)
                            End If
                            _sw1 = New StreamWriter(filename1)
                            ExportToExcelForReg(ds, _sw1)
                            _sw1.Close()
                            _sw1.Dispose()


                            lblmsg.Text = ""
                            Response.Clear()

                            Response.ContentType = "application/zip"
                            Response.AddHeader("content-disposition", "filename=MusterRollRegister.zip")
                            Using zip As New ZipFile()
                                zip.AddFile(filename1, "MusterRollRegister")
                                zip.Save(Response.OutputStream)
                            End Using

                            If File.Exists(filename1) Then
                                File.Delete(filename1)
                            End If
                            Response.End()
                        End If

                    Else
                        _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !", False)
                    End If

                End If
                'Added by geeta PTC
            ElseIf DDLPaySlipType.SelectedValue = "53" Then

                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString _
                & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & Hold.ToString & "~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString & "~~~~~~~~~"

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                ''For other Format
            ElseIf DDLPaySlipType.SelectedValue = "40" Then
                countemp = "0"
                'here we store all selected list item in the paycodesel1 variable.
                For Each lstitem1 In cbldeduction.Items
                    If lstitem1.Selected = True Then
                        paycodeSel1 = paycodeSel1 + lstitem1.Value + ","
                    End If
                Next
                'here we store all selected list item in the paycodesel variable.
                For Each lstitem In cbladd.Items
                    If lstitem.Selected = True Then
                        paycodeSel = paycodeSel + lstitem.Value + ","
                    End If
                Next
                'here we add all paycode in hidden control
                PayCodeId = paycodeSel1 + paycodeSel

                arrpamwork(0) = New SqlParameter("@CC", USearch.UCddlcostcenter.ToString())
                arrpamwork(1) = New SqlParameter("@Loc", USearch.UCddllocation.ToString())
                arrpamwork(2) = New SqlParameter("@unit", USearch.UCddlunit.ToString())
                arrpamwork(3) = New SqlParameter("@salbase", USearch.UCddlsalbasis.ToString())
                arrpamwork(4) = New SqlParameter("@Empcode", USearch.UCTextcode.ToString)
                arrpamwork(5) = New SqlParameter("@Dep", USearch.UCddldept.ToString())
                arrpamwork(6) = New SqlParameter("@Desig", USearch.UCddldesig.ToString())
                arrpamwork(7) = New SqlParameter("@Grad", USearch.UCddlgrade.ToString())
                arrpamwork(8) = New SqlParameter("@Lable", USearch.UCddllevel.ToString())
                arrpamwork(9) = New SqlParameter("@EmpFName", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
                arrpamwork(10) = New SqlParameter("@EmpLName", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                arrpamwork(11) = New SqlParameter("@UserGroup", Session("Ugroup").ToString)
                arrpamwork(12) = New SqlParameter("@EmpStatus", USearch.UCddlEmp.ToString)

                ds = _ObjData.GetDataSetProc("PaySp_Rpt_FacRegWorkMan", arrpamwork)
                If DDLPaySlipType.SelectedValue.ToString = "40" Then
                    If ds.Tables(2).Rows.Count > 0 Then
                        'Geeta2012
                        Dim _sw1 As StreamWriter, filename1 As String = "", complexID As Guid = Guid.NewGuid()
                        filename1 = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\WorkmenRegister" & "_" & Right(complexID.ToString, 6) & ".xls"
                        If System.IO.File.Exists(filename1) Then
                            System.IO.File.Delete(filename1)
                        End If
                        _sw1 = New StreamWriter(filename1)
                        ExportToExcelworkman(ds, _sw1)
                        _sw1.Close()
                        _sw1.Dispose()


                        lblmsg.Text = ""
                        Response.Clear()

                        Response.ContentType = "application/zip"
                        Response.AddHeader("content-disposition", "filename=WorkmenRegister.zip")
                        Using zip As New ZipFile()
                            zip.AddFile(filename1, "WorkmenRegister")
                            zip.Save(Response.OutputStream)
                        End Using

                        If File.Exists(filename1) Then
                            File.Delete(filename1)
                        End If
                        Response.End()
                    Else
                        _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !", False)
                    End If
                Else
                    _objCommon.ShowMessage("M", lblmsg, "Basic Or HRA are not mapped !", False)
                End If
                'Added this condition by Rohtas Singh on 08 Dec 2017 for "Monthly Salary Slip (MAX Life)"
            ElseIf DDLPaySlipType.SelectedValue = "54" Then
                arrparam(0) = New SqlParameter("@fk_costcenter_code", USearch.UCddlcostcenter.ToString())
                arrparam(1) = New SqlParameter("@fk_loc_code", USearch.UCddllocation.ToString())
                arrparam(2) = New SqlParameter("@fk_unit", USearch.UCddlunit.ToString())
                arrparam(3) = New SqlParameter("@salaried", USearch.UCddlsalbasis.ToString())
                arrparam(4) = New SqlParameter("@pk_emp_code", USearch.UCTextcode.ToString)
                arrparam(5) = New SqlParameter("@fk_dept_code", USearch.UCddldept.ToString())
                arrparam(6) = New SqlParameter("@fk_desig_code", USearch.UCddldesig.ToString())
                arrparam(7) = New SqlParameter("@fk_grade_code", USearch.UCddlgrade.ToString())
                arrparam(8) = New SqlParameter("@fk_level_Code", USearch.UCddllevel.ToString())
                arrparam(9) = New SqlClient.SqlParameter("@month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
                arrparam(10) = New SqlClient.SqlParameter("@year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                arrparam(11) = New SqlParameter("@first_name", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
                arrparam(12) = New SqlParameter("@last_name", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                arrparam(13) = New SqlParameter("@Hold", ddlshowsal.SelectedValue.ToString)
                arrparam(14) = New SqlParameter("@UserGroup", Session("Ugroup").ToString)
                arrparam(15) = New SqlParameter("@EmpStatus", USearch.UCddlEmp.ToString)
                arrparam(16) = New SqlParameter("@RepID", DDLPaySlipType.SelectedValue.ToString)

                ds = _ObjData.GetDataSetProc("PaySP_SalarySlipFor_CSV_Excel", arrparam)
                If ds.Tables(0).Rows.Count > 0 Then
                    If ddlRptFormat.SelectedValue.ToString = "1" Then
                        Export_CSV(ds.Tables(0))
                    Else
                        Dim filename As String = "", _sw As StreamWriter, complexID As Guid = Guid.NewGuid()

                        filename = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\MHC_PAYSLIP" & "_" & Right(complexID.ToString, 6) & ".xls"
                        If System.IO.File.Exists(filename) Then
                            System.IO.File.Delete(filename)
                        End If
                        _sw = New StreamWriter(filename)
                        ExportToExcelXML(ds, _sw)
                        _sw.Close()
                        _sw.Dispose()

                        Response.Clear()
                        Response.BufferOutput = False

                        Response.ContentType = "application/zip"
                        Response.AddHeader("content-disposition", "filename=MHC_PAYSLIP.zip")
                        Using zip As New ZipFile()
                            zip.AddFile(filename, "MHC_PAYSLIP")
                            zip.Save(Response.OutputStream)
                        End Using

                        If File.Exists(filename) Then
                            File.Delete(filename)
                        End If

                        lblmsg.Text = ""

                        Response.End()
                    End If
                Else
                    _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !", False)
                End If
                ds.Clear()
                ds.Dispose()
            ElseIf DDLPaySlipType.SelectedValue = "55" Then
                HidPreVal.Value = DDLPaySlipType.SelectedValue.Trim & "~H1~" & empcode.ToString & "~" & EmpFName.ToString & "~" & EmpLName.ToString & "~" &
                                CC.ToString & "~" & Dep.ToString & "~" & Grad.ToString & "~" & Desig.ToString & "~" &
                               Loc.ToString & "~" & unit.ToString & "~" & SalBase.ToString & "~" &
                                Level.ToString & "~" & EmpType.ToString & "~" & Session("ugroup").ToString & "~" & Month.ToString & "~" _
                                & Year.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
            ElseIf DDLPaySlipType.SelectedValue = "57" Then
                HidPreVal.Value = DDLPaySlipType.SelectedValue.Trim & "~H1~" & empcode.ToString & "~" & EmpFName.ToString & "~" & EmpLName.ToString & "~" &
                                CC.ToString & "~" & Dep.ToString & "~" & Grad.ToString & "~" & Desig.ToString & "~" &
                               Loc.ToString & "~" & unit.ToString & "~" & SalBase.ToString & "~" &
                                Level.ToString & "~" & EmpType.ToString & "~" & Session("ugroup").ToString & "~" & Month.ToString & "~" _
                                & Year.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                ''Rep ID 56 Pay slip with reimbursement details
            ElseIf DDLPaySlipType.SelectedValue = "56" Then
                HidPreVal.Value = "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString &
                "~" & Year.ToString & "~" & EmpLName.ToString & "~" &
                Comzero.ToString & "~" &
                Comlogo.ToString & "~" & Hold.ToString _
                & "~~" & Comzero.ToString & "~" & Comlogo.ToString & "~S" _
                & "~~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString _
                & "~~~~" & "N" 'Added by geeta on 11 sep 2012

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'Added by geeta for time card 
                'added by Ritu Malik : Trainee stipend Salary Slip Equals("59")
                'added by Geeta : Marathi payslip("60")
            ElseIf DDLPaySlipType.SelectedValue = "58" Or DDLPaySlipType.SelectedValue.Equals("59") Or DDLPaySlipType.SelectedValue.Equals("60") Then
                HidPreVal.Value = "K~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                                Month.ToString & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & Salut.ToString & "~" _
                                & chkPF_Old.ToString & "~" & chkESI_Old.ToString & "~" & Comzero.ToString & "~" &
                                Comlogo.ToString & "~" & Hold.ToString & "~" & Loan.ToString & "~" _
                                & Advance.ToString & "~" & "~S" & "~" & StaffId.ToString & "~~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
            ElseIf DDLPaySlipType.SelectedValue.Trim.Equals("61") Then
                ReDim arrparam(13)
                arrparam(0) = New SqlParameter("@fk_costcenter_code", USearch.UCddlcostcenter.ToString())
                arrparam(1) = New SqlParameter("@fk_loc_code", USearch.UCddllocation.ToString())
                arrparam(2) = New SqlParameter("@fk_unit", USearch.UCddlunit.ToString())
                arrparam(3) = New SqlParameter("@salaried", USearch.UCddlsalbasis.ToString())
                arrparam(4) = New SqlParameter("@pk_emp_code", USearch.UCTextcode.ToString)
                arrparam(5) = New SqlParameter("@fk_dept_code", USearch.UCddldept.ToString())
                arrparam(6) = New SqlParameter("@fk_desig_code", USearch.UCddldesig.ToString())
                arrparam(7) = New SqlParameter("@fk_grade_code", USearch.UCddlgrade.ToString())
                arrparam(8) = New SqlParameter("@fk_level_Code", USearch.UCddllevel.ToString())
                arrparam(9) = New SqlParameter("@last_name", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                arrparam(10) = New SqlClient.SqlParameter("@month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
                arrparam(11) = New SqlClient.SqlParameter("@year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                arrparam(12) = New SqlParameter("@EmpStatus", USearch.UCddlEmp.ToString)
                arrparam(13) = New SqlParameter("@UserGroup", Session("Ugroup").ToString)
                ds = _ObjData.GetDataSetProc("PaySP_TrnEmpPaidHolidaysProcessArr", arrparam)
                If ds.Tables(0).Rows.Count > 0 Then
                    Dim filename As String = "", _sw As StreamWriter, complexID As Guid = Guid.NewGuid()

                    filename = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\paidHoliday" & "_" & Right(complexID.ToString, 6) & ".xls"
                    If System.IO.File.Exists(filename) Then
                        System.IO.File.Delete(filename)
                    End If
                    _sw = New StreamWriter(filename)
                    ExportToExcelXML(ds, _sw)
                    _sw.Close()
                    _sw.Dispose()

                    Response.Clear()
                    Response.BufferOutput = False

                    'Dim ReadmeText As [String] = "Hello!" & vbLf & vbLf & "This is a README..." & DateTime.Now.ToString("G")
                    Response.ContentType = "application/zip"
                    Response.AddHeader("content-disposition", "filename=paidHoliday.zip")
                    Using zip As New ZipFile()
                        zip.AddFile(filename, "paidHoliday")
                        zip.Save(Response.OutputStream)
                    End Using

                    If File.Exists(filename) Then
                        File.Delete(filename)
                    End If

                    lblmsg.Text = ""
                    Response.End()
                Else
                    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "No record(s) Found according to the selection criteria !"})
                    _objCommon.ShowMessage(_msg)
                End If
            ElseIf DDLPaySlipType.SelectedValue.Equals("62") Then
                Fromdate = Convert.ToString(ddlmonthyearS.SelectedValue)

                HidPreVal.Value = "H~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
            Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
            SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString _
            & "~" & Year.ToString & "~" & EmpLName.ToString & "~" & "~" & EmpType.ToString & "~" & Fromdate.ToString & "~0~" & DDLPaySlipType.SelectedValue
                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
            ElseIf DDLPaySlipType.SelectedValue.Equals("63") Then
                HidPreVal.Value = "~" & DDLPaySlipType.SelectedValue.ToString & "~" & empcode.ToString & "~" & Month.ToString & "~" & Year.ToString & "~" &
                CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" & Level.ToString & "~" &
                EmpFName.ToString & "~" & EmpLName.ToString & "~" & SalBase.ToString & "~" & EmpType.ToString & "~"

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
            ElseIf DDLPaySlipType.SelectedValue.Equals("64") Then
                HidPreVal.Value = "~" & DDLPaySlipType.SelectedValue.ToString & "~" & empcode.ToString & "~" & Month.ToString & "~" & Year.ToString & "~" &
                CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" & Level.ToString & "~" &
                EmpFName.ToString & "~" & EmpLName.ToString & "~" & SalBase.ToString & "~" & EmpType.ToString & "~" & Comzero.ToString & "~" & Hold.ToString & "~"

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
            ElseIf DDLPaySlipType.SelectedValue.Equals("65") Then
                HidPreVal.Value = "~" & DDLPaySlipType.SelectedValue.ToString & "~" & empcode.ToString & "~" & Month.ToString & "~" & Year.ToString & "~" &
                CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" & Level.ToString & "~" &
                EmpFName.ToString & "~" & EmpLName.ToString & "~" & SalBase.ToString & "~" & EmpType.ToString & "~" & Comzero.ToString & "~" & Hold.ToString & "~"

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
            ElseIf DDLPaySlipType.SelectedValue.Equals("66") Then
                HidPreVal.Value = "~" & DDLPaySlipType.SelectedValue.ToString & "~" & empcode.ToString & "~" & Month.ToString & "~" & Year.ToString & "~" &
                CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" & Level.ToString & "~" &
                EmpFName.ToString & "~" & EmpLName.ToString & "~" & SalBase.ToString & "~" & EmpType.ToString & "~" & Comzero.ToString & "~" & Hold.ToString & "~"

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
            ElseIf DDLPaySlipType.SelectedValue.Equals("67") Then

                HidPreVal.Value = Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" &
                Month.ToString & "~" & Year.ToString & "~" & EmpLName.ToString & "~" _
                & ddllingual.SelectedValue & "~" & Hold.ToString & "~" _
                & "~S" & "~" & "~~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
                'This statement is used for display "Monthly Pay Slip Exclude Details"
                rbtnmail.Checked = False
                rbtnslip.Checked = True
            ElseIf DDLPaySlipType.SelectedValue.Equals("68") Then
                HidPreVal.Value = "~" & DDLPaySlipType.SelectedValue.ToString & "~" & empcode.ToString & "~" & Month.ToString & "~" & Year.ToString & "~" &
                CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" & Level.ToString & "~" &
                EmpFName.ToString & "~" & EmpLName.ToString & "~" & SalBase.ToString & "~" & EmpType.ToString & "~" & Comzero.ToString & "~" & Hold.ToString & "~"

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
            ElseIf DDLPaySlipType.SelectedValue.Equals("69") Then
                HidPreVal.Value = DDLPaySlipType.SelectedValue.Trim & "~H1~" & empcode.ToString & "~" & EmpFName.ToString & "~" & EmpLName.ToString & "~" &
                                CC.ToString & "~" & Dep.ToString & "~" & Grad.ToString & "~" & Desig.ToString & "~" &
                               Loc.ToString & "~" & unit.ToString & "~" & SalBase.ToString & "~" &
                                Level.ToString & "~" & EmpType.ToString & "~" & Session("ugroup").ToString & "~" & Month.ToString & "~" _
                                & Year.ToString & "~~~"

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)
            ElseIf DDLPaySlipType.SelectedValue.Equals("70") Then
                ReDim arrparam(15)
                arrparam(0) = New SqlParameter("@fk_costcenter_code", USearch.UCddlcostcenter.ToString())
                arrparam(1) = New SqlParameter("@fk_loc_code", USearch.UCddllocation.ToString())
                arrparam(2) = New SqlParameter("@fk_unit", USearch.UCddlunit.ToString())
                arrparam(3) = New SqlParameter("@salaried", USearch.UCddlsalbasis.ToString())
                arrparam(4) = New SqlParameter("@pk_emp_code", USearch.UCTextcode.ToString)
                arrparam(5) = New SqlParameter("@fk_dept_code", USearch.UCddldept.ToString())
                arrparam(6) = New SqlParameter("@fk_desig_code", USearch.UCddldesig.ToString())
                arrparam(7) = New SqlParameter("@fk_grade_code", USearch.UCddlgrade.ToString())
                arrparam(8) = New SqlParameter("@fk_level_Code", USearch.UCddllevel.ToString())
                arrparam(9) = New SqlClient.SqlParameter("@mm", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
                arrparam(10) = New SqlClient.SqlParameter("@yyyy", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                arrparam(11) = New SqlParameter("@first_name", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
                arrparam(12) = New SqlParameter("@last_name", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                arrparam(13) = New SqlParameter("@flag", "R")
                arrparam(14) = New SqlParameter("@UserGroup", Session("Ugroup").ToString)
                arrparam(15) = New SqlParameter("@EmpStatus", USearch.UCddlEmp.ToString)
                ds = _ObjData.GetDataSetProc("PaySp_REFNF_Employee", arrparam)
                If ds.Tables(0).Rows.Count > 0 Then
                    Dim filename As String = "", _sw As StreamWriter, complexID As Guid = Guid.NewGuid()
                    filename = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\RefnfRegister" & "_" & Right(complexID.ToString, 6) & ".xls"
                    If System.IO.File.Exists(filename) Then
                        System.IO.File.Delete(filename)
                    End If
                    _sw = New StreamWriter(filename)
                    lblmsg.Text = ""
                    ExportToExcelXMLREFNF(ds, _sw)
                    _sw.Close()
                    _sw.Dispose()

                    Response.Clear()
                    'Dim ReadmeText As [String] = "Hello!" & vbLf & vbLf & "This is a README..." & DateTime.Now.ToString("G")
                    Response.ContentType = "application/zip"
                    Response.AddHeader("content-disposition", "filename=Refnf Register.zip")
                    Using zip As New ZipFile()
                        zip.AddFile(filename, "Refnf Register")
                        zip.Save(Response.OutputStream)
                    End Using

                    If File.Exists(filename) Then
                        File.Delete(filename)
                    End If

                    Response.End()
                Else
                    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "M", .MessageString = "No record(s) Found according to the selection criteria !"})
                    _objCommon.ShowMessage(_msg)
                    ' _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !", False)
                End If
            ElseIf DDLPaySlipType.SelectedValue = "76" Or DDLPaySlipType.SelectedValue = "77" Then

                HidPreVal.Value = "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & CC.ToString & "~" & Loc.ToString & "~" & unit.ToString & "~" &
                SalBase.ToString & "~" & empcode.ToString & "~" & EmpFName.ToString & "~" & Month.ToString &
                "~" & Year.ToString & "~" & EmpLName.ToString & "~" &
                Comzero.ToString & "~" &
                Comlogo.ToString & "~" & Hold.ToString _
                & "~~" & Comzero.ToString & "~" & Comlogo.ToString & "~S" _
                & "~~" & EmpType.ToString & "~" & DDLPaySlipType.SelectedValue.ToString _
                & "~~~~" & "N"

                'check session is blank or store value
                If Not Session(_strVal) Is Nothing Then
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                Else
                    Session.Remove(_strVal)
                    Session(_strVal) = HidPreVal.Value.ToString
                    HidPreVal.Value = _strVal.ToString
                End If
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "blankcheck1();", True)


            End If

        End Sub
        Private Sub ExportCSV()
            Dim ds As New DataSet, arrparam(16) As SqlClient.SqlParameter, stringWrite As New System.IO.StringWriter, myHTMLTextWriter As New System.Web.UI.HtmlTextWriter(stringWrite) _
          , arrcode() As String = Nothing, countemp As Integer = 0, empcode As String = "", lstitem As ListItem, paycodeSel1 As String = "", lstitem1 As ListItem _
          , paycodeSel As String = "", Dep As String = "", Desig As String = "", Grad As String = "", Level As String = "", CC As String = "", Loc As String = "" _
          , unit As String = "", SalBase As String = "", EmpFName As String = "", EmpLName As String = "", EmpType As String = "", Month As String = "", Year As String = "" _
          , Sorttype As String = "", ShortType As String = "", PayCode As String = "", reptype As String = "", PayCodeAdd As String = "", Ids As String = "" _
          , srttp As String = "", Hold As String = "", doj As String = "", PF As String = "", Leave As String = "", Absent As String = "", ESI As String = "" _
          , BankAcc As String = "", PayCodeId As String = "", Salut As String = "", _strVal As String = Guid.NewGuid.ToString, _str As New System.Text.StringBuilder _
          , str As String = "", chkitem As ListItem, Comzero As String = "", ReptType As String = "", RepFormat As String = "", chkPF_Old As String = "" _
          , chkESI_Old As String = "", Comlogo As String = "", Loan As String = "", Advance As String = "", Ent As String = "", otherinc As String = "" _
          , StaffId As String = "", NegitiveSalFlg As String = "", arrp(20) As SqlClient.SqlParameter, arrpam(18) As SqlClient.SqlParameter _
          , arrpamwork(12) As SqlClient.SqlParameter, arrpamm(18) As SqlClient.SqlParameter, dtRepType As DataTable, ReportType As String = "", Fromdate As String = ""
            hdnFileFormat.Value = "CSV"
            'CheckExcelProcessbarAlreadyProcessing()
            'If (lblProcessStatusExcel.Text <> "") Then
            '    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = lblProcessStatusExcel.Text})
            '    _objCommon.ShowMessage(_msg)
            '    Exit Sub
            'End If
            'Added by Debargha on 21 Oct 2024
            Dim APIConfigParam(2) As SqlClient.SqlParameter, IsNewUrl As String = "N"
            Dim AppPathStr As String = HttpRuntime.AppDomainAppVirtualPath.ToString, _array() As String
            _array = Split(AppPathStr, "/")
            AppPathStr = _array(_array.Length - 1)

            hidothrpaycode.Value = ""
            hdquery.Value = ""
            'Common variable start here
            Dep = USearch.UCddldept.ToString()
            Desig = USearch.UCddldesig.ToString()
            Grad = USearch.UCddlgrade.ToString()
            Level = USearch.UCddllevel.ToString()
            CC = USearch.UCddlcostcenter.ToString()
            Loc = USearch.UCddllocation.ToString()
            unit = USearch.UCddlunit.ToString()
            SalBase = USearch.UCddlsalbasis.ToString()
            empcode = USearch.UCTextcode.ToString
            EmpFName = IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), "")
            Month = _objCommon.nNz(ddlMonthYear.SelectedValue.ToString)
            Year = Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4)
            EmpLName = IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), "")
            EmpType = USearch.UCddlEmp.ToString()
            Hold = ddlshowsal.SelectedValue.ToString
            'If DDLPaySlipType.SelectedValue = "38" Then
            '    'Added by Debargha on 21-Oct-2024
            '    APIConfigParam(0) = New SqlClient.SqlParameter("@SP_Name", "PaySP_SalaryRegInExcel_Dynamic")
            '    APIConfigParam(1) = New SqlClient.SqlParameter("@ReportName", "HRD Report")
            '    APIConfigParam(2) = New SqlClient.SqlParameter("@IsNewURL", SqlDbType.VarChar, 1)
            '    APIConfigParam(2).Direction = ParameterDirection.Output
            '    _ObjData.ExecuteStoredProcMsg("PaySP_ReportAPI_ConfigSel", APIConfigParam)
            '    IsNewUrl = APIConfigParam(2).Value.ToString

            '    If IsNewUrl = Nothing Or IsNewUrl = "N" Then
            '        _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = DDLPaySlipType.SelectedValue & " Report is not configured till for CSV!"})
            '        _objCommon.ShowMessage(_msg)
            '    Else
            '        Dim arprm(8) As SqlClient.SqlParameter, apiurl As String
            '        arprm(0) = New SqlClient.SqlParameter("@UserID", Session("UID").ToString)
            '        arprm(1) = New SqlClient.SqlParameter("@Process_Type", GetApiProcessType(DDLPaySlipType.SelectedValue))
            '        arprm(2) = New SqlClient.SqlParameter("@ActionType", "Init")
            '        arprm(3) = New SqlClient.SqlParameter("@Sys_IP", "::1")
            '        arprm(4) = New SqlClient.SqlParameter("@HostIP", ConfigurationManager.AppSettings("Hostip").ToString())
            '        arprm(5) = New SqlClient.SqlParameter("@ProcName", "PaySP_SalaryRegInExcel_Dynamic")
            '        arprm(6) = New SqlClient.SqlParameter("@DdlRptName", DDLPaySlipType.SelectedItem.Text.Replace("'", ""))
            '        arprm(7) = New SqlClient.SqlParameter("@DdlRptId", DDLPaySlipType.SelectedValue)
            '        arprm(8) = New SqlClient.SqlParameter("@RPTFRMT", "CSV")
            '        Dim _dt As DataTable = _ObjData.GetDataTableProc("PaySP_ReportApi_ProcessBar", arprm)
            '        If (_dt.Rows.Count > 0) Then
            '            If (_dt.Rows(0)("IsAbleToStart").ToString = "1" AndAlso _dt.Rows(0)("BatchId").ToString <> "") Then
            '                Dim scripttag As String = "StartCSVProcessbar('" & DDLPaySlipType.SelectedItem.Text.Replace("'", "") & "');"
            '                'ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpooce89sbar231", scripttag, True)
            '                ScriptManager.RegisterStartupScript(Me, Me.GetType(), "openpooce89sbar231", scripttag, True)
            '                hdnBatchId.Value = _dt.Rows(0)("BatchId").ToString
            '                btnProgressbarExcel.Visible = False
            '                lblProcessStatusExcel.Text = ""
            '                divSocialExcel.Visible = False
            '                apiurl = _dt.Rows(0)("apiurl").ToString
            '            Else
            '                divSocialExcel.Visible = True
            '                lblProcessStatusExcel.Text = DDLPaySlipType.SelectedItem.Text.Replace("'", "") & " is already processing. Please wait till the completion."
            '                If (_dt.Rows(0)("UserId").ToString.ToUpper <> Session("UID").ToUpper.ToString) Then
            '                    btnProgressbarExcel.Visible = True
            '                Else
            '                    btnProgressbarExcel.Visible = True
            '                End If
            '                _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = lblProcessStatusExcel.Text})
            '                _objCommon.ShowMessage(_msg)
            '                ScriptManager.RegisterStartupScript(Me.Page, GetType(Page), "refreshExeclProcessStatus97", "ShowExcelLockSummaryDetails('" & GetApiProcessType(DDLPaySlipType.SelectedValue) & "')", True)
            '                Exit Sub
            '            End If
            '        End If
            '        Dim keyValuePairs As New Dictionary(Of String, Object) From {
            '            {"hostIp", ConfigurationManager.AppSettings("Hostip").ToString()},
            '            {"userId", HttpContext.Current.Session("UID").ToString()},
            '            {"moduleType", HttpContext.Current.Session("ModuleType").ToString()},
            '            {"domainName", Session("CompCode").ToString},
            '            {"showClr", rbtshowclr.SelectedValue.ToString.ToUpper},
            '            {"fk_costcenter_code", USearch.UCddlcostcenter.ToString()},
            '            {"fk_loc_code", USearch.UCddllocation.ToString()},
            '            {"fk_unit", USearch.UCddlunit.ToString()},
            '            {"salaried", USearch.UCddlsalbasis.ToString()},
            '            {"pk_emp_code", USearch.UCTextcode.ToString.ToString},
            '            {"fk_dept_code", USearch.UCddldept.ToString()},
            '            {"fk_desig_code", USearch.UCddldesig.ToString()},
            '            {"fk_grade_code", USearch.UCddlgrade.ToString()},
            '            {"fk_level_code", USearch.UCddllevel.ToString()},
            '            {"month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString)},
            '            {"year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4)},
            '            {"firstName", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), "")},
            '            {"lastName", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), "")},
            '            {"hold", ddlshowsal.SelectedValue.ToString},
            '            {"userGroup", Session("Ugroup").ToString},
            '            {"empStatus", USearch.UCddlEmp.ToString},
            '            {"repId", DDLPaySlipType.SelectedValue.ToString},
            '            {"BatchId", hdnBatchId.Value.ToString},
            '            {"FileFormat", hdnFileFormat.Value.ToString},
            '            {"FileName", txtrptName.Text.ToString}
            '        }


            '        Dim requestBody As String = JsonConvert.SerializeObject(keyValuePairs)
            '        'CallAPIReport(requestBody, "SalaryRegister")
            '        CallReportAPIOnNewThread(requestBody, "SalaryRegister", AppPathStr, apiurl, "CSVSalaryRegister")
            '    End If
            '    ' Add new payslip [Salary register group wise.], by praveen verma on 23 Aug 2013.
            'End If
            If DDLPaySlipType.SelectedValue = "38" Then
                arrparam(0) = New SqlParameter("@fk_costcenter_code", USearch.UCddlcostcenter.ToString())
                arrparam(1) = New SqlParameter("@fk_loc_code", USearch.UCddllocation.ToString())
                arrparam(2) = New SqlParameter("@fk_unit", USearch.UCddlunit.ToString())
                arrparam(3) = New SqlParameter("@salaried", USearch.UCddlsalbasis.ToString())
                arrparam(4) = New SqlParameter("@pk_emp_code", USearch.UCTextcode.ToString)
                arrparam(5) = New SqlParameter("@fk_dept_code", USearch.UCddldept.ToString())
                arrparam(6) = New SqlParameter("@fk_desig_code", USearch.UCddldesig.ToString())
                arrparam(7) = New SqlParameter("@fk_grade_code", USearch.UCddlgrade.ToString())
                arrparam(8) = New SqlParameter("@fk_level_Code", USearch.UCddllevel.ToString())
                arrparam(9) = New SqlClient.SqlParameter("@month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
                arrparam(10) = New SqlClient.SqlParameter("@year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                arrparam(11) = New SqlParameter("@first_name", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
                arrparam(12) = New SqlParameter("@last_name", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                arrparam(13) = New SqlParameter("@Hold", ddlshowsal.SelectedValue.ToString)
                arrparam(14) = New SqlParameter("@UserGroup", Session("Ugroup").ToString)
                arrparam(15) = New SqlParameter("@EmpStatus", USearch.UCddlEmp.ToString)
                arrparam(16) = New SqlParameter("@RepID", DDLPaySlipType.SelectedValue.ToString)
                ds = _ObjData.GetDataSetProc("PaySP_SalaryRegInExcel_Dynamic", arrparam)
                If ds.Tables(0).Rows.Count > 0 Then
                    Dim _ExcelName As String = "", complexID As Guid = Guid.NewGuid(), folderPath As String = "", filePath As String = ""
                    If String.IsNullOrEmpty(txtrptName.Text.Trim.ToString()) Then
                        _ExcelName = "Salary Register" & "_" & Right(complexID.ToString, 6)
                    Else
                        _ExcelName = txtrptName.Text.Trim.ToString()
                    End If
                    'If ddllEncrType.SelectedValue.Trim.ToString <> "WP" Then
                    folderPath = _objCommon.GetDirpath(Session("CompCode").ToString) & "\Documents\DynamicSalaryRegister\"
                    'Else
                    '    folderPath = _objCommon.GetDirpath(Session("CompCode").ToString) & "\Documents\DynamicSalaryRegister\EncryptedCSVFiles\"
                    'End If


                    If Directory.Exists(folderPath) Then
                        For Each filepaths As String In Directory.GetFiles(folderPath)
                            File.Delete(filepaths)
                        Next
                    End If

                    If Not Directory.Exists(folderPath) Then
                        Directory.CreateDirectory(folderPath)
                    End If
                    Dim outputCsvPath As String = folderPath & _ExcelName & ".csv"
                    filePath = GenerateCsvFromDataTable(ds.Tables(0), outputCsvPath)
                    'Dim filename As String = "", _sw As StreamWriter
                    'Dim folderPath As String = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\"
                    'filename = folderPath & _ExcelName

                    If ddllEncrType.SelectedValue.ToString.Trim = "WE" Then
                        hdfile.Value = folderPath + "~" + _ExcelName + "~" + "A" + "~" + (IIf(chkSFTP.Checked, "Y", "N")).ToString + "~" + "38" + "~" + "DynamicSalaryRegister"
                        Dim popupScript As String = "<script language='javascript' type='text/javascript'>ShowDownload('" & hdfile.Value & "')</script>"
                        ClientScript.RegisterStartupScript(GetType(String), "PopupScript", popupScript)
                    Else
                        Dim param(1) As SqlClient.SqlParameter
                        param(0) = New SqlClient.SqlParameter("@ReportId", "38")
                        param(1) = New SqlClient.SqlParameter("@RptType", "DYNSALREG")
                        Dim dt As DataTable = _ObjData.GetDataTableProc("Paysp_MstPGPEncryptionConfig_GetEncrKey", param)

                        Dim PublicKeyPath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\SalaryRegister\PGP\Enc\" & dt.Rows(0)("EncrFileName").ToString.Trim
                        Dim ofileName As String = _ExcelName
                        Dim PGPFilePath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\EncryptedFiles\" & ofileName & ".asc"

                        If Not Directory.Exists(Path.GetDirectoryName(PGPFilePath)) Then
                            Directory.CreateDirectory(Path.GetDirectoryName(PGPFilePath))
                        End If

                        ' Send file for encryption
                        Try
                            Dim url As String = ConfigurationManager.AppSettings("PGPCommonApi").ToString
                            Dim boundary As String = "---------------------------" & DateTime.Now.Ticks.ToString("x")
                            Dim encoding As Encoding = Encoding.UTF8
                            ServicePointManager.SecurityProtocol = CType(SecurityProtocolType.Tls12, SecurityProtocolType)

                            Dim request As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
                            request.Method = "POST"
                            request.ContentType = "multipart/form-data; boundary=" & boundary
                            request.KeepAlive = True
                            ServicePointManager.ServerCertificateValidationCallback = Function(sender, cert, chain, sslPolicyErrors) True

                            Using requestStream As Stream = request.GetRequestStream()
                                WriteFormField(requestStream, "AuthKey", "ENC8C786CD454950A0A01609AD767DA2", boundary, encoding)
                                WriteFormField(requestStream, "Username", Session("UID").ToString.Trim, boundary, encoding)
                                WriteFormField(requestStream, "DomainCode", Session("Compcode").ToString.ToUpper, boundary, encoding)
                                WriteFileField(requestStream, "inputpgpfile", filePath, "csv", boundary, encoding)
                                WriteFileField(requestStream, "inputencrkey", PublicKeyPath, "application/octet-stream", boundary, encoding)
                                Dim trailer As String = "--" & boundary & "--" & vbCrLf
                                Dim trailerBytes As Byte() = encoding.GetBytes(trailer)
                                requestStream.Write(trailerBytes, 0, trailerBytes.Length)
                            End Using

                            Using response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                                Using reader As New StreamReader(response.GetResponseStream())
                                    Dim responseText As String = reader.ReadToEnd()
                                    File.WriteAllText(PGPFilePath, responseText)
                                End Using
                            End Using

                            hdfile.Value = PGPFilePath & "~" & ofileName & "~ASC~" & IIf(chkSFTP.Checked, "Y", "N").ToString() + "~" + "38" + "~" + "DynamicSalaryRegister"
                            Dim popupScript As String = "<script language='javascript' type='text/javascript'>ShowDownload('" & hdfile.Value & "')</script>"
                            ClientScript.RegisterStartupScript(GetType(String), "PopupScript", popupScript)

                        Catch ex As WebException
                            Using errorResponse As HttpWebResponse = CType(ex.Response, HttpWebResponse)
                                Using reader As New StreamReader(errorResponse.GetResponseStream())
                                    Dim errorText As String = reader.ReadToEnd()
                                    Console.WriteLine("Encryption API Error: " & errorText)
                                End Using
                            End Using
                            Throw
                        End Try
                    End If
                Else
                    Dim _msg As New List(Of PayrollUtility.UserMessage)
                    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "M", .MessageString = "No record(s) Found according to the selection criteria !"})
                    _objCommon.ShowMessage(_msg)
                    Exit Sub
                End If
            End If

        End Sub
        Private Sub btnExport2CSV_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExport2CSV.Click
            Dim ds As New DataSet, arrparam(16) As SqlClient.SqlParameter, stringWrite As New System.IO.StringWriter, myHTMLTextWriter As New System.Web.UI.HtmlTextWriter(stringWrite) _
            , arrcode() As String = Nothing, countemp As Integer = 0, empcode As String = "", lstitem As ListItem, paycodeSel1 As String = "", lstitem1 As ListItem _
            , paycodeSel As String = "", Dep As String = "", Desig As String = "", Grad As String = "", Level As String = "", CC As String = "", Loc As String = "" _
            , unit As String = "", SalBase As String = "", EmpFName As String = "", EmpLName As String = "", EmpType As String = "", Month As String = "", Year As String = "" _
            , Sorttype As String = "", ShortType As String = "", PayCode As String = "", reptype As String = "", PayCodeAdd As String = "", Ids As String = "" _
            , srttp As String = "", Hold As String = "", doj As String = "", PF As String = "", Leave As String = "", Absent As String = "", ESI As String = "" _
            , BankAcc As String = "", PayCodeId As String = "", Salut As String = "", _strVal As String = Guid.NewGuid.ToString, _str As New System.Text.StringBuilder _
            , str As String = "", chkitem As ListItem, Comzero As String = "", ReptType As String = "", RepFormat As String = "", chkPF_Old As String = "" _
            , chkESI_Old As String = "", Comlogo As String = "", Loan As String = "", Advance As String = "", Ent As String = "", otherinc As String = "" _
            , StaffId As String = "", NegitiveSalFlg As String = "", arrp(20) As SqlClient.SqlParameter, arrpam(18) As SqlClient.SqlParameter _
            , arrpamwork(12) As SqlClient.SqlParameter, arrpamm(18) As SqlClient.SqlParameter, dtRepType As DataTable, ReportType As String = "", Fromdate As String = ""
            hdnFileFormat.Value = "CSV"
            'CheckExcelProcessbarAlreadyProcessing()
            'If (lblProcessStatusExcel.Text <> "") Then
            '    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = lblProcessStatusExcel.Text})
            '    _objCommon.ShowMessage(_msg)
            '    Exit Sub
            'End If
            'Added by Debargha on 21 Oct 2024
            Dim APIConfigParam(2) As SqlClient.SqlParameter, IsNewUrl As String = "N"
            Dim AppPathStr As String = HttpRuntime.AppDomainAppVirtualPath.ToString, _array() As String
            _array = Split(AppPathStr, "/")
            AppPathStr = _array(_array.Length - 1)

            hidothrpaycode.Value = ""
            hdquery.Value = ""
            'Common variable start here
            Dep = USearch.UCddldept.ToString()
            Desig = USearch.UCddldesig.ToString()
            Grad = USearch.UCddlgrade.ToString()
            Level = USearch.UCddllevel.ToString()
            CC = USearch.UCddlcostcenter.ToString()
            Loc = USearch.UCddllocation.ToString()
            unit = USearch.UCddlunit.ToString()
            SalBase = USearch.UCddlsalbasis.ToString()
            empcode = USearch.UCTextcode.ToString
            EmpFName = IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), "")
            Month = _objCommon.nNz(ddlMonthYear.SelectedValue.ToString)
            Year = Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4)
            EmpLName = IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), "")
            EmpType = USearch.UCddlEmp.ToString()
            Hold = ddlshowsal.SelectedValue.ToString
            If DDLPaySlipType.SelectedValue = "21" Then
                'Added by Debargha on 21-Oct-2024
                APIConfigParam(0) = New SqlClient.SqlParameter("@SP_Name", "PaySP_SalaryRegInExcel")
                APIConfigParam(1) = New SqlClient.SqlParameter("@ReportName", "HRD Report")
                APIConfigParam(2) = New SqlClient.SqlParameter("@IsNewURL", SqlDbType.VarChar, 1)
                APIConfigParam(2).Direction = ParameterDirection.Output
                _ObjData.ExecuteStoredProcMsg("PaySP_ReportAPI_ConfigSel", APIConfigParam)
                IsNewUrl = APIConfigParam(2).Value.ToString
                If IsNewUrl = Nothing OrElse IsNewUrl = "N" Then
                    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = DDLPaySlipType.SelectedValue & " Report is not configured till for CSV!"})
                    _objCommon.ShowMessage(_msg)
                Else
                    Dim arprm(7) As SqlClient.SqlParameter, apiurl As String
                    arprm(0) = New SqlClient.SqlParameter("@UserID", Session("UID").ToString)
                    arprm(1) = New SqlClient.SqlParameter("@Process_Type", GetApiProcessType(DDLPaySlipType.SelectedValue))
                    arprm(2) = New SqlClient.SqlParameter("@ActionType", "Init")
                    arprm(3) = New SqlClient.SqlParameter("@Sys_IP", "::1")
                    arprm(4) = New SqlClient.SqlParameter("@HostIP", ConfigurationManager.AppSettings("Hostip").ToString())
                    arprm(5) = New SqlClient.SqlParameter("@ProcName", "PaySP_SalaryRegInExcel")
                    arprm(6) = New SqlClient.SqlParameter("@DdlRptName", DDLPaySlipType.SelectedItem.Text.Replace("'", ""))
                    arprm(7) = New SqlClient.SqlParameter("@DdlRptId", DDLPaySlipType.SelectedValue)
                    Dim _dt As DataTable = _ObjData.GetDataTableProc("PaySP_ReportApi_ProcessBar", arprm)
                    If (_dt.Rows.Count > 0) Then
                        If (_dt.Rows(0)("IsAbleToStart").ToString = "1" AndAlso _dt.Rows(0)("BatchId").ToString <> "") Then
                            Dim scripttag As String = "StartProcessbar('" & DDLPaySlipType.SelectedItem.Text.Replace("'", "") & "');"
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpooce89sbar231", scripttag, True)
                            hdnBatchId.Value = _dt.Rows(0)("BatchId").ToString
                            btnProgressbarExcel.Visible = False
                            divSocialExcel.Visible = False
                            lblProcessStatusExcel.Text = ""
                            apiurl = _dt.Rows(0)("apiurl").ToString
                        Else
                            divSocialExcel.Visible = True
                            lblProcessStatusExcel.Text = DDLPaySlipType.SelectedItem.Text.Replace("'", "") & " is already processing. Please wait till the completion."
                            If (_dt.Rows(0)("UserId").ToString.ToUpper <> Session("UID").ToUpper.ToString) Then
                                btnProgressbarExcel.Visible = False
                            Else
                                btnProgressbarExcel.Visible = True
                            End If
                            _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = lblProcessStatusExcel.Text})
                            _objCommon.ShowMessage(_msg)
                            ScriptManager.RegisterStartupScript(Me.Page, GetType(Page), "refreshExeclProcessStatus98", "ShowExcelLockSummaryDetails('" & GetApiProcessType(DDLPaySlipType.SelectedValue) & "')", True)
                            Exit Sub
                        End If
                    End If
                    Dim keyValuePairs As New Dictionary(Of String, Object) From {
                        {"hostIp", ConfigurationManager.AppSettings("Hostip").ToString()},
                        {"userId", HttpContext.Current.Session("UID").ToString()},
                        {"moduleType", HttpContext.Current.Session("ModuleType").ToString()},
                        {"domainName", Session("CompCode").ToString},
                        {"showClr", rbtshowclr.SelectedValue.ToString.ToUpper},
                        {"fk_costcenter_code", USearch.UCddlcostcenter.ToString()},
                        {"fk_loc_code", USearch.UCddllocation.ToString()},
                        {"fk_unit", USearch.UCddlunit.ToString()},
                        {"salaried", USearch.UCddlsalbasis.ToString()},
                        {"pk_emp_code", USearch.UCTextcode.ToString.ToString},
                        {"fk_dept_code", USearch.UCddldept.ToString()},
                        {"fk_desig_code", USearch.UCddldesig.ToString()},
                        {"fk_grade_code", USearch.UCddlgrade.ToString()},
                        {"fk_level_code", USearch.UCddllevel.ToString()},
                        {"month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString)},
                        {"year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4)},
                        {"firstName", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), "")},
                        {"lastName", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), "")},
                        {"hold", ddlshowsal.SelectedValue.ToString},
                        {"userGroup", Session("Ugroup").ToString},
                        {"empStatus", USearch.UCddlEmp.ToString},
                        {"repId", DDLPaySlipType.SelectedValue.ToString},
                        {"BatchId", hdnBatchId.Value.ToString},
                        {"FileFormat", hdnFileFormat.Value.ToString}
                    }
                    Dim requestBody As String = JsonConvert.SerializeObject(keyValuePairs)
                    'Modified by Vishal Chauhan to call HeavyExcel API
                    CallReportAPIOnNewThreadHeavyExcel(requestBody, "StaticSalRegister", AppPathStr, apiurl)
                End If
            ElseIf DDLPaySlipType.SelectedValue = "38" Then
                'Added by Debargha on 21-Oct-2024
                APIConfigParam(0) = New SqlClient.SqlParameter("@SP_Name", "PaySP_SalaryRegInExcel_Dynamic")
                APIConfigParam(1) = New SqlClient.SqlParameter("@ReportName", "HRD Report")
                APIConfigParam(2) = New SqlClient.SqlParameter("@IsNewURL", SqlDbType.VarChar, 1)
                APIConfigParam(2).Direction = ParameterDirection.Output
                _ObjData.ExecuteStoredProcMsg("PaySP_ReportAPI_ConfigSel", APIConfigParam)
                IsNewUrl = APIConfigParam(2).Value.ToString

                If IsNewUrl = Nothing Or IsNewUrl = "N" Then
                    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = DDLPaySlipType.SelectedValue & " Report is not configured till for CSV!"})
                    _objCommon.ShowMessage(_msg)
                Else
                    Dim arprm(7) As SqlClient.SqlParameter, apiurl As String
                    arprm(0) = New SqlClient.SqlParameter("@UserID", Session("UID").ToString)
                    arprm(1) = New SqlClient.SqlParameter("@Process_Type", GetApiProcessType(DDLPaySlipType.SelectedValue))
                    arprm(2) = New SqlClient.SqlParameter("@ActionType", "Init")
                    arprm(3) = New SqlClient.SqlParameter("@Sys_IP", "::1")
                    arprm(4) = New SqlClient.SqlParameter("@HostIP", ConfigurationManager.AppSettings("Hostip").ToString())
                    arprm(5) = New SqlClient.SqlParameter("@ProcName", "PaySP_SalaryRegInExcel_Dynamic")
                    arprm(6) = New SqlClient.SqlParameter("@DdlRptName", DDLPaySlipType.SelectedItem.Text.Replace("'", ""))
                    arprm(7) = New SqlClient.SqlParameter("@DdlRptId", DDLPaySlipType.SelectedValue)
                    Dim _dt As DataTable = _ObjData.GetDataTableProc("PaySP_ReportApi_ProcessBar", arprm)
                    If (_dt.Rows.Count > 0) Then
                        If (_dt.Rows(0)("IsAbleToStart").ToString = "1" AndAlso _dt.Rows(0)("BatchId").ToString <> "") Then
                            Dim scripttag As String = "StartProcessbar('" & DDLPaySlipType.SelectedItem.Text.Replace("'", "") & "');"
                            'ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpooce89sbar231", scripttag, True)
                            ScriptManager.RegisterStartupScript(Me, Me.GetType(), "openpooce89sbar231", scripttag, True)
                            hdnBatchId.Value = _dt.Rows(0)("BatchId").ToString
                            btnProgressbarExcel.Visible = False
                            lblProcessStatusExcel.Text = ""
                            divSocialExcel.Visible = False
                            apiurl = _dt.Rows(0)("apiurl").ToString
                        Else
                            divSocialExcel.Visible = True
                            lblProcessStatusExcel.Text = DDLPaySlipType.SelectedItem.Text.Replace("'", "") & " is already processing. Please wait till the completion."
                            If (_dt.Rows(0)("UserId").ToString.ToUpper <> Session("UID").ToUpper.ToString) Then
                                btnProgressbarExcel.Visible = True
                            Else
                                btnProgressbarExcel.Visible = True
                            End If
                            _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = lblProcessStatusExcel.Text})
                            _objCommon.ShowMessage(_msg)
                            ScriptManager.RegisterStartupScript(Me.Page, GetType(Page), "refreshExeclProcessStatus97", "ShowExcelLockSummaryDetails('" & GetApiProcessType(DDLPaySlipType.SelectedValue) & "')", True)
                            Exit Sub
                        End If
                    End If
                    Dim keyValuePairs As New Dictionary(Of String, Object) From {
                        {"hostIp", ConfigurationManager.AppSettings("Hostip").ToString()},
                        {"userId", HttpContext.Current.Session("UID").ToString()},
                        {"moduleType", HttpContext.Current.Session("ModuleType").ToString()},
                        {"domainName", Session("CompCode").ToString},
                        {"showClr", rbtshowclr.SelectedValue.ToString.ToUpper},
                        {"fk_costcenter_code", USearch.UCddlcostcenter.ToString()},
                        {"fk_loc_code", USearch.UCddllocation.ToString()},
                        {"fk_unit", USearch.UCddlunit.ToString()},
                        {"salaried", USearch.UCddlsalbasis.ToString()},
                        {"pk_emp_code", USearch.UCTextcode.ToString.ToString},
                        {"fk_dept_code", USearch.UCddldept.ToString()},
                        {"fk_desig_code", USearch.UCddldesig.ToString()},
                        {"fk_grade_code", USearch.UCddlgrade.ToString()},
                        {"fk_level_code", USearch.UCddllevel.ToString()},
                        {"month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString)},
                        {"year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4)},
                        {"firstName", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), "")},
                        {"lastName", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), "")},
                        {"hold", ddlshowsal.SelectedValue.ToString},
                        {"userGroup", Session("Ugroup").ToString},
                        {"empStatus", USearch.UCddlEmp.ToString},
                        {"repId", DDLPaySlipType.SelectedValue.ToString},
                        {"BatchId", hdnBatchId.Value.ToString},
                        {"FileFormat", hdnFileFormat.Value.ToString}
                    }
                    Dim requestBody As String = JsonConvert.SerializeObject(keyValuePairs)
                    'CallAPIReport(requestBody, "SalaryRegister")
                    CallReportAPIOnNewThread(requestBody, "SalaryRegister", AppPathStr, apiurl)
                End If
                ' Add new payslip [Salary register group wise.], by praveen verma on 23 Aug 2013.
            End If

        End Sub
        Protected Sub Btnsearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Btnsearch.Click
            lblMsgSlip.Text = ""
            lblMailMsg.Text = ""
            lblProcessBarMsg.Text = ""
            LnkPDF.Style.Value = "display:none"
            LnkPDFWOPWD.Style.Value = "display:None"
            download_pdf1.Style.Value = "display:none;"
            download_pdf2.Style.Value = "display:none;"
            process_status_id.Value = ""
            lblMailMsgWOPWD.Text = ""
            'Excel Process locking validation checking
            'CheckExcelProcessbarAlreadyProcessing()
            'If (lblProcessStatusExcel.Text <> "") Then
            '    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocessF8201", "UnLoadPaySlipProgress();", True)
            '    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = lblProcessStatusExcel.Text})
            '    _objCommon.ShowMessage(_msg)
            '    Exit Sub
            'End If
            CheckProcessLocked()
            If (hdnAlreadyRunRptName.Value.Trim().Length > 1) Then
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocessF8211", "UnLoadPaySlipProgress();", True)
                _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = hdnAlreadyRunRptName.Value})
                _objCommon.ShowMessage(_msg)
                Exit Sub
            End If

            If ddlRepIn.SelectedValue.ToUpper = "P" Then
                tblpwd.Style.Value = "display:"
                TrNoSearch.Style.Value = "display:"
                If Convert.ToString(DdlreportType.SelectedValue).Equals("55") Or Convert.ToString(DdlreportType.SelectedValue).Equals("56") _
                    Or Convert.ToString(DdlreportType.SelectedValue).Equals("57") Or Convert.ToString(DdlreportType.SelectedValue).Equals("62") _
                    Or Convert.ToString(DdlreportType.SelectedValue).Equals("63") Then
                    Tr2.Style.Value = "display:none"
                ElseIf DdlreportType.SelectedValue.ToString <> "43" Or DdlreportType.SelectedValue.ToString <> "49" Then
                    trselall.Style.Value = "display:"
                    Tr2.Style.Value = "display:"
                    BtnSend.Visible = "true"
                    BtnPublishGrpBy.Visible = True
                    BtnSendCCBCC.Visible = "true"
                    BtnLog.Visible = True
                ElseIf DdlreportType.SelectedValue.Equals("74") Then
                    BtnPublishGrpBy.Visible = False
                End If
                If Convert.ToString(DdlreportType.SelectedValue).Equals("R") Then
                    btnWOPWD.Visible = True
                Else
                    btnWOPWD.Visible = False
                End If
            ElseIf ddlRepIn.SelectedValue.ToUpper = "L" Then
                trselall.Style.Value = "display:"
            ElseIf ddlRepIn.SelectedValue.ToUpper = "H" Then
                Tr2.Style.Value = "display:none"
                btnWOPWD.Visible = False
                If DdlreportType.SelectedValue.Equals("74") Then
                    BtnSend.Visible = False
                Else
                    BtnSend.Visible = True
                End If

            End If
            Try
                If DdlreportType.SelectedValue = "I" Then
                    SendInvestmentDetials()
                Else
                    DgPayslip.Columns(6).Visible = True
                    DgPayslip.Columns(7).Visible = True
                    populateDgPayslip()
                End If
            Catch ex As Exception
                _objcommonExp.PublishError("Btnsearch_Click", ex)
            End Try
        End Sub
        Private Sub populateDgPayslip(Optional ByVal _SortBY As String = "")
            Try
                Dim _Ds As New DataSet, _msg As New List(Of PayrollUtility.UserMessage), _MsgReturn As String = ""
                HidYear.Value = Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4)
                _Ds = ReturnDsSearch(_MsgReturn)
                If _Ds.Tables.Count > 0 Then
                    If _Ds.Tables(0).Rows.Count > 0 Then
                        DgPayslip.DataSource = _Ds.Tables(0)
                        DgPayslip.DataBind()
                        For _i As Integer = 0 To DgPayslip.Items.Count - 1
                            If CType(DgPayslip.Items(_i).Cells(6).Text.Trim.ToString, String) = "Processed Salary" And (CType(DgPayslip.Items(_i).Cells(9).Text.Trim.ToString, String) = "Hold" Or DgPayslip.Items(_i).Cells(5).Text.Trim.ToString = "N/A") Then
                                DgPayslip.Items(_i).BackColor = Color.Pink
                                CType(DgPayslip.Items(_i).FindControl("chkEmpHold"), CheckBox).Checked = False
                                CType(DgPayslip.Items(_i).FindControl("chkEmpHold"), CheckBox).Enabled = True
                            ElseIf CType(DgPayslip.Items(_i).Cells(6).Text.Trim.ToString, String) = "Salary Not Processed" And DdlreportType.SelectedValue.ToString.ToUpper <> "RN" Then
                                CType(DgPayslip.Items(_i).FindControl("chkEmpHold"), CheckBox).Checked = False
                                If DdlreportType.SelectedValue.ToString.ToUpper <> "62" Then
                                    CType(DgPayslip.Items(_i).FindControl("LinkButton1"), LinkButton).Enabled = False
                                End If
                                DgPayslip.Items(_i).BackColor = Color.SkyBlue
                            ElseIf CType(DgPayslip.Items(_i).Cells(6).Text.Trim.ToString, String) = "Salary Not Processed" And CType(DgPayslip.Items(_i).Cells(5).Text.Trim.ToString, String) <> "N/A" Then
                                DgPayslip.Items(_i).BackColor = Color.SkyBlue
                                CType(DgPayslip.Items(_i).FindControl("chkEmpHold"), CheckBox).Checked = False
                            End If
                        Next
                        If _MsgReturn.ToString <> "" Then
                            lblProcessBarMsg.Text = _MsgReturn.ToString
                            '_msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = _MsgReturn.ToString})
                        End If
                        tableshow.Style("display") = ""
                        trbutton.Style("display") = ""
                        TrDg.Style("display") = ""
                        Hidden1.Value = DgPayslip.Items.Count.ToString
                        lblmsg1.Text = "Total " & DgPayslip.Items.Count.ToString & " Record(s) Found !"
                    Else
                        DgPayslip.DataSource = Nothing
                        DgPayslip.DataBind()
                        If _MsgReturn.ToString <> "" Then
                            lblProcessBarMsg.Text = _MsgReturn.ToString
                            '_msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = _MsgReturn.ToString})
                        Else
                            _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "No record(s) Found according to the selection criteria !"})
                        End If
                        tableshow.Style("display") = "none"
                        trbutton.Style("display") = "none"
                        Hidden1.Value = "0"
                        lblmsg1.Text = ""
                    End If

                    If DgPayslip.Items.Count > 2000 Then
                        DgPayslip.Columns(2).Visible = False
                        DgPayslip.Columns(3).Visible = False
                        DgPayslip.Columns(4).Visible = False
                        DgPayslip.Columns(5).Visible = False
                        DgPayslip.Columns(6).Visible = False
                        DgPayslip.Columns(7).Visible = False
                        DgPayslip.Columns(8).Visible = False
                    ElseIf DgPayslip.Items.Count > 500 Then
                        DgPayslip.Columns(2).Visible = True
                        DgPayslip.Columns(3).Visible = False
                        DgPayslip.Columns(4).Visible = False
                        DgPayslip.Columns(5).Visible = False
                        DgPayslip.Columns(6).Visible = False
                        DgPayslip.Columns(7).Visible = False
                        DgPayslip.Columns(8).Visible = False
                    Else
                        DgPayslip.Columns(2).Visible = True
                        DgPayslip.Columns(3).Visible = True
                        DgPayslip.Columns(4).Visible = True
                        DgPayslip.Columns(5).Visible = True
                        DgPayslip.Columns(6).Visible = True
                        DgPayslip.Columns(7).Visible = True
                        DgPayslip.Columns(8).Visible = True
                    End If
                Else
                    DgPayslip.DataSource = Nothing
                    DgPayslip.DataBind()
                    If _MsgReturn.ToString <> "" Then
                        '_msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = _MsgReturn.ToString})
                        lblProcessBarMsg.Text = _MsgReturn.ToString
                    Else
                        _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "No record(s) Found according to the selection criteria !"})
                    End If
                    tableshow.Style("display") = "none"
                    trbutton.Style("display") = "none"
                    Hidden1.Value = "0"
                    lblmsg1.Text = ""
                End If
                If _Ds.Tables.Count > 1 Then
                    If _Ds.Tables(1).Rows.Count > 0 Then
                        hid.Value = CType(_Ds.Tables(1).Rows(0).Item("email"), String)
                    End If
                End If

                'added by Geeta : Marathi payslip("60")
                If DdlreportType.SelectedValue.ToString = "43" Or DdlreportType.SelectedValue.ToString = "49" Or DdlreportType.SelectedValue.ToString = "51" _
                    Or DdlreportType.SelectedValue.ToString = "55" Or Convert.ToString(DdlreportType.SelectedValue).Equals("56") _
                    Or Convert.ToString(DdlreportType.SelectedValue).Equals("57") Or Convert.ToString(DdlreportType.SelectedValue).Equals("58") _
                    Or Convert.ToString(DdlreportType.SelectedValue).Equals("59") Or Convert.ToString(DdlreportType.SelectedValue).Equals("60") Then
                    BtnSend.Visible = False
                    btnSave.Visible = True
                    Tr2.Visible = False
                    BtnSendCCBCC.Visible = False
                Else
                    BtnSend.Visible = True
                    btnSave.Visible = True
                    Tr2.Visible = True
                    BtnSendCCBCC.Visible = True
                    If Convert.ToString(DdlreportType.SelectedValue).Equals("62") Or Convert.ToString(DdlreportType.SelectedValue).Equals("63") Then
                        BtnSend.Visible = False
                        BtnSendCCBCC.Visible = False
                        BtnPublishGrpBy.Visible = False
                        BtnLog.Visible = False
                    ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("64") Then
                        BtnSendCCBCC.Visible = False
                        BtnPublishGrpBy.Visible = False
                        BtnLog.Visible = False
                    ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("65") Then
                        BtnSendCCBCC.Visible = False
                        BtnPublishGrpBy.Visible = False
                        BtnLog.Visible = False
                        trselall.Style("display") = "none"
                    ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("66") Then
                        BtnSendCCBCC.Visible = False
                        BtnPublishGrpBy.Visible = False
                        BtnLog.Visible = False
                        trselall.Style("display") = "none"
                    ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("68") Then
                        BtnSendCCBCC.Visible = False
                        BtnPublishGrpBy.Visible = False
                        BtnLog.Visible = False
                        trselall.Style("display") = "none"
                    ElseIf DdlreportType.SelectedValue.Equals("74") Then
                        BtnPublishGrpBy.Visible = False
                    End If
                End If
                If ddlRepIn.SelectedValue.ToUpper = "H" Then
                    If DdlreportType.SelectedValue.Equals("74") Then
                        BtnSend.Visible = False
                    Else
                        BtnSend.Visible = True
                    End If

                End If
                _objCommon.ShowMessage(_msg)
            Catch ex As Exception
                _objcommonExp.PublishError("Error in Search Records(populateDgPayslip())", ex)
            End Try
        End Sub
        Private Sub SendInvestmentDetials()
            Try
                Dim ds As DataSet, _count As Integer = 0, _msg As New List(Of PayrollUtility.UserMessage)
                HidYear.Value = Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4)
                ds = ReturnInvestmentDeclaration()
                If ds.Tables.Count > 0 Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        DgPayslip.DataSource = ds.Tables(0)
                        DgPayslip.DataBind()
                        If DgPayslip.Items.Count > 2000 Then
                            DgPayslip.Columns(2).Visible = False
                            DgPayslip.Columns(3).Visible = False
                            DgPayslip.Columns(4).Visible = False
                            DgPayslip.Columns(5).Visible = False
                            DgPayslip.Columns(6).Visible = False
                            DgPayslip.Columns(7).Visible = False
                            DgPayslip.Columns(8).Visible = False
                        ElseIf DgPayslip.Items.Count > 500 Then
                            DgPayslip.Columns(2).Visible = True
                            DgPayslip.Columns(3).Visible = False
                            DgPayslip.Columns(4).Visible = False
                            DgPayslip.Columns(5).Visible = False
                            DgPayslip.Columns(6).Visible = False
                            DgPayslip.Columns(7).Visible = False
                            DgPayslip.Columns(8).Visible = False
                        Else
                            DgPayslip.Columns(2).Visible = True
                            DgPayslip.Columns(3).Visible = True
                            DgPayslip.Columns(4).Visible = True
                            DgPayslip.Columns(5).Visible = True
                            DgPayslip.Columns(6).Visible = True
                            DgPayslip.Columns(7).Visible = True
                            DgPayslip.Columns(8).Visible = True
                        End If
                        DgPayslip.Columns(6).Visible = False
                        DgPayslip.Columns(7).Visible = False
                        For _count = 0 To DgPayslip.Items.Count - 1
                            If CType(DgPayslip.Items(_count).Cells(5).Text.Trim.ToString, String) = "N/A" Then
                                DgPayslip.Items(_count).BackColor = Color.Pink
                                CType(DgPayslip.Items(_count).FindControl("chkEmpHold"), CheckBox).Checked = False
                            End If
                        Next
                        tableshow.Style("display") = ""
                        TrDg.Style("display") = ""
                        Hidden1.Value = DgPayslip.Items.Count.ToString
                        lblmsg1.Text = "Total " & DgPayslip.Items.Count.ToString & " Record(s) Found ! "
                    Else
                        _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "No record(s) Found according to the selection criteria !"})
                        tableshow.Style("display") = "none"
                    End If
                Else
                    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "No record(s) Found according to the selection criteria !"})
                    tableshow.Style("display") = "none"
                End If
                If ds.Tables.Count > 1 Then
                    If ds.Tables(1).Rows.Count > 0 Then
                        hid.Value = CType(ds.Tables(1).Rows(0).Item("email").ToString, String)
                    End If
                End If
                _objCommon.ShowMessage(_msg)
            Catch ex As Exception
                _objcommonExp.PublishError("Error in Search Records(SendInvestmentDetials())", ex)
            End Try
        End Sub

        'Added by Quadir on 27 Nov 2020 for changing logic Payslip Publish Mode on Search
        Private Function ReturnDsSpecific(Optional ByRef MsgReturn As String = "", Optional ByRef Flag As String = "", Optional ByRef EmployeeCodes As String = "") As DataSet
            Dim _DsNS As New DataSet, arrparam(24) As SqlClient.SqlParameter, EmpCodes As String = ""

            'Added by Quadir on 24 Nov 2020 for changing logic Payslip Publish Mode on Search
            arrparam(0) = New SqlClient.SqlParameter("@Pk_Emp_Code", USearch.UCTextcode.Trim.ToString)
            arrparam(1) = New SqlClient.SqlParameter("@Name", USearch.UCTextname.ToString())
            arrparam(2) = New SqlClient.SqlParameter("@COC", _objCommon.nNz(USearch.UCddlcostcenter).ToString())
            arrparam(3) = New SqlClient.SqlParameter("@DEP", _objCommon.nNz(USearch.UCddldept).ToString())
            arrparam(4) = New SqlClient.SqlParameter("@GRD", _objCommon.nNz(USearch.UCddlgrade).ToString())
            arrparam(5) = New SqlClient.SqlParameter("@DES", _objCommon.nNz(USearch.UCddldesig).ToString())
            arrparam(6) = New SqlClient.SqlParameter("@LOC", _objCommon.nNz(USearch.UCddllocation).ToString())
            arrparam(7) = New SqlClient.SqlParameter("@UNT", _objCommon.nNz(USearch.UCddlunit).ToString())
            arrparam(8) = New SqlClient.SqlParameter("@Salaried", _objCommon.nNz(USearch.UCddlsalbasis).ToString())
            arrparam(9) = New SqlClient.SqlParameter("@LVL", _objCommon.nNz(USearch.UCddllevel).ToString())
            arrparam(10) = New SqlClient.SqlParameter("@EmpStatus", _objCommon.nNz(USearch.UCddlEmp).ToString)
            arrparam(11) = New SqlClient.SqlParameter("@UserGroup", Session("UGroup").ToString)
            arrparam(12) = New SqlClient.SqlParameter("@Month", ddlMonthYear.SelectedValue.ToString)
            arrparam(13) = New SqlClient.SqlParameter("@Year", Right(ddlMonthYear.SelectedItem.ToString, 4))
            arrparam(14) = New SqlClient.SqlParameter("@Msg", SqlDbType.VarChar, 8000)
            arrparam(14).Direction = ParameterDirection.InputOutput
            arrparam(15) = New SqlClient.SqlParameter("@EmailExist", ddlEMailExist.SelectedValue.ToString)
            arrparam(16) = New SqlClient.SqlParameter("@SalProcessed", ddlSalaryProcess.SelectedValue.ToString)
            arrparam(17) = New SqlClient.SqlParameter("@EmailSent", ddlEMailSend.SelectedValue.ToString)
            arrparam(18) = New SqlClient.SqlParameter("@SalonHold", ddlSalWithHeld.SelectedValue.ToString)
            arrparam(19) = New SqlClient.SqlParameter("@HoldType", ddlshowsal.SelectedValue.ToString)
            arrparam(20) = New SqlClient.SqlParameter("@RepType", DdlreportType.SelectedValue.ToString)
            arrparam(21) = New SqlClient.SqlParameter("@Reppermail", IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString)
            arrparam(22) = New SqlClient.SqlParameter("@Date", ddloffcycledt.SelectedValue.Trim)
            arrparam(23) = New SqlClient.SqlParameter("@EmpCodes", EmpCodes.ToString)
            arrparam(24) = New SqlClient.SqlParameter("@UserID", Session("UID").ToString)
            _DsNS = _ObjData.GetDataSetProc("PaySp_Rpt_Sel_EmpPaySlip", arrparam)

            Return _DsNS
        End Function

        Private Function ReturnDsSearch(Optional ByRef MsgReturn As String = "", Optional ByRef Flag As String = "", Optional ByRef EmployeeCodes As String = "", Optional ByRef slipWOPWD As String = "") As DataSet
            Dim _Ds As New DataSet, arrparam(24) As SqlClient.SqlParameter, i As Integer, EmpCodes As String = "",
        _dst As New DataSet, _DtRow As DataRow = Nothing, _FilePath As String, fileName As String = "", PdfEmp As String = "", EmpCodesCheck As String = "", _arrEmpDet() As String, Ext As String = ""
            Dim _DsNS As New DataSet, _StrEmpCode As String = "", _dRowDoc As DataRow = Nothing, _dRow As DataRow,
        _RecCount As Integer = 0, _Counter As Integer, EmpCode As String = "", _dt As New DataTable
            Dim EmpCodeEntered As String = USearch.UCTextcode.Trim.ToString
            If Flag.ToString.ToUpper = "PDF" Then
                'Added by Quadir on 24 Nov 2020 for changing logic Payslip Publish Mode on Search
                If EmployeeCodes <> "" Then
                    _arrEmpDet = Split(EmployeeCodes, ",")
                Else
                    _arrEmpDet = Split(EmpCodeEntered, ",")
                End If

                _dst.Tables.Add("Table1")
                _dst.Tables(0).Columns.Add(New DataColumn("EmpCode"))
                _FilePath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Left(MonthName(CType(ddlMonthYear.SelectedValue, Integer)), 3) & Right(ddlMonthYear.SelectedItem.ToString, 4) & "\TaxPaySlip\"
                'Added by Quadir on 14 OCT 2020- Payslip Publish Overwrite Mode
                If rbtSlipPubMode.SelectedValue.ToString.ToUpper <> "O" And slipWOPWD <> "W" Then
                    If Directory.Exists(_FilePath) Then
                        If i = 0 Then
                            For Each filepaths As String In Directory.GetFiles(_FilePath)
                                _DtRow = _dst.Tables(0).NewRow
                                fileName = Path.GetFileName(filepaths)
                                Ext = Path.GetExtension(filepaths)
                                If Ext.ToString.ToUpper = ".PDF" Then
                                    EmpCodes = Left(fileName, fileName.Length - 29)

                                    'Added by Quadir on 24 Nov 2020 for changing logic Payslip Publish Mode on Search
                                    If RblNoSearch.SelectedValue = "S" Then
                                        For et As Integer = 0 To _arrEmpDet.Length - 1
                                            If _arrEmpDet(et) = EmpCodes Then
                                                PdfEmp = PdfEmp + "," + EmpCodes.ToString
                                                _DtRow(0) = _objCommon.nNz(EmpCodes)
                                                _dst.Tables(0).Rows.Add(_DtRow)
                                            End If
                                        Next
                                    Else
                                        PdfEmp = PdfEmp + "," + EmpCodes.ToString
                                        _DtRow(0) = _objCommon.nNz(EmpCodes)
                                        _dst.Tables(0).Rows.Add(_DtRow)
                                    End If

                                End If
                            Next
                        End If

                        EmpCodes = ""
                        EmpCodes = _dst.GetXml()
                        _dst.Clear()
                        _dst.Dispose()
                    End If
                Else
                    If Directory.Exists(_FilePath) Then
                        'Added by Quadir on 27 Nov 2020 for changing logic Payslip Publish Mode on Search
                        If EmpCodeEntered = "" And RblNoSearch.SelectedValue <> "S" Then
                            _DsNS = ReturnDsSpecific()
                            _dt.Columns.Add(New DataColumn("EmpCode"))
                            _Ds.Tables.Add(_dt)
                            For _Counter = 0 To _DsNS.Tables(0).Rows.Count - 1
                                PK_emp_code = _DsNS.Tables(0).Rows(_Counter)("fk_emp_code").ToString
                                _dRow = _Ds.Tables(0).NewRow
                                _dRow(0) = PK_emp_code
                                _Ds.Tables(0).Rows.Add(_dRow)
                                EmpCode = EmpCode + PK_emp_code + ","
                                _RecCount = _RecCount + 1
                            Next
                            _arrEmpDet = Split(EmpCode, ",")
                        End If

                        For Each filepaths As String In Directory.GetFiles(_FilePath)
                            fileName = Path.GetFileName(filepaths)
                            EmpCodesCheck = Left(fileName, fileName.Length - 29)
                            For et As Integer = 0 To _arrEmpDet.Length - 1
                                If _arrEmpDet(et) = EmpCodesCheck Then
                                    File.Delete(filepaths)
                                Else
                                    'Added by Quadir on 24 Nov 2020 for changing logic Payslip Publish Mode on Search
                                    If RblNoSearch.SelectedValue <> "S" Then
                                        If EmpCodeEntered = "" And EmpCode <> "" Then
                                            If _arrEmpDet(et) = EmpCode Then
                                                File.Delete(filepaths)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                End If
            ElseIf Flag.ToUpper.Equals("SPDF") Then
                If EmployeeCodes <> "" Then
                    _arrEmpDet = Split(EmployeeCodes, ",")
                Else
                    _arrEmpDet = Split(EmpCodeEntered, ",")
                End If

                If DdlreportType.SelectedValue.Equals("S") Then
                    _FilePath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Left(MonthName(CType(ddlMonthYear.SelectedValue, Integer)), 3) & Right(ddlMonthYear.SelectedItem.ToString, 4) & "\LeaveWoPaySlip\"
                Else
                    'If DdlreportType.SelectedValue.Equals("57") Then
                    _FilePath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Left(MonthName(CType(ddlMonthYear.SelectedValue, Integer)), 3) & Right(ddlMonthYear.SelectedItem.ToString, 4) & "\YTDTaxComputationSheet\"
                End If

                If Directory.Exists(_FilePath) Then
                    If EmpCodeEntered = "" And RblNoSearch.SelectedValue <> "S" Then
                        _DsNS = ReturnDsSpecific()
                        _dt.Columns.Add(New DataColumn("EmpCode"))
                        _Ds.Tables.Add(_dt)
                        For _Counter = 0 To _DsNS.Tables(0).Rows.Count - 1
                            PK_emp_code = _DsNS.Tables(0).Rows(_Counter)("fk_emp_code").ToString
                            _dRow = _Ds.Tables(0).NewRow
                            _dRow(0) = PK_emp_code
                            _Ds.Tables(0).Rows.Add(_dRow)
                            EmpCode = EmpCode + PK_emp_code + ","
                            _RecCount = _RecCount + 1
                        Next
                        _arrEmpDet = Split(EmpCode, ",")
                    End If

                    For Each filepaths As String In Directory.GetFiles(_FilePath)
                        fileName = Path.GetFileName(filepaths)
                        If DdlreportType.SelectedValue.Equals("S") Then
                            EmpCodesCheck = Left(fileName, fileName.Length - 22)
                        Else
                            EmpCodesCheck = Left(fileName, fileName.Length - 17)
                        End If

                        For et As Integer = 0 To _arrEmpDet.Length - 1
                            If _arrEmpDet(et) = EmpCodesCheck Then
                                File.Delete(filepaths)
                            Else
                                If RblNoSearch.SelectedValue <> "S" Then
                                    If EmpCodeEntered = "" And EmpCode <> "" Then
                                        If _arrEmpDet(et) = EmpCode Then
                                            File.Delete(filepaths)
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Next
                End If
            End If
            'Added by Quadir on 24 Nov 2020 for changing logic Payslip Publish Mode on Search
            If EmployeeCodes <> "" Then
                arrparam(0) = New SqlClient.SqlParameter("@Pk_Emp_Code", EmployeeCodes.Trim.ToString)
            Else
                arrparam(0) = New SqlClient.SqlParameter("@Pk_Emp_Code", USearch.UCTextcode.Trim.ToString)
            End If
            arrparam(1) = New SqlClient.SqlParameter("@Name", USearch.UCTextname.ToString())
            arrparam(2) = New SqlClient.SqlParameter("@COC", _objCommon.nNz(USearch.UCddlcostcenter).ToString())
            arrparam(3) = New SqlClient.SqlParameter("@DEP", _objCommon.nNz(USearch.UCddldept).ToString())
            arrparam(4) = New SqlClient.SqlParameter("@GRD", _objCommon.nNz(USearch.UCddlgrade).ToString())
            arrparam(5) = New SqlClient.SqlParameter("@DES", _objCommon.nNz(USearch.UCddldesig).ToString())
            arrparam(6) = New SqlClient.SqlParameter("@LOC", _objCommon.nNz(USearch.UCddllocation).ToString())
            arrparam(7) = New SqlClient.SqlParameter("@UNT", _objCommon.nNz(USearch.UCddlunit).ToString())
            arrparam(8) = New SqlClient.SqlParameter("@Salaried", _objCommon.nNz(USearch.UCddlsalbasis).ToString())
            arrparam(9) = New SqlClient.SqlParameter("@LVL", _objCommon.nNz(USearch.UCddllevel).ToString())
            arrparam(10) = New SqlClient.SqlParameter("@EmpStatus", _objCommon.nNz(USearch.UCddlEmp).ToString)
            arrparam(11) = New SqlClient.SqlParameter("@UserGroup", Session("UGroup").ToString)
            arrparam(12) = New SqlClient.SqlParameter("@Month", ddlMonthYear.SelectedValue.ToString)
            arrparam(13) = New SqlClient.SqlParameter("@Year", Right(ddlMonthYear.SelectedItem.ToString, 4))
            arrparam(14) = New SqlClient.SqlParameter("@Msg", SqlDbType.VarChar, 8000)
            arrparam(14).Direction = ParameterDirection.InputOutput
            arrparam(15) = New SqlClient.SqlParameter("@EmailExist", ddlEMailExist.SelectedValue.ToString)
            arrparam(16) = New SqlClient.SqlParameter("@SalProcessed", ddlSalaryProcess.SelectedValue.ToString)
            arrparam(17) = New SqlClient.SqlParameter("@EmailSent", ddlEMailSend.SelectedValue.ToString)
            arrparam(18) = New SqlClient.SqlParameter("@SalonHold", ddlSalWithHeld.SelectedValue.ToString)
            arrparam(19) = New SqlClient.SqlParameter("@HoldType", ddlshowsal.SelectedValue.ToString)
            arrparam(20) = New SqlClient.SqlParameter("@RepType", DdlreportType.SelectedValue.ToString)
            arrparam(21) = New SqlClient.SqlParameter("@Reppermail", IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString)
            arrparam(22) = New SqlClient.SqlParameter("@Date", ddloffcycledt.SelectedValue.Trim)
            arrparam(23) = New SqlClient.SqlParameter("@EmpCodes", EmpCodes.ToString)
            arrparam(24) = New SqlClient.SqlParameter("@UserID", Session("UID").ToString)
            _Ds = _ObjData.GetDataSetProc("PaySp_Rpt_Sel_EmpPaySlip", arrparam)
            If _Ds.Tables.Count > 2 Then
                If _Ds.Tables(2).Rows.Count > 0 Then
                    'If RblNoSearch.SelectedValue <> "S" Then
                    '    HidEmpPdf.Value = _Ds.Tables(2).Rows(0).Item("AllEmp").ToString + ","
                    'End If
                    'Added by Quadir on 24 Nov 2020 for changing logic Payslip Publish Mode on Search
                    If _Ds.Tables(2).Rows(0).Item("AllEmp").ToString <> "" Then
                        HidEmpPdf.Value = _Ds.Tables(2).Rows(0).Item("AllEmp").ToString + ","
                    End If

                End If
            End If
            If Not IsNothing(arrparam(14).Value) Then
                MsgReturn = arrparam(14).Value.ToString
            Else
                MsgReturn = ""
            End If
            EmpCodes = ""
            Return _Ds
        End Function
        Private Function ReturnInvestmentDeclaration(Optional ByRef MsgReturn As String = "") As DataSet
            Dim _Ds As New DataSet, arrparam(16) As SqlClient.SqlParameter
            arrparam(0) = New SqlClient.SqlParameter("@fk_Costcenter_code", _objCommon.nNz(USearch.UCddlcostcenter).ToString())
            arrparam(1) = New SqlClient.SqlParameter("@fk_loc_code", _objCommon.nNz(USearch.UCddllocation).ToString())
            arrparam(2) = New SqlClient.SqlParameter("@fk_unit", _objCommon.nNz(USearch.UCddlunit).ToString())
            arrparam(3) = New SqlClient.SqlParameter("@salaried", _objCommon.nNz(USearch.UCddlsalbasis).ToString())
            arrparam(4) = New SqlClient.SqlParameter("@pk_emp_code", _objCommon.nNz(USearch.UCTextcode.ToString))
            arrparam(5) = New SqlClient.SqlParameter("@fk_dept_code", _objCommon.nNz(USearch.UCddldept).ToString())
            arrparam(6) = New SqlClient.SqlParameter("@fk_desig_code", _objCommon.nNz(USearch.UCddldesig).ToString())
            arrparam(7) = New SqlClient.SqlParameter("@fk_grade_code", _objCommon.nNz(USearch.UCddlgrade).ToString())
            arrparam(8) = New SqlClient.SqlParameter("@fk_level_code", _objCommon.nNz(USearch.UCddllevel).ToString())
            arrparam(9) = New SqlClient.SqlParameter("@first_Name", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
            arrparam(10) = New SqlClient.SqlParameter("@lastname", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
            arrparam(11) = New SqlClient.SqlParameter("@UserGroup", Session("UGroup"))
            arrparam(12) = New SqlClient.SqlParameter("@EmpType", USearch.UCddlEmp.ToString)
            arrparam(13) = New SqlClient.SqlParameter("@Sfindate", _objCommon.nNz(CType(Session("Sfindate"), Date)))
            arrparam(14) = New SqlClient.SqlParameter("@Efindate", _objCommon.nNz(CType(Session("Efindate"), Date)))
            arrparam(15) = New SqlClient.SqlParameter("@month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
            arrparam(16) = New SqlClient.SqlParameter("@year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
            _Ds = _ObjData.GetDataSetProc("Paysp_Investment_RptMailSearch", arrparam)
            Return _Ds
        End Function
        Protected Sub btnresetsearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnresetsearch.Click
            Response.Redirect("RptMultipleSalarySlip.aspx")
        End Sub
        Protected Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender
            Dim objSecurityManager As New PayrollUtility.SecurityManager
            objSecurityManager.UserAccessintoForm("RptMultipleSalarySlip.aspx", btnPreview, , False)
        End Sub
        Private Sub Showreport()
            Try
                Dim dt As New DataTable, arrparam(0) As SqlClient.SqlParameter
                arrparam(0) = New SqlClient.SqlParameter("@id", DDLPaySlipType.SelectedValue)
                dt = _ObjData.GetDataTableProc("paysp_reportdetails", arrparam)
                If dt.Rows.Count > 0 Then
                    ViewState("SalaryData") = Nothing
                    If Not DDLPaySlipType.SelectedValue.Equals("67") Then
                        SalaryDataNew(dt)
                        trview.Style.Value = "display:"
                    End If
                Else
                    dt = Nothing
                    ViewState("SalaryData") = Nothing
                    SalaryDataNew(dt)
                    trview.Style.Value = "display:none"
                End If

            Catch ex As Exception
                _objcommonExp.PublishError("Salarydata())", ex)
            End Try
        End Sub
        'to populate the Report format
        Private Sub PopulateReportFormat()
            Try
                ddlformat.Items.Add(New ListItem("--Select Report Format--", ""))
                ddlformat.Items.Add(New ListItem("With Paycode", "0"))
                ddlformat.Items.Add(New ListItem("Without Paycode", "1"))
            Catch ex As Exception
                _objcommonExp.PublishError("Error in Populate the Report Format(PopulateReportFormat())", ex)
            End Try
        End Sub
        Protected Sub SalaryDataNew(Optional ByVal _DT As DataTable = Nothing)
            If IsNothing(ViewState("SalaryData")) Then
                Dim HtmlStr As New System.Text.StringBuilder, _cnt As Int32 = 0
                HtmlStr.Append("<Table width='100%' class='grid' cellspacing='0' cellpadding='0' border='0'>")
                'HtmlStr.Append("<Tr style='height:20px'>")
                'HtmlStr.Append("<th colspan = '12'>")
                'HtmlStr.Append("Options")
                'HtmlStr.Append("</th>")
                'HtmlStr.Append("</Tr>")
                For _cnt = 0 To _DT.Columns.Count - 1
                    If DDLPaySlipType.SelectedValue.ToString = "38" And _cnt = _DT.Columns.Count - 1 Then
                        HtmlStr.Append("<Tr>")
                        HtmlStr.Append("<Td colspan='12'>")
                        HtmlStr.Append("&nbsp;")
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("</Tr>")

                        HtmlStr.Append("<Tr>")
                        HtmlStr.Append("<Td  width='11%'>")
                        HtmlStr.Append(_DT.Columns(_cnt).ColumnName.ToString)
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("<Td  width='3%'>")
                        HtmlStr.Append(":")
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("<Td colspan='10'>")
                        HtmlStr.Append(_DT.Rows(0).Item(_cnt))
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("</Tr>")

                        HtmlStr.Append("<Tr>")
                        HtmlStr.Append("<Td colspan='12'>")
                        HtmlStr.Append("&nbsp;")
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("</Tr>")
                        Exit For
                    End If
                    HtmlStr.Append("<Tr>")
                    HtmlStr.Append("<Td  width='11%'>")
                    HtmlStr.Append(_DT.Columns(_cnt).ColumnName.ToString)
                    HtmlStr.Append("</Td>")
                    HtmlStr.Append("<Td  width='3%'>")
                    HtmlStr.Append(":")
                    HtmlStr.Append("</Td>")
                    HtmlStr.Append("<Td  align='left' width='11%'>")
                    HtmlStr.Append(_DT.Rows(0).Item(_cnt))
                    HtmlStr.Append("</Td>")
                    If _cnt < _DT.Columns.Count - 1 Then
                        _cnt = _cnt + 1
                        HtmlStr.Append("<Td  width='11%'>")
                        HtmlStr.Append(_DT.Columns(_cnt).ColumnName.ToString)
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("<Td  width='3%'>")
                        HtmlStr.Append(":")
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("<Td  align ='left' width='11%'>")
                        HtmlStr.Append(_DT.Rows(0).Item(_cnt))
                        HtmlStr.Append("</Td>")
                    Else
                        HtmlStr.Append("<Td  width='11%'>")
                        HtmlStr.Append("&nbsp;")
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("<Td  width='3%'>")
                        HtmlStr.Append("&nbsp;")
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("<Td  align ='left' width='11%'>")
                        HtmlStr.Append("&nbsp;")
                        HtmlStr.Append("</Td>")
                    End If
                    If _cnt < _DT.Columns.Count - 1 Then
                        _cnt = _cnt + 1
                        HtmlStr.Append("<Td  width='11%'>")
                        HtmlStr.Append(_DT.Columns(_cnt).ColumnName.ToString)
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("<Td  width='3%'>")
                        HtmlStr.Append(":")
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("<Td  align ='left' width='11%'>")
                        HtmlStr.Append(_DT.Rows(0).Item(_cnt))
                        HtmlStr.Append("</Td>")
                    Else
                        HtmlStr.Append("<Td  width='11%'>")
                        HtmlStr.Append("&nbsp;")
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("<Td  width='3%'>")
                        HtmlStr.Append("&nbsp;")
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("<Td  align ='right' width='11%'>")
                        HtmlStr.Append("&nbsp;")
                        HtmlStr.Append("</Td>")
                    End If
                    If _cnt < _DT.Columns.Count - 1 Then
                        _cnt = _cnt + 1
                        HtmlStr.Append("<Td  width='11%'>")
                        HtmlStr.Append(_DT.Columns(_cnt).ColumnName.ToString)
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("<Td  width='3%'>")
                        HtmlStr.Append(":")
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("<Td  align ='left' width='11%'>")
                        HtmlStr.Append(_DT.Rows(0).Item(_cnt))
                        HtmlStr.Append("</Td>")
                    Else
                        HtmlStr.Append("<Td  width='11%'>")
                        HtmlStr.Append("&nbsp;")
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("<Td  width='3%'>")
                        HtmlStr.Append("&nbsp;")
                        HtmlStr.Append("</Td>")
                        HtmlStr.Append("<Td  align ='right' width='11%'>")
                        HtmlStr.Append("&nbsp;")
                        HtmlStr.Append("</Td>")
                    End If
                    HtmlStr.Append("</Tr>")
                Next
                HtmlStr.Append("</Table>")
                ViewState("SalaryData") = HtmlStr.ToString
            End If
        End Sub
        'Added by Rajesh for Luxor Register on 10 oct 13
        Protected Sub populatestrddl1(ByVal ddl As DropDownList)
            ddl.Items.Add(New ListItem("Department", "DET"))
        End Sub
        Protected Sub DDLPaySlipType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDLPaySlipType.SelectedIndexChanged
            lblProcessBarMsg.Text = ""
            Showreport()
            Try
                trEncrType.Style.Value = "display:none"
                trSftpID.Style.Value = "display:none"
                trFileName.Style.Value = "display:none"
                trformat.Style.Value = "display:none"
                txtrptName.Text = ""
                chkSFTP.Checked = False
                ddllEncrType.SelectedIndex = -1
                ddlrepformat.SelectedIndex = -1
                btnExport2CSV.Visible = False
                If Convert.ToString(DDLPaySlipType.SelectedValue).Trim.Equals("67") Then
                    trling.Style("display") = ""
                Else
                    trling.Style("display") = "none"
                End If
                Me.lblmsg2.Text = ""
                If DDLPaySlipType.SelectedValue.ToString = "49" Then
                    trHelp.Style("display") = ""
                Else
                    trHelp.Style("display") = "none"
                End If
                If DDLPaySlipType.SelectedValue = "5" Then
                    tblsection.Style.Value = "display:"
                Else
                    tblsection.Style.Value = "display:none"
                End If
                If DDLPaySlipType.SelectedValue = "6" Then
                    trrepformat.Style.Value = "display:"
                Else
                    trrepformat.Style.Value = "display:none"
                End If
                'here we checl the report type and according to the report type we show/Hide the group2 and group3 selection
                If DDLPaySlipType.SelectedValue = "16" Then
                    populatestrddl(ddlGroup2)
                    populatestrddl(ddlGroup3)
                    trGroupBY.Style.Value = "display:"
                    trsortbasis2.Style.Value = "display:"
                    trsortbasis3.Style.Value = "display:"
                Else
                    populatestrddl(ddlGroup2)
                    populatestrddl(ddlGroup3)
                    trGroupBY.Style.Value = "display:none"
                    trsortbasis2.Style.Value = "display:none"
                    trsortbasis3.Style.Value = "display:none"
                End If
                If DDLPaySlipType.SelectedValue = "5" Then
                    tblsection.Style.Value = "display:"
                    SectionSel()
                Else
                    tblsection.Style.Value = "display:none"
                End If
                'to populate the paycodes dropdownlist if report type is salary register
                If DDLPaySlipType.SelectedValue = "8" Or DDLPaySlipType.SelectedValue = "9" Or DDLPaySlipType.SelectedValue = "13" Or DDLPaySlipType.SelectedValue = "15" Or DDLPaySlipType.SelectedValue = "16" Or DDLPaySlipType.SelectedValue = "17" Or DDLPaySlipType.SelectedValue = "27" Or DDLPaySlipType.SelectedValue = "19" Or DDLPaySlipType.SelectedValue = "47" Or DDLPaySlipType.SelectedValue = "48" Or DDLPaySlipType.SelectedValue = "22" Or DDLPaySlipType.SelectedValue = "29" Or DDLPaySlipType.SelectedValue = "30" Or DDLPaySlipType.SelectedValue = "31" Or DDLPaySlipType.SelectedValue = "39" Or DDLPaySlipType.SelectedValue = "46" Then
                    tblpaycode.Style.Value = "display:"
                    'added by Rajesh for show help msg when report type REGISTER OF WAGES (FORM-XVII)
                    If DDLPaySlipType.SelectedValue = "39" Then
                        tblhelp.Style.Value = "display:"
                    Else
                        tblhelp.Style.Value = "display:none"
                    End If
                    If DDLPaySlipType.SelectedValue = "15" Or DDLPaySlipType.SelectedValue = "19" Or DDLPaySlipType.SelectedValue = "47" Or DDLPaySlipType.SelectedValue = "48" Then
                        trsortbasis.Style.Value = "display:"
                        If DDLPaySlipType.SelectedValue = "47" Then
                            trGroupBY.Style.Value = "display:none"
                        Else
                            trGroupBY.Style.Value = "display:"
                        End If
                        TRDIV.Style.Value = "display:"
                    Else
                        If DDLPaySlipType.SelectedValue = "16" Or DDLPaySlipType.SelectedValue = "9" Or DDLPaySlipType.SelectedValue = "13" Then
                            trsortbasis.Style.Value = "display:"
                            trGroupBY.Style.Value = "display:"
                            TRDIV.Style.Value = "display:"
                        Else
                            trsortbasis.Style.Value = "display:none"
                            trGroupBY.Style.Value = "display:none"
                            TRDIV.Style.Value = "display:none"
                        End If
                    End If
                    'Added by Rajesh for Luxor Register on 10 oct 13
                    If DDLPaySlipType.SelectedValue = "46" Then
                        trsortbasis.Style.Value = "display:"
                        trGroupBY.Style.Value = "display:"
                        TRDIV.Style.Value = "display:"
                        ddlshortbasis.Items.Clear()
                        populatestrddl1(ddlshortbasis)
                        ddlshortbasis.Enabled = False
                        tblpaycode.Style.Value = "display:none"
                    Else
                        populatestrddl(ddlshortbasis)
                        ddlshortbasis.Enabled = True
                    End If

                    If DDLPaySlipType.SelectedValue = "22" Then
                        tblpaycode.Style.Value = "display:"
                    Else
                    End If
                    Bind_Paycode_Check()
                Else
                    tblpaycode.Style.Value = "display:none"
                End If
                EnableUSearchDdl()
                If DDLPaySlipType.SelectedValue = "25" Then

                    TblReimb.Style.Value = "display:"
                    Bind_Paycode_Check()
                Else
                    TblReimb.Style.Value = "display:none"
                End If
                If DDLPaySlipType.SelectedValue = "31" Then
                    lblmsg2.Text = "This is a client specific report. In this report value of processed arrear and processed loan does not get publish, only selected paycode are published"
                End If
                If DDLPaySlipType.SelectedValue = "" Then
                    trview.Style.Value = "display:none"
                End If
                '----Added By Geeta on 30 Aug 2012
                If DDLPaySlipType.SelectedValue = "17" Or DDLPaySlipType.SelectedValue = "27" Then
                    trview.Style.Value = "display:none"
                    tblpaycode.Style.Value = "display:none"
                End If
                'Added by Rohtas Singh on 06 Dec 2017
                If DDLPaySlipType.SelectedValue = "54" Then
                    trRptformat.Style.Value = "display:"
                    lblmsg2.Text = "This is a custom made report."
                Else
                    trRptformat.Style.Value = "display:none"
                End If

                If DDLPaySlipType.SelectedValue.ToString = "55" And rbtnslip.Checked.Equals(True) Then
                    ddlshowsal.Enabled = False
                Else
                    ddlshowsal.Enabled = True
                End If

                If DDLPaySlipType.SelectedValue.ToString = "62" Then
                    troffcycle.Style.Value = "display:"
                    populateFromdate("H")
                Else
                    troffcycle.Style.Value = "display:none"
                    btnPreview.Enabled = True
                End If

                'Modified by Vishal Chauhan on 01 Feb 2025
                If DDLPaySlipType.SelectedValue.ToString = "21" Then
                    trShowClr.Style.Value = "display:"
                    btnPreview.Text = "Export to Excel"
                    btnPreview.Width = "100"

                    Dim IsNewUrl As String
                    IsNewUrl = IsRptapiConfigured4CSV("PaySP_SalaryRegInExcel")
                    If IsNewUrl IsNot Nothing AndAlso IsNewUrl.ToUpper.Trim = "Y" Then
                        btnExport2CSV.Visible = True
                    Else
                        btnExport2CSV.Visible = False
                    End If
                ElseIf DDLPaySlipType.SelectedValue.ToString = "38" Then
                    trShowClr.Style.Value = "display:"

                    trEncrType.Style.Value = "display:none"
                    trSftpID.Style.Value = "display:none"
                    trFileName.Style.Value = "display:none"
                    trformat.Style.Value = "display:"
                    txtrptName.Text = ""
                    chkSFTP.Checked = False
                    ddllEncrType.SelectedIndex = -1
                    ddlrepformat.SelectedIndex = -1
                    btnPreview.Text = "Export to Excel"
                    btnPreview.Width = "100"
                    btnExport2CSV.Visible = False
                    If ddlrepformat.SelectedValue.ToUpper.Equals("CSV") Then
                        btnPreview.Text = "Export to CSV"
                        trEncrType.Style.Value = "display:"
                        trFileName.Style.Value = "display:"
                    End If
                    ShowSFTP()
                Else
                    trShowClr.Style.Value = "display:none"
                    btnPreview.Text = "Preview"
                    btnPreview.Width = "80"
                    btnExport2CSV.Visible = False
                End If
            Catch ex As Exception
                _objcommonExp.PublishError("Error in ddlreporttype_SelectedIndexChanged()", ex)
            End Try
        End Sub
        Private Function IsRptapiConfigured4CSV(ByVal ProcName As String) As String
            Try
                Dim sParam(1) As SqlClient.SqlParameter
                sParam(0) = New SqlClient.SqlParameter("@SP_Name", ProcName)
                sParam(1) = New SqlClient.SqlParameter("@IsNewURL", SqlDbType.VarChar, 1)
                sParam(1).Direction = ParameterDirection.Output
                '_ObjData.ExecuteStoredProcMsg("PaySP_ReportAPI_ConfigSel", sParam)
                'Return sParam(1).Value.ToString
                Dim dt As DataTable = _ObjData.GetDataTableProc("PaySP_ReportAPI_ConfigSel", sParam)
                If dt.Rows.Count = 0 OrElse dt.Rows(0)("isNewURL").ToString.ToUpper.Trim() <> "Y" Then
                    Return "N"
                End If
                If dt.Rows.Count > 0 AndAlso dt.Rows(0)("WithoutProcessbar").ToString.ToUpper.Trim() = "Y" Then
                    Return "N"
                Else
                    Return "Y"
                End If
            Catch ex As Exception
                Return "N"
            End Try
        End Function
        'To populate the State Code
        Private Sub populateFromdate(ByVal str As String)
            Try
                'LogMessage.log.Debug("Start: Popluate Date in drop down (populateFromdate)")
                ddlmonthyearS.Items.Clear()
                ddlmonthyearS.Items.Add(New ListItem("-- Select All --", ""))
                Dim arrparam(1) As SqlParameter, dt As New DataTable
                arrparam(0) = New SqlParameter("@month", ddlMonthYear.SelectedValue.ToString)
                arrparam(1) = New SqlParameter("@year", Right(ddlMonthYear.SelectedItem.Text.Trim, 4).ToString) 'Sandeep
                ' LogMessage.log.Debug(objLogMessage.MessagingFormat("@month:" & Convert.ToString(ddlmonthyear.SelectedValue) & "@year:" & Convert.ToString(Right(ddlmonthyear.SelectedItem.Text.Trim, 4)), , "Bind Date", "Paysp_TrnEmpMissPaycodeDataoffcycle_DDL"))
                dt = _ObjData.GetDataTableProc("Paysp_TrnEmpMissPaycodeDataoffcycle_DDL", arrparam)
                If dt.Rows.Count > 0 Then
                    If str.Equals("P") Then
                        ddloffcycledt.DataSource = dt
                        ddloffcycledt.DataTextField = "start_Date"
                        ddloffcycledt.DataValueField = "start_Date"
                        ddloffcycledt.DataBind()
                    Else
                        ddlmonthyearS.DataSource = dt
                        ddlmonthyearS.DataTextField = "start_Date"
                        ddlmonthyearS.DataValueField = "start_Date"
                        ddlmonthyearS.DataBind()
                    End If


                    btnPreview.Enabled = True
                Else
                    btnPreview.Enabled = False
                    lblmsg2.Text = "No offcycle Payment data found for the selected month!"
                End If
                'LogMessage.log.Debug("End: Popluate Date in drop down (populateFromdate)")
            Catch Ex As Exception
                'LogMessage.log.Error("populateFromdate()", Ex)
                _objcommonExp.PublishError("To bind the offcycle date(populateFromdate)", Ex)
            End Try
        End Sub
        Private Sub Bind_Paycode_Check()
            Try
                lit.Text = ""
                Dim ds As New DataSet, lstitem As ListItem = Nothing, lstitemReimb As ListItem = Nothing
                'for execute store procedure
                ds = _ObjData.GetDataSetProc("sp_sel_Check_Box_salreg")
                'for check the existance of records
                If ds.Tables(0).Rows.Count > 0 Then
                    HidDed.Value = CType(ds.Tables(0).Rows.Count, String)
                    cbldeduction.DataTextField = "Pay_desc"
                    cbldeduction.DataValueField = "pk_pay_code"
                    cbldeduction.DataSource = ds.Tables(0)
                    cbldeduction.DataBind()
                    'Changed by Rohtas Singh on 16 Dec 2008
                    'for select all checkboxes of checkbox list on first time
                    For Each lstitem In cbldeduction.Items
                        lstitem.Selected = True
                    Next
                End If
                'for check the existance of record
                If ds.Tables(1).Rows.Count > 0 Then
                    HidAdd.Value = CType(ds.Tables(1).Rows.Count, String)
                    cbladd.DataTextField = "Pay_desc"
                    cbladd.DataValueField = "pk_pay_code"
                    cbladd.DataSource = ds.Tables(1)
                    cbladd.DataBind()
                    'Changed by Rohtas Singh on 16 Dec 2008
                    'for select all checkboxes of checkbox list on first time
                    For Each lstitem In cbladd.Items
                        lstitem.Selected = True
                    Next
                End If
                'By Dharmendra Rawat [30 Jun 2010]
                'for check the reimbursement type paycode check box.
                If TblReimb.Style.Value = "display:" Then
                    If ds.Tables(2).Rows.Count > 0 Then
                        HidReimb.Value = CType(ds.Tables(2).Rows.Count, String)
                        cblReimb.DataTextField = "Pay_desc"
                        cblReimb.DataValueField = "pk_pay_code"
                        cblReimb.DataSource = ds.Tables(2)
                        cblReimb.DataBind()
                        ''for select all checkboxes of checkbox list on first time
                        For Each lstitemReimb In cblReimb.Items
                            lstitemReimb.Selected = False
                        Next
                    End If
                End If
            Catch ex As Exception
                _objcommonExp.PublishError("Error in Bind_Paycode_Check()", ex)
            End Try
        End Sub
        Private Sub populatestrddl(ByRef _ddl As DropDownList)
            Try
                _ddl.Items.Clear() ' Added by Jay on 11 Mar 2014
                If UCase(_ddl.ID.ToString) = UCase("ddlshortbasis") Then
                    _ddl.Items.Add(New ListItem(" -- Select All -- ", ""))
                    _ddl.Items.Add(New ListItem("Cost Center", "COC"))
                    _ddl.Items.Add(New ListItem("Location", "LOC"))
                    _ddl.Items.Add(New ListItem("Unit", "UNT"))
                    _ddl.Items.Add(New ListItem("Department", "DET"))
                    _ddl.Items.Add(New ListItem("Designation", "DES"))
                    _ddl.Items.Add(New ListItem("Grade", "GRD"))
                    _ddl.Items.Add(New ListItem("Level", "LVL"))
                ElseIf UCase(_ddl.ID.ToString) = UCase("ddlGroup2") Then
                    If ddlshortbasis.SelectedValue.ToString = "" Then
                        _ddl.Items.Clear()
                        _ddl.Items.Add(New ListItem(" -- Select All -- ", ""))
                    Else
                        _ddl.Items.Clear()
                        _ddl.Items.Add(New ListItem(" -- Select All -- ", ""))
                        _ddl.Items.Add(New ListItem("Cost Center", "COC"))
                        _ddl.Items.Add(New ListItem("Location", "LOC"))
                        _ddl.Items.Add(New ListItem("Unit", "UNT"))
                        _ddl.Items.Add(New ListItem("Department", "DET"))
                        _ddl.Items.Add(New ListItem("Sub Department", "SDET"))
                        _ddl.Items.Add(New ListItem("Designation", "DES"))
                        _ddl.Items.Add(New ListItem("Grade", "GRD"))
                        _ddl.Items.Add(New ListItem("Level", "LVL"))

                        If ddlshortbasis.SelectedValue.ToString = "LOC" Then
                            Dim _li As ListItem = ddlshortbasis.Items.FindByValue("COC")

                            _li = ddlshortbasis.Items.FindByValue("LOC")
                            _ddl.Items.Remove(_li)


                        ElseIf ddlshortbasis.SelectedValue.ToString = "UNT" Then
                            Dim _li As ListItem = ddlshortbasis.Items.FindByValue("COC")

                            _li = ddlshortbasis.Items.FindByValue("LOC")
                            _ddl.Items.Remove(_li)
                            _li = ddlshortbasis.Items.FindByValue("UNT")
                            _ddl.Items.Remove(_li)

                        Else
                            Dim _li As ListItem = ddlshortbasis.Items.FindByValue(ddlshortbasis.SelectedValue.ToString)
                            _ddl.Items.Remove(_li)

                        End If
                    End If
                Else
                    If ddlGroup2.SelectedValue.ToString = "" Then
                        _ddl.Items.Clear()
                        _ddl.Items.Add(New ListItem(" -- Select All -- ", ""))
                    Else
                        _ddl.Items.Clear()
                        _ddl.Items.Add(New ListItem(" -- Select All -- ", ""))
                        _ddl.Items.Add(New ListItem("Cost Center", "COC"))
                        _ddl.Items.Add(New ListItem("Location", "LOC"))
                        _ddl.Items.Add(New ListItem("Unit", "UNT"))
                        _ddl.Items.Add(New ListItem("Department", "DET"))
                        _ddl.Items.Add(New ListItem("Sub Department", "SDET"))
                        _ddl.Items.Add(New ListItem("Designation", "DES"))
                        _ddl.Items.Add(New ListItem("Grade", "GRD"))
                        _ddl.Items.Add(New ListItem("Level", "LVL"))

                        If ddlshortbasis.SelectedValue.ToString = "LOC" Then
                            Dim _li As ListItem = ddlshortbasis.Items.FindByValue("COC")
                            _ddl.Items.Remove(_li)
                            _li = ddlshortbasis.Items.FindByValue("LOC")

                        ElseIf ddlshortbasis.SelectedValue.ToString = "UNT" Then
                            Dim _li As ListItem = ddlshortbasis.Items.FindByValue("COC")
                            _ddl.Items.Remove(_li)
                            _li = ddlshortbasis.Items.FindByValue("LOC")

                            _li = ddlshortbasis.Items.FindByValue("UNT")

                        Else
                            Dim _li As ListItem = ddlshortbasis.Items.FindByValue(ddlshortbasis.SelectedValue.ToString)
                            _ddl.Items.Remove(_li)

                        End If

                        If ddlGroup2.SelectedValue.ToString = "LOC" Then
                            Dim _li As ListItem = ddlGroup2.Items.FindByValue("COC")

                            _li = ddlGroup2.Items.FindByValue("LOC")
                            _ddl.Items.Remove(_li)

                        ElseIf ddlGroup2.SelectedValue.ToString = "UNT" Then
                            Dim _li As ListItem = ddlGroup2.Items.FindByValue("COC")

                            _li = ddlGroup2.Items.FindByValue("LOC")
                            _ddl.Items.Remove(_li)
                            _li = ddlGroup2.Items.FindByValue("UNT")
                            _ddl.Items.Remove(_li)

                        ElseIf ddlGroup2.SelectedValue.ToString = "SDET" Then
                            Dim _li As ListItem = ddlGroup2.Items.FindByValue("DET")
                            _ddl.Items.Remove(_li)
                            _li = ddlGroup2.Items.FindByValue("SDET")
                            _ddl.Items.Remove(_li)

                        Else
                            Dim _li As ListItem = ddlGroup2.Items.FindByValue(ddlGroup2.SelectedValue.ToString)
                            _ddl.Items.Remove(_li)

                        End If
                    End If
                End If
            Catch ex As Exception
                _objcommonExp.PublishError("For populate the shortbased in the drop downlist(populatestrddl())", ex)
            End Try
        End Sub
        Private Sub EnableUSearchDdl()
            USearch.UCddlCOCEnable = True
            USearch.UCddlLOCEnable = True
            USearch.UCddlUNTEnable = True
            USearch.UCddlDepEnable = True
            USearch.UCddlDESIGEnable = True
            USearch.UCddlGrdEnable = True
            USearch.UCddlLVLEnable = True
            USearch.UCddlSalBasisEnable = True
            USearch.UCddlEmpTypeEnable = True
            ddlshortbasis.SelectedIndex = 0
            ddlGroup2.SelectedIndex = 0
            ddlGroup3.SelectedIndex = 0
            DdlreportType.Enabled = True
            DdlreportType.SelectedIndex = 0
        End Sub
        Protected Function getdatalist() As String
            Dim Htmlstring As New System.Text.StringBuilder, introw As Integer, _gtdt As New DataTable _
            , checkList As String = "", DoCheck, DocheckAll As String, arrparams(1) As SqlParameter _
            , strId As String()
            DoCheck = ""
            DocheckAll = ""
            Try
                'when in checked box value in hidden control not found then 'docheck' and 'docheckall' option 
                'should be checked otherwise unchecked
                If hidShortVal.Value.ToString.ToUpper <> ddlshortbasis.SelectedValue.ToString.ToUpper Then
                    DoCheck = "checked"
                    DocheckAll = "checked"
                    hiddls_id.Value = ""
                Else
                    DoCheck = ""
                    DocheckAll = ""
                End If
                hidShortVal.Value = ddlshortbasis.SelectedValue.ToString
                arrparams(0) = New SqlParameter("@SelType", ddlshortbasis.SelectedValue.ToString)
                '@pk_userid added by sonia on 20 sep 08 to implement filteration according to the 
                'usergroup location mapping 
                arrparams(1) = New SqlParameter("@pk_userid", Session("UID").ToString)
                'execute storeporcedure for get record.
                _gtdt = _ObjData.GetDataTableProc("sp_sel_shortby_onselectbasis", arrparams)
                'for split the value & store in arry
                strId = hiddls_id.Value.ToString.Split(",")
                'for compare the data table record with string length for check all.
                If _gtdt.Rows.Count = strId.Length - 1 Then
                    DocheckAll = "checked"
                End If
                With _gtdt
                    Htmlstring.Append("<table border=0 " & IIf(_gtdt.Rows.Count <= 0, "style=display:none", "") & ">")
                    'for check the existance of data
                    If _gtdt.Rows.Count > 0 Then
                        Htmlstring.Append("<tr><td><input type='checkbox' id='Chk_SelAll' onclick=CheckAllList(this);")
                        Htmlstring.Append(" " & DocheckAll)
                        Htmlstring.Append(" ></td>")
                        Htmlstring.Append("<td width=3></td>")
                        Htmlstring.Append("<td class=TDCaptionBold>" & "Select All")
                    End If
                    'for checked the select short by type
                    For introw = 0 To _gtdt.Rows.Count - 1
                        Htmlstring.Append("<tr><td><input type=checkbox " & DoCheck & " id=dls_" & introw & " value=" & .Rows(introw).Item("ddlvalue").ToString & " onclick=generateIdString('dlist','dls_'); ")
                        'for check the value of hidden control
                        If hiddls_id.Value.ToString.Length > 0 Then
                            Htmlstring.Append(getchecked(.Rows(introw).Item("ddlvalue").ToString, hiddls_id.Value.ToString))
                        End If
                        Htmlstring.Append(" ></td>")
                        Htmlstring.Append("<td width=3></td>")
                        Htmlstring.Append("<td class=tdcaption>" & .Rows(introw).Item("ddldesc").ToString)
                        Htmlstring.Append("</td></tr>")
                        checkList = checkList & .Rows(introw).Item("ddlvalue").ToString & ","
                    Next
                    'for check the value of hidden control
                    If RTrim(LTrim(hiddls_id.Value.ToString)) = "" Then
                        hiddls_id.Value = checkList
                    End If
                    hiddls_count.Value = .Rows.Count.ToString
                    Htmlstring.Append("</table>")
                End With
                Return Htmlstring.ToString
            Catch ex As Exception
                _objcommonExp.PublishError("Error in getdatalist()", ex)
            End Try
            Return ""
        End Function
        'to check or uncheck the checkbox created by string builder
        Protected Function getchecked(ByVal strval As String, ByVal ctrlname As String) As String
            Try
                Dim strId As String(), intcount As Integer
                strId = ctrlname.Split(",")
                'for check or uncheck the checkbox.
                For intcount = 0 To UBound(strId)
                    If strval.Trim = strId(intcount).Trim Then
                        Return " checked"
                    End If
                Next
            Catch ex As Exception
                _objcommonExp.PublishError("Error in getchecked()", ex)
            End Try
            Return ""
        End Function
        'for change index
        Private Sub ddlshortbasis_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlshortbasis.SelectedIndexChanged
            Try
                If ddlshortbasis.SelectedValue.ToString <> "" Then
                    'for check the drop down list selected value is cost center
                    If ddlshortbasis.SelectedValue.ToString = "COC" Then
                        USearch.UCddlCOCEnable = False
                        USearch.UCddlUNTEnable = False
                        USearch.UCddlLOCEnable = False
                        'for check the drop down list selected value is Location
                    ElseIf ddlshortbasis.SelectedValue.ToString = "LOC" Then
                        USearch.UCddlCOCEnable = False
                        USearch.UCddlUNTEnable = False
                        USearch.UCddlLOCEnable = False
                        'for check the drop down list selected value is Unit
                    ElseIf ddlshortbasis.SelectedValue.ToString = "UNT" Then
                        USearch.UCddlCOCEnable = False
                        USearch.UCddlUNTEnable = False
                        USearch.UCddlLOCEnable = False
                    Else
                        USearch.UCddlCOCEnable = True
                        USearch.UCddlUNTEnable = True
                        USearch.UCddlLOCEnable = True
                    End If
                    'for check the drop down list selected value is Department
                    If ddlshortbasis.SelectedValue.ToString = "DET" Then
                        USearch.UCddlDepEnable = False
                    Else
                        USearch.UCddlDepEnable = True
                    End If
                    'for check the drop down list selected value is Designation
                    If ddlshortbasis.SelectedValue.ToString = "DES" Then
                        USearch.UCddlDESIGEnable = False
                    Else
                        USearch.UCddlDESIGEnable = True
                    End If
                    'for check the drop down list selected value is Grade
                    If ddlshortbasis.SelectedValue.ToString = "GRD" Then
                        USearch.UCddlGrdEnable = False
                    Else
                        USearch.UCddlGrdEnable = True
                    End If
                    'for check the drop down list selected value is Level
                    If ddlshortbasis.SelectedValue.ToString = "LVL" Then
                        USearch.UCddlLVLEnable = False
                    Else
                        USearch.UCddlLVLEnable = True
                    End If
                Else
                    hiddls_id.Value = ""
                    hiddls_selcount.Value = ""
                    USearch.UCddlCOCEnable = True
                    USearch.UCddlLOCEnable = True
                    USearch.UCddlUNTEnable = True
                    USearch.UCddlDepEnable = True
                    USearch.UCddlDESIGEnable = True
                    USearch.UCddlGrdEnable = True
                    USearch.UCddlLVLEnable = True
                End If
                populatestrddl(ddlGroup2)
                populatestrddl(ddlGroup3)
                PopulateGroup(ddlGroup2, ddlGroup2.SelectedValue.ToString)
                PopulateGroup(ddlGroup3, ddlGroup3.SelectedValue.ToString)
            Catch ex As Exception
                _objcommonExp.PublishError("Error in ddlshortbasis_SelectedIndexChanged()", ex)
            End Try
        End Sub
        Private Sub PopulateGroup(ByRef _ddl As DropDownList, Optional ByVal _selVal As String = "")
            Dim arrparams(1) As SqlParameter, _gtdt As New DataTable, i As Integer = 0
            arrparams(0) = New SqlParameter("@SelType", _selVal.ToString)
            arrparams(1) = New SqlParameter("@pk_userid", Session("UID").ToString)
            _gtdt = _ObjData.GetDataTableProc("sp_sel_shortby_onselectbasis", arrparams)
            If UCase(_ddl.ID.ToString) = UCase("ddlGroup2") Then
                chkListGroup1.DataSource = _gtdt
                chkListGroup1.DataTextField = "ddldesc"
                chkListGroup1.DataValueField = "ddlvalue"
                chkListGroup1.DataBind()
                chkAllGr1.Checked = True
                For i = 0 To _gtdt.Rows.Count - 1
                    chkListGroup1.Items(i).Selected = True
                Next
            Else
                chkListGroup2.DataSource = _gtdt
                chkListGroup2.DataTextField = "ddldesc"
                chkListGroup2.DataValueField = "ddlvalue"
                chkListGroup2.DataBind()
                hidGroup2Count.Value = _gtdt.Rows.Count
                chkAllGr2.Checked = True
                For i = 0 To _gtdt.Rows.Count - 1
                    chkListGroup2.Items(i).Selected = True
                Next
            End If
        End Sub
        Protected Sub ddlGroup2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlGroup2.SelectedIndexChanged
            populatestrddl(ddlGroup3)
            PopulateGroup(ddlGroup2, ddlGroup2.SelectedValue.ToString)
            PopulateGroup(ddlGroup3, ddlGroup3.SelectedValue.ToString)
        End Sub
        Protected Sub ddlGroup3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlGroup3.SelectedIndexChanged
            PopulateGroup(ddlGroup3, ddlGroup3.SelectedValue.ToString)
        End Sub
        Private Sub getGroupVal()
            Dim i As Integer = 0
            Dim _strVal As String = ""
            For i = 0 To chkListGroup1.Items.Count - 1
                If chkListGroup1.Items(i).Selected = True Then
                    _strVal = _strVal & chkListGroup1.Items(i).Value & ","
                End If
            Next
            If _strVal.Length > 0 Then
                _strVal = Left(_strVal, _strVal.Length - 1)
            End If
            hidGroup2Val.Value = _strVal
            _strVal = ""
            For i = 0 To chkListGroup2.Items.Count - 1
                If chkListGroup2.Items(i).Selected = True Then
                    _strVal = _strVal & chkListGroup2.Items(i).Value & ","
                End If
            Next
            If _strVal.Length > 0 Then
                _strVal = Left(_strVal, _strVal.Length - 1)
            End If
            hidGroup3Val.Value = _strVal
        End Sub
        Protected Sub Btnclick_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Btnclick.Click
            Dim dt As DataTable, Name As String = Nothing, Type As String = Nothing, TargetURL As String
            dt = _ObjData.ExecSQLQuery("Select isnull(Type,'') Type from trnsalaryreport where repid = " + Server.UrlEncode(DDLPaySlipType.SelectedValue))
            Type = _objCommon.Nz(dt.Rows(0)("Type"), "").ToString()
            If Type = "R" Then
                Name = "lstRegister"
            Else
                Name = "lstsalslip"
            End If
            TargetURL = "frmpayslipconfigure.aspx?ID=" + Server.UrlEncode(DDLPaySlipType.SelectedValue) + "~" + Server.UrlEncode(Name.ToString)
            Response.Redirect(TargetURL)
        End Sub
        '''''''''''''''''''
        Private Sub ExportToExcelForReg(ByVal source As DataSet, ByRef _ExcelDoc As StreamWriter)
            Dim _RowCount As Integer = 0
            'Dim _ExcelDoc As System.IO.StringWriter
            ' _ExcelDoc = New System.IO.StringWriter()
            Try
                Dim _ExcelXML As New StringBuilder()
                _ExcelDoc.Write("<?xml version=""1.0""?><?mso-application progid=""Excel.Sheet""?>" & Chr(13) & "" & Chr(10) & "<Workbook ")
                _ExcelDoc.Write("xmlns=""urn:schemas-microsoft-com:office:spreadsheet""" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write(" xmlns:o=""urn:schemas-microsoft-com:office:office""" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write(" xmlns:x=""urn:schemas-microsoft-com:office:excel""" & Chr(13) & "" & Chr(10))
                _ExcelDoc.Write(" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">" & Chr(13) & "" & Chr(10) & " <Styles>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""Default"" ss:Name=""Normal"">" & Chr(13) & "" & Chr(10) & " ")
                '_ExcelXML.Append("<Alignment ss:Horizontal=\"Right\" ss:Vertical=\"Bottom\"/>\r\n <Borders/>"); 
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>" & Chr(13) & "" & Chr(10) & " <Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /></Borders>")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Font/>" & Chr(13) & "" & Chr(10) & " <Interior/>" & Chr(13) & "" & Chr(10) & " <NumberFormat/>")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Protection/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")

                'Cumstomise format 
                _ExcelDoc.Write("<Style ss:ID=""CD"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Center"" ss:Vertical=""Bottom"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelDoc.Write("ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Bold=""1"" />" & Chr(13) & "" & Chr(10) & " <Interior ss:Color=""#CDAF95"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & " ")

                '_ExcelXML.Append("\r\n <Style ss:ID=\"StyleBorder\"><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>"); 
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Style ss:ID=""LC"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelDoc.Write("x:Family=""Swiss"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")

                '_ExcelXML.Append("\r\n <Style ss:ID=\"StyleBorder\"><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>"); 
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Style ss:ID=""BC"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Center"" ss:Vertical=""Bottom"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelDoc.Write("x:Family=""Swiss"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")

                'FOR EXCEL HEARDER ONLY
                _ExcelDoc.Write("<Style ss:ID=""C2"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Color=""#000000"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")



                'FOR COLUMN ONLY
                _ExcelDoc.Write("<Style ss:ID=""C3"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""20"" ss:Color=""#000000"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#FFFFFF"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelDoc.Write("<Style ss:ID=""SL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat")
                _ExcelDoc.Write(" ss:Format=""@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _ExcelDoc.Write("ss:ID=""DC"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""0.0000""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""IT"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""0""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _ExcelDoc.Write("ss:ID=""DL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""mm/dd/yyyy;@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""s2"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Top"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:Size=""9"" ss:Color=""#808080"" ss:Underline=""Single""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#FFFFFF"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("</Styles>" & Chr(13) & "" & Chr(10) & " ")

                Dim start_ExcelXML As String = _ExcelDoc.ToString()
                Const end_ExcelXML As String = "</Workbook>"
                Dim sheetCount As Integer = 1

                '_ExcelDoc.Write(start_ExcelXML)
                _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                _ExcelDoc.Write("<Table><Column ss:Width=""40""/><Column ss:Width=""60""/><Column ss:Width=""120""/><Column ss:Width=""80""/><Column ss:Width=""120""/><Column ss:Width=""40""/><Column ss:Width=""80""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""50""/><Column ss:Width=""80""/><Column ss:Width=""80""/><Column ss:Width=""100""/><Column ss:Width=""65""/><Column ss:Width=""65""/><Column ss:Width=""65""/>")

                Dim Px As Integer = (source.Tables(1).Columns.Count - 1) - 3
                Dim Wx, Zx, Rt As String
                Wx = source.Tables(3).Rows(0).Item("AddCount").ToString
                Zx = source.Tables(4).Rows(0).Item("DedCount").ToString
                Rt = source.Tables(6).Rows(0).Item("Rate").ToString
                _ExcelDoc.Write("<Row>")

                For x As Integer = 0 To 1
                    If x = 1 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""" & ((Rt - 1) + Wx + 11 + (Zx - 1)) & """ ss:StyleID=""C3""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("REGISTER OF WAGES")
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell  ss:StyleID=""C3""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")
                _ExcelDoc.Write("<Row>")



                For x As Integer = 0 To 1
                    If x = 1 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""" & ((Rt - 1) + Wx + 11 + (Zx - 1)) & """ ss:StyleID=""LC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("FORM-XVII")
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell  ss:StyleID=""LC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")
                _ExcelDoc.Write("<Row>")

                For x As Integer = 0 To 1
                    If x = 1 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""" & ((Rt - 1) + Wx + 11 + (Zx - 1)) & """ ss:StyleID=""LC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("{See Rule 78(1) (a)(i)}")
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell  ss:StyleID=""LC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To 10
                    If x = 1 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""" & ((Rt - 1) + Wx + 11 + (Zx - 1)) & """  ss:StyleID=""LC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Name and address of Contractor  ")
                        _ExcelDoc.Write(":-")
                        _ExcelDoc.Write(source.Tables(0).Rows(0).Item("Comp_Name").ToString)
                        _ExcelDoc.Write("&nbsp;")
                        _ExcelDoc.Write(",")
                        _ExcelDoc.Write(source.Tables(0).Rows(0).Item("Address").ToString)
                        _ExcelDoc.Write("&nbsp;")
                        'added by rajesh for show zip code in excel 
                        _ExcelDoc.Write(source.Tables(0).Rows(0).Item("ZIP_Code").ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell  ss:StyleID=""LC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To 10
                    If x = 1 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""" & ((Rt - 1) + Wx + 11 + (Zx - 1)) & """  ss:StyleID=""LC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Name and address of Establishment in/ Under which contract is carried on  ")
                        _ExcelDoc.Write(":-")
                        _ExcelDoc.Write(source.Tables(5).Rows(0).Item("Grp_Name").ToString)
                        _ExcelDoc.Write("&nbsp;")
                        _ExcelDoc.Write(",")
                        _ExcelDoc.Write(source.Tables(5).Rows(0).Item("Grp_Address").ToString)
                        _ExcelDoc.Write("&nbsp;")
                        'added by rajesh for show zip code in excel 
                        _ExcelDoc.Write(source.Tables(0).Rows(0).Item("ZIP_Code").ToString)
                        _ExcelDoc.Write("</Data></Cell>")

                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""LC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")

                For x As Integer = 0 To 1
                    If x = 1 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""" & ((Rt - 1) + Wx + 11 + (Zx - 1)) & """ ss:StyleID=""LC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Nature and location of work………………………")
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell  ss:StyleID=""LC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")
                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To 1
                    If x = 1 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""" & ((Rt - 1) + Wx + 11 + (Zx - 1)) & """ ss:StyleID=""LC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Wages period Monthly............")
                        _ExcelDoc.Write(MonthName(source.Tables(0).Rows(0).Item("SelMonth").ToString))
                        _ExcelDoc.Write("   ")
                        _ExcelDoc.Write("Year ")
                        _ExcelDoc.Write(source.Tables(0).Rows(0).Item("SelYear").ToString)
                        _ExcelDoc.Write("</Data></Cell>")

                    Else
                        _ExcelDoc.Write("<Cell  ss:StyleID=""LC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To 1 - 1
                    If x < 1 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""" & ((Rt - 1) + Wx + 12 + (Zx - 1)) & """ ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")

                '''''''''''''''''''''''''''

                '''''''''''''''''''''''''''
                _ExcelDoc.Write("<Row>")
                'For x As Integer = 0 To 6

                'If source.Tables(1).Columns(x).ColumnName = "SrNo" Then
                '    _ExcelDoc.Write("<Cell ss:MergeAcross=""6""   ss:StyleID=""C2"" ><Data ss:Type=""String"">")
                '    _ExcelDoc.Write("")
                'ElseIf source.Tables(1).Columns(x).ColumnName = "EmpCode" Then
                '    _ExcelDoc.Write("<Cell  ss:StyleID=""C2"" ><Data ss:Type=""String"">")
                '    _ExcelDoc.Write("")
                'ElseIf source.Tables(1).Columns(x).ColumnName = "NAME" Then
                '    _ExcelDoc.Write("<Cell  ss:StyleID=""C2"" ><Data ss:Type=""String"">")
                '    _ExcelDoc.Write("")
                'ElseIf source.Tables(1).Columns(x).ColumnName = "Serial No. in the Register of workman" Then
                '    _ExcelDoc.Write("<Cell  ss:StyleID=""C2"" ><Data ss:Type=""String"">")
                '    _ExcelDoc.Write("")
                'ElseIf source.Tables(1).Columns(x).ColumnName = "Designation" Then
                '    _ExcelDoc.Write("<Cell  ss:StyleID=""C2"" ><Data ss:Type=""String"">")
                '    _ExcelDoc.Write("")
                'ElseIf source.Tables(1).Columns(x).ColumnName = "NoOfDays" Then
                '    _ExcelDoc.Write("<Cell  ss:StyleID=""C2"" ><Data ss:Type=""String"">")
                '    _ExcelDoc.Write("")
                'ElseIf source.Tables(1).Columns(x).ColumnName = "Units of work done" Then
                '    _ExcelDoc.Write("<Cell  ss:StyleID=""C2""><Data ss:Type=""String"">")
                '    _ExcelDoc.Write("")

                'Else

                '    _ExcelDoc.Write("<Cell  ss:StyleID=""C2""><Data ss:Type=""String"">")
                '    _ExcelDoc.Write(source.Tables(1).Columns(x).ColumnName)

                'End If
                _ExcelDoc.Write("<Cell ss:MergeAcross=""6""   ss:StyleID=""C2"" ><Data ss:Type=""String"">")
                _ExcelDoc.Write("")
                _ExcelDoc.Write("</Data></Cell>")
                'Next
                'Dim Px As Integer = (source.Tables(1).Columns.Count - 1) - 3
                'Dim Wx, Zx, Rt As String
                'Wx = source.Tables(3).Rows(0).Item("AddCount").ToString
                'Zx = source.Tables(4).Rows(0).Item("DedCount").ToString
                'Rt = source.Tables(6).Rows(0).Item("Rate").ToString
                ' For p As Integer = (source.Tables(3).Rows(0).Item("AddCount").ToString) To source.Tables(1).Columns.Count - 1


                _ExcelDoc.Write("<Cell ss:MergeAcross=""" & (Rt - 1) & """ ss:StyleID=""C2""><Data ss:Type=""String"">")
                _ExcelDoc.Write("Daily-rate of wages/Piece rate(Wages P/M)")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:MergeAcross=""" & (Wx) & """ ss:StyleID=""C2""><Data ss:Type=""String"">")
                _ExcelDoc.Write("Amount of wages earned")
                _ExcelDoc.Write("</Data></Cell>")

                '_ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                '_ExcelDoc.Write("Deductions,if any, (indicate nature)")
                '_ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:MergeAcross=""" & (Zx - 1) & """ ss:StyleID=""C2""><Data ss:Type=""String"">")
                _ExcelDoc.Write("Deductions,if any, (indicate nature)")
                _ExcelDoc.Write("</Data></Cell>")


                _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""C2""><Data ss:Type=""String"">")
                _ExcelDoc.Write("")
                _ExcelDoc.Write("</Data></Cell>")

                '_ExcelDoc.Write("<Cell ss:StyleID=""C2""><Data ss:Type=""String"">")
                '_ExcelDoc.Write("")
                '_ExcelDoc.Write("</Data></Cell>")


                '_ExcelDoc.Write("<Cell ss:StyleID=""C2""><Data ss:Type=""String"">")
                '_ExcelDoc.Write("")
                '_ExcelDoc.Write("</Data></Cell>")

                '_ExcelDoc.Write("<Cell ss:StyleID=""C2""><Data ss:Type=""String"">")
                '_ExcelDoc.Write("")
                '_ExcelDoc.Write("</Data></Cell>")

                ' Next
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                For T As Integer = 0 To source.Tables(1).Columns.Count - 1
                    If T = 0 Then
                        _ExcelDoc.Write("<Cell   ss:StyleID=""C2""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Sr.No")
                        _ExcelDoc.Write("</Data></Cell>")
                    ElseIf T = 1 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""C2""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Emp Code")
                        _ExcelDoc.Write("</Data></Cell>")

                    ElseIf T = 2 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""C2""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Name  of Workman")
                        _ExcelDoc.Write("</Data></Cell>")
                    ElseIf T = 3 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""C2""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Serial No. in the Register of workman")
                        _ExcelDoc.Write("</Data></Cell>")
                    ElseIf T = 4 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""C2""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Designation")
                        _ExcelDoc.Write("</Data></Cell>")
                    ElseIf T = 5 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""C2""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("No. of Days worked")
                        _ExcelDoc.Write("</Data></Cell>")
                    ElseIf T = 6 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""C2""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Units of work done")
                        _ExcelDoc.Write("</Data></Cell>")

                    Else

                        _ExcelDoc.Write("<Cell ss:StyleID=""C2""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(source.Tables(1).Columns(T).ColumnName)
                        _ExcelDoc.Write("</Data></Cell>")
                    End If

                Next
                _ExcelDoc.Write("</Row>")

                '''''''''''''''''''''''''''''''''''
                For Each x As DataRow In source.Tables(1).Rows
                    _RowCount += 1
                    'if the number of rows is > 63000 create a new page to continue output 
                    'If _RowCount = 63000 Then
                    '    Dim test As String = ""
                    '    test = "god"
                    'End If
                    'If _RowCount = 63000 Then
                    '    _RowCount = 0
                    '    sheetCount += 1
                    '    _ExcelDoc.Write("</Table>")
                    '    _ExcelDoc.Write(" </Worksheet>")
                    '    _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                    '    _ExcelDoc.Write("<Table>")
                    '    _ExcelDoc.Write("<Row>")
                    '    For xi As Integer = 0 To source.Tables(1).Columns.Count - 1
                    '        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write(source.Tables(1).Columns(xi).ColumnName)
                    '        _ExcelDoc.Write("</Data></Cell>")
                    '    Next
                    '    _ExcelDoc.Write("</Row>")
                    'End If
                    _ExcelDoc.Write("<Row>")
                    For y As Integer = 0 To source.Tables(1).Columns.Count - 1
                        Dim XMLstring As String = x(y).ToString()
                        XMLstring = XMLstring.Trim()

                        If IsNumeric(XMLstring) Then
                            If y = source.Tables(1).Columns.Count - 2 Then
                                _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""String"">")
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""Number"">")
                            End If


                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""String"">")
                        End If
                        _ExcelDoc.Write(XMLstring)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    _ExcelDoc.Write("</Row>")
                Next
                '---By GRand
                For Each x As DataRow In source.Tables(2).Rows
                    _RowCount += 1
                    'if the number of rows is > 63000 create a new page to continue output 
                    'If _RowCount = 63000 Then
                    '    Dim test As String = ""
                    '    test = "god"
                    'End If
                    'If _RowCount = 63000 Then
                    '    _RowCount = 0
                    '    sheetCount += 1
                    '    _ExcelDoc.Write("</Table>")
                    '    _ExcelDoc.Write(" </Worksheet>")
                    '    _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                    '    _ExcelDoc.Write("<Table>")
                    '    _ExcelDoc.Write("<Row>")
                    '    For xi As Integer = 0 To source.Tables(2).Columns.Count - 1
                    '        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write(source.Tables(2).Columns(xi).ColumnName)
                    '        _ExcelDoc.Write("</Data></Cell>")
                    '    Next
                    '    _ExcelDoc.Write("</Row>")
                    'End If
                    _ExcelDoc.Write("<Row>")
                    For y As Integer = 0 To source.Tables(2).Columns.Count - 1
                        Dim XMLstring As String = x(y).ToString()
                        XMLstring = XMLstring.Trim()
                        If IsNumeric(XMLstring) Then
                            _ExcelDoc.Write("<Cell ss:StyleID=""BC"">" + "<Data ss:Type=""Number"">")
                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""BC"">" + "<Data ss:Type=""String"">")
                        End If
                        _ExcelDoc.Write(XMLstring)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    _ExcelDoc.Write("</Row>")
                Next
                _ExcelDoc.Write("</Table>")
                _ExcelDoc.Write(" </Worksheet>")
                _ExcelDoc.Write(end_ExcelXML)
            Catch ex As Exception
                'for catch the error
                _objcommonExp.PublishError("For gebnerate the excel file (fnExportTo_ExcelXML())", ex)
            End Try
            ' Return _ExcelDoc
        End Sub
        Private Sub ExportToExcelworkman(ByVal source As DataSet, ByRef _ExcelDoc As StreamWriter)
            Dim _RowCount As Integer = 0
            'Dim _ExcelDoc As System.IO.StringWriter
            ' _ExcelDoc = New System.IO.StringWriter()
            Try
                Dim _ExcelXML As New StringBuilder()
                _ExcelDoc.Write("<?xml version=""1.0""?><?mso-application progid=""Excel.Sheet""?>" & Chr(13) & "" & Chr(10) & "<Workbook ")
                _ExcelDoc.Write("xmlns=""urn:schemas-microsoft-com:office:spreadsheet""" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write(" xmlns:o=""urn:schemas-microsoft-com:office:office""" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write(" xmlns:x=""urn:schemas-microsoft-com:office:excel""" & Chr(13) & "" & Chr(10))
                _ExcelDoc.Write(" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">" & Chr(13) & "" & Chr(10) & " <Styles>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""Default"" ss:Name=""Normal"">" & Chr(13) & "" & Chr(10) & " ")
                '_ExcelXML.Append("<Alignment ss:Horizontal=\"Right\" ss:Vertical=\"Bottom\"/>\r\n <Borders/>"); 
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>" & Chr(13) & "" & Chr(10) & " <Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /></Borders>")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Font/>" & Chr(13) & "" & Chr(10) & " <Interior/>" & Chr(13) & "" & Chr(10) & " <NumberFormat/>")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Protection/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")

                'Cumstomise format 
                _ExcelDoc.Write("<Style ss:ID=""CD"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Center"" ss:Vertical=""Bottom"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelDoc.Write("ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Bold=""1"" />" & Chr(13) & "" & Chr(10) & " <Interior ss:Color=""#CDAF95"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & " ")


                '_ExcelXML.Append("\r\n <Style ss:ID=\"StyleBorder\"><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>"); 
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Style ss:ID=""BC"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelDoc.Write("x:Family=""Swiss"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""SL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat")
                _ExcelDoc.Write(" ss:Format=""@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _ExcelDoc.Write("ss:ID=""DC"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""0.0000""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""IT"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""0""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _ExcelDoc.Write("ss:ID=""DL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""mm/dd/yyyy;@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""s2"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Top"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:Size=""9"" ss:Color=""#808080"" ss:Underline=""Single""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#FFFFFF"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("</Styles>" & Chr(13) & "" & Chr(10) & " ")

                Dim start_ExcelXML As String = _ExcelDoc.ToString()
                Const end_ExcelXML As String = "</Workbook>"
                Dim sheetCount As Integer = 1

                '_ExcelDoc.Write(start_ExcelXML)
                _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                _ExcelDoc.Write("<Table><Column ss:Width=""40""/><Column ss:Width=""100""/><Column ss:Width=""120""/><Column ss:Width=""80""/><Column ss:Width=""120""/><Column ss:Width=""100""/><Column ss:Width=""80""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""50""/><Column ss:Width=""80""/><Column ss:Width=""80""/><Column ss:Width=""100""/><Column ss:Width=""65""/><Column ss:Width=""65""/><Column ss:Width=""65""/>")
                _ExcelDoc.Write("<Row>")

                _ExcelDoc.Write("<Cell ss:MergeAcross=""13"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                _ExcelDoc.Write("FORM XIII")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")

                _ExcelDoc.Write("<Cell ss:MergeAcross=""13"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                _ExcelDoc.Write("[See Rule 75]")
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")

                _ExcelDoc.Write("<Cell ss:MergeAcross=""13"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                _ExcelDoc.Write("Register of Workmen Employed by Contractor")
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell ss:MergeAcross=""13""  ss:StyleID=""BC""><Data ss:Type=""String"">")
                _ExcelDoc.Write("Name and address of Contractor  ")
                _ExcelDoc.Write(":-")
                _ExcelDoc.Write(source.Tables(0).Rows(0).Item("Comp_Name").ToString)
                _ExcelDoc.Write("&nbsp;")
                _ExcelDoc.Write(",")
                _ExcelDoc.Write(source.Tables(0).Rows(0).Item("Address").ToString)
                _ExcelDoc.Write("&nbsp;")
                _ExcelDoc.Write(source.Tables(0).Rows(0).Item("ZIP_Code").ToString)
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell ss:MergeAcross=""13""  ss:StyleID=""BC""><Data ss:Type=""String"">")

                _ExcelDoc.Write("Name and address of Establishment in/ Under which contract is carried on  ")
                _ExcelDoc.Write(":-")
                _ExcelDoc.Write(source.Tables(1).Rows(0).Item("Grp_Name").ToString)
                _ExcelDoc.Write("&nbsp;")
                _ExcelDoc.Write(",")
                _ExcelDoc.Write(source.Tables(1).Rows(0).Item("Grp_Address").ToString)
                _ExcelDoc.Write("&nbsp;")
                _ExcelDoc.Write(source.Tables(1).Rows(0).Item("ZIP_Code").ToString)
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To 1 - 1
                    If x < 1 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""13"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")


                '''''''''''''''''''''''''''
                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To source.Tables(2).Columns.Count - 2
                    _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(source.Tables(2).Columns(x).ColumnName)
                    _ExcelDoc.Write("</Data></Cell>")
                Next
                _ExcelDoc.Write("</Row>")

                '''''''''''''''''''''''''''
                _ExcelDoc.Write("<Row>")



                For Z As Integer = 0 To 12
                    _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("")
                    _ExcelDoc.Write("</Data></Cell>")
                Next
                _ExcelDoc.Write("</Row>")

                '''''''''''''''''''''''''''''''''''
                For Each x As DataRow In source.Tables(2).Rows
                    _RowCount += 1
                    'if the number of rows is > 63000 create a new page to continue output 
                    'If _RowCount = 63000 Then
                    '    Dim test As String = ""
                    '    test = "god"
                    'End If
                    'If _RowCount = 63000 Then
                    '    _RowCount = 0
                    '    sheetCount += 1
                    '    _ExcelDoc.Write("</Table>")
                    '    _ExcelDoc.Write(" </Worksheet>")
                    '    _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                    '    _ExcelDoc.Write("<Table>")
                    '    _ExcelDoc.Write("<Row>")
                    '    For xi As Integer = 0 To source.Tables(2).Columns.Count - 2
                    '        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write(source.Tables(2).Columns(xi).ColumnName)
                    '        _ExcelDoc.Write("</Data></Cell>")
                    '    Next
                    '    _ExcelDoc.Write("</Row>")
                    'End If
                    _ExcelDoc.Write("<Row>")
                    For y As Integer = 0 To source.Tables(2).Columns.Count - 2
                        Dim XMLstring As String = x(y).ToString()
                        XMLstring = XMLstring.Trim()

                        If IsNumeric(XMLstring) Then
                            _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""Number"">")

                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""String"">")
                        End If
                        _ExcelDoc.Write(XMLstring)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    _ExcelDoc.Write("</Row>")
                Next

                'added by rajesh for show current date in excel in excel 
                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell ss:MergeAcross=""13""  ss:StyleID=""BC""><Data ss:Type=""String"">")
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")

                _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                _ExcelDoc.Write("")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                _ExcelDoc.Write("")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:MergeAcross=""11""  ss:StyleID=""SL""><Data ss:Type=""String"">")
                _ExcelDoc.Write("Ins by GSR 948, dated 12-7-1978 (22-7-1978).")
                '_ExcelDoc.Write(source.Tables(0).Rows(0).Item("dated").ToString)
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("</Table>")
                _ExcelDoc.Write(" </Worksheet>")
                _ExcelDoc.Write(end_ExcelXML)
            Catch ex As Exception
                'for catch the error
                _objcommonExp.PublishError("For gebnerate the excel file (fnExportTo_ExcelXML())", ex)
            End Try
            ' Return _ExcelDoc
        End Sub
        'Add new Excel Export funtion for salary register group wise, by praveen verma on 23 Aug 2013..
        Private Sub ExportToExcelXMLGrpWise(ByVal source As DataSet, ByRef _ExcelDoc As StreamWriter)
            Dim _RowCount As Integer = 0, headstr As String = "", K As Integer = 0, Stlname As String = ""
            'Dim _ExcelDoc As System.IO.StringWriter
            '_ExcelDoc = New System.IO.StringWriter()
            Try
                Dim _ExcelXML As New StringBuilder()
                _ExcelDoc.Write("<?xml version=""1.0""?><?mso-application progid=""Excel.Sheet""?>" & Chr(13) & "" & Chr(10) & "<Workbook ")
                _ExcelDoc.Write("xmlns=""urn:schemas-microsoft-com:office:spreadsheet""" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write(" xmlns:o=""urn:schemas-microsoft-com:office:office""" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write(" xmlns:x=""urn:schemas-microsoft-com:office:excel""" & Chr(13) & "" & Chr(10))
                _ExcelDoc.Write(" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">" & Chr(13) & "" & Chr(10) & " <Styles>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""Default"" ss:Name=""Normal"">" & Chr(13) & "" & Chr(10) & " ")
                '_ExcelXML.Append("<Alignment ss:Horizontal=\"Right\" ss:Vertical=\"Bottom\"/>\r\n <Borders/>"); 
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>" & Chr(13) & "" & Chr(10) & " <Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders>")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Font/>" & Chr(13) & "" & Chr(10) & " <Interior/>" & Chr(13) & "" & Chr(10) & " <NumberFormat/>")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Protection/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")

                '_ExcelXML.Append("\r\n <Style ss:ID=\"StyleBorder\"><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>"); 
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Style ss:ID=""BC"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelDoc.Write("x:Family=""Swiss"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""SL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat")
                _ExcelDoc.Write(" ss:Format=""@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _ExcelDoc.Write("ss:ID=""DC"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""0.0000""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""IT"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""0""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _ExcelDoc.Write("ss:ID=""DL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""mm/dd/yyyy;@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""s2"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Top"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:Size=""9"" ss:Color=""#808080"" ss:Underline=""Single""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#FFFFFF"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelDoc.Write("<Style ss:ID=""s32"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Vertical=""Center"" ss:Horizontal=""Center""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" ss:Family=""Swiss"" ss:Size=""9"" ss:Color=""#000000"" ss:Bold=""0""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#CCFFFF"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelDoc.Write("<Style ss:ID=""s33"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Vertical=""Center"" ss:Horizontal=""Center""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" ss:Family=""Swiss"" ss:Size=""9"" ss:Color=""#000000"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#CCFFFF"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelDoc.Write("<Style ss:ID=""s1"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Vertical=""Center"" ss:Horizontal=""Center""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" ss:Family=""Swiss"" ss:Size=""11"" ss:Color=""#FFE4E1"" ss:Bold=""0""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#800080"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                'FOR EMPLOYEE DETAILS COLUMN ONLY
                _ExcelDoc.Write("<Style ss:ID=""ED"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000"" ss:Bold=""0""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Pattern=""Solid""/></Style>" & Chr(13) & "" & Chr(10) & "")

                'FOR EXCEL HEARDER ONLY
                _ExcelDoc.Write("<Style ss:ID=""C2"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Color=""#000000"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                'FOR COLUMN ONLY
                _ExcelDoc.Write("<Style ss:ID=""C3"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                'FOR Header ONLY
                _ExcelDoc.Write("<Style ss:ID=""C4"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""10"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#C2FDE2"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelDoc.Write("<Style ss:ID=""C5"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""10"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#F9966B"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelDoc.Write("<Style ss:ID=""C6"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""10"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#FFFF00"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelDoc.Write("<Style ss:ID=""C7"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""10"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#FF9900"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelDoc.Write("<Style ss:ID=""C8"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""10"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#66CDAA"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelDoc.Write("<Style ss:ID=""C9"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""10"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#6495ED"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelDoc.Write("<Style ss:ID=""A1"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""10"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#FFDAB9"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelDoc.Write("</Styles>" & Chr(13) & "" & Chr(10) & " ")

                Dim start_ExcelXML As String = _ExcelDoc.ToString()
                Const end_ExcelXML As String = "</Workbook>"
                Dim sheetCount As Integer = 1, tcnt As Integer = 0

                '_ExcelDoc.Write(start_ExcelXML)
                _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")

                _ExcelDoc.Write("<Table><Column ss:Width=""40""/><Column ss:Width=""180""/><Column ss:Width=""180""/><Column ss:Width=""150""/><Column ss:Width=""225""/><Column ss:Width=""180""/><Column ss:Width=""180""/><Column ss:Width=""180""/><Column ss:Width=""180""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/>")

                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell ss:StyleID=""s1"" ss:MergeAcross=""" & 10 & """><Data ss:Type=""String"">")
                _ExcelDoc.Write("Cost sheet for the Month of " & ddlMonthYear.SelectedItem.ToString)
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""s1"" ss:MergeAcross=""" & source.Tables(0).Columns.Count - 11 & """><Data ss:Type=""String"">")
                _ExcelDoc.Write("&nbsp;")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")

                _ExcelDoc.Write("<Cell ss:StyleID=""A1"" ss:MergeAcross=""" & 2 & """><Data ss:Type=""String"">")
                _ExcelDoc.Write("Particulare")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""C4"" ss:MergeAcross=""" & 6 & """><Data ss:Type=""String"">")
                _ExcelDoc.Write("Employee Details")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""C8"" ss:MergeAcross=""" & CType(source.Tables(1).Rows(0).Item(0).ToString, Integer) - 1 & """><Data ss:Type=""String"">")
                _ExcelDoc.Write("Salary(Rate)")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""C5"" ss:MergeAcross=""" & CType(source.Tables(1).Rows(0).Item(1).ToString, Integer) - 1 & """><Data ss:Type=""String"">")
                _ExcelDoc.Write("Salary(Earned)")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""C9"" ss:MergeAcross=""" & CType(source.Tables(1).Rows(0).Item(2).ToString, Integer) - 1 & """><Data ss:Type=""String"">")
                _ExcelDoc.Write("Employee Deduction")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""C6"" ss:MergeAcross=""" & 3 & """><Data ss:Type=""String"">")
                _ExcelDoc.Write("Employer Share")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""C7"" ss:MergeAcross=""" & 1 & """><Data ss:Type=""String"">")
                _ExcelDoc.Write("Payable")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("</Row>")


                If source.Tables.Count > 0 Then
                    If source.Tables(0).Rows.Count > 0 Then
                        _ExcelDoc.Write("<Row>")

                        _ExcelDoc.Write("<Cell ss:StyleID=""C2"" ><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Sno.")
                        _ExcelDoc.Write("</Data></Cell>")

                        For x As Integer = 0 To source.Tables(0).Columns.Count - 1
                            _ExcelDoc.Write("<Cell ss:StyleID=""C2"" ><Data ss:Type=""String"">")
                            _ExcelDoc.Write(IIf(source.Tables(0).Columns(x).ToString.ToUpper = "ELWF", "LWF", source.Tables(0).Columns(x).ToString.ToUpper))
                            _ExcelDoc.Write("</Data></Cell>")
                        Next
                        _ExcelDoc.Write("</Row>")
                    End If
                End If

                For Each x As DataRow In source.Tables(0).Rows
                    _RowCount += 1
                    'If _RowCount = 63000 Then
                    '    _RowCount = 0
                    '    sheetCount += 1
                    '    _ExcelDoc.Write("</Table>")
                    '    _ExcelDoc.Write(" </Worksheet>")
                    '    _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                    '    _ExcelDoc.Write("<Table>")
                    '    _ExcelDoc.Write("<Row>")
                    '    For xi As Integer = 0 To source.Tables(0).Columns.Count - 1
                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C2""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write(source.Tables(0).Columns(xi).ColumnName)
                    '        _ExcelDoc.Write("</Data></Cell>")
                    '    Next
                    '    _ExcelDoc.Write("</Row>")
                    'End If
                    _ExcelDoc.Write("<Row>")
                    If _RowCount < source.Tables(0).Rows.Count Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""s32""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(_RowCount)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""s32""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("&nbsp;")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If

                    For y As Integer = 0 To source.Tables(0).Columns.Count - 1
                        Dim XMLstring As String = x(y).ToString()
                        XMLstring = XMLstring.Trim()

                        If _RowCount = source.Tables(0).Rows.Count Then
                            If IsNumeric(XMLstring) Then
                                _ExcelDoc.Write("<Cell ss:StyleID=""s33""><Data ss:Type=""Number"">")
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""s33""><Data ss:Type=""String"">")
                            End If

                        Else
                            If IsNumeric(XMLstring) Then
                                _ExcelDoc.Write("<Cell ss:StyleID=""s32""><Data ss:Type=""Number"">")
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""s32""><Data ss:Type=""String"">")
                            End If
                        End If

                        If _RowCount = source.Tables(0).Rows.Count And y = 0 Then
                            _ExcelDoc.Write("Total")
                        Else
                            _ExcelDoc.Write(XMLstring)
                        End If

                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    _ExcelDoc.Write("</Row>")

                Next
                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell><Data ss:Type=""String"">")
                _ExcelDoc.Write("&nbsp;")
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell><Data ss:Type=""String"">")
                _ExcelDoc.Write("&nbsp;")
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell><Data ss:Type=""String"">")
                _ExcelDoc.Write("&nbsp;")
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                If source.Tables(2).Rows.Count > 0 Then

                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell><Data ss:Type=""String"">")
                    _ExcelDoc.Write("&nbsp;")
                    _ExcelDoc.Write("</Data></Cell>")

                    For x As Integer = 0 To source.Tables(2).Columns.Count - 1
                        _ExcelDoc.Write("<Cell ss:StyleID=""C2"" ><Data ss:Type=""String"">")
                        _ExcelDoc.Write(source.Tables(2).Columns(x).ToString.ToUpper)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    _ExcelDoc.Write("</Row>")
                End If

                _RowCount = 0
                For Each x As DataRow In source.Tables(2).Rows
                    _RowCount += 1
                    'If _RowCount = 63000 Then
                    '    _RowCount = 0
                    '    sheetCount += 1
                    '    _ExcelDoc.Write("</Table>")
                    '    _ExcelDoc.Write(" </Worksheet>")
                    '    _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                    '    _ExcelDoc.Write("<Table>")
                    '    _ExcelDoc.Write("<Row>")
                    '    For xi As Integer = 0 To source.Tables(2).Columns.Count - 1
                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C2""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write(source.Tables(0).Columns(xi).ColumnName)
                    '        _ExcelDoc.Write("</Data></Cell>")
                    '    Next
                    '    _ExcelDoc.Write("</Row>")
                    'End If
                    _ExcelDoc.Write("<Row>")

                    _ExcelDoc.Write("<Cell><Data ss:Type=""String"">")
                    _ExcelDoc.Write("&nbsp;")
                    _ExcelDoc.Write("</Data></Cell>")
                    For y As Integer = 0 To source.Tables(2).Columns.Count - 1
                        Dim XMLstring As String = x(y).ToString()
                        XMLstring = XMLstring.Trim()
                        If _RowCount = source.Tables(2).Rows.Count Then
                            If IsNumeric(XMLstring) Then
                                _ExcelDoc.Write("<Cell ss:StyleID=""s33""><Data ss:Type=""Number"">")
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""s33""><Data ss:Type=""String"">")
                            End If

                        Else
                            If IsNumeric(XMLstring) Then
                                _ExcelDoc.Write("<Cell ss:StyleID=""s32""><Data ss:Type=""Number"">")
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""s32""><Data ss:Type=""String"">")
                            End If
                        End If
                        _ExcelDoc.Write(XMLstring)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    _ExcelDoc.Write("</Row>")

                Next

                _ExcelDoc.Write("</Table>")
                '_ExcelDoc.Write("<WorksheetOptions xmlns=""urn:schemas-microsoft-com:office:excel""><Selected/><FreezePanes/><FrozenNoSplit/><SplitHorizontal>3</SplitHorizontal><TopRowBottomPane>3</TopRowBottomPane><SplitVertical>3</SplitVertical><LeftColumnRightPane>3</LeftColumnRightPane><ActivePane>0</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>1</Number></Pane><Pane><Number>2</Number></Pane><Pane><Number>0</Number><ActiveCol>0</ActiveCol></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions>")
                _ExcelDoc.Write("</Worksheet>")
                _ExcelDoc.Write(end_ExcelXML)
            Catch ex As Exception
                'for catch the error
                _objcommonExp.PublishError("For gebnerate the excel file (fnExportTo_ExcelXML())", ex)
            End Try
            'Return _ExcelDoc
        End Sub
        'Add optional argument flag , it is used to display column width accroding to report type. Added by praveen verma on 19 Feb 2013.
        Private Sub ExportToExcelXML(ByVal source As DataSet, ByRef _ExcelDoc As StreamWriter, Optional ByVal flag As String = "")
            Dim _RowCount As Integer = 0, headstr As String = "", K As Integer = 0, Stlname As String = "", TblNumber As Integer = 1
            'Dim _ExcelDoc As System.IO.StringWriter
            '_ExcelDoc = New System.IO.StringWriter()
            Try
                Dim _ExcelXML As New StringBuilder()
                _ExcelDoc.Write("<?xml version=""1.0""?><?mso-application progid=""Excel.Sheet""?>" & Chr(13) & "" & Chr(10) & "<Workbook ")
                _ExcelDoc.Write("xmlns=""urn:schemas-microsoft-com:office:spreadsheet""" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write(" xmlns:o=""urn:schemas-microsoft-com:office:office""" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write(" xmlns:x=""urn:schemas-microsoft-com:office:excel""" & Chr(13) & "" & Chr(10))
                _ExcelDoc.Write(" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">" & Chr(13) & "" & Chr(10) & " <Styles>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""Default"" ss:Name=""Normal"">" & Chr(13) & "" & Chr(10) & " ")
                '_ExcelXML.Append("<Alignment ss:Horizontal=\"Right\" ss:Vertical=\"Bottom\"/>\r\n <Borders/>"); 
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>" & Chr(13) & "" & Chr(10) & " <Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders>")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Font/>" & Chr(13) & "" & Chr(10) & " <Interior/>" & Chr(13) & "" & Chr(10) & " <NumberFormat/>")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Protection/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")

                '_ExcelXML.Append("\r\n <Style ss:ID=\"StyleBorder\"><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>"); 
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Style ss:ID=""BC"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelDoc.Write("x:Family=""Swiss"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""SL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat")
                _ExcelDoc.Write(" ss:Format=""@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _ExcelDoc.Write("ss:ID=""DC"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""0.0000""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""IT"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""0""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _ExcelDoc.Write("ss:ID=""DL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""mm/dd/yyyy;@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""s2"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Top"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")

                If DDLPaySlipType.SelectedValue.ToString = "38" And rbtshowclr.SelectedValue.ToString.ToUpper = "N" Then
                    _ExcelDoc.Write("<Font x:CharSet=""1"" ss:Size=""9"" ss:Underline=""Single""/>" & Chr(13) & "" & Chr(10) & "")
                    _ExcelDoc.Write("<Interior ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                Else
                    _ExcelDoc.Write("<Font x:CharSet=""1"" ss:Size=""9"" ss:Color=""#808080"" ss:Underline=""Single""/>" & Chr(13) & "" & Chr(10) & "")
                    _ExcelDoc.Write("<Interior ss:Color=""#FFFFFF"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                End If

                _ExcelDoc.Write("<Style ss:ID=""s32"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Vertical=""Center""/>" & Chr(13) & "" & Chr(10) & "")
                If DDLPaySlipType.SelectedValue.ToString = "38" And rbtshowclr.SelectedValue.ToString.ToUpper = "N" Then
                    _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" ss:Family=""Swiss"" ss:Size=""9"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                    _ExcelDoc.Write("<Interior ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                Else
                    _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" ss:Family=""Swiss"" ss:Size=""9"" ss:Color=""#000000"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                    _ExcelDoc.Write("<Interior ss:Color=""#CCFFFF"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                End If



                'FOR EMPLOYEE DETAILS COLUMN ONLY
                _ExcelDoc.Write("<Style ss:ID=""ED"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")


                If DDLPaySlipType.SelectedValue.ToString = "38" And rbtshowclr.SelectedValue.ToString.ToUpper = "N" Then
                    _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""0""/>" & Chr(13) & "" & Chr(10) & "")
                    _ExcelDoc.Write("<Interior ss:Pattern=""Solid""/></Style>" & Chr(13) & "" & Chr(10) & "")
                Else
                    _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000"" ss:Bold=""0""/>" & Chr(13) & "" & Chr(10) & "")
                    _ExcelDoc.Write("<Interior ss:Color=""#CCFFFF"" ss:Pattern=""Solid""/></Style>" & Chr(13) & "" & Chr(10) & "")
                End If


                'FOR EXCEL HEARDER ONLY
                _ExcelDoc.Write("<Style ss:ID=""C2"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")


                If DDLPaySlipType.SelectedValue.ToString = "38" And rbtshowclr.SelectedValue.ToString.ToUpper = "N" Then
                    _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                    _ExcelDoc.Write("<Interior ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                Else
                    _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Color=""#000000"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                    _ExcelDoc.Write("<Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                End If


                'FOR COLUMN ONLY
                _ExcelDoc.Write("<Style ss:ID=""C3"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")

                If DDLPaySlipType.SelectedValue.ToString = "38" And rbtshowclr.SelectedValue.ToString.ToUpper = "N" Then
                    _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                    _ExcelDoc.Write("<Interior ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                Else
                    _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                    _ExcelDoc.Write("<Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                End If


                'FOR Header ONLY
                _ExcelDoc.Write("<Style ss:ID=""C4"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")

                If DDLPaySlipType.SelectedValue.ToString = "38" And rbtshowclr.SelectedValue.ToString.ToUpper = "N" Then
                    _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""10"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                    _ExcelDoc.Write("<Interior ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                Else
                    _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""10"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                    _ExcelDoc.Write("<Interior ss:Color=""#C2FDE2"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                End If


                _ExcelDoc.Write("<Style ss:ID=""C5"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""10"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")

                If DDLPaySlipType.SelectedValue.ToString = "38" And rbtshowclr.SelectedValue.ToString.ToUpper = "N" Then
                    _ExcelDoc.Write("<Interior ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                Else
                    _ExcelDoc.Write("<Interior ss:Color=""#F9966B"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                End If

                _ExcelDoc.Write("<Style ss:ID=""C6"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""10"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")

                If DDLPaySlipType.SelectedValue.ToString = "38" And rbtshowclr.SelectedValue.ToString.ToUpper = "N" Then
                    _ExcelDoc.Write("<Interior ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                Else
                    _ExcelDoc.Write("<Interior ss:Color=""#FFFF00"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                End If

                _ExcelDoc.Write("<Style ss:ID=""C7"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""10"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")

                If DDLPaySlipType.SelectedValue.ToString = "38" And rbtshowclr.SelectedValue.ToString.ToUpper = "N" Then
                    _ExcelDoc.Write("<Interior ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                Else
                    _ExcelDoc.Write("<Interior ss:Color=""#FF9900"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                End If

                _ExcelDoc.Write("<Style ss:ID=""C8"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""10"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")

                If DDLPaySlipType.SelectedValue.ToString = "38" And rbtshowclr.SelectedValue.ToString.ToUpper = "N" Then
                    _ExcelDoc.Write("<Interior ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                Else
                    _ExcelDoc.Write("<Interior ss:Color=""#FF6600"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                End If

                _ExcelDoc.Write("<Style ss:ID=""C9"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""10"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")

                If DDLPaySlipType.SelectedValue.ToString = "38" And rbtshowclr.SelectedValue.ToString.ToUpper = "N" Then
                    _ExcelDoc.Write("<Interior ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                Else
                    _ExcelDoc.Write("<Interior ss:Color=""#6495ED"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                End If

                'For Excel Header in Left Align
                _ExcelDoc.Write("<Style ss:ID=""C10"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")

                If DDLPaySlipType.SelectedValue.ToString = "38" And rbtshowclr.SelectedValue.ToString.ToUpper = "N" Then
                    _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                    _ExcelDoc.Write("<Interior ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                Else
                    _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Color=""#000000"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                    _ExcelDoc.Write("<Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                End If

                'For Excel Header in Left Align
                _ExcelDoc.Write("<Style ss:ID=""C11"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Right"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")

                If DDLPaySlipType.SelectedValue.ToString = "38" And rbtshowclr.SelectedValue.ToString.ToUpper = "N" Then
                    _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                    _ExcelDoc.Write("<Interior ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                Else
                    _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Color=""#000000"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                    _ExcelDoc.Write("<Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                End If

                _ExcelDoc.Write("</Styles>" & Chr(13) & "" & Chr(10) & " ")

                Dim start_ExcelXML As String = _ExcelDoc.ToString()
                Const end_ExcelXML As String = "</Workbook>"
                Dim sheetCount As Integer = 1

                ' _ExcelDoc.Write(start_ExcelXML)
                _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                If flag = "" Then
                    If DDLPaySlipType.SelectedValue.ToString = "41" Then
                        _ExcelDoc.Write("<Table><Column ss:Width=""40""/><Column ss:Width=""100""/><Column ss:Width=""180""/><Column ss:Width=""100""/><Column ss:Width=""225""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/>")
                    ElseIf DDLPaySlipType.SelectedValue.ToString = "42" Then
                        _ExcelDoc.Write("<Table><Column ss:Width=""40""/><Column ss:Width=""100""/><Column ss:Width=""180""/><Column ss:Width=""100""/><Column ss:Width=""225""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""140""/><Column ss:Width=""120""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/>")
                    ElseIf DDLPaySlipType.SelectedValue.ToString = "54" Then
                        _ExcelDoc.Write("<Table><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""100""/><Column ss:Width=""50""/><Column ss:Width=""225""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""50""/>")
                    Else
                        _ExcelDoc.Write("<Table><Column ss:Width=""40""/><Column ss:Width=""100""/><Column ss:Width=""140""/><Column ss:Width=""100""/><Column ss:Width=""80""/><Column ss:Width=""80""/><Column ss:Width=""50""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""240""/><Column ss:Width=""240""/><Column ss:Width=""240""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""240""/><Column ss:Width=""60""/><Column ss:Width=""100""/><Column ss:Width=""100""/><Column ss:Width=""120""/><Column ss:Width=""130""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/>")
                    End If
                Else
                    If DDLPaySlipType.SelectedValue.ToString = "41" Then
                        _ExcelDoc.Write("<Table><Column ss:Width=""40""/><Column ss:Width=""100""/><Column ss:Width=""180""/><Column ss:Width=""100""/><Column ss:Width=""225""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/>")
                    ElseIf DDLPaySlipType.SelectedValue.ToString = "42" Then
                        _ExcelDoc.Write("<Table><Column ss:Width=""40""/><Column ss:Width=""100""/><Column ss:Width=""180""/><Column ss:Width=""100""/><Column ss:Width=""225""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""140""/><Column ss:Width=""120""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/>")
                    Else
                        'START: Changed by Rohtas Singh on 22 Nov 2017 for fix all column width of excel sheet
                        Dim ColStr As String = ""
                        If source.Tables(0).Columns.Count > 138 Then
                            For colC As Integer = 139 To source.Tables(0).Columns.Count
                                ColStr = ColStr & "<Column ss:Width=""130""/>"
                            Next
                        End If
                        _ExcelDoc.Write("<Table><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/>" & ColStr)
                        'START: Changed by Rohtas Singh on 22 Nov 2017
                    End If
                End If

                'Added By geeta on 25 Jan 2013

                If DDLPaySlipType.SelectedValue.ToString = "41" Then

                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:StyleID=""C2"" ss:MergeAcross=""7""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("FORM - R")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")

                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:StyleID=""C2"" ss:MergeAcross=""7""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("REGISTER OF WAGES")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")

                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:StyleID=""C2"" ss:MergeAcross=""7""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("[See sub-rule (5) of Rule 11] Tamil Nadu Shops and Establishments Rules, 1948")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")

                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:MergeAcross=""7""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")

                    Dim _dt As New DataTable, _Arrparam(0) As SqlClient.SqlParameter
                    _Arrparam(0) = New SqlClient.SqlParameter("@CC_Code", USearch.UCddlcostcenter.ToString)
                    _dt = _ObjData.GetDataTableProc("sp_sel_companydetails_forreport", _Arrparam)

                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:StyleID=""C10"" ss:MergeAcross=""5""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("Name and Address of the Establishment : " & _dt.Rows(0)("Comp_Name").ToString & ", " & _dt.Rows(0)("Add1").ToString & " " & _dt.Rows(0)("Add2").ToString & IIf(_dt.Rows(0)("Add3").ToString.Trim <> "", ", " & _dt.Rows(0)("Add3").ToString, "").ToString)
                    _ExcelDoc.Write("</Data></Cell>")

                    _ExcelDoc.Write("<Cell ss:StyleID=""C10"" ss:MergeAcross=""1""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("Wages Period : " & ddlMonthYear.SelectedItem.ToString)
                    _ExcelDoc.Write("</Data></Cell>")

                    _ExcelDoc.Write("</Row>")

                End If
                If DDLPaySlipType.SelectedValue.ToString <> "41" And DDLPaySlipType.SelectedValue.ToString <> "54" Then
                    _ExcelDoc.Write("<Row>")
                    If flag = "" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""C2"" ss:MergeAcross=""2""><Data ss:Type=""String"">")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""C2"" ss:MergeAcross=""1""><Data ss:Type=""String"">")
                    End If

                    If DDLPaySlipType.SelectedValue.ToString = "42" Then
                        _ExcelDoc.Write("Maharashtra Minimum Wages Salary Register (" & ddlMonthYear.SelectedItem.ToString & ")")
                    Else
                        If DDLPaySlipType.SelectedValue.ToString = "61" Then
                            _ExcelDoc.Write("Paid Holiday Report for the Month of " & ddlMonthYear.SelectedItem.ToString)
                        Else
                            _ExcelDoc.Write("Salary Register for the Month of " & ddlMonthYear.SelectedItem.ToString)
                        End If

                    End If
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                End If

                'added by Rohtas Singh on 08 Dec 2017 for not print blank line when select "Monthly Salary Slip (MAX Life)"
                If DDLPaySlipType.SelectedValue.ToString <> "54" Then
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell><Data ss:Type=""String"">")
                    _ExcelDoc.Write("")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                End If

                'below code is used to display Header according to paycode wise, Added by praveen verma on 19 Feb 2013.
                If source.Tables.Count > 1 And DDLPaySlipType.SelectedValue.ToString <> "21" And DDLPaySlipType.SelectedValue.ToString <> "61" Then
                    If source.Tables(1).Rows.Count > 0 Then
                        _ExcelDoc.Write("<Row>")

                        For j As Integer = 0 To source.Tables(1).Rows.Count - 1
                            Select Case source.Tables(1).Rows(j).Item(2).ToString.ToUpper.Trim
                                Case "D"
                                    headstr = "DEDUCTION"
                                    Stlname = "C8"
                                Case "A"
                                    headstr = "EARNINGS"
                                    Stlname = "C7"
                                Case "R"
                                    headstr = "RATE"
                                    Stlname = "C6"
                                Case "AR"
                                    headstr = "ARREAR"
                                    Stlname = "C4"
                                Case "DR"
                                    headstr = "ARREAR"
                                    Stlname = "C9"
                                Case "L"
                                    headstr = "LEAVE"
                                    Stlname = "C5"
                            End Select
                            For x As Integer = K To source.Tables(0).Columns.Count - 1
                                If x = CInt(source.Tables(1).Rows(j).Item(0)) - 1 Then
                                    _ExcelDoc.Write("<Cell ss:StyleID=""" & Stlname & """ ss:MergeAcross=""" & source.Tables(1).Rows(j).Item(1).ToString - 1 & """ ><Data ss:Type=""String"">")
                                    _ExcelDoc.Write(headstr)
                                    _ExcelDoc.Write("</Data></Cell>")
                                    K = x + CInt(source.Tables(1).Rows(j).Item(1))
                                    Exit For
                                Else
                                    _ExcelDoc.Write("<Cell><Data ss:Type=""String"">")
                                    _ExcelDoc.Write("")
                                    _ExcelDoc.Write("</Data></Cell>")
                                End If
                            Next
                        Next
                        _ExcelDoc.Write("</Row>")
                    End If
                End If

                _ExcelDoc.Write("<Row>")
                _ExcelXML.Append("<Row>")

                Dim dtCopy As DataTable = New DataTable()
                dtCopy = source.Tables(0).Clone

                If (DDLPaySlipType.SelectedValue.Equals("38")) Then

                    Dim RC1 = source.Tables(0).Select("SNO<>'" & source.Tables(0).Rows.Count.ToString & "'")

                    For Each drtableOld As DataRow In RC1
                        dtCopy.ImportRow(drtableOld)
                    Next
                End If

                'Pankaj Sachan
                For x As Integer = 0 To source.Tables(0).Columns.Count - 1
                    Dim _dt As New DataTable, _colName As String = source.Tables(0).Columns(x).ColumnName.ToUpper
                    _ExcelDoc.Write("<Cell ss:StyleID=""C2""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(_colName)
                    _ExcelDoc.Write("</Data></Cell>")

                    If (DDLPaySlipType.SelectedValue.Equals("38")) Then

                        Dim strColDataTypeCheck As String = IIf(IsDBNull(source.Tables(0).Columns(x)), "", source.Tables(0).Columns(x).DataType.ToString())
                        Dim filtered = From r In dtCopy.AsEnumerable Where (IsDBNull(_colName) = False And IsNumeric(r(_colName)) = 0)
                                       Select r

                        If filtered.Count = 0 And strColDataTypeCheck <> "System.String" And _colName <> "PROFIT_CENTRE" And _colName <> "AADHAR CARD" And _colName <> "ESINO" And _colName <> "AADHAR ENROLEMENT NUMBER" And _colName <> "BANK_ACC_NO" And _colName <> "PFNO" And _colName <> "PF UAN" And _colName <> "SITE CODE" And _colName <> "COST CENTER NO." And _colName <> "SAP_CODE" And _colName <> "IFSC" And _colName <> "CODE" And _colName <> "EMPLOYEECODE" And _colName <> "BANKACNO" And _colName <> "ESINO" And _colName <> "PFNO" And _colName <> "CMSGINO" And _colName <> "STAFF_ID" And _colName <> "IFSC" And _colName.ToString.ToUpper.Trim <> _objCommon.DisplayCaption("EMPCODE").ToString.ToUpper.Trim Then
                            _ExcelXML.Append("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">{" & x & "}</Data></Cell>")
                        Else
                            _ExcelXML.Append("<Cell ss:StyleID=""ED""><Data ss:Type=""String"">{" & x & "}</Data></Cell>")
                        End If

                    Else
                        Dim filtered = From r In source.Tables(0).AsEnumerable Where (IsDBNull(_colName) = False And IsNumeric(r(_colName)) = 0)
                                       Select r
                        If filtered.Count = 0 And _colName <> "CODE" And _colName <> "EMPLOYEECODE" And _colName <> "BANKACNO" And _colName <> "ESINO" And _colName <> "PFNO" And _colName <> "CMSGINO" And _colName <> "STAFF_ID" And _colName <> "IFSC" And _colName.ToString.ToUpper.Trim <> _objCommon.DisplayCaption("EMPCODE").ToString.ToUpper.Trim Then
                            _ExcelXML.Append("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">{" & x & "}</Data></Cell>")
                        Else
                            _ExcelXML.Append("<Cell ss:StyleID=""ED""><Data ss:Type=""String"">{" & x & "}</Data></Cell>")

                        End If
                    End If

                Next

                dtCopy.Dispose()

                _ExcelXML.Append("</Row>")
                _ExcelDoc.Write("</Row>")
                Dim rflag As Integer = 0
                For Each x As DataRow In source.Tables(0).Rows
                    _RowCount += 1

                    If _RowCount = source.Tables(0).Rows.Count And Not DDLPaySlipType.SelectedValue.Equals("61") Then
                        _ExcelDoc.Write("<Row>")
                        For y As Integer = 0 To source.Tables(0).Columns.Count - 1
                            Dim XMLstring As String = x(y).ToString()
                            XMLstring = XMLstring.Trim()
                            If _RowCount = source.Tables(0).Rows.Count And y = 0 And DDLPaySlipType.SelectedValue.ToString <> "54" Then
                                XMLstring = ""
                            End If
                            If Convert.ToString(source.Tables(0).Columns(y).DataType).Equals("System.Decimal") Or Convert.ToString(source.Tables(0).Columns(y).DataType).Equals("System.Int64") Then
                                _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write(XMLstring)
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                                _ExcelDoc.Write(XMLstring.Replace("&", "&amp;"))
                            End If
                            'If IsNumeric(XMLstring) Then
                            '    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""Number"">")
                            '    _ExcelDoc.Write(XMLstring)
                            'Else
                            '    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                            '    _ExcelDoc.Write(XMLstring.Replace("&", "&amp;"))
                            'End If


                            _ExcelDoc.Write("</Data></Cell>")
                        Next
                        _ExcelDoc.Write("</Row>")
                    Else
                        _ExcelDoc.Write(String.Format(_ExcelXML.ToString, x.ItemArray))
                    End If
                Next
                _ExcelDoc.Write("</Table>")
                If DDLPaySlipType.SelectedValue.ToString = "41" Then
                    _ExcelDoc.Write("<WorksheetOptions xmlns=""urn:schemas-microsoft-com:office:excel""><Selected/><FreezePanes/><FrozenNoSplit/><SplitHorizontal>7</SplitHorizontal><TopRowBottomPane>7</TopRowBottomPane><SplitVertical>6</SplitVertical><LeftColumnRightPane>6</LeftColumnRightPane><ActivePane>0</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>1</Number></Pane><Pane><Number>2</Number></Pane><Pane><Number>0</Number></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions>")
                Else
                    _ExcelDoc.Write("<WorksheetOptions xmlns=""urn:schemas-microsoft-com:office:excel""><Selected/><FreezePanes/><FrozenNoSplit/><SplitHorizontal>3</SplitHorizontal><TopRowBottomPane>3</TopRowBottomPane><SplitVertical>3</SplitVertical><LeftColumnRightPane>3</LeftColumnRightPane><ActivePane>0</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>1</Number></Pane><Pane><Number>2</Number></Pane><Pane><Number>0</Number><ActiveCol>0</ActiveCol></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions>")
                End If

                _ExcelDoc.Write(" </Worksheet>")

                If DDLPaySlipType.SelectedValue.ToString = "38" Then
                    TblNumber = 2
                    'Added by Rohtas Singh on 01 Jun 2020 
                ElseIf DDLPaySlipType.SelectedValue.Equals("61") Or DDLPaySlipType.SelectedValue.Equals("41") Then
                    TblNumber = 0
                Else
                    TblNumber = 1
                End If


                'Salary register in excel with hold
                If source.Tables(TblNumber).Rows.Count > 0 Then

                    If DDLPaySlipType.SelectedValue.ToString = "21" Or DDLPaySlipType.SelectedValue.ToString = "38" Then

                        _ExcelXML = New StringBuilder

                        _RowCount = 0
                        headstr = ""
                        K = 0
                        Stlname = ""

                        _ExcelDoc.Write("<Worksheet ss:Name=""Release Salary"">")
                        '_ExcelDoc.Write("<Table>")


                        ''Start
                        flag = ""
                        If flag = "" Then
                            _ExcelDoc.Write("<Table><Column ss:Width=""40""/><Column ss:Width=""100""/><Column ss:Width=""140""/><Column ss:Width=""100""/><Column ss:Width=""80""/><Column ss:Width=""80""/><Column ss:Width=""50""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""240""/><Column ss:Width=""240""/><Column ss:Width=""240""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""240""/><Column ss:Width=""60""/><Column ss:Width=""100""/><Column ss:Width=""100""/><Column ss:Width=""120""/><Column ss:Width=""130""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""120""/><Column ss:Width=""120""/><Column ss:Width=""150""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/>")
                        Else
                            'START: Changed by Rohtas Singh on 22 Nov 2017 for fix all column width of excel sheet
                            Dim ColStr As String = ""
                            If source.Tables(TblNumber).Columns.Count > 138 Then
                                For colC As Integer = 139 To source.Tables(TblNumber).Columns.Count
                                    ColStr = ColStr & "<Column ss:Width=""130""/>"
                                Next
                            End If
                            _ExcelDoc.Write("<Table><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/><Column ss:Width=""130""/>" & ColStr)
                            'START: Changed by Rohtas Singh on 22 Nov 2017
                        End If

                        If DDLPaySlipType.SelectedValue.ToString <> "41" And DDLPaySlipType.SelectedValue.ToString <> "54" Then
                            _ExcelDoc.Write("<Row>")
                            If flag = "" Then
                                _ExcelDoc.Write("<Cell ss:StyleID=""C2"" ss:MergeAcross=""2""><Data ss:Type=""String"">")
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""C2"" ss:MergeAcross=""1""><Data ss:Type=""String"">")
                            End If

                            'If DDLPaySlipType.SelectedValue.ToString = "42" Then
                            '   _ExcelDoc.Write("Maharashtra Minimum Wages Salary Register (" & ddlMonthYear.SelectedItem.ToString & ")")
                            'Else
                            _ExcelDoc.Write("Release Salary Register for the Month of " & ddlMonthYear.SelectedItem.ToString)
                            'End If
                            _ExcelDoc.Write("</Data></Cell>")
                            _ExcelDoc.Write("</Row>")
                        End If

                        'added by Rohtas Singh on 08 Dec 2017 for not print blank line when select "Monthly Salary Slip (MAX Life)"
                        If DDLPaySlipType.SelectedValue.ToString <> "54" Then
                            _ExcelDoc.Write("<Row>")
                            _ExcelDoc.Write("<Cell><Data ss:Type=""String"">")
                            _ExcelDoc.Write("")
                            _ExcelDoc.Write("</Data></Cell>")
                            _ExcelDoc.Write("</Row>")
                        End If


                        _ExcelDoc.Write("<Row>")
                        _ExcelXML.Append("<Row>")
                        For x As Integer = 0 To source.Tables(TblNumber).Columns.Count - 1
                            Dim _dt As New DataTable, _colName As String = source.Tables(TblNumber).Columns(x).ColumnName.ToUpper
                            _ExcelDoc.Write("<Cell ss:StyleID=""C2""><Data ss:Type=""String"">")
                            _ExcelDoc.Write(_colName)
                            _ExcelDoc.Write("</Data></Cell>")

                            Dim filtered = From r In source.Tables(TblNumber).AsEnumerable Where (IsDBNull(_colName) = False And IsNumeric(r(_colName)) = 0)
                                           Select r


                            If filtered.Count = 0 And _colName <> "CODE" And _colName <> "BANKACNO" And _colName <> "ESINO" And _colName <> "PFNO" And _colName <> "CMSGINO" And _colName <> "STAFF_ID" And _colName <> "IFSC" And _colName.ToString.ToUpper.Trim <> _objCommon.DisplayCaption("EMPCODE").ToString.ToUpper.Trim Then
                                _ExcelXML.Append("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">{" & x & "}</Data></Cell>")
                            Else
                                _ExcelXML.Append("<Cell ss:StyleID=""ED""><Data ss:Type=""String"">{" & x & "}</Data></Cell>")
                            End If
                        Next
                        _ExcelXML.Append("</Row>")
                        _ExcelDoc.Write("</Row>")
                        For Each x As DataRow In source.Tables(TblNumber).Rows
                            _RowCount += 1

                            If _RowCount = source.Tables(TblNumber).Rows.Count Then
                                _ExcelDoc.Write("<Row>")
                                For y As Integer = 0 To source.Tables(TblNumber).Columns.Count - 1
                                    Dim XMLstring As String = x(y).ToString()
                                    XMLstring = XMLstring.Trim()
                                    If _RowCount = source.Tables(TblNumber).Rows.Count And y = 0 And DDLPaySlipType.SelectedValue.ToString <> "54" Then
                                        XMLstring = ""
                                    End If
                                    If Convert.ToString(DDLPaySlipType.SelectedValue).Equals(38) Then
                                        Dim strcol As String = source.Tables(TblNumber).Columns(1).DataType.ToString()
                                        'If System.Type.GetType(source.Tables(TblNumber).Columns(1).DataType.ToString()).Equals() Then

                                        'End If
                                        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                                    Else
                                        If IsNumeric(XMLstring) Then
                                            _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""Number"">")
                                        Else
                                            _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                                        End If
                                    End If
                                    _ExcelDoc.Write(XMLstring)
                                    _ExcelDoc.Write("</Data></Cell>")
                                Next
                                _ExcelDoc.Write("</Row>")
                            Else
                                _ExcelDoc.Write(String.Format(_ExcelXML.ToString, x.ItemArray))
                            End If
                        Next
                        _ExcelDoc.Write("</Table>")

                        _ExcelDoc.Write("<WorksheetOptions xmlns=""urn:schemas-microsoft-com:office:excel""><Selected/><FreezePanes/><FrozenNoSplit/><SplitHorizontal>3</SplitHorizontal><TopRowBottomPane>3</TopRowBottomPane><SplitVertical>3</SplitVertical><LeftColumnRightPane>3</LeftColumnRightPane><ActivePane>0</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>1</Number></Pane><Pane><Number>2</Number></Pane><Pane><Number>0</Number><ActiveCol>0</ActiveCol></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions>")
                        ''end

                        'DISPLAYPayheadcount(source.Tables(3), _ExcelDoc)
                        '_ExcelDoc.Write("</Table>")
                        _ExcelDoc.Write(" </Worksheet>")

                    End If
                End If
                _ExcelDoc.Write(end_ExcelXML)
            Catch ex As Exception
                'for catch the error
                _objcommonExp.PublishError("For gebnerate the excel file (fnExportTo_ExcelXML())", ex)
            End Try
            'Return _ExcelDoc
        End Sub

        Private Sub ExportToExcel(ByVal source As DataSet, ByRef _ExcelDoc As StreamWriter)
            Dim _RowCount As Integer = 0
            'Dim _ExcelDoc As System.IO.StringWriter
            '_ExcelDoc = New System.IO.StringWriter()
            Dim Z, W As Integer
            Try
                Dim _ExcelXML As New StringBuilder()
                _ExcelDoc.Write("<?xml version=""1.0""?><?mso-application progid=""Excel.Sheet""?>" & Chr(13) & "" & Chr(10) & "<Workbook ")
                _ExcelDoc.Write("xmlns=""urn:schemas-microsoft-com:office:spreadsheet""" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write(" xmlns:o=""urn:schemas-microsoft-com:office:office""" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write(" xmlns:x=""urn:schemas-microsoft-com:office:excel""" & Chr(13) & "" & Chr(10))
                _ExcelDoc.Write(" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">" & Chr(13) & "" & Chr(10) & " <Styles>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""Default"" ss:Name=""Normal"">" & Chr(13) & "" & Chr(10) & " ")
                '_ExcelXML.Append("<Alignment ss:Horizontal=\"Right\" ss:Vertical=\"Bottom\"/>\r\n <Borders/>"); 
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>" & Chr(13) & "" & Chr(10) & " <Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /></Borders>")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Font/>" & Chr(13) & "" & Chr(10) & " <Interior/>" & Chr(13) & "" & Chr(10) & " <NumberFormat/>")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Protection/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")

                'Cumstomise format 
                _ExcelDoc.Write("<Style ss:ID=""CD"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelDoc.Write("ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Bold=""1"" />" & Chr(13) & "" & Chr(10) & " <Interior ss:Color=""#CDAF95"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & " ")

                '_ExcelXML.Append("\r\n <Style ss:ID=\"StyleBorder\"><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>"); 
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Style ss:ID=""BC"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelDoc.Write("x:Family=""Swiss"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""SL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat")
                _ExcelDoc.Write(" ss:Format=""@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _ExcelDoc.Write("ss:ID=""DC"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""0.0000""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")

                _ExcelDoc.Write("<Style ss:ID=""IT"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""0""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _ExcelDoc.Write("ss:ID=""DL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""mm/dd/yyyy;@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")

                _ExcelDoc.Write("<Style ss:ID=""s2"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Top"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:Size=""9"" ss:Color=""#808080"" ss:Underline=""Single""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#FFFFFF"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("</Styles>" & Chr(13) & "" & Chr(10) & " ")

                Dim start_ExcelXML As String = _ExcelDoc.ToString()
                Const end_ExcelXML As String = "</Workbook>"
                Dim sheetCount As Integer = 1

                '_ExcelDoc.Write(start_ExcelXML)
                _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                _ExcelDoc.Write("<Table><Column ss:Width=""40""/><Column ss:Width=""60""/><Column ss:Width=""120""/><Column ss:Width=""80""/><Column ss:Width=""120""/><Column ss:Width=""70""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""60""/><Column ss:Width=""50""/><Column ss:Width=""80""/><Column ss:Width=""80""/><Column ss:Width=""120""/>")
                _ExcelDoc.Write("<Row>")

                For x As Integer = 0 To 1
                    If x = 1 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("FORM  II")
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")
                _ExcelDoc.Write("<Row>")

                For x As Integer = 0 To 1
                    If x = 1 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""3"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("[Rule 27 (1) }")
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")
                _ExcelDoc.Write("<Row>")

                For x As Integer = 0 To 1
                    If x = 1 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""4"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Muster-roll-cum Wages Register")
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")
                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To 1 - 1
                    If x < 1 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""4"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Name of Employer/Controller  :")
                        _ExcelDoc.Write(source.Tables(0).Rows(0).Item("Comp_Name").ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")
                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To 1 - 1
                    If x < 1 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""6"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("For the month of ")
                        _ExcelDoc.Write(MonthName(source.Tables(0).Rows(0).Item("SelMonth").ToString))
                        _ExcelDoc.Write("   ")
                        _ExcelDoc.Write("Year ")
                        _ExcelDoc.Write(source.Tables(0).Rows(0).Item("SelYear").ToString)
                        _ExcelDoc.Write("</Data></Cell>")

                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To 1 - 1
                    If x < 1 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")

                '''''''''''''''''''''''''''''''Change
                _ExcelDoc.Write("<Row>")
                For Lx As Integer = 0 To source.Tables(2).Columns.Count - 4
                    If Lx = (source.Tables(2).Columns.Count - 6) Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""3"" ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Leave")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                    End If
                    _ExcelDoc.Write("</Data></Cell>")
                Next
                _ExcelDoc.Write("</Row>")


                Dim Ax As Integer = source.Tables(4).Rows(0).Item("AddCount").ToString
                Dim Dx As Integer = source.Tables(5).Rows(0).Item("DedCount").ToString
                '''''''''''''''''''''''''''''''
                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To source.Tables(1).Columns.Count - 1

                    If source.Tables(1).Columns(x).ColumnName = "Srno" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("sr no")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "Dob" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Age and Sex")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "Desig" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Nature of work Destination")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "Doj" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Date of entry into service")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "Worhr" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Working Hours From To")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "interval" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Interval for rest From To")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "Hrswrkd" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Hours worked on 1,2,3,4,..31")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "paiddays" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Total days worked")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "RatBasic" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Minimum Rates of wages payable")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "piece" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Total production in case of piece rate")

                    ElseIf source.Tables(1).Columns(x).ColumnName = "othrs" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Total overtime hours worked")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "rategross" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Actual Rates of wages payable")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "RateHra" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Rate of H.R.A.")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "PayableHra" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("H.R.A. payable")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "Gross" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Gross wages payable")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "Arrears_Ot" Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Arrears")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "Nominal Wages" Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""" & Ax - 1 & """ ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Nominal Wages")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "Overtime earnings" Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""1"" ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Overtime earnings")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "Dedection -  Advanced,  Fine,  Demage,  Others" Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""" & Dx - 1 & """ ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Dedection -  Advanced,  Fine,  Demage,  Others")

                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(source.Tables(1).Columns(x).ColumnName)
                    End If
                    _ExcelDoc.Write("</Data></Cell>")
                Next
                _ExcelDoc.Write("</Row>")
                ''''''''''''''''''''''''''
                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To source.Tables(2).Columns.Count - 1
                    If x >= (13) And x <= (13 + Ax) Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(source.Tables(2).Columns(x).ColumnName)

                    ElseIf x > (13 + Ax) And x <= (16 + Ax) Then

                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(source.Tables(2).Columns(x).ColumnName)

                    ElseIf x > (16 + Ax) And x <= (17 + Ax + Dx) Then

                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(source.Tables(2).Columns(x).ColumnName)
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                    End If
                    _ExcelDoc.Write("</Data></Cell>")

                Next
                _ExcelDoc.Write("</Row>")

                '''''''''''''''''''''''''
                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To source.Tables(1).Columns.Count - 1

                    If source.Tables(1).Columns(x).ColumnName = "Nominal Wages" Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""" & Ax - 1 & """ ss:StyleID=""CD""><Data ss:Type=""Number"">")
                        _ExcelDoc.Write(x + 1)
                    ElseIf source.Tables(1).Columns(x).ColumnName = "Overtime earnings" Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""1"" ss:StyleID=""CD""><Data ss:Type=""Number"">")
                        _ExcelDoc.Write(x + 1)
                    ElseIf source.Tables(1).Columns(x).ColumnName = "Dedection -  Advanced,  Fine,  Demage,  Others" Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""" & Dx - 1 & """ ss:StyleID=""CD""><Data ss:Type=""Number"">")
                        _ExcelDoc.Write(x + 1)

                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""Number"">")
                        _ExcelDoc.Write(x + 1)
                    End If
                    _ExcelDoc.Write("</Data></Cell>")
                Next
                _ExcelDoc.Write("</Row>")
                '''''''''''''''''''''''''



                For Each x As DataRow In source.Tables(2).Rows
                    _RowCount += 1
                    'if the number of rows is > 63000 create a new page to continue output 
                    'If _RowCount = 63000 Then
                    '    _RowCount = 0
                    '    sheetCount += 1
                    '    _ExcelDoc.Write("</Table>")
                    '    _ExcelDoc.Write(" </Worksheet>")
                    '    _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                    '    _ExcelDoc.Write("<Table>")
                    '    _ExcelDoc.Write("<Row>")
                    '    For xi As Integer = 0 To source.Tables(2).Columns.Count - 1
                    '        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write(source.Tables(1).Columns(xi).ColumnName)
                    '        _ExcelDoc.Write("</Data></Cell>")
                    '    Next
                    '    _ExcelDoc.Write("</Row>")
                    'End If
                    _ExcelDoc.Write("<Row>")
                    For y As Integer = 0 To source.Tables(2).Columns.Count - 1
                        Dim XMLstring As String = x(y).ToString()
                        XMLstring = XMLstring.Trim()

                        If IsNumeric(XMLstring) Then
                            _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""Number"">")
                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""String"">")
                        End If

                        _ExcelDoc.Write(XMLstring)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    _ExcelDoc.Write("</Row>")
                Next
                '---By GRand
                For Each x As DataRow In source.Tables(3).Rows
                    _RowCount += 1
                    'if the number of rows is > 63000 create a new page to continue output 
                    'If _RowCount = 63000 Then
                    '    _RowCount = 0
                    '    sheetCount += 1
                    '    _ExcelDoc.Write("</Table>")
                    '    _ExcelDoc.Write(" </Worksheet>")
                    '    _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                    '    _ExcelDoc.Write("<Table>")
                    '    _ExcelDoc.Write("<Row>")
                    '    For xi As Integer = 0 To source.Tables(3).Columns.Count - 1
                    '        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write(source.Tables(2).Columns(xi).ColumnName)
                    '        _ExcelDoc.Write("</Data></Cell>")
                    '    Next
                    '    _ExcelDoc.Write("</Row>")
                    'End If
                    _ExcelDoc.Write("<Row>")
                    For y As Integer = 0 To source.Tables(3).Columns.Count - 1
                        Z = y
                        W = y
                        Dim XMLstring As String = x(y).ToString()
                        XMLstring = XMLstring.Trim()
                        If IsNumeric(XMLstring) Then
                            _ExcelDoc.Write("<Cell ss:StyleID=""BC"">" + "<Data ss:Type=""Number"">")
                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""BC"">" + "<Data ss:Type=""String"">")
                        End If
                        _ExcelDoc.Write(XMLstring)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    _ExcelDoc.Write("</Row>")
                    Dim XMLCusto As String
                    Dim XMLst As String
                    XMLst = ""
                    XMLCusto = ""
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                    XMLCusto = "Signatue of authorised"
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:MergeAcross=""4"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(XMLCusto)
                    _ExcelDoc.Write("</Data></Cell>")

                    '''''''''''''''''''''''''''''''''
                    For P As Integer = 1 To Z - 1
                        If P = W - 1 Then
                            _ExcelDoc.Write("<Cell ss:MergeAcross=""4"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                            XMLst = "Signature of the Employer or the Person authorised"
                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""BC"">" + "<Data ss:Type=""String"">")
                        End If

                        _ExcelDoc.Write(XMLst)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next


                    '''''''''''''''''''''''''''''''''
                    XMLst = ""
                    _ExcelDoc.Write("</Row>")
                    XMLCusto = "Representative Principal Employer"
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:MergeAcross=""4"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(XMLCusto)
                    _ExcelDoc.Write("</Data></Cell>")
                    '''''''''''''''''''''''''''''''''
                    For P As Integer = 1 To Z - 1
                        If P = W - 1 Then
                            _ExcelDoc.Write("<Cell ss:MergeAcross=""4"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                            XMLst = "by him to authenticate the above entries."
                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""BC"">" + "<Data ss:Type=""String"">")
                        End If

                        _ExcelDoc.Write(XMLst)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next


                    '''''''''''''''''''''''''''''''''
                    _ExcelDoc.Write("</Row>")
                    XMLCusto = "(in case of Contract Labour)"
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:MergeAcross=""4"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(XMLCusto)
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")


                Next


                _ExcelDoc.Write("</Table>")
                _ExcelDoc.Write(" </Worksheet>")
                _ExcelDoc.Write(end_ExcelXML)
            Catch ex As Exception
                'for catch the error
                _objcommonExp.PublishError("For gebnerate the excel file (fnExportTo_ExcelXML())", ex)
            End Try
            'Return _ExcelDoc
        End Sub
        Private Sub ExportToExcelFoStaff(ByVal source As DataSet, ByRef _ExcelDoc As StreamWriter)
            Dim _RowCount As Integer = 0
            ' Dim _ExcelDoc As System.IO.StringWriter
            '_ExcelDoc = New System.IO.StringWriter()
            Try
                Dim _ExcelXML As New StringBuilder()
                _ExcelDoc.Write("<?xml version=""1.0""?><?mso-application progid=""Excel.Sheet""?>" & Chr(13) & "" & Chr(10) & "<Workbook ")
                _ExcelDoc.Write("xmlns=""urn:schemas-microsoft-com:office:spreadsheet""" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write(" xmlns:o=""urn:schemas-microsoft-com:office:office""" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write(" xmlns:x=""urn:schemas-microsoft-com:office:excel""" & Chr(13) & "" & Chr(10))
                _ExcelDoc.Write(" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">" & Chr(13) & "" & Chr(10) & " <Styles>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""Default"" ss:Name=""Normal"">" & Chr(13) & "" & Chr(10) & " ")
                '_ExcelXML.Append("<Alignment ss:Horizontal=\"Right\" ss:Vertical=\"Bottom\"/>\r\n <Borders/>"); 
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>" & Chr(13) & "" & Chr(10) & " <Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /></Borders>")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Font/>" & Chr(13) & "" & Chr(10) & " <Interior/>" & Chr(13) & "" & Chr(10) & " <NumberFormat/>")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Protection/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                'Cumstomise format 
                _ExcelDoc.Write("<Style ss:ID=""CD"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Center"" ss:Vertical=""Bottom"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelDoc.Write("ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Bold=""1"" />" & Chr(13) & "" & Chr(10) & " <Interior ss:Color=""#CDAF95"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & " ")
                '_ExcelXML.Append("\r\n <Style ss:ID=\"StyleBorder\"><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>"); 
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Style ss:ID=""BC"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelDoc.Write("x:Family=""Swiss"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""SL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat")
                _ExcelDoc.Write(" ss:Format=""@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _ExcelDoc.Write("ss:ID=""DC"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""0.0000""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""IT"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""0""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _ExcelDoc.Write("ss:ID=""DL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""mm/dd/yyyy;@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""s2"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Top"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:Size=""9"" ss:Color=""#808080"" ss:Underline=""Single""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#FFFFFF"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("</Styles>" & Chr(13) & "" & Chr(10) & " ")

                Dim start_ExcelXML As String = _ExcelDoc.ToString()
                Const end_ExcelXML As String = "</Workbook>"
                Dim sheetCount As Integer = 1

                '_ExcelDoc.Write(start_ExcelXML)
                _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                _ExcelDoc.Write("<Table><Column ss:Width=""40""/><Column ss:Width=""60""/><Column ss:Width=""120""/><Column ss:Width=""80""/><Column ss:Width=""120""/><Column ss:Width=""80""/><Column ss:Width=""80""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""50""/><Column ss:Width=""80""/><Column ss:Width=""80""/><Column ss:Width=""100""/><Column ss:Width=""65""/><Column ss:Width=""65""/><Column ss:Width=""65""/>")
                _ExcelDoc.Write("<Row>")

                Dim str As String = ""
                str = System.Configuration.ConfigurationManager.AppSettings("BASE_HREF").ToString() & Session("CompCode").ToString & "\CompanyBanner\" & source.Tables(0).Rows(0).Item("LogoName").ToString
                For x As Integer = 0 To 10
                    If x = 9 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(source.Tables(0).Rows(0).Item("Comp_Name").ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")
                _ExcelDoc.Write("<Row>")

                For x As Integer = 0 To 10
                    If x = 8 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""4"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("SALARY FOR THE MONTH OF--")
                        _ExcelDoc.Write(source.Tables(0).Rows(0).Item("Mon").ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")
                _ExcelDoc.Write("<Row>")

                For x As Integer = 0 To 10
                    If x = 9 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""4"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("LOCATION--")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write(source.Tables(0).Rows(0).Item("LOCATION").ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To 1 - 1
                    If x < 1 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Dim Px As Integer
                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To 9
                    _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                    If source.Tables(1).Columns(x).ColumnName = "SrNo" Then
                        _ExcelDoc.Write("SR.NO.")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "EmpCode" Then
                        _ExcelDoc.Write("EMPLOYEE CODE")


                    Else
                        _ExcelDoc.Write(source.Tables(1).Columns(x).ColumnName)
                    End If
                    _ExcelDoc.Write("</Data></Cell>")
                Next

                Px = (source.Tables(1).Columns.Count - 1) - 3
                Dim Wx, Zx As String
                Wx = source.Tables(3).Rows(0).Item("AddCount").ToString
                Zx = source.Tables(4).Rows(0).Item("DedCount").ToString
                For x As Integer = 10 To 17
                    If x = 10 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""3"" ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("LEAVE")
                    End If
                    If x = 11 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("TNDP")
                    End If
                    If x = 12 Then
                        '_ExcelDoc.Write("<Cell ss:MergeAcross=""6"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                        '_ExcelDoc.Write("EARNINGS")
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""" & (Wx - 1) & """ ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("EARNINGS")

                    End If
                    If x = 13 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("GROSS AMT")
                    End If
                    If x = 14 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""" & (Zx - 1) & """ ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("DEDUCTIONS")

                    End If
                    If x = 15 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("TOTAL DEDUCTION")

                    End If
                    If x = 16 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("NET PAYMENT")

                    End If

                    If x = 17 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("SIGNATURE")

                    End If
                    _ExcelDoc.Write("</Data></Cell>")
                Next
                _ExcelDoc.Write("</Row>")



                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To source.Tables(1).Columns.Count - 1
                    If x >= 10 And x <= Px Then

                        If source.Tables(1).Columns(x).ColumnName = "ArrearGross" Then
                            _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                            _ExcelDoc.Write("Arrears")

                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                            _ExcelDoc.Write(source.Tables(1).Columns(x).ColumnName)

                        End If

                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                    End If
                    _ExcelDoc.Write("</Data></Cell>")
                Next
                _ExcelDoc.Write("</Row>")


                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                For Each x As DataRow In source.Tables(1).Rows
                    _RowCount += 1
                    'if the number of rows is > 63000 create a new page to continue output 
                    'If _RowCount = 63000 Then
                    '    Dim test As String = ""
                    '    test = "god"
                    'End If
                    'If _RowCount = 63000 Then
                    '    _RowCount = 0
                    '    sheetCount += 1
                    '    _ExcelDoc.Write("</Table>")
                    '    _ExcelDoc.Write(" </Worksheet>")
                    '    _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                    '    _ExcelDoc.Write("<Table>")
                    '    _ExcelDoc.Write("<Row>")
                    '    For xi As Integer = 0 To source.Tables(1).Columns.Count - 1
                    '        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write(source.Tables(1).Columns(xi).ColumnName)
                    '        _ExcelDoc.Write("</Data></Cell>")
                    '    Next
                    '    _ExcelDoc.Write("</Row>")
                    'End If
                    _ExcelDoc.Write("<Row>")
                    For y As Integer = 0 To source.Tables(1).Columns.Count - 1
                        Dim XMLstring As String = x(y).ToString()
                        XMLstring = XMLstring.Trim()

                        If IsNumeric(XMLstring) Then
                            If y = 6 Then
                                _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""String"">")
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""Number"">")
                            End If

                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""String"">")
                        End If
                        _ExcelDoc.Write(XMLstring)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    _ExcelDoc.Write("</Row>")
                Next
                '---By GRand
                For Each x As DataRow In source.Tables(2).Rows
                    _RowCount += 1
                    'if the number of rows is > 63000 create a new page to continue output 
                    'If _RowCount = 63000 Then
                    '    Dim test As String = ""
                    '    test = "god"
                    'End If
                    'If _RowCount = 63000 Then
                    '    _RowCount = 0
                    '    sheetCount += 1
                    '    _ExcelDoc.Write("</Table>")
                    '    _ExcelDoc.Write(" </Worksheet>")
                    '    _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                    '    _ExcelDoc.Write("<Table>")
                    '    _ExcelDoc.Write("<Row>")
                    '    For xi As Integer = 0 To source.Tables(2).Columns.Count - 1
                    '        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write(source.Tables(2).Columns(xi).ColumnName)
                    '        _ExcelDoc.Write("</Data></Cell>")
                    '    Next
                    '    _ExcelDoc.Write("</Row>")
                    'End If
                    _ExcelDoc.Write("<Row>")
                    For y As Integer = 0 To source.Tables(2).Columns.Count - 1
                        Dim XMLstring As String = x(y).ToString()
                        XMLstring = XMLstring.Trim()
                        If IsNumeric(XMLstring) Then
                            _ExcelDoc.Write("<Cell ss:StyleID=""BC"">" + "<Data ss:Type=""Number"">")
                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""BC"">" + "<Data ss:Type=""String"">")
                        End If
                        _ExcelDoc.Write(XMLstring)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    _ExcelDoc.Write("</Row>")
                    Dim XMLCusto As String
                    XMLCusto = ""
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                    XMLCusto = "Abbreviations"
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(XMLCusto)
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                    XMLCusto = ""
                    XMLCusto = "NDP = NO  OF DAYS  PRESENT"
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(XMLCusto)
                    _ExcelDoc.Write("</Data></Cell>")
                    XMLCusto = ""
                    '''''''''''''''''''''''''''''''''
                    For P As Integer = 1 To 3
                        If P = 3 Then
                            _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                            XMLCusto = "CL = CASUAL LEAVE"
                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""String"">")
                            XMLCusto = ""
                        End If

                        _ExcelDoc.Write(XMLCusto)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    '''''''''''''''''''''''''''''''''
                    XMLCusto = ""
                    _ExcelDoc.Write("</Row>")
                    XMLCusto = "WO = WEEKLY OFF"
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(XMLCusto)
                    _ExcelDoc.Write("</Data></Cell>")
                    '''''''''''''''''''''''''''''''''
                    For P As Integer = 1 To 3
                        If P = 3 Then
                            _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                            XMLCusto = "LOP = LOSS OF PAY"
                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""String"">")
                            XMLCusto = ""
                        End If

                        _ExcelDoc.Write(XMLCusto)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    '''''''''''''''''''''''''''''''''
                    _ExcelDoc.Write("</Row>")
                    XMLCusto = ""
                    XMLCusto = "PH = PUBLIC HOLIDAY"
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(XMLCusto)
                    _ExcelDoc.Write("</Data></Cell>")
                    XMLCusto = ""
                    '''''''''''''''''''''''''''''''''
                    For P As Integer = 1 To 3
                        If P = 3 Then
                            _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                            XMLCusto = "HRA = HOUSE RENT ALLOWANCE"
                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""String"">")
                            XMLCusto = ""
                        End If

                        _ExcelDoc.Write(XMLCusto)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    '''''''''''''''''''''''''''''''''
                    _ExcelDoc.Write("</Row>")
                    XMLCusto = ""
                    XMLCusto = "PL =  PRIVILEGE LEAVE"
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(XMLCusto)
                    _ExcelDoc.Write("</Data></Cell>")
                    '''''''''''''''''''''''''''''''''
                    For P As Integer = 1 To 3
                        If P = 3 Then
                            _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                            XMLCusto = "LTA = LEAVE TRAVEL  ALLOWNACE"
                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""String"">")
                            XMLCusto = ""
                        End If

                        _ExcelDoc.Write(XMLCusto)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    '''''''''''''''''''''''''''''''''
                    _ExcelDoc.Write("</Row>")
                    XMLCusto = ""
                    XMLCusto = "SL = SICK LEAVE"
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(XMLCusto)
                    _ExcelDoc.Write("</Data></Cell>")
                    '''''''''''''''''''''''''''''''''
                    For P As Integer = 1 To 3
                        If P = 3 Then
                            _ExcelDoc.Write("<Cell ss:MergeAcross=""3"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                            XMLCusto = "MLWF = MAHARASHTRA LABOUR WELFARE FUND"
                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""String"">")
                            XMLCusto = ""
                        End If

                        _ExcelDoc.Write(XMLCusto)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    '''''''''''''''''''''''''''''''''
                    _ExcelDoc.Write("</Row>")
                    XMLCusto = ""
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                    XMLCusto = "PREPARED BY________"
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(XMLCusto)
                    _ExcelDoc.Write("</Data></Cell>")
                    '''''''''''''''''''''''''''''''''
                    For P As Integer = 1 To 6
                        If P = 3 Then
                            _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                            XMLCusto = "CHECKED BY________"
                        ElseIf P = 6 Then
                            _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                            XMLCusto = "AUTHORISED BY________"
                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""String"">")
                            XMLCusto = ""
                        End If

                        _ExcelDoc.Write(XMLCusto)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    '''''''''''''''''''''''''''''''''
                    _ExcelDoc.Write("</Row>")

                Next

                _ExcelDoc.Write("</Table>")
                _ExcelDoc.Write(" </Worksheet>")
                _ExcelDoc.Write(end_ExcelXML)
            Catch ex As Exception
                'for catch the error
                _objcommonExp.PublishError("For gebnerate the excel file (fnExportTo_ExcelXML())", ex)
            End Try
            'Return _ExcelDoc
        End Sub
        Private Sub ExportToExcelFoWorker(ByVal source As DataSet, ByRef _ExcelDoc As StreamWriter)
            Dim _RowCount As Integer = 0
            'Dim _ExcelDoc As System.IO.StringWriter
            ' _ExcelDoc = New System.IO.StringWriter()
            Try
                Dim _ExcelXML As New StringBuilder()
                _ExcelDoc.Write("<?xml version=""1.0""?><?mso-application progid=""Excel.Sheet""?>" & Chr(13) & "" & Chr(10) & "<Workbook ")
                _ExcelDoc.Write("xmlns=""urn:schemas-microsoft-com:office:spreadsheet""" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write(" xmlns:o=""urn:schemas-microsoft-com:office:office""" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write(" xmlns:x=""urn:schemas-microsoft-com:office:excel""" & Chr(13) & "" & Chr(10))
                _ExcelDoc.Write(" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">" & Chr(13) & "" & Chr(10) & " <Styles>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""Default"" ss:Name=""Normal"">" & Chr(13) & "" & Chr(10) & " ")
                '_ExcelXML.Append("<Alignment ss:Horizontal=\"Right\" ss:Vertical=\"Bottom\"/>\r\n <Borders/>"); 
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>" & Chr(13) & "" & Chr(10) & " <Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /></Borders>")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Font/>" & Chr(13) & "" & Chr(10) & " <Interior/>" & Chr(13) & "" & Chr(10) & " <NumberFormat/>")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Protection/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")

                'Cumstomise format 
                _ExcelDoc.Write("<Style ss:ID=""CD"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Center"" ss:Vertical=""Bottom"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelDoc.Write("ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Bold=""1"" />" & Chr(13) & "" & Chr(10) & " <Interior ss:Color=""#CDAF95"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & " ")


                '_ExcelXML.Append("\r\n <Style ss:ID=\"StyleBorder\"><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/> <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>"); 
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Style ss:ID=""BC"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelDoc.Write("x:Family=""Swiss"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""SL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat")
                _ExcelDoc.Write(" ss:Format=""@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _ExcelDoc.Write("ss:ID=""DC"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""0.0000""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""IT"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""0""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _ExcelDoc.Write("ss:ID=""DL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""mm/dd/yyyy;@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""s2"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Top"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:Size=""9"" ss:Color=""#808080"" ss:Underline=""Single""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#FFFFFF"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("</Styles>" & Chr(13) & "" & Chr(10) & " ")

                Dim start_ExcelXML As String = _ExcelDoc.ToString()
                Const end_ExcelXML As String = "</Workbook>"
                Dim sheetCount As Integer = 1

                '_ExcelDoc.Write(start_ExcelXML)
                _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                _ExcelDoc.Write("<Table><Column ss:Width=""40""/><Column ss:Width=""60""/><Column ss:Width=""120""/><Column ss:Width=""80""/><Column ss:Width=""120""/><Column ss:Width=""80""/><Column ss:Width=""80""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""50""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""60""/><Column ss:Width=""50""/><Column ss:Width=""80""/><Column ss:Width=""80""/><Column ss:Width=""100""/><Column ss:Width=""65""/><Column ss:Width=""65""/><Column ss:Width=""65""/>")
                _ExcelDoc.Write("<Row>")

                For x As Integer = 0 To 6
                    If x = 6 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(source.Tables(0).Rows(0).Item("Comp_Name").ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")
                _ExcelDoc.Write("<Row>")

                For x As Integer = 0 To 5
                    If x = 5 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""4"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("WAGES FOR THE MONTH OF--")
                        _ExcelDoc.Write(source.Tables(0).Rows(0).Item("Mon").ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")
                _ExcelDoc.Write("<Row>")

                For x As Integer = 0 To 6
                    If x = 6 Then
                        _ExcelDoc.Write("<Cell ss:MergeAcross=""4"" ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("LOCATION--")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write(source.Tables(0).Rows(0).Item("LOCATION").ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To 1 - 1
                    If x < 1 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next
                _ExcelDoc.Write("</Row>")

                '''''''''''''''''''''''''''

                '''''''''''''''''''''''''''
                _ExcelDoc.Write("<Row>")
                For x As Integer = 0 To (3 + source.Tables(3).Rows(0).Item("AddCount").ToString - 4)

                    If source.Tables(1).Columns(x).ColumnName = "SrNo" Then
                        _ExcelDoc.Write("<Cell   ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Sr.No")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "EmpCode" Then
                        _ExcelDoc.Write("<Cell  ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Emp Code")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "NAME" Then
                        _ExcelDoc.Write("<Cell  ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Name  of the Employee")
                    ElseIf source.Tables(1).Columns(x).ColumnName = "Designation" Then
                        _ExcelDoc.Write("<Cell   ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("Designation")

                    Else

                        _ExcelDoc.Write("<Cell  ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(source.Tables(1).Columns(x).ColumnName)

                    End If
                    _ExcelDoc.Write("</Data></Cell>")
                Next
                Dim Px As Integer = (source.Tables(1).Columns.Count - 1) - 3
                Dim Wx, Zx As String
                Wx = source.Tables(3).Rows(0).Item("AddCount").ToString
                Zx = source.Tables(4).Rows(0).Item("DedCount").ToString
                ' For p As Integer = (source.Tables(3).Rows(0).Item("AddCount").ToString) To source.Tables(1).Columns.Count - 1


                _ExcelDoc.Write("<Cell ss:MergeAcross=""1"" ss:StyleID=""CD""><Data ss:Type=""String"">")
                _ExcelDoc.Write("Attendance Calculation")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:MergeAcross=""1"" ss:StyleID=""CD""><Data ss:Type=""String"">")
                _ExcelDoc.Write("")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                _ExcelDoc.Write("GROSS AMOUNT")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:MergeAcross=""" & (Zx - 1) & """ ss:StyleID=""CD""><Data ss:Type=""String"">")
                _ExcelDoc.Write("Deductions")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                _ExcelDoc.Write("Total Deductions")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                _ExcelDoc.Write("Net Amount")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                _ExcelDoc.Write("Signature")
                _ExcelDoc.Write("</Data></Cell>")

                ' Next
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                For T As Integer = 0 To source.Tables(1).Columns.Count - 1
                    If T >= (source.Tables(3).Rows(0).Item("AddCount").ToString) And T <= Px + 1 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(source.Tables(1).Columns(T).ColumnName)
                    Else
                        _ExcelDoc.Write("<Cell  ss:StyleID=""CD""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("")

                    End If
                    _ExcelDoc.Write("</Data></Cell>")

                Next
                _ExcelDoc.Write("</Row>")



                '''''''''''''''''''''''''''''''''''
                For Each x As DataRow In source.Tables(1).Rows
                    _RowCount += 1
                    'if the number of rows is > 63000 create a new page to continue output 
                    'If _RowCount = 63000 Then
                    '    Dim test As String = ""
                    '    test = "god"
                    'End If
                    'If _RowCount = 63000 Then
                    '    _RowCount = 0
                    '    sheetCount += 1
                    '    _ExcelDoc.Write("</Table>")
                    '    _ExcelDoc.Write(" </Worksheet>")
                    '    _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                    '    _ExcelDoc.Write("<Table>")
                    '    _ExcelDoc.Write("<Row>")
                    '    For xi As Integer = 0 To source.Tables(1).Columns.Count - 1
                    '        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write(source.Tables(1).Columns(xi).ColumnName)
                    '        _ExcelDoc.Write("</Data></Cell>")
                    '    Next
                    '    _ExcelDoc.Write("</Row>")
                    'End If
                    _ExcelDoc.Write("<Row>")
                    For y As Integer = 0 To source.Tables(1).Columns.Count - 1
                        Dim XMLstring As String = x(y).ToString()
                        XMLstring = XMLstring.Trim()

                        If IsNumeric(XMLstring) Then
                            If y = 6 Then
                                _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""String"">")
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""Number"">")
                            End If

                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""String"">")
                        End If
                        _ExcelDoc.Write(XMLstring)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    _ExcelDoc.Write("</Row>")
                Next
                '---By GRand
                For Each x As DataRow In source.Tables(2).Rows
                    _RowCount += 1
                    'if the number of rows is > 63000 create a new page to continue output 
                    'If _RowCount = 63000 Then
                    '    Dim test As String = ""
                    '    test = "god"
                    'End If
                    'If _RowCount = 63000 Then
                    '    _RowCount = 0
                    '    sheetCount += 1
                    '    _ExcelDoc.Write("</Table>")
                    '    _ExcelDoc.Write(" </Worksheet>")
                    '    _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                    '    _ExcelDoc.Write("<Table>")
                    '    _ExcelDoc.Write("<Row>")
                    '    For xi As Integer = 0 To source.Tables(2).Columns.Count - 1
                    '        _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write(source.Tables(2).Columns(xi).ColumnName)
                    '        _ExcelDoc.Write("</Data></Cell>")
                    '    Next
                    '    _ExcelDoc.Write("</Row>")
                    'End If
                    _ExcelDoc.Write("<Row>")
                    For y As Integer = 0 To source.Tables(2).Columns.Count - 1
                        Dim XMLstring As String = x(y).ToString()
                        XMLstring = XMLstring.Trim()
                        If IsNumeric(XMLstring) Then
                            _ExcelDoc.Write("<Cell ss:StyleID=""BC"">" + "<Data ss:Type=""Number"">")
                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""BC"">" + "<Data ss:Type=""String"">")
                        End If
                        _ExcelDoc.Write(XMLstring)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    _ExcelDoc.Write("</Row>")
                    Dim XMLCusto As String
                    XMLCusto = ""
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:StyleID=""BC""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                    XMLCusto = "Abbreviations"
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(XMLCusto)
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                    XMLCusto = ""
                    XMLCusto = "NDP = NO  OF DAYS  PRESENT"
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(XMLCusto)
                    _ExcelDoc.Write("</Data></Cell>")
                    XMLCusto = ""
                    _ExcelDoc.Write("</Row>")
                    XMLCusto = "OT =OVERTIME"
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(XMLCusto)
                    _ExcelDoc.Write("</Data></Cell>")
                    _ExcelDoc.Write("</Row>")
                    XMLCusto = ""
                    XMLCusto = "PREPARED BY________"
                    _ExcelDoc.Write("<Row>")
                    _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(XMLCusto)
                    _ExcelDoc.Write("</Data></Cell>")
                    '''''''''''''''''''''''''''''''''
                    For P As Integer = 1 To 6
                        If P = 3 Then
                            _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                            XMLCusto = "CHECKED BY________"
                        ElseIf P = 6 Then
                            _ExcelDoc.Write("<Cell ss:MergeAcross=""2"" ss:StyleID=""SL""><Data ss:Type=""String"">")
                            XMLCusto = "AUTHORISED BY________"
                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""SL"">" + "<Data ss:Type=""String"">")
                            XMLCusto = ""
                        End If

                        _ExcelDoc.Write(XMLCusto)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next
                    '''''''''''''''''''''''''''''''''
                    _ExcelDoc.Write("</Row>")

                Next

                _ExcelDoc.Write("</Table>")
                _ExcelDoc.Write(" </Worksheet>")
                _ExcelDoc.Write(end_ExcelXML)
            Catch ex As Exception
                'for catch the error
                _objcommonExp.PublishError("For generate the excel file (fnExportTo_ExcelXML())", ex)
            End Try
            ' Return _ExcelDoc
        End Sub
        'Added by praveen on 17 Aug 2011.
        Protected Sub DdlreportType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlreportType.SelectedIndexChanged
            reportvisible()
        End Sub
        'Added by praveen on 17 Aug 2011.
        Private Sub reportvisible()
            Try
                lblMailMsgWOPWD.Text = ""
                lblMailMsg.Text = ""
                lblProcessBarMsg.Text = ""
                LnkPDF.Style.Value = "display:none;"
                LnkPDFWOPWD.Style.Value = "display:none;"
                download_pdf1.Style.Value = "display:none;"
                download_pdf2.Style.Value = "display:none;"
                chkMerge.Checked = False
                CheckProcessLocked()
                If (DdlreportType.SelectedValue.ToString.ToUpper = "T" OrElse DdlreportType.SelectedValue.ToString.ToUpper = "R" _
                    OrElse DdlreportType.SelectedValue.ToString.ToUpper = "S" OrElse DdlreportType.SelectedValue.ToString.ToUpper = "57") _
                    AndAlso ddlRepIn.SelectedValue.ToString.ToUpper = "P" AndAlso rbtnmail.Checked Then
                    btnPublishedPDF.Visible = True
                    btnPublishedPDF.ToolTip = "Download Already Published " & DdlreportType.SelectedItem.Text
                Else
                    btnPublishedPDF.Visible = False
                    btnPublishedPDF.ToolTip = "Download Already Published Pay Slips"
                End If
                If DdlreportType.SelectedValue.ToString = "49" Then
                    trhelp1.Style("display") = ""

                Else
                    trhelp1.Style("display") = "none"
                End If
                If Convert.ToString(DdlreportType.SelectedValue).Equals("67") Then
                    trlingpdf.Style("display") = ""
                    tblpwd.Style("display") = "none"
                Else
                    trlingpdf.Style("display") = "none"
                End If
                trrepformat.Style.Value = "display:none"
                tableshow.Style.Value = "display:none"
                tblrepin.Style.Value = "display:"
                'Added by Rohtas Singh on 06 Dec 2017
                trRptformat.Style.Value = "display:none"
                If DdlreportType.SelectedValue.ToString = "62" Then
                    troffcycledt.Style.Value = "display:"
                    populateFromdate("P")
                Else
                    troffcycledt.Style.Value = "display:none"
                End If
                'Added by geeta on 11 Sep 2012 display reporting manager Email id checkbox when Salary slip with leave balance report selected
                trRepEmail.Style.Value = "display:none"
                chkemailrepmanager.Checked = False
                chkmailformat.Checked = False
                If DdlreportType.SelectedValue.ToUpper.ToString = "I" Or DdlreportType.SelectedValue.ToUpper.Equals("67") Then
                    tblsp.Style.Value = "display:none"
                    tblSh.Style.Value = "display:none"
                    tblpwd.Style.Value = "display:none"
                    TrNoSearch.Style.Value = "display:none"
                    tremail.Style.Value = "display:none"
                Else
                    tblsp.Style.Value = "display:"
                    tblSh.Style.Value = "display:"
                    tremail.Style.Value = "display:"

                    If DdlreportType.SelectedValue.ToUpper.ToString = "SL" Or DdlreportType.SelectedValue.ToUpper.ToString = "S" Or DdlreportType.SelectedValue.ToUpper.ToString = "T" Or DdlreportType.SelectedValue.ToUpper.ToString = "R" Or DdlreportType.SelectedValue.ToUpper.ToString = "52" Or DdlreportType.SelectedValue.ToUpper.ToString = "SI" Or DdlreportType.SelectedValue.ToUpper.ToString = "SH" Or DdlreportType.SelectedValue.ToUpper.ToString = "TL" Or DdlreportType.SelectedValue.ToUpper.ToString = "RS" Or DdlreportType.SelectedValue.ToUpper.Equals("62") Then
                        tblpwd.Style.Value = "display:"
                        TrNoSearch.Style.Value = "display:"
                        'Added by geeta on 11 Sep 2012 display reporting manager Email id checkbox when Salary slip with leave balance report selected
                        If DdlreportType.SelectedValue.ToUpper.ToString = "SL" Then
                            trRepEmail.Style.Value = "display:"
                        End If
                    Else
                        tblpwd.Style.Value = "display:none"
                        TrNoSearch.Style.Value = "display:none"
                    End If

                    If DdlreportType.SelectedValue.ToUpper.ToString = "RN" Then
                        BtnPreviewdivActive.Style.Value = "display:"
                    Else
                        BtnPreviewdivActive.Style.Value = "display:none"
                    End If
                End If
                'Geeta
                Dim dt As New DataTable
                dt = _ObjData.GetDataTableProc("paysp_bind_DDLPaySlipType")
                HidRepId.Value = dt.Rows(DdlreportType.SelectedIndex).Item("Reptype").ToString
                ddlRepIn.Items.Clear()
                If DdlreportType.SelectedValue = "I" Then
                    ddlRepIn.Items.Add(New ListItem("HTML", "H"))
                ElseIf DdlreportType.SelectedValue = "SI" Then
                    ddlRepIn.Items.Add(New ListItem("PDF", "P"))
                    ddlRepIn.Items.Add(New ListItem("PaySlip Link", "L"))
                    'added by Geeta : Marathi payslip("60")
                ElseIf DdlreportType.SelectedValue = "SH" Or DdlreportType.SelectedValue = "TL" Or DdlreportType.SelectedValue = "RS" Or DdlreportType.SelectedValue = "RN" _
                    Or DdlreportType.SelectedValue = "43" Or DdlreportType.SelectedValue = "49" Or DdlreportType.SelectedValue = "51" Or DdlreportType.SelectedValue = "55" _
                    Or Convert.ToString(DdlreportType.SelectedValue).Equals("56") Or Convert.ToString(DdlreportType.SelectedValue).Equals("57") _
                    Or Convert.ToString(DdlreportType.SelectedValue).Equals("58") Or Convert.ToString(DdlreportType.SelectedValue).Equals("59") _
                    Or Convert.ToString(DdlreportType.SelectedValue).Equals("60") Or Convert.ToString(DdlreportType.SelectedValue).Equals("63") _
                     Or Convert.ToString(DdlreportType.SelectedValue).Equals("64") Or Convert.ToString(DdlreportType.SelectedValue).Equals("65") _
                      Or Convert.ToString(DdlreportType.SelectedValue).Equals("66") Or Convert.ToString(DdlreportType.SelectedValue).Equals("68") Then
                    ddlRepIn.Items.Add(New ListItem("PDF", "P"))
                    'PTC
                ElseIf DdlreportType.SelectedValue = "YTD" Or DdlreportType.SelectedValue = "53" Then
                    ddlRepIn.Items.Add(New ListItem("HTML", "H"))
                    ddlRepIn.Items.Add(New ListItem("PDF", "P"))
                ElseIf DdlreportType.SelectedValue.Equals("67") Then
                    ddlRepIn.Items.Add(New ListItem("HTML", "H"))
                    ddlRepIn.Items.Add(New ListItem("PDF", "P"))
                ElseIf DdlreportType.SelectedValue.Equals("74") Then
                    ddlRepIn.Items.Add(New ListItem("HTML", "H"))
                    ddlRepIn.Items.Add(New ListItem("PDF", "P"))
                Else
                    ddlRepIn.Items.Add(New ListItem("HTML", "H"))
                    ddlRepIn.Items.Add(New ListItem("PDF", "P"))
                    ddlRepIn.Items.Add(New ListItem("PaySlip Link", "L"))
                End If

                If ddlRepIn.SelectedValue = "P" Then
                    tblpwd.Style.Value = "display:"
                    TrNoSearch.Style.Value = "display:"
                Else
                    tblpwd.Style.Value = "display:none"
                    TrNoSearch.Style.Value = "display:none"
                End If
                If ddlRepIn.SelectedValue <> "P" Then
                    RblNoSearch.SelectedValue = "S"
                Else
                    RblNoSearch.SelectedValue = "P"
                End If
                ShowHideNoSearch()
                'Added by Rohtas Singh on 14 Feb 2018
                If DdlreportType.SelectedValue.ToString = "55" Or Convert.ToString(DdlreportType.SelectedValue).Equals("56") Or Convert.ToString(DdlreportType.SelectedValue).Equals("57") Then
                    trMergeMsg.Style("display") = ""
                Else
                    trMergeMsg.Style("display") = "none"
                End If
                If DdlreportType.SelectedValue.Equals("62") Or DdlreportType.SelectedValue.Equals("63") Then
                    tremail.Style("display") = "none"
                    tblsp.Style("display") = "none"
                    tblSh.Style("display") = "none"
                    trselall.Style("display") = "none"
                Else
                    tremail.Style("display") = ""
                    tblsp.Style("display") = ""
                    tblSh.Style("display") = ""
                    trselall.Style("display") = ""
                End If
                If DdlreportType.SelectedValue.ToUpper.Equals("S") Then
                    If ddlEmpPass.Items.Count = 6 Then
                        ddlEmpPass.Items.RemoveAt(5)
                    End If
                    If ddlEmpPass.Items.Count = 8 Then
                        ddlEmpPass.Items.RemoveAt(5)
                        ddlEmpPass.Items.Add(New ListItem("Pan No. & DOB(DDMM)", "8"))
                        ddlEmpPass.Items.Add(New ListItem("Emp. Code & DOB(DDMM)", "9"))
                    End If
                    If ddlEmpPass.Items.Count = 5 Then
                        ddlEmpPass.Items.Add(New ListItem("Pan No. & DOB(DDMMYY)", "6"))
                        ddlEmpPass.Items.Add(New ListItem("Emp. Code & DOB(DDMMYY)", "7"))
                        ddlEmpPass.Items.Add(New ListItem("Pan No. & DOB(DDMM)", "8"))
                        ddlEmpPass.Items.Add(New ListItem("Emp. Code & DOB(DDMM)", "9"))
                    End If
                ElseIf DdlreportType.SelectedValue.ToUpper.Equals("R") Then
                    If ddlEmpPass.Items.Count = 9 Then
                        ddlEmpPass.Items.RemoveAt(8)
                        ddlEmpPass.Items.RemoveAt(7)
                    End If
                    If ddlEmpPass.Items.Count = 7 Then
                        ddlEmpPass.Items.RemoveAt(6)
                        ddlEmpPass.Items.RemoveAt(5)
                        ddlEmpPass.Items.Add(New ListItem("First Name & DOB(DDMMYYYY)", "5"))
                        ddlEmpPass.Items.Add(New ListItem("Pan No. & DOB(DDMMYY)", "6"))
                        ddlEmpPass.Items.Add(New ListItem("Emp. Code & DOB(DDMMYY)", "7"))
                    End If
                    If ddlEmpPass.Items.Count = 6 Then
                        ddlEmpPass.Items.Add(New ListItem("Pan No. & DOB(DDMMYY)", "6"))
                        ddlEmpPass.Items.Add(New ListItem("Emp. Code & DOB(DDMMYY)", "7"))
                    End If
                    If ddlEmpPass.Items.Count = 5 Then
                        ddlEmpPass.Items.Add(New ListItem("First Name & DOB(DDMMYYYY)", "5"))
                        ddlEmpPass.Items.Add(New ListItem("Pan No. & DOB(DDMMYY)", "6"))
                        ddlEmpPass.Items.Add(New ListItem("Emp. Code & DOB(DDMMYY)", "7"))
                    End If
                Else
                    If ddlEmpPass.Items.Count = 9 Then
                        ddlEmpPass.Items.RemoveAt(8)
                        ddlEmpPass.Items.RemoveAt(7)
                        ddlEmpPass.Items.RemoveAt(6)
                        ddlEmpPass.Items.RemoveAt(5)
                    ElseIf ddlEmpPass.Items.Count = 8 Then
                        ddlEmpPass.Items.RemoveAt(7)
                        ddlEmpPass.Items.RemoveAt(6)
                    End If
                    If ddlEmpPass.Items.Count = 5 Then
                        ddlEmpPass.Items.Add(New ListItem("First Name & DOB(DDMMYYYY)", "5"))
                    End If
                End If
                'Added by Quadir on 14 OCT 2020
                If DdlreportType.SelectedValue.ToString.ToUpper = "R" And ddlRepIn.SelectedValue.ToString.ToUpper = "P" Then
                    TrSlipPubMode.Style.Value = "display:"
                Else
                    TrSlipPubMode.Style.Value = "display:none"
                End If
            Catch ex As Exception
                _objcommonExp.PublishError("DdlreportType_SelectedIndexChanged()", ex)
            End Try
        End Sub
        Private Sub BtnSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSend.Click
            lblMsgSlip.Text = ""
            lblMailMsg.Text = ""
            lblProcessBarMsg.Text = ""
            LnkPDF.Style.Value = "display:none"
            HidEmailCCBCC.Value = ""
            'Excel Process locking validation checking
            'CheckExcelProcessbarAlreadyProcessing()
            'If (lblProcessStatusExcel.Text <> "") Then
            '    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openproces9F821", "UnLoadPaySlipProgress();", True)
            '    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = lblProcessStatusExcel.Text})
            '    _objCommon.ShowMessage(_msg)
            '    Exit Sub
            'End If
            CheckProcessLocked()
            If (hdnAlreadyRunRptName.Value.Trim().Length > 1) Then
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocessF8221", "UnLoadPaySlipProgress();", True)
                _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = hdnAlreadyRunRptName.Value})
                _objCommon.ShowMessage(_msg)
                Exit Sub
            End If
            If ddlRepIn.SelectedValue = "L" Then
                If chkmailformat.Checked = True Then
                    _SaveMailBody()
                End If
                _LinkSendMail()
            Else
                If UCase(ddlRepIn.SelectedValue.ToString) = "P" Then
                    'Condition added by Rohtas Singh on 27 Dec 2017 for check user wants to save mail body or not
                    If chkmailformat.Checked = True And hidSaveMail.Value.ToUpper = "Y" Then
                        _SaveMailBody()
                    End If
                    SendReportPDF("")

                    ' Server.Execute("Reports/PreSalSlip.aspx?id=Pankaj")
                Else
                    If chkmailformat.Checked = True Then
                        _SaveMailBody()
                    End If
                    If DdlreportType.SelectedValue.ToString = "53" Then
                        _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "M", .MessageString = "Please select PDF format to send email to employee.HTML format is not available for this report."})
                        _objCommon.ShowMessage(_msg)
                    Else
                        SendReport()
                    End If

                End If
            End If
        End Sub
        Private Sub _SaveMailBody()
            Try
                Dim _arrparam(3) As SqlClient.SqlParameter
                _arrparam(0) = New SqlClient.SqlParameter("@Fk_report_Id", DdlreportType.SelectedValue.ToString)
                _arrparam(1) = New SqlClient.SqlParameter("@Header", txtheader.Text.ToString)
                _arrparam(2) = New SqlClient.SqlParameter("@Footer", Textfooter.Text.ToString)
                _arrparam(3) = New SqlClient.SqlParameter("@MailBody", _objCommon.nNz(EasyWebMAilBody.HTMLValue.ToString))

                _ObjData.ExecuteStoredProc("Paysp_mstmailbody_ForIns", _arrparam)
            Catch ex As Exception
                _objcommonExp.PublishError("error in (_SaveMailBody())", ex)
            End Try
        End Sub

        Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
            Dim EmpCodeSearch As String = "", StrPath As String
            HidEmailCCBCC.Value = ""
            HidEmpPdf.Value = ""
            lblProcessBarMsg.Text = ""
            lblMsgSlip.Text = ""
            lblMailMsg.Text = ""
            LnkPDF.Style.Value = "display:none"
            LnkPDFWOPWD.Style.Value = "display:None"
            download_pdf2.Style.Value = "display:None"
            lblMailMsgWOPWD.Text = ""
            'Excel Process locking validation checking
            'CheckExcelProcessbarAlreadyProcessing()
            'If (lblProcessStatusExcel.Text <> "") Then
            '    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocessF821", "UnLoadPaySlipProgress();", True)
            '    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = lblProcessStatusExcel.Text})
            '    _objCommon.ShowMessage(_msg)
            '    Exit Sub
            'End If
            CheckProcessLocked()
            If (hdnAlreadyRunRptName.Value.Trim().Length > 1) Then
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocessF8241", "UnLoadPaySlipProgress();", True)
                _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = hdnAlreadyRunRptName.Value})
                _objCommon.ShowMessage(_msg)
                Exit Sub
            End If
            Dim _mm As String = "", _yyyy As String = ""
            If ddlMonthYear.SelectedItem IsNot Nothing Then
                _mm = CType(ddlMonthYear.SelectedValue.ToString, Integer)
                _yyyy = Right(ddlMonthYear.SelectedItem.Text.ToString, 4)
            End If
            Dim gcs_service As Integer = 0
            Dim process_type = ""

            Try
                StrPath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\PaySlipAcessCheck\"
                If Not Directory.Exists(StrPath) Then
                    LogMessage.log.Info("Directory does not exist")
                    LogMessage.log.Debug("Directory Creation initialization")
                    Directory.CreateDirectory(StrPath)
                    LogMessage.log.Debug("Directory Created")
                Else
                    LogMessage.log.Info("Directory exist")
                    LogMessage.log.Debug("Directory deletion initialization")
                    Directory.Delete(StrPath)
                    LogMessage.log.Debug("Delete Directory")
                End If
                LogMessage.log.Debug("Directory exist done")
            Catch ex As Exception
                LogMessage.log.Error("PDFFiles directory Access issue", ex)
                Dim Allexemption As String = ""
                Allexemption = Allexemption & "Message*" & ex.Message.ToString
                Allexemption = Allexemption & "InnerException*" & ex.InnerException.ToString
                Allexemption = Allexemption & "GetBaseException*" & ex.GetBaseException.ToString
                Allexemption = Allexemption & "Source*" & ex.Source.ToString
                Allexemption = Allexemption & "TargetSite*" & ex.TargetSite.ToString
                _objcommonExp.PublishError("Error:Access:PDFFiles ***" & Allexemption, ex)
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocessF8221", "UnLoadPaySlipProgress();", True)
                Exit Sub
            End Try

            If UCase(ddlRepIn.SelectedValue.ToString) = "P" Then
                If DdlreportType.SelectedValue.ToString = "R" Then
                    Dim dst As New DataSet
                    'LoadPaySlipProgress
                    'ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocess800", "LoadPaySlipProgress('" + DdlreportType.SelectedItem.Text.Replace("'", "") + "');", True)
                    'Added by Quadir on 24 Nov 2020 for changing logic Payslip Publish Mode on Search
                    If RblNoSearch.SelectedValue = "S" Then

                        For counter = 0 To DgPayslip.Items.Count - 1
                            If CType(DgPayslip.Items(counter).FindControl("chkEmpHold"), CheckBox).Checked = True Then
                                EmpCodeSearch = EmpCodeSearch + DgPayslip.Items(counter).Cells(1).Text.ToString + ","
                            End If
                        Next
                        dst = ReturnDsSearch("", "PDF", EmpCodeSearch)
                    Else
                        dst = ReturnDsSearch("", "PDF")
                    End If

                    If dst.Tables(0).Rows.Count > 0 Then
                        Dim _array() As String, _AppPath As String = HttpRuntime.AppDomainAppVirtualPath.ToString, Path As String, cnt As Integer, dtp As DataTable, row_count As Integer
                        row_count = dst.Tables(0).Rows.Count
                        _array = Split(_AppPath, "/")
                        _AppPath = _array(_array.Length - 1)
                        HidAppPath.Value = _AppPath
                        Path = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Left(MonthName(CType(ddlMonthYear.SelectedValue.ToString, Integer)), 3) & Right(ddlMonthYear.SelectedItem.Text.ToString, 4) & "\TaxPaySlip\"
                        If gcs_service_obj.IsPaySlipAllowedByMonthYear(Session("compCode"), "TAXSLIP", _mm, _yyyy) Then
                            gcs_service = 1
                        End If
                        If gcs_service = 1 Then
                            Session("pdf_file_location") = Path.ToString
                        End If

                        If Not Directory.Exists(Path) Then
                            Directory.CreateDirectory(Path)
                        End If
                        HidPath.Value = Replace(Replace(Path, "\", "~").ToString, "/", "~").ToString
                        SendReportPDF("S", "PDF", EmpCodeSearch)
                    Else
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocessF831", "UnLoadPaySlipProgress();", True)
                        If dst.Tables(2).Rows.Count > 0 Then
                            If dst.Tables(2).Rows(0).Item("AllEmp").ToString.ToUpper = "N" Then
                                IsProcessBarStated = False
                                _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "Employee(s) salary not processed as per selected month-year."})
                                _objCommon.ShowMessage(_msg)
                                LnkPDF.Style.Value = "display:None"
                                lblMailMsg.Text = ""
                                Exit Sub
                            Else
                                If rbtSlipPubMode.SelectedValue.ToString.ToUpper <> "O" Then

                                    If dst.Tables(2).Rows(0).Item("EmpSearch").ToString.ToUpper = "N" Then
                                        _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "M", .MessageString = "Already PDF files generated as per selected criteria on file server.<br>"})
                                        _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "M", .MessageString = "Please click on below PDF icon to open the already genetated PDF files.<br>"})
                                        _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "Record does not exists as per selected criteria!"})
                                    Else
                                        _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "M", .MessageString = "Already PDF files generated as per selected criteria on file server.<br>"})
                                        _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "M", .MessageString = "Please click on below PDF icon to open the already genetated PDF files.<br>"})
                                    End If
                                Else
                                    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "Record does not exists as per selected criteria!"})
                                    _objCommon.ShowMessage(_msg)
                                    LnkPDF.Style.Value = "display:None"
                                    lblMailMsg.Text = ""
                                    Exit Sub
                                End If

                                _objCommon.ShowMessage(_msg)
                            End If
                        Else
                            _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "M", .MessageString = "Already PDF files generated as per selected criteria.<br>"})
                            _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "M", .MessageString = "Please click on below PDF icon to open the already genetated PDF files."})
                            _objCommon.ShowMessage(_msg)
                        End If

                    End If
                ElseIf DdlreportType.SelectedValue.ToString = "S" Then
                    Dim dst As New DataSet
                    'ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocessS800", "LoadPaySlipProgress('" + DdlreportType.SelectedItem.Text.Replace("'", "") + "');", True)
                    If RblNoSearch.SelectedValue = "S" Then
                        For counter = 0 To DgPayslip.Items.Count - 1
                            If CType(DgPayslip.Items(counter).FindControl("chkEmpHold"), CheckBox).Checked = True Then
                                EmpCodeSearch = EmpCodeSearch + DgPayslip.Items(counter).Cells(1).Text.ToString + ","
                            End If
                        Next
                        dst = ReturnDsSearch("S", "SPDF", EmpCodeSearch)
                    Else
                        dst = ReturnDsSearch("", "SPDF")
                    End If

                    If dst.Tables(0).Rows.Count > 0 Then
                        Dim _array() As String, _AppPath As String = HttpRuntime.AppDomainAppVirtualPath.ToString, Path As String, cnt As Integer
                        cnt = dst.Tables(0).Rows.Count
                        _array = Split(_AppPath, "/")
                        _AppPath = _array(_array.Length - 1)
                        HidAppPath.Value = _AppPath
                        Path = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Left(MonthName(CType(ddlMonthYear.SelectedValue.ToString, Integer)), 3) & Right(ddlMonthYear.SelectedItem.Text.ToString, 4) & "\LeaveWoPaySlip\"

                        If gcs_service_obj.IsPaySlipAllowedByMonthYear(Session("compCode"), "SLIPWOLVE", _mm, _yyyy) Then
                            gcs_service = 1
                        End If
                        If gcs_service = 1 Then
                            Session("pdf_file_location") = Path.ToString
                        End If
                        Try
                            LogMessage.log.Debug("Directory exist")
                            If Not Directory.Exists(Path) Then
                                Directory.CreateDirectory(Path)
                            End If
                            LogMessage.log.Debug("Directory exist done")
                        Catch ex As Exception
                            LogMessage.log.Error("PDFFiles directory issue", ex)
                        End Try

                        HidPath.Value = Replace(Replace(Path, "\", "~").ToString, "/", "~").ToString
                        SendReportPDF("S", "SPDF", EmpCodeSearch)
                    Else
                        SendReportPDF("S")
                    End If
                ElseIf DdlreportType.SelectedValue.ToString.Equals("57") Then
                    Dim dst As New DataSet
                    'ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocessF800", "LoadPaySlipProgress('" + DdlreportType.SelectedItem.Text.Replace("'", "") + "');", True)
                    If RblNoSearch.SelectedValue = "S" Then
                        For counter = 0 To DgPayslip.Items.Count - 1
                            If CType(DgPayslip.Items(counter).FindControl("chkEmpHold"), CheckBox).Checked = True Then
                                EmpCodeSearch = EmpCodeSearch + DgPayslip.Items(counter).Cells(1).Text.ToString + ","
                            End If
                        Next
                        dst = ReturnDsSearch("S", "SPDF", EmpCodeSearch)
                    Else
                        dst = ReturnDsSearch("", "SPDF")
                    End If

                    If dst.Tables(0).Rows.Count > 0 Then
                        Dim _array() As String, _AppPath As String = HttpRuntime.AppDomainAppVirtualPath.ToString, Path As String, cnt As Integer
                        cnt = dst.Tables(0).Rows.Count
                        _array = Split(_AppPath, "/")
                        _AppPath = _array(_array.Length - 1)
                        HidAppPath.Value = _AppPath
                        Path = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Left(MonthName(CType(ddlMonthYear.SelectedValue.ToString, Integer)), 3) & Right(ddlMonthYear.SelectedItem.Text.ToString, 4) & "\YTDTaxComputationSheet\"
                        If gcs_service_obj.IsPaySlipAllowedByMonthYear(Session("compCode"), "FORCAST", _mm, _yyyy) Then
                            gcs_service = 1
                        End If
                        If gcs_service = 1 Then
                            Session("pdf_file_location") = Path.ToString
                        End If
                        Try
                            LogMessage.log.Debug("Directory exist")
                            If Not Directory.Exists(Path) Then
                                Directory.CreateDirectory(Path)
                            End If
                            LogMessage.log.Debug("Directory exist done")
                        Catch ex As Exception
                            LogMessage.log.Error("PDFFiles directory issue", ex)
                        End Try

                        HidPath.Value = Replace(Replace(Path, "\", "~").ToString, "/", "~").ToString
                        SendReportPDF("S", "SPDF", EmpCodeSearch)
                    Else
                        SendReportPDF("S")
                    End If
                Else
                    SendReportPDF("S")
                End If
                'Comment both line Rohtas(22Jan)
                If (IsProcessBarStated) Then
                    lblMailMsg.Text = "Click on PDF icon to open"
                    LnkPDF.Style.Value = "display:"
                End If
            End If
        End Sub
        ' following proc send the mail as attachement to those employee which is selected in the div
        Private Sub SendReport()
            Dim _TotRecord As Integer, _Counter As Integer, From_Mail As String = hid.Value.ToString _
            , To_mail As String, arrparam(3) As SqlClient.SqlParameter, MailNotSentCount As String = "" _
            , COSTCENTER As String = "", Reportid As String, var As String = "" _
            , monthvalue As String = "", YearVal As String = "", PM13 As String = "", EmpPassType As String = "" _
            , flag As String = "", counter As Integer = 0, Fk_Emp_Code As String = "", _Dtbl As New DataTable _
            , _Arr() As String = Nothing, _CCID As String = "", _BCCID As String = "", LOCATION As String = USearch.UCddllocation.ToString

            Reportid = HidRepId.Value.ToString
            monthvalue = Me.ddlMonthYear.SelectedValue
            YearVal = HidYear.Value.ToString
            PM13 = ddlshowsal.SelectedValue.ToString
            EmpPassType = ddlEmpPass.SelectedValue.ToString
            COSTCENTER = USearch.UCddlcostcenter.ToString
            'Loop for send mail
            If DdlreportType.SelectedValue = "T" Or DdlreportType.SelectedValue = "I" Then
                For _Counter = 0 To DgPayslip.Items.Count - 1
                    If CType(DgPayslip.Items(_Counter).FindControl("chkEmpHold"), CheckBox).Checked = True Then
                        To_mail = DgPayslip.Items(_Counter).Cells(5).Text.ToString
                        PK_emp_code = DgPayslip.Items(_Counter).Cells(1).Text.ToString
                        If To_mail.ToString.Replace("N/A", "") <> "" Then
                            'This code is used for send "TDS Estimation Slip"
                            If DdlreportType.SelectedValue = "T" Then
                                var = "H~" + monthvalue + "~" + YearVal + "~" + PK_emp_code + "~" + Reportid + "~" + EmpPassType + "~" + COSTCENTER + "~" + "" + "~" + flag + "~" + "N" + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~T"
                                _Arr = Split(var, "~")
                                Dim webrp As New WebReport.WebReport
                                _Dtbl = ClsNewTdsEstimationSlip.getSourceTable(, _Arr)

                                If txtccc.Text.ToString.Trim <> "" And txtBCC.Text.ToString.Trim <> "" Then
                                    lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(0).Item("MailID").ToString, "Tds-EstimationSlip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(0).Item("StrBuild").ToString, txtccc.Text.ToString.Trim, , txtBCC.Text.ToString.Trim)
                                ElseIf txtccc.Text.ToString.Trim <> "" Then
                                    lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(0).Item("MailID").ToString, "Tds-EstimationSlip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(0).Item("StrBuild").ToString, txtccc.Text.ToString.Trim)
                                ElseIf txtBCC.Text.ToString.Trim <> "" Then
                                    lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(0).Item("MailID").ToString, "Tds-EstimationSlip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(0).Item("StrBuild").ToString, From_Mail, , txtBCC.Text.ToString.Trim)
                                Else
                                    lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(0).Item("MailID").ToString, "Tds-EstimationSlip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(0).Item("StrBuild").ToString, From_Mail)
                                End If

                                'This code is used for send "Invest Declaration(Confirm Amount)"
                            ElseIf DdlreportType.SelectedValue = "I" Then
                                var = "M~" + PK_emp_code.ToString
                                _Arr = Split(var, "~")
                                _Dtbl = ClsConfInvest.getSourceTable(, _Arr)
                                If txtccc.Text.ToString.Trim <> "" And txtBCC.Text.ToString.Trim <> "" Then
                                    lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, To_mail, "Investment Approval of " & PK_emp_code.ToString, _Dtbl.Rows(0).Item("StrTemp").ToString, txtccc.Text.ToString.Trim, , txtBCC.Text.ToString.Trim, From_Mail.ToString)
                                ElseIf txtccc.Text.ToString.Trim <> "" Then
                                    lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, To_mail, "Investment Approval of " & PK_emp_code.ToString, _Dtbl.Rows(0).Item("StrTemp").ToString, txtccc.Text.ToString.Trim, , , From_Mail.ToString)
                                ElseIf txtBCC.Text.ToString.Trim <> "" Then
                                    lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, To_mail, "Investment Approval of " & PK_emp_code.ToString, _Dtbl.Rows(0).Item("StrTemp").ToString, From_Mail.ToString, , txtBCC.Text.ToString.Trim, From_Mail.ToString)
                                Else
                                    lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, To_mail, "Investment Approval of " & PK_emp_code.ToString, _Dtbl.Rows(0).Item("StrTemp").ToString, From_Mail.ToString, , , From_Mail.ToString)
                                End If
                            End If
                            If lblmsg.Text.ToUpper = "MAIL SENT SUCCESSFULLY !" Then
                                arrparam(0) = New SqlClient.SqlParameter("@pk_emp_code", PK_emp_code)
                                arrparam(1) = New SqlClient.SqlParameter("@MM", ddlMonthYear.SelectedValue.ToString)
                                arrparam(2) = New SqlClient.SqlParameter("@YY", HidYear.Value.ToString)
                                arrparam(3) = New SqlClient.SqlParameter("@ReportTtype", DdlreportType.SelectedValue.ToString)
                                'following sp save those employee in the database whom mail has been sent
                                _ObjData.GetDataSetProc("Sp_Ins_TrnDocMailStatus", arrparam)
                                'for count the record
                                _TotRecord = _TotRecord + 1
                            End If
                        Else
                            If MailNotSentCount.ToString = "" Then
                                MailNotSentCount = PK_emp_code
                            Else
                                MailNotSentCount += ", " & PK_emp_code
                            End If
                        End If
                        To_mail = ""
                    End If
                Next
            End If
            'PTC
            If DdlreportType.SelectedValue = "YTD" Or DdlreportType.SelectedValue = "R" Or DdlreportType.SelectedValue = "52" Or DdlreportType.SelectedValue = "SL" Or DdlreportType.SelectedValue = "S" Or DdlreportType.SelectedValue = "53" Then
                For counter = 0 To DgPayslip.Items.Count - 1
                    If CType(DgPayslip.Items(counter).FindControl("chkEmpHold"), CheckBox).Checked = True Then
                        Fk_Emp_Code = Fk_Emp_Code + DgPayslip.Items(counter).Cells(1).Text.ToString + ","
                    End If
                Next
                If Fk_Emp_Code.ToString <> "" Then
                    Fk_Emp_Code = Left(Fk_Emp_Code, Len(Fk_Emp_Code) - 1)
                End If
            End If

            'This code is used for send "Year To Date Salary Slip"
            If DdlreportType.SelectedValue = "YTD" Then
                var = "~~~~" + COSTCENTER + "~~~~" + Fk_Emp_Code + "~~" + monthvalue + "~" + YearVal + "~~" + PM13 + "~" + "A" + "~" + Reportid + "~H" + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~" + "" + "~" + EmpPassType.ToString
                _Arr = Split(var, "~")
                _Dtbl = _ClsYTD.getSourceTable(, _Arr)
                For counter = 0 To _Dtbl.Rows.Count - 1
                    If txtccc.Text.ToString.Trim <> "" And txtBCC.Text.ToString.Trim <> "" Then
                        lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Year To Date Salary Slip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, txtccc.Text.ToString.Trim, , txtBCC.Text.ToString.Trim)
                    ElseIf txtccc.Text.ToString.Trim <> "" Then
                        lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Year To Date Salary Slip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, txtccc.Text.ToString.Trim, , From_Mail)
                    ElseIf txtBCC.Text.ToString.Trim <> "" Then
                        lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Year To Date Salary Slip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, From_Mail, , txtBCC.Text.ToString.Trim)
                    Else
                        lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Year To Date Salary Slip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, From_Mail, , From_Mail)
                    End If
                Next

                'This code is used for send "For Salary Slip With Tax Details"
            ElseIf DdlreportType.SelectedValue = "R" Or DdlreportType.SelectedValue = "52" Then
                var = "H~~~~~" + COSTCENTER + "~~~~" + Fk_Emp_Code + "~~" + monthvalue + "~" + YearVal + "~" + "" + "~" + "" + "~" + "" + "~1~" + "true" + "~" + PM13 + "~" + "" + "~L" + "" + "~" + EmpPassType.ToString + "~" + "N" + "~" + "" + "~" + Reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString
                _Arr = Split(var, "~")
                If DdlreportType.SelectedValue = "52" Then
                    _Dtbl = ClsTaxDetails_Bangladesh.getSourceTable(, _Arr, , , , , _CCID, _BCCID)
                Else
                    _Dtbl = ClsTaxDetails.getSourceTable(, _Arr, , , , , _CCID, _BCCID)
                End If
                For counter = 0 To _Dtbl.Rows.Count - 1
                    If DdlreportType.SelectedValue = "52" Then
                        If txtccc.Text.ToString.Trim <> "" And txtBCC.Text.ToString.Trim <> "" Then
                            lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary Slip With Tax Details(Bangladesh) for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, txtccc.Text.ToString.Trim, , txtBCC.Text.ToString.Trim)
                        ElseIf txtccc.Text.ToString.Trim <> "" Then
                            lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary Slip With Tax Details(Bangladesh) for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, txtccc.Text.ToString.Trim, , _BCCID)
                        ElseIf txtBCC.Text.ToString.Trim <> "" Then
                            lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary Slip With Tax Details(Bangladesh) for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, _CCID, , txtBCC.Text.ToString.Trim)
                        Else
                            lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary Slip With Tax Details(Bangladesh) for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, _CCID, , _BCCID)
                        End If
                    Else
                        If txtccc.Text.ToString.Trim <> "" And txtBCC.Text.ToString.Trim <> "" Then
                            lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary slip with tax details for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, txtccc.Text.ToString.Trim, , txtBCC.Text.ToString.Trim)
                        ElseIf txtccc.Text.ToString.Trim <> "" Then
                            lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary slip with tax details for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, txtccc.Text.ToString.Trim, , _BCCID)
                        ElseIf txtBCC.Text.ToString.Trim <> "" Then
                            lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary slip with tax details for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, _CCID, , txtBCC.Text.ToString.Trim)
                        Else
                            lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary slip with tax details for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, _CCID, , _BCCID)
                        End If
                    End If
                Next

                'This code is used for send "Pay Slip With Leave Details"
            ElseIf DdlreportType.SelectedValue = "SL" Then
                'added by geeta IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString() on 11 sep 2012
                var = "M" + "~" + "" + "~" + "" + "~" + "" + "~" + "" + "~" + COSTCENTER + "~" + LOCATION + "~" + "" + "~" + "" + "~" + Fk_Emp_Code.ToString + "~" + "" + "~" + monthvalue + "~" + YearVal + "~" + "" + "~" + "1" + "~" + "true" + "~" + PM13 + "~~1~" + "true" + "~" + "" + "~" + EmpPassType.ToString + "~L" + "~" + Reportid + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~SL" + "~" + IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString
                _Arr = Split(var, "~")
                _Dtbl = SalarySlipsWithLeave.getSourceTable(, _Arr)
                For counter = 0 To _Dtbl.Rows.Count - 1
                    If txtccc.Text.ToString.Trim <> "" And txtBCC.Text.ToString.Trim <> "" Then
                        lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary Slip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, txtccc.Text.ToString.Trim, , txtBCC.Text.ToString.Trim)
                    ElseIf txtccc.Text.ToString.Trim <> "" Then
                        lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary Slip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, txtccc.Text.ToString.Trim)
                    ElseIf txtBCC.Text.ToString.Trim <> "" Then
                        lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary Slip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, From_Mail, , txtBCC.Text.ToString.Trim)
                    Else
                        lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary Slip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, From_Mail)
                    End If
                Next

                'This code is used for send "Pay Slip Without Leave Details"
            ElseIf DdlreportType.SelectedValue = "S" Then
                var = "" & "~~~~" & COSTCENTER & "~" & LOCATION & "~~~" & Fk_Emp_Code.ToString & "~~" & monthvalue & "~" + YearVal & "~~" & "1" & "~" & "true" & "~" & PM13 & "~" & "M" & "~" & "" & "~" & EmpPassType.ToString & "~L" & "~" & Reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~S"
                _Arr = Split(var, "~")
                _Dtbl = SalSlipsWithoutLeave.getSourceTable(, _Arr)
                For counter = 0 To _Dtbl.Rows.Count - 1
                    If txtccc.Text.ToString.Trim <> "" And txtBCC.Text.ToString.Trim <> "" Then
                        lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary slip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, txtccc.Text.ToString.Trim, , txtBCC.Text.ToString.Trim)
                    ElseIf txtccc.Text.ToString.Trim <> "" Then
                        lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary slip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, txtccc.Text.ToString.Trim)
                    ElseIf txtBCC.Text.ToString.Trim <> "" Then
                        lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary slip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, From_Mail, , txtBCC.Text.ToString.Trim)
                    Else
                        lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary slip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, From_Mail)
                    End If
                Next
                'This code is used for send "PTC"
            ElseIf DdlreportType.SelectedValue = "53" Then
                var = "~~~~" + COSTCENTER + "~~~~" + Fk_Emp_Code + "~~" + monthvalue + "~" + YearVal + "~~" + PM13 + "~" + "A" + "~" + Reportid + "~H" + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~" + "" + "~" + EmpPassType.ToString + "~PTC" + "~" + "~~~"
                _Arr = Split(var, "~")
                _Dtbl = _ClsPTC.getSourceTable(, _Arr)
                For counter = 0 To _Dtbl.Rows.Count - 1
                    If txtccc.Text.ToString.Trim <> "" And txtBCC.Text.ToString.Trim <> "" Then
                        lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary Slip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, txtccc.Text.ToString.Trim, , txtBCC.Text.ToString.Trim)
                    ElseIf txtccc.Text.ToString.Trim <> "" Then
                        lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary Slip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, txtccc.Text.ToString.Trim, , From_Mail)
                    ElseIf txtBCC.Text.ToString.Trim <> "" Then
                        lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary Slip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, From_Mail, , txtBCC.Text.ToString.Trim)
                    Else
                        lblmsg.Text = _objCommon.SendMailBySMTPNew(From_Mail, _Dtbl.Rows(counter).Item("MailID").ToString, "Salary Slip for the month of " & ddlMonthYear.SelectedItem.ToString, _Dtbl.Rows(counter).Item("StrBuild").ToString, From_Mail, , From_Mail)
                    End If
                Next

            End If

            If DdlreportType.SelectedValue = "YTD" Or DdlreportType.SelectedValue = "R" Or DdlreportType.SelectedValue = "52" Or DdlreportType.SelectedValue = "SL" Or DdlreportType.SelectedValue = "S" Then
                If _Dtbl.Rows.Count > 0 Then
                    For counter = 0 To _Dtbl.Rows.Count - 1
                        If _Dtbl.Rows(counter).Item("MailID").ToString.Replace("N/A", "") <> "" Then
                            arrparam(0) = New SqlClient.SqlParameter("@pk_emp_code", _Dtbl.Rows(counter).Item("EmpCode").ToString)
                            arrparam(1) = New SqlClient.SqlParameter("@MM", ddlMonthYear.SelectedValue.ToString)
                            arrparam(2) = New SqlClient.SqlParameter("@YY", HidYear.Value.ToString)
                            arrparam(3) = New SqlClient.SqlParameter("@ReportTtype", DdlreportType.SelectedValue.ToString)
                            'following sp save those employee in the database whom mail has been sent
                            _ObjData.GetDataSetProc("Sp_Ins_TrnDocMailStatus", arrparam)
                            'for count the record
                            _TotRecord = _TotRecord + 1
                        Else
                            If MailNotSentCount.ToString = "" Then
                                MailNotSentCount = _Dtbl.Rows(counter).Item("EmpCode").ToString
                            Else
                                MailNotSentCount += ", " & _Dtbl.Rows(counter).Item("EmpCode").ToString
                            End If
                        End If
                    Next
                End If
            End If
            'Number of Records Sent
            lblMsgSlip.Text = CType(_TotRecord, String) + " email(s) have been sent."
            lblmsg.Text = ""

            'Mail Not Sent of Employee
            If MailNotSentCount.ToString <> "" Then
                _objCommon.ShowMessage("M", lblMailMsg, "Mail not sent for the following employee code(s): " & MailNotSentCount.ToString, True)
            End If

            BtnSend.CssClass = "btn"

            If DdlreportType.SelectedValue = "I" Then
                SendInvestmentDetials()
            Else
                populateDgPayslip()
            End If
        End Sub
        'following proc send the mail as attachement to those employee which is selected in the DIV
        Private Sub SendReportPDF(ByVal _A As String, Optional ByRef UsedFor As String = "", Optional ByRef EmployeeCodes As String = "")
            Dim _TotRec As String, _RecCount As Integer = 0, _Counter As Integer, To_mail As String, arrparam(4) As SqlClient.SqlParameter, _ds As New DataSet _
            , _dt As New DataTable, _dRow As DataRow, _StrEmpCode As String = "", MailNotSentCount As String = "", _strVal As String = Guid.NewGuid.ToString _
            , EmpPassType As String = "", EmpCode As String = "", _DsMailDoc As New DataSet, _dRowDoc As DataRow = Nothing, _DsNS As New DataSet _
            , MsgReturn As String = "", _msg As New List(Of PayrollUtility.UserMessage),
            Dep As String = "", Desig As String = "", Grad As String = "", Level As String = "", CC As String = "", Loc As String = "" _
            , unit As String = "", SalBase As String = "", EmpFName As String = "", EmpLName As String = "", EmpType As String = "", Dtt As New DataTable, EmpCodeSearch As String,
            _array() As String, _AppPath As String = HttpRuntime.AppDomainAppVirtualPath.ToString, Path As String _
            , ArrParams(1) As SqlClient.SqlParameter
            ArrParams(0) = New SqlClient.SqlParameter("@Month", ddlMonthYear.SelectedValue.ToString)
            ArrParams(1) = New SqlClient.SqlParameter("@Year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
            Dtt = _ObjData.GetDataTableProc("PaySP_PaySlipConfigure_Sel", ArrParams)
            If Dtt.Rows.Count > 0 Then
                HidPdfName.Value = Dtt.Rows(0).Item("PdfName").ToString
            End If
            _dt.Columns.Add(New DataColumn("EmpCode"))
            _ds.Tables.Add(_dt)
            _DsMailDoc.Tables.Add("Table1")
            _DsMailDoc.Tables(0).Columns.Add(New DataColumn("EmpCode"))
            lblMsgSlip.Text = ""
            lblMailMsg.Text = ""
            HidYear.Value = Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4)
            Dim row_count As Integer = 0

            'Added by Quadir on 24 Nov 2020 for changing logic Payslip Publish Mode on Search
            If RblNoSearch.SelectedValue = "S" Then
                For counter = 0 To DgPayslip.Items.Count - 1
                    If CType(DgPayslip.Items(counter).FindControl("chkEmpHold"), CheckBox).Checked = True Then
                        EmpCodeSearch = EmpCodeSearch + DgPayslip.Items(counter).Cells(1).Text.ToString + ","
                    End If
                Next
            End If

            If RblNoSearch.SelectedValue = "P" Or RblNoSearch.SelectedValue = "S" And ddlRepIn.SelectedValue = "P" Then
                If DdlreportType.SelectedValue = "I" Then
                    _DsNS = ReturnInvestmentDeclaration()
                Else
                    If UsedFor.ToString.ToUpper = "PDF" Then
                        _DsNS = ReturnDsSearch(MsgReturn, "PDF", EmpCodeSearch)
                    Else
                        _DsNS = ReturnDsSearch(MsgReturn, "", EmpCodeSearch)
                    End If
                End If
                If MsgReturn.ToString <> "" Then
                    If (DdlreportType.SelectedValue.ToUpper.Trim = "R" OrElse DdlreportType.SelectedValue.ToUpper.Trim = "S" OrElse DdlreportType.SelectedValue.ToUpper.Trim = "T" OrElse DdlreportType.SelectedValue.Trim = "57") Then
                        lblProcessBarMsg.Text = MsgReturn.ToString
                    Else
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocessF821", "UnLoadPaySlipProgress();", True)
                        _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = MsgReturn.ToString})
                        _objCommon.ShowMessage(_msg)
                        Exit Sub
                    End If
                End If
                If _DsNS.Tables.Count > 0 Then
                    If _DsNS.Tables(0).Rows.Count > 0 Then
                        _TotRec = _DsNS.Tables(0).Rows.Count.ToString
                        row_count = _DsNS.Tables(0).Rows.Count
                        For _Counter = 0 To _DsNS.Tables(0).Rows.Count - 1
                            To_mail = _DsNS.Tables(0).Rows(_Counter)("Email").ToString
                            PK_emp_code = _DsNS.Tables(0).Rows(_Counter)("fk_emp_code").ToString
                            If _DsNS.Tables(0).Rows(_Counter)("Status").ToString.Trim.ToUpper = "SALARY NOT PROCESSED" And Not DdlreportType.SelectedValue.Equals("62") Then
                                If MailNotSentCount.ToString = "" Then
                                    MailNotSentCount = PK_emp_code
                                Else
                                    MailNotSentCount += ", " & PK_emp_code
                                End If
                            Else
                                If To_mail.ToString.Replace("N/A", "") <> "" Or _A.ToString.ToUpper = "S" Then
                                    _dRow = _ds.Tables(0).NewRow
                                    _dRow(0) = PK_emp_code
                                    _ds.Tables(0).Rows.Add(_dRow)
                                    EmpCode = EmpCode + PK_emp_code + ","
                                    If _A.ToString = "" Then
                                        _StrEmpCode = _StrEmpCode + "," + PK_emp_code
                                        _dRowDoc = _DsMailDoc.Tables(0).NewRow
                                        _dRowDoc(0) = PK_emp_code
                                        _DsMailDoc.Tables(0).Rows.Add(_dRowDoc)
                                    End If
                                    _RecCount = _RecCount + 1
                                Else
                                    If MailNotSentCount.ToString = "" Then
                                        MailNotSentCount = PK_emp_code
                                    Else
                                        MailNotSentCount += ", " & PK_emp_code
                                    End If
                                End If
                            End If

                        Next
                        GenerateLogFile(EmpCode, MailNotSentCount)
                    Else
                        IsProcessBarStated = False
                        _TotRec = "0"
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocessF822", "UnLoadPaySlipProgress();", True)
                        _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "No record(s) found according to the selection criteria for publish payslip !"})
                        _objCommon.ShowMessage(_msg)
                        Exit Sub
                    End If
                Else
                    IsProcessBarStated = False
                    _TotRec = "0"
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocessF823", "UnLoadPaySlipProgress();", True)
                    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "No record(s) found according to the selection criteria for publish payslip !"})
                    _objCommon.ShowMessage(_msg)
                    Exit Sub
                End If
                _DsNS.Dispose()
            Else
                _TotRec = DgPayslip.Items.Count.ToString
                For _Counter = 0 To DgPayslip.Items.Count - 1
                    If CType(DgPayslip.Items(_Counter).FindControl("chkEmpHold"), CheckBox).Checked = True Then
                        To_mail = DgPayslip.Items(_Counter).Cells(5).Text.ToString
                        PK_emp_code = DgPayslip.Items(_Counter).Cells(1).Text.ToString
                        If To_mail.ToString.Replace("N/A", "") <> "" Or _A.ToString.ToUpper = "S" Then
                            _dRow = _ds.Tables(0).NewRow
                            _dRow(0) = PK_emp_code
                            _ds.Tables(0).Rows.Add(_dRow)
                            EmpCode = EmpCode + PK_emp_code + ","
                            If _A.ToString = "" Then
                                _StrEmpCode = _StrEmpCode + "," + PK_emp_code
                                _dRowDoc = _DsMailDoc.Tables(0).NewRow
                                _dRowDoc(0) = PK_emp_code
                                _DsMailDoc.Tables(0).Rows.Add(_dRowDoc)
                            End If
                            _RecCount = _RecCount + 1
                        Else
                            If MailNotSentCount.ToString = "" Then
                                MailNotSentCount = PK_emp_code
                            Else
                                MailNotSentCount += ", " & PK_emp_code
                            End If
                        End If
                    End If
                Next
                _TotRec = _RecCount
            End If
            If _A = "" Then
                arrparam(0) = New SqlClient.SqlParameter("@Pk_emp_Code", "")
                arrparam(1) = New SqlClient.SqlParameter("@MM", ddlMonthYear.SelectedValue.ToString)
                arrparam(2) = New SqlClient.SqlParameter("@YY", HidYear.Value.ToString)
                arrparam(3) = New SqlClient.SqlParameter("@ReportTtype", DdlreportType.SelectedValue.ToString)
                arrparam(4) = New SqlClient.SqlParameter("@XML", _DsMailDoc.GetXml())
                _ObjData.GetDataSetProc("Sp_Ins_TrnDocMailStatus", arrparam)
            End If

            If EmpCode.ToString <> "" Then
                If DdlreportType.SelectedValue.Equals("S") Or DdlreportType.SelectedValue.Equals("57") Then
                    HidEmpPdf.Value = ""
                End If
                'Added by Quadir on 24 Nov 2020 for changing logic Payslip Publish Mode on Search
                If RblNoSearch.SelectedValue = "S" Then
                    If rbtSlipPubMode.SelectedValue.ToString.ToUpper <> "I" Then
                        HidEmpPdf.Value = ""
                    End If
                    HidEmpPdf.Value = HidEmpPdf.Value + "," + EmpCode.ToString.Trim
                    EmpCode = Left(EmpCode, Len(EmpCode) - 1).ToString
                Else
                    HidEmpPdf.Value = HidEmpPdf.Value + "," + EmpCode.ToString.Trim
                    EmpCode = Left(EmpCode, Len(EmpCode) - 1).ToString
                End If

            End If
            SavePwdRec(_A, EmpCode)
            If _A = "S" Then
                If HidEmailCCBCC.Value.ToString <> "" And RblGrpbyPublish.SelectedValue = "" Then
                    lblMsgSlip.Text = "Email Sent to " & txtccc.Text.ToString & IIf(txtBCC.Text.ToString <> "", IIf(txtccc.Text.ToString <> "", ",", "") & txtBCC.Text.ToString, txtBCC.Text.ToString)
                ElseIf RblGrpbyPublish.SelectedValue = "" Then
                    lblMsgSlip.Text = CType(_RecCount, String) + " of " + _TotRec + " pay slips published...."
                Else
                    lblMsgSlip.Text = "Email sent to unit authority!"
                End If
            End If

            If MailNotSentCount.ToString <> "" And _A.ToString = "" Then
                _objCommon.ShowMessage("M", lblMailMsg, "Mail not sent for the following employee code(s): " + MailNotSentCount.ToString, True)
            End If

            If Len(_StrEmpCode) > 0 Then
                _StrEmpCode = Right(_StrEmpCode, Len(_StrEmpCode) - 1)
            End If

            If _ds.Tables(0).Rows.Count > 0 Then
                row_count = _ds.Tables(0).Rows.Count
                _ds.WriteXml(Server.MapPath("XMLFiles\" & _strVal.ToString & ".xml"))
                Dim monthvalue As String, YearVal As String, HDocStatus As String, HReimbSts As String, QryString As String = "", PM1 As String = "",
                PM11 As String = "0", PM13 As String = "A", PM14 As String = "true", PopupScript As String, var As String, ReportType As String _
                , RepType As String, flag As String = "", EmpStatus As String = USearch.UCddlEmp.ToString, COSTCENTER As String = "", Grpwise As String = "",
                reportid As String = "", MailFrom As String = "", _MonthValue As String = ddlMonthYear.SelectedValue.ToString, _YearValue As String = "",
                dt As DataTable = Nothing, LoopVar As Integer = 0, InvCode As String = "", LOCATION As String = USearch.UCddllocation.ToString, PayCode As String = ""
                HDocStatus = "P"
                HReimbSts = "N"
                monthvalue = Me.ddlMonthYear.SelectedValue
                YearVal = HidYear.Value.ToString
                ReportType = "V"
                RepType = "P"
                HttpContext.Current.Session.Add("MsgPDF", "")
                flag = _A
                EmpPassType = ddlEmpPass.SelectedValue.ToString
                PM13 = ddlshowsal.SelectedValue.ToString
                reportid = HidRepId.Value.ToString
                COSTCENTER = USearch.UCddlcostcenter.ToString
                'This code is used for Save/mail "Pay Slip With Tax Details"
                If DdlreportType.SelectedValue = "R" Or DdlreportType.SelectedValue = "52" Or DdlreportType.SelectedValue = "49" Or DdlreportType.SelectedValue = "51" Then
                    If HidEmailCCBCC.Value.ToString.Trim = "SP" Then
                        var = "P~~~~~" + COSTCENTER + "~~~~~~" + monthvalue + "~" + YearVal + "~" + PM1 + "~" + PM1 + "~" + PM1 + "~1~" + PM14 + "~" + PM13 + "~" + _A + "~" + EmpStatus.ToString + "~" + EmpPassType.ToString + "~" + "N" + "~" + "" + "~" + reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~SP~" & _strVal.ToString & ".xml~" & IIf(chkHelp1.Checked = True, "Y", "N").ToString & "~" & HidPdfName.Value.ToString
                    ElseIf HidEmailCCBCC.Value.ToString.Trim = "R" Or HidEmailCCBCC.Value.ToString.Trim = "52" Then
                        var = "P~~~~~" + COSTCENTER + "~~~~~~" + monthvalue + "~" + YearVal + "~" + PM1 + "~" + PM1 + "~" + PM1 + "~1~" + PM14 + "~" + PM13 + "~" + _A + "~" + EmpStatus.ToString + "~" + EmpPassType.ToString + "~" + "N" + "~" + "" + "~" + reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~R~" & _strVal.ToString & ".xml~" & IIf(chkHelp1.Checked = True, "Y", "N") & "~" & HidPdfName.Value.ToString
                    Else
                        If UsedFor.ToString.ToUpper = "PDF" Then
                            Dim BPath As String = _objCommon.GetBaseHref(), filepat As String = HttpContext.Current.Server.MapPath(HttpContext.Current.Request.ApplicationPath())
                            var = "P~~~~~" + COSTCENTER + "~~~~~~" + monthvalue + "~" + YearVal + "~" + PM1 + "~" + PM1 + "~" + PM1 + "~1~" + PM14 + "~" + PM13 + "~" + _A + "~" + EmpStatus.ToString + "~" + EmpPassType.ToString + "~" + "N" + "~" + "" + "~" + reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~51~" & _strVal.ToString & ".xml~" & IIf(chkHelp1.Checked = True, "Y", "N").ToString & "~" & HidPdfName.Value.ToString & "~" & BPath.ToString & "~" & filepat.ToString
                        Else
                            var = "P~~~~~" + COSTCENTER + "~~~~~~" + monthvalue + "~" + YearVal + "~" + PM1 + "~" + PM1 + "~" + PM1 + "~1~" + PM14 + "~" + PM13 + "~" + _A + "~" + EmpStatus.ToString + "~" + EmpPassType.ToString + "~" + "N" + "~" + "" + "~" + reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~51~" & _strVal.ToString & ".xml~" & IIf(chkHelp1.Checked = True, "Y", "N").ToString & "~" & HidPdfName.Value.ToString
                        End If
                    End If
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    If DdlreportType.SelectedValue.ToString.ToUpper = "R" And UsedFor.ToString.ToUpper = "PDF" Then
                        Dim dtp As DataTable = InitializeProcessBar(_ds.GetXml(), "I", "TAXSLIP", DdlreportType.SelectedValue.Replace("'", "").ToUpper, DdlreportType.SelectedItem.Text.Replace("'", ""), _ds.Tables(0).Rows.Count)
                        If (dtp.Rows.Count > 0 AndAlso dtp.Rows(0)("IsAbleToStart").ToString = "0") Then
                            lblMailMsg.Text = dtp.Rows(0)("Msg").ToString
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocessF824", "UnLoadPaySlipProgress();", True)
                            _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = dtp.Rows(0)("Msg").ToString})
                            _objCommon.ShowMessage(_msg)
                            LnkPDF.Style.Value = "display:None"
                            IsProcessBarStated = False
                            Exit Sub
                        End If
                        HidPreVal.Value = var.ToString
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "ShowTaxDetails();", True)
                    Else
                        PopupScript = ""
                        PopupScript = PopupScript & "window.open('Reports/PreSalSlip.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                        ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                        PopupScript = ""
                    End If

                ElseIf DdlreportType.SelectedValue = "TL" Then
                    If HidEmailCCBCC.Value.ToString.Trim = "SP" Then
                        var = "P~~~~~" + COSTCENTER + "~" + LOCATION + "~~~~~" + monthvalue + "~" + YearVal + "~" + PM1 + "~" + PM1 + "~" + PM1 + "~1~" + PM14 + "~" + PM13 + "~" + _A + "~" + EmpStatus.ToString + "~" + EmpPassType.ToString + "~" + reportid + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~SP~" & _strVal.ToString & ".xml"
                    Else
                        var = "P~~~~~" + COSTCENTER + "~" + LOCATION + "~~~~~" + monthvalue + "~" + YearVal + "~" + PM1 + "~" + PM1 + "~" + PM1 + "~1~" + PM14 + "~" + PM13 + "~" + _A + "~" + EmpStatus.ToString + "~" + EmpPassType.ToString + "~" + reportid + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~TL~" & _strVal.ToString & ".xml"
                    End If
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/PreSalSlip_LoanDetail.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                    'This code is used for Save/mail "Pay Slip With Leave Details"
                ElseIf DdlreportType.SelectedValue.ToUpper = "SL" Or DdlreportType.SelectedValue.ToString.ToUpper = "56" Then
                    Dim _array1() As String, _AppPath1 As String = HttpRuntime.AppDomainAppVirtualPath.ToString

                    _array1 = Split(_AppPath1, "/")
                    _AppPath1 = _array1(_array1.Length - 1)
                    HidAppPath.Value = _AppPath1
                    If DdlreportType.SelectedValue.ToString.ToUpper = "56" Then
                        var = "P" & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & COSTCENTER & "~" & LOCATION & "~" & PM1 & "~" & PM1 & "~~" & PM1 & "~" & monthvalue & "~" & YearVal & "~" & PM1 & "~" & "1" & "~" & PM14 & "~" & PM13 & "~~1~" & PM14 & "~" & _A & "~" & EmpPassType.ToString & "~L" & "~" & reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~56" & "~" & IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString & "~" & _strVal.ToString & ".xml"
                    Else
                        If HidEmailCCBCC.Value = "SP" Then
                            var = "P" & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & COSTCENTER & "~" & LOCATION & "~" & PM1 & "~" & PM1 & "~~" & PM1 & "~" & monthvalue & "~" & YearVal & "~" & PM1 & "~" & "1" & "~" & PM14 & "~" & PM13 & "~~1~" & PM14 & "~" & _A & "~" & EmpPassType.ToString & "~L" & "~" & reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~SP" & "~" & IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString & "~" & _strVal.ToString & ".xml"
                        Else
                            var = "P" & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & COSTCENTER & "~" & LOCATION & "~" & PM1 & "~" & PM1 & "~~" & PM1 & "~" & monthvalue & "~" & YearVal & "~" & PM1 & "~" & "1" & "~" & PM14 & "~" & PM13 & "~~1~" & PM14 & "~" & _A & "~" & EmpPassType.ToString & "~L" & "~" & reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~SL" & "~" & IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString & "~" & _strVal.ToString & ".xml"
                        End If
                    End If


                    Dim _mm As String = "", _yyyy As String = ""
                    If ddlMonthYear.SelectedItem IsNot Nothing Then
                        _mm = CType(ddlMonthYear.SelectedValue.ToString, Integer)
                        _yyyy = Right(ddlMonthYear.SelectedItem.Text.ToString, 4)
                    End If
                    Dim gcs_service As Integer = 0
                    If DdlreportType.SelectedValue.ToString.ToUpper = "SL" And HidEmailCCBCC.Value <> "SP" And _A <> "" And gcs_service_obj.IsPaySlipAllowedByMonthYear(Session("compCode"), "SLIPWILVE", _mm, _yyyy) Then
                        gcs_service = 1
                    End If
                    If gcs_service = 1 Then
                        Dim _dt_temp, temp As New DataTable
                        _dt_temp = _ObjData.ExecSQLQuery("Insert Into GcsSalaryProcessStatus(process_user_id, status, total_processed, total_to_process, record_created, Process_Type, mm, YYYY) VALUES('" & Convert.ToString(HttpContext.Current.Session("UId")) & "','START','0', '" & Convert.ToString(row_count) & "',GetDate(), 'SLIPWILVE', '" & monthvalue & "', '" & YearVal & "'); SELECT SCOPE_IDENTITY() AS id;")
                        If _dt_temp.Rows.Count > 0 Then
                            Session("process_status_id") = _dt_temp.Rows(0)("id")
                            process_status_id.Value = _dt_temp.Rows(0)("id").ToString
                            LnkPDF.Style.Value = "display:none"
                            download_pdf1.Style.Value = "border:0;cursor:pointer;"
                        End If
                    Else
                        download_pdf1.Style.Value = "display:none"

                    End If
                    'download_pdf2.Style.Value = "border:0;cursor:pointer;"

                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    HidPreVal.Value = var.ToString
                    If gcs_service = 1 And HidEmailCCBCC.Value <> "SP" Then
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopupwithleave", "ShowWithLeaveDetails();", True)
                    Else
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "closepopup", "CloseSlipProgressbar();", True)
                        PopupScript = ""
                        PopupScript = PopupScript & "window.open('Reports/Pre_SalarySlip.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                        ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                        PopupScript = ""
                    End If

                    'added by geeta on 21 sep 09 to send mail on "TDS Estimation Slip" in pdf
                ElseIf DdlreportType.SelectedValue.ToUpper = "T" Then
                    If HidEmailCCBCC.Value = "SP" Then
                        var = HDocStatus + "~" + monthvalue + "~" + YearVal + "~" + "" + "~" + reportid + "~" + EmpPassType + "~" + COSTCENTER + "~" + ReportType + "~" + flag + "~" + HReimbSts + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~SP~" & _strVal.ToString & ".xml"
                    Else
                        var = HDocStatus + "~" + monthvalue + "~" + YearVal + "~" + "" + "~" + reportid + "~" + EmpPassType + "~" + COSTCENTER + "~" + ReportType + "~" + flag + "~" + HReimbSts + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~T~" & _strVal.ToString & ".xml"
                    End If
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    If ddlRepIn.SelectedValue.ToString.ToUpper = "P" Then
                        Dim dtp As DataTable = InitializeProcessBar(_ds.GetXml(), "I", "SLIPTDSV", DdlreportType.SelectedValue.Replace("'", "").ToUpper, DdlreportType.SelectedItem.Text.Replace("'", ""), _ds.Tables(0).Rows.Count)
                        If (dtp.Rows.Count > 0 AndAlso dtp.Rows(0)("IsAbleToStart").ToString = "0") Then
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocessF825", "UnLoadPaySlipProgress();", True)
                            _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = dtp.Rows(0)("Msg").ToString})
                            _objCommon.ShowMessage(_msg)
                            LnkPDF.Style.Value = "display:None"
                            IsProcessBarStated = False
                            lblMailMsg.Text = ""
                            Exit Sub
                        End If
                        _array = Split(_AppPath, "/")
                        _AppPath = _array(_array.Length - 1)
                        HidAppPath.Value = _AppPath
                        Path = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Left(MonthName(CType(ddlMonthYear.SelectedValue.ToString, Integer)), 3) & Right(ddlMonthYear.SelectedItem.Text.ToString, 4) & "\TDSEstimationSlip\"
                        Dim _mm As String = "", _yyyy As String = ""
                        If ddlMonthYear.SelectedItem IsNot Nothing Then
                            _mm = CType(ddlMonthYear.SelectedValue.ToString, Integer)
                            _yyyy = Right(ddlMonthYear.SelectedItem.Text.ToString, 4)
                        End If
                        Dim gcs_service As Integer = 0
                        If gcs_service_obj.IsPaySlipAllowedByMonthYear(Session("compCode"), "SLIPTDSV", _mm, _yyyy) Then
                            gcs_service = 1
                        End If
                        If gcs_service = 1 Then
                            Session("pdf_file_location") = Path.ToString
                        End If
                        If Not Directory.Exists(Path) Then
                            Directory.CreateDirectory(Path)
                        End If
                        HidPath.Value = Replace(Replace(Path, "\", "~").ToString, "/", "~").ToString
                        HidPreVal.Value = var.ToString
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopuptds", "ShowTdsDetails();", True)
                    Else
                        PopupScript = ""
                        PopupScript = PopupScript & "window.open('Reports/PreNewTdsEstimationSlip.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                        ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                        PopupScript = ""
                    End If
                    'PopupScript = ""
                    'PopupScript = PopupScript & "window.open('Reports/PreNewTdsEstimationSlip.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    'ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    'PopupScript = ""
                    'Salary Slip Without Leave Details
                ElseIf DdlreportType.SelectedValue.ToUpper = "S" Then
                    Grpwise = ""
                    If HidEmailCCBCC.Value = "UNT" And RblGrpbyPublish.SelectedValue <> "" And ddlRepIn.SelectedValue = "P" Then
                        Grpwise = RblGrpbyPublish.SelectedValue.ToString
                    End If
                    Dim reportmodule As String = HttpContext.Current.Session("ModuleType").ToString
                    If HidEmailCCBCC.Value = "SP" Then
                        var = "" & "~~~~" & COSTCENTER & "~" & LOCATION & "~~~~~" & monthvalue & "~" + YearVal & "~~" & "1" & "~" & PM14 & "~" & PM13 & "~" & "P" & "~" & _A & "~" & EmpPassType.ToString & "~L" & "~" & reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~SP~" & _strVal.ToString & ".xml" & "~" & Grpwise
                    ElseIf UsedFor.ToString.ToUpper = "SPDF" Then
                        Dim BPath As String = _objCommon.GetBaseHref(), filepat As String = HttpContext.Current.Server.MapPath(HttpContext.Current.Request.ApplicationPath())
                        var = "" & "~~~~" & COSTCENTER & "~" & LOCATION & "~~~~~" & monthvalue & "~" + YearVal & "~~" & "1" & "~" & PM14 & "~" & PM13 & "~" & "P" & "~" & _A & "~" & EmpPassType.ToString & "~L" & "~" & reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~S~" & _strVal.ToString & ".xml" & "~" & Grpwise & "~" & BPath.ToString & "~" & filepat.ToString
                    Else
                        var = "" & "~~~~" & COSTCENTER & "~" & LOCATION & "~~~~~" & monthvalue & "~" + YearVal & "~~" & "1" & "~" & PM14 & "~" & PM13 & "~" & "P" & "~" & _A & "~" & EmpPassType.ToString & "~L" & "~" & reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~S~" & _strVal.ToString & ".xml" & "~" & Grpwise
                    End If

                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If

                    If DdlreportType.SelectedValue.ToString.ToUpper = "S" And UsedFor.ToString.ToUpper = "SPDF" Then
                        Dim dtp As DataTable = InitializeProcessBar(_ds.GetXml(), "I", "SLIPWOLVE", DdlreportType.SelectedValue.Replace("'", "").ToUpper, DdlreportType.SelectedItem.Text.Replace("'", ""), _ds.Tables(0).Rows.Count)
                        If (dtp.Rows.Count > 0 AndAlso dtp.Rows(0)("IsAbleToStart").ToString = "0") Then
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocessF826", "UnLoadPaySlipProgress();", True)
                            lblMailMsg.Text = dtp.Rows(0)("Msg").ToString
                            _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = dtp.Rows(0)("Msg").ToString})
                            _objCommon.ShowMessage(_msg)
                            LnkPDF.Style.Value = "display:None"
                            IsProcessBarStated = False
                            Exit Sub
                        End If
                        HidPreVal.Value = var.ToString
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "ShowSlipWOLeave();", True)
                    Else
                        PopupScript = ""
                        PopupScript = PopupScript & "window.open('Reports/Pre_SalarySlipInclude.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                        ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    End If

                    PopupScript = ""
                ElseIf DdlreportType.SelectedValue.ToUpper = "L" Then
                    If CType(_MonthValue, Integer) < 4 Then
                        _YearValue = CType(Session("Efindate"), DateTime).Year.ToString
                    Else
                        _YearValue = CType(Session("Sfindate"), DateTime).Year.ToString
                    End If
                    QryString = _MonthValue + "~" + _YearValue
                    If Not IsNothing(System.Configuration.ConfigurationManager.AppSettings("AuthenticationType").ToString()) Then
                        If System.Configuration.ConfigurationManager.AppSettings("AuthenticationType").ToString() = "AD" Then
                            QryString = System.Configuration.ConfigurationManager.AppSettings("BASE_HREF").ToString() & "/frmadlogin.aspx?id=" & QryString & "', 'CustomPopUp','width=500, height=250,left=3,top=0,location=0, menubar=1, resizable=no');"
                        Else
                            QryString = System.Configuration.ConfigurationManager.AppSettings("BASE_HREF").ToString() & "/frmMainLogin.aspx?id=" & QryString & "', 'CustomPopUp','width=500, height=250,left=3,top=0,location=0, menubar=1, resizable=no');"
                        End If
                    Else
                        QryString = System.Configuration.ConfigurationManager.AppSettings("BASE_HREF").ToString() & "/frmMainLogin.aspx?id=" & QryString & "', 'CustomPopUp','width=500, height=250,left=3,top=0,location=0, menubar=1, resizable=no');"
                    End If

                    'This is used to Save/mail in PDF "Salary Slip with Investment"
                ElseIf DdlreportType.SelectedValue.ToUpper = "SI" Then
                    dt = _ObjData.GetDataTableProc("sel_sections")
                    For LoopVar = 0 To dt.Rows.Count - 1
                        InvCode = InvCode + "'" + dt.Rows(LoopVar).Item("sectionnumber").ToString + "',"
                    Next
                    If HidEmailCCBCC.Value = "SP" Then
                        var = "" + "~~~~" + COSTCENTER.ToString + "~~~~~~" + monthvalue + "~" + YearVal + "~~0~1~1" + "~" + PM13 + "~P" + "~" + _A + "~~0~" + InvCode.ToString + "~" + EmpPassType.ToString & "~L" & "~" + reportid + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~SP~" & _strVal.ToString & ".xml"
                    Else
                        var = "" + "~~~~" + COSTCENTER.ToString + "~~~~~~" + monthvalue + "~" + YearVal + "~~0~1~1" + "~" + PM13 + "~P" + "~" + _A + "~~0~" + InvCode.ToString + "~" + EmpPassType.ToString & "~L" & "~" + reportid + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~SI~" & _strVal.ToString & ".xml"
                    End If
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/Pre_SalSlipwithINVEST.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf DdlreportType.SelectedValue.ToUpper = "SH" Then
                    If HidEmailCCBCC.Value = "SP" Then
                        var = "" + EmpCode + "~" + "F" + "~~~~~~~~~~" + "A" + "~" + monthvalue + "~" + YearVal + "~" + "A" + "~~" + EmpPassType.ToString + "~~" + ddlRepIn.Text + "~" + _A + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~SP" + "~" + reportid.ToString
                    Else
                        var = "" + EmpCode + "~" + "F" + "~~~~~~~~~~" + "A" + "~" + monthvalue + "~" + YearVal + "~" + "A" + "~~" + EmpPassType.ToString + "~~" + ddlRepIn.Text + "~" + _A + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~SH" + "~" + reportid.ToString
                    End If
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/Pre_MonthlySalarySlipHindiforpgl.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                    ''SUSHIL 6 FEB 2019 COSTCENTER ADDED
                ElseIf DdlreportType.SelectedValue.ToUpper = "RS" Then
                    If HidEmailCCBCC.Value = "SP" Then
                        var = "P~~~~~" + COSTCENTER.ToString + "~~~" + "S" + "~" + EmpCode + "~~" + monthvalue + "~" + YearVal + "~~~~~~~~~~" + "A" + "~" + _A + "~" + EmpPassType.ToString + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~SP" + "~" + reportid.ToString & "~~~~" & "~" & ddlshowsal.SelectedValue.ToString & "~" & _strVal.ToString & ".xml"
                    Else
                        var = "P~~~~~" + COSTCENTER.ToString + "~~~" + "S" + "~" + EmpCode + "~~" + monthvalue + "~" + YearVal + "~~~~~~~~~~" + "A" + "~" + _A + "~" + EmpPassType.ToString + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~RS" + "~" + reportid.ToString & "~~~~" & "~" & ddlshowsal.SelectedValue.ToString & "~" & _strVal.ToString & ".xml"
                    End If
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/Pre_SalSlipWithReimb.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf DdlreportType.SelectedValue.ToUpper = "YTD" Then
                    If HidEmailCCBCC.Value = "SP" Then
                        var = "~~~~" + COSTCENTER + "~~~~" + EmpCode + "~~" + monthvalue + "~" + YearVal + "~~" + PM13 + "~" + "A" + "~" + reportid + "~P" + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~" + _A + "~" + EmpPassType.ToString + "~SP"
                    Else
                        var = "~~~~" + COSTCENTER + "~~~~" + EmpCode + "~~" + monthvalue + "~" + YearVal + "~~" + PM13 + "~" + "A" + "~" + reportid + "~P" + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~" + _A + "~" + EmpPassType.ToString + "~YTD"
                    End If
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/PreSalSlipYTD.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf DdlreportType.SelectedValue.ToUpper = "43" Then
                    var = "~~~~" + COSTCENTER.ToString + "~~~~" + EmpCode.ToString + "~~" + monthvalue.ToString + "~" + YearVal.ToString + "~~A~" + DDLPaySlipType.SelectedValue + "~P~" + EmpPassType.ToString
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/Pre_WageSlipForEasySource.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf DdlreportType.SelectedValue.ToUpper = "RN" Then
                    If _A.ToString.ToUpper = "S" Then
                        If HidEmailCCBCC.Value = "SP" Then
                            var = "P~~~~~" + COSTCENTER + "~~~" + "S" + "~" + EmpCode + "~~" + monthvalue + "~" + YearVal + "~~~~~1~~A~~~" + "A" + "~" + _A + "~" + "3" + "~" + "R" + "~" + "P1" & "~" & EmpPassType.ToString & "~" & "PP" & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~SP" & "~" & reportid.ToString & "~" & ddlshowsal.SelectedValue.ToString
                        Else
                            var = "P~~~~~" + COSTCENTER + "~~~" + "S" + "~" + EmpCode + "~~" + monthvalue + "~" + YearVal + "~~~~~1~~A~~~" + "A" + "~" + _A + "~" + "3" + "~" + "R" + "~" + "P1" & "~" & EmpPassType.ToString & "~" & "PP" & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~RN" & "~" & reportid.ToString & "~" & ddlshowsal.SelectedValue.ToString
                        End If
                    Else
                        var = "P~~~~~" + COSTCENTER + "~~~" + "S" + "~" + EmpCode + "~~" + monthvalue + "~" + YearVal + "~~~~~1~~A~~~" + "A" + "~" + _A + "~" + "3" + "~" + "R" + "~" + "P1" & "~" & EmpPassType.ToString & "~" & "PE" & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~RN" & "~" & reportid.ToString & "~" & ddlshowsal.SelectedValue.ToString
                    End If
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/Pre_SalSlipWithReimbNewFormat.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf DdlreportType.SelectedValue = "50" Then
                    If HidEmailCCBCC.Value = "SP" Then
                        var = "P" & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & COSTCENTER & "~" & LOCATION & "~" & PM1 & "~" & PM1 & "~~" & PM1 & "~" & monthvalue & "~" & YearVal & "~" & PM1 & "~" & "1" & "~" & PM14 & "~" & PM13 & "~~1~" & PM14 & "~" & _A & "~" & EmpPassType.ToString & "~L" & "~" & reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~SP" & "~" & IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString & "~" & _strVal.ToString & ".xml"
                    Else
                        var = "P" & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & COSTCENTER & "~" & LOCATION & "~" & PM1 & "~" & PM1 & "~~" & PM1 & "~" & monthvalue & "~" & YearVal & "~" & PM1 & "~" & "1" & "~" & PM14 & "~" & PM13 & "~~1~" & PM14 & "~" & _A & "~" & EmpPassType.ToString & "~L" & "~" & reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~SL" & "~" & IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString & "~" & _strVal.ToString & ".xml"
                    End If
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/Pre_SalarySlip.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                    'PTC
                ElseIf DdlreportType.SelectedValue.ToUpper = "53" Then
                    If HidEmailCCBCC.Value = "SP" Then
                        var = "~~~~" + COSTCENTER + "~~~~" + EmpCode + "~~" + monthvalue + "~" + YearVal + "~~" + PM13 + "~" + "A" + "~" + reportid + "~P" + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~" + _A + "~SP" + "~" + EmpPassType.ToString + "~PTC" + "~" + "SP~"
                    Else
                        var = "~~~~" + COSTCENTER + "~~~~" + EmpCode + "~~" + monthvalue + "~" + YearVal + "~~" + PM13 + "~" + "A" + "~" + reportid + "~P" + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~" + _A + "~" + _A + "~" + EmpPassType.ToString + "~PTC" + "~" + "~"
                    End If

                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/PreSalSlipPTC.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf DdlreportType.SelectedValue = "55" And chkMerge.Checked = False Then
                    var = DdlreportType.SelectedValue & "~" & HDocStatus + "~" + monthvalue + "~" + YearVal + "~" + EmpPassType + "~" + COSTCENTER + "~" + HReimbSts + "~" _
                        & _strVal.ToString & ".xml~SP" '& "~" & IIf(chkMerge.Checked = True, "Y", "N")
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/PreNewTdsEstimationSlip.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf DdlreportType.SelectedValue = "55" And chkMerge.Checked = True Then
                    'Code commented by Rohtas Singh on 20 Mar 2018 for worked on future plan 
                    'This code is used for merge "Salary Slip with TAX Details with TAX Competetion Slip
                    'var = DdlreportType.SelectedValue + "~" + "P~~~~~" + COSTCENTER + "~~~~~~" + monthvalue + "~" + YearVal + "~" + PM1 + "~" + PM1 + "~" + PM1 + "~1~" + _
                    '    PM14 + "~" + PM13 + "~" + _A + "~" + EmpStatus.ToString + "~" + EmpPassType.ToString + "~" + "N" + "~" + "" + "~" + reportid + "~" + _
                    '    txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~51~" + _strVal.ToString + ".xml~" + IIf(chkHelp1.Checked = True, "Y", "N").ToString + _
                    '    "~" + IIf(chkMerge.Checked = True, "Y", "N").ToString + "~" + HReimbSts

                    'Added by Rohtas Singh on 20 Mar 2018
                    'This code is used for merge "Salary Slip with TAX Details with TAX Competetion Slip
                    var = DdlreportType.SelectedValue + "~" + "P" & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & COSTCENTER & "~" & LOCATION & "~" & PM1 & "~" &
                        PM1 & "~~" & PM1 & "~" & monthvalue & "~" & YearVal & "~" & PM1 & "~" & "1" & "~" & PM14 & "~" & PM13 & "~~1~" & PM14 & "~" & _A & "~" &
                       EmpPassType.ToString & "~L" & "~" & reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~SL" & "~" &
                       IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString & "~" & _strVal.ToString & ".xml~" + IIf(chkHelp1.Checked = True, "Y", "N").ToString +
                       "~" + IIf(chkMerge.Checked = True, "Y", "N").ToString + "~" + HReimbSts

                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/PreNewTdsEstimationSlip.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("57") And chkMerge.Checked.Equals(False) Then  'added by ritu malik on 22 may 2018 regarding tax computation IDFC

                    If UsedFor.ToString.ToUpper.Equals("SPDF") Then
                        Dim BPath As String = _objCommon.GetBaseHref(), filepat As String = HttpContext.Current.Server.MapPath(HttpContext.Current.Request.ApplicationPath())
                        var = DdlreportType.SelectedValue & "~" & HDocStatus + "~" + monthvalue + "~" + YearVal + "~" + EmpPassType + "~" + COSTCENTER + "~" _
                            & _strVal.ToString & ".xml~SP" & "~" & BPath.ToString & "~" & filepat.ToString
                    Else
                        var = DdlreportType.SelectedValue & "~" & HDocStatus + "~" + monthvalue + "~" + YearVal + "~" + EmpPassType + "~" + COSTCENTER + "~" _
                            & _strVal.ToString & ".xml~SP"
                    End If


                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If

                    If DdlreportType.SelectedValue.ToString.ToUpper.Equals("57") And UsedFor.ToString.ToUpper.Equals("SPDF") Then
                        Dim dtp As DataTable = InitializeProcessBar(_ds.GetXml(), "I", "FORCAST", DdlreportType.SelectedValue.Replace("'", "").ToUpper, DdlreportType.SelectedItem.Text.Replace("'", ""), _ds.Tables(0).Rows.Count)
                        If (dtp.Rows.Count > 0 AndAlso dtp.Rows(0)("IsAbleToStart").ToString = "0") Then
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openprocessF827", "UnLoadPaySlipProgress();", True)
                            lblMailMsg.Text = dtp.Rows(0)("Msg").ToString
                            _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = dtp.Rows(0)("Msg").ToString})
                            _objCommon.ShowMessage(_msg)
                            LnkPDF.Style.Value = "display:None"
                            IsProcessBarStated = False
                            Exit Sub
                        End If
                        HidPreVal.Value = var.ToString
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "ShowTaxForcast();", True)
                    Else
                        PopupScript = ""
                        PopupScript = PopupScript & "window.open('Reports/PreNewTdsEstimationSlip.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                        ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    End If
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("57") And chkMerge.Checked.Equals(True) Then  'added by ritu malik on 22 may 2018 regarding merge functionality payslip and tax computation IDFC
                    Dim gcs_service As Integer = 0


                    Try
                        Dim _mm As String = "", _yyyy As String = ""
                        If ddlMonthYear.SelectedItem IsNot Nothing Then
                            _mm = CType(ddlMonthYear.SelectedValue.ToString, Integer)
                            _yyyy = Right(ddlMonthYear.SelectedItem.Text.ToString, 4)
                        End If
                        If gcs_service_obj.IsPaySlipAllowedByMonthYear(Session("compCode"), "FORCAST", _mm, _yyyy) Then
                            gcs_service = 1
                        End If

                    Catch ex As Exception

                    End Try
                    If gcs_service = 1 Then
                        Dim dtp As DataTable = InitializeProcessBar(_ds.GetXml(), "I", "FORCAST", DdlreportType.SelectedValue.Replace("'", "").ToUpper, DdlreportType.SelectedItem.Text.Replace("'", ""), _ds.Tables(0).Rows.Count)
                    End If

                    If UsedFor.ToString.ToUpper.Equals("SPDF") Then
                        Dim BPath As String = _objCommon.GetBaseHref(), filepat As String = HttpContext.Current.Server.MapPath(HttpContext.Current.Request.ApplicationPath())
                        var = DdlreportType.SelectedValue & "~" & "P" & "~" & monthvalue & "~" & YearVal & "~" & EmpPassType & "~" & COSTCENTER & "~" _
                       & _strVal.ToString & ".xml~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & LOCATION & "~" & PM1 & "~" &
                        PM1 & "~~" & PM1 & "~" & PM1 & "~" & "1" & "~" & PM14 & "~" & PM13 & "~~1~" & PM14 & "~" & _A & "~" &
                        "~L" & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~SL" & "~" &
                       IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString & "~" & IIf(chkHelp1.Checked = True, "Y", "N").ToString &
                       "~" & IIf(chkMerge.Checked = True, "Y", "N").ToString & "~" & HReimbSts & "~" & BPath.ToString & "~" & filepat.ToString
                    Else
                        var = DdlreportType.SelectedValue & "~" & "P" & "~" & monthvalue & "~" & YearVal & "~" & EmpPassType & "~" & COSTCENTER & "~" _
                       & _strVal.ToString & ".xml~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & LOCATION & "~" & PM1 & "~" &
                        PM1 & "~~" & PM1 & "~" & PM1 & "~" & "1" & "~" & PM14 & "~" & PM13 & "~~1~" & PM14 & "~" & _A & "~" &
                        "~L" & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~SL" & "~" &
                       IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString & "~" & IIf(chkHelp1.Checked = True, "Y", "N").ToString &
                       "~" & IIf(chkMerge.Checked = True, "Y", "N").ToString & "~" & HReimbSts
                    End If


                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If

                    If DdlreportType.SelectedValue.ToString.ToUpper.Equals("57") And UsedFor.ToString.ToUpper.Equals("SPDF") Then
                        HidPreVal.Value = var.ToString
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "ShowTaxForcast();", True)
                    Else
                        PopupScript = ""
                        PopupScript = PopupScript & "window.open('Reports/PreNewTdsEstimationSlip.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                        ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    End If
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("58") Or Convert.ToString(DdlreportType.SelectedValue).Equals("59") Or Convert.ToString(DdlreportType.SelectedValue).Equals("60") Then  'added by ritu malik on 22 may 2018 regarding merge functionality payslip and tax computation IDFC
                    var = "P" & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & COSTCENTER & "~" & LOCATION & "~" & PM1 & "~" & PM1 & "~" & EmpCode & "~" & PM1 & "~" & monthvalue & "~" & YearVal & "~" & PM1 &
                        "~" & PM1 & "~" & PM14 & "~~~~" & PM13 & "~~~S~~~~A~" & reportid & "~" & EmpPassType.ToString & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~" & IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString & "~" & _strVal.ToString & ".xml"
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/pre_SalarySlipTimecard.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("62") Then
                    var = "P~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
              Level.ToString & "~" & COSTCENTER & "~" & LOCATION & "~" & unit.ToString & "~" &
              SalBase.ToString & "~" & EmpCode.ToString & "~" & EmpFName.ToString & "~" & monthvalue _
              & "~" & YearVal & "~" & EmpLName.ToString & "~" & "~" & EmpType.ToString & "~" & Convert.ToString(ddloffcycledt.SelectedValue) & "~" & EmpPassType.ToString & "~" & DdlreportType.SelectedValue

                    'var = "P" & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & COSTCENTER & "~" & LOCATION & "~" & PM1 & "~" & PM1 & "~" & EmpCode & "~" & PM1 & "~" & monthvalue & "~" & YearVal & "~" & PM1 & _
                    '    "~" & PM1 & "~" & PM14 & "~~~~" & PM13 & "~~~S~~~~A~" & reportid & "~" & EmpPassType.ToString & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~" & IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString & "~" & _strVal.ToString & ".xml"
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/Pre_OffCycle.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("63") Then
                    CC = USearch.UCddlcostcenter.ToString
                    Dep = USearch.UCddldept.ToString()
                    Desig = USearch.UCddldesig.ToString()
                    Grad = USearch.UCddlgrade.ToString()
                    Level = USearch.UCddllevel.ToString()
                    Loc = USearch.UCddllocation.ToString()
                    unit = USearch.UCddlunit.ToString()
                    SalBase = USearch.UCddlsalbasis.ToString()
                    EmpCode = USearch.UCTextcode.ToString
                    EmpFName = IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), "")
                    EmpLName = IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), "")
                    EmpType = USearch.UCddlEmp.ToString()

                    var = "P~" & DdlreportType.SelectedValue.ToString & "~" & EmpCode.ToString & "~" & monthvalue.ToString & "~" & YearVal.ToString & "~" & CC.ToString & "~" _
                        & Loc.ToString & "~" & unit.ToString & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" & Level.ToString & "~" & EmpFName.ToString _
                        & "~" & EmpLName.ToString & "~" & SalBase.ToString & "~" & EmpType.ToString & "~" & _strVal.ToString & ".xml"

                    'check session is blank or store value
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/PreSalRegPDF.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("64") Then
                    CC = USearch.UCddlcostcenter.ToString
                    Dep = USearch.UCddldept.ToString()
                    Desig = USearch.UCddldesig.ToString()
                    Grad = USearch.UCddlgrade.ToString()
                    Level = USearch.UCddllevel.ToString()
                    Loc = USearch.UCddllocation.ToString()
                    unit = USearch.UCddlunit.ToString()
                    SalBase = USearch.UCddlsalbasis.ToString()
                    EmpCode = USearch.UCTextcode.ToString
                    EmpFName = IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), "")
                    EmpLName = IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), "")
                    EmpType = USearch.UCddlEmp.ToString()

                    var = "P~" & DdlreportType.SelectedValue.ToString & "~" & EmpCode.ToString & "~" & monthvalue.ToString & "~" & YearVal.ToString & "~" & CC.ToString & "~" _
                    & Loc.ToString & "~" & unit.ToString & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" & Level.ToString & "~" & EmpFName.ToString _
                    & "~" & EmpLName.ToString & "~" & SalBase.ToString & "~" & EmpType.ToString & "~" & _strVal.ToString & ".xml~" & ddlshowsal.SelectedValue.ToString _
                    & "~" & EmpPassType.ToString & "~" & _A & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString

                    'check session is blank or store value
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/PreSalSlipPFA.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("65") Then
                    CC = USearch.UCddlcostcenter.ToString
                    Dep = USearch.UCddldept.ToString()
                    Desig = USearch.UCddldesig.ToString()
                    Grad = USearch.UCddlgrade.ToString()
                    Level = USearch.UCddllevel.ToString()
                    Loc = USearch.UCddllocation.ToString()
                    unit = USearch.UCddlunit.ToString()
                    SalBase = USearch.UCddlsalbasis.ToString()
                    EmpCode = USearch.UCTextcode.ToString
                    EmpFName = IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), "")
                    EmpLName = IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), "")
                    EmpType = USearch.UCddlEmp.ToString()

                    var = "P~" & DdlreportType.SelectedValue.ToString & "~" & EmpCode.ToString & "~" & monthvalue.ToString & "~" & YearVal.ToString & "~" & CC.ToString & "~" _
                    & Loc.ToString & "~" & unit.ToString & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" & Level.ToString & "~" & EmpFName.ToString _
                    & "~" & EmpLName.ToString & "~" & SalBase.ToString & "~" & EmpType.ToString & "~" & _strVal.ToString & ".xml~" & ddlshowsal.SelectedValue.ToString _
                    & "~" & EmpPassType.ToString & "~" & _A & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString

                    'check session is blank or store value
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/PreSalSlipMNF.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("66") Then
                    CC = USearch.UCddlcostcenter.ToString
                    Dep = USearch.UCddldept.ToString()
                    Desig = USearch.UCddldesig.ToString()
                    Grad = USearch.UCddlgrade.ToString()
                    Level = USearch.UCddllevel.ToString()
                    Loc = USearch.UCddllocation.ToString()
                    unit = USearch.UCddlunit.ToString()
                    SalBase = USearch.UCddlsalbasis.ToString()
                    EmpCode = USearch.UCTextcode.ToString
                    EmpFName = IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), "")
                    EmpLName = IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), "")
                    EmpType = USearch.UCddlEmp.ToString()

                    var = "P~" & DdlreportType.SelectedValue.ToString & "~" & EmpCode.ToString & "~" & monthvalue.ToString & "~" & YearVal.ToString & "~" & CC.ToString & "~" _
                    & Loc.ToString & "~" & unit.ToString & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" & Level.ToString & "~" & EmpFName.ToString _
                    & "~" & EmpLName.ToString & "~" & SalBase.ToString & "~" & EmpType.ToString & "~" & _strVal.ToString & ".xml~" & ddlshowsal.SelectedValue.ToString _
                    & "~" & EmpPassType.ToString & "~" & _A & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString

                    'check session is blank or store value
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/PreSalSlipMiddleEast.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf DdlreportType.SelectedValue.ToUpper.Equals("67") Then
                    Grpwise = ""
                    'If HidEmailCCBCC.Value = "UNT" And RblGrpbyPublish.SelectedValue <> "" And ddlRepIn.SelectedValue = "P" Then
                    '    Grpwise = RblGrpbyPublish.SelectedValue.ToString
                    'End If

                    If HidEmailCCBCC.Value = "SP" Then
                        var = "" & "~~~~" & COSTCENTER & "~" & LOCATION & "~~~" & USearch.UCTextcode.ToString & "~~" & monthvalue & "~" + YearVal & "~~" & ddlmultilingual.SelectedValue & "~A~" & PM13 & "~" & "P" & "~" & _A & "~" & EmpPassType.ToString & "~L" & "~" & reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~SP~" & _strVal.ToString & ".xml" & "~" & Grpwise
                    Else
                        var = "" & "~~~~" & COSTCENTER & "~" & LOCATION & "~~~" & USearch.UCTextcode.ToString & "~~" & monthvalue & "~" + YearVal & "~~" & ddlmultilingual.SelectedValue & "~A~" & PM13 & "~" & "P" & "~" & _A & "~" & EmpPassType.ToString & "~L" & "~" & reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~S~" & _strVal.ToString & ".xml" & "~" & Grpwise
                    End If

                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/Pre_SalarySlipInclude.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("68") Then
                    CC = USearch.UCddlcostcenter.ToString
                    Dep = USearch.UCddldept.ToString()
                    Desig = USearch.UCddldesig.ToString()
                    Grad = USearch.UCddlgrade.ToString()
                    Level = USearch.UCddllevel.ToString()
                    Loc = USearch.UCddllocation.ToString()
                    unit = USearch.UCddlunit.ToString()
                    SalBase = USearch.UCddlsalbasis.ToString()
                    EmpCode = USearch.UCTextcode.ToString
                    EmpFName = IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), "")
                    EmpLName = IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), "")
                    EmpType = USearch.UCddlEmp.ToString()

                    var = "P~" & DdlreportType.SelectedValue.ToString & "~" & EmpCode.ToString & "~" & monthvalue.ToString & "~" & YearVal.ToString & "~" & CC.ToString & "~" _
                    & Loc.ToString & "~" & unit.ToString & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" & Level.ToString & "~" & EmpFName.ToString _
                    & "~" & EmpLName.ToString & "~" & SalBase.ToString & "~" & EmpType.ToString & "~" & _strVal.ToString & ".xml~" & ddlshowsal.SelectedValue.ToString _
                    & "~" & EmpPassType.ToString & "~" & _A & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString

                    'check session is blank or store value
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/PreSalSlipTrainee.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("69") Then  'added by ritu malik on 25 jun 2021 regarding tax computation ABFRL
                    Hidforecast.Value = String.Empty
                    Hidforecast.Value = DateTime.Now.ToString("hhmm")
                    var = DdlreportType.SelectedValue.Trim & "~" & HDocStatus + "~~~~" &
                                COSTCENTER & "~~~~~~~~~~" & monthvalue.ToString & "~" _
                                & YearVal.ToString & "~" & EmpPassType & "~" & _strVal.ToString & ".xml~" & IIf(String.IsNullOrEmpty(_A) And String.IsNullOrEmpty(HidEmailCCBCC.Value), "SP", IIf(_A.Equals("S") And HidEmailCCBCC.Value.Equals("SP"), "CC", HidEmailCCBCC.Value)) & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~" & Hidforecast.Value.Trim
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/PreNewTdsEstimationSlip.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                ElseIf DdlreportType.SelectedValue = "76" Or DdlreportType.SelectedValue = "77" Then

                    var = "P" & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & PM1 & "~" & COSTCENTER & "~" & LOCATION & "~" & PM1 & "~" & PM1 & "~~" & PM1 & "~" & monthvalue & "~" & YearVal & "~" & PM1 & "~" & "1" & "~" & PM14 & "~" & PM13 & "~~1~" & PM14 & "~" & _A & "~" & EmpPassType.ToString & "~L" & "~" & reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~SL" & "~" & IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString & "~" & _strVal.ToString & ".xml"

                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/pre_salaryslipAttra.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                ElseIf DdlreportType.SelectedValue.ToUpper.Equals("74") Then
                    If HidEmailCCBCC.Value.ToString.Trim = "SP" Then
                        var = "P~~~~~" + COSTCENTER + "~~~~~~" + monthvalue + "~" + YearVal + "~" + PM1 + "~" + PM1 + "~" + PM1 + "~1~" + PM14 + "~" + PM13 + "~" + _A + "~" + EmpStatus.ToString + "~" + EmpPassType.ToString + "~" + "N" + "~" + "" + "~" + reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~SP~" & _strVal.ToString & ".xml~" & IIf(chkHelp1.Checked = True, "Y", "N").ToString & "~" & HidPdfName.Value.ToString
                    ElseIf HidEmailCCBCC.Value.ToString.Trim.Equals("") And _A.Equals("S") Then
                        var = "P~~~~~" + COSTCENTER + "~~~~~~" + monthvalue + "~" + YearVal + "~" + PM1 + "~" + PM1 + "~" + PM1 + "~1~" + PM14 + "~" + PM13 + "~" + _A + "~" + EmpStatus.ToString + "~" + EmpPassType.ToString + "~" + "N" + "~" + "" + "~" + reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~~" & _strVal.ToString & ".xml~" & IIf(chkHelp1.Checked = True, "Y", "N") & "~" & HidPdfName.Value.ToString
                    Else
                        var = "P~~~~~" + COSTCENTER + "~~~~~~" + monthvalue + "~" + YearVal + "~" + PM1 + "~" + PM1 + "~" + PM1 + "~1~" + PM14 + "~" + PM13 + "~" + _A + "~" + EmpStatus.ToString + "~" + EmpPassType.ToString + "~" + "N" + "~" + "" + "~" + reportid & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~R~" & _strVal.ToString & ".xml~" & IIf(chkHelp1.Checked = True, "Y", "N") & "~" & HidPdfName.Value.ToString
                    End If
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    PopupScript = ""
                    PopupScript = PopupScript & "window.open('Reports/PreSlipArrDetails.aspx?id=" & var & "', '','directories=no,width=600,height=250,top=259,bottom=259, left=212,screenX=212,screenY=212,toolbar=no,scrollbars=no,location=no,resizable =no');"
                    ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", PopupScript, True)
                    PopupScript = ""
                End If
                If _A = "" Then
                    lblMsgSlip.Text = DdlreportType.SelectedItem.Text.ToString & " has been emailed to " + CType(_RecCount, String) + " employees."
                End If
            Else
                If _A.ToString = "S" Then
                    lblMsgSlip.Text = "No payslip published."
                Else
                    lblMsgSlip.Text = "No mail sent."
                End If
            End If
            'For refresh send datetime in Datagrid 'Added by Rohtas Singh on 14 Sep 2009
            If RblNoSearch.SelectedValue <> "P" Then
                If DdlreportType.SelectedValue = "I" Then
                    SendInvestmentDetials()
                Else
                    populateDgPayslip()
                End If
            End If
        End Sub

        Private Function InitializeProcessBar(EmpXml As String, flag As String, Process_Type As String, PaySlipId As String, PaySlipName As String, Optional count As Integer = 0) As DataTable
            Dim gcs_service As Integer = 0
            Dim month As String = ddlMonthYear.SelectedValue.ToString, year As String = Right(ddlMonthYear.SelectedItem.Text.ToString, 4)

            Try
                Dim gcs As New DataTable
                Dim _mm As String = "", _yyyy As String = "", _Month As String = ""
                If ddlMonthYear.SelectedItem IsNot Nothing Then
                    _Month = Left(MonthName(CType(ddlMonthYear.SelectedValue.ToString, Integer)), 3)
                    'Left(MonthName(CType(ddlMonthYear.SelectedValue.ToString, Integer)), 3) & Right(ddlMonthYear.SelectedItem.Text.ToString, 4)
                    _mm = CType(ddlMonthYear.SelectedValue.ToString, Integer)
                    _yyyy = Right(ddlMonthYear.SelectedItem.Text.ToString, 4)
                End If

                If gcs_service_obj.IsPaySlipAllowedByMonthYear(Session("compCode"), Process_Type, _mm, _yyyy) Then
                    gcs_service = 1
                End If

            Catch ex As Exception

            End Try
            'Array.IndexOf(allowed_gcs_service, Process_Type) >= 0
            If gcs_service = 0 Then
                process_status_id.Value = ""
                Dim ArrParam(5) As SqlClient.SqlParameter
                Try
                    ArrParam(0) = New SqlClient.SqlParameter("@flag", flag)
                    ArrParam(1) = New SqlClient.SqlParameter("@Process_Type", Process_Type)
                    ArrParam(2) = New SqlClient.SqlParameter("@UserID", Session("UID"))
                    ArrParam(3) = New SqlClient.SqlParameter("@Emp_Code_XML", EmpXml)
                    ArrParam(4) = New SqlClient.SqlParameter("@PaySlipId", PaySlipId)
                    ArrParam(5) = New SqlClient.SqlParameter("@PaySlipName", PaySlipName)
                    Return _ObjData.GetDataTableProc("PaySP_ReportProcess_ProcessBar", ArrParam)
                Catch ex As Exception
                    _objcommonExp.PublishError("Error in InitializeProcessBar()", ex)
                    Return Nothing
                End Try
            Else
                Try
                    Dim _ArrParam(1) As SqlParameter, dst As New DataTable
                    Dim _array() As String, _AppPath As String = HttpRuntime.AppDomainAppVirtualPath.ToString
                    _ArrParam(0) = New SqlClient.SqlParameter("@flag", "L")
                    _ArrParam(1) = New SqlClient.SqlParameter("@UserID", Session("UID"))

                    dst = _ObjData.GetDataTableProc("PaySP_GetGcsSalarySlip", _ArrParam)
                    If (dst.Rows.Count > 0 AndAlso dst.Rows(0)("IsAbleToStart").ToString = "0") Then
                        Dim slip As String = "Pay Slip is already publishing. Please wait till the completion."
                        If dst.Rows(0)("Process_Type").ToString = "TAXSLIP" Then
                            dst.Rows(0)("Msg") = "Salary Slip With Tax Details is already publishing. Please wait till the completion."
                        ElseIf dst.Rows(0)("Process_Type").ToString = "SLIPWOLVE" Then
                            dst.Rows(0)("Msg") = "Pay Slip Without Leave Details is already publishing. Please wait till the completion."
                        ElseIf dst.Rows(0)("Process_Type").ToString = "FORCAST" Then
                            dst.Rows(0)("Msg") = "Tax sheet-forecast is already publishing. Please wait till the completion."
                        ElseIf dst.Rows(0)("Process_Type").ToString = "SLIPTDSV" Then
                            dst.Rows(0)("Msg") = "TDS Estimation is already publishing. Please wait till the completion."
                        Else
                            dst.Rows(0)("Msg") = slip
                        End If
                        Session("process_status_id") = dst.Rows(0)("id")
                        process_status_id.Value = dst.Rows(0)("id")
                        Return dst
                    End If

                    Dim _dt_temp, temp As New DataTable

                    _dt_temp = _ObjData.ExecSQLQuery("DELETE FROM GcsSalaryProcessStatus WHERE process_user_id = '" & Convert.ToString(HttpContext.Current.Session("UId")) & "' and Process_Type='" & Process_Type & "' " & " and status = 'DONE' and process_user_id = '" & Convert.ToString(HttpContext.Current.Session("UId")) & "' " &
                                                   "Insert Into GcsSalaryProcessStatus(process_user_id, status, total_processed, total_to_process, record_created, Process_Type, mm, YYYY) VALUES('" & Convert.ToString(HttpContext.Current.Session("UId")) & "','START','0', '" & Convert.ToString(count) & "',GetDate(),'" & Process_Type & "', '" & month & "', '" & year & "'); SELECT SCOPE_IDENTITY() AS id;")

                    If _dt_temp.Rows.Count > 0 Then
                        Session("process_status_id") = _dt_temp.Rows(0)("id")
                        process_status_id.Value = _dt_temp.Rows(0)("id")
                    End If
                    _dt_temp.Dispose()
                    Return temp
                Catch ex As Exception
                    _objcommonExp.PublishError("Error in InitializeProcessBar()", ex)
                    Return Nothing
                End Try

            End If

        End Function

        Private Sub SavePwdRec(ByVal _A As String, ByVal EmpCode As String)
            Dim Save_Mail As String = "", ArrParam(7) As SqlClient.SqlParameter, PassType As String = ""
            Try
                If _A.ToString.ToUpper = "S" Then
                    Save_Mail = "S"
                Else
                    Save_Mail = "M"
                End If
                PassType = ddlEmpPass.SelectedValue.ToString
                ArrParam(0) = New SqlClient.SqlParameter("@RptType", DdlreportType.SelectedValue.ToString)
                ArrParam(1) = New SqlClient.SqlParameter("@PassType", PassType.ToString.ToUpper)
                ArrParam(2) = New SqlClient.SqlParameter("@Month", ddlMonthYear.SelectedValue.ToString)
                ArrParam(3) = New SqlClient.SqlParameter("@Year", HidYear.Value.ToString)
                ArrParam(4) = New SqlClient.SqlParameter("@UID", Session("UID"))
                ArrParam(5) = New SqlClient.SqlParameter("@SystemIP", Request.UserHostAddress.ToUpper.ToString)
                ArrParam(6) = New SqlClient.SqlParameter("@SaveMail", Save_Mail.ToString.ToUpper)
                ArrParam(7) = New SqlClient.SqlParameter("@EmpCode", EmpCode.ToString.ToUpper)
                'following sp save those employee in the database whom mail has been sent
                _ObjData.ExecuteStoredProc("PaySP_Ins_TrnSalslipMailPwd", ArrParam)
            Catch ex As Exception
                _objcommonExp.PublishError("SavePwdRec()", ex)
            End Try
        End Sub
        'Custom Mail Body
        Private Function Link_MailBodyDynamic(ByVal EmpCode As String, ByVal EmpName As String) As String
            Try
                Dim _DSET As New DataSet, arrparam(3) As SqlClient.SqlParameter, htmlstr As New System.Text.StringBuilder _
                , AuthType As String = "", CompCode As String = "", apppath As String = "", LinkName As String = "" _
                , EncryptCode As String = "", queryApprove As String = "", EncryptString As New clsEncryptDecrypt

                'Add common function to replace Base href, by praveen on 11 Feb 2013.               
                Dim _s As String = _objCommon.GetBaseHref() & "/"
                AuthType = System.Configuration.ConfigurationManager.AppSettings("AuthenticationType").ToString()

                If AuthType.ToString.ToUpper = "AD" Then
                    apppath = _s & "frmAdlogin.aspx?id="
                ElseIf AuthType.ToString.ToUpper = "DB" Then
                    CompCode = Session("COMPCODE").ToString
                    apppath = _s & CompCode & "/" & "EmpLogin.aspx?id="
                End If

                queryApprove = DdlreportType.SelectedItem.ToString & "~" & EmpCode.ToString & "~" &
                ddlMonthYear.SelectedValue.ToString & "~" & HidYear.Value.ToString

                'Call function for Encrypt code
                EncryptCode = EncryptString.EncryptString(queryApprove, "!#$a54?3")

                apppath = apppath + EncryptCode
                htmlstr = New System.Text.StringBuilder

                'Geeta****************************************************
                Dim contents As String = "", HeaderContents As String = "", NewHeadcontents As String = "" _
                , _arrparam(0) As SqlClient.SqlParameter, _DT As New DataTable, Newcontents As String = "" _
                , value As String = "", parm(1) As SqlClient.SqlParameter, _dst As New DataSet, i As Integer = 0 _
                , Headerval As String = "", FooterContents As String = "", NewFooterContents As String = "" _
                , FooterVal As String = "", monthvalue As String = "", HMonthName As String = ""

                monthvalue = Me.ddlMonthYear.SelectedValue
                HMonthName = Left(MonthName(CType(monthvalue, Integer)), 3)

                _arrparam(0) = New SqlClient.SqlParameter("@RepId", HidRepId.Value.ToString)
                _DT = _ObjData.GetDataTableProc("Paysp_mstmailbody_Getmailbody", _arrparam)

                parm(0) = New SqlClient.SqlParameter("@Fk_emp_code", EmpCode.ToString)
                parm(1) = New SqlClient.SqlParameter("@Fk_cost_center", USearch.UCddlcostcenter.ToString.Trim)
                _dst = _ObjData.GetDsetProc("paysp_mstmailbody_ForGetempDet", parm)

                If _DT.Rows.Count > 0 Then
                    For i = 0 To _dst.Tables(0).Rows.Count - 1
                        HeaderContents = _DT.Rows(0).Item("Header").ToString
                        NewHeadcontents = HeaderContents.Replace("[EMPNAME]", _dst.Tables(0).Rows(i).Item("EmpName").ToString)
                        NewHeadcontents = NewHeadcontents.Replace("[EMPCODE]", _dst.Tables(0).Rows(i).Item("pk_emp_code").ToString)
                        NewHeadcontents = NewHeadcontents.Replace("[MONTH]", HMonthName.ToString)
                        NewHeadcontents = NewHeadcontents.Replace("[YEAR]", HidYear.Value.ToString)
                        If _dst.Tables(1).Rows.Count > 0 Then
                            NewHeadcontents = NewHeadcontents.Replace("[COMPNAME]", _dst.Tables(1).Rows(0).Item("Comp_name").ToString)
                            NewHeadcontents = NewHeadcontents.Replace("[COMPADDRESS]", _dst.Tables(1).Rows(0).Item("Comp_Add").ToString)
                        End If

                        contents = _DT.Rows(0).Item("MailBody").ToString
                        Newcontents = contents.Replace("[EMPNAME]", _dst.Tables(0).Rows(i).Item("EmpName").ToString)
                        Newcontents = Newcontents.Replace("[EMPNAME]", _dst.Tables(0).Rows(i).Item("EmpName").ToString)
                        Newcontents = Newcontents.Replace("[EMPCODE]", _dst.Tables(0).Rows(i).Item("pk_emp_code").ToString)
                        Newcontents = Newcontents.Replace("[MONTH]", HMonthName.ToString)
                        Newcontents = Newcontents.Replace("[YEAR]", HidYear.Value.ToString)
                        If _dst.Tables(1).Rows.Count > 0 Then
                            Newcontents = Newcontents.Replace("[COMPNAME]", _dst.Tables(1).Rows(0).Item("Comp_name").ToString)
                            Newcontents = Newcontents.Replace("[COMPADDRESS]", _dst.Tables(1).Rows(0).Item("Comp_Add").ToString)
                        End If


                        FooterContents = _DT.Rows(0).Item("Footer")
                        NewFooterContents = FooterContents.Replace("[EMPNAME]", _dst.Tables(0).Rows(i).Item("EmpName").ToString)
                        NewFooterContents = NewFooterContents.Replace("[EMPCODE]", _dst.Tables(0).Rows(i).Item("pk_emp_code").ToString)
                        NewFooterContents = NewFooterContents.Replace("[MONTH]", HMonthName.ToString)
                        NewFooterContents = NewFooterContents.Replace("[YEAR]", HidYear.Value.ToString)
                        If _dst.Tables(1).Rows.Count > 0 Then
                            NewFooterContents = NewFooterContents.Replace("[COMPNAME]", _dst.Tables(1).Rows(0).Item("Comp_name").ToString)
                            NewFooterContents = NewFooterContents.Replace("[COMPADDRESS]", _dst.Tables(1).Rows(0).Item("Comp_Add").ToString)
                        End If
                    Next

                    value = Newcontents.ToString
                    Headerval = NewHeadcontents.ToString
                    FooterVal = NewFooterContents.ToString

                    htmlstr.Append("<table align='center' valign='top' id='OuterTable' cellpadding='1' cellspacing='1' width='100%' border='0'>")
                    htmlstr.Append("<tr>")
                    htmlstr.Append("<Td style='font-size: 11px; font-family: Verdana, 'Times New Roman';' colspan='2' valign='top' align='left'>")
                    htmlstr.Append(Headerval.ToString)
                    htmlstr.Append("</Td>")
                    htmlstr.Append("</tr>")

                    htmlstr.Append("<Tr>")
                    htmlstr.Append("<Td colspan='2'>")
                    htmlstr.Append("&nbsp;")
                    htmlstr.Append("</Td>")
                    htmlstr.Append("</Tr>")

                    htmlstr.Append("<tr>")
                    htmlstr.Append("<td style='font-size: 11px; font-family: Verdana, 'Times New Roman';' colspan='2' valign='top' align='left'>")
                    htmlstr.Append(value.ToString)
                    htmlstr.Append("</td>")
                    htmlstr.Append("</tr>")


                    htmlstr.Append("<Tr height='15'>")
                    htmlstr.Append("<Td  style='font-size: 11px; font-family: Verdana, 'Times New Roman';'  colspan='2' valign='top' align='left'>")
                    If DdlreportType.SelectedValue.ToString.ToUpper = "T" Then
                        htmlstr.Append("Click on the following link and enter your username - password and you can view your TDS Estimation Slip.")
                    ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "I" Then
                        htmlstr.Append("Click on the following link and enter your username - password and you can view your Investment Details.")
                    Else
                        htmlstr.Append("Click on the following link and enter your username - password and you can view your Payslip.")
                    End If
                    htmlstr.Append("</Td>")
                    htmlstr.Append("</Tr>")

                    htmlstr.Append("<Tr height='15'>")
                    htmlstr.Append("<Td class='ReportTitle' colspan='2'>")
                    htmlstr.Append("&nbsp;")
                    htmlstr.Append("</Td>")
                    htmlstr.Append("</Tr>")

                    htmlstr.Append("<Tr height='15'>")
                    htmlstr.Append("<Td style='font-size: 11px; font-family: Verdana, 'Times New Roman';'  colspan='2' valign='top' align='left'>")
                    'htmlstr.Append("<a>")
                    If DdlreportType.SelectedValue.ToString.ToUpper = "T" Then
                        htmlstr.Append("<a href='" & apppath.ToString & "'>View TDS Estimation Slip</a>")
                    ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "I" Then
                        htmlstr.Append("<a href='" & apppath.ToString & "'>View Investment Details</a>")
                    Else
                        htmlstr.Append("<a href='" & apppath.ToString & "'>View Payslip </a>")
                    End If

                    htmlstr.Append("</Td>")
                    htmlstr.Append("</Tr>")

                    htmlstr.Append("<Tr>")
                    htmlstr.Append("<Td colspan='2'>")
                    htmlstr.Append("&nbsp;")
                    htmlstr.Append("</Td>")
                    htmlstr.Append("</Tr>")

                    htmlstr.Append("<tr>")
                    htmlstr.Append("<Td style='font-size: 11px; font-family: Verdana, 'Times New Roman';' colspan='2' valign='top' align='left'>")
                    htmlstr.Append(FooterVal.ToString)
                    htmlstr.Append("</Td>")
                    htmlstr.Append("</tr>")
                    htmlstr.Append("</Table>")
                End If
                Return htmlstr.ToString
            Catch ex As Exception
                _objcommonExp.PublishError("Link_MailBody()", ex)
            End Try
            Return ""
        End Function
        Private Function _LinkSendMail() As String
            Dim _msg As String = "", _Counter As Integer = 0, _param(0) As SqlParameter, _DT As New DataTable _
            , MailCount As Integer = 0, MailNotSend As String = "", EmpCode As String = "", EmpCodeName As String = "" _
            , EmpMailId As String = "", arrparam(3) As SqlClient.SqlParameter, MailNotSentCount As String = "" _
            , _MsgSubject As String = "", Dt As New DataTable
            Try
                _param(0) = New SqlParameter("@UID", HidEmpCode.Value)
                _DT = _ObjData.GetDataTableProc("PaySp_Company_Email", _param)

                For _Counter = 0 To DgPayslip.Items.Count - 1
                    If CType(DgPayslip.Items(_Counter).FindControl("chkEmpHold"), CheckBox).Checked = True Then
                        EmpCode = DgPayslip.Items(_Counter).Cells(1).Text.ToString
                        EmpCodeName = DgPayslip.Items(_Counter).Cells(2).Text.ToString & " (" & DgPayslip.Items(_Counter).Cells(1).Text.ToString & ")"
                        EmpMailId = DgPayslip.Items(_Counter).Cells(5).Text.ToString

                        If _DT.Rows.Count > 0 Then
                            If EmpMailId.ToString.Replace("N/A", "") <> "" Then
                                If DdlreportType.SelectedValue.ToString.ToUpper = "T" Then
                                    _MsgSubject = "TDS Estimation Slip"
                                ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "I" Then

                                    _MsgSubject = "Investment Details"
                                Else
                                    _MsgSubject = "Pay Slip"
                                End If

                                If _DT.Rows(0)("CompEmail").ToString.Trim <> "" Then

                                    Dt = _ObjData.ExecSQLQuery("select * from mstmailbody where Fk_rep_id= '" & HidRepId.Value.ToString & "'" & "")

                                    If Dt.Rows.Count > 0 Then
                                        If Dt.Rows(0).Item("FK_rep_id").ToString <> "" Then
                                            If txtccc.Text.ToString.Trim <> "" And txtBCC.Text.ToString.Trim <> "" Then
                                                _msg = _objCommon.SendMailBySMTPNew(System.Configuration.ConfigurationManager.AppSettings("AuthMail").ToString(), EmpMailId.ToString.Trim, "Salary slip " & ddlMonthYear.SelectedItem.ToString, Link_MailBodyDynamic(EmpCode, DgPayslip.Items(_Counter).Cells(10).Text.ToString), txtccc.Text.ToString.Trim, "", txtBCC.Text.ToString.Trim, _DT.Rows(0)("CompEmail").ToString.Trim)
                                            ElseIf txtccc.Text.ToString.Trim <> "" Then
                                                _msg = _objCommon.SendMailBySMTPNew(System.Configuration.ConfigurationManager.AppSettings("AuthMail").ToString(), EmpMailId.ToString.Trim, "Salary slip " & ddlMonthYear.SelectedItem.ToString, Link_MailBodyDynamic(EmpCode, DgPayslip.Items(_Counter).Cells(10).Text.ToString), txtccc.Text.ToString.Trim, "", "", _DT.Rows(0)("CompEmail").ToString.Trim)
                                            ElseIf txtBCC.Text.ToString.Trim <> "" Then
                                                _msg = _objCommon.SendMailBySMTPNew(System.Configuration.ConfigurationManager.AppSettings("AuthMail").ToString(), EmpMailId.ToString.Trim, "Salary slip " & ddlMonthYear.SelectedItem.ToString, Link_MailBodyDynamic(EmpCode, DgPayslip.Items(_Counter).Cells(10).Text.ToString), _DT.Rows(0)("CompEmail").ToString.Trim, "", txtBCC.Text.ToString.Trim, _DT.Rows(0)("CompEmail").ToString.Trim)
                                            Else
                                                _msg = _objCommon.SendMailBySMTPNew(System.Configuration.ConfigurationManager.AppSettings("AuthMail").ToString(), EmpMailId.ToString.Trim, "Salary slip " & ddlMonthYear.SelectedItem.ToString, Link_MailBodyDynamic(EmpCode, DgPayslip.Items(_Counter).Cells(10).Text.ToString), _DT.Rows(0)("CompEmail").ToString.Trim, "", "", _DT.Rows(0)("CompEmail").ToString.Trim)
                                            End If
                                        End If
                                    Else
                                        If txtccc.Text.ToString.Trim <> "" And txtBCC.Text.ToString.Trim <> "" Then
                                            _msg = _objCommon.SendMailBySMTPNew(System.Configuration.ConfigurationManager.AppSettings("AuthMail").ToString(), EmpMailId.ToString.Trim, "Salary slip " & ddlMonthYear.SelectedItem.ToString, Link_MailBody(EmpCode, DgPayslip.Items(_Counter).Cells(10).Text.ToString), txtccc.Text.ToString.Trim, "", txtBCC.Text.ToString.Trim, _DT.Rows(0)("CompEmail").ToString.Trim)
                                        ElseIf txtccc.Text.ToString.Trim <> "" Then
                                            _msg = _objCommon.SendMailBySMTPNew(System.Configuration.ConfigurationManager.AppSettings("AuthMail").ToString(), EmpMailId.ToString.Trim, "Salary slip " & ddlMonthYear.SelectedItem.ToString, Link_MailBody(EmpCode, DgPayslip.Items(_Counter).Cells(10).Text.ToString), txtccc.Text.ToString.Trim, "", "", _DT.Rows(0)("CompEmail").ToString.Trim)
                                        ElseIf txtBCC.Text.ToString.Trim <> "" Then
                                            _msg = _objCommon.SendMailBySMTPNew(System.Configuration.ConfigurationManager.AppSettings("AuthMail").ToString(), EmpMailId.ToString.Trim, "Salary slip " & ddlMonthYear.SelectedItem.ToString, Link_MailBody(EmpCode, DgPayslip.Items(_Counter).Cells(10).Text.ToString), _DT.Rows(0)("CompEmail").ToString.Trim, "", txtBCC.Text.ToString.Trim, _DT.Rows(0)("CompEmail").ToString.Trim)
                                        Else
                                            _msg = _objCommon.SendMailBySMTPNew(System.Configuration.ConfigurationManager.AppSettings("AuthMail").ToString(), EmpMailId.ToString.Trim, "Salary slip " & ddlMonthYear.SelectedItem.ToString, Link_MailBody(EmpCode, DgPayslip.Items(_Counter).Cells(10).Text.ToString), _DT.Rows(0)("CompEmail").ToString.Trim, "", "", _DT.Rows(0)("CompEmail").ToString.Trim)
                                        End If
                                    End If

                                    If _msg.ToString.ToUpper = "MAIL SENT SUCCESSFULLY !" Then
                                        arrparam(0) = New SqlClient.SqlParameter("@Pk_emp_code", EmpCode)
                                        arrparam(1) = New SqlClient.SqlParameter("@MM", ddlMonthYear.SelectedValue.ToString)
                                        arrparam(2) = New SqlClient.SqlParameter("@YY", HidYear.Value.ToString)
                                        arrparam(3) = New SqlClient.SqlParameter("@ReportTtype", DdlreportType.SelectedValue.ToString)
                                        'following sp save those employee in the database whom mail has been sent
                                        _ObjData.GetDataSetProc("Sp_Ins_TrnDocMailStatus", arrparam)
                                        MailCount = MailCount + 1
                                    Else
                                        If MailNotSend.ToString = "" Then
                                            MailNotSend = EmpCode.ToString
                                        Else
                                            MailNotSend = MailNotSend + ", " + EmpCode.ToString
                                        End If
                                    End If
                                End If
                            Else
                                If MailNotSentCount.ToString = "" Then
                                    MailNotSentCount = EmpCode.ToString
                                Else
                                    MailNotSentCount = MailNotSentCount + ", " + EmpCode.ToString
                                End If
                            End If
                        End If
                    End If
                    EmpCode = ""
                    EmpCodeName = ""
                    _msg = ""
                Next

                If MailCount <> 0 Then
                    lblMsgSlip.Text = CType(MailCount, String) + " email(s) have been sent."
                End If

                If MailNotSend.ToString <> "" Then
                    If MailNotSentCount.ToString = "" Then
                        MailNotSentCount = MailNotSend
                    Else
                        MailNotSentCount = MailNotSentCount + ", " + MailNotSend
                    End If
                End If

                If MailNotSentCount.ToString <> "" Then
                    _objCommon.ShowMessage("M", lblMailMsg, "Mail not sent for the following employee code(s): " + MailNotSentCount.ToString, True)
                Else
                    lblMailMsg.Text = ""
                End If

                BtnSend.CssClass = "btn"

                populateDgPayslip()

            Catch ex As Exception
                _objcommonExp.PublishError("_LinkSendMail()", ex)
            End Try
            Return _msg
        End Function
        'Default Mail Body
        Private Function Link_MailBody(ByVal EmpCode As String, ByVal EmpName As String) As String
            Try
                Dim _DSET As New DataSet, arrparam(3) As SqlClient.SqlParameter, htmlstr As New System.Text.StringBuilder _
                , AuthType As String = "", CompCode As String = "", apppath As String = "" _
                , EncryptCode As String = "", queryApprove As String = "", EncryptString As New clsEncryptDecrypt
                'Add common function to replace base Href, Added by praveen verma on 11 Feb 2013.
                Dim _s As String = _objCommon.GetBaseHref() & "/"
                AuthType = System.Configuration.ConfigurationManager.AppSettings("AuthenticationType").ToString()

                If AuthType.ToString.ToUpper = "AD" Then
                    apppath = _s & "frmAdlogin.aspx?id="
                ElseIf AuthType.ToString.ToUpper = "DB" Then
                    CompCode = Session("COMPCODE").ToString
                    apppath = _s & CompCode & "/" & "EmpLogin.aspx?id="
                End If

                queryApprove = DdlreportType.SelectedItem.ToString & "~" & EmpCode.ToString & "~" &
                ddlMonthYear.SelectedValue.ToString & "~" & HidYear.Value.ToString & "~" & USearch.UCddlcostcenter.ToString()

                'Call function for Encrypt code
                EncryptCode = EncryptString.EncryptString(queryApprove, "!#$a54?3")

                apppath = apppath + EncryptCode
                htmlstr = New System.Text.StringBuilder
                ''Tblmain
                htmlstr.Append("<Table cellSpacing='0' id='tblmain' cellPadding='0' border='0' width='96%'>")
                htmlstr.Append("<Tr>")
                htmlstr.Append("<Td class='ReportFieldCaptionForSide' valign='top'>")

                htmlstr.Append("<Table cellSpacing='0' id='tblEmp' cellPadding='0' border='0' width='95%' align='right'>")

                htmlstr.Append("<Tr height='35' colspan='2'>")
                htmlstr.Append("<Td style='font-weight: bold; font-size: 11px; font-family: Verdana, 'Times New Roman';' colspan='2' valign='middle' align='left' >")
                htmlstr.Append("Greeting of the Day!")
                htmlstr.Append("</Td>")
                htmlstr.Append("</Tr>")

                htmlstr.Append("<Tr height='15'>")
                htmlstr.Append("<Td class='ReportTitle' colspan='2'>")
                htmlstr.Append("&nbsp;")
                htmlstr.Append("</Td>")
                htmlstr.Append("</Tr>")

                htmlstr.Append("<Tr colspan='2'>")
                htmlstr.Append("<Td style='font-size: 11px; font-family: Verdana, 'Times New Roman';' colspan='2' valign='top' align='left' >")
                htmlstr.Append("Dear ")
                htmlstr.Append(EmpName)
                htmlstr.Append(" ,")
                htmlstr.Append("</Td>")
                htmlstr.Append("</Tr>")

                htmlstr.Append("<Tr height='15'>")
                htmlstr.Append("<Td class='ReportTitle' colspan='2'>")
                htmlstr.Append("&nbsp;")
                htmlstr.Append("</Td>")
                htmlstr.Append("</Tr>")

                htmlstr.Append("<Tr height='15'>")
                htmlstr.Append("<Td style='font-size: 11px; font-family: Verdana, 'Times New Roman';'  colspan='2' valign='top' align='left'>")
                If DdlreportType.SelectedValue.ToString.ToUpper = "T" Then
                    htmlstr.Append("Click on the following link and enter your username - password and you can view your TDS Estimation Slip.")
                ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "I" Then
                    htmlstr.Append("Click on the following link and enter your username - password and you can view your Investment Details.")
                Else
                    htmlstr.Append("Click on the following link and enter your username - password and you can view your Payslip.")
                End If
                htmlstr.Append("</Td>")
                htmlstr.Append("</Tr>")

                htmlstr.Append("<Tr height='15'>")
                htmlstr.Append("<Td class='ReportTitle' colspan='2'>")
                htmlstr.Append("&nbsp;")
                htmlstr.Append("</Td>")
                htmlstr.Append("</Tr>")

                htmlstr.Append("<Tr height='15'>")
                htmlstr.Append("<Td style='font-size: 11px; font-family: Verdana, 'Times New Roman';'  colspan='2' valign='top' align='left'>")
                'htmlstr.Append("<a>")
                If DdlreportType.SelectedValue.ToString.ToUpper = "T" Then
                    htmlstr.Append("<a href='" & apppath.ToString & "'>View TDS Estimation Slip</a>")
                ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "I" Then
                    htmlstr.Append("<a href='" & apppath.ToString & "'>View Investment Details</a>")
                Else
                    htmlstr.Append("<a href='" & apppath.ToString & "'>View Payslip </a>")
                End If

                htmlstr.Append("</Td>")
                htmlstr.Append("</Tr>")

                htmlstr.Append("<Tr height='15'>")
                htmlstr.Append("<Td class='ReportTitle' colspan='2'>")
                htmlstr.Append("&nbsp;")
                htmlstr.Append("</Td>")
                htmlstr.Append("</Tr>")

                htmlstr.Append("<Tr height='15'>")
                htmlstr.Append("<Td style='font-size: 11px; font-family: Verdana, 'Times New Roman';'  colspan='2' valign='top' align='left'>")
                htmlstr.Append("Thanks & Regards")
                htmlstr.Append("</Td>")
                htmlstr.Append("</Tr>")

                htmlstr.Append("<Tr height='15'>")
                htmlstr.Append("<Td class='ReportTitle' colspan='2'>")
                htmlstr.Append("&nbsp;")
                htmlstr.Append("</Td>")
                htmlstr.Append("</Tr>")

                htmlstr.Append("<Tr height='15'>")
                htmlstr.Append("<Td style='font-size: 11px; font-family: Verdana, 'Times New Roman';'  colspan='2' valign='top' align='left'>")
                htmlstr.Append("Payroll Help Desk")
                htmlstr.Append("</Td>")
                htmlstr.Append("</Tr>")
                htmlstr.Append("</Table>")
                htmlstr.Append("</Td>")
                htmlstr.Append("</Tr>")
                htmlstr.Append("</Table>") ''End TblMain
                Return htmlstr.ToString
            Catch ex As Exception
                _objcommonExp.PublishError("Link_MailBody()", ex)
            End Try
            Return ""
        End Function
        Protected Sub DgPayslip_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DgPayslip.ItemCommand
            litJava.Text = ""
            If UCase(e.CommandName) = UCase("Preview") Then
                'Rohtas(25Feb) Start
                Dim monthvalue As String = "", YearVal As String = "", HDocStatus As String = "P", HEmpCode As String = "", HMonthName As String = "", HYearName As String = "",
                HReimbSts As String = "N", ReportType As String = "V", RepType As String = "A", flag As String = "S",
                Empcode As String = CType(DgPayslip.DataKeys(e.Item.ItemIndex), String), PM1 As String = "", PM11 As String = "0", PM13 As String = "A",
                PM14 As String = "true", var As String, EmpStatus As String = USearch.UCddlEmp.ToString, Loan As String = "", Advance As String = "",
                _strVal As String = Guid.NewGuid.ToString, StaffId As String = "", LeaveDisp As String = "", EmpPassType As String = "",
                COSTCENTER As String = USearch.UCddlcostcenter.ToString, Repid As String = "", LOCATION As String = USearch.UCddllocation.ToString, PayCode As String = "",
                Dep As String, Desig As String, Grad As String, Level As String, Unit As String, SalBase As String, EmpFName As String, EmpLName As String, EmpType As String
                'for assign value to variable
                monthvalue = Me.ddlMonthYear.SelectedValue
                HMonthName = monthvalue
                YearVal = HidYear.Value.ToString
                HYearName = YearVal
                HEmpCode = Empcode
                PM13 = ddlshowsal.SelectedValue.ToString
                EmpPassType = ddlEmpPass.SelectedValue.ToString
                Repid = HidRepId.Value.ToString
                HidEmailCCBCC.Value = ""

                Dep = USearch.UCddldept.ToString()
                Desig = USearch.UCddldesig.ToString()
                Grad = USearch.UCddlgrade.ToString()
                Level = USearch.UCddllevel.ToString()
                Unit = USearch.UCddlunit.ToString()
                SalBase = USearch.UCddlsalbasis.ToString()
                EmpFName = IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), "")
                EmpLName = IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), "")
                EmpType = USearch.UCddlEmp.ToString()
                'Hold = ddlshowsal.SelectedValue.ToString

                If DdlreportType.SelectedValue.ToUpper = "SL" Then
                    'Added by geeta (& "~~~~" & IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString) on 11 sep 2012
                    var = "" & "~~~~~" & COSTCENTER & "~" & LOCATION & "~~~" & Empcode & "~~" & monthvalue & "~" & YearVal & "~~" & "1" & "~" & PM14 & "~" & PM13 & "~1~1~" & PM14 & "~S~" & EmpPassType.ToString & "~L" & "~" & Repid & "~~~~" & IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue.ToUpper = "S" Then

                    var = "" & "~~~~" & COSTCENTER & "~" & LOCATION & "~~~" & Empcode & "~~" + monthvalue & "~" & YearVal & "~~" & "1" & "~" & PM14 & "~" & PM13 & "~" & "~" & "S" & "~" & EmpPassType.ToString & "~L" & "~" & Repid

                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    'To display "TDS Estimation Slip"
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue.ToUpper = "T" Then
                    var = "~" + monthvalue + "~" + YearVal + "~" + Empcode + "~" + Repid + "~" + EmpPassType + "~" + COSTCENTER + "~" + ReportType + "~" + flag + "~" + HReimbSts + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~T"
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue = "R" Or DdlreportType.SelectedValue = "52" Or DdlreportType.SelectedValue = "49" Or DdlreportType.SelectedValue = "51" Or DdlreportType.SelectedValue.Equals("74") Then
                    var = "~~~~~" + COSTCENTER + "~~~~" + Empcode + "~~" + monthvalue + "~" + YearVal + "~~~~1~true" + "~" + PM13 + "~S~" + EmpStatus.ToString + "~" + EmpPassType.ToString + "~" + "N" + "~" + "" + "~" & Repid & "~~~~~" & If(chkHelp1.Checked = True, "Y", "N").ToString


                    'this is used for convert Session method
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue = "TL" Then
                    'var = "~~~~~" + COSTCENTER + "~~~~" + Empcode + "~~" + monthvalue + "~" + YearVal + "~~~~1~true" + "~" + PM13 + "~S~" + EmpStatus.ToString + "~" + EmpPassType.ToString + "~" + "N" + "~" + "" + "~" & Repid
                    var = "~~~~~~~~~" + Empcode + "~~" + monthvalue + "~" + YearVal + "~~~~1~true" + "~" + PM13 + "~S~" + EmpStatus.ToString + "~" + EmpPassType.ToString + "~" + Repid
                    'this is used for convert Session method
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    'Rohtas(21Jan)
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue = "I" Then
                    var = "C~" + Empcode.ToString
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                    'This is used to open in PDF "Salary Slip with Investment"
                ElseIf DdlreportType.SelectedValue.ToUpper = "SI" Then
                    Dim dt As DataTable = Nothing, LoopVar As Integer = 0, InvCode As String = ""
                    dt = _ObjData.GetDataTableProc("sel_sections")
                    For LoopVar = 0 To dt.Rows.Count - 1
                        InvCode = InvCode + "'" + dt.Rows(LoopVar).Item("sectionnumber").ToString + "',"
                    Next

                    var = "" + "~~~~" + COSTCENTER.ToString + "~~~~" + Empcode.ToString + "~~" + monthvalue + "~" + YearVal + "~~0~1~1" + "~" + PM13 + "~~~" + "~0~" + InvCode.ToString + "~" + EmpPassType.ToString + "~" + EmpStatus.ToString + "~" + Repid

                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                    '==============================================================================================
                    'By: Dharmendra Rawat [24 Jun 2010]
                    'This code is used for send "Pay Slip in hindi With Arrear Details
                ElseIf DdlreportType.SelectedValue.ToUpper = "SH" Then
                    var = "" + Empcode + "~" + "F" + "~~~~~~~~~~" + "A" + "~" + monthvalue + "~" + YearVal + "~" + "A" + "~~" + EmpPassType.ToString + "~~~"
                    'this is used for convert Session method
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue.ToUpper = "RS" Then
                    'Comment by Niraj kumar on 05-jun-2013
                    'var = "~~~~~~~~" + "S" + "~" + Empcode + "~~" + monthvalue + "~" + YearVal + "~~~~~~~~~~" + "A" + "~" + "S" + "~" + EmpPassType.ToString
                    'Add by Niraj on 05-jun-2013 
                    var = "~~~~~" + COSTCENTER.ToString + "~~~" + "S" + "~" + Empcode + "~~" + monthvalue + "~" + YearVal + "~~~~~~~~~~" + "A" + "~" + "S" + "~" + EmpPassType.ToString + "~" + txtccc.Text.ToString + "~" + txtBCC.Text.ToString + "~RS" + "~~" & "~~~~" & "~" & ddlshowsal.SelectedValue.ToString
                    'Added by Geeta on 19 Nov 2010 to send parameter in session
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)

                    'ADDED BY JAY ON 27 JAN 12 FOR THE YEAR TO DATE SALARY SLIP
                ElseIf DdlreportType.SelectedValue.ToUpper = "YTD" Then
                    var = "~~~~" + COSTCENTER + "~~~~" + Empcode + "~~" + monthvalue + "~" + YearVal + "~~" + PM13 + "~" + "A" + "~" + Repid + "~" + "" + "~~~~"
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)

                ElseIf DdlreportType.SelectedValue.ToUpper = "43" Then
                    var = "~~~~" + COSTCENTER.ToString + "~~~~" + Empcode.ToString + "~~" + monthvalue.ToString + "~" + YearVal.ToString + "~~A~" + DDLPaySlipType.SelectedValue + "~H~0"
                    'var = "~~~~" + COSTCENTER + "~~~~" + Empcode + "~~" + monthvalue + "~" + YearVal + "~~" + PM13 + "~" + "A" + "~" + Repid + "~" + "" + "~~~~"
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue.ToUpper = "RN" Then
                    'Comment by Niraj on 05-jun-2013
                    'var = "P~~~~~" + COSTCENTER + "~~~" + "S" + "~" + Empcode + "~~" + monthvalue + "~" + YearVal + "~~~~~1~~A~~~" + "A" + "~" + "" + "~" + "3" + "~" + "R" + "~" + "H1" & "~" & EmpPassType.ToString & "~" & "PE"
                    'Add by Niraj on 05-jun-2013
                    var = "P~~~~~" + COSTCENTER + "~~~" + "S" + "~" + Empcode + "~~" + monthvalue + "~" + YearVal + "~~~~~1~~A~~~" + "A" + "~" + "" + "~" + "3" + "~" + "R" + "~" + "H1" & "~" & EmpPassType.ToString & "~" & "PE" & "~" & txtccc.Text.ToString & "~" & txtBCC.Text.ToString & "~RN" & "~" & "~" & "~" & ddlshowsal.SelectedValue.ToString
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue = "50" Then

                    var = "" & "~~~~~" & COSTCENTER & "~" & LOCATION & "~~~" & Empcode & "~~" & monthvalue & "~" & YearVal & "~~" & "1" & "~" & PM14 & "~" & PM13 & "~1~1~" & PM14 & "~S~" & EmpPassType.ToString & "~L" & "~" & Repid & "~~~~" & IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                    'PTC
                ElseIf DdlreportType.SelectedValue.ToUpper = "53" Then
                    var = "~~~~" + COSTCENTER + "~~~~" + Empcode + "~~" + monthvalue + "~" + YearVal + "~~" + PM13 + "~" + "A" + "~" + Repid + "~" + "" + "~~~~~~~~~"
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue.ToUpper = "55" Then
                    var = DdlreportType.SelectedValue.Trim & "~H1~" & Empcode.ToString & "~" & EmpFName.ToString & "~" & EmpLName.ToString & "~" &
                            COSTCENTER.ToString & "~" & Dep.ToString & "~" & Grad.ToString & "~" & Desig.ToString & "~" &
                           LOCATION.ToString & "~" & Unit.ToString & "~" & SalBase.ToString & "~" &
                            Level.ToString & "~" & EmpType.ToString & "~" & Session("ugroup").ToString & "~" & monthvalue.ToString & "~" _
                            & YearVal.ToString
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf Convert.ToString(DdlreportType.SelectedValue).ToUpper.Equals("57") Then
                    var = DdlreportType.SelectedValue.Trim & "~H1~" & Empcode.ToString & "~" & EmpFName.ToString & "~" & EmpLName.ToString & "~" &
                            COSTCENTER.ToString & "~" & Dep.ToString & "~" & Grad.ToString & "~" & Desig.ToString & "~" &
                           LOCATION.ToString & "~" & Unit.ToString & "~" & SalBase.ToString & "~" &
                            Level.ToString & "~" & EmpType.ToString & "~" & Session("ugroup").ToString & "~" & monthvalue.ToString & "~" _
                            & YearVal.ToString
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf Convert.ToString(DdlreportType.SelectedValue).ToUpper.Equals("56") Then
                    var = "" & "~~~~~" & COSTCENTER & "~" & LOCATION & "~~~" & Empcode & "~~" & monthvalue & "~" & YearVal & "~~" & "1" & "~" & PM14 & "~" & PM13 & "~1~1~" & PM14 & "~S~" & EmpPassType.ToString & "~L" & "~" & Repid & "~~~~" & IIf(chkemailrepmanager.Checked = True, "Y", "N").ToString
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue.Equals("59") Or DdlreportType.SelectedValue.Equals("58") Or DdlreportType.SelectedValue.Equals("60") Then
                    var = "K~~~~~" & COSTCENTER & "~" & LOCATION & "~~~" & Empcode.ToString & "~" & EmpFName.ToString & "~" &
                                    monthvalue & "~" & YearVal & "~" & EmpLName.ToString & "~~~~~~" & PM13 & "~~~~S~~~A~" & Repid

                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue.Equals("62") Then
                    var = "H~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" &
                Level.ToString & "~" & COSTCENTER & "~" & LOCATION & "~" & Unit.ToString & "~" &
                SalBase.ToString & "~" & Empcode.ToString & "~" & EmpFName.ToString & "~" & monthvalue _
                & "~" & YearVal & "~" & EmpLName.ToString & "~" & "~" & EmpType.ToString & "~" & Convert.ToString(ddloffcycledt.SelectedValue) & "~0~" & Repid
                    'check session is blank or store value
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue.Equals("63") Then
                    var = "H~" & Repid & "~" & Empcode.ToString & "~" & monthvalue & "~" & YearVal & "~" & COSTCENTER & "~" & LOCATION & "~" & Unit.ToString _
                        & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" & Level.ToString & "~" & EmpFName.ToString & "~" & EmpLName.ToString _
                        & "~" & SalBase.ToString & "~" & EmpType.ToString & "~~"
                    'check session is blank or store value
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue.Equals("64") Then
                    var = "H~" & Repid & "~" & Empcode.ToString & "~" & monthvalue & "~" & YearVal & "~" & COSTCENTER & "~" & LOCATION & "~" & Unit.ToString _
                        & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" & Level.ToString & "~" & EmpFName.ToString & "~" & EmpLName.ToString _
                        & "~" & SalBase.ToString & "~" & EmpType.ToString & "~~" & ddlshowsal.SelectedValue.ToString & "~"
                    'check session is blank or store value
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue.Equals("65") Then
                    var = "H~" & Repid & "~" & Empcode.ToString & "~" & monthvalue & "~" & YearVal & "~" & COSTCENTER & "~" & LOCATION & "~" & Unit.ToString _
                        & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" & Level.ToString & "~" & EmpFName.ToString & "~" & EmpLName.ToString _
                        & "~" & SalBase.ToString & "~" & EmpType.ToString & "~~" & ddlshowsal.SelectedValue.ToString & "~"
                    'check session is blank or store value
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue.Equals("66") Then
                    var = "H~" & Repid & "~" & Empcode.ToString & "~" & monthvalue & "~" & YearVal & "~" & COSTCENTER & "~" & LOCATION & "~" & Unit.ToString _
                        & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" & Level.ToString & "~" & EmpFName.ToString & "~" & EmpLName.ToString _
                        & "~" & SalBase.ToString & "~" & EmpType.ToString & "~~" & ddlshowsal.SelectedValue.ToString & "~"
                    'check session is blank or store value
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue.ToUpper.Equals("67") Then

                    var = "" & "~~~~" & COSTCENTER & "~" & LOCATION & "~~~" & Empcode & "~~" + monthvalue & "~" & YearVal & "~~" & ddlmultilingual.SelectedValue & "~A~" & PM13 & "~" & "~" & "S" & "~~L" & "~" & Repid

                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    'To display "TDS Estimation Slip"
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf DdlreportType.SelectedValue.Equals("68") Then
                    var = "H~" & Repid & "~" & Empcode.ToString & "~" & monthvalue & "~" & YearVal & "~" & COSTCENTER & "~" & LOCATION & "~" & Unit.ToString _
                        & "~" & Dep.ToString & "~" & Desig.ToString & "~" & Grad.ToString & "~" & Level.ToString & "~" & EmpFName.ToString & "~" & EmpLName.ToString _
                        & "~" & SalBase.ToString & "~" & EmpType.ToString & "~~" & ddlshowsal.SelectedValue.ToString & "~"
                    'check session is blank or store value
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)
                ElseIf Convert.ToString(DdlreportType.SelectedValue).ToUpper.Equals("69") Then
                    Hidforecast.Value = String.Empty
                    Hidforecast.Value = DateTime.Now.ToString("hhmm")

                    var = DdlreportType.SelectedValue.Trim & "~H1~" & Empcode.ToString & "~" & EmpFName.ToString & "~" & EmpLName.ToString & "~" &
                            COSTCENTER.ToString & "~" & Dep.ToString & "~" & Grad.ToString & "~" & Desig.ToString & "~" &
                           LOCATION.ToString & "~" & Unit.ToString & "~" & SalBase.ToString & "~" &
                            Level.ToString & "~" & EmpType.ToString & "~" & Session("ugroup").ToString & "~" & monthvalue.ToString & "~" _
                            & YearVal.ToString & "~~~~~~" & Hidforecast.Value.Trim
                    If Not Session(_strVal) Is Nothing Then
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    Else
                        Session.Remove(_strVal)
                        Session(_strVal) = var.ToString
                        var = _strVal.ToString
                    End If
                    hidstring1.Value = var.ToString
                    Hidden5.Value = DdlreportType.SelectedValue.ToUpper
                    ScriptManager.RegisterStartupScript(Me, Me.[GetType](), "test", "showReport(); ", True)

                End If

            End If
        End Sub
        Protected Sub chkmailformat_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkmailformat.CheckedChanged
            If chkmailformat.Checked = True And DdlreportType.SelectedValue.ToString <> "" And (DdlreportType.SelectedValue.ToString <> "43" Or DdlreportType.SelectedValue.ToString <> "49") Then
                trselall.Style.Value = "display:"
                tblmail.Style.Value = "display:"
                Btndelete.Style.Value = "display:"

                _PopulateDataMail()
            Else
                tblmail.Style.Add("display", "none")
                'tblmail.Style.Value = "display:none"
                Btndelete.Style.Value = "display:none"
            End If
        End Sub
        Private Sub _PopulateDataMail()
            Dim _Dt As New DataTable, Arrparam(0) As SqlClient.SqlParameter
            Arrparam(0) = New SqlClient.SqlParameter("@Fk_Rep_Id", DdlreportType.SelectedValue.ToString)
            _Dt = _ObjData.GetDataTableProc("Paysp_mstmailbody_Sel", Arrparam)
            If _Dt.Rows.Count > 0 Then
                txtheader.Text = _Dt.Rows(0).Item("Header").ToString
                EasyWebMAilBody.HTMLValue = _Dt.Rows(0).Item("MailBody").ToString
                Textfooter.Text = _Dt.Rows(0).Item("Footer").ToString
            End If
        End Sub
        Protected Sub lnkPDF_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LnkPDF.Click
            Try
                Dim Month As String, Year As String, _strJava As New StringBuilder, filepath As String, FileName() As String = Nothing, FName As String = ""
                lblMailMsg.Text = ""
                lblmsg.Text = ""
                Year = _objCommon.nNz(Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                Month = Left(MonthName(CType(ddlMonthYear.SelectedValue, Integer)), 3)
                If Convert.ToString(DdlreportType.SelectedValue).Equals("55") Then
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\TaxComputationSheet"
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("57") Then
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\YTDTaxComputationSheet"
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("62") Then
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\OffCyclePaySlip"
                ElseIf DdlreportType.SelectedValue.ToString = "S" Then
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\LeaveWoPaySlip"
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("69") Then
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\ForecastTaxComputationSheet_" & Hidforecast.Value & "\"
                ElseIf DdlreportType.SelectedValue.ToString <> "43" And DdlreportType.SelectedValue.ToString <> "49" And DdlreportType.SelectedValue.ToString <> "63" _
                     And DdlreportType.SelectedValue.ToString <> "64" And DdlreportType.SelectedValue.ToString <> "65" And DdlreportType.SelectedValue.ToString <> "66" _
                     And DdlreportType.SelectedValue.ToString <> "67" And DdlreportType.SelectedValue.ToString <> "68" Then
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("49") Then
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("63") Then
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\SalRegPDF\"
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("64") Then
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\PaySlipPFA\"
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("65") Then
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\PaySlipMNF\"
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("66") Then
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\PaySlipMiddleEast\"
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("67") Then
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\multilingPaySlip\"
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("68") Then
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\SalarySlipTrainee\"
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("76") Or Convert.ToString(DdlreportType.SelectedValue).Equals("77") Then ' Added by Deepankar Raizada - For Attra
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\SalarySlipAttra\"
                ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("74") Then
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\TaxPaySlipArrDetails\"
                Else
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\WageSlipFormXIX"
                End If
                If DdlreportType.SelectedValue.ToString.ToUpper = "RN" Then
                    filepath = ""
                    filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles" & "\NewReimbPDFFiles\" & "Publish" & Month & Year & "\ReimbursementSlip.Zip"
                    If System.IO.File.Exists(filepath) Then
                        Response.ContentType = "application/x-pdf"
                        Response.AddHeader("Content-Disposition", "attachment; filename=ReimbursementSlip.Zip")
                        Response.WriteFile(filepath)
                        Response.End()
                    Else
                        lblMailMsg.Text = "Please first generate report in PDF format in selected month Or <br>Could not publish PDF slip because reimbursement is not process !"
                        lblMailMsg.CssClass = "UserMessage"
                        LnkPDF.Style.Value = "display:none"
                        Exit Sub
                    End If
                Else
                    If Directory.Exists(filepath) Then
                        If DdlreportType.SelectedValue = "R" Then
                            If HidPdfName.Value.ToString <> "" Then
                                FName = "_" & HidPdfName.Value & ".pdf"
                            Else
                                FName = "_SalarySlipwithTaxDetails.pdf"
                            End If
                        ElseIf DdlreportType.SelectedValue = "52" Then
                            FName = "_SalarySlipwithTaxDetails_Bangladesh.pdf"
                        ElseIf DdlreportType.SelectedValue = "51" Then
                            FName = "_SalarySlipwithTaxDetails_Yum.pdf"
                        ElseIf DdlreportType.SelectedValue = "49" Then
                            FName = "_SalSlipWithTaxDetailsMisc.pdf"
                        ElseIf DdlreportType.SelectedValue = "TL" Then
                            FName = "_SalarySlipwithTaxLoanDetails.pdf"
                        ElseIf DdlreportType.SelectedValue.ToUpper = "SL" Then
                            FName = "_SalarySlip.pdf"
                        ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "56" Then
                            FName = "_PaySlipwithReimbDetails.pdf"
                        ElseIf DdlreportType.SelectedValue.ToUpper = "S" Then
                            FName = "_SalarySlipInclude.pdf"
                        ElseIf DdlreportType.SelectedValue.ToUpper = "T" Then
                            FName = "_TDSEstimationSlip.pdf"
                        ElseIf DdlreportType.SelectedValue.ToUpper = "SI" Then
                            FName = "_Payslip.pdf"
                        ElseIf DdlreportType.SelectedValue.ToUpper = "SH" Then
                            FName = "_SalarySlipInHindiWithArrearDetails.pdf"
                        ElseIf DdlreportType.SelectedValue.ToUpper = "RS" Then
                            FName = "_ReimbSlip.pdf"
                        ElseIf DdlreportType.SelectedValue.ToUpper = "YTD" Then
                            FName = "_YTDSalSlip.pdf"
                        ElseIf DdlreportType.SelectedValue.ToUpper = "43" Then
                            FName = "_WageSlipFormXIX.pdf"
                        ElseIf DdlreportType.SelectedValue.ToUpper = "RN" Then
                            FName = "_ReimbSlipNew.pdf"
                        ElseIf DdlreportType.SelectedValue = "50" Then
                            FName = "_SalSlipOthDetails.pdf"
                            'PTC
                        ElseIf DdlreportType.SelectedValue = "53" Then
                            FName = "_SalarySlipPTC.pdf"
                        ElseIf DdlreportType.SelectedValue = "55" And chkMerge.Checked = False Then
                            If Len(ddlMonthYear.SelectedValue.ToString) > 1 Then
                                FName = "_estax_" & Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4) & ddlMonthYear.SelectedValue.ToString & ".pdf"

                            Else
                                FName = "_estax_" & Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4) & "0" & ddlMonthYear.SelectedValue.ToString & ".pdf"
                            End If
                            'Rohtas Singh for Merge two PDF Slip on 14 Feb 2018
                        ElseIf DdlreportType.SelectedValue = "55" And chkMerge.Checked = True Then
                            'Add payslip name according to client requirement on 06 Apr 2018
                            Dim FlName As String = ""
                            FlName = "_ITCS"
                            FName = "_Payslip" & FlName.ToString & ".pdf"
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("57") Then  'And chkMerge.Checked.Equals(False)
                            If Len(ddlMonthYear.SelectedValue.ToString) > 1 Then
                                FName = "_estax_" & Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4) & ddlMonthYear.SelectedValue.ToString & ".pdf"

                            Else
                                FName = "_estax_" & Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4) & "0" & ddlMonthYear.SelectedValue.ToString & ".pdf"
                            End If
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("58") Then
                            filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\PaySlipTimeCard"
                            FName = "_SalSlip_" & ddlMonthYear.SelectedValue.ToString & "_" & Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4) & ".pdf"
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("59") Then
                            filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\PaySlipStipend"
                            FName = "_SalSlip_" & ddlMonthYear.SelectedValue.ToString & "_" & Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4) & ".pdf"
                            'added by Geeta : Marathi payslip("60")
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("60") Then
                            filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year & "\PaySlipMarathi"
                            'FName = "_SalSlip_" & ddlMonthYear.SelectedValue.ToString & "_" & Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4) & ".pdf"
                            FName = "_PaySlipMarathi.pdf"
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("62") Then
                            FName = "_offcyclePaySlip.pdf"
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("63") Then
                            FName = "SalaryRegister_" & USearch.UCddlunit.ToString() & "_" & USearch.UCddldesig.ToString() & ".pdf"
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("64") Then
                            FName = "_SalarySlipPFA.pdf"
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("65") Then
                            FName = "_SalarySlipMNF.pdf"
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("66") Then
                            FName = "_SalarySlipMiddleEast.pdf"
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("67") Then
                            FName = "_MultiLingualSalarySlip.pdf"
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("68") Then
                            FName = "_SalarySlipTrainee.pdf"
                            'Added by Deepankar Raizada - For Attra
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("76") Or Convert.ToString(DdlreportType.SelectedValue).Equals("77") Then
                            FName = "_SalarySlipAttra.pdf"
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("69") Then
                            FName = "_ForecastTaxComputationSheet.pdf"
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("74") Then
                            FName = "_TaxPaySlipArrDetails.pdf"
                        End If
                        If Convert.ToString(DdlreportType.SelectedValue).Equals("63") Then
                            HidEmpPdf.Value = FName.ToString
                        Else
                            HidEmpPdf.Value = HidEmpPdf.Value.Replace(",", FName & ",")
                            If HidEmpPdf.Value.ToString.Trim <> "" Then
                                HidEmpPdf.Value = Left(HidEmpPdf.Value.ToString.Trim, Len(HidEmpPdf.Value.ToString.Trim) - 1)
                            End If
                        End If


                        FileName = Split(HidEmpPdf.Value.ToString.Trim, ",")
                        AddZipFiles(filepath, FileName)
                        Response.Clear()
                        Response.BufferOutput = False
                        ' for large files...
                        Dim c As System.Web.HttpContext = System.Web.HttpContext.Current
                        'Dim ReadmeText As [String] = "Hello!" & vbLf & vbLf & "This is a README..." & DateTime.Now.ToString("G")
                        Dim archiveName As String = [String].Format("Slips-{0}.zip", DateTime.Now.ToString("dd-MMM-yyyy"))
                        If Convert.ToString(DdlreportType.SelectedValue).Equals("59") Then
                            archiveName = [String].Format("PaySlipStipend-{0}.zip", DateTime.Now.ToString("dd-MMM-yyyy"))
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("58") Then
                            archiveName = [String].Format("PaySlipTimeCard-{0}.zip", DateTime.Now.ToString("dd-MMM-yyyy"))
                            'added by Geeta : Marathi payslip("60")
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("60") Then
                            archiveName = [String].Format("PaySlipMarathi-{0}.zip", DateTime.Now.ToString("dd-MMM-yyyy"))
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("62") Then
                            archiveName = [String].Format("OffCyclePaySlips-{0}.zip", DateTime.Now.ToString("dd-MMM-yyyy"))
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("63") Then
                            archiveName = [String].Format("SalaryRegister-{0}.zip", DateTime.Now.ToString("dd-MMM-yyyy"))
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("64") Then
                            archiveName = [String].Format("SalarySlipPFA-{0}.zip", DateTime.Now.ToString("dd-MMM-yyyy"))
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("65") Then
                            archiveName = [String].Format("SalarySlipMNF-{0}.zip", DateTime.Now.ToString("dd-MMM-yyyy"))
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("66") Then
                            archiveName = [String].Format("SalarySlipMiddleEast-{0}.zip", DateTime.Now.ToString("dd-MMM-yyyy"))
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("67") Then
                            archiveName = [String].Format("MultiLingualSalarySlip-{0}.zip", DateTime.Now.ToString("dd-MMM-yyyy"))
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("68") Then
                            archiveName = [String].Format("SalarySlipTrainee-{0}.zip", DateTime.Now.ToString("dd-MMM-yyyy"))
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("69") Then
                            archiveName = [String].Format("ForecastTaxComputationSheet-{0}.zip", DateTime.Now.ToString("dd-MMM-yyyy"))
                            ' Added By Deepankar Raizada - For Attra
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("76") Or Convert.ToString(DdlreportType.SelectedValue).Equals("77") Then
                            archiveName = [String].Format("SalarySlipAttra-{0}.zip", DateTime.Now.ToString("dd-MMM-yyyy"))
                        ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("74") Then
                            archiveName = [String].Format("TaxPaySlipArrDetails-{0}.zip", DateTime.Now.ToString("dd-MMM-yyyy"))
                        End If

                        Response.ContentType = "application/zip"
                        Response.AddHeader("content-disposition", "filename=" & archiveName)
                        Using zip As New ZipFile()
                            If Convert.ToString(DdlreportType.SelectedValue).Equals("59") Then
                                zip.AddFiles(filesToInclude, "PaySlipStipend")
                            ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("58") Then
                                zip.AddFiles(filesToInclude, "PaySlipTimeCard")
                                'added by Geeta : Marathi payslip("60")
                            ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("60") Then
                                zip.AddFiles(filesToInclude, "PaySlipMarathi")
                            ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("62") Then
                                zip.AddFiles(filesToInclude, "OffCyclePaySlips")
                            ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("63") Then
                                zip.AddFiles(filesToInclude, "SalaryRegister")
                            ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("64") Then
                                zip.AddFiles(filesToInclude, "SalarySlipPFA")
                            ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("65") Then
                                zip.AddFiles(filesToInclude, "SalarySlipMNF")
                            ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("66") Then
                                zip.AddFiles(filesToInclude, "SalarySlipMiddleEast")
                            ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("67") Then
                                zip.AddFiles(filesToInclude, "MultiLingualSalarySlip")
                            ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("68") Then
                                zip.AddFiles(filesToInclude, "SalarySlipTrainee")
                            ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("69") Then
                                zip.AddFiles(filesToInclude, "ForecastTaxComputationSheet")
                            ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("76") Or Convert.ToString(DdlreportType.SelectedValue).Equals("77") Then
                                zip.AddFiles(filesToInclude, "SalarySlipAttra") ' Added By Deepankar Raizada - For Attra
                            ElseIf Convert.ToString(DdlreportType.SelectedValue).Equals("74") Then
                                zip.AddFiles(filesToInclude, "TaxPaySlipArrDetails")
                            Else
                                zip.AddFiles(filesToInclude, "PaySlips")
                            End If
                            zip.Save(Response.OutputStream)
                        End Using

                        Response.End()

                        'Response.Close()
                    Else
                        lblMailMsg.Text = "Please first generate report in PDF format in selected month !"
                        lblMailMsg.CssClass = "UserMessage"
                        LnkPDF.Style.Value = "display:none"
                        'Added by Rajarshi on 2 Nov 2017
                        rbtnmail.Checked = False
                        rbtnslip.Checked = True
                    End If
                End If
            Catch ex As Exception
                lblMailMsg.Text = "Message Published On Convert Into PDF " & ex.Message.ToString
                lblMailMsg.CssClass = "ErrorMessage"
                LnkPDF.Style.Value = "display:none"
            End Try
        End Sub
        Private Sub AddZipFiles(ByVal targetDirectory As String, ByVal FName() As String)
            Dim fileEntries As String() = Directory.GetFiles(targetDirectory), subdirectoryEntries As String() = Directory.GetDirectories(targetDirectory)
            For Each fileName As String In fileEntries
                If FName.Contains(System.IO.Path.GetFileName(fileName)) Then
                    filesToInclude.Add(System.IO.Path.Combine(fileName, lblmsg.Text))
                End If
            Next
            For Each subdirectory As String In subdirectoryEntries
                AddZipFiles(subdirectory, FName)
            Next
        End Sub
        Protected Sub BtnPreviewdivActive_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnPreviewdivActive.Click
            Response.Redirect("FrmReimbCoverLetter.aspx", False)
        End Sub
        Protected Sub BtnSendCCBCC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnSendCCBCC.Click
            lblMailMsg.Text = ""
            lblMsgSlip.Text = ""
            lblProcessBarMsg.Text = ""
            LnkPDF.Style.Value = "display:none"
            HidEmailCCBCC.Value = "SP"
            'Excel Process locking validation checking
            'CheckExcelProcessbarAlreadyProcessing()
            'If (lblProcessStatusExcel.Text <> "") Then
            '    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = lblProcessStatusExcel.Text})
            '    _objCommon.ShowMessage(_msg)
            '    Exit Sub
            'End If
            If txtccc.Text.ToString.Replace("&nbsp;", "") = "" And txtBCC.Text.ToString.Replace("&nbsp;", "") = "" Then
                _objCommon.ShowMessage("M", lblMailMsg, "Please enter at least one email id in CC or BCC", True)
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "closepopup", "CloseSlipProgressbar();", True)
                Exit Sub
            End If
            'Added by Rohtas singh on 14 Dec 2017 for save email contant.
            If chkmailformat.Checked = True Then
                _SaveMailBody()
            End If

            SendReportPDF("S")
        End Sub
        Protected Sub btnreset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnreset.Click
            'Excel Process locking validation checking
            CheckExcelProcessbarAlreadyProcessing()
            trrepformat.Style.Value = "display:none"
            tableshow.Style.Value = "display:none"
            RblNoSearch.SelectedValue = "P"
            TdSearch1.Style("display") = ""
            TdSearch2.Style("display") = ""
            'Added by Rohtas Singh on 06 Dec 2017
            trRptformat.Style.Value = "display:none"
            'Added by Quadir on 14 OCT 2020
            TrSlipPubMode.Style.Value = "display:none"

            RblGrpbyPublish.SelectedValue = ""
            Dim _dt As New DataTable
            _dt.Columns.Add(New DataColumn("fk_emp_code"))
            _dt.Columns.Add(New DataColumn("EmpName"))
            _dt.Columns.Add(New DataColumn("Dept_desc"))
            _dt.Columns.Add(New DataColumn("desig_desc"))
            _dt.Columns.Add(New DataColumn("Email"))
            _dt.Columns.Add(New DataColumn("Status"))
            _dt.Columns.Add(New DataColumn("Sent_Date"))
            _dt.Columns.Add(New DataColumn("Flag"))
            _dt.Columns.Add(New DataColumn("SalHold"))
            _dt.Columns.Add(New DataColumn("EMailExist"))
            _dt.Columns.Add(New DataColumn("EMailSend"))
            _dt.Columns.Add(New DataColumn("EmpLastName"))
            DgPayslip.DataSource = _dt
            DgPayslip.DataBind()

            'Reset by Rohtas Singh on 28 Dec 2017
            EasyWebMAilBody.HTMLValue = ""
            txtheader.Text = ""
            Textfooter.Text = ""
            'Rohtas Singh for Merge two PDF Slip on 14 Feb 2018
            chkMerge.Checked = False
            btnPreview.Enabled = True
            ddlSalWithHeld.SelectedValue = "N"
        End Sub
        'Added by Nisha on 04 Sep 2013
        Private Sub ExporttoExcelforFinalSalary()
            Dim j As Integer = 0, PAYCODE As String = "", arrparam(17) As SqlClient.SqlParameter, dst As New DataSet,
            filename As String = "", _sw As StreamWriter, complexID As Guid = Guid.NewGuid()
            Try
                arrparam(0) = New SqlClient.SqlParameter("@month", _objCommon.nNz(ddlMonthYear.SelectedValue.ToString))
                arrparam(1) = New SqlClient.SqlParameter("@year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                arrparam(2) = New SqlClient.SqlParameter("@Dep", USearch.UCddldept.ToString())
                arrparam(3) = New SqlClient.SqlParameter("@Desig", USearch.UCddldesig.ToString())
                arrparam(4) = New SqlClient.SqlParameter("@Grad", USearch.UCddlgrade.ToString())
                arrparam(5) = New SqlClient.SqlParameter("@Lable", USearch.UCddllevel.ToString())
                arrparam(6) = New SqlClient.SqlParameter("@CC", USearch.UCddlcostcenter.ToString())
                arrparam(7) = New SqlClient.SqlParameter("@Loc", USearch.UCddllocation.ToString())
                arrparam(8) = New SqlClient.SqlParameter("@unit", USearch.UCddlunit.ToString())
                arrparam(9) = New SqlClient.SqlParameter("@SalBase", USearch.UCddlsalbasis.ToString())
                arrparam(10) = New SqlClient.SqlParameter("@EmpCode", USearch.UCTextcode.ToString)
                arrparam(11) = New SqlClient.SqlParameter("@EmpFName", IIf(UCase(USearch.UCrbtfirst.ToString()) = "F", USearch.UCTextname.ToString(), ""))
                arrparam(12) = New SqlClient.SqlParameter("@EmpLName", IIf(UCase(USearch.UCrbtlast.ToString()) = "L", USearch.UCTextname.ToString(), ""))
                arrparam(13) = New SqlParameter("@Hold", ddlshowsal.SelectedValue.ToString)
                arrparam(14) = New SqlParameter("@userGroup", Session("UGroup").ToString)
                arrparam(15) = New SqlParameter("@EmpType", USearch.UCddlEmp.ToString())
                arrparam(16) = New SqlClient.SqlParameter("@RepId", DDLPaySlipType.SelectedValue.ToString)
                arrparam(17) = New SqlClient.SqlParameter("@userid", Session("uid").ToString)
                'Change SP Name by Nisha on 21 Jan 2014
                dst = _ObjData.GetDataSetProc("Paysp_Rpt_Sel_SalaryRegisterUpdated", arrparam)

            Catch ex As Exception
                _objcommonExp.PublishError("ExporttoExcelforFinalSalary()", ex)
            End Try

            If dst.Tables(3).Rows.Count > 0 Then
                filename = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\Final_Salary.xls"
                If System.IO.File.Exists(filename) Then
                    System.IO.File.Delete(filename)
                End If
                _sw = New StreamWriter(filename)
                ExportToExcelXMLFinalSalary(dst, _sw)
                _sw.Close()
                _sw.Dispose()
                Response.Clear()
                Response.BufferOutput = False
                Response.ContentType = "application/zip"
                Response.AddHeader("content-disposition", "filename=Salary_Register.zip")
                Using zip As New ZipFile()
                    zip.AddFile(filename, "Final_Salary")
                    zip.Save(Response.OutputStream)
                End Using

                If File.Exists(filename) Then
                    File.Delete(filename)
                End If

                lblmsg.Text = ""

                Response.End()
                'Response.Close()
            Else
                lblmsg2.Text = "No Record Found!"
                lblmsg2.CssClass = "ErrorMessage"
            End If
        End Sub
        'Added by Nisha on 04 Sep 2013
        Private Sub ExportToExcelXMLFinalSalary(ByVal source As DataSet, ByRef _ExcelDoc As StreamWriter, Optional ByVal flag As String = "")
            Dim _RowCount As Integer = 0, headstr As String = "", K As Integer = 0, Stlname As String = "", start_ExcelXML As String = _ExcelDoc.ToString(),
            sheetCount As Integer = 1, TTotalGross, TIT, TPF, TESI, TDec, TNetSal, TPaidDays, TGrossRate, TPT As Decimal, Row1(), Row2(), Row3() As DataRow, GTotalDec() As DataRow,
            DTab As DataTable = source.Tables(3), TotBasicRate As Decimal = 0, GPCA As Integer, GTotal() As DataRow, VarTotal1 As Decimal = 0, VarTotal As Decimal = 0

            Const end_ExcelXML As String = "</Workbook>"
            Try
                Dim _ExcelXML As New StringBuilder()
                _ExcelDoc.Write("<?xml version=""1.0""?><?mso-application progid=""Excel.Sheet""?>" & Chr(13) & "" & Chr(10) & "<Workbook ")
                _ExcelDoc.Write("xmlns=""urn:schemas-microsoft-com:office:spreadsheet""" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write(" xmlns:o=""urn:schemas-microsoft-com:office:office""" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write(" xmlns:x=""urn:schemas-microsoft-com:office:excel""" & Chr(13) & "" & Chr(10))
                _ExcelDoc.Write(" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">" & Chr(13) & "" & Chr(10) & " <Styles>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""Default"" ss:Name=""Normal"">" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>" & Chr(13) & "" & Chr(10) & " <Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders>")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Font/>" & Chr(13) & "" & Chr(10) & " <Interior/>" & Chr(13) & "" & Chr(10) & " <NumberFormat/>")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Protection/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("" & Chr(13) & "" & Chr(10) & " <Style ss:ID=""BC"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelDoc.Write("x:Family=""Swiss"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""SL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat")
                _ExcelDoc.Write(" ss:Format=""@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _ExcelDoc.Write("ss:ID=""DC"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""0.0000""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelDoc.Write("<Style ss:ID=""IT"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""0""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")
                _ExcelDoc.Write("ss:ID=""DL"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelDoc.Write("ss:Format=""mm/dd/yyyy;@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")

                _ExcelDoc.Write("<Style ss:ID=""s21""><Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""2""/><Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders><Font x:Family=""Verdana"" ss:Bold=""1"" ss:Size=""10""/></Style>")

                _ExcelDoc.Write("<Style ss:ID=""s2"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Top"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:Size=""9"" ss:Color=""#808080"" ss:Underline=""Single""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#FFFFFF"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelDoc.Write("<Style ss:ID=""s32"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Vertical=""Center""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" ss:Family=""Swiss"" ss:Size=""9"" ss:Color=""#000000"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#CCFFFF"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                'FOR EMPLOYEE DETAILS COLUMN ONLY
                _ExcelDoc.Write("<Style ss:ID=""ED"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000"" ss:Bold=""0""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#CCFFFF"" ss:Pattern=""Solid""/></Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelDoc.Write("<Style ss:ID=""ED1"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000"" ss:Bold=""0""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior/></Style>" & Chr(13) & "" & Chr(10) & "")

                'FOR EXCEL HEARDER ONLY
                _ExcelDoc.Write("<Style ss:ID=""C2"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Color=""#000000"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#FF8080"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                'FOR COLUMN ONLY
                _ExcelDoc.Write("<Style ss:ID=""C3"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Center"" ss:Vertical=""Bottom""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelDoc.Write("ss:FontName=""Verdana"" ss:Bold=""1"" ss:Color=""black""  x:Family=""Swiss"" ss:Size=""10"" />" & Chr(13) & "" & Chr(10) & " <Interior ss:Color=""Silver"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & " ")

                _ExcelDoc.Write("<Style ss:ID=""CompHearder"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""0""/>" & Chr(13) & "" & Chr(10) & "<Borders></Borders>")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8.5"" ss:Color=""#000000"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelDoc.Write("<Style ss:ID=""Total"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#FBFBFB"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelDoc.Write("<Interior ss:Color=""#447B60"" ss:Pattern=""Solid""/></Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelDoc.Write("</Styles>" & Chr(13) & "" & Chr(10) & " ")

                _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                _ExcelDoc.Write("<Table><Column ss:Width=""40""/><Column ss:Width=""100""/><Column ss:Width=""180""/><Column ss:Width=""180""/><Column ss:Width=""180""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/><Column ss:Width=""110""/>")

                Dim colspan As Integer = 0, Count As Integer = 0
                colspan = 16
                'For _RecordCounter = 0 To source.Tables(2).Rows.Count - 1
                '    If source.Tables(2).Rows(_RecordCounter)("RateGross").ToString = "Y" Then
                '        Count = Count + 1
                '    End If
                'Next
                'colspan = colspan + Count
                'Count = 0

                For _RecordCounter = 0 To source.Tables(2).Rows.Count - 1
                    Count = Count + 1
                Next
                colspan = colspan + Count
                Count = 0

                For _RecordCounter = 0 To source.Tables(5).Rows.Count - 1
                    Count = Count + 1
                Next
                colspan = colspan + Count
                colspan = colspan - 1

                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell ss:StyleID=""CompHearder"" ss:MergeAcross=""" & colspan & """><Data ss:Type=""String"">")
                _ExcelDoc.Write(source.Tables(0).Rows(0)("Comp_Name").ToString)
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell ss:StyleID=""CompHearder"" ss:MergeAcross=""" & colspan & """><Data ss:Type=""String"">")
                _ExcelDoc.Write(source.Tables(0).Rows(0)("Address").ToString)
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell ss:StyleID=""CompHearder"" ss:MergeAcross=""" & colspan & """><Data ss:Type=""String"">")
                _ExcelDoc.Write("Location : " & source.Tables(0).Rows(0)("Location").ToString)
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell ss:StyleID=""CompHearder"" ss:MergeAcross=""" & colspan & """><Data ss:Type=""String"">")
                _ExcelDoc.Write("Unit : " & source.Tables(0).Rows(0)("Unit").ToString)
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell ss:StyleID=""CompHearder"" ss:MergeAcross=""" & colspan & """><Data ss:Type=""String"">")
                _ExcelDoc.Write("No. of Employees : " & source.Tables(6).Rows.Count.ToString)
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell ss:StyleID=""CompHearder"" ss:MergeAcross=""" & colspan & """><Data ss:Type=""String"">")
                _ExcelDoc.Write("Payable Month : " + source.Tables(0).Rows(0)("Mon").ToString)
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell ss:StyleID=""CompHearder"" ss:MergeAcross=""" & colspan & """><Data ss:Type=""String"">")
                _ExcelDoc.Write("Salary Period : " + source.Tables(0).Rows(0)("SalaryPeriod").ToString)
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell ss:MergeAcross=""" & colspan & """><Data ss:Type=""String"">")
                _ExcelDoc.Write("")
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                If source.Tables(3).Rows.Count > 0 Then
                    _ExcelDoc.Write("<Row>")

                    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("S.No.")
                    _ExcelDoc.Write("</Data></Cell>")

                    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("Employee Id")
                    _ExcelDoc.Write("</Data></Cell>")

                    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("Employee Name")
                    _ExcelDoc.Write("</Data></Cell>")

                    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(_objCommon.DisplayCaption("DES"))
                    _ExcelDoc.Write("</Data></Cell>")

                    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(_objCommon.DisplayCaption("GRD"))
                    _ExcelDoc.Write("</Data></Cell>")

                    'Added by Nisha on 25 Sep 2013
                    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("Basic Rate")
                    _ExcelDoc.Write("</Data></Cell>")

                    'For _RecordCounter = 0 To source.Tables(2).Rows.Count - 1
                    '    If source.Tables(2).Rows(_RecordCounter)("RateGross").ToString = "Y" Then
                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write(source.Tables(2).Rows(_RecordCounter)("PayHead").ToString & " Rate")
                    '        _ExcelDoc.Write("</Data></Cell>")
                    '    End If
                    'Next

                    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("Monthly Gross Rate")
                    _ExcelDoc.Write("</Data></Cell>")

                    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("Salary Days")
                    _ExcelDoc.Write("</Data></Cell>")

                    For _RecordCounter = 0 To source.Tables(2).Rows.Count - 1
                        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(source.Tables(2).Rows(_RecordCounter)("PayHead").ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next

                    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("Total Gross Salary")
                    _ExcelDoc.Write("</Data></Cell>")

                    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("PF")
                    _ExcelDoc.Write("</Data></Cell>")

                    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("ESI")
                    _ExcelDoc.Write("</Data></Cell>")

                    'Added by Nisha on 06 Jan 2014
                    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("PT")
                    _ExcelDoc.Write("</Data></Cell>")

                    For _RecordCounter = 0 To source.Tables(5).Rows.Count - 1
                        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(source.Tables(5).Rows(_RecordCounter)("PayHead").ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Next

                    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("TDS")
                    _ExcelDoc.Write("</Data></Cell>")

                    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("Total Deduction")
                    _ExcelDoc.Write("</Data></Cell>")

                    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("Net Payable")
                    _ExcelDoc.Write("</Data></Cell>")

                    _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("Mode of Payment")
                    _ExcelDoc.Write("</Data></Cell>")

                    _ExcelDoc.Write("</Row>")

                    TTotalGross = 0
                    TIT = 0
                    TPF = 0
                    TESI = 0
                    TDec = 0
                    TNetSal = 0
                    TPaidDays = 0
                    TGrossRate = 0
                    TPT = 0

                    'Add TPT to calculate total PT by Nisha on 06 Jan 2014
                    'loop for calculate gross,grosstotal,IT,PF,ESI,VOLPF,TotDed,NetSal,ProfTax,Arrears,ESIEmr and PaidDays
                    For FixTotal = 0 To source.Tables(9).Rows.Count - 1
                        TTotalGross = TTotalGross + CType(_objCommon.Nz(source.Tables(9).Rows(FixTotal)("GrossTotal"), 0), Decimal)
                        TIT = TIT + CType(_objCommon.Nz(source.Tables(9).Rows(FixTotal)("IT"), 0), Decimal)
                        TPF = TPF + CType(_objCommon.Nz(source.Tables(9).Rows(FixTotal)("PF"), 0), Decimal)
                        TESI = TESI + CType(_objCommon.Nz(source.Tables(9).Rows(FixTotal)("ESI"), 0), Decimal)
                        TDec = TDec + CType(_objCommon.Nz(source.Tables(9).Rows(FixTotal)("TotDed"), 0), Decimal)
                        TNetSal = TNetSal + CType(_objCommon.Nz(source.Tables(9).Rows(FixTotal)("NetSal"), 0), Decimal)
                        TPaidDays = TPaidDays + CType(_objCommon.Nz(source.Tables(9).Rows(FixTotal)("PaidDays"), 2), Decimal)
                        TGrossRate = TGrossRate + CType(_objCommon.Nz(source.Tables(9).Rows(FixTotal)("RateGross"), 2), Decimal)
                        TPT = TPT + CType(_objCommon.Nz(source.Tables(9).Rows(FixTotal)("ProfTax"), 2), Decimal)
                    Next

                    'For Each x As DataRow In DTab.Rows
                    '    _RowCount += 1
                    '    If _RowCount = 63000 Then
                    '        _RowCount = 0
                    '        sheetCount += 1
                    '        _ExcelDoc.Write("</Table>")
                    '        _ExcelDoc.Write(" </Worksheet>")
                    '        _ExcelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                    '        _ExcelDoc.Write("<Table>")
                    '        _ExcelDoc.Write("<Row>")
                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write("S.No.")
                    '        _ExcelDoc.Write("</Data></Cell>")

                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write("Employee Id")
                    '        _ExcelDoc.Write("</Data></Cell>")

                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write("Employee Name")
                    '        _ExcelDoc.Write("</Data></Cell>")

                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write(_objCommon.DisplayCaption("DES"))
                    '        _ExcelDoc.Write("</Data></Cell>")

                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write(_objCommon.DisplayCaption("GRD"))
                    '        _ExcelDoc.Write("</Data></Cell>")

                    '        'Added by Nisha on 25 Sep 2013
                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write("Basic Rate")
                    '        _ExcelDoc.Write("</Data></Cell>")

                    '        'For _RecordCounter = 0 To source.Tables(2).Rows.Count - 1
                    '        '    If source.Tables(2).Rows(_RecordCounter)("RateGross").ToString = "Y" Then
                    '        '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        '        _ExcelDoc.Write(source.Tables(2).Rows(_RecordCounter)("PayHead").ToString & " Rate")
                    '        '        _ExcelDoc.Write("</Data></Cell>")
                    '        '    End If
                    '        'Next

                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write("Monthly Gross Rate")
                    '        _ExcelDoc.Write("</Data></Cell>")

                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write("Salary Days")
                    '        _ExcelDoc.Write("</Data></Cell>")

                    '        For _RecordCounter = 0 To source.Tables(2).Rows.Count - 1
                    '            _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '            _ExcelDoc.Write(source.Tables(2).Rows(_RecordCounter)("PayHead").ToString)
                    '            _ExcelDoc.Write("</Data></Cell>")
                    '        Next

                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write("Total Gross Salary")
                    '        _ExcelDoc.Write("</Data></Cell>")

                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write("PF")
                    '        _ExcelDoc.Write("</Data></Cell>")

                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write("ESI")
                    '        _ExcelDoc.Write("</Data></Cell>")

                    '        'Added by Nisha on 06 Jan 2014
                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write("PT")
                    '        _ExcelDoc.Write("</Data></Cell>")

                    '        For _RecordCounter = 0 To source.Tables(5).Rows.Count - 1
                    '            _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '            _ExcelDoc.Write(source.Tables(5).Rows(_RecordCounter)("PayHead").ToString)
                    '            _ExcelDoc.Write("</Data></Cell>")
                    '        Next

                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write("TDS")
                    '        _ExcelDoc.Write("</Data></Cell>")

                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write("Total Deduction")
                    '        _ExcelDoc.Write("</Data></Cell>")

                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write("Net Payable")
                    '        _ExcelDoc.Write("</Data></Cell>")

                    '        _ExcelDoc.Write("<Cell ss:StyleID=""C3""><Data ss:Type=""String"">")
                    '        _ExcelDoc.Write("Mode of Payment")
                    '        _ExcelDoc.Write("</Data></Cell>")
                    '        _ExcelDoc.Write("</Row>")
                    '    End If
                    'Next

                    For _counter1 As Integer = 0 To DTab.Rows.Count - 1
                        _ExcelDoc.Write("<Row>")

                        _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                        _ExcelDoc.Write((_counter1 + 1).ToString)
                        _ExcelDoc.Write("</Data></Cell>")

                        'Row to display the records related to the Employee
                        Row1 = GetRecRow(DTab.Rows(_counter1)("pk_emp_code").ToString, source.Tables(6))

                        _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(DTab.Rows(_counter1)("pk_emp_code").ToString)
                        _ExcelDoc.Write("</Data></Cell>")

                        _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(DTab.Rows(_counter1)("empname").ToString)
                        _ExcelDoc.Write("</Data></Cell>")

                        _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(DTab.Rows(_counter1)("desig_desc").ToString)
                        _ExcelDoc.Write("</Data></Cell>")

                        _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(DTab.Rows(_counter1)("grade_desc").ToString)
                        _ExcelDoc.Write("</Data></Cell>")

                        Dim cnt1 As Integer = 0
                        Row3 = GetRecRow(DTab.Rows(_counter1)("pk_emp_code").ToString, source.Tables(6))
                        'Loop for count the earning of paycodes
                        For _RecordCounter = 0 To source.Tables(2).Rows.Count - 1
                            'Change condition for Basic Rate by Nisha on 25 Sep 2013
                            If source.Tables(2).Rows(_RecordCounter).Item("Payhead").ToString.ToUpper = "BASIC" Then
                                cnt1 = 1
                                Row2 = GetRecRowPay(DTab.Rows(_counter1)("pk_emp_code").ToString, source.Tables(2).Rows(_RecordCounter).Item("Paycode").ToString, source.Tables(4))
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")

                                If Row2.Length > 0 Then
                                    'Check for no of Records related to paycode Earnings
                                    If Row2(0)("Actual_Basic").ToString <> "" Then
                                        _ExcelDoc.Write(Row2(0)("Actual_Basic").ToString)
                                    Else
                                        _ExcelDoc.Write("0")
                                    End If
                                Else
                                    _ExcelDoc.Write("0")
                                End If
                                _ExcelDoc.Write("</Data></Cell>")
                            End If
                        Next
                        If cnt1 = 0 Then
                            _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                            _ExcelDoc.Write("0")
                            _ExcelDoc.Write("</Data></Cell>")
                        End If

                        'Row for show the gross or total values
                        Row3 = GetRecRow(DTab.Rows(_counter1)("pk_emp_code").ToString, source.Tables(6))
                        'Loop for count the earning of paycodes
                        _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                        If Row3.Length > 0 Then
                            'Check for no of Records related to paycode Earnings
                            If Row3(0)("RateGross").ToString <> "" Then
                                _ExcelDoc.Write(Row3(0)("RateGross").ToString)
                            Else
                                _ExcelDoc.Write("0")
                            End If
                        Else
                            _ExcelDoc.Write("0")
                        End If
                        _ExcelDoc.Write("</Data></Cell>")

                        If Row1.Length > 0 Then
                            _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                            _ExcelDoc.Write(Row1(0)("PaidDays").ToString)
                            _ExcelDoc.Write("</Data></Cell>")
                        End If

                        'Row for show the gross or total values
                        Row3 = GetRecRow(DTab.Rows(_counter1)("pk_emp_code").ToString, source.Tables(6))
                        'Loop for count the earning of paycodes
                        For _RecordCounter = 0 To source.Tables(2).Rows.Count - 1
                            'for featch the record according employee code in data row
                            Row2 = GetRecRowPay(DTab.Rows(_counter1)("pk_emp_code").ToString, source.Tables(2).Rows(_RecordCounter).Item("Paycode").ToString, source.Tables(4))
                            _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                            'for check the existance of data 
                            If Row2.Length > 0 Then
                                'Check for no of Records related to paycode Earnings
                                If Row2(0)("Earnings").ToString <> "" Then
                                    _ExcelDoc.Write(Row2(0)("Earnings").ToString)
                                Else
                                    _ExcelDoc.Write("0")
                                End If
                            Else
                                _ExcelDoc.Write("0")
                            End If
                            _ExcelDoc.Write("</Data></Cell>")
                        Next

                        'to print the gross total value
                        If source.Tables(13).Rows(0)("ZeroColumns").ToString.ToUpper = "Y" Then
                            If CType(source.Tables(14).Rows(0)("GrossTotal"), Integer) <> 0 Then
                                If Row3.Length > 0 Then
                                    _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                    _ExcelDoc.Write(Row3(0)("GrossTotal").ToString)
                                    _ExcelDoc.Write("</Data></Cell>")
                                Else
                                    _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                    _ExcelDoc.Write("0")
                                    _ExcelDoc.Write("</Data></Cell>")
                                End If
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write("0")
                                _ExcelDoc.Write("</Data></Cell>")
                            End If
                        Else
                            If Row3.Length > 0 Then
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write(Row3(0)("GrossTotal").ToString)
                                _ExcelDoc.Write("</Data></Cell>")
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write("0")
                                _ExcelDoc.Write("</Data></Cell>")
                            End If
                        End If

                        'print pf value
                        If source.Tables(13).Rows(0)("ZeroColumns").ToString.ToUpper = "Y" Then
                            If CType(source.Tables(14).Rows(0)("PF"), Integer) <> 0 Then
                                If Row3.Length > 0 Then
                                    If CType(Row3(0)("PF"), Decimal) <> 0 Then
                                        _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                        _ExcelDoc.Write(Row3(0)("PF").ToString)
                                        _ExcelDoc.Write("</Data></Cell>")
                                    Else
                                        _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                        _ExcelDoc.Write("0")
                                        _ExcelDoc.Write("</Data></Cell>")
                                    End If
                                Else
                                    _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                    _ExcelDoc.Write("0")
                                    _ExcelDoc.Write("</Data></Cell>")
                                End If
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write("0")
                                _ExcelDoc.Write("</Data></Cell>")
                            End If
                        Else
                            If Row3.Length > 0 Then
                                If CType(Row3(0)("PF"), Decimal) <> 0 Then
                                    _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                    _ExcelDoc.Write(Row3(0)("PF").ToString)
                                    _ExcelDoc.Write("</Data></Cell>")
                                Else
                                    _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                    _ExcelDoc.Write("0")
                                    _ExcelDoc.Write("</Data></Cell>")
                                End If
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write("0")
                                _ExcelDoc.Write("</Data></Cell>")
                            End If
                        End If

                        'print esi value
                        If source.Tables(13).Rows(0)("ZeroColumns").ToString.ToUpper = "Y" Then
                            If CType(source.Tables(14).Rows(0)("ESI"), Integer) <> 0 Then
                                If Row3.Length > 0 Then
                                    _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                    _ExcelDoc.Write(Row3(0)("ESI").ToString)
                                    _ExcelDoc.Write("</Data></Cell>")
                                Else
                                    _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                    _ExcelDoc.Write("0")
                                    _ExcelDoc.Write("</Data></Cell>")
                                End If
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write("0")
                                _ExcelDoc.Write("</Data></Cell>")
                            End If
                        Else
                            If Row3.Length > 0 Then
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write(Row3(0)("ESI").ToString)
                                _ExcelDoc.Write("</Data></Cell>")
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write("0")
                                _ExcelDoc.Write("</Data></Cell>")
                            End If
                        End If

                        'Added by Nisha on 06 Jan 2014
                        'print PT value
                        If source.Tables(13).Rows(0)("ZeroColumns").ToString.ToUpper = "Y" Then
                            If CType(source.Tables(14).Rows(0)("ProfTax"), Integer) <> 0 Then
                                If Row3.Length > 0 Then
                                    _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                    _ExcelDoc.Write(Row3(0)("ProfTax").ToString)
                                    _ExcelDoc.Write("</Data></Cell>")
                                Else
                                    _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                    _ExcelDoc.Write("0")
                                    _ExcelDoc.Write("</Data></Cell>")
                                End If
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write("0")
                                _ExcelDoc.Write("</Data></Cell>")
                            End If
                        Else
                            If Row3.Length > 0 Then
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write(Row3(0)("ProfTax").ToString)
                                _ExcelDoc.Write("</Data></Cell>")
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write("0")
                                _ExcelDoc.Write("</Data></Cell>")
                            End If
                        End If

                        'Loop for Deductions payCode value print
                        For _RecordCounter = 0 To source.Tables(5).Rows.Count - 1
                            'for featch record in the data row according employee code wise.
                            Row2 = GetRecRowPay(DTab.Rows(_counter1)("pk_emp_code").ToString, source.Tables(5).Rows(_RecordCounter).Item("Paycode").ToString, source.Tables(7))
                            _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                            'Count the paycode
                            If Row2.Length > 0 Then
                                _ExcelDoc.Write(Row2(0)("Deductions").ToString)
                            Else
                                _ExcelDoc.Write("0")
                            End If
                            _ExcelDoc.Write("</Data></Cell>")
                        Next

                        'print it value
                        If source.Tables(13).Rows(0)("ZeroColumns").ToString.ToUpper = "Y" Then
                            If CType(source.Tables(14).Rows(0)("IT"), Integer) <> 0 Then
                                If Row3.Length > 0 Then
                                    _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                    _ExcelDoc.Write(Row3(0)("IT").ToString)
                                    _ExcelDoc.Write("</Data></Cell>")
                                Else
                                    _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                    _ExcelDoc.Write("0")
                                    _ExcelDoc.Write("</Data></Cell>")
                                End If
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write("0")
                                _ExcelDoc.Write("</Data></Cell>")
                            End If
                        Else
                            If Row3.Length > 0 Then
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write(Row3(0)("IT").ToString)
                                _ExcelDoc.Write("</Data></Cell>")
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write("0")
                                _ExcelDoc.Write("</Data></Cell>")
                            End If
                        End If

                        'print TotDeduction value
                        If source.Tables(13).Rows(0)("ZeroColumns").ToString.ToUpper = "Y" Then
                            If CType(source.Tables(14).Rows(0)("TotDed"), Integer) <> 0 Then
                                If Row3.Length > 0 Then
                                    _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                    _ExcelDoc.Write(Row3(0)("TotDed").ToString)
                                    _ExcelDoc.Write("</Data></Cell>")
                                Else
                                    _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                    _ExcelDoc.Write("0")
                                    _ExcelDoc.Write("</Data></Cell>")
                                End If
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write("0")
                                _ExcelDoc.Write("</Data></Cell>")
                            End If
                        Else
                            If Row3.Length > 0 Then
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write(Row3(0)("TotDed").ToString)
                                _ExcelDoc.Write("</Data></Cell>")
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write("0")
                                _ExcelDoc.Write("</Data></Cell>")
                            End If
                        End If

                        'print NetSalary value
                        If source.Tables(13).Rows(0)("ZeroColumns").ToString.ToUpper = "Y" Then
                            If CType(source.Tables(14).Rows(0)("NetSal"), Integer) <> 0 Then
                                If Row3.Length > 0 Then
                                    _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                    _ExcelDoc.Write(Row3(0)("NetSal").ToString)
                                    _ExcelDoc.Write("</Data></Cell>")
                                Else
                                    _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                    _ExcelDoc.Write("0")
                                    _ExcelDoc.Write("</Data></Cell>")
                                End If
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write("0")
                                _ExcelDoc.Write("</Data></Cell>")
                            End If
                        Else
                            If Row3.Length > 0 Then
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write(Row3(0)("NetSal").ToString)
                                _ExcelDoc.Write("</Data></Cell>")
                            Else
                                _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""Number"">")
                                _ExcelDoc.Write("0")
                                _ExcelDoc.Write("</Data></Cell>")
                            End If
                        End If

                        If Row3.Length > 0 Then
                            _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""String"">")
                            _ExcelDoc.Write(Row3(0)("PayMode").ToString)
                            _ExcelDoc.Write("</Data></Cell>")
                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""ED""><Data ss:Type=""String"">")
                            _ExcelDoc.Write("")
                            _ExcelDoc.Write("</Data></Cell>")
                        End If
                        'for page break on the given no record and print header part of column in next page.
                        _ExcelDoc.Write("</Row>")
                    Next
                End If

                _ExcelDoc.Write("<Row>")

                _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                _ExcelDoc.Write("")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                _ExcelDoc.Write("")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                _ExcelDoc.Write("")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                _ExcelDoc.Write("")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                _ExcelDoc.Write("Total")
                _ExcelDoc.Write("</Data></Cell>")

                Dim cnt As Integer = 0

                For GPCA = 0 To source.Tables(2).Rows.Count - 1
                    'for featch record in the data row according employee code wise.
                    GTotal = GetTotal(source.Tables(2).Rows(GPCA)("PayCode").ToString, source.Tables(4))
                    'for check the record existance
                    'Change condition for Basic Rate by Nisha on 25 Sep 2013
                    If source.Tables(2).Rows(GPCA)("Payhead").ToString.ToUpper = "BASIC" Then
                        cnt = 1
                        If GTotal.Length - 1 >= 0 Then
                            VarTotal1 = 0
                            For _cnt = 0 To GTotal.Length - 1
                                VarTotal1 = VarTotal1 + CType(_objCommon.Nz(GTotal(_cnt)("Actual_Basic"), 0), Decimal)
                            Next

                            _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                            _ExcelDoc.Write(VarTotal1.ToString)
                            _ExcelDoc.Write("</Data></Cell>")
                        Else
                            _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                            _ExcelDoc.Write("0")
                            _ExcelDoc.Write("</Data></Cell>")
                        End If
                    End If
                Next
                If cnt = 0 Then
                    _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                    _ExcelDoc.Write("0")
                    _ExcelDoc.Write("</Data></Cell>")
                End If

                _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                _ExcelDoc.Write(TGrossRate.ToString)
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                _ExcelDoc.Write(TPaidDays.ToString)
                _ExcelDoc.Write("</Data></Cell>")

                For GPCA = 0 To source.Tables(2).Rows.Count - 1

                    'for featch record in the data row according employee code wise.
                    GTotal = GetTotal(source.Tables(2).Rows(GPCA)("PayCode").ToString, source.Tables(8))
                    'for check the record existance
                    If GTotal.Length - 1 >= 0 Then
                        VarTotal = 0
                        For _cnt = 0 To GTotal.Length - 1
                            VarTotal = VarTotal + CType(_objCommon.Nz(GTotal(_cnt)("Earnings"), 0), Decimal)
                        Next

                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(VarTotal.ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("0")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next

                'print grosstotal
                If source.Tables(13).Rows(0)("ZeroColumns").ToString.ToUpper = "Y" Then
                    If CType(source.Tables(14).Rows(0)("GrossTotal"), Integer) <> 0 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(TTotalGross.ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(TTotalGross.ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Else
                    _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(TTotalGross.ToString)
                    _ExcelDoc.Write("</Data></Cell>")
                End If
                'print total pf
                If source.Tables(13).Rows(0)("ZeroColumns").ToString.ToUpper = "Y" Then
                    If CType(source.Tables(14).Rows(0)("PF"), Integer) <> 0 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(TPF.ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(TPF.ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Else
                    _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(TPF.ToString)
                    _ExcelDoc.Write("</Data></Cell>")
                End If

                'print total empesi
                If source.Tables(13).Rows(0)("ZeroColumns").ToString.ToUpper = "Y" Then
                    If CType(source.Tables(14).Rows(0)("ESI"), Integer) <> 0 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(TESI.ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(TESI.ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Else
                    _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(TESI.ToString)
                    _ExcelDoc.Write("</Data></Cell>")
                End If

                'Added by Nisha on 06 Jan 2014
                'print total empesi
                If source.Tables(13).Rows(0)("ZeroColumns").ToString.ToUpper = "Y" Then
                    If CType(source.Tables(14).Rows(0)("ProfTax"), Integer) <> 0 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(TPT.ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(TPT.ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Else
                    _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(TPT.ToString)
                    _ExcelDoc.Write("</Data></Cell>")
                End If

                'for add earnings in a variable
                For TPCodeDec = 0 To source.Tables(5).Rows.Count - 1
                    'for featch record in the data row according employee code wise.
                    GTotalDec = GetTotal(source.Tables(5).Rows(TPCodeDec)("PayCode").ToString, source.Tables(8))
                    'for check the data existance
                    If GTotalDec.Length - 1 >= 0 Then
                        VarTotal = 0
                        For Tot = 0 To GTotalDec.Length - 1
                            VarTotal = VarTotal + CType(_objCommon.Nz(GTotalDec(Tot)("Earnings"), 0), Decimal)
                        Next
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(VarTotal.ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write("0")
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Next

                'print total it
                If source.Tables(13).Rows(0)("ZeroColumns").ToString.ToUpper = "Y" Then
                    If CType(source.Tables(14).Rows(0)("IT"), Integer) <> 0 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(TIT.ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(TIT.ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Else
                    _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(TIT.ToString)
                    _ExcelDoc.Write("</Data></Cell>")
                End If

                'print total deduction
                If source.Tables(13).Rows(0)("ZeroColumns").ToString.ToUpper = "Y" Then
                    If CType(source.Tables(14).Rows(0)("TotDed"), Integer) <> 0 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(TDec.ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(TIT.ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Else
                    _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(TDec.ToString)
                    _ExcelDoc.Write("</Data></Cell>")
                End If

                'print total net salary
                If source.Tables(13).Rows(0)("ZeroColumns").ToString.ToUpper = "Y" Then
                    If CType(source.Tables(14).Rows(0)("NetSal"), Integer) <> 0 Then
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(TNetSal.ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    Else
                        _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                        _ExcelDoc.Write(TIT.ToString)
                        _ExcelDoc.Write("</Data></Cell>")
                    End If
                Else
                    _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                    _ExcelDoc.Write(TNetSal.ToString)
                    _ExcelDoc.Write("</Data></Cell>")
                End If

                _ExcelDoc.Write("<Cell ss:StyleID=""Total""><Data ss:Type=""String"">")
                _ExcelDoc.Write("")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("</Row>")

                'ss:MergeAcross=""" & colspan & """
                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell><Data ss:Type=""String"">")
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell><Data ss:Type=""String"">")
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("<Row>")
                _ExcelDoc.Write("<Cell ss:StyleID=""s21"" ss:MergeAcross=""2""><Data ss:Type=""String"">")
                _ExcelDoc.Write("Prepared By")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""s21"" ss:MergeAcross=""1""><Data ss:Type=""String"">")
                _ExcelDoc.Write("Checked By")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""s21"" ss:MergeAcross=""2""><Data ss:Type=""String"">")
                _ExcelDoc.Write("Verified By")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""s21"" ss:MergeAcross=""1""><Data ss:Type=""String"">")
                _ExcelDoc.Write("Approved By")
                _ExcelDoc.Write("</Data></Cell>")

                _ExcelDoc.Write("<Cell ss:StyleID=""s21"" ss:MergeAcross=""" & colspan - 10 & """><Data ss:Type=""String"">")
                _ExcelDoc.Write("")
                _ExcelDoc.Write("</Data></Cell>")
                _ExcelDoc.Write("</Row>")

                _ExcelDoc.Write("</Table>")
                _ExcelDoc.Write("<WorksheetOptions xmlns=""urn:schemas-microsoft-com:office:excel""><Selected/><FreezePanes/><FrozenNoSplit/><SplitHorizontal>9</SplitHorizontal><TopRowBottomPane>9</TopRowBottomPane><ActivePane>2</ActivePane><Panes><Pane><Number>3</Number></Pane><Pane><Number>2</Number><ActiveRow>11</ActiveRow><ActiveCol>2</ActiveCol></Pane></Panes><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions>")
                _ExcelDoc.Write(" </Worksheet>")
                _ExcelDoc.Write(end_ExcelXML)
            Catch ex As Exception
                _objcommonExp.PublishError("For generate the excel file ExportToExcelXMLFinalSalary()", ex)
            End Try
        End Sub
        Private Function GetRecRow(ByVal fk_emp_code As String, ByVal dt As DataTable) As DataRow()
            Try
                Dim Drow() As DataRow
                Drow = dt.Select("fk_emp_code='" & fk_emp_code & "'")
                Return Drow
            Catch ex As Exception
                _objcommonExp.PublishError("GetRecRow()", ex)
            End Try
            Return Nothing
        End Function
        'to filter the datatable rows according to the employee code and pay code
        Private Function GetRecRowPay(ByVal fk_emp_code As String, ByVal PayCode As String, ByVal dt As DataTable) As DataRow()
            Try
                Dim Drow() As DataRow
                Drow = dt.Select("fk_emp_code='" & fk_emp_code & "' and Paycode='" & PayCode & "'")
                Return Drow
            Catch ex As Exception
                _objcommonExp.PublishError("GetRecRowPay()", ex)
            End Try
            Return Nothing
        End Function
        Private Function GetTotal(ByVal paycode As String, ByVal dt As DataTable) As DataRow()
            Try
                Dim Drow() As DataRow
                Drow = dt.Select("paycode='" & paycode & "'")
                Return Drow
            Catch ex As Exception
                _objcommonExp.PublishError("GetTotal()", ex)
            End Try
            Return Nothing
        End Function
        'added by moksh
        Protected Sub Btndelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Btndelete.Click
            Try
                Dim arrparam(1) As SqlClient.SqlParameter, dt As DataTable
                arrparam(0) = New SqlClient.SqlParameter("@Fk_report_Id", DdlreportType.SelectedValue.ToString)
                arrparam(1) = New SqlClient.SqlParameter("@fk_userlog", Session("userlogKey").ToString)
                dt = _ObjData.GetDataTableProc("paysp_clearMailFormat", arrparam)
                If dt.Rows.Count < 0 Then
                    _objCommon.ShowMessage("D", lblmsg)
                Else
                    _objCommon.ShowMessage("M", lblmsg, "No record(s) Found according to the selection criteria !", False)
                End If
            Catch ex As Exception
                _objcommonExp.PublishError("ClearMailFormat", ex)
            End Try
        End Sub
        Protected Sub RblNoSearch_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RblNoSearch.SelectedIndexChanged
            ShowHideNoSearch()
        End Sub
        Protected Sub ddlRepIn_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlRepIn.SelectedIndexChanged
            If ddlRepIn.SelectedValue <> "P" Then
                RblNoSearch.SelectedValue = "S"
            Else
                RblNoSearch.SelectedValue = "P"
            End If
            ShowHideNoSearch()
        End Sub
        Private Sub ShowHideNoSearch()
            CheckProcessLocked()
            If (DdlreportType.SelectedValue.ToString.ToUpper = "T" OrElse DdlreportType.SelectedValue.ToString.ToUpper = "R" _
                    OrElse DdlreportType.SelectedValue.ToString.ToUpper = "S" OrElse DdlreportType.SelectedValue.ToString.ToUpper = "57") _
                    AndAlso ddlRepIn.SelectedValue.ToString.ToUpper = "P" AndAlso rbtnmail.Checked Then
                btnPublishedPDF.Visible = True
                btnPublishedPDF.ToolTip = "Download Already Published " & DdlreportType.SelectedItem.Text
            Else
                btnPublishedPDF.Visible = False
                btnPublishedPDF.ToolTip = "Download Already Published Pay Slips"
            End If
            If ddlRepIn.SelectedValue = "P" Then
                TrNoSearch.Style("display") = ""
                tblpwd.Style("display") = ""
                trselall.Style("display") = ""
                If DdlreportType.SelectedValue = "SL" Then
                    trRepEmail.Style("display") = ""
                Else
                    trRepEmail.Style("display") = "none"
                    chkemailrepmanager.Checked = False
                End If
                If DdlreportType.SelectedValue.Equals("67") Then
                    tblpwd.Style("display") = "none"
                End If
                'Rohtas Singh for Merge two PDF Slip on 14 Feb 2018
                If DdlreportType.SelectedValue = "55" Or DdlreportType.SelectedValue.Equals("57") Then
                    trMerge.Style("display") = ""
                Else
                    trMerge.Style("display") = "none"
                End If
                If DdlreportType.SelectedValue.Equals("R") Then
                    btnWOPWD.Visible = True
                Else
                    btnWOPWD.Visible = False
                End If

            ElseIf ddlRepIn.SelectedValue = "H" Then
                TrNoSearch.Style("display") = "none"
                tblpwd.Style("display") = "none"
                trselall.Style("display") = "none"
                If DdlreportType.SelectedValue = "SL" Then
                    trRepEmail.Style("display") = ""
                Else
                    trRepEmail.Style("display") = "none"
                    chkemailrepmanager.Checked = False
                End If
                'Rohtas Singh for Merge two PDF Slip on 14 Feb 2018
                trMerge.Style("display") = "none"
            Else
                TrNoSearch.Style("display") = "none"
                tblpwd.Style("display") = "none"
                trselall.Style("display") = ""
                trRepEmail.Style("display") = "none"
                chkemailrepmanager.Checked = False
                'Rohtas Singh for Merge two PDF Slip on 14 Feb 2018
                trMerge.Style("display") = "none"
            End If
            If RblNoSearch.SelectedValue = "P" Then
                tableshow.Style("display") = ""
                trbutton.Style("display") = ""
                TrDg.Style("display") = "none"
                Tr2.Style("display") = "none"
                'Tr3.Style("display") = "none"
                BtnSend.Visible = False
                BtnSendCCBCC.Visible = False
                'Btnsearch.Visible = False
                btnSave.Visible = True
                TdSearch1.Style("display") = "none"
                TdSearch2.Style("display") = "none"
            Else
                tableshow.Style("display") = "none"
                trbutton.Style("display") = ""
                TrDg.Style("display") = ""
                Tr2.Style("display") = ""
                BtnSend.Visible = True
                BtnSendCCBCC.Visible = True
                'Btnsearch.Visible = True
                btnSave.Visible = True
                TdSearch1.Style("display") = ""
                TdSearch2.Style("display") = ""
            End If
            If DdlreportType.SelectedValue = "S" And ddlRepIn.SelectedValue = "P" Then
                If RblGrpbyPublish.SelectedValue <> "" Then
                    BtnPublishGrpBy.Visible = True
                Else
                    BtnPublishGrpBy.Visible = False
                End If
                TrGrpbyPublish.Style("display") = ""
            Else
                BtnPublishGrpBy.Visible = False
                TrGrpbyPublish.Style("display") = "none"
            End If

            If DdlreportType.SelectedValue.Equals("62") Then
                tremail.Style("display") = "none"
                tblsp.Style("display") = "none"
                tblSh.Style("display") = "none"
                trselall.Style("display") = "none"
            End If

            'Added by Quadir on 14 OCT 2020
            If DdlreportType.SelectedValue.Equals("R") And ddlRepIn.SelectedValue = "P" Then
                TrSlipPubMode.Style("display") = ""
            Else
                TrSlipPubMode.Style("display") = "none"
            End If


            lblMsgSlip.Text = ""
            lblMailMsg.Text = ""
            LnkPDF.Style.Value = "display:none"
            DgPayslip.DataSource = Nothing
            DgPayslip.DataBind()
            download_pdf1.Style.Value = "display:none;border:0;cursor:pointer;"
        End Sub
        Private Sub GenerateLogFile(ByVal Published As String, ByVal NotPublished As String)
            Dim oWrite As System.IO.StreamWriter, _fs As FileStream, _fileAdd As String = ""
            If Not HttpContext.Current.Session("compcode") Is Nothing Then
                _fileAdd = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles"
            Else
                _fileAdd = _objCommon.GetDirpath() & ""
            End If
            If Not Directory.Exists(_fileAdd) Then
                Directory.CreateDirectory(_fileAdd)
            End If
            _fileAdd = _fileAdd & "\Log_PaySlips.txt"
            If File.Exists(_fileAdd) Then
                _fs = New FileStream(_fileAdd, FileMode.Append, FileAccess.Write)
            Else
                _fs = New FileStream(_fileAdd, FileMode.Create, FileAccess.ReadWrite)
            End If
            oWrite = New StreamWriter(_fs)
            Try
                oWrite.WriteLine("")
                oWrite.WriteLine("*******************************" & Date.Now().ToString("dd-MMM-yyyy HH:mm:ss") & "*****************************************************")
                If Not HttpContext.Current.Session("compcode") Is Nothing Then
                    oWrite.WriteLine("User ID: " & HttpContext.Current.Session("UID").ToString & "    Report Type: " & DdlreportType.SelectedItem.ToString & "    Month-Year: " & ddlMonthYear.SelectedItem.ToString & "    System IP: " & Request.UserHostAddress.ToString.Trim())
                Else
                    oWrite.WriteLine("User ID: Not Avalable          Module Type : Not Avalable")
                End If
                If Published.ToString <> "" Then
                    oWrite.WriteLine("********************************Published Employee Code(s):****************************************")
                    oWrite.WriteLine(Published.ToString)
                    oWrite.WriteLine("***************************************************************************************************")
                End If
                If NotPublished.ToString.Trim <> "" Then
                    oWrite.WriteLine("********************************Not Published Employee Code(s):************************************")
                    oWrite.WriteLine(NotPublished.ToString)
                    oWrite.WriteLine("***************************************************************************************************")
                End If
                oWrite.Close()
                oWrite.Dispose()
                _fs.Close()
                _fs.Dispose()
            Catch ex As Exception

            Finally
                oWrite.Close()
                oWrite.Dispose()
                _fs.Close()
                _fs.Dispose()
            End Try
        End Sub
        Protected Sub RblGrpbyPublish_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RblGrpbyPublish.SelectedIndexChanged
            ShowHideNoSearch()
        End Sub
        Protected Sub BtnPublishGrpBy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnPublishGrpBy.Click
            lblMailMsg.Text = ""
            lblMsgSlip.Text = ""
            lblProcessBarMsg.Text = ""
            LnkPDF.Style.Value = "display:none"
            HidEmailCCBCC.Value = "UNT"
            SendReportPDF("S")
        End Sub
        Private Sub COC_MaindatoryCheck()
            Try
                Dim dt As DataTable, ArrParam(2) As SqlParameter
                ArrParam(0) = New SqlParameter("@SFinYear", _objCommon.nNz(Session("SfinDate").ToString))
                ArrParam(1) = New SqlParameter("@EFinYear", _objCommon.nNz(Session("EfinDate").ToString))
                ArrParam(2) = New SqlParameter("@CostcenterWise", SqlDbType.VarChar, 500)
                ArrParam(2).Direction = ParameterDirection.Output
                dt = _ObjData.GetDataTableProc("PaySP_Mandatory_Costcenter", ArrParam)
                HidCocManCheck.Value = ArrParam(2).Value.ToString()
                dt.Clear()
                dt.Dispose()
            Catch ex As Exception
                _objcommonExp.PublishError("Error in CheckCostCenter_Maindatory()", ex)
            End Try
        End Sub
        'Added by Rohtas Singh on 08 Dec 2017 for generate the Monthly Salary Slip (MAX Life)
        Private Sub Export_CSV(ByVal Dt As DataTable)
            Dim filename As String = "", Recordtype As String = "", webCurrApplication As New System.Web.HttpApplication, complexID As Guid = Guid.NewGuid(),
              sAppPath As String = "", Year As String, Month As String, Emp_Code As String, Order As String, Fld As String, Rate As String,
              Value As String, Arr As String, Total As String, Type As String, X As Integer = 0, Amt As String = "", str As String = ""

            sAppPath = Server.MapPath(Request.ApplicationPath)
            If Not System.IO.Directory.CreateDirectory(sAppPath).Exists Then
                System.IO.Directory.CreateDirectory(sAppPath)
            End If

            filename = Server.MapPath(Request.ApplicationPath) & "\" & Session("COMPCODE").ToString & "\Documents\MHC_PAYSLIP" & "_" & Right(complexID.ToString, 6) & ".CSV"

            oWrite = IO.File.CreateText(filename)
            If Dt.Rows.Count > 0 Then
                oWrite.WriteLine("Year,Month,Emp_Code,Order,Fld,Rate,Value,Arr,Total,Type")
                For count As Integer = 0 To Dt.Rows.Count - 1
                    Year = Dt.Rows(count)("Year").ToString
                    Month = Dt.Rows(count)("Month").ToString
                    Emp_Code = Dt.Rows(count)("Emp_Code").ToString
                    Order = Dt.Rows(count)("Order").ToString
                    Fld = Dt.Rows(count)("Fld").ToString
                    Rate = Dt.Rows(count)("Rate").ToString
                    Value = Dt.Rows(count)("Value").ToString
                    Arr = Dt.Rows(count)("Arr").ToString
                    Total = Dt.Rows(count)("Total").ToString
                    Type = Dt.Rows(count)("Type").ToString


                    oWrite.WriteLine(Year & "," & Month & "," & Emp_Code & "," & Order & "," & Fld & "," & Rate & "," & Value & "," & Arr & "," & Total & "," & Type)
                Next
                oWrite.Close()

                Response.Clear()
                Response.BufferOutput = False

                Response.ContentType = "application/zip"
                Response.AddHeader("content-disposition", "filename=MHC_PAYSLIP.zip")
                Using zip As New ZipFile()
                    zip.AddFile(filename, "MHC_PAYSLIP")
                    zip.Save(Response.OutputStream)
                End Using

                If File.Exists(filename) Then
                    File.Delete(filename)
                End If

                lblmsg.Text = ""

                Response.End()
            Else
                oWrite.WriteLine("No Record Found!")
                oWrite.Close()
            End If

        End Sub
        Protected Sub BtnLog_Click(sender As Object, e As System.EventArgs) Handles BtnLog.Click
            Try
                Dim ds As New DataSet, FileName As String = "", _fileAdd As String = "", dtrow As DataRow
                If Not HttpContext.Current.Session("compcode") Is Nothing Then
                    _fileAdd = _objCommon.GetDirpath(HttpContext.Current.Session("COMPCODE").ToString) & "\PDFFiles"
                Else
                    _fileAdd = _objCommon.GetDirpath() & ""
                End If
                FileName = _fileAdd & "\Log_PaySlips_" & Left(Trim(ddlMonthYear.SelectedItem.Text.ToString), 3) & "_" & Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4) & ".xml"
                If System.IO.File.Exists(FileName.ToString) = True Then
                    ds.ReadXml(FileName.ToString)
                End If

                If ds.Tables.Count > 0 Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        Export_Text(ds, 1)
                    End If
                Else
                    ds.Tables.Add("T1")
                    ds.Tables(0).Columns.Add(New DataColumn("E"))
                    dtrow = ds.Tables(0).NewRow
                    dtrow(0) = "PDF file not published for selected month-year."
                    ds.Tables(0).Rows.Add(dtrow)
                    Export_Text(ds, 0)
                End If
            Catch ex As Exception
                _objcommonExp.PublishError("Error in BtnLog_Click()", ex)
            End Try
        End Sub
        Private Sub Export_Text(ByVal ds As DataSet, TotRow As Integer)
            If ds.Tables(0).Rows.Count > 0 Then
                Dim _filePath As String = "", _strAutoNum As Guid = Guid.NewGuid(), intFileNbr As Integer = FreeFile() _
                    , _filePathtxt As String, ZipName As String = "", EmpCode As String = "", TotCount As Integer = 0
                _filePath = Server.MapPath(Request.ApplicationPath).ToString & "\" & Session("COMPCODE").ToString & "\Documents\Log_Payslips\"

                If Not System.IO.Directory.Exists(_filePath) Then
                    System.IO.Directory.CreateDirectory(_filePath)
                End If
                _filePathtxt = Server.MapPath(Request.ApplicationPath).ToString & "\" & Session("COMPCODE").ToString & "\Documents\Log_Payslips\Log_Payslips_" & Right(_strAutoNum.ToString, 6) & ".txt"
                ZipName = Path.GetDirectoryName(_filePathtxt) & "\Log_Payslips.Zip"

                FileOpen(intFileNbr, _filePathtxt, OpenMode.Output, OpenAccess.Write)
                For I As Integer = 0 To ds.Tables(0).Rows.Count - 1
                    TotCount = TotCount + 1
                    EmpCode = EmpCode.ToString + "," + ds.Tables(0).Rows(I).Item("E").ToString
                Next
                EmpCode = Right(EmpCode, Len(EmpCode) - 1)
                If TotRow > 0 Then
                    PrintLine(intFileNbr, ("Bellow mentioned employee's PDF has been generated on server :-"))
                    PrintLine(intFileNbr, (""))
                    PrintLine(intFileNbr, ("Total Published : " & TotCount))
                    PrintLine(intFileNbr, (""))
                    PrintLine(intFileNbr, (EmpCode.ToString))
                    FileClose(intFileNbr)
                Else
                    PrintLine(intFileNbr, (""))
                    PrintLine(intFileNbr, (EmpCode.ToString))
                    FileClose(intFileNbr)
                End If

                _objCommon.CreateZipFile(Path.GetDirectoryName(_filePathtxt), ZipName, _filePathtxt)
                getFileDownLoad("PaySlipLogFiles", ZipName)

                If File.Exists(_filePath) Then
                    File.Delete(_filePath)
                End If
                If File.Exists(_filePathtxt) Then
                    File.Delete(_filePathtxt)
                End If
                Response.End()
                ds.Clear()
                ds.Dispose()
            Else
                Dim _Msg As New List(Of PayrollUtility.UserMessage)
                _Msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "PDF file not published for selected month-year !"})
                _objCommon.ShowMessage(_Msg)
            End If
        End Sub
        Private Function getFileDownLoad(Optional ByVal FileName As String = "", Optional ByVal FilePath As String = "") As String
            Response.Clear()
            HttpContext.Current.Response.ContentType = "application/zip"
            Response.AddHeader("Refresh", "12;URL=Rpt_SalaryStructureSelCriteria.aspx")
            Response.AddHeader("Content-Disposition", "attachment; filename = " & FileName & ".zip")
            HttpContext.Current.Response.WriteFile(FilePath)
            Response.End()
            Return ""
        End Function

        Protected Sub populate_multilingDDl()
            Dim dt As DataTable, Name As String = Nothing, Type As String = Nothing

            dt = _ObjData.ExecSQLQuery("Select pk_ling_id,multilingualName from mstPayslipmultilingual order by pk_ling_id")
            ddllingual.Items.Clear()
            ddllingual.DataTextField = "multilingualName"
            ddllingual.DataValueField = "pk_ling_id"
            ddllingual.DataSource = dt
            ddllingual.DataBind()
            ddlmultilingual.Items.Clear()
            ddlmultilingual.DataTextField = "multilingualName"
            ddlmultilingual.DataValueField = "pk_ling_id"
            ddlmultilingual.DataSource = dt
            ddlmultilingual.DataBind()

        End Sub
        'Protected Sub ddlMonthYear_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlMonthYear.SelectedIndexChanged
        '    If (DDLPaySlipType.SelectedValue.Equals("62") And rbtnslip.Checked) Or (DdlreportType.SelectedValue.Equals("62") And rbtnmail.Checked) Then
        '        troffcycle.Style("display") = ""
        '    Else
        '        troffcycle.Style("display") = "none"
        '    End If
        '    populateFromdate()
        'End Sub


        Private Sub ExportToExcelXMLREFNF(ByVal source As DataSet, ByRef swr As StreamWriter)
            Dim _RowCount As Integer = 0
            Try
                Dim _ExcelXML As New StringBuilder()
                _ExcelXML.Append("<?xml version=""1.0""?><?mso-application progid=""Excel.Sheet""?>" & Chr(13) & "" & Chr(10) & "<Workbook ")
                _ExcelXML.Append("xmlns=""urn:schemas-microsoft-com:office:spreadsheet""" & Chr(13) & "" & Chr(10) & "")
                _ExcelXML.Append(" xmlns:o=""urn:schemas-microsoft-com:office:office""" & Chr(13) & "" & Chr(10) & " ")
                _ExcelXML.Append(" xmlns:x=""urn:schemas-microsoft-com:office:excel""" & Chr(13) & "" & Chr(10))
                _ExcelXML.Append(" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">" & Chr(13) & "" & Chr(10) & " <Styles>" & Chr(13) & "" & Chr(10) & " ")
                _ExcelXML.Append("<Style ss:ID=""Default"" ss:Name=""Normal"">" & Chr(13) & "" & Chr(10) & " ")
                _ExcelXML.Append("<Alignment ss:Horizontal=""Left"" ss:Vertical=""Bottom""/>" & Chr(13) & "" & Chr(10) & " <Borders><Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1""/> <Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1""/><Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1""/></Borders>")
                _ExcelXML.Append("" & Chr(13) & "" & Chr(10) & " <Font ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9""/>" & Chr(13) & "" & Chr(10) & " <Interior/>" & Chr(13) & "" & Chr(10) & " <NumberFormat/>")
                _ExcelXML.Append("" & Chr(13) & "" & Chr(10) & " <Protection/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")

                _ExcelXML.Append("" & Chr(13) & "" & Chr(10) & " <Style ss:ID=""BoldColumn"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelXML.Append("ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & " <Interior ss:Color=""#F2DDDC"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & " ")

                _ExcelXML.Append("<Style ss:ID=""CompanyDet"">" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Center"" ss:Vertical=""Center""/>" & Chr(13) & "" & Chr(10) & " <Font ")
                _ExcelXML.Append("ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""9"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & " <Interior ss:Color=""#DBEEF3"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & " ")

                _ExcelXML.Append("<Style ss:ID=""StringLiteral"">" & Chr(13) & "" & Chr(10) & " <NumberFormat")
                _ExcelXML.Append(" ss:Format=""@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " <Style ")

                _ExcelXML.Append("ss:ID=""Decimal"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelXML.Append("ss:Format=""0.00""/>" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Right"" ss:Vertical=""Bottom""/></Style>" & Chr(13) & "" & Chr(10) & " ")

                _ExcelXML.Append("<Style ss:ID=""Integer"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelXML.Append("ss:Format=""0""/>" & Chr(13) & "" & Chr(10) & " <Alignment ss:Horizontal=""Right"" ss:Vertical=""Bottom""/></Style>" & Chr(13) & "" & Chr(10) & " <Style ")

                _ExcelXML.Append("ss:ID=""DateLiteral"">" & Chr(13) & "" & Chr(10) & " <NumberFormat ")
                _ExcelXML.Append("ss:Format=""dd-mmm-yyyy;@""/>" & Chr(13) & "" & Chr(10) & " </Style>" & Chr(13) & "" & Chr(10) & " ")

                _ExcelXML.Append("<Style ss:ID=""Custum2"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelXML.Append("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelXML.Append("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelXML.Append("<Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelXML.Append("<Style ss:ID=""Custum3"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelXML.Append("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelXML.Append("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000"" ss:Bold=""1""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelXML.Append("<Interior ss:Color=""Yellow"" ss:Pattern=""Solid""/> </Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelXML.Append("<Style ss:ID=""EmpDetail"">" & Chr(13) & "" & Chr(10) & "")
                _ExcelXML.Append("<Alignment ss:Horizontal=""Center"" ss:Vertical=""Center"" ss:ReadingOrder=""LeftToRight"" ss:WrapText=""0""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelXML.Append("<Font x:CharSet=""1"" ss:FontName=""Verdana"" x:Family=""Swiss"" ss:Size=""8"" ss:Color=""#000000"" ss:Bold=""0""/>" & Chr(13) & "" & Chr(10) & "")
                _ExcelXML.Append("<Interior ss:Color=""#CCFFFF"" ss:Pattern=""Solid""/></Style>" & Chr(13) & "" & Chr(10) & "")

                _ExcelXML.Append("</Styles>" & Chr(13) & "" & Chr(10) & " ")

                Dim start_ExcelXML As String = _ExcelXML.ToString(), sheetCount As Integer = 1
                Const end_ExcelXML As String = "</Workbook>"
                swr.Write(start_ExcelXML)

                swr.Write("<Worksheet ss:Name=""Sheet" + sheetCount.ToString + """>")
                swr.Write("<Table><Column ss:Width=""150""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""300""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/><Column ss:Width=""200""/>")

                swr.Write("<Row>")

                For x As Integer = 0 To source.Tables(0).Columns.Count - 1
                    swr.Write("<Cell ss:StyleID=""Custum2""><Data ss:Type=""String"">")
                    swr.Write(source.Tables(0).Columns(x).ColumnName)
                    swr.Write("</Data></Cell>")
                Next

                swr.Write("</Row>")
                For Each x As DataRow In source.Tables(0).Rows
                    _RowCount += 1
                    swr.Write("<Row>")

                    For y As Integer = 0 To source.Tables(0).Columns.Count - 1
                        Dim XMLstring As String = x(y).ToString()
                        XMLstring = XMLstring.Trim()
                        swr.Write("<Cell ss:StyleID=""EmpDetail"">" + "<Data ss:Type=""String"">")
                        swr.Write(XMLstring)
                        swr.Write("</Data></Cell>")
                    Next
                    swr.Write("</Row>")
                Next

                swr.Write("</Table>")
                swr.Write(" </Worksheet>")
                swr.Write(end_ExcelXML)

            Catch ex As Exception
                'LogMessage.log.Error("ExportToExcelXML()", ex)
                Throw (New Exception("CreateXLSheetException", ex))
            End Try

        End Sub

        Protected Sub btnWOPWD_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnWOPWD.Click
            lblProcessBarMsg.Text = ""
            Try
                Dim _TotRec As String, _RecCount As Integer = 0, _Counter As Integer, To_mail As String, arrparam(4) As SqlClient.SqlParameter, _ds As New DataSet _
             , _dt As New DataTable, _dRow As DataRow, _StrEmpCode As String = "", MailNotSentCount As String = "", _strVal As String = Guid.NewGuid.ToString _
             , EmpPassType As String = "", EmpCode As String = "", _DsMailDoc As New DataSet, _dRowDoc As DataRow = Nothing, _DsNS As New DataSet _
             , MsgReturn As String = "", _msg As New List(Of PayrollUtility.UserMessage), EmpCodeSearch As String = String.Empty,
             Dep As String = "", Desig As String = "", Grad As String = "", Level As String = "", CC As String = "", Loc As String = "" _
             , unit As String = "", SalBase As String = "", EmpFName As String = "", EmpLName As String = "", EmpType As String = "", Dtt As New DataTable,
             ArrParams(1) As SqlClient.SqlParameter, dst As DataSet
                Dim gcs_service As Integer = 0
                Try
                    Dim gcs As New DataTable
                    Dim _mm As String = "", _yyyy As String = "", _Month As String = ""
                    If ddlMonthYear.SelectedItem IsNot Nothing Then
                        _Month = Left(MonthName(CType(ddlMonthYear.SelectedValue.ToString, Integer)), 3)
                        'Left(MonthName(CType(ddlMonthYear.SelectedValue.ToString, Integer)), 3) & Right(ddlMonthYear.SelectedItem.Text.ToString, 4)
                        _mm = CType(ddlMonthYear.SelectedValue.ToString, Integer)
                        _yyyy = Right(ddlMonthYear.SelectedItem.Text.ToString, 4)
                    End If
                    If gcs_service_obj.IsPaySlipAllowedByMonthYear(Session("compCode"), "TAXSLIP", _mm, _yyyy) Then
                        gcs_service = 1
                    End If

                Catch ex As Exception

                End Try

                If RblNoSearch.SelectedValue = "S" Then
                    For counter = 0 To DgPayslip.Items.Count - 1
                        If CType(DgPayslip.Items(counter).FindControl("chkEmpHold"), CheckBox).Checked = True Then
                            EmpCodeSearch = EmpCodeSearch + DgPayslip.Items(counter).Cells(1).Text.ToString + ","
                        End If
                    Next
                    dst = ReturnDsSearch("", "PDF", EmpCodeSearch, "W")
                Else
                    dst = ReturnDsSearch("", "PDF", "", "W")
                End If
                If dst.Tables(0).Rows.Count > 0 Then
                    Dim _array() As String, _AppPath As String = HttpRuntime.AppDomainAppVirtualPath.ToString, Path As String, cnt As Integer
                    cnt = dst.Tables(0).Rows.Count
                    Dim Process_Type As String = "TAXSLIP"
                    If gcs_service = 0 Then
                        process_status_id.Value = ""
                        _ObjData.ExecSQLQuery("DELETE FROM SalaryProcessStatus WHERE process_user_id = '" & Convert.ToString(HttpContext.Current.Session("UId")) & "' and Process_Type='TAXSLIP' " & "" &
                                              "DELETE FROM Log_Error_SalaryProcess WHERE process_user_id = '" & Convert.ToString(HttpContext.Current.Session("UId")) & "' and Process_Type='TAXSLIP' " & "" &
                                              "Insert Into SalaryProcessStatus(process_user_id,status,total_to_process,record_created,Process_Type) VALUES('" & Convert.ToString(HttpContext.Current.Session("UId")) & "','START', '" & Convert.ToString(cnt) & "',GetDate(),'TAXSLIP')")

                    Else
                        Dim _dt_temp, temp As New DataTable
                        Dim month As String = ddlMonthYear.SelectedValue.ToString, year As String = Right(ddlMonthYear.SelectedItem.Text.ToString, 4)

                        _dt_temp = _ObjData.ExecSQLQuery("DELETE FROM GcsSalaryProcessStatus WHERE process_user_id = '" & Convert.ToString(HttpContext.Current.Session("UId")) & "' and Process_Type='" & Process_Type & "' " & "" &
                                                       "Insert Into GcsSalaryProcessStatus(process_user_id, status, total_processed, total_to_process, record_created, Process_Type, mm, YYYY) VALUES('" & Convert.ToString(HttpContext.Current.Session("UId")) & "','START','0', '" & Convert.ToString(cnt) & "',GetDate(),'" & Process_Type & "', '" & month & "', '" & year & "'); SELECT SCOPE_IDENTITY() AS id;")

                        If _dt_temp.Rows.Count > 0 Then
                            Session("process_status_id") = _dt_temp.Rows(0)("id")
                            process_status_id.Value = _dt_temp.Rows(0)("id")
                        End If
                        _dt_temp.Dispose()

                    End If

                    _array = Split(_AppPath, "/")
                    _AppPath = _array(_array.Length - 1)
                    HidAppPath.Value = _AppPath
                    Path = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Left(MonthName(CType(ddlMonthYear.SelectedValue.ToString, Integer)), 3) & Right(ddlMonthYear.SelectedItem.Text.ToString, 4) & "\TaxPaySlipWOPWD\"
                    If Not Directory.Exists(Path) Then
                        Directory.CreateDirectory(Path)
                    End If
                    If gcs_service = 1 Then
                        Session("pdf_file_location") = Path.ToString
                    End If
                    HidPath.Value = Replace(Replace(Path, "\", "~").ToString, "/", "~").ToString

                    ArrParams(0) = New SqlClient.SqlParameter("@Month", ddlMonthYear.SelectedValue.ToString)
                    ArrParams(1) = New SqlClient.SqlParameter("@Year", Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                    Dtt = _ObjData.GetDataTableProc("PaySP_PaySlipConfigure_Sel", ArrParams)
                    If Dtt.Rows.Count > 0 Then
                        HidPdfName.Value = Dtt.Rows(0).Item("PdfName").ToString
                    End If
                    _dt.Columns.Add(New DataColumn("EmpCode"))
                    _ds.Tables.Add(_dt)
                    _DsMailDoc.Tables.Add("Table1")
                    _DsMailDoc.Tables(0).Columns.Add(New DataColumn("EmpCode"))
                    lblMsgSlip.Text = ""
                    lblMailMsg.Text = ""
                    HidYear.Value = Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4)
                    If RblNoSearch.SelectedValue = "S" Then
                        For counter = 0 To DgPayslip.Items.Count - 1
                            If CType(DgPayslip.Items(counter).FindControl("chkEmpHold"), CheckBox).Checked = True Then
                                EmpCodeSearch = EmpCodeSearch + DgPayslip.Items(counter).Cells(1).Text.ToString + ","
                            End If
                        Next
                    End If

                    _DsNS = ReturnDsSearch(MsgReturn, "PDF", EmpCodeSearch)

                    If MsgReturn.ToString <> "" Then
                        _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = MsgReturn.ToString})
                        _objCommon.ShowMessage(_msg)
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "closepopup1224", "CloseSlipProgressbar();", True)
                        Exit Sub
                    End If
                    If _DsNS.Tables.Count > 0 Then
                        If _DsNS.Tables(0).Rows.Count > 0 Then
                            _TotRec = _DsNS.Tables(0).Rows.Count.ToString
                            For _Counter = 0 To _DsNS.Tables(0).Rows.Count - 1
                                _dRow = _ds.Tables(0).NewRow
                                _dRow(0) = Convert.ToString(_DsNS.Tables(0).Rows(_Counter)("fk_emp_code"))
                                _ds.Tables(0).Rows.Add(_dRow)
                                EmpCode = EmpCode + Convert.ToString(_DsNS.Tables(0).Rows(_Counter)("fk_emp_code")) + ","
                            Next
                        End If
                    End If

                    If RblNoSearch.SelectedValue = "S" Then
                        HidEmpPdf.Value = HidEmpPdf.Value + "," + EmpCode.ToString.Trim
                        EmpCode = Left(EmpCode, Len(EmpCode) - 1).ToString
                    Else
                        HidEmpPdf.Value = HidEmpPdf.Value + "," + EmpCode.ToString.Trim
                        EmpCode = Left(EmpCode, Len(EmpCode) - 1).ToString
                    End If

                    If _ds.Tables(0).Rows.Count > 0 Then
                        _ds.WriteXml(Server.MapPath("XMLFiles\" & _strVal.ToString & ".xml"))

                        Dim monthvalue As String, YearVal As String, HDocStatus As String, HReimbSts As String, QryString As String = "", PM1 As String = "",
                           PM11 As String = "0", PM13 As String = "A", PM14 As String = "true", PopupScript As String, var As String, ReportType As String _
                           , RepType As String, flag As String = "", EmpStatus As String = USearch.UCddlEmp.ToString, COSTCENTER As String = "", Grpwise As String = "",
                           reportid As String = "", MailFrom As String = "", _MonthValue As String = ddlMonthYear.SelectedValue.ToString, _YearValue As String = "",
                           dt As DataTable = Nothing, LoopVar As Integer = 0, InvCode As String = "", LOCATION As String = USearch.UCddllocation.ToString, PayCode As String = ""
                        HDocStatus = "P"
                        HReimbSts = "N"
                        monthvalue = Me.ddlMonthYear.SelectedValue
                        YearVal = HidYear.Value.ToString
                        ReportType = "V"
                        RepType = "P"
                        HttpContext.Current.Session.Add("MsgPDF", "")
                        flag = "S"
                        EmpPassType = "0"
                        PM13 = ddlshowsal.SelectedValue.ToString
                        reportid = HidRepId.Value.ToString
                        COSTCENTER = USearch.UCddlcostcenter.ToString
                        lblMailMsgWOPWD.Text = "Click on PDF icon to open payslips without password"
                        If gcs_service = 0 Then
                            LnkPDFWOPWD.Style.Value = "display:''"
                            download_pdf2.Style.Value = "display:'none'"
                        Else
                            download_pdf2.Style.Value = "display:''"
                            LnkPDFWOPWD.Style.Value = "display:'none'"
                        End If

                        Dim BPath As String = _objCommon.GetBaseHref(), filepat As String = HttpContext.Current.Server.MapPath(HttpContext.Current.Request.ApplicationPath())
                        var = "P~~~~~" + COSTCENTER + "~~~~~~" + monthvalue + "~" + YearVal + "~" + PM1 + "~" + PM1 + "~" + PM1 + "~1~" + PM14 + "~" + PM13 + "~" + "S" + "~" + EmpStatus.ToString + "~" + EmpPassType.ToString + "~" + "N" + "~" + "" + "~" + reportid & "~~~51~" & _strVal.ToString & ".xml~" & IIf(chkHelp1.Checked = True, "Y", "N").ToString & "~" & HidPdfName.Value.ToString & "~" & BPath.ToString & "~" & filepat.ToString & "~" & "WOPWD"

                        If Not Session(_strVal) Is Nothing Then
                            Session(_strVal) = var.ToString
                            var = _strVal.ToString
                        Else
                            Session.Remove(_strVal)
                            Session(_strVal) = var.ToString
                            var = _strVal.ToString
                        End If
                        If DdlreportType.SelectedValue.ToString.ToUpper = "R" And ddlRepIn.SelectedValue.ToUpper = "P" Then
                            HidPreVal.Value = var.ToString
                            lblMsgSlip.Text = ""
                            lblMailMsg.Text = ""
                            LnkPDF.Style.Value = "display:none"
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup", "ShowTaxDetailsWOPWD();", True)
                        End If
                    End If
                Else
                    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "Record does not exists as per selected criteria!"})
                    _objCommon.ShowMessage(_msg)
                    LnkPDFWOPWD.Style.Value = "display:None"
                    lblMailMsgWOPWD.Text = ""
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "closepopup1222", "CloseSlipProgressbar();", True)
                    Exit Sub
                End If

            Catch ex As Exception
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "closepopup1221", "CloseSlipProgressbar();", True)
                _objcommonExp.PublishError("Error in btnWOPWD_Click()", ex)

            End Try
        End Sub


        Protected Sub LnkPDFWOPWD_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LnkPDFWOPWD.Click
            Try
                Dim Month As String, Year As String, _strJava As New StringBuilder, filepath As String, FileName() As String = Nothing, FName As String = ""
                lblMailMsg.Text = ""
                lblmsg.Text = ""
                Year = _objCommon.nNz(Right(Trim(ddlMonthYear.SelectedItem.Text.ToString), 4))
                Month = Left(MonthName(CType(ddlMonthYear.SelectedValue, Integer)), 3)
                filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & Month & Year '& "\TaxPaySlipWOPWD"
                If Directory.Exists(filepath) Then
                    FName = "_SalarySlipwithTaxDetails_WOPWD.pdf"
                End If
                HidEmpPdf.Value = HidEmpPdf.Value.Replace(",", FName & ",")
                If HidEmpPdf.Value.ToString.Trim <> "" Then
                    HidEmpPdf.Value = Left(HidEmpPdf.Value.ToString.Trim, Len(HidEmpPdf.Value.ToString.Trim) - 1)
                End If
                FileName = Split(HidEmpPdf.Value.ToString.Trim, ",")
                AddZipFiles(filepath, FileName)
                Response.Clear()
                Response.BufferOutput = False
                ' for large files...
                Dim c As System.Web.HttpContext = System.Web.HttpContext.Current
                'Dim ReadmeText As [String] = "Hello!" & vbLf & vbLf & "This is a README..." & DateTime.Now.ToString("G")
                Dim archiveName As String = [String].Format("SlipsWOPWD-{0}.zip", DateTime.Now.ToString("dd-MMM-yyyy"))
                Response.ContentType = "application/zip"
                Response.AddHeader("content-disposition", "filename=" & archiveName)
                Using zip As New ZipFile()
                    zip.AddFiles(filesToInclude, "PaySlipsWOPWD")
                    zip.Save(Response.OutputStream)
                    ' File.Delete(FileName(InStr))
                End Using
                'For index As Integer = 0 To filesToInclude.Count - 1
                '    File.Delete(Convert.ToString(filesToInclude(index)))
                'Next

                Response.End()

            Catch ex As Exception
                _objcommonExp.PublishError("Error in BtnLog_Click()", ex)
            End Try
        End Sub

        'Added by Debargha On 17 May 2024 For 'Please Wait' clickbait when Exporting Salary Register in Excel[Dynamic]
        Private Sub MakeZipFolderbyXml(ByVal _Ds As DataSet, ByVal _fileName As String, ByVal _ExcelName As String)
            Dim _sw As StreamWriter

            If System.IO.File.Exists(_fileName) Then
                System.IO.File.Delete(_fileName)
            End If

            _sw = New StreamWriter(_fileName)
            lblmsg.Text = ""
            ExportToExcelXML(_Ds, _sw, "D")
            _sw.Close()
            _sw.Dispose()

            hdfile.Value = _fileName + "~" + _ExcelName + "~" + "N"

            Dim popupScript As String = "<script language='javascript' type='text/javascript'>ShowDownload('" & hdfile.Value & "')</script>"
            'register the script
            ClientScript.RegisterStartupScript(GetType(String), "PopupScript", popupScript)
        End Sub

        'Created by Vishal Chauhan for progress bar'
        Private Sub CallReportAPIOnNewThread(ByVal requestBody As String, ByVal controller As String, ByVal AppPathStr As String, ByVal apiUrl As String, Optional ByVal CSVdynamic As String = "")
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Try
                'Dim apiUrl As String
                If (apiUrl Is Nothing OrElse apiUrl.Trim().Length < 10) Then
                    apiUrl = System.Configuration.ConfigurationManager.AppSettings("ApiRptExcel").ToString & controller
                Else
                    apiUrl = apiUrl & controller
                End If
                Dim jobResponse = PostAPICall(apiUrl & "/generate-file", requestBody)
                Dim jobResponseData = Newtonsoft.Json.JsonConvert.DeserializeObject(Of JobResponse)(jobResponse)
                JobUniqueId.Value = jobResponseData.JobId.ToString
                If CSVdynamic.Trim.ToUpper.Equals("CSVSALARYREGISTER") Then
                    Dim CommonPath As String = _objCommon.GetDirpath(Session("COMPCODE").ToString)
                    CommonPath = CommonPath.Replace("\", "\\")
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup9181", "CSVWithPGPProcessBar('" & AppPathStr & "','" & GetApiProcessType(DDLPaySlipType.SelectedValue) & "', '" & DDLPaySlipType.SelectedItem.Text.Replace("'", "") & "','" & IIf(DDLPaySlipType.SelectedValue.Trim.Equals("38") And ddllEncrType.SelectedValue.Trim.Equals("WP"), "Y", "N") & "','" & IIf(DDLPaySlipType.SelectedValue.Trim.Equals("38") And chkSFTP.Checked = True, "Y", "N") & "','" & CommonPath & "');", True)
                Else
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup9181", "OpenAttendanceProcessBar('" & AppPathStr & "','" & GetApiProcessType(DDLPaySlipType.SelectedValue) & "', '" & DDLPaySlipType.SelectedItem.Text.Replace("'", "") & "');", True)
                End If
            Catch ex As Exception
                _objcommonExp.PublishError("Error in CallAPIReport()", ex)
                Dim ErrorCaseparams(4) As SqlClient.SqlParameter
                ErrorCaseparams(0) = New SqlParameter("@userId", Session("UID").ToString)
                ErrorCaseparams(1) = New SqlParameter("@Process_Type", GetApiProcessType(DDLPaySlipType.SelectedValue))
                ErrorCaseparams(2) = New SqlParameter("@ErrorMsg", ex.Message)
                ErrorCaseparams(3) = New SqlParameter("@ActionType", "ErrorInExcel")
                ErrorCaseparams(4) = New SqlParameter("@BatchId", hdnBatchId.Value)
                _ObjData.GetDataTableProc("PaySP_ReportApi_ProcessBar", ErrorCaseparams)
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "hideModal908", "hideModal();", True)
            End Try
        End Sub


        'Created by Vishal Chauhan for Salary Register in Excel with HeavyExcel web.config tag call
        Private Sub CallReportAPIOnNewThreadHeavyExcel(ByVal requestBody As String, ByVal controller As String, ByVal AppPathStr As String, ByVal apiUrl As String, Optional ByVal CSVdynamic As String = "")
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Try
                'Dim apiUrl As String
                If (apiUrl Is Nothing OrElse apiUrl.Trim().Length < 10) Then
                    apiUrl = System.Configuration.ConfigurationManager.AppSettings("HeavyExcel").ToString & controller
                Else
                    apiUrl = apiUrl & controller
                End If
                Dim jobResponse = PostAPICall(apiUrl & "/generate-file", requestBody)
                Dim jobResponseData = Newtonsoft.Json.JsonConvert.DeserializeObject(Of JobResponse)(jobResponse)
                JobUniqueId.Value = jobResponseData.JobId.ToString
                If CSVdynamic.Trim.ToUpper.Equals("CSVSALARYREGISTER") Then
                    Dim CommonPath As String = _objCommon.GetDirpath(Session("COMPCODE").ToString)
                    CommonPath = CommonPath.Replace("\", "\\")
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup9181", "CSVWithPGPProcessBar('" & AppPathStr & "','" & GetApiProcessType(DDLPaySlipType.SelectedValue) & "', '" & DDLPaySlipType.SelectedItem.Text.Replace("'", "") & "','" & IIf(DDLPaySlipType.SelectedValue.Trim.Equals("38") And ddllEncrType.SelectedValue.Trim.Equals("WP"), "Y", "N") & "','" & IIf(DDLPaySlipType.SelectedValue.Trim.Equals("38") And chkSFTP.Checked = True, "Y", "N") & "','" & CommonPath & "');", True)
                Else
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup9181", "OpenAttendanceProcessBar('" & AppPathStr & "','" & GetApiProcessType(DDLPaySlipType.SelectedValue) & "', '" & DDLPaySlipType.SelectedItem.Text.Replace("'", "") & "');", True)
                End If
            Catch ex As Exception
                _objcommonExp.PublishError("Error in CallAPIReport()", ex)
                Dim ErrorCaseparams(4) As SqlClient.SqlParameter
                ErrorCaseparams(0) = New SqlParameter("@userId", Session("UID").ToString)
                ErrorCaseparams(1) = New SqlParameter("@Process_Type", GetApiProcessType(DDLPaySlipType.SelectedValue))
                ErrorCaseparams(2) = New SqlParameter("@ErrorMsg", ex.Message)
                ErrorCaseparams(3) = New SqlParameter("@ActionType", "ErrorInExcel")
                ErrorCaseparams(4) = New SqlParameter("@BatchId", hdnBatchId.Value)
                _ObjData.GetDataTableProc("PaySP_ReportApi_ProcessBar", ErrorCaseparams)
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "hideModal908", "hideModal();", True)
            End Try
        End Sub

        Private Sub CallAPIReport(ByVal requestBody As String, ByVal controller As String)
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Dim filePath = Server.MapPath(Request.ApplicationPath).ToString & "\" & Session("COMPCODE").ToString & "\TempExcelFiles\"

            If Not Directory.Exists(filePath) Then
                Directory.CreateDirectory(filePath)
            End If
            Session("ServerFilePath") = filePath
            Try
                Dim apiUrl As String
                apiUrl = System.Configuration.ConfigurationManager.AppSettings("APIReport").ToString & controller

                Dim jobResponse = PostAPICall(apiUrl & "/generate-file", requestBody)
                Dim jobResponseData = Newtonsoft.Json.JsonConvert.DeserializeObject(Of JobResponse)(jobResponse)
                JobUniqueId.Value = jobResponseData.JobId.ToString
                Session("JobUniqueId") = jobResponseData.JobId.ToString

                Dim jScript As String = "<script language='javascript' type='text/javascript'>ChkRptDownloadStatus('" + controller + "','" + JobUniqueId.Value.ToString + "')</script>"
                'register the script
                ClientScript.RegisterStartupScript(Me.GetType(), "Job", jScript)
            Catch ex As Exception
                _objcommonExp.PublishError("Error in CallAPIReport()", ex)
                Console.WriteLine(ex)
            End Try
        End Sub

        Private Function PostAPICall(url As String, postData As String) As String
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Try
                ServicePointManager.ServerCertificateValidationCallback = Function(sender, cert, chain, sslPolicyErrors) True

                Dim request As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
                request.Method = "POST"
                request.ContentType = "application/json"
                request.Timeout = 5000000

                If postData IsNot Nothing Then
                    Dim data = Encoding.UTF8.GetBytes(postData)
                    request.ContentLength = data.Length
                    Using requestStream = request.GetRequestStream()
                        requestStream.Write(data, 0, data.Length)
                    End Using
                End If
                Using response As HttpWebResponse = CType(Request.GetResponse(), HttpWebResponse)
                    Using responseStream = response.GetResponseStream()
                        Using reader = New StreamReader(responseStream)
                            Return reader.ReadToEnd()
                        End Using
                    End Using
                End Using
            Catch ex As WebException
                If ex.Response IsNot Nothing Then
                    Using resp = CType(ex.Response, HttpWebResponse)
                        Using respStream = resp.GetResponseStream()
                            Using reader = New StreamReader(respStream)
                                Dim errorDetails = reader.ReadToEnd()
                                Console.WriteLine("Response status: " & resp.StatusCode)
                                Console.WriteLine("Error body: " & errorDetails)
                            End Using
                        End Using
                    End Using
                End If
                _objcommonExp.PublishError("Error in PostAPICall()", ex)
                Return Nothing
            End Try
        End Function

        Protected Function GetAppPath() As String
            Dim _array() As String, _AppPath As String = HttpRuntime.AppDomainAppVirtualPath.ToString
            _array = Split(_AppPath, "/")
            _AppPath = _array(_array.Length - 1)
            Return _AppPath
        End Function

        Private Sub imgpdfprocess_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgpdfprocess.Click
            IsShowlnkIcon.Value = "Y"
            CheckProcessLocked()
        End Sub
        Private Function SplitTimes(ByVal timeInSecs As Double) As String
            Dim hours As Double = Math.Floor(timeInSecs / 3600)
            Dim minutes As Double = Math.Floor(timeInSecs / 60) - (hours * 60)
            Dim seconds As Double = timeInSecs - (hours * 3600) - (minutes * 60)
            Dim hs As String = " hour"
            Dim ms As String = " minute"
            Dim ss As String = " second"

            If hours <> 1 Then
                hs &= "s"
            End If
            If minutes <> 1 Then
                ms &= "s"
            End If
            If seconds <> 1 Then
                ss &= "s"
            End If

            Dim time As String = String.Empty
            If hours <> 0 Then
                time &= hours.ToString() & hs & ", "
            End If
            If minutes <> 0 Then
                time &= minutes.ToString() & ms & ", "
            End If
            time &= seconds.ToString() & ss
            Return time
        End Function
        Private Sub CheckProcessLocked()
            Dim _ArrParam(1) As SqlParameter, dst As New DataTable, _empCodes As String = ""
            Dim gcs_service As Integer = 0

            Try

                hdnAlreadyRunRptName.Value = ""
                Dim _array() As String, _AppPath As String = HttpRuntime.AppDomainAppVirtualPath.ToString
                _array = Split(_AppPath, "/")
                _AppPath = _array(_array.Length - 1)
                HidAppPath.Value = _AppPath
                _ArrParam(0) = New SqlClient.SqlParameter("@flag", "L")
                _ArrParam(1) = New SqlClient.SqlParameter("@UserID", Session("UID"))
                dst = _ObjData.GetDataTableProc("PaySP_ReportProcess_ProcessBar", _ArrParam)
                If (dst.Rows.Count > 0 AndAlso dst.Rows(0)("IsAbleToStart").ToString = "0") Then
                    process_status_id.Value = ""
                    'divSocialExcel.Visible = True
                    'lblProcessStatus.Text = "" & dst.Rows(0)("Msg").ToString
                    hdnAlreadyRunRptName.Value = dst.Rows(0)("Msg").ToString
                    ScriptManager.RegisterStartupScript(Me.Page, GetType(Page), "refreshprocessstatus", "ShowPayslipsLockSummaryDetails('" & dst.Rows(0)("Process_Type").ToString.ToUpper & "')", True)
                    Exit Sub
                End If
            Catch ex As Exception
                _objcommonExp.PublishError("Error in CheckProcessLocked()", ex)
            End Try


            Try
                hdnAlreadyRunRptName.Value = ""
                Dim _array() As String, _AppPath As String = HttpRuntime.AppDomainAppVirtualPath.ToString
                _array = Split(_AppPath, "/")
                _AppPath = _array(_array.Length - 1)
                HidAppPath.Value = _AppPath
                _ArrParam(0) = New SqlClient.SqlParameter("@flag", "L")
                _ArrParam(1) = New SqlClient.SqlParameter("@UserID", Session("UID"))
                dst = _ObjData.GetDataTableProc("PaySP_GetGcsSalarySlip", _ArrParam)
                If (dst.Rows.Count > 0 AndAlso dst.Rows(0)("IsAbleToStart").ToString = "0") Then

                    Dim slip As String = "Pay Slip is already publishing. Please wait till the completion."
                    If dst.Rows(0)("Process_Type").ToString = "TAXSLIP" Then
                        slip = "Salary Slip With Tax Details is already publishing. Please wait till the completion."
                    ElseIf dst.Rows(0)("Process_Type").ToString = "SLIPWOLVE" Then
                        slip = "Pay Slip Without Leave Details is already publishing. Please wait till the completion."
                    ElseIf dst.Rows(0)("Process_Type").ToString = "FORCAST" Then
                        slip = "Tax sheet-forecast is already publishing. Please wait till the completion."
                    ElseIf dst.Rows(0)("Process_Type").ToString = "SLIPTDSV" Then
                        slip = "TDS Estimation is already publishing. Please wait till the completion."
                    ElseIf dst.Rows(0)("Process_Type").ToString = "SLIPWILVE" Then
                        slip = "Pay Slip With Leave Details is already publishing. Please wait till the completion."
                    End If
                    hdnAlreadyRunRptName.Value = slip
                    If dst.Rows(0).Table.Columns.Contains("ID") Then
                        process_status_id.Value = dst.Rows(0)("ID").ToString
                    End If
                    ScriptManager.RegisterStartupScript(Me.Page, GetType(Page), "refreshprocessstatus", "ShowPayslipsLockSummaryDetails('" & dst.Rows(0)("Process_Type").ToString.ToUpper & "')", True)
                    Exit Sub

                End If
            Catch ex As Exception
                _objcommonExp.PublishError("Error in CheckProcessLocked()", ex)
            End Try


        End Sub
        Protected Sub btnPublishedPDF_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPublishedPDF.Click
            Dim _strJava As New StringBuilder, filepath As String, FileName() As String = Nothing, FName As String = "", PublishedPdfEmpCodes As String = "" _
            , _msg As New List(Of PayrollUtility.UserMessage), arrCd(21) As String, _dt As DataTable _
            , allFileNames As String(), rowsToRemove As New List(Of DataRow), mmStr As String = Left(MonthName(CType(ddlMonthYear.SelectedValue.ToString, Integer)), 3),
            yyyy As String = Right(ddlMonthYear.SelectedItem.Text.ToString, 4), mmVal As String = ddlMonthYear.SelectedValue, EmpCodeSearch As String = ""
            'Excel Process locking validation checking
            'CheckExcelProcessbarAlreadyProcessing()
            'If (lblProcessStatusExcel.Text <> "") Then
            '    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = lblProcessStatusExcel.Text})
            '    _objCommon.ShowMessage(_msg)
            '    Exit Sub
            'End If
            Try
                CheckProcessLocked()
                If (hdnAlreadyRunRptName.Value.Trim().Length > 1) Then
                    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = hdnAlreadyRunRptName.Value})
                    _objCommon.ShowMessage(_msg)
                    Exit Sub
                End If
                Dim gcs_service As Integer = 0
                Dim _slip As String = ""
                Try
                    Dim _mm As String = "", _yyyy As String = ""
                    If ddlMonthYear.SelectedItem IsNot Nothing Then
                        If DdlreportType.SelectedValue.ToString.ToUpper = "R" Then
                            _slip = "TAXSLIP"
                        ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "S" Then
                            _slip = "SLIPWOLVE"
                        ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "57" Then
                            _slip = "FORCAST"
                        ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "T" Then
                            _slip = "SLIPTDSV"
                        ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "SL" Then
                            _slip = "SLIPWILVE"
                        End If
                        _mm = CType(ddlMonthYear.SelectedValue.ToString, Integer)
                        _yyyy = Right(ddlMonthYear.SelectedItem.Text.ToString, 4)
                    End If
                    If gcs_service_obj.IsPaySlipAllowedByMonthYear(Session("compCode"), _slip, _mm, _yyyy) Then
                        gcs_service = 1
                    End If
                Catch ex As Exception

                End Try


                If gcs_service = 0 Then
                    lblProcessBarMsg.Text = ""
                    Session("PublishedTDSEmpCode") = Nothing
                    Session("PublishedTaxSlipEmpCode") = Nothing
                    Session("PublishedForecastEmpCode") = Nothing
                    Session("PublishedSLIPWOLVEEmpCode") = Nothing
                    If RblNoSearch.SelectedValue = "S" Then
                        For counter = 0 To DgPayslip.Items.Count - 1
                            If CType(DgPayslip.Items(counter).FindControl("chkEmpHold"), CheckBox).Checked = True Then
                                EmpCodeSearch = EmpCodeSearch + DgPayslip.Items(counter).Cells(1).Text.ToString + ","
                            End If
                        Next
                        _dt = ReturnDsSearch("", "", EmpCodeSearch).Tables(0)
                    Else
                        _dt = ReturnDsSearch().Tables(0)
                    End If
                    If ddlRepIn.SelectedValue.ToString.ToUpper = "P" AndAlso rbtnmail.Checked Then
                        If DdlreportType.SelectedValue.ToString.ToUpper.Equals("T") Then
                            filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & mmStr & yyyy & "\TDSEstimationSlip\"
                            FName = "_TDSEstimationSlip.pdf"
                        ElseIf DdlreportType.SelectedValue.ToString.ToUpper.Equals("R") Then
                            filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & mmStr & yyyy & "\TaxPaySlip\"
                            If HidPdfName.Value.ToString <> "" Then
                                FName = "_" & HidPdfName.Value & ".pdf"
                            Else
                                FName = "_SalarySlipwithTaxDetails.pdf"
                            End If
                        ElseIf DdlreportType.SelectedValue.ToString.ToUpper.Equals("S") Then
                            filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & mmStr & yyyy & "\LeaveWoPaySlip\"
                            FName = "_SalarySlipInclude.pdf"
                        ElseIf DdlreportType.SelectedValue.ToString.Equals("57") Then
                            filepath = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\" & mmStr & yyyy & "\YTDTaxComputationSheet\"
                            If (mmVal.Length > 1) Then
                                FName = "_estax_" & yyyy & mmVal & ".pdf"
                            Else
                                FName = "_estax_" & yyyy & "0" & mmVal & ".pdf"
                            End If
                        End If
                    End If
                    If Directory.Exists(filepath) Then
                        allFileNames = Directory.GetFiles(filepath, "*" & FName, SearchOption.AllDirectories)
                        For Each row As DataRow In _dt.Rows
                            Dim empCode As String = row("Fk_emp_code").ToString()
                            Dim fileExists As Boolean = False
                            For Each fPath As String In allFileNames
                                Dim flName As String = Path.GetFileName(fPath)
                                If flName.StartsWith(empCode & "_") AndAlso flName.EndsWith(FName) Then
                                    fileExists = True
                                    Exit For
                                End If
                            Next
                            If Not fileExists Then
                                rowsToRemove.Add(row)
                            End If
                        Next
                        For Each rowToRemove As DataRow In rowsToRemove
                            _dt.Rows.Remove(rowToRemove)
                        Next
                        If (_dt.Rows.Count = 0) Then
                            _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "Payslip(s) are not published as of now or you are not authorised to view payslip as per user rights !"})
                            _objCommon.ShowMessage(_msg)
                            Exit Sub
                        End If
                        HidEmpPdf.Value = String.Join(",",
                                                  _dt.AsEnumerable().
                                                  Where(Function(row) Not IsDBNull(row("Fk_emp_code"))).
                                                  Select(Function(row) row("Fk_emp_code").ToString()))
                        If (HidEmpPdf.Value.ToString.Last() <> ",") Then
                            HidEmpPdf.Value += ","
                        End If
                        If ddlRepIn.SelectedValue.ToString.ToUpper = "P" AndAlso rbtnmail.Checked Then
                            If DdlreportType.SelectedValue.ToString.ToUpper.Equals("T") Then
                                Session("PublishedTDSEmpCode") = HidEmpPdf.Value
                                hdfile.Value = filepath.Replace("\", "\\") & "~" & FName & "~Progessbar"
                            ElseIf DdlreportType.SelectedValue.ToString.ToUpper.Equals("R") Then
                                Session("PublishedTaxSlipEmpCode") = HidEmpPdf.Value
                                hdfile.Value = filepath.Replace("\", "\\") & "~" & FName & "~Progessbar"
                            ElseIf DdlreportType.SelectedValue.ToString.ToUpper.Equals("S") Then
                                Session("PublishedSLIPWOLVEEmpCode") = HidEmpPdf.Value
                                hdfile.Value = filepath.Replace("\", "\\") & "~" & FName & "~Progessbar"
                            ElseIf DdlreportType.SelectedValue.ToString.Equals("57") Then
                                Session("PublishedForecastEmpCode") = HidEmpPdf.Value
                                hdfile.Value = filepath.Replace("\", "\\") & "~" & FName & "~Progessbar"
                            End If
                        End If
                        Dim popupScript As String = "<script language='javascript' type='text/javascript'>OpenDownloadDiaog('" & hdfile.Value & "')</script>"
                        ClientScript.RegisterStartupScript(GetType(String), "PopupScriptPayslipTDS", popupScript)
                    Else
                        _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "Payslip(s) are not published as of now or you are not authorised to view payslip as per user rights !"})
                        _objCommon.ShowMessage(_msg)
                        Exit Sub
                    End If

                Else
                    If ddlRepIn.SelectedValue.ToString.ToUpper = "P" Then
                        LnkPDF.Style.Value = "display:none;"
                        LnkPDFWOPWD.Style.Value = "display:none;"
                        download_pdf2.Style.Value = "border:0;cursor:pointer;"
                        Dim id As String = process_status_id.Value
                        Dim gcp_path As String = "Payroll/", month_val As String = ""
                        Dim slip As String = "", slipType As String = ""
                        'Dim Month As String = Left(MonthName(CType(DdlMonthYear.SelectedValue.ToString, Integer)), 3)
                        'Dim Year As String = Right(Trim(DdlMonthYear.SelectedItem.Text.ToString), 4)).ToString
                        month_val = ddlMonthYear.SelectedValue.ToString
                        If DdlreportType.SelectedValue.ToString.ToUpper = "R" Then
                            slip = "TAXSLIP"
                            slipType = "TaxPaySlip"
                            gcp_path = gcp_path & Session("COMPCODE").ToString & "/PDFFiles/" & mmStr & yyyy & "/" & "TaxPaySlip"
                        ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "S" Then
                            slip = "LeaveWoPaySlip"
                            slipType = "LeaveWoPaySlip"
                            gcp_path = gcp_path & Session("COMPCODE").ToString & "/PDFFiles/" & mmStr & yyyy & "/" & "LeaveWoPaySlip"
                        ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "57" Then
                            slipType = "YTDTaxComputationSheet"
                            slip = "FORCAST"
                            gcp_path = gcp_path & Session("COMPCODE").ToString & "/PDFFiles/" & mmStr & yyyy & "/" & "YTDTaxComputationSheet"
                        ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "T" Then
                            slipType = "TDSEstimationSlip"
                            slip = "SLIPTDSV"
                            gcp_path = gcp_path & Session("COMPCODE").ToString & "/PDFFiles/" & mmStr & yyyy & "/" & "TDSEstimationSlip"
                        ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "SL" Then
                            slipType = ""
                            slip = "SLIPWILVE"
                            gcp_path = gcp_path & Session("COMPCODE").ToString & "/PDFFiles/" & mmStr & yyyy & "/"
                        End If

                        If month_val.Length = 1 Then
                            month_val = "0" & month_val
                        End If
                        Dim dt_temp As New DataTable
                        If RblNoSearch.SelectedValue = "S" Then
                            For counter = 0 To DgPayslip.Items.Count - 1
                                If CType(DgPayslip.Items(counter).FindControl("chkEmpHold"), CheckBox).Checked = True Then
                                    EmpCodeSearch = EmpCodeSearch + DgPayslip.Items(counter).Cells(1).Text.ToString + ","
                                End If
                            Next
                            _dt = ReturnDsSearch("", "", EmpCodeSearch).Tables(0)
                        Else
                            _dt = ReturnDsSearch().Tables(0)
                        End If

                        If _dt.Rows.Count = 0 Then
                            _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "No record(s) found according to the selection criteria"})
                            _objCommon.ShowMessage(_msg)
                            Exit Sub
                        End If
                        Dim file_path As New List(Of String)
                        Dim all_path As String = "", fileExt As String = ""
                        If DdlreportType.SelectedValue.ToString.ToUpper = "57" Then
                            fileExt = "_estax_" & yyyy & month_val
                        ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "S" Then
                            fileExt = "_SalarySlipInclude"
                        ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "R" Then
                            fileExt = "_SalarySlipwithTaxDetails"
                        ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "T" Then
                            fileExt = "_TDSEstimationSlip"
                        ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "SL" Then
                            fileExt = "_SalarySlip"
                        End If

                        For Each row As DataRow In _dt.Rows
                            Dim empCode As String = row("Fk_emp_code").ToString()
                            If DdlreportType.SelectedValue.ToString.ToUpper = "57" Then
                                all_path &= empCode & ","
                            ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "S" Then
                                all_path &= empCode & ","
                            ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "R" Then
                                all_path &= empCode & ","
                            ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "T" Then
                                all_path &= empCode & ","
                            ElseIf DdlreportType.SelectedValue.ToString.ToUpper = "SL" Then
                                all_path &= empCode & ","
                            End If
                        Next
                        all_path = all_path.TrimEnd(","c)



                        Dim java_url As String = ""
                        If Not System.Configuration.ConfigurationManager.AppSettings("JavaServiceDomain") Is Nothing Then
                            If System.Configuration.ConfigurationManager.AppSettings("JavaServiceDomain").ToString <> "" Then
                                java_url = System.Configuration.ConfigurationManager.AppSettings("JavaServiceDomain").ToString
                            End If
                        End If
                        'Dim postDataEscaped As String = postData.Replace("\", "\\").Replace("'", "\'")
                        Dim download_pdf_url As String = ""
                        If id.Length > 0 Then
                            download_pdf_url = java_url & "/api/payslip/download-zip"
                        Else
                            download_pdf1.Style.Value = "display:none;border:0;cursor:pointer;"
                            download_pdf2.Style.Value = "display:none;border:0;cursor:pointer;"
                        End If

                        java_url = java_url & "/api/payslip/download-zip"

                        ' Check if there is already a process running
                        Dim _ArrParam(1) As SqlParameter, dst As New DataTable
                        Dim _array() As String, _AppPath As String = HttpRuntime.AppDomainAppVirtualPath.ToString
                        _ArrParam(0) = New SqlClient.SqlParameter("@flag", "L")
                        _ArrParam(1) = New SqlClient.SqlParameter("@UserID", Session("UID"))

                        dst = _ObjData.GetDataTableProc("PaySP_GetGcsSalarySlip", _ArrParam)
                        If (dst.Rows.Count > 0 AndAlso dst.Rows(0)("IsAbleToStart").ToString = "0") Then
                            Dim slip_name As String = "Pay Slip is already publishing. Please wait till the completion."
                            If dst.Rows(0)("Process_Type").ToString = "TAXSLIP" Then
                                dst.Rows(0)("Msg") = "Salary Slip With Tax Details is already publishing. Please wait till the completion."
                            ElseIf dst.Rows(0)("Process_Type").ToString = "SLIPWOLVE" Then
                                dst.Rows(0)("Msg") = "Pay Slip Without Leave Details is already publishing. Please wait till the completion."
                            ElseIf dst.Rows(0)("Process_Type").ToString = "FORCAST" Then
                                dst.Rows(0)("Msg") = "Tax sheet-forecast is already publishing. Please wait till the completion."
                            ElseIf dst.Rows(0)("Process_Type").ToString = "SLIPTDSV" Then
                                dst.Rows(0)("Msg") = "TDS Estimation is already publishing. Please wait till the completion."
                            ElseIf dst.Rows(0)("Process_Type").ToString = "SLIPWILVE" Then
                                dst.Rows(0)("Msg") = "Pay Slip With Leave Details already publishing. Please wait till the completion."
                            Else
                                dst.Rows(0)("Msg") = slip_name
                            End If
                        End If
                        If (dst.Rows.Count > 0 AndAlso dst.Rows(0)("IsAbleToStart").ToString = "0") Then
                            _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = dst.Rows(0)("Msg")})
                            _objCommon.ShowMessage(_msg)
                            Exit Sub
                        End If
                        ' Exit sub if there is a process running

                        '(url, filePath, compCode, fileExt, slipType, mm, yyyy )
                        ScriptManager.RegisterStartupScript(Me.Page,
                                                            GetType(Page),
                                                            "openDownloadWindow",
                                                            "openDownloadWindow('" & java_url & "', '" & all_path & "', '" & Session("compCode") & "', '" & fileExt &
                                                             "', '" & slipType & "', '" & mmStr & "', '" & yyyy & "')",
                                                            True)

                    Else
                        _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "No record(s) found"})
                        _objCommon.ShowMessage(_msg)
                        Exit Sub

                    End If
                End If

            Catch ex As Exception
                _objcommonExp.PublishError("Error in btnPublishedPDF_Click()", ex)
                _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "Payslip(s) are not published as of now or you are not authorised to view payslip as per user rights !!"})
                _objCommon.ShowMessage(_msg)
            End Try
        End Sub


        Private Function GetPaySlipProcessType(ByVal PaySlipType As String) As String
            If (PaySlipType.ToUpper.Trim = "R") Then
                Return "TAXSLIP"
            ElseIf (PaySlipType.ToUpper.Trim = "S") Then
                Return "SLIPWOLVE"
            ElseIf (PaySlipType.ToUpper.Trim = "T") Then
                Return "SLIPTDSV"
            ElseIf (PaySlipType = "57") Then
                Return "FORCAST"
            Else
                Return ""
            End If
        End Function
        Private Function GetApiProcessType(ByVal PaySlipType As String) As String
            If (PaySlipType = "38") And ddlrepformat.SelectedValue.Trim.ToUpper.Equals("CSV") Then
                Return "DYNAMICSALREGCSV"
            ElseIf (PaySlipType = "38") Then
                Return "DYNAMICSALREG"
            ElseIf (PaySlipType = "21") Then
                Return "SALARYREGISTER"
            Else
                Return ""
            End If
        End Function

        Protected Sub btnProgressbarExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProgressbarExcel.Click
            Dim AppPathStr As String = HttpRuntime.AppDomainAppVirtualPath.ToString
            Dim _array() As String
            _array = Split(AppPathStr, "/")
            AppPathStr = _array(_array.Length - 1)
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopupprogress21", "OpenAttendanceProcessBar('" & AppPathStr & "','" & GetApiProcessType(hdnAlreadyRunRptId.Value) & "', '" & hdnAlreadyRunRptName.Value.Replace("'", "") & "');", True)
            lblProcessStatusExcel.Text = ""
            divSocialExcel.Visible = False
        End Sub
        Private Sub CheckExcelProcessbarAlreadyProcessing()
            Try
                Dim arprm(2) As SqlClient.SqlParameter
                arprm(0) = New SqlClient.SqlParameter("@UserID", Session("UID").ToString)
                arprm(1) = New SqlClient.SqlParameter("@Process_Type", GetApiProcessType(DDLPaySlipType.SelectedValue))
                arprm(2) = New SqlClient.SqlParameter("@ActionType", "isalreadyrun")
                Dim _dt As DataTable = _ObjData.GetDataTableProc("PaySP_ReportApi_ProcessBar", arprm)
                If (_dt.Rows.Count > 0 AndAlso _dt.Rows(0)("IsAbleToStart").ToString = "0") Then
                    divSocialExcel.Visible = True
                    hdnAlreadyRunRptId.Value = _dt.Rows(0)("DdlRptId").ToString
                    hdnAlreadyRunRptName.Value = _dt.Rows(0)("RptName").ToString
                    lblProcessStatusExcel.Text = "" & _dt.Rows(0)("RptName").ToString & " is already processing. Please wait till the completion."
                    If (_dt.Rows(0)("StartedByUserId").ToString.ToUpper <> Session("UID").ToUpper.ToString) Then
                        btnProgressbarExcel.Visible = False
                    Else
                        btnProgressbarExcel.Visible = True
                    End If
                    ScriptManager.RegisterStartupScript(Me.Page, GetType(Page), "refreshExeclProcessStatus123", "ShowExcelLockSummaryDetails('" & _dt.Rows(0)("Process_Type").ToString.ToUpper & "')", True)
                Else
                    hdnAlreadyRunRptName.Value = ""
                    hdnAlreadyRunRptId.Value = ""
                    lblProcessStatusExcel.Text = ""
                    divSocialExcel.Visible = False
                End If
            Catch ex As Exception
                _objcommonExp.PublishError("Error in CheckExcelProcessbarAlreadyProcessing()", ex)
            End Try
        End Sub


        Private Class JobResponse
            Public Property JobId As String
            Public Property StatusUrl As String
        End Class

        Protected Sub Ddlrepformat_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlrepformat.SelectedIndexChanged
            trEncrType.Style.Value = "display:none"
            trSftpID.Style.Value = "display:none"
            trFileName.Style.Value = "display:none"
            txtrptName.Text = ""
            chkSFTP.Checked = False
            ddllEncrType.SelectedIndex = -1
            btnPreview.Text = "Export to Excel"
            ShowPGP()
            If ddlrepformat.SelectedValue.ToString.ToUpper.Trim = "CSV" Then
                trEncrType.Style.Value = "display:"
                trSftpID.Style.Value = "display:"
                trFileName.Style.Value = "display:"
                btnPreview.Text = "Export to CSV"
            End If
            ShowSFTP()
        End Sub

        Private Sub ShowSFTP()
            Dim Ds As New DataSet, arrparam(0) As SqlClient.SqlParameter
            arrparam(0) = New SqlParameter("@id", DDLPaySlipType.SelectedValue)
            Ds = _ObjData.GetDsetProc("paysp_reportdetails", arrparam)
            If Ds.Tables.Count > 1 Then
                If Ds.Tables(1).Rows(0)("sftConfig").ToString().Equals("Y") And ddlrepformat.SelectedValue.ToString.ToUpper.Trim = "CSV" Then
                    trSftpID.Style.Value = "display:"
                Else
                    trSftpID.Style.Value = "display:none"
                End If
            Else
                trSftpID.Style.Value = "display:none"
            End If
        End Sub

        Private Sub ShowPGP()
            Dim ds As New DataSet, _sqlParam(1) As SqlClient.SqlParameter
            _sqlParam(0) = New SqlClient.SqlParameter("@ReportId", "38")
            _sqlParam(1) = New SqlClient.SqlParameter("@RptType", "DYNREG")
            ds = _ObjData.GetDsetProc("Paysp_MstPGPEncryptionConfig_GetEncrKey", _sqlParam)
            If ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
                hdnPGP.Value = "Y"
            Else
                hdnPGP.Value = "N"
            End If
        End Sub

        'Public Function ExportToCSV_SalRegister(ds As DataSet, salReg As SalRegModel, configuration As IConfiguration, _objConfig As IConfigProvider, _filename As String) As (String, String)
        '    Dim mName As String = "ExportToCSV_SalRegister"
        '    Dim TotalRowCount As Integer = ds.Tables(0).Rows.Count
        '    Dim csvBuilder As New StringBuilder()

        '    Try
        '        Dim loopdivider As Integer = _objConfig.GetLoopDividerToSaveLogInDB(ds.Tables(0).Rows.Count)
        '        Dim loopSavedivider As Integer = _objConfig.GetLoopDividerToSaveLogInTxt(ds.Tables(0).Rows.Count)

        '        Dim _path As String = configuration.GetConnectionString("ExcelFilePath")
        '        _path = Path.Combine(_path, salReg.DomainName, "TempExcelFiles", FileFolderName.CommonReportFolder)

        '        Dim _path As String = _objCommon.GetDirpath(Session("COMPCODE").ToString) & "\TempExcelFiles"

        '        If Directory.Exists(_path) Then
        '            Dim timeLimit As TimeSpan = TimeSpan.FromHours(12)
        '            Dim now As DateTime = DateTime.Now
        '            For Each file As String In Directory.GetFiles(_path)
        '                If now - System.IO.File.GetCreationTime(file) >= timeLimit Then
        '                    System.IO.File.Delete(file)
        '                End If
        '            Next
        '        End If

        '        If Not Directory.Exists(_path) Then
        '            Directory.CreateDirectory(_path)
        '        End If

        '        Dim recordsPerSheet As Integer = 1000000
        '        Dim dtconfig As DataTable = _objConfig.GetConfigDetails("configs", salReg.UserId, "PaySP_SalaryRegInExcel_Dynamic", salReg.DomainName)
        '        If dtconfig.Rows.Count > 0 AndAlso Not IsDBNull(dtconfig.Rows(0)("RecordsPerSheetCSV")) Then
        '            Dim val As Integer = Convert.ToInt32(dtconfig.Rows(0)("RecordsPerSheetCSV"))
        '            If val > 1000 Then
        '                recordsPerSheet = val
        '            End If
        '        End If

        '        Dim TotalFilesTobeCreated As Integer = If(TotalRowCount < recordsPerSheet, 1, Math.Ceiling(TotalRowCount / recordsPerSheet))
        '        Dim totalrowsprocessed As Integer = 0

        '        For i As Integer = 1 To TotalFilesTobeCreated
        '            Dim NextSheetLoopCounter As Integer = totalrowsprocessed
        '            Dim SheetRowsProcessed As Integer = 0
        '            Dim tempExcelFilePath As String = Path.Combine(_path, $"{_filename}_{i}.csv")

        '            Using writer As New StreamWriter(tempExcelFilePath)
        '                ' Write header
        '                For col As Integer = 0 To ds.Tables(0).Columns.Count - 1
        '                    writer.Write(ds.Tables(0).Columns(col).ColumnName)
        '                    If col < ds.Tables(0).Columns.Count - 1 Then writer.Write(",")
        '                Next
        '                writer.WriteLine()

        '                ' Write data rows
        '                For row As Integer = NextSheetLoopCounter To TotalRowCount - 1
        '                    If SheetRowsProcessed > recordsPerSheet Then Exit For

        '                    totalrowsprocessed += 1
        '                    SheetRowsProcessed += 1

        '                    For col As Integer = 0 To ds.Tables(0).Columns.Count - 1
        '                        Dim tableData = ds.Tables(0).Rows(row)(col)
        '                        Dim CSVstring As String = tableData.ToString().Trim()
        '                        Dim CSVnumber As Double
        '                        Dim isNumeric As Boolean = Double.TryParse(CSVstring, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, CSVnumber)

        '                        If TypeOf tableData Is DBNull Then
        '                            csvBuilder.Append(",")
        '                        ElseIf isNumeric AndAlso (ds.Tables(0).Columns(col).DataType = GetType(Decimal) OrElse ds.Tables(0).Columns(col).DataType = GetType(Long)) Then
        '                            csvBuilder.Append(CSVnumber.ToString(System.Globalization.CultureInfo.InvariantCulture) & ",")
        '                        Else
        '                            csvBuilder.Append("""" & CSVstring.Replace("""", """""") & """,")
        '                        End If
        '                    Next

        '                    csvBuilder.Length -= 1
        '                    writer.WriteLine(csvBuilder.ToString())
        '                    csvBuilder.Clear()

        '                    If row Mod loopSavedivider = 0 Then writer.Flush()
        '                Next
        '            End Using
        '        Next

        '        Dim finalFileInfo As New FileInfo(Path.Combine(_path, $"{_filename}.csv"))
        '        Dim fileSize As Double = Math.Round(finalFileInfo.Length / 1024.0)

        '        Return (_path, fileSize.ToString())
        '    Catch ex As Exception

        '    End Try
        'End Function

        Private Function GenerateCsvFromDataTable(dt As DataTable, csvFilePath As String) As String
            Try
                ' Ensure directory exists
                Dim folderPath As String = Path.GetDirectoryName(csvFilePath)
                If Not Directory.Exists(folderPath) Then
                    Directory.CreateDirectory(folderPath)
                End If

                Using sw As New StreamWriter(csvFilePath, False, Encoding.UTF8)
                    ' Write column headers
                    Dim columnNames As IEnumerable(Of String) = dt.Columns.Cast(Of DataColumn)().
                Select(Function(column) """" & column.ColumnName.Replace("""", """""") & """")
                    sw.WriteLine(String.Join(",", columnNames))

                    ' Write data rows
                    For Each row As DataRow In dt.Rows
                        Dim fields As IEnumerable(Of String) = row.ItemArray.Select(Function(field) FormatCsvValue(field))
                        sw.WriteLine(String.Join(",", fields))
                    Next
                End Using

                Return csvFilePath
            Catch ex As Exception
                Throw New Exception("Error generating CSV: " & ex.Message)
            End Try
        End Function

        Private Function FormatCsvValue(value As Object) As String
            If value Is Nothing OrElse IsDBNull(value) Then
                Return """"""
            End If

            If TypeOf value Is DateTime Then
                Return """" & DirectCast(value, DateTime).ToString("dd-MMM-yyyy HH:mm:ss") & """"
            ElseIf TypeOf value Is Boolean Then
                Return If(CBool(value), """Yes""", """No""")
            Else
                Dim strValue As String = value.ToString().Replace("""", """""")
                Return """" & strValue & """"
            End If
        End Function

        Private Sub WriteFileField(stream As Stream, fieldName As String, filePath As String, contentType As String, boundary As String, encoding As Encoding)
            Dim fileName As String = Path.GetFileName(filePath)
            Dim header As String = "--" & boundary & vbCrLf &
                                   "Content-Disposition: form-data; name=""" & fieldName & """; filename=""" & fileName & """" & vbCrLf &
                                   "Content-Type: " & contentType & vbCrLf & vbCrLf
            Dim headerBytes As Byte() = encoding.GetBytes(header)
            stream.Write(headerBytes, 0, headerBytes.Length)

            Using fileStream As FileStream = File.OpenRead(filePath)
                fileStream.CopyTo(stream)
            End Using

            Dim newlineBytes As Byte() = encoding.GetBytes(vbCrLf)
            stream.Write(newlineBytes, 0, newlineBytes.Length)
        End Sub
        Private Sub WriteFormField(stream As Stream, fieldName As String, fieldValue As String, boundary As String, encoding As Encoding)
            Dim formData As String = "--" & boundary & vbCrLf &
                                     "Content-Disposition: form-data; name=""" & fieldName & """" & vbCrLf & vbCrLf &
                                     fieldValue & vbCrLf
            Dim formDataBytes As Byte() = encoding.GetBytes(formData)
            stream.Write(formDataBytes, 0, formDataBytes.Length)
        End Sub

    End Class
End Namespace



