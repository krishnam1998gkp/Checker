'=======================================================================================================
'Created by            :    Umesh Saini
'Started On Date       :    30-March-2006
'Completion Date       :    30-March-2006
'What is the form for  :    for Preview the Report of particular Employee for Selected Period
'=======================================================================================================
'For Bug fix
'SNo. ===== Name ======== Date-Time ========= Purpose ==================================================
'12.    Jay Sharma      25 Aug 2016     Get path by objCommon.GetDirpath()
'13.    Swapnil         02 Apr 2018     Saved report configuration.
'14.    Rohtas Singh    17 May 2018     Change the Export Excel logic
'15.    Quadir Nawaj    14 July 2020    Added new code for Folder Structure related work in YTD_Salary_Slip
'16.    Ritu            5 jul 2022       regarding password apply
'17.    Huzefa          3 APR 2023       Replace "_" with "_" in File Name
'18.    Debargha        24 Jul 2024     Added 'Please Wait' clickbait on YTD Salary Register report download
'19.    Debargha        16 Sep 2024     Added API Integration for YTD Salary Register report
'20.    Vishal Chauhan  12 Sep 2024     Progress bar added on salary slip
'21.    Debargha        07 Oct 2024     Added logic to hit the download status within 4 seconds for first time and 10 seconds thereafter 
'22.    Vishal Chauhan  27 Jan 2025     Progress bar on Excel with Report API
'=======================================================================================================
Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.IO
Imports Ionic.Zip
Imports Newtonsoft.Json
Imports System.Net
Namespace Payroll
    Partial Class Rpt_IndEmpSalaryRegister
        Inherits System.Web.UI.Page
        Private _objTableManager As New clsEncryptDecrypt
        Private ExportToExcel As New WriteExcelFileByXML
        Dim filesToInclude As New System.Collections.Generic.List(Of [String])()
#Region "Developer Generated Code"
        Protected ObjCommon As New PayrollUtility.common
        Private Objdatamanager As New PayrollUtility.Utilities, ObjException As New PayrollUtility.ExceptionManager
        Public enable_report_service As Integer = 0
#End Region
#Region " Web Form Designer Generated Code "
        'This call is required by the Web Form Designer.
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        End Sub
        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: This method call is required by the Web Form Designer
            'Do not modify it using the code editor.
            InitializeComponent()
        End Sub
#End Region
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            'To check the Session
            Try
                report_service.Value = "N"
                hdnusername.Value = Session("UId").ToString
                Session("PublishedEmpCode") = Nothing
                ObjCommon.sessionCheck(Form1)
                ScriptManager.RegisterStartupScript(Me.Page, GetType(Page), "CheckChangeScript", "CheckChange()", True)
                ScriptManager.RegisterStartupScript(Me.Page, GetType(Page), String.Format("StartupScript"), "<script type=""text"">GetAttributeName();</script>", False)
                'ScriptManager.RegisterStartupScript(Page, GetType(Page), "Key", "<script>GetAttributeName();</script>", False)
                ScriptManager.RegisterStartupScript(upd1, upd1.GetType(), Guid.NewGuid().ToString(), "CloseDialog();", True)
                Dim ArrParam(0) As SqlClient.SqlParameter, Dt As New DataTable
                ArrParam(0) = New SqlParameter("@CompanyDomain", Session("CompCode").ToString) 'CompanyDomain
                Dt = Objdatamanager.GetDataTableProc("PaySp_GetGCSModuleConfigDetails", ArrParam)
                If Dt.Rows.Count > 0 And Dt.Columns.Contains("enable_config_report") AndAlso Dt.Rows(0)("enable_config_report") IsNot DBNull.Value Then
                    enable_report_service = If(Dt.Rows(0)("enable_config_report").ToString() = "Y", 1, 0)
                End If

                If Not IsPostBack Then
                    PopddlMonth(ddlAEmonth)
                    PopddlMonth(ddlASmonth)
                    ObjCommon.PopulateDDL_SalaryProcMonthYr(ddlmonth, CType(Session("Sfindate"), DateTime),
                    CType(Session("Efindate"), DateTime))
                    ObjCommon.PopulateDDL_SalaryProcMonthYr(ddlmonth1, CType(Session("Sfindate"), DateTime),
                    CType(Session("Efindate"), DateTime))
                    ddlmonth.SelectedIndex = 0
                    PnlAfromdatatodate.Style.Value = "display:none"
                    PnlSMonth.Style.Value = "display:"
                    txtpagevalue.Text = "2"
                    EmpDetailsBound()
                    PopulateReportRecord()
                    Dim IsNewUrl As String
                    IsNewUrl = IsRptapiConfigured4CSV("Paysp_YearToDate_ForExcel")
                    If IsNewUrl IsNot Nothing AndAlso IsNewUrl.ToUpper.Trim = "Y" Then
                        btnDownloadCSV.Visible = True
                    Else
                        btnDownloadCSV.Visible = False
                    End If
                    CheckYTDProcessLock()
                    CheckYTDExcelAlreadyProcessing()
                End If
            Catch ex As Exception
                ObjException.PublishError("Page_Load()", ex)
            End Try
        End Sub

        Private Sub btnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReset.Click
            Try
                Response.Redirect("Rpt_IndEmpSalaryRegister.aspx")
                btnEdit.Style.Value = "display:"
                TrEmpDetails.Style.Value = "display:none"
                TrFormat.Style.Value = "display:none"
                TrExtrapolate.Style.Value = "display:none"
                Trpublish.Style.Value = "display:none"
                TrSvConfig.Style.Value = "display:none"
                TrSave.Style.Value = "display:"
            Catch ex As Exception
                ObjException.PublishError("Error in clear the control(btnreset_Click)", ex)
            End Try
        End Sub
        Protected Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender
            Dim objSecurityManager As New PayrollUtility.SecurityManager
            objSecurityManager.UserAccessintoForm("Rpt_IndEmpSalaryRegister.aspx", , , False)
        End Sub
        'This procuedre is used for populate the month dropdownlist on the form.add by sameer on 7 Aug_09
        Private Sub PopddlMonth(Optional ByRef ddl As System.Web.UI.WebControls.DropDownList = Nothing, Optional ByVal dt As String = "")
            Try
                Dim _YEar As String
                Dim _EYEar As String
                _YEar = CDate(Session("Sfindate")).Year.ToString
                _EYEar = CDate(Session("Efindate")).Year.ToString
                ddl.Items.Add(New ListItem("Apr - " & _YEar, "4"))
                ddl.Items.Add(New ListItem("May - " & _YEar, "5"))
                ddl.Items.Add(New ListItem("Jun - " & _YEar, "6"))
                ddl.Items.Add(New ListItem("Jul - " & _YEar, "7"))
                ddl.Items.Add(New ListItem("Aug - " & _YEar, "8"))
                ddl.Items.Add(New ListItem("Sep - " & _YEar, "9"))
                ddl.Items.Add(New ListItem("Oct - " & _YEar, "10"))
                ddl.Items.Add(New ListItem("Nov - " & _YEar, "11"))
                ddl.Items.Add(New ListItem("Dec - " & _YEar, "12"))
                ddl.Items.Add(New ListItem("Jan - " & _EYEar, "1"))
                ddl.Items.Add(New ListItem("Feb - " & _EYEar, "2"))
                ddl.Items.Add(New ListItem("Mar - " & _EYEar, "3"))
                ddlASmonth.SelectedValue = "4"
                ddlAEmonth.SelectedValue = "3"
            Catch ex As Exception
                ObjException.PublishError("For populating the month drop down list(PopddlMonth())", ex)
            End Try
        End Sub
        'For check extrapolate salary register if check yes then dropdown list populate
        'from April to March financial year add by sameer on 7 Aug_09
        Protected Sub rblextrapolate_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rblextrapolate.SelectedIndexChanged
            Try
                btnEdit.Style.Value = "display:"
                btnEdit.Text = "DONE"
                TrEmpDetails.Style.Value = "display:"
                TrFormat.Style.Value = "display:"
                TrExtrapolate.Style.Value = "display:"
                Trpublish.Style.Value = "display:"
                TrSvConfig.Style.Value = "display:"
                TrSave.Style.Value = "display:none"

                tdDwn.Style.Value = "display:none"
                If rblextrapolate.SelectedValue.ToString = "Y" Then
                    PnlAfromdatatodate.Style.Value = "display:"
                    PnlSMonth.Style.Value = "display:none"
                Else
                    PnlSMonth.Style.Value = "display:"
                    PnlAfromdatatodate.Style.Value = "display:none"
                    ObjCommon.PopulateDDL_SalaryProcMonthYr(ddlmonth, CType(Session("Sfindate"), DateTime), CType(Session("Efindate"), DateTime))
                    ObjCommon.PopulateDDL_SalaryProcMonthYr(ddlmonth1, CType(Session("Sfindate"), DateTime), CType(Session("Efindate"), DateTime))
                    ddlmonth.SelectedIndex = 0
                End If

                If rbtnreportformate.SelectedValue = "V" Then
                    trRepIn.Style.Value = "display:"
                    If ddlreportIn.SelectedValue.Equals("Y") Then
                        trpswd.Style.Value = "display:"
                    Else
                        trpswd.Style.Value = "display:none"
                    End If
                Else
                    trRepIn.Style.Value = "display:none"
                    trpswd.Style.Value = "display:none"
                End If
            Catch ex As Exception
                ObjException.PublishError("rblextrapolate_SelectedIndexChanged()", ex)
            End Try
        End Sub

        Protected Sub btnProgressbarExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProgressbarExcel.Click
            Dim AppPathStr As String = HttpRuntime.AppDomainAppVirtualPath.ToString
            Dim _array() As String
            _array = Split(AppPathStr, "/")
            AppPathStr = _array(_array.Length - 1)
            lblProcessStatusExcel.Text = ""
            divSocialExcel.Visible = False
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openporess809", "OpenAttendanceProcessBar('" & AppPathStr & "','" & hdnProcessType.Value & "', '" & hdnRptName.Value.Replace("'", "") & "');", True)
        End Sub
        Private Sub CheckYTDExcelAlreadyProcessing()
            Dim arprm(2) As SqlClient.SqlParameter
            arprm(0) = New SqlClient.SqlParameter("@UserID", Session("UID").ToString)
            arprm(1) = New SqlClient.SqlParameter("@Process_Type", hdnProcessType.Value)
            arprm(2) = New SqlClient.SqlParameter("@ActionType", "isalreadyrun")
            Dim _dt As DataTable = Objdatamanager.GetDataTableProc("PaySP_ReportApi_ProcessBar", arprm)
            If (_dt.Rows.Count > 0 AndAlso _dt.Rows(0)("IsAbleToStart").ToString = "0") Then
                divSocialExcel.Visible = True
                lblProcessStatusExcel.Text = hdnRptName.Value.Replace("'", "") & " is already processing. Please wait till the completion."
                If (_dt.Rows(0)("UserId").ToString.ToUpper <> Session("UID").ToUpper.ToString) Then
                    btnProgressbarExcel.Visible = False
                Else
                    btnProgressbarExcel.Visible = True
                End If
            Else
                lblProcessStatusExcel.Text = ""
                divSocialExcel.Visible = False
            End If
        End Sub

        'This Function used to YTD SALARY REGISTER in excel sheet added by sameer on 21_Nov_2009
        Private Sub GetEmployeeYTD()
            Dim ArrParam(22) As SqlClient.SqlParameter, dt As New DataTable, _msg As New List(Of PayrollUtility.UserMessage)
            Dim APIConfigParam(2) As SqlClient.SqlParameter, IsNewUrl As String = "N"
            Dim AppPathStr As String = HttpRuntime.AppDomainAppVirtualPath.ToString, _array() As String
            _array = Split(AppPathStr, "/")
            AppPathStr = _array(_array.Length - 1)
            Dim IpAddress As String = ObjCommon.nNz(Request.UserHostAddress.ToUpper.ToString).ToString
            Dim controller As String = "YTDSalaryRegister"

            'Added by Debargha on 09-Oct-2024
            APIConfigParam(0) = New SqlClient.SqlParameter("@SP_Name", "Paysp_YearToDate_ForExcel")
            APIConfigParam(1) = New SqlClient.SqlParameter("@ReportName", "YTD Report in Excel")
            APIConfigParam(2) = New SqlClient.SqlParameter("@IsNewURL", SqlDbType.VarChar, 1)
            APIConfigParam(2).Direction = ParameterDirection.Output
            Objdatamanager.ExecuteStoredProcMsg("PaySP_ReportAPI_ConfigSel", APIConfigParam)
            IsNewUrl = APIConfigParam(2).Value.ToString
            IsNewUrl = "Y"

            'If IsDBNull(reportConfigDt.Rows(0).Item("isNewURL")) Or reportConfigDt.Rows(0).Item("isNewURL").ToString().ToUpper.Trim() = "N" Then
            If IsNewUrl = "N" Or IsNewUrl = Nothing Then
                'Vishal Chuahan multi select option added on common search paramter
                ArrParam(0) = New SqlClient.SqlParameter("@pk_emp_Code", USearchMulti.UCTextcode.ToString)
                ArrParam(1) = New SqlClient.SqlParameter("@first_name", USearchMulti.UCTextname.ToString())
                ArrParam(2) = New SqlClient.SqlParameter("@last_name", "")
                ArrParam(3) = New SqlClient.SqlParameter("@fk_costcenter_code", USearchMulti.UCddlcostcenter.ToString())
                ArrParam(4) = New SqlClient.SqlParameter("@fk_dept_code", USearchMulti.UCddldept.ToString())
                ArrParam(5) = New SqlClient.SqlParameter("@fk_desig_code", USearchMulti.UCddldesig.ToString())
                ArrParam(6) = New SqlClient.SqlParameter("@fk_grade_code", USearchMulti.UCddlgrade.ToString())
                ArrParam(7) = New SqlClient.SqlParameter("@fk_loc_code", USearchMulti.UCddllocation.ToString())
                ArrParam(8) = New SqlClient.SqlParameter("@fk_unit", USearchMulti.UCddlunit.ToString())
                ArrParam(9) = New SqlClient.SqlParameter("@salaried", USearchMulti.UCddlsalbasis.ToString())
                ArrParam(10) = New SqlClient.SqlParameter("@fk_level_Code", USearchMulti.UCddllevel.ToString())
                ArrParam(11) = New SqlParameter("@FMonth", ddlmonth.SelectedValue.ToString)
                ArrParam(12) = New SqlParameter("@FYear", Right(Trim(ddlmonth.SelectedItem.Text.ToString), 4))
                ArrParam(13) = New SqlParameter("@TMonth", ddlmonth1.SelectedValue.ToString)
                ArrParam(14) = New SqlParameter("@TYear", Right(Trim(ddlmonth1.SelectedItem.Text.ToString), 4))
                ArrParam(15) = New SqlClient.SqlParameter("@SFYear", Session("Sfindate"))
                ArrParam(16) = New SqlClient.SqlParameter("@EFYear", Session("Efindate"))
                ArrParam(17) = New SqlClient.SqlParameter("@EmpType", USearchMulti.UCddlEmp.ToString)
                ArrParam(18) = New SqlClient.SqlParameter("@UserGroup", Session("UGroup"))
                ArrParam(19) = New SqlClient.SqlParameter("@userid", Session("uid").ToString)
                ArrParam(20) = New SqlClient.SqlParameter("@SameMonthArrPay", CType(IIf(chkArrSmeMnth.Checked = True, "Y", "N"), String))
                ArrParam(21) = New SqlClient.SqlParameter("@Extrapolate", rblextrapolate.SelectedValue.ToString)
                ArrParam(22) = New SqlClient.SqlParameter("@PanApp", CType(IIf(chkEmpdet.Items(9).Selected = True, "Y", "N"), String))

                'for executing the procedure accoding to parameter
                dt = Objdatamanager.GetDataTableProc("Paysp_YearToDate_ForExcel", ArrParam)
                If dt.Rows.Count > 0 Then
                    Export_Excel(dt)
                    dt.Clear()
                    dt.Dispose()
                Else
                    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "No record(s) found according to the selection criteria !"})
                    ObjCommon.ShowMessage(_msg)
                End If
            Else
                Dim arprm(5) As SqlClient.SqlParameter
                arprm(0) = New SqlClient.SqlParameter("@UserID", Session("UID").ToString)
                arprm(1) = New SqlClient.SqlParameter("@Process_Type", hdnProcessType.Value)
                arprm(2) = New SqlClient.SqlParameter("@ActionType", "Init")
                arprm(3) = New SqlClient.SqlParameter("@Sys_IP", "::1")
                arprm(4) = New SqlClient.SqlParameter("@HostIP", ConfigurationManager.AppSettings("Hostip").ToString())
                arprm(5) = New SqlClient.SqlParameter("@ProcName", "Paysp_YearToDate_ForExcel")
                Dim _dt As DataTable = Objdatamanager.GetDataTableProc("PaySP_ReportApi_ProcessBar", arprm)
                If (_dt.Rows.Count > 0) Then
                    If (_dt.Rows(0)("IsAbleToStart").ToString = "1" AndAlso _dt.Rows(0)("BatchId").ToString <> "") Then
                        Dim scripttag As String = "StartProcessbar('" & hdnRptName.Value.Replace("'", "") & "');"
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpoocessbar231", scripttag, True)
                        hdnBatchId.Value = _dt.Rows(0)("BatchId").ToString
                        btnProgressbarExcel.Visible = False
                        lblProcessStatusExcel.Text = ""
                        divSocialExcel.Visible = False
                    Else
                        divSocialExcel.Visible = True
                        lblProcessStatusExcel.Text = hdnRptName.Value.Replace("'", "") & " is already processing. Please wait till the completion."
                        If (_dt.Rows(0)("UserId").ToString.ToUpper <> Session("UID").ToUpper.ToString) Then
                            btnProgressbarExcel.Visible = False
                        Else
                            btnProgressbarExcel.Visible = True
                        End If
                        _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = lblProcessStatusExcel.Text})
                        ObjCommon.ShowMessage(_msg)
                        Exit Sub
                    End If
                End If
                Dim keyValuePairs As New Dictionary(Of String, Object) From
                {
                 {"hostIp", ConfigurationManager.AppSettings("Hostip").ToString()},
                 {"moduleType", HttpContext.Current.Session("ModuleType").ToString()},
                 {"domainName", Session("CompCode").ToString},
                 {"pk_emp_code", USearchMulti.UCTextcode.ToString},
                 {"firstName", USearchMulti.UCTextname.ToString},
                 {"lastName", ""},
                 {"fk_costcenter_code", USearchMulti.UCddlcostcenter.ToString},
                 {"fk_dept_code", USearchMulti.UCddldept.ToString},
                 {"fk_desig_code", USearchMulti.UCddldesig.ToString},
                 {"fk_grade_code", USearchMulti.UCddlgrade.ToString},
                 {"fk_loc_code", USearchMulti.UCddllocation.ToString},
                 {"fk_unit", USearchMulti.UCddlunit.ToString},
                 {"salaried", USearchMulti.UCddlsalbasis.ToString},
                 {"fk_level_code", USearchMulti.UCddllevel.ToString},
                 {"fMonth", ddlmonth.SelectedValue.ToString},
                 {"fYear", Right(Trim(ddlmonth.SelectedItem.Text.ToString), 4)},
                 {"tMonth", ddlmonth1.SelectedValue.ToString},
                 {"tYear", Right(Trim(ddlmonth1.SelectedItem.Text.ToString), 4)},
                 {"sFYear", Session("Sfindate")},
                 {"eFYear", Session("Efindate")},
                 {"empType", USearchMulti.UCddlEmp.ToString},
                 {"userGroup", Session("UGroup")},
                 {"userId", Session("uid").ToString},
                 {"sameMonthArrPay", CType(IIf(chkArrSmeMnth.Checked = True, "Y", "N"), String)},
                 {"extrapolate", rblextrapolate.SelectedValue.ToString},
                 {"panApp", CType(IIf(chkEmpdet.Items(9).Selected = True, "Y", "N"), String)},
                 {"BatchId", hdnBatchId.Value.ToString},
                 {"FileFormat", hdnFileFormat.Value.ToString}
                }

                Dim requestBody As String = JsonConvert.SerializeObject(keyValuePairs)
                'CallAPIReport(requestBody, "YTDSalaryRegister")
                CallReportAPIOnNewThread(requestBody, controller, AppPathStr)
            End If
        End Sub
        'For create the excel file and allow the client to upload the excel file added by sameer on 21_Nov_2009
        Private Sub Export_Excel(ByVal dt As DataTable)
            Dim filepath As String = "", complexID As String = ""
            complexID = "YTD_SALARY_REGISTER" & Left(Guid.NewGuid.ToString, 5).ToString
            filepath = ExportToExcel.ExportToExcelXML_ByDataTable(dt, "", "", False, complexID.ToString)
            hdfile.Value = filepath + "~" + "YTD_SALARY_REGISTER" + "~" + "D"
            Dim popupscript As String = "<script language='javascript' type='text/javascript'>ShowDownload('" & hdfile.Value & "')</script>"
            ClientScript.RegisterStartupScript(GetType(String), "PopupScript", popupscript)
        End Sub
        'To filter the row Pay code according to employee code and month and year added by sameer on 21_Nov_2009
        Private Function GetPayCodeRow(ByVal fk_Emp_code As String, ByVal fk_Pay_code As String, ByVal Curr_MM As String, ByVal Curr_YY As String, ByVal dt As DataTable) As DataRow()
            Try
                Dim Drow() As DataRow
                Drow = dt.Select("fk_emp_code='" & fk_Emp_code & "' AND fk_Pay_code='" & fk_Pay_code & "' AND Curr_MM='" & Curr_MM & "' AND Curr_YY='" & Curr_YY & "'")
                Return Drow
            Catch ex As Exception
                ObjException.PublishError("Error in GetRecRow()", ex)
            End Try
            Return Nothing
        End Function
        'To filter the row Month , Year according to employee code added by sameer on 21_Nov_2009
        Private Function GetEmpNetRow(ByVal fk_Emp_code As String, ByVal Curr_MM As String, ByVal Curr_YY As String, ByVal dt As DataTable) As DataRow()
            Try
                Dim Drow() As DataRow
                Drow = dt.Select("fk_emp_code='" & fk_Emp_code & "' AND Curr_MM='" & Curr_MM & "' AND Curr_YY='" & Curr_YY & "'")
                Return Drow
            Catch ex As Exception
                ObjException.PublishError("Error in GetRecRow()", ex)
            End Try
            Return Nothing
        End Function
        'This procdure used to YTD SALARY REGISTER exprt in excel sheet added by sameer on 21_Nov_2009
        Protected Sub btnDownloadCSV_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDownloadCSV.Click
            hdnRptName.Value = "YTD Salary Register CSV"
            hdnFileFormat.Value = "CSV"


        End Sub
        Protected Sub btnexcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnexcel.Click
            hdnRptName.Value = "YTD Salary Register EXCEL"
            hdnFileFormat.Value = "EXCEL"
            If enable_report_service = 1 Then
                report_service.Value = "Y"
                If Not System.Configuration.ConfigurationManager.AppSettings("ReportServiceDomain") Is Nothing Then
                    report_service_url.Value = System.Configuration.ConfigurationManager.AppSettings("ReportServiceDomain").ToString
                End If
                CallReportService()
            Else
                report_service.Value = "N"
                GetEmployeeYTD()
            End If
        End Sub
        Protected Sub LnkPDF_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LnkPDF.Click
            Dim _strJava As New StringBuilder, filepath As String, FileName() As String = Nothing, FName As String = ""
            filepath = ObjCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\YTD_Salary_Slip" & "\FY_" _
                                   & Year(CDate(Session("Sfindate").ToString)) & "-" & Year(CDate(Session("Efindate").ToString))
            If Directory.Exists(filepath) Then
                FName = "_YearToDateSalarySlip.pdf"
                If (HidEmpCodes4Pdf.Value.ToString.Last() <> ",") Then
                    HidEmpCodes4Pdf.Value += ","
                End If
                HidEmpCodes4Pdf.Value = HidEmpCodes4Pdf.Value.Replace(",", FName & ",")
                If HidEmpCodes4Pdf.Value.ToString.Trim <> "" Then
                    HidEmpCodes4Pdf.Value = Left(HidEmpCodes4Pdf.Value.ToString.Trim, Len(HidEmpCodes4Pdf.Value.ToString.Trim) - 1)
                End If
                FileName = Split(HidEmpCodes4Pdf.Value.ToString.Trim, ",")

                AddZipFiles(filepath, FileName)
                Response.Clear()
                Response.BufferOutput = False
                Dim c As System.Web.HttpContext = System.Web.HttpContext.Current

                Dim archiveName = [String].Format("YearToDateSalarySlip-{0}.zip", DateTime.Now.ToString("dd-MMM-yyyy"))
                Response.ContentType = "application/zip"
                Response.AddHeader("content-disposition", "filename=" & archiveName)

                Using zip As New ZipFile()
                    zip.AddFiles(filesToInclude, "YearToDateSalarySlip")
                    zip.Save(Response.OutputStream)
                End Using
                Response.End()
            End If
        End Sub
        'Button Added by Vishal Chauhan to Donwload already publsihed pdf
        Protected Sub btnPublishedPDF_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPublishedPDF.Click
            Dim _strJava As New StringBuilder, filepath As String, FileName() As String = Nothing, FName As String = "", PublishedPdfEmpCodes As String = "" _
            , _msg As New List(Of PayrollUtility.UserMessage), arrCd(21) As String, _dt As DataTable _
            , allFileNames As String(), rowsToRemove As New List(Of DataRow)
            Session("PublishedEmpCode") = Nothing
            If (ddlreportIn.SelectedValue.ToString.ToUpper = "P" AndAlso rbtnreportformate.SelectedValue.ToString.ToUpper = "V") Then
                arrCd(0) = USearchMulti.UCddlsalbasis.ToString
                arrCd(1) = USearchMulti.UCddldesig.ToString
                arrCd(2) = USearchMulti.UCddlgrade.ToString
                arrCd(3) = USearchMulti.UCddllevel.ToString
                arrCd(4) = USearchMulti.UCddlcostcenter.ToString
                arrCd(5) = USearchMulti.UCddllocation.ToString
                arrCd(6) = USearchMulti.UCddlunit.ToString
                arrCd(7) = USearchMulti.UCddlsalbasis.ToString
                arrCd(8) = USearchMulti.UCTextcode.ToString
                arrCd(9) = ddlmonth.SelectedValue
                arrCd(10) = ddlmonth1.SelectedValue
                arrCd(11) = rblextrapolate.SelectedValue
                arrCd(12) = ""
                arrCd(13) = USearchMulti.UCTextname.ToString
                arrCd(14) = ""
                arrCd(15) = USearchMulti.UCddlEmp.ToString()
                arrCd(16) = ""
                arrCd(17) = ddlmonth.SelectedItem.Text.Split("-"c)(1)
                arrCd(18) = ddlmonth1.SelectedItem.Text.Split("-"c)(1)
                arrCd(19) = ""
                arrCd(20) = ""
                arrCd(21) = ddlEmpPass.SelectedValue
                HidEmpCodes4Pdf.Value = ""
                _dt = SetSalSlipYTDProcessInfo(arrCd, "Y")
                If (_dt.Rows.Count = 0) Then
                    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "No record(s) published till!"})
                    ObjCommon.ShowMessage(_msg)
                    Exit Sub
                End If
            End If

            filepath = ObjCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\YTD_Salary_Slip" & "\FY_" _
                                   & Year(CDate(Session("Sfindate").ToString)) & "-" & Year(CDate(Session("Efindate").ToString))

            If Directory.Exists(filepath) Then
                allFileNames = Directory.GetFiles(filepath, "*YearToDateSalarySlip.PDF", SearchOption.AllDirectories)
                FName = "_YearToDateSalarySlip.pdf"
                For Each row As DataRow In _dt.Rows
                    Dim empCode As String = row("EmpCode").ToString()
                    Dim fileExists As Boolean = False
                    For Each fPath As String In allFileNames
                        Dim flName As String = Path.GetFileName(fPath)
                        If flName.StartsWith(empCode & "_") AndAlso flName.EndsWith("YearToDateSalarySlip.pdf") Then
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
                    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "No record published till!"})
                    ObjCommon.ShowMessage(_msg)
                    Exit Sub
                End If
                HidEmpCodes4Pdf.Value = String.Join(",",
                                      _dt.AsEnumerable().
                                      Where(Function(row) Not IsDBNull(row("EmpCode"))).
                                      Select(Function(row) row("EmpCode").ToString()))
                If (HidEmpCodes4Pdf.Value.ToString.Last() <> ",") Then
                    HidEmpCodes4Pdf.Value += ","
                End If
                Session("PublishedEmpCode") = HidEmpCodes4Pdf.Value
                hdfile.Value = "NA~YTD~YTD"
                Dim popupScript As String = "<script language='javascript' type='text/javascript'>ShowDownload('" & hdfile.Value & "')</script>"
                ClientScript.RegisterStartupScript(GetType(String), "PopupScriptPayslip", popupScript)
            Else
                _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "No record(s) published till!"})
                ObjCommon.ShowMessage(_msg)
                Exit Sub
            End If
        End Sub

        'Added AddZipFiles Function for Folder Structure related work in YTD_Salary_Slip by Quadir on 14 July 2020
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

        Protected Sub btnPriview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPriview.Click
            Dim _strVal As String = "", _EmpDet As String = "", cnt As Integer = 0, _AppPath As String = HttpRuntime.AppDomainAppVirtualPath.ToString, _Path As String, _array() As String _
                , _msg As New List(Of PayrollUtility.UserMessage)
            Dim SalBase As String = USearchMulti.UCddlsalbasis.ToString _
                , EmpCode = USearchMulti.UCTextcode.ToString _
                , Ename = USearchMulti.UCTextname.ToString _
                , es As String = "F" _
                , Month0 As String = ddlmonth.SelectedValue _
                , Month1 As String = ddlmonth1.SelectedValue _
                , Reptype As String = rblextrapolate.SelectedValue _
                , EmpType = USearchMulti.UCddlEmp.ToString() _
                , PageNo = txtpagevalue.Text _
                , EmpSts = USearchMulti.UCddlEmp.ToString _
                , password As String = ddlEmpPass.SelectedValue _
                , year0 As String = ddlmonth.SelectedItem.Text.Split("-"c)(1) _
                , year1 As String = ddlmonth1.SelectedItem.Text.Split("-"c)(1) _
                , ArrSmeMnth As String = "N"
            hdf_USearchMulti_DdlCostCenter.Value = USearchMulti.UCddlcostcenter.ToString
            hdf_USearchMulti_DdlDept.Value = USearchMulti.UCddldept.ToString
            hdf_USearchMulti_ddllocation.Value = USearchMulti.UCddllocation.ToString
            hdf_USearchMulti_ddldesignation.Value = USearchMulti.UCddldesig.ToString
            hdf_USearchMulti_ddlunit.Value = USearchMulti.UCddlunit.ToString
            hdf_USearchMulti_ddlGrade.Value = USearchMulti.UCddlgrade.ToString
            hdf_USearchMulti_ddllevel.Value = USearchMulti.UCddllevel.ToString
            HidEmpCodes4Pdf.Value = ""
            spnmsgtolink.InnerHtml = "Click to PDF icon to download the .zip file."
            'divSocial.Style.Value = "display:none"
            tdDwn.Style.Value = "display:none"

            For i As Integer = 0 To chkEmpdet.Items.Count - 1
                If chkEmpdet.Items(i).Selected Then
                    _EmpDet = _EmpDet & chkEmpdet.Items(i).Value & ","
                    cnt = cnt + 1
                Else
                    _EmpDet = _EmpDet & ","
                End If
            Next
            If USearchMulti.UCTextcode <> "" Then
                Dim oWrite As System.IO.StreamWriter = Nothing, _fs As FileStream = Nothing, _fileAdd As String = ""

                Try
                    _fileAdd = ObjCommon.GetDirpath(Session("COMPCODE").ToString) & "\TempExcelFiles"

                    If Not Directory.Exists(_fileAdd) Then
                        Directory.CreateDirectory(_fileAdd)
                    End If

                    _fileAdd = _fileAdd & "\" & Guid.NewGuid().ToString & ".txt"
                    hidEmpCode.Value = Path.GetFileNameWithoutExtension(_fileAdd)
                    If File.Exists(_fileAdd) Then
                        File.Delete(_fileAdd)
                    End If
                    _fs = New FileStream(_fileAdd, FileMode.Create, FileAccess.ReadWrite)

                    oWrite = New StreamWriter(_fs)
                    oWrite.Write(USearchMulti.UCTextcode.ToString)
                Catch ex As Exception
                    ScriptManager.RegisterStartupScript(Me.Page, GetType(Page), "previewscrt", "CloseProcessbarDialog()", True)
                    ObjException.PublishError("Generate txt file", ex)
                Finally
                    oWrite.Close()
                    oWrite.Dispose()
                    _fs.Close()
                    _fs.Dispose()
                    'Added by Rohtas Singh on 06 Apr 2015 (Resolve when click on Preview button "trRepIn" hide automatic
                    If rbtnreportformate.SelectedValue = "V" Then
                        trRepIn.Style.Value = "display:"
                        If ddlreportIn.SelectedValue.Equals("Y") Then
                            trpswd.Style.Value = "display:"
                        Else
                            trpswd.Style.Value = "display:none"
                        End If
                    Else
                        trRepIn.Style.Value = "display:none"
                        trpswd.Style.Value = "display:none"
                    End If
                End Try
            Else
                hidEmpCode.Value = ""
            End If
            'added by Vishal Chauhan for progress bar
            If (ddlreportIn.SelectedValue.ToString.ToUpper = "P" AndAlso rbtnreportformate.SelectedValue.ToString.ToUpper = "V") Then

                If (chkArrSmeMnth.Checked) Then
                    ArrSmeMnth = "Y"
                End If
                Dim var As String = hdf_USearchMulti_DdlDept.Value + "~" + hdf_USearchMulti_ddldesignation.Value + "~" + hdf_USearchMulti_ddlGrade.Value + "~" _
                + hdf_USearchMulti_ddllevel.Value + "~" + hdf_USearchMulti_DdlCostCenter.Value + "~" + hdf_USearchMulti_ddllocation.Value + "~" + hdf_USearchMulti_ddlunit.Value _
                + "~" + SalBase + "~" + EmpCode + "~" + Month0 + "~" + Month1 + "~" + Reptype + "~" + PageNo + "~" + es + "~" _
                + Ename + "~" + EmpSts + "~" + _EmpDet + "~" + year0 + "~" + year1 + "~P~" + ArrSmeMnth + "~" + password + "~" + cnt.ToString
                _strVal = Guid.NewGuid.ToString
                If Not Session(_strVal) Is Nothing Then
                    Session.Remove(_strVal)
                End If
                Session(_strVal) = var.ToString
                var = _strVal.ToString
                HidPreVal.Value = var.ToString
                _array = Split(_AppPath, "/")
                _AppPath = _array(_array.Length - 1)
                HidAppPath.Value = _AppPath
                _Path = ObjCommon.GetDirpath(Session("COMPCODE").ToString) & "\PDFFiles\YTD_Salary_Slip" & "\FY_" _
                    & Year(CDate(Session("Sfindate").ToString)) & "-" & Year(CDate(Session("Efindate").ToString))

                If Not Directory.Exists(_Path) Then
                    Directory.CreateDirectory(_Path)
                End If
                Dim arrCd() As String = Split(Session(_strVal).ToString, "~")
                Dim _dt As DataTable = SetSalSlipYTDProcessInfo(arrCd)
                If (_dt.Rows.Count = 0) Then
                    LnkPDF.Style.Value = "display:none"
                    ScriptManager.RegisterStartupScript(Me.Page, GetType(Page), "previewscrt", "CloseProcessbarDialog()", True)
                    _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = "No record(s) found according to the selection criteria !"})
                    ObjCommon.ShowMessage(_msg)
                    Exit Sub
                End If
                LnkPDF.Style.Value = ""
                HidEmpCodes4Pdf.Value = String.Join(",",
                                      _dt.AsEnumerable().
                                      Where(Function(row) Not IsDBNull(row("EmpCode"))).
                                      Select(Function(row) row("EmpCode").ToString()))
                HidPath.Value = Replace(Replace(_Path, "\", "~").ToString, "/", "~").ToString
            End If
            ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "openpopup", "datecheck('" & _EmpDet & "','" & cnt & "');", True)
        End Sub
        'New function added by Vishal Chauhan to get total records
        Private Function SetSalSlipYTDProcessInfo(ByVal arrCode() As String, Optional ByVal IsPublished As String = "N") As DataTable
            Dim _ArrParam(23) As SqlParameter, dst As New DataTable, _empCodes As String = ""
            Try
                _ArrParam(0) = New SqlParameter("@fk_costcenter_code", arrCode(4).Trim.ToString)
                _ArrParam(1) = New SqlParameter("@fk_loc_code", arrCode(5).Trim.ToString)
                _ArrParam(2) = New SqlParameter("@Fk_unit", arrCode(6).Trim.ToString)
                _ArrParam(3) = New SqlParameter("@salaried", arrCode(7).Trim.ToString)
                _ArrParam(4) = New SqlParameter("@pk_emp_code", arrCode(8).Trim.ToString)
                _ArrParam(5) = New SqlParameter("@fk_dept_code", arrCode(0).Trim.ToString)
                _ArrParam(6) = New SqlParameter("@fk_desig_code", arrCode(1).Trim.ToString)
                _ArrParam(7) = New SqlParameter("@fk_grade_code", arrCode(2).Trim.ToString)
                _ArrParam(8) = New SqlParameter("@fk_level_Code", arrCode(3).Trim.ToString)
                _ArrParam(9) = New SqlParameter("@FMonth", arrCode(9).ToString)
                _ArrParam(10) = New SqlParameter("@FYear", Right(Trim(arrCode(17).ToString), 4))
                _ArrParam(11) = New SqlParameter("@TMonth", arrCode(10).ToString)
                _ArrParam(12) = New SqlParameter("@TYear", Right(Trim(arrCode(18).ToString), 4))
                _ArrParam(13) = New SqlParameter("@Extrapolate", arrCode(11).ToString)
                _ArrParam(14) = New SqlParameter("@SFYear", Session("Sfindate").ToString)
                _ArrParam(15) = New SqlParameter("@EFYear", Session("Efindate").ToString)
                _ArrParam(16) = New SqlParameter("@EmpType", arrCode(15).ToString)
                _ArrParam(17) = New SqlParameter("@UserGroup", Session("UGroup").ToString)
                _ArrParam(18) = New SqlParameter("@first_name", IIf(arrCode(13).ToString.ToUpper = "F", arrCode(14).ToString, ""))
                _ArrParam(19) = New SqlParameter("@last_name", IIf(arrCode(13).ToString.ToUpper = "L", arrCode(14).ToString, ""))
                _ArrParam(20) = New SqlClient.SqlParameter("@SameMonthArrPay", "N")  'arrCode(20).ToString.ToUpper /* after discussing Pankaj Sachan,its commented (Ritu MAlik)
                _ArrParam(21) = New SqlParameter("@PassType", arrCode(21).Trim.ToString)
                _ArrParam(22) = New SqlParameter("@userid", Session("uid").ToString)
                _ArrParam(23) = New SqlParameter("@IsPublished", IsPublished)
                dst = Objdatamanager.GetDataTableProc("PaySP_YTDSalarySlip_ProcessInfo", _ArrParam)
            Catch ex As Exception
                ObjException.PublishError("Error in SetSalSlipYTDProcessInfo()", ex)
            End Try
            Return dst
        End Function

        Private Sub imgpdfprocess_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgpdfprocess.Click
            IsShowlnkIcon.Value = "Y"
            CheckYTDProcessLock()
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
        Private Sub CheckYTDProcessLock()
            Dim _ArrParam(2) As SqlParameter, dst As New DataSet, _empCodes As String = ""
            imgpdfprocess.Visible = False
            LnkPDF.Style.Value = "display:none"
            tdDwn.Style.Value = "display:none"
            divSocial.Style.Value = "display:none"
            spnmsgtolink.InnerHtml = "Click to PDF icon to download the .zip file."
            lblProcessStatus.Text = ""
            Try
                _ArrParam(0) = New SqlClient.SqlParameter("@UserID", Session("uid").ToString)
                _ArrParam(1) = New SqlClient.SqlParameter("@flag", "V")
                _ArrParam(2) = New SqlClient.SqlParameter("@Process_Type", "YTDSLIPV")
                dst = Objdatamanager.GetDataSetProc("PaySP_ReportProcess_ProcessBar", _ArrParam)
                If (dst.Tables(0).Rows.Count > 0 AndAlso dst.Tables(0).Rows(0)("process_status").ToString.ToUpper.Trim = "START" AndAlso Convert.ToInt32(dst.Tables(0).Rows(0)("total_to_process").ToString) > 0) Then
                    imgpdfprocess.Visible = True
                    divSocial.Style.Value = ""
                    lblProcessStatus.Text = "Please wait while YTD Payslip is being published... (" & dst.Tables(0).Rows(0)("total_processed").ToString & "/" & dst.Tables(0).Rows(0)("total_to_process").ToString & "). Estimated time left: " + SplitTimes(Convert.ToDouble(dst.Tables(0).Rows(0)("estimated_time_left").ToString))
                ElseIf dst.Tables(0).Rows.Count > 0 AndAlso dst.Tables(1).Rows.Count > 0 AndAlso dst.Tables(0).Rows(0)("process_status").ToString.ToUpper.Trim = "DONE" Then
                    If (IsShowlnkIcon.Value.ToUpper = "Y" AndAlso ddlreportIn.SelectedValue.ToString.ToUpper = "P" AndAlso rbtnreportformate.SelectedValue.ToString.ToUpper = "V") Then
                        HidEmpCodes4Pdf.Value = String.Join(",",
                                          dst.Tables(1).AsEnumerable().
                                          Where(Function(row) Not IsDBNull(row("fk_Emp_Code"))).
                                          Select(Function(row) row("fk_Emp_Code").ToString()))
                        LnkPDF.Style.Value = ""
                        tdDwn.Style.Value = ""
                        'spnmsgtolink.InnerHtml = "Download Already published .zip."
                    End If
                End If
            Catch ex As Exception
                ObjException.PublishError("Error in CheckYTDProcessLock()", ex)
            End Try
        End Sub

        Private Sub EmpDetailsBound()
            Dim _arr As String(,) = {{ObjCommon.DisplayCaption("COC"), "C"}, {ObjCommon.DisplayCaption("LOC"), "L"}, {ObjCommon.DisplayCaption("UNT"), "U"},
                                     {ObjCommon.DisplayCaption("DEP"), "D"}, {ObjCommon.DisplayCaption("DES"), "S"}, {ObjCommon.DisplayCaption("LVL"), "V"} _
                                , {ObjCommon.DisplayCaption("GRD"), "G"}, {"DOJ", "J"}, {"DOL", "O"}, {"PAN", "T"}}

            For i As Integer = 0 To _arr.GetUpperBound(0)
                chkEmpdet.Items.Add(New ListItem With {.Text = _arr(i, 0), .Value = _arr(i, 1)})
            Next
        End Sub
        Protected Sub btnSaveConfig_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveConfig.Click
            Dim ArrParam(8) As SqlClient.SqlParameter, Dt As New DataTable, EmpDetail As String = "", _msg As New List(Of PayrollUtility.UserMessage)
            For _P As Integer = 0 To chkEmpdet.Items.Count - 1
                If chkEmpdet.Items(_P).Selected = True Then
                    EmpDetail = EmpDetail & chkEmpdet.Items(_P).Value.ToString & ","
                End If
            Next
            If EmpDetail.ToString.Trim <> "" Then
                EmpDetail = Left(EmpDetail, Len(EmpDetail) - 1)
            End If
            ArrParam(0) = New SqlParameter("@Flag", "Y")
            ArrParam(1) = New SqlParameter("@RptType", "Y")
            ArrParam(2) = New SqlParameter("@RptFormat", ObjCommon.nNz(rbtnreportformate.SelectedValue.ToString))
            ArrParam(3) = New SqlParameter("@ArrSmeMnth", ObjCommon.nNz(IIf(chkArrSmeMnth.Checked, "Y", "N")).ToString)
            ArrParam(4) = New SqlParameter("@Extrapolate", ObjCommon.nNz(rblextrapolate.SelectedValue.ToString))
            ArrParam(5) = New SqlParameter("@EmpDetail", ObjCommon.nNz(EmpDetail.ToString))
            ArrParam(6) = New SqlParameter("@fk_userid", Session("UID").ToString)
            ArrParam(7) = New SqlParameter("@IPAddress", ObjCommon.GetIPAddress())
            ArrParam(8) = New SqlParameter("@fk_userlog", Session("UserLogKey").ToString)
            Dt = Objdatamanager.GetDataTableProc("PaySP_Trn_PFESI_ReportDetail_Sel_Ins", ArrParam)

            _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "S"})
            ObjCommon.ShowMessage(_msg)
            PopulateReportRecord()
            btnEdit.Style.Value = "display:"
            btnEdit.Text = "EDIT"
            TrEmpDetails.Style.Value = "display:none"
            TrFormat.Style.Value = "display:none"
            TrExtrapolate.Style.Value = "display:none"
            Trpublish.Style.Value = "display:none"
            TrSvConfig.Style.Value = "display:none"
            TrSave.Style.Value = "display:"
            trpswd.Style.Value = "display:none"
        End Sub
        Private Sub PopulateReportRecord()
            Dim dt As New DataTable, arrparam(2) As SqlClient.SqlParameter, EmpDetail As String()
            arrparam(0) = New SqlClient.SqlParameter("@RptType", "Y")
            arrparam(1) = New SqlParameter("@fk_userid", Session("UID").ToString)
            arrparam(2) = New SqlClient.SqlParameter("@Flag", "D")
            dt = Objdatamanager.GetDataTableProc("PaySP_Trn_PFESI_ReportDetail_Sel_Ins", arrparam)
            If dt.Rows.Count > 0 Then
                chkArrSmeMnth.Checked = CBool(IIf(dt.Rows(0)("ArrSmeMnth").ToString.ToUpper = "Y", True, False))
                rblextrapolate.SelectedValue = dt.Rows(0)("Extrapolate").ToString
                rbtnreportformate.SelectedValue = dt.Rows(0)("RptFormat").ToString
                If rblextrapolate.SelectedValue.ToString = "Y" Then
                    PnlAfromdatatodate.Style.Value = "display:"
                    PnlSMonth.Style.Value = "display:none"
                Else
                    PnlSMonth.Style.Value = "display:"
                    PnlAfromdatatodate.Style.Value = "display:none"
                End If
                If dt.Rows(0)("EmpDetail").ToString <> "" Then
                    EmpDetail = Split(dt.Rows(0)("EmpDetail").ToString, ",")
                    For LoopVal = 0 To EmpDetail.Length - 1
                        chkEmpdet.Items.FindByValue(EmpDetail(LoopVal)).Selected = True
                    Next
                End If

            End If
        End Sub
        'Created by Vishal Chauhan for progress bar'
        Private Sub CallReportAPIOnNewThread(ByVal requestBody As String, ByVal controller As String, ByVal AppPathStr As String)
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Try
                Dim apiUrl As String
                apiUrl = System.Configuration.ConfigurationManager.AppSettings("ApiRptExcel").ToString & controller
                Dim jobResponse = PostAPICall(apiUrl & "/generate-file", requestBody)
                Dim jobResponseData = Newtonsoft.Json.JsonConvert.DeserializeObject(Of JobResponse)(jobResponse)
                JobUniqueId.Value = jobResponseData.JobId.ToString
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup918", "OpenAttendanceProcessBar('" & AppPathStr & "','" & hdnProcessType.Value & "', '" & hdnRptName.Value.Replace("'", "") & "');", True)
            Catch ex As Exception
                ObjException.PublishError("Error in CallAPIReport()", ex)
                Dim ErrorCaseparams(4) As SqlClient.SqlParameter
                ErrorCaseparams(0) = New SqlParameter("@userId", Session("UID").ToString)
                ErrorCaseparams(1) = New SqlParameter("@Process_Type", hdnProcessType.Value)
                ErrorCaseparams(2) = New SqlParameter("@ErrorMsg", ex.Message)
                ErrorCaseparams(3) = New SqlParameter("@ActionType", "ErrorInExcel")
                ErrorCaseparams(4) = New SqlParameter("@BatchId", hdnBatchId.Value)
                Objdatamanager.GetDataTableProc("PaySP_ReportApi_ProcessBar", ErrorCaseparams)
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
                ObjException.PublishError("Error in CallAPIReport()", ex)
                Console.WriteLine(ex)
            End Try
        End Sub
        Private Function PostAPICall(url As String, postData As String) As String
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Try
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
                Using response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                    Using responseStream = response.GetResponseStream()
                        Using reader = New StreamReader(responseStream)
                            Return reader.ReadToEnd()
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                ObjException.PublishError("Error in PostAPICall()", ex)
                Console.WriteLine(ex)
                Return Nothing
            End Try
        End Function

        Private Function IsRptapiConfigured4CSV(ByVal ProcName As String) As String
            Try
                Dim sParam(1) As SqlClient.SqlParameter
                sParam(0) = New SqlClient.SqlParameter("@SP_Name", ProcName)
                sParam(1) = New SqlClient.SqlParameter("@IsNewURL", SqlDbType.VarChar, 1)
                sParam(1).Direction = ParameterDirection.Output
                '_ObjData.ExecuteStoredProcMsg("PaySP_ReportAPI_ConfigSel", sParam)
                'Return sParam(1).Value.ToString
                Dim dt As DataTable = Objdatamanager.GetDataTableProc("PaySP_ReportAPI_ConfigSel", sParam)
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
        Protected Function GetAppPath() As String
            Dim _array() As String, _AppPath As String = HttpRuntime.AppDomainAppVirtualPath.ToString
            _array = Split(_AppPath, "/")
            _AppPath = _array(_array.Length - 1)
            Return _AppPath
        End Function
        Private Class JobResponse
            Public Property JobId As String
            Public Property StatusUrl As String
        End Class

        Private Sub CallReportService()
            Try
                'objCommon.GetDirpath(Session("COMPCODE").ToString)
                Dim filePath As String = ObjCommon.GetDirpath(Session("COMPCODE").ToString) & "\" & Session("COMPCODE").ToString & "\TempExcelFiles\"
                Dim apiUrl As String = "", _msg As New List(Of PayrollUtility.UserMessage)
                If Not System.Configuration.ConfigurationManager.AppSettings("ReportServiceDomain") Is Nothing Then
                    If System.Configuration.ConfigurationManager.AppSettings("ReportServiceDomain").ToString <> "" Then
                        apiUrl = System.Configuration.ConfigurationManager.AppSettings("ReportServiceDomain").ToString
                    End If
                End If
                Dim g As String = Guid.NewGuid().ToString()
                Dim fileName As String = "YTD_SALARY_REGISTER_" & g.Substring(g.Length - 5)
                report_service_file_name.Value = fileName
                Dim arprm(5) As SqlClient.SqlParameter
                arprm(0) = New SqlClient.SqlParameter("@UserID", Session("UID").ToString)
                arprm(1) = New SqlClient.SqlParameter("@Process_Type", hdnProcessType.Value)
                arprm(2) = New SqlClient.SqlParameter("@ActionType", "Init")
                arprm(3) = New SqlClient.SqlParameter("@Sys_IP", "::1")
                arprm(4) = New SqlClient.SqlParameter("@HostIP", ConfigurationManager.AppSettings("Hostip").ToString())
                arprm(5) = New SqlClient.SqlParameter("@ProcName", "Paysp_YearToDate_ForExcel")
                Dim _dt As DataTable = Objdatamanager.GetDataTableProc("PaySP_ReportApi_ProcessBar", arprm)
                If (_dt.Rows.Count > 0) Then
                    If (_dt.Rows(0)("IsAbleToStart").ToString = "1" AndAlso _dt.Rows(0)("BatchId").ToString <> "") Then
                        Dim scripttag As String = "StartProcessbar('" & hdnRptName.Value.Replace("'", "") & "');"
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpoocessbar231", scripttag, True)
                        hdnBatchId.Value = _dt.Rows(0)("BatchId").ToString
                        btnProgressbarExcel.Visible = False
                        lblProcessStatusExcel.Text = ""
                        divSocialExcel.Visible = False
                    Else
                        divSocialExcel.Visible = True
                        lblProcessStatusExcel.Text = hdnRptName.Value.Replace("'", "") & " is already processing. Please wait till the completion."
                        If (_dt.Rows(0)("UserId").ToString.ToUpper <> Session("UID").ToUpper.ToString) Then
                            btnProgressbarExcel.Visible = False
                        Else
                            btnProgressbarExcel.Visible = True
                        End If
                        _msg.Add(New PayrollUtility.UserMessage With {.MessageType = "E", .MessageString = lblProcessStatusExcel.Text})
                        ObjCommon.ShowMessage(_msg)
                        Exit Sub
                    End If
                End If

                Dim execPlanParams As New Dictionary(Of String, Object) From
                {
                 {"HostIP", ConfigurationManager.AppSettings("Hostip").ToString()},
                 {"pk_emp_code", USearchMulti.UCTextcode.ToString},
                 {"first_name", USearchMulti.UCTextname.ToString},
                 {"last_name", ""},
                 {"fk_costcenter_code", USearchMulti.UCddlcostcenter.ToString},
                 {"fk_dept_code", USearchMulti.UCddldept.ToString},
                 {"fk_desig_code", USearchMulti.UCddldesig.ToString},
                 {"fk_grade_code", USearchMulti.UCddlgrade.ToString},
                 {"fk_loc_code", USearchMulti.UCddllocation.ToString},
                 {"Fk_unit", USearchMulti.UCddlunit.ToString},
                 {"salaried", USearchMulti.UCddlsalbasis.ToString},
                 {"fk_level_Code", USearchMulti.UCddllevel.ToString},
                 {"FMonth", ddlmonth.SelectedValue.ToString},
                 {"FYear", Right(Trim(ddlmonth.SelectedItem.Text.ToString), 4)},
                 {"TMonth", ddlmonth1.SelectedValue.ToString},
                 {"TYear", Right(Trim(ddlmonth1.SelectedItem.Text.ToString), 4)},
                 {"SFYear", Session("Sfindate")},
                 {"EFYear", Session("Efindate")},
                 {"EmpType", USearchMulti.UCddlEmp.ToString},
                 {"UserGroup", Session("UGroup")},
                 {"userid", Session("uid").ToString},
                 {"SameMonthArrPay", CType(IIf(chkArrSmeMnth.Checked = True, "Y", "N"), String)},
                 {"Extrapolate", rblextrapolate.SelectedValue.ToString},
                 {"PanApp", CType(IIf(chkEmpdet.Items(9).Selected = True, "Y", "N"), String)},
                 {"Sys_IP", ""},
                 {"IsAPI", "Y"},
                 {"BatchId", hdnBatchId.Value.ToString}
                }

                Dim executionPlan As New Dictionary(Of String, Object) From {
                    {"reportName", "FlatReport"},
                    {"procedureName", "Paysp_YearToDate_ForExcel_backup"},
                    {"executionPlanParameters", execPlanParams},
                    {"writerLibrary", "FlatAsposeCellsExcelWriter"}}

                Dim payload As New Dictionary(Of String, Object) From {
                {"fileName", fileName},
                {"filePath", ""},
                {"fileType", "xlsx"},
                {"filePrefix", "YTD_SALARY_REGISTER_"},
                {"progressType", "YTDSALREG"},
                {"statusUrl", "/api/YTDSalaryRegister/status/"},
                {"executionPlan", executionPlan},
                {"compCode", Session("CompCode").ToString()}
                }
                Dim requestBody As String = JsonConvert.SerializeObject(payload)
                apiUrl = apiUrl & "/api/service/payrollReportFramework/generateReportAsync"
                Try

                    Dim controller As String = "YTDSalaryRegister"


                    Dim AppPathStr As String = HttpRuntime.AppDomainAppVirtualPath.ToString, _array() As String
                    _array = Split(AppPathStr, "/")
                    AppPathStr = _array(_array.Length - 1)
                    Dim jobResponse = PostAPICall(apiUrl, requestBody)
                    Dim jobResponseData = Newtonsoft.Json.JsonConvert.DeserializeObject(Of JobResponse)(jobResponse)
                    JobUniqueId.Value = jobResponseData.JobId.ToString
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "openpopup918", "OpenAttendanceProcessBar('" & AppPathStr & "','" & hdnProcessType.Value & "', '" & hdnRptName.Value.Replace("'", "") & "');", True)
                Catch ex As Exception
                    ObjException.PublishError("Error in CallReportService()", ex)
                End Try
            Catch ex As Exception
                ObjException.PublishError("Error in CallReportService()", ex)
            End Try

        End Sub
    End Class
End Namespace
