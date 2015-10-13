<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_departmentinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim department_edit
Set department_edit = New cdepartment_edit
Set Page = department_edit

' Page init processing
department_edit.Page_Init()

' Page main processing
department_edit.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
department_edit.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var department_edit = new ew_Page("department_edit");
department_edit.PageID = "edit"; // Page ID
var EW_PAGE_ID = department_edit.PageID; // For backward compatibility
// Form object
var fdepartmentedit = new ew_Form("fdepartmentedit");
// Validate form
fdepartmentedit.Validate = function() {
	if (!this.ValidateRequired)
		return true; // Ignore validation
	var $ = jQuery, fobj = this.GetForm(), $fobj = $(fobj);
	this.PostAutoSuggest();
	if ($fobj.find("#a_confirm").val() == "F")
		return true;
	var elm, felm, uelm, addcnt = 0;
	var $k = $fobj.find("#" + this.FormKeyCountName); // Get key_count
	var rowcnt = ($k[0]) ? parseInt($k.val(), 10) : 1;
	var startcnt = (rowcnt == 0) ? 0 : 1; // Check rowcnt == 0 => Inline-Add
	var gridinsert = $fobj.find("#a_list").val() == "gridinsert";
	for (var i = startcnt; i <= rowcnt; i++) {
		var infix = ($k[0]) ? String(i) : "";
		$fobj.data("rowindex", infix);
			elm = this.GetElements("x" + infix + "_dept_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(department.dept_id.FldErrMsg) %>");
			elm = this.GetElements("x" + infix + "_office_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(department.office_id.FldErrMsg) %>");
			elm = this.GetElements("x" + infix + "_dept_sort");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(department.dept_sort.FldErrMsg) %>");
			// Set up row object
			ew_ElementsToRow(fobj);
			// Fire Form_CustomValidate event
			if (!this.Form_CustomValidate(fobj))
				return false;
	}
	// Process detail forms
	var dfs = $fobj.find("input[name='detailpage']").get();
	for (var i = 0; i < dfs.length; i++) {
		var df = dfs[i], val = df.value;
		if (val && ewForms[val])
			if (!ewForms[val].Validate())
				return false;
	}
	return true;
}
// Form_CustomValidate event
fdepartmentedit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fdepartmentedit.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fdepartmentedit.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If department.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% department_edit.ShowPageHeader() %>
<% department_edit.ShowMessage %>
<form name="fdepartmentedit" id="fdepartmentedit" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="a_table" id="a_table" value="department">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table class="ewGrid"><tr><td>
<table id="tbl_departmentedit" class="table table-bordered table-striped">
<% If department.dept_id.Visible Then ' dept_id %>
	<tr id="r_dept_id">
		<td><span id="elh_department_dept_id"><%= department.dept_id.FldCaption %></span></td>
		<td<%= department.dept_id.CellAttributes %>>
<span id="el_department_dept_id" class="control-group">
<span<%= department.dept_id.ViewAttributes %>>
<%= department.dept_id.EditValue %>
</span>
</span>
<input type="hidden" data-field="x_dept_id" name="x_dept_id" id="x_dept_id" value="<%= Server.HTMLEncode(department.dept_id.CurrentValue&"") %>">
<%= department.dept_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If department.office_id.Visible Then ' office_id %>
	<tr id="r_office_id">
		<td><span id="elh_department_office_id"><%= department.office_id.FldCaption %></span></td>
		<td<%= department.office_id.CellAttributes %>>
<span id="el_department_office_id" class="control-group">
<input type="text" data-field="x_office_id" name="x_office_id" id="x_office_id" size="30" placeholder="<%= department.office_id.PlaceHolder %>" value="<%= department.office_id.EditValue %>"<%= department.office_id.EditAttributes %>>
</span>
<%= department.office_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If department.dept_name.Visible Then ' dept_name %>
	<tr id="r_dept_name">
		<td><span id="elh_department_dept_name"><%= department.dept_name.FldCaption %></span></td>
		<td<%= department.dept_name.CellAttributes %>>
<span id="el_department_dept_name" class="control-group">
<input type="text" data-field="x_dept_name" name="x_dept_name" id="x_dept_name" size="30" maxlength="255" placeholder="<%= department.dept_name.PlaceHolder %>" value="<%= department.dept_name.EditValue %>"<%= department.dept_name.EditAttributes %>>
</span>
<%= department.dept_name.CustomMsg %></td>
	</tr>
<% End If %>
<% If department.dept_sort.Visible Then ' dept_sort %>
	<tr id="r_dept_sort">
		<td><span id="elh_department_dept_sort"><%= department.dept_sort.FldCaption %></span></td>
		<td<%= department.dept_sort.CellAttributes %>>
<span id="el_department_dept_sort" class="control-group">
<input type="text" data-field="x_dept_sort" name="x_dept_sort" id="x_dept_sort" size="30" placeholder="<%= department.dept_sort.PlaceHolder %>" value="<%= department.dept_sort.EditValue %>"<%= department.dept_sort.EditAttributes %>>
</span>
<%= department.dept_sort.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("EditBtn") %></button>
</form>
<script type="text/javascript">
fdepartmentedit.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
department_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set department_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cdepartment_edit

	' Page ID
	Public Property Get PageID()
		PageID = "edit"
	End Property

	' Project ID
	Public Property Get ProjectID()
		ProjectID = "{324ED72D-DE20-46F7-B12E-7AF8CE8711A6}"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "department"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "department_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If department.UseTokenInUrl Then PageUrl = PageUrl & "t=" & department.TableVar & "&" ' add page token
	End Property

	' Message
	Public Property Get Message()
		Message = Session(EW_SESSION_MESSAGE)
	End Property

	Public Property Let Message(v)
		Dim msg
		msg = Session(EW_SESSION_MESSAGE)
		Call ew_AddMessage(msg, v)
		Session(EW_SESSION_MESSAGE) = msg
	End Property

	Public Property Get FailureMessage()
		FailureMessage = Session(EW_SESSION_FAILURE_MESSAGE)
	End Property

	Public Property Let FailureMessage(v)
		Dim msg
		msg = Session(EW_SESSION_FAILURE_MESSAGE)
		Call ew_AddMessage(msg, v)
		Session(EW_SESSION_FAILURE_MESSAGE) = msg
	End Property

	Public Property Get SuccessMessage()
		SuccessMessage = Session(EW_SESSION_SUCCESS_MESSAGE)
	End Property

	Public Property Let SuccessMessage(v)
		Dim msg
		msg = Session(EW_SESSION_SUCCESS_MESSAGE)
		Call ew_AddMessage(msg, v)
		Session(EW_SESSION_SUCCESS_MESSAGE) = msg
	End Property

	Public Property Get WarningMessage()
		WarningMessage = Session(EW_SESSION_WARNING_MESSAGE)
	End Property

	Public Property Let WarningMessage(v)
		Dim msg
		msg = Session(EW_SESSION_WARNING_MESSAGE)
		Call ew_AddMessage(msg, v)
		Session(EW_SESSION_WARNING_MESSAGE) = msg
	End Property

	' Show Message
	Public Sub ShowMessage()
		Dim hidden, html, sMessage
		hidden = False
		html = ""

		' Message
		sMessage = Message
		Call Message_Showing(sMessage, "")
		If sMessage <> "" Then ' Message in Session, display
			If Not hidden Then sMessage = "<button type=""button"" class=""close"" data-dismiss=""alert"">&times;</button>" & sMessage
			html = html & "<div class=""alert alert-success ewSuccess"">" & sMessage & "</div>"
			Session(EW_SESSION_MESSAGE) = "" ' Clear message in Session
		End If

		' Warning message
		Dim sWarningMessage
		sWarningMessage = WarningMessage
		Call Message_Showing(sWarningMessage, "warning")
		If sWarningMessage <> "" Then ' Message in Session, display
			If Not hidden Then sWarningMessage = "<button type=""button"" class=""close"" data-dismiss=""alert"">&times;</button>" & sWarningMessage
			html = html & "<div class=""alert alert-warning ewWarning"">" & sWarningMessage & "</div>"
			Session(EW_SESSION_WARNING_MESSAGE) = "" ' Clear message in Session
		End If

		' Success message
		Dim sSuccessMessage
		sSuccessMessage = SuccessMessage
		Call Message_Showing(sSuccessMessage, "success")
		If sSuccessMessage <> "" Then ' Message in Session, display
			If Not hidden Then sSuccessMessage = "<button type=""button"" class=""close"" data-dismiss=""alert"">&times;</button>" & sSuccessMessage
			html = html & "<div class=""alert alert-success ewSuccess"">" & sSuccessMessage & "</div>"
			Session(EW_SESSION_SUCCESS_MESSAGE) = "" ' Clear message in Session
		End If

		' Failure message
		Dim sErrorMessage
		sErrorMessage = FailureMessage
		Call Message_Showing(sErrorMessage, "failure")
		If sErrorMessage <> "" Then ' Message in Session, display
			If Not hidden Then sErrorMessage = "<button type=""button"" class=""close"" data-dismiss=""alert"">&times;</button>" & sErrorMessage
			html = html & "<div class=""alert alert-error ewError"">" & sErrorMessage & "</div>"
			Session(EW_SESSION_FAILURE_MESSAGE) = "" ' Clear message in Session
		End If
		Response.Write "<table class=""ewStdTable""><tr><td><div class=""ewMessageDialog""" & ew_IIf(hidden, " style=""display: none;""", "") & ">" & html & "</div></td></tr></table>"
	End Sub
	Dim PageHeader
	Dim PageFooter

	' Show Page Header
	Public Sub ShowPageHeader()
		Dim sHeader
		sHeader = PageHeader
		Call Page_DataRendering(sHeader)
		If sHeader <> "" Then ' Header exists, display
			Response.Write "<p>" & sHeader & "</p>"
		End If
	End Sub

	' Show Page Footer
	Public Sub ShowPageFooter()
		Dim sFooter
		sFooter = PageFooter
		Call Page_DataRendered(sFooter)
		If sFooter <> "" Then ' Footer exists, display
			Response.Write "<p>" & sFooter & "</p>"
		End If
	End Sub

	' -----------------------
	'  Validate Page request
	'
	Public Function IsPageRequest()
		If department.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (department.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (department.TableVar = Request.QueryString("t"))
			End If
		Else
			IsPageRequest = True
		End If
	End Function

	' -----------------------------------------------------------------
	'  Class initialize
	'  - init objects
	'  - open ADO connection
	'
	Private Sub Class_Initialize()
		If IsEmpty(StartTimer) Then StartTimer = Timer ' Init start time

		' Initialize language object
		If IsEmpty(Language) Then
			Set Language = New cLanguage
			Call Language.LoadPhrases()
		End If

		' Initialize table object
		If IsEmpty(department) Then Set department = New cdepartment
		Set Table = department

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "department"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()
	End Sub

	' -----------------------------------------------------------------
	'  Subroutine Page_Init
	'  - called before page main
	'  - check Security
	'  - set up response header
	'  - call page load events
	'
	Sub Page_Init()
		Set Security = New cAdvancedSecurity
		If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
		If Not Security.IsLoggedIn() Then
			Call Security.SaveLastUrl()
			Call Page_Terminate("pom_login.asp")
		End If

		' Create form object
		'If Request.ServerVariables("HTTP_CONTENT_TYPE") = "application/x-www-form-urlencoded" Then

			Set ObjForm = New cFormObj

		'Else
		'	Set ObjForm = ew_GetUploadObj()
		'End If

		department.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

		' Global page loading event (in userfn7.asp)
		Page_Loading()

		' Page load event, used in current page
		Page_Load()
	End Sub

	' -----------------------------------------------------------------
	'  Class terminate
	'  - clean up page object
	'
	Private Sub Class_Terminate()
		Call Page_Terminate("")
	End Sub

	' -----------------------------------------------------------------
	'  Subroutine Page_Terminate
	'  - called when exit page
	'  - clean up ADO connection and objects
	'  - if url specified, redirect to url
	'
	Sub Page_Terminate(url)

		' Page unload event, used in current page
		Call Page_Unload()

		' Global page unloaded event (in userfn60.asp)
		Call Page_Unloaded()
		Dim sRedirectUrl
		sReDirectUrl = url
		Call Page_Redirecting(sReDirectUrl)
		If Not (Conn Is Nothing) Then Conn.Close ' Close Connection
		Set Conn = Nothing
		Set Security = Nothing
		Set department = Nothing
		Set ObjForm = Nothing

		' Go to url if specified
		If sReDirectUrl <> "" Then
			If Response.Buffer Then Response.Clear
			Response.Redirect sReDirectUrl
		End If
	End Sub

	'
	'  Subroutine Page_Terminate (End)
	' ----------------------------------------

	Dim DbMasterFilter, DbDetailFilter

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		Dim sReturnUrl
		sReturnUrl = ""

		' Load key from QueryString
		If Request.QueryString("dept_id").Count > 0 Then
			department.dept_id.QueryStringValue = Request.QueryString("dept_id")
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Process form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			department.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values
		Else
			department.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If department.dept_id.CurrentValue = "" Then Call Page_Terminate("pom_departmentlist.asp") ' Invalid key, return to list

		' Validate form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			If Not ValidateForm() Then
				department.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				department.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		End If
		Select Case department.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_departmentlist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				department.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					sReturnUrl = department.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					department.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		department.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call department.ResetAttrs()
		Call RenderRow()
	End Sub
	Dim Pager

	' -----------------------------------------------------------------
	' Set up Starting Record parameters based on Pager Navigation
	'
	Sub SetUpStartRec()
		Dim PageNo

		' Exit if DisplayRecs = 0
		If DisplayRecs = 0 Then Exit Sub
		If IsPageRequest Then ' Validate request

			' Check for a START parameter
			If Request.QueryString(EW_TABLE_START_REC).Count > 0 Then
				StartRec = Request.QueryString(EW_TABLE_START_REC)
				department.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					department.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = department.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			department.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			department.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			department.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Function Get upload files
	'
	Function GetUploadFiles()

		' Get upload data
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not department.dept_id.FldIsDetailKey Then department.dept_id.FormValue = ObjForm.GetValue("x_dept_id")
		If Not department.office_id.FldIsDetailKey Then department.office_id.FormValue = ObjForm.GetValue("x_office_id")
		If Not department.dept_name.FldIsDetailKey Then department.dept_name.FormValue = ObjForm.GetValue("x_dept_name")
		If Not department.dept_sort.FldIsDetailKey Then department.dept_sort.FormValue = ObjForm.GetValue("x_dept_sort")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		department.dept_id.CurrentValue = department.dept_id.FormValue
		department.office_id.CurrentValue = department.office_id.FormValue
		department.dept_name.CurrentValue = department.dept_name.FormValue
		department.dept_sort.CurrentValue = department.dept_sort.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = department.KeyFilter

		' Call Row Selecting event
		Call department.Row_Selecting(sFilter)

		' Load sql based on filter
		department.CurrentFilter = sFilter
		sSql = department.SQL
		Call ew_SetDebugMsg("LoadRow: " & sSql) ' Show SQL for debugging
		Set RsRow = ew_LoadRow(sSql)
		If RsRow.Eof Then
			LoadRow = False
		Else
			LoadRow = True
			RsRow.MoveFirst
			Call LoadRowValues(RsRow) ' Load row values
		End If
		RsRow.Close
		Set RsRow = Nothing
	End Function

	' -----------------------------------------------------------------
	' Load row values from recordset
	'
	Sub LoadRowValues(RsRow)
		Dim sDetailFilter
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If RsRow.Eof Then Exit Sub

		' Call Row Selected event
		Call department.Row_Selected(RsRow)
		department.dept_id.DbValue = RsRow("dept_id")
		department.office_id.DbValue = RsRow("office_id")
		department.dept_name.DbValue = RsRow("dept_name")
		department.dept_sort.DbValue = RsRow("dept_sort")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		department.dept_id.m_DbValue = Rs("dept_id")
		department.office_id.m_DbValue = Rs("office_id")
		department.dept_name.m_DbValue = Rs("dept_name")
		department.dept_sort.m_DbValue = Rs("dept_sort")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call department.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' dept_id
		' office_id
		' dept_name
		' dept_sort
		' -----------
		'  View  Row
		' -----------

		If department.RowType = EW_ROWTYPE_VIEW Then ' View row

			' dept_id
			department.dept_id.ViewValue = department.dept_id.CurrentValue
			department.dept_id.ViewCustomAttributes = ""

			' office_id
			department.office_id.ViewValue = department.office_id.CurrentValue
			department.office_id.ViewCustomAttributes = ""

			' dept_name
			department.dept_name.ViewValue = department.dept_name.CurrentValue
			department.dept_name.ViewCustomAttributes = ""

			' dept_sort
			department.dept_sort.ViewValue = department.dept_sort.CurrentValue
			department.dept_sort.ViewCustomAttributes = ""

			' View refer script
			' dept_id

			department.dept_id.LinkCustomAttributes = ""
			department.dept_id.HrefValue = ""
			department.dept_id.TooltipValue = ""

			' office_id
			department.office_id.LinkCustomAttributes = ""
			department.office_id.HrefValue = ""
			department.office_id.TooltipValue = ""

			' dept_name
			department.dept_name.LinkCustomAttributes = ""
			department.dept_name.HrefValue = ""
			department.dept_name.TooltipValue = ""

			' dept_sort
			department.dept_sort.LinkCustomAttributes = ""
			department.dept_sort.HrefValue = ""
			department.dept_sort.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf department.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' dept_id
			department.dept_id.EditCustomAttributes = ""
			department.dept_id.EditValue = department.dept_id.CurrentValue
			department.dept_id.ViewCustomAttributes = ""

			' office_id
			department.office_id.EditCustomAttributes = ""
			department.office_id.EditValue = ew_HtmlEncode(department.office_id.CurrentValue)
			department.office_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(department.office_id.FldCaption))

			' dept_name
			department.dept_name.EditCustomAttributes = ""
			department.dept_name.EditValue = ew_HtmlEncode(department.dept_name.CurrentValue)
			department.dept_name.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(department.dept_name.FldCaption))

			' dept_sort
			department.dept_sort.EditCustomAttributes = ""
			department.dept_sort.EditValue = ew_HtmlEncode(department.dept_sort.CurrentValue)
			department.dept_sort.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(department.dept_sort.FldCaption))

			' Edit refer script
			' dept_id

			department.dept_id.HrefValue = ""

			' office_id
			department.office_id.HrefValue = ""

			' dept_name
			department.dept_name.HrefValue = ""

			' dept_sort
			department.dept_sort.HrefValue = ""
		End If
		If department.RowType = EW_ROWTYPE_ADD Or department.RowType = EW_ROWTYPE_EDIT Or department.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call department.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If department.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call department.Row_Rendered()
		End If
	End Sub

	' -----------------------------------------------------------------
	' Validate form
	'
	Function ValidateForm()

		' Initialize
		gsFormError = ""

		' Check if validation required
		If Not EW_SERVER_VALIDATE Then
			ValidateForm = (gsFormError = "")
			Exit Function
		End If
		If Not ew_CheckInteger(department.dept_id.FormValue) Then
			Call ew_AddMessage(gsFormError, department.dept_id.FldErrMsg)
		End If
		If Not ew_CheckInteger(department.office_id.FormValue) Then
			Call ew_AddMessage(gsFormError, department.office_id.FldErrMsg)
		End If
		If Not ew_CheckInteger(department.dept_sort.FormValue) Then
			Call ew_AddMessage(gsFormError, department.dept_sort.FldErrMsg)
		End If

		' Return validate result
		ValidateForm = (gsFormError = "")

		' Call Form Custom Validate event
		Dim sFormCustomError
		sFormCustomError = ""
		ValidateForm = ValidateForm And Form_CustomValidate(sFormCustomError)
		If sFormCustomError <> "" Then
			Call ew_AddMessage(gsFormError, sFormCustomError)
		End If
	End Function

	' -----------------------------------------------------------------
	' Update record based on key values
	'
	Function EditRow()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim Rs, sSql, sFilter
		Dim RsChk, sSqlChk, sFilterChk
		Dim bUpdateRow
		Dim RsOld, RsNew
		Dim sIdxErrMsg

		' Clear any previous errors
		Err.Clear
		sFilter = department.KeyFilter
		department.CurrentFilter  = sFilter
		sSql = department.SQL
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = EW_CURSORLOCATION
		Rs.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		If Err.Number <> 0 Then
			Message = Err.Description
			Rs.Close
			Set Rs = Nothing
			EditRow = False
			Exit Function
		End If

		' Clone old recordset object
		Set RsOld = ew_CloneRs(Rs)
		Call LoadDbValues(RsOld)
		If Rs.Eof Then
			EditRow = False ' Update Failed
		Else

			' Field dept_id
			' Field office_id

			Call department.office_id.SetDbValue(Rs, department.office_id.CurrentValue, Null, department.office_id.ReadOnly)

			' Field dept_name
			Call department.dept_name.SetDbValue(Rs, department.dept_name.CurrentValue, Null, department.dept_name.ReadOnly)

			' Field dept_sort
			Call department.dept_sort.SetDbValue(Rs, department.dept_sort.CurrentValue, Null, department.dept_sort.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = department.Row_Updating(RsOld, Rs)
			If bUpdateRow Then

				' Clone new recordset object
				Set RsNew = ew_CloneRs(Rs)
				EditRow = True
				If EditRow Then
					Rs.Update
				End If
				If Err.Number <> 0 Or Not EditRow Then
					If Err.Description <> "" Then FailureMessage = Err.Description
					EditRow = False
				Else
					EditRow = True
				End If
				If EditRow Then
				End If
			Else
				Rs.CancelUpdate

				' Set up error message
				If SuccessMessage <> "" Or FailureMessage <> "" Then

					' Use the message, do nothing
				ElseIf department.CancelMessage <> "" Then
					FailureMessage = department.CancelMessage
					department.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call department.Row_Updated(RsOld, RsNew)
		End If
		Rs.Close
		Set Rs = Nothing
		If IsObject(RsOld) Then
			RsOld.Close
			Set RsOld = Nothing
		End If
		If IsObject(RsNew) Then
			RsNew.Close
			Set RsNew = Nothing
		End If
	End Function

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", department.TableVar, "pom_departmentlist.asp", department.TableVar, True)
		PageId = "edit"
		Call Breadcrumb.Add("edit", PageId, ew_CurrentUrl, "", False)
	End Sub

	' Page Load event
	Sub Page_Load()

		'Response.Write "Page Load"
	End Sub

	' Page Unload event
	Sub Page_Unload()

		'Response.Write "Page Unload"
	End Sub

	' Page Redirecting event
	Sub Page_Redirecting(url)

		'url = newurl
	End Sub

	' Message Showing event
	' typ = ""|"success"|"failure"|"warning"
	Sub Message_Showing(msg, typ)

		' Example:
		'If typ = "success" Then
		'	msg = "your success message"
		'ElseIf typ = "failure" Then
		'	msg = "your failure message"
		'ElseIf typ = "warning" Then
		'	msg = "your warning message"
		'Else
		'	msg = "your message"
		'End If

	End Sub

	' Page Render event
	Sub Page_Render()

		'Response.Write "Page Render"
	End Sub

	' Page Data Rendering event
	Sub Page_DataRendering(header)

		' Example:
		'header = "your header"

	End Sub

	' Page Data Rendered event
	Sub Page_DataRendered(footer)

		' Example:
		'footer = "your footer"

	End Sub

	' Form Custom Validate event
	Function Form_CustomValidate(CustomError)

		'Return error message in CustomError
		Form_CustomValidate = True
	End Function
End Class
%>
