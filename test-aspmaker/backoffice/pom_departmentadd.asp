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
Dim department_add
Set department_add = New cdepartment_add
Set Page = department_add

' Page init processing
department_add.Page_Init()

' Page main processing
department_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
department_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var department_add = new ew_Page("department_add");
department_add.PageID = "add"; // Page ID
var EW_PAGE_ID = department_add.PageID; // For backward compatibility
// Form object
var fdepartmentadd = new ew_Form("fdepartmentadd");
// Validate form
fdepartmentadd.Validate = function() {
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
fdepartmentadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fdepartmentadd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fdepartmentadd.ValidateRequired = false; // No JavaScript validation
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
<% department_add.ShowPageHeader() %>
<% department_add.ShowMessage %>
<form name="fdepartmentadd" id="fdepartmentadd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="department">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_departmentadd" class="table table-bordered table-striped">
<% If department.dept_id.Visible Then ' dept_id %>
	<tr id="r_dept_id">
		<td><span id="elh_department_dept_id"><%= department.dept_id.FldCaption %></span></td>
		<td<%= department.dept_id.CellAttributes %>>
<span id="el_department_dept_id" class="control-group">
<input type="text" data-field="x_dept_id" name="x_dept_id" id="x_dept_id" size="30" placeholder="<%= department.dept_id.PlaceHolder %>" value="<%= department.dept_id.EditValue %>"<%= department.dept_id.EditAttributes %>>
</span>
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
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
fdepartmentadd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
department_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set department_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cdepartment_add

	' Page ID
	Public Property Get PageID()
		PageID = "add"
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
		PageObjName = "department_add"
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
		EW_PAGE_ID = "add"

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
	Dim Priv
	Dim OldRecordset
	Dim CopyRecord

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Process form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			department.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("dept_id").Count > 0 Then
				department.dept_id.QueryStringValue = Request.QueryString("dept_id")
				Call department.SetKey("dept_id", department.dept_id.CurrentValue) ' Set up key
			Else
				Call department.SetKey("dept_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				department.CurrentAction = "C" ' Copy Record
			Else
				department.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				department.CurrentAction = "I" ' Form error, reset action
				department.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case department.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_departmentlist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				department.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = department.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_departmentview.asp" Then sReturnUrl = department.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					department.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		department.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call department.ResetAttrs()
		Call RenderRow()
	End Sub

	' -----------------------------------------------------------------
	' Function Get upload files
	'
	Function GetUploadFiles()

		' Get upload data
	End Function

	' -----------------------------------------------------------------
	' Load default values
	'
	Function LoadDefaultValues()
		department.dept_id.CurrentValue = Null
		department.dept_id.OldValue = department.dept_id.CurrentValue
		department.office_id.CurrentValue = Null
		department.office_id.OldValue = department.office_id.CurrentValue
		department.dept_name.CurrentValue = Null
		department.dept_name.OldValue = department.dept_name.CurrentValue
		department.dept_sort.CurrentValue = Null
		department.dept_sort.OldValue = department.dept_sort.CurrentValue
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
		Call LoadOldRecord()
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

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If department.GetKey("dept_id")&"" <> "" Then
			department.dept_id.CurrentValue = department.GetKey("dept_id") ' dept_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			department.CurrentFilter = department.KeyFilter
			Dim sSql
			sSql = department.SQL
			Set OldRecordset = ew_LoadRecordset(sSql)
			Call LoadRowValues(OldRecordset) ' Load row values
		Else
			OldRecordset = Null
		End If
		LoadOldRecord = bValidKey
	End Function

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

		' ---------
		'  Add Row
		' ---------

		ElseIf department.RowType = EW_ROWTYPE_ADD Then ' Add row

			' dept_id
			department.dept_id.EditCustomAttributes = ""
			department.dept_id.EditValue = ew_HtmlEncode(department.dept_id.CurrentValue)
			department.dept_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(department.dept_id.FldCaption))

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
	' Add record
	'
	Function AddRow(RsOld)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim Rs, sSql, sFilter
		Dim RsNew
		Dim bInsertRow
		Dim RsChk
		Dim sIdxErrMsg

		' Clear any previous errors
		Err.Clear
		Dim RsMaster, sMasterUserIdMsg, sMasterFilter, bCheckMasterRecord
		If department.dept_id.CurrentValue <> "" Then ' Check field with unique index
			sFilter = "([dept_id] = " & ew_AdjustSql(department.dept_id.CurrentValue) & ")"
			Set RsChk = department.LoadRs(sFilter)
			If Not (RsChk Is Nothing) Then
				sIdxErrMsg = Replace(Language.Phrase("DupIndex"), "%f", department.dept_id.FldCaption)
				sIdxErrMsg = Replace(sIdxErrMsg, "%v", department.dept_id.CurrentValue)
				FailureMessage = sIdxErrMsg
				RsChk.Close
				Set RsChk = Nothing
				AddRow = False
				Exit Function
			End If
		End If

		' Load db values from rsold
		If Not IsNull(RsOld) Then
			Call LoadDbValues(RsOld)
		End If

		' Add new record
		sFilter = "(0 = 1)"
		department.CurrentFilter = sFilter
		sSql = department.SQL
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = EW_CURSORLOCATION
		Rs.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		Rs.AddNew
		If Err.Number <> 0 Then
			Message = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Field dept_id
		Call department.dept_id.SetDbValue(Rs, department.dept_id.CurrentValue, Null, False)

		' Field office_id
		Call department.office_id.SetDbValue(Rs, department.office_id.CurrentValue, Null, False)

		' Field dept_name
		Call department.dept_name.SetDbValue(Rs, department.dept_name.CurrentValue, Null, False)

		' Field dept_sort
		Call department.dept_sort.SetDbValue(Rs, department.dept_sort.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = department.Row_Inserting(RsOld, Rs)

		' Check if key value entered
		If bInsertRow And department.ValidateKey And department.dept_id.CurrentValue = "" And department.dept_id.SessionValue = "" Then
			FailureMessage = Language.Phrase("InvalidKeyValue")
			bInsertRow = False
		End If

		' Check for duplicate key
		Dim sKeyErrMsg
		If bInsertRow And department.ValidateKey Then
			sFilter = department.KeyFilter
			Set RsChk = department.LoadRs(sFilter)
			If Not (RsChk Is Nothing) Then
				sKeyErrMsg = Replace(Language.Phrase("DupKey"), "%f", sFilter)
				FailureMessage = sKeyErrMsg
				RsChk.Close
				Set RsChk = Nothing
				bInsertRow = False
			End If
		End If
		If bInsertRow Then

			' Clone new recordset object
			Set RsNew = ew_CloneRs(Rs)
			Rs.Update
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				AddRow = False
			Else
				AddRow = True
			End If
			If AddRow Then
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
				FailureMessage = Language.Phrase("InsertCancelled")
			End If
			AddRow = False
		End If
		Rs.Close
		Set Rs = Nothing
		If AddRow Then
		End If
		If AddRow Then

			' Call Row Inserted event
			Call department.Row_Inserted(RsOld, RsNew)
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
		PageId = ew_IIf(department.CurrentAction = "C", "Copy", "Add")
		Call Breadcrumb.Add("add", PageId, ew_CurrentUrl, "", False)
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
