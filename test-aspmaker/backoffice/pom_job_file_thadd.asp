<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_job_file_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim job_file_th_add
Set job_file_th_add = New cjob_file_th_add
Set Page = job_file_th_add

' Page init processing
job_file_th_add.Page_Init()

' Page main processing
job_file_th_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
job_file_th_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var job_file_th_add = new ew_Page("job_file_th_add");
job_file_th_add.PageID = "add"; // Page ID
var EW_PAGE_ID = job_file_th_add.PageID; // For backward compatibility
// Form object
var fjob_file_thadd = new ew_Form("fjob_file_thadd");
// Validate form
fjob_file_thadd.Validate = function() {
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
			elm = this.GetElements("x" + infix + "_job_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(job_file_th.job_id.FldErrMsg) %>");
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
fjob_file_thadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fjob_file_thadd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fjob_file_thadd.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If job_file_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% job_file_th_add.ShowPageHeader() %>
<% job_file_th_add.ShowMessage %>
<form name="fjob_file_thadd" id="fjob_file_thadd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="job_file_th">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_job_file_thadd" class="table table-bordered table-striped">
<% If job_file_th.job_id.Visible Then ' job_id %>
	<tr id="r_job_id">
		<td><span id="elh_job_file_th_job_id"><%= job_file_th.job_id.FldCaption %></span></td>
		<td<%= job_file_th.job_id.CellAttributes %>>
<span id="el_job_file_th_job_id" class="control-group">
<input type="text" data-field="x_job_id" name="x_job_id" id="x_job_id" size="30" placeholder="<%= job_file_th.job_id.PlaceHolder %>" value="<%= job_file_th.job_id.EditValue %>"<%= job_file_th.job_id.EditAttributes %>>
</span>
<%= job_file_th.job_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If job_file_th.job_file_name.Visible Then ' job_file_name %>
	<tr id="r_job_file_name">
		<td><span id="elh_job_file_th_job_file_name"><%= job_file_th.job_file_name.FldCaption %></span></td>
		<td<%= job_file_th.job_file_name.CellAttributes %>>
<span id="el_job_file_th_job_file_name" class="control-group">
<input type="text" data-field="x_job_file_name" name="x_job_file_name" id="x_job_file_name" size="30" maxlength="255" placeholder="<%= job_file_th.job_file_name.PlaceHolder %>" value="<%= job_file_th.job_file_name.EditValue %>"<%= job_file_th.job_file_name.EditAttributes %>>
</span>
<%= job_file_th.job_file_name.CustomMsg %></td>
	</tr>
<% End If %>
<% If job_file_th.job_file_title.Visible Then ' job_file_title %>
	<tr id="r_job_file_title">
		<td><span id="elh_job_file_th_job_file_title"><%= job_file_th.job_file_title.FldCaption %></span></td>
		<td<%= job_file_th.job_file_title.CellAttributes %>>
<span id="el_job_file_th_job_file_title" class="control-group">
<input type="text" data-field="x_job_file_title" name="x_job_file_title" id="x_job_file_title" size="30" maxlength="255" placeholder="<%= job_file_th.job_file_title.PlaceHolder %>" value="<%= job_file_th.job_file_title.EditValue %>"<%= job_file_th.job_file_title.EditAttributes %>>
</span>
<%= job_file_th.job_file_title.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
fjob_file_thadd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
job_file_th_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set job_file_th_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cjob_file_th_add

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
		TableName = "job_file_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "job_file_th_add"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If job_file_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & job_file_th.TableVar & "&" ' add page token
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
		If job_file_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (job_file_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (job_file_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(job_file_th) Then Set job_file_th = New cjob_file_th
		Set Table = job_file_th

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "add"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "job_file_th"

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

		job_file_th.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set job_file_th = Nothing
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
			job_file_th.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("job_file_id").Count > 0 Then
				job_file_th.job_file_id.QueryStringValue = Request.QueryString("job_file_id")
				Call job_file_th.SetKey("job_file_id", job_file_th.job_file_id.CurrentValue) ' Set up key
			Else
				Call job_file_th.SetKey("job_file_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				job_file_th.CurrentAction = "C" ' Copy Record
			Else
				job_file_th.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				job_file_th.CurrentAction = "I" ' Form error, reset action
				job_file_th.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case job_file_th.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_job_file_thlist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				job_file_th.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = job_file_th.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_job_file_thview.asp" Then sReturnUrl = job_file_th.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					job_file_th.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		job_file_th.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call job_file_th.ResetAttrs()
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
		job_file_th.job_id.CurrentValue = Null
		job_file_th.job_id.OldValue = job_file_th.job_id.CurrentValue
		job_file_th.job_file_name.CurrentValue = Null
		job_file_th.job_file_name.OldValue = job_file_th.job_file_name.CurrentValue
		job_file_th.job_file_title.CurrentValue = Null
		job_file_th.job_file_title.OldValue = job_file_th.job_file_title.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not job_file_th.job_id.FldIsDetailKey Then job_file_th.job_id.FormValue = ObjForm.GetValue("x_job_id")
		If Not job_file_th.job_file_name.FldIsDetailKey Then job_file_th.job_file_name.FormValue = ObjForm.GetValue("x_job_file_name")
		If Not job_file_th.job_file_title.FldIsDetailKey Then job_file_th.job_file_title.FormValue = ObjForm.GetValue("x_job_file_title")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		job_file_th.job_id.CurrentValue = job_file_th.job_id.FormValue
		job_file_th.job_file_name.CurrentValue = job_file_th.job_file_name.FormValue
		job_file_th.job_file_title.CurrentValue = job_file_th.job_file_title.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = job_file_th.KeyFilter

		' Call Row Selecting event
		Call job_file_th.Row_Selecting(sFilter)

		' Load sql based on filter
		job_file_th.CurrentFilter = sFilter
		sSql = job_file_th.SQL
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
		Call job_file_th.Row_Selected(RsRow)
		job_file_th.job_file_id.DbValue = RsRow("job_file_id")
		job_file_th.job_id.DbValue = RsRow("job_id")
		job_file_th.job_file_name.DbValue = RsRow("job_file_name")
		job_file_th.job_file_title.DbValue = RsRow("job_file_title")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		job_file_th.job_file_id.m_DbValue = Rs("job_file_id")
		job_file_th.job_id.m_DbValue = Rs("job_id")
		job_file_th.job_file_name.m_DbValue = Rs("job_file_name")
		job_file_th.job_file_title.m_DbValue = Rs("job_file_title")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If job_file_th.GetKey("job_file_id")&"" <> "" Then
			job_file_th.job_file_id.CurrentValue = job_file_th.GetKey("job_file_id") ' job_file_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			job_file_th.CurrentFilter = job_file_th.KeyFilter
			Dim sSql
			sSql = job_file_th.SQL
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

		Call job_file_th.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' job_file_id
		' job_id
		' job_file_name
		' job_file_title
		' -----------
		'  View  Row
		' -----------

		If job_file_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' job_file_id
			job_file_th.job_file_id.ViewValue = job_file_th.job_file_id.CurrentValue
			job_file_th.job_file_id.ViewCustomAttributes = ""

			' job_id
			job_file_th.job_id.ViewValue = job_file_th.job_id.CurrentValue
			job_file_th.job_id.ViewCustomAttributes = ""

			' job_file_name
			job_file_th.job_file_name.ViewValue = job_file_th.job_file_name.CurrentValue
			job_file_th.job_file_name.ViewCustomAttributes = ""

			' job_file_title
			job_file_th.job_file_title.ViewValue = job_file_th.job_file_title.CurrentValue
			job_file_th.job_file_title.ViewCustomAttributes = ""

			' View refer script
			' job_id

			job_file_th.job_id.LinkCustomAttributes = ""
			job_file_th.job_id.HrefValue = ""
			job_file_th.job_id.TooltipValue = ""

			' job_file_name
			job_file_th.job_file_name.LinkCustomAttributes = ""
			job_file_th.job_file_name.HrefValue = ""
			job_file_th.job_file_name.TooltipValue = ""

			' job_file_title
			job_file_th.job_file_title.LinkCustomAttributes = ""
			job_file_th.job_file_title.HrefValue = ""
			job_file_th.job_file_title.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf job_file_th.RowType = EW_ROWTYPE_ADD Then ' Add row

			' job_id
			job_file_th.job_id.EditCustomAttributes = ""
			job_file_th.job_id.EditValue = ew_HtmlEncode(job_file_th.job_id.CurrentValue)
			job_file_th.job_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job_file_th.job_id.FldCaption))

			' job_file_name
			job_file_th.job_file_name.EditCustomAttributes = ""
			job_file_th.job_file_name.EditValue = ew_HtmlEncode(job_file_th.job_file_name.CurrentValue)
			job_file_th.job_file_name.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job_file_th.job_file_name.FldCaption))

			' job_file_title
			job_file_th.job_file_title.EditCustomAttributes = ""
			job_file_th.job_file_title.EditValue = ew_HtmlEncode(job_file_th.job_file_title.CurrentValue)
			job_file_th.job_file_title.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job_file_th.job_file_title.FldCaption))

			' Edit refer script
			' job_id

			job_file_th.job_id.HrefValue = ""

			' job_file_name
			job_file_th.job_file_name.HrefValue = ""

			' job_file_title
			job_file_th.job_file_title.HrefValue = ""
		End If
		If job_file_th.RowType = EW_ROWTYPE_ADD Or job_file_th.RowType = EW_ROWTYPE_EDIT Or job_file_th.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call job_file_th.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If job_file_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call job_file_th.Row_Rendered()
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
		If Not ew_CheckInteger(job_file_th.job_id.FormValue) Then
			Call ew_AddMessage(gsFormError, job_file_th.job_id.FldErrMsg)
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

		' Load db values from rsold
		If Not IsNull(RsOld) Then
			Call LoadDbValues(RsOld)
		End If

		' Add new record
		sFilter = "(0 = 1)"
		job_file_th.CurrentFilter = sFilter
		sSql = job_file_th.SQL
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

		' Field job_id
		Call job_file_th.job_id.SetDbValue(Rs, job_file_th.job_id.CurrentValue, Null, False)

		' Field job_file_name
		Call job_file_th.job_file_name.SetDbValue(Rs, job_file_th.job_file_name.CurrentValue, Null, False)

		' Field job_file_title
		Call job_file_th.job_file_title.SetDbValue(Rs, job_file_th.job_file_title.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = job_file_th.Row_Inserting(RsOld, Rs)
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
			ElseIf job_file_th.CancelMessage <> "" Then
				FailureMessage = job_file_th.CancelMessage
				job_file_th.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("InsertCancelled")
			End If
			AddRow = False
		End If
		Rs.Close
		Set Rs = Nothing
		If AddRow Then
			job_file_th.job_file_id.DbValue = RsNew("job_file_id")
		End If
		If AddRow Then

			' Call Row Inserted event
			Call job_file_th.Row_Inserted(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", job_file_th.TableVar, "pom_job_file_thlist.asp", job_file_th.TableVar, True)
		PageId = ew_IIf(job_file_th.CurrentAction = "C", "Copy", "Add")
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
