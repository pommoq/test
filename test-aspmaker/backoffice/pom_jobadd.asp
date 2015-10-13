<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_jobinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim job_add
Set job_add = New cjob_add
Set Page = job_add

' Page init processing
job_add.Page_Init()

' Page main processing
job_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
job_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var job_add = new ew_Page("job_add");
job_add.PageID = "add"; // Page ID
var EW_PAGE_ID = job_add.PageID; // For backward compatibility
// Form object
var fjobadd = new ew_Form("fjobadd");
// Validate form
fjobadd.Validate = function() {
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
				return this.OnError(elm, "<%= ew_JsEncode2(job.job_id.FldErrMsg) %>");
			elm = this.GetElements("x" + infix + "_company_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(job.company_id.FldErrMsg) %>");
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
fjobadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fjobadd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fjobadd.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If job.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% job_add.ShowPageHeader() %>
<% job_add.ShowMessage %>
<form name="fjobadd" id="fjobadd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="job">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_jobadd" class="table table-bordered table-striped">
<% If job.job_id.Visible Then ' job_id %>
	<tr id="r_job_id">
		<td><span id="elh_job_job_id"><%= job.job_id.FldCaption %></span></td>
		<td<%= job.job_id.CellAttributes %>>
<span id="el_job_job_id" class="control-group">
<input type="text" data-field="x_job_id" name="x_job_id" id="x_job_id" size="30" placeholder="<%= job.job_id.PlaceHolder %>" value="<%= job.job_id.EditValue %>"<%= job.job_id.EditAttributes %>>
</span>
<%= job.job_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If job.company_id.Visible Then ' company_id %>
	<tr id="r_company_id">
		<td><span id="elh_job_company_id"><%= job.company_id.FldCaption %></span></td>
		<td<%= job.company_id.CellAttributes %>>
<span id="el_job_company_id" class="control-group">
<input type="text" data-field="x_company_id" name="x_company_id" id="x_company_id" size="30" placeholder="<%= job.company_id.PlaceHolder %>" value="<%= job.company_id.EditValue %>"<%= job.company_id.EditAttributes %>>
</span>
<%= job.company_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If job.job_date.Visible Then ' job_date %>
	<tr id="r_job_date">
		<td><span id="elh_job_job_date"><%= job.job_date.FldCaption %></span></td>
		<td<%= job.job_date.CellAttributes %>>
<span id="el_job_job_date" class="control-group">
<input type="text" data-field="x_job_date" name="x_job_date" id="x_job_date" placeholder="<%= job.job_date.PlaceHolder %>" value="<%= job.job_date.EditValue %>"<%= job.job_date.EditAttributes %>>
</span>
<%= job.job_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If job.job_title.Visible Then ' job_title %>
	<tr id="r_job_title">
		<td><span id="elh_job_job_title"><%= job.job_title.FldCaption %></span></td>
		<td<%= job.job_title.CellAttributes %>>
<span id="el_job_job_title" class="control-group">
<input type="text" data-field="x_job_title" name="x_job_title" id="x_job_title" size="30" maxlength="255" placeholder="<%= job.job_title.PlaceHolder %>" value="<%= job.job_title.EditValue %>"<%= job.job_title.EditAttributes %>>
</span>
<%= job.job_title.CustomMsg %></td>
	</tr>
<% End If %>
<% If job.job_intro.Visible Then ' job_intro %>
	<tr id="r_job_intro">
		<td><span id="elh_job_job_intro"><%= job.job_intro.FldCaption %></span></td>
		<td<%= job.job_intro.CellAttributes %>>
<span id="el_job_job_intro" class="control-group">
<textarea data-field="x_job_intro" name="x_job_intro" id="x_job_intro" cols="35" rows="4" placeholder="<%= job.job_intro.PlaceHolder %>"<%= job.job_intro.EditAttributes %>><%= job.job_intro.EditValue %></textarea>
</span>
<%= job.job_intro.CustomMsg %></td>
	</tr>
<% End If %>
<% If job.job_detail.Visible Then ' job_detail %>
	<tr id="r_job_detail">
		<td><span id="elh_job_job_detail"><%= job.job_detail.FldCaption %></span></td>
		<td<%= job.job_detail.CellAttributes %>>
<span id="el_job_job_detail" class="control-group">
<textarea data-field="x_job_detail" name="x_job_detail" id="x_job_detail" cols="35" rows="4" placeholder="<%= job.job_detail.PlaceHolder %>"<%= job.job_detail.EditAttributes %>><%= job.job_detail.EditValue %></textarea>
</span>
<%= job.job_detail.CustomMsg %></td>
	</tr>
<% End If %>
<% If job.job_create.Visible Then ' job_create %>
	<tr id="r_job_create">
		<td><span id="elh_job_job_create"><%= job.job_create.FldCaption %></span></td>
		<td<%= job.job_create.CellAttributes %>>
<span id="el_job_job_create" class="control-group">
<input type="text" data-field="x_job_create" name="x_job_create" id="x_job_create" size="30" maxlength="255" placeholder="<%= job.job_create.PlaceHolder %>" value="<%= job.job_create.EditValue %>"<%= job.job_create.EditAttributes %>>
</span>
<%= job.job_create.CustomMsg %></td>
	</tr>
<% End If %>
<% If job.job_update.Visible Then ' job_update %>
	<tr id="r_job_update">
		<td><span id="elh_job_job_update"><%= job.job_update.FldCaption %></span></td>
		<td<%= job.job_update.CellAttributes %>>
<span id="el_job_job_update" class="control-group">
<input type="text" data-field="x_job_update" name="x_job_update" id="x_job_update" size="30" maxlength="255" placeholder="<%= job.job_update.PlaceHolder %>" value="<%= job.job_update.EditValue %>"<%= job.job_update.EditAttributes %>>
</span>
<%= job.job_update.CustomMsg %></td>
	</tr>
<% End If %>
<% If job.job_show.Visible Then ' job_show %>
	<tr id="r_job_show">
		<td><span id="elh_job_job_show"><%= job.job_show.FldCaption %></span></td>
		<td<%= job.job_show.CellAttributes %>>
<span id="el_job_job_show" class="control-group">
<input type="text" data-field="x_job_show" name="x_job_show" id="x_job_show" size="30" maxlength="1" placeholder="<%= job.job_show.PlaceHolder %>" value="<%= job.job_show.EditValue %>"<%= job.job_show.EditAttributes %>>
</span>
<%= job.job_show.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
fjobadd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
job_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set job_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cjob_add

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
		TableName = "job"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "job_add"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If job.UseTokenInUrl Then PageUrl = PageUrl & "t=" & job.TableVar & "&" ' add page token
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
		If job.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (job.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (job.TableVar = Request.QueryString("t"))
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
		If IsEmpty(job) Then Set job = New cjob
		Set Table = job

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "add"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "job"

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

		job.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set job = Nothing
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
			job.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("job_id").Count > 0 Then
				job.job_id.QueryStringValue = Request.QueryString("job_id")
				Call job.SetKey("job_id", job.job_id.CurrentValue) ' Set up key
			Else
				Call job.SetKey("job_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				job.CurrentAction = "C" ' Copy Record
			Else
				job.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				job.CurrentAction = "I" ' Form error, reset action
				job.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case job.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_joblist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				job.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = job.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_jobview.asp" Then sReturnUrl = job.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					job.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		job.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call job.ResetAttrs()
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
		job.job_id.CurrentValue = Null
		job.job_id.OldValue = job.job_id.CurrentValue
		job.company_id.CurrentValue = Null
		job.company_id.OldValue = job.company_id.CurrentValue
		job.job_date.CurrentValue = Null
		job.job_date.OldValue = job.job_date.CurrentValue
		job.job_title.CurrentValue = Null
		job.job_title.OldValue = job.job_title.CurrentValue
		job.job_intro.CurrentValue = Null
		job.job_intro.OldValue = job.job_intro.CurrentValue
		job.job_detail.CurrentValue = Null
		job.job_detail.OldValue = job.job_detail.CurrentValue
		job.job_create.CurrentValue = Null
		job.job_create.OldValue = job.job_create.CurrentValue
		job.job_update.CurrentValue = Null
		job.job_update.OldValue = job.job_update.CurrentValue
		job.job_show.CurrentValue = "N"
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not job.job_id.FldIsDetailKey Then job.job_id.FormValue = ObjForm.GetValue("x_job_id")
		If Not job.company_id.FldIsDetailKey Then job.company_id.FormValue = ObjForm.GetValue("x_company_id")
		If Not job.job_date.FldIsDetailKey Then job.job_date.FormValue = ObjForm.GetValue("x_job_date")
		If Not job.job_date.FldIsDetailKey Then job.job_date.CurrentValue = ew_UnFormatDateTime(job.job_date.CurrentValue, 8)
		If Not job.job_title.FldIsDetailKey Then job.job_title.FormValue = ObjForm.GetValue("x_job_title")
		If Not job.job_intro.FldIsDetailKey Then job.job_intro.FormValue = ObjForm.GetValue("x_job_intro")
		If Not job.job_detail.FldIsDetailKey Then job.job_detail.FormValue = ObjForm.GetValue("x_job_detail")
		If Not job.job_create.FldIsDetailKey Then job.job_create.FormValue = ObjForm.GetValue("x_job_create")
		If Not job.job_update.FldIsDetailKey Then job.job_update.FormValue = ObjForm.GetValue("x_job_update")
		If Not job.job_show.FldIsDetailKey Then job.job_show.FormValue = ObjForm.GetValue("x_job_show")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		job.job_id.CurrentValue = job.job_id.FormValue
		job.company_id.CurrentValue = job.company_id.FormValue
		job.job_date.CurrentValue = job.job_date.FormValue
		job.job_date.CurrentValue = ew_UnFormatDateTime(job.job_date.CurrentValue, 8)
		job.job_title.CurrentValue = job.job_title.FormValue
		job.job_intro.CurrentValue = job.job_intro.FormValue
		job.job_detail.CurrentValue = job.job_detail.FormValue
		job.job_create.CurrentValue = job.job_create.FormValue
		job.job_update.CurrentValue = job.job_update.FormValue
		job.job_show.CurrentValue = job.job_show.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = job.KeyFilter

		' Call Row Selecting event
		Call job.Row_Selecting(sFilter)

		' Load sql based on filter
		job.CurrentFilter = sFilter
		sSql = job.SQL
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
		Call job.Row_Selected(RsRow)
		job.job_id.DbValue = RsRow("job_id")
		job.company_id.DbValue = RsRow("company_id")
		job.job_date.DbValue = RsRow("job_date")
		job.job_title.DbValue = RsRow("job_title")
		job.job_intro.DbValue = RsRow("job_intro")
		job.job_detail.DbValue = RsRow("job_detail")
		job.job_create.DbValue = RsRow("job_create")
		job.job_update.DbValue = RsRow("job_update")
		job.job_show.DbValue = RsRow("job_show")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		job.job_id.m_DbValue = Rs("job_id")
		job.company_id.m_DbValue = Rs("company_id")
		job.job_date.m_DbValue = Rs("job_date")
		job.job_title.m_DbValue = Rs("job_title")
		job.job_intro.m_DbValue = Rs("job_intro")
		job.job_detail.m_DbValue = Rs("job_detail")
		job.job_create.m_DbValue = Rs("job_create")
		job.job_update.m_DbValue = Rs("job_update")
		job.job_show.m_DbValue = Rs("job_show")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If job.GetKey("job_id")&"" <> "" Then
			job.job_id.CurrentValue = job.GetKey("job_id") ' job_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			job.CurrentFilter = job.KeyFilter
			Dim sSql
			sSql = job.SQL
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

		Call job.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' job_id
		' company_id
		' job_date
		' job_title
		' job_intro
		' job_detail
		' job_create
		' job_update
		' job_show
		' -----------
		'  View  Row
		' -----------

		If job.RowType = EW_ROWTYPE_VIEW Then ' View row

			' job_id
			job.job_id.ViewValue = job.job_id.CurrentValue
			job.job_id.ViewCustomAttributes = ""

			' company_id
			job.company_id.ViewValue = job.company_id.CurrentValue
			job.company_id.ViewCustomAttributes = ""

			' job_date
			job.job_date.ViewValue = job.job_date.CurrentValue
			job.job_date.ViewCustomAttributes = ""

			' job_title
			job.job_title.ViewValue = job.job_title.CurrentValue
			job.job_title.ViewCustomAttributes = ""

			' job_intro
			job.job_intro.ViewValue = job.job_intro.CurrentValue
			job.job_intro.ViewCustomAttributes = ""

			' job_detail
			job.job_detail.ViewValue = job.job_detail.CurrentValue
			job.job_detail.ViewCustomAttributes = ""

			' job_create
			job.job_create.ViewValue = job.job_create.CurrentValue
			job.job_create.ViewCustomAttributes = ""

			' job_update
			job.job_update.ViewValue = job.job_update.CurrentValue
			job.job_update.ViewCustomAttributes = ""

			' job_show
			job.job_show.ViewValue = job.job_show.CurrentValue
			job.job_show.ViewCustomAttributes = ""

			' View refer script
			' job_id

			job.job_id.LinkCustomAttributes = ""
			job.job_id.HrefValue = ""
			job.job_id.TooltipValue = ""

			' company_id
			job.company_id.LinkCustomAttributes = ""
			job.company_id.HrefValue = ""
			job.company_id.TooltipValue = ""

			' job_date
			job.job_date.LinkCustomAttributes = ""
			job.job_date.HrefValue = ""
			job.job_date.TooltipValue = ""

			' job_title
			job.job_title.LinkCustomAttributes = ""
			job.job_title.HrefValue = ""
			job.job_title.TooltipValue = ""

			' job_intro
			job.job_intro.LinkCustomAttributes = ""
			job.job_intro.HrefValue = ""
			job.job_intro.TooltipValue = ""

			' job_detail
			job.job_detail.LinkCustomAttributes = ""
			job.job_detail.HrefValue = ""
			job.job_detail.TooltipValue = ""

			' job_create
			job.job_create.LinkCustomAttributes = ""
			job.job_create.HrefValue = ""
			job.job_create.TooltipValue = ""

			' job_update
			job.job_update.LinkCustomAttributes = ""
			job.job_update.HrefValue = ""
			job.job_update.TooltipValue = ""

			' job_show
			job.job_show.LinkCustomAttributes = ""
			job.job_show.HrefValue = ""
			job.job_show.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf job.RowType = EW_ROWTYPE_ADD Then ' Add row

			' job_id
			job.job_id.EditCustomAttributes = ""
			job.job_id.EditValue = ew_HtmlEncode(job.job_id.CurrentValue)
			job.job_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job.job_id.FldCaption))

			' company_id
			job.company_id.EditCustomAttributes = ""
			job.company_id.EditValue = ew_HtmlEncode(job.company_id.CurrentValue)
			job.company_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job.company_id.FldCaption))

			' job_date
			job.job_date.EditCustomAttributes = ""
			job.job_date.EditValue = ew_HtmlEncode(job.job_date.CurrentValue)
			job.job_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job.job_date.FldCaption))

			' job_title
			job.job_title.EditCustomAttributes = ""
			job.job_title.EditValue = ew_HtmlEncode(job.job_title.CurrentValue)
			job.job_title.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job.job_title.FldCaption))

			' job_intro
			job.job_intro.EditCustomAttributes = ""
			job.job_intro.EditValue = job.job_intro.CurrentValue
			job.job_intro.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job.job_intro.FldCaption))

			' job_detail
			job.job_detail.EditCustomAttributes = ""
			job.job_detail.EditValue = job.job_detail.CurrentValue
			job.job_detail.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job.job_detail.FldCaption))

			' job_create
			job.job_create.EditCustomAttributes = ""
			job.job_create.EditValue = ew_HtmlEncode(job.job_create.CurrentValue)
			job.job_create.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job.job_create.FldCaption))

			' job_update
			job.job_update.EditCustomAttributes = ""
			job.job_update.EditValue = ew_HtmlEncode(job.job_update.CurrentValue)
			job.job_update.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job.job_update.FldCaption))

			' job_show
			job.job_show.EditCustomAttributes = ""
			job.job_show.EditValue = ew_HtmlEncode(job.job_show.CurrentValue)
			job.job_show.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job.job_show.FldCaption))

			' Edit refer script
			' job_id

			job.job_id.HrefValue = ""

			' company_id
			job.company_id.HrefValue = ""

			' job_date
			job.job_date.HrefValue = ""

			' job_title
			job.job_title.HrefValue = ""

			' job_intro
			job.job_intro.HrefValue = ""

			' job_detail
			job.job_detail.HrefValue = ""

			' job_create
			job.job_create.HrefValue = ""

			' job_update
			job.job_update.HrefValue = ""

			' job_show
			job.job_show.HrefValue = ""
		End If
		If job.RowType = EW_ROWTYPE_ADD Or job.RowType = EW_ROWTYPE_EDIT Or job.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call job.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If job.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call job.Row_Rendered()
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
		If Not ew_CheckInteger(job.job_id.FormValue) Then
			Call ew_AddMessage(gsFormError, job.job_id.FldErrMsg)
		End If
		If Not ew_CheckInteger(job.company_id.FormValue) Then
			Call ew_AddMessage(gsFormError, job.company_id.FldErrMsg)
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
		If job.job_id.CurrentValue <> "" Then ' Check field with unique index
			sFilter = "([job_id] = " & ew_AdjustSql(job.job_id.CurrentValue) & ")"
			Set RsChk = job.LoadRs(sFilter)
			If Not (RsChk Is Nothing) Then
				sIdxErrMsg = Replace(Language.Phrase("DupIndex"), "%f", job.job_id.FldCaption)
				sIdxErrMsg = Replace(sIdxErrMsg, "%v", job.job_id.CurrentValue)
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
		job.CurrentFilter = sFilter
		sSql = job.SQL
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
		Call job.job_id.SetDbValue(Rs, job.job_id.CurrentValue, Null, False)

		' Field company_id
		Call job.company_id.SetDbValue(Rs, job.company_id.CurrentValue, Null, False)

		' Field job_date
		Call job.job_date.SetDbValue(Rs, job.job_date.CurrentValue, Null, False)

		' Field job_title
		Call job.job_title.SetDbValue(Rs, job.job_title.CurrentValue, Null, False)

		' Field job_intro
		Call job.job_intro.SetDbValue(Rs, job.job_intro.CurrentValue, Null, False)

		' Field job_detail
		Call job.job_detail.SetDbValue(Rs, job.job_detail.CurrentValue, Null, False)

		' Field job_create
		Call job.job_create.SetDbValue(Rs, job.job_create.CurrentValue, Null, False)

		' Field job_update
		Call job.job_update.SetDbValue(Rs, job.job_update.CurrentValue, Null, False)

		' Field job_show
		Call job.job_show.SetDbValue(Rs, job.job_show.CurrentValue, Null, (job.job_show.CurrentValue&"" = ""))

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = job.Row_Inserting(RsOld, Rs)

		' Check if key value entered
		If bInsertRow And job.ValidateKey And job.job_id.CurrentValue = "" And job.job_id.SessionValue = "" Then
			FailureMessage = Language.Phrase("InvalidKeyValue")
			bInsertRow = False
		End If

		' Check for duplicate key
		Dim sKeyErrMsg
		If bInsertRow And job.ValidateKey Then
			sFilter = job.KeyFilter
			Set RsChk = job.LoadRs(sFilter)
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
			ElseIf job.CancelMessage <> "" Then
				FailureMessage = job.CancelMessage
				job.CancelMessage = ""
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
			Call job.Row_Inserted(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", job.TableVar, "pom_joblist.asp", job.TableVar, True)
		PageId = ew_IIf(job.CurrentAction = "C", "Copy", "Add")
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
