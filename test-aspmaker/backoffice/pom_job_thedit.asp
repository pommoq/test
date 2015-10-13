<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_job_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim job_th_edit
Set job_th_edit = New cjob_th_edit
Set Page = job_th_edit

' Page init processing
job_th_edit.Page_Init()

' Page main processing
job_th_edit.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
job_th_edit.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var job_th_edit = new ew_Page("job_th_edit");
job_th_edit.PageID = "edit"; // Page ID
var EW_PAGE_ID = job_th_edit.PageID; // For backward compatibility
// Form object
var fjob_thedit = new ew_Form("fjob_thedit");
// Validate form
fjob_thedit.Validate = function() {
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
				return this.OnError(elm, "<%= ew_JsEncode2(job_th.job_id.FldErrMsg) %>");
			elm = this.GetElements("x" + infix + "_company_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(job_th.company_id.FldErrMsg) %>");
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
fjob_thedit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fjob_thedit.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fjob_thedit.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If job_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% job_th_edit.ShowPageHeader() %>
<% job_th_edit.ShowMessage %>
<form name="fjob_thedit" id="fjob_thedit" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="a_table" id="a_table" value="job_th">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table class="ewGrid"><tr><td>
<table id="tbl_job_thedit" class="table table-bordered table-striped">
<% If job_th.job_id.Visible Then ' job_id %>
	<tr id="r_job_id">
		<td><span id="elh_job_th_job_id"><%= job_th.job_id.FldCaption %></span></td>
		<td<%= job_th.job_id.CellAttributes %>>
<span id="el_job_th_job_id" class="control-group">
<span<%= job_th.job_id.ViewAttributes %>>
<%= job_th.job_id.EditValue %>
</span>
</span>
<input type="hidden" data-field="x_job_id" name="x_job_id" id="x_job_id" value="<%= Server.HTMLEncode(job_th.job_id.CurrentValue&"") %>">
<%= job_th.job_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If job_th.company_id.Visible Then ' company_id %>
	<tr id="r_company_id">
		<td><span id="elh_job_th_company_id"><%= job_th.company_id.FldCaption %></span></td>
		<td<%= job_th.company_id.CellAttributes %>>
<span id="el_job_th_company_id" class="control-group">
<input type="text" data-field="x_company_id" name="x_company_id" id="x_company_id" size="30" placeholder="<%= job_th.company_id.PlaceHolder %>" value="<%= job_th.company_id.EditValue %>"<%= job_th.company_id.EditAttributes %>>
</span>
<%= job_th.company_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If job_th.job_date.Visible Then ' job_date %>
	<tr id="r_job_date">
		<td><span id="elh_job_th_job_date"><%= job_th.job_date.FldCaption %></span></td>
		<td<%= job_th.job_date.CellAttributes %>>
<span id="el_job_th_job_date" class="control-group">
<input type="text" data-field="x_job_date" name="x_job_date" id="x_job_date" placeholder="<%= job_th.job_date.PlaceHolder %>" value="<%= job_th.job_date.EditValue %>"<%= job_th.job_date.EditAttributes %>>
</span>
<%= job_th.job_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If job_th.job_title.Visible Then ' job_title %>
	<tr id="r_job_title">
		<td><span id="elh_job_th_job_title"><%= job_th.job_title.FldCaption %></span></td>
		<td<%= job_th.job_title.CellAttributes %>>
<span id="el_job_th_job_title" class="control-group">
<input type="text" data-field="x_job_title" name="x_job_title" id="x_job_title" size="30" maxlength="255" placeholder="<%= job_th.job_title.PlaceHolder %>" value="<%= job_th.job_title.EditValue %>"<%= job_th.job_title.EditAttributes %>>
</span>
<%= job_th.job_title.CustomMsg %></td>
	</tr>
<% End If %>
<% If job_th.job_intro.Visible Then ' job_intro %>
	<tr id="r_job_intro">
		<td><span id="elh_job_th_job_intro"><%= job_th.job_intro.FldCaption %></span></td>
		<td<%= job_th.job_intro.CellAttributes %>>
<span id="el_job_th_job_intro" class="control-group">
<textarea data-field="x_job_intro" name="x_job_intro" id="x_job_intro" cols="35" rows="4" placeholder="<%= job_th.job_intro.PlaceHolder %>"<%= job_th.job_intro.EditAttributes %>><%= job_th.job_intro.EditValue %></textarea>
</span>
<%= job_th.job_intro.CustomMsg %></td>
	</tr>
<% End If %>
<% If job_th.job_detail.Visible Then ' job_detail %>
	<tr id="r_job_detail">
		<td><span id="elh_job_th_job_detail"><%= job_th.job_detail.FldCaption %></span></td>
		<td<%= job_th.job_detail.CellAttributes %>>
<span id="el_job_th_job_detail" class="control-group">
<textarea data-field="x_job_detail" name="x_job_detail" id="x_job_detail" cols="35" rows="4" placeholder="<%= job_th.job_detail.PlaceHolder %>"<%= job_th.job_detail.EditAttributes %>><%= job_th.job_detail.EditValue %></textarea>
</span>
<%= job_th.job_detail.CustomMsg %></td>
	</tr>
<% End If %>
<% If job_th.job_create.Visible Then ' job_create %>
	<tr id="r_job_create">
		<td><span id="elh_job_th_job_create"><%= job_th.job_create.FldCaption %></span></td>
		<td<%= job_th.job_create.CellAttributes %>>
<span id="el_job_th_job_create" class="control-group">
<input type="text" data-field="x_job_create" name="x_job_create" id="x_job_create" size="30" maxlength="255" placeholder="<%= job_th.job_create.PlaceHolder %>" value="<%= job_th.job_create.EditValue %>"<%= job_th.job_create.EditAttributes %>>
</span>
<%= job_th.job_create.CustomMsg %></td>
	</tr>
<% End If %>
<% If job_th.job_update.Visible Then ' job_update %>
	<tr id="r_job_update">
		<td><span id="elh_job_th_job_update"><%= job_th.job_update.FldCaption %></span></td>
		<td<%= job_th.job_update.CellAttributes %>>
<span id="el_job_th_job_update" class="control-group">
<input type="text" data-field="x_job_update" name="x_job_update" id="x_job_update" size="30" maxlength="255" placeholder="<%= job_th.job_update.PlaceHolder %>" value="<%= job_th.job_update.EditValue %>"<%= job_th.job_update.EditAttributes %>>
</span>
<%= job_th.job_update.CustomMsg %></td>
	</tr>
<% End If %>
<% If job_th.job_show.Visible Then ' job_show %>
	<tr id="r_job_show">
		<td><span id="elh_job_th_job_show"><%= job_th.job_show.FldCaption %></span></td>
		<td<%= job_th.job_show.CellAttributes %>>
<span id="el_job_th_job_show" class="control-group">
<input type="text" data-field="x_job_show" name="x_job_show" id="x_job_show" size="30" maxlength="1" placeholder="<%= job_th.job_show.PlaceHolder %>" value="<%= job_th.job_show.EditValue %>"<%= job_th.job_show.EditAttributes %>>
</span>
<%= job_th.job_show.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("EditBtn") %></button>
</form>
<script type="text/javascript">
fjob_thedit.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
job_th_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set job_th_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cjob_th_edit

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
		TableName = "job_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "job_th_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If job_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & job_th.TableVar & "&" ' add page token
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
		If job_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (job_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (job_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(job_th) Then Set job_th = New cjob_th
		Set Table = job_th

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "job_th"

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

		job_th.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set job_th = Nothing
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
		If Request.QueryString("job_id").Count > 0 Then
			job_th.job_id.QueryStringValue = Request.QueryString("job_id")
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Process form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			job_th.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values
		Else
			job_th.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If job_th.job_id.CurrentValue = "" Then Call Page_Terminate("pom_job_thlist.asp") ' Invalid key, return to list

		' Validate form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			If Not ValidateForm() Then
				job_th.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				job_th.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		End If
		Select Case job_th.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_job_thlist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				job_th.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					sReturnUrl = job_th.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					job_th.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		job_th.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call job_th.ResetAttrs()
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
				job_th.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					job_th.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = job_th.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			job_th.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			job_th.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			job_th.StartRecordNumber = StartRec
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
		If Not job_th.job_id.FldIsDetailKey Then job_th.job_id.FormValue = ObjForm.GetValue("x_job_id")
		If Not job_th.company_id.FldIsDetailKey Then job_th.company_id.FormValue = ObjForm.GetValue("x_company_id")
		If Not job_th.job_date.FldIsDetailKey Then job_th.job_date.FormValue = ObjForm.GetValue("x_job_date")
		If Not job_th.job_date.FldIsDetailKey Then job_th.job_date.CurrentValue = ew_UnFormatDateTime(job_th.job_date.CurrentValue, 8)
		If Not job_th.job_title.FldIsDetailKey Then job_th.job_title.FormValue = ObjForm.GetValue("x_job_title")
		If Not job_th.job_intro.FldIsDetailKey Then job_th.job_intro.FormValue = ObjForm.GetValue("x_job_intro")
		If Not job_th.job_detail.FldIsDetailKey Then job_th.job_detail.FormValue = ObjForm.GetValue("x_job_detail")
		If Not job_th.job_create.FldIsDetailKey Then job_th.job_create.FormValue = ObjForm.GetValue("x_job_create")
		If Not job_th.job_update.FldIsDetailKey Then job_th.job_update.FormValue = ObjForm.GetValue("x_job_update")
		If Not job_th.job_show.FldIsDetailKey Then job_th.job_show.FormValue = ObjForm.GetValue("x_job_show")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		job_th.job_id.CurrentValue = job_th.job_id.FormValue
		job_th.company_id.CurrentValue = job_th.company_id.FormValue
		job_th.job_date.CurrentValue = job_th.job_date.FormValue
		job_th.job_date.CurrentValue = ew_UnFormatDateTime(job_th.job_date.CurrentValue, 8)
		job_th.job_title.CurrentValue = job_th.job_title.FormValue
		job_th.job_intro.CurrentValue = job_th.job_intro.FormValue
		job_th.job_detail.CurrentValue = job_th.job_detail.FormValue
		job_th.job_create.CurrentValue = job_th.job_create.FormValue
		job_th.job_update.CurrentValue = job_th.job_update.FormValue
		job_th.job_show.CurrentValue = job_th.job_show.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = job_th.KeyFilter

		' Call Row Selecting event
		Call job_th.Row_Selecting(sFilter)

		' Load sql based on filter
		job_th.CurrentFilter = sFilter
		sSql = job_th.SQL
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
		Call job_th.Row_Selected(RsRow)
		job_th.job_id.DbValue = RsRow("job_id")
		job_th.company_id.DbValue = RsRow("company_id")
		job_th.job_date.DbValue = RsRow("job_date")
		job_th.job_title.DbValue = RsRow("job_title")
		job_th.job_intro.DbValue = RsRow("job_intro")
		job_th.job_detail.DbValue = RsRow("job_detail")
		job_th.job_create.DbValue = RsRow("job_create")
		job_th.job_update.DbValue = RsRow("job_update")
		job_th.job_show.DbValue = RsRow("job_show")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		job_th.job_id.m_DbValue = Rs("job_id")
		job_th.company_id.m_DbValue = Rs("company_id")
		job_th.job_date.m_DbValue = Rs("job_date")
		job_th.job_title.m_DbValue = Rs("job_title")
		job_th.job_intro.m_DbValue = Rs("job_intro")
		job_th.job_detail.m_DbValue = Rs("job_detail")
		job_th.job_create.m_DbValue = Rs("job_create")
		job_th.job_update.m_DbValue = Rs("job_update")
		job_th.job_show.m_DbValue = Rs("job_show")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call job_th.Row_Rendering()

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

		If job_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' job_id
			job_th.job_id.ViewValue = job_th.job_id.CurrentValue
			job_th.job_id.ViewCustomAttributes = ""

			' company_id
			job_th.company_id.ViewValue = job_th.company_id.CurrentValue
			job_th.company_id.ViewCustomAttributes = ""

			' job_date
			job_th.job_date.ViewValue = job_th.job_date.CurrentValue
			job_th.job_date.ViewCustomAttributes = ""

			' job_title
			job_th.job_title.ViewValue = job_th.job_title.CurrentValue
			job_th.job_title.ViewCustomAttributes = ""

			' job_intro
			job_th.job_intro.ViewValue = job_th.job_intro.CurrentValue
			job_th.job_intro.ViewCustomAttributes = ""

			' job_detail
			job_th.job_detail.ViewValue = job_th.job_detail.CurrentValue
			job_th.job_detail.ViewCustomAttributes = ""

			' job_create
			job_th.job_create.ViewValue = job_th.job_create.CurrentValue
			job_th.job_create.ViewCustomAttributes = ""

			' job_update
			job_th.job_update.ViewValue = job_th.job_update.CurrentValue
			job_th.job_update.ViewCustomAttributes = ""

			' job_show
			job_th.job_show.ViewValue = job_th.job_show.CurrentValue
			job_th.job_show.ViewCustomAttributes = ""

			' View refer script
			' job_id

			job_th.job_id.LinkCustomAttributes = ""
			job_th.job_id.HrefValue = ""
			job_th.job_id.TooltipValue = ""

			' company_id
			job_th.company_id.LinkCustomAttributes = ""
			job_th.company_id.HrefValue = ""
			job_th.company_id.TooltipValue = ""

			' job_date
			job_th.job_date.LinkCustomAttributes = ""
			job_th.job_date.HrefValue = ""
			job_th.job_date.TooltipValue = ""

			' job_title
			job_th.job_title.LinkCustomAttributes = ""
			job_th.job_title.HrefValue = ""
			job_th.job_title.TooltipValue = ""

			' job_intro
			job_th.job_intro.LinkCustomAttributes = ""
			job_th.job_intro.HrefValue = ""
			job_th.job_intro.TooltipValue = ""

			' job_detail
			job_th.job_detail.LinkCustomAttributes = ""
			job_th.job_detail.HrefValue = ""
			job_th.job_detail.TooltipValue = ""

			' job_create
			job_th.job_create.LinkCustomAttributes = ""
			job_th.job_create.HrefValue = ""
			job_th.job_create.TooltipValue = ""

			' job_update
			job_th.job_update.LinkCustomAttributes = ""
			job_th.job_update.HrefValue = ""
			job_th.job_update.TooltipValue = ""

			' job_show
			job_th.job_show.LinkCustomAttributes = ""
			job_th.job_show.HrefValue = ""
			job_th.job_show.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf job_th.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' job_id
			job_th.job_id.EditCustomAttributes = ""
			job_th.job_id.EditValue = job_th.job_id.CurrentValue
			job_th.job_id.ViewCustomAttributes = ""

			' company_id
			job_th.company_id.EditCustomAttributes = ""
			job_th.company_id.EditValue = ew_HtmlEncode(job_th.company_id.CurrentValue)
			job_th.company_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job_th.company_id.FldCaption))

			' job_date
			job_th.job_date.EditCustomAttributes = ""
			job_th.job_date.EditValue = ew_HtmlEncode(job_th.job_date.CurrentValue)
			job_th.job_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job_th.job_date.FldCaption))

			' job_title
			job_th.job_title.EditCustomAttributes = ""
			job_th.job_title.EditValue = ew_HtmlEncode(job_th.job_title.CurrentValue)
			job_th.job_title.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job_th.job_title.FldCaption))

			' job_intro
			job_th.job_intro.EditCustomAttributes = ""
			job_th.job_intro.EditValue = job_th.job_intro.CurrentValue
			job_th.job_intro.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job_th.job_intro.FldCaption))

			' job_detail
			job_th.job_detail.EditCustomAttributes = ""
			job_th.job_detail.EditValue = job_th.job_detail.CurrentValue
			job_th.job_detail.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job_th.job_detail.FldCaption))

			' job_create
			job_th.job_create.EditCustomAttributes = ""
			job_th.job_create.EditValue = ew_HtmlEncode(job_th.job_create.CurrentValue)
			job_th.job_create.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job_th.job_create.FldCaption))

			' job_update
			job_th.job_update.EditCustomAttributes = ""
			job_th.job_update.EditValue = ew_HtmlEncode(job_th.job_update.CurrentValue)
			job_th.job_update.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job_th.job_update.FldCaption))

			' job_show
			job_th.job_show.EditCustomAttributes = ""
			job_th.job_show.EditValue = ew_HtmlEncode(job_th.job_show.CurrentValue)
			job_th.job_show.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job_th.job_show.FldCaption))

			' Edit refer script
			' job_id

			job_th.job_id.HrefValue = ""

			' company_id
			job_th.company_id.HrefValue = ""

			' job_date
			job_th.job_date.HrefValue = ""

			' job_title
			job_th.job_title.HrefValue = ""

			' job_intro
			job_th.job_intro.HrefValue = ""

			' job_detail
			job_th.job_detail.HrefValue = ""

			' job_create
			job_th.job_create.HrefValue = ""

			' job_update
			job_th.job_update.HrefValue = ""

			' job_show
			job_th.job_show.HrefValue = ""
		End If
		If job_th.RowType = EW_ROWTYPE_ADD Or job_th.RowType = EW_ROWTYPE_EDIT Or job_th.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call job_th.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If job_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call job_th.Row_Rendered()
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
		If Not ew_CheckInteger(job_th.job_id.FormValue) Then
			Call ew_AddMessage(gsFormError, job_th.job_id.FldErrMsg)
		End If
		If Not ew_CheckInteger(job_th.company_id.FormValue) Then
			Call ew_AddMessage(gsFormError, job_th.company_id.FldErrMsg)
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
		sFilter = job_th.KeyFilter
		job_th.CurrentFilter  = sFilter
		sSql = job_th.SQL
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

			' Field job_id
			' Field company_id

			Call job_th.company_id.SetDbValue(Rs, job_th.company_id.CurrentValue, Null, job_th.company_id.ReadOnly)

			' Field job_date
			Call job_th.job_date.SetDbValue(Rs, job_th.job_date.CurrentValue, Null, job_th.job_date.ReadOnly)

			' Field job_title
			Call job_th.job_title.SetDbValue(Rs, job_th.job_title.CurrentValue, Null, job_th.job_title.ReadOnly)

			' Field job_intro
			Call job_th.job_intro.SetDbValue(Rs, job_th.job_intro.CurrentValue, Null, job_th.job_intro.ReadOnly)

			' Field job_detail
			Call job_th.job_detail.SetDbValue(Rs, job_th.job_detail.CurrentValue, Null, job_th.job_detail.ReadOnly)

			' Field job_create
			Call job_th.job_create.SetDbValue(Rs, job_th.job_create.CurrentValue, Null, job_th.job_create.ReadOnly)

			' Field job_update
			Call job_th.job_update.SetDbValue(Rs, job_th.job_update.CurrentValue, Null, job_th.job_update.ReadOnly)

			' Field job_show
			Call job_th.job_show.SetDbValue(Rs, job_th.job_show.CurrentValue, Null, job_th.job_show.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = job_th.Row_Updating(RsOld, Rs)
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
				ElseIf job_th.CancelMessage <> "" Then
					FailureMessage = job_th.CancelMessage
					job_th.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call job_th.Row_Updated(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", job_th.TableVar, "pom_job_thlist.asp", job_th.TableVar, True)
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
