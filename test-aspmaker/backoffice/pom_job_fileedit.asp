<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_job_fileinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim job_file_edit
Set job_file_edit = New cjob_file_edit
Set Page = job_file_edit

' Page init processing
job_file_edit.Page_Init()

' Page main processing
job_file_edit.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
job_file_edit.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var job_file_edit = new ew_Page("job_file_edit");
job_file_edit.PageID = "edit"; // Page ID
var EW_PAGE_ID = job_file_edit.PageID; // For backward compatibility
// Form object
var fjob_fileedit = new ew_Form("fjob_fileedit");
// Validate form
fjob_fileedit.Validate = function() {
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
				return this.OnError(elm, "<%= ew_JsEncode2(job_file.job_id.FldErrMsg) %>");
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
fjob_fileedit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fjob_fileedit.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fjob_fileedit.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If job_file.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% job_file_edit.ShowPageHeader() %>
<% job_file_edit.ShowMessage %>
<form name="fjob_fileedit" id="fjob_fileedit" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="a_table" id="a_table" value="job_file">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table class="ewGrid"><tr><td>
<table id="tbl_job_fileedit" class="table table-bordered table-striped">
<% If job_file.job_file_id.Visible Then ' job_file_id %>
	<tr id="r_job_file_id">
		<td><span id="elh_job_file_job_file_id"><%= job_file.job_file_id.FldCaption %></span></td>
		<td<%= job_file.job_file_id.CellAttributes %>>
<span id="el_job_file_job_file_id" class="control-group">
<span<%= job_file.job_file_id.ViewAttributes %>>
<%= job_file.job_file_id.EditValue %>
</span>
</span>
<input type="hidden" data-field="x_job_file_id" name="x_job_file_id" id="x_job_file_id" value="<%= Server.HTMLEncode(job_file.job_file_id.CurrentValue&"") %>">
<%= job_file.job_file_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If job_file.job_id.Visible Then ' job_id %>
	<tr id="r_job_id">
		<td><span id="elh_job_file_job_id"><%= job_file.job_id.FldCaption %></span></td>
		<td<%= job_file.job_id.CellAttributes %>>
<span id="el_job_file_job_id" class="control-group">
<input type="text" data-field="x_job_id" name="x_job_id" id="x_job_id" size="30" placeholder="<%= job_file.job_id.PlaceHolder %>" value="<%= job_file.job_id.EditValue %>"<%= job_file.job_id.EditAttributes %>>
</span>
<%= job_file.job_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If job_file.job_file_name.Visible Then ' job_file_name %>
	<tr id="r_job_file_name">
		<td><span id="elh_job_file_job_file_name"><%= job_file.job_file_name.FldCaption %></span></td>
		<td<%= job_file.job_file_name.CellAttributes %>>
<span id="el_job_file_job_file_name" class="control-group">
<input type="text" data-field="x_job_file_name" name="x_job_file_name" id="x_job_file_name" size="30" maxlength="255" placeholder="<%= job_file.job_file_name.PlaceHolder %>" value="<%= job_file.job_file_name.EditValue %>"<%= job_file.job_file_name.EditAttributes %>>
</span>
<%= job_file.job_file_name.CustomMsg %></td>
	</tr>
<% End If %>
<% If job_file.job_file_title.Visible Then ' job_file_title %>
	<tr id="r_job_file_title">
		<td><span id="elh_job_file_job_file_title"><%= job_file.job_file_title.FldCaption %></span></td>
		<td<%= job_file.job_file_title.CellAttributes %>>
<span id="el_job_file_job_file_title" class="control-group">
<input type="text" data-field="x_job_file_title" name="x_job_file_title" id="x_job_file_title" size="30" maxlength="255" placeholder="<%= job_file.job_file_title.PlaceHolder %>" value="<%= job_file.job_file_title.EditValue %>"<%= job_file.job_file_title.EditAttributes %>>
</span>
<%= job_file.job_file_title.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("EditBtn") %></button>
</form>
<script type="text/javascript">
fjob_fileedit.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
job_file_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set job_file_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cjob_file_edit

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
		TableName = "job_file"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "job_file_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If job_file.UseTokenInUrl Then PageUrl = PageUrl & "t=" & job_file.TableVar & "&" ' add page token
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
		If job_file.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (job_file.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (job_file.TableVar = Request.QueryString("t"))
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
		If IsEmpty(job_file) Then Set job_file = New cjob_file
		Set Table = job_file

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "job_file"

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

		job_file.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action
		job_file.job_file_id.Visible = Not job_file.IsAdd() And Not job_file.IsCopy() And Not job_file.IsGridAdd()

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
		Set job_file = Nothing
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
		If Request.QueryString("job_file_id").Count > 0 Then
			job_file.job_file_id.QueryStringValue = Request.QueryString("job_file_id")
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Process form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			job_file.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values
		Else
			job_file.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If job_file.job_file_id.CurrentValue = "" Then Call Page_Terminate("pom_job_filelist.asp") ' Invalid key, return to list

		' Validate form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			If Not ValidateForm() Then
				job_file.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				job_file.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		End If
		Select Case job_file.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_job_filelist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				job_file.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					sReturnUrl = job_file.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					job_file.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		job_file.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call job_file.ResetAttrs()
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
				job_file.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					job_file.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = job_file.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			job_file.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			job_file.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			job_file.StartRecordNumber = StartRec
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
		If Not job_file.job_file_id.FldIsDetailKey Then job_file.job_file_id.FormValue = ObjForm.GetValue("x_job_file_id")
		If Not job_file.job_id.FldIsDetailKey Then job_file.job_id.FormValue = ObjForm.GetValue("x_job_id")
		If Not job_file.job_file_name.FldIsDetailKey Then job_file.job_file_name.FormValue = ObjForm.GetValue("x_job_file_name")
		If Not job_file.job_file_title.FldIsDetailKey Then job_file.job_file_title.FormValue = ObjForm.GetValue("x_job_file_title")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		job_file.job_file_id.CurrentValue = job_file.job_file_id.FormValue
		job_file.job_id.CurrentValue = job_file.job_id.FormValue
		job_file.job_file_name.CurrentValue = job_file.job_file_name.FormValue
		job_file.job_file_title.CurrentValue = job_file.job_file_title.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = job_file.KeyFilter

		' Call Row Selecting event
		Call job_file.Row_Selecting(sFilter)

		' Load sql based on filter
		job_file.CurrentFilter = sFilter
		sSql = job_file.SQL
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
		Call job_file.Row_Selected(RsRow)
		job_file.job_file_id.DbValue = RsRow("job_file_id")
		job_file.job_id.DbValue = RsRow("job_id")
		job_file.job_file_name.DbValue = RsRow("job_file_name")
		job_file.job_file_title.DbValue = RsRow("job_file_title")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		job_file.job_file_id.m_DbValue = Rs("job_file_id")
		job_file.job_id.m_DbValue = Rs("job_id")
		job_file.job_file_name.m_DbValue = Rs("job_file_name")
		job_file.job_file_title.m_DbValue = Rs("job_file_title")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call job_file.Row_Rendering()

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

		If job_file.RowType = EW_ROWTYPE_VIEW Then ' View row

			' job_file_id
			job_file.job_file_id.ViewValue = job_file.job_file_id.CurrentValue
			job_file.job_file_id.ViewCustomAttributes = ""

			' job_id
			job_file.job_id.ViewValue = job_file.job_id.CurrentValue
			job_file.job_id.ViewCustomAttributes = ""

			' job_file_name
			job_file.job_file_name.ViewValue = job_file.job_file_name.CurrentValue
			job_file.job_file_name.ViewCustomAttributes = ""

			' job_file_title
			job_file.job_file_title.ViewValue = job_file.job_file_title.CurrentValue
			job_file.job_file_title.ViewCustomAttributes = ""

			' View refer script
			' job_file_id

			job_file.job_file_id.LinkCustomAttributes = ""
			job_file.job_file_id.HrefValue = ""
			job_file.job_file_id.TooltipValue = ""

			' job_id
			job_file.job_id.LinkCustomAttributes = ""
			job_file.job_id.HrefValue = ""
			job_file.job_id.TooltipValue = ""

			' job_file_name
			job_file.job_file_name.LinkCustomAttributes = ""
			job_file.job_file_name.HrefValue = ""
			job_file.job_file_name.TooltipValue = ""

			' job_file_title
			job_file.job_file_title.LinkCustomAttributes = ""
			job_file.job_file_title.HrefValue = ""
			job_file.job_file_title.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf job_file.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' job_file_id
			job_file.job_file_id.EditCustomAttributes = ""
			job_file.job_file_id.EditValue = job_file.job_file_id.CurrentValue
			job_file.job_file_id.ViewCustomAttributes = ""

			' job_id
			job_file.job_id.EditCustomAttributes = ""
			job_file.job_id.EditValue = ew_HtmlEncode(job_file.job_id.CurrentValue)
			job_file.job_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job_file.job_id.FldCaption))

			' job_file_name
			job_file.job_file_name.EditCustomAttributes = ""
			job_file.job_file_name.EditValue = ew_HtmlEncode(job_file.job_file_name.CurrentValue)
			job_file.job_file_name.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job_file.job_file_name.FldCaption))

			' job_file_title
			job_file.job_file_title.EditCustomAttributes = ""
			job_file.job_file_title.EditValue = ew_HtmlEncode(job_file.job_file_title.CurrentValue)
			job_file.job_file_title.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(job_file.job_file_title.FldCaption))

			' Edit refer script
			' job_file_id

			job_file.job_file_id.HrefValue = ""

			' job_id
			job_file.job_id.HrefValue = ""

			' job_file_name
			job_file.job_file_name.HrefValue = ""

			' job_file_title
			job_file.job_file_title.HrefValue = ""
		End If
		If job_file.RowType = EW_ROWTYPE_ADD Or job_file.RowType = EW_ROWTYPE_EDIT Or job_file.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call job_file.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If job_file.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call job_file.Row_Rendered()
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
		If Not ew_CheckInteger(job_file.job_id.FormValue) Then
			Call ew_AddMessage(gsFormError, job_file.job_id.FldErrMsg)
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
		sFilter = job_file.KeyFilter
		job_file.CurrentFilter  = sFilter
		sSql = job_file.SQL
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
			Call job_file.job_id.SetDbValue(Rs, job_file.job_id.CurrentValue, Null, job_file.job_id.ReadOnly)

			' Field job_file_name
			Call job_file.job_file_name.SetDbValue(Rs, job_file.job_file_name.CurrentValue, Null, job_file.job_file_name.ReadOnly)

			' Field job_file_title
			Call job_file.job_file_title.SetDbValue(Rs, job_file.job_file_title.CurrentValue, Null, job_file.job_file_title.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = job_file.Row_Updating(RsOld, Rs)
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
				ElseIf job_file.CancelMessage <> "" Then
					FailureMessage = job_file.CancelMessage
					job_file.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call job_file.Row_Updated(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", job_file.TableVar, "pom_job_filelist.asp", job_file.TableVar, True)
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
