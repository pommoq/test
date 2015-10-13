<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_eventcalendar_pdf_file_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim eventcalendar_pdf_file_th_add
Set eventcalendar_pdf_file_th_add = New ceventcalendar_pdf_file_th_add
Set Page = eventcalendar_pdf_file_th_add

' Page init processing
eventcalendar_pdf_file_th_add.Page_Init()

' Page main processing
eventcalendar_pdf_file_th_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
eventcalendar_pdf_file_th_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var eventcalendar_pdf_file_th_add = new ew_Page("eventcalendar_pdf_file_th_add");
eventcalendar_pdf_file_th_add.PageID = "add"; // Page ID
var EW_PAGE_ID = eventcalendar_pdf_file_th_add.PageID; // For backward compatibility
// Form object
var feventcalendar_pdf_file_thadd = new ew_Form("feventcalendar_pdf_file_thadd");
// Validate form
feventcalendar_pdf_file_thadd.Validate = function() {
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
			elm = this.GetElements("x" + infix + "_eventcalendar_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(eventcalendar_pdf_file_th.eventcalendar_id.FldErrMsg) %>");
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
feventcalendar_pdf_file_thadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
feventcalendar_pdf_file_thadd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
feventcalendar_pdf_file_thadd.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If eventcalendar_pdf_file_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% eventcalendar_pdf_file_th_add.ShowPageHeader() %>
<% eventcalendar_pdf_file_th_add.ShowMessage %>
<form name="feventcalendar_pdf_file_thadd" id="feventcalendar_pdf_file_thadd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="eventcalendar_pdf_file_th">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_eventcalendar_pdf_file_thadd" class="table table-bordered table-striped">
<% If eventcalendar_pdf_file_th.eventcalendar_id.Visible Then ' eventcalendar_id %>
	<tr id="r_eventcalendar_id">
		<td><span id="elh_eventcalendar_pdf_file_th_eventcalendar_id"><%= eventcalendar_pdf_file_th.eventcalendar_id.FldCaption %></span></td>
		<td<%= eventcalendar_pdf_file_th.eventcalendar_id.CellAttributes %>>
<span id="el_eventcalendar_pdf_file_th_eventcalendar_id" class="control-group">
<input type="text" data-field="x_eventcalendar_id" name="x_eventcalendar_id" id="x_eventcalendar_id" size="30" placeholder="<%= eventcalendar_pdf_file_th.eventcalendar_id.PlaceHolder %>" value="<%= eventcalendar_pdf_file_th.eventcalendar_id.EditValue %>"<%= eventcalendar_pdf_file_th.eventcalendar_id.EditAttributes %>>
</span>
<%= eventcalendar_pdf_file_th.eventcalendar_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_pdf_file_th.eventcalendar_pdf_file.Visible Then ' eventcalendar_pdf_file %>
	<tr id="r_eventcalendar_pdf_file">
		<td><span id="elh_eventcalendar_pdf_file_th_eventcalendar_pdf_file"><%= eventcalendar_pdf_file_th.eventcalendar_pdf_file.FldCaption %></span></td>
		<td<%= eventcalendar_pdf_file_th.eventcalendar_pdf_file.CellAttributes %>>
<span id="el_eventcalendar_pdf_file_th_eventcalendar_pdf_file" class="control-group">
<input type="text" data-field="x_eventcalendar_pdf_file" name="x_eventcalendar_pdf_file" id="x_eventcalendar_pdf_file" size="30" maxlength="255" placeholder="<%= eventcalendar_pdf_file_th.eventcalendar_pdf_file.PlaceHolder %>" value="<%= eventcalendar_pdf_file_th.eventcalendar_pdf_file.EditValue %>"<%= eventcalendar_pdf_file_th.eventcalendar_pdf_file.EditAttributes %>>
</span>
<%= eventcalendar_pdf_file_th.eventcalendar_pdf_file.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_pdf_file_th.eventcalendar_pdf_title.Visible Then ' eventcalendar_pdf_title %>
	<tr id="r_eventcalendar_pdf_title">
		<td><span id="elh_eventcalendar_pdf_file_th_eventcalendar_pdf_title"><%= eventcalendar_pdf_file_th.eventcalendar_pdf_title.FldCaption %></span></td>
		<td<%= eventcalendar_pdf_file_th.eventcalendar_pdf_title.CellAttributes %>>
<span id="el_eventcalendar_pdf_file_th_eventcalendar_pdf_title" class="control-group">
<input type="text" data-field="x_eventcalendar_pdf_title" name="x_eventcalendar_pdf_title" id="x_eventcalendar_pdf_title" size="30" maxlength="255" placeholder="<%= eventcalendar_pdf_file_th.eventcalendar_pdf_title.PlaceHolder %>" value="<%= eventcalendar_pdf_file_th.eventcalendar_pdf_title.EditValue %>"<%= eventcalendar_pdf_file_th.eventcalendar_pdf_title.EditAttributes %>>
</span>
<%= eventcalendar_pdf_file_th.eventcalendar_pdf_title.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
feventcalendar_pdf_file_thadd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
eventcalendar_pdf_file_th_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set eventcalendar_pdf_file_th_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class ceventcalendar_pdf_file_th_add

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
		TableName = "eventcalendar_pdf_file_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "eventcalendar_pdf_file_th_add"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If eventcalendar_pdf_file_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & eventcalendar_pdf_file_th.TableVar & "&" ' add page token
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
		If eventcalendar_pdf_file_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (eventcalendar_pdf_file_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (eventcalendar_pdf_file_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(eventcalendar_pdf_file_th) Then Set eventcalendar_pdf_file_th = New ceventcalendar_pdf_file_th
		Set Table = eventcalendar_pdf_file_th

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "add"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "eventcalendar_pdf_file_th"

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

		eventcalendar_pdf_file_th.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set eventcalendar_pdf_file_th = Nothing
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
			eventcalendar_pdf_file_th.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("eventcalendar_pdf_id").Count > 0 Then
				eventcalendar_pdf_file_th.eventcalendar_pdf_id.QueryStringValue = Request.QueryString("eventcalendar_pdf_id")
				Call eventcalendar_pdf_file_th.SetKey("eventcalendar_pdf_id", eventcalendar_pdf_file_th.eventcalendar_pdf_id.CurrentValue) ' Set up key
			Else
				Call eventcalendar_pdf_file_th.SetKey("eventcalendar_pdf_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				eventcalendar_pdf_file_th.CurrentAction = "C" ' Copy Record
			Else
				eventcalendar_pdf_file_th.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				eventcalendar_pdf_file_th.CurrentAction = "I" ' Form error, reset action
				eventcalendar_pdf_file_th.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case eventcalendar_pdf_file_th.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_eventcalendar_pdf_file_thlist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				eventcalendar_pdf_file_th.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = eventcalendar_pdf_file_th.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_eventcalendar_pdf_file_thview.asp" Then sReturnUrl = eventcalendar_pdf_file_th.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					eventcalendar_pdf_file_th.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		eventcalendar_pdf_file_th.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call eventcalendar_pdf_file_th.ResetAttrs()
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
		eventcalendar_pdf_file_th.eventcalendar_id.CurrentValue = Null
		eventcalendar_pdf_file_th.eventcalendar_id.OldValue = eventcalendar_pdf_file_th.eventcalendar_id.CurrentValue
		eventcalendar_pdf_file_th.eventcalendar_pdf_file.CurrentValue = Null
		eventcalendar_pdf_file_th.eventcalendar_pdf_file.OldValue = eventcalendar_pdf_file_th.eventcalendar_pdf_file.CurrentValue
		eventcalendar_pdf_file_th.eventcalendar_pdf_title.CurrentValue = Null
		eventcalendar_pdf_file_th.eventcalendar_pdf_title.OldValue = eventcalendar_pdf_file_th.eventcalendar_pdf_title.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not eventcalendar_pdf_file_th.eventcalendar_id.FldIsDetailKey Then eventcalendar_pdf_file_th.eventcalendar_id.FormValue = ObjForm.GetValue("x_eventcalendar_id")
		If Not eventcalendar_pdf_file_th.eventcalendar_pdf_file.FldIsDetailKey Then eventcalendar_pdf_file_th.eventcalendar_pdf_file.FormValue = ObjForm.GetValue("x_eventcalendar_pdf_file")
		If Not eventcalendar_pdf_file_th.eventcalendar_pdf_title.FldIsDetailKey Then eventcalendar_pdf_file_th.eventcalendar_pdf_title.FormValue = ObjForm.GetValue("x_eventcalendar_pdf_title")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		eventcalendar_pdf_file_th.eventcalendar_id.CurrentValue = eventcalendar_pdf_file_th.eventcalendar_id.FormValue
		eventcalendar_pdf_file_th.eventcalendar_pdf_file.CurrentValue = eventcalendar_pdf_file_th.eventcalendar_pdf_file.FormValue
		eventcalendar_pdf_file_th.eventcalendar_pdf_title.CurrentValue = eventcalendar_pdf_file_th.eventcalendar_pdf_title.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = eventcalendar_pdf_file_th.KeyFilter

		' Call Row Selecting event
		Call eventcalendar_pdf_file_th.Row_Selecting(sFilter)

		' Load sql based on filter
		eventcalendar_pdf_file_th.CurrentFilter = sFilter
		sSql = eventcalendar_pdf_file_th.SQL
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
		Call eventcalendar_pdf_file_th.Row_Selected(RsRow)
		eventcalendar_pdf_file_th.eventcalendar_pdf_id.DbValue = RsRow("eventcalendar_pdf_id")
		eventcalendar_pdf_file_th.eventcalendar_id.DbValue = RsRow("eventcalendar_id")
		eventcalendar_pdf_file_th.eventcalendar_pdf_file.DbValue = RsRow("eventcalendar_pdf_file")
		eventcalendar_pdf_file_th.eventcalendar_pdf_title.DbValue = RsRow("eventcalendar_pdf_title")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		eventcalendar_pdf_file_th.eventcalendar_pdf_id.m_DbValue = Rs("eventcalendar_pdf_id")
		eventcalendar_pdf_file_th.eventcalendar_id.m_DbValue = Rs("eventcalendar_id")
		eventcalendar_pdf_file_th.eventcalendar_pdf_file.m_DbValue = Rs("eventcalendar_pdf_file")
		eventcalendar_pdf_file_th.eventcalendar_pdf_title.m_DbValue = Rs("eventcalendar_pdf_title")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If eventcalendar_pdf_file_th.GetKey("eventcalendar_pdf_id")&"" <> "" Then
			eventcalendar_pdf_file_th.eventcalendar_pdf_id.CurrentValue = eventcalendar_pdf_file_th.GetKey("eventcalendar_pdf_id") ' eventcalendar_pdf_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			eventcalendar_pdf_file_th.CurrentFilter = eventcalendar_pdf_file_th.KeyFilter
			Dim sSql
			sSql = eventcalendar_pdf_file_th.SQL
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

		Call eventcalendar_pdf_file_th.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' eventcalendar_pdf_id
		' eventcalendar_id
		' eventcalendar_pdf_file
		' eventcalendar_pdf_title
		' -----------
		'  View  Row
		' -----------

		If eventcalendar_pdf_file_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' eventcalendar_pdf_id
			eventcalendar_pdf_file_th.eventcalendar_pdf_id.ViewValue = eventcalendar_pdf_file_th.eventcalendar_pdf_id.CurrentValue
			eventcalendar_pdf_file_th.eventcalendar_pdf_id.ViewCustomAttributes = ""

			' eventcalendar_id
			eventcalendar_pdf_file_th.eventcalendar_id.ViewValue = eventcalendar_pdf_file_th.eventcalendar_id.CurrentValue
			eventcalendar_pdf_file_th.eventcalendar_id.ViewCustomAttributes = ""

			' eventcalendar_pdf_file
			eventcalendar_pdf_file_th.eventcalendar_pdf_file.ViewValue = eventcalendar_pdf_file_th.eventcalendar_pdf_file.CurrentValue
			eventcalendar_pdf_file_th.eventcalendar_pdf_file.ViewCustomAttributes = ""

			' eventcalendar_pdf_title
			eventcalendar_pdf_file_th.eventcalendar_pdf_title.ViewValue = eventcalendar_pdf_file_th.eventcalendar_pdf_title.CurrentValue
			eventcalendar_pdf_file_th.eventcalendar_pdf_title.ViewCustomAttributes = ""

			' View refer script
			' eventcalendar_id

			eventcalendar_pdf_file_th.eventcalendar_id.LinkCustomAttributes = ""
			eventcalendar_pdf_file_th.eventcalendar_id.HrefValue = ""
			eventcalendar_pdf_file_th.eventcalendar_id.TooltipValue = ""

			' eventcalendar_pdf_file
			eventcalendar_pdf_file_th.eventcalendar_pdf_file.LinkCustomAttributes = ""
			eventcalendar_pdf_file_th.eventcalendar_pdf_file.HrefValue = ""
			eventcalendar_pdf_file_th.eventcalendar_pdf_file.TooltipValue = ""

			' eventcalendar_pdf_title
			eventcalendar_pdf_file_th.eventcalendar_pdf_title.LinkCustomAttributes = ""
			eventcalendar_pdf_file_th.eventcalendar_pdf_title.HrefValue = ""
			eventcalendar_pdf_file_th.eventcalendar_pdf_title.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf eventcalendar_pdf_file_th.RowType = EW_ROWTYPE_ADD Then ' Add row

			' eventcalendar_id
			eventcalendar_pdf_file_th.eventcalendar_id.EditCustomAttributes = ""
			eventcalendar_pdf_file_th.eventcalendar_id.EditValue = ew_HtmlEncode(eventcalendar_pdf_file_th.eventcalendar_id.CurrentValue)
			eventcalendar_pdf_file_th.eventcalendar_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_pdf_file_th.eventcalendar_id.FldCaption))

			' eventcalendar_pdf_file
			eventcalendar_pdf_file_th.eventcalendar_pdf_file.EditCustomAttributes = ""
			eventcalendar_pdf_file_th.eventcalendar_pdf_file.EditValue = ew_HtmlEncode(eventcalendar_pdf_file_th.eventcalendar_pdf_file.CurrentValue)
			eventcalendar_pdf_file_th.eventcalendar_pdf_file.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_pdf_file_th.eventcalendar_pdf_file.FldCaption))

			' eventcalendar_pdf_title
			eventcalendar_pdf_file_th.eventcalendar_pdf_title.EditCustomAttributes = ""
			eventcalendar_pdf_file_th.eventcalendar_pdf_title.EditValue = ew_HtmlEncode(eventcalendar_pdf_file_th.eventcalendar_pdf_title.CurrentValue)
			eventcalendar_pdf_file_th.eventcalendar_pdf_title.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_pdf_file_th.eventcalendar_pdf_title.FldCaption))

			' Edit refer script
			' eventcalendar_id

			eventcalendar_pdf_file_th.eventcalendar_id.HrefValue = ""

			' eventcalendar_pdf_file
			eventcalendar_pdf_file_th.eventcalendar_pdf_file.HrefValue = ""

			' eventcalendar_pdf_title
			eventcalendar_pdf_file_th.eventcalendar_pdf_title.HrefValue = ""
		End If
		If eventcalendar_pdf_file_th.RowType = EW_ROWTYPE_ADD Or eventcalendar_pdf_file_th.RowType = EW_ROWTYPE_EDIT Or eventcalendar_pdf_file_th.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call eventcalendar_pdf_file_th.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If eventcalendar_pdf_file_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call eventcalendar_pdf_file_th.Row_Rendered()
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
		If Not ew_CheckInteger(eventcalendar_pdf_file_th.eventcalendar_id.FormValue) Then
			Call ew_AddMessage(gsFormError, eventcalendar_pdf_file_th.eventcalendar_id.FldErrMsg)
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
		eventcalendar_pdf_file_th.CurrentFilter = sFilter
		sSql = eventcalendar_pdf_file_th.SQL
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

		' Field eventcalendar_id
		Call eventcalendar_pdf_file_th.eventcalendar_id.SetDbValue(Rs, eventcalendar_pdf_file_th.eventcalendar_id.CurrentValue, Null, False)

		' Field eventcalendar_pdf_file
		Call eventcalendar_pdf_file_th.eventcalendar_pdf_file.SetDbValue(Rs, eventcalendar_pdf_file_th.eventcalendar_pdf_file.CurrentValue, Null, False)

		' Field eventcalendar_pdf_title
		Call eventcalendar_pdf_file_th.eventcalendar_pdf_title.SetDbValue(Rs, eventcalendar_pdf_file_th.eventcalendar_pdf_title.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = eventcalendar_pdf_file_th.Row_Inserting(RsOld, Rs)
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
			ElseIf eventcalendar_pdf_file_th.CancelMessage <> "" Then
				FailureMessage = eventcalendar_pdf_file_th.CancelMessage
				eventcalendar_pdf_file_th.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("InsertCancelled")
			End If
			AddRow = False
		End If
		Rs.Close
		Set Rs = Nothing
		If AddRow Then
			eventcalendar_pdf_file_th.eventcalendar_pdf_id.DbValue = RsNew("eventcalendar_pdf_id")
		End If
		If AddRow Then

			' Call Row Inserted event
			Call eventcalendar_pdf_file_th.Row_Inserted(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", eventcalendar_pdf_file_th.TableVar, "pom_eventcalendar_pdf_file_thlist.asp", eventcalendar_pdf_file_th.TableVar, True)
		PageId = ew_IIf(eventcalendar_pdf_file_th.CurrentAction = "C", "Copy", "Add")
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
