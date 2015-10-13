<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_journalinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim journal_add
Set journal_add = New cjournal_add
Set Page = journal_add

' Page init processing
journal_add.Page_Init()

' Page main processing
journal_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
journal_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var journal_add = new ew_Page("journal_add");
journal_add.PageID = "add"; // Page ID
var EW_PAGE_ID = journal_add.PageID; // For backward compatibility
// Form object
var fjournaladd = new ew_Form("fjournaladd");
// Validate form
fjournaladd.Validate = function() {
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
fjournaladd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fjournaladd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fjournaladd.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If journal.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% journal_add.ShowPageHeader() %>
<% journal_add.ShowMessage %>
<form name="fjournaladd" id="fjournaladd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="journal">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_journaladd" class="table table-bordered table-striped">
<% If journal.jrl_category.Visible Then ' jrl_category %>
	<tr id="r_jrl_category">
		<td><span id="elh_journal_jrl_category"><%= journal.jrl_category.FldCaption %></span></td>
		<td<%= journal.jrl_category.CellAttributes %>>
<span id="el_journal_jrl_category" class="control-group">
<input type="text" data-field="x_jrl_category" name="x_jrl_category" id="x_jrl_category" size="30" maxlength="255" placeholder="<%= journal.jrl_category.PlaceHolder %>" value="<%= journal.jrl_category.EditValue %>"<%= journal.jrl_category.EditAttributes %>>
</span>
<%= journal.jrl_category.CustomMsg %></td>
	</tr>
<% End If %>
<% If journal.jrl_date.Visible Then ' jrl_date %>
	<tr id="r_jrl_date">
		<td><span id="elh_journal_jrl_date"><%= journal.jrl_date.FldCaption %></span></td>
		<td<%= journal.jrl_date.CellAttributes %>>
<span id="el_journal_jrl_date" class="control-group">
<input type="text" data-field="x_jrl_date" name="x_jrl_date" id="x_jrl_date" placeholder="<%= journal.jrl_date.PlaceHolder %>" value="<%= journal.jrl_date.EditValue %>"<%= journal.jrl_date.EditAttributes %>>
</span>
<%= journal.jrl_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If journal.jrl_title.Visible Then ' jrl_title %>
	<tr id="r_jrl_title">
		<td><span id="elh_journal_jrl_title"><%= journal.jrl_title.FldCaption %></span></td>
		<td<%= journal.jrl_title.CellAttributes %>>
<span id="el_journal_jrl_title" class="control-group">
<input type="text" data-field="x_jrl_title" name="x_jrl_title" id="x_jrl_title" size="30" maxlength="255" placeholder="<%= journal.jrl_title.PlaceHolder %>" value="<%= journal.jrl_title.EditValue %>"<%= journal.jrl_title.EditAttributes %>>
</span>
<%= journal.jrl_title.CustomMsg %></td>
	</tr>
<% End If %>
<% If journal.jrl_title_th.Visible Then ' jrl_title_th %>
	<tr id="r_jrl_title_th">
		<td><span id="elh_journal_jrl_title_th"><%= journal.jrl_title_th.FldCaption %></span></td>
		<td<%= journal.jrl_title_th.CellAttributes %>>
<span id="el_journal_jrl_title_th" class="control-group">
<input type="text" data-field="x_jrl_title_th" name="x_jrl_title_th" id="x_jrl_title_th" size="30" maxlength="255" placeholder="<%= journal.jrl_title_th.PlaceHolder %>" value="<%= journal.jrl_title_th.EditValue %>"<%= journal.jrl_title_th.EditAttributes %>>
</span>
<%= journal.jrl_title_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If journal.jrl_pdf.Visible Then ' jrl_pdf %>
	<tr id="r_jrl_pdf">
		<td><span id="elh_journal_jrl_pdf"><%= journal.jrl_pdf.FldCaption %></span></td>
		<td<%= journal.jrl_pdf.CellAttributes %>>
<span id="el_journal_jrl_pdf" class="control-group">
<input type="text" data-field="x_jrl_pdf" name="x_jrl_pdf" id="x_jrl_pdf" size="30" maxlength="255" placeholder="<%= journal.jrl_pdf.PlaceHolder %>" value="<%= journal.jrl_pdf.EditValue %>"<%= journal.jrl_pdf.EditAttributes %>>
</span>
<%= journal.jrl_pdf.CustomMsg %></td>
	</tr>
<% End If %>
<% If journal.jrl_img.Visible Then ' jrl_img %>
	<tr id="r_jrl_img">
		<td><span id="elh_journal_jrl_img"><%= journal.jrl_img.FldCaption %></span></td>
		<td<%= journal.jrl_img.CellAttributes %>>
<span id="el_journal_jrl_img" class="control-group">
<input type="text" data-field="x_jrl_img" name="x_jrl_img" id="x_jrl_img" size="30" maxlength="255" placeholder="<%= journal.jrl_img.PlaceHolder %>" value="<%= journal.jrl_img.EditValue %>"<%= journal.jrl_img.EditAttributes %>>
</span>
<%= journal.jrl_img.CustomMsg %></td>
	</tr>
<% End If %>
<% If journal.jrl_create.Visible Then ' jrl_create %>
	<tr id="r_jrl_create">
		<td><span id="elh_journal_jrl_create"><%= journal.jrl_create.FldCaption %></span></td>
		<td<%= journal.jrl_create.CellAttributes %>>
<span id="el_journal_jrl_create" class="control-group">
<input type="text" data-field="x_jrl_create" name="x_jrl_create" id="x_jrl_create" size="30" maxlength="255" placeholder="<%= journal.jrl_create.PlaceHolder %>" value="<%= journal.jrl_create.EditValue %>"<%= journal.jrl_create.EditAttributes %>>
</span>
<%= journal.jrl_create.CustomMsg %></td>
	</tr>
<% End If %>
<% If journal.jrl_update.Visible Then ' jrl_update %>
	<tr id="r_jrl_update">
		<td><span id="elh_journal_jrl_update"><%= journal.jrl_update.FldCaption %></span></td>
		<td<%= journal.jrl_update.CellAttributes %>>
<span id="el_journal_jrl_update" class="control-group">
<input type="text" data-field="x_jrl_update" name="x_jrl_update" id="x_jrl_update" size="30" maxlength="255" placeholder="<%= journal.jrl_update.PlaceHolder %>" value="<%= journal.jrl_update.EditValue %>"<%= journal.jrl_update.EditAttributes %>>
</span>
<%= journal.jrl_update.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
fjournaladd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
journal_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set journal_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cjournal_add

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
		TableName = "journal"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "journal_add"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If journal.UseTokenInUrl Then PageUrl = PageUrl & "t=" & journal.TableVar & "&" ' add page token
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
		If journal.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (journal.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (journal.TableVar = Request.QueryString("t"))
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
		If IsEmpty(journal) Then Set journal = New cjournal
		Set Table = journal

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "add"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "journal"

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

		journal.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set journal = Nothing
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
			journal.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("jrl_id").Count > 0 Then
				journal.jrl_id.QueryStringValue = Request.QueryString("jrl_id")
				Call journal.SetKey("jrl_id", journal.jrl_id.CurrentValue) ' Set up key
			Else
				Call journal.SetKey("jrl_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				journal.CurrentAction = "C" ' Copy Record
			Else
				journal.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				journal.CurrentAction = "I" ' Form error, reset action
				journal.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case journal.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_journallist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				journal.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = journal.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_journalview.asp" Then sReturnUrl = journal.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					journal.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		journal.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call journal.ResetAttrs()
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
		journal.jrl_category.CurrentValue = Null
		journal.jrl_category.OldValue = journal.jrl_category.CurrentValue
		journal.jrl_date.CurrentValue = Null
		journal.jrl_date.OldValue = journal.jrl_date.CurrentValue
		journal.jrl_title.CurrentValue = Null
		journal.jrl_title.OldValue = journal.jrl_title.CurrentValue
		journal.jrl_title_th.CurrentValue = Null
		journal.jrl_title_th.OldValue = journal.jrl_title_th.CurrentValue
		journal.jrl_pdf.CurrentValue = Null
		journal.jrl_pdf.OldValue = journal.jrl_pdf.CurrentValue
		journal.jrl_img.CurrentValue = Null
		journal.jrl_img.OldValue = journal.jrl_img.CurrentValue
		journal.jrl_create.CurrentValue = Null
		journal.jrl_create.OldValue = journal.jrl_create.CurrentValue
		journal.jrl_update.CurrentValue = Null
		journal.jrl_update.OldValue = journal.jrl_update.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not journal.jrl_category.FldIsDetailKey Then journal.jrl_category.FormValue = ObjForm.GetValue("x_jrl_category")
		If Not journal.jrl_date.FldIsDetailKey Then journal.jrl_date.FormValue = ObjForm.GetValue("x_jrl_date")
		If Not journal.jrl_date.FldIsDetailKey Then journal.jrl_date.CurrentValue = ew_UnFormatDateTime(journal.jrl_date.CurrentValue, 8)
		If Not journal.jrl_title.FldIsDetailKey Then journal.jrl_title.FormValue = ObjForm.GetValue("x_jrl_title")
		If Not journal.jrl_title_th.FldIsDetailKey Then journal.jrl_title_th.FormValue = ObjForm.GetValue("x_jrl_title_th")
		If Not journal.jrl_pdf.FldIsDetailKey Then journal.jrl_pdf.FormValue = ObjForm.GetValue("x_jrl_pdf")
		If Not journal.jrl_img.FldIsDetailKey Then journal.jrl_img.FormValue = ObjForm.GetValue("x_jrl_img")
		If Not journal.jrl_create.FldIsDetailKey Then journal.jrl_create.FormValue = ObjForm.GetValue("x_jrl_create")
		If Not journal.jrl_update.FldIsDetailKey Then journal.jrl_update.FormValue = ObjForm.GetValue("x_jrl_update")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		journal.jrl_category.CurrentValue = journal.jrl_category.FormValue
		journal.jrl_date.CurrentValue = journal.jrl_date.FormValue
		journal.jrl_date.CurrentValue = ew_UnFormatDateTime(journal.jrl_date.CurrentValue, 8)
		journal.jrl_title.CurrentValue = journal.jrl_title.FormValue
		journal.jrl_title_th.CurrentValue = journal.jrl_title_th.FormValue
		journal.jrl_pdf.CurrentValue = journal.jrl_pdf.FormValue
		journal.jrl_img.CurrentValue = journal.jrl_img.FormValue
		journal.jrl_create.CurrentValue = journal.jrl_create.FormValue
		journal.jrl_update.CurrentValue = journal.jrl_update.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = journal.KeyFilter

		' Call Row Selecting event
		Call journal.Row_Selecting(sFilter)

		' Load sql based on filter
		journal.CurrentFilter = sFilter
		sSql = journal.SQL
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
		Call journal.Row_Selected(RsRow)
		journal.jrl_id.DbValue = RsRow("jrl_id")
		journal.jrl_category.DbValue = RsRow("jrl_category")
		journal.jrl_date.DbValue = RsRow("jrl_date")
		journal.jrl_title.DbValue = RsRow("jrl_title")
		journal.jrl_title_th.DbValue = RsRow("jrl_title_th")
		journal.jrl_pdf.DbValue = RsRow("jrl_pdf")
		journal.jrl_img.DbValue = RsRow("jrl_img")
		journal.jrl_create.DbValue = RsRow("jrl_create")
		journal.jrl_update.DbValue = RsRow("jrl_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		journal.jrl_id.m_DbValue = Rs("jrl_id")
		journal.jrl_category.m_DbValue = Rs("jrl_category")
		journal.jrl_date.m_DbValue = Rs("jrl_date")
		journal.jrl_title.m_DbValue = Rs("jrl_title")
		journal.jrl_title_th.m_DbValue = Rs("jrl_title_th")
		journal.jrl_pdf.m_DbValue = Rs("jrl_pdf")
		journal.jrl_img.m_DbValue = Rs("jrl_img")
		journal.jrl_create.m_DbValue = Rs("jrl_create")
		journal.jrl_update.m_DbValue = Rs("jrl_update")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If journal.GetKey("jrl_id")&"" <> "" Then
			journal.jrl_id.CurrentValue = journal.GetKey("jrl_id") ' jrl_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			journal.CurrentFilter = journal.KeyFilter
			Dim sSql
			sSql = journal.SQL
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

		Call journal.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' jrl_id
		' jrl_category
		' jrl_date
		' jrl_title
		' jrl_title_th
		' jrl_pdf
		' jrl_img
		' jrl_create
		' jrl_update
		' -----------
		'  View  Row
		' -----------

		If journal.RowType = EW_ROWTYPE_VIEW Then ' View row

			' jrl_id
			journal.jrl_id.ViewValue = journal.jrl_id.CurrentValue
			journal.jrl_id.ViewCustomAttributes = ""

			' jrl_category
			journal.jrl_category.ViewValue = journal.jrl_category.CurrentValue
			journal.jrl_category.ViewCustomAttributes = ""

			' jrl_date
			journal.jrl_date.ViewValue = journal.jrl_date.CurrentValue
			journal.jrl_date.ViewCustomAttributes = ""

			' jrl_title
			journal.jrl_title.ViewValue = journal.jrl_title.CurrentValue
			journal.jrl_title.ViewCustomAttributes = ""

			' jrl_title_th
			journal.jrl_title_th.ViewValue = journal.jrl_title_th.CurrentValue
			journal.jrl_title_th.ViewCustomAttributes = ""

			' jrl_pdf
			journal.jrl_pdf.ViewValue = journal.jrl_pdf.CurrentValue
			journal.jrl_pdf.ViewCustomAttributes = ""

			' jrl_img
			journal.jrl_img.ViewValue = journal.jrl_img.CurrentValue
			journal.jrl_img.ViewCustomAttributes = ""

			' jrl_create
			journal.jrl_create.ViewValue = journal.jrl_create.CurrentValue
			journal.jrl_create.ViewCustomAttributes = ""

			' jrl_update
			journal.jrl_update.ViewValue = journal.jrl_update.CurrentValue
			journal.jrl_update.ViewCustomAttributes = ""

			' View refer script
			' jrl_category

			journal.jrl_category.LinkCustomAttributes = ""
			journal.jrl_category.HrefValue = ""
			journal.jrl_category.TooltipValue = ""

			' jrl_date
			journal.jrl_date.LinkCustomAttributes = ""
			journal.jrl_date.HrefValue = ""
			journal.jrl_date.TooltipValue = ""

			' jrl_title
			journal.jrl_title.LinkCustomAttributes = ""
			journal.jrl_title.HrefValue = ""
			journal.jrl_title.TooltipValue = ""

			' jrl_title_th
			journal.jrl_title_th.LinkCustomAttributes = ""
			journal.jrl_title_th.HrefValue = ""
			journal.jrl_title_th.TooltipValue = ""

			' jrl_pdf
			journal.jrl_pdf.LinkCustomAttributes = ""
			journal.jrl_pdf.HrefValue = ""
			journal.jrl_pdf.TooltipValue = ""

			' jrl_img
			journal.jrl_img.LinkCustomAttributes = ""
			journal.jrl_img.HrefValue = ""
			journal.jrl_img.TooltipValue = ""

			' jrl_create
			journal.jrl_create.LinkCustomAttributes = ""
			journal.jrl_create.HrefValue = ""
			journal.jrl_create.TooltipValue = ""

			' jrl_update
			journal.jrl_update.LinkCustomAttributes = ""
			journal.jrl_update.HrefValue = ""
			journal.jrl_update.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf journal.RowType = EW_ROWTYPE_ADD Then ' Add row

			' jrl_category
			journal.jrl_category.EditCustomAttributes = ""
			journal.jrl_category.EditValue = ew_HtmlEncode(journal.jrl_category.CurrentValue)
			journal.jrl_category.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(journal.jrl_category.FldCaption))

			' jrl_date
			journal.jrl_date.EditCustomAttributes = ""
			journal.jrl_date.EditValue = ew_HtmlEncode(journal.jrl_date.CurrentValue)
			journal.jrl_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(journal.jrl_date.FldCaption))

			' jrl_title
			journal.jrl_title.EditCustomAttributes = ""
			journal.jrl_title.EditValue = ew_HtmlEncode(journal.jrl_title.CurrentValue)
			journal.jrl_title.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(journal.jrl_title.FldCaption))

			' jrl_title_th
			journal.jrl_title_th.EditCustomAttributes = ""
			journal.jrl_title_th.EditValue = ew_HtmlEncode(journal.jrl_title_th.CurrentValue)
			journal.jrl_title_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(journal.jrl_title_th.FldCaption))

			' jrl_pdf
			journal.jrl_pdf.EditCustomAttributes = ""
			journal.jrl_pdf.EditValue = ew_HtmlEncode(journal.jrl_pdf.CurrentValue)
			journal.jrl_pdf.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(journal.jrl_pdf.FldCaption))

			' jrl_img
			journal.jrl_img.EditCustomAttributes = ""
			journal.jrl_img.EditValue = ew_HtmlEncode(journal.jrl_img.CurrentValue)
			journal.jrl_img.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(journal.jrl_img.FldCaption))

			' jrl_create
			journal.jrl_create.EditCustomAttributes = ""
			journal.jrl_create.EditValue = ew_HtmlEncode(journal.jrl_create.CurrentValue)
			journal.jrl_create.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(journal.jrl_create.FldCaption))

			' jrl_update
			journal.jrl_update.EditCustomAttributes = ""
			journal.jrl_update.EditValue = ew_HtmlEncode(journal.jrl_update.CurrentValue)
			journal.jrl_update.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(journal.jrl_update.FldCaption))

			' Edit refer script
			' jrl_category

			journal.jrl_category.HrefValue = ""

			' jrl_date
			journal.jrl_date.HrefValue = ""

			' jrl_title
			journal.jrl_title.HrefValue = ""

			' jrl_title_th
			journal.jrl_title_th.HrefValue = ""

			' jrl_pdf
			journal.jrl_pdf.HrefValue = ""

			' jrl_img
			journal.jrl_img.HrefValue = ""

			' jrl_create
			journal.jrl_create.HrefValue = ""

			' jrl_update
			journal.jrl_update.HrefValue = ""
		End If
		If journal.RowType = EW_ROWTYPE_ADD Or journal.RowType = EW_ROWTYPE_EDIT Or journal.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call journal.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If journal.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call journal.Row_Rendered()
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
		journal.CurrentFilter = sFilter
		sSql = journal.SQL
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

		' Field jrl_category
		Call journal.jrl_category.SetDbValue(Rs, journal.jrl_category.CurrentValue, Null, False)

		' Field jrl_date
		Call journal.jrl_date.SetDbValue(Rs, journal.jrl_date.CurrentValue, Null, False)

		' Field jrl_title
		Call journal.jrl_title.SetDbValue(Rs, journal.jrl_title.CurrentValue, Null, False)

		' Field jrl_title_th
		Call journal.jrl_title_th.SetDbValue(Rs, journal.jrl_title_th.CurrentValue, Null, False)

		' Field jrl_pdf
		Call journal.jrl_pdf.SetDbValue(Rs, journal.jrl_pdf.CurrentValue, Null, False)

		' Field jrl_img
		Call journal.jrl_img.SetDbValue(Rs, journal.jrl_img.CurrentValue, Null, False)

		' Field jrl_create
		Call journal.jrl_create.SetDbValue(Rs, journal.jrl_create.CurrentValue, Null, False)

		' Field jrl_update
		Call journal.jrl_update.SetDbValue(Rs, journal.jrl_update.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = journal.Row_Inserting(RsOld, Rs)
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
			ElseIf journal.CancelMessage <> "" Then
				FailureMessage = journal.CancelMessage
				journal.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("InsertCancelled")
			End If
			AddRow = False
		End If
		Rs.Close
		Set Rs = Nothing
		If AddRow Then
			journal.jrl_id.DbValue = RsNew("jrl_id")
		End If
		If AddRow Then

			' Call Row Inserted event
			Call journal.Row_Inserted(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", journal.TableVar, "pom_journallist.asp", journal.TableVar, True)
		PageId = ew_IIf(journal.CurrentAction = "C", "Copy", "Add")
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
