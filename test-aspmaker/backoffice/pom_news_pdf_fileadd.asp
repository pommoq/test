<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_news_pdf_fileinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim news_pdf_file_add
Set news_pdf_file_add = New cnews_pdf_file_add
Set Page = news_pdf_file_add

' Page init processing
news_pdf_file_add.Page_Init()

' Page main processing
news_pdf_file_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
news_pdf_file_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var news_pdf_file_add = new ew_Page("news_pdf_file_add");
news_pdf_file_add.PageID = "add"; // Page ID
var EW_PAGE_ID = news_pdf_file_add.PageID; // For backward compatibility
// Form object
var fnews_pdf_fileadd = new ew_Form("fnews_pdf_fileadd");
// Validate form
fnews_pdf_fileadd.Validate = function() {
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
			elm = this.GetElements("x" + infix + "_news_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(news_pdf_file.news_id.FldErrMsg) %>");
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
fnews_pdf_fileadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fnews_pdf_fileadd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fnews_pdf_fileadd.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If news_pdf_file.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% news_pdf_file_add.ShowPageHeader() %>
<% news_pdf_file_add.ShowMessage %>
<form name="fnews_pdf_fileadd" id="fnews_pdf_fileadd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="news_pdf_file">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_news_pdf_fileadd" class="table table-bordered table-striped">
<% If news_pdf_file.news_id.Visible Then ' news_id %>
	<tr id="r_news_id">
		<td><span id="elh_news_pdf_file_news_id"><%= news_pdf_file.news_id.FldCaption %></span></td>
		<td<%= news_pdf_file.news_id.CellAttributes %>>
<span id="el_news_pdf_file_news_id" class="control-group">
<input type="text" data-field="x_news_id" name="x_news_id" id="x_news_id" size="30" placeholder="<%= news_pdf_file.news_id.PlaceHolder %>" value="<%= news_pdf_file.news_id.EditValue %>"<%= news_pdf_file.news_id.EditAttributes %>>
</span>
<%= news_pdf_file.news_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_pdf_file.news_pdf_file_1.Visible Then ' news_pdf_file %>
	<tr id="r_news_pdf_file_1">
		<td><span id="elh_news_pdf_file_news_pdf_file_1"><%= news_pdf_file.news_pdf_file_1.FldCaption %></span></td>
		<td<%= news_pdf_file.news_pdf_file_1.CellAttributes %>>
<span id="el_news_pdf_file_news_pdf_file_1" class="control-group">
<input type="text" data-field="x_news_pdf_file_1" name="x_news_pdf_file_1" id="x_news_pdf_file_1" size="30" maxlength="255" placeholder="<%= news_pdf_file.news_pdf_file_1.PlaceHolder %>" value="<%= news_pdf_file.news_pdf_file_1.EditValue %>"<%= news_pdf_file.news_pdf_file_1.EditAttributes %>>
</span>
<%= news_pdf_file.news_pdf_file_1.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_pdf_file.news_pdf_title.Visible Then ' news_pdf_title %>
	<tr id="r_news_pdf_title">
		<td><span id="elh_news_pdf_file_news_pdf_title"><%= news_pdf_file.news_pdf_title.FldCaption %></span></td>
		<td<%= news_pdf_file.news_pdf_title.CellAttributes %>>
<span id="el_news_pdf_file_news_pdf_title" class="control-group">
<input type="text" data-field="x_news_pdf_title" name="x_news_pdf_title" id="x_news_pdf_title" size="30" maxlength="255" placeholder="<%= news_pdf_file.news_pdf_title.PlaceHolder %>" value="<%= news_pdf_file.news_pdf_title.EditValue %>"<%= news_pdf_file.news_pdf_title.EditAttributes %>>
</span>
<%= news_pdf_file.news_pdf_title.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
fnews_pdf_fileadd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
news_pdf_file_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set news_pdf_file_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cnews_pdf_file_add

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
		TableName = "news_pdf_file"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "news_pdf_file_add"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If news_pdf_file.UseTokenInUrl Then PageUrl = PageUrl & "t=" & news_pdf_file.TableVar & "&" ' add page token
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
		If news_pdf_file.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (news_pdf_file.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (news_pdf_file.TableVar = Request.QueryString("t"))
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
		If IsEmpty(news_pdf_file) Then Set news_pdf_file = New cnews_pdf_file
		Set Table = news_pdf_file

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "add"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "news_pdf_file"

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

		news_pdf_file.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set news_pdf_file = Nothing
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
			news_pdf_file.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("news_pdf_id").Count > 0 Then
				news_pdf_file.news_pdf_id.QueryStringValue = Request.QueryString("news_pdf_id")
				Call news_pdf_file.SetKey("news_pdf_id", news_pdf_file.news_pdf_id.CurrentValue) ' Set up key
			Else
				Call news_pdf_file.SetKey("news_pdf_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				news_pdf_file.CurrentAction = "C" ' Copy Record
			Else
				news_pdf_file.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				news_pdf_file.CurrentAction = "I" ' Form error, reset action
				news_pdf_file.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case news_pdf_file.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_news_pdf_filelist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				news_pdf_file.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = news_pdf_file.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_news_pdf_fileview.asp" Then sReturnUrl = news_pdf_file.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					news_pdf_file.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		news_pdf_file.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call news_pdf_file.ResetAttrs()
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
		news_pdf_file.news_id.CurrentValue = Null
		news_pdf_file.news_id.OldValue = news_pdf_file.news_id.CurrentValue
		news_pdf_file.news_pdf_file_1.CurrentValue = Null
		news_pdf_file.news_pdf_file_1.OldValue = news_pdf_file.news_pdf_file_1.CurrentValue
		news_pdf_file.news_pdf_title.CurrentValue = Null
		news_pdf_file.news_pdf_title.OldValue = news_pdf_file.news_pdf_title.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not news_pdf_file.news_id.FldIsDetailKey Then news_pdf_file.news_id.FormValue = ObjForm.GetValue("x_news_id")
		If Not news_pdf_file.news_pdf_file_1.FldIsDetailKey Then news_pdf_file.news_pdf_file_1.FormValue = ObjForm.GetValue("x_news_pdf_file_1")
		If Not news_pdf_file.news_pdf_title.FldIsDetailKey Then news_pdf_file.news_pdf_title.FormValue = ObjForm.GetValue("x_news_pdf_title")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		news_pdf_file.news_id.CurrentValue = news_pdf_file.news_id.FormValue
		news_pdf_file.news_pdf_file_1.CurrentValue = news_pdf_file.news_pdf_file_1.FormValue
		news_pdf_file.news_pdf_title.CurrentValue = news_pdf_file.news_pdf_title.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = news_pdf_file.KeyFilter

		' Call Row Selecting event
		Call news_pdf_file.Row_Selecting(sFilter)

		' Load sql based on filter
		news_pdf_file.CurrentFilter = sFilter
		sSql = news_pdf_file.SQL
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
		Call news_pdf_file.Row_Selected(RsRow)
		news_pdf_file.news_pdf_id.DbValue = RsRow("news_pdf_id")
		news_pdf_file.news_id.DbValue = RsRow("news_id")
		news_pdf_file.news_pdf_file_1.DbValue = RsRow("news_pdf_file")
		news_pdf_file.news_pdf_title.DbValue = RsRow("news_pdf_title")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		news_pdf_file.news_pdf_id.m_DbValue = Rs("news_pdf_id")
		news_pdf_file.news_id.m_DbValue = Rs("news_id")
		news_pdf_file.news_pdf_file_1.m_DbValue = Rs("news_pdf_file")
		news_pdf_file.news_pdf_title.m_DbValue = Rs("news_pdf_title")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If news_pdf_file.GetKey("news_pdf_id")&"" <> "" Then
			news_pdf_file.news_pdf_id.CurrentValue = news_pdf_file.GetKey("news_pdf_id") ' news_pdf_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			news_pdf_file.CurrentFilter = news_pdf_file.KeyFilter
			Dim sSql
			sSql = news_pdf_file.SQL
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

		Call news_pdf_file.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' news_pdf_id
		' news_id
		' news_pdf_file
		' news_pdf_title
		' -----------
		'  View  Row
		' -----------

		If news_pdf_file.RowType = EW_ROWTYPE_VIEW Then ' View row

			' news_pdf_id
			news_pdf_file.news_pdf_id.ViewValue = news_pdf_file.news_pdf_id.CurrentValue
			news_pdf_file.news_pdf_id.ViewCustomAttributes = ""

			' news_id
			news_pdf_file.news_id.ViewValue = news_pdf_file.news_id.CurrentValue
			news_pdf_file.news_id.ViewCustomAttributes = ""

			' news_pdf_file
			news_pdf_file.news_pdf_file_1.ViewValue = news_pdf_file.news_pdf_file_1.CurrentValue
			news_pdf_file.news_pdf_file_1.ViewCustomAttributes = ""

			' news_pdf_title
			news_pdf_file.news_pdf_title.ViewValue = news_pdf_file.news_pdf_title.CurrentValue
			news_pdf_file.news_pdf_title.ViewCustomAttributes = ""

			' View refer script
			' news_id

			news_pdf_file.news_id.LinkCustomAttributes = ""
			news_pdf_file.news_id.HrefValue = ""
			news_pdf_file.news_id.TooltipValue = ""

			' news_pdf_file
			news_pdf_file.news_pdf_file_1.LinkCustomAttributes = ""
			news_pdf_file.news_pdf_file_1.HrefValue = ""
			news_pdf_file.news_pdf_file_1.TooltipValue = ""

			' news_pdf_title
			news_pdf_file.news_pdf_title.LinkCustomAttributes = ""
			news_pdf_file.news_pdf_title.HrefValue = ""
			news_pdf_file.news_pdf_title.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf news_pdf_file.RowType = EW_ROWTYPE_ADD Then ' Add row

			' news_id
			news_pdf_file.news_id.EditCustomAttributes = ""
			news_pdf_file.news_id.EditValue = ew_HtmlEncode(news_pdf_file.news_id.CurrentValue)
			news_pdf_file.news_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_pdf_file.news_id.FldCaption))

			' news_pdf_file
			news_pdf_file.news_pdf_file_1.EditCustomAttributes = ""
			news_pdf_file.news_pdf_file_1.EditValue = ew_HtmlEncode(news_pdf_file.news_pdf_file_1.CurrentValue)
			news_pdf_file.news_pdf_file_1.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_pdf_file.news_pdf_file_1.FldCaption))

			' news_pdf_title
			news_pdf_file.news_pdf_title.EditCustomAttributes = ""
			news_pdf_file.news_pdf_title.EditValue = ew_HtmlEncode(news_pdf_file.news_pdf_title.CurrentValue)
			news_pdf_file.news_pdf_title.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_pdf_file.news_pdf_title.FldCaption))

			' Edit refer script
			' news_id

			news_pdf_file.news_id.HrefValue = ""

			' news_pdf_file
			news_pdf_file.news_pdf_file_1.HrefValue = ""

			' news_pdf_title
			news_pdf_file.news_pdf_title.HrefValue = ""
		End If
		If news_pdf_file.RowType = EW_ROWTYPE_ADD Or news_pdf_file.RowType = EW_ROWTYPE_EDIT Or news_pdf_file.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call news_pdf_file.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If news_pdf_file.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call news_pdf_file.Row_Rendered()
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
		If Not ew_CheckInteger(news_pdf_file.news_id.FormValue) Then
			Call ew_AddMessage(gsFormError, news_pdf_file.news_id.FldErrMsg)
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
		news_pdf_file.CurrentFilter = sFilter
		sSql = news_pdf_file.SQL
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

		' Field news_id
		Call news_pdf_file.news_id.SetDbValue(Rs, news_pdf_file.news_id.CurrentValue, Null, False)

		' Field news_pdf_file
		Call news_pdf_file.news_pdf_file_1.SetDbValue(Rs, news_pdf_file.news_pdf_file_1.CurrentValue, Null, False)

		' Field news_pdf_title
		Call news_pdf_file.news_pdf_title.SetDbValue(Rs, news_pdf_file.news_pdf_title.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = news_pdf_file.Row_Inserting(RsOld, Rs)
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
			ElseIf news_pdf_file.CancelMessage <> "" Then
				FailureMessage = news_pdf_file.CancelMessage
				news_pdf_file.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("InsertCancelled")
			End If
			AddRow = False
		End If
		Rs.Close
		Set Rs = Nothing
		If AddRow Then
			news_pdf_file.news_pdf_id.DbValue = RsNew("news_pdf_id")
		End If
		If AddRow Then

			' Call Row Inserted event
			Call news_pdf_file.Row_Inserted(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", news_pdf_file.TableVar, "pom_news_pdf_filelist.asp", news_pdf_file.TableVar, True)
		PageId = ew_IIf(news_pdf_file.CurrentAction = "C", "Copy", "Add")
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
