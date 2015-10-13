<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim admins_view
Set admins_view = New cadmins_view
Set Page = admins_view

' Page init processing
admins_view.Page_Init()

' Page main processing
admins_view.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
admins_view.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If admins.Export = "" Then %>
<script type="text/javascript">
// Page object
var admins_view = new ew_Page("admins_view");
admins_view.PageID = "view"; // Page ID
var EW_PAGE_ID = admins_view.PageID; // For backward compatibility
// Form object
var fadminsview = new ew_Form("fadminsview");
// Form_CustomValidate event
fadminsview.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fadminsview.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fadminsview.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If admins.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If admins.Export = "" Then %>
<div class="ewViewExportOptions">
<% admins_view.ExportOptions.Render "body", "", "", "", "", "" %>
<% If Not admins_view.ExportOptions.UseDropDownButton Then %>
</div>
<div class="ewViewOtherOptions">
<% End If %>
<%
	admins_view.ActionOptions.Render "body", "", "", "", "", ""
	admins_view.DetailOptions.Render "body", "", "", "", "", ""
%>
</div>
<% End If %>
<% admins_view.ShowPageHeader() %>
<% admins_view.ShowMessage %>
<form name="fadminsview" id="fadminsview" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="admins">
<table class="ewGrid"><tr><td>
<table id="tbl_adminsview" class="table table-bordered table-striped">
<% If admins.admin_id.Visible Then ' admin_id %>
	<tr id="r_admin_id">
		<td><span id="elh_admins_admin_id"><%= admins.admin_id.FldCaption %></span></td>
		<td<%= admins.admin_id.CellAttributes %>>
<span id="el_admins_admin_id" class="control-group">
<span<%= admins.admin_id.ViewAttributes %>>
<%= admins.admin_id.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If admins.admin_username.Visible Then ' admin_username %>
	<tr id="r_admin_username">
		<td><span id="elh_admins_admin_username"><%= admins.admin_username.FldCaption %></span></td>
		<td<%= admins.admin_username.CellAttributes %>>
<span id="el_admins_admin_username" class="control-group">
<span<%= admins.admin_username.ViewAttributes %>>
<%= admins.admin_username.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If admins.admin_password.Visible Then ' admin_password %>
	<tr id="r_admin_password">
		<td><span id="elh_admins_admin_password"><%= admins.admin_password.FldCaption %></span></td>
		<td<%= admins.admin_password.CellAttributes %>>
<span id="el_admins_admin_password" class="control-group">
<span<%= admins.admin_password.ViewAttributes %>>
<%= admins.admin_password.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If admins.admin_name.Visible Then ' admin_name %>
	<tr id="r_admin_name">
		<td><span id="elh_admins_admin_name"><%= admins.admin_name.FldCaption %></span></td>
		<td<%= admins.admin_name.CellAttributes %>>
<span id="el_admins_admin_name" class="control-group">
<span<%= admins.admin_name.ViewAttributes %>>
<%= admins.admin_name.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If admins.admin_email.Visible Then ' admin_email %>
	<tr id="r_admin_email">
		<td><span id="elh_admins_admin_email"><%= admins.admin_email.FldCaption %></span></td>
		<td<%= admins.admin_email.CellAttributes %>>
<span id="el_admins_admin_email" class="control-group">
<span<%= admins.admin_email.ViewAttributes %>>
<%= admins.admin_email.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If admins.admin_tel.Visible Then ' admin_tel %>
	<tr id="r_admin_tel">
		<td><span id="elh_admins_admin_tel"><%= admins.admin_tel.FldCaption %></span></td>
		<td<%= admins.admin_tel.CellAttributes %>>
<span id="el_admins_admin_tel" class="control-group">
<span<%= admins.admin_tel.ViewAttributes %>>
<%= admins.admin_tel.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If admins.admin_permis.Visible Then ' admin_permis %>
	<tr id="r_admin_permis">
		<td><span id="elh_admins_admin_permis"><%= admins.admin_permis.FldCaption %></span></td>
		<td<%= admins.admin_permis.CellAttributes %>>
<span id="el_admins_admin_permis" class="control-group">
<span<%= admins.admin_permis.ViewAttributes %>>
<%= admins.admin_permis.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If admins.admin_create.Visible Then ' admin_create %>
	<tr id="r_admin_create">
		<td><span id="elh_admins_admin_create"><%= admins.admin_create.FldCaption %></span></td>
		<td<%= admins.admin_create.CellAttributes %>>
<span id="el_admins_admin_create" class="control-group">
<span<%= admins.admin_create.ViewAttributes %>>
<%= admins.admin_create.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If admins.admin_update.Visible Then ' admin_update %>
	<tr id="r_admin_update">
		<td><span id="elh_admins_admin_update"><%= admins.admin_update.FldCaption %></span></td>
		<td<%= admins.admin_update.CellAttributes %>>
<span id="el_admins_admin_update" class="control-group">
<span<%= admins.admin_update.ViewAttributes %>>
<%= admins.admin_update.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If admins.last_online.Visible Then ' last_online %>
	<tr id="r_last_online">
		<td><span id="elh_admins_last_online"><%= admins.last_online.FldCaption %></span></td>
		<td<%= admins.last_online.CellAttributes %>>
<span id="el_admins_last_online" class="control-group">
<span<%= admins.last_online.ViewAttributes %>>
<%= admins.last_online.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
</table>
</td></tr></table>
</form>
<script type="text/javascript">
fadminsview.Init();
</script>
<%
admins_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If admins.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set admins_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cadmins_view

	' Page ID
	Public Property Get PageID()
		PageID = "view"
	End Property

	' Project ID
	Public Property Get ProjectID()
		ProjectID = "{324ED72D-DE20-46F7-B12E-7AF8CE8711A6}"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "admins"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "admins_view"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If admins.UseTokenInUrl Then PageUrl = PageUrl & "t=" & admins.TableVar & "&" ' add page token
	End Property

	' Common urls
	Dim AddUrl
	Dim EditUrl
	Dim CopyUrl
	Dim DeleteUrl
	Dim ViewUrl
	Dim ListUrl

	' Export urls
	Dim ExportPrintUrl
	Dim ExportHtmlUrl
	Dim ExportExcelUrl
	Dim ExportWordUrl
	Dim ExportXmlUrl
	Dim ExportCsvUrl
	Dim ExportPdfUrl

	' Inline urls
	Dim InlineAddUrl
	Dim InlineCopyUrl
	Dim InlineEditUrl
	Dim GridAddUrl
	Dim GridEditUrl
	Dim MultiDeleteUrl
	Dim MultiUpdateUrl

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
		If admins.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (admins.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (admins.TableVar = Request.QueryString("t"))
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
		If IsEmpty(admins) Then Set admins = New cadmins
		Set Table = admins

		' Initialize urls
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("admin_id").Count > 0 Then
			ew_AddKey RecKey, "admin_id", Request.QueryString("admin_id")
			KeyUrl = KeyUrl & "&amp;admin_id=" & Server.URLEncode(Request.QueryString("admin_id"))
		End If
		ExportPrintUrl = PageUrl & "export=print" & KeyUrl
		ExportHtmlUrl = PageUrl & "export=html" & KeyUrl
		ExportExcelUrl = PageUrl & "export=excel" & KeyUrl
		ExportWordUrl = PageUrl & "export=word" & KeyUrl
		ExportXmlUrl = PageUrl & "export=xml" & KeyUrl
		ExportCsvUrl = PageUrl & "export=csv" & KeyUrl
		ExportPdfUrl = PageUrl & "export=pdf" & KeyUrl

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "view"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "admins"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = admins.TableVar
		ExportOptions.Tag = "div"
		ExportOptions.TagClassName = "ewExportOption"

		' Other options
		Set ActionOptions = New cListOptions
		ActionOptions.Tag = "div"
		ActionOptions.TagClassName = "ewActionOption"
		Set DetailOptions = New cListOptions
		DetailOptions.Tag = "div"
		DetailOptions.TagClassName = "ewDetailOption"
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
		Set admins = Nothing
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

	Dim DisplayRecs ' Number of display records
	Dim StartRec, StopRec, TotalRecs, RecRange
	Dim RecCnt
	Dim RecKey
	Dim ExportOptions ' Export options
	Dim DetailOptions ' Other options (detail)
	Dim ActionOptions ' Other options (action)
	Dim Recordset

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Paging variables
		DisplayRecs = 1
		RecRange = 10

		' Load current record
		Dim bLoadCurrentRecord
		bLoadCurrentRecord = False
		Dim sReturnUrl
		sReturnUrl = ""
		Dim bMatchRecord
		bMatchRecord = False

		' Set up Breadcrumb
		If admins.Export = "" Then
			SetupBreadcrumb()
		End If
		If IsPageRequest Then ' Validate request
			If Request.QueryString("admin_id").Count > 0 Then
				admins.admin_id.QueryStringValue = Request.QueryString("admin_id")
			Else
				sReturnUrl = "pom_adminslist.asp" ' Return to list
			End If

			' Get action
			admins.CurrentAction = "I" ' Display form
			Select Case admins.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "pom_adminslist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "pom_adminslist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		admins.RowType = EW_ROWTYPE_VIEW
		Call admins.ResetAttrs()
		Call RenderRow()
	End Sub

	' Set up other options
	Sub SetupOtherOptions()
		Dim opt, item
		Set opt = ActionOptions

		' Add
		Call opt.Add("add")
		Set item = opt.GetItem("add")
		item.Body = "<a class=""ewAction ewAdd"" href=""" & ew_HtmlEncode(AddUrl) & """>" & Language.Phrase("ViewPageAddLink") & "</a>"
		item.Visible = (AddUrl <> "" And Security.IsLoggedIn())

		' Edit
		Call opt.Add("edit")
		Set item = opt.GetItem("edit")
		item.Body = "<a class=""ewAction ewEdit"" href=""" & ew_HtmlEncode(EditUrl) & """>" & Language.Phrase("ViewPageEditLink") & "</a>"
		item.Visible = (EditUrl <> "" And Security.IsLoggedIn())

		' Copy
		Call opt.Add("copy")
		Set item = opt.GetItem("copy")
		item.Body = "<a class=""ewAction ewCopy"" href=""" & ew_HtmlEncode(CopyUrl) & """>" & Language.Phrase("ViewPageCopyLink") & "</a>"
		item.Visible = (CopyUrl <> "" And Security.IsLoggedIn())

		' Delete
		Call opt.Add("delete")
		Set item = opt.GetItem("delete")
		item.Body = "<a class=""ewAction ewDelete"" href=""" & ew_HtmlEncode(DeleteUrl) & """>" & Language.Phrase("ViewPageDeleteLink") & "</a>"
		item.Visible = (DeleteUrl <> "" And Security.IsLoggedIn())

		' Set up options default
		Set opt = ActionOptions
		opt.DropDownButtonPhrase = Language.Phrase("ButtonActions")
		opt.UseDropDownButton = False
		opt.UseButtonGroup = True
		Call opt.Add(opt.GroupOptionName)
		Set item = opt.GetItem(opt.GroupOptionName)
		item.Body = ""
		item.Visible = False
		Set opt = DetailOptions
		opt.DropDownButtonPhrase = Language.Phrase("ButtonDetails")
		opt.UseDropDownButton = False
		opt.UseButtonGroup = True
		Call opt.Add(opt.GroupOptionName)
		Set item = opt.GetItem(opt.GroupOptionName)
		item.Body = ""
		item.Visible = False
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
				admins.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					admins.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = admins.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			admins.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			admins.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			admins.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = admins.KeyFilter

		' Call Row Selecting event
		Call admins.Row_Selecting(sFilter)

		' Load sql based on filter
		admins.CurrentFilter = sFilter
		sSql = admins.SQL
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
		Call admins.Row_Selected(RsRow)
		admins.admin_id.DbValue = RsRow("admin_id")
		admins.admin_username.DbValue = RsRow("admin_username")
		admins.admin_password.DbValue = RsRow("admin_password")
		admins.admin_name.DbValue = RsRow("admin_name")
		admins.admin_email.DbValue = RsRow("admin_email")
		admins.admin_tel.DbValue = RsRow("admin_tel")
		admins.admin_permis.DbValue = RsRow("admin_permis")
		admins.admin_create.DbValue = RsRow("admin_create")
		admins.admin_update.DbValue = RsRow("admin_update")
		admins.last_online.DbValue = RsRow("last_online")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		admins.admin_id.m_DbValue = Rs("admin_id")
		admins.admin_username.m_DbValue = Rs("admin_username")
		admins.admin_password.m_DbValue = Rs("admin_password")
		admins.admin_name.m_DbValue = Rs("admin_name")
		admins.admin_email.m_DbValue = Rs("admin_email")
		admins.admin_tel.m_DbValue = Rs("admin_tel")
		admins.admin_permis.m_DbValue = Rs("admin_permis")
		admins.admin_create.m_DbValue = Rs("admin_create")
		admins.admin_update.m_DbValue = Rs("admin_update")
		admins.last_online.m_DbValue = Rs("last_online")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = admins.AddUrl
		EditUrl = admins.EditUrl("")
		CopyUrl = admins.CopyUrl("")
		DeleteUrl = admins.DeleteUrl
		ListUrl = admins.ListUrl
		SetupOtherOptions()

		' Call Row Rendering event
		Call admins.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' admin_id
		' admin_username
		' admin_password
		' admin_name
		' admin_email
		' admin_tel
		' admin_permis
		' admin_create
		' admin_update
		' last_online
		' -----------
		'  View  Row
		' -----------

		If admins.RowType = EW_ROWTYPE_VIEW Then ' View row

			' admin_id
			admins.admin_id.ViewValue = admins.admin_id.CurrentValue
			admins.admin_id.ViewCustomAttributes = ""

			' admin_username
			admins.admin_username.ViewValue = admins.admin_username.CurrentValue
			admins.admin_username.ViewCustomAttributes = ""

			' admin_password
			admins.admin_password.ViewValue = admins.admin_password.CurrentValue
			admins.admin_password.ViewCustomAttributes = ""

			' admin_name
			admins.admin_name.ViewValue = admins.admin_name.CurrentValue
			admins.admin_name.ViewCustomAttributes = ""

			' admin_email
			admins.admin_email.ViewValue = admins.admin_email.CurrentValue
			admins.admin_email.ViewCustomAttributes = ""

			' admin_tel
			admins.admin_tel.ViewValue = admins.admin_tel.CurrentValue
			admins.admin_tel.ViewCustomAttributes = ""

			' admin_permis
			admins.admin_permis.ViewValue = admins.admin_permis.CurrentValue
			admins.admin_permis.ViewCustomAttributes = ""

			' admin_create
			admins.admin_create.ViewValue = admins.admin_create.CurrentValue
			admins.admin_create.ViewCustomAttributes = ""

			' admin_update
			admins.admin_update.ViewValue = admins.admin_update.CurrentValue
			admins.admin_update.ViewCustomAttributes = ""

			' last_online
			admins.last_online.ViewValue = admins.last_online.CurrentValue
			admins.last_online.ViewCustomAttributes = ""

			' View refer script
			' admin_id

			admins.admin_id.LinkCustomAttributes = ""
			admins.admin_id.HrefValue = ""
			admins.admin_id.TooltipValue = ""

			' admin_username
			admins.admin_username.LinkCustomAttributes = ""
			admins.admin_username.HrefValue = ""
			admins.admin_username.TooltipValue = ""

			' admin_password
			admins.admin_password.LinkCustomAttributes = ""
			admins.admin_password.HrefValue = ""
			admins.admin_password.TooltipValue = ""

			' admin_name
			admins.admin_name.LinkCustomAttributes = ""
			admins.admin_name.HrefValue = ""
			admins.admin_name.TooltipValue = ""

			' admin_email
			admins.admin_email.LinkCustomAttributes = ""
			admins.admin_email.HrefValue = ""
			admins.admin_email.TooltipValue = ""

			' admin_tel
			admins.admin_tel.LinkCustomAttributes = ""
			admins.admin_tel.HrefValue = ""
			admins.admin_tel.TooltipValue = ""

			' admin_permis
			admins.admin_permis.LinkCustomAttributes = ""
			admins.admin_permis.HrefValue = ""
			admins.admin_permis.TooltipValue = ""

			' admin_create
			admins.admin_create.LinkCustomAttributes = ""
			admins.admin_create.HrefValue = ""
			admins.admin_create.TooltipValue = ""

			' admin_update
			admins.admin_update.LinkCustomAttributes = ""
			admins.admin_update.HrefValue = ""
			admins.admin_update.TooltipValue = ""

			' last_online
			admins.last_online.LinkCustomAttributes = ""
			admins.last_online.HrefValue = ""
			admins.last_online.TooltipValue = ""
		End If

		' Call Row Rendered event
		If admins.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call admins.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", admins.TableVar, "pom_adminslist.asp", admins.TableVar, True)
		PageId = "view"
		Call Breadcrumb.Add("view", PageId, ew_CurrentUrl, "", False)
	End Sub

	Sub ExportPdf(html)
		Response.Write html
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
End Class
%>
