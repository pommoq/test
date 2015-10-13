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
Dim journal_view
Set journal_view = New cjournal_view
Set Page = journal_view

' Page init processing
journal_view.Page_Init()

' Page main processing
journal_view.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
journal_view.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If journal.Export = "" Then %>
<script type="text/javascript">
// Page object
var journal_view = new ew_Page("journal_view");
journal_view.PageID = "view"; // Page ID
var EW_PAGE_ID = journal_view.PageID; // For backward compatibility
// Form object
var fjournalview = new ew_Form("fjournalview");
// Form_CustomValidate event
fjournalview.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fjournalview.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fjournalview.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If journal.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If journal.Export = "" Then %>
<div class="ewViewExportOptions">
<% journal_view.ExportOptions.Render "body", "", "", "", "", "" %>
<% If Not journal_view.ExportOptions.UseDropDownButton Then %>
</div>
<div class="ewViewOtherOptions">
<% End If %>
<%
	journal_view.ActionOptions.Render "body", "", "", "", "", ""
	journal_view.DetailOptions.Render "body", "", "", "", "", ""
%>
</div>
<% End If %>
<% journal_view.ShowPageHeader() %>
<% journal_view.ShowMessage %>
<form name="fjournalview" id="fjournalview" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="journal">
<table class="ewGrid"><tr><td>
<table id="tbl_journalview" class="table table-bordered table-striped">
<% If journal.jrl_id.Visible Then ' jrl_id %>
	<tr id="r_jrl_id">
		<td><span id="elh_journal_jrl_id"><%= journal.jrl_id.FldCaption %></span></td>
		<td<%= journal.jrl_id.CellAttributes %>>
<span id="el_journal_jrl_id" class="control-group">
<span<%= journal.jrl_id.ViewAttributes %>>
<%= journal.jrl_id.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If journal.jrl_category.Visible Then ' jrl_category %>
	<tr id="r_jrl_category">
		<td><span id="elh_journal_jrl_category"><%= journal.jrl_category.FldCaption %></span></td>
		<td<%= journal.jrl_category.CellAttributes %>>
<span id="el_journal_jrl_category" class="control-group">
<span<%= journal.jrl_category.ViewAttributes %>>
<%= journal.jrl_category.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If journal.jrl_date.Visible Then ' jrl_date %>
	<tr id="r_jrl_date">
		<td><span id="elh_journal_jrl_date"><%= journal.jrl_date.FldCaption %></span></td>
		<td<%= journal.jrl_date.CellAttributes %>>
<span id="el_journal_jrl_date" class="control-group">
<span<%= journal.jrl_date.ViewAttributes %>>
<%= journal.jrl_date.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If journal.jrl_title.Visible Then ' jrl_title %>
	<tr id="r_jrl_title">
		<td><span id="elh_journal_jrl_title"><%= journal.jrl_title.FldCaption %></span></td>
		<td<%= journal.jrl_title.CellAttributes %>>
<span id="el_journal_jrl_title" class="control-group">
<span<%= journal.jrl_title.ViewAttributes %>>
<%= journal.jrl_title.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If journal.jrl_title_th.Visible Then ' jrl_title_th %>
	<tr id="r_jrl_title_th">
		<td><span id="elh_journal_jrl_title_th"><%= journal.jrl_title_th.FldCaption %></span></td>
		<td<%= journal.jrl_title_th.CellAttributes %>>
<span id="el_journal_jrl_title_th" class="control-group">
<span<%= journal.jrl_title_th.ViewAttributes %>>
<%= journal.jrl_title_th.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If journal.jrl_pdf.Visible Then ' jrl_pdf %>
	<tr id="r_jrl_pdf">
		<td><span id="elh_journal_jrl_pdf"><%= journal.jrl_pdf.FldCaption %></span></td>
		<td<%= journal.jrl_pdf.CellAttributes %>>
<span id="el_journal_jrl_pdf" class="control-group">
<span<%= journal.jrl_pdf.ViewAttributes %>>
<%= journal.jrl_pdf.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If journal.jrl_img.Visible Then ' jrl_img %>
	<tr id="r_jrl_img">
		<td><span id="elh_journal_jrl_img"><%= journal.jrl_img.FldCaption %></span></td>
		<td<%= journal.jrl_img.CellAttributes %>>
<span id="el_journal_jrl_img" class="control-group">
<span<%= journal.jrl_img.ViewAttributes %>>
<%= journal.jrl_img.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If journal.jrl_create.Visible Then ' jrl_create %>
	<tr id="r_jrl_create">
		<td><span id="elh_journal_jrl_create"><%= journal.jrl_create.FldCaption %></span></td>
		<td<%= journal.jrl_create.CellAttributes %>>
<span id="el_journal_jrl_create" class="control-group">
<span<%= journal.jrl_create.ViewAttributes %>>
<%= journal.jrl_create.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If journal.jrl_update.Visible Then ' jrl_update %>
	<tr id="r_jrl_update">
		<td><span id="elh_journal_jrl_update"><%= journal.jrl_update.FldCaption %></span></td>
		<td<%= journal.jrl_update.CellAttributes %>>
<span id="el_journal_jrl_update" class="control-group">
<span<%= journal.jrl_update.ViewAttributes %>>
<%= journal.jrl_update.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
</table>
</td></tr></table>
</form>
<script type="text/javascript">
fjournalview.Init();
</script>
<%
journal_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If journal.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set journal_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cjournal_view

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
		TableName = "journal"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "journal_view"
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
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("jrl_id").Count > 0 Then
			ew_AddKey RecKey, "jrl_id", Request.QueryString("jrl_id")
			KeyUrl = KeyUrl & "&amp;jrl_id=" & Server.URLEncode(Request.QueryString("jrl_id"))
		End If
		ExportPrintUrl = PageUrl & "export=print" & KeyUrl
		ExportHtmlUrl = PageUrl & "export=html" & KeyUrl
		ExportExcelUrl = PageUrl & "export=excel" & KeyUrl
		ExportWordUrl = PageUrl & "export=word" & KeyUrl
		ExportXmlUrl = PageUrl & "export=xml" & KeyUrl
		ExportCsvUrl = PageUrl & "export=csv" & KeyUrl
		ExportPdfUrl = PageUrl & "export=pdf" & KeyUrl

		' Initialize other table object
		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "view"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "journal"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = journal.TableVar
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
		If journal.Export = "" Then
			SetupBreadcrumb()
		End If
		If IsPageRequest Then ' Validate request
			If Request.QueryString("jrl_id").Count > 0 Then
				journal.jrl_id.QueryStringValue = Request.QueryString("jrl_id")
			Else
				sReturnUrl = "pom_journallist.asp" ' Return to list
			End If

			' Get action
			journal.CurrentAction = "I" ' Display form
			Select Case journal.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "pom_journallist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "pom_journallist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		journal.RowType = EW_ROWTYPE_VIEW
		Call journal.ResetAttrs()
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
				journal.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					journal.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = journal.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			journal.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			journal.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			journal.StartRecordNumber = StartRec
		End If
	End Sub

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

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = journal.AddUrl
		EditUrl = journal.EditUrl("")
		CopyUrl = journal.CopyUrl("")
		DeleteUrl = journal.DeleteUrl
		ListUrl = journal.ListUrl
		SetupOtherOptions()

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
			' jrl_id

			journal.jrl_id.LinkCustomAttributes = ""
			journal.jrl_id.HrefValue = ""
			journal.jrl_id.TooltipValue = ""

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
		End If

		' Call Row Rendered event
		If journal.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call journal.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", journal.TableVar, "pom_journallist.asp", journal.TableVar, True)
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
