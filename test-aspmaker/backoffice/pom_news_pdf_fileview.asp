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
Dim news_pdf_file_view
Set news_pdf_file_view = New cnews_pdf_file_view
Set Page = news_pdf_file_view

' Page init processing
news_pdf_file_view.Page_Init()

' Page main processing
news_pdf_file_view.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
news_pdf_file_view.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If news_pdf_file.Export = "" Then %>
<script type="text/javascript">
// Page object
var news_pdf_file_view = new ew_Page("news_pdf_file_view");
news_pdf_file_view.PageID = "view"; // Page ID
var EW_PAGE_ID = news_pdf_file_view.PageID; // For backward compatibility
// Form object
var fnews_pdf_fileview = new ew_Form("fnews_pdf_fileview");
// Form_CustomValidate event
fnews_pdf_fileview.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fnews_pdf_fileview.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fnews_pdf_fileview.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If news_pdf_file.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If news_pdf_file.Export = "" Then %>
<div class="ewViewExportOptions">
<% news_pdf_file_view.ExportOptions.Render "body", "", "", "", "", "" %>
<% If Not news_pdf_file_view.ExportOptions.UseDropDownButton Then %>
</div>
<div class="ewViewOtherOptions">
<% End If %>
<%
	news_pdf_file_view.ActionOptions.Render "body", "", "", "", "", ""
	news_pdf_file_view.DetailOptions.Render "body", "", "", "", "", ""
%>
</div>
<% End If %>
<% news_pdf_file_view.ShowPageHeader() %>
<% news_pdf_file_view.ShowMessage %>
<form name="fnews_pdf_fileview" id="fnews_pdf_fileview" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="news_pdf_file">
<table class="ewGrid"><tr><td>
<table id="tbl_news_pdf_fileview" class="table table-bordered table-striped">
<% If news_pdf_file.news_pdf_id.Visible Then ' news_pdf_id %>
	<tr id="r_news_pdf_id">
		<td><span id="elh_news_pdf_file_news_pdf_id"><%= news_pdf_file.news_pdf_id.FldCaption %></span></td>
		<td<%= news_pdf_file.news_pdf_id.CellAttributes %>>
<span id="el_news_pdf_file_news_pdf_id" class="control-group">
<span<%= news_pdf_file.news_pdf_id.ViewAttributes %>>
<%= news_pdf_file.news_pdf_id.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news_pdf_file.news_id.Visible Then ' news_id %>
	<tr id="r_news_id">
		<td><span id="elh_news_pdf_file_news_id"><%= news_pdf_file.news_id.FldCaption %></span></td>
		<td<%= news_pdf_file.news_id.CellAttributes %>>
<span id="el_news_pdf_file_news_id" class="control-group">
<span<%= news_pdf_file.news_id.ViewAttributes %>>
<%= news_pdf_file.news_id.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news_pdf_file.news_pdf_file_1.Visible Then ' news_pdf_file %>
	<tr id="r_news_pdf_file_1">
		<td><span id="elh_news_pdf_file_news_pdf_file_1"><%= news_pdf_file.news_pdf_file_1.FldCaption %></span></td>
		<td<%= news_pdf_file.news_pdf_file_1.CellAttributes %>>
<span id="el_news_pdf_file_news_pdf_file_1" class="control-group">
<span<%= news_pdf_file.news_pdf_file_1.ViewAttributes %>>
<%= news_pdf_file.news_pdf_file_1.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news_pdf_file.news_pdf_title.Visible Then ' news_pdf_title %>
	<tr id="r_news_pdf_title">
		<td><span id="elh_news_pdf_file_news_pdf_title"><%= news_pdf_file.news_pdf_title.FldCaption %></span></td>
		<td<%= news_pdf_file.news_pdf_title.CellAttributes %>>
<span id="el_news_pdf_file_news_pdf_title" class="control-group">
<span<%= news_pdf_file.news_pdf_title.ViewAttributes %>>
<%= news_pdf_file.news_pdf_title.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
</table>
</td></tr></table>
</form>
<script type="text/javascript">
fnews_pdf_fileview.Init();
</script>
<%
news_pdf_file_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If news_pdf_file.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set news_pdf_file_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cnews_pdf_file_view

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
		TableName = "news_pdf_file"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "news_pdf_file_view"
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
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("news_pdf_id").Count > 0 Then
			ew_AddKey RecKey, "news_pdf_id", Request.QueryString("news_pdf_id")
			KeyUrl = KeyUrl & "&amp;news_pdf_id=" & Server.URLEncode(Request.QueryString("news_pdf_id"))
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
		EW_TABLE_NAME = "news_pdf_file"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = news_pdf_file.TableVar
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
		If news_pdf_file.Export = "" Then
			SetupBreadcrumb()
		End If
		If IsPageRequest Then ' Validate request
			If Request.QueryString("news_pdf_id").Count > 0 Then
				news_pdf_file.news_pdf_id.QueryStringValue = Request.QueryString("news_pdf_id")
			Else
				sReturnUrl = "pom_news_pdf_filelist.asp" ' Return to list
			End If

			' Get action
			news_pdf_file.CurrentAction = "I" ' Display form
			Select Case news_pdf_file.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "pom_news_pdf_filelist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "pom_news_pdf_filelist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		news_pdf_file.RowType = EW_ROWTYPE_VIEW
		Call news_pdf_file.ResetAttrs()
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
				news_pdf_file.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					news_pdf_file.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = news_pdf_file.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			news_pdf_file.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			news_pdf_file.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			news_pdf_file.StartRecordNumber = StartRec
		End If
	End Sub

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

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = news_pdf_file.AddUrl
		EditUrl = news_pdf_file.EditUrl("")
		CopyUrl = news_pdf_file.CopyUrl("")
		DeleteUrl = news_pdf_file.DeleteUrl
		ListUrl = news_pdf_file.ListUrl
		SetupOtherOptions()

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
			' news_pdf_id

			news_pdf_file.news_pdf_id.LinkCustomAttributes = ""
			news_pdf_file.news_pdf_id.HrefValue = ""
			news_pdf_file.news_pdf_id.TooltipValue = ""

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
		End If

		' Call Row Rendered event
		If news_pdf_file.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call news_pdf_file.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", news_pdf_file.TableVar, "pom_news_pdf_filelist.asp", news_pdf_file.TableVar, True)
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
