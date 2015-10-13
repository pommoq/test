<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_news_saleinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim news_sale_view
Set news_sale_view = New cnews_sale_view
Set Page = news_sale_view

' Page init processing
news_sale_view.Page_Init()

' Page main processing
news_sale_view.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
news_sale_view.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If news_sale.Export = "" Then %>
<script type="text/javascript">
// Page object
var news_sale_view = new ew_Page("news_sale_view");
news_sale_view.PageID = "view"; // Page ID
var EW_PAGE_ID = news_sale_view.PageID; // For backward compatibility
// Form object
var fnews_saleview = new ew_Form("fnews_saleview");
// Form_CustomValidate event
fnews_saleview.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fnews_saleview.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fnews_saleview.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If news_sale.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If news_sale.Export = "" Then %>
<div class="ewViewExportOptions">
<% news_sale_view.ExportOptions.Render "body", "", "", "", "", "" %>
<% If Not news_sale_view.ExportOptions.UseDropDownButton Then %>
</div>
<div class="ewViewOtherOptions">
<% End If %>
<%
	news_sale_view.ActionOptions.Render "body", "", "", "", "", ""
	news_sale_view.DetailOptions.Render "body", "", "", "", "", ""
%>
</div>
<% End If %>
<% news_sale_view.ShowPageHeader() %>
<% news_sale_view.ShowMessage %>
<form name="fnews_saleview" id="fnews_saleview" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="news_sale">
<table class="ewGrid"><tr><td>
<table id="tbl_news_saleview" class="table table-bordered table-striped">
<% If news_sale.news_sale_id.Visible Then ' news_sale_id %>
	<tr id="r_news_sale_id">
		<td><span id="elh_news_sale_news_sale_id"><%= news_sale.news_sale_id.FldCaption %></span></td>
		<td<%= news_sale.news_sale_id.CellAttributes %>>
<span id="el_news_sale_news_sale_id" class="control-group">
<span<%= news_sale.news_sale_id.ViewAttributes %>>
<%= news_sale.news_sale_id.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news_sale.news_sale_pdf.Visible Then ' news_sale_pdf %>
	<tr id="r_news_sale_pdf">
		<td><span id="elh_news_sale_news_sale_pdf"><%= news_sale.news_sale_pdf.FldCaption %></span></td>
		<td<%= news_sale.news_sale_pdf.CellAttributes %>>
<span id="el_news_sale_news_sale_pdf" class="control-group">
<span<%= news_sale.news_sale_pdf.ViewAttributes %>>
<%= news_sale.news_sale_pdf.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news_sale.news_sale_title.Visible Then ' news_sale_title %>
	<tr id="r_news_sale_title">
		<td><span id="elh_news_sale_news_sale_title"><%= news_sale.news_sale_title.FldCaption %></span></td>
		<td<%= news_sale.news_sale_title.CellAttributes %>>
<span id="el_news_sale_news_sale_title" class="control-group">
<span<%= news_sale.news_sale_title.ViewAttributes %>>
<%= news_sale.news_sale_title.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news_sale.start_date.Visible Then ' start_date %>
	<tr id="r_start_date">
		<td><span id="elh_news_sale_start_date"><%= news_sale.start_date.FldCaption %></span></td>
		<td<%= news_sale.start_date.CellAttributes %>>
<span id="el_news_sale_start_date" class="control-group">
<span<%= news_sale.start_date.ViewAttributes %>>
<%= news_sale.start_date.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news_sale.end_date.Visible Then ' end_date %>
	<tr id="r_end_date">
		<td><span id="elh_news_sale_end_date"><%= news_sale.end_date.FldCaption %></span></td>
		<td<%= news_sale.end_date.CellAttributes %>>
<span id="el_news_sale_end_date" class="control-group">
<span<%= news_sale.end_date.ViewAttributes %>>
<%= news_sale.end_date.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
</table>
</td></tr></table>
</form>
<script type="text/javascript">
fnews_saleview.Init();
</script>
<%
news_sale_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If news_sale.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set news_sale_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cnews_sale_view

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
		TableName = "news_sale"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "news_sale_view"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If news_sale.UseTokenInUrl Then PageUrl = PageUrl & "t=" & news_sale.TableVar & "&" ' add page token
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
		If news_sale.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (news_sale.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (news_sale.TableVar = Request.QueryString("t"))
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
		If IsEmpty(news_sale) Then Set news_sale = New cnews_sale
		Set Table = news_sale

		' Initialize urls
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("news_sale_id").Count > 0 Then
			ew_AddKey RecKey, "news_sale_id", Request.QueryString("news_sale_id")
			KeyUrl = KeyUrl & "&amp;news_sale_id=" & Server.URLEncode(Request.QueryString("news_sale_id"))
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
		EW_TABLE_NAME = "news_sale"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = news_sale.TableVar
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
		Set news_sale = Nothing
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
		If news_sale.Export = "" Then
			SetupBreadcrumb()
		End If
		If IsPageRequest Then ' Validate request
			If Request.QueryString("news_sale_id").Count > 0 Then
				news_sale.news_sale_id.QueryStringValue = Request.QueryString("news_sale_id")
			Else
				sReturnUrl = "pom_news_salelist.asp" ' Return to list
			End If

			' Get action
			news_sale.CurrentAction = "I" ' Display form
			Select Case news_sale.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "pom_news_salelist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "pom_news_salelist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		news_sale.RowType = EW_ROWTYPE_VIEW
		Call news_sale.ResetAttrs()
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
				news_sale.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					news_sale.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = news_sale.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			news_sale.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			news_sale.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			news_sale.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = news_sale.KeyFilter

		' Call Row Selecting event
		Call news_sale.Row_Selecting(sFilter)

		' Load sql based on filter
		news_sale.CurrentFilter = sFilter
		sSql = news_sale.SQL
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
		Call news_sale.Row_Selected(RsRow)
		news_sale.news_sale_id.DbValue = RsRow("news_sale_id")
		news_sale.news_sale_pdf.DbValue = RsRow("news_sale_pdf")
		news_sale.news_sale_title.DbValue = RsRow("news_sale_title")
		news_sale.start_date.DbValue = RsRow("start_date")
		news_sale.end_date.DbValue = RsRow("end_date")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		news_sale.news_sale_id.m_DbValue = Rs("news_sale_id")
		news_sale.news_sale_pdf.m_DbValue = Rs("news_sale_pdf")
		news_sale.news_sale_title.m_DbValue = Rs("news_sale_title")
		news_sale.start_date.m_DbValue = Rs("start_date")
		news_sale.end_date.m_DbValue = Rs("end_date")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = news_sale.AddUrl
		EditUrl = news_sale.EditUrl("")
		CopyUrl = news_sale.CopyUrl("")
		DeleteUrl = news_sale.DeleteUrl
		ListUrl = news_sale.ListUrl
		SetupOtherOptions()

		' Call Row Rendering event
		Call news_sale.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' news_sale_id
		' news_sale_pdf
		' news_sale_title
		' start_date
		' end_date
		' -----------
		'  View  Row
		' -----------

		If news_sale.RowType = EW_ROWTYPE_VIEW Then ' View row

			' news_sale_id
			news_sale.news_sale_id.ViewValue = news_sale.news_sale_id.CurrentValue
			news_sale.news_sale_id.ViewCustomAttributes = ""

			' news_sale_pdf
			news_sale.news_sale_pdf.ViewValue = news_sale.news_sale_pdf.CurrentValue
			news_sale.news_sale_pdf.ViewCustomAttributes = ""

			' news_sale_title
			news_sale.news_sale_title.ViewValue = news_sale.news_sale_title.CurrentValue
			news_sale.news_sale_title.ViewCustomAttributes = ""

			' start_date
			news_sale.start_date.ViewValue = news_sale.start_date.CurrentValue
			news_sale.start_date.ViewCustomAttributes = ""

			' end_date
			news_sale.end_date.ViewValue = news_sale.end_date.CurrentValue
			news_sale.end_date.ViewCustomAttributes = ""

			' View refer script
			' news_sale_id

			news_sale.news_sale_id.LinkCustomAttributes = ""
			news_sale.news_sale_id.HrefValue = ""
			news_sale.news_sale_id.TooltipValue = ""

			' news_sale_pdf
			news_sale.news_sale_pdf.LinkCustomAttributes = ""
			news_sale.news_sale_pdf.HrefValue = ""
			news_sale.news_sale_pdf.TooltipValue = ""

			' news_sale_title
			news_sale.news_sale_title.LinkCustomAttributes = ""
			news_sale.news_sale_title.HrefValue = ""
			news_sale.news_sale_title.TooltipValue = ""

			' start_date
			news_sale.start_date.LinkCustomAttributes = ""
			news_sale.start_date.HrefValue = ""
			news_sale.start_date.TooltipValue = ""

			' end_date
			news_sale.end_date.LinkCustomAttributes = ""
			news_sale.end_date.HrefValue = ""
			news_sale.end_date.TooltipValue = ""
		End If

		' Call Row Rendered event
		If news_sale.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call news_sale.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", news_sale.TableVar, "pom_news_salelist.asp", news_sale.TableVar, True)
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
