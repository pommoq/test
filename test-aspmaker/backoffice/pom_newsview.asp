<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_newsinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim news_view
Set news_view = New cnews_view
Set Page = news_view

' Page init processing
news_view.Page_Init()

' Page main processing
news_view.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
news_view.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If news.Export = "" Then %>
<script type="text/javascript">
// Page object
var news_view = new ew_Page("news_view");
news_view.PageID = "view"; // Page ID
var EW_PAGE_ID = news_view.PageID; // For backward compatibility
// Form object
var fnewsview = new ew_Form("fnewsview");
// Form_CustomValidate event
fnewsview.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fnewsview.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fnewsview.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If news.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If news.Export = "" Then %>
<div class="ewViewExportOptions">
<% news_view.ExportOptions.Render "body", "", "", "", "", "" %>
<% If Not news_view.ExportOptions.UseDropDownButton Then %>
</div>
<div class="ewViewOtherOptions">
<% End If %>
<%
	news_view.ActionOptions.Render "body", "", "", "", "", ""
	news_view.DetailOptions.Render "body", "", "", "", "", ""
%>
</div>
<% End If %>
<% news_view.ShowPageHeader() %>
<% news_view.ShowMessage %>
<form name="fnewsview" id="fnewsview" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="news">
<table class="ewGrid"><tr><td>
<table id="tbl_newsview" class="table table-bordered table-striped">
<% If news.news_id.Visible Then ' news_id %>
	<tr id="r_news_id">
		<td><span id="elh_news_news_id"><%= news.news_id.FldCaption %></span></td>
		<td<%= news.news_id.CellAttributes %>>
<span id="el_news_news_id" class="control-group">
<span<%= news.news_id.ViewAttributes %>>
<%= news.news_id.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.news_img.Visible Then ' news_img %>
	<tr id="r_news_img">
		<td><span id="elh_news_news_img"><%= news.news_img.FldCaption %></span></td>
		<td<%= news.news_img.CellAttributes %>>
<span id="el_news_news_img" class="control-group">
<span>
<%= ew_GetFileViewTag(news.news_img, news.news_img.ViewValue) %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.news_date.Visible Then ' news_date %>
	<tr id="r_news_date">
		<td><span id="elh_news_news_date"><%= news.news_date.FldCaption %></span></td>
		<td<%= news.news_date.CellAttributes %>>
<span id="el_news_news_date" class="control-group">
<span<%= news.news_date.ViewAttributes %>>
<%= news.news_date.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.news_category.Visible Then ' news_category %>
	<tr id="r_news_category">
		<td><span id="elh_news_news_category"><%= news.news_category.FldCaption %></span></td>
		<td<%= news.news_category.CellAttributes %>>
<span id="el_news_news_category" class="control-group">
<span<%= news.news_category.ViewAttributes %>>
<%= news.news_category.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.news_category_sub.Visible Then ' news_category_sub %>
	<tr id="r_news_category_sub">
		<td><span id="elh_news_news_category_sub"><%= news.news_category_sub.FldCaption %></span></td>
		<td<%= news.news_category_sub.CellAttributes %>>
<span id="el_news_news_category_sub" class="control-group">
<span<%= news.news_category_sub.ViewAttributes %>>
<%= news.news_category_sub.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.start_date.Visible Then ' start_date %>
	<tr id="r_start_date">
		<td><span id="elh_news_start_date"><%= news.start_date.FldCaption %></span></td>
		<td<%= news.start_date.CellAttributes %>>
<span id="el_news_start_date" class="control-group">
<span<%= news.start_date.ViewAttributes %>>
<%= news.start_date.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.end_date.Visible Then ' end_date %>
	<tr id="r_end_date">
		<td><span id="elh_news_end_date"><%= news.end_date.FldCaption %></span></td>
		<td<%= news.end_date.CellAttributes %>>
<span id="el_news_end_date" class="control-group">
<span<%= news.end_date.ViewAttributes %>>
<%= news.end_date.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.news_pdf.Visible Then ' news_pdf %>
	<tr id="r_news_pdf">
		<td><span id="elh_news_news_pdf"><%= news.news_pdf.FldCaption %></span></td>
		<td<%= news.news_pdf.CellAttributes %>>
<span id="el_news_news_pdf" class="control-group">
<span<%= news.news_pdf.ViewAttributes %>>
<%= news.news_pdf.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.news_subject.Visible Then ' news_subject %>
	<tr id="r_news_subject">
		<td><span id="elh_news_news_subject"><%= news.news_subject.FldCaption %></span></td>
		<td<%= news.news_subject.CellAttributes %>>
<span id="el_news_news_subject" class="control-group">
<span<%= news.news_subject.ViewAttributes %>>
<%= news.news_subject.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.news_subject_th.Visible Then ' news_subject_th %>
	<tr id="r_news_subject_th">
		<td><span id="elh_news_news_subject_th"><%= news.news_subject_th.FldCaption %></span></td>
		<td<%= news.news_subject_th.CellAttributes %>>
<span id="el_news_news_subject_th" class="control-group">
<span<%= news.news_subject_th.ViewAttributes %>>
<%= news.news_subject_th.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.news_intro.Visible Then ' news_intro %>
	<tr id="r_news_intro">
		<td><span id="elh_news_news_intro"><%= news.news_intro.FldCaption %></span></td>
		<td<%= news.news_intro.CellAttributes %>>
<span id="el_news_news_intro" class="control-group">
<span<%= news.news_intro.ViewAttributes %>>
<%= news.news_intro.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.news_intro_th.Visible Then ' news_intro_th %>
	<tr id="r_news_intro_th">
		<td><span id="elh_news_news_intro_th"><%= news.news_intro_th.FldCaption %></span></td>
		<td<%= news.news_intro_th.CellAttributes %>>
<span id="el_news_news_intro_th" class="control-group">
<span<%= news.news_intro_th.ViewAttributes %>>
<%= news.news_intro_th.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.news_content.Visible Then ' news_content %>
	<tr id="r_news_content">
		<td><span id="elh_news_news_content"><%= news.news_content.FldCaption %></span></td>
		<td<%= news.news_content.CellAttributes %>>
<span id="el_news_news_content" class="control-group">
<span<%= news.news_content.ViewAttributes %>>
<%= news.news_content.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.news_content_th.Visible Then ' news_content_th %>
	<tr id="r_news_content_th">
		<td><span id="elh_news_news_content_th"><%= news.news_content_th.FldCaption %></span></td>
		<td<%= news.news_content_th.CellAttributes %>>
<span id="el_news_news_content_th" class="control-group">
<span<%= news.news_content_th.ViewAttributes %>>
<%= news.news_content_th.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.news_show_en.Visible Then ' news_show_en %>
	<tr id="r_news_show_en">
		<td><span id="elh_news_news_show_en"><%= news.news_show_en.FldCaption %></span></td>
		<td<%= news.news_show_en.CellAttributes %>>
<span id="el_news_news_show_en" class="control-group">
<span<%= news.news_show_en.ViewAttributes %>>
<%= news.news_show_en.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.news_show.Visible Then ' news_show %>
	<tr id="r_news_show">
		<td><span id="elh_news_news_show"><%= news.news_show.FldCaption %></span></td>
		<td<%= news.news_show.CellAttributes %>>
<span id="el_news_news_show" class="control-group">
<span<%= news.news_show.ViewAttributes %>>
<%= news.news_show.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.news_show_home.Visible Then ' news_show_home %>
	<tr id="r_news_show_home">
		<td><span id="elh_news_news_show_home"><%= news.news_show_home.FldCaption %></span></td>
		<td<%= news.news_show_home.CellAttributes %>>
<span id="el_news_news_show_home" class="control-group">
<span<%= news.news_show_home.ViewAttributes %>>
<%= news.news_show_home.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.news_create.Visible Then ' news_create %>
	<tr id="r_news_create">
		<td><span id="elh_news_news_create"><%= news.news_create.FldCaption %></span></td>
		<td<%= news.news_create.CellAttributes %>>
<span id="el_news_news_create" class="control-group">
<span<%= news.news_create.ViewAttributes %>>
<%= news.news_create.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If news.news_update.Visible Then ' news_update %>
	<tr id="r_news_update">
		<td><span id="elh_news_news_update"><%= news.news_update.FldCaption %></span></td>
		<td<%= news.news_update.CellAttributes %>>
<span id="el_news_news_update" class="control-group">
<span<%= news.news_update.ViewAttributes %>>
<%= news.news_update.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
</table>
</td></tr></table>
</form>
<script type="text/javascript">
fnewsview.Init();
</script>
<%
news_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If news.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set news_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cnews_view

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
		TableName = "news"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "news_view"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If news.UseTokenInUrl Then PageUrl = PageUrl & "t=" & news.TableVar & "&" ' add page token
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
		If news.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (news.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (news.TableVar = Request.QueryString("t"))
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
		If IsEmpty(news) Then Set news = New cnews
		Set Table = news

		' Initialize urls
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("news_id").Count > 0 Then
			ew_AddKey RecKey, "news_id", Request.QueryString("news_id")
			KeyUrl = KeyUrl & "&amp;news_id=" & Server.URLEncode(Request.QueryString("news_id"))
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
		EW_TABLE_NAME = "news"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = news.TableVar
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
		Set news = Nothing
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
		If news.Export = "" Then
			SetupBreadcrumb()
		End If
		If IsPageRequest Then ' Validate request
			If Request.QueryString("news_id").Count > 0 Then
				news.news_id.QueryStringValue = Request.QueryString("news_id")
			Else
				sReturnUrl = "pom_newslist.asp" ' Return to list
			End If

			' Get action
			news.CurrentAction = "I" ' Display form
			Select Case news.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "pom_newslist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "pom_newslist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		news.RowType = EW_ROWTYPE_VIEW
		Call news.ResetAttrs()
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
				news.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					news.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = news.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			news.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			news.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			news.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = news.KeyFilter

		' Call Row Selecting event
		Call news.Row_Selecting(sFilter)

		' Load sql based on filter
		news.CurrentFilter = sFilter
		sSql = news.SQL
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
		Call news.Row_Selected(RsRow)
		news.news_id.DbValue = RsRow("news_id")
		news.news_img.Upload.DbValue = RsRow("news_img")
		news.news_img.CurrentValue = news.news_img.Upload.DbValue
		news.news_date.DbValue = RsRow("news_date")
		news.news_category.DbValue = RsRow("news_category")
		news.news_category_sub.DbValue = RsRow("news_category_sub")
		news.start_date.DbValue = RsRow("start_date")
		news.end_date.DbValue = RsRow("end_date")
		news.news_pdf.DbValue = RsRow("news_pdf")
		news.news_subject.DbValue = RsRow("news_subject")
		news.news_subject_th.DbValue = RsRow("news_subject_th")
		news.news_intro.DbValue = RsRow("news_intro")
		news.news_intro_th.DbValue = RsRow("news_intro_th")
		news.news_content.DbValue = RsRow("news_content")
		news.news_content_th.DbValue = RsRow("news_content_th")
		news.news_show_en.DbValue = RsRow("news_show_en")
		news.news_show.DbValue = RsRow("news_show")
		news.news_show_home.DbValue = RsRow("news_show_home")
		news.news_create.DbValue = RsRow("news_create")
		news.news_update.DbValue = RsRow("news_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		news.news_id.m_DbValue = Rs("news_id")
		news.news_img.Upload.DbValue = Rs("news_img")
		news.news_date.m_DbValue = Rs("news_date")
		news.news_category.m_DbValue = Rs("news_category")
		news.news_category_sub.m_DbValue = Rs("news_category_sub")
		news.start_date.m_DbValue = Rs("start_date")
		news.end_date.m_DbValue = Rs("end_date")
		news.news_pdf.m_DbValue = Rs("news_pdf")
		news.news_subject.m_DbValue = Rs("news_subject")
		news.news_subject_th.m_DbValue = Rs("news_subject_th")
		news.news_intro.m_DbValue = Rs("news_intro")
		news.news_intro_th.m_DbValue = Rs("news_intro_th")
		news.news_content.m_DbValue = Rs("news_content")
		news.news_content_th.m_DbValue = Rs("news_content_th")
		news.news_show_en.m_DbValue = Rs("news_show_en")
		news.news_show.m_DbValue = Rs("news_show")
		news.news_show_home.m_DbValue = Rs("news_show_home")
		news.news_create.m_DbValue = Rs("news_create")
		news.news_update.m_DbValue = Rs("news_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = news.AddUrl
		EditUrl = news.EditUrl("")
		CopyUrl = news.CopyUrl("")
		DeleteUrl = news.DeleteUrl
		ListUrl = news.ListUrl
		SetupOtherOptions()

		' Call Row Rendering event
		Call news.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' news_id
		' news_img
		' news_date
		' news_category
		' news_category_sub
		' start_date
		' end_date
		' news_pdf
		' news_subject
		' news_subject_th
		' news_intro
		' news_intro_th
		' news_content
		' news_content_th
		' news_show_en
		' news_show
		' news_show_home
		' news_create
		' news_update
		' -----------
		'  View  Row
		' -----------

		If news.RowType = EW_ROWTYPE_VIEW Then ' View row

			' news_id
			news.news_id.ViewValue = news.news_id.CurrentValue
			news.news_id.ViewCustomAttributes = ""

			' news_img
			news.news_img.UploadPath = "./Upload/news"
			If Not ew_Empty(news.news_img.Upload.DbValue) Then
				news.news_img.ViewValue = news.news_img.Upload.DbValue
				news.news_img.ImageAlt = news.news_img.FldAlt
				news.news_img.ViewValue = ew_UploadPathEx(False, news.news_img.UploadPath) & news.news_img.Upload.DbValue
			Else
				news.news_img.ViewValue = ""
			End If
			news.news_img.ViewCustomAttributes = ""

			' news_date
			news.news_date.ViewValue = news.news_date.CurrentValue
			news.news_date.ViewCustomAttributes = ""

			' news_category
			news.news_category.ViewValue = news.news_category.CurrentValue
			news.news_category.ViewCustomAttributes = ""

			' news_category_sub
			news.news_category_sub.ViewValue = news.news_category_sub.CurrentValue
			news.news_category_sub.ViewCustomAttributes = ""

			' start_date
			news.start_date.ViewValue = news.start_date.CurrentValue
			news.start_date.ViewCustomAttributes = ""

			' end_date
			news.end_date.ViewValue = news.end_date.CurrentValue
			news.end_date.ViewCustomAttributes = ""

			' news_pdf
			news.news_pdf.ViewValue = news.news_pdf.CurrentValue
			news.news_pdf.ViewCustomAttributes = ""

			' news_subject
			news.news_subject.ViewValue = news.news_subject.CurrentValue
			news.news_subject.ViewCustomAttributes = ""

			' news_subject_th
			news.news_subject_th.ViewValue = news.news_subject_th.CurrentValue
			news.news_subject_th.ViewCustomAttributes = ""

			' news_intro
			news.news_intro.ViewValue = news.news_intro.CurrentValue
			news.news_intro.ViewCustomAttributes = ""

			' news_intro_th
			news.news_intro_th.ViewValue = news.news_intro_th.CurrentValue
			news.news_intro_th.ViewCustomAttributes = ""

			' news_content
			news.news_content.ViewValue = news.news_content.CurrentValue
			news.news_content.ViewCustomAttributes = ""

			' news_content_th
			news.news_content_th.ViewValue = news.news_content_th.CurrentValue
			news.news_content_th.ViewCustomAttributes = ""

			' news_show_en
			news.news_show_en.ViewValue = news.news_show_en.CurrentValue
			news.news_show_en.ViewCustomAttributes = ""

			' news_show
			news.news_show.ViewValue = news.news_show.CurrentValue
			news.news_show.ViewCustomAttributes = ""

			' news_show_home
			news.news_show_home.ViewValue = news.news_show_home.CurrentValue
			news.news_show_home.ViewCustomAttributes = ""

			' news_create
			news.news_create.ViewValue = news.news_create.CurrentValue
			news.news_create.ViewCustomAttributes = ""

			' news_update
			news.news_update.ViewValue = news.news_update.CurrentValue
			news.news_update.ViewCustomAttributes = ""

			' View refer script
			' news_id

			news.news_id.LinkCustomAttributes = ""
			news.news_id.HrefValue = ""
			news.news_id.TooltipValue = ""

			' news_img
			news.news_img.LinkCustomAttributes = ""
			news.news_img.HrefValue = ""
			news.news_img.HrefValue2 = news.news_img.UploadPath & news.news_img.Upload.DbValue
			news.news_img.TooltipValue = ""

			' news_date
			news.news_date.LinkCustomAttributes = ""
			news.news_date.HrefValue = ""
			news.news_date.TooltipValue = ""

			' news_category
			news.news_category.LinkCustomAttributes = ""
			news.news_category.HrefValue = ""
			news.news_category.TooltipValue = ""

			' news_category_sub
			news.news_category_sub.LinkCustomAttributes = ""
			news.news_category_sub.HrefValue = ""
			news.news_category_sub.TooltipValue = ""

			' start_date
			news.start_date.LinkCustomAttributes = ""
			news.start_date.HrefValue = ""
			news.start_date.TooltipValue = ""

			' end_date
			news.end_date.LinkCustomAttributes = ""
			news.end_date.HrefValue = ""
			news.end_date.TooltipValue = ""

			' news_pdf
			news.news_pdf.LinkCustomAttributes = ""
			news.news_pdf.HrefValue = ""
			news.news_pdf.TooltipValue = ""

			' news_subject
			news.news_subject.LinkCustomAttributes = ""
			news.news_subject.HrefValue = ""
			news.news_subject.TooltipValue = ""

			' news_subject_th
			news.news_subject_th.LinkCustomAttributes = ""
			news.news_subject_th.HrefValue = ""
			news.news_subject_th.TooltipValue = ""

			' news_intro
			news.news_intro.LinkCustomAttributes = ""
			news.news_intro.HrefValue = ""
			news.news_intro.TooltipValue = ""

			' news_intro_th
			news.news_intro_th.LinkCustomAttributes = ""
			news.news_intro_th.HrefValue = ""
			news.news_intro_th.TooltipValue = ""

			' news_content
			news.news_content.LinkCustomAttributes = ""
			news.news_content.HrefValue = ""
			news.news_content.TooltipValue = ""

			' news_content_th
			news.news_content_th.LinkCustomAttributes = ""
			news.news_content_th.HrefValue = ""
			news.news_content_th.TooltipValue = ""

			' news_show_en
			news.news_show_en.LinkCustomAttributes = ""
			news.news_show_en.HrefValue = ""
			news.news_show_en.TooltipValue = ""

			' news_show
			news.news_show.LinkCustomAttributes = ""
			news.news_show.HrefValue = ""
			news.news_show.TooltipValue = ""

			' news_show_home
			news.news_show_home.LinkCustomAttributes = ""
			news.news_show_home.HrefValue = ""
			news.news_show_home.TooltipValue = ""

			' news_create
			news.news_create.LinkCustomAttributes = ""
			news.news_create.HrefValue = ""
			news.news_create.TooltipValue = ""

			' news_update
			news.news_update.LinkCustomAttributes = ""
			news.news_update.HrefValue = ""
			news.news_update.TooltipValue = ""
		End If

		' Call Row Rendered event
		If news.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call news.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", news.TableVar, "pom_newslist.asp", news.TableVar, True)
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
