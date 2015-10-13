<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_banner_logo_01info.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim banner_logo_01_view
Set banner_logo_01_view = New cbanner_logo_01_view
Set Page = banner_logo_01_view

' Page init processing
banner_logo_01_view.Page_Init()

' Page main processing
banner_logo_01_view.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
banner_logo_01_view.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If banner_logo_01.Export = "" Then %>
<script type="text/javascript">
// Page object
var banner_logo_01_view = new ew_Page("banner_logo_01_view");
banner_logo_01_view.PageID = "view"; // Page ID
var EW_PAGE_ID = banner_logo_01_view.PageID; // For backward compatibility
// Form object
var fbanner_logo_01view = new ew_Form("fbanner_logo_01view");
// Form_CustomValidate event
fbanner_logo_01view.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fbanner_logo_01view.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fbanner_logo_01view.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If banner_logo_01.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If banner_logo_01.Export = "" Then %>
<div class="ewViewExportOptions">
<% banner_logo_01_view.ExportOptions.Render "body", "", "", "", "", "" %>
<% If Not banner_logo_01_view.ExportOptions.UseDropDownButton Then %>
</div>
<div class="ewViewOtherOptions">
<% End If %>
<%
	banner_logo_01_view.ActionOptions.Render "body", "", "", "", "", ""
	banner_logo_01_view.DetailOptions.Render "body", "", "", "", "", ""
%>
</div>
<% End If %>
<% banner_logo_01_view.ShowPageHeader() %>
<% banner_logo_01_view.ShowMessage %>
<form name="fbanner_logo_01view" id="fbanner_logo_01view" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="banner_logo_01">
<table class="ewGrid"><tr><td>
<table id="tbl_banner_logo_01view" class="table table-bordered table-striped">
<% If banner_logo_01.banner_id.Visible Then ' banner_id %>
	<tr id="r_banner_id">
		<td><span id="elh_banner_logo_01_banner_id"><%= banner_logo_01.banner_id.FldCaption %></span></td>
		<td<%= banner_logo_01.banner_id.CellAttributes %>>
<span id="el_banner_logo_01_banner_id" class="control-group">
<span<%= banner_logo_01.banner_id.ViewAttributes %>>
<%= banner_logo_01.banner_id.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If banner_logo_01.banner_img.Visible Then ' banner_img %>
	<tr id="r_banner_img">
		<td><span id="elh_banner_logo_01_banner_img"><%= banner_logo_01.banner_img.FldCaption %></span></td>
		<td<%= banner_logo_01.banner_img.CellAttributes %>>
<span id="el_banner_logo_01_banner_img" class="control-group">
<span<%= banner_logo_01.banner_img.ViewAttributes %>>
<%= banner_logo_01.banner_img.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If banner_logo_01.banner_link.Visible Then ' banner_link %>
	<tr id="r_banner_link">
		<td><span id="elh_banner_logo_01_banner_link"><%= banner_logo_01.banner_link.FldCaption %></span></td>
		<td<%= banner_logo_01.banner_link.CellAttributes %>>
<span id="el_banner_logo_01_banner_link" class="control-group">
<span<%= banner_logo_01.banner_link.ViewAttributes %>>
<%= banner_logo_01.banner_link.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If banner_logo_01.banner_sort.Visible Then ' banner_sort %>
	<tr id="r_banner_sort">
		<td><span id="elh_banner_logo_01_banner_sort"><%= banner_logo_01.banner_sort.FldCaption %></span></td>
		<td<%= banner_logo_01.banner_sort.CellAttributes %>>
<span id="el_banner_logo_01_banner_sort" class="control-group">
<span<%= banner_logo_01.banner_sort.ViewAttributes %>>
<%= banner_logo_01.banner_sort.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If banner_logo_01.start_date.Visible Then ' start_date %>
	<tr id="r_start_date">
		<td><span id="elh_banner_logo_01_start_date"><%= banner_logo_01.start_date.FldCaption %></span></td>
		<td<%= banner_logo_01.start_date.CellAttributes %>>
<span id="el_banner_logo_01_start_date" class="control-group">
<span<%= banner_logo_01.start_date.ViewAttributes %>>
<%= banner_logo_01.start_date.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If banner_logo_01.end_date.Visible Then ' end_date %>
	<tr id="r_end_date">
		<td><span id="elh_banner_logo_01_end_date"><%= banner_logo_01.end_date.FldCaption %></span></td>
		<td<%= banner_logo_01.end_date.CellAttributes %>>
<span id="el_banner_logo_01_end_date" class="control-group">
<span<%= banner_logo_01.end_date.ViewAttributes %>>
<%= banner_logo_01.end_date.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
</table>
</td></tr></table>
</form>
<script type="text/javascript">
fbanner_logo_01view.Init();
</script>
<%
banner_logo_01_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If banner_logo_01.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set banner_logo_01_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cbanner_logo_01_view

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
		TableName = "banner_logo_01"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "banner_logo_01_view"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If banner_logo_01.UseTokenInUrl Then PageUrl = PageUrl & "t=" & banner_logo_01.TableVar & "&" ' add page token
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
		If banner_logo_01.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (banner_logo_01.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (banner_logo_01.TableVar = Request.QueryString("t"))
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
		If IsEmpty(banner_logo_01) Then Set banner_logo_01 = New cbanner_logo_01
		Set Table = banner_logo_01

		' Initialize urls
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("banner_id").Count > 0 Then
			ew_AddKey RecKey, "banner_id", Request.QueryString("banner_id")
			KeyUrl = KeyUrl & "&amp;banner_id=" & Server.URLEncode(Request.QueryString("banner_id"))
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
		EW_TABLE_NAME = "banner_logo_01"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = banner_logo_01.TableVar
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
		Set banner_logo_01 = Nothing
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
		If banner_logo_01.Export = "" Then
			SetupBreadcrumb()
		End If
		If IsPageRequest Then ' Validate request
			If Request.QueryString("banner_id").Count > 0 Then
				banner_logo_01.banner_id.QueryStringValue = Request.QueryString("banner_id")
			Else
				sReturnUrl = "pom_banner_logo_01list.asp" ' Return to list
			End If

			' Get action
			banner_logo_01.CurrentAction = "I" ' Display form
			Select Case banner_logo_01.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "pom_banner_logo_01list.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "pom_banner_logo_01list.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		banner_logo_01.RowType = EW_ROWTYPE_VIEW
		Call banner_logo_01.ResetAttrs()
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
				banner_logo_01.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					banner_logo_01.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = banner_logo_01.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			banner_logo_01.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			banner_logo_01.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			banner_logo_01.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = banner_logo_01.KeyFilter

		' Call Row Selecting event
		Call banner_logo_01.Row_Selecting(sFilter)

		' Load sql based on filter
		banner_logo_01.CurrentFilter = sFilter
		sSql = banner_logo_01.SQL
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
		Call banner_logo_01.Row_Selected(RsRow)
		banner_logo_01.banner_id.DbValue = RsRow("banner_id")
		banner_logo_01.banner_img.DbValue = RsRow("banner_img")
		banner_logo_01.banner_link.DbValue = RsRow("banner_link")
		banner_logo_01.banner_sort.DbValue = RsRow("banner_sort")
		banner_logo_01.start_date.DbValue = RsRow("start_date")
		banner_logo_01.end_date.DbValue = RsRow("end_date")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		banner_logo_01.banner_id.m_DbValue = Rs("banner_id")
		banner_logo_01.banner_img.m_DbValue = Rs("banner_img")
		banner_logo_01.banner_link.m_DbValue = Rs("banner_link")
		banner_logo_01.banner_sort.m_DbValue = Rs("banner_sort")
		banner_logo_01.start_date.m_DbValue = Rs("start_date")
		banner_logo_01.end_date.m_DbValue = Rs("end_date")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = banner_logo_01.AddUrl
		EditUrl = banner_logo_01.EditUrl("")
		CopyUrl = banner_logo_01.CopyUrl("")
		DeleteUrl = banner_logo_01.DeleteUrl
		ListUrl = banner_logo_01.ListUrl
		SetupOtherOptions()

		' Call Row Rendering event
		Call banner_logo_01.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' banner_id
		' banner_img
		' banner_link
		' banner_sort
		' start_date
		' end_date
		' -----------
		'  View  Row
		' -----------

		If banner_logo_01.RowType = EW_ROWTYPE_VIEW Then ' View row

			' banner_id
			banner_logo_01.banner_id.ViewValue = banner_logo_01.banner_id.CurrentValue
			banner_logo_01.banner_id.ViewCustomAttributes = ""

			' banner_img
			banner_logo_01.banner_img.ViewValue = banner_logo_01.banner_img.CurrentValue
			banner_logo_01.banner_img.ViewCustomAttributes = ""

			' banner_link
			banner_logo_01.banner_link.ViewValue = banner_logo_01.banner_link.CurrentValue
			banner_logo_01.banner_link.ViewCustomAttributes = ""

			' banner_sort
			banner_logo_01.banner_sort.ViewValue = banner_logo_01.banner_sort.CurrentValue
			banner_logo_01.banner_sort.ViewCustomAttributes = ""

			' start_date
			banner_logo_01.start_date.ViewValue = banner_logo_01.start_date.CurrentValue
			banner_logo_01.start_date.ViewCustomAttributes = ""

			' end_date
			banner_logo_01.end_date.ViewValue = banner_logo_01.end_date.CurrentValue
			banner_logo_01.end_date.ViewCustomAttributes = ""

			' View refer script
			' banner_id

			banner_logo_01.banner_id.LinkCustomAttributes = ""
			banner_logo_01.banner_id.HrefValue = ""
			banner_logo_01.banner_id.TooltipValue = ""

			' banner_img
			banner_logo_01.banner_img.LinkCustomAttributes = ""
			banner_logo_01.banner_img.HrefValue = ""
			banner_logo_01.banner_img.TooltipValue = ""

			' banner_link
			banner_logo_01.banner_link.LinkCustomAttributes = ""
			banner_logo_01.banner_link.HrefValue = ""
			banner_logo_01.banner_link.TooltipValue = ""

			' banner_sort
			banner_logo_01.banner_sort.LinkCustomAttributes = ""
			banner_logo_01.banner_sort.HrefValue = ""
			banner_logo_01.banner_sort.TooltipValue = ""

			' start_date
			banner_logo_01.start_date.LinkCustomAttributes = ""
			banner_logo_01.start_date.HrefValue = ""
			banner_logo_01.start_date.TooltipValue = ""

			' end_date
			banner_logo_01.end_date.LinkCustomAttributes = ""
			banner_logo_01.end_date.HrefValue = ""
			banner_logo_01.end_date.TooltipValue = ""
		End If

		' Call Row Rendered event
		If banner_logo_01.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call banner_logo_01.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", banner_logo_01.TableVar, "pom_banner_logo_01list.asp", banner_logo_01.TableVar, True)
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
