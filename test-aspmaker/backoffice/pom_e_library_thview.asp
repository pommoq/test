<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_e_library_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim e_library_th_view
Set e_library_th_view = New ce_library_th_view
Set Page = e_library_th_view

' Page init processing
e_library_th_view.Page_Init()

' Page main processing
e_library_th_view.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
e_library_th_view.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If e_library_th.Export = "" Then %>
<script type="text/javascript">
// Page object
var e_library_th_view = new ew_Page("e_library_th_view");
e_library_th_view.PageID = "view"; // Page ID
var EW_PAGE_ID = e_library_th_view.PageID; // For backward compatibility
// Form object
var fe_library_thview = new ew_Form("fe_library_thview");
// Form_CustomValidate event
fe_library_thview.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fe_library_thview.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fe_library_thview.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If e_library_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If e_library_th.Export = "" Then %>
<div class="ewViewExportOptions">
<% e_library_th_view.ExportOptions.Render "body", "", "", "", "", "" %>
<% If Not e_library_th_view.ExportOptions.UseDropDownButton Then %>
</div>
<div class="ewViewOtherOptions">
<% End If %>
<%
	e_library_th_view.ActionOptions.Render "body", "", "", "", "", ""
	e_library_th_view.DetailOptions.Render "body", "", "", "", "", ""
%>
</div>
<% End If %>
<% e_library_th_view.ShowPageHeader() %>
<% e_library_th_view.ShowMessage %>
<form name="fe_library_thview" id="fe_library_thview" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="e_library_th">
<table class="ewGrid"><tr><td>
<table id="tbl_e_library_thview" class="table table-bordered table-striped">
<% If e_library_th.el_id.Visible Then ' el_id %>
	<tr id="r_el_id">
		<td><span id="elh_e_library_th_el_id"><%= e_library_th.el_id.FldCaption %></span></td>
		<td<%= e_library_th.el_id.CellAttributes %>>
<span id="el_e_library_th_el_id" class="control-group">
<span<%= e_library_th.el_id.ViewAttributes %>>
<%= e_library_th.el_id.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If e_library_th.el_date.Visible Then ' el_date %>
	<tr id="r_el_date">
		<td><span id="elh_e_library_th_el_date"><%= e_library_th.el_date.FldCaption %></span></td>
		<td<%= e_library_th.el_date.CellAttributes %>>
<span id="el_e_library_th_el_date" class="control-group">
<span<%= e_library_th.el_date.ViewAttributes %>>
<%= e_library_th.el_date.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If e_library_th.el_title.Visible Then ' el_title %>
	<tr id="r_el_title">
		<td><span id="elh_e_library_th_el_title"><%= e_library_th.el_title.FldCaption %></span></td>
		<td<%= e_library_th.el_title.CellAttributes %>>
<span id="el_e_library_th_el_title" class="control-group">
<span<%= e_library_th.el_title.ViewAttributes %>>
<%= e_library_th.el_title.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If e_library_th.el_pdf.Visible Then ' el_pdf %>
	<tr id="r_el_pdf">
		<td><span id="elh_e_library_th_el_pdf"><%= e_library_th.el_pdf.FldCaption %></span></td>
		<td<%= e_library_th.el_pdf.CellAttributes %>>
<span id="el_e_library_th_el_pdf" class="control-group">
<span<%= e_library_th.el_pdf.ViewAttributes %>>
<%= e_library_th.el_pdf.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If e_library_th.el_img.Visible Then ' el_img %>
	<tr id="r_el_img">
		<td><span id="elh_e_library_th_el_img"><%= e_library_th.el_img.FldCaption %></span></td>
		<td<%= e_library_th.el_img.CellAttributes %>>
<span id="el_e_library_th_el_img" class="control-group">
<span<%= e_library_th.el_img.ViewAttributes %>>
<%= e_library_th.el_img.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If e_library_th.el_detail.Visible Then ' el_detail %>
	<tr id="r_el_detail">
		<td><span id="elh_e_library_th_el_detail"><%= e_library_th.el_detail.FldCaption %></span></td>
		<td<%= e_library_th.el_detail.CellAttributes %>>
<span id="el_e_library_th_el_detail" class="control-group">
<span<%= e_library_th.el_detail.ViewAttributes %>>
<%= e_library_th.el_detail.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If e_library_th.el_create.Visible Then ' el_create %>
	<tr id="r_el_create">
		<td><span id="elh_e_library_th_el_create"><%= e_library_th.el_create.FldCaption %></span></td>
		<td<%= e_library_th.el_create.CellAttributes %>>
<span id="el_e_library_th_el_create" class="control-group">
<span<%= e_library_th.el_create.ViewAttributes %>>
<%= e_library_th.el_create.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If e_library_th.el_update.Visible Then ' el_update %>
	<tr id="r_el_update">
		<td><span id="elh_e_library_th_el_update"><%= e_library_th.el_update.FldCaption %></span></td>
		<td<%= e_library_th.el_update.CellAttributes %>>
<span id="el_e_library_th_el_update" class="control-group">
<span<%= e_library_th.el_update.ViewAttributes %>>
<%= e_library_th.el_update.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
</table>
</td></tr></table>
</form>
<script type="text/javascript">
fe_library_thview.Init();
</script>
<%
e_library_th_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If e_library_th.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set e_library_th_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class ce_library_th_view

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
		TableName = "e_library_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "e_library_th_view"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If e_library_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & e_library_th.TableVar & "&" ' add page token
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
		If e_library_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (e_library_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (e_library_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(e_library_th) Then Set e_library_th = New ce_library_th
		Set Table = e_library_th

		' Initialize urls
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("el_id").Count > 0 Then
			ew_AddKey RecKey, "el_id", Request.QueryString("el_id")
			KeyUrl = KeyUrl & "&amp;el_id=" & Server.URLEncode(Request.QueryString("el_id"))
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
		EW_TABLE_NAME = "e_library_th"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = e_library_th.TableVar
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
		Set e_library_th = Nothing
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
		If e_library_th.Export = "" Then
			SetupBreadcrumb()
		End If
		If IsPageRequest Then ' Validate request
			If Request.QueryString("el_id").Count > 0 Then
				e_library_th.el_id.QueryStringValue = Request.QueryString("el_id")
			Else
				sReturnUrl = "pom_e_library_thlist.asp" ' Return to list
			End If

			' Get action
			e_library_th.CurrentAction = "I" ' Display form
			Select Case e_library_th.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "pom_e_library_thlist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "pom_e_library_thlist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		e_library_th.RowType = EW_ROWTYPE_VIEW
		Call e_library_th.ResetAttrs()
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
				e_library_th.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					e_library_th.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = e_library_th.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			e_library_th.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			e_library_th.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			e_library_th.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = e_library_th.KeyFilter

		' Call Row Selecting event
		Call e_library_th.Row_Selecting(sFilter)

		' Load sql based on filter
		e_library_th.CurrentFilter = sFilter
		sSql = e_library_th.SQL
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
		Call e_library_th.Row_Selected(RsRow)
		e_library_th.el_id.DbValue = RsRow("el_id")
		e_library_th.el_date.DbValue = RsRow("el_date")
		e_library_th.el_title.DbValue = RsRow("el_title")
		e_library_th.el_pdf.DbValue = RsRow("el_pdf")
		e_library_th.el_img.DbValue = RsRow("el_img")
		e_library_th.el_detail.DbValue = RsRow("el_detail")
		e_library_th.el_create.DbValue = RsRow("el_create")
		e_library_th.el_update.DbValue = RsRow("el_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		e_library_th.el_id.m_DbValue = Rs("el_id")
		e_library_th.el_date.m_DbValue = Rs("el_date")
		e_library_th.el_title.m_DbValue = Rs("el_title")
		e_library_th.el_pdf.m_DbValue = Rs("el_pdf")
		e_library_th.el_img.m_DbValue = Rs("el_img")
		e_library_th.el_detail.m_DbValue = Rs("el_detail")
		e_library_th.el_create.m_DbValue = Rs("el_create")
		e_library_th.el_update.m_DbValue = Rs("el_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = e_library_th.AddUrl
		EditUrl = e_library_th.EditUrl("")
		CopyUrl = e_library_th.CopyUrl("")
		DeleteUrl = e_library_th.DeleteUrl
		ListUrl = e_library_th.ListUrl
		SetupOtherOptions()

		' Call Row Rendering event
		Call e_library_th.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' el_id
		' el_date
		' el_title
		' el_pdf
		' el_img
		' el_detail
		' el_create
		' el_update
		' -----------
		'  View  Row
		' -----------

		If e_library_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' el_id
			e_library_th.el_id.ViewValue = e_library_th.el_id.CurrentValue
			e_library_th.el_id.ViewCustomAttributes = ""

			' el_date
			e_library_th.el_date.ViewValue = e_library_th.el_date.CurrentValue
			e_library_th.el_date.ViewCustomAttributes = ""

			' el_title
			e_library_th.el_title.ViewValue = e_library_th.el_title.CurrentValue
			e_library_th.el_title.ViewCustomAttributes = ""

			' el_pdf
			e_library_th.el_pdf.ViewValue = e_library_th.el_pdf.CurrentValue
			e_library_th.el_pdf.ViewCustomAttributes = ""

			' el_img
			e_library_th.el_img.ViewValue = e_library_th.el_img.CurrentValue
			e_library_th.el_img.ViewCustomAttributes = ""

			' el_detail
			e_library_th.el_detail.ViewValue = e_library_th.el_detail.CurrentValue
			e_library_th.el_detail.ViewCustomAttributes = ""

			' el_create
			e_library_th.el_create.ViewValue = e_library_th.el_create.CurrentValue
			e_library_th.el_create.ViewCustomAttributes = ""

			' el_update
			e_library_th.el_update.ViewValue = e_library_th.el_update.CurrentValue
			e_library_th.el_update.ViewCustomAttributes = ""

			' View refer script
			' el_id

			e_library_th.el_id.LinkCustomAttributes = ""
			e_library_th.el_id.HrefValue = ""
			e_library_th.el_id.TooltipValue = ""

			' el_date
			e_library_th.el_date.LinkCustomAttributes = ""
			e_library_th.el_date.HrefValue = ""
			e_library_th.el_date.TooltipValue = ""

			' el_title
			e_library_th.el_title.LinkCustomAttributes = ""
			e_library_th.el_title.HrefValue = ""
			e_library_th.el_title.TooltipValue = ""

			' el_pdf
			e_library_th.el_pdf.LinkCustomAttributes = ""
			e_library_th.el_pdf.HrefValue = ""
			e_library_th.el_pdf.TooltipValue = ""

			' el_img
			e_library_th.el_img.LinkCustomAttributes = ""
			e_library_th.el_img.HrefValue = ""
			e_library_th.el_img.TooltipValue = ""

			' el_detail
			e_library_th.el_detail.LinkCustomAttributes = ""
			e_library_th.el_detail.HrefValue = ""
			e_library_th.el_detail.TooltipValue = ""

			' el_create
			e_library_th.el_create.LinkCustomAttributes = ""
			e_library_th.el_create.HrefValue = ""
			e_library_th.el_create.TooltipValue = ""

			' el_update
			e_library_th.el_update.LinkCustomAttributes = ""
			e_library_th.el_update.HrefValue = ""
			e_library_th.el_update.TooltipValue = ""
		End If

		' Call Row Rendered event
		If e_library_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call e_library_th.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", e_library_th.TableVar, "pom_e_library_thlist.asp", e_library_th.TableVar, True)
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
