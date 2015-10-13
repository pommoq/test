<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_research_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim research_th_view
Set research_th_view = New cresearch_th_view
Set Page = research_th_view

' Page init processing
research_th_view.Page_Init()

' Page main processing
research_th_view.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
research_th_view.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If research_th.Export = "" Then %>
<script type="text/javascript">
// Page object
var research_th_view = new ew_Page("research_th_view");
research_th_view.PageID = "view"; // Page ID
var EW_PAGE_ID = research_th_view.PageID; // For backward compatibility
// Form object
var fresearch_thview = new ew_Form("fresearch_thview");
// Form_CustomValidate event
fresearch_thview.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fresearch_thview.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fresearch_thview.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If research_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If research_th.Export = "" Then %>
<div class="ewViewExportOptions">
<% research_th_view.ExportOptions.Render "body", "", "", "", "", "" %>
<% If Not research_th_view.ExportOptions.UseDropDownButton Then %>
</div>
<div class="ewViewOtherOptions">
<% End If %>
<%
	research_th_view.ActionOptions.Render "body", "", "", "", "", ""
	research_th_view.DetailOptions.Render "body", "", "", "", "", ""
%>
</div>
<% End If %>
<% research_th_view.ShowPageHeader() %>
<% research_th_view.ShowMessage %>
<form name="fresearch_thview" id="fresearch_thview" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="research_th">
<table class="ewGrid"><tr><td>
<table id="tbl_research_thview" class="table table-bordered table-striped">
<% If research_th.rsh_id.Visible Then ' rsh_id %>
	<tr id="r_rsh_id">
		<td><span id="elh_research_th_rsh_id"><%= research_th.rsh_id.FldCaption %></span></td>
		<td<%= research_th.rsh_id.CellAttributes %>>
<span id="el_research_th_rsh_id" class="control-group">
<span<%= research_th.rsh_id.ViewAttributes %>>
<%= research_th.rsh_id.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If research_th.rsh_img.Visible Then ' rsh_img %>
	<tr id="r_rsh_img">
		<td><span id="elh_research_th_rsh_img"><%= research_th.rsh_img.FldCaption %></span></td>
		<td<%= research_th.rsh_img.CellAttributes %>>
<span id="el_research_th_rsh_img" class="control-group">
<span<%= research_th.rsh_img.ViewAttributes %>>
<%= research_th.rsh_img.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If research_th.rsh_date.Visible Then ' rsh_date %>
	<tr id="r_rsh_date">
		<td><span id="elh_research_th_rsh_date"><%= research_th.rsh_date.FldCaption %></span></td>
		<td<%= research_th.rsh_date.CellAttributes %>>
<span id="el_research_th_rsh_date" class="control-group">
<span<%= research_th.rsh_date.ViewAttributes %>>
<%= research_th.rsh_date.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If research_th.rsh_pdf.Visible Then ' rsh_pdf %>
	<tr id="r_rsh_pdf">
		<td><span id="elh_research_th_rsh_pdf"><%= research_th.rsh_pdf.FldCaption %></span></td>
		<td<%= research_th.rsh_pdf.CellAttributes %>>
<span id="el_research_th_rsh_pdf" class="control-group">
<span<%= research_th.rsh_pdf.ViewAttributes %>>
<%= research_th.rsh_pdf.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If research_th.rsh_category.Visible Then ' rsh_category %>
	<tr id="r_rsh_category">
		<td><span id="elh_research_th_rsh_category"><%= research_th.rsh_category.FldCaption %></span></td>
		<td<%= research_th.rsh_category.CellAttributes %>>
<span id="el_research_th_rsh_category" class="control-group">
<span<%= research_th.rsh_category.ViewAttributes %>>
<%= research_th.rsh_category.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If research_th.rsh_subject.Visible Then ' rsh_subject %>
	<tr id="r_rsh_subject">
		<td><span id="elh_research_th_rsh_subject"><%= research_th.rsh_subject.FldCaption %></span></td>
		<td<%= research_th.rsh_subject.CellAttributes %>>
<span id="el_research_th_rsh_subject" class="control-group">
<span<%= research_th.rsh_subject.ViewAttributes %>>
<%= research_th.rsh_subject.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If research_th.rsh_subject_th.Visible Then ' rsh_subject_th %>
	<tr id="r_rsh_subject_th">
		<td><span id="elh_research_th_rsh_subject_th"><%= research_th.rsh_subject_th.FldCaption %></span></td>
		<td<%= research_th.rsh_subject_th.CellAttributes %>>
<span id="el_research_th_rsh_subject_th" class="control-group">
<span<%= research_th.rsh_subject_th.ViewAttributes %>>
<%= research_th.rsh_subject_th.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If research_th.rsh_intro.Visible Then ' rsh_intro %>
	<tr id="r_rsh_intro">
		<td><span id="elh_research_th_rsh_intro"><%= research_th.rsh_intro.FldCaption %></span></td>
		<td<%= research_th.rsh_intro.CellAttributes %>>
<span id="el_research_th_rsh_intro" class="control-group">
<span<%= research_th.rsh_intro.ViewAttributes %>>
<%= research_th.rsh_intro.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If research_th.rsh_intro_th.Visible Then ' rsh_intro_th %>
	<tr id="r_rsh_intro_th">
		<td><span id="elh_research_th_rsh_intro_th"><%= research_th.rsh_intro_th.FldCaption %></span></td>
		<td<%= research_th.rsh_intro_th.CellAttributes %>>
<span id="el_research_th_rsh_intro_th" class="control-group">
<span<%= research_th.rsh_intro_th.ViewAttributes %>>
<%= research_th.rsh_intro_th.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If research_th.rsh_content.Visible Then ' rsh_content %>
	<tr id="r_rsh_content">
		<td><span id="elh_research_th_rsh_content"><%= research_th.rsh_content.FldCaption %></span></td>
		<td<%= research_th.rsh_content.CellAttributes %>>
<span id="el_research_th_rsh_content" class="control-group">
<span<%= research_th.rsh_content.ViewAttributes %>>
<%= research_th.rsh_content.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If research_th.rsh_content_th.Visible Then ' rsh_content_th %>
	<tr id="r_rsh_content_th">
		<td><span id="elh_research_th_rsh_content_th"><%= research_th.rsh_content_th.FldCaption %></span></td>
		<td<%= research_th.rsh_content_th.CellAttributes %>>
<span id="el_research_th_rsh_content_th" class="control-group">
<span<%= research_th.rsh_content_th.ViewAttributes %>>
<%= research_th.rsh_content_th.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If research_th.rsh_show.Visible Then ' rsh_show %>
	<tr id="r_rsh_show">
		<td><span id="elh_research_th_rsh_show"><%= research_th.rsh_show.FldCaption %></span></td>
		<td<%= research_th.rsh_show.CellAttributes %>>
<span id="el_research_th_rsh_show" class="control-group">
<span<%= research_th.rsh_show.ViewAttributes %>>
<%= research_th.rsh_show.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If research_th.rsh_show_home.Visible Then ' rsh_show_home %>
	<tr id="r_rsh_show_home">
		<td><span id="elh_research_th_rsh_show_home"><%= research_th.rsh_show_home.FldCaption %></span></td>
		<td<%= research_th.rsh_show_home.CellAttributes %>>
<span id="el_research_th_rsh_show_home" class="control-group">
<span<%= research_th.rsh_show_home.ViewAttributes %>>
<%= research_th.rsh_show_home.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If research_th.rsh_create.Visible Then ' rsh_create %>
	<tr id="r_rsh_create">
		<td><span id="elh_research_th_rsh_create"><%= research_th.rsh_create.FldCaption %></span></td>
		<td<%= research_th.rsh_create.CellAttributes %>>
<span id="el_research_th_rsh_create" class="control-group">
<span<%= research_th.rsh_create.ViewAttributes %>>
<%= research_th.rsh_create.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If research_th.rsh_update.Visible Then ' rsh_update %>
	<tr id="r_rsh_update">
		<td><span id="elh_research_th_rsh_update"><%= research_th.rsh_update.FldCaption %></span></td>
		<td<%= research_th.rsh_update.CellAttributes %>>
<span id="el_research_th_rsh_update" class="control-group">
<span<%= research_th.rsh_update.ViewAttributes %>>
<%= research_th.rsh_update.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
</table>
</td></tr></table>
</form>
<script type="text/javascript">
fresearch_thview.Init();
</script>
<%
research_th_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If research_th.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set research_th_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cresearch_th_view

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
		TableName = "research_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "research_th_view"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If research_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & research_th.TableVar & "&" ' add page token
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
		If research_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (research_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (research_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(research_th) Then Set research_th = New cresearch_th
		Set Table = research_th

		' Initialize urls
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("rsh_id").Count > 0 Then
			ew_AddKey RecKey, "rsh_id", Request.QueryString("rsh_id")
			KeyUrl = KeyUrl & "&amp;rsh_id=" & Server.URLEncode(Request.QueryString("rsh_id"))
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
		EW_TABLE_NAME = "research_th"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = research_th.TableVar
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
		Set research_th = Nothing
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
		If research_th.Export = "" Then
			SetupBreadcrumb()
		End If
		If IsPageRequest Then ' Validate request
			If Request.QueryString("rsh_id").Count > 0 Then
				research_th.rsh_id.QueryStringValue = Request.QueryString("rsh_id")
			Else
				sReturnUrl = "pom_research_thlist.asp" ' Return to list
			End If

			' Get action
			research_th.CurrentAction = "I" ' Display form
			Select Case research_th.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "pom_research_thlist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "pom_research_thlist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		research_th.RowType = EW_ROWTYPE_VIEW
		Call research_th.ResetAttrs()
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
				research_th.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					research_th.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = research_th.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			research_th.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			research_th.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			research_th.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = research_th.KeyFilter

		' Call Row Selecting event
		Call research_th.Row_Selecting(sFilter)

		' Load sql based on filter
		research_th.CurrentFilter = sFilter
		sSql = research_th.SQL
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
		Call research_th.Row_Selected(RsRow)
		research_th.rsh_id.DbValue = RsRow("rsh_id")
		research_th.rsh_img.DbValue = RsRow("rsh_img")
		research_th.rsh_date.DbValue = RsRow("rsh_date")
		research_th.rsh_pdf.DbValue = RsRow("rsh_pdf")
		research_th.rsh_category.DbValue = RsRow("rsh_category")
		research_th.rsh_subject.DbValue = RsRow("rsh_subject")
		research_th.rsh_subject_th.DbValue = RsRow("rsh_subject_th")
		research_th.rsh_intro.DbValue = RsRow("rsh_intro")
		research_th.rsh_intro_th.DbValue = RsRow("rsh_intro_th")
		research_th.rsh_content.DbValue = RsRow("rsh_content")
		research_th.rsh_content_th.DbValue = RsRow("rsh_content_th")
		research_th.rsh_show.DbValue = RsRow("rsh_show")
		research_th.rsh_show_home.DbValue = RsRow("rsh_show_home")
		research_th.rsh_create.DbValue = RsRow("rsh_create")
		research_th.rsh_update.DbValue = RsRow("rsh_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		research_th.rsh_id.m_DbValue = Rs("rsh_id")
		research_th.rsh_img.m_DbValue = Rs("rsh_img")
		research_th.rsh_date.m_DbValue = Rs("rsh_date")
		research_th.rsh_pdf.m_DbValue = Rs("rsh_pdf")
		research_th.rsh_category.m_DbValue = Rs("rsh_category")
		research_th.rsh_subject.m_DbValue = Rs("rsh_subject")
		research_th.rsh_subject_th.m_DbValue = Rs("rsh_subject_th")
		research_th.rsh_intro.m_DbValue = Rs("rsh_intro")
		research_th.rsh_intro_th.m_DbValue = Rs("rsh_intro_th")
		research_th.rsh_content.m_DbValue = Rs("rsh_content")
		research_th.rsh_content_th.m_DbValue = Rs("rsh_content_th")
		research_th.rsh_show.m_DbValue = Rs("rsh_show")
		research_th.rsh_show_home.m_DbValue = Rs("rsh_show_home")
		research_th.rsh_create.m_DbValue = Rs("rsh_create")
		research_th.rsh_update.m_DbValue = Rs("rsh_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = research_th.AddUrl
		EditUrl = research_th.EditUrl("")
		CopyUrl = research_th.CopyUrl("")
		DeleteUrl = research_th.DeleteUrl
		ListUrl = research_th.ListUrl
		SetupOtherOptions()

		' Call Row Rendering event
		Call research_th.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' rsh_id
		' rsh_img
		' rsh_date
		' rsh_pdf
		' rsh_category
		' rsh_subject
		' rsh_subject_th
		' rsh_intro
		' rsh_intro_th
		' rsh_content
		' rsh_content_th
		' rsh_show
		' rsh_show_home
		' rsh_create
		' rsh_update
		' -----------
		'  View  Row
		' -----------

		If research_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' rsh_id
			research_th.rsh_id.ViewValue = research_th.rsh_id.CurrentValue
			research_th.rsh_id.ViewCustomAttributes = ""

			' rsh_img
			research_th.rsh_img.ViewValue = research_th.rsh_img.CurrentValue
			research_th.rsh_img.ViewCustomAttributes = ""

			' rsh_date
			research_th.rsh_date.ViewValue = research_th.rsh_date.CurrentValue
			research_th.rsh_date.ViewCustomAttributes = ""

			' rsh_pdf
			research_th.rsh_pdf.ViewValue = research_th.rsh_pdf.CurrentValue
			research_th.rsh_pdf.ViewCustomAttributes = ""

			' rsh_category
			research_th.rsh_category.ViewValue = research_th.rsh_category.CurrentValue
			research_th.rsh_category.ViewCustomAttributes = ""

			' rsh_subject
			research_th.rsh_subject.ViewValue = research_th.rsh_subject.CurrentValue
			research_th.rsh_subject.ViewCustomAttributes = ""

			' rsh_subject_th
			research_th.rsh_subject_th.ViewValue = research_th.rsh_subject_th.CurrentValue
			research_th.rsh_subject_th.ViewCustomAttributes = ""

			' rsh_intro
			research_th.rsh_intro.ViewValue = research_th.rsh_intro.CurrentValue
			research_th.rsh_intro.ViewCustomAttributes = ""

			' rsh_intro_th
			research_th.rsh_intro_th.ViewValue = research_th.rsh_intro_th.CurrentValue
			research_th.rsh_intro_th.ViewCustomAttributes = ""

			' rsh_content
			research_th.rsh_content.ViewValue = research_th.rsh_content.CurrentValue
			research_th.rsh_content.ViewCustomAttributes = ""

			' rsh_content_th
			research_th.rsh_content_th.ViewValue = research_th.rsh_content_th.CurrentValue
			research_th.rsh_content_th.ViewCustomAttributes = ""

			' rsh_show
			research_th.rsh_show.ViewValue = research_th.rsh_show.CurrentValue
			research_th.rsh_show.ViewCustomAttributes = ""

			' rsh_show_home
			research_th.rsh_show_home.ViewValue = research_th.rsh_show_home.CurrentValue
			research_th.rsh_show_home.ViewCustomAttributes = ""

			' rsh_create
			research_th.rsh_create.ViewValue = research_th.rsh_create.CurrentValue
			research_th.rsh_create.ViewCustomAttributes = ""

			' rsh_update
			research_th.rsh_update.ViewValue = research_th.rsh_update.CurrentValue
			research_th.rsh_update.ViewCustomAttributes = ""

			' View refer script
			' rsh_id

			research_th.rsh_id.LinkCustomAttributes = ""
			research_th.rsh_id.HrefValue = ""
			research_th.rsh_id.TooltipValue = ""

			' rsh_img
			research_th.rsh_img.LinkCustomAttributes = ""
			research_th.rsh_img.HrefValue = ""
			research_th.rsh_img.TooltipValue = ""

			' rsh_date
			research_th.rsh_date.LinkCustomAttributes = ""
			research_th.rsh_date.HrefValue = ""
			research_th.rsh_date.TooltipValue = ""

			' rsh_pdf
			research_th.rsh_pdf.LinkCustomAttributes = ""
			research_th.rsh_pdf.HrefValue = ""
			research_th.rsh_pdf.TooltipValue = ""

			' rsh_category
			research_th.rsh_category.LinkCustomAttributes = ""
			research_th.rsh_category.HrefValue = ""
			research_th.rsh_category.TooltipValue = ""

			' rsh_subject
			research_th.rsh_subject.LinkCustomAttributes = ""
			research_th.rsh_subject.HrefValue = ""
			research_th.rsh_subject.TooltipValue = ""

			' rsh_subject_th
			research_th.rsh_subject_th.LinkCustomAttributes = ""
			research_th.rsh_subject_th.HrefValue = ""
			research_th.rsh_subject_th.TooltipValue = ""

			' rsh_intro
			research_th.rsh_intro.LinkCustomAttributes = ""
			research_th.rsh_intro.HrefValue = ""
			research_th.rsh_intro.TooltipValue = ""

			' rsh_intro_th
			research_th.rsh_intro_th.LinkCustomAttributes = ""
			research_th.rsh_intro_th.HrefValue = ""
			research_th.rsh_intro_th.TooltipValue = ""

			' rsh_content
			research_th.rsh_content.LinkCustomAttributes = ""
			research_th.rsh_content.HrefValue = ""
			research_th.rsh_content.TooltipValue = ""

			' rsh_content_th
			research_th.rsh_content_th.LinkCustomAttributes = ""
			research_th.rsh_content_th.HrefValue = ""
			research_th.rsh_content_th.TooltipValue = ""

			' rsh_show
			research_th.rsh_show.LinkCustomAttributes = ""
			research_th.rsh_show.HrefValue = ""
			research_th.rsh_show.TooltipValue = ""

			' rsh_show_home
			research_th.rsh_show_home.LinkCustomAttributes = ""
			research_th.rsh_show_home.HrefValue = ""
			research_th.rsh_show_home.TooltipValue = ""

			' rsh_create
			research_th.rsh_create.LinkCustomAttributes = ""
			research_th.rsh_create.HrefValue = ""
			research_th.rsh_create.TooltipValue = ""

			' rsh_update
			research_th.rsh_update.LinkCustomAttributes = ""
			research_th.rsh_update.HrefValue = ""
			research_th.rsh_update.TooltipValue = ""
		End If

		' Call Row Rendered event
		If research_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call research_th.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", research_th.TableVar, "pom_research_thlist.asp", research_th.TableVar, True)
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
