<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_person_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim person_th_view
Set person_th_view = New cperson_th_view
Set Page = person_th_view

' Page init processing
person_th_view.Page_Init()

' Page main processing
person_th_view.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
person_th_view.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If person_th.Export = "" Then %>
<script type="text/javascript">
// Page object
var person_th_view = new ew_Page("person_th_view");
person_th_view.PageID = "view"; // Page ID
var EW_PAGE_ID = person_th_view.PageID; // For backward compatibility
// Form object
var fperson_thview = new ew_Form("fperson_thview");
// Form_CustomValidate event
fperson_thview.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fperson_thview.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fperson_thview.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If person_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If person_th.Export = "" Then %>
<div class="ewViewExportOptions">
<% person_th_view.ExportOptions.Render "body", "", "", "", "", "" %>
<% If Not person_th_view.ExportOptions.UseDropDownButton Then %>
</div>
<div class="ewViewOtherOptions">
<% End If %>
<%
	person_th_view.ActionOptions.Render "body", "", "", "", "", ""
	person_th_view.DetailOptions.Render "body", "", "", "", "", ""
%>
</div>
<% End If %>
<% person_th_view.ShowPageHeader() %>
<% person_th_view.ShowMessage %>
<form name="fperson_thview" id="fperson_thview" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="person_th">
<table class="ewGrid"><tr><td>
<table id="tbl_person_thview" class="table table-bordered table-striped">
<% If person_th.per_id.Visible Then ' per_id %>
	<tr id="r_per_id">
		<td><span id="elh_person_th_per_id"><%= person_th.per_id.FldCaption %></span></td>
		<td<%= person_th.per_id.CellAttributes %>>
<span id="el_person_th_per_id" class="control-group">
<span<%= person_th.per_id.ViewAttributes %>>
<%= person_th.per_id.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If person_th.dept_id.Visible Then ' dept_id %>
	<tr id="r_dept_id">
		<td><span id="elh_person_th_dept_id"><%= person_th.dept_id.FldCaption %></span></td>
		<td<%= person_th.dept_id.CellAttributes %>>
<span id="el_person_th_dept_id" class="control-group">
<span<%= person_th.dept_id.ViewAttributes %>>
<%= person_th.dept_id.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If person_th.office_id.Visible Then ' office_id %>
	<tr id="r_office_id">
		<td><span id="elh_person_th_office_id"><%= person_th.office_id.FldCaption %></span></td>
		<td<%= person_th.office_id.CellAttributes %>>
<span id="el_person_th_office_id" class="control-group">
<span<%= person_th.office_id.ViewAttributes %>>
<%= person_th.office_id.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If person_th.per_img.Visible Then ' per_img %>
	<tr id="r_per_img">
		<td><span id="elh_person_th_per_img"><%= person_th.per_img.FldCaption %></span></td>
		<td<%= person_th.per_img.CellAttributes %>>
<span id="el_person_th_per_img" class="control-group">
<span<%= person_th.per_img.ViewAttributes %>>
<%= person_th.per_img.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If person_th.per_en_name.Visible Then ' per_en_name %>
	<tr id="r_per_en_name">
		<td><span id="elh_person_th_per_en_name"><%= person_th.per_en_name.FldCaption %></span></td>
		<td<%= person_th.per_en_name.CellAttributes %>>
<span id="el_person_th_per_en_name" class="control-group">
<span<%= person_th.per_en_name.ViewAttributes %>>
<%= person_th.per_en_name.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If person_th.per_th_name.Visible Then ' per_th_name %>
	<tr id="r_per_th_name">
		<td><span id="elh_person_th_per_th_name"><%= person_th.per_th_name.FldCaption %></span></td>
		<td<%= person_th.per_th_name.CellAttributes %>>
<span id="el_person_th_per_th_name" class="control-group">
<span<%= person_th.per_th_name.ViewAttributes %>>
<%= person_th.per_th_name.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If person_th.per_position.Visible Then ' per_position %>
	<tr id="r_per_position">
		<td><span id="elh_person_th_per_position"><%= person_th.per_position.FldCaption %></span></td>
		<td<%= person_th.per_position.CellAttributes %>>
<span id="el_person_th_per_position" class="control-group">
<span<%= person_th.per_position.ViewAttributes %>>
<%= person_th.per_position.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If person_th.per_mobile.Visible Then ' per_mobile %>
	<tr id="r_per_mobile">
		<td><span id="elh_person_th_per_mobile"><%= person_th.per_mobile.FldCaption %></span></td>
		<td<%= person_th.per_mobile.CellAttributes %>>
<span id="el_person_th_per_mobile" class="control-group">
<span<%= person_th.per_mobile.ViewAttributes %>>
<%= person_th.per_mobile.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If person_th.per_tel.Visible Then ' per_tel %>
	<tr id="r_per_tel">
		<td><span id="elh_person_th_per_tel"><%= person_th.per_tel.FldCaption %></span></td>
		<td<%= person_th.per_tel.CellAttributes %>>
<span id="el_person_th_per_tel" class="control-group">
<span<%= person_th.per_tel.ViewAttributes %>>
<%= person_th.per_tel.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If person_th.per_fax.Visible Then ' per_fax %>
	<tr id="r_per_fax">
		<td><span id="elh_person_th_per_fax"><%= person_th.per_fax.FldCaption %></span></td>
		<td<%= person_th.per_fax.CellAttributes %>>
<span id="el_person_th_per_fax" class="control-group">
<span<%= person_th.per_fax.ViewAttributes %>>
<%= person_th.per_fax.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If person_th.per_email.Visible Then ' per_email %>
	<tr id="r_per_email">
		<td><span id="elh_person_th_per_email"><%= person_th.per_email.FldCaption %></span></td>
		<td<%= person_th.per_email.CellAttributes %>>
<span id="el_person_th_per_email" class="control-group">
<span<%= person_th.per_email.ViewAttributes %>>
<%= person_th.per_email.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If person_th.per_address.Visible Then ' per_address %>
	<tr id="r_per_address">
		<td><span id="elh_person_th_per_address"><%= person_th.per_address.FldCaption %></span></td>
		<td<%= person_th.per_address.CellAttributes %>>
<span id="el_person_th_per_address" class="control-group">
<span<%= person_th.per_address.ViewAttributes %>>
<%= person_th.per_address.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If person_th.per_show.Visible Then ' per_show %>
	<tr id="r_per_show">
		<td><span id="elh_person_th_per_show"><%= person_th.per_show.FldCaption %></span></td>
		<td<%= person_th.per_show.CellAttributes %>>
<span id="el_person_th_per_show" class="control-group">
<span<%= person_th.per_show.ViewAttributes %>>
<%= person_th.per_show.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If person_th.per_create.Visible Then ' per_create %>
	<tr id="r_per_create">
		<td><span id="elh_person_th_per_create"><%= person_th.per_create.FldCaption %></span></td>
		<td<%= person_th.per_create.CellAttributes %>>
<span id="el_person_th_per_create" class="control-group">
<span<%= person_th.per_create.ViewAttributes %>>
<%= person_th.per_create.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If person_th.per_update.Visible Then ' per_update %>
	<tr id="r_per_update">
		<td><span id="elh_person_th_per_update"><%= person_th.per_update.FldCaption %></span></td>
		<td<%= person_th.per_update.CellAttributes %>>
<span id="el_person_th_per_update" class="control-group">
<span<%= person_th.per_update.ViewAttributes %>>
<%= person_th.per_update.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If person_th.per_sort.Visible Then ' per_sort %>
	<tr id="r_per_sort">
		<td><span id="elh_person_th_per_sort"><%= person_th.per_sort.FldCaption %></span></td>
		<td<%= person_th.per_sort.CellAttributes %>>
<span id="el_person_th_per_sort" class="control-group">
<span<%= person_th.per_sort.ViewAttributes %>>
<%= person_th.per_sort.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If person_th.per_department.Visible Then ' per_department %>
	<tr id="r_per_department">
		<td><span id="elh_person_th_per_department"><%= person_th.per_department.FldCaption %></span></td>
		<td<%= person_th.per_department.CellAttributes %>>
<span id="el_person_th_per_department" class="control-group">
<span<%= person_th.per_department.ViewAttributes %>>
<%= person_th.per_department.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
</table>
</td></tr></table>
</form>
<script type="text/javascript">
fperson_thview.Init();
</script>
<%
person_th_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If person_th.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set person_th_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cperson_th_view

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
		TableName = "person_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "person_th_view"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If person_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & person_th.TableVar & "&" ' add page token
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
		If person_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (person_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (person_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(person_th) Then Set person_th = New cperson_th
		Set Table = person_th

		' Initialize urls
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("per_id").Count > 0 Then
			ew_AddKey RecKey, "per_id", Request.QueryString("per_id")
			KeyUrl = KeyUrl & "&amp;per_id=" & Server.URLEncode(Request.QueryString("per_id"))
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
		EW_TABLE_NAME = "person_th"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = person_th.TableVar
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
		Set person_th = Nothing
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
		If person_th.Export = "" Then
			SetupBreadcrumb()
		End If
		If IsPageRequest Then ' Validate request
			If Request.QueryString("per_id").Count > 0 Then
				person_th.per_id.QueryStringValue = Request.QueryString("per_id")
			Else
				sReturnUrl = "pom_person_thlist.asp" ' Return to list
			End If

			' Get action
			person_th.CurrentAction = "I" ' Display form
			Select Case person_th.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "pom_person_thlist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "pom_person_thlist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		person_th.RowType = EW_ROWTYPE_VIEW
		Call person_th.ResetAttrs()
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
				person_th.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					person_th.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = person_th.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			person_th.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			person_th.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			person_th.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = person_th.KeyFilter

		' Call Row Selecting event
		Call person_th.Row_Selecting(sFilter)

		' Load sql based on filter
		person_th.CurrentFilter = sFilter
		sSql = person_th.SQL
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
		Call person_th.Row_Selected(RsRow)
		person_th.per_id.DbValue = RsRow("per_id")
		person_th.dept_id.DbValue = RsRow("dept_id")
		person_th.office_id.DbValue = RsRow("office_id")
		person_th.per_img.DbValue = RsRow("per_img")
		person_th.per_en_name.DbValue = RsRow("per_en_name")
		person_th.per_th_name.DbValue = RsRow("per_th_name")
		person_th.per_position.DbValue = RsRow("per_position")
		person_th.per_mobile.DbValue = RsRow("per_mobile")
		person_th.per_tel.DbValue = RsRow("per_tel")
		person_th.per_fax.DbValue = RsRow("per_fax")
		person_th.per_email.DbValue = RsRow("per_email")
		person_th.per_address.DbValue = RsRow("per_address")
		person_th.per_show.DbValue = RsRow("per_show")
		person_th.per_create.DbValue = RsRow("per_create")
		person_th.per_update.DbValue = RsRow("per_update")
		person_th.per_sort.DbValue = RsRow("per_sort")
		person_th.per_department.DbValue = RsRow("per_department")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		person_th.per_id.m_DbValue = Rs("per_id")
		person_th.dept_id.m_DbValue = Rs("dept_id")
		person_th.office_id.m_DbValue = Rs("office_id")
		person_th.per_img.m_DbValue = Rs("per_img")
		person_th.per_en_name.m_DbValue = Rs("per_en_name")
		person_th.per_th_name.m_DbValue = Rs("per_th_name")
		person_th.per_position.m_DbValue = Rs("per_position")
		person_th.per_mobile.m_DbValue = Rs("per_mobile")
		person_th.per_tel.m_DbValue = Rs("per_tel")
		person_th.per_fax.m_DbValue = Rs("per_fax")
		person_th.per_email.m_DbValue = Rs("per_email")
		person_th.per_address.m_DbValue = Rs("per_address")
		person_th.per_show.m_DbValue = Rs("per_show")
		person_th.per_create.m_DbValue = Rs("per_create")
		person_th.per_update.m_DbValue = Rs("per_update")
		person_th.per_sort.m_DbValue = Rs("per_sort")
		person_th.per_department.m_DbValue = Rs("per_department")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = person_th.AddUrl
		EditUrl = person_th.EditUrl("")
		CopyUrl = person_th.CopyUrl("")
		DeleteUrl = person_th.DeleteUrl
		ListUrl = person_th.ListUrl
		SetupOtherOptions()

		' Call Row Rendering event
		Call person_th.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' per_id
		' dept_id
		' office_id
		' per_img
		' per_en_name
		' per_th_name
		' per_position
		' per_mobile
		' per_tel
		' per_fax
		' per_email
		' per_address
		' per_show
		' per_create
		' per_update
		' per_sort
		' per_department
		' -----------
		'  View  Row
		' -----------

		If person_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' per_id
			person_th.per_id.ViewValue = person_th.per_id.CurrentValue
			person_th.per_id.ViewCustomAttributes = ""

			' dept_id
			person_th.dept_id.ViewValue = person_th.dept_id.CurrentValue
			person_th.dept_id.ViewCustomAttributes = ""

			' office_id
			person_th.office_id.ViewValue = person_th.office_id.CurrentValue
			person_th.office_id.ViewCustomAttributes = ""

			' per_img
			person_th.per_img.ViewValue = person_th.per_img.CurrentValue
			person_th.per_img.ViewCustomAttributes = ""

			' per_en_name
			person_th.per_en_name.ViewValue = person_th.per_en_name.CurrentValue
			person_th.per_en_name.ViewCustomAttributes = ""

			' per_th_name
			person_th.per_th_name.ViewValue = person_th.per_th_name.CurrentValue
			person_th.per_th_name.ViewCustomAttributes = ""

			' per_position
			person_th.per_position.ViewValue = person_th.per_position.CurrentValue
			person_th.per_position.ViewCustomAttributes = ""

			' per_mobile
			person_th.per_mobile.ViewValue = person_th.per_mobile.CurrentValue
			person_th.per_mobile.ViewCustomAttributes = ""

			' per_tel
			person_th.per_tel.ViewValue = person_th.per_tel.CurrentValue
			person_th.per_tel.ViewCustomAttributes = ""

			' per_fax
			person_th.per_fax.ViewValue = person_th.per_fax.CurrentValue
			person_th.per_fax.ViewCustomAttributes = ""

			' per_email
			person_th.per_email.ViewValue = person_th.per_email.CurrentValue
			person_th.per_email.ViewCustomAttributes = ""

			' per_address
			person_th.per_address.ViewValue = person_th.per_address.CurrentValue
			person_th.per_address.ViewCustomAttributes = ""

			' per_show
			person_th.per_show.ViewValue = person_th.per_show.CurrentValue
			person_th.per_show.ViewCustomAttributes = ""

			' per_create
			person_th.per_create.ViewValue = person_th.per_create.CurrentValue
			person_th.per_create.ViewCustomAttributes = ""

			' per_update
			person_th.per_update.ViewValue = person_th.per_update.CurrentValue
			person_th.per_update.ViewCustomAttributes = ""

			' per_sort
			person_th.per_sort.ViewValue = person_th.per_sort.CurrentValue
			person_th.per_sort.ViewCustomAttributes = ""

			' per_department
			person_th.per_department.ViewValue = person_th.per_department.CurrentValue
			person_th.per_department.ViewCustomAttributes = ""

			' View refer script
			' per_id

			person_th.per_id.LinkCustomAttributes = ""
			person_th.per_id.HrefValue = ""
			person_th.per_id.TooltipValue = ""

			' dept_id
			person_th.dept_id.LinkCustomAttributes = ""
			person_th.dept_id.HrefValue = ""
			person_th.dept_id.TooltipValue = ""

			' office_id
			person_th.office_id.LinkCustomAttributes = ""
			person_th.office_id.HrefValue = ""
			person_th.office_id.TooltipValue = ""

			' per_img
			person_th.per_img.LinkCustomAttributes = ""
			person_th.per_img.HrefValue = ""
			person_th.per_img.TooltipValue = ""

			' per_en_name
			person_th.per_en_name.LinkCustomAttributes = ""
			person_th.per_en_name.HrefValue = ""
			person_th.per_en_name.TooltipValue = ""

			' per_th_name
			person_th.per_th_name.LinkCustomAttributes = ""
			person_th.per_th_name.HrefValue = ""
			person_th.per_th_name.TooltipValue = ""

			' per_position
			person_th.per_position.LinkCustomAttributes = ""
			person_th.per_position.HrefValue = ""
			person_th.per_position.TooltipValue = ""

			' per_mobile
			person_th.per_mobile.LinkCustomAttributes = ""
			person_th.per_mobile.HrefValue = ""
			person_th.per_mobile.TooltipValue = ""

			' per_tel
			person_th.per_tel.LinkCustomAttributes = ""
			person_th.per_tel.HrefValue = ""
			person_th.per_tel.TooltipValue = ""

			' per_fax
			person_th.per_fax.LinkCustomAttributes = ""
			person_th.per_fax.HrefValue = ""
			person_th.per_fax.TooltipValue = ""

			' per_email
			person_th.per_email.LinkCustomAttributes = ""
			person_th.per_email.HrefValue = ""
			person_th.per_email.TooltipValue = ""

			' per_address
			person_th.per_address.LinkCustomAttributes = ""
			person_th.per_address.HrefValue = ""
			person_th.per_address.TooltipValue = ""

			' per_show
			person_th.per_show.LinkCustomAttributes = ""
			person_th.per_show.HrefValue = ""
			person_th.per_show.TooltipValue = ""

			' per_create
			person_th.per_create.LinkCustomAttributes = ""
			person_th.per_create.HrefValue = ""
			person_th.per_create.TooltipValue = ""

			' per_update
			person_th.per_update.LinkCustomAttributes = ""
			person_th.per_update.HrefValue = ""
			person_th.per_update.TooltipValue = ""

			' per_sort
			person_th.per_sort.LinkCustomAttributes = ""
			person_th.per_sort.HrefValue = ""
			person_th.per_sort.TooltipValue = ""

			' per_department
			person_th.per_department.LinkCustomAttributes = ""
			person_th.per_department.HrefValue = ""
			person_th.per_department.TooltipValue = ""
		End If

		' Call Row Rendered event
		If person_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call person_th.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", person_th.TableVar, "pom_person_thlist.asp", person_th.TableVar, True)
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
