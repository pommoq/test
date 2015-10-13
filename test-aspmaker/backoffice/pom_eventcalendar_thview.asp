<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_eventcalendar_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim eventcalendar_th_view
Set eventcalendar_th_view = New ceventcalendar_th_view
Set Page = eventcalendar_th_view

' Page init processing
eventcalendar_th_view.Page_Init()

' Page main processing
eventcalendar_th_view.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
eventcalendar_th_view.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If eventcalendar_th.Export = "" Then %>
<script type="text/javascript">
// Page object
var eventcalendar_th_view = new ew_Page("eventcalendar_th_view");
eventcalendar_th_view.PageID = "view"; // Page ID
var EW_PAGE_ID = eventcalendar_th_view.PageID; // For backward compatibility
// Form object
var feventcalendar_thview = new ew_Form("feventcalendar_thview");
// Form_CustomValidate event
feventcalendar_thview.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
feventcalendar_thview.ValidateRequired = true; // Use JavaScript validation
<% Else %>
feventcalendar_thview.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If eventcalendar_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If eventcalendar_th.Export = "" Then %>
<div class="ewViewExportOptions">
<% eventcalendar_th_view.ExportOptions.Render "body", "", "", "", "", "" %>
<% If Not eventcalendar_th_view.ExportOptions.UseDropDownButton Then %>
</div>
<div class="ewViewOtherOptions">
<% End If %>
<%
	eventcalendar_th_view.ActionOptions.Render "body", "", "", "", "", ""
	eventcalendar_th_view.DetailOptions.Render "body", "", "", "", "", ""
%>
</div>
<% End If %>
<% eventcalendar_th_view.ShowPageHeader() %>
<% eventcalendar_th_view.ShowMessage %>
<form name="feventcalendar_thview" id="feventcalendar_thview" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="eventcalendar_th">
<table class="ewGrid"><tr><td>
<table id="tbl_eventcalendar_thview" class="table table-bordered table-striped">
<% If eventcalendar_th.eventcalendar_id.Visible Then ' eventcalendar_id %>
	<tr id="r_eventcalendar_id">
		<td><span id="elh_eventcalendar_th_eventcalendar_id"><%= eventcalendar_th.eventcalendar_id.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_id.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_id" class="control-group">
<span<%= eventcalendar_th.eventcalendar_id.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_id.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_img.Visible Then ' eventcalendar_img %>
	<tr id="r_eventcalendar_img">
		<td><span id="elh_eventcalendar_th_eventcalendar_img"><%= eventcalendar_th.eventcalendar_img.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_img.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_img" class="control-group">
<span<%= eventcalendar_th.eventcalendar_img.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_img.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_date.Visible Then ' eventcalendar_date %>
	<tr id="r_eventcalendar_date">
		<td><span id="elh_eventcalendar_th_eventcalendar_date"><%= eventcalendar_th.eventcalendar_date.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_date.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_date" class="control-group">
<span<%= eventcalendar_th.eventcalendar_date.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_date.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_category.Visible Then ' eventcalendar_category %>
	<tr id="r_eventcalendar_category">
		<td><span id="elh_eventcalendar_th_eventcalendar_category"><%= eventcalendar_th.eventcalendar_category.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_category.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_category" class="control-group">
<span<%= eventcalendar_th.eventcalendar_category.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_category.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_category_sub.Visible Then ' eventcalendar_category_sub %>
	<tr id="r_eventcalendar_category_sub">
		<td><span id="elh_eventcalendar_th_eventcalendar_category_sub"><%= eventcalendar_th.eventcalendar_category_sub.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_category_sub.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_category_sub" class="control-group">
<span<%= eventcalendar_th.eventcalendar_category_sub.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_category_sub.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.start_date.Visible Then ' start_date %>
	<tr id="r_start_date">
		<td><span id="elh_eventcalendar_th_start_date"><%= eventcalendar_th.start_date.FldCaption %></span></td>
		<td<%= eventcalendar_th.start_date.CellAttributes %>>
<span id="el_eventcalendar_th_start_date" class="control-group">
<span<%= eventcalendar_th.start_date.ViewAttributes %>>
<%= eventcalendar_th.start_date.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.end_date.Visible Then ' end_date %>
	<tr id="r_end_date">
		<td><span id="elh_eventcalendar_th_end_date"><%= eventcalendar_th.end_date.FldCaption %></span></td>
		<td<%= eventcalendar_th.end_date.CellAttributes %>>
<span id="el_eventcalendar_th_end_date" class="control-group">
<span<%= eventcalendar_th.end_date.ViewAttributes %>>
<%= eventcalendar_th.end_date.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_pdf.Visible Then ' eventcalendar_pdf %>
	<tr id="r_eventcalendar_pdf">
		<td><span id="elh_eventcalendar_th_eventcalendar_pdf"><%= eventcalendar_th.eventcalendar_pdf.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_pdf.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_pdf" class="control-group">
<span<%= eventcalendar_th.eventcalendar_pdf.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_pdf.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_subject.Visible Then ' eventcalendar_subject %>
	<tr id="r_eventcalendar_subject">
		<td><span id="elh_eventcalendar_th_eventcalendar_subject"><%= eventcalendar_th.eventcalendar_subject.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_subject.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_subject" class="control-group">
<span<%= eventcalendar_th.eventcalendar_subject.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_subject.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_subject_th.Visible Then ' eventcalendar_subject_th %>
	<tr id="r_eventcalendar_subject_th">
		<td><span id="elh_eventcalendar_th_eventcalendar_subject_th"><%= eventcalendar_th.eventcalendar_subject_th.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_subject_th.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_subject_th" class="control-group">
<span<%= eventcalendar_th.eventcalendar_subject_th.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_subject_th.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_intro.Visible Then ' eventcalendar_intro %>
	<tr id="r_eventcalendar_intro">
		<td><span id="elh_eventcalendar_th_eventcalendar_intro"><%= eventcalendar_th.eventcalendar_intro.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_intro.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_intro" class="control-group">
<span<%= eventcalendar_th.eventcalendar_intro.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_intro.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_intro_th.Visible Then ' eventcalendar_intro_th %>
	<tr id="r_eventcalendar_intro_th">
		<td><span id="elh_eventcalendar_th_eventcalendar_intro_th"><%= eventcalendar_th.eventcalendar_intro_th.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_intro_th.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_intro_th" class="control-group">
<span<%= eventcalendar_th.eventcalendar_intro_th.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_intro_th.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_content.Visible Then ' eventcalendar_content %>
	<tr id="r_eventcalendar_content">
		<td><span id="elh_eventcalendar_th_eventcalendar_content"><%= eventcalendar_th.eventcalendar_content.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_content.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_content" class="control-group">
<span<%= eventcalendar_th.eventcalendar_content.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_content.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_content_th.Visible Then ' eventcalendar_content_th %>
	<tr id="r_eventcalendar_content_th">
		<td><span id="elh_eventcalendar_th_eventcalendar_content_th"><%= eventcalendar_th.eventcalendar_content_th.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_content_th.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_content_th" class="control-group">
<span<%= eventcalendar_th.eventcalendar_content_th.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_content_th.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_show_en.Visible Then ' eventcalendar_show_en %>
	<tr id="r_eventcalendar_show_en">
		<td><span id="elh_eventcalendar_th_eventcalendar_show_en"><%= eventcalendar_th.eventcalendar_show_en.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_show_en.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_show_en" class="control-group">
<span<%= eventcalendar_th.eventcalendar_show_en.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_show_en.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_show.Visible Then ' eventcalendar_show %>
	<tr id="r_eventcalendar_show">
		<td><span id="elh_eventcalendar_th_eventcalendar_show"><%= eventcalendar_th.eventcalendar_show.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_show.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_show" class="control-group">
<span<%= eventcalendar_th.eventcalendar_show.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_show.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_show_home.Visible Then ' eventcalendar_show_home %>
	<tr id="r_eventcalendar_show_home">
		<td><span id="elh_eventcalendar_th_eventcalendar_show_home"><%= eventcalendar_th.eventcalendar_show_home.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_show_home.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_show_home" class="control-group">
<span<%= eventcalendar_th.eventcalendar_show_home.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_show_home.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_create.Visible Then ' eventcalendar_create %>
	<tr id="r_eventcalendar_create">
		<td><span id="elh_eventcalendar_th_eventcalendar_create"><%= eventcalendar_th.eventcalendar_create.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_create.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_create" class="control-group">
<span<%= eventcalendar_th.eventcalendar_create.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_create.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_update.Visible Then ' eventcalendar_update %>
	<tr id="r_eventcalendar_update">
		<td><span id="elh_eventcalendar_th_eventcalendar_update"><%= eventcalendar_th.eventcalendar_update.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_update.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_update" class="control-group">
<span<%= eventcalendar_th.eventcalendar_update.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_update.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
</table>
</td></tr></table>
</form>
<script type="text/javascript">
feventcalendar_thview.Init();
</script>
<%
eventcalendar_th_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If eventcalendar_th.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set eventcalendar_th_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class ceventcalendar_th_view

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
		TableName = "eventcalendar_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "eventcalendar_th_view"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If eventcalendar_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & eventcalendar_th.TableVar & "&" ' add page token
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
		If eventcalendar_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (eventcalendar_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (eventcalendar_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(eventcalendar_th) Then Set eventcalendar_th = New ceventcalendar_th
		Set Table = eventcalendar_th

		' Initialize urls
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("eventcalendar_id").Count > 0 Then
			ew_AddKey RecKey, "eventcalendar_id", Request.QueryString("eventcalendar_id")
			KeyUrl = KeyUrl & "&amp;eventcalendar_id=" & Server.URLEncode(Request.QueryString("eventcalendar_id"))
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
		EW_TABLE_NAME = "eventcalendar_th"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = eventcalendar_th.TableVar
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
		Set eventcalendar_th = Nothing
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
		If eventcalendar_th.Export = "" Then
			SetupBreadcrumb()
		End If
		If IsPageRequest Then ' Validate request
			If Request.QueryString("eventcalendar_id").Count > 0 Then
				eventcalendar_th.eventcalendar_id.QueryStringValue = Request.QueryString("eventcalendar_id")
			Else
				sReturnUrl = "pom_eventcalendar_thlist.asp" ' Return to list
			End If

			' Get action
			eventcalendar_th.CurrentAction = "I" ' Display form
			Select Case eventcalendar_th.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "pom_eventcalendar_thlist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "pom_eventcalendar_thlist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		eventcalendar_th.RowType = EW_ROWTYPE_VIEW
		Call eventcalendar_th.ResetAttrs()
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
				eventcalendar_th.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					eventcalendar_th.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = eventcalendar_th.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			eventcalendar_th.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			eventcalendar_th.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			eventcalendar_th.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = eventcalendar_th.KeyFilter

		' Call Row Selecting event
		Call eventcalendar_th.Row_Selecting(sFilter)

		' Load sql based on filter
		eventcalendar_th.CurrentFilter = sFilter
		sSql = eventcalendar_th.SQL
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
		Call eventcalendar_th.Row_Selected(RsRow)
		eventcalendar_th.eventcalendar_id.DbValue = RsRow("eventcalendar_id")
		eventcalendar_th.eventcalendar_img.DbValue = RsRow("eventcalendar_img")
		eventcalendar_th.eventcalendar_date.DbValue = RsRow("eventcalendar_date")
		eventcalendar_th.eventcalendar_category.DbValue = RsRow("eventcalendar_category")
		eventcalendar_th.eventcalendar_category_sub.DbValue = RsRow("eventcalendar_category_sub")
		eventcalendar_th.start_date.DbValue = RsRow("start_date")
		eventcalendar_th.end_date.DbValue = RsRow("end_date")
		eventcalendar_th.eventcalendar_pdf.DbValue = RsRow("eventcalendar_pdf")
		eventcalendar_th.eventcalendar_subject.DbValue = RsRow("eventcalendar_subject")
		eventcalendar_th.eventcalendar_subject_th.DbValue = RsRow("eventcalendar_subject_th")
		eventcalendar_th.eventcalendar_intro.DbValue = RsRow("eventcalendar_intro")
		eventcalendar_th.eventcalendar_intro_th.DbValue = RsRow("eventcalendar_intro_th")
		eventcalendar_th.eventcalendar_content.DbValue = RsRow("eventcalendar_content")
		eventcalendar_th.eventcalendar_content_th.DbValue = RsRow("eventcalendar_content_th")
		eventcalendar_th.eventcalendar_show_en.DbValue = RsRow("eventcalendar_show_en")
		eventcalendar_th.eventcalendar_show.DbValue = RsRow("eventcalendar_show")
		eventcalendar_th.eventcalendar_show_home.DbValue = RsRow("eventcalendar_show_home")
		eventcalendar_th.eventcalendar_create.DbValue = RsRow("eventcalendar_create")
		eventcalendar_th.eventcalendar_update.DbValue = RsRow("eventcalendar_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		eventcalendar_th.eventcalendar_id.m_DbValue = Rs("eventcalendar_id")
		eventcalendar_th.eventcalendar_img.m_DbValue = Rs("eventcalendar_img")
		eventcalendar_th.eventcalendar_date.m_DbValue = Rs("eventcalendar_date")
		eventcalendar_th.eventcalendar_category.m_DbValue = Rs("eventcalendar_category")
		eventcalendar_th.eventcalendar_category_sub.m_DbValue = Rs("eventcalendar_category_sub")
		eventcalendar_th.start_date.m_DbValue = Rs("start_date")
		eventcalendar_th.end_date.m_DbValue = Rs("end_date")
		eventcalendar_th.eventcalendar_pdf.m_DbValue = Rs("eventcalendar_pdf")
		eventcalendar_th.eventcalendar_subject.m_DbValue = Rs("eventcalendar_subject")
		eventcalendar_th.eventcalendar_subject_th.m_DbValue = Rs("eventcalendar_subject_th")
		eventcalendar_th.eventcalendar_intro.m_DbValue = Rs("eventcalendar_intro")
		eventcalendar_th.eventcalendar_intro_th.m_DbValue = Rs("eventcalendar_intro_th")
		eventcalendar_th.eventcalendar_content.m_DbValue = Rs("eventcalendar_content")
		eventcalendar_th.eventcalendar_content_th.m_DbValue = Rs("eventcalendar_content_th")
		eventcalendar_th.eventcalendar_show_en.m_DbValue = Rs("eventcalendar_show_en")
		eventcalendar_th.eventcalendar_show.m_DbValue = Rs("eventcalendar_show")
		eventcalendar_th.eventcalendar_show_home.m_DbValue = Rs("eventcalendar_show_home")
		eventcalendar_th.eventcalendar_create.m_DbValue = Rs("eventcalendar_create")
		eventcalendar_th.eventcalendar_update.m_DbValue = Rs("eventcalendar_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = eventcalendar_th.AddUrl
		EditUrl = eventcalendar_th.EditUrl("")
		CopyUrl = eventcalendar_th.CopyUrl("")
		DeleteUrl = eventcalendar_th.DeleteUrl
		ListUrl = eventcalendar_th.ListUrl
		SetupOtherOptions()

		' Call Row Rendering event
		Call eventcalendar_th.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' eventcalendar_id
		' eventcalendar_img
		' eventcalendar_date
		' eventcalendar_category
		' eventcalendar_category_sub
		' start_date
		' end_date
		' eventcalendar_pdf
		' eventcalendar_subject
		' eventcalendar_subject_th
		' eventcalendar_intro
		' eventcalendar_intro_th
		' eventcalendar_content
		' eventcalendar_content_th
		' eventcalendar_show_en
		' eventcalendar_show
		' eventcalendar_show_home
		' eventcalendar_create
		' eventcalendar_update
		' -----------
		'  View  Row
		' -----------

		If eventcalendar_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' eventcalendar_id
			eventcalendar_th.eventcalendar_id.ViewValue = eventcalendar_th.eventcalendar_id.CurrentValue
			eventcalendar_th.eventcalendar_id.ViewCustomAttributes = ""

			' eventcalendar_img
			eventcalendar_th.eventcalendar_img.ViewValue = eventcalendar_th.eventcalendar_img.CurrentValue
			eventcalendar_th.eventcalendar_img.ViewCustomAttributes = ""

			' eventcalendar_date
			eventcalendar_th.eventcalendar_date.ViewValue = eventcalendar_th.eventcalendar_date.CurrentValue
			eventcalendar_th.eventcalendar_date.ViewCustomAttributes = ""

			' eventcalendar_category
			eventcalendar_th.eventcalendar_category.ViewValue = eventcalendar_th.eventcalendar_category.CurrentValue
			eventcalendar_th.eventcalendar_category.ViewCustomAttributes = ""

			' eventcalendar_category_sub
			eventcalendar_th.eventcalendar_category_sub.ViewValue = eventcalendar_th.eventcalendar_category_sub.CurrentValue
			eventcalendar_th.eventcalendar_category_sub.ViewCustomAttributes = ""

			' start_date
			eventcalendar_th.start_date.ViewValue = eventcalendar_th.start_date.CurrentValue
			eventcalendar_th.start_date.ViewCustomAttributes = ""

			' end_date
			eventcalendar_th.end_date.ViewValue = eventcalendar_th.end_date.CurrentValue
			eventcalendar_th.end_date.ViewCustomAttributes = ""

			' eventcalendar_pdf
			eventcalendar_th.eventcalendar_pdf.ViewValue = eventcalendar_th.eventcalendar_pdf.CurrentValue
			eventcalendar_th.eventcalendar_pdf.ViewCustomAttributes = ""

			' eventcalendar_subject
			eventcalendar_th.eventcalendar_subject.ViewValue = eventcalendar_th.eventcalendar_subject.CurrentValue
			eventcalendar_th.eventcalendar_subject.ViewCustomAttributes = ""

			' eventcalendar_subject_th
			eventcalendar_th.eventcalendar_subject_th.ViewValue = eventcalendar_th.eventcalendar_subject_th.CurrentValue
			eventcalendar_th.eventcalendar_subject_th.ViewCustomAttributes = ""

			' eventcalendar_intro
			eventcalendar_th.eventcalendar_intro.ViewValue = eventcalendar_th.eventcalendar_intro.CurrentValue
			eventcalendar_th.eventcalendar_intro.ViewCustomAttributes = ""

			' eventcalendar_intro_th
			eventcalendar_th.eventcalendar_intro_th.ViewValue = eventcalendar_th.eventcalendar_intro_th.CurrentValue
			eventcalendar_th.eventcalendar_intro_th.ViewCustomAttributes = ""

			' eventcalendar_content
			eventcalendar_th.eventcalendar_content.ViewValue = eventcalendar_th.eventcalendar_content.CurrentValue
			eventcalendar_th.eventcalendar_content.ViewCustomAttributes = ""

			' eventcalendar_content_th
			eventcalendar_th.eventcalendar_content_th.ViewValue = eventcalendar_th.eventcalendar_content_th.CurrentValue
			eventcalendar_th.eventcalendar_content_th.ViewCustomAttributes = ""

			' eventcalendar_show_en
			eventcalendar_th.eventcalendar_show_en.ViewValue = eventcalendar_th.eventcalendar_show_en.CurrentValue
			eventcalendar_th.eventcalendar_show_en.ViewCustomAttributes = ""

			' eventcalendar_show
			eventcalendar_th.eventcalendar_show.ViewValue = eventcalendar_th.eventcalendar_show.CurrentValue
			eventcalendar_th.eventcalendar_show.ViewCustomAttributes = ""

			' eventcalendar_show_home
			eventcalendar_th.eventcalendar_show_home.ViewValue = eventcalendar_th.eventcalendar_show_home.CurrentValue
			eventcalendar_th.eventcalendar_show_home.ViewCustomAttributes = ""

			' eventcalendar_create
			eventcalendar_th.eventcalendar_create.ViewValue = eventcalendar_th.eventcalendar_create.CurrentValue
			eventcalendar_th.eventcalendar_create.ViewCustomAttributes = ""

			' eventcalendar_update
			eventcalendar_th.eventcalendar_update.ViewValue = eventcalendar_th.eventcalendar_update.CurrentValue
			eventcalendar_th.eventcalendar_update.ViewCustomAttributes = ""

			' View refer script
			' eventcalendar_id

			eventcalendar_th.eventcalendar_id.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_id.HrefValue = ""
			eventcalendar_th.eventcalendar_id.TooltipValue = ""

			' eventcalendar_img
			eventcalendar_th.eventcalendar_img.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_img.HrefValue = ""
			eventcalendar_th.eventcalendar_img.TooltipValue = ""

			' eventcalendar_date
			eventcalendar_th.eventcalendar_date.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_date.HrefValue = ""
			eventcalendar_th.eventcalendar_date.TooltipValue = ""

			' eventcalendar_category
			eventcalendar_th.eventcalendar_category.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_category.HrefValue = ""
			eventcalendar_th.eventcalendar_category.TooltipValue = ""

			' eventcalendar_category_sub
			eventcalendar_th.eventcalendar_category_sub.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_category_sub.HrefValue = ""
			eventcalendar_th.eventcalendar_category_sub.TooltipValue = ""

			' start_date
			eventcalendar_th.start_date.LinkCustomAttributes = ""
			eventcalendar_th.start_date.HrefValue = ""
			eventcalendar_th.start_date.TooltipValue = ""

			' end_date
			eventcalendar_th.end_date.LinkCustomAttributes = ""
			eventcalendar_th.end_date.HrefValue = ""
			eventcalendar_th.end_date.TooltipValue = ""

			' eventcalendar_pdf
			eventcalendar_th.eventcalendar_pdf.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_pdf.HrefValue = ""
			eventcalendar_th.eventcalendar_pdf.TooltipValue = ""

			' eventcalendar_subject
			eventcalendar_th.eventcalendar_subject.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_subject.HrefValue = ""
			eventcalendar_th.eventcalendar_subject.TooltipValue = ""

			' eventcalendar_subject_th
			eventcalendar_th.eventcalendar_subject_th.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_subject_th.HrefValue = ""
			eventcalendar_th.eventcalendar_subject_th.TooltipValue = ""

			' eventcalendar_intro
			eventcalendar_th.eventcalendar_intro.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_intro.HrefValue = ""
			eventcalendar_th.eventcalendar_intro.TooltipValue = ""

			' eventcalendar_intro_th
			eventcalendar_th.eventcalendar_intro_th.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_intro_th.HrefValue = ""
			eventcalendar_th.eventcalendar_intro_th.TooltipValue = ""

			' eventcalendar_content
			eventcalendar_th.eventcalendar_content.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_content.HrefValue = ""
			eventcalendar_th.eventcalendar_content.TooltipValue = ""

			' eventcalendar_content_th
			eventcalendar_th.eventcalendar_content_th.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_content_th.HrefValue = ""
			eventcalendar_th.eventcalendar_content_th.TooltipValue = ""

			' eventcalendar_show_en
			eventcalendar_th.eventcalendar_show_en.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_show_en.HrefValue = ""
			eventcalendar_th.eventcalendar_show_en.TooltipValue = ""

			' eventcalendar_show
			eventcalendar_th.eventcalendar_show.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_show.HrefValue = ""
			eventcalendar_th.eventcalendar_show.TooltipValue = ""

			' eventcalendar_show_home
			eventcalendar_th.eventcalendar_show_home.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_show_home.HrefValue = ""
			eventcalendar_th.eventcalendar_show_home.TooltipValue = ""

			' eventcalendar_create
			eventcalendar_th.eventcalendar_create.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_create.HrefValue = ""
			eventcalendar_th.eventcalendar_create.TooltipValue = ""

			' eventcalendar_update
			eventcalendar_th.eventcalendar_update.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_update.HrefValue = ""
			eventcalendar_th.eventcalendar_update.TooltipValue = ""
		End If

		' Call Row Rendered event
		If eventcalendar_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call eventcalendar_th.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", eventcalendar_th.TableVar, "pom_eventcalendar_thlist.asp", eventcalendar_th.TableVar, True)
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
