<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_eventcalendarinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim eventcalendar_delete
Set eventcalendar_delete = New ceventcalendar_delete
Set Page = eventcalendar_delete

' Page init processing
eventcalendar_delete.Page_Init()

' Page main processing
eventcalendar_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
eventcalendar_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var eventcalendar_delete = new ew_Page("eventcalendar_delete");
eventcalendar_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = eventcalendar_delete.PageID; // For backward compatibility
// Form object
var feventcalendardelete = new ew_Form("feventcalendardelete");
// Form_CustomValidate event
feventcalendardelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
feventcalendardelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
feventcalendardelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set eventcalendar_delete.Recordset = eventcalendar_delete.LoadRecordset()
eventcalendar_delete.TotalRecs = eventcalendar_delete.Recordset.RecordCount ' Get record count
If eventcalendar_delete.TotalRecs <= 0 Then ' No record found, exit
	eventcalendar_delete.Recordset.Close
	Set eventcalendar_delete.Recordset = Nothing
	Call eventcalendar_delete.Page_Terminate("pom_eventcalendarlist.asp") ' Return to list
End If
%>
<% If eventcalendar.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% eventcalendar_delete.ShowPageHeader() %>
<% eventcalendar_delete.ShowMessage %>
<form name="feventcalendardelete" id="feventcalendardelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="eventcalendar">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(eventcalendar_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(eventcalendar_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_eventcalendardelete" class="ewTable ewTableSeparate">
<%= eventcalendar.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If eventcalendar.eventcalendar_id.Visible Then ' eventcalendar_id %>
		<td><span id="elh_eventcalendar_eventcalendar_id" class="eventcalendar_eventcalendar_id"><%= eventcalendar.eventcalendar_id.FldCaption %></span></td>
<% End If %>
<% If eventcalendar.eventcalendar_date.Visible Then ' eventcalendar_date %>
		<td><span id="elh_eventcalendar_eventcalendar_date" class="eventcalendar_eventcalendar_date"><%= eventcalendar.eventcalendar_date.FldCaption %></span></td>
<% End If %>
<% If eventcalendar.eventcalendar_category.Visible Then ' eventcalendar_category %>
		<td><span id="elh_eventcalendar_eventcalendar_category" class="eventcalendar_eventcalendar_category"><%= eventcalendar.eventcalendar_category.FldCaption %></span></td>
<% End If %>
<% If eventcalendar.eventcalendar_category_sub.Visible Then ' eventcalendar_category_sub %>
		<td><span id="elh_eventcalendar_eventcalendar_category_sub" class="eventcalendar_eventcalendar_category_sub"><%= eventcalendar.eventcalendar_category_sub.FldCaption %></span></td>
<% End If %>
<% If eventcalendar.start_date.Visible Then ' start_date %>
		<td><span id="elh_eventcalendar_start_date" class="eventcalendar_start_date"><%= eventcalendar.start_date.FldCaption %></span></td>
<% End If %>
<% If eventcalendar.end_date.Visible Then ' end_date %>
		<td><span id="elh_eventcalendar_end_date" class="eventcalendar_end_date"><%= eventcalendar.end_date.FldCaption %></span></td>
<% End If %>
<% If eventcalendar.eventcalendar_pdf.Visible Then ' eventcalendar_pdf %>
		<td><span id="elh_eventcalendar_eventcalendar_pdf" class="eventcalendar_eventcalendar_pdf"><%= eventcalendar.eventcalendar_pdf.FldCaption %></span></td>
<% End If %>
<% If eventcalendar.eventcalendar_subject.Visible Then ' eventcalendar_subject %>
		<td><span id="elh_eventcalendar_eventcalendar_subject" class="eventcalendar_eventcalendar_subject"><%= eventcalendar.eventcalendar_subject.FldCaption %></span></td>
<% End If %>
<% If eventcalendar.eventcalendar_subject_th.Visible Then ' eventcalendar_subject_th %>
		<td><span id="elh_eventcalendar_eventcalendar_subject_th" class="eventcalendar_eventcalendar_subject_th"><%= eventcalendar.eventcalendar_subject_th.FldCaption %></span></td>
<% End If %>
<% If eventcalendar.eventcalendar_show_en.Visible Then ' eventcalendar_show_en %>
		<td><span id="elh_eventcalendar_eventcalendar_show_en" class="eventcalendar_eventcalendar_show_en"><%= eventcalendar.eventcalendar_show_en.FldCaption %></span></td>
<% End If %>
<% If eventcalendar.eventcalendar_show.Visible Then ' eventcalendar_show %>
		<td><span id="elh_eventcalendar_eventcalendar_show" class="eventcalendar_eventcalendar_show"><%= eventcalendar.eventcalendar_show.FldCaption %></span></td>
<% End If %>
<% If eventcalendar.eventcalendar_show_home.Visible Then ' eventcalendar_show_home %>
		<td><span id="elh_eventcalendar_eventcalendar_show_home" class="eventcalendar_eventcalendar_show_home"><%= eventcalendar.eventcalendar_show_home.FldCaption %></span></td>
<% End If %>
<% If eventcalendar.eventcalendar_create.Visible Then ' eventcalendar_create %>
		<td><span id="elh_eventcalendar_eventcalendar_create" class="eventcalendar_eventcalendar_create"><%= eventcalendar.eventcalendar_create.FldCaption %></span></td>
<% End If %>
<% If eventcalendar.eventcalendar_update.Visible Then ' eventcalendar_update %>
		<td><span id="elh_eventcalendar_eventcalendar_update" class="eventcalendar_eventcalendar_update"><%= eventcalendar.eventcalendar_update.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
eventcalendar_delete.RecCnt = 0
eventcalendar_delete.RowCnt = 0
Do While (Not eventcalendar_delete.Recordset.Eof)
	eventcalendar_delete.RecCnt = eventcalendar_delete.RecCnt + 1
	eventcalendar_delete.RowCnt = eventcalendar_delete.RowCnt + 1

	' Set row properties
	Call eventcalendar.ResetAttrs()
	eventcalendar.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call eventcalendar_delete.LoadRowValues(eventcalendar_delete.Recordset)

	' Render row
	Call eventcalendar_delete.RenderRow()
%>
	<tr<%= eventcalendar.RowAttributes %>>
<% If eventcalendar.eventcalendar_id.Visible Then ' eventcalendar_id %>
		<td<%= eventcalendar.eventcalendar_id.CellAttributes %>>
<span id="el<%= eventcalendar_delete.RowCnt %>_eventcalendar_eventcalendar_id" class="control-group eventcalendar_eventcalendar_id">
<span<%= eventcalendar.eventcalendar_id.ViewAttributes %>>
<%= eventcalendar.eventcalendar_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If eventcalendar.eventcalendar_date.Visible Then ' eventcalendar_date %>
		<td<%= eventcalendar.eventcalendar_date.CellAttributes %>>
<span id="el<%= eventcalendar_delete.RowCnt %>_eventcalendar_eventcalendar_date" class="control-group eventcalendar_eventcalendar_date">
<span<%= eventcalendar.eventcalendar_date.ViewAttributes %>>
<%= eventcalendar.eventcalendar_date.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If eventcalendar.eventcalendar_category.Visible Then ' eventcalendar_category %>
		<td<%= eventcalendar.eventcalendar_category.CellAttributes %>>
<span id="el<%= eventcalendar_delete.RowCnt %>_eventcalendar_eventcalendar_category" class="control-group eventcalendar_eventcalendar_category">
<span<%= eventcalendar.eventcalendar_category.ViewAttributes %>>
<%= eventcalendar.eventcalendar_category.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If eventcalendar.eventcalendar_category_sub.Visible Then ' eventcalendar_category_sub %>
		<td<%= eventcalendar.eventcalendar_category_sub.CellAttributes %>>
<span id="el<%= eventcalendar_delete.RowCnt %>_eventcalendar_eventcalendar_category_sub" class="control-group eventcalendar_eventcalendar_category_sub">
<span<%= eventcalendar.eventcalendar_category_sub.ViewAttributes %>>
<%= eventcalendar.eventcalendar_category_sub.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If eventcalendar.start_date.Visible Then ' start_date %>
		<td<%= eventcalendar.start_date.CellAttributes %>>
<span id="el<%= eventcalendar_delete.RowCnt %>_eventcalendar_start_date" class="control-group eventcalendar_start_date">
<span<%= eventcalendar.start_date.ViewAttributes %>>
<%= eventcalendar.start_date.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If eventcalendar.end_date.Visible Then ' end_date %>
		<td<%= eventcalendar.end_date.CellAttributes %>>
<span id="el<%= eventcalendar_delete.RowCnt %>_eventcalendar_end_date" class="control-group eventcalendar_end_date">
<span<%= eventcalendar.end_date.ViewAttributes %>>
<%= eventcalendar.end_date.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If eventcalendar.eventcalendar_pdf.Visible Then ' eventcalendar_pdf %>
		<td<%= eventcalendar.eventcalendar_pdf.CellAttributes %>>
<span id="el<%= eventcalendar_delete.RowCnt %>_eventcalendar_eventcalendar_pdf" class="control-group eventcalendar_eventcalendar_pdf">
<span<%= eventcalendar.eventcalendar_pdf.ViewAttributes %>>
<%= eventcalendar.eventcalendar_pdf.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If eventcalendar.eventcalendar_subject.Visible Then ' eventcalendar_subject %>
		<td<%= eventcalendar.eventcalendar_subject.CellAttributes %>>
<span id="el<%= eventcalendar_delete.RowCnt %>_eventcalendar_eventcalendar_subject" class="control-group eventcalendar_eventcalendar_subject">
<span<%= eventcalendar.eventcalendar_subject.ViewAttributes %>>
<%= eventcalendar.eventcalendar_subject.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If eventcalendar.eventcalendar_subject_th.Visible Then ' eventcalendar_subject_th %>
		<td<%= eventcalendar.eventcalendar_subject_th.CellAttributes %>>
<span id="el<%= eventcalendar_delete.RowCnt %>_eventcalendar_eventcalendar_subject_th" class="control-group eventcalendar_eventcalendar_subject_th">
<span<%= eventcalendar.eventcalendar_subject_th.ViewAttributes %>>
<%= eventcalendar.eventcalendar_subject_th.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If eventcalendar.eventcalendar_show_en.Visible Then ' eventcalendar_show_en %>
		<td<%= eventcalendar.eventcalendar_show_en.CellAttributes %>>
<span id="el<%= eventcalendar_delete.RowCnt %>_eventcalendar_eventcalendar_show_en" class="control-group eventcalendar_eventcalendar_show_en">
<span<%= eventcalendar.eventcalendar_show_en.ViewAttributes %>>
<%= eventcalendar.eventcalendar_show_en.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If eventcalendar.eventcalendar_show.Visible Then ' eventcalendar_show %>
		<td<%= eventcalendar.eventcalendar_show.CellAttributes %>>
<span id="el<%= eventcalendar_delete.RowCnt %>_eventcalendar_eventcalendar_show" class="control-group eventcalendar_eventcalendar_show">
<span<%= eventcalendar.eventcalendar_show.ViewAttributes %>>
<%= eventcalendar.eventcalendar_show.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If eventcalendar.eventcalendar_show_home.Visible Then ' eventcalendar_show_home %>
		<td<%= eventcalendar.eventcalendar_show_home.CellAttributes %>>
<span id="el<%= eventcalendar_delete.RowCnt %>_eventcalendar_eventcalendar_show_home" class="control-group eventcalendar_eventcalendar_show_home">
<span<%= eventcalendar.eventcalendar_show_home.ViewAttributes %>>
<%= eventcalendar.eventcalendar_show_home.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If eventcalendar.eventcalendar_create.Visible Then ' eventcalendar_create %>
		<td<%= eventcalendar.eventcalendar_create.CellAttributes %>>
<span id="el<%= eventcalendar_delete.RowCnt %>_eventcalendar_eventcalendar_create" class="control-group eventcalendar_eventcalendar_create">
<span<%= eventcalendar.eventcalendar_create.ViewAttributes %>>
<%= eventcalendar.eventcalendar_create.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If eventcalendar.eventcalendar_update.Visible Then ' eventcalendar_update %>
		<td<%= eventcalendar.eventcalendar_update.CellAttributes %>>
<span id="el<%= eventcalendar_delete.RowCnt %>_eventcalendar_eventcalendar_update" class="control-group eventcalendar_eventcalendar_update">
<span<%= eventcalendar.eventcalendar_update.ViewAttributes %>>
<%= eventcalendar.eventcalendar_update.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	eventcalendar_delete.Recordset.MoveNext
Loop
eventcalendar_delete.Recordset.Close
Set eventcalendar_delete.Recordset = Nothing
%>
	</tbody>
</table>
</div>
</td></tr></table>
<div class="btn-group ewButtonGroup">
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("DeleteBtn") %></button>
</div>
</form>
<script type="text/javascript">
feventcalendardelete.Init();
</script>
<%
eventcalendar_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set eventcalendar_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class ceventcalendar_delete

	' Page ID
	Public Property Get PageID()
		PageID = "delete"
	End Property

	' Project ID
	Public Property Get ProjectID()
		ProjectID = "{324ED72D-DE20-46F7-B12E-7AF8CE8711A6}"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "eventcalendar"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "eventcalendar_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If eventcalendar.UseTokenInUrl Then PageUrl = PageUrl & "t=" & eventcalendar.TableVar & "&" ' add page token
	End Property

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
		If eventcalendar.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (eventcalendar.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (eventcalendar.TableVar = Request.QueryString("t"))
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
		If IsEmpty(eventcalendar) Then Set eventcalendar = New ceventcalendar
		Set Table = eventcalendar

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "eventcalendar"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()
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
		Set eventcalendar = Nothing
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

	Dim TotalRecs
	Dim RecCnt
	Dim RecKeys
	Dim Recordset
	Dim StartRowCnt
	Dim RowCnt

	' Page main processing
	Sub Page_Main()
		Dim sFilter
		StartRowCnt = 1

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Load Key Parameters
		RecKeys = eventcalendar.GetRecordKeys() ' Load record keys
		sFilter = eventcalendar.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_eventcalendarlist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in eventcalendar class, eventcalendarinfo.asp

		eventcalendar.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			eventcalendar.CurrentAction = Request.Form("a_delete")
		Else
			eventcalendar.CurrentAction = "I"	' Display record
		End If
		Select Case eventcalendar.CurrentAction
			Case "D" ' Delete
				eventcalendar.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(eventcalendar.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = eventcalendar.CurrentFilter
		Call eventcalendar.Recordset_Selecting(sFilter)
		eventcalendar.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = eventcalendar.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call eventcalendar.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = eventcalendar.KeyFilter

		' Call Row Selecting event
		Call eventcalendar.Row_Selecting(sFilter)

		' Load sql based on filter
		eventcalendar.CurrentFilter = sFilter
		sSql = eventcalendar.SQL
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
		Call eventcalendar.Row_Selected(RsRow)
		eventcalendar.eventcalendar_id.DbValue = RsRow("eventcalendar_id")
		eventcalendar.eventcalendar_img.DbValue = RsRow("eventcalendar_img")
		eventcalendar.eventcalendar_date.DbValue = RsRow("eventcalendar_date")
		eventcalendar.eventcalendar_category.DbValue = RsRow("eventcalendar_category")
		eventcalendar.eventcalendar_category_sub.DbValue = RsRow("eventcalendar_category_sub")
		eventcalendar.start_date.DbValue = RsRow("start_date")
		eventcalendar.end_date.DbValue = RsRow("end_date")
		eventcalendar.eventcalendar_pdf.DbValue = RsRow("eventcalendar_pdf")
		eventcalendar.eventcalendar_subject.DbValue = RsRow("eventcalendar_subject")
		eventcalendar.eventcalendar_subject_th.DbValue = RsRow("eventcalendar_subject_th")
		eventcalendar.eventcalendar_intro.DbValue = RsRow("eventcalendar_intro")
		eventcalendar.eventcalendar_intro_th.DbValue = RsRow("eventcalendar_intro_th")
		eventcalendar.eventcalendar_content.DbValue = RsRow("eventcalendar_content")
		eventcalendar.eventcalendar_content_th.DbValue = RsRow("eventcalendar_content_th")
		eventcalendar.eventcalendar_show_en.DbValue = RsRow("eventcalendar_show_en")
		eventcalendar.eventcalendar_show.DbValue = RsRow("eventcalendar_show")
		eventcalendar.eventcalendar_show_home.DbValue = RsRow("eventcalendar_show_home")
		eventcalendar.eventcalendar_create.DbValue = RsRow("eventcalendar_create")
		eventcalendar.eventcalendar_update.DbValue = RsRow("eventcalendar_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		eventcalendar.eventcalendar_id.m_DbValue = Rs("eventcalendar_id")
		eventcalendar.eventcalendar_img.m_DbValue = Rs("eventcalendar_img")
		eventcalendar.eventcalendar_date.m_DbValue = Rs("eventcalendar_date")
		eventcalendar.eventcalendar_category.m_DbValue = Rs("eventcalendar_category")
		eventcalendar.eventcalendar_category_sub.m_DbValue = Rs("eventcalendar_category_sub")
		eventcalendar.start_date.m_DbValue = Rs("start_date")
		eventcalendar.end_date.m_DbValue = Rs("end_date")
		eventcalendar.eventcalendar_pdf.m_DbValue = Rs("eventcalendar_pdf")
		eventcalendar.eventcalendar_subject.m_DbValue = Rs("eventcalendar_subject")
		eventcalendar.eventcalendar_subject_th.m_DbValue = Rs("eventcalendar_subject_th")
		eventcalendar.eventcalendar_intro.m_DbValue = Rs("eventcalendar_intro")
		eventcalendar.eventcalendar_intro_th.m_DbValue = Rs("eventcalendar_intro_th")
		eventcalendar.eventcalendar_content.m_DbValue = Rs("eventcalendar_content")
		eventcalendar.eventcalendar_content_th.m_DbValue = Rs("eventcalendar_content_th")
		eventcalendar.eventcalendar_show_en.m_DbValue = Rs("eventcalendar_show_en")
		eventcalendar.eventcalendar_show.m_DbValue = Rs("eventcalendar_show")
		eventcalendar.eventcalendar_show_home.m_DbValue = Rs("eventcalendar_show_home")
		eventcalendar.eventcalendar_create.m_DbValue = Rs("eventcalendar_create")
		eventcalendar.eventcalendar_update.m_DbValue = Rs("eventcalendar_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call eventcalendar.Row_Rendering()

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

		If eventcalendar.RowType = EW_ROWTYPE_VIEW Then ' View row

			' eventcalendar_id
			eventcalendar.eventcalendar_id.ViewValue = eventcalendar.eventcalendar_id.CurrentValue
			eventcalendar.eventcalendar_id.ViewCustomAttributes = ""

			' eventcalendar_date
			eventcalendar.eventcalendar_date.ViewValue = eventcalendar.eventcalendar_date.CurrentValue
			eventcalendar.eventcalendar_date.ViewCustomAttributes = ""

			' eventcalendar_category
			eventcalendar.eventcalendar_category.ViewValue = eventcalendar.eventcalendar_category.CurrentValue
			eventcalendar.eventcalendar_category.ViewCustomAttributes = ""

			' eventcalendar_category_sub
			eventcalendar.eventcalendar_category_sub.ViewValue = eventcalendar.eventcalendar_category_sub.CurrentValue
			eventcalendar.eventcalendar_category_sub.ViewCustomAttributes = ""

			' start_date
			eventcalendar.start_date.ViewValue = eventcalendar.start_date.CurrentValue
			eventcalendar.start_date.ViewCustomAttributes = ""

			' end_date
			eventcalendar.end_date.ViewValue = eventcalendar.end_date.CurrentValue
			eventcalendar.end_date.ViewCustomAttributes = ""

			' eventcalendar_pdf
			eventcalendar.eventcalendar_pdf.ViewValue = eventcalendar.eventcalendar_pdf.CurrentValue
			eventcalendar.eventcalendar_pdf.ViewCustomAttributes = ""

			' eventcalendar_subject
			eventcalendar.eventcalendar_subject.ViewValue = eventcalendar.eventcalendar_subject.CurrentValue
			eventcalendar.eventcalendar_subject.ViewCustomAttributes = ""

			' eventcalendar_subject_th
			eventcalendar.eventcalendar_subject_th.ViewValue = eventcalendar.eventcalendar_subject_th.CurrentValue
			eventcalendar.eventcalendar_subject_th.ViewCustomAttributes = ""

			' eventcalendar_show_en
			eventcalendar.eventcalendar_show_en.ViewValue = eventcalendar.eventcalendar_show_en.CurrentValue
			eventcalendar.eventcalendar_show_en.ViewCustomAttributes = ""

			' eventcalendar_show
			eventcalendar.eventcalendar_show.ViewValue = eventcalendar.eventcalendar_show.CurrentValue
			eventcalendar.eventcalendar_show.ViewCustomAttributes = ""

			' eventcalendar_show_home
			eventcalendar.eventcalendar_show_home.ViewValue = eventcalendar.eventcalendar_show_home.CurrentValue
			eventcalendar.eventcalendar_show_home.ViewCustomAttributes = ""

			' eventcalendar_create
			eventcalendar.eventcalendar_create.ViewValue = eventcalendar.eventcalendar_create.CurrentValue
			eventcalendar.eventcalendar_create.ViewCustomAttributes = ""

			' eventcalendar_update
			eventcalendar.eventcalendar_update.ViewValue = eventcalendar.eventcalendar_update.CurrentValue
			eventcalendar.eventcalendar_update.ViewCustomAttributes = ""

			' View refer script
			' eventcalendar_id

			eventcalendar.eventcalendar_id.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_id.HrefValue = ""
			eventcalendar.eventcalendar_id.TooltipValue = ""

			' eventcalendar_date
			eventcalendar.eventcalendar_date.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_date.HrefValue = ""
			eventcalendar.eventcalendar_date.TooltipValue = ""

			' eventcalendar_category
			eventcalendar.eventcalendar_category.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_category.HrefValue = ""
			eventcalendar.eventcalendar_category.TooltipValue = ""

			' eventcalendar_category_sub
			eventcalendar.eventcalendar_category_sub.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_category_sub.HrefValue = ""
			eventcalendar.eventcalendar_category_sub.TooltipValue = ""

			' start_date
			eventcalendar.start_date.LinkCustomAttributes = ""
			eventcalendar.start_date.HrefValue = ""
			eventcalendar.start_date.TooltipValue = ""

			' end_date
			eventcalendar.end_date.LinkCustomAttributes = ""
			eventcalendar.end_date.HrefValue = ""
			eventcalendar.end_date.TooltipValue = ""

			' eventcalendar_pdf
			eventcalendar.eventcalendar_pdf.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_pdf.HrefValue = ""
			eventcalendar.eventcalendar_pdf.TooltipValue = ""

			' eventcalendar_subject
			eventcalendar.eventcalendar_subject.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_subject.HrefValue = ""
			eventcalendar.eventcalendar_subject.TooltipValue = ""

			' eventcalendar_subject_th
			eventcalendar.eventcalendar_subject_th.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_subject_th.HrefValue = ""
			eventcalendar.eventcalendar_subject_th.TooltipValue = ""

			' eventcalendar_show_en
			eventcalendar.eventcalendar_show_en.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_show_en.HrefValue = ""
			eventcalendar.eventcalendar_show_en.TooltipValue = ""

			' eventcalendar_show
			eventcalendar.eventcalendar_show.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_show.HrefValue = ""
			eventcalendar.eventcalendar_show.TooltipValue = ""

			' eventcalendar_show_home
			eventcalendar.eventcalendar_show_home.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_show_home.HrefValue = ""
			eventcalendar.eventcalendar_show_home.TooltipValue = ""

			' eventcalendar_create
			eventcalendar.eventcalendar_create.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_create.HrefValue = ""
			eventcalendar.eventcalendar_create.TooltipValue = ""

			' eventcalendar_update
			eventcalendar.eventcalendar_update.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_update.HrefValue = ""
			eventcalendar.eventcalendar_update.TooltipValue = ""
		End If

		' Call Row Rendered event
		If eventcalendar.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call eventcalendar.Row_Rendered()
		End If
	End Sub

	'
	' Delete records based on current filter
	'
	Function DeleteRows()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim sKey, sThisKey, sKeyFld, arKeyFlds
		Dim sSql, RsDelete
		Dim RsOld, RsDetail
		DeleteRows = True
		sSql = eventcalendar.SQL
		Conn.BeginTrans
		Set RsDelete = Server.CreateObject("ADODB.Recordset")
		RsDelete.CursorLocation = EW_CURSORLOCATION
		RsDelete.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			RsDelete.Close
			Set RsDelete = Nothing
			DeleteRows = False
			Exit Function
		ElseIf RsDelete.Eof Then
			FailureMessage = Language.Phrase("NoRecord") ' No record found
			RsDelete.Close
			Set RsDelete = Nothing
			DeleteRows = False
			Exit Function
		End If

		' Clone old recordset object
		Set RsOld = ew_CloneRs(RsDelete)

		' Call row deleting event
		If DeleteRows Then
			RsDelete.MoveFirst
			Do While Not RsDelete.Eof
				DeleteRows = eventcalendar.Row_Deleting(RsDelete)
				If Not DeleteRows Then Exit Do
				RsDelete.MoveNext
			Loop
			RsDelete.MoveFirst
		End If
		If DeleteRows Then
			sKey = ""
			RsDelete.MoveFirst
			Do While Not RsDelete.Eof
				sThisKey = ""
				If sThisKey <> "" Then sThisKey = sThisKey & EW_COMPOSITE_KEY_SEPARATOR
				sThisKey = sThisKey & RsDelete("eventcalendar_id")
				Call LoadDbValues(RsDelete)
				If DeleteRows Then
					RsDelete.Delete
				End If
				If Err.Number <> 0 Or Not DeleteRows Then
					If Err.Description <> "" Then FailureMessage = Err.Description ' Set up error message
					DeleteRows = False
					Exit Do
				End If
				If sKey <> "" Then sKey = sKey & ", "
				sKey = sKey & sThisKey
				RsDelete.MoveNext
			Loop
		Else

			' Set up error message
			If SuccessMessage <> "" Or FailureMessage <> "" Then

				' Use the message, do nothing
			ElseIf eventcalendar.CancelMessage <> "" Then
				FailureMessage = eventcalendar.CancelMessage
				eventcalendar.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("DeleteCancelled")
			End If
		End If
		If DeleteRows Then
			Conn.CommitTrans ' Commit the changes
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				DeleteRows = False ' Delete failed
			End If
		Else
			Conn.RollbackTrans ' Rollback changes
		End If
		RsDelete.Close
		Set RsDelete = Nothing

		' Call row deleting event
		If DeleteRows Then
			RsOld.MoveFirst
			Do While Not RsOld.Eof
				Call eventcalendar.Row_Deleted(RsOld)
				RsOld.MoveNext
			Loop
		End If
		RsOld.Close
		Set RsOld = Nothing
	End Function

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", eventcalendar.TableVar, "pom_eventcalendarlist.asp", eventcalendar.TableVar, True)
		PageId = "delete"
		Call Breadcrumb.Add("delete", PageId, ew_CurrentUrl, "", False)
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
