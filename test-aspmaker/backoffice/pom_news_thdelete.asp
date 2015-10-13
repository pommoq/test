<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_news_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim news_th_delete
Set news_th_delete = New cnews_th_delete
Set Page = news_th_delete

' Page init processing
news_th_delete.Page_Init()

' Page main processing
news_th_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
news_th_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var news_th_delete = new ew_Page("news_th_delete");
news_th_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = news_th_delete.PageID; // For backward compatibility
// Form object
var fnews_thdelete = new ew_Form("fnews_thdelete");
// Form_CustomValidate event
fnews_thdelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fnews_thdelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fnews_thdelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set news_th_delete.Recordset = news_th_delete.LoadRecordset()
news_th_delete.TotalRecs = news_th_delete.Recordset.RecordCount ' Get record count
If news_th_delete.TotalRecs <= 0 Then ' No record found, exit
	news_th_delete.Recordset.Close
	Set news_th_delete.Recordset = Nothing
	Call news_th_delete.Page_Terminate("pom_news_thlist.asp") ' Return to list
End If
%>
<% If news_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% news_th_delete.ShowPageHeader() %>
<% news_th_delete.ShowMessage %>
<form name="fnews_thdelete" id="fnews_thdelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="news_th">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(news_th_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(news_th_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_news_thdelete" class="ewTable ewTableSeparate">
<%= news_th.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If news_th.news_id.Visible Then ' news_id %>
		<td><span id="elh_news_th_news_id" class="news_th_news_id"><%= news_th.news_id.FldCaption %></span></td>
<% End If %>
<% If news_th.news_img.Visible Then ' news_img %>
		<td><span id="elh_news_th_news_img" class="news_th_news_img"><%= news_th.news_img.FldCaption %></span></td>
<% End If %>
<% If news_th.news_date.Visible Then ' news_date %>
		<td><span id="elh_news_th_news_date" class="news_th_news_date"><%= news_th.news_date.FldCaption %></span></td>
<% End If %>
<% If news_th.news_category.Visible Then ' news_category %>
		<td><span id="elh_news_th_news_category" class="news_th_news_category"><%= news_th.news_category.FldCaption %></span></td>
<% End If %>
<% If news_th.news_category_sub.Visible Then ' news_category_sub %>
		<td><span id="elh_news_th_news_category_sub" class="news_th_news_category_sub"><%= news_th.news_category_sub.FldCaption %></span></td>
<% End If %>
<% If news_th.start_date.Visible Then ' start_date %>
		<td><span id="elh_news_th_start_date" class="news_th_start_date"><%= news_th.start_date.FldCaption %></span></td>
<% End If %>
<% If news_th.end_date.Visible Then ' end_date %>
		<td><span id="elh_news_th_end_date" class="news_th_end_date"><%= news_th.end_date.FldCaption %></span></td>
<% End If %>
<% If news_th.news_pdf.Visible Then ' news_pdf %>
		<td><span id="elh_news_th_news_pdf" class="news_th_news_pdf"><%= news_th.news_pdf.FldCaption %></span></td>
<% End If %>
<% If news_th.news_subject.Visible Then ' news_subject %>
		<td><span id="elh_news_th_news_subject" class="news_th_news_subject"><%= news_th.news_subject.FldCaption %></span></td>
<% End If %>
<% If news_th.news_subject_th.Visible Then ' news_subject_th %>
		<td><span id="elh_news_th_news_subject_th" class="news_th_news_subject_th"><%= news_th.news_subject_th.FldCaption %></span></td>
<% End If %>
<% If news_th.news_show_en.Visible Then ' news_show_en %>
		<td><span id="elh_news_th_news_show_en" class="news_th_news_show_en"><%= news_th.news_show_en.FldCaption %></span></td>
<% End If %>
<% If news_th.news_show.Visible Then ' news_show %>
		<td><span id="elh_news_th_news_show" class="news_th_news_show"><%= news_th.news_show.FldCaption %></span></td>
<% End If %>
<% If news_th.news_show_home.Visible Then ' news_show_home %>
		<td><span id="elh_news_th_news_show_home" class="news_th_news_show_home"><%= news_th.news_show_home.FldCaption %></span></td>
<% End If %>
<% If news_th.news_create.Visible Then ' news_create %>
		<td><span id="elh_news_th_news_create" class="news_th_news_create"><%= news_th.news_create.FldCaption %></span></td>
<% End If %>
<% If news_th.news_update.Visible Then ' news_update %>
		<td><span id="elh_news_th_news_update" class="news_th_news_update"><%= news_th.news_update.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
news_th_delete.RecCnt = 0
news_th_delete.RowCnt = 0
Do While (Not news_th_delete.Recordset.Eof)
	news_th_delete.RecCnt = news_th_delete.RecCnt + 1
	news_th_delete.RowCnt = news_th_delete.RowCnt + 1

	' Set row properties
	Call news_th.ResetAttrs()
	news_th.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call news_th_delete.LoadRowValues(news_th_delete.Recordset)

	' Render row
	Call news_th_delete.RenderRow()
%>
	<tr<%= news_th.RowAttributes %>>
<% If news_th.news_id.Visible Then ' news_id %>
		<td<%= news_th.news_id.CellAttributes %>>
<span id="el<%= news_th_delete.RowCnt %>_news_th_news_id" class="control-group news_th_news_id">
<span<%= news_th.news_id.ViewAttributes %>>
<%= news_th.news_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_th.news_img.Visible Then ' news_img %>
		<td<%= news_th.news_img.CellAttributes %>>
<span id="el<%= news_th_delete.RowCnt %>_news_th_news_img" class="control-group news_th_news_img">
<span<%= news_th.news_img.ViewAttributes %>>
<%= news_th.news_img.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_th.news_date.Visible Then ' news_date %>
		<td<%= news_th.news_date.CellAttributes %>>
<span id="el<%= news_th_delete.RowCnt %>_news_th_news_date" class="control-group news_th_news_date">
<span<%= news_th.news_date.ViewAttributes %>>
<%= news_th.news_date.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_th.news_category.Visible Then ' news_category %>
		<td<%= news_th.news_category.CellAttributes %>>
<span id="el<%= news_th_delete.RowCnt %>_news_th_news_category" class="control-group news_th_news_category">
<span<%= news_th.news_category.ViewAttributes %>>
<%= news_th.news_category.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_th.news_category_sub.Visible Then ' news_category_sub %>
		<td<%= news_th.news_category_sub.CellAttributes %>>
<span id="el<%= news_th_delete.RowCnt %>_news_th_news_category_sub" class="control-group news_th_news_category_sub">
<span<%= news_th.news_category_sub.ViewAttributes %>>
<%= news_th.news_category_sub.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_th.start_date.Visible Then ' start_date %>
		<td<%= news_th.start_date.CellAttributes %>>
<span id="el<%= news_th_delete.RowCnt %>_news_th_start_date" class="control-group news_th_start_date">
<span<%= news_th.start_date.ViewAttributes %>>
<%= news_th.start_date.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_th.end_date.Visible Then ' end_date %>
		<td<%= news_th.end_date.CellAttributes %>>
<span id="el<%= news_th_delete.RowCnt %>_news_th_end_date" class="control-group news_th_end_date">
<span<%= news_th.end_date.ViewAttributes %>>
<%= news_th.end_date.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_th.news_pdf.Visible Then ' news_pdf %>
		<td<%= news_th.news_pdf.CellAttributes %>>
<span id="el<%= news_th_delete.RowCnt %>_news_th_news_pdf" class="control-group news_th_news_pdf">
<span<%= news_th.news_pdf.ViewAttributes %>>
<%= news_th.news_pdf.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_th.news_subject.Visible Then ' news_subject %>
		<td<%= news_th.news_subject.CellAttributes %>>
<span id="el<%= news_th_delete.RowCnt %>_news_th_news_subject" class="control-group news_th_news_subject">
<span<%= news_th.news_subject.ViewAttributes %>>
<%= news_th.news_subject.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_th.news_subject_th.Visible Then ' news_subject_th %>
		<td<%= news_th.news_subject_th.CellAttributes %>>
<span id="el<%= news_th_delete.RowCnt %>_news_th_news_subject_th" class="control-group news_th_news_subject_th">
<span<%= news_th.news_subject_th.ViewAttributes %>>
<%= news_th.news_subject_th.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_th.news_show_en.Visible Then ' news_show_en %>
		<td<%= news_th.news_show_en.CellAttributes %>>
<span id="el<%= news_th_delete.RowCnt %>_news_th_news_show_en" class="control-group news_th_news_show_en">
<span<%= news_th.news_show_en.ViewAttributes %>>
<%= news_th.news_show_en.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_th.news_show.Visible Then ' news_show %>
		<td<%= news_th.news_show.CellAttributes %>>
<span id="el<%= news_th_delete.RowCnt %>_news_th_news_show" class="control-group news_th_news_show">
<span<%= news_th.news_show.ViewAttributes %>>
<%= news_th.news_show.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_th.news_show_home.Visible Then ' news_show_home %>
		<td<%= news_th.news_show_home.CellAttributes %>>
<span id="el<%= news_th_delete.RowCnt %>_news_th_news_show_home" class="control-group news_th_news_show_home">
<span<%= news_th.news_show_home.ViewAttributes %>>
<%= news_th.news_show_home.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_th.news_create.Visible Then ' news_create %>
		<td<%= news_th.news_create.CellAttributes %>>
<span id="el<%= news_th_delete.RowCnt %>_news_th_news_create" class="control-group news_th_news_create">
<span<%= news_th.news_create.ViewAttributes %>>
<%= news_th.news_create.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_th.news_update.Visible Then ' news_update %>
		<td<%= news_th.news_update.CellAttributes %>>
<span id="el<%= news_th_delete.RowCnt %>_news_th_news_update" class="control-group news_th_news_update">
<span<%= news_th.news_update.ViewAttributes %>>
<%= news_th.news_update.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	news_th_delete.Recordset.MoveNext
Loop
news_th_delete.Recordset.Close
Set news_th_delete.Recordset = Nothing
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
fnews_thdelete.Init();
</script>
<%
news_th_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set news_th_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cnews_th_delete

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
		TableName = "news_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "news_th_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If news_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & news_th.TableVar & "&" ' add page token
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
		If news_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (news_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (news_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(news_th) Then Set news_th = New cnews_th
		Set Table = news_th

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "news_th"

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
		Set news_th = Nothing
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
		RecKeys = news_th.GetRecordKeys() ' Load record keys
		sFilter = news_th.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_news_thlist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in news_th class, news_thinfo.asp

		news_th.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			news_th.CurrentAction = Request.Form("a_delete")
		Else
			news_th.CurrentAction = "I"	' Display record
		End If
		Select Case news_th.CurrentAction
			Case "D" ' Delete
				news_th.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(news_th.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = news_th.CurrentFilter
		Call news_th.Recordset_Selecting(sFilter)
		news_th.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = news_th.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call news_th.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = news_th.KeyFilter

		' Call Row Selecting event
		Call news_th.Row_Selecting(sFilter)

		' Load sql based on filter
		news_th.CurrentFilter = sFilter
		sSql = news_th.SQL
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
		Call news_th.Row_Selected(RsRow)
		news_th.news_id.DbValue = RsRow("news_id")
		news_th.news_img.DbValue = RsRow("news_img")
		news_th.news_date.DbValue = RsRow("news_date")
		news_th.news_category.DbValue = RsRow("news_category")
		news_th.news_category_sub.DbValue = RsRow("news_category_sub")
		news_th.start_date.DbValue = RsRow("start_date")
		news_th.end_date.DbValue = RsRow("end_date")
		news_th.news_pdf.DbValue = RsRow("news_pdf")
		news_th.news_subject.DbValue = RsRow("news_subject")
		news_th.news_subject_th.DbValue = RsRow("news_subject_th")
		news_th.news_intro.DbValue = RsRow("news_intro")
		news_th.news_intro_th.DbValue = RsRow("news_intro_th")
		news_th.news_content.DbValue = RsRow("news_content")
		news_th.news_content_th.DbValue = RsRow("news_content_th")
		news_th.news_show_en.DbValue = RsRow("news_show_en")
		news_th.news_show.DbValue = RsRow("news_show")
		news_th.news_show_home.DbValue = RsRow("news_show_home")
		news_th.news_create.DbValue = RsRow("news_create")
		news_th.news_update.DbValue = RsRow("news_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		news_th.news_id.m_DbValue = Rs("news_id")
		news_th.news_img.m_DbValue = Rs("news_img")
		news_th.news_date.m_DbValue = Rs("news_date")
		news_th.news_category.m_DbValue = Rs("news_category")
		news_th.news_category_sub.m_DbValue = Rs("news_category_sub")
		news_th.start_date.m_DbValue = Rs("start_date")
		news_th.end_date.m_DbValue = Rs("end_date")
		news_th.news_pdf.m_DbValue = Rs("news_pdf")
		news_th.news_subject.m_DbValue = Rs("news_subject")
		news_th.news_subject_th.m_DbValue = Rs("news_subject_th")
		news_th.news_intro.m_DbValue = Rs("news_intro")
		news_th.news_intro_th.m_DbValue = Rs("news_intro_th")
		news_th.news_content.m_DbValue = Rs("news_content")
		news_th.news_content_th.m_DbValue = Rs("news_content_th")
		news_th.news_show_en.m_DbValue = Rs("news_show_en")
		news_th.news_show.m_DbValue = Rs("news_show")
		news_th.news_show_home.m_DbValue = Rs("news_show_home")
		news_th.news_create.m_DbValue = Rs("news_create")
		news_th.news_update.m_DbValue = Rs("news_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call news_th.Row_Rendering()

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

		If news_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' news_id
			news_th.news_id.ViewValue = news_th.news_id.CurrentValue
			news_th.news_id.ViewCustomAttributes = ""

			' news_img
			news_th.news_img.ViewValue = news_th.news_img.CurrentValue
			news_th.news_img.ViewCustomAttributes = ""

			' news_date
			news_th.news_date.ViewValue = news_th.news_date.CurrentValue
			news_th.news_date.ViewCustomAttributes = ""

			' news_category
			news_th.news_category.ViewValue = news_th.news_category.CurrentValue
			news_th.news_category.ViewCustomAttributes = ""

			' news_category_sub
			news_th.news_category_sub.ViewValue = news_th.news_category_sub.CurrentValue
			news_th.news_category_sub.ViewCustomAttributes = ""

			' start_date
			news_th.start_date.ViewValue = news_th.start_date.CurrentValue
			news_th.start_date.ViewCustomAttributes = ""

			' end_date
			news_th.end_date.ViewValue = news_th.end_date.CurrentValue
			news_th.end_date.ViewCustomAttributes = ""

			' news_pdf
			news_th.news_pdf.ViewValue = news_th.news_pdf.CurrentValue
			news_th.news_pdf.ViewCustomAttributes = ""

			' news_subject
			news_th.news_subject.ViewValue = news_th.news_subject.CurrentValue
			news_th.news_subject.ViewCustomAttributes = ""

			' news_subject_th
			news_th.news_subject_th.ViewValue = news_th.news_subject_th.CurrentValue
			news_th.news_subject_th.ViewCustomAttributes = ""

			' news_show_en
			news_th.news_show_en.ViewValue = news_th.news_show_en.CurrentValue
			news_th.news_show_en.ViewCustomAttributes = ""

			' news_show
			news_th.news_show.ViewValue = news_th.news_show.CurrentValue
			news_th.news_show.ViewCustomAttributes = ""

			' news_show_home
			news_th.news_show_home.ViewValue = news_th.news_show_home.CurrentValue
			news_th.news_show_home.ViewCustomAttributes = ""

			' news_create
			news_th.news_create.ViewValue = news_th.news_create.CurrentValue
			news_th.news_create.ViewCustomAttributes = ""

			' news_update
			news_th.news_update.ViewValue = news_th.news_update.CurrentValue
			news_th.news_update.ViewCustomAttributes = ""

			' View refer script
			' news_id

			news_th.news_id.LinkCustomAttributes = ""
			news_th.news_id.HrefValue = ""
			news_th.news_id.TooltipValue = ""

			' news_img
			news_th.news_img.LinkCustomAttributes = ""
			news_th.news_img.HrefValue = ""
			news_th.news_img.TooltipValue = ""

			' news_date
			news_th.news_date.LinkCustomAttributes = ""
			news_th.news_date.HrefValue = ""
			news_th.news_date.TooltipValue = ""

			' news_category
			news_th.news_category.LinkCustomAttributes = ""
			news_th.news_category.HrefValue = ""
			news_th.news_category.TooltipValue = ""

			' news_category_sub
			news_th.news_category_sub.LinkCustomAttributes = ""
			news_th.news_category_sub.HrefValue = ""
			news_th.news_category_sub.TooltipValue = ""

			' start_date
			news_th.start_date.LinkCustomAttributes = ""
			news_th.start_date.HrefValue = ""
			news_th.start_date.TooltipValue = ""

			' end_date
			news_th.end_date.LinkCustomAttributes = ""
			news_th.end_date.HrefValue = ""
			news_th.end_date.TooltipValue = ""

			' news_pdf
			news_th.news_pdf.LinkCustomAttributes = ""
			news_th.news_pdf.HrefValue = ""
			news_th.news_pdf.TooltipValue = ""

			' news_subject
			news_th.news_subject.LinkCustomAttributes = ""
			news_th.news_subject.HrefValue = ""
			news_th.news_subject.TooltipValue = ""

			' news_subject_th
			news_th.news_subject_th.LinkCustomAttributes = ""
			news_th.news_subject_th.HrefValue = ""
			news_th.news_subject_th.TooltipValue = ""

			' news_show_en
			news_th.news_show_en.LinkCustomAttributes = ""
			news_th.news_show_en.HrefValue = ""
			news_th.news_show_en.TooltipValue = ""

			' news_show
			news_th.news_show.LinkCustomAttributes = ""
			news_th.news_show.HrefValue = ""
			news_th.news_show.TooltipValue = ""

			' news_show_home
			news_th.news_show_home.LinkCustomAttributes = ""
			news_th.news_show_home.HrefValue = ""
			news_th.news_show_home.TooltipValue = ""

			' news_create
			news_th.news_create.LinkCustomAttributes = ""
			news_th.news_create.HrefValue = ""
			news_th.news_create.TooltipValue = ""

			' news_update
			news_th.news_update.LinkCustomAttributes = ""
			news_th.news_update.HrefValue = ""
			news_th.news_update.TooltipValue = ""
		End If

		' Call Row Rendered event
		If news_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call news_th.Row_Rendered()
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
		sSql = news_th.SQL
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
				DeleteRows = news_th.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("news_id")
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
			ElseIf news_th.CancelMessage <> "" Then
				FailureMessage = news_th.CancelMessage
				news_th.CancelMessage = ""
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
				Call news_th.Row_Deleted(RsOld)
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
		Call Breadcrumb.Add("list", news_th.TableVar, "pom_news_thlist.asp", news_th.TableVar, True)
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
