<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_z40newsinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim z40news_delete
Set z40news_delete = New cz40news_delete
Set Page = z40news_delete

' Page init processing
z40news_delete.Page_Init()

' Page main processing
z40news_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
z40news_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var z40news_delete = new ew_Page("z40news_delete");
z40news_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = z40news_delete.PageID; // For backward compatibility
// Form object
var fz40newsdelete = new ew_Form("fz40newsdelete");
// Form_CustomValidate event
fz40newsdelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fz40newsdelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fz40newsdelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set z40news_delete.Recordset = z40news_delete.LoadRecordset()
z40news_delete.TotalRecs = z40news_delete.Recordset.RecordCount ' Get record count
If z40news_delete.TotalRecs <= 0 Then ' No record found, exit
	z40news_delete.Recordset.Close
	Set z40news_delete.Recordset = Nothing
	Call z40news_delete.Page_Terminate("pom_z40newslist.asp") ' Return to list
End If
%>
<% If z40news.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% z40news_delete.ShowPageHeader() %>
<% z40news_delete.ShowMessage %>
<form name="fz40newsdelete" id="fz40newsdelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="z40news">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(z40news_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(z40news_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_z40newsdelete" class="ewTable ewTableSeparate">
<%= z40news.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If z40news.news_id.Visible Then ' news_id %>
		<td><span id="elh_z40news_news_id" class="z40news_news_id"><%= z40news.news_id.FldCaption %></span></td>
<% End If %>
<% If z40news.news_img.Visible Then ' news_img %>
		<td><span id="elh_z40news_news_img" class="z40news_news_img"><%= z40news.news_img.FldCaption %></span></td>
<% End If %>
<% If z40news.news_category.Visible Then ' news_category %>
		<td><span id="elh_z40news_news_category" class="z40news_news_category"><%= z40news.news_category.FldCaption %></span></td>
<% End If %>
<% If z40news.news_subject.Visible Then ' news_subject %>
		<td><span id="elh_z40news_news_subject" class="z40news_news_subject"><%= z40news.news_subject.FldCaption %></span></td>
<% End If %>
<% If z40news.news_subject_th.Visible Then ' news_subject_th %>
		<td><span id="elh_z40news_news_subject_th" class="z40news_news_subject_th"><%= z40news.news_subject_th.FldCaption %></span></td>
<% End If %>
<% If z40news.news_show_en.Visible Then ' news_show_en %>
		<td><span id="elh_z40news_news_show_en" class="z40news_news_show_en"><%= z40news.news_show_en.FldCaption %></span></td>
<% End If %>
<% If z40news.news_show.Visible Then ' news_show %>
		<td><span id="elh_z40news_news_show" class="z40news_news_show"><%= z40news.news_show.FldCaption %></span></td>
<% End If %>
<% If z40news.news_create.Visible Then ' news_create %>
		<td><span id="elh_z40news_news_create" class="z40news_news_create"><%= z40news.news_create.FldCaption %></span></td>
<% End If %>
<% If z40news.news_update.Visible Then ' news_update %>
		<td><span id="elh_z40news_news_update" class="z40news_news_update"><%= z40news.news_update.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
z40news_delete.RecCnt = 0
z40news_delete.RowCnt = 0
Do While (Not z40news_delete.Recordset.Eof)
	z40news_delete.RecCnt = z40news_delete.RecCnt + 1
	z40news_delete.RowCnt = z40news_delete.RowCnt + 1

	' Set row properties
	Call z40news.ResetAttrs()
	z40news.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call z40news_delete.LoadRowValues(z40news_delete.Recordset)

	' Render row
	Call z40news_delete.RenderRow()
%>
	<tr<%= z40news.RowAttributes %>>
<% If z40news.news_id.Visible Then ' news_id %>
		<td<%= z40news.news_id.CellAttributes %>>
<span id="el<%= z40news_delete.RowCnt %>_z40news_news_id" class="control-group z40news_news_id">
<span<%= z40news.news_id.ViewAttributes %>>
<%= z40news.news_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If z40news.news_img.Visible Then ' news_img %>
		<td<%= z40news.news_img.CellAttributes %>>
<span id="el<%= z40news_delete.RowCnt %>_z40news_news_img" class="control-group z40news_news_img">
<span<%= z40news.news_img.ViewAttributes %>>
<%= z40news.news_img.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If z40news.news_category.Visible Then ' news_category %>
		<td<%= z40news.news_category.CellAttributes %>>
<span id="el<%= z40news_delete.RowCnt %>_z40news_news_category" class="control-group z40news_news_category">
<span<%= z40news.news_category.ViewAttributes %>>
<%= z40news.news_category.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If z40news.news_subject.Visible Then ' news_subject %>
		<td<%= z40news.news_subject.CellAttributes %>>
<span id="el<%= z40news_delete.RowCnt %>_z40news_news_subject" class="control-group z40news_news_subject">
<span<%= z40news.news_subject.ViewAttributes %>>
<%= z40news.news_subject.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If z40news.news_subject_th.Visible Then ' news_subject_th %>
		<td<%= z40news.news_subject_th.CellAttributes %>>
<span id="el<%= z40news_delete.RowCnt %>_z40news_news_subject_th" class="control-group z40news_news_subject_th">
<span<%= z40news.news_subject_th.ViewAttributes %>>
<%= z40news.news_subject_th.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If z40news.news_show_en.Visible Then ' news_show_en %>
		<td<%= z40news.news_show_en.CellAttributes %>>
<span id="el<%= z40news_delete.RowCnt %>_z40news_news_show_en" class="control-group z40news_news_show_en">
<span<%= z40news.news_show_en.ViewAttributes %>>
<%= z40news.news_show_en.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If z40news.news_show.Visible Then ' news_show %>
		<td<%= z40news.news_show.CellAttributes %>>
<span id="el<%= z40news_delete.RowCnt %>_z40news_news_show" class="control-group z40news_news_show">
<span<%= z40news.news_show.ViewAttributes %>>
<%= z40news.news_show.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If z40news.news_create.Visible Then ' news_create %>
		<td<%= z40news.news_create.CellAttributes %>>
<span id="el<%= z40news_delete.RowCnt %>_z40news_news_create" class="control-group z40news_news_create">
<span<%= z40news.news_create.ViewAttributes %>>
<%= z40news.news_create.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If z40news.news_update.Visible Then ' news_update %>
		<td<%= z40news.news_update.CellAttributes %>>
<span id="el<%= z40news_delete.RowCnt %>_z40news_news_update" class="control-group z40news_news_update">
<span<%= z40news.news_update.ViewAttributes %>>
<%= z40news.news_update.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	z40news_delete.Recordset.MoveNext
Loop
z40news_delete.Recordset.Close
Set z40news_delete.Recordset = Nothing
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
fz40newsdelete.Init();
</script>
<%
z40news_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set z40news_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cz40news_delete

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
		TableName = "@news"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "z40news_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If z40news.UseTokenInUrl Then PageUrl = PageUrl & "t=" & z40news.TableVar & "&" ' add page token
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
		If z40news.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (z40news.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (z40news.TableVar = Request.QueryString("t"))
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
		If IsEmpty(z40news) Then Set z40news = New cz40news
		Set Table = z40news

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "@news"

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
		Set z40news = Nothing
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
		RecKeys = z40news.GetRecordKeys() ' Load record keys
		sFilter = z40news.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_z40newslist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in z40news class, z40newsinfo.asp

		z40news.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			z40news.CurrentAction = Request.Form("a_delete")
		Else
			z40news.CurrentAction = "I"	' Display record
		End If
		Select Case z40news.CurrentAction
			Case "D" ' Delete
				z40news.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(z40news.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = z40news.CurrentFilter
		Call z40news.Recordset_Selecting(sFilter)
		z40news.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = z40news.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call z40news.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = z40news.KeyFilter

		' Call Row Selecting event
		Call z40news.Row_Selecting(sFilter)

		' Load sql based on filter
		z40news.CurrentFilter = sFilter
		sSql = z40news.SQL
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
		Call z40news.Row_Selected(RsRow)
		z40news.news_id.DbValue = RsRow("news_id")
		z40news.news_img.DbValue = RsRow("news_img")
		z40news.news_category.DbValue = RsRow("news_category")
		z40news.news_subject.DbValue = RsRow("news_subject")
		z40news.news_subject_th.DbValue = RsRow("news_subject_th")
		z40news.news_intro.DbValue = RsRow("news_intro")
		z40news.news_intro_th.DbValue = RsRow("news_intro_th")
		z40news.news_content.DbValue = RsRow("news_content")
		z40news.news_content_th.DbValue = RsRow("news_content_th")
		z40news.news_show_en.DbValue = RsRow("news_show_en")
		z40news.news_show.DbValue = RsRow("news_show")
		z40news.news_create.DbValue = RsRow("news_create")
		z40news.news_update.DbValue = RsRow("news_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		z40news.news_id.m_DbValue = Rs("news_id")
		z40news.news_img.m_DbValue = Rs("news_img")
		z40news.news_category.m_DbValue = Rs("news_category")
		z40news.news_subject.m_DbValue = Rs("news_subject")
		z40news.news_subject_th.m_DbValue = Rs("news_subject_th")
		z40news.news_intro.m_DbValue = Rs("news_intro")
		z40news.news_intro_th.m_DbValue = Rs("news_intro_th")
		z40news.news_content.m_DbValue = Rs("news_content")
		z40news.news_content_th.m_DbValue = Rs("news_content_th")
		z40news.news_show_en.m_DbValue = Rs("news_show_en")
		z40news.news_show.m_DbValue = Rs("news_show")
		z40news.news_create.m_DbValue = Rs("news_create")
		z40news.news_update.m_DbValue = Rs("news_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call z40news.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' news_id
		' news_img
		' news_category
		' news_subject
		' news_subject_th
		' news_intro
		' news_intro_th
		' news_content
		' news_content_th
		' news_show_en
		' news_show
		' news_create
		' news_update
		' -----------
		'  View  Row
		' -----------

		If z40news.RowType = EW_ROWTYPE_VIEW Then ' View row

			' news_id
			z40news.news_id.ViewValue = z40news.news_id.CurrentValue
			z40news.news_id.ViewCustomAttributes = ""

			' news_img
			z40news.news_img.ViewValue = z40news.news_img.CurrentValue
			z40news.news_img.ViewCustomAttributes = ""

			' news_category
			z40news.news_category.ViewValue = z40news.news_category.CurrentValue
			z40news.news_category.ViewCustomAttributes = ""

			' news_subject
			z40news.news_subject.ViewValue = z40news.news_subject.CurrentValue
			z40news.news_subject.ViewCustomAttributes = ""

			' news_subject_th
			z40news.news_subject_th.ViewValue = z40news.news_subject_th.CurrentValue
			z40news.news_subject_th.ViewCustomAttributes = ""

			' news_show_en
			z40news.news_show_en.ViewValue = z40news.news_show_en.CurrentValue
			z40news.news_show_en.ViewCustomAttributes = ""

			' news_show
			z40news.news_show.ViewValue = z40news.news_show.CurrentValue
			z40news.news_show.ViewCustomAttributes = ""

			' news_create
			z40news.news_create.ViewValue = z40news.news_create.CurrentValue
			z40news.news_create.ViewCustomAttributes = ""

			' news_update
			z40news.news_update.ViewValue = z40news.news_update.CurrentValue
			z40news.news_update.ViewCustomAttributes = ""

			' View refer script
			' news_id

			z40news.news_id.LinkCustomAttributes = ""
			z40news.news_id.HrefValue = ""
			z40news.news_id.TooltipValue = ""

			' news_img
			z40news.news_img.LinkCustomAttributes = ""
			z40news.news_img.HrefValue = ""
			z40news.news_img.TooltipValue = ""

			' news_category
			z40news.news_category.LinkCustomAttributes = ""
			z40news.news_category.HrefValue = ""
			z40news.news_category.TooltipValue = ""

			' news_subject
			z40news.news_subject.LinkCustomAttributes = ""
			z40news.news_subject.HrefValue = ""
			z40news.news_subject.TooltipValue = ""

			' news_subject_th
			z40news.news_subject_th.LinkCustomAttributes = ""
			z40news.news_subject_th.HrefValue = ""
			z40news.news_subject_th.TooltipValue = ""

			' news_show_en
			z40news.news_show_en.LinkCustomAttributes = ""
			z40news.news_show_en.HrefValue = ""
			z40news.news_show_en.TooltipValue = ""

			' news_show
			z40news.news_show.LinkCustomAttributes = ""
			z40news.news_show.HrefValue = ""
			z40news.news_show.TooltipValue = ""

			' news_create
			z40news.news_create.LinkCustomAttributes = ""
			z40news.news_create.HrefValue = ""
			z40news.news_create.TooltipValue = ""

			' news_update
			z40news.news_update.LinkCustomAttributes = ""
			z40news.news_update.HrefValue = ""
			z40news.news_update.TooltipValue = ""
		End If

		' Call Row Rendered event
		If z40news.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call z40news.Row_Rendered()
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
		sSql = z40news.SQL
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
				DeleteRows = z40news.Row_Deleting(RsDelete)
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
			ElseIf z40news.CancelMessage <> "" Then
				FailureMessage = z40news.CancelMessage
				z40news.CancelMessage = ""
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
				Call z40news.Row_Deleted(RsOld)
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
		Call Breadcrumb.Add("list", z40news.TableVar, "pom_z40newslist.asp", z40news.TableVar, True)
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
