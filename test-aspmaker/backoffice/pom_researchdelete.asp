<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_researchinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim research_delete
Set research_delete = New cresearch_delete
Set Page = research_delete

' Page init processing
research_delete.Page_Init()

' Page main processing
research_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
research_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var research_delete = new ew_Page("research_delete");
research_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = research_delete.PageID; // For backward compatibility
// Form object
var fresearchdelete = new ew_Form("fresearchdelete");
// Form_CustomValidate event
fresearchdelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fresearchdelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fresearchdelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set research_delete.Recordset = research_delete.LoadRecordset()
research_delete.TotalRecs = research_delete.Recordset.RecordCount ' Get record count
If research_delete.TotalRecs <= 0 Then ' No record found, exit
	research_delete.Recordset.Close
	Set research_delete.Recordset = Nothing
	Call research_delete.Page_Terminate("pom_researchlist.asp") ' Return to list
End If
%>
<% If research.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% research_delete.ShowPageHeader() %>
<% research_delete.ShowMessage %>
<form name="fresearchdelete" id="fresearchdelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="research">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(research_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(research_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_researchdelete" class="ewTable ewTableSeparate">
<%= research.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If research.rsh_id.Visible Then ' rsh_id %>
		<td><span id="elh_research_rsh_id" class="research_rsh_id"><%= research.rsh_id.FldCaption %></span></td>
<% End If %>
<% If research.rsh_img.Visible Then ' rsh_img %>
		<td><span id="elh_research_rsh_img" class="research_rsh_img"><%= research.rsh_img.FldCaption %></span></td>
<% End If %>
<% If research.rsh_date.Visible Then ' rsh_date %>
		<td><span id="elh_research_rsh_date" class="research_rsh_date"><%= research.rsh_date.FldCaption %></span></td>
<% End If %>
<% If research.rsh_pdf.Visible Then ' rsh_pdf %>
		<td><span id="elh_research_rsh_pdf" class="research_rsh_pdf"><%= research.rsh_pdf.FldCaption %></span></td>
<% End If %>
<% If research.rsh_category.Visible Then ' rsh_category %>
		<td><span id="elh_research_rsh_category" class="research_rsh_category"><%= research.rsh_category.FldCaption %></span></td>
<% End If %>
<% If research.rsh_subject.Visible Then ' rsh_subject %>
		<td><span id="elh_research_rsh_subject" class="research_rsh_subject"><%= research.rsh_subject.FldCaption %></span></td>
<% End If %>
<% If research.rsh_subject_th.Visible Then ' rsh_subject_th %>
		<td><span id="elh_research_rsh_subject_th" class="research_rsh_subject_th"><%= research.rsh_subject_th.FldCaption %></span></td>
<% End If %>
<% If research.rsh_intro_th.Visible Then ' rsh_intro_th %>
		<td><span id="elh_research_rsh_intro_th" class="research_rsh_intro_th"><%= research.rsh_intro_th.FldCaption %></span></td>
<% End If %>
<% If research.rsh_show.Visible Then ' rsh_show %>
		<td><span id="elh_research_rsh_show" class="research_rsh_show"><%= research.rsh_show.FldCaption %></span></td>
<% End If %>
<% If research.rsh_show_home.Visible Then ' rsh_show_home %>
		<td><span id="elh_research_rsh_show_home" class="research_rsh_show_home"><%= research.rsh_show_home.FldCaption %></span></td>
<% End If %>
<% If research.rsh_create.Visible Then ' rsh_create %>
		<td><span id="elh_research_rsh_create" class="research_rsh_create"><%= research.rsh_create.FldCaption %></span></td>
<% End If %>
<% If research.rsh_update.Visible Then ' rsh_update %>
		<td><span id="elh_research_rsh_update" class="research_rsh_update"><%= research.rsh_update.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
research_delete.RecCnt = 0
research_delete.RowCnt = 0
Do While (Not research_delete.Recordset.Eof)
	research_delete.RecCnt = research_delete.RecCnt + 1
	research_delete.RowCnt = research_delete.RowCnt + 1

	' Set row properties
	Call research.ResetAttrs()
	research.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call research_delete.LoadRowValues(research_delete.Recordset)

	' Render row
	Call research_delete.RenderRow()
%>
	<tr<%= research.RowAttributes %>>
<% If research.rsh_id.Visible Then ' rsh_id %>
		<td<%= research.rsh_id.CellAttributes %>>
<span id="el<%= research_delete.RowCnt %>_research_rsh_id" class="control-group research_rsh_id">
<span<%= research.rsh_id.ViewAttributes %>>
<%= research.rsh_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If research.rsh_img.Visible Then ' rsh_img %>
		<td<%= research.rsh_img.CellAttributes %>>
<span id="el<%= research_delete.RowCnt %>_research_rsh_img" class="control-group research_rsh_img">
<span<%= research.rsh_img.ViewAttributes %>>
<%= research.rsh_img.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If research.rsh_date.Visible Then ' rsh_date %>
		<td<%= research.rsh_date.CellAttributes %>>
<span id="el<%= research_delete.RowCnt %>_research_rsh_date" class="control-group research_rsh_date">
<span<%= research.rsh_date.ViewAttributes %>>
<%= research.rsh_date.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If research.rsh_pdf.Visible Then ' rsh_pdf %>
		<td<%= research.rsh_pdf.CellAttributes %>>
<span id="el<%= research_delete.RowCnt %>_research_rsh_pdf" class="control-group research_rsh_pdf">
<span<%= research.rsh_pdf.ViewAttributes %>>
<%= research.rsh_pdf.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If research.rsh_category.Visible Then ' rsh_category %>
		<td<%= research.rsh_category.CellAttributes %>>
<span id="el<%= research_delete.RowCnt %>_research_rsh_category" class="control-group research_rsh_category">
<span<%= research.rsh_category.ViewAttributes %>>
<%= research.rsh_category.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If research.rsh_subject.Visible Then ' rsh_subject %>
		<td<%= research.rsh_subject.CellAttributes %>>
<span id="el<%= research_delete.RowCnt %>_research_rsh_subject" class="control-group research_rsh_subject">
<span<%= research.rsh_subject.ViewAttributes %>>
<%= research.rsh_subject.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If research.rsh_subject_th.Visible Then ' rsh_subject_th %>
		<td<%= research.rsh_subject_th.CellAttributes %>>
<span id="el<%= research_delete.RowCnt %>_research_rsh_subject_th" class="control-group research_rsh_subject_th">
<span<%= research.rsh_subject_th.ViewAttributes %>>
<%= research.rsh_subject_th.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If research.rsh_intro_th.Visible Then ' rsh_intro_th %>
		<td<%= research.rsh_intro_th.CellAttributes %>>
<span id="el<%= research_delete.RowCnt %>_research_rsh_intro_th" class="control-group research_rsh_intro_th">
<span<%= research.rsh_intro_th.ViewAttributes %>>
<%= research.rsh_intro_th.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If research.rsh_show.Visible Then ' rsh_show %>
		<td<%= research.rsh_show.CellAttributes %>>
<span id="el<%= research_delete.RowCnt %>_research_rsh_show" class="control-group research_rsh_show">
<span<%= research.rsh_show.ViewAttributes %>>
<%= research.rsh_show.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If research.rsh_show_home.Visible Then ' rsh_show_home %>
		<td<%= research.rsh_show_home.CellAttributes %>>
<span id="el<%= research_delete.RowCnt %>_research_rsh_show_home" class="control-group research_rsh_show_home">
<span<%= research.rsh_show_home.ViewAttributes %>>
<%= research.rsh_show_home.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If research.rsh_create.Visible Then ' rsh_create %>
		<td<%= research.rsh_create.CellAttributes %>>
<span id="el<%= research_delete.RowCnt %>_research_rsh_create" class="control-group research_rsh_create">
<span<%= research.rsh_create.ViewAttributes %>>
<%= research.rsh_create.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If research.rsh_update.Visible Then ' rsh_update %>
		<td<%= research.rsh_update.CellAttributes %>>
<span id="el<%= research_delete.RowCnt %>_research_rsh_update" class="control-group research_rsh_update">
<span<%= research.rsh_update.ViewAttributes %>>
<%= research.rsh_update.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	research_delete.Recordset.MoveNext
Loop
research_delete.Recordset.Close
Set research_delete.Recordset = Nothing
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
fresearchdelete.Init();
</script>
<%
research_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set research_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cresearch_delete

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
		TableName = "research"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "research_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If research.UseTokenInUrl Then PageUrl = PageUrl & "t=" & research.TableVar & "&" ' add page token
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
		If research.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (research.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (research.TableVar = Request.QueryString("t"))
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
		If IsEmpty(research) Then Set research = New cresearch
		Set Table = research

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "research"

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
		Set research = Nothing
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
		RecKeys = research.GetRecordKeys() ' Load record keys
		sFilter = research.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_researchlist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in research class, researchinfo.asp

		research.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			research.CurrentAction = Request.Form("a_delete")
		Else
			research.CurrentAction = "I"	' Display record
		End If
		Select Case research.CurrentAction
			Case "D" ' Delete
				research.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(research.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = research.CurrentFilter
		Call research.Recordset_Selecting(sFilter)
		research.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = research.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call research.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = research.KeyFilter

		' Call Row Selecting event
		Call research.Row_Selecting(sFilter)

		' Load sql based on filter
		research.CurrentFilter = sFilter
		sSql = research.SQL
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
		Call research.Row_Selected(RsRow)
		research.rsh_id.DbValue = RsRow("rsh_id")
		research.rsh_img.DbValue = RsRow("rsh_img")
		research.rsh_date.DbValue = RsRow("rsh_date")
		research.rsh_pdf.DbValue = RsRow("rsh_pdf")
		research.rsh_category.DbValue = RsRow("rsh_category")
		research.rsh_subject.DbValue = RsRow("rsh_subject")
		research.rsh_subject_th.DbValue = RsRow("rsh_subject_th")
		research.rsh_intro.DbValue = RsRow("rsh_intro")
		research.rsh_intro_th.DbValue = RsRow("rsh_intro_th")
		research.rsh_content.DbValue = RsRow("rsh_content")
		research.rsh_content_th.DbValue = RsRow("rsh_content_th")
		research.rsh_show.DbValue = RsRow("rsh_show")
		research.rsh_show_home.DbValue = RsRow("rsh_show_home")
		research.rsh_create.DbValue = RsRow("rsh_create")
		research.rsh_update.DbValue = RsRow("rsh_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		research.rsh_id.m_DbValue = Rs("rsh_id")
		research.rsh_img.m_DbValue = Rs("rsh_img")
		research.rsh_date.m_DbValue = Rs("rsh_date")
		research.rsh_pdf.m_DbValue = Rs("rsh_pdf")
		research.rsh_category.m_DbValue = Rs("rsh_category")
		research.rsh_subject.m_DbValue = Rs("rsh_subject")
		research.rsh_subject_th.m_DbValue = Rs("rsh_subject_th")
		research.rsh_intro.m_DbValue = Rs("rsh_intro")
		research.rsh_intro_th.m_DbValue = Rs("rsh_intro_th")
		research.rsh_content.m_DbValue = Rs("rsh_content")
		research.rsh_content_th.m_DbValue = Rs("rsh_content_th")
		research.rsh_show.m_DbValue = Rs("rsh_show")
		research.rsh_show_home.m_DbValue = Rs("rsh_show_home")
		research.rsh_create.m_DbValue = Rs("rsh_create")
		research.rsh_update.m_DbValue = Rs("rsh_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call research.Row_Rendering()

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

		If research.RowType = EW_ROWTYPE_VIEW Then ' View row

			' rsh_id
			research.rsh_id.ViewValue = research.rsh_id.CurrentValue
			research.rsh_id.ViewCustomAttributes = ""

			' rsh_img
			research.rsh_img.ViewValue = research.rsh_img.CurrentValue
			research.rsh_img.ViewCustomAttributes = ""

			' rsh_date
			research.rsh_date.ViewValue = research.rsh_date.CurrentValue
			research.rsh_date.ViewCustomAttributes = ""

			' rsh_pdf
			research.rsh_pdf.ViewValue = research.rsh_pdf.CurrentValue
			research.rsh_pdf.ViewCustomAttributes = ""

			' rsh_category
			research.rsh_category.ViewValue = research.rsh_category.CurrentValue
			research.rsh_category.ViewCustomAttributes = ""

			' rsh_subject
			research.rsh_subject.ViewValue = research.rsh_subject.CurrentValue
			research.rsh_subject.ViewCustomAttributes = ""

			' rsh_subject_th
			research.rsh_subject_th.ViewValue = research.rsh_subject_th.CurrentValue
			research.rsh_subject_th.ViewCustomAttributes = ""

			' rsh_intro_th
			research.rsh_intro_th.ViewValue = research.rsh_intro_th.CurrentValue
			research.rsh_intro_th.ViewCustomAttributes = ""

			' rsh_show
			research.rsh_show.ViewValue = research.rsh_show.CurrentValue
			research.rsh_show.ViewCustomAttributes = ""

			' rsh_show_home
			research.rsh_show_home.ViewValue = research.rsh_show_home.CurrentValue
			research.rsh_show_home.ViewCustomAttributes = ""

			' rsh_create
			research.rsh_create.ViewValue = research.rsh_create.CurrentValue
			research.rsh_create.ViewCustomAttributes = ""

			' rsh_update
			research.rsh_update.ViewValue = research.rsh_update.CurrentValue
			research.rsh_update.ViewCustomAttributes = ""

			' View refer script
			' rsh_id

			research.rsh_id.LinkCustomAttributes = ""
			research.rsh_id.HrefValue = ""
			research.rsh_id.TooltipValue = ""

			' rsh_img
			research.rsh_img.LinkCustomAttributes = ""
			research.rsh_img.HrefValue = ""
			research.rsh_img.TooltipValue = ""

			' rsh_date
			research.rsh_date.LinkCustomAttributes = ""
			research.rsh_date.HrefValue = ""
			research.rsh_date.TooltipValue = ""

			' rsh_pdf
			research.rsh_pdf.LinkCustomAttributes = ""
			research.rsh_pdf.HrefValue = ""
			research.rsh_pdf.TooltipValue = ""

			' rsh_category
			research.rsh_category.LinkCustomAttributes = ""
			research.rsh_category.HrefValue = ""
			research.rsh_category.TooltipValue = ""

			' rsh_subject
			research.rsh_subject.LinkCustomAttributes = ""
			research.rsh_subject.HrefValue = ""
			research.rsh_subject.TooltipValue = ""

			' rsh_subject_th
			research.rsh_subject_th.LinkCustomAttributes = ""
			research.rsh_subject_th.HrefValue = ""
			research.rsh_subject_th.TooltipValue = ""

			' rsh_intro_th
			research.rsh_intro_th.LinkCustomAttributes = ""
			research.rsh_intro_th.HrefValue = ""
			research.rsh_intro_th.TooltipValue = ""

			' rsh_show
			research.rsh_show.LinkCustomAttributes = ""
			research.rsh_show.HrefValue = ""
			research.rsh_show.TooltipValue = ""

			' rsh_show_home
			research.rsh_show_home.LinkCustomAttributes = ""
			research.rsh_show_home.HrefValue = ""
			research.rsh_show_home.TooltipValue = ""

			' rsh_create
			research.rsh_create.LinkCustomAttributes = ""
			research.rsh_create.HrefValue = ""
			research.rsh_create.TooltipValue = ""

			' rsh_update
			research.rsh_update.LinkCustomAttributes = ""
			research.rsh_update.HrefValue = ""
			research.rsh_update.TooltipValue = ""
		End If

		' Call Row Rendered event
		If research.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call research.Row_Rendered()
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
		sSql = research.SQL
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
				DeleteRows = research.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("rsh_id")
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
			ElseIf research.CancelMessage <> "" Then
				FailureMessage = research.CancelMessage
				research.CancelMessage = ""
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
				Call research.Row_Deleted(RsOld)
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
		Call Breadcrumb.Add("list", research.TableVar, "pom_researchlist.asp", research.TableVar, True)
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
