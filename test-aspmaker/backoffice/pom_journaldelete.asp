<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_journalinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim journal_delete
Set journal_delete = New cjournal_delete
Set Page = journal_delete

' Page init processing
journal_delete.Page_Init()

' Page main processing
journal_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
journal_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var journal_delete = new ew_Page("journal_delete");
journal_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = journal_delete.PageID; // For backward compatibility
// Form object
var fjournaldelete = new ew_Form("fjournaldelete");
// Form_CustomValidate event
fjournaldelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fjournaldelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fjournaldelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set journal_delete.Recordset = journal_delete.LoadRecordset()
journal_delete.TotalRecs = journal_delete.Recordset.RecordCount ' Get record count
If journal_delete.TotalRecs <= 0 Then ' No record found, exit
	journal_delete.Recordset.Close
	Set journal_delete.Recordset = Nothing
	Call journal_delete.Page_Terminate("pom_journallist.asp") ' Return to list
End If
%>
<% If journal.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% journal_delete.ShowPageHeader() %>
<% journal_delete.ShowMessage %>
<form name="fjournaldelete" id="fjournaldelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="journal">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(journal_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(journal_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_journaldelete" class="ewTable ewTableSeparate">
<%= journal.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If journal.jrl_id.Visible Then ' jrl_id %>
		<td><span id="elh_journal_jrl_id" class="journal_jrl_id"><%= journal.jrl_id.FldCaption %></span></td>
<% End If %>
<% If journal.jrl_category.Visible Then ' jrl_category %>
		<td><span id="elh_journal_jrl_category" class="journal_jrl_category"><%= journal.jrl_category.FldCaption %></span></td>
<% End If %>
<% If journal.jrl_date.Visible Then ' jrl_date %>
		<td><span id="elh_journal_jrl_date" class="journal_jrl_date"><%= journal.jrl_date.FldCaption %></span></td>
<% End If %>
<% If journal.jrl_title.Visible Then ' jrl_title %>
		<td><span id="elh_journal_jrl_title" class="journal_jrl_title"><%= journal.jrl_title.FldCaption %></span></td>
<% End If %>
<% If journal.jrl_title_th.Visible Then ' jrl_title_th %>
		<td><span id="elh_journal_jrl_title_th" class="journal_jrl_title_th"><%= journal.jrl_title_th.FldCaption %></span></td>
<% End If %>
<% If journal.jrl_pdf.Visible Then ' jrl_pdf %>
		<td><span id="elh_journal_jrl_pdf" class="journal_jrl_pdf"><%= journal.jrl_pdf.FldCaption %></span></td>
<% End If %>
<% If journal.jrl_img.Visible Then ' jrl_img %>
		<td><span id="elh_journal_jrl_img" class="journal_jrl_img"><%= journal.jrl_img.FldCaption %></span></td>
<% End If %>
<% If journal.jrl_create.Visible Then ' jrl_create %>
		<td><span id="elh_journal_jrl_create" class="journal_jrl_create"><%= journal.jrl_create.FldCaption %></span></td>
<% End If %>
<% If journal.jrl_update.Visible Then ' jrl_update %>
		<td><span id="elh_journal_jrl_update" class="journal_jrl_update"><%= journal.jrl_update.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
journal_delete.RecCnt = 0
journal_delete.RowCnt = 0
Do While (Not journal_delete.Recordset.Eof)
	journal_delete.RecCnt = journal_delete.RecCnt + 1
	journal_delete.RowCnt = journal_delete.RowCnt + 1

	' Set row properties
	Call journal.ResetAttrs()
	journal.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call journal_delete.LoadRowValues(journal_delete.Recordset)

	' Render row
	Call journal_delete.RenderRow()
%>
	<tr<%= journal.RowAttributes %>>
<% If journal.jrl_id.Visible Then ' jrl_id %>
		<td<%= journal.jrl_id.CellAttributes %>>
<span id="el<%= journal_delete.RowCnt %>_journal_jrl_id" class="control-group journal_jrl_id">
<span<%= journal.jrl_id.ViewAttributes %>>
<%= journal.jrl_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If journal.jrl_category.Visible Then ' jrl_category %>
		<td<%= journal.jrl_category.CellAttributes %>>
<span id="el<%= journal_delete.RowCnt %>_journal_jrl_category" class="control-group journal_jrl_category">
<span<%= journal.jrl_category.ViewAttributes %>>
<%= journal.jrl_category.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If journal.jrl_date.Visible Then ' jrl_date %>
		<td<%= journal.jrl_date.CellAttributes %>>
<span id="el<%= journal_delete.RowCnt %>_journal_jrl_date" class="control-group journal_jrl_date">
<span<%= journal.jrl_date.ViewAttributes %>>
<%= journal.jrl_date.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If journal.jrl_title.Visible Then ' jrl_title %>
		<td<%= journal.jrl_title.CellAttributes %>>
<span id="el<%= journal_delete.RowCnt %>_journal_jrl_title" class="control-group journal_jrl_title">
<span<%= journal.jrl_title.ViewAttributes %>>
<%= journal.jrl_title.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If journal.jrl_title_th.Visible Then ' jrl_title_th %>
		<td<%= journal.jrl_title_th.CellAttributes %>>
<span id="el<%= journal_delete.RowCnt %>_journal_jrl_title_th" class="control-group journal_jrl_title_th">
<span<%= journal.jrl_title_th.ViewAttributes %>>
<%= journal.jrl_title_th.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If journal.jrl_pdf.Visible Then ' jrl_pdf %>
		<td<%= journal.jrl_pdf.CellAttributes %>>
<span id="el<%= journal_delete.RowCnt %>_journal_jrl_pdf" class="control-group journal_jrl_pdf">
<span<%= journal.jrl_pdf.ViewAttributes %>>
<%= journal.jrl_pdf.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If journal.jrl_img.Visible Then ' jrl_img %>
		<td<%= journal.jrl_img.CellAttributes %>>
<span id="el<%= journal_delete.RowCnt %>_journal_jrl_img" class="control-group journal_jrl_img">
<span<%= journal.jrl_img.ViewAttributes %>>
<%= journal.jrl_img.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If journal.jrl_create.Visible Then ' jrl_create %>
		<td<%= journal.jrl_create.CellAttributes %>>
<span id="el<%= journal_delete.RowCnt %>_journal_jrl_create" class="control-group journal_jrl_create">
<span<%= journal.jrl_create.ViewAttributes %>>
<%= journal.jrl_create.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If journal.jrl_update.Visible Then ' jrl_update %>
		<td<%= journal.jrl_update.CellAttributes %>>
<span id="el<%= journal_delete.RowCnt %>_journal_jrl_update" class="control-group journal_jrl_update">
<span<%= journal.jrl_update.ViewAttributes %>>
<%= journal.jrl_update.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	journal_delete.Recordset.MoveNext
Loop
journal_delete.Recordset.Close
Set journal_delete.Recordset = Nothing
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
fjournaldelete.Init();
</script>
<%
journal_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set journal_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cjournal_delete

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
		TableName = "journal"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "journal_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If journal.UseTokenInUrl Then PageUrl = PageUrl & "t=" & journal.TableVar & "&" ' add page token
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
		If journal.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (journal.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (journal.TableVar = Request.QueryString("t"))
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
		If IsEmpty(journal) Then Set journal = New cjournal
		Set Table = journal

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "journal"

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
		Set journal = Nothing
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
		RecKeys = journal.GetRecordKeys() ' Load record keys
		sFilter = journal.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_journallist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in journal class, journalinfo.asp

		journal.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			journal.CurrentAction = Request.Form("a_delete")
		Else
			journal.CurrentAction = "I"	' Display record
		End If
		Select Case journal.CurrentAction
			Case "D" ' Delete
				journal.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(journal.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = journal.CurrentFilter
		Call journal.Recordset_Selecting(sFilter)
		journal.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = journal.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call journal.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = journal.KeyFilter

		' Call Row Selecting event
		Call journal.Row_Selecting(sFilter)

		' Load sql based on filter
		journal.CurrentFilter = sFilter
		sSql = journal.SQL
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
		Call journal.Row_Selected(RsRow)
		journal.jrl_id.DbValue = RsRow("jrl_id")
		journal.jrl_category.DbValue = RsRow("jrl_category")
		journal.jrl_date.DbValue = RsRow("jrl_date")
		journal.jrl_title.DbValue = RsRow("jrl_title")
		journal.jrl_title_th.DbValue = RsRow("jrl_title_th")
		journal.jrl_pdf.DbValue = RsRow("jrl_pdf")
		journal.jrl_img.DbValue = RsRow("jrl_img")
		journal.jrl_create.DbValue = RsRow("jrl_create")
		journal.jrl_update.DbValue = RsRow("jrl_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		journal.jrl_id.m_DbValue = Rs("jrl_id")
		journal.jrl_category.m_DbValue = Rs("jrl_category")
		journal.jrl_date.m_DbValue = Rs("jrl_date")
		journal.jrl_title.m_DbValue = Rs("jrl_title")
		journal.jrl_title_th.m_DbValue = Rs("jrl_title_th")
		journal.jrl_pdf.m_DbValue = Rs("jrl_pdf")
		journal.jrl_img.m_DbValue = Rs("jrl_img")
		journal.jrl_create.m_DbValue = Rs("jrl_create")
		journal.jrl_update.m_DbValue = Rs("jrl_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call journal.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' jrl_id
		' jrl_category
		' jrl_date
		' jrl_title
		' jrl_title_th
		' jrl_pdf
		' jrl_img
		' jrl_create
		' jrl_update
		' -----------
		'  View  Row
		' -----------

		If journal.RowType = EW_ROWTYPE_VIEW Then ' View row

			' jrl_id
			journal.jrl_id.ViewValue = journal.jrl_id.CurrentValue
			journal.jrl_id.ViewCustomAttributes = ""

			' jrl_category
			journal.jrl_category.ViewValue = journal.jrl_category.CurrentValue
			journal.jrl_category.ViewCustomAttributes = ""

			' jrl_date
			journal.jrl_date.ViewValue = journal.jrl_date.CurrentValue
			journal.jrl_date.ViewCustomAttributes = ""

			' jrl_title
			journal.jrl_title.ViewValue = journal.jrl_title.CurrentValue
			journal.jrl_title.ViewCustomAttributes = ""

			' jrl_title_th
			journal.jrl_title_th.ViewValue = journal.jrl_title_th.CurrentValue
			journal.jrl_title_th.ViewCustomAttributes = ""

			' jrl_pdf
			journal.jrl_pdf.ViewValue = journal.jrl_pdf.CurrentValue
			journal.jrl_pdf.ViewCustomAttributes = ""

			' jrl_img
			journal.jrl_img.ViewValue = journal.jrl_img.CurrentValue
			journal.jrl_img.ViewCustomAttributes = ""

			' jrl_create
			journal.jrl_create.ViewValue = journal.jrl_create.CurrentValue
			journal.jrl_create.ViewCustomAttributes = ""

			' jrl_update
			journal.jrl_update.ViewValue = journal.jrl_update.CurrentValue
			journal.jrl_update.ViewCustomAttributes = ""

			' View refer script
			' jrl_id

			journal.jrl_id.LinkCustomAttributes = ""
			journal.jrl_id.HrefValue = ""
			journal.jrl_id.TooltipValue = ""

			' jrl_category
			journal.jrl_category.LinkCustomAttributes = ""
			journal.jrl_category.HrefValue = ""
			journal.jrl_category.TooltipValue = ""

			' jrl_date
			journal.jrl_date.LinkCustomAttributes = ""
			journal.jrl_date.HrefValue = ""
			journal.jrl_date.TooltipValue = ""

			' jrl_title
			journal.jrl_title.LinkCustomAttributes = ""
			journal.jrl_title.HrefValue = ""
			journal.jrl_title.TooltipValue = ""

			' jrl_title_th
			journal.jrl_title_th.LinkCustomAttributes = ""
			journal.jrl_title_th.HrefValue = ""
			journal.jrl_title_th.TooltipValue = ""

			' jrl_pdf
			journal.jrl_pdf.LinkCustomAttributes = ""
			journal.jrl_pdf.HrefValue = ""
			journal.jrl_pdf.TooltipValue = ""

			' jrl_img
			journal.jrl_img.LinkCustomAttributes = ""
			journal.jrl_img.HrefValue = ""
			journal.jrl_img.TooltipValue = ""

			' jrl_create
			journal.jrl_create.LinkCustomAttributes = ""
			journal.jrl_create.HrefValue = ""
			journal.jrl_create.TooltipValue = ""

			' jrl_update
			journal.jrl_update.LinkCustomAttributes = ""
			journal.jrl_update.HrefValue = ""
			journal.jrl_update.TooltipValue = ""
		End If

		' Call Row Rendered event
		If journal.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call journal.Row_Rendered()
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
		sSql = journal.SQL
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
				DeleteRows = journal.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("jrl_id")
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
			ElseIf journal.CancelMessage <> "" Then
				FailureMessage = journal.CancelMessage
				journal.CancelMessage = ""
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
				Call journal.Row_Deleted(RsOld)
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
		Call Breadcrumb.Add("list", journal.TableVar, "pom_journallist.asp", journal.TableVar, True)
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
