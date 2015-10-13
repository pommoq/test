<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_e_libraryinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim e_library_delete
Set e_library_delete = New ce_library_delete
Set Page = e_library_delete

' Page init processing
e_library_delete.Page_Init()

' Page main processing
e_library_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
e_library_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var e_library_delete = new ew_Page("e_library_delete");
e_library_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = e_library_delete.PageID; // For backward compatibility
// Form object
var fe_librarydelete = new ew_Form("fe_librarydelete");
// Form_CustomValidate event
fe_librarydelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fe_librarydelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fe_librarydelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set e_library_delete.Recordset = e_library_delete.LoadRecordset()
e_library_delete.TotalRecs = e_library_delete.Recordset.RecordCount ' Get record count
If e_library_delete.TotalRecs <= 0 Then ' No record found, exit
	e_library_delete.Recordset.Close
	Set e_library_delete.Recordset = Nothing
	Call e_library_delete.Page_Terminate("pom_e_librarylist.asp") ' Return to list
End If
%>
<% If e_library.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% e_library_delete.ShowPageHeader() %>
<% e_library_delete.ShowMessage %>
<form name="fe_librarydelete" id="fe_librarydelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="e_library">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(e_library_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(e_library_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_e_librarydelete" class="ewTable ewTableSeparate">
<%= e_library.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If e_library.el_id.Visible Then ' el_id %>
		<td><span id="elh_e_library_el_id" class="e_library_el_id"><%= e_library.el_id.FldCaption %></span></td>
<% End If %>
<% If e_library.el_date.Visible Then ' el_date %>
		<td><span id="elh_e_library_el_date" class="e_library_el_date"><%= e_library.el_date.FldCaption %></span></td>
<% End If %>
<% If e_library.el_title.Visible Then ' el_title %>
		<td><span id="elh_e_library_el_title" class="e_library_el_title"><%= e_library.el_title.FldCaption %></span></td>
<% End If %>
<% If e_library.el_pdf.Visible Then ' el_pdf %>
		<td><span id="elh_e_library_el_pdf" class="e_library_el_pdf"><%= e_library.el_pdf.FldCaption %></span></td>
<% End If %>
<% If e_library.el_img.Visible Then ' el_img %>
		<td><span id="elh_e_library_el_img" class="e_library_el_img"><%= e_library.el_img.FldCaption %></span></td>
<% End If %>
<% If e_library.el_create.Visible Then ' el_create %>
		<td><span id="elh_e_library_el_create" class="e_library_el_create"><%= e_library.el_create.FldCaption %></span></td>
<% End If %>
<% If e_library.el_update.Visible Then ' el_update %>
		<td><span id="elh_e_library_el_update" class="e_library_el_update"><%= e_library.el_update.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
e_library_delete.RecCnt = 0
e_library_delete.RowCnt = 0
Do While (Not e_library_delete.Recordset.Eof)
	e_library_delete.RecCnt = e_library_delete.RecCnt + 1
	e_library_delete.RowCnt = e_library_delete.RowCnt + 1

	' Set row properties
	Call e_library.ResetAttrs()
	e_library.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call e_library_delete.LoadRowValues(e_library_delete.Recordset)

	' Render row
	Call e_library_delete.RenderRow()
%>
	<tr<%= e_library.RowAttributes %>>
<% If e_library.el_id.Visible Then ' el_id %>
		<td<%= e_library.el_id.CellAttributes %>>
<span id="el<%= e_library_delete.RowCnt %>_e_library_el_id" class="control-group e_library_el_id">
<span<%= e_library.el_id.ViewAttributes %>>
<%= e_library.el_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If e_library.el_date.Visible Then ' el_date %>
		<td<%= e_library.el_date.CellAttributes %>>
<span id="el<%= e_library_delete.RowCnt %>_e_library_el_date" class="control-group e_library_el_date">
<span<%= e_library.el_date.ViewAttributes %>>
<%= e_library.el_date.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If e_library.el_title.Visible Then ' el_title %>
		<td<%= e_library.el_title.CellAttributes %>>
<span id="el<%= e_library_delete.RowCnt %>_e_library_el_title" class="control-group e_library_el_title">
<span<%= e_library.el_title.ViewAttributes %>>
<%= e_library.el_title.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If e_library.el_pdf.Visible Then ' el_pdf %>
		<td<%= e_library.el_pdf.CellAttributes %>>
<span id="el<%= e_library_delete.RowCnt %>_e_library_el_pdf" class="control-group e_library_el_pdf">
<span<%= e_library.el_pdf.ViewAttributes %>>
<%= e_library.el_pdf.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If e_library.el_img.Visible Then ' el_img %>
		<td<%= e_library.el_img.CellAttributes %>>
<span id="el<%= e_library_delete.RowCnt %>_e_library_el_img" class="control-group e_library_el_img">
<span<%= e_library.el_img.ViewAttributes %>>
<%= e_library.el_img.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If e_library.el_create.Visible Then ' el_create %>
		<td<%= e_library.el_create.CellAttributes %>>
<span id="el<%= e_library_delete.RowCnt %>_e_library_el_create" class="control-group e_library_el_create">
<span<%= e_library.el_create.ViewAttributes %>>
<%= e_library.el_create.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If e_library.el_update.Visible Then ' el_update %>
		<td<%= e_library.el_update.CellAttributes %>>
<span id="el<%= e_library_delete.RowCnt %>_e_library_el_update" class="control-group e_library_el_update">
<span<%= e_library.el_update.ViewAttributes %>>
<%= e_library.el_update.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	e_library_delete.Recordset.MoveNext
Loop
e_library_delete.Recordset.Close
Set e_library_delete.Recordset = Nothing
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
fe_librarydelete.Init();
</script>
<%
e_library_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set e_library_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class ce_library_delete

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
		TableName = "e_library"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "e_library_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If e_library.UseTokenInUrl Then PageUrl = PageUrl & "t=" & e_library.TableVar & "&" ' add page token
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
		If e_library.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (e_library.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (e_library.TableVar = Request.QueryString("t"))
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
		If IsEmpty(e_library) Then Set e_library = New ce_library
		Set Table = e_library

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "e_library"

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
		Set e_library = Nothing
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
		RecKeys = e_library.GetRecordKeys() ' Load record keys
		sFilter = e_library.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_e_librarylist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in e_library class, e_libraryinfo.asp

		e_library.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			e_library.CurrentAction = Request.Form("a_delete")
		Else
			e_library.CurrentAction = "I"	' Display record
		End If
		Select Case e_library.CurrentAction
			Case "D" ' Delete
				e_library.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(e_library.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = e_library.CurrentFilter
		Call e_library.Recordset_Selecting(sFilter)
		e_library.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = e_library.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call e_library.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = e_library.KeyFilter

		' Call Row Selecting event
		Call e_library.Row_Selecting(sFilter)

		' Load sql based on filter
		e_library.CurrentFilter = sFilter
		sSql = e_library.SQL
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
		Call e_library.Row_Selected(RsRow)
		e_library.el_id.DbValue = RsRow("el_id")
		e_library.el_date.DbValue = RsRow("el_date")
		e_library.el_title.DbValue = RsRow("el_title")
		e_library.el_pdf.DbValue = RsRow("el_pdf")
		e_library.el_img.DbValue = RsRow("el_img")
		e_library.el_detail.DbValue = RsRow("el_detail")
		e_library.el_create.DbValue = RsRow("el_create")
		e_library.el_update.DbValue = RsRow("el_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		e_library.el_id.m_DbValue = Rs("el_id")
		e_library.el_date.m_DbValue = Rs("el_date")
		e_library.el_title.m_DbValue = Rs("el_title")
		e_library.el_pdf.m_DbValue = Rs("el_pdf")
		e_library.el_img.m_DbValue = Rs("el_img")
		e_library.el_detail.m_DbValue = Rs("el_detail")
		e_library.el_create.m_DbValue = Rs("el_create")
		e_library.el_update.m_DbValue = Rs("el_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call e_library.Row_Rendering()

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

		If e_library.RowType = EW_ROWTYPE_VIEW Then ' View row

			' el_id
			e_library.el_id.ViewValue = e_library.el_id.CurrentValue
			e_library.el_id.ViewCustomAttributes = ""

			' el_date
			e_library.el_date.ViewValue = e_library.el_date.CurrentValue
			e_library.el_date.ViewCustomAttributes = ""

			' el_title
			e_library.el_title.ViewValue = e_library.el_title.CurrentValue
			e_library.el_title.ViewCustomAttributes = ""

			' el_pdf
			e_library.el_pdf.ViewValue = e_library.el_pdf.CurrentValue
			e_library.el_pdf.ViewCustomAttributes = ""

			' el_img
			e_library.el_img.ViewValue = e_library.el_img.CurrentValue
			e_library.el_img.ViewCustomAttributes = ""

			' el_create
			e_library.el_create.ViewValue = e_library.el_create.CurrentValue
			e_library.el_create.ViewCustomAttributes = ""

			' el_update
			e_library.el_update.ViewValue = e_library.el_update.CurrentValue
			e_library.el_update.ViewCustomAttributes = ""

			' View refer script
			' el_id

			e_library.el_id.LinkCustomAttributes = ""
			e_library.el_id.HrefValue = ""
			e_library.el_id.TooltipValue = ""

			' el_date
			e_library.el_date.LinkCustomAttributes = ""
			e_library.el_date.HrefValue = ""
			e_library.el_date.TooltipValue = ""

			' el_title
			e_library.el_title.LinkCustomAttributes = ""
			e_library.el_title.HrefValue = ""
			e_library.el_title.TooltipValue = ""

			' el_pdf
			e_library.el_pdf.LinkCustomAttributes = ""
			e_library.el_pdf.HrefValue = ""
			e_library.el_pdf.TooltipValue = ""

			' el_img
			e_library.el_img.LinkCustomAttributes = ""
			e_library.el_img.HrefValue = ""
			e_library.el_img.TooltipValue = ""

			' el_create
			e_library.el_create.LinkCustomAttributes = ""
			e_library.el_create.HrefValue = ""
			e_library.el_create.TooltipValue = ""

			' el_update
			e_library.el_update.LinkCustomAttributes = ""
			e_library.el_update.HrefValue = ""
			e_library.el_update.TooltipValue = ""
		End If

		' Call Row Rendered event
		If e_library.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call e_library.Row_Rendered()
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
		sSql = e_library.SQL
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
				DeleteRows = e_library.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("el_id")
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
			ElseIf e_library.CancelMessage <> "" Then
				FailureMessage = e_library.CancelMessage
				e_library.CancelMessage = ""
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
				Call e_library.Row_Deleted(RsOld)
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
		Call Breadcrumb.Add("list", e_library.TableVar, "pom_e_librarylist.asp", e_library.TableVar, True)
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
