<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_department_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim department_th_delete
Set department_th_delete = New cdepartment_th_delete
Set Page = department_th_delete

' Page init processing
department_th_delete.Page_Init()

' Page main processing
department_th_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
department_th_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var department_th_delete = new ew_Page("department_th_delete");
department_th_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = department_th_delete.PageID; // For backward compatibility
// Form object
var fdepartment_thdelete = new ew_Form("fdepartment_thdelete");
// Form_CustomValidate event
fdepartment_thdelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fdepartment_thdelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fdepartment_thdelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set department_th_delete.Recordset = department_th_delete.LoadRecordset()
department_th_delete.TotalRecs = department_th_delete.Recordset.RecordCount ' Get record count
If department_th_delete.TotalRecs <= 0 Then ' No record found, exit
	department_th_delete.Recordset.Close
	Set department_th_delete.Recordset = Nothing
	Call department_th_delete.Page_Terminate("pom_department_thlist.asp") ' Return to list
End If
%>
<% If department_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% department_th_delete.ShowPageHeader() %>
<% department_th_delete.ShowMessage %>
<form name="fdepartment_thdelete" id="fdepartment_thdelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="department_th">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(department_th_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(department_th_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_department_thdelete" class="ewTable ewTableSeparate">
<%= department_th.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If department_th.dept_id.Visible Then ' dept_id %>
		<td><span id="elh_department_th_dept_id" class="department_th_dept_id"><%= department_th.dept_id.FldCaption %></span></td>
<% End If %>
<% If department_th.office_id.Visible Then ' office_id %>
		<td><span id="elh_department_th_office_id" class="department_th_office_id"><%= department_th.office_id.FldCaption %></span></td>
<% End If %>
<% If department_th.dept_name.Visible Then ' dept_name %>
		<td><span id="elh_department_th_dept_name" class="department_th_dept_name"><%= department_th.dept_name.FldCaption %></span></td>
<% End If %>
<% If department_th.dept_sort.Visible Then ' dept_sort %>
		<td><span id="elh_department_th_dept_sort" class="department_th_dept_sort"><%= department_th.dept_sort.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
department_th_delete.RecCnt = 0
department_th_delete.RowCnt = 0
Do While (Not department_th_delete.Recordset.Eof)
	department_th_delete.RecCnt = department_th_delete.RecCnt + 1
	department_th_delete.RowCnt = department_th_delete.RowCnt + 1

	' Set row properties
	Call department_th.ResetAttrs()
	department_th.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call department_th_delete.LoadRowValues(department_th_delete.Recordset)

	' Render row
	Call department_th_delete.RenderRow()
%>
	<tr<%= department_th.RowAttributes %>>
<% If department_th.dept_id.Visible Then ' dept_id %>
		<td<%= department_th.dept_id.CellAttributes %>>
<span id="el<%= department_th_delete.RowCnt %>_department_th_dept_id" class="control-group department_th_dept_id">
<span<%= department_th.dept_id.ViewAttributes %>>
<%= department_th.dept_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If department_th.office_id.Visible Then ' office_id %>
		<td<%= department_th.office_id.CellAttributes %>>
<span id="el<%= department_th_delete.RowCnt %>_department_th_office_id" class="control-group department_th_office_id">
<span<%= department_th.office_id.ViewAttributes %>>
<%= department_th.office_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If department_th.dept_name.Visible Then ' dept_name %>
		<td<%= department_th.dept_name.CellAttributes %>>
<span id="el<%= department_th_delete.RowCnt %>_department_th_dept_name" class="control-group department_th_dept_name">
<span<%= department_th.dept_name.ViewAttributes %>>
<%= department_th.dept_name.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If department_th.dept_sort.Visible Then ' dept_sort %>
		<td<%= department_th.dept_sort.CellAttributes %>>
<span id="el<%= department_th_delete.RowCnt %>_department_th_dept_sort" class="control-group department_th_dept_sort">
<span<%= department_th.dept_sort.ViewAttributes %>>
<%= department_th.dept_sort.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	department_th_delete.Recordset.MoveNext
Loop
department_th_delete.Recordset.Close
Set department_th_delete.Recordset = Nothing
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
fdepartment_thdelete.Init();
</script>
<%
department_th_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set department_th_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cdepartment_th_delete

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
		TableName = "department_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "department_th_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If department_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & department_th.TableVar & "&" ' add page token
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
		If department_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (department_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (department_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(department_th) Then Set department_th = New cdepartment_th
		Set Table = department_th

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "department_th"

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
		Set department_th = Nothing
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
		RecKeys = department_th.GetRecordKeys() ' Load record keys
		sFilter = department_th.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_department_thlist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in department_th class, department_thinfo.asp

		department_th.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			department_th.CurrentAction = Request.Form("a_delete")
		Else
			department_th.CurrentAction = "I"	' Display record
		End If
		Select Case department_th.CurrentAction
			Case "D" ' Delete
				department_th.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(department_th.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = department_th.CurrentFilter
		Call department_th.Recordset_Selecting(sFilter)
		department_th.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = department_th.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call department_th.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = department_th.KeyFilter

		' Call Row Selecting event
		Call department_th.Row_Selecting(sFilter)

		' Load sql based on filter
		department_th.CurrentFilter = sFilter
		sSql = department_th.SQL
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
		Call department_th.Row_Selected(RsRow)
		department_th.dept_id.DbValue = RsRow("dept_id")
		department_th.office_id.DbValue = RsRow("office_id")
		department_th.dept_name.DbValue = RsRow("dept_name")
		department_th.dept_sort.DbValue = RsRow("dept_sort")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		department_th.dept_id.m_DbValue = Rs("dept_id")
		department_th.office_id.m_DbValue = Rs("office_id")
		department_th.dept_name.m_DbValue = Rs("dept_name")
		department_th.dept_sort.m_DbValue = Rs("dept_sort")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call department_th.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' dept_id
		' office_id
		' dept_name
		' dept_sort
		' -----------
		'  View  Row
		' -----------

		If department_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' dept_id
			department_th.dept_id.ViewValue = department_th.dept_id.CurrentValue
			department_th.dept_id.ViewCustomAttributes = ""

			' office_id
			department_th.office_id.ViewValue = department_th.office_id.CurrentValue
			department_th.office_id.ViewCustomAttributes = ""

			' dept_name
			department_th.dept_name.ViewValue = department_th.dept_name.CurrentValue
			department_th.dept_name.ViewCustomAttributes = ""

			' dept_sort
			department_th.dept_sort.ViewValue = department_th.dept_sort.CurrentValue
			department_th.dept_sort.ViewCustomAttributes = ""

			' View refer script
			' dept_id

			department_th.dept_id.LinkCustomAttributes = ""
			department_th.dept_id.HrefValue = ""
			department_th.dept_id.TooltipValue = ""

			' office_id
			department_th.office_id.LinkCustomAttributes = ""
			department_th.office_id.HrefValue = ""
			department_th.office_id.TooltipValue = ""

			' dept_name
			department_th.dept_name.LinkCustomAttributes = ""
			department_th.dept_name.HrefValue = ""
			department_th.dept_name.TooltipValue = ""

			' dept_sort
			department_th.dept_sort.LinkCustomAttributes = ""
			department_th.dept_sort.HrefValue = ""
			department_th.dept_sort.TooltipValue = ""
		End If

		' Call Row Rendered event
		If department_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call department_th.Row_Rendered()
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
		sSql = department_th.SQL
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
				DeleteRows = department_th.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("dept_id")
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
			ElseIf department_th.CancelMessage <> "" Then
				FailureMessage = department_th.CancelMessage
				department_th.CancelMessage = ""
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
				Call department_th.Row_Deleted(RsOld)
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
		Call Breadcrumb.Add("list", department_th.TableVar, "pom_department_thlist.asp", department_th.TableVar, True)
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
