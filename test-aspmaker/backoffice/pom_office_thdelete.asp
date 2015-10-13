<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_office_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim office_th_delete
Set office_th_delete = New coffice_th_delete
Set Page = office_th_delete

' Page init processing
office_th_delete.Page_Init()

' Page main processing
office_th_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
office_th_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var office_th_delete = new ew_Page("office_th_delete");
office_th_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = office_th_delete.PageID; // For backward compatibility
// Form object
var foffice_thdelete = new ew_Form("foffice_thdelete");
// Form_CustomValidate event
foffice_thdelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
foffice_thdelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
foffice_thdelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set office_th_delete.Recordset = office_th_delete.LoadRecordset()
office_th_delete.TotalRecs = office_th_delete.Recordset.RecordCount ' Get record count
If office_th_delete.TotalRecs <= 0 Then ' No record found, exit
	office_th_delete.Recordset.Close
	Set office_th_delete.Recordset = Nothing
	Call office_th_delete.Page_Terminate("pom_office_thlist.asp") ' Return to list
End If
%>
<% If office_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% office_th_delete.ShowPageHeader() %>
<% office_th_delete.ShowMessage %>
<form name="foffice_thdelete" id="foffice_thdelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="office_th">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(office_th_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(office_th_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_office_thdelete" class="ewTable ewTableSeparate">
<%= office_th.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If office_th.office_id.Visible Then ' office_id %>
		<td><span id="elh_office_th_office_id" class="office_th_office_id"><%= office_th.office_id.FldCaption %></span></td>
<% End If %>
<% If office_th.office_name_en.Visible Then ' office_name_en %>
		<td><span id="elh_office_th_office_name_en" class="office_th_office_name_en"><%= office_th.office_name_en.FldCaption %></span></td>
<% End If %>
<% If office_th.office_name_th.Visible Then ' office_name_th %>
		<td><span id="elh_office_th_office_name_th" class="office_th_office_name_th"><%= office_th.office_name_th.FldCaption %></span></td>
<% End If %>
<% If office_th.office_sort.Visible Then ' office_sort %>
		<td><span id="elh_office_th_office_sort" class="office_th_office_sort"><%= office_th.office_sort.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
office_th_delete.RecCnt = 0
office_th_delete.RowCnt = 0
Do While (Not office_th_delete.Recordset.Eof)
	office_th_delete.RecCnt = office_th_delete.RecCnt + 1
	office_th_delete.RowCnt = office_th_delete.RowCnt + 1

	' Set row properties
	Call office_th.ResetAttrs()
	office_th.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call office_th_delete.LoadRowValues(office_th_delete.Recordset)

	' Render row
	Call office_th_delete.RenderRow()
%>
	<tr<%= office_th.RowAttributes %>>
<% If office_th.office_id.Visible Then ' office_id %>
		<td<%= office_th.office_id.CellAttributes %>>
<span id="el<%= office_th_delete.RowCnt %>_office_th_office_id" class="control-group office_th_office_id">
<span<%= office_th.office_id.ViewAttributes %>>
<%= office_th.office_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If office_th.office_name_en.Visible Then ' office_name_en %>
		<td<%= office_th.office_name_en.CellAttributes %>>
<span id="el<%= office_th_delete.RowCnt %>_office_th_office_name_en" class="control-group office_th_office_name_en">
<span<%= office_th.office_name_en.ViewAttributes %>>
<%= office_th.office_name_en.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If office_th.office_name_th.Visible Then ' office_name_th %>
		<td<%= office_th.office_name_th.CellAttributes %>>
<span id="el<%= office_th_delete.RowCnt %>_office_th_office_name_th" class="control-group office_th_office_name_th">
<span<%= office_th.office_name_th.ViewAttributes %>>
<%= office_th.office_name_th.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If office_th.office_sort.Visible Then ' office_sort %>
		<td<%= office_th.office_sort.CellAttributes %>>
<span id="el<%= office_th_delete.RowCnt %>_office_th_office_sort" class="control-group office_th_office_sort">
<span<%= office_th.office_sort.ViewAttributes %>>
<%= office_th.office_sort.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	office_th_delete.Recordset.MoveNext
Loop
office_th_delete.Recordset.Close
Set office_th_delete.Recordset = Nothing
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
foffice_thdelete.Init();
</script>
<%
office_th_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set office_th_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class coffice_th_delete

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
		TableName = "office_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "office_th_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If office_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & office_th.TableVar & "&" ' add page token
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
		If office_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (office_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (office_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(office_th) Then Set office_th = New coffice_th
		Set Table = office_th

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "office_th"

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
		Set office_th = Nothing
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
		RecKeys = office_th.GetRecordKeys() ' Load record keys
		sFilter = office_th.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_office_thlist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in office_th class, office_thinfo.asp

		office_th.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			office_th.CurrentAction = Request.Form("a_delete")
		Else
			office_th.CurrentAction = "I"	' Display record
		End If
		Select Case office_th.CurrentAction
			Case "D" ' Delete
				office_th.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(office_th.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = office_th.CurrentFilter
		Call office_th.Recordset_Selecting(sFilter)
		office_th.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = office_th.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call office_th.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = office_th.KeyFilter

		' Call Row Selecting event
		Call office_th.Row_Selecting(sFilter)

		' Load sql based on filter
		office_th.CurrentFilter = sFilter
		sSql = office_th.SQL
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
		Call office_th.Row_Selected(RsRow)
		office_th.office_id.DbValue = RsRow("office_id")
		office_th.office_name_en.DbValue = RsRow("office_name_en")
		office_th.office_name_th.DbValue = RsRow("office_name_th")
		office_th.office_sort.DbValue = RsRow("office_sort")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		office_th.office_id.m_DbValue = Rs("office_id")
		office_th.office_name_en.m_DbValue = Rs("office_name_en")
		office_th.office_name_th.m_DbValue = Rs("office_name_th")
		office_th.office_sort.m_DbValue = Rs("office_sort")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call office_th.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' office_id
		' office_name_en
		' office_name_th
		' office_sort
		' -----------
		'  View  Row
		' -----------

		If office_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' office_id
			office_th.office_id.ViewValue = office_th.office_id.CurrentValue
			office_th.office_id.ViewCustomAttributes = ""

			' office_name_en
			office_th.office_name_en.ViewValue = office_th.office_name_en.CurrentValue
			office_th.office_name_en.ViewCustomAttributes = ""

			' office_name_th
			office_th.office_name_th.ViewValue = office_th.office_name_th.CurrentValue
			office_th.office_name_th.ViewCustomAttributes = ""

			' office_sort
			office_th.office_sort.ViewValue = office_th.office_sort.CurrentValue
			office_th.office_sort.ViewCustomAttributes = ""

			' View refer script
			' office_id

			office_th.office_id.LinkCustomAttributes = ""
			office_th.office_id.HrefValue = ""
			office_th.office_id.TooltipValue = ""

			' office_name_en
			office_th.office_name_en.LinkCustomAttributes = ""
			office_th.office_name_en.HrefValue = ""
			office_th.office_name_en.TooltipValue = ""

			' office_name_th
			office_th.office_name_th.LinkCustomAttributes = ""
			office_th.office_name_th.HrefValue = ""
			office_th.office_name_th.TooltipValue = ""

			' office_sort
			office_th.office_sort.LinkCustomAttributes = ""
			office_th.office_sort.HrefValue = ""
			office_th.office_sort.TooltipValue = ""
		End If

		' Call Row Rendered event
		If office_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call office_th.Row_Rendered()
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
		sSql = office_th.SQL
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
				DeleteRows = office_th.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("office_id")
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
			ElseIf office_th.CancelMessage <> "" Then
				FailureMessage = office_th.CancelMessage
				office_th.CancelMessage = ""
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
				Call office_th.Row_Deleted(RsOld)
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
		Call Breadcrumb.Add("list", office_th.TableVar, "pom_office_thlist.asp", office_th.TableVar, True)
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
