<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_job_fileinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim job_file_delete
Set job_file_delete = New cjob_file_delete
Set Page = job_file_delete

' Page init processing
job_file_delete.Page_Init()

' Page main processing
job_file_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
job_file_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var job_file_delete = new ew_Page("job_file_delete");
job_file_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = job_file_delete.PageID; // For backward compatibility
// Form object
var fjob_filedelete = new ew_Form("fjob_filedelete");
// Form_CustomValidate event
fjob_filedelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fjob_filedelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fjob_filedelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set job_file_delete.Recordset = job_file_delete.LoadRecordset()
job_file_delete.TotalRecs = job_file_delete.Recordset.RecordCount ' Get record count
If job_file_delete.TotalRecs <= 0 Then ' No record found, exit
	job_file_delete.Recordset.Close
	Set job_file_delete.Recordset = Nothing
	Call job_file_delete.Page_Terminate("pom_job_filelist.asp") ' Return to list
End If
%>
<% If job_file.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% job_file_delete.ShowPageHeader() %>
<% job_file_delete.ShowMessage %>
<form name="fjob_filedelete" id="fjob_filedelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="job_file">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(job_file_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(job_file_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_job_filedelete" class="ewTable ewTableSeparate">
<%= job_file.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If job_file.job_file_id.Visible Then ' job_file_id %>
		<td><span id="elh_job_file_job_file_id" class="job_file_job_file_id"><%= job_file.job_file_id.FldCaption %></span></td>
<% End If %>
<% If job_file.job_id.Visible Then ' job_id %>
		<td><span id="elh_job_file_job_id" class="job_file_job_id"><%= job_file.job_id.FldCaption %></span></td>
<% End If %>
<% If job_file.job_file_name.Visible Then ' job_file_name %>
		<td><span id="elh_job_file_job_file_name" class="job_file_job_file_name"><%= job_file.job_file_name.FldCaption %></span></td>
<% End If %>
<% If job_file.job_file_title.Visible Then ' job_file_title %>
		<td><span id="elh_job_file_job_file_title" class="job_file_job_file_title"><%= job_file.job_file_title.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
job_file_delete.RecCnt = 0
job_file_delete.RowCnt = 0
Do While (Not job_file_delete.Recordset.Eof)
	job_file_delete.RecCnt = job_file_delete.RecCnt + 1
	job_file_delete.RowCnt = job_file_delete.RowCnt + 1

	' Set row properties
	Call job_file.ResetAttrs()
	job_file.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call job_file_delete.LoadRowValues(job_file_delete.Recordset)

	' Render row
	Call job_file_delete.RenderRow()
%>
	<tr<%= job_file.RowAttributes %>>
<% If job_file.job_file_id.Visible Then ' job_file_id %>
		<td<%= job_file.job_file_id.CellAttributes %>>
<span id="el<%= job_file_delete.RowCnt %>_job_file_job_file_id" class="control-group job_file_job_file_id">
<span<%= job_file.job_file_id.ViewAttributes %>>
<%= job_file.job_file_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If job_file.job_id.Visible Then ' job_id %>
		<td<%= job_file.job_id.CellAttributes %>>
<span id="el<%= job_file_delete.RowCnt %>_job_file_job_id" class="control-group job_file_job_id">
<span<%= job_file.job_id.ViewAttributes %>>
<%= job_file.job_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If job_file.job_file_name.Visible Then ' job_file_name %>
		<td<%= job_file.job_file_name.CellAttributes %>>
<span id="el<%= job_file_delete.RowCnt %>_job_file_job_file_name" class="control-group job_file_job_file_name">
<span<%= job_file.job_file_name.ViewAttributes %>>
<%= job_file.job_file_name.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If job_file.job_file_title.Visible Then ' job_file_title %>
		<td<%= job_file.job_file_title.CellAttributes %>>
<span id="el<%= job_file_delete.RowCnt %>_job_file_job_file_title" class="control-group job_file_job_file_title">
<span<%= job_file.job_file_title.ViewAttributes %>>
<%= job_file.job_file_title.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	job_file_delete.Recordset.MoveNext
Loop
job_file_delete.Recordset.Close
Set job_file_delete.Recordset = Nothing
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
fjob_filedelete.Init();
</script>
<%
job_file_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set job_file_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cjob_file_delete

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
		TableName = "job_file"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "job_file_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If job_file.UseTokenInUrl Then PageUrl = PageUrl & "t=" & job_file.TableVar & "&" ' add page token
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
		If job_file.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (job_file.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (job_file.TableVar = Request.QueryString("t"))
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
		If IsEmpty(job_file) Then Set job_file = New cjob_file
		Set Table = job_file

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "job_file"

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
		Set job_file = Nothing
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
		RecKeys = job_file.GetRecordKeys() ' Load record keys
		sFilter = job_file.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_job_filelist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in job_file class, job_fileinfo.asp

		job_file.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			job_file.CurrentAction = Request.Form("a_delete")
		Else
			job_file.CurrentAction = "I"	' Display record
		End If
		Select Case job_file.CurrentAction
			Case "D" ' Delete
				job_file.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(job_file.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = job_file.CurrentFilter
		Call job_file.Recordset_Selecting(sFilter)
		job_file.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = job_file.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call job_file.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = job_file.KeyFilter

		' Call Row Selecting event
		Call job_file.Row_Selecting(sFilter)

		' Load sql based on filter
		job_file.CurrentFilter = sFilter
		sSql = job_file.SQL
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
		Call job_file.Row_Selected(RsRow)
		job_file.job_file_id.DbValue = RsRow("job_file_id")
		job_file.job_id.DbValue = RsRow("job_id")
		job_file.job_file_name.DbValue = RsRow("job_file_name")
		job_file.job_file_title.DbValue = RsRow("job_file_title")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		job_file.job_file_id.m_DbValue = Rs("job_file_id")
		job_file.job_id.m_DbValue = Rs("job_id")
		job_file.job_file_name.m_DbValue = Rs("job_file_name")
		job_file.job_file_title.m_DbValue = Rs("job_file_title")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call job_file.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' job_file_id
		' job_id
		' job_file_name
		' job_file_title
		' -----------
		'  View  Row
		' -----------

		If job_file.RowType = EW_ROWTYPE_VIEW Then ' View row

			' job_file_id
			job_file.job_file_id.ViewValue = job_file.job_file_id.CurrentValue
			job_file.job_file_id.ViewCustomAttributes = ""

			' job_id
			job_file.job_id.ViewValue = job_file.job_id.CurrentValue
			job_file.job_id.ViewCustomAttributes = ""

			' job_file_name
			job_file.job_file_name.ViewValue = job_file.job_file_name.CurrentValue
			job_file.job_file_name.ViewCustomAttributes = ""

			' job_file_title
			job_file.job_file_title.ViewValue = job_file.job_file_title.CurrentValue
			job_file.job_file_title.ViewCustomAttributes = ""

			' View refer script
			' job_file_id

			job_file.job_file_id.LinkCustomAttributes = ""
			job_file.job_file_id.HrefValue = ""
			job_file.job_file_id.TooltipValue = ""

			' job_id
			job_file.job_id.LinkCustomAttributes = ""
			job_file.job_id.HrefValue = ""
			job_file.job_id.TooltipValue = ""

			' job_file_name
			job_file.job_file_name.LinkCustomAttributes = ""
			job_file.job_file_name.HrefValue = ""
			job_file.job_file_name.TooltipValue = ""

			' job_file_title
			job_file.job_file_title.LinkCustomAttributes = ""
			job_file.job_file_title.HrefValue = ""
			job_file.job_file_title.TooltipValue = ""
		End If

		' Call Row Rendered event
		If job_file.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call job_file.Row_Rendered()
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
		sSql = job_file.SQL
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
				DeleteRows = job_file.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("job_file_id")
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
			ElseIf job_file.CancelMessage <> "" Then
				FailureMessage = job_file.CancelMessage
				job_file.CancelMessage = ""
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
				Call job_file.Row_Deleted(RsOld)
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
		Call Breadcrumb.Add("list", job_file.TableVar, "pom_job_filelist.asp", job_file.TableVar, True)
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
