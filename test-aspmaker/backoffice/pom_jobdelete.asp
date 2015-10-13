<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_jobinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim job_delete
Set job_delete = New cjob_delete
Set Page = job_delete

' Page init processing
job_delete.Page_Init()

' Page main processing
job_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
job_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var job_delete = new ew_Page("job_delete");
job_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = job_delete.PageID; // For backward compatibility
// Form object
var fjobdelete = new ew_Form("fjobdelete");
// Form_CustomValidate event
fjobdelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fjobdelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fjobdelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set job_delete.Recordset = job_delete.LoadRecordset()
job_delete.TotalRecs = job_delete.Recordset.RecordCount ' Get record count
If job_delete.TotalRecs <= 0 Then ' No record found, exit
	job_delete.Recordset.Close
	Set job_delete.Recordset = Nothing
	Call job_delete.Page_Terminate("pom_joblist.asp") ' Return to list
End If
%>
<% If job.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% job_delete.ShowPageHeader() %>
<% job_delete.ShowMessage %>
<form name="fjobdelete" id="fjobdelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="job">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(job_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(job_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_jobdelete" class="ewTable ewTableSeparate">
<%= job.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If job.job_id.Visible Then ' job_id %>
		<td><span id="elh_job_job_id" class="job_job_id"><%= job.job_id.FldCaption %></span></td>
<% End If %>
<% If job.company_id.Visible Then ' company_id %>
		<td><span id="elh_job_company_id" class="job_company_id"><%= job.company_id.FldCaption %></span></td>
<% End If %>
<% If job.job_date.Visible Then ' job_date %>
		<td><span id="elh_job_job_date" class="job_job_date"><%= job.job_date.FldCaption %></span></td>
<% End If %>
<% If job.job_title.Visible Then ' job_title %>
		<td><span id="elh_job_job_title" class="job_job_title"><%= job.job_title.FldCaption %></span></td>
<% End If %>
<% If job.job_create.Visible Then ' job_create %>
		<td><span id="elh_job_job_create" class="job_job_create"><%= job.job_create.FldCaption %></span></td>
<% End If %>
<% If job.job_update.Visible Then ' job_update %>
		<td><span id="elh_job_job_update" class="job_job_update"><%= job.job_update.FldCaption %></span></td>
<% End If %>
<% If job.job_show.Visible Then ' job_show %>
		<td><span id="elh_job_job_show" class="job_job_show"><%= job.job_show.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
job_delete.RecCnt = 0
job_delete.RowCnt = 0
Do While (Not job_delete.Recordset.Eof)
	job_delete.RecCnt = job_delete.RecCnt + 1
	job_delete.RowCnt = job_delete.RowCnt + 1

	' Set row properties
	Call job.ResetAttrs()
	job.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call job_delete.LoadRowValues(job_delete.Recordset)

	' Render row
	Call job_delete.RenderRow()
%>
	<tr<%= job.RowAttributes %>>
<% If job.job_id.Visible Then ' job_id %>
		<td<%= job.job_id.CellAttributes %>>
<span id="el<%= job_delete.RowCnt %>_job_job_id" class="control-group job_job_id">
<span<%= job.job_id.ViewAttributes %>>
<%= job.job_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If job.company_id.Visible Then ' company_id %>
		<td<%= job.company_id.CellAttributes %>>
<span id="el<%= job_delete.RowCnt %>_job_company_id" class="control-group job_company_id">
<span<%= job.company_id.ViewAttributes %>>
<%= job.company_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If job.job_date.Visible Then ' job_date %>
		<td<%= job.job_date.CellAttributes %>>
<span id="el<%= job_delete.RowCnt %>_job_job_date" class="control-group job_job_date">
<span<%= job.job_date.ViewAttributes %>>
<%= job.job_date.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If job.job_title.Visible Then ' job_title %>
		<td<%= job.job_title.CellAttributes %>>
<span id="el<%= job_delete.RowCnt %>_job_job_title" class="control-group job_job_title">
<span<%= job.job_title.ViewAttributes %>>
<%= job.job_title.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If job.job_create.Visible Then ' job_create %>
		<td<%= job.job_create.CellAttributes %>>
<span id="el<%= job_delete.RowCnt %>_job_job_create" class="control-group job_job_create">
<span<%= job.job_create.ViewAttributes %>>
<%= job.job_create.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If job.job_update.Visible Then ' job_update %>
		<td<%= job.job_update.CellAttributes %>>
<span id="el<%= job_delete.RowCnt %>_job_job_update" class="control-group job_job_update">
<span<%= job.job_update.ViewAttributes %>>
<%= job.job_update.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If job.job_show.Visible Then ' job_show %>
		<td<%= job.job_show.CellAttributes %>>
<span id="el<%= job_delete.RowCnt %>_job_job_show" class="control-group job_job_show">
<span<%= job.job_show.ViewAttributes %>>
<%= job.job_show.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	job_delete.Recordset.MoveNext
Loop
job_delete.Recordset.Close
Set job_delete.Recordset = Nothing
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
fjobdelete.Init();
</script>
<%
job_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set job_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cjob_delete

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
		TableName = "job"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "job_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If job.UseTokenInUrl Then PageUrl = PageUrl & "t=" & job.TableVar & "&" ' add page token
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
		If job.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (job.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (job.TableVar = Request.QueryString("t"))
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
		If IsEmpty(job) Then Set job = New cjob
		Set Table = job

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "job"

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
		Set job = Nothing
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
		RecKeys = job.GetRecordKeys() ' Load record keys
		sFilter = job.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_joblist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in job class, jobinfo.asp

		job.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			job.CurrentAction = Request.Form("a_delete")
		Else
			job.CurrentAction = "I"	' Display record
		End If
		Select Case job.CurrentAction
			Case "D" ' Delete
				job.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(job.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = job.CurrentFilter
		Call job.Recordset_Selecting(sFilter)
		job.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = job.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call job.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = job.KeyFilter

		' Call Row Selecting event
		Call job.Row_Selecting(sFilter)

		' Load sql based on filter
		job.CurrentFilter = sFilter
		sSql = job.SQL
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
		Call job.Row_Selected(RsRow)
		job.job_id.DbValue = RsRow("job_id")
		job.company_id.DbValue = RsRow("company_id")
		job.job_date.DbValue = RsRow("job_date")
		job.job_title.DbValue = RsRow("job_title")
		job.job_intro.DbValue = RsRow("job_intro")
		job.job_detail.DbValue = RsRow("job_detail")
		job.job_create.DbValue = RsRow("job_create")
		job.job_update.DbValue = RsRow("job_update")
		job.job_show.DbValue = RsRow("job_show")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		job.job_id.m_DbValue = Rs("job_id")
		job.company_id.m_DbValue = Rs("company_id")
		job.job_date.m_DbValue = Rs("job_date")
		job.job_title.m_DbValue = Rs("job_title")
		job.job_intro.m_DbValue = Rs("job_intro")
		job.job_detail.m_DbValue = Rs("job_detail")
		job.job_create.m_DbValue = Rs("job_create")
		job.job_update.m_DbValue = Rs("job_update")
		job.job_show.m_DbValue = Rs("job_show")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call job.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' job_id
		' company_id
		' job_date
		' job_title
		' job_intro
		' job_detail
		' job_create
		' job_update
		' job_show
		' -----------
		'  View  Row
		' -----------

		If job.RowType = EW_ROWTYPE_VIEW Then ' View row

			' job_id
			job.job_id.ViewValue = job.job_id.CurrentValue
			job.job_id.ViewCustomAttributes = ""

			' company_id
			job.company_id.ViewValue = job.company_id.CurrentValue
			job.company_id.ViewCustomAttributes = ""

			' job_date
			job.job_date.ViewValue = job.job_date.CurrentValue
			job.job_date.ViewCustomAttributes = ""

			' job_title
			job.job_title.ViewValue = job.job_title.CurrentValue
			job.job_title.ViewCustomAttributes = ""

			' job_create
			job.job_create.ViewValue = job.job_create.CurrentValue
			job.job_create.ViewCustomAttributes = ""

			' job_update
			job.job_update.ViewValue = job.job_update.CurrentValue
			job.job_update.ViewCustomAttributes = ""

			' job_show
			job.job_show.ViewValue = job.job_show.CurrentValue
			job.job_show.ViewCustomAttributes = ""

			' View refer script
			' job_id

			job.job_id.LinkCustomAttributes = ""
			job.job_id.HrefValue = ""
			job.job_id.TooltipValue = ""

			' company_id
			job.company_id.LinkCustomAttributes = ""
			job.company_id.HrefValue = ""
			job.company_id.TooltipValue = ""

			' job_date
			job.job_date.LinkCustomAttributes = ""
			job.job_date.HrefValue = ""
			job.job_date.TooltipValue = ""

			' job_title
			job.job_title.LinkCustomAttributes = ""
			job.job_title.HrefValue = ""
			job.job_title.TooltipValue = ""

			' job_create
			job.job_create.LinkCustomAttributes = ""
			job.job_create.HrefValue = ""
			job.job_create.TooltipValue = ""

			' job_update
			job.job_update.LinkCustomAttributes = ""
			job.job_update.HrefValue = ""
			job.job_update.TooltipValue = ""

			' job_show
			job.job_show.LinkCustomAttributes = ""
			job.job_show.HrefValue = ""
			job.job_show.TooltipValue = ""
		End If

		' Call Row Rendered event
		If job.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call job.Row_Rendered()
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
		sSql = job.SQL
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
				DeleteRows = job.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("job_id")
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
			ElseIf job.CancelMessage <> "" Then
				FailureMessage = job.CancelMessage
				job.CancelMessage = ""
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
				Call job.Row_Deleted(RsOld)
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
		Call Breadcrumb.Add("list", job.TableVar, "pom_joblist.asp", job.TableVar, True)
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
