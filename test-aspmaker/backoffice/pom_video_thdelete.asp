<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_video_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim video_th_delete
Set video_th_delete = New cvideo_th_delete
Set Page = video_th_delete

' Page init processing
video_th_delete.Page_Init()

' Page main processing
video_th_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
video_th_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var video_th_delete = new ew_Page("video_th_delete");
video_th_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = video_th_delete.PageID; // For backward compatibility
// Form object
var fvideo_thdelete = new ew_Form("fvideo_thdelete");
// Form_CustomValidate event
fvideo_thdelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fvideo_thdelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fvideo_thdelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set video_th_delete.Recordset = video_th_delete.LoadRecordset()
video_th_delete.TotalRecs = video_th_delete.Recordset.RecordCount ' Get record count
If video_th_delete.TotalRecs <= 0 Then ' No record found, exit
	video_th_delete.Recordset.Close
	Set video_th_delete.Recordset = Nothing
	Call video_th_delete.Page_Terminate("pom_video_thlist.asp") ' Return to list
End If
%>
<% If video_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% video_th_delete.ShowPageHeader() %>
<% video_th_delete.ShowMessage %>
<form name="fvideo_thdelete" id="fvideo_thdelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="video_th">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(video_th_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(video_th_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_video_thdelete" class="ewTable ewTableSeparate">
<%= video_th.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If video_th.video_id.Visible Then ' video_id %>
		<td><span id="elh_video_th_video_id" class="video_th_video_id"><%= video_th.video_id.FldCaption %></span></td>
<% End If %>
<% If video_th.video_title.Visible Then ' video_title %>
		<td><span id="elh_video_th_video_title" class="video_th_video_title"><%= video_th.video_title.FldCaption %></span></td>
<% End If %>
<% If video_th.video_link.Visible Then ' video_link %>
		<td><span id="elh_video_th_video_link" class="video_th_video_link"><%= video_th.video_link.FldCaption %></span></td>
<% End If %>
<% If video_th.video_create.Visible Then ' video_create %>
		<td><span id="elh_video_th_video_create" class="video_th_video_create"><%= video_th.video_create.FldCaption %></span></td>
<% End If %>
<% If video_th.video_update.Visible Then ' video_update %>
		<td><span id="elh_video_th_video_update" class="video_th_video_update"><%= video_th.video_update.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
video_th_delete.RecCnt = 0
video_th_delete.RowCnt = 0
Do While (Not video_th_delete.Recordset.Eof)
	video_th_delete.RecCnt = video_th_delete.RecCnt + 1
	video_th_delete.RowCnt = video_th_delete.RowCnt + 1

	' Set row properties
	Call video_th.ResetAttrs()
	video_th.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call video_th_delete.LoadRowValues(video_th_delete.Recordset)

	' Render row
	Call video_th_delete.RenderRow()
%>
	<tr<%= video_th.RowAttributes %>>
<% If video_th.video_id.Visible Then ' video_id %>
		<td<%= video_th.video_id.CellAttributes %>>
<span id="el<%= video_th_delete.RowCnt %>_video_th_video_id" class="control-group video_th_video_id">
<span<%= video_th.video_id.ViewAttributes %>>
<%= video_th.video_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If video_th.video_title.Visible Then ' video_title %>
		<td<%= video_th.video_title.CellAttributes %>>
<span id="el<%= video_th_delete.RowCnt %>_video_th_video_title" class="control-group video_th_video_title">
<span<%= video_th.video_title.ViewAttributes %>>
<%= video_th.video_title.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If video_th.video_link.Visible Then ' video_link %>
		<td<%= video_th.video_link.CellAttributes %>>
<span id="el<%= video_th_delete.RowCnt %>_video_th_video_link" class="control-group video_th_video_link">
<span<%= video_th.video_link.ViewAttributes %>>
<%= video_th.video_link.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If video_th.video_create.Visible Then ' video_create %>
		<td<%= video_th.video_create.CellAttributes %>>
<span id="el<%= video_th_delete.RowCnt %>_video_th_video_create" class="control-group video_th_video_create">
<span<%= video_th.video_create.ViewAttributes %>>
<%= video_th.video_create.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If video_th.video_update.Visible Then ' video_update %>
		<td<%= video_th.video_update.CellAttributes %>>
<span id="el<%= video_th_delete.RowCnt %>_video_th_video_update" class="control-group video_th_video_update">
<span<%= video_th.video_update.ViewAttributes %>>
<%= video_th.video_update.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	video_th_delete.Recordset.MoveNext
Loop
video_th_delete.Recordset.Close
Set video_th_delete.Recordset = Nothing
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
fvideo_thdelete.Init();
</script>
<%
video_th_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set video_th_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cvideo_th_delete

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
		TableName = "video_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "video_th_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If video_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & video_th.TableVar & "&" ' add page token
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
		If video_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (video_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (video_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(video_th) Then Set video_th = New cvideo_th
		Set Table = video_th

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "video_th"

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
		Set video_th = Nothing
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
		RecKeys = video_th.GetRecordKeys() ' Load record keys
		sFilter = video_th.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_video_thlist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in video_th class, video_thinfo.asp

		video_th.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			video_th.CurrentAction = Request.Form("a_delete")
		Else
			video_th.CurrentAction = "I"	' Display record
		End If
		Select Case video_th.CurrentAction
			Case "D" ' Delete
				video_th.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(video_th.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = video_th.CurrentFilter
		Call video_th.Recordset_Selecting(sFilter)
		video_th.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = video_th.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call video_th.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = video_th.KeyFilter

		' Call Row Selecting event
		Call video_th.Row_Selecting(sFilter)

		' Load sql based on filter
		video_th.CurrentFilter = sFilter
		sSql = video_th.SQL
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
		Call video_th.Row_Selected(RsRow)
		video_th.video_id.DbValue = RsRow("video_id")
		video_th.video_title.DbValue = RsRow("video_title")
		video_th.video_link.DbValue = RsRow("video_link")
		video_th.video_detail.DbValue = RsRow("video_detail")
		video_th.video_create.DbValue = RsRow("video_create")
		video_th.video_update.DbValue = RsRow("video_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		video_th.video_id.m_DbValue = Rs("video_id")
		video_th.video_title.m_DbValue = Rs("video_title")
		video_th.video_link.m_DbValue = Rs("video_link")
		video_th.video_detail.m_DbValue = Rs("video_detail")
		video_th.video_create.m_DbValue = Rs("video_create")
		video_th.video_update.m_DbValue = Rs("video_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call video_th.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' video_id
		' video_title
		' video_link
		' video_detail
		' video_create
		' video_update
		' -----------
		'  View  Row
		' -----------

		If video_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' video_id
			video_th.video_id.ViewValue = video_th.video_id.CurrentValue
			video_th.video_id.ViewCustomAttributes = ""

			' video_title
			video_th.video_title.ViewValue = video_th.video_title.CurrentValue
			video_th.video_title.ViewCustomAttributes = ""

			' video_link
			video_th.video_link.ViewValue = video_th.video_link.CurrentValue
			video_th.video_link.ViewCustomAttributes = ""

			' video_create
			video_th.video_create.ViewValue = video_th.video_create.CurrentValue
			video_th.video_create.ViewCustomAttributes = ""

			' video_update
			video_th.video_update.ViewValue = video_th.video_update.CurrentValue
			video_th.video_update.ViewCustomAttributes = ""

			' View refer script
			' video_id

			video_th.video_id.LinkCustomAttributes = ""
			video_th.video_id.HrefValue = ""
			video_th.video_id.TooltipValue = ""

			' video_title
			video_th.video_title.LinkCustomAttributes = ""
			video_th.video_title.HrefValue = ""
			video_th.video_title.TooltipValue = ""

			' video_link
			video_th.video_link.LinkCustomAttributes = ""
			video_th.video_link.HrefValue = ""
			video_th.video_link.TooltipValue = ""

			' video_create
			video_th.video_create.LinkCustomAttributes = ""
			video_th.video_create.HrefValue = ""
			video_th.video_create.TooltipValue = ""

			' video_update
			video_th.video_update.LinkCustomAttributes = ""
			video_th.video_update.HrefValue = ""
			video_th.video_update.TooltipValue = ""
		End If

		' Call Row Rendered event
		If video_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call video_th.Row_Rendered()
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
		sSql = video_th.SQL
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
				DeleteRows = video_th.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("video_id")
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
			ElseIf video_th.CancelMessage <> "" Then
				FailureMessage = video_th.CancelMessage
				video_th.CancelMessage = ""
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
				Call video_th.Row_Deleted(RsOld)
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
		Call Breadcrumb.Add("list", video_th.TableVar, "pom_video_thlist.asp", video_th.TableVar, True)
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
