<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_bannerinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim banner_delete
Set banner_delete = New cbanner_delete
Set Page = banner_delete

' Page init processing
banner_delete.Page_Init()

' Page main processing
banner_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
banner_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var banner_delete = new ew_Page("banner_delete");
banner_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = banner_delete.PageID; // For backward compatibility
// Form object
var fbannerdelete = new ew_Form("fbannerdelete");
// Form_CustomValidate event
fbannerdelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fbannerdelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fbannerdelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set banner_delete.Recordset = banner_delete.LoadRecordset()
banner_delete.TotalRecs = banner_delete.Recordset.RecordCount ' Get record count
If banner_delete.TotalRecs <= 0 Then ' No record found, exit
	banner_delete.Recordset.Close
	Set banner_delete.Recordset = Nothing
	Call banner_delete.Page_Terminate("pom_bannerlist.asp") ' Return to list
End If
%>
<% If banner.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% banner_delete.ShowPageHeader() %>
<% banner_delete.ShowMessage %>
<form name="fbannerdelete" id="fbannerdelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="banner">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(banner_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(banner_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_bannerdelete" class="ewTable ewTableSeparate">
<%= banner.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If banner.banner_id.Visible Then ' banner_id %>
		<td><span id="elh_banner_banner_id" class="banner_banner_id"><%= banner.banner_id.FldCaption %></span></td>
<% End If %>
<% If banner.banner_img.Visible Then ' banner_img %>
		<td><span id="elh_banner_banner_img" class="banner_banner_img"><%= banner.banner_img.FldCaption %></span></td>
<% End If %>
<% If banner.banner_link.Visible Then ' banner_link %>
		<td><span id="elh_banner_banner_link" class="banner_banner_link"><%= banner.banner_link.FldCaption %></span></td>
<% End If %>
<% If banner.banner_sort.Visible Then ' banner_sort %>
		<td><span id="elh_banner_banner_sort" class="banner_banner_sort"><%= banner.banner_sort.FldCaption %></span></td>
<% End If %>
<% If banner.start_date.Visible Then ' start_date %>
		<td><span id="elh_banner_start_date" class="banner_start_date"><%= banner.start_date.FldCaption %></span></td>
<% End If %>
<% If banner.end_date.Visible Then ' end_date %>
		<td><span id="elh_banner_end_date" class="banner_end_date"><%= banner.end_date.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
banner_delete.RecCnt = 0
banner_delete.RowCnt = 0
Do While (Not banner_delete.Recordset.Eof)
	banner_delete.RecCnt = banner_delete.RecCnt + 1
	banner_delete.RowCnt = banner_delete.RowCnt + 1

	' Set row properties
	Call banner.ResetAttrs()
	banner.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call banner_delete.LoadRowValues(banner_delete.Recordset)

	' Render row
	Call banner_delete.RenderRow()
%>
	<tr<%= banner.RowAttributes %>>
<% If banner.banner_id.Visible Then ' banner_id %>
		<td<%= banner.banner_id.CellAttributes %>>
<span id="el<%= banner_delete.RowCnt %>_banner_banner_id" class="control-group banner_banner_id">
<span<%= banner.banner_id.ViewAttributes %>>
<%= banner.banner_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If banner.banner_img.Visible Then ' banner_img %>
		<td<%= banner.banner_img.CellAttributes %>>
<span id="el<%= banner_delete.RowCnt %>_banner_banner_img" class="control-group banner_banner_img">
<span<%= banner.banner_img.ViewAttributes %>>
<%= banner.banner_img.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If banner.banner_link.Visible Then ' banner_link %>
		<td<%= banner.banner_link.CellAttributes %>>
<span id="el<%= banner_delete.RowCnt %>_banner_banner_link" class="control-group banner_banner_link">
<span<%= banner.banner_link.ViewAttributes %>>
<%= banner.banner_link.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If banner.banner_sort.Visible Then ' banner_sort %>
		<td<%= banner.banner_sort.CellAttributes %>>
<span id="el<%= banner_delete.RowCnt %>_banner_banner_sort" class="control-group banner_banner_sort">
<span<%= banner.banner_sort.ViewAttributes %>>
<%= banner.banner_sort.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If banner.start_date.Visible Then ' start_date %>
		<td<%= banner.start_date.CellAttributes %>>
<span id="el<%= banner_delete.RowCnt %>_banner_start_date" class="control-group banner_start_date">
<span<%= banner.start_date.ViewAttributes %>>
<%= banner.start_date.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If banner.end_date.Visible Then ' end_date %>
		<td<%= banner.end_date.CellAttributes %>>
<span id="el<%= banner_delete.RowCnt %>_banner_end_date" class="control-group banner_end_date">
<span<%= banner.end_date.ViewAttributes %>>
<%= banner.end_date.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	banner_delete.Recordset.MoveNext
Loop
banner_delete.Recordset.Close
Set banner_delete.Recordset = Nothing
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
fbannerdelete.Init();
</script>
<%
banner_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set banner_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cbanner_delete

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
		TableName = "banner"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "banner_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If banner.UseTokenInUrl Then PageUrl = PageUrl & "t=" & banner.TableVar & "&" ' add page token
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
		If banner.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (banner.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (banner.TableVar = Request.QueryString("t"))
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
		If IsEmpty(banner) Then Set banner = New cbanner
		Set Table = banner

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "banner"

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
		Set banner = Nothing
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
		RecKeys = banner.GetRecordKeys() ' Load record keys
		sFilter = banner.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_bannerlist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in banner class, bannerinfo.asp

		banner.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			banner.CurrentAction = Request.Form("a_delete")
		Else
			banner.CurrentAction = "I"	' Display record
		End If
		Select Case banner.CurrentAction
			Case "D" ' Delete
				banner.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(banner.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = banner.CurrentFilter
		Call banner.Recordset_Selecting(sFilter)
		banner.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = banner.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call banner.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = banner.KeyFilter

		' Call Row Selecting event
		Call banner.Row_Selecting(sFilter)

		' Load sql based on filter
		banner.CurrentFilter = sFilter
		sSql = banner.SQL
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
		Call banner.Row_Selected(RsRow)
		banner.banner_id.DbValue = RsRow("banner_id")
		banner.banner_img.DbValue = RsRow("banner_img")
		banner.banner_link.DbValue = RsRow("banner_link")
		banner.banner_sort.DbValue = RsRow("banner_sort")
		banner.start_date.DbValue = RsRow("start_date")
		banner.end_date.DbValue = RsRow("end_date")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		banner.banner_id.m_DbValue = Rs("banner_id")
		banner.banner_img.m_DbValue = Rs("banner_img")
		banner.banner_link.m_DbValue = Rs("banner_link")
		banner.banner_sort.m_DbValue = Rs("banner_sort")
		banner.start_date.m_DbValue = Rs("start_date")
		banner.end_date.m_DbValue = Rs("end_date")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call banner.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' banner_id
		' banner_img
		' banner_link
		' banner_sort
		' start_date
		' end_date
		' -----------
		'  View  Row
		' -----------

		If banner.RowType = EW_ROWTYPE_VIEW Then ' View row

			' banner_id
			banner.banner_id.ViewValue = banner.banner_id.CurrentValue
			banner.banner_id.ViewCustomAttributes = ""

			' banner_img
			banner.banner_img.ViewValue = banner.banner_img.CurrentValue
			banner.banner_img.ViewCustomAttributes = ""

			' banner_link
			banner.banner_link.ViewValue = banner.banner_link.CurrentValue
			banner.banner_link.ViewCustomAttributes = ""

			' banner_sort
			banner.banner_sort.ViewValue = banner.banner_sort.CurrentValue
			banner.banner_sort.ViewCustomAttributes = ""

			' start_date
			banner.start_date.ViewValue = banner.start_date.CurrentValue
			banner.start_date.ViewCustomAttributes = ""

			' end_date
			banner.end_date.ViewValue = banner.end_date.CurrentValue
			banner.end_date.ViewCustomAttributes = ""

			' View refer script
			' banner_id

			banner.banner_id.LinkCustomAttributes = ""
			banner.banner_id.HrefValue = ""
			banner.banner_id.TooltipValue = ""

			' banner_img
			banner.banner_img.LinkCustomAttributes = ""
			banner.banner_img.HrefValue = ""
			banner.banner_img.TooltipValue = ""

			' banner_link
			banner.banner_link.LinkCustomAttributes = ""
			banner.banner_link.HrefValue = ""
			banner.banner_link.TooltipValue = ""

			' banner_sort
			banner.banner_sort.LinkCustomAttributes = ""
			banner.banner_sort.HrefValue = ""
			banner.banner_sort.TooltipValue = ""

			' start_date
			banner.start_date.LinkCustomAttributes = ""
			banner.start_date.HrefValue = ""
			banner.start_date.TooltipValue = ""

			' end_date
			banner.end_date.LinkCustomAttributes = ""
			banner.end_date.HrefValue = ""
			banner.end_date.TooltipValue = ""
		End If

		' Call Row Rendered event
		If banner.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call banner.Row_Rendered()
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
		sSql = banner.SQL
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
				DeleteRows = banner.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("banner_id")
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
			ElseIf banner.CancelMessage <> "" Then
				FailureMessage = banner.CancelMessage
				banner.CancelMessage = ""
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
				Call banner.Row_Deleted(RsOld)
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
		Call Breadcrumb.Add("list", banner.TableVar, "pom_bannerlist.asp", banner.TableVar, True)
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
