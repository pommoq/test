<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_homepageinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim homepage_delete
Set homepage_delete = New chomepage_delete
Set Page = homepage_delete

' Page init processing
homepage_delete.Page_Init()

' Page main processing
homepage_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
homepage_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var homepage_delete = new ew_Page("homepage_delete");
homepage_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = homepage_delete.PageID; // For backward compatibility
// Form object
var fhomepagedelete = new ew_Form("fhomepagedelete");
// Form_CustomValidate event
fhomepagedelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fhomepagedelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fhomepagedelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set homepage_delete.Recordset = homepage_delete.LoadRecordset()
homepage_delete.TotalRecs = homepage_delete.Recordset.RecordCount ' Get record count
If homepage_delete.TotalRecs <= 0 Then ' No record found, exit
	homepage_delete.Recordset.Close
	Set homepage_delete.Recordset = Nothing
	Call homepage_delete.Page_Terminate("pom_homepagelist.asp") ' Return to list
End If
%>
<% If homepage.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% homepage_delete.ShowPageHeader() %>
<% homepage_delete.ShowMessage %>
<form name="fhomepagedelete" id="fhomepagedelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="homepage">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(homepage_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(homepage_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_homepagedelete" class="ewTable ewTableSeparate">
<%= homepage.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If homepage.hp_id.Visible Then ' hp_id %>
		<td><span id="elh_homepage_hp_id" class="homepage_hp_id"><%= homepage.hp_id.FldCaption %></span></td>
<% End If %>
<% If homepage.hp_img.Visible Then ' hp_img %>
		<td><span id="elh_homepage_hp_img" class="homepage_hp_img"><%= homepage.hp_img.FldCaption %></span></td>
<% End If %>
<% If homepage.hp_content.Visible Then ' hp_content %>
		<td><span id="elh_homepage_hp_content" class="homepage_hp_content"><%= homepage.hp_content.FldCaption %></span></td>
<% End If %>
<% If homepage.hp_show.Visible Then ' hp_show %>
		<td><span id="elh_homepage_hp_show" class="homepage_hp_show"><%= homepage.hp_show.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
homepage_delete.RecCnt = 0
homepage_delete.RowCnt = 0
Do While (Not homepage_delete.Recordset.Eof)
	homepage_delete.RecCnt = homepage_delete.RecCnt + 1
	homepage_delete.RowCnt = homepage_delete.RowCnt + 1

	' Set row properties
	Call homepage.ResetAttrs()
	homepage.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call homepage_delete.LoadRowValues(homepage_delete.Recordset)

	' Render row
	Call homepage_delete.RenderRow()
%>
	<tr<%= homepage.RowAttributes %>>
<% If homepage.hp_id.Visible Then ' hp_id %>
		<td<%= homepage.hp_id.CellAttributes %>>
<span id="el<%= homepage_delete.RowCnt %>_homepage_hp_id" class="control-group homepage_hp_id">
<span<%= homepage.hp_id.ViewAttributes %>>
<%= homepage.hp_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If homepage.hp_img.Visible Then ' hp_img %>
		<td<%= homepage.hp_img.CellAttributes %>>
<span id="el<%= homepage_delete.RowCnt %>_homepage_hp_img" class="control-group homepage_hp_img">
<span<%= homepage.hp_img.ViewAttributes %>>
<%= homepage.hp_img.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If homepage.hp_content.Visible Then ' hp_content %>
		<td<%= homepage.hp_content.CellAttributes %>>
<span id="el<%= homepage_delete.RowCnt %>_homepage_hp_content" class="control-group homepage_hp_content">
<span<%= homepage.hp_content.ViewAttributes %>>
<%= homepage.hp_content.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If homepage.hp_show.Visible Then ' hp_show %>
		<td<%= homepage.hp_show.CellAttributes %>>
<span id="el<%= homepage_delete.RowCnt %>_homepage_hp_show" class="control-group homepage_hp_show">
<span<%= homepage.hp_show.ViewAttributes %>>
<%= homepage.hp_show.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	homepage_delete.Recordset.MoveNext
Loop
homepage_delete.Recordset.Close
Set homepage_delete.Recordset = Nothing
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
fhomepagedelete.Init();
</script>
<%
homepage_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set homepage_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class chomepage_delete

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
		TableName = "homepage"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "homepage_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If homepage.UseTokenInUrl Then PageUrl = PageUrl & "t=" & homepage.TableVar & "&" ' add page token
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
		If homepage.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (homepage.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (homepage.TableVar = Request.QueryString("t"))
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
		If IsEmpty(homepage) Then Set homepage = New chomepage
		Set Table = homepage

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "homepage"

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
		Set homepage = Nothing
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
		RecKeys = homepage.GetRecordKeys() ' Load record keys
		sFilter = homepage.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_homepagelist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in homepage class, homepageinfo.asp

		homepage.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			homepage.CurrentAction = Request.Form("a_delete")
		Else
			homepage.CurrentAction = "I"	' Display record
		End If
		Select Case homepage.CurrentAction
			Case "D" ' Delete
				homepage.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(homepage.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = homepage.CurrentFilter
		Call homepage.Recordset_Selecting(sFilter)
		homepage.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = homepage.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call homepage.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = homepage.KeyFilter

		' Call Row Selecting event
		Call homepage.Row_Selecting(sFilter)

		' Load sql based on filter
		homepage.CurrentFilter = sFilter
		sSql = homepage.SQL
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
		Call homepage.Row_Selected(RsRow)
		homepage.hp_id.DbValue = RsRow("hp_id")
		homepage.hp_img.DbValue = RsRow("hp_img")
		homepage.hp_content.DbValue = RsRow("hp_content")
		homepage.hp_show.DbValue = RsRow("hp_show")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		homepage.hp_id.m_DbValue = Rs("hp_id")
		homepage.hp_img.m_DbValue = Rs("hp_img")
		homepage.hp_content.m_DbValue = Rs("hp_content")
		homepage.hp_show.m_DbValue = Rs("hp_show")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call homepage.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' hp_id
		' hp_img
		' hp_content
		' hp_show
		' -----------
		'  View  Row
		' -----------

		If homepage.RowType = EW_ROWTYPE_VIEW Then ' View row

			' hp_id
			homepage.hp_id.ViewValue = homepage.hp_id.CurrentValue
			homepage.hp_id.ViewCustomAttributes = ""

			' hp_img
			homepage.hp_img.ViewValue = homepage.hp_img.CurrentValue
			homepage.hp_img.ViewCustomAttributes = ""

			' hp_content
			homepage.hp_content.ViewValue = homepage.hp_content.CurrentValue
			homepage.hp_content.ViewCustomAttributes = ""

			' hp_show
			homepage.hp_show.ViewValue = homepage.hp_show.CurrentValue
			homepage.hp_show.ViewCustomAttributes = ""

			' View refer script
			' hp_id

			homepage.hp_id.LinkCustomAttributes = ""
			homepage.hp_id.HrefValue = ""
			homepage.hp_id.TooltipValue = ""

			' hp_img
			homepage.hp_img.LinkCustomAttributes = ""
			homepage.hp_img.HrefValue = ""
			homepage.hp_img.TooltipValue = ""

			' hp_content
			homepage.hp_content.LinkCustomAttributes = ""
			homepage.hp_content.HrefValue = ""
			homepage.hp_content.TooltipValue = ""

			' hp_show
			homepage.hp_show.LinkCustomAttributes = ""
			homepage.hp_show.HrefValue = ""
			homepage.hp_show.TooltipValue = ""
		End If

		' Call Row Rendered event
		If homepage.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call homepage.Row_Rendered()
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
		sSql = homepage.SQL
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
				DeleteRows = homepage.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("hp_id")
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
			ElseIf homepage.CancelMessage <> "" Then
				FailureMessage = homepage.CancelMessage
				homepage.CancelMessage = ""
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
				Call homepage.Row_Deleted(RsOld)
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
		Call Breadcrumb.Add("list", homepage.TableVar, "pom_homepagelist.asp", homepage.TableVar, True)
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
