<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim admins_delete
Set admins_delete = New cadmins_delete
Set Page = admins_delete

' Page init processing
admins_delete.Page_Init()

' Page main processing
admins_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
admins_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var admins_delete = new ew_Page("admins_delete");
admins_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = admins_delete.PageID; // For backward compatibility
// Form object
var fadminsdelete = new ew_Form("fadminsdelete");
// Form_CustomValidate event
fadminsdelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fadminsdelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fadminsdelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set admins_delete.Recordset = admins_delete.LoadRecordset()
admins_delete.TotalRecs = admins_delete.Recordset.RecordCount ' Get record count
If admins_delete.TotalRecs <= 0 Then ' No record found, exit
	admins_delete.Recordset.Close
	Set admins_delete.Recordset = Nothing
	Call admins_delete.Page_Terminate("pom_adminslist.asp") ' Return to list
End If
%>
<% If admins.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% admins_delete.ShowPageHeader() %>
<% admins_delete.ShowMessage %>
<form name="fadminsdelete" id="fadminsdelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="admins">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(admins_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(admins_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_adminsdelete" class="ewTable ewTableSeparate">
<%= admins.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If admins.admin_id.Visible Then ' admin_id %>
		<td><span id="elh_admins_admin_id" class="admins_admin_id"><%= admins.admin_id.FldCaption %></span></td>
<% End If %>
<% If admins.admin_username.Visible Then ' admin_username %>
		<td><span id="elh_admins_admin_username" class="admins_admin_username"><%= admins.admin_username.FldCaption %></span></td>
<% End If %>
<% If admins.admin_password.Visible Then ' admin_password %>
		<td><span id="elh_admins_admin_password" class="admins_admin_password"><%= admins.admin_password.FldCaption %></span></td>
<% End If %>
<% If admins.admin_name.Visible Then ' admin_name %>
		<td><span id="elh_admins_admin_name" class="admins_admin_name"><%= admins.admin_name.FldCaption %></span></td>
<% End If %>
<% If admins.admin_email.Visible Then ' admin_email %>
		<td><span id="elh_admins_admin_email" class="admins_admin_email"><%= admins.admin_email.FldCaption %></span></td>
<% End If %>
<% If admins.admin_tel.Visible Then ' admin_tel %>
		<td><span id="elh_admins_admin_tel" class="admins_admin_tel"><%= admins.admin_tel.FldCaption %></span></td>
<% End If %>
<% If admins.admin_permis.Visible Then ' admin_permis %>
		<td><span id="elh_admins_admin_permis" class="admins_admin_permis"><%= admins.admin_permis.FldCaption %></span></td>
<% End If %>
<% If admins.admin_create.Visible Then ' admin_create %>
		<td><span id="elh_admins_admin_create" class="admins_admin_create"><%= admins.admin_create.FldCaption %></span></td>
<% End If %>
<% If admins.admin_update.Visible Then ' admin_update %>
		<td><span id="elh_admins_admin_update" class="admins_admin_update"><%= admins.admin_update.FldCaption %></span></td>
<% End If %>
<% If admins.last_online.Visible Then ' last_online %>
		<td><span id="elh_admins_last_online" class="admins_last_online"><%= admins.last_online.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
admins_delete.RecCnt = 0
admins_delete.RowCnt = 0
Do While (Not admins_delete.Recordset.Eof)
	admins_delete.RecCnt = admins_delete.RecCnt + 1
	admins_delete.RowCnt = admins_delete.RowCnt + 1

	' Set row properties
	Call admins.ResetAttrs()
	admins.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call admins_delete.LoadRowValues(admins_delete.Recordset)

	' Render row
	Call admins_delete.RenderRow()
%>
	<tr<%= admins.RowAttributes %>>
<% If admins.admin_id.Visible Then ' admin_id %>
		<td<%= admins.admin_id.CellAttributes %>>
<span id="el<%= admins_delete.RowCnt %>_admins_admin_id" class="control-group admins_admin_id">
<span<%= admins.admin_id.ViewAttributes %>>
<%= admins.admin_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If admins.admin_username.Visible Then ' admin_username %>
		<td<%= admins.admin_username.CellAttributes %>>
<span id="el<%= admins_delete.RowCnt %>_admins_admin_username" class="control-group admins_admin_username">
<span<%= admins.admin_username.ViewAttributes %>>
<%= admins.admin_username.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If admins.admin_password.Visible Then ' admin_password %>
		<td<%= admins.admin_password.CellAttributes %>>
<span id="el<%= admins_delete.RowCnt %>_admins_admin_password" class="control-group admins_admin_password">
<span<%= admins.admin_password.ViewAttributes %>>
<%= admins.admin_password.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If admins.admin_name.Visible Then ' admin_name %>
		<td<%= admins.admin_name.CellAttributes %>>
<span id="el<%= admins_delete.RowCnt %>_admins_admin_name" class="control-group admins_admin_name">
<span<%= admins.admin_name.ViewAttributes %>>
<%= admins.admin_name.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If admins.admin_email.Visible Then ' admin_email %>
		<td<%= admins.admin_email.CellAttributes %>>
<span id="el<%= admins_delete.RowCnt %>_admins_admin_email" class="control-group admins_admin_email">
<span<%= admins.admin_email.ViewAttributes %>>
<%= admins.admin_email.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If admins.admin_tel.Visible Then ' admin_tel %>
		<td<%= admins.admin_tel.CellAttributes %>>
<span id="el<%= admins_delete.RowCnt %>_admins_admin_tel" class="control-group admins_admin_tel">
<span<%= admins.admin_tel.ViewAttributes %>>
<%= admins.admin_tel.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If admins.admin_permis.Visible Then ' admin_permis %>
		<td<%= admins.admin_permis.CellAttributes %>>
<span id="el<%= admins_delete.RowCnt %>_admins_admin_permis" class="control-group admins_admin_permis">
<span<%= admins.admin_permis.ViewAttributes %>>
<%= admins.admin_permis.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If admins.admin_create.Visible Then ' admin_create %>
		<td<%= admins.admin_create.CellAttributes %>>
<span id="el<%= admins_delete.RowCnt %>_admins_admin_create" class="control-group admins_admin_create">
<span<%= admins.admin_create.ViewAttributes %>>
<%= admins.admin_create.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If admins.admin_update.Visible Then ' admin_update %>
		<td<%= admins.admin_update.CellAttributes %>>
<span id="el<%= admins_delete.RowCnt %>_admins_admin_update" class="control-group admins_admin_update">
<span<%= admins.admin_update.ViewAttributes %>>
<%= admins.admin_update.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If admins.last_online.Visible Then ' last_online %>
		<td<%= admins.last_online.CellAttributes %>>
<span id="el<%= admins_delete.RowCnt %>_admins_last_online" class="control-group admins_last_online">
<span<%= admins.last_online.ViewAttributes %>>
<%= admins.last_online.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	admins_delete.Recordset.MoveNext
Loop
admins_delete.Recordset.Close
Set admins_delete.Recordset = Nothing
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
fadminsdelete.Init();
</script>
<%
admins_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set admins_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cadmins_delete

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
		TableName = "admins"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "admins_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If admins.UseTokenInUrl Then PageUrl = PageUrl & "t=" & admins.TableVar & "&" ' add page token
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
		If admins.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (admins.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (admins.TableVar = Request.QueryString("t"))
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
		If IsEmpty(admins) Then Set admins = New cadmins
		Set Table = admins

		' Initialize urls
		' Initialize form object

		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "admins"

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
		Set admins = Nothing
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
		RecKeys = admins.GetRecordKeys() ' Load record keys
		sFilter = admins.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_adminslist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in admins class, adminsinfo.asp

		admins.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			admins.CurrentAction = Request.Form("a_delete")
		Else
			admins.CurrentAction = "I"	' Display record
		End If
		Select Case admins.CurrentAction
			Case "D" ' Delete
				admins.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(admins.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = admins.CurrentFilter
		Call admins.Recordset_Selecting(sFilter)
		admins.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = admins.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call admins.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = admins.KeyFilter

		' Call Row Selecting event
		Call admins.Row_Selecting(sFilter)

		' Load sql based on filter
		admins.CurrentFilter = sFilter
		sSql = admins.SQL
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
		Call admins.Row_Selected(RsRow)
		admins.admin_id.DbValue = RsRow("admin_id")
		admins.admin_username.DbValue = RsRow("admin_username")
		admins.admin_password.DbValue = RsRow("admin_password")
		admins.admin_name.DbValue = RsRow("admin_name")
		admins.admin_email.DbValue = RsRow("admin_email")
		admins.admin_tel.DbValue = RsRow("admin_tel")
		admins.admin_permis.DbValue = RsRow("admin_permis")
		admins.admin_create.DbValue = RsRow("admin_create")
		admins.admin_update.DbValue = RsRow("admin_update")
		admins.last_online.DbValue = RsRow("last_online")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		admins.admin_id.m_DbValue = Rs("admin_id")
		admins.admin_username.m_DbValue = Rs("admin_username")
		admins.admin_password.m_DbValue = Rs("admin_password")
		admins.admin_name.m_DbValue = Rs("admin_name")
		admins.admin_email.m_DbValue = Rs("admin_email")
		admins.admin_tel.m_DbValue = Rs("admin_tel")
		admins.admin_permis.m_DbValue = Rs("admin_permis")
		admins.admin_create.m_DbValue = Rs("admin_create")
		admins.admin_update.m_DbValue = Rs("admin_update")
		admins.last_online.m_DbValue = Rs("last_online")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call admins.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' admin_id
		' admin_username
		' admin_password
		' admin_name
		' admin_email
		' admin_tel
		' admin_permis
		' admin_create
		' admin_update
		' last_online
		' -----------
		'  View  Row
		' -----------

		If admins.RowType = EW_ROWTYPE_VIEW Then ' View row

			' admin_id
			admins.admin_id.ViewValue = admins.admin_id.CurrentValue
			admins.admin_id.ViewCustomAttributes = ""

			' admin_username
			admins.admin_username.ViewValue = admins.admin_username.CurrentValue
			admins.admin_username.ViewCustomAttributes = ""

			' admin_password
			admins.admin_password.ViewValue = admins.admin_password.CurrentValue
			admins.admin_password.ViewCustomAttributes = ""

			' admin_name
			admins.admin_name.ViewValue = admins.admin_name.CurrentValue
			admins.admin_name.ViewCustomAttributes = ""

			' admin_email
			admins.admin_email.ViewValue = admins.admin_email.CurrentValue
			admins.admin_email.ViewCustomAttributes = ""

			' admin_tel
			admins.admin_tel.ViewValue = admins.admin_tel.CurrentValue
			admins.admin_tel.ViewCustomAttributes = ""

			' admin_permis
			admins.admin_permis.ViewValue = admins.admin_permis.CurrentValue
			admins.admin_permis.ViewCustomAttributes = ""

			' admin_create
			admins.admin_create.ViewValue = admins.admin_create.CurrentValue
			admins.admin_create.ViewCustomAttributes = ""

			' admin_update
			admins.admin_update.ViewValue = admins.admin_update.CurrentValue
			admins.admin_update.ViewCustomAttributes = ""

			' last_online
			admins.last_online.ViewValue = admins.last_online.CurrentValue
			admins.last_online.ViewCustomAttributes = ""

			' View refer script
			' admin_id

			admins.admin_id.LinkCustomAttributes = ""
			admins.admin_id.HrefValue = ""
			admins.admin_id.TooltipValue = ""

			' admin_username
			admins.admin_username.LinkCustomAttributes = ""
			admins.admin_username.HrefValue = ""
			admins.admin_username.TooltipValue = ""

			' admin_password
			admins.admin_password.LinkCustomAttributes = ""
			admins.admin_password.HrefValue = ""
			admins.admin_password.TooltipValue = ""

			' admin_name
			admins.admin_name.LinkCustomAttributes = ""
			admins.admin_name.HrefValue = ""
			admins.admin_name.TooltipValue = ""

			' admin_email
			admins.admin_email.LinkCustomAttributes = ""
			admins.admin_email.HrefValue = ""
			admins.admin_email.TooltipValue = ""

			' admin_tel
			admins.admin_tel.LinkCustomAttributes = ""
			admins.admin_tel.HrefValue = ""
			admins.admin_tel.TooltipValue = ""

			' admin_permis
			admins.admin_permis.LinkCustomAttributes = ""
			admins.admin_permis.HrefValue = ""
			admins.admin_permis.TooltipValue = ""

			' admin_create
			admins.admin_create.LinkCustomAttributes = ""
			admins.admin_create.HrefValue = ""
			admins.admin_create.TooltipValue = ""

			' admin_update
			admins.admin_update.LinkCustomAttributes = ""
			admins.admin_update.HrefValue = ""
			admins.admin_update.TooltipValue = ""

			' last_online
			admins.last_online.LinkCustomAttributes = ""
			admins.last_online.HrefValue = ""
			admins.last_online.TooltipValue = ""
		End If

		' Call Row Rendered event
		If admins.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call admins.Row_Rendered()
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
		sSql = admins.SQL
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
				DeleteRows = admins.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("admin_id")
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
			ElseIf admins.CancelMessage <> "" Then
				FailureMessage = admins.CancelMessage
				admins.CancelMessage = ""
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
				Call admins.Row_Deleted(RsOld)
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
		Call Breadcrumb.Add("list", admins.TableVar, "pom_adminslist.asp", admins.TableVar, True)
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
