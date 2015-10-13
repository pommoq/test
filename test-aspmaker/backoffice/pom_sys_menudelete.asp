<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_sys_menuinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim sys_menu_delete
Set sys_menu_delete = New csys_menu_delete
Set Page = sys_menu_delete

' Page init processing
sys_menu_delete.Page_Init()

' Page main processing
sys_menu_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
sys_menu_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var sys_menu_delete = new ew_Page("sys_menu_delete");
sys_menu_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = sys_menu_delete.PageID; // For backward compatibility
// Form object
var fsys_menudelete = new ew_Form("fsys_menudelete");
// Form_CustomValidate event
fsys_menudelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fsys_menudelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fsys_menudelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set sys_menu_delete.Recordset = sys_menu_delete.LoadRecordset()
sys_menu_delete.TotalRecs = sys_menu_delete.Recordset.RecordCount ' Get record count
If sys_menu_delete.TotalRecs <= 0 Then ' No record found, exit
	sys_menu_delete.Recordset.Close
	Set sys_menu_delete.Recordset = Nothing
	Call sys_menu_delete.Page_Terminate("pom_sys_menulist.asp") ' Return to list
End If
%>
<% If sys_menu.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% sys_menu_delete.ShowPageHeader() %>
<% sys_menu_delete.ShowMessage %>
<form name="fsys_menudelete" id="fsys_menudelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="sys_menu">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(sys_menu_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(sys_menu_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_sys_menudelete" class="ewTable ewTableSeparate">
<%= sys_menu.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If sys_menu.menu_id.Visible Then ' menu_id %>
		<td><span id="elh_sys_menu_menu_id" class="sys_menu_menu_id"><%= sys_menu.menu_id.FldCaption %></span></td>
<% End If %>
<% If sys_menu.menu_name.Visible Then ' menu_name %>
		<td><span id="elh_sys_menu_menu_name" class="sys_menu_menu_name"><%= sys_menu.menu_name.FldCaption %></span></td>
<% End If %>
<% If sys_menu.menu_parent_id.Visible Then ' menu_parent_id %>
		<td><span id="elh_sys_menu_menu_parent_id" class="sys_menu_menu_parent_id"><%= sys_menu.menu_parent_id.FldCaption %></span></td>
<% End If %>
<% If sys_menu.menu_thai.Visible Then ' menu_thai %>
		<td><span id="elh_sys_menu_menu_thai" class="sys_menu_menu_thai"><%= sys_menu.menu_thai.FldCaption %></span></td>
<% End If %>
<% If sys_menu.menu_idname.Visible Then ' menu_idname %>
		<td><span id="elh_sys_menu_menu_idname" class="sys_menu_menu_idname"><%= sys_menu.menu_idname.FldCaption %></span></td>
<% End If %>
<% If sys_menu.menu_filename.Visible Then ' menu_filename %>
		<td><span id="elh_sys_menu_menu_filename" class="sys_menu_menu_filename"><%= sys_menu.menu_filename.FldCaption %></span></td>
<% End If %>
<% If sys_menu.target.Visible Then ' target %>
		<td><span id="elh_sys_menu_target" class="sys_menu_target"><%= sys_menu.target.FldCaption %></span></td>
<% End If %>
<% If sys_menu.OrderList.Visible Then ' OrderList %>
		<td><span id="elh_sys_menu_OrderList" class="sys_menu_OrderList"><%= sys_menu.OrderList.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
sys_menu_delete.RecCnt = 0
sys_menu_delete.RowCnt = 0
Do While (Not sys_menu_delete.Recordset.Eof)
	sys_menu_delete.RecCnt = sys_menu_delete.RecCnt + 1
	sys_menu_delete.RowCnt = sys_menu_delete.RowCnt + 1

	' Set row properties
	Call sys_menu.ResetAttrs()
	sys_menu.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call sys_menu_delete.LoadRowValues(sys_menu_delete.Recordset)

	' Render row
	Call sys_menu_delete.RenderRow()
%>
	<tr<%= sys_menu.RowAttributes %>>
<% If sys_menu.menu_id.Visible Then ' menu_id %>
		<td<%= sys_menu.menu_id.CellAttributes %>>
<span id="el<%= sys_menu_delete.RowCnt %>_sys_menu_menu_id" class="control-group sys_menu_menu_id">
<span<%= sys_menu.menu_id.ViewAttributes %>>
<%= sys_menu.menu_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If sys_menu.menu_name.Visible Then ' menu_name %>
		<td<%= sys_menu.menu_name.CellAttributes %>>
<span id="el<%= sys_menu_delete.RowCnt %>_sys_menu_menu_name" class="control-group sys_menu_menu_name">
<span<%= sys_menu.menu_name.ViewAttributes %>>
<%= sys_menu.menu_name.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If sys_menu.menu_parent_id.Visible Then ' menu_parent_id %>
		<td<%= sys_menu.menu_parent_id.CellAttributes %>>
<span id="el<%= sys_menu_delete.RowCnt %>_sys_menu_menu_parent_id" class="control-group sys_menu_menu_parent_id">
<span<%= sys_menu.menu_parent_id.ViewAttributes %>>
<%= sys_menu.menu_parent_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If sys_menu.menu_thai.Visible Then ' menu_thai %>
		<td<%= sys_menu.menu_thai.CellAttributes %>>
<span id="el<%= sys_menu_delete.RowCnt %>_sys_menu_menu_thai" class="control-group sys_menu_menu_thai">
<span<%= sys_menu.menu_thai.ViewAttributes %>>
<%= sys_menu.menu_thai.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If sys_menu.menu_idname.Visible Then ' menu_idname %>
		<td<%= sys_menu.menu_idname.CellAttributes %>>
<span id="el<%= sys_menu_delete.RowCnt %>_sys_menu_menu_idname" class="control-group sys_menu_menu_idname">
<span<%= sys_menu.menu_idname.ViewAttributes %>>
<%= sys_menu.menu_idname.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If sys_menu.menu_filename.Visible Then ' menu_filename %>
		<td<%= sys_menu.menu_filename.CellAttributes %>>
<span id="el<%= sys_menu_delete.RowCnt %>_sys_menu_menu_filename" class="control-group sys_menu_menu_filename">
<span<%= sys_menu.menu_filename.ViewAttributes %>>
<%= sys_menu.menu_filename.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If sys_menu.target.Visible Then ' target %>
		<td<%= sys_menu.target.CellAttributes %>>
<span id="el<%= sys_menu_delete.RowCnt %>_sys_menu_target" class="control-group sys_menu_target">
<span<%= sys_menu.target.ViewAttributes %>>
<%= sys_menu.target.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If sys_menu.OrderList.Visible Then ' OrderList %>
		<td<%= sys_menu.OrderList.CellAttributes %>>
<span id="el<%= sys_menu_delete.RowCnt %>_sys_menu_OrderList" class="control-group sys_menu_OrderList">
<span<%= sys_menu.OrderList.ViewAttributes %>>
<%= sys_menu.OrderList.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	sys_menu_delete.Recordset.MoveNext
Loop
sys_menu_delete.Recordset.Close
Set sys_menu_delete.Recordset = Nothing
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
fsys_menudelete.Init();
</script>
<%
sys_menu_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set sys_menu_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class csys_menu_delete

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
		TableName = "sys_menu"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "sys_menu_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If sys_menu.UseTokenInUrl Then PageUrl = PageUrl & "t=" & sys_menu.TableVar & "&" ' add page token
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
		If sys_menu.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (sys_menu.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (sys_menu.TableVar = Request.QueryString("t"))
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
		If IsEmpty(sys_menu) Then Set sys_menu = New csys_menu
		Set Table = sys_menu

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "sys_menu"

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
		Set sys_menu = Nothing
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
		RecKeys = sys_menu.GetRecordKeys() ' Load record keys
		sFilter = sys_menu.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_sys_menulist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in sys_menu class, sys_menuinfo.asp

		sys_menu.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			sys_menu.CurrentAction = Request.Form("a_delete")
		Else
			sys_menu.CurrentAction = "I"	' Display record
		End If
		Select Case sys_menu.CurrentAction
			Case "D" ' Delete
				sys_menu.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(sys_menu.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = sys_menu.CurrentFilter
		Call sys_menu.Recordset_Selecting(sFilter)
		sys_menu.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = sys_menu.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call sys_menu.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = sys_menu.KeyFilter

		' Call Row Selecting event
		Call sys_menu.Row_Selecting(sFilter)

		' Load sql based on filter
		sys_menu.CurrentFilter = sFilter
		sSql = sys_menu.SQL
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
		Call sys_menu.Row_Selected(RsRow)
		sys_menu.menu_id.DbValue = RsRow("menu_id")
		sys_menu.menu_name.DbValue = RsRow("menu_name")
		sys_menu.menu_parent_id.DbValue = RsRow("menu_parent_id")
		sys_menu.menu_thai.DbValue = RsRow("menu_thai")
		sys_menu.menu_idname.DbValue = RsRow("menu_idname")
		sys_menu.menu_filename.DbValue = RsRow("menu_filename")
		sys_menu.target.DbValue = RsRow("target")
		sys_menu.OrderList.DbValue = RsRow("OrderList")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		sys_menu.menu_id.m_DbValue = Rs("menu_id")
		sys_menu.menu_name.m_DbValue = Rs("menu_name")
		sys_menu.menu_parent_id.m_DbValue = Rs("menu_parent_id")
		sys_menu.menu_thai.m_DbValue = Rs("menu_thai")
		sys_menu.menu_idname.m_DbValue = Rs("menu_idname")
		sys_menu.menu_filename.m_DbValue = Rs("menu_filename")
		sys_menu.target.m_DbValue = Rs("target")
		sys_menu.OrderList.m_DbValue = Rs("OrderList")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Convert decimal values if posted back

		If sys_menu.OrderList.FormValue = sys_menu.OrderList.CurrentValue And IsNumeric(sys_menu.OrderList.CurrentValue) Then
			sys_menu.OrderList.CurrentValue = ew_StrToFloat(sys_menu.OrderList.CurrentValue)
		End If

		' Call Row Rendering event
		Call sys_menu.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' menu_id
		' menu_name
		' menu_parent_id
		' menu_thai
		' menu_idname
		' menu_filename
		' target
		' OrderList
		' -----------
		'  View  Row
		' -----------

		If sys_menu.RowType = EW_ROWTYPE_VIEW Then ' View row

			' menu_id
			sys_menu.menu_id.ViewValue = sys_menu.menu_id.CurrentValue
			sys_menu.menu_id.ViewCustomAttributes = ""

			' menu_name
			sys_menu.menu_name.ViewValue = sys_menu.menu_name.CurrentValue
			sys_menu.menu_name.ViewCustomAttributes = ""

			' menu_parent_id
			sys_menu.menu_parent_id.ViewValue = sys_menu.menu_parent_id.CurrentValue
			sys_menu.menu_parent_id.ViewCustomAttributes = ""

			' menu_thai
			sys_menu.menu_thai.ViewValue = sys_menu.menu_thai.CurrentValue
			sys_menu.menu_thai.ViewCustomAttributes = ""

			' menu_idname
			sys_menu.menu_idname.ViewValue = sys_menu.menu_idname.CurrentValue
			sys_menu.menu_idname.ViewCustomAttributes = ""

			' menu_filename
			sys_menu.menu_filename.ViewValue = sys_menu.menu_filename.CurrentValue
			sys_menu.menu_filename.ViewCustomAttributes = ""

			' target
			sys_menu.target.ViewValue = sys_menu.target.CurrentValue
			sys_menu.target.ViewCustomAttributes = ""

			' OrderList
			sys_menu.OrderList.ViewValue = sys_menu.OrderList.CurrentValue
			sys_menu.OrderList.ViewCustomAttributes = ""

			' View refer script
			' menu_id

			sys_menu.menu_id.LinkCustomAttributes = ""
			sys_menu.menu_id.HrefValue = ""
			sys_menu.menu_id.TooltipValue = ""

			' menu_name
			sys_menu.menu_name.LinkCustomAttributes = ""
			sys_menu.menu_name.HrefValue = ""
			sys_menu.menu_name.TooltipValue = ""

			' menu_parent_id
			sys_menu.menu_parent_id.LinkCustomAttributes = ""
			sys_menu.menu_parent_id.HrefValue = ""
			sys_menu.menu_parent_id.TooltipValue = ""

			' menu_thai
			sys_menu.menu_thai.LinkCustomAttributes = ""
			sys_menu.menu_thai.HrefValue = ""
			sys_menu.menu_thai.TooltipValue = ""

			' menu_idname
			sys_menu.menu_idname.LinkCustomAttributes = ""
			sys_menu.menu_idname.HrefValue = ""
			sys_menu.menu_idname.TooltipValue = ""

			' menu_filename
			sys_menu.menu_filename.LinkCustomAttributes = ""
			sys_menu.menu_filename.HrefValue = ""
			sys_menu.menu_filename.TooltipValue = ""

			' target
			sys_menu.target.LinkCustomAttributes = ""
			sys_menu.target.HrefValue = ""
			sys_menu.target.TooltipValue = ""

			' OrderList
			sys_menu.OrderList.LinkCustomAttributes = ""
			sys_menu.OrderList.HrefValue = ""
			sys_menu.OrderList.TooltipValue = ""
		End If

		' Call Row Rendered event
		If sys_menu.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call sys_menu.Row_Rendered()
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
		sSql = sys_menu.SQL
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
				DeleteRows = sys_menu.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("menu_id")
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
			ElseIf sys_menu.CancelMessage <> "" Then
				FailureMessage = sys_menu.CancelMessage
				sys_menu.CancelMessage = ""
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
				Call sys_menu.Row_Deleted(RsOld)
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
		Call Breadcrumb.Add("list", sys_menu.TableVar, "pom_sys_menulist.asp", sys_menu.TableVar, True)
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
