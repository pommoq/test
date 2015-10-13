<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_personinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim person_delete
Set person_delete = New cperson_delete
Set Page = person_delete

' Page init processing
person_delete.Page_Init()

' Page main processing
person_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
person_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var person_delete = new ew_Page("person_delete");
person_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = person_delete.PageID; // For backward compatibility
// Form object
var fpersondelete = new ew_Form("fpersondelete");
// Form_CustomValidate event
fpersondelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fpersondelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fpersondelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set person_delete.Recordset = person_delete.LoadRecordset()
person_delete.TotalRecs = person_delete.Recordset.RecordCount ' Get record count
If person_delete.TotalRecs <= 0 Then ' No record found, exit
	person_delete.Recordset.Close
	Set person_delete.Recordset = Nothing
	Call person_delete.Page_Terminate("pom_personlist.asp") ' Return to list
End If
%>
<% If person.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% person_delete.ShowPageHeader() %>
<% person_delete.ShowMessage %>
<form name="fpersondelete" id="fpersondelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="person">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(person_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(person_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_persondelete" class="ewTable ewTableSeparate">
<%= person.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If person.per_id.Visible Then ' per_id %>
		<td><span id="elh_person_per_id" class="person_per_id"><%= person.per_id.FldCaption %></span></td>
<% End If %>
<% If person.dept_id.Visible Then ' dept_id %>
		<td><span id="elh_person_dept_id" class="person_dept_id"><%= person.dept_id.FldCaption %></span></td>
<% End If %>
<% If person.office_id.Visible Then ' office_id %>
		<td><span id="elh_person_office_id" class="person_office_id"><%= person.office_id.FldCaption %></span></td>
<% End If %>
<% If person.per_img.Visible Then ' per_img %>
		<td><span id="elh_person_per_img" class="person_per_img"><%= person.per_img.FldCaption %></span></td>
<% End If %>
<% If person.per_en_name.Visible Then ' per_en_name %>
		<td><span id="elh_person_per_en_name" class="person_per_en_name"><%= person.per_en_name.FldCaption %></span></td>
<% End If %>
<% If person.per_th_name.Visible Then ' per_th_name %>
		<td><span id="elh_person_per_th_name" class="person_per_th_name"><%= person.per_th_name.FldCaption %></span></td>
<% End If %>
<% If person.per_position.Visible Then ' per_position %>
		<td><span id="elh_person_per_position" class="person_per_position"><%= person.per_position.FldCaption %></span></td>
<% End If %>
<% If person.per_mobile.Visible Then ' per_mobile %>
		<td><span id="elh_person_per_mobile" class="person_per_mobile"><%= person.per_mobile.FldCaption %></span></td>
<% End If %>
<% If person.per_tel.Visible Then ' per_tel %>
		<td><span id="elh_person_per_tel" class="person_per_tel"><%= person.per_tel.FldCaption %></span></td>
<% End If %>
<% If person.per_fax.Visible Then ' per_fax %>
		<td><span id="elh_person_per_fax" class="person_per_fax"><%= person.per_fax.FldCaption %></span></td>
<% End If %>
<% If person.per_email.Visible Then ' per_email %>
		<td><span id="elh_person_per_email" class="person_per_email"><%= person.per_email.FldCaption %></span></td>
<% End If %>
<% If person.per_address.Visible Then ' per_address %>
		<td><span id="elh_person_per_address" class="person_per_address"><%= person.per_address.FldCaption %></span></td>
<% End If %>
<% If person.per_show.Visible Then ' per_show %>
		<td><span id="elh_person_per_show" class="person_per_show"><%= person.per_show.FldCaption %></span></td>
<% End If %>
<% If person.per_create.Visible Then ' per_create %>
		<td><span id="elh_person_per_create" class="person_per_create"><%= person.per_create.FldCaption %></span></td>
<% End If %>
<% If person.per_update.Visible Then ' per_update %>
		<td><span id="elh_person_per_update" class="person_per_update"><%= person.per_update.FldCaption %></span></td>
<% End If %>
<% If person.per_sort.Visible Then ' per_sort %>
		<td><span id="elh_person_per_sort" class="person_per_sort"><%= person.per_sort.FldCaption %></span></td>
<% End If %>
<% If person.per_department.Visible Then ' per_department %>
		<td><span id="elh_person_per_department" class="person_per_department"><%= person.per_department.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
person_delete.RecCnt = 0
person_delete.RowCnt = 0
Do While (Not person_delete.Recordset.Eof)
	person_delete.RecCnt = person_delete.RecCnt + 1
	person_delete.RowCnt = person_delete.RowCnt + 1

	' Set row properties
	Call person.ResetAttrs()
	person.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call person_delete.LoadRowValues(person_delete.Recordset)

	' Render row
	Call person_delete.RenderRow()
%>
	<tr<%= person.RowAttributes %>>
<% If person.per_id.Visible Then ' per_id %>
		<td<%= person.per_id.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_per_id" class="control-group person_per_id">
<span<%= person.per_id.ViewAttributes %>>
<%= person.per_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If person.dept_id.Visible Then ' dept_id %>
		<td<%= person.dept_id.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_dept_id" class="control-group person_dept_id">
<span<%= person.dept_id.ViewAttributes %>>
<%= person.dept_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If person.office_id.Visible Then ' office_id %>
		<td<%= person.office_id.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_office_id" class="control-group person_office_id">
<span<%= person.office_id.ViewAttributes %>>
<%= person.office_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If person.per_img.Visible Then ' per_img %>
		<td<%= person.per_img.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_per_img" class="control-group person_per_img">
<span<%= person.per_img.ViewAttributes %>>
<%= person.per_img.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If person.per_en_name.Visible Then ' per_en_name %>
		<td<%= person.per_en_name.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_per_en_name" class="control-group person_per_en_name">
<span<%= person.per_en_name.ViewAttributes %>>
<%= person.per_en_name.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If person.per_th_name.Visible Then ' per_th_name %>
		<td<%= person.per_th_name.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_per_th_name" class="control-group person_per_th_name">
<span<%= person.per_th_name.ViewAttributes %>>
<%= person.per_th_name.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If person.per_position.Visible Then ' per_position %>
		<td<%= person.per_position.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_per_position" class="control-group person_per_position">
<span<%= person.per_position.ViewAttributes %>>
<%= person.per_position.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If person.per_mobile.Visible Then ' per_mobile %>
		<td<%= person.per_mobile.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_per_mobile" class="control-group person_per_mobile">
<span<%= person.per_mobile.ViewAttributes %>>
<%= person.per_mobile.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If person.per_tel.Visible Then ' per_tel %>
		<td<%= person.per_tel.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_per_tel" class="control-group person_per_tel">
<span<%= person.per_tel.ViewAttributes %>>
<%= person.per_tel.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If person.per_fax.Visible Then ' per_fax %>
		<td<%= person.per_fax.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_per_fax" class="control-group person_per_fax">
<span<%= person.per_fax.ViewAttributes %>>
<%= person.per_fax.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If person.per_email.Visible Then ' per_email %>
		<td<%= person.per_email.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_per_email" class="control-group person_per_email">
<span<%= person.per_email.ViewAttributes %>>
<%= person.per_email.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If person.per_address.Visible Then ' per_address %>
		<td<%= person.per_address.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_per_address" class="control-group person_per_address">
<span<%= person.per_address.ViewAttributes %>>
<%= person.per_address.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If person.per_show.Visible Then ' per_show %>
		<td<%= person.per_show.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_per_show" class="control-group person_per_show">
<span<%= person.per_show.ViewAttributes %>>
<%= person.per_show.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If person.per_create.Visible Then ' per_create %>
		<td<%= person.per_create.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_per_create" class="control-group person_per_create">
<span<%= person.per_create.ViewAttributes %>>
<%= person.per_create.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If person.per_update.Visible Then ' per_update %>
		<td<%= person.per_update.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_per_update" class="control-group person_per_update">
<span<%= person.per_update.ViewAttributes %>>
<%= person.per_update.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If person.per_sort.Visible Then ' per_sort %>
		<td<%= person.per_sort.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_per_sort" class="control-group person_per_sort">
<span<%= person.per_sort.ViewAttributes %>>
<%= person.per_sort.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If person.per_department.Visible Then ' per_department %>
		<td<%= person.per_department.CellAttributes %>>
<span id="el<%= person_delete.RowCnt %>_person_per_department" class="control-group person_per_department">
<span<%= person.per_department.ViewAttributes %>>
<%= person.per_department.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	person_delete.Recordset.MoveNext
Loop
person_delete.Recordset.Close
Set person_delete.Recordset = Nothing
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
fpersondelete.Init();
</script>
<%
person_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set person_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cperson_delete

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
		TableName = "person"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "person_delete"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If person.UseTokenInUrl Then PageUrl = PageUrl & "t=" & person.TableVar & "&" ' add page token
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
		If person.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (person.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (person.TableVar = Request.QueryString("t"))
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
		If IsEmpty(person) Then Set person = New cperson
		Set Table = person

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "delete"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "person"

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
		Set person = Nothing
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
		RecKeys = person.GetRecordKeys() ' Load record keys
		sFilter = person.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_personlist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in person class, personinfo.asp

		person.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			person.CurrentAction = Request.Form("a_delete")
		Else
			person.CurrentAction = "I"	' Display record
		End If
		Select Case person.CurrentAction
			Case "D" ' Delete
				person.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(person.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = person.CurrentFilter
		Call person.Recordset_Selecting(sFilter)
		person.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = person.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call person.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = person.KeyFilter

		' Call Row Selecting event
		Call person.Row_Selecting(sFilter)

		' Load sql based on filter
		person.CurrentFilter = sFilter
		sSql = person.SQL
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
		Call person.Row_Selected(RsRow)
		person.per_id.DbValue = RsRow("per_id")
		person.dept_id.DbValue = RsRow("dept_id")
		person.office_id.DbValue = RsRow("office_id")
		person.per_img.DbValue = RsRow("per_img")
		person.per_en_name.DbValue = RsRow("per_en_name")
		person.per_th_name.DbValue = RsRow("per_th_name")
		person.per_position.DbValue = RsRow("per_position")
		person.per_mobile.DbValue = RsRow("per_mobile")
		person.per_tel.DbValue = RsRow("per_tel")
		person.per_fax.DbValue = RsRow("per_fax")
		person.per_email.DbValue = RsRow("per_email")
		person.per_address.DbValue = RsRow("per_address")
		person.per_show.DbValue = RsRow("per_show")
		person.per_create.DbValue = RsRow("per_create")
		person.per_update.DbValue = RsRow("per_update")
		person.per_sort.DbValue = RsRow("per_sort")
		person.per_department.DbValue = RsRow("per_department")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		person.per_id.m_DbValue = Rs("per_id")
		person.dept_id.m_DbValue = Rs("dept_id")
		person.office_id.m_DbValue = Rs("office_id")
		person.per_img.m_DbValue = Rs("per_img")
		person.per_en_name.m_DbValue = Rs("per_en_name")
		person.per_th_name.m_DbValue = Rs("per_th_name")
		person.per_position.m_DbValue = Rs("per_position")
		person.per_mobile.m_DbValue = Rs("per_mobile")
		person.per_tel.m_DbValue = Rs("per_tel")
		person.per_fax.m_DbValue = Rs("per_fax")
		person.per_email.m_DbValue = Rs("per_email")
		person.per_address.m_DbValue = Rs("per_address")
		person.per_show.m_DbValue = Rs("per_show")
		person.per_create.m_DbValue = Rs("per_create")
		person.per_update.m_DbValue = Rs("per_update")
		person.per_sort.m_DbValue = Rs("per_sort")
		person.per_department.m_DbValue = Rs("per_department")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call person.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' per_id
		' dept_id
		' office_id
		' per_img
		' per_en_name
		' per_th_name
		' per_position
		' per_mobile
		' per_tel
		' per_fax
		' per_email
		' per_address
		' per_show
		' per_create
		' per_update
		' per_sort
		' per_department
		' -----------
		'  View  Row
		' -----------

		If person.RowType = EW_ROWTYPE_VIEW Then ' View row

			' per_id
			person.per_id.ViewValue = person.per_id.CurrentValue
			person.per_id.ViewCustomAttributes = ""

			' dept_id
			person.dept_id.ViewValue = person.dept_id.CurrentValue
			person.dept_id.ViewCustomAttributes = ""

			' office_id
			person.office_id.ViewValue = person.office_id.CurrentValue
			person.office_id.ViewCustomAttributes = ""

			' per_img
			person.per_img.ViewValue = person.per_img.CurrentValue
			person.per_img.ViewCustomAttributes = ""

			' per_en_name
			person.per_en_name.ViewValue = person.per_en_name.CurrentValue
			person.per_en_name.ViewCustomAttributes = ""

			' per_th_name
			person.per_th_name.ViewValue = person.per_th_name.CurrentValue
			person.per_th_name.ViewCustomAttributes = ""

			' per_position
			person.per_position.ViewValue = person.per_position.CurrentValue
			person.per_position.ViewCustomAttributes = ""

			' per_mobile
			person.per_mobile.ViewValue = person.per_mobile.CurrentValue
			person.per_mobile.ViewCustomAttributes = ""

			' per_tel
			person.per_tel.ViewValue = person.per_tel.CurrentValue
			person.per_tel.ViewCustomAttributes = ""

			' per_fax
			person.per_fax.ViewValue = person.per_fax.CurrentValue
			person.per_fax.ViewCustomAttributes = ""

			' per_email
			person.per_email.ViewValue = person.per_email.CurrentValue
			person.per_email.ViewCustomAttributes = ""

			' per_address
			person.per_address.ViewValue = person.per_address.CurrentValue
			person.per_address.ViewCustomAttributes = ""

			' per_show
			person.per_show.ViewValue = person.per_show.CurrentValue
			person.per_show.ViewCustomAttributes = ""

			' per_create
			person.per_create.ViewValue = person.per_create.CurrentValue
			person.per_create.ViewCustomAttributes = ""

			' per_update
			person.per_update.ViewValue = person.per_update.CurrentValue
			person.per_update.ViewCustomAttributes = ""

			' per_sort
			person.per_sort.ViewValue = person.per_sort.CurrentValue
			person.per_sort.ViewCustomAttributes = ""

			' per_department
			person.per_department.ViewValue = person.per_department.CurrentValue
			person.per_department.ViewCustomAttributes = ""

			' View refer script
			' per_id

			person.per_id.LinkCustomAttributes = ""
			person.per_id.HrefValue = ""
			person.per_id.TooltipValue = ""

			' dept_id
			person.dept_id.LinkCustomAttributes = ""
			person.dept_id.HrefValue = ""
			person.dept_id.TooltipValue = ""

			' office_id
			person.office_id.LinkCustomAttributes = ""
			person.office_id.HrefValue = ""
			person.office_id.TooltipValue = ""

			' per_img
			person.per_img.LinkCustomAttributes = ""
			person.per_img.HrefValue = ""
			person.per_img.TooltipValue = ""

			' per_en_name
			person.per_en_name.LinkCustomAttributes = ""
			person.per_en_name.HrefValue = ""
			person.per_en_name.TooltipValue = ""

			' per_th_name
			person.per_th_name.LinkCustomAttributes = ""
			person.per_th_name.HrefValue = ""
			person.per_th_name.TooltipValue = ""

			' per_position
			person.per_position.LinkCustomAttributes = ""
			person.per_position.HrefValue = ""
			person.per_position.TooltipValue = ""

			' per_mobile
			person.per_mobile.LinkCustomAttributes = ""
			person.per_mobile.HrefValue = ""
			person.per_mobile.TooltipValue = ""

			' per_tel
			person.per_tel.LinkCustomAttributes = ""
			person.per_tel.HrefValue = ""
			person.per_tel.TooltipValue = ""

			' per_fax
			person.per_fax.LinkCustomAttributes = ""
			person.per_fax.HrefValue = ""
			person.per_fax.TooltipValue = ""

			' per_email
			person.per_email.LinkCustomAttributes = ""
			person.per_email.HrefValue = ""
			person.per_email.TooltipValue = ""

			' per_address
			person.per_address.LinkCustomAttributes = ""
			person.per_address.HrefValue = ""
			person.per_address.TooltipValue = ""

			' per_show
			person.per_show.LinkCustomAttributes = ""
			person.per_show.HrefValue = ""
			person.per_show.TooltipValue = ""

			' per_create
			person.per_create.LinkCustomAttributes = ""
			person.per_create.HrefValue = ""
			person.per_create.TooltipValue = ""

			' per_update
			person.per_update.LinkCustomAttributes = ""
			person.per_update.HrefValue = ""
			person.per_update.TooltipValue = ""

			' per_sort
			person.per_sort.LinkCustomAttributes = ""
			person.per_sort.HrefValue = ""
			person.per_sort.TooltipValue = ""

			' per_department
			person.per_department.LinkCustomAttributes = ""
			person.per_department.HrefValue = ""
			person.per_department.TooltipValue = ""
		End If

		' Call Row Rendered event
		If person.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call person.Row_Rendered()
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
		sSql = person.SQL
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
				DeleteRows = person.Row_Deleting(RsDelete)
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
				sThisKey = sThisKey & RsDelete("per_id")
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
			ElseIf person.CancelMessage <> "" Then
				FailureMessage = person.CancelMessage
				person.CancelMessage = ""
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
				Call person.Row_Deleted(RsOld)
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
		Call Breadcrumb.Add("list", person.TableVar, "pom_personlist.asp", person.TableVar, True)
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
