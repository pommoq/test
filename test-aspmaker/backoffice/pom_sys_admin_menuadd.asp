<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_sys_admin_menuinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim sys_admin_menu_add
Set sys_admin_menu_add = New csys_admin_menu_add
Set Page = sys_admin_menu_add

' Page init processing
sys_admin_menu_add.Page_Init()

' Page main processing
sys_admin_menu_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
sys_admin_menu_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var sys_admin_menu_add = new ew_Page("sys_admin_menu_add");
sys_admin_menu_add.PageID = "add"; // Page ID
var EW_PAGE_ID = sys_admin_menu_add.PageID; // For backward compatibility
// Form object
var fsys_admin_menuadd = new ew_Form("fsys_admin_menuadd");
// Validate form
fsys_admin_menuadd.Validate = function() {
	if (!this.ValidateRequired)
		return true; // Ignore validation
	var $ = jQuery, fobj = this.GetForm(), $fobj = $(fobj);
	this.PostAutoSuggest();
	if ($fobj.find("#a_confirm").val() == "F")
		return true;
	var elm, felm, uelm, addcnt = 0;
	var $k = $fobj.find("#" + this.FormKeyCountName); // Get key_count
	var rowcnt = ($k[0]) ? parseInt($k.val(), 10) : 1;
	var startcnt = (rowcnt == 0) ? 0 : 1; // Check rowcnt == 0 => Inline-Add
	var gridinsert = $fobj.find("#a_list").val() == "gridinsert";
	for (var i = startcnt; i <= rowcnt; i++) {
		var infix = ($k[0]) ? String(i) : "";
		$fobj.data("rowindex", infix);
			elm = this.GetElements("x" + infix + "_admin_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(sys_admin_menu.admin_id.FldErrMsg) %>");
			elm = this.GetElements("x" + infix + "_menu_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(sys_admin_menu.menu_id.FldErrMsg) %>");
			// Set up row object
			ew_ElementsToRow(fobj);
			// Fire Form_CustomValidate event
			if (!this.Form_CustomValidate(fobj))
				return false;
	}
	// Process detail forms
	var dfs = $fobj.find("input[name='detailpage']").get();
	for (var i = 0; i < dfs.length; i++) {
		var df = dfs[i], val = df.value;
		if (val && ewForms[val])
			if (!ewForms[val].Validate())
				return false;
	}
	return true;
}
// Form_CustomValidate event
fsys_admin_menuadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fsys_admin_menuadd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fsys_admin_menuadd.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If sys_admin_menu.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% sys_admin_menu_add.ShowPageHeader() %>
<% sys_admin_menu_add.ShowMessage %>
<form name="fsys_admin_menuadd" id="fsys_admin_menuadd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="sys_admin_menu">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_sys_admin_menuadd" class="table table-bordered table-striped">
<% If sys_admin_menu.admin_id.Visible Then ' admin_id %>
	<tr id="r_admin_id">
		<td><span id="elh_sys_admin_menu_admin_id"><%= sys_admin_menu.admin_id.FldCaption %></span></td>
		<td<%= sys_admin_menu.admin_id.CellAttributes %>>
<span id="el_sys_admin_menu_admin_id" class="control-group">
<input type="text" data-field="x_admin_id" name="x_admin_id" id="x_admin_id" size="30" placeholder="<%= sys_admin_menu.admin_id.PlaceHolder %>" value="<%= sys_admin_menu.admin_id.EditValue %>"<%= sys_admin_menu.admin_id.EditAttributes %>>
</span>
<%= sys_admin_menu.admin_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If sys_admin_menu.menu_id.Visible Then ' menu_id %>
	<tr id="r_menu_id">
		<td><span id="elh_sys_admin_menu_menu_id"><%= sys_admin_menu.menu_id.FldCaption %></span></td>
		<td<%= sys_admin_menu.menu_id.CellAttributes %>>
<span id="el_sys_admin_menu_menu_id" class="control-group">
<input type="text" data-field="x_menu_id" name="x_menu_id" id="x_menu_id" size="30" placeholder="<%= sys_admin_menu.menu_id.PlaceHolder %>" value="<%= sys_admin_menu.menu_id.EditValue %>"<%= sys_admin_menu.menu_id.EditAttributes %>>
</span>
<%= sys_admin_menu.menu_id.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
fsys_admin_menuadd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
sys_admin_menu_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set sys_admin_menu_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class csys_admin_menu_add

	' Page ID
	Public Property Get PageID()
		PageID = "add"
	End Property

	' Project ID
	Public Property Get ProjectID()
		ProjectID = "{324ED72D-DE20-46F7-B12E-7AF8CE8711A6}"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "sys_admin_menu"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "sys_admin_menu_add"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If sys_admin_menu.UseTokenInUrl Then PageUrl = PageUrl & "t=" & sys_admin_menu.TableVar & "&" ' add page token
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
		If sys_admin_menu.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (sys_admin_menu.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (sys_admin_menu.TableVar = Request.QueryString("t"))
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
		If IsEmpty(sys_admin_menu) Then Set sys_admin_menu = New csys_admin_menu
		Set Table = sys_admin_menu

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "add"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "sys_admin_menu"

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

		' Create form object
		'If Request.ServerVariables("HTTP_CONTENT_TYPE") = "application/x-www-form-urlencoded" Then

			Set ObjForm = New cFormObj

		'Else
		'	Set ObjForm = ew_GetUploadObj()
		'End If

		sys_admin_menu.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set sys_admin_menu = Nothing
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

	Dim DbMasterFilter, DbDetailFilter
	Dim Priv
	Dim OldRecordset
	Dim CopyRecord

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Process form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			sys_admin_menu.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("sys_admin_menu_id").Count > 0 Then
				sys_admin_menu.sys_admin_menu_id.QueryStringValue = Request.QueryString("sys_admin_menu_id")
				Call sys_admin_menu.SetKey("sys_admin_menu_id", sys_admin_menu.sys_admin_menu_id.CurrentValue) ' Set up key
			Else
				Call sys_admin_menu.SetKey("sys_admin_menu_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				sys_admin_menu.CurrentAction = "C" ' Copy Record
			Else
				sys_admin_menu.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				sys_admin_menu.CurrentAction = "I" ' Form error, reset action
				sys_admin_menu.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case sys_admin_menu.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_sys_admin_menulist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				sys_admin_menu.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = sys_admin_menu.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_sys_admin_menuview.asp" Then sReturnUrl = sys_admin_menu.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					sys_admin_menu.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		sys_admin_menu.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call sys_admin_menu.ResetAttrs()
		Call RenderRow()
	End Sub

	' -----------------------------------------------------------------
	' Function Get upload files
	'
	Function GetUploadFiles()

		' Get upload data
	End Function

	' -----------------------------------------------------------------
	' Load default values
	'
	Function LoadDefaultValues()
		sys_admin_menu.admin_id.CurrentValue = Null
		sys_admin_menu.admin_id.OldValue = sys_admin_menu.admin_id.CurrentValue
		sys_admin_menu.menu_id.CurrentValue = Null
		sys_admin_menu.menu_id.OldValue = sys_admin_menu.menu_id.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not sys_admin_menu.admin_id.FldIsDetailKey Then sys_admin_menu.admin_id.FormValue = ObjForm.GetValue("x_admin_id")
		If Not sys_admin_menu.menu_id.FldIsDetailKey Then sys_admin_menu.menu_id.FormValue = ObjForm.GetValue("x_menu_id")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		sys_admin_menu.admin_id.CurrentValue = sys_admin_menu.admin_id.FormValue
		sys_admin_menu.menu_id.CurrentValue = sys_admin_menu.menu_id.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = sys_admin_menu.KeyFilter

		' Call Row Selecting event
		Call sys_admin_menu.Row_Selecting(sFilter)

		' Load sql based on filter
		sys_admin_menu.CurrentFilter = sFilter
		sSql = sys_admin_menu.SQL
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
		Call sys_admin_menu.Row_Selected(RsRow)
		sys_admin_menu.sys_admin_menu_id.DbValue = RsRow("sys_admin_menu_id")
		sys_admin_menu.admin_id.DbValue = RsRow("admin_id")
		sys_admin_menu.menu_id.DbValue = RsRow("menu_id")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		sys_admin_menu.sys_admin_menu_id.m_DbValue = Rs("sys_admin_menu_id")
		sys_admin_menu.admin_id.m_DbValue = Rs("admin_id")
		sys_admin_menu.menu_id.m_DbValue = Rs("menu_id")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If sys_admin_menu.GetKey("sys_admin_menu_id")&"" <> "" Then
			sys_admin_menu.sys_admin_menu_id.CurrentValue = sys_admin_menu.GetKey("sys_admin_menu_id") ' sys_admin_menu_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			sys_admin_menu.CurrentFilter = sys_admin_menu.KeyFilter
			Dim sSql
			sSql = sys_admin_menu.SQL
			Set OldRecordset = ew_LoadRecordset(sSql)
			Call LoadRowValues(OldRecordset) ' Load row values
		Else
			OldRecordset = Null
		End If
		LoadOldRecord = bValidKey
	End Function

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call sys_admin_menu.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' sys_admin_menu_id
		' admin_id
		' menu_id
		' -----------
		'  View  Row
		' -----------

		If sys_admin_menu.RowType = EW_ROWTYPE_VIEW Then ' View row

			' sys_admin_menu_id
			sys_admin_menu.sys_admin_menu_id.ViewValue = sys_admin_menu.sys_admin_menu_id.CurrentValue
			sys_admin_menu.sys_admin_menu_id.ViewCustomAttributes = ""

			' admin_id
			sys_admin_menu.admin_id.ViewValue = sys_admin_menu.admin_id.CurrentValue
			sys_admin_menu.admin_id.ViewCustomAttributes = ""

			' menu_id
			sys_admin_menu.menu_id.ViewValue = sys_admin_menu.menu_id.CurrentValue
			sys_admin_menu.menu_id.ViewCustomAttributes = ""

			' View refer script
			' admin_id

			sys_admin_menu.admin_id.LinkCustomAttributes = ""
			sys_admin_menu.admin_id.HrefValue = ""
			sys_admin_menu.admin_id.TooltipValue = ""

			' menu_id
			sys_admin_menu.menu_id.LinkCustomAttributes = ""
			sys_admin_menu.menu_id.HrefValue = ""
			sys_admin_menu.menu_id.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf sys_admin_menu.RowType = EW_ROWTYPE_ADD Then ' Add row

			' admin_id
			sys_admin_menu.admin_id.EditCustomAttributes = ""
			sys_admin_menu.admin_id.EditValue = ew_HtmlEncode(sys_admin_menu.admin_id.CurrentValue)
			sys_admin_menu.admin_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(sys_admin_menu.admin_id.FldCaption))

			' menu_id
			sys_admin_menu.menu_id.EditCustomAttributes = ""
			sys_admin_menu.menu_id.EditValue = ew_HtmlEncode(sys_admin_menu.menu_id.CurrentValue)
			sys_admin_menu.menu_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(sys_admin_menu.menu_id.FldCaption))

			' Edit refer script
			' admin_id

			sys_admin_menu.admin_id.HrefValue = ""

			' menu_id
			sys_admin_menu.menu_id.HrefValue = ""
		End If
		If sys_admin_menu.RowType = EW_ROWTYPE_ADD Or sys_admin_menu.RowType = EW_ROWTYPE_EDIT Or sys_admin_menu.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call sys_admin_menu.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If sys_admin_menu.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call sys_admin_menu.Row_Rendered()
		End If
	End Sub

	' -----------------------------------------------------------------
	' Validate form
	'
	Function ValidateForm()

		' Initialize
		gsFormError = ""

		' Check if validation required
		If Not EW_SERVER_VALIDATE Then
			ValidateForm = (gsFormError = "")
			Exit Function
		End If
		If Not ew_CheckInteger(sys_admin_menu.admin_id.FormValue) Then
			Call ew_AddMessage(gsFormError, sys_admin_menu.admin_id.FldErrMsg)
		End If
		If Not ew_CheckInteger(sys_admin_menu.menu_id.FormValue) Then
			Call ew_AddMessage(gsFormError, sys_admin_menu.menu_id.FldErrMsg)
		End If

		' Return validate result
		ValidateForm = (gsFormError = "")

		' Call Form Custom Validate event
		Dim sFormCustomError
		sFormCustomError = ""
		ValidateForm = ValidateForm And Form_CustomValidate(sFormCustomError)
		If sFormCustomError <> "" Then
			Call ew_AddMessage(gsFormError, sFormCustomError)
		End If
	End Function

	' -----------------------------------------------------------------
	' Add record
	'
	Function AddRow(RsOld)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim Rs, sSql, sFilter
		Dim RsNew
		Dim bInsertRow
		Dim RsChk
		Dim sIdxErrMsg

		' Clear any previous errors
		Err.Clear
		Dim RsMaster, sMasterUserIdMsg, sMasterFilter, bCheckMasterRecord

		' Load db values from rsold
		If Not IsNull(RsOld) Then
			Call LoadDbValues(RsOld)
		End If

		' Add new record
		sFilter = "(0 = 1)"
		sys_admin_menu.CurrentFilter = sFilter
		sSql = sys_admin_menu.SQL
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = EW_CURSORLOCATION
		Rs.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		Rs.AddNew
		If Err.Number <> 0 Then
			Message = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Field admin_id
		Call sys_admin_menu.admin_id.SetDbValue(Rs, sys_admin_menu.admin_id.CurrentValue, Null, False)

		' Field menu_id
		Call sys_admin_menu.menu_id.SetDbValue(Rs, sys_admin_menu.menu_id.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = sys_admin_menu.Row_Inserting(RsOld, Rs)
		If bInsertRow Then

			' Clone new recordset object
			Set RsNew = ew_CloneRs(Rs)
			Rs.Update
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				AddRow = False
			Else
				AddRow = True
			End If
			If AddRow Then
			End If
		Else
			Rs.CancelUpdate

			' Set up error message
			If SuccessMessage <> "" Or FailureMessage <> "" Then

				' Use the message, do nothing
			ElseIf sys_admin_menu.CancelMessage <> "" Then
				FailureMessage = sys_admin_menu.CancelMessage
				sys_admin_menu.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("InsertCancelled")
			End If
			AddRow = False
		End If
		Rs.Close
		Set Rs = Nothing
		If AddRow Then
			sys_admin_menu.sys_admin_menu_id.DbValue = RsNew("sys_admin_menu_id")
		End If
		If AddRow Then

			' Call Row Inserted event
			Call sys_admin_menu.Row_Inserted(RsOld, RsNew)
		End If
		If IsObject(RsNew) Then
			RsNew.Close
			Set RsNew = Nothing
		End If
	End Function

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", sys_admin_menu.TableVar, "pom_sys_admin_menulist.asp", sys_admin_menu.TableVar, True)
		PageId = ew_IIf(sys_admin_menu.CurrentAction = "C", "Copy", "Add")
		Call Breadcrumb.Add("add", PageId, ew_CurrentUrl, "", False)
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

	' Form Custom Validate event
	Function Form_CustomValidate(CustomError)

		'Return error message in CustomError
		Form_CustomValidate = True
	End Function
End Class
%>
