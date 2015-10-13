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
Dim sys_admin_menu_edit
Set sys_admin_menu_edit = New csys_admin_menu_edit
Set Page = sys_admin_menu_edit

' Page init processing
sys_admin_menu_edit.Page_Init()

' Page main processing
sys_admin_menu_edit.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
sys_admin_menu_edit.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var sys_admin_menu_edit = new ew_Page("sys_admin_menu_edit");
sys_admin_menu_edit.PageID = "edit"; // Page ID
var EW_PAGE_ID = sys_admin_menu_edit.PageID; // For backward compatibility
// Form object
var fsys_admin_menuedit = new ew_Form("fsys_admin_menuedit");
// Validate form
fsys_admin_menuedit.Validate = function() {
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
fsys_admin_menuedit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fsys_admin_menuedit.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fsys_admin_menuedit.ValidateRequired = false; // No JavaScript validation
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
<% sys_admin_menu_edit.ShowPageHeader() %>
<% sys_admin_menu_edit.ShowMessage %>
<form name="fsys_admin_menuedit" id="fsys_admin_menuedit" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="a_table" id="a_table" value="sys_admin_menu">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table class="ewGrid"><tr><td>
<table id="tbl_sys_admin_menuedit" class="table table-bordered table-striped">
<% If sys_admin_menu.sys_admin_menu_id.Visible Then ' sys_admin_menu_id %>
	<tr id="r_sys_admin_menu_id">
		<td><span id="elh_sys_admin_menu_sys_admin_menu_id"><%= sys_admin_menu.sys_admin_menu_id.FldCaption %></span></td>
		<td<%= sys_admin_menu.sys_admin_menu_id.CellAttributes %>>
<span id="el_sys_admin_menu_sys_admin_menu_id" class="control-group">
<span<%= sys_admin_menu.sys_admin_menu_id.ViewAttributes %>>
<%= sys_admin_menu.sys_admin_menu_id.EditValue %>
</span>
</span>
<input type="hidden" data-field="x_sys_admin_menu_id" name="x_sys_admin_menu_id" id="x_sys_admin_menu_id" value="<%= Server.HTMLEncode(sys_admin_menu.sys_admin_menu_id.CurrentValue&"") %>">
<%= sys_admin_menu.sys_admin_menu_id.CustomMsg %></td>
	</tr>
<% End If %>
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
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("EditBtn") %></button>
</form>
<script type="text/javascript">
fsys_admin_menuedit.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
sys_admin_menu_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set sys_admin_menu_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class csys_admin_menu_edit

	' Page ID
	Public Property Get PageID()
		PageID = "edit"
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
		PageObjName = "sys_admin_menu_edit"
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
		EW_PAGE_ID = "edit"

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
		sys_admin_menu.sys_admin_menu_id.Visible = Not sys_admin_menu.IsAdd() And Not sys_admin_menu.IsCopy() And Not sys_admin_menu.IsGridAdd()

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

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		Dim sReturnUrl
		sReturnUrl = ""

		' Load key from QueryString
		If Request.QueryString("sys_admin_menu_id").Count > 0 Then
			sys_admin_menu.sys_admin_menu_id.QueryStringValue = Request.QueryString("sys_admin_menu_id")
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Process form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			sys_admin_menu.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values
		Else
			sys_admin_menu.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If sys_admin_menu.sys_admin_menu_id.CurrentValue = "" Then Call Page_Terminate("pom_sys_admin_menulist.asp") ' Invalid key, return to list

		' Validate form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			If Not ValidateForm() Then
				sys_admin_menu.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				sys_admin_menu.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		End If
		Select Case sys_admin_menu.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_sys_admin_menulist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				sys_admin_menu.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					sReturnUrl = sys_admin_menu.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					sys_admin_menu.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		sys_admin_menu.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call sys_admin_menu.ResetAttrs()
		Call RenderRow()
	End Sub
	Dim Pager

	' -----------------------------------------------------------------
	' Set up Starting Record parameters based on Pager Navigation
	'
	Sub SetUpStartRec()
		Dim PageNo

		' Exit if DisplayRecs = 0
		If DisplayRecs = 0 Then Exit Sub
		If IsPageRequest Then ' Validate request

			' Check for a START parameter
			If Request.QueryString(EW_TABLE_START_REC).Count > 0 Then
				StartRec = Request.QueryString(EW_TABLE_START_REC)
				sys_admin_menu.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					sys_admin_menu.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = sys_admin_menu.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			sys_admin_menu.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			sys_admin_menu.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			sys_admin_menu.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Function Get upload files
	'
	Function GetUploadFiles()

		' Get upload data
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not sys_admin_menu.sys_admin_menu_id.FldIsDetailKey Then sys_admin_menu.sys_admin_menu_id.FormValue = ObjForm.GetValue("x_sys_admin_menu_id")
		If Not sys_admin_menu.admin_id.FldIsDetailKey Then sys_admin_menu.admin_id.FormValue = ObjForm.GetValue("x_admin_id")
		If Not sys_admin_menu.menu_id.FldIsDetailKey Then sys_admin_menu.menu_id.FormValue = ObjForm.GetValue("x_menu_id")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		sys_admin_menu.sys_admin_menu_id.CurrentValue = sys_admin_menu.sys_admin_menu_id.FormValue
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
			' sys_admin_menu_id

			sys_admin_menu.sys_admin_menu_id.LinkCustomAttributes = ""
			sys_admin_menu.sys_admin_menu_id.HrefValue = ""
			sys_admin_menu.sys_admin_menu_id.TooltipValue = ""

			' admin_id
			sys_admin_menu.admin_id.LinkCustomAttributes = ""
			sys_admin_menu.admin_id.HrefValue = ""
			sys_admin_menu.admin_id.TooltipValue = ""

			' menu_id
			sys_admin_menu.menu_id.LinkCustomAttributes = ""
			sys_admin_menu.menu_id.HrefValue = ""
			sys_admin_menu.menu_id.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf sys_admin_menu.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' sys_admin_menu_id
			sys_admin_menu.sys_admin_menu_id.EditCustomAttributes = ""
			sys_admin_menu.sys_admin_menu_id.EditValue = sys_admin_menu.sys_admin_menu_id.CurrentValue
			sys_admin_menu.sys_admin_menu_id.ViewCustomAttributes = ""

			' admin_id
			sys_admin_menu.admin_id.EditCustomAttributes = ""
			sys_admin_menu.admin_id.EditValue = ew_HtmlEncode(sys_admin_menu.admin_id.CurrentValue)
			sys_admin_menu.admin_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(sys_admin_menu.admin_id.FldCaption))

			' menu_id
			sys_admin_menu.menu_id.EditCustomAttributes = ""
			sys_admin_menu.menu_id.EditValue = ew_HtmlEncode(sys_admin_menu.menu_id.CurrentValue)
			sys_admin_menu.menu_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(sys_admin_menu.menu_id.FldCaption))

			' Edit refer script
			' sys_admin_menu_id

			sys_admin_menu.sys_admin_menu_id.HrefValue = ""

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
	' Update record based on key values
	'
	Function EditRow()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim Rs, sSql, sFilter
		Dim RsChk, sSqlChk, sFilterChk
		Dim bUpdateRow
		Dim RsOld, RsNew
		Dim sIdxErrMsg

		' Clear any previous errors
		Err.Clear
		sFilter = sys_admin_menu.KeyFilter
		sys_admin_menu.CurrentFilter  = sFilter
		sSql = sys_admin_menu.SQL
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = EW_CURSORLOCATION
		Rs.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		If Err.Number <> 0 Then
			Message = Err.Description
			Rs.Close
			Set Rs = Nothing
			EditRow = False
			Exit Function
		End If

		' Clone old recordset object
		Set RsOld = ew_CloneRs(Rs)
		Call LoadDbValues(RsOld)
		If Rs.Eof Then
			EditRow = False ' Update Failed
		Else

			' Field admin_id
			Call sys_admin_menu.admin_id.SetDbValue(Rs, sys_admin_menu.admin_id.CurrentValue, Null, sys_admin_menu.admin_id.ReadOnly)

			' Field menu_id
			Call sys_admin_menu.menu_id.SetDbValue(Rs, sys_admin_menu.menu_id.CurrentValue, Null, sys_admin_menu.menu_id.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = sys_admin_menu.Row_Updating(RsOld, Rs)
			If bUpdateRow Then

				' Clone new recordset object
				Set RsNew = ew_CloneRs(Rs)
				EditRow = True
				If EditRow Then
					Rs.Update
				End If
				If Err.Number <> 0 Or Not EditRow Then
					If Err.Description <> "" Then FailureMessage = Err.Description
					EditRow = False
				Else
					EditRow = True
				End If
				If EditRow Then
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
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call sys_admin_menu.Row_Updated(RsOld, RsNew)
		End If
		Rs.Close
		Set Rs = Nothing
		If IsObject(RsOld) Then
			RsOld.Close
			Set RsOld = Nothing
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
		PageId = "edit"
		Call Breadcrumb.Add("edit", PageId, ew_CurrentUrl, "", False)
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
