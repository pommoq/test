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
Dim sys_menu_edit
Set sys_menu_edit = New csys_menu_edit
Set Page = sys_menu_edit

' Page init processing
sys_menu_edit.Page_Init()

' Page main processing
sys_menu_edit.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
sys_menu_edit.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var sys_menu_edit = new ew_Page("sys_menu_edit");
sys_menu_edit.PageID = "edit"; // Page ID
var EW_PAGE_ID = sys_menu_edit.PageID; // For backward compatibility
// Form object
var fsys_menuedit = new ew_Form("fsys_menuedit");
// Validate form
fsys_menuedit.Validate = function() {
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
			elm = this.GetElements("x" + infix + "_menu_parent_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(sys_menu.menu_parent_id.FldErrMsg) %>");
			elm = this.GetElements("x" + infix + "_OrderList");
			if (elm && !ew_CheckNumber(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(sys_menu.OrderList.FldErrMsg) %>");
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
fsys_menuedit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fsys_menuedit.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fsys_menuedit.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If sys_menu.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% sys_menu_edit.ShowPageHeader() %>
<% sys_menu_edit.ShowMessage %>
<form name="fsys_menuedit" id="fsys_menuedit" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="a_table" id="a_table" value="sys_menu">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table class="ewGrid"><tr><td>
<table id="tbl_sys_menuedit" class="table table-bordered table-striped">
<% If sys_menu.menu_id.Visible Then ' menu_id %>
	<tr id="r_menu_id">
		<td><span id="elh_sys_menu_menu_id"><%= sys_menu.menu_id.FldCaption %></span></td>
		<td<%= sys_menu.menu_id.CellAttributes %>>
<span id="el_sys_menu_menu_id" class="control-group">
<span<%= sys_menu.menu_id.ViewAttributes %>>
<%= sys_menu.menu_id.EditValue %>
</span>
</span>
<input type="hidden" data-field="x_menu_id" name="x_menu_id" id="x_menu_id" value="<%= Server.HTMLEncode(sys_menu.menu_id.CurrentValue&"") %>">
<%= sys_menu.menu_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If sys_menu.menu_name.Visible Then ' menu_name %>
	<tr id="r_menu_name">
		<td><span id="elh_sys_menu_menu_name"><%= sys_menu.menu_name.FldCaption %></span></td>
		<td<%= sys_menu.menu_name.CellAttributes %>>
<span id="el_sys_menu_menu_name" class="control-group">
<input type="text" data-field="x_menu_name" name="x_menu_name" id="x_menu_name" size="30" maxlength="255" placeholder="<%= sys_menu.menu_name.PlaceHolder %>" value="<%= sys_menu.menu_name.EditValue %>"<%= sys_menu.menu_name.EditAttributes %>>
</span>
<%= sys_menu.menu_name.CustomMsg %></td>
	</tr>
<% End If %>
<% If sys_menu.menu_parent_id.Visible Then ' menu_parent_id %>
	<tr id="r_menu_parent_id">
		<td><span id="elh_sys_menu_menu_parent_id"><%= sys_menu.menu_parent_id.FldCaption %></span></td>
		<td<%= sys_menu.menu_parent_id.CellAttributes %>>
<span id="el_sys_menu_menu_parent_id" class="control-group">
<input type="text" data-field="x_menu_parent_id" name="x_menu_parent_id" id="x_menu_parent_id" size="30" placeholder="<%= sys_menu.menu_parent_id.PlaceHolder %>" value="<%= sys_menu.menu_parent_id.EditValue %>"<%= sys_menu.menu_parent_id.EditAttributes %>>
</span>
<%= sys_menu.menu_parent_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If sys_menu.menu_thai.Visible Then ' menu_thai %>
	<tr id="r_menu_thai">
		<td><span id="elh_sys_menu_menu_thai"><%= sys_menu.menu_thai.FldCaption %></span></td>
		<td<%= sys_menu.menu_thai.CellAttributes %>>
<span id="el_sys_menu_menu_thai" class="control-group">
<input type="text" data-field="x_menu_thai" name="x_menu_thai" id="x_menu_thai" size="30" maxlength="255" placeholder="<%= sys_menu.menu_thai.PlaceHolder %>" value="<%= sys_menu.menu_thai.EditValue %>"<%= sys_menu.menu_thai.EditAttributes %>>
</span>
<%= sys_menu.menu_thai.CustomMsg %></td>
	</tr>
<% End If %>
<% If sys_menu.menu_idname.Visible Then ' menu_idname %>
	<tr id="r_menu_idname">
		<td><span id="elh_sys_menu_menu_idname"><%= sys_menu.menu_idname.FldCaption %></span></td>
		<td<%= sys_menu.menu_idname.CellAttributes %>>
<span id="el_sys_menu_menu_idname" class="control-group">
<input type="text" data-field="x_menu_idname" name="x_menu_idname" id="x_menu_idname" size="30" maxlength="255" placeholder="<%= sys_menu.menu_idname.PlaceHolder %>" value="<%= sys_menu.menu_idname.EditValue %>"<%= sys_menu.menu_idname.EditAttributes %>>
</span>
<%= sys_menu.menu_idname.CustomMsg %></td>
	</tr>
<% End If %>
<% If sys_menu.menu_filename.Visible Then ' menu_filename %>
	<tr id="r_menu_filename">
		<td><span id="elh_sys_menu_menu_filename"><%= sys_menu.menu_filename.FldCaption %></span></td>
		<td<%= sys_menu.menu_filename.CellAttributes %>>
<span id="el_sys_menu_menu_filename" class="control-group">
<input type="text" data-field="x_menu_filename" name="x_menu_filename" id="x_menu_filename" size="30" maxlength="255" placeholder="<%= sys_menu.menu_filename.PlaceHolder %>" value="<%= sys_menu.menu_filename.EditValue %>"<%= sys_menu.menu_filename.EditAttributes %>>
</span>
<%= sys_menu.menu_filename.CustomMsg %></td>
	</tr>
<% End If %>
<% If sys_menu.target.Visible Then ' target %>
	<tr id="r_target">
		<td><span id="elh_sys_menu_target"><%= sys_menu.target.FldCaption %></span></td>
		<td<%= sys_menu.target.CellAttributes %>>
<span id="el_sys_menu_target" class="control-group">
<input type="text" data-field="x_target" name="x_target" id="x_target" size="30" maxlength="255" placeholder="<%= sys_menu.target.PlaceHolder %>" value="<%= sys_menu.target.EditValue %>"<%= sys_menu.target.EditAttributes %>>
</span>
<%= sys_menu.target.CustomMsg %></td>
	</tr>
<% End If %>
<% If sys_menu.OrderList.Visible Then ' OrderList %>
	<tr id="r_OrderList">
		<td><span id="elh_sys_menu_OrderList"><%= sys_menu.OrderList.FldCaption %></span></td>
		<td<%= sys_menu.OrderList.CellAttributes %>>
<span id="el_sys_menu_OrderList" class="control-group">
<input type="text" data-field="x_OrderList" name="x_OrderList" id="x_OrderList" size="30" placeholder="<%= sys_menu.OrderList.PlaceHolder %>" value="<%= sys_menu.OrderList.EditValue %>"<%= sys_menu.OrderList.EditAttributes %>>
</span>
<%= sys_menu.OrderList.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("EditBtn") %></button>
</form>
<script type="text/javascript">
fsys_menuedit.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
sys_menu_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set sys_menu_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class csys_menu_edit

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
		TableName = "sys_menu"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "sys_menu_edit"
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
		EW_PAGE_ID = "edit"

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

		' Create form object
		'If Request.ServerVariables("HTTP_CONTENT_TYPE") = "application/x-www-form-urlencoded" Then

			Set ObjForm = New cFormObj

		'Else
		'	Set ObjForm = ew_GetUploadObj()
		'End If

		sys_menu.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action
		sys_menu.menu_id.Visible = Not sys_menu.IsAdd() And Not sys_menu.IsCopy() And Not sys_menu.IsGridAdd()

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

	Dim DbMasterFilter, DbDetailFilter

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		Dim sReturnUrl
		sReturnUrl = ""

		' Load key from QueryString
		If Request.QueryString("menu_id").Count > 0 Then
			sys_menu.menu_id.QueryStringValue = Request.QueryString("menu_id")
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Process form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			sys_menu.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values
		Else
			sys_menu.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If sys_menu.menu_id.CurrentValue = "" Then Call Page_Terminate("pom_sys_menulist.asp") ' Invalid key, return to list

		' Validate form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			If Not ValidateForm() Then
				sys_menu.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				sys_menu.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		End If
		Select Case sys_menu.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_sys_menulist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				sys_menu.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					sReturnUrl = sys_menu.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					sys_menu.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		sys_menu.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call sys_menu.ResetAttrs()
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
				sys_menu.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					sys_menu.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = sys_menu.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			sys_menu.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			sys_menu.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			sys_menu.StartRecordNumber = StartRec
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
		If Not sys_menu.menu_id.FldIsDetailKey Then sys_menu.menu_id.FormValue = ObjForm.GetValue("x_menu_id")
		If Not sys_menu.menu_name.FldIsDetailKey Then sys_menu.menu_name.FormValue = ObjForm.GetValue("x_menu_name")
		If Not sys_menu.menu_parent_id.FldIsDetailKey Then sys_menu.menu_parent_id.FormValue = ObjForm.GetValue("x_menu_parent_id")
		If Not sys_menu.menu_thai.FldIsDetailKey Then sys_menu.menu_thai.FormValue = ObjForm.GetValue("x_menu_thai")
		If Not sys_menu.menu_idname.FldIsDetailKey Then sys_menu.menu_idname.FormValue = ObjForm.GetValue("x_menu_idname")
		If Not sys_menu.menu_filename.FldIsDetailKey Then sys_menu.menu_filename.FormValue = ObjForm.GetValue("x_menu_filename")
		If Not sys_menu.target.FldIsDetailKey Then sys_menu.target.FormValue = ObjForm.GetValue("x_target")
		If Not sys_menu.OrderList.FldIsDetailKey Then sys_menu.OrderList.FormValue = ObjForm.GetValue("x_OrderList")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		sys_menu.menu_id.CurrentValue = sys_menu.menu_id.FormValue
		sys_menu.menu_name.CurrentValue = sys_menu.menu_name.FormValue
		sys_menu.menu_parent_id.CurrentValue = sys_menu.menu_parent_id.FormValue
		sys_menu.menu_thai.CurrentValue = sys_menu.menu_thai.FormValue
		sys_menu.menu_idname.CurrentValue = sys_menu.menu_idname.FormValue
		sys_menu.menu_filename.CurrentValue = sys_menu.menu_filename.FormValue
		sys_menu.target.CurrentValue = sys_menu.target.FormValue
		sys_menu.OrderList.CurrentValue = sys_menu.OrderList.FormValue
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

		' ----------
		'  Edit Row
		' ----------

		ElseIf sys_menu.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' menu_id
			sys_menu.menu_id.EditCustomAttributes = ""
			sys_menu.menu_id.EditValue = sys_menu.menu_id.CurrentValue
			sys_menu.menu_id.ViewCustomAttributes = ""

			' menu_name
			sys_menu.menu_name.EditCustomAttributes = ""
			sys_menu.menu_name.EditValue = ew_HtmlEncode(sys_menu.menu_name.CurrentValue)
			sys_menu.menu_name.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(sys_menu.menu_name.FldCaption))

			' menu_parent_id
			sys_menu.menu_parent_id.EditCustomAttributes = ""
			sys_menu.menu_parent_id.EditValue = ew_HtmlEncode(sys_menu.menu_parent_id.CurrentValue)
			sys_menu.menu_parent_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(sys_menu.menu_parent_id.FldCaption))

			' menu_thai
			sys_menu.menu_thai.EditCustomAttributes = ""
			sys_menu.menu_thai.EditValue = ew_HtmlEncode(sys_menu.menu_thai.CurrentValue)
			sys_menu.menu_thai.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(sys_menu.menu_thai.FldCaption))

			' menu_idname
			sys_menu.menu_idname.EditCustomAttributes = ""
			sys_menu.menu_idname.EditValue = ew_HtmlEncode(sys_menu.menu_idname.CurrentValue)
			sys_menu.menu_idname.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(sys_menu.menu_idname.FldCaption))

			' menu_filename
			sys_menu.menu_filename.EditCustomAttributes = ""
			sys_menu.menu_filename.EditValue = ew_HtmlEncode(sys_menu.menu_filename.CurrentValue)
			sys_menu.menu_filename.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(sys_menu.menu_filename.FldCaption))

			' target
			sys_menu.target.EditCustomAttributes = ""
			sys_menu.target.EditValue = ew_HtmlEncode(sys_menu.target.CurrentValue)
			sys_menu.target.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(sys_menu.target.FldCaption))

			' OrderList
			sys_menu.OrderList.EditCustomAttributes = ""
			sys_menu.OrderList.EditValue = ew_HtmlEncode(sys_menu.OrderList.CurrentValue)
			sys_menu.OrderList.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(sys_menu.OrderList.FldCaption))
			If sys_menu.OrderList.EditValue&"" <> "" And IsNumeric(sys_menu.OrderList.EditValue) Then sys_menu.OrderList.EditValue = ew_FormatNumber(sys_menu.OrderList.EditValue, -2, -1, -2, 0)

			' Edit refer script
			' menu_id

			sys_menu.menu_id.HrefValue = ""

			' menu_name
			sys_menu.menu_name.HrefValue = ""

			' menu_parent_id
			sys_menu.menu_parent_id.HrefValue = ""

			' menu_thai
			sys_menu.menu_thai.HrefValue = ""

			' menu_idname
			sys_menu.menu_idname.HrefValue = ""

			' menu_filename
			sys_menu.menu_filename.HrefValue = ""

			' target
			sys_menu.target.HrefValue = ""

			' OrderList
			sys_menu.OrderList.HrefValue = ""
		End If
		If sys_menu.RowType = EW_ROWTYPE_ADD Or sys_menu.RowType = EW_ROWTYPE_EDIT Or sys_menu.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call sys_menu.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If sys_menu.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call sys_menu.Row_Rendered()
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
		If Not ew_CheckInteger(sys_menu.menu_parent_id.FormValue) Then
			Call ew_AddMessage(gsFormError, sys_menu.menu_parent_id.FldErrMsg)
		End If
		If Not ew_CheckNumber(sys_menu.OrderList.FormValue) Then
			Call ew_AddMessage(gsFormError, sys_menu.OrderList.FldErrMsg)
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
		sFilter = sys_menu.KeyFilter
		sys_menu.CurrentFilter  = sFilter
		sSql = sys_menu.SQL
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

			' Field menu_name
			Call sys_menu.menu_name.SetDbValue(Rs, sys_menu.menu_name.CurrentValue, Null, sys_menu.menu_name.ReadOnly)

			' Field menu_parent_id
			Call sys_menu.menu_parent_id.SetDbValue(Rs, sys_menu.menu_parent_id.CurrentValue, Null, sys_menu.menu_parent_id.ReadOnly)

			' Field menu_thai
			Call sys_menu.menu_thai.SetDbValue(Rs, sys_menu.menu_thai.CurrentValue, Null, sys_menu.menu_thai.ReadOnly)

			' Field menu_idname
			Call sys_menu.menu_idname.SetDbValue(Rs, sys_menu.menu_idname.CurrentValue, Null, sys_menu.menu_idname.ReadOnly)

			' Field menu_filename
			Call sys_menu.menu_filename.SetDbValue(Rs, sys_menu.menu_filename.CurrentValue, Null, sys_menu.menu_filename.ReadOnly)

			' Field target
			Call sys_menu.target.SetDbValue(Rs, sys_menu.target.CurrentValue, Null, sys_menu.target.ReadOnly)

			' Field OrderList
			Call sys_menu.OrderList.SetDbValue(Rs, sys_menu.OrderList.CurrentValue, Null, sys_menu.OrderList.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = sys_menu.Row_Updating(RsOld, Rs)
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
				ElseIf sys_menu.CancelMessage <> "" Then
					FailureMessage = sys_menu.CancelMessage
					sys_menu.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call sys_menu.Row_Updated(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", sys_menu.TableVar, "pom_sys_menulist.asp", sys_menu.TableVar, True)
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
