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
Dim admins_add
Set admins_add = New cadmins_add
Set Page = admins_add

' Page init processing
admins_add.Page_Init()

' Page main processing
admins_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
admins_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var admins_add = new ew_Page("admins_add");
admins_add.PageID = "add"; // Page ID
var EW_PAGE_ID = admins_add.PageID; // For backward compatibility
// Form object
var fadminsadd = new ew_Form("fadminsadd");
// Validate form
fadminsadd.Validate = function() {
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
fadminsadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fadminsadd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fadminsadd.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If admins.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% admins_add.ShowPageHeader() %>
<% admins_add.ShowMessage %>
<form name="fadminsadd" id="fadminsadd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="admins">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_adminsadd" class="table table-bordered table-striped">
<% If admins.admin_username.Visible Then ' admin_username %>
	<tr id="r_admin_username">
		<td><span id="elh_admins_admin_username"><%= admins.admin_username.FldCaption %></span></td>
		<td<%= admins.admin_username.CellAttributes %>>
<span id="el_admins_admin_username" class="control-group">
<input type="text" data-field="x_admin_username" name="x_admin_username" id="x_admin_username" size="30" maxlength="15" placeholder="<%= admins.admin_username.PlaceHolder %>" value="<%= admins.admin_username.EditValue %>"<%= admins.admin_username.EditAttributes %>>
</span>
<%= admins.admin_username.CustomMsg %></td>
	</tr>
<% End If %>
<% If admins.admin_password.Visible Then ' admin_password %>
	<tr id="r_admin_password">
		<td><span id="elh_admins_admin_password"><%= admins.admin_password.FldCaption %></span></td>
		<td<%= admins.admin_password.CellAttributes %>>
<span id="el_admins_admin_password" class="control-group">
<input type="text" data-field="x_admin_password" name="x_admin_password" id="x_admin_password" size="30" maxlength="60" placeholder="<%= admins.admin_password.PlaceHolder %>" value="<%= admins.admin_password.EditValue %>"<%= admins.admin_password.EditAttributes %>>
</span>
<%= admins.admin_password.CustomMsg %></td>
	</tr>
<% End If %>
<% If admins.admin_name.Visible Then ' admin_name %>
	<tr id="r_admin_name">
		<td><span id="elh_admins_admin_name"><%= admins.admin_name.FldCaption %></span></td>
		<td<%= admins.admin_name.CellAttributes %>>
<span id="el_admins_admin_name" class="control-group">
<input type="text" data-field="x_admin_name" name="x_admin_name" id="x_admin_name" size="30" maxlength="50" placeholder="<%= admins.admin_name.PlaceHolder %>" value="<%= admins.admin_name.EditValue %>"<%= admins.admin_name.EditAttributes %>>
</span>
<%= admins.admin_name.CustomMsg %></td>
	</tr>
<% End If %>
<% If admins.admin_email.Visible Then ' admin_email %>
	<tr id="r_admin_email">
		<td><span id="elh_admins_admin_email"><%= admins.admin_email.FldCaption %></span></td>
		<td<%= admins.admin_email.CellAttributes %>>
<span id="el_admins_admin_email" class="control-group">
<input type="text" data-field="x_admin_email" name="x_admin_email" id="x_admin_email" size="30" maxlength="50" placeholder="<%= admins.admin_email.PlaceHolder %>" value="<%= admins.admin_email.EditValue %>"<%= admins.admin_email.EditAttributes %>>
</span>
<%= admins.admin_email.CustomMsg %></td>
	</tr>
<% End If %>
<% If admins.admin_tel.Visible Then ' admin_tel %>
	<tr id="r_admin_tel">
		<td><span id="elh_admins_admin_tel"><%= admins.admin_tel.FldCaption %></span></td>
		<td<%= admins.admin_tel.CellAttributes %>>
<span id="el_admins_admin_tel" class="control-group">
<input type="text" data-field="x_admin_tel" name="x_admin_tel" id="x_admin_tel" size="30" maxlength="20" placeholder="<%= admins.admin_tel.PlaceHolder %>" value="<%= admins.admin_tel.EditValue %>"<%= admins.admin_tel.EditAttributes %>>
</span>
<%= admins.admin_tel.CustomMsg %></td>
	</tr>
<% End If %>
<% If admins.admin_permis.Visible Then ' admin_permis %>
	<tr id="r_admin_permis">
		<td><span id="elh_admins_admin_permis"><%= admins.admin_permis.FldCaption %></span></td>
		<td<%= admins.admin_permis.CellAttributes %>>
<span id="el_admins_admin_permis" class="control-group">
<input type="text" data-field="x_admin_permis" name="x_admin_permis" id="x_admin_permis" size="30" maxlength="30" placeholder="<%= admins.admin_permis.PlaceHolder %>" value="<%= admins.admin_permis.EditValue %>"<%= admins.admin_permis.EditAttributes %>>
</span>
<%= admins.admin_permis.CustomMsg %></td>
	</tr>
<% End If %>
<% If admins.admin_create.Visible Then ' admin_create %>
	<tr id="r_admin_create">
		<td><span id="elh_admins_admin_create"><%= admins.admin_create.FldCaption %></span></td>
		<td<%= admins.admin_create.CellAttributes %>>
<span id="el_admins_admin_create" class="control-group">
<input type="text" data-field="x_admin_create" name="x_admin_create" id="x_admin_create" placeholder="<%= admins.admin_create.PlaceHolder %>" value="<%= admins.admin_create.EditValue %>"<%= admins.admin_create.EditAttributes %>>
</span>
<%= admins.admin_create.CustomMsg %></td>
	</tr>
<% End If %>
<% If admins.admin_update.Visible Then ' admin_update %>
	<tr id="r_admin_update">
		<td><span id="elh_admins_admin_update"><%= admins.admin_update.FldCaption %></span></td>
		<td<%= admins.admin_update.CellAttributes %>>
<span id="el_admins_admin_update" class="control-group">
<input type="text" data-field="x_admin_update" name="x_admin_update" id="x_admin_update" placeholder="<%= admins.admin_update.PlaceHolder %>" value="<%= admins.admin_update.EditValue %>"<%= admins.admin_update.EditAttributes %>>
</span>
<%= admins.admin_update.CustomMsg %></td>
	</tr>
<% End If %>
<% If admins.last_online.Visible Then ' last_online %>
	<tr id="r_last_online">
		<td><span id="elh_admins_last_online"><%= admins.last_online.FldCaption %></span></td>
		<td<%= admins.last_online.CellAttributes %>>
<span id="el_admins_last_online" class="control-group">
<input type="text" data-field="x_last_online" name="x_last_online" id="x_last_online" placeholder="<%= admins.last_online.PlaceHolder %>" value="<%= admins.last_online.EditValue %>"<%= admins.last_online.EditAttributes %>>
</span>
<%= admins.last_online.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
fadminsadd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
admins_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set admins_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cadmins_add

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
		TableName = "admins"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "admins_add"
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
		EW_PAGE_ID = "add"

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

		' Create form object
		'If Request.ServerVariables("HTTP_CONTENT_TYPE") = "application/x-www-form-urlencoded" Then

			Set ObjForm = New cFormObj

		'Else
		'	Set ObjForm = ew_GetUploadObj()
		'End If

		admins.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
			admins.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("admin_id").Count > 0 Then
				admins.admin_id.QueryStringValue = Request.QueryString("admin_id")
				Call admins.SetKey("admin_id", admins.admin_id.CurrentValue) ' Set up key
			Else
				Call admins.SetKey("admin_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				admins.CurrentAction = "C" ' Copy Record
			Else
				admins.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				admins.CurrentAction = "I" ' Form error, reset action
				admins.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case admins.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_adminslist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				admins.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = admins.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_adminsview.asp" Then sReturnUrl = admins.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					admins.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		admins.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call admins.ResetAttrs()
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
		admins.admin_username.CurrentValue = Null
		admins.admin_username.OldValue = admins.admin_username.CurrentValue
		admins.admin_password.CurrentValue = Null
		admins.admin_password.OldValue = admins.admin_password.CurrentValue
		admins.admin_name.CurrentValue = Null
		admins.admin_name.OldValue = admins.admin_name.CurrentValue
		admins.admin_email.CurrentValue = Null
		admins.admin_email.OldValue = admins.admin_email.CurrentValue
		admins.admin_tel.CurrentValue = Null
		admins.admin_tel.OldValue = admins.admin_tel.CurrentValue
		admins.admin_permis.CurrentValue = "All"
		admins.admin_create.CurrentValue = Null
		admins.admin_create.OldValue = admins.admin_create.CurrentValue
		admins.admin_update.CurrentValue = Null
		admins.admin_update.OldValue = admins.admin_update.CurrentValue
		admins.last_online.CurrentValue = Null
		admins.last_online.OldValue = admins.last_online.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not admins.admin_username.FldIsDetailKey Then admins.admin_username.FormValue = ObjForm.GetValue("x_admin_username")
		If Not admins.admin_password.FldIsDetailKey Then admins.admin_password.FormValue = ObjForm.GetValue("x_admin_password")
		If Not admins.admin_name.FldIsDetailKey Then admins.admin_name.FormValue = ObjForm.GetValue("x_admin_name")
		If Not admins.admin_email.FldIsDetailKey Then admins.admin_email.FormValue = ObjForm.GetValue("x_admin_email")
		If Not admins.admin_tel.FldIsDetailKey Then admins.admin_tel.FormValue = ObjForm.GetValue("x_admin_tel")
		If Not admins.admin_permis.FldIsDetailKey Then admins.admin_permis.FormValue = ObjForm.GetValue("x_admin_permis")
		If Not admins.admin_create.FldIsDetailKey Then admins.admin_create.FormValue = ObjForm.GetValue("x_admin_create")
		If Not admins.admin_create.FldIsDetailKey Then admins.admin_create.CurrentValue = ew_UnFormatDateTime(admins.admin_create.CurrentValue, 8)
		If Not admins.admin_update.FldIsDetailKey Then admins.admin_update.FormValue = ObjForm.GetValue("x_admin_update")
		If Not admins.admin_update.FldIsDetailKey Then admins.admin_update.CurrentValue = ew_UnFormatDateTime(admins.admin_update.CurrentValue, 8)
		If Not admins.last_online.FldIsDetailKey Then admins.last_online.FormValue = ObjForm.GetValue("x_last_online")
		If Not admins.last_online.FldIsDetailKey Then admins.last_online.CurrentValue = ew_UnFormatDateTime(admins.last_online.CurrentValue, 8)
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		admins.admin_username.CurrentValue = admins.admin_username.FormValue
		admins.admin_password.CurrentValue = admins.admin_password.FormValue
		admins.admin_name.CurrentValue = admins.admin_name.FormValue
		admins.admin_email.CurrentValue = admins.admin_email.FormValue
		admins.admin_tel.CurrentValue = admins.admin_tel.FormValue
		admins.admin_permis.CurrentValue = admins.admin_permis.FormValue
		admins.admin_create.CurrentValue = admins.admin_create.FormValue
		admins.admin_create.CurrentValue = ew_UnFormatDateTime(admins.admin_create.CurrentValue, 8)
		admins.admin_update.CurrentValue = admins.admin_update.FormValue
		admins.admin_update.CurrentValue = ew_UnFormatDateTime(admins.admin_update.CurrentValue, 8)
		admins.last_online.CurrentValue = admins.last_online.FormValue
		admins.last_online.CurrentValue = ew_UnFormatDateTime(admins.last_online.CurrentValue, 8)
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

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If admins.GetKey("admin_id")&"" <> "" Then
			admins.admin_id.CurrentValue = admins.GetKey("admin_id") ' admin_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			admins.CurrentFilter = admins.KeyFilter
			Dim sSql
			sSql = admins.SQL
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

		' ---------
		'  Add Row
		' ---------

		ElseIf admins.RowType = EW_ROWTYPE_ADD Then ' Add row

			' admin_username
			admins.admin_username.EditCustomAttributes = ""
			admins.admin_username.EditValue = ew_HtmlEncode(admins.admin_username.CurrentValue)
			admins.admin_username.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(admins.admin_username.FldCaption))

			' admin_password
			admins.admin_password.EditCustomAttributes = ""
			admins.admin_password.EditValue = ew_HtmlEncode(admins.admin_password.CurrentValue)
			admins.admin_password.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(admins.admin_password.FldCaption))

			' admin_name
			admins.admin_name.EditCustomAttributes = ""
			admins.admin_name.EditValue = ew_HtmlEncode(admins.admin_name.CurrentValue)
			admins.admin_name.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(admins.admin_name.FldCaption))

			' admin_email
			admins.admin_email.EditCustomAttributes = ""
			admins.admin_email.EditValue = ew_HtmlEncode(admins.admin_email.CurrentValue)
			admins.admin_email.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(admins.admin_email.FldCaption))

			' admin_tel
			admins.admin_tel.EditCustomAttributes = ""
			admins.admin_tel.EditValue = ew_HtmlEncode(admins.admin_tel.CurrentValue)
			admins.admin_tel.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(admins.admin_tel.FldCaption))

			' admin_permis
			admins.admin_permis.EditCustomAttributes = ""
			admins.admin_permis.EditValue = ew_HtmlEncode(admins.admin_permis.CurrentValue)
			admins.admin_permis.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(admins.admin_permis.FldCaption))

			' admin_create
			admins.admin_create.EditCustomAttributes = ""
			admins.admin_create.EditValue = ew_HtmlEncode(admins.admin_create.CurrentValue)
			admins.admin_create.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(admins.admin_create.FldCaption))

			' admin_update
			admins.admin_update.EditCustomAttributes = ""
			admins.admin_update.EditValue = ew_HtmlEncode(admins.admin_update.CurrentValue)
			admins.admin_update.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(admins.admin_update.FldCaption))

			' last_online
			admins.last_online.EditCustomAttributes = ""
			admins.last_online.EditValue = ew_HtmlEncode(admins.last_online.CurrentValue)
			admins.last_online.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(admins.last_online.FldCaption))

			' Edit refer script
			' admin_username

			admins.admin_username.HrefValue = ""

			' admin_password
			admins.admin_password.HrefValue = ""

			' admin_name
			admins.admin_name.HrefValue = ""

			' admin_email
			admins.admin_email.HrefValue = ""

			' admin_tel
			admins.admin_tel.HrefValue = ""

			' admin_permis
			admins.admin_permis.HrefValue = ""

			' admin_create
			admins.admin_create.HrefValue = ""

			' admin_update
			admins.admin_update.HrefValue = ""

			' last_online
			admins.last_online.HrefValue = ""
		End If
		If admins.RowType = EW_ROWTYPE_ADD Or admins.RowType = EW_ROWTYPE_EDIT Or admins.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call admins.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If admins.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call admins.Row_Rendered()
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
		admins.CurrentFilter = sFilter
		sSql = admins.SQL
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

		' Field admin_username
		Call admins.admin_username.SetDbValue(Rs, admins.admin_username.CurrentValue, Null, False)

		' Field admin_password
		If Not EW_CASE_SENSITIVE_PASSWORD And Not IsNull(admins.admin_password.CurrentValue) Then admins.admin_password.CurrentValue = LCase(admins.admin_password.CurrentValue)
		If EW_ENCRYPTED_PASSWORD And Not IsNull(admins.admin_password.CurrentValue) Then admins.admin_password.CurrentValue = MD5(admins.admin_password.CurrentValue)
		Call admins.admin_password.SetDbValue(Rs, admins.admin_password.CurrentValue, Null, False)

		' Field admin_name
		Call admins.admin_name.SetDbValue(Rs, admins.admin_name.CurrentValue, Null, False)

		' Field admin_email
		Call admins.admin_email.SetDbValue(Rs, admins.admin_email.CurrentValue, Null, False)

		' Field admin_tel
		Call admins.admin_tel.SetDbValue(Rs, admins.admin_tel.CurrentValue, Null, False)

		' Field admin_permis
		Call admins.admin_permis.SetDbValue(Rs, admins.admin_permis.CurrentValue, Null, (admins.admin_permis.CurrentValue&"" = ""))

		' Field admin_create
		Call admins.admin_create.SetDbValue(Rs, admins.admin_create.CurrentValue, Null, False)

		' Field admin_update
		Call admins.admin_update.SetDbValue(Rs, admins.admin_update.CurrentValue, Null, False)

		' Field last_online
		Call admins.last_online.SetDbValue(Rs, admins.last_online.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = admins.Row_Inserting(RsOld, Rs)
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
			ElseIf admins.CancelMessage <> "" Then
				FailureMessage = admins.CancelMessage
				admins.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("InsertCancelled")
			End If
			AddRow = False
		End If
		Rs.Close
		Set Rs = Nothing
		If AddRow Then
			admins.admin_id.DbValue = RsNew("admin_id")
		End If
		If AddRow Then

			' Call Row Inserted event
			Call admins.Row_Inserted(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", admins.TableVar, "pom_adminslist.asp", admins.TableVar, True)
		PageId = ew_IIf(admins.CurrentAction = "C", "Copy", "Add")
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
