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
Dim admins_edit
Set admins_edit = New cadmins_edit
Set Page = admins_edit

' Page init processing
admins_edit.Page_Init()

' Page main processing
admins_edit.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
admins_edit.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var admins_edit = new ew_Page("admins_edit");
admins_edit.PageID = "edit"; // Page ID
var EW_PAGE_ID = admins_edit.PageID; // For backward compatibility
// Form object
var fadminsedit = new ew_Form("fadminsedit");
// Validate form
fadminsedit.Validate = function() {
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
fadminsedit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fadminsedit.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fadminsedit.ValidateRequired = false; // No JavaScript validation
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
<% admins_edit.ShowPageHeader() %>
<% admins_edit.ShowMessage %>
<form name="fadminsedit" id="fadminsedit" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="a_table" id="a_table" value="admins">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table class="ewGrid"><tr><td>
<table id="tbl_adminsedit" class="table table-bordered table-striped">
<% If admins.admin_id.Visible Then ' admin_id %>
	<tr id="r_admin_id">
		<td><span id="elh_admins_admin_id"><%= admins.admin_id.FldCaption %></span></td>
		<td<%= admins.admin_id.CellAttributes %>>
<span id="el_admins_admin_id" class="control-group">
<span<%= admins.admin_id.ViewAttributes %>>
<%= admins.admin_id.EditValue %>
</span>
</span>
<input type="hidden" data-field="x_admin_id" name="x_admin_id" id="x_admin_id" value="<%= Server.HTMLEncode(admins.admin_id.CurrentValue&"") %>">
<%= admins.admin_id.CustomMsg %></td>
	</tr>
<% End If %>
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
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("EditBtn") %></button>
</form>
<script type="text/javascript">
fadminsedit.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
admins_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set admins_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cadmins_edit

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
		TableName = "admins"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "admins_edit"
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
		EW_PAGE_ID = "edit"

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
		admins.admin_id.Visible = Not admins.IsAdd() And Not admins.IsCopy() And Not admins.IsGridAdd()

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

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		Dim sReturnUrl
		sReturnUrl = ""

		' Load key from QueryString
		If Request.QueryString("admin_id").Count > 0 Then
			admins.admin_id.QueryStringValue = Request.QueryString("admin_id")
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Process form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			admins.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values
		Else
			admins.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If admins.admin_id.CurrentValue = "" Then Call Page_Terminate("pom_adminslist.asp") ' Invalid key, return to list

		' Validate form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			If Not ValidateForm() Then
				admins.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				admins.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		End If
		Select Case admins.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_adminslist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				admins.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					sReturnUrl = admins.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					admins.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		admins.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call admins.ResetAttrs()
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
				admins.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					admins.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = admins.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			admins.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			admins.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			admins.StartRecordNumber = StartRec
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
		If Not admins.admin_id.FldIsDetailKey Then admins.admin_id.FormValue = ObjForm.GetValue("x_admin_id")
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
		Call LoadRow()
		admins.admin_id.CurrentValue = admins.admin_id.FormValue
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

		' ----------
		'  Edit Row
		' ----------

		ElseIf admins.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' admin_id
			admins.admin_id.EditCustomAttributes = ""
			admins.admin_id.EditValue = admins.admin_id.CurrentValue
			admins.admin_id.ViewCustomAttributes = ""

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
			' admin_id

			admins.admin_id.HrefValue = ""

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
		sFilter = admins.KeyFilter
		admins.CurrentFilter  = sFilter
		sSql = admins.SQL
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

			' Field admin_username
			Call admins.admin_username.SetDbValue(Rs, admins.admin_username.CurrentValue, Null, admins.admin_username.ReadOnly)

			' Field admin_password
			If Not EW_CASE_SENSITIVE_PASSWORD And Not IsNull(admins.admin_password.CurrentValue) Then admins.admin_password.CurrentValue = LCase(admins.admin_password.CurrentValue)
			If EW_ENCRYPTED_PASSWORD And Not IsNull(admins.admin_password.CurrentValue) And (admins.admin_password.CurrentValue <> Rs("admin_password")) Then admins.admin_password.CurrentValue = MD5(admins.admin_password.CurrentValue)
			Call admins.admin_password.SetDbValue(Rs, admins.admin_password.CurrentValue, Null, admins.admin_password.ReadOnly)

			' Field admin_name
			Call admins.admin_name.SetDbValue(Rs, admins.admin_name.CurrentValue, Null, admins.admin_name.ReadOnly)

			' Field admin_email
			Call admins.admin_email.SetDbValue(Rs, admins.admin_email.CurrentValue, Null, admins.admin_email.ReadOnly)

			' Field admin_tel
			Call admins.admin_tel.SetDbValue(Rs, admins.admin_tel.CurrentValue, Null, admins.admin_tel.ReadOnly)

			' Field admin_permis
			Call admins.admin_permis.SetDbValue(Rs, admins.admin_permis.CurrentValue, Null, admins.admin_permis.ReadOnly)

			' Field admin_create
			Call admins.admin_create.SetDbValue(Rs, admins.admin_create.CurrentValue, Null, admins.admin_create.ReadOnly)

			' Field admin_update
			Call admins.admin_update.SetDbValue(Rs, admins.admin_update.CurrentValue, Null, admins.admin_update.ReadOnly)

			' Field last_online
			Call admins.last_online.SetDbValue(Rs, admins.last_online.CurrentValue, Null, admins.last_online.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = admins.Row_Updating(RsOld, Rs)
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
				ElseIf admins.CancelMessage <> "" Then
					FailureMessage = admins.CancelMessage
					admins.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call admins.Row_Updated(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", admins.TableVar, "pom_adminslist.asp", admins.TableVar, True)
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
