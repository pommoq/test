<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_person_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim person_th_edit
Set person_th_edit = New cperson_th_edit
Set Page = person_th_edit

' Page init processing
person_th_edit.Page_Init()

' Page main processing
person_th_edit.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
person_th_edit.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var person_th_edit = new ew_Page("person_th_edit");
person_th_edit.PageID = "edit"; // Page ID
var EW_PAGE_ID = person_th_edit.PageID; // For backward compatibility
// Form object
var fperson_thedit = new ew_Form("fperson_thedit");
// Validate form
fperson_thedit.Validate = function() {
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
			elm = this.GetElements("x" + infix + "_per_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(person_th.per_id.FldErrMsg) %>");
			elm = this.GetElements("x" + infix + "_dept_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(person_th.dept_id.FldErrMsg) %>");
			elm = this.GetElements("x" + infix + "_office_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(person_th.office_id.FldErrMsg) %>");
			elm = this.GetElements("x" + infix + "_per_sort");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(person_th.per_sort.FldErrMsg) %>");
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
fperson_thedit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fperson_thedit.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fperson_thedit.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If person_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% person_th_edit.ShowPageHeader() %>
<% person_th_edit.ShowMessage %>
<form name="fperson_thedit" id="fperson_thedit" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="a_table" id="a_table" value="person_th">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table class="ewGrid"><tr><td>
<table id="tbl_person_thedit" class="table table-bordered table-striped">
<% If person_th.per_id.Visible Then ' per_id %>
	<tr id="r_per_id">
		<td><span id="elh_person_th_per_id"><%= person_th.per_id.FldCaption %></span></td>
		<td<%= person_th.per_id.CellAttributes %>>
<span id="el_person_th_per_id" class="control-group">
<span<%= person_th.per_id.ViewAttributes %>>
<%= person_th.per_id.EditValue %>
</span>
</span>
<input type="hidden" data-field="x_per_id" name="x_per_id" id="x_per_id" value="<%= Server.HTMLEncode(person_th.per_id.CurrentValue&"") %>">
<%= person_th.per_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If person_th.dept_id.Visible Then ' dept_id %>
	<tr id="r_dept_id">
		<td><span id="elh_person_th_dept_id"><%= person_th.dept_id.FldCaption %></span></td>
		<td<%= person_th.dept_id.CellAttributes %>>
<span id="el_person_th_dept_id" class="control-group">
<input type="text" data-field="x_dept_id" name="x_dept_id" id="x_dept_id" size="30" placeholder="<%= person_th.dept_id.PlaceHolder %>" value="<%= person_th.dept_id.EditValue %>"<%= person_th.dept_id.EditAttributes %>>
</span>
<%= person_th.dept_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If person_th.office_id.Visible Then ' office_id %>
	<tr id="r_office_id">
		<td><span id="elh_person_th_office_id"><%= person_th.office_id.FldCaption %></span></td>
		<td<%= person_th.office_id.CellAttributes %>>
<span id="el_person_th_office_id" class="control-group">
<input type="text" data-field="x_office_id" name="x_office_id" id="x_office_id" size="30" placeholder="<%= person_th.office_id.PlaceHolder %>" value="<%= person_th.office_id.EditValue %>"<%= person_th.office_id.EditAttributes %>>
</span>
<%= person_th.office_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If person_th.per_img.Visible Then ' per_img %>
	<tr id="r_per_img">
		<td><span id="elh_person_th_per_img"><%= person_th.per_img.FldCaption %></span></td>
		<td<%= person_th.per_img.CellAttributes %>>
<span id="el_person_th_per_img" class="control-group">
<input type="text" data-field="x_per_img" name="x_per_img" id="x_per_img" size="30" maxlength="255" placeholder="<%= person_th.per_img.PlaceHolder %>" value="<%= person_th.per_img.EditValue %>"<%= person_th.per_img.EditAttributes %>>
</span>
<%= person_th.per_img.CustomMsg %></td>
	</tr>
<% End If %>
<% If person_th.per_en_name.Visible Then ' per_en_name %>
	<tr id="r_per_en_name">
		<td><span id="elh_person_th_per_en_name"><%= person_th.per_en_name.FldCaption %></span></td>
		<td<%= person_th.per_en_name.CellAttributes %>>
<span id="el_person_th_per_en_name" class="control-group">
<input type="text" data-field="x_per_en_name" name="x_per_en_name" id="x_per_en_name" size="30" maxlength="255" placeholder="<%= person_th.per_en_name.PlaceHolder %>" value="<%= person_th.per_en_name.EditValue %>"<%= person_th.per_en_name.EditAttributes %>>
</span>
<%= person_th.per_en_name.CustomMsg %></td>
	</tr>
<% End If %>
<% If person_th.per_th_name.Visible Then ' per_th_name %>
	<tr id="r_per_th_name">
		<td><span id="elh_person_th_per_th_name"><%= person_th.per_th_name.FldCaption %></span></td>
		<td<%= person_th.per_th_name.CellAttributes %>>
<span id="el_person_th_per_th_name" class="control-group">
<input type="text" data-field="x_per_th_name" name="x_per_th_name" id="x_per_th_name" size="30" maxlength="255" placeholder="<%= person_th.per_th_name.PlaceHolder %>" value="<%= person_th.per_th_name.EditValue %>"<%= person_th.per_th_name.EditAttributes %>>
</span>
<%= person_th.per_th_name.CustomMsg %></td>
	</tr>
<% End If %>
<% If person_th.per_position.Visible Then ' per_position %>
	<tr id="r_per_position">
		<td><span id="elh_person_th_per_position"><%= person_th.per_position.FldCaption %></span></td>
		<td<%= person_th.per_position.CellAttributes %>>
<span id="el_person_th_per_position" class="control-group">
<input type="text" data-field="x_per_position" name="x_per_position" id="x_per_position" size="30" maxlength="255" placeholder="<%= person_th.per_position.PlaceHolder %>" value="<%= person_th.per_position.EditValue %>"<%= person_th.per_position.EditAttributes %>>
</span>
<%= person_th.per_position.CustomMsg %></td>
	</tr>
<% End If %>
<% If person_th.per_mobile.Visible Then ' per_mobile %>
	<tr id="r_per_mobile">
		<td><span id="elh_person_th_per_mobile"><%= person_th.per_mobile.FldCaption %></span></td>
		<td<%= person_th.per_mobile.CellAttributes %>>
<span id="el_person_th_per_mobile" class="control-group">
<input type="text" data-field="x_per_mobile" name="x_per_mobile" id="x_per_mobile" size="30" maxlength="255" placeholder="<%= person_th.per_mobile.PlaceHolder %>" value="<%= person_th.per_mobile.EditValue %>"<%= person_th.per_mobile.EditAttributes %>>
</span>
<%= person_th.per_mobile.CustomMsg %></td>
	</tr>
<% End If %>
<% If person_th.per_tel.Visible Then ' per_tel %>
	<tr id="r_per_tel">
		<td><span id="elh_person_th_per_tel"><%= person_th.per_tel.FldCaption %></span></td>
		<td<%= person_th.per_tel.CellAttributes %>>
<span id="el_person_th_per_tel" class="control-group">
<input type="text" data-field="x_per_tel" name="x_per_tel" id="x_per_tel" size="30" maxlength="255" placeholder="<%= person_th.per_tel.PlaceHolder %>" value="<%= person_th.per_tel.EditValue %>"<%= person_th.per_tel.EditAttributes %>>
</span>
<%= person_th.per_tel.CustomMsg %></td>
	</tr>
<% End If %>
<% If person_th.per_fax.Visible Then ' per_fax %>
	<tr id="r_per_fax">
		<td><span id="elh_person_th_per_fax"><%= person_th.per_fax.FldCaption %></span></td>
		<td<%= person_th.per_fax.CellAttributes %>>
<span id="el_person_th_per_fax" class="control-group">
<input type="text" data-field="x_per_fax" name="x_per_fax" id="x_per_fax" size="30" maxlength="255" placeholder="<%= person_th.per_fax.PlaceHolder %>" value="<%= person_th.per_fax.EditValue %>"<%= person_th.per_fax.EditAttributes %>>
</span>
<%= person_th.per_fax.CustomMsg %></td>
	</tr>
<% End If %>
<% If person_th.per_email.Visible Then ' per_email %>
	<tr id="r_per_email">
		<td><span id="elh_person_th_per_email"><%= person_th.per_email.FldCaption %></span></td>
		<td<%= person_th.per_email.CellAttributes %>>
<span id="el_person_th_per_email" class="control-group">
<input type="text" data-field="x_per_email" name="x_per_email" id="x_per_email" size="30" maxlength="255" placeholder="<%= person_th.per_email.PlaceHolder %>" value="<%= person_th.per_email.EditValue %>"<%= person_th.per_email.EditAttributes %>>
</span>
<%= person_th.per_email.CustomMsg %></td>
	</tr>
<% End If %>
<% If person_th.per_address.Visible Then ' per_address %>
	<tr id="r_per_address">
		<td><span id="elh_person_th_per_address"><%= person_th.per_address.FldCaption %></span></td>
		<td<%= person_th.per_address.CellAttributes %>>
<span id="el_person_th_per_address" class="control-group">
<input type="text" data-field="x_per_address" name="x_per_address" id="x_per_address" size="30" maxlength="255" placeholder="<%= person_th.per_address.PlaceHolder %>" value="<%= person_th.per_address.EditValue %>"<%= person_th.per_address.EditAttributes %>>
</span>
<%= person_th.per_address.CustomMsg %></td>
	</tr>
<% End If %>
<% If person_th.per_show.Visible Then ' per_show %>
	<tr id="r_per_show">
		<td><span id="elh_person_th_per_show"><%= person_th.per_show.FldCaption %></span></td>
		<td<%= person_th.per_show.CellAttributes %>>
<span id="el_person_th_per_show" class="control-group">
<input type="text" data-field="x_per_show" name="x_per_show" id="x_per_show" size="30" maxlength="255" placeholder="<%= person_th.per_show.PlaceHolder %>" value="<%= person_th.per_show.EditValue %>"<%= person_th.per_show.EditAttributes %>>
</span>
<%= person_th.per_show.CustomMsg %></td>
	</tr>
<% End If %>
<% If person_th.per_create.Visible Then ' per_create %>
	<tr id="r_per_create">
		<td><span id="elh_person_th_per_create"><%= person_th.per_create.FldCaption %></span></td>
		<td<%= person_th.per_create.CellAttributes %>>
<span id="el_person_th_per_create" class="control-group">
<input type="text" data-field="x_per_create" name="x_per_create" id="x_per_create" size="30" maxlength="255" placeholder="<%= person_th.per_create.PlaceHolder %>" value="<%= person_th.per_create.EditValue %>"<%= person_th.per_create.EditAttributes %>>
</span>
<%= person_th.per_create.CustomMsg %></td>
	</tr>
<% End If %>
<% If person_th.per_update.Visible Then ' per_update %>
	<tr id="r_per_update">
		<td><span id="elh_person_th_per_update"><%= person_th.per_update.FldCaption %></span></td>
		<td<%= person_th.per_update.CellAttributes %>>
<span id="el_person_th_per_update" class="control-group">
<input type="text" data-field="x_per_update" name="x_per_update" id="x_per_update" size="30" maxlength="255" placeholder="<%= person_th.per_update.PlaceHolder %>" value="<%= person_th.per_update.EditValue %>"<%= person_th.per_update.EditAttributes %>>
</span>
<%= person_th.per_update.CustomMsg %></td>
	</tr>
<% End If %>
<% If person_th.per_sort.Visible Then ' per_sort %>
	<tr id="r_per_sort">
		<td><span id="elh_person_th_per_sort"><%= person_th.per_sort.FldCaption %></span></td>
		<td<%= person_th.per_sort.CellAttributes %>>
<span id="el_person_th_per_sort" class="control-group">
<input type="text" data-field="x_per_sort" name="x_per_sort" id="x_per_sort" size="30" placeholder="<%= person_th.per_sort.PlaceHolder %>" value="<%= person_th.per_sort.EditValue %>"<%= person_th.per_sort.EditAttributes %>>
</span>
<%= person_th.per_sort.CustomMsg %></td>
	</tr>
<% End If %>
<% If person_th.per_department.Visible Then ' per_department %>
	<tr id="r_per_department">
		<td><span id="elh_person_th_per_department"><%= person_th.per_department.FldCaption %></span></td>
		<td<%= person_th.per_department.CellAttributes %>>
<span id="el_person_th_per_department" class="control-group">
<input type="text" data-field="x_per_department" name="x_per_department" id="x_per_department" size="30" maxlength="255" placeholder="<%= person_th.per_department.PlaceHolder %>" value="<%= person_th.per_department.EditValue %>"<%= person_th.per_department.EditAttributes %>>
</span>
<%= person_th.per_department.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("EditBtn") %></button>
</form>
<script type="text/javascript">
fperson_thedit.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
person_th_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set person_th_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cperson_th_edit

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
		TableName = "person_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "person_th_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If person_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & person_th.TableVar & "&" ' add page token
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
		If person_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (person_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (person_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(person_th) Then Set person_th = New cperson_th
		Set Table = person_th

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "person_th"

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

		person_th.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set person_th = Nothing
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
		If Request.QueryString("per_id").Count > 0 Then
			person_th.per_id.QueryStringValue = Request.QueryString("per_id")
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Process form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			person_th.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values
		Else
			person_th.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If person_th.per_id.CurrentValue = "" Then Call Page_Terminate("pom_person_thlist.asp") ' Invalid key, return to list

		' Validate form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			If Not ValidateForm() Then
				person_th.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				person_th.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		End If
		Select Case person_th.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_person_thlist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				person_th.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					sReturnUrl = person_th.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					person_th.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		person_th.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call person_th.ResetAttrs()
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
				person_th.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					person_th.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = person_th.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			person_th.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			person_th.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			person_th.StartRecordNumber = StartRec
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
		If Not person_th.per_id.FldIsDetailKey Then person_th.per_id.FormValue = ObjForm.GetValue("x_per_id")
		If Not person_th.dept_id.FldIsDetailKey Then person_th.dept_id.FormValue = ObjForm.GetValue("x_dept_id")
		If Not person_th.office_id.FldIsDetailKey Then person_th.office_id.FormValue = ObjForm.GetValue("x_office_id")
		If Not person_th.per_img.FldIsDetailKey Then person_th.per_img.FormValue = ObjForm.GetValue("x_per_img")
		If Not person_th.per_en_name.FldIsDetailKey Then person_th.per_en_name.FormValue = ObjForm.GetValue("x_per_en_name")
		If Not person_th.per_th_name.FldIsDetailKey Then person_th.per_th_name.FormValue = ObjForm.GetValue("x_per_th_name")
		If Not person_th.per_position.FldIsDetailKey Then person_th.per_position.FormValue = ObjForm.GetValue("x_per_position")
		If Not person_th.per_mobile.FldIsDetailKey Then person_th.per_mobile.FormValue = ObjForm.GetValue("x_per_mobile")
		If Not person_th.per_tel.FldIsDetailKey Then person_th.per_tel.FormValue = ObjForm.GetValue("x_per_tel")
		If Not person_th.per_fax.FldIsDetailKey Then person_th.per_fax.FormValue = ObjForm.GetValue("x_per_fax")
		If Not person_th.per_email.FldIsDetailKey Then person_th.per_email.FormValue = ObjForm.GetValue("x_per_email")
		If Not person_th.per_address.FldIsDetailKey Then person_th.per_address.FormValue = ObjForm.GetValue("x_per_address")
		If Not person_th.per_show.FldIsDetailKey Then person_th.per_show.FormValue = ObjForm.GetValue("x_per_show")
		If Not person_th.per_create.FldIsDetailKey Then person_th.per_create.FormValue = ObjForm.GetValue("x_per_create")
		If Not person_th.per_update.FldIsDetailKey Then person_th.per_update.FormValue = ObjForm.GetValue("x_per_update")
		If Not person_th.per_sort.FldIsDetailKey Then person_th.per_sort.FormValue = ObjForm.GetValue("x_per_sort")
		If Not person_th.per_department.FldIsDetailKey Then person_th.per_department.FormValue = ObjForm.GetValue("x_per_department")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		person_th.per_id.CurrentValue = person_th.per_id.FormValue
		person_th.dept_id.CurrentValue = person_th.dept_id.FormValue
		person_th.office_id.CurrentValue = person_th.office_id.FormValue
		person_th.per_img.CurrentValue = person_th.per_img.FormValue
		person_th.per_en_name.CurrentValue = person_th.per_en_name.FormValue
		person_th.per_th_name.CurrentValue = person_th.per_th_name.FormValue
		person_th.per_position.CurrentValue = person_th.per_position.FormValue
		person_th.per_mobile.CurrentValue = person_th.per_mobile.FormValue
		person_th.per_tel.CurrentValue = person_th.per_tel.FormValue
		person_th.per_fax.CurrentValue = person_th.per_fax.FormValue
		person_th.per_email.CurrentValue = person_th.per_email.FormValue
		person_th.per_address.CurrentValue = person_th.per_address.FormValue
		person_th.per_show.CurrentValue = person_th.per_show.FormValue
		person_th.per_create.CurrentValue = person_th.per_create.FormValue
		person_th.per_update.CurrentValue = person_th.per_update.FormValue
		person_th.per_sort.CurrentValue = person_th.per_sort.FormValue
		person_th.per_department.CurrentValue = person_th.per_department.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = person_th.KeyFilter

		' Call Row Selecting event
		Call person_th.Row_Selecting(sFilter)

		' Load sql based on filter
		person_th.CurrentFilter = sFilter
		sSql = person_th.SQL
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
		Call person_th.Row_Selected(RsRow)
		person_th.per_id.DbValue = RsRow("per_id")
		person_th.dept_id.DbValue = RsRow("dept_id")
		person_th.office_id.DbValue = RsRow("office_id")
		person_th.per_img.DbValue = RsRow("per_img")
		person_th.per_en_name.DbValue = RsRow("per_en_name")
		person_th.per_th_name.DbValue = RsRow("per_th_name")
		person_th.per_position.DbValue = RsRow("per_position")
		person_th.per_mobile.DbValue = RsRow("per_mobile")
		person_th.per_tel.DbValue = RsRow("per_tel")
		person_th.per_fax.DbValue = RsRow("per_fax")
		person_th.per_email.DbValue = RsRow("per_email")
		person_th.per_address.DbValue = RsRow("per_address")
		person_th.per_show.DbValue = RsRow("per_show")
		person_th.per_create.DbValue = RsRow("per_create")
		person_th.per_update.DbValue = RsRow("per_update")
		person_th.per_sort.DbValue = RsRow("per_sort")
		person_th.per_department.DbValue = RsRow("per_department")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		person_th.per_id.m_DbValue = Rs("per_id")
		person_th.dept_id.m_DbValue = Rs("dept_id")
		person_th.office_id.m_DbValue = Rs("office_id")
		person_th.per_img.m_DbValue = Rs("per_img")
		person_th.per_en_name.m_DbValue = Rs("per_en_name")
		person_th.per_th_name.m_DbValue = Rs("per_th_name")
		person_th.per_position.m_DbValue = Rs("per_position")
		person_th.per_mobile.m_DbValue = Rs("per_mobile")
		person_th.per_tel.m_DbValue = Rs("per_tel")
		person_th.per_fax.m_DbValue = Rs("per_fax")
		person_th.per_email.m_DbValue = Rs("per_email")
		person_th.per_address.m_DbValue = Rs("per_address")
		person_th.per_show.m_DbValue = Rs("per_show")
		person_th.per_create.m_DbValue = Rs("per_create")
		person_th.per_update.m_DbValue = Rs("per_update")
		person_th.per_sort.m_DbValue = Rs("per_sort")
		person_th.per_department.m_DbValue = Rs("per_department")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call person_th.Row_Rendering()

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

		If person_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' per_id
			person_th.per_id.ViewValue = person_th.per_id.CurrentValue
			person_th.per_id.ViewCustomAttributes = ""

			' dept_id
			person_th.dept_id.ViewValue = person_th.dept_id.CurrentValue
			person_th.dept_id.ViewCustomAttributes = ""

			' office_id
			person_th.office_id.ViewValue = person_th.office_id.CurrentValue
			person_th.office_id.ViewCustomAttributes = ""

			' per_img
			person_th.per_img.ViewValue = person_th.per_img.CurrentValue
			person_th.per_img.ViewCustomAttributes = ""

			' per_en_name
			person_th.per_en_name.ViewValue = person_th.per_en_name.CurrentValue
			person_th.per_en_name.ViewCustomAttributes = ""

			' per_th_name
			person_th.per_th_name.ViewValue = person_th.per_th_name.CurrentValue
			person_th.per_th_name.ViewCustomAttributes = ""

			' per_position
			person_th.per_position.ViewValue = person_th.per_position.CurrentValue
			person_th.per_position.ViewCustomAttributes = ""

			' per_mobile
			person_th.per_mobile.ViewValue = person_th.per_mobile.CurrentValue
			person_th.per_mobile.ViewCustomAttributes = ""

			' per_tel
			person_th.per_tel.ViewValue = person_th.per_tel.CurrentValue
			person_th.per_tel.ViewCustomAttributes = ""

			' per_fax
			person_th.per_fax.ViewValue = person_th.per_fax.CurrentValue
			person_th.per_fax.ViewCustomAttributes = ""

			' per_email
			person_th.per_email.ViewValue = person_th.per_email.CurrentValue
			person_th.per_email.ViewCustomAttributes = ""

			' per_address
			person_th.per_address.ViewValue = person_th.per_address.CurrentValue
			person_th.per_address.ViewCustomAttributes = ""

			' per_show
			person_th.per_show.ViewValue = person_th.per_show.CurrentValue
			person_th.per_show.ViewCustomAttributes = ""

			' per_create
			person_th.per_create.ViewValue = person_th.per_create.CurrentValue
			person_th.per_create.ViewCustomAttributes = ""

			' per_update
			person_th.per_update.ViewValue = person_th.per_update.CurrentValue
			person_th.per_update.ViewCustomAttributes = ""

			' per_sort
			person_th.per_sort.ViewValue = person_th.per_sort.CurrentValue
			person_th.per_sort.ViewCustomAttributes = ""

			' per_department
			person_th.per_department.ViewValue = person_th.per_department.CurrentValue
			person_th.per_department.ViewCustomAttributes = ""

			' View refer script
			' per_id

			person_th.per_id.LinkCustomAttributes = ""
			person_th.per_id.HrefValue = ""
			person_th.per_id.TooltipValue = ""

			' dept_id
			person_th.dept_id.LinkCustomAttributes = ""
			person_th.dept_id.HrefValue = ""
			person_th.dept_id.TooltipValue = ""

			' office_id
			person_th.office_id.LinkCustomAttributes = ""
			person_th.office_id.HrefValue = ""
			person_th.office_id.TooltipValue = ""

			' per_img
			person_th.per_img.LinkCustomAttributes = ""
			person_th.per_img.HrefValue = ""
			person_th.per_img.TooltipValue = ""

			' per_en_name
			person_th.per_en_name.LinkCustomAttributes = ""
			person_th.per_en_name.HrefValue = ""
			person_th.per_en_name.TooltipValue = ""

			' per_th_name
			person_th.per_th_name.LinkCustomAttributes = ""
			person_th.per_th_name.HrefValue = ""
			person_th.per_th_name.TooltipValue = ""

			' per_position
			person_th.per_position.LinkCustomAttributes = ""
			person_th.per_position.HrefValue = ""
			person_th.per_position.TooltipValue = ""

			' per_mobile
			person_th.per_mobile.LinkCustomAttributes = ""
			person_th.per_mobile.HrefValue = ""
			person_th.per_mobile.TooltipValue = ""

			' per_tel
			person_th.per_tel.LinkCustomAttributes = ""
			person_th.per_tel.HrefValue = ""
			person_th.per_tel.TooltipValue = ""

			' per_fax
			person_th.per_fax.LinkCustomAttributes = ""
			person_th.per_fax.HrefValue = ""
			person_th.per_fax.TooltipValue = ""

			' per_email
			person_th.per_email.LinkCustomAttributes = ""
			person_th.per_email.HrefValue = ""
			person_th.per_email.TooltipValue = ""

			' per_address
			person_th.per_address.LinkCustomAttributes = ""
			person_th.per_address.HrefValue = ""
			person_th.per_address.TooltipValue = ""

			' per_show
			person_th.per_show.LinkCustomAttributes = ""
			person_th.per_show.HrefValue = ""
			person_th.per_show.TooltipValue = ""

			' per_create
			person_th.per_create.LinkCustomAttributes = ""
			person_th.per_create.HrefValue = ""
			person_th.per_create.TooltipValue = ""

			' per_update
			person_th.per_update.LinkCustomAttributes = ""
			person_th.per_update.HrefValue = ""
			person_th.per_update.TooltipValue = ""

			' per_sort
			person_th.per_sort.LinkCustomAttributes = ""
			person_th.per_sort.HrefValue = ""
			person_th.per_sort.TooltipValue = ""

			' per_department
			person_th.per_department.LinkCustomAttributes = ""
			person_th.per_department.HrefValue = ""
			person_th.per_department.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf person_th.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' per_id
			person_th.per_id.EditCustomAttributes = ""
			person_th.per_id.EditValue = person_th.per_id.CurrentValue
			person_th.per_id.ViewCustomAttributes = ""

			' dept_id
			person_th.dept_id.EditCustomAttributes = ""
			person_th.dept_id.EditValue = ew_HtmlEncode(person_th.dept_id.CurrentValue)
			person_th.dept_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person_th.dept_id.FldCaption))

			' office_id
			person_th.office_id.EditCustomAttributes = ""
			person_th.office_id.EditValue = ew_HtmlEncode(person_th.office_id.CurrentValue)
			person_th.office_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person_th.office_id.FldCaption))

			' per_img
			person_th.per_img.EditCustomAttributes = ""
			person_th.per_img.EditValue = ew_HtmlEncode(person_th.per_img.CurrentValue)
			person_th.per_img.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person_th.per_img.FldCaption))

			' per_en_name
			person_th.per_en_name.EditCustomAttributes = ""
			person_th.per_en_name.EditValue = ew_HtmlEncode(person_th.per_en_name.CurrentValue)
			person_th.per_en_name.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person_th.per_en_name.FldCaption))

			' per_th_name
			person_th.per_th_name.EditCustomAttributes = ""
			person_th.per_th_name.EditValue = ew_HtmlEncode(person_th.per_th_name.CurrentValue)
			person_th.per_th_name.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person_th.per_th_name.FldCaption))

			' per_position
			person_th.per_position.EditCustomAttributes = ""
			person_th.per_position.EditValue = ew_HtmlEncode(person_th.per_position.CurrentValue)
			person_th.per_position.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person_th.per_position.FldCaption))

			' per_mobile
			person_th.per_mobile.EditCustomAttributes = ""
			person_th.per_mobile.EditValue = ew_HtmlEncode(person_th.per_mobile.CurrentValue)
			person_th.per_mobile.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person_th.per_mobile.FldCaption))

			' per_tel
			person_th.per_tel.EditCustomAttributes = ""
			person_th.per_tel.EditValue = ew_HtmlEncode(person_th.per_tel.CurrentValue)
			person_th.per_tel.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person_th.per_tel.FldCaption))

			' per_fax
			person_th.per_fax.EditCustomAttributes = ""
			person_th.per_fax.EditValue = ew_HtmlEncode(person_th.per_fax.CurrentValue)
			person_th.per_fax.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person_th.per_fax.FldCaption))

			' per_email
			person_th.per_email.EditCustomAttributes = ""
			person_th.per_email.EditValue = ew_HtmlEncode(person_th.per_email.CurrentValue)
			person_th.per_email.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person_th.per_email.FldCaption))

			' per_address
			person_th.per_address.EditCustomAttributes = ""
			person_th.per_address.EditValue = ew_HtmlEncode(person_th.per_address.CurrentValue)
			person_th.per_address.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person_th.per_address.FldCaption))

			' per_show
			person_th.per_show.EditCustomAttributes = ""
			person_th.per_show.EditValue = ew_HtmlEncode(person_th.per_show.CurrentValue)
			person_th.per_show.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person_th.per_show.FldCaption))

			' per_create
			person_th.per_create.EditCustomAttributes = ""
			person_th.per_create.EditValue = ew_HtmlEncode(person_th.per_create.CurrentValue)
			person_th.per_create.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person_th.per_create.FldCaption))

			' per_update
			person_th.per_update.EditCustomAttributes = ""
			person_th.per_update.EditValue = ew_HtmlEncode(person_th.per_update.CurrentValue)
			person_th.per_update.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person_th.per_update.FldCaption))

			' per_sort
			person_th.per_sort.EditCustomAttributes = ""
			person_th.per_sort.EditValue = ew_HtmlEncode(person_th.per_sort.CurrentValue)
			person_th.per_sort.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person_th.per_sort.FldCaption))

			' per_department
			person_th.per_department.EditCustomAttributes = ""
			person_th.per_department.EditValue = ew_HtmlEncode(person_th.per_department.CurrentValue)
			person_th.per_department.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person_th.per_department.FldCaption))

			' Edit refer script
			' per_id

			person_th.per_id.HrefValue = ""

			' dept_id
			person_th.dept_id.HrefValue = ""

			' office_id
			person_th.office_id.HrefValue = ""

			' per_img
			person_th.per_img.HrefValue = ""

			' per_en_name
			person_th.per_en_name.HrefValue = ""

			' per_th_name
			person_th.per_th_name.HrefValue = ""

			' per_position
			person_th.per_position.HrefValue = ""

			' per_mobile
			person_th.per_mobile.HrefValue = ""

			' per_tel
			person_th.per_tel.HrefValue = ""

			' per_fax
			person_th.per_fax.HrefValue = ""

			' per_email
			person_th.per_email.HrefValue = ""

			' per_address
			person_th.per_address.HrefValue = ""

			' per_show
			person_th.per_show.HrefValue = ""

			' per_create
			person_th.per_create.HrefValue = ""

			' per_update
			person_th.per_update.HrefValue = ""

			' per_sort
			person_th.per_sort.HrefValue = ""

			' per_department
			person_th.per_department.HrefValue = ""
		End If
		If person_th.RowType = EW_ROWTYPE_ADD Or person_th.RowType = EW_ROWTYPE_EDIT Or person_th.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call person_th.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If person_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call person_th.Row_Rendered()
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
		If Not ew_CheckInteger(person_th.per_id.FormValue) Then
			Call ew_AddMessage(gsFormError, person_th.per_id.FldErrMsg)
		End If
		If Not ew_CheckInteger(person_th.dept_id.FormValue) Then
			Call ew_AddMessage(gsFormError, person_th.dept_id.FldErrMsg)
		End If
		If Not ew_CheckInteger(person_th.office_id.FormValue) Then
			Call ew_AddMessage(gsFormError, person_th.office_id.FldErrMsg)
		End If
		If Not ew_CheckInteger(person_th.per_sort.FormValue) Then
			Call ew_AddMessage(gsFormError, person_th.per_sort.FldErrMsg)
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
		sFilter = person_th.KeyFilter
		person_th.CurrentFilter  = sFilter
		sSql = person_th.SQL
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

			' Field per_id
			' Field dept_id

			Call person_th.dept_id.SetDbValue(Rs, person_th.dept_id.CurrentValue, Null, person_th.dept_id.ReadOnly)

			' Field office_id
			Call person_th.office_id.SetDbValue(Rs, person_th.office_id.CurrentValue, Null, person_th.office_id.ReadOnly)

			' Field per_img
			Call person_th.per_img.SetDbValue(Rs, person_th.per_img.CurrentValue, Null, person_th.per_img.ReadOnly)

			' Field per_en_name
			Call person_th.per_en_name.SetDbValue(Rs, person_th.per_en_name.CurrentValue, Null, person_th.per_en_name.ReadOnly)

			' Field per_th_name
			Call person_th.per_th_name.SetDbValue(Rs, person_th.per_th_name.CurrentValue, Null, person_th.per_th_name.ReadOnly)

			' Field per_position
			Call person_th.per_position.SetDbValue(Rs, person_th.per_position.CurrentValue, Null, person_th.per_position.ReadOnly)

			' Field per_mobile
			Call person_th.per_mobile.SetDbValue(Rs, person_th.per_mobile.CurrentValue, Null, person_th.per_mobile.ReadOnly)

			' Field per_tel
			Call person_th.per_tel.SetDbValue(Rs, person_th.per_tel.CurrentValue, Null, person_th.per_tel.ReadOnly)

			' Field per_fax
			Call person_th.per_fax.SetDbValue(Rs, person_th.per_fax.CurrentValue, Null, person_th.per_fax.ReadOnly)

			' Field per_email
			Call person_th.per_email.SetDbValue(Rs, person_th.per_email.CurrentValue, Null, person_th.per_email.ReadOnly)

			' Field per_address
			Call person_th.per_address.SetDbValue(Rs, person_th.per_address.CurrentValue, Null, person_th.per_address.ReadOnly)

			' Field per_show
			Call person_th.per_show.SetDbValue(Rs, person_th.per_show.CurrentValue, Null, person_th.per_show.ReadOnly)

			' Field per_create
			Call person_th.per_create.SetDbValue(Rs, person_th.per_create.CurrentValue, Null, person_th.per_create.ReadOnly)

			' Field per_update
			Call person_th.per_update.SetDbValue(Rs, person_th.per_update.CurrentValue, Null, person_th.per_update.ReadOnly)

			' Field per_sort
			Call person_th.per_sort.SetDbValue(Rs, person_th.per_sort.CurrentValue, Null, person_th.per_sort.ReadOnly)

			' Field per_department
			Call person_th.per_department.SetDbValue(Rs, person_th.per_department.CurrentValue, Null, person_th.per_department.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = person_th.Row_Updating(RsOld, Rs)
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
				ElseIf person_th.CancelMessage <> "" Then
					FailureMessage = person_th.CancelMessage
					person_th.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call person_th.Row_Updated(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", person_th.TableVar, "pom_person_thlist.asp", person_th.TableVar, True)
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
