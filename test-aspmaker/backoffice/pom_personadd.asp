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
Dim person_add
Set person_add = New cperson_add
Set Page = person_add

' Page init processing
person_add.Page_Init()

' Page main processing
person_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
person_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var person_add = new ew_Page("person_add");
person_add.PageID = "add"; // Page ID
var EW_PAGE_ID = person_add.PageID; // For backward compatibility
// Form object
var fpersonadd = new ew_Form("fpersonadd");
// Validate form
fpersonadd.Validate = function() {
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
				return this.OnError(elm, "<%= ew_JsEncode2(person.per_id.FldErrMsg) %>");
			elm = this.GetElements("x" + infix + "_dept_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(person.dept_id.FldErrMsg) %>");
			elm = this.GetElements("x" + infix + "_office_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(person.office_id.FldErrMsg) %>");
			elm = this.GetElements("x" + infix + "_per_sort");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(person.per_sort.FldErrMsg) %>");
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
fpersonadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fpersonadd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fpersonadd.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If person.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% person_add.ShowPageHeader() %>
<% person_add.ShowMessage %>
<form name="fpersonadd" id="fpersonadd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="person">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_personadd" class="table table-bordered table-striped">
<% If person.per_id.Visible Then ' per_id %>
	<tr id="r_per_id">
		<td><span id="elh_person_per_id"><%= person.per_id.FldCaption %></span></td>
		<td<%= person.per_id.CellAttributes %>>
<span id="el_person_per_id" class="control-group">
<input type="text" data-field="x_per_id" name="x_per_id" id="x_per_id" size="30" placeholder="<%= person.per_id.PlaceHolder %>" value="<%= person.per_id.EditValue %>"<%= person.per_id.EditAttributes %>>
</span>
<%= person.per_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If person.dept_id.Visible Then ' dept_id %>
	<tr id="r_dept_id">
		<td><span id="elh_person_dept_id"><%= person.dept_id.FldCaption %></span></td>
		<td<%= person.dept_id.CellAttributes %>>
<span id="el_person_dept_id" class="control-group">
<input type="text" data-field="x_dept_id" name="x_dept_id" id="x_dept_id" size="30" placeholder="<%= person.dept_id.PlaceHolder %>" value="<%= person.dept_id.EditValue %>"<%= person.dept_id.EditAttributes %>>
</span>
<%= person.dept_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If person.office_id.Visible Then ' office_id %>
	<tr id="r_office_id">
		<td><span id="elh_person_office_id"><%= person.office_id.FldCaption %></span></td>
		<td<%= person.office_id.CellAttributes %>>
<span id="el_person_office_id" class="control-group">
<input type="text" data-field="x_office_id" name="x_office_id" id="x_office_id" size="30" placeholder="<%= person.office_id.PlaceHolder %>" value="<%= person.office_id.EditValue %>"<%= person.office_id.EditAttributes %>>
</span>
<%= person.office_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If person.per_img.Visible Then ' per_img %>
	<tr id="r_per_img">
		<td><span id="elh_person_per_img"><%= person.per_img.FldCaption %></span></td>
		<td<%= person.per_img.CellAttributes %>>
<span id="el_person_per_img" class="control-group">
<input type="text" data-field="x_per_img" name="x_per_img" id="x_per_img" size="30" maxlength="255" placeholder="<%= person.per_img.PlaceHolder %>" value="<%= person.per_img.EditValue %>"<%= person.per_img.EditAttributes %>>
</span>
<%= person.per_img.CustomMsg %></td>
	</tr>
<% End If %>
<% If person.per_en_name.Visible Then ' per_en_name %>
	<tr id="r_per_en_name">
		<td><span id="elh_person_per_en_name"><%= person.per_en_name.FldCaption %></span></td>
		<td<%= person.per_en_name.CellAttributes %>>
<span id="el_person_per_en_name" class="control-group">
<input type="text" data-field="x_per_en_name" name="x_per_en_name" id="x_per_en_name" size="30" maxlength="255" placeholder="<%= person.per_en_name.PlaceHolder %>" value="<%= person.per_en_name.EditValue %>"<%= person.per_en_name.EditAttributes %>>
</span>
<%= person.per_en_name.CustomMsg %></td>
	</tr>
<% End If %>
<% If person.per_th_name.Visible Then ' per_th_name %>
	<tr id="r_per_th_name">
		<td><span id="elh_person_per_th_name"><%= person.per_th_name.FldCaption %></span></td>
		<td<%= person.per_th_name.CellAttributes %>>
<span id="el_person_per_th_name" class="control-group">
<input type="text" data-field="x_per_th_name" name="x_per_th_name" id="x_per_th_name" size="30" maxlength="255" placeholder="<%= person.per_th_name.PlaceHolder %>" value="<%= person.per_th_name.EditValue %>"<%= person.per_th_name.EditAttributes %>>
</span>
<%= person.per_th_name.CustomMsg %></td>
	</tr>
<% End If %>
<% If person.per_position.Visible Then ' per_position %>
	<tr id="r_per_position">
		<td><span id="elh_person_per_position"><%= person.per_position.FldCaption %></span></td>
		<td<%= person.per_position.CellAttributes %>>
<span id="el_person_per_position" class="control-group">
<input type="text" data-field="x_per_position" name="x_per_position" id="x_per_position" size="30" maxlength="255" placeholder="<%= person.per_position.PlaceHolder %>" value="<%= person.per_position.EditValue %>"<%= person.per_position.EditAttributes %>>
</span>
<%= person.per_position.CustomMsg %></td>
	</tr>
<% End If %>
<% If person.per_mobile.Visible Then ' per_mobile %>
	<tr id="r_per_mobile">
		<td><span id="elh_person_per_mobile"><%= person.per_mobile.FldCaption %></span></td>
		<td<%= person.per_mobile.CellAttributes %>>
<span id="el_person_per_mobile" class="control-group">
<input type="text" data-field="x_per_mobile" name="x_per_mobile" id="x_per_mobile" size="30" maxlength="255" placeholder="<%= person.per_mobile.PlaceHolder %>" value="<%= person.per_mobile.EditValue %>"<%= person.per_mobile.EditAttributes %>>
</span>
<%= person.per_mobile.CustomMsg %></td>
	</tr>
<% End If %>
<% If person.per_tel.Visible Then ' per_tel %>
	<tr id="r_per_tel">
		<td><span id="elh_person_per_tel"><%= person.per_tel.FldCaption %></span></td>
		<td<%= person.per_tel.CellAttributes %>>
<span id="el_person_per_tel" class="control-group">
<input type="text" data-field="x_per_tel" name="x_per_tel" id="x_per_tel" size="30" maxlength="255" placeholder="<%= person.per_tel.PlaceHolder %>" value="<%= person.per_tel.EditValue %>"<%= person.per_tel.EditAttributes %>>
</span>
<%= person.per_tel.CustomMsg %></td>
	</tr>
<% End If %>
<% If person.per_fax.Visible Then ' per_fax %>
	<tr id="r_per_fax">
		<td><span id="elh_person_per_fax"><%= person.per_fax.FldCaption %></span></td>
		<td<%= person.per_fax.CellAttributes %>>
<span id="el_person_per_fax" class="control-group">
<input type="text" data-field="x_per_fax" name="x_per_fax" id="x_per_fax" size="30" maxlength="255" placeholder="<%= person.per_fax.PlaceHolder %>" value="<%= person.per_fax.EditValue %>"<%= person.per_fax.EditAttributes %>>
</span>
<%= person.per_fax.CustomMsg %></td>
	</tr>
<% End If %>
<% If person.per_email.Visible Then ' per_email %>
	<tr id="r_per_email">
		<td><span id="elh_person_per_email"><%= person.per_email.FldCaption %></span></td>
		<td<%= person.per_email.CellAttributes %>>
<span id="el_person_per_email" class="control-group">
<input type="text" data-field="x_per_email" name="x_per_email" id="x_per_email" size="30" maxlength="255" placeholder="<%= person.per_email.PlaceHolder %>" value="<%= person.per_email.EditValue %>"<%= person.per_email.EditAttributes %>>
</span>
<%= person.per_email.CustomMsg %></td>
	</tr>
<% End If %>
<% If person.per_address.Visible Then ' per_address %>
	<tr id="r_per_address">
		<td><span id="elh_person_per_address"><%= person.per_address.FldCaption %></span></td>
		<td<%= person.per_address.CellAttributes %>>
<span id="el_person_per_address" class="control-group">
<input type="text" data-field="x_per_address" name="x_per_address" id="x_per_address" size="30" maxlength="255" placeholder="<%= person.per_address.PlaceHolder %>" value="<%= person.per_address.EditValue %>"<%= person.per_address.EditAttributes %>>
</span>
<%= person.per_address.CustomMsg %></td>
	</tr>
<% End If %>
<% If person.per_show.Visible Then ' per_show %>
	<tr id="r_per_show">
		<td><span id="elh_person_per_show"><%= person.per_show.FldCaption %></span></td>
		<td<%= person.per_show.CellAttributes %>>
<span id="el_person_per_show" class="control-group">
<input type="text" data-field="x_per_show" name="x_per_show" id="x_per_show" size="30" maxlength="255" placeholder="<%= person.per_show.PlaceHolder %>" value="<%= person.per_show.EditValue %>"<%= person.per_show.EditAttributes %>>
</span>
<%= person.per_show.CustomMsg %></td>
	</tr>
<% End If %>
<% If person.per_create.Visible Then ' per_create %>
	<tr id="r_per_create">
		<td><span id="elh_person_per_create"><%= person.per_create.FldCaption %></span></td>
		<td<%= person.per_create.CellAttributes %>>
<span id="el_person_per_create" class="control-group">
<input type="text" data-field="x_per_create" name="x_per_create" id="x_per_create" size="30" maxlength="255" placeholder="<%= person.per_create.PlaceHolder %>" value="<%= person.per_create.EditValue %>"<%= person.per_create.EditAttributes %>>
</span>
<%= person.per_create.CustomMsg %></td>
	</tr>
<% End If %>
<% If person.per_update.Visible Then ' per_update %>
	<tr id="r_per_update">
		<td><span id="elh_person_per_update"><%= person.per_update.FldCaption %></span></td>
		<td<%= person.per_update.CellAttributes %>>
<span id="el_person_per_update" class="control-group">
<input type="text" data-field="x_per_update" name="x_per_update" id="x_per_update" size="30" maxlength="255" placeholder="<%= person.per_update.PlaceHolder %>" value="<%= person.per_update.EditValue %>"<%= person.per_update.EditAttributes %>>
</span>
<%= person.per_update.CustomMsg %></td>
	</tr>
<% End If %>
<% If person.per_sort.Visible Then ' per_sort %>
	<tr id="r_per_sort">
		<td><span id="elh_person_per_sort"><%= person.per_sort.FldCaption %></span></td>
		<td<%= person.per_sort.CellAttributes %>>
<span id="el_person_per_sort" class="control-group">
<input type="text" data-field="x_per_sort" name="x_per_sort" id="x_per_sort" size="30" placeholder="<%= person.per_sort.PlaceHolder %>" value="<%= person.per_sort.EditValue %>"<%= person.per_sort.EditAttributes %>>
</span>
<%= person.per_sort.CustomMsg %></td>
	</tr>
<% End If %>
<% If person.per_department.Visible Then ' per_department %>
	<tr id="r_per_department">
		<td><span id="elh_person_per_department"><%= person.per_department.FldCaption %></span></td>
		<td<%= person.per_department.CellAttributes %>>
<span id="el_person_per_department" class="control-group">
<input type="text" data-field="x_per_department" name="x_per_department" id="x_per_department" size="30" maxlength="255" placeholder="<%= person.per_department.PlaceHolder %>" value="<%= person.per_department.EditValue %>"<%= person.per_department.EditAttributes %>>
</span>
<%= person.per_department.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
fpersonadd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
person_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set person_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cperson_add

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
		TableName = "person"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "person_add"
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
		EW_PAGE_ID = "add"

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

		' Create form object
		'If Request.ServerVariables("HTTP_CONTENT_TYPE") = "application/x-www-form-urlencoded" Then

			Set ObjForm = New cFormObj

		'Else
		'	Set ObjForm = ew_GetUploadObj()
		'End If

		person.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
			person.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("per_id").Count > 0 Then
				person.per_id.QueryStringValue = Request.QueryString("per_id")
				Call person.SetKey("per_id", person.per_id.CurrentValue) ' Set up key
			Else
				Call person.SetKey("per_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				person.CurrentAction = "C" ' Copy Record
			Else
				person.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				person.CurrentAction = "I" ' Form error, reset action
				person.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case person.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_personlist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				person.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = person.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_personview.asp" Then sReturnUrl = person.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					person.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		person.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call person.ResetAttrs()
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
		person.per_id.CurrentValue = Null
		person.per_id.OldValue = person.per_id.CurrentValue
		person.dept_id.CurrentValue = Null
		person.dept_id.OldValue = person.dept_id.CurrentValue
		person.office_id.CurrentValue = Null
		person.office_id.OldValue = person.office_id.CurrentValue
		person.per_img.CurrentValue = Null
		person.per_img.OldValue = person.per_img.CurrentValue
		person.per_en_name.CurrentValue = Null
		person.per_en_name.OldValue = person.per_en_name.CurrentValue
		person.per_th_name.CurrentValue = Null
		person.per_th_name.OldValue = person.per_th_name.CurrentValue
		person.per_position.CurrentValue = Null
		person.per_position.OldValue = person.per_position.CurrentValue
		person.per_mobile.CurrentValue = Null
		person.per_mobile.OldValue = person.per_mobile.CurrentValue
		person.per_tel.CurrentValue = Null
		person.per_tel.OldValue = person.per_tel.CurrentValue
		person.per_fax.CurrentValue = Null
		person.per_fax.OldValue = person.per_fax.CurrentValue
		person.per_email.CurrentValue = Null
		person.per_email.OldValue = person.per_email.CurrentValue
		person.per_address.CurrentValue = Null
		person.per_address.OldValue = person.per_address.CurrentValue
		person.per_show.CurrentValue = Null
		person.per_show.OldValue = person.per_show.CurrentValue
		person.per_create.CurrentValue = Null
		person.per_create.OldValue = person.per_create.CurrentValue
		person.per_update.CurrentValue = Null
		person.per_update.OldValue = person.per_update.CurrentValue
		person.per_sort.CurrentValue = Null
		person.per_sort.OldValue = person.per_sort.CurrentValue
		person.per_department.CurrentValue = Null
		person.per_department.OldValue = person.per_department.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not person.per_id.FldIsDetailKey Then person.per_id.FormValue = ObjForm.GetValue("x_per_id")
		If Not person.dept_id.FldIsDetailKey Then person.dept_id.FormValue = ObjForm.GetValue("x_dept_id")
		If Not person.office_id.FldIsDetailKey Then person.office_id.FormValue = ObjForm.GetValue("x_office_id")
		If Not person.per_img.FldIsDetailKey Then person.per_img.FormValue = ObjForm.GetValue("x_per_img")
		If Not person.per_en_name.FldIsDetailKey Then person.per_en_name.FormValue = ObjForm.GetValue("x_per_en_name")
		If Not person.per_th_name.FldIsDetailKey Then person.per_th_name.FormValue = ObjForm.GetValue("x_per_th_name")
		If Not person.per_position.FldIsDetailKey Then person.per_position.FormValue = ObjForm.GetValue("x_per_position")
		If Not person.per_mobile.FldIsDetailKey Then person.per_mobile.FormValue = ObjForm.GetValue("x_per_mobile")
		If Not person.per_tel.FldIsDetailKey Then person.per_tel.FormValue = ObjForm.GetValue("x_per_tel")
		If Not person.per_fax.FldIsDetailKey Then person.per_fax.FormValue = ObjForm.GetValue("x_per_fax")
		If Not person.per_email.FldIsDetailKey Then person.per_email.FormValue = ObjForm.GetValue("x_per_email")
		If Not person.per_address.FldIsDetailKey Then person.per_address.FormValue = ObjForm.GetValue("x_per_address")
		If Not person.per_show.FldIsDetailKey Then person.per_show.FormValue = ObjForm.GetValue("x_per_show")
		If Not person.per_create.FldIsDetailKey Then person.per_create.FormValue = ObjForm.GetValue("x_per_create")
		If Not person.per_update.FldIsDetailKey Then person.per_update.FormValue = ObjForm.GetValue("x_per_update")
		If Not person.per_sort.FldIsDetailKey Then person.per_sort.FormValue = ObjForm.GetValue("x_per_sort")
		If Not person.per_department.FldIsDetailKey Then person.per_department.FormValue = ObjForm.GetValue("x_per_department")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		person.per_id.CurrentValue = person.per_id.FormValue
		person.dept_id.CurrentValue = person.dept_id.FormValue
		person.office_id.CurrentValue = person.office_id.FormValue
		person.per_img.CurrentValue = person.per_img.FormValue
		person.per_en_name.CurrentValue = person.per_en_name.FormValue
		person.per_th_name.CurrentValue = person.per_th_name.FormValue
		person.per_position.CurrentValue = person.per_position.FormValue
		person.per_mobile.CurrentValue = person.per_mobile.FormValue
		person.per_tel.CurrentValue = person.per_tel.FormValue
		person.per_fax.CurrentValue = person.per_fax.FormValue
		person.per_email.CurrentValue = person.per_email.FormValue
		person.per_address.CurrentValue = person.per_address.FormValue
		person.per_show.CurrentValue = person.per_show.FormValue
		person.per_create.CurrentValue = person.per_create.FormValue
		person.per_update.CurrentValue = person.per_update.FormValue
		person.per_sort.CurrentValue = person.per_sort.FormValue
		person.per_department.CurrentValue = person.per_department.FormValue
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

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If person.GetKey("per_id")&"" <> "" Then
			person.per_id.CurrentValue = person.GetKey("per_id") ' per_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			person.CurrentFilter = person.KeyFilter
			Dim sSql
			sSql = person.SQL
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

		' ---------
		'  Add Row
		' ---------

		ElseIf person.RowType = EW_ROWTYPE_ADD Then ' Add row

			' per_id
			person.per_id.EditCustomAttributes = ""
			person.per_id.EditValue = ew_HtmlEncode(person.per_id.CurrentValue)
			person.per_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.per_id.FldCaption))

			' dept_id
			person.dept_id.EditCustomAttributes = ""
			person.dept_id.EditValue = ew_HtmlEncode(person.dept_id.CurrentValue)
			person.dept_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.dept_id.FldCaption))

			' office_id
			person.office_id.EditCustomAttributes = ""
			person.office_id.EditValue = ew_HtmlEncode(person.office_id.CurrentValue)
			person.office_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.office_id.FldCaption))

			' per_img
			person.per_img.EditCustomAttributes = ""
			person.per_img.EditValue = ew_HtmlEncode(person.per_img.CurrentValue)
			person.per_img.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.per_img.FldCaption))

			' per_en_name
			person.per_en_name.EditCustomAttributes = ""
			person.per_en_name.EditValue = ew_HtmlEncode(person.per_en_name.CurrentValue)
			person.per_en_name.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.per_en_name.FldCaption))

			' per_th_name
			person.per_th_name.EditCustomAttributes = ""
			person.per_th_name.EditValue = ew_HtmlEncode(person.per_th_name.CurrentValue)
			person.per_th_name.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.per_th_name.FldCaption))

			' per_position
			person.per_position.EditCustomAttributes = ""
			person.per_position.EditValue = ew_HtmlEncode(person.per_position.CurrentValue)
			person.per_position.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.per_position.FldCaption))

			' per_mobile
			person.per_mobile.EditCustomAttributes = ""
			person.per_mobile.EditValue = ew_HtmlEncode(person.per_mobile.CurrentValue)
			person.per_mobile.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.per_mobile.FldCaption))

			' per_tel
			person.per_tel.EditCustomAttributes = ""
			person.per_tel.EditValue = ew_HtmlEncode(person.per_tel.CurrentValue)
			person.per_tel.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.per_tel.FldCaption))

			' per_fax
			person.per_fax.EditCustomAttributes = ""
			person.per_fax.EditValue = ew_HtmlEncode(person.per_fax.CurrentValue)
			person.per_fax.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.per_fax.FldCaption))

			' per_email
			person.per_email.EditCustomAttributes = ""
			person.per_email.EditValue = ew_HtmlEncode(person.per_email.CurrentValue)
			person.per_email.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.per_email.FldCaption))

			' per_address
			person.per_address.EditCustomAttributes = ""
			person.per_address.EditValue = ew_HtmlEncode(person.per_address.CurrentValue)
			person.per_address.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.per_address.FldCaption))

			' per_show
			person.per_show.EditCustomAttributes = ""
			person.per_show.EditValue = ew_HtmlEncode(person.per_show.CurrentValue)
			person.per_show.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.per_show.FldCaption))

			' per_create
			person.per_create.EditCustomAttributes = ""
			person.per_create.EditValue = ew_HtmlEncode(person.per_create.CurrentValue)
			person.per_create.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.per_create.FldCaption))

			' per_update
			person.per_update.EditCustomAttributes = ""
			person.per_update.EditValue = ew_HtmlEncode(person.per_update.CurrentValue)
			person.per_update.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.per_update.FldCaption))

			' per_sort
			person.per_sort.EditCustomAttributes = ""
			person.per_sort.EditValue = ew_HtmlEncode(person.per_sort.CurrentValue)
			person.per_sort.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.per_sort.FldCaption))

			' per_department
			person.per_department.EditCustomAttributes = ""
			person.per_department.EditValue = ew_HtmlEncode(person.per_department.CurrentValue)
			person.per_department.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(person.per_department.FldCaption))

			' Edit refer script
			' per_id

			person.per_id.HrefValue = ""

			' dept_id
			person.dept_id.HrefValue = ""

			' office_id
			person.office_id.HrefValue = ""

			' per_img
			person.per_img.HrefValue = ""

			' per_en_name
			person.per_en_name.HrefValue = ""

			' per_th_name
			person.per_th_name.HrefValue = ""

			' per_position
			person.per_position.HrefValue = ""

			' per_mobile
			person.per_mobile.HrefValue = ""

			' per_tel
			person.per_tel.HrefValue = ""

			' per_fax
			person.per_fax.HrefValue = ""

			' per_email
			person.per_email.HrefValue = ""

			' per_address
			person.per_address.HrefValue = ""

			' per_show
			person.per_show.HrefValue = ""

			' per_create
			person.per_create.HrefValue = ""

			' per_update
			person.per_update.HrefValue = ""

			' per_sort
			person.per_sort.HrefValue = ""

			' per_department
			person.per_department.HrefValue = ""
		End If
		If person.RowType = EW_ROWTYPE_ADD Or person.RowType = EW_ROWTYPE_EDIT Or person.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call person.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If person.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call person.Row_Rendered()
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
		If Not ew_CheckInteger(person.per_id.FormValue) Then
			Call ew_AddMessage(gsFormError, person.per_id.FldErrMsg)
		End If
		If Not ew_CheckInteger(person.dept_id.FormValue) Then
			Call ew_AddMessage(gsFormError, person.dept_id.FldErrMsg)
		End If
		If Not ew_CheckInteger(person.office_id.FormValue) Then
			Call ew_AddMessage(gsFormError, person.office_id.FldErrMsg)
		End If
		If Not ew_CheckInteger(person.per_sort.FormValue) Then
			Call ew_AddMessage(gsFormError, person.per_sort.FldErrMsg)
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
		If person.per_id.CurrentValue <> "" Then ' Check field with unique index
			sFilter = "([per_id] = " & ew_AdjustSql(person.per_id.CurrentValue) & ")"
			Set RsChk = person.LoadRs(sFilter)
			If Not (RsChk Is Nothing) Then
				sIdxErrMsg = Replace(Language.Phrase("DupIndex"), "%f", person.per_id.FldCaption)
				sIdxErrMsg = Replace(sIdxErrMsg, "%v", person.per_id.CurrentValue)
				FailureMessage = sIdxErrMsg
				RsChk.Close
				Set RsChk = Nothing
				AddRow = False
				Exit Function
			End If
		End If

		' Load db values from rsold
		If Not IsNull(RsOld) Then
			Call LoadDbValues(RsOld)
		End If

		' Add new record
		sFilter = "(0 = 1)"
		person.CurrentFilter = sFilter
		sSql = person.SQL
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

		' Field per_id
		Call person.per_id.SetDbValue(Rs, person.per_id.CurrentValue, Null, False)

		' Field dept_id
		Call person.dept_id.SetDbValue(Rs, person.dept_id.CurrentValue, Null, False)

		' Field office_id
		Call person.office_id.SetDbValue(Rs, person.office_id.CurrentValue, Null, False)

		' Field per_img
		Call person.per_img.SetDbValue(Rs, person.per_img.CurrentValue, Null, False)

		' Field per_en_name
		Call person.per_en_name.SetDbValue(Rs, person.per_en_name.CurrentValue, Null, False)

		' Field per_th_name
		Call person.per_th_name.SetDbValue(Rs, person.per_th_name.CurrentValue, Null, False)

		' Field per_position
		Call person.per_position.SetDbValue(Rs, person.per_position.CurrentValue, Null, False)

		' Field per_mobile
		Call person.per_mobile.SetDbValue(Rs, person.per_mobile.CurrentValue, Null, False)

		' Field per_tel
		Call person.per_tel.SetDbValue(Rs, person.per_tel.CurrentValue, Null, False)

		' Field per_fax
		Call person.per_fax.SetDbValue(Rs, person.per_fax.CurrentValue, Null, False)

		' Field per_email
		Call person.per_email.SetDbValue(Rs, person.per_email.CurrentValue, Null, False)

		' Field per_address
		Call person.per_address.SetDbValue(Rs, person.per_address.CurrentValue, Null, False)

		' Field per_show
		Call person.per_show.SetDbValue(Rs, person.per_show.CurrentValue, Null, False)

		' Field per_create
		Call person.per_create.SetDbValue(Rs, person.per_create.CurrentValue, Null, False)

		' Field per_update
		Call person.per_update.SetDbValue(Rs, person.per_update.CurrentValue, Null, False)

		' Field per_sort
		Call person.per_sort.SetDbValue(Rs, person.per_sort.CurrentValue, Null, False)

		' Field per_department
		Call person.per_department.SetDbValue(Rs, person.per_department.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = person.Row_Inserting(RsOld, Rs)

		' Check if key value entered
		If bInsertRow And person.ValidateKey And person.per_id.CurrentValue = "" And person.per_id.SessionValue = "" Then
			FailureMessage = Language.Phrase("InvalidKeyValue")
			bInsertRow = False
		End If

		' Check for duplicate key
		Dim sKeyErrMsg
		If bInsertRow And person.ValidateKey Then
			sFilter = person.KeyFilter
			Set RsChk = person.LoadRs(sFilter)
			If Not (RsChk Is Nothing) Then
				sKeyErrMsg = Replace(Language.Phrase("DupKey"), "%f", sFilter)
				FailureMessage = sKeyErrMsg
				RsChk.Close
				Set RsChk = Nothing
				bInsertRow = False
			End If
		End If
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
			ElseIf person.CancelMessage <> "" Then
				FailureMessage = person.CancelMessage
				person.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("InsertCancelled")
			End If
			AddRow = False
		End If
		Rs.Close
		Set Rs = Nothing
		If AddRow Then
		End If
		If AddRow Then

			' Call Row Inserted event
			Call person.Row_Inserted(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", person.TableVar, "pom_personlist.asp", person.TableVar, True)
		PageId = ew_IIf(person.CurrentAction = "C", "Copy", "Add")
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
