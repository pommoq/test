<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_research_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim research_th_edit
Set research_th_edit = New cresearch_th_edit
Set Page = research_th_edit

' Page init processing
research_th_edit.Page_Init()

' Page main processing
research_th_edit.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
research_th_edit.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var research_th_edit = new ew_Page("research_th_edit");
research_th_edit.PageID = "edit"; // Page ID
var EW_PAGE_ID = research_th_edit.PageID; // For backward compatibility
// Form object
var fresearch_thedit = new ew_Form("fresearch_thedit");
// Validate form
fresearch_thedit.Validate = function() {
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
			elm = this.GetElements("x" + infix + "_rsh_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(research_th.rsh_id.FldErrMsg) %>");
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
fresearch_thedit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fresearch_thedit.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fresearch_thedit.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If research_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% research_th_edit.ShowPageHeader() %>
<% research_th_edit.ShowMessage %>
<form name="fresearch_thedit" id="fresearch_thedit" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="a_table" id="a_table" value="research_th">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table class="ewGrid"><tr><td>
<table id="tbl_research_thedit" class="table table-bordered table-striped">
<% If research_th.rsh_id.Visible Then ' rsh_id %>
	<tr id="r_rsh_id">
		<td><span id="elh_research_th_rsh_id"><%= research_th.rsh_id.FldCaption %></span></td>
		<td<%= research_th.rsh_id.CellAttributes %>>
<span id="el_research_th_rsh_id" class="control-group">
<span<%= research_th.rsh_id.ViewAttributes %>>
<%= research_th.rsh_id.EditValue %>
</span>
</span>
<input type="hidden" data-field="x_rsh_id" name="x_rsh_id" id="x_rsh_id" value="<%= Server.HTMLEncode(research_th.rsh_id.CurrentValue&"") %>">
<%= research_th.rsh_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_th.rsh_img.Visible Then ' rsh_img %>
	<tr id="r_rsh_img">
		<td><span id="elh_research_th_rsh_img"><%= research_th.rsh_img.FldCaption %></span></td>
		<td<%= research_th.rsh_img.CellAttributes %>>
<span id="el_research_th_rsh_img" class="control-group">
<input type="text" data-field="x_rsh_img" name="x_rsh_img" id="x_rsh_img" size="30" maxlength="255" placeholder="<%= research_th.rsh_img.PlaceHolder %>" value="<%= research_th.rsh_img.EditValue %>"<%= research_th.rsh_img.EditAttributes %>>
</span>
<%= research_th.rsh_img.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_th.rsh_date.Visible Then ' rsh_date %>
	<tr id="r_rsh_date">
		<td><span id="elh_research_th_rsh_date"><%= research_th.rsh_date.FldCaption %></span></td>
		<td<%= research_th.rsh_date.CellAttributes %>>
<span id="el_research_th_rsh_date" class="control-group">
<input type="text" data-field="x_rsh_date" name="x_rsh_date" id="x_rsh_date" placeholder="<%= research_th.rsh_date.PlaceHolder %>" value="<%= research_th.rsh_date.EditValue %>"<%= research_th.rsh_date.EditAttributes %>>
</span>
<%= research_th.rsh_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_th.rsh_pdf.Visible Then ' rsh_pdf %>
	<tr id="r_rsh_pdf">
		<td><span id="elh_research_th_rsh_pdf"><%= research_th.rsh_pdf.FldCaption %></span></td>
		<td<%= research_th.rsh_pdf.CellAttributes %>>
<span id="el_research_th_rsh_pdf" class="control-group">
<input type="text" data-field="x_rsh_pdf" name="x_rsh_pdf" id="x_rsh_pdf" size="30" maxlength="255" placeholder="<%= research_th.rsh_pdf.PlaceHolder %>" value="<%= research_th.rsh_pdf.EditValue %>"<%= research_th.rsh_pdf.EditAttributes %>>
</span>
<%= research_th.rsh_pdf.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_th.rsh_category.Visible Then ' rsh_category %>
	<tr id="r_rsh_category">
		<td><span id="elh_research_th_rsh_category"><%= research_th.rsh_category.FldCaption %></span></td>
		<td<%= research_th.rsh_category.CellAttributes %>>
<span id="el_research_th_rsh_category" class="control-group">
<input type="text" data-field="x_rsh_category" name="x_rsh_category" id="x_rsh_category" size="30" maxlength="255" placeholder="<%= research_th.rsh_category.PlaceHolder %>" value="<%= research_th.rsh_category.EditValue %>"<%= research_th.rsh_category.EditAttributes %>>
</span>
<%= research_th.rsh_category.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_th.rsh_subject.Visible Then ' rsh_subject %>
	<tr id="r_rsh_subject">
		<td><span id="elh_research_th_rsh_subject"><%= research_th.rsh_subject.FldCaption %></span></td>
		<td<%= research_th.rsh_subject.CellAttributes %>>
<span id="el_research_th_rsh_subject" class="control-group">
<input type="text" data-field="x_rsh_subject" name="x_rsh_subject" id="x_rsh_subject" size="30" maxlength="255" placeholder="<%= research_th.rsh_subject.PlaceHolder %>" value="<%= research_th.rsh_subject.EditValue %>"<%= research_th.rsh_subject.EditAttributes %>>
</span>
<%= research_th.rsh_subject.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_th.rsh_subject_th.Visible Then ' rsh_subject_th %>
	<tr id="r_rsh_subject_th">
		<td><span id="elh_research_th_rsh_subject_th"><%= research_th.rsh_subject_th.FldCaption %></span></td>
		<td<%= research_th.rsh_subject_th.CellAttributes %>>
<span id="el_research_th_rsh_subject_th" class="control-group">
<input type="text" data-field="x_rsh_subject_th" name="x_rsh_subject_th" id="x_rsh_subject_th" size="30" maxlength="255" placeholder="<%= research_th.rsh_subject_th.PlaceHolder %>" value="<%= research_th.rsh_subject_th.EditValue %>"<%= research_th.rsh_subject_th.EditAttributes %>>
</span>
<%= research_th.rsh_subject_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_th.rsh_intro.Visible Then ' rsh_intro %>
	<tr id="r_rsh_intro">
		<td><span id="elh_research_th_rsh_intro"><%= research_th.rsh_intro.FldCaption %></span></td>
		<td<%= research_th.rsh_intro.CellAttributes %>>
<span id="el_research_th_rsh_intro" class="control-group">
<textarea data-field="x_rsh_intro" name="x_rsh_intro" id="x_rsh_intro" cols="35" rows="4" placeholder="<%= research_th.rsh_intro.PlaceHolder %>"<%= research_th.rsh_intro.EditAttributes %>><%= research_th.rsh_intro.EditValue %></textarea>
</span>
<%= research_th.rsh_intro.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_th.rsh_intro_th.Visible Then ' rsh_intro_th %>
	<tr id="r_rsh_intro_th">
		<td><span id="elh_research_th_rsh_intro_th"><%= research_th.rsh_intro_th.FldCaption %></span></td>
		<td<%= research_th.rsh_intro_th.CellAttributes %>>
<span id="el_research_th_rsh_intro_th" class="control-group">
<input type="text" data-field="x_rsh_intro_th" name="x_rsh_intro_th" id="x_rsh_intro_th" size="30" maxlength="255" placeholder="<%= research_th.rsh_intro_th.PlaceHolder %>" value="<%= research_th.rsh_intro_th.EditValue %>"<%= research_th.rsh_intro_th.EditAttributes %>>
</span>
<%= research_th.rsh_intro_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_th.rsh_content.Visible Then ' rsh_content %>
	<tr id="r_rsh_content">
		<td><span id="elh_research_th_rsh_content"><%= research_th.rsh_content.FldCaption %></span></td>
		<td<%= research_th.rsh_content.CellAttributes %>>
<span id="el_research_th_rsh_content" class="control-group">
<textarea data-field="x_rsh_content" name="x_rsh_content" id="x_rsh_content" cols="35" rows="4" placeholder="<%= research_th.rsh_content.PlaceHolder %>"<%= research_th.rsh_content.EditAttributes %>><%= research_th.rsh_content.EditValue %></textarea>
</span>
<%= research_th.rsh_content.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_th.rsh_content_th.Visible Then ' rsh_content_th %>
	<tr id="r_rsh_content_th">
		<td><span id="elh_research_th_rsh_content_th"><%= research_th.rsh_content_th.FldCaption %></span></td>
		<td<%= research_th.rsh_content_th.CellAttributes %>>
<span id="el_research_th_rsh_content_th" class="control-group">
<textarea data-field="x_rsh_content_th" name="x_rsh_content_th" id="x_rsh_content_th" cols="35" rows="4" placeholder="<%= research_th.rsh_content_th.PlaceHolder %>"<%= research_th.rsh_content_th.EditAttributes %>><%= research_th.rsh_content_th.EditValue %></textarea>
</span>
<%= research_th.rsh_content_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_th.rsh_show.Visible Then ' rsh_show %>
	<tr id="r_rsh_show">
		<td><span id="elh_research_th_rsh_show"><%= research_th.rsh_show.FldCaption %></span></td>
		<td<%= research_th.rsh_show.CellAttributes %>>
<span id="el_research_th_rsh_show" class="control-group">
<input type="text" data-field="x_rsh_show" name="x_rsh_show" id="x_rsh_show" size="30" maxlength="255" placeholder="<%= research_th.rsh_show.PlaceHolder %>" value="<%= research_th.rsh_show.EditValue %>"<%= research_th.rsh_show.EditAttributes %>>
</span>
<%= research_th.rsh_show.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_th.rsh_show_home.Visible Then ' rsh_show_home %>
	<tr id="r_rsh_show_home">
		<td><span id="elh_research_th_rsh_show_home"><%= research_th.rsh_show_home.FldCaption %></span></td>
		<td<%= research_th.rsh_show_home.CellAttributes %>>
<span id="el_research_th_rsh_show_home" class="control-group">
<input type="text" data-field="x_rsh_show_home" name="x_rsh_show_home" id="x_rsh_show_home" size="30" maxlength="255" placeholder="<%= research_th.rsh_show_home.PlaceHolder %>" value="<%= research_th.rsh_show_home.EditValue %>"<%= research_th.rsh_show_home.EditAttributes %>>
</span>
<%= research_th.rsh_show_home.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_th.rsh_create.Visible Then ' rsh_create %>
	<tr id="r_rsh_create">
		<td><span id="elh_research_th_rsh_create"><%= research_th.rsh_create.FldCaption %></span></td>
		<td<%= research_th.rsh_create.CellAttributes %>>
<span id="el_research_th_rsh_create" class="control-group">
<input type="text" data-field="x_rsh_create" name="x_rsh_create" id="x_rsh_create" size="30" maxlength="255" placeholder="<%= research_th.rsh_create.PlaceHolder %>" value="<%= research_th.rsh_create.EditValue %>"<%= research_th.rsh_create.EditAttributes %>>
</span>
<%= research_th.rsh_create.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_th.rsh_update.Visible Then ' rsh_update %>
	<tr id="r_rsh_update">
		<td><span id="elh_research_th_rsh_update"><%= research_th.rsh_update.FldCaption %></span></td>
		<td<%= research_th.rsh_update.CellAttributes %>>
<span id="el_research_th_rsh_update" class="control-group">
<input type="text" data-field="x_rsh_update" name="x_rsh_update" id="x_rsh_update" size="30" maxlength="255" placeholder="<%= research_th.rsh_update.PlaceHolder %>" value="<%= research_th.rsh_update.EditValue %>"<%= research_th.rsh_update.EditAttributes %>>
</span>
<%= research_th.rsh_update.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("EditBtn") %></button>
</form>
<script type="text/javascript">
fresearch_thedit.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
research_th_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set research_th_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cresearch_th_edit

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
		TableName = "research_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "research_th_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If research_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & research_th.TableVar & "&" ' add page token
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
		If research_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (research_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (research_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(research_th) Then Set research_th = New cresearch_th
		Set Table = research_th

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "research_th"

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

		research_th.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set research_th = Nothing
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
		If Request.QueryString("rsh_id").Count > 0 Then
			research_th.rsh_id.QueryStringValue = Request.QueryString("rsh_id")
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Process form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			research_th.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values
		Else
			research_th.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If research_th.rsh_id.CurrentValue = "" Then Call Page_Terminate("pom_research_thlist.asp") ' Invalid key, return to list

		' Validate form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			If Not ValidateForm() Then
				research_th.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				research_th.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		End If
		Select Case research_th.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_research_thlist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				research_th.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					sReturnUrl = research_th.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					research_th.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		research_th.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call research_th.ResetAttrs()
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
				research_th.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					research_th.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = research_th.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			research_th.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			research_th.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			research_th.StartRecordNumber = StartRec
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
		If Not research_th.rsh_id.FldIsDetailKey Then research_th.rsh_id.FormValue = ObjForm.GetValue("x_rsh_id")
		If Not research_th.rsh_img.FldIsDetailKey Then research_th.rsh_img.FormValue = ObjForm.GetValue("x_rsh_img")
		If Not research_th.rsh_date.FldIsDetailKey Then research_th.rsh_date.FormValue = ObjForm.GetValue("x_rsh_date")
		If Not research_th.rsh_date.FldIsDetailKey Then research_th.rsh_date.CurrentValue = ew_UnFormatDateTime(research_th.rsh_date.CurrentValue, 8)
		If Not research_th.rsh_pdf.FldIsDetailKey Then research_th.rsh_pdf.FormValue = ObjForm.GetValue("x_rsh_pdf")
		If Not research_th.rsh_category.FldIsDetailKey Then research_th.rsh_category.FormValue = ObjForm.GetValue("x_rsh_category")
		If Not research_th.rsh_subject.FldIsDetailKey Then research_th.rsh_subject.FormValue = ObjForm.GetValue("x_rsh_subject")
		If Not research_th.rsh_subject_th.FldIsDetailKey Then research_th.rsh_subject_th.FormValue = ObjForm.GetValue("x_rsh_subject_th")
		If Not research_th.rsh_intro.FldIsDetailKey Then research_th.rsh_intro.FormValue = ObjForm.GetValue("x_rsh_intro")
		If Not research_th.rsh_intro_th.FldIsDetailKey Then research_th.rsh_intro_th.FormValue = ObjForm.GetValue("x_rsh_intro_th")
		If Not research_th.rsh_content.FldIsDetailKey Then research_th.rsh_content.FormValue = ObjForm.GetValue("x_rsh_content")
		If Not research_th.rsh_content_th.FldIsDetailKey Then research_th.rsh_content_th.FormValue = ObjForm.GetValue("x_rsh_content_th")
		If Not research_th.rsh_show.FldIsDetailKey Then research_th.rsh_show.FormValue = ObjForm.GetValue("x_rsh_show")
		If Not research_th.rsh_show_home.FldIsDetailKey Then research_th.rsh_show_home.FormValue = ObjForm.GetValue("x_rsh_show_home")
		If Not research_th.rsh_create.FldIsDetailKey Then research_th.rsh_create.FormValue = ObjForm.GetValue("x_rsh_create")
		If Not research_th.rsh_update.FldIsDetailKey Then research_th.rsh_update.FormValue = ObjForm.GetValue("x_rsh_update")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		research_th.rsh_id.CurrentValue = research_th.rsh_id.FormValue
		research_th.rsh_img.CurrentValue = research_th.rsh_img.FormValue
		research_th.rsh_date.CurrentValue = research_th.rsh_date.FormValue
		research_th.rsh_date.CurrentValue = ew_UnFormatDateTime(research_th.rsh_date.CurrentValue, 8)
		research_th.rsh_pdf.CurrentValue = research_th.rsh_pdf.FormValue
		research_th.rsh_category.CurrentValue = research_th.rsh_category.FormValue
		research_th.rsh_subject.CurrentValue = research_th.rsh_subject.FormValue
		research_th.rsh_subject_th.CurrentValue = research_th.rsh_subject_th.FormValue
		research_th.rsh_intro.CurrentValue = research_th.rsh_intro.FormValue
		research_th.rsh_intro_th.CurrentValue = research_th.rsh_intro_th.FormValue
		research_th.rsh_content.CurrentValue = research_th.rsh_content.FormValue
		research_th.rsh_content_th.CurrentValue = research_th.rsh_content_th.FormValue
		research_th.rsh_show.CurrentValue = research_th.rsh_show.FormValue
		research_th.rsh_show_home.CurrentValue = research_th.rsh_show_home.FormValue
		research_th.rsh_create.CurrentValue = research_th.rsh_create.FormValue
		research_th.rsh_update.CurrentValue = research_th.rsh_update.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = research_th.KeyFilter

		' Call Row Selecting event
		Call research_th.Row_Selecting(sFilter)

		' Load sql based on filter
		research_th.CurrentFilter = sFilter
		sSql = research_th.SQL
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
		Call research_th.Row_Selected(RsRow)
		research_th.rsh_id.DbValue = RsRow("rsh_id")
		research_th.rsh_img.DbValue = RsRow("rsh_img")
		research_th.rsh_date.DbValue = RsRow("rsh_date")
		research_th.rsh_pdf.DbValue = RsRow("rsh_pdf")
		research_th.rsh_category.DbValue = RsRow("rsh_category")
		research_th.rsh_subject.DbValue = RsRow("rsh_subject")
		research_th.rsh_subject_th.DbValue = RsRow("rsh_subject_th")
		research_th.rsh_intro.DbValue = RsRow("rsh_intro")
		research_th.rsh_intro_th.DbValue = RsRow("rsh_intro_th")
		research_th.rsh_content.DbValue = RsRow("rsh_content")
		research_th.rsh_content_th.DbValue = RsRow("rsh_content_th")
		research_th.rsh_show.DbValue = RsRow("rsh_show")
		research_th.rsh_show_home.DbValue = RsRow("rsh_show_home")
		research_th.rsh_create.DbValue = RsRow("rsh_create")
		research_th.rsh_update.DbValue = RsRow("rsh_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		research_th.rsh_id.m_DbValue = Rs("rsh_id")
		research_th.rsh_img.m_DbValue = Rs("rsh_img")
		research_th.rsh_date.m_DbValue = Rs("rsh_date")
		research_th.rsh_pdf.m_DbValue = Rs("rsh_pdf")
		research_th.rsh_category.m_DbValue = Rs("rsh_category")
		research_th.rsh_subject.m_DbValue = Rs("rsh_subject")
		research_th.rsh_subject_th.m_DbValue = Rs("rsh_subject_th")
		research_th.rsh_intro.m_DbValue = Rs("rsh_intro")
		research_th.rsh_intro_th.m_DbValue = Rs("rsh_intro_th")
		research_th.rsh_content.m_DbValue = Rs("rsh_content")
		research_th.rsh_content_th.m_DbValue = Rs("rsh_content_th")
		research_th.rsh_show.m_DbValue = Rs("rsh_show")
		research_th.rsh_show_home.m_DbValue = Rs("rsh_show_home")
		research_th.rsh_create.m_DbValue = Rs("rsh_create")
		research_th.rsh_update.m_DbValue = Rs("rsh_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call research_th.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' rsh_id
		' rsh_img
		' rsh_date
		' rsh_pdf
		' rsh_category
		' rsh_subject
		' rsh_subject_th
		' rsh_intro
		' rsh_intro_th
		' rsh_content
		' rsh_content_th
		' rsh_show
		' rsh_show_home
		' rsh_create
		' rsh_update
		' -----------
		'  View  Row
		' -----------

		If research_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' rsh_id
			research_th.rsh_id.ViewValue = research_th.rsh_id.CurrentValue
			research_th.rsh_id.ViewCustomAttributes = ""

			' rsh_img
			research_th.rsh_img.ViewValue = research_th.rsh_img.CurrentValue
			research_th.rsh_img.ViewCustomAttributes = ""

			' rsh_date
			research_th.rsh_date.ViewValue = research_th.rsh_date.CurrentValue
			research_th.rsh_date.ViewCustomAttributes = ""

			' rsh_pdf
			research_th.rsh_pdf.ViewValue = research_th.rsh_pdf.CurrentValue
			research_th.rsh_pdf.ViewCustomAttributes = ""

			' rsh_category
			research_th.rsh_category.ViewValue = research_th.rsh_category.CurrentValue
			research_th.rsh_category.ViewCustomAttributes = ""

			' rsh_subject
			research_th.rsh_subject.ViewValue = research_th.rsh_subject.CurrentValue
			research_th.rsh_subject.ViewCustomAttributes = ""

			' rsh_subject_th
			research_th.rsh_subject_th.ViewValue = research_th.rsh_subject_th.CurrentValue
			research_th.rsh_subject_th.ViewCustomAttributes = ""

			' rsh_intro
			research_th.rsh_intro.ViewValue = research_th.rsh_intro.CurrentValue
			research_th.rsh_intro.ViewCustomAttributes = ""

			' rsh_intro_th
			research_th.rsh_intro_th.ViewValue = research_th.rsh_intro_th.CurrentValue
			research_th.rsh_intro_th.ViewCustomAttributes = ""

			' rsh_content
			research_th.rsh_content.ViewValue = research_th.rsh_content.CurrentValue
			research_th.rsh_content.ViewCustomAttributes = ""

			' rsh_content_th
			research_th.rsh_content_th.ViewValue = research_th.rsh_content_th.CurrentValue
			research_th.rsh_content_th.ViewCustomAttributes = ""

			' rsh_show
			research_th.rsh_show.ViewValue = research_th.rsh_show.CurrentValue
			research_th.rsh_show.ViewCustomAttributes = ""

			' rsh_show_home
			research_th.rsh_show_home.ViewValue = research_th.rsh_show_home.CurrentValue
			research_th.rsh_show_home.ViewCustomAttributes = ""

			' rsh_create
			research_th.rsh_create.ViewValue = research_th.rsh_create.CurrentValue
			research_th.rsh_create.ViewCustomAttributes = ""

			' rsh_update
			research_th.rsh_update.ViewValue = research_th.rsh_update.CurrentValue
			research_th.rsh_update.ViewCustomAttributes = ""

			' View refer script
			' rsh_id

			research_th.rsh_id.LinkCustomAttributes = ""
			research_th.rsh_id.HrefValue = ""
			research_th.rsh_id.TooltipValue = ""

			' rsh_img
			research_th.rsh_img.LinkCustomAttributes = ""
			research_th.rsh_img.HrefValue = ""
			research_th.rsh_img.TooltipValue = ""

			' rsh_date
			research_th.rsh_date.LinkCustomAttributes = ""
			research_th.rsh_date.HrefValue = ""
			research_th.rsh_date.TooltipValue = ""

			' rsh_pdf
			research_th.rsh_pdf.LinkCustomAttributes = ""
			research_th.rsh_pdf.HrefValue = ""
			research_th.rsh_pdf.TooltipValue = ""

			' rsh_category
			research_th.rsh_category.LinkCustomAttributes = ""
			research_th.rsh_category.HrefValue = ""
			research_th.rsh_category.TooltipValue = ""

			' rsh_subject
			research_th.rsh_subject.LinkCustomAttributes = ""
			research_th.rsh_subject.HrefValue = ""
			research_th.rsh_subject.TooltipValue = ""

			' rsh_subject_th
			research_th.rsh_subject_th.LinkCustomAttributes = ""
			research_th.rsh_subject_th.HrefValue = ""
			research_th.rsh_subject_th.TooltipValue = ""

			' rsh_intro
			research_th.rsh_intro.LinkCustomAttributes = ""
			research_th.rsh_intro.HrefValue = ""
			research_th.rsh_intro.TooltipValue = ""

			' rsh_intro_th
			research_th.rsh_intro_th.LinkCustomAttributes = ""
			research_th.rsh_intro_th.HrefValue = ""
			research_th.rsh_intro_th.TooltipValue = ""

			' rsh_content
			research_th.rsh_content.LinkCustomAttributes = ""
			research_th.rsh_content.HrefValue = ""
			research_th.rsh_content.TooltipValue = ""

			' rsh_content_th
			research_th.rsh_content_th.LinkCustomAttributes = ""
			research_th.rsh_content_th.HrefValue = ""
			research_th.rsh_content_th.TooltipValue = ""

			' rsh_show
			research_th.rsh_show.LinkCustomAttributes = ""
			research_th.rsh_show.HrefValue = ""
			research_th.rsh_show.TooltipValue = ""

			' rsh_show_home
			research_th.rsh_show_home.LinkCustomAttributes = ""
			research_th.rsh_show_home.HrefValue = ""
			research_th.rsh_show_home.TooltipValue = ""

			' rsh_create
			research_th.rsh_create.LinkCustomAttributes = ""
			research_th.rsh_create.HrefValue = ""
			research_th.rsh_create.TooltipValue = ""

			' rsh_update
			research_th.rsh_update.LinkCustomAttributes = ""
			research_th.rsh_update.HrefValue = ""
			research_th.rsh_update.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf research_th.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' rsh_id
			research_th.rsh_id.EditCustomAttributes = ""
			research_th.rsh_id.EditValue = research_th.rsh_id.CurrentValue
			research_th.rsh_id.ViewCustomAttributes = ""

			' rsh_img
			research_th.rsh_img.EditCustomAttributes = ""
			research_th.rsh_img.EditValue = ew_HtmlEncode(research_th.rsh_img.CurrentValue)
			research_th.rsh_img.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_th.rsh_img.FldCaption))

			' rsh_date
			research_th.rsh_date.EditCustomAttributes = ""
			research_th.rsh_date.EditValue = ew_HtmlEncode(research_th.rsh_date.CurrentValue)
			research_th.rsh_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_th.rsh_date.FldCaption))

			' rsh_pdf
			research_th.rsh_pdf.EditCustomAttributes = ""
			research_th.rsh_pdf.EditValue = ew_HtmlEncode(research_th.rsh_pdf.CurrentValue)
			research_th.rsh_pdf.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_th.rsh_pdf.FldCaption))

			' rsh_category
			research_th.rsh_category.EditCustomAttributes = ""
			research_th.rsh_category.EditValue = ew_HtmlEncode(research_th.rsh_category.CurrentValue)
			research_th.rsh_category.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_th.rsh_category.FldCaption))

			' rsh_subject
			research_th.rsh_subject.EditCustomAttributes = ""
			research_th.rsh_subject.EditValue = ew_HtmlEncode(research_th.rsh_subject.CurrentValue)
			research_th.rsh_subject.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_th.rsh_subject.FldCaption))

			' rsh_subject_th
			research_th.rsh_subject_th.EditCustomAttributes = ""
			research_th.rsh_subject_th.EditValue = ew_HtmlEncode(research_th.rsh_subject_th.CurrentValue)
			research_th.rsh_subject_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_th.rsh_subject_th.FldCaption))

			' rsh_intro
			research_th.rsh_intro.EditCustomAttributes = ""
			research_th.rsh_intro.EditValue = research_th.rsh_intro.CurrentValue
			research_th.rsh_intro.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_th.rsh_intro.FldCaption))

			' rsh_intro_th
			research_th.rsh_intro_th.EditCustomAttributes = ""
			research_th.rsh_intro_th.EditValue = ew_HtmlEncode(research_th.rsh_intro_th.CurrentValue)
			research_th.rsh_intro_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_th.rsh_intro_th.FldCaption))

			' rsh_content
			research_th.rsh_content.EditCustomAttributes = ""
			research_th.rsh_content.EditValue = research_th.rsh_content.CurrentValue
			research_th.rsh_content.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_th.rsh_content.FldCaption))

			' rsh_content_th
			research_th.rsh_content_th.EditCustomAttributes = ""
			research_th.rsh_content_th.EditValue = research_th.rsh_content_th.CurrentValue
			research_th.rsh_content_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_th.rsh_content_th.FldCaption))

			' rsh_show
			research_th.rsh_show.EditCustomAttributes = ""
			research_th.rsh_show.EditValue = ew_HtmlEncode(research_th.rsh_show.CurrentValue)
			research_th.rsh_show.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_th.rsh_show.FldCaption))

			' rsh_show_home
			research_th.rsh_show_home.EditCustomAttributes = ""
			research_th.rsh_show_home.EditValue = ew_HtmlEncode(research_th.rsh_show_home.CurrentValue)
			research_th.rsh_show_home.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_th.rsh_show_home.FldCaption))

			' rsh_create
			research_th.rsh_create.EditCustomAttributes = ""
			research_th.rsh_create.EditValue = ew_HtmlEncode(research_th.rsh_create.CurrentValue)
			research_th.rsh_create.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_th.rsh_create.FldCaption))

			' rsh_update
			research_th.rsh_update.EditCustomAttributes = ""
			research_th.rsh_update.EditValue = ew_HtmlEncode(research_th.rsh_update.CurrentValue)
			research_th.rsh_update.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_th.rsh_update.FldCaption))

			' Edit refer script
			' rsh_id

			research_th.rsh_id.HrefValue = ""

			' rsh_img
			research_th.rsh_img.HrefValue = ""

			' rsh_date
			research_th.rsh_date.HrefValue = ""

			' rsh_pdf
			research_th.rsh_pdf.HrefValue = ""

			' rsh_category
			research_th.rsh_category.HrefValue = ""

			' rsh_subject
			research_th.rsh_subject.HrefValue = ""

			' rsh_subject_th
			research_th.rsh_subject_th.HrefValue = ""

			' rsh_intro
			research_th.rsh_intro.HrefValue = ""

			' rsh_intro_th
			research_th.rsh_intro_th.HrefValue = ""

			' rsh_content
			research_th.rsh_content.HrefValue = ""

			' rsh_content_th
			research_th.rsh_content_th.HrefValue = ""

			' rsh_show
			research_th.rsh_show.HrefValue = ""

			' rsh_show_home
			research_th.rsh_show_home.HrefValue = ""

			' rsh_create
			research_th.rsh_create.HrefValue = ""

			' rsh_update
			research_th.rsh_update.HrefValue = ""
		End If
		If research_th.RowType = EW_ROWTYPE_ADD Or research_th.RowType = EW_ROWTYPE_EDIT Or research_th.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call research_th.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If research_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call research_th.Row_Rendered()
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
		If Not ew_CheckInteger(research_th.rsh_id.FormValue) Then
			Call ew_AddMessage(gsFormError, research_th.rsh_id.FldErrMsg)
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
		sFilter = research_th.KeyFilter
		research_th.CurrentFilter  = sFilter
		sSql = research_th.SQL
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

			' Field rsh_id
			' Field rsh_img

			Call research_th.rsh_img.SetDbValue(Rs, research_th.rsh_img.CurrentValue, Null, research_th.rsh_img.ReadOnly)

			' Field rsh_date
			Call research_th.rsh_date.SetDbValue(Rs, research_th.rsh_date.CurrentValue, Null, research_th.rsh_date.ReadOnly)

			' Field rsh_pdf
			Call research_th.rsh_pdf.SetDbValue(Rs, research_th.rsh_pdf.CurrentValue, Null, research_th.rsh_pdf.ReadOnly)

			' Field rsh_category
			Call research_th.rsh_category.SetDbValue(Rs, research_th.rsh_category.CurrentValue, Null, research_th.rsh_category.ReadOnly)

			' Field rsh_subject
			Call research_th.rsh_subject.SetDbValue(Rs, research_th.rsh_subject.CurrentValue, Null, research_th.rsh_subject.ReadOnly)

			' Field rsh_subject_th
			Call research_th.rsh_subject_th.SetDbValue(Rs, research_th.rsh_subject_th.CurrentValue, Null, research_th.rsh_subject_th.ReadOnly)

			' Field rsh_intro
			Call research_th.rsh_intro.SetDbValue(Rs, research_th.rsh_intro.CurrentValue, Null, research_th.rsh_intro.ReadOnly)

			' Field rsh_intro_th
			Call research_th.rsh_intro_th.SetDbValue(Rs, research_th.rsh_intro_th.CurrentValue, Null, research_th.rsh_intro_th.ReadOnly)

			' Field rsh_content
			Call research_th.rsh_content.SetDbValue(Rs, research_th.rsh_content.CurrentValue, Null, research_th.rsh_content.ReadOnly)

			' Field rsh_content_th
			Call research_th.rsh_content_th.SetDbValue(Rs, research_th.rsh_content_th.CurrentValue, Null, research_th.rsh_content_th.ReadOnly)

			' Field rsh_show
			Call research_th.rsh_show.SetDbValue(Rs, research_th.rsh_show.CurrentValue, Null, research_th.rsh_show.ReadOnly)

			' Field rsh_show_home
			Call research_th.rsh_show_home.SetDbValue(Rs, research_th.rsh_show_home.CurrentValue, Null, research_th.rsh_show_home.ReadOnly)

			' Field rsh_create
			Call research_th.rsh_create.SetDbValue(Rs, research_th.rsh_create.CurrentValue, Null, research_th.rsh_create.ReadOnly)

			' Field rsh_update
			Call research_th.rsh_update.SetDbValue(Rs, research_th.rsh_update.CurrentValue, Null, research_th.rsh_update.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = research_th.Row_Updating(RsOld, Rs)
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
				ElseIf research_th.CancelMessage <> "" Then
					FailureMessage = research_th.CancelMessage
					research_th.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call research_th.Row_Updated(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", research_th.TableVar, "pom_research_thlist.asp", research_th.TableVar, True)
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
