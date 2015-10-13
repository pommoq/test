<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_researchinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim research_add
Set research_add = New cresearch_add
Set Page = research_add

' Page init processing
research_add.Page_Init()

' Page main processing
research_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
research_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var research_add = new ew_Page("research_add");
research_add.PageID = "add"; // Page ID
var EW_PAGE_ID = research_add.PageID; // For backward compatibility
// Form object
var fresearchadd = new ew_Form("fresearchadd");
// Validate form
fresearchadd.Validate = function() {
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
				return this.OnError(elm, "<%= ew_JsEncode2(research.rsh_id.FldErrMsg) %>");
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
fresearchadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fresearchadd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fresearchadd.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If research.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% research_add.ShowPageHeader() %>
<% research_add.ShowMessage %>
<form name="fresearchadd" id="fresearchadd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="research">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_researchadd" class="table table-bordered table-striped">
<% If research.rsh_id.Visible Then ' rsh_id %>
	<tr id="r_rsh_id">
		<td><span id="elh_research_rsh_id"><%= research.rsh_id.FldCaption %></span></td>
		<td<%= research.rsh_id.CellAttributes %>>
<span id="el_research_rsh_id" class="control-group">
<input type="text" data-field="x_rsh_id" name="x_rsh_id" id="x_rsh_id" size="30" placeholder="<%= research.rsh_id.PlaceHolder %>" value="<%= research.rsh_id.EditValue %>"<%= research.rsh_id.EditAttributes %>>
</span>
<%= research.rsh_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If research.rsh_img.Visible Then ' rsh_img %>
	<tr id="r_rsh_img">
		<td><span id="elh_research_rsh_img"><%= research.rsh_img.FldCaption %></span></td>
		<td<%= research.rsh_img.CellAttributes %>>
<span id="el_research_rsh_img" class="control-group">
<input type="text" data-field="x_rsh_img" name="x_rsh_img" id="x_rsh_img" size="30" maxlength="255" placeholder="<%= research.rsh_img.PlaceHolder %>" value="<%= research.rsh_img.EditValue %>"<%= research.rsh_img.EditAttributes %>>
</span>
<%= research.rsh_img.CustomMsg %></td>
	</tr>
<% End If %>
<% If research.rsh_date.Visible Then ' rsh_date %>
	<tr id="r_rsh_date">
		<td><span id="elh_research_rsh_date"><%= research.rsh_date.FldCaption %></span></td>
		<td<%= research.rsh_date.CellAttributes %>>
<span id="el_research_rsh_date" class="control-group">
<input type="text" data-field="x_rsh_date" name="x_rsh_date" id="x_rsh_date" placeholder="<%= research.rsh_date.PlaceHolder %>" value="<%= research.rsh_date.EditValue %>"<%= research.rsh_date.EditAttributes %>>
</span>
<%= research.rsh_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If research.rsh_pdf.Visible Then ' rsh_pdf %>
	<tr id="r_rsh_pdf">
		<td><span id="elh_research_rsh_pdf"><%= research.rsh_pdf.FldCaption %></span></td>
		<td<%= research.rsh_pdf.CellAttributes %>>
<span id="el_research_rsh_pdf" class="control-group">
<input type="text" data-field="x_rsh_pdf" name="x_rsh_pdf" id="x_rsh_pdf" size="30" maxlength="255" placeholder="<%= research.rsh_pdf.PlaceHolder %>" value="<%= research.rsh_pdf.EditValue %>"<%= research.rsh_pdf.EditAttributes %>>
</span>
<%= research.rsh_pdf.CustomMsg %></td>
	</tr>
<% End If %>
<% If research.rsh_category.Visible Then ' rsh_category %>
	<tr id="r_rsh_category">
		<td><span id="elh_research_rsh_category"><%= research.rsh_category.FldCaption %></span></td>
		<td<%= research.rsh_category.CellAttributes %>>
<span id="el_research_rsh_category" class="control-group">
<input type="text" data-field="x_rsh_category" name="x_rsh_category" id="x_rsh_category" size="30" maxlength="255" placeholder="<%= research.rsh_category.PlaceHolder %>" value="<%= research.rsh_category.EditValue %>"<%= research.rsh_category.EditAttributes %>>
</span>
<%= research.rsh_category.CustomMsg %></td>
	</tr>
<% End If %>
<% If research.rsh_subject.Visible Then ' rsh_subject %>
	<tr id="r_rsh_subject">
		<td><span id="elh_research_rsh_subject"><%= research.rsh_subject.FldCaption %></span></td>
		<td<%= research.rsh_subject.CellAttributes %>>
<span id="el_research_rsh_subject" class="control-group">
<input type="text" data-field="x_rsh_subject" name="x_rsh_subject" id="x_rsh_subject" size="30" maxlength="255" placeholder="<%= research.rsh_subject.PlaceHolder %>" value="<%= research.rsh_subject.EditValue %>"<%= research.rsh_subject.EditAttributes %>>
</span>
<%= research.rsh_subject.CustomMsg %></td>
	</tr>
<% End If %>
<% If research.rsh_subject_th.Visible Then ' rsh_subject_th %>
	<tr id="r_rsh_subject_th">
		<td><span id="elh_research_rsh_subject_th"><%= research.rsh_subject_th.FldCaption %></span></td>
		<td<%= research.rsh_subject_th.CellAttributes %>>
<span id="el_research_rsh_subject_th" class="control-group">
<input type="text" data-field="x_rsh_subject_th" name="x_rsh_subject_th" id="x_rsh_subject_th" size="30" maxlength="255" placeholder="<%= research.rsh_subject_th.PlaceHolder %>" value="<%= research.rsh_subject_th.EditValue %>"<%= research.rsh_subject_th.EditAttributes %>>
</span>
<%= research.rsh_subject_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If research.rsh_intro.Visible Then ' rsh_intro %>
	<tr id="r_rsh_intro">
		<td><span id="elh_research_rsh_intro"><%= research.rsh_intro.FldCaption %></span></td>
		<td<%= research.rsh_intro.CellAttributes %>>
<span id="el_research_rsh_intro" class="control-group">
<textarea data-field="x_rsh_intro" name="x_rsh_intro" id="x_rsh_intro" cols="35" rows="4" placeholder="<%= research.rsh_intro.PlaceHolder %>"<%= research.rsh_intro.EditAttributes %>><%= research.rsh_intro.EditValue %></textarea>
</span>
<%= research.rsh_intro.CustomMsg %></td>
	</tr>
<% End If %>
<% If research.rsh_intro_th.Visible Then ' rsh_intro_th %>
	<tr id="r_rsh_intro_th">
		<td><span id="elh_research_rsh_intro_th"><%= research.rsh_intro_th.FldCaption %></span></td>
		<td<%= research.rsh_intro_th.CellAttributes %>>
<span id="el_research_rsh_intro_th" class="control-group">
<input type="text" data-field="x_rsh_intro_th" name="x_rsh_intro_th" id="x_rsh_intro_th" size="30" maxlength="255" placeholder="<%= research.rsh_intro_th.PlaceHolder %>" value="<%= research.rsh_intro_th.EditValue %>"<%= research.rsh_intro_th.EditAttributes %>>
</span>
<%= research.rsh_intro_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If research.rsh_content.Visible Then ' rsh_content %>
	<tr id="r_rsh_content">
		<td><span id="elh_research_rsh_content"><%= research.rsh_content.FldCaption %></span></td>
		<td<%= research.rsh_content.CellAttributes %>>
<span id="el_research_rsh_content" class="control-group">
<textarea data-field="x_rsh_content" name="x_rsh_content" id="x_rsh_content" cols="35" rows="4" placeholder="<%= research.rsh_content.PlaceHolder %>"<%= research.rsh_content.EditAttributes %>><%= research.rsh_content.EditValue %></textarea>
</span>
<%= research.rsh_content.CustomMsg %></td>
	</tr>
<% End If %>
<% If research.rsh_content_th.Visible Then ' rsh_content_th %>
	<tr id="r_rsh_content_th">
		<td><span id="elh_research_rsh_content_th"><%= research.rsh_content_th.FldCaption %></span></td>
		<td<%= research.rsh_content_th.CellAttributes %>>
<span id="el_research_rsh_content_th" class="control-group">
<textarea data-field="x_rsh_content_th" name="x_rsh_content_th" id="x_rsh_content_th" cols="35" rows="4" placeholder="<%= research.rsh_content_th.PlaceHolder %>"<%= research.rsh_content_th.EditAttributes %>><%= research.rsh_content_th.EditValue %></textarea>
</span>
<%= research.rsh_content_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If research.rsh_show.Visible Then ' rsh_show %>
	<tr id="r_rsh_show">
		<td><span id="elh_research_rsh_show"><%= research.rsh_show.FldCaption %></span></td>
		<td<%= research.rsh_show.CellAttributes %>>
<span id="el_research_rsh_show" class="control-group">
<input type="text" data-field="x_rsh_show" name="x_rsh_show" id="x_rsh_show" size="30" maxlength="255" placeholder="<%= research.rsh_show.PlaceHolder %>" value="<%= research.rsh_show.EditValue %>"<%= research.rsh_show.EditAttributes %>>
</span>
<%= research.rsh_show.CustomMsg %></td>
	</tr>
<% End If %>
<% If research.rsh_show_home.Visible Then ' rsh_show_home %>
	<tr id="r_rsh_show_home">
		<td><span id="elh_research_rsh_show_home"><%= research.rsh_show_home.FldCaption %></span></td>
		<td<%= research.rsh_show_home.CellAttributes %>>
<span id="el_research_rsh_show_home" class="control-group">
<input type="text" data-field="x_rsh_show_home" name="x_rsh_show_home" id="x_rsh_show_home" size="30" maxlength="255" placeholder="<%= research.rsh_show_home.PlaceHolder %>" value="<%= research.rsh_show_home.EditValue %>"<%= research.rsh_show_home.EditAttributes %>>
</span>
<%= research.rsh_show_home.CustomMsg %></td>
	</tr>
<% End If %>
<% If research.rsh_create.Visible Then ' rsh_create %>
	<tr id="r_rsh_create">
		<td><span id="elh_research_rsh_create"><%= research.rsh_create.FldCaption %></span></td>
		<td<%= research.rsh_create.CellAttributes %>>
<span id="el_research_rsh_create" class="control-group">
<input type="text" data-field="x_rsh_create" name="x_rsh_create" id="x_rsh_create" size="30" maxlength="255" placeholder="<%= research.rsh_create.PlaceHolder %>" value="<%= research.rsh_create.EditValue %>"<%= research.rsh_create.EditAttributes %>>
</span>
<%= research.rsh_create.CustomMsg %></td>
	</tr>
<% End If %>
<% If research.rsh_update.Visible Then ' rsh_update %>
	<tr id="r_rsh_update">
		<td><span id="elh_research_rsh_update"><%= research.rsh_update.FldCaption %></span></td>
		<td<%= research.rsh_update.CellAttributes %>>
<span id="el_research_rsh_update" class="control-group">
<input type="text" data-field="x_rsh_update" name="x_rsh_update" id="x_rsh_update" size="30" maxlength="255" placeholder="<%= research.rsh_update.PlaceHolder %>" value="<%= research.rsh_update.EditValue %>"<%= research.rsh_update.EditAttributes %>>
</span>
<%= research.rsh_update.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
fresearchadd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
research_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set research_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cresearch_add

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
		TableName = "research"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "research_add"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If research.UseTokenInUrl Then PageUrl = PageUrl & "t=" & research.TableVar & "&" ' add page token
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
		If research.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (research.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (research.TableVar = Request.QueryString("t"))
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
		If IsEmpty(research) Then Set research = New cresearch
		Set Table = research

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "add"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "research"

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

		research.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set research = Nothing
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
			research.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("rsh_id").Count > 0 Then
				research.rsh_id.QueryStringValue = Request.QueryString("rsh_id")
				Call research.SetKey("rsh_id", research.rsh_id.CurrentValue) ' Set up key
			Else
				Call research.SetKey("rsh_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				research.CurrentAction = "C" ' Copy Record
			Else
				research.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				research.CurrentAction = "I" ' Form error, reset action
				research.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case research.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_researchlist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				research.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = research.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_researchview.asp" Then sReturnUrl = research.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					research.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		research.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call research.ResetAttrs()
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
		research.rsh_id.CurrentValue = Null
		research.rsh_id.OldValue = research.rsh_id.CurrentValue
		research.rsh_img.CurrentValue = Null
		research.rsh_img.OldValue = research.rsh_img.CurrentValue
		research.rsh_date.CurrentValue = Null
		research.rsh_date.OldValue = research.rsh_date.CurrentValue
		research.rsh_pdf.CurrentValue = Null
		research.rsh_pdf.OldValue = research.rsh_pdf.CurrentValue
		research.rsh_category.CurrentValue = Null
		research.rsh_category.OldValue = research.rsh_category.CurrentValue
		research.rsh_subject.CurrentValue = Null
		research.rsh_subject.OldValue = research.rsh_subject.CurrentValue
		research.rsh_subject_th.CurrentValue = Null
		research.rsh_subject_th.OldValue = research.rsh_subject_th.CurrentValue
		research.rsh_intro.CurrentValue = Null
		research.rsh_intro.OldValue = research.rsh_intro.CurrentValue
		research.rsh_intro_th.CurrentValue = Null
		research.rsh_intro_th.OldValue = research.rsh_intro_th.CurrentValue
		research.rsh_content.CurrentValue = Null
		research.rsh_content.OldValue = research.rsh_content.CurrentValue
		research.rsh_content_th.CurrentValue = Null
		research.rsh_content_th.OldValue = research.rsh_content_th.CurrentValue
		research.rsh_show.CurrentValue = Null
		research.rsh_show.OldValue = research.rsh_show.CurrentValue
		research.rsh_show_home.CurrentValue = Null
		research.rsh_show_home.OldValue = research.rsh_show_home.CurrentValue
		research.rsh_create.CurrentValue = Null
		research.rsh_create.OldValue = research.rsh_create.CurrentValue
		research.rsh_update.CurrentValue = Null
		research.rsh_update.OldValue = research.rsh_update.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not research.rsh_id.FldIsDetailKey Then research.rsh_id.FormValue = ObjForm.GetValue("x_rsh_id")
		If Not research.rsh_img.FldIsDetailKey Then research.rsh_img.FormValue = ObjForm.GetValue("x_rsh_img")
		If Not research.rsh_date.FldIsDetailKey Then research.rsh_date.FormValue = ObjForm.GetValue("x_rsh_date")
		If Not research.rsh_date.FldIsDetailKey Then research.rsh_date.CurrentValue = ew_UnFormatDateTime(research.rsh_date.CurrentValue, 8)
		If Not research.rsh_pdf.FldIsDetailKey Then research.rsh_pdf.FormValue = ObjForm.GetValue("x_rsh_pdf")
		If Not research.rsh_category.FldIsDetailKey Then research.rsh_category.FormValue = ObjForm.GetValue("x_rsh_category")
		If Not research.rsh_subject.FldIsDetailKey Then research.rsh_subject.FormValue = ObjForm.GetValue("x_rsh_subject")
		If Not research.rsh_subject_th.FldIsDetailKey Then research.rsh_subject_th.FormValue = ObjForm.GetValue("x_rsh_subject_th")
		If Not research.rsh_intro.FldIsDetailKey Then research.rsh_intro.FormValue = ObjForm.GetValue("x_rsh_intro")
		If Not research.rsh_intro_th.FldIsDetailKey Then research.rsh_intro_th.FormValue = ObjForm.GetValue("x_rsh_intro_th")
		If Not research.rsh_content.FldIsDetailKey Then research.rsh_content.FormValue = ObjForm.GetValue("x_rsh_content")
		If Not research.rsh_content_th.FldIsDetailKey Then research.rsh_content_th.FormValue = ObjForm.GetValue("x_rsh_content_th")
		If Not research.rsh_show.FldIsDetailKey Then research.rsh_show.FormValue = ObjForm.GetValue("x_rsh_show")
		If Not research.rsh_show_home.FldIsDetailKey Then research.rsh_show_home.FormValue = ObjForm.GetValue("x_rsh_show_home")
		If Not research.rsh_create.FldIsDetailKey Then research.rsh_create.FormValue = ObjForm.GetValue("x_rsh_create")
		If Not research.rsh_update.FldIsDetailKey Then research.rsh_update.FormValue = ObjForm.GetValue("x_rsh_update")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		research.rsh_id.CurrentValue = research.rsh_id.FormValue
		research.rsh_img.CurrentValue = research.rsh_img.FormValue
		research.rsh_date.CurrentValue = research.rsh_date.FormValue
		research.rsh_date.CurrentValue = ew_UnFormatDateTime(research.rsh_date.CurrentValue, 8)
		research.rsh_pdf.CurrentValue = research.rsh_pdf.FormValue
		research.rsh_category.CurrentValue = research.rsh_category.FormValue
		research.rsh_subject.CurrentValue = research.rsh_subject.FormValue
		research.rsh_subject_th.CurrentValue = research.rsh_subject_th.FormValue
		research.rsh_intro.CurrentValue = research.rsh_intro.FormValue
		research.rsh_intro_th.CurrentValue = research.rsh_intro_th.FormValue
		research.rsh_content.CurrentValue = research.rsh_content.FormValue
		research.rsh_content_th.CurrentValue = research.rsh_content_th.FormValue
		research.rsh_show.CurrentValue = research.rsh_show.FormValue
		research.rsh_show_home.CurrentValue = research.rsh_show_home.FormValue
		research.rsh_create.CurrentValue = research.rsh_create.FormValue
		research.rsh_update.CurrentValue = research.rsh_update.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = research.KeyFilter

		' Call Row Selecting event
		Call research.Row_Selecting(sFilter)

		' Load sql based on filter
		research.CurrentFilter = sFilter
		sSql = research.SQL
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
		Call research.Row_Selected(RsRow)
		research.rsh_id.DbValue = RsRow("rsh_id")
		research.rsh_img.DbValue = RsRow("rsh_img")
		research.rsh_date.DbValue = RsRow("rsh_date")
		research.rsh_pdf.DbValue = RsRow("rsh_pdf")
		research.rsh_category.DbValue = RsRow("rsh_category")
		research.rsh_subject.DbValue = RsRow("rsh_subject")
		research.rsh_subject_th.DbValue = RsRow("rsh_subject_th")
		research.rsh_intro.DbValue = RsRow("rsh_intro")
		research.rsh_intro_th.DbValue = RsRow("rsh_intro_th")
		research.rsh_content.DbValue = RsRow("rsh_content")
		research.rsh_content_th.DbValue = RsRow("rsh_content_th")
		research.rsh_show.DbValue = RsRow("rsh_show")
		research.rsh_show_home.DbValue = RsRow("rsh_show_home")
		research.rsh_create.DbValue = RsRow("rsh_create")
		research.rsh_update.DbValue = RsRow("rsh_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		research.rsh_id.m_DbValue = Rs("rsh_id")
		research.rsh_img.m_DbValue = Rs("rsh_img")
		research.rsh_date.m_DbValue = Rs("rsh_date")
		research.rsh_pdf.m_DbValue = Rs("rsh_pdf")
		research.rsh_category.m_DbValue = Rs("rsh_category")
		research.rsh_subject.m_DbValue = Rs("rsh_subject")
		research.rsh_subject_th.m_DbValue = Rs("rsh_subject_th")
		research.rsh_intro.m_DbValue = Rs("rsh_intro")
		research.rsh_intro_th.m_DbValue = Rs("rsh_intro_th")
		research.rsh_content.m_DbValue = Rs("rsh_content")
		research.rsh_content_th.m_DbValue = Rs("rsh_content_th")
		research.rsh_show.m_DbValue = Rs("rsh_show")
		research.rsh_show_home.m_DbValue = Rs("rsh_show_home")
		research.rsh_create.m_DbValue = Rs("rsh_create")
		research.rsh_update.m_DbValue = Rs("rsh_update")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If research.GetKey("rsh_id")&"" <> "" Then
			research.rsh_id.CurrentValue = research.GetKey("rsh_id") ' rsh_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			research.CurrentFilter = research.KeyFilter
			Dim sSql
			sSql = research.SQL
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

		Call research.Row_Rendering()

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

		If research.RowType = EW_ROWTYPE_VIEW Then ' View row

			' rsh_id
			research.rsh_id.ViewValue = research.rsh_id.CurrentValue
			research.rsh_id.ViewCustomAttributes = ""

			' rsh_img
			research.rsh_img.ViewValue = research.rsh_img.CurrentValue
			research.rsh_img.ViewCustomAttributes = ""

			' rsh_date
			research.rsh_date.ViewValue = research.rsh_date.CurrentValue
			research.rsh_date.ViewCustomAttributes = ""

			' rsh_pdf
			research.rsh_pdf.ViewValue = research.rsh_pdf.CurrentValue
			research.rsh_pdf.ViewCustomAttributes = ""

			' rsh_category
			research.rsh_category.ViewValue = research.rsh_category.CurrentValue
			research.rsh_category.ViewCustomAttributes = ""

			' rsh_subject
			research.rsh_subject.ViewValue = research.rsh_subject.CurrentValue
			research.rsh_subject.ViewCustomAttributes = ""

			' rsh_subject_th
			research.rsh_subject_th.ViewValue = research.rsh_subject_th.CurrentValue
			research.rsh_subject_th.ViewCustomAttributes = ""

			' rsh_intro
			research.rsh_intro.ViewValue = research.rsh_intro.CurrentValue
			research.rsh_intro.ViewCustomAttributes = ""

			' rsh_intro_th
			research.rsh_intro_th.ViewValue = research.rsh_intro_th.CurrentValue
			research.rsh_intro_th.ViewCustomAttributes = ""

			' rsh_content
			research.rsh_content.ViewValue = research.rsh_content.CurrentValue
			research.rsh_content.ViewCustomAttributes = ""

			' rsh_content_th
			research.rsh_content_th.ViewValue = research.rsh_content_th.CurrentValue
			research.rsh_content_th.ViewCustomAttributes = ""

			' rsh_show
			research.rsh_show.ViewValue = research.rsh_show.CurrentValue
			research.rsh_show.ViewCustomAttributes = ""

			' rsh_show_home
			research.rsh_show_home.ViewValue = research.rsh_show_home.CurrentValue
			research.rsh_show_home.ViewCustomAttributes = ""

			' rsh_create
			research.rsh_create.ViewValue = research.rsh_create.CurrentValue
			research.rsh_create.ViewCustomAttributes = ""

			' rsh_update
			research.rsh_update.ViewValue = research.rsh_update.CurrentValue
			research.rsh_update.ViewCustomAttributes = ""

			' View refer script
			' rsh_id

			research.rsh_id.LinkCustomAttributes = ""
			research.rsh_id.HrefValue = ""
			research.rsh_id.TooltipValue = ""

			' rsh_img
			research.rsh_img.LinkCustomAttributes = ""
			research.rsh_img.HrefValue = ""
			research.rsh_img.TooltipValue = ""

			' rsh_date
			research.rsh_date.LinkCustomAttributes = ""
			research.rsh_date.HrefValue = ""
			research.rsh_date.TooltipValue = ""

			' rsh_pdf
			research.rsh_pdf.LinkCustomAttributes = ""
			research.rsh_pdf.HrefValue = ""
			research.rsh_pdf.TooltipValue = ""

			' rsh_category
			research.rsh_category.LinkCustomAttributes = ""
			research.rsh_category.HrefValue = ""
			research.rsh_category.TooltipValue = ""

			' rsh_subject
			research.rsh_subject.LinkCustomAttributes = ""
			research.rsh_subject.HrefValue = ""
			research.rsh_subject.TooltipValue = ""

			' rsh_subject_th
			research.rsh_subject_th.LinkCustomAttributes = ""
			research.rsh_subject_th.HrefValue = ""
			research.rsh_subject_th.TooltipValue = ""

			' rsh_intro
			research.rsh_intro.LinkCustomAttributes = ""
			research.rsh_intro.HrefValue = ""
			research.rsh_intro.TooltipValue = ""

			' rsh_intro_th
			research.rsh_intro_th.LinkCustomAttributes = ""
			research.rsh_intro_th.HrefValue = ""
			research.rsh_intro_th.TooltipValue = ""

			' rsh_content
			research.rsh_content.LinkCustomAttributes = ""
			research.rsh_content.HrefValue = ""
			research.rsh_content.TooltipValue = ""

			' rsh_content_th
			research.rsh_content_th.LinkCustomAttributes = ""
			research.rsh_content_th.HrefValue = ""
			research.rsh_content_th.TooltipValue = ""

			' rsh_show
			research.rsh_show.LinkCustomAttributes = ""
			research.rsh_show.HrefValue = ""
			research.rsh_show.TooltipValue = ""

			' rsh_show_home
			research.rsh_show_home.LinkCustomAttributes = ""
			research.rsh_show_home.HrefValue = ""
			research.rsh_show_home.TooltipValue = ""

			' rsh_create
			research.rsh_create.LinkCustomAttributes = ""
			research.rsh_create.HrefValue = ""
			research.rsh_create.TooltipValue = ""

			' rsh_update
			research.rsh_update.LinkCustomAttributes = ""
			research.rsh_update.HrefValue = ""
			research.rsh_update.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf research.RowType = EW_ROWTYPE_ADD Then ' Add row

			' rsh_id
			research.rsh_id.EditCustomAttributes = ""
			research.rsh_id.EditValue = ew_HtmlEncode(research.rsh_id.CurrentValue)
			research.rsh_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research.rsh_id.FldCaption))

			' rsh_img
			research.rsh_img.EditCustomAttributes = ""
			research.rsh_img.EditValue = ew_HtmlEncode(research.rsh_img.CurrentValue)
			research.rsh_img.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research.rsh_img.FldCaption))

			' rsh_date
			research.rsh_date.EditCustomAttributes = ""
			research.rsh_date.EditValue = ew_HtmlEncode(research.rsh_date.CurrentValue)
			research.rsh_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research.rsh_date.FldCaption))

			' rsh_pdf
			research.rsh_pdf.EditCustomAttributes = ""
			research.rsh_pdf.EditValue = ew_HtmlEncode(research.rsh_pdf.CurrentValue)
			research.rsh_pdf.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research.rsh_pdf.FldCaption))

			' rsh_category
			research.rsh_category.EditCustomAttributes = ""
			research.rsh_category.EditValue = ew_HtmlEncode(research.rsh_category.CurrentValue)
			research.rsh_category.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research.rsh_category.FldCaption))

			' rsh_subject
			research.rsh_subject.EditCustomAttributes = ""
			research.rsh_subject.EditValue = ew_HtmlEncode(research.rsh_subject.CurrentValue)
			research.rsh_subject.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research.rsh_subject.FldCaption))

			' rsh_subject_th
			research.rsh_subject_th.EditCustomAttributes = ""
			research.rsh_subject_th.EditValue = ew_HtmlEncode(research.rsh_subject_th.CurrentValue)
			research.rsh_subject_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research.rsh_subject_th.FldCaption))

			' rsh_intro
			research.rsh_intro.EditCustomAttributes = ""
			research.rsh_intro.EditValue = research.rsh_intro.CurrentValue
			research.rsh_intro.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research.rsh_intro.FldCaption))

			' rsh_intro_th
			research.rsh_intro_th.EditCustomAttributes = ""
			research.rsh_intro_th.EditValue = ew_HtmlEncode(research.rsh_intro_th.CurrentValue)
			research.rsh_intro_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research.rsh_intro_th.FldCaption))

			' rsh_content
			research.rsh_content.EditCustomAttributes = ""
			research.rsh_content.EditValue = research.rsh_content.CurrentValue
			research.rsh_content.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research.rsh_content.FldCaption))

			' rsh_content_th
			research.rsh_content_th.EditCustomAttributes = ""
			research.rsh_content_th.EditValue = research.rsh_content_th.CurrentValue
			research.rsh_content_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research.rsh_content_th.FldCaption))

			' rsh_show
			research.rsh_show.EditCustomAttributes = ""
			research.rsh_show.EditValue = ew_HtmlEncode(research.rsh_show.CurrentValue)
			research.rsh_show.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research.rsh_show.FldCaption))

			' rsh_show_home
			research.rsh_show_home.EditCustomAttributes = ""
			research.rsh_show_home.EditValue = ew_HtmlEncode(research.rsh_show_home.CurrentValue)
			research.rsh_show_home.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research.rsh_show_home.FldCaption))

			' rsh_create
			research.rsh_create.EditCustomAttributes = ""
			research.rsh_create.EditValue = ew_HtmlEncode(research.rsh_create.CurrentValue)
			research.rsh_create.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research.rsh_create.FldCaption))

			' rsh_update
			research.rsh_update.EditCustomAttributes = ""
			research.rsh_update.EditValue = ew_HtmlEncode(research.rsh_update.CurrentValue)
			research.rsh_update.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research.rsh_update.FldCaption))

			' Edit refer script
			' rsh_id

			research.rsh_id.HrefValue = ""

			' rsh_img
			research.rsh_img.HrefValue = ""

			' rsh_date
			research.rsh_date.HrefValue = ""

			' rsh_pdf
			research.rsh_pdf.HrefValue = ""

			' rsh_category
			research.rsh_category.HrefValue = ""

			' rsh_subject
			research.rsh_subject.HrefValue = ""

			' rsh_subject_th
			research.rsh_subject_th.HrefValue = ""

			' rsh_intro
			research.rsh_intro.HrefValue = ""

			' rsh_intro_th
			research.rsh_intro_th.HrefValue = ""

			' rsh_content
			research.rsh_content.HrefValue = ""

			' rsh_content_th
			research.rsh_content_th.HrefValue = ""

			' rsh_show
			research.rsh_show.HrefValue = ""

			' rsh_show_home
			research.rsh_show_home.HrefValue = ""

			' rsh_create
			research.rsh_create.HrefValue = ""

			' rsh_update
			research.rsh_update.HrefValue = ""
		End If
		If research.RowType = EW_ROWTYPE_ADD Or research.RowType = EW_ROWTYPE_EDIT Or research.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call research.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If research.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call research.Row_Rendered()
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
		If Not ew_CheckInteger(research.rsh_id.FormValue) Then
			Call ew_AddMessage(gsFormError, research.rsh_id.FldErrMsg)
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
		If research.rsh_id.CurrentValue <> "" Then ' Check field with unique index
			sFilter = "([rsh_id] = " & ew_AdjustSql(research.rsh_id.CurrentValue) & ")"
			Set RsChk = research.LoadRs(sFilter)
			If Not (RsChk Is Nothing) Then
				sIdxErrMsg = Replace(Language.Phrase("DupIndex"), "%f", research.rsh_id.FldCaption)
				sIdxErrMsg = Replace(sIdxErrMsg, "%v", research.rsh_id.CurrentValue)
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
		research.CurrentFilter = sFilter
		sSql = research.SQL
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

		' Field rsh_id
		Call research.rsh_id.SetDbValue(Rs, research.rsh_id.CurrentValue, Null, False)

		' Field rsh_img
		Call research.rsh_img.SetDbValue(Rs, research.rsh_img.CurrentValue, Null, False)

		' Field rsh_date
		Call research.rsh_date.SetDbValue(Rs, research.rsh_date.CurrentValue, Null, False)

		' Field rsh_pdf
		Call research.rsh_pdf.SetDbValue(Rs, research.rsh_pdf.CurrentValue, Null, False)

		' Field rsh_category
		Call research.rsh_category.SetDbValue(Rs, research.rsh_category.CurrentValue, Null, False)

		' Field rsh_subject
		Call research.rsh_subject.SetDbValue(Rs, research.rsh_subject.CurrentValue, Null, False)

		' Field rsh_subject_th
		Call research.rsh_subject_th.SetDbValue(Rs, research.rsh_subject_th.CurrentValue, Null, False)

		' Field rsh_intro
		Call research.rsh_intro.SetDbValue(Rs, research.rsh_intro.CurrentValue, Null, False)

		' Field rsh_intro_th
		Call research.rsh_intro_th.SetDbValue(Rs, research.rsh_intro_th.CurrentValue, Null, False)

		' Field rsh_content
		Call research.rsh_content.SetDbValue(Rs, research.rsh_content.CurrentValue, Null, False)

		' Field rsh_content_th
		Call research.rsh_content_th.SetDbValue(Rs, research.rsh_content_th.CurrentValue, Null, False)

		' Field rsh_show
		Call research.rsh_show.SetDbValue(Rs, research.rsh_show.CurrentValue, Null, False)

		' Field rsh_show_home
		Call research.rsh_show_home.SetDbValue(Rs, research.rsh_show_home.CurrentValue, Null, False)

		' Field rsh_create
		Call research.rsh_create.SetDbValue(Rs, research.rsh_create.CurrentValue, Null, False)

		' Field rsh_update
		Call research.rsh_update.SetDbValue(Rs, research.rsh_update.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = research.Row_Inserting(RsOld, Rs)

		' Check if key value entered
		If bInsertRow And research.ValidateKey And research.rsh_id.CurrentValue = "" And research.rsh_id.SessionValue = "" Then
			FailureMessage = Language.Phrase("InvalidKeyValue")
			bInsertRow = False
		End If

		' Check for duplicate key
		Dim sKeyErrMsg
		If bInsertRow And research.ValidateKey Then
			sFilter = research.KeyFilter
			Set RsChk = research.LoadRs(sFilter)
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
			ElseIf research.CancelMessage <> "" Then
				FailureMessage = research.CancelMessage
				research.CancelMessage = ""
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
			Call research.Row_Inserted(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", research.TableVar, "pom_researchlist.asp", research.TableVar, True)
		PageId = ew_IIf(research.CurrentAction = "C", "Copy", "Add")
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
