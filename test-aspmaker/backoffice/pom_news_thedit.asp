<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_news_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim news_th_edit
Set news_th_edit = New cnews_th_edit
Set Page = news_th_edit

' Page init processing
news_th_edit.Page_Init()

' Page main processing
news_th_edit.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
news_th_edit.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var news_th_edit = new ew_Page("news_th_edit");
news_th_edit.PageID = "edit"; // Page ID
var EW_PAGE_ID = news_th_edit.PageID; // For backward compatibility
// Form object
var fnews_thedit = new ew_Form("fnews_thedit");
// Validate form
fnews_thedit.Validate = function() {
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
			elm = this.GetElements("x" + infix + "_news_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(news_th.news_id.FldErrMsg) %>");
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
fnews_thedit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fnews_thedit.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fnews_thedit.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If news_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% news_th_edit.ShowPageHeader() %>
<% news_th_edit.ShowMessage %>
<form name="fnews_thedit" id="fnews_thedit" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="a_table" id="a_table" value="news_th">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table class="ewGrid"><tr><td>
<table id="tbl_news_thedit" class="table table-bordered table-striped">
<% If news_th.news_id.Visible Then ' news_id %>
	<tr id="r_news_id">
		<td><span id="elh_news_th_news_id"><%= news_th.news_id.FldCaption %></span></td>
		<td<%= news_th.news_id.CellAttributes %>>
<span id="el_news_th_news_id" class="control-group">
<span<%= news_th.news_id.ViewAttributes %>>
<%= news_th.news_id.EditValue %>
</span>
</span>
<input type="hidden" data-field="x_news_id" name="x_news_id" id="x_news_id" value="<%= Server.HTMLEncode(news_th.news_id.CurrentValue&"") %>">
<%= news_th.news_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.news_img.Visible Then ' news_img %>
	<tr id="r_news_img">
		<td><span id="elh_news_th_news_img"><%= news_th.news_img.FldCaption %></span></td>
		<td<%= news_th.news_img.CellAttributes %>>
<span id="el_news_th_news_img" class="control-group">
<input type="text" data-field="x_news_img" name="x_news_img" id="x_news_img" size="30" maxlength="255" placeholder="<%= news_th.news_img.PlaceHolder %>" value="<%= news_th.news_img.EditValue %>"<%= news_th.news_img.EditAttributes %>>
</span>
<%= news_th.news_img.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.news_date.Visible Then ' news_date %>
	<tr id="r_news_date">
		<td><span id="elh_news_th_news_date"><%= news_th.news_date.FldCaption %></span></td>
		<td<%= news_th.news_date.CellAttributes %>>
<span id="el_news_th_news_date" class="control-group">
<input type="text" data-field="x_news_date" name="x_news_date" id="x_news_date" placeholder="<%= news_th.news_date.PlaceHolder %>" value="<%= news_th.news_date.EditValue %>"<%= news_th.news_date.EditAttributes %>>
</span>
<%= news_th.news_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.news_category.Visible Then ' news_category %>
	<tr id="r_news_category">
		<td><span id="elh_news_th_news_category"><%= news_th.news_category.FldCaption %></span></td>
		<td<%= news_th.news_category.CellAttributes %>>
<span id="el_news_th_news_category" class="control-group">
<input type="text" data-field="x_news_category" name="x_news_category" id="x_news_category" size="30" maxlength="255" placeholder="<%= news_th.news_category.PlaceHolder %>" value="<%= news_th.news_category.EditValue %>"<%= news_th.news_category.EditAttributes %>>
</span>
<%= news_th.news_category.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.news_category_sub.Visible Then ' news_category_sub %>
	<tr id="r_news_category_sub">
		<td><span id="elh_news_th_news_category_sub"><%= news_th.news_category_sub.FldCaption %></span></td>
		<td<%= news_th.news_category_sub.CellAttributes %>>
<span id="el_news_th_news_category_sub" class="control-group">
<input type="text" data-field="x_news_category_sub" name="x_news_category_sub" id="x_news_category_sub" size="30" maxlength="255" placeholder="<%= news_th.news_category_sub.PlaceHolder %>" value="<%= news_th.news_category_sub.EditValue %>"<%= news_th.news_category_sub.EditAttributes %>>
</span>
<%= news_th.news_category_sub.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.start_date.Visible Then ' start_date %>
	<tr id="r_start_date">
		<td><span id="elh_news_th_start_date"><%= news_th.start_date.FldCaption %></span></td>
		<td<%= news_th.start_date.CellAttributes %>>
<span id="el_news_th_start_date" class="control-group">
<input type="text" data-field="x_start_date" name="x_start_date" id="x_start_date" placeholder="<%= news_th.start_date.PlaceHolder %>" value="<%= news_th.start_date.EditValue %>"<%= news_th.start_date.EditAttributes %>>
</span>
<%= news_th.start_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.end_date.Visible Then ' end_date %>
	<tr id="r_end_date">
		<td><span id="elh_news_th_end_date"><%= news_th.end_date.FldCaption %></span></td>
		<td<%= news_th.end_date.CellAttributes %>>
<span id="el_news_th_end_date" class="control-group">
<input type="text" data-field="x_end_date" name="x_end_date" id="x_end_date" placeholder="<%= news_th.end_date.PlaceHolder %>" value="<%= news_th.end_date.EditValue %>"<%= news_th.end_date.EditAttributes %>>
</span>
<%= news_th.end_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.news_pdf.Visible Then ' news_pdf %>
	<tr id="r_news_pdf">
		<td><span id="elh_news_th_news_pdf"><%= news_th.news_pdf.FldCaption %></span></td>
		<td<%= news_th.news_pdf.CellAttributes %>>
<span id="el_news_th_news_pdf" class="control-group">
<input type="text" data-field="x_news_pdf" name="x_news_pdf" id="x_news_pdf" size="30" maxlength="255" placeholder="<%= news_th.news_pdf.PlaceHolder %>" value="<%= news_th.news_pdf.EditValue %>"<%= news_th.news_pdf.EditAttributes %>>
</span>
<%= news_th.news_pdf.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.news_subject.Visible Then ' news_subject %>
	<tr id="r_news_subject">
		<td><span id="elh_news_th_news_subject"><%= news_th.news_subject.FldCaption %></span></td>
		<td<%= news_th.news_subject.CellAttributes %>>
<span id="el_news_th_news_subject" class="control-group">
<input type="text" data-field="x_news_subject" name="x_news_subject" id="x_news_subject" size="30" maxlength="255" placeholder="<%= news_th.news_subject.PlaceHolder %>" value="<%= news_th.news_subject.EditValue %>"<%= news_th.news_subject.EditAttributes %>>
</span>
<%= news_th.news_subject.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.news_subject_th.Visible Then ' news_subject_th %>
	<tr id="r_news_subject_th">
		<td><span id="elh_news_th_news_subject_th"><%= news_th.news_subject_th.FldCaption %></span></td>
		<td<%= news_th.news_subject_th.CellAttributes %>>
<span id="el_news_th_news_subject_th" class="control-group">
<input type="text" data-field="x_news_subject_th" name="x_news_subject_th" id="x_news_subject_th" size="30" maxlength="255" placeholder="<%= news_th.news_subject_th.PlaceHolder %>" value="<%= news_th.news_subject_th.EditValue %>"<%= news_th.news_subject_th.EditAttributes %>>
</span>
<%= news_th.news_subject_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.news_intro.Visible Then ' news_intro %>
	<tr id="r_news_intro">
		<td><span id="elh_news_th_news_intro"><%= news_th.news_intro.FldCaption %></span></td>
		<td<%= news_th.news_intro.CellAttributes %>>
<span id="el_news_th_news_intro" class="control-group">
<textarea data-field="x_news_intro" name="x_news_intro" id="x_news_intro" cols="35" rows="4" placeholder="<%= news_th.news_intro.PlaceHolder %>"<%= news_th.news_intro.EditAttributes %>><%= news_th.news_intro.EditValue %></textarea>
</span>
<%= news_th.news_intro.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.news_intro_th.Visible Then ' news_intro_th %>
	<tr id="r_news_intro_th">
		<td><span id="elh_news_th_news_intro_th"><%= news_th.news_intro_th.FldCaption %></span></td>
		<td<%= news_th.news_intro_th.CellAttributes %>>
<span id="el_news_th_news_intro_th" class="control-group">
<textarea data-field="x_news_intro_th" name="x_news_intro_th" id="x_news_intro_th" cols="35" rows="4" placeholder="<%= news_th.news_intro_th.PlaceHolder %>"<%= news_th.news_intro_th.EditAttributes %>><%= news_th.news_intro_th.EditValue %></textarea>
</span>
<%= news_th.news_intro_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.news_content.Visible Then ' news_content %>
	<tr id="r_news_content">
		<td><span id="elh_news_th_news_content"><%= news_th.news_content.FldCaption %></span></td>
		<td<%= news_th.news_content.CellAttributes %>>
<span id="el_news_th_news_content" class="control-group">
<textarea data-field="x_news_content" name="x_news_content" id="x_news_content" cols="35" rows="4" placeholder="<%= news_th.news_content.PlaceHolder %>"<%= news_th.news_content.EditAttributes %>><%= news_th.news_content.EditValue %></textarea>
</span>
<%= news_th.news_content.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.news_content_th.Visible Then ' news_content_th %>
	<tr id="r_news_content_th">
		<td><span id="elh_news_th_news_content_th"><%= news_th.news_content_th.FldCaption %></span></td>
		<td<%= news_th.news_content_th.CellAttributes %>>
<span id="el_news_th_news_content_th" class="control-group">
<textarea data-field="x_news_content_th" name="x_news_content_th" id="x_news_content_th" cols="35" rows="4" placeholder="<%= news_th.news_content_th.PlaceHolder %>"<%= news_th.news_content_th.EditAttributes %>><%= news_th.news_content_th.EditValue %></textarea>
</span>
<%= news_th.news_content_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.news_show_en.Visible Then ' news_show_en %>
	<tr id="r_news_show_en">
		<td><span id="elh_news_th_news_show_en"><%= news_th.news_show_en.FldCaption %></span></td>
		<td<%= news_th.news_show_en.CellAttributes %>>
<span id="el_news_th_news_show_en" class="control-group">
<input type="text" data-field="x_news_show_en" name="x_news_show_en" id="x_news_show_en" size="30" maxlength="255" placeholder="<%= news_th.news_show_en.PlaceHolder %>" value="<%= news_th.news_show_en.EditValue %>"<%= news_th.news_show_en.EditAttributes %>>
</span>
<%= news_th.news_show_en.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.news_show.Visible Then ' news_show %>
	<tr id="r_news_show">
		<td><span id="elh_news_th_news_show"><%= news_th.news_show.FldCaption %></span></td>
		<td<%= news_th.news_show.CellAttributes %>>
<span id="el_news_th_news_show" class="control-group">
<input type="text" data-field="x_news_show" name="x_news_show" id="x_news_show" size="30" maxlength="255" placeholder="<%= news_th.news_show.PlaceHolder %>" value="<%= news_th.news_show.EditValue %>"<%= news_th.news_show.EditAttributes %>>
</span>
<%= news_th.news_show.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.news_show_home.Visible Then ' news_show_home %>
	<tr id="r_news_show_home">
		<td><span id="elh_news_th_news_show_home"><%= news_th.news_show_home.FldCaption %></span></td>
		<td<%= news_th.news_show_home.CellAttributes %>>
<span id="el_news_th_news_show_home" class="control-group">
<input type="text" data-field="x_news_show_home" name="x_news_show_home" id="x_news_show_home" size="30" maxlength="255" placeholder="<%= news_th.news_show_home.PlaceHolder %>" value="<%= news_th.news_show_home.EditValue %>"<%= news_th.news_show_home.EditAttributes %>>
</span>
<%= news_th.news_show_home.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.news_create.Visible Then ' news_create %>
	<tr id="r_news_create">
		<td><span id="elh_news_th_news_create"><%= news_th.news_create.FldCaption %></span></td>
		<td<%= news_th.news_create.CellAttributes %>>
<span id="el_news_th_news_create" class="control-group">
<input type="text" data-field="x_news_create" name="x_news_create" id="x_news_create" size="30" maxlength="255" placeholder="<%= news_th.news_create.PlaceHolder %>" value="<%= news_th.news_create.EditValue %>"<%= news_th.news_create.EditAttributes %>>
</span>
<%= news_th.news_create.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_th.news_update.Visible Then ' news_update %>
	<tr id="r_news_update">
		<td><span id="elh_news_th_news_update"><%= news_th.news_update.FldCaption %></span></td>
		<td<%= news_th.news_update.CellAttributes %>>
<span id="el_news_th_news_update" class="control-group">
<input type="text" data-field="x_news_update" name="x_news_update" id="x_news_update" size="30" maxlength="255" placeholder="<%= news_th.news_update.PlaceHolder %>" value="<%= news_th.news_update.EditValue %>"<%= news_th.news_update.EditAttributes %>>
</span>
<%= news_th.news_update.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("EditBtn") %></button>
</form>
<script type="text/javascript">
fnews_thedit.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
news_th_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set news_th_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cnews_th_edit

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
		TableName = "news_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "news_th_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If news_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & news_th.TableVar & "&" ' add page token
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
		If news_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (news_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (news_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(news_th) Then Set news_th = New cnews_th
		Set Table = news_th

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "news_th"

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

		news_th.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set news_th = Nothing
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
		If Request.QueryString("news_id").Count > 0 Then
			news_th.news_id.QueryStringValue = Request.QueryString("news_id")
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Process form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			news_th.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values
		Else
			news_th.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If news_th.news_id.CurrentValue = "" Then Call Page_Terminate("pom_news_thlist.asp") ' Invalid key, return to list

		' Validate form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			If Not ValidateForm() Then
				news_th.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				news_th.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		End If
		Select Case news_th.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_news_thlist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				news_th.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					sReturnUrl = news_th.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					news_th.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		news_th.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call news_th.ResetAttrs()
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
				news_th.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					news_th.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = news_th.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			news_th.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			news_th.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			news_th.StartRecordNumber = StartRec
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
		If Not news_th.news_id.FldIsDetailKey Then news_th.news_id.FormValue = ObjForm.GetValue("x_news_id")
		If Not news_th.news_img.FldIsDetailKey Then news_th.news_img.FormValue = ObjForm.GetValue("x_news_img")
		If Not news_th.news_date.FldIsDetailKey Then news_th.news_date.FormValue = ObjForm.GetValue("x_news_date")
		If Not news_th.news_date.FldIsDetailKey Then news_th.news_date.CurrentValue = ew_UnFormatDateTime(news_th.news_date.CurrentValue, 8)
		If Not news_th.news_category.FldIsDetailKey Then news_th.news_category.FormValue = ObjForm.GetValue("x_news_category")
		If Not news_th.news_category_sub.FldIsDetailKey Then news_th.news_category_sub.FormValue = ObjForm.GetValue("x_news_category_sub")
		If Not news_th.start_date.FldIsDetailKey Then news_th.start_date.FormValue = ObjForm.GetValue("x_start_date")
		If Not news_th.start_date.FldIsDetailKey Then news_th.start_date.CurrentValue = ew_UnFormatDateTime(news_th.start_date.CurrentValue, 8)
		If Not news_th.end_date.FldIsDetailKey Then news_th.end_date.FormValue = ObjForm.GetValue("x_end_date")
		If Not news_th.end_date.FldIsDetailKey Then news_th.end_date.CurrentValue = ew_UnFormatDateTime(news_th.end_date.CurrentValue, 8)
		If Not news_th.news_pdf.FldIsDetailKey Then news_th.news_pdf.FormValue = ObjForm.GetValue("x_news_pdf")
		If Not news_th.news_subject.FldIsDetailKey Then news_th.news_subject.FormValue = ObjForm.GetValue("x_news_subject")
		If Not news_th.news_subject_th.FldIsDetailKey Then news_th.news_subject_th.FormValue = ObjForm.GetValue("x_news_subject_th")
		If Not news_th.news_intro.FldIsDetailKey Then news_th.news_intro.FormValue = ObjForm.GetValue("x_news_intro")
		If Not news_th.news_intro_th.FldIsDetailKey Then news_th.news_intro_th.FormValue = ObjForm.GetValue("x_news_intro_th")
		If Not news_th.news_content.FldIsDetailKey Then news_th.news_content.FormValue = ObjForm.GetValue("x_news_content")
		If Not news_th.news_content_th.FldIsDetailKey Then news_th.news_content_th.FormValue = ObjForm.GetValue("x_news_content_th")
		If Not news_th.news_show_en.FldIsDetailKey Then news_th.news_show_en.FormValue = ObjForm.GetValue("x_news_show_en")
		If Not news_th.news_show.FldIsDetailKey Then news_th.news_show.FormValue = ObjForm.GetValue("x_news_show")
		If Not news_th.news_show_home.FldIsDetailKey Then news_th.news_show_home.FormValue = ObjForm.GetValue("x_news_show_home")
		If Not news_th.news_create.FldIsDetailKey Then news_th.news_create.FormValue = ObjForm.GetValue("x_news_create")
		If Not news_th.news_update.FldIsDetailKey Then news_th.news_update.FormValue = ObjForm.GetValue("x_news_update")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		news_th.news_id.CurrentValue = news_th.news_id.FormValue
		news_th.news_img.CurrentValue = news_th.news_img.FormValue
		news_th.news_date.CurrentValue = news_th.news_date.FormValue
		news_th.news_date.CurrentValue = ew_UnFormatDateTime(news_th.news_date.CurrentValue, 8)
		news_th.news_category.CurrentValue = news_th.news_category.FormValue
		news_th.news_category_sub.CurrentValue = news_th.news_category_sub.FormValue
		news_th.start_date.CurrentValue = news_th.start_date.FormValue
		news_th.start_date.CurrentValue = ew_UnFormatDateTime(news_th.start_date.CurrentValue, 8)
		news_th.end_date.CurrentValue = news_th.end_date.FormValue
		news_th.end_date.CurrentValue = ew_UnFormatDateTime(news_th.end_date.CurrentValue, 8)
		news_th.news_pdf.CurrentValue = news_th.news_pdf.FormValue
		news_th.news_subject.CurrentValue = news_th.news_subject.FormValue
		news_th.news_subject_th.CurrentValue = news_th.news_subject_th.FormValue
		news_th.news_intro.CurrentValue = news_th.news_intro.FormValue
		news_th.news_intro_th.CurrentValue = news_th.news_intro_th.FormValue
		news_th.news_content.CurrentValue = news_th.news_content.FormValue
		news_th.news_content_th.CurrentValue = news_th.news_content_th.FormValue
		news_th.news_show_en.CurrentValue = news_th.news_show_en.FormValue
		news_th.news_show.CurrentValue = news_th.news_show.FormValue
		news_th.news_show_home.CurrentValue = news_th.news_show_home.FormValue
		news_th.news_create.CurrentValue = news_th.news_create.FormValue
		news_th.news_update.CurrentValue = news_th.news_update.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = news_th.KeyFilter

		' Call Row Selecting event
		Call news_th.Row_Selecting(sFilter)

		' Load sql based on filter
		news_th.CurrentFilter = sFilter
		sSql = news_th.SQL
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
		Call news_th.Row_Selected(RsRow)
		news_th.news_id.DbValue = RsRow("news_id")
		news_th.news_img.DbValue = RsRow("news_img")
		news_th.news_date.DbValue = RsRow("news_date")
		news_th.news_category.DbValue = RsRow("news_category")
		news_th.news_category_sub.DbValue = RsRow("news_category_sub")
		news_th.start_date.DbValue = RsRow("start_date")
		news_th.end_date.DbValue = RsRow("end_date")
		news_th.news_pdf.DbValue = RsRow("news_pdf")
		news_th.news_subject.DbValue = RsRow("news_subject")
		news_th.news_subject_th.DbValue = RsRow("news_subject_th")
		news_th.news_intro.DbValue = RsRow("news_intro")
		news_th.news_intro_th.DbValue = RsRow("news_intro_th")
		news_th.news_content.DbValue = RsRow("news_content")
		news_th.news_content_th.DbValue = RsRow("news_content_th")
		news_th.news_show_en.DbValue = RsRow("news_show_en")
		news_th.news_show.DbValue = RsRow("news_show")
		news_th.news_show_home.DbValue = RsRow("news_show_home")
		news_th.news_create.DbValue = RsRow("news_create")
		news_th.news_update.DbValue = RsRow("news_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		news_th.news_id.m_DbValue = Rs("news_id")
		news_th.news_img.m_DbValue = Rs("news_img")
		news_th.news_date.m_DbValue = Rs("news_date")
		news_th.news_category.m_DbValue = Rs("news_category")
		news_th.news_category_sub.m_DbValue = Rs("news_category_sub")
		news_th.start_date.m_DbValue = Rs("start_date")
		news_th.end_date.m_DbValue = Rs("end_date")
		news_th.news_pdf.m_DbValue = Rs("news_pdf")
		news_th.news_subject.m_DbValue = Rs("news_subject")
		news_th.news_subject_th.m_DbValue = Rs("news_subject_th")
		news_th.news_intro.m_DbValue = Rs("news_intro")
		news_th.news_intro_th.m_DbValue = Rs("news_intro_th")
		news_th.news_content.m_DbValue = Rs("news_content")
		news_th.news_content_th.m_DbValue = Rs("news_content_th")
		news_th.news_show_en.m_DbValue = Rs("news_show_en")
		news_th.news_show.m_DbValue = Rs("news_show")
		news_th.news_show_home.m_DbValue = Rs("news_show_home")
		news_th.news_create.m_DbValue = Rs("news_create")
		news_th.news_update.m_DbValue = Rs("news_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call news_th.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' news_id
		' news_img
		' news_date
		' news_category
		' news_category_sub
		' start_date
		' end_date
		' news_pdf
		' news_subject
		' news_subject_th
		' news_intro
		' news_intro_th
		' news_content
		' news_content_th
		' news_show_en
		' news_show
		' news_show_home
		' news_create
		' news_update
		' -----------
		'  View  Row
		' -----------

		If news_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' news_id
			news_th.news_id.ViewValue = news_th.news_id.CurrentValue
			news_th.news_id.ViewCustomAttributes = ""

			' news_img
			news_th.news_img.ViewValue = news_th.news_img.CurrentValue
			news_th.news_img.ViewCustomAttributes = ""

			' news_date
			news_th.news_date.ViewValue = news_th.news_date.CurrentValue
			news_th.news_date.ViewCustomAttributes = ""

			' news_category
			news_th.news_category.ViewValue = news_th.news_category.CurrentValue
			news_th.news_category.ViewCustomAttributes = ""

			' news_category_sub
			news_th.news_category_sub.ViewValue = news_th.news_category_sub.CurrentValue
			news_th.news_category_sub.ViewCustomAttributes = ""

			' start_date
			news_th.start_date.ViewValue = news_th.start_date.CurrentValue
			news_th.start_date.ViewCustomAttributes = ""

			' end_date
			news_th.end_date.ViewValue = news_th.end_date.CurrentValue
			news_th.end_date.ViewCustomAttributes = ""

			' news_pdf
			news_th.news_pdf.ViewValue = news_th.news_pdf.CurrentValue
			news_th.news_pdf.ViewCustomAttributes = ""

			' news_subject
			news_th.news_subject.ViewValue = news_th.news_subject.CurrentValue
			news_th.news_subject.ViewCustomAttributes = ""

			' news_subject_th
			news_th.news_subject_th.ViewValue = news_th.news_subject_th.CurrentValue
			news_th.news_subject_th.ViewCustomAttributes = ""

			' news_intro
			news_th.news_intro.ViewValue = news_th.news_intro.CurrentValue
			news_th.news_intro.ViewCustomAttributes = ""

			' news_intro_th
			news_th.news_intro_th.ViewValue = news_th.news_intro_th.CurrentValue
			news_th.news_intro_th.ViewCustomAttributes = ""

			' news_content
			news_th.news_content.ViewValue = news_th.news_content.CurrentValue
			news_th.news_content.ViewCustomAttributes = ""

			' news_content_th
			news_th.news_content_th.ViewValue = news_th.news_content_th.CurrentValue
			news_th.news_content_th.ViewCustomAttributes = ""

			' news_show_en
			news_th.news_show_en.ViewValue = news_th.news_show_en.CurrentValue
			news_th.news_show_en.ViewCustomAttributes = ""

			' news_show
			news_th.news_show.ViewValue = news_th.news_show.CurrentValue
			news_th.news_show.ViewCustomAttributes = ""

			' news_show_home
			news_th.news_show_home.ViewValue = news_th.news_show_home.CurrentValue
			news_th.news_show_home.ViewCustomAttributes = ""

			' news_create
			news_th.news_create.ViewValue = news_th.news_create.CurrentValue
			news_th.news_create.ViewCustomAttributes = ""

			' news_update
			news_th.news_update.ViewValue = news_th.news_update.CurrentValue
			news_th.news_update.ViewCustomAttributes = ""

			' View refer script
			' news_id

			news_th.news_id.LinkCustomAttributes = ""
			news_th.news_id.HrefValue = ""
			news_th.news_id.TooltipValue = ""

			' news_img
			news_th.news_img.LinkCustomAttributes = ""
			news_th.news_img.HrefValue = ""
			news_th.news_img.TooltipValue = ""

			' news_date
			news_th.news_date.LinkCustomAttributes = ""
			news_th.news_date.HrefValue = ""
			news_th.news_date.TooltipValue = ""

			' news_category
			news_th.news_category.LinkCustomAttributes = ""
			news_th.news_category.HrefValue = ""
			news_th.news_category.TooltipValue = ""

			' news_category_sub
			news_th.news_category_sub.LinkCustomAttributes = ""
			news_th.news_category_sub.HrefValue = ""
			news_th.news_category_sub.TooltipValue = ""

			' start_date
			news_th.start_date.LinkCustomAttributes = ""
			news_th.start_date.HrefValue = ""
			news_th.start_date.TooltipValue = ""

			' end_date
			news_th.end_date.LinkCustomAttributes = ""
			news_th.end_date.HrefValue = ""
			news_th.end_date.TooltipValue = ""

			' news_pdf
			news_th.news_pdf.LinkCustomAttributes = ""
			news_th.news_pdf.HrefValue = ""
			news_th.news_pdf.TooltipValue = ""

			' news_subject
			news_th.news_subject.LinkCustomAttributes = ""
			news_th.news_subject.HrefValue = ""
			news_th.news_subject.TooltipValue = ""

			' news_subject_th
			news_th.news_subject_th.LinkCustomAttributes = ""
			news_th.news_subject_th.HrefValue = ""
			news_th.news_subject_th.TooltipValue = ""

			' news_intro
			news_th.news_intro.LinkCustomAttributes = ""
			news_th.news_intro.HrefValue = ""
			news_th.news_intro.TooltipValue = ""

			' news_intro_th
			news_th.news_intro_th.LinkCustomAttributes = ""
			news_th.news_intro_th.HrefValue = ""
			news_th.news_intro_th.TooltipValue = ""

			' news_content
			news_th.news_content.LinkCustomAttributes = ""
			news_th.news_content.HrefValue = ""
			news_th.news_content.TooltipValue = ""

			' news_content_th
			news_th.news_content_th.LinkCustomAttributes = ""
			news_th.news_content_th.HrefValue = ""
			news_th.news_content_th.TooltipValue = ""

			' news_show_en
			news_th.news_show_en.LinkCustomAttributes = ""
			news_th.news_show_en.HrefValue = ""
			news_th.news_show_en.TooltipValue = ""

			' news_show
			news_th.news_show.LinkCustomAttributes = ""
			news_th.news_show.HrefValue = ""
			news_th.news_show.TooltipValue = ""

			' news_show_home
			news_th.news_show_home.LinkCustomAttributes = ""
			news_th.news_show_home.HrefValue = ""
			news_th.news_show_home.TooltipValue = ""

			' news_create
			news_th.news_create.LinkCustomAttributes = ""
			news_th.news_create.HrefValue = ""
			news_th.news_create.TooltipValue = ""

			' news_update
			news_th.news_update.LinkCustomAttributes = ""
			news_th.news_update.HrefValue = ""
			news_th.news_update.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf news_th.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' news_id
			news_th.news_id.EditCustomAttributes = ""
			news_th.news_id.EditValue = news_th.news_id.CurrentValue
			news_th.news_id.ViewCustomAttributes = ""

			' news_img
			news_th.news_img.EditCustomAttributes = ""
			news_th.news_img.EditValue = ew_HtmlEncode(news_th.news_img.CurrentValue)
			news_th.news_img.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.news_img.FldCaption))

			' news_date
			news_th.news_date.EditCustomAttributes = ""
			news_th.news_date.EditValue = ew_HtmlEncode(news_th.news_date.CurrentValue)
			news_th.news_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.news_date.FldCaption))

			' news_category
			news_th.news_category.EditCustomAttributes = ""
			news_th.news_category.EditValue = ew_HtmlEncode(news_th.news_category.CurrentValue)
			news_th.news_category.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.news_category.FldCaption))

			' news_category_sub
			news_th.news_category_sub.EditCustomAttributes = ""
			news_th.news_category_sub.EditValue = ew_HtmlEncode(news_th.news_category_sub.CurrentValue)
			news_th.news_category_sub.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.news_category_sub.FldCaption))

			' start_date
			news_th.start_date.EditCustomAttributes = ""
			news_th.start_date.EditValue = ew_HtmlEncode(news_th.start_date.CurrentValue)
			news_th.start_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.start_date.FldCaption))

			' end_date
			news_th.end_date.EditCustomAttributes = ""
			news_th.end_date.EditValue = ew_HtmlEncode(news_th.end_date.CurrentValue)
			news_th.end_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.end_date.FldCaption))

			' news_pdf
			news_th.news_pdf.EditCustomAttributes = ""
			news_th.news_pdf.EditValue = ew_HtmlEncode(news_th.news_pdf.CurrentValue)
			news_th.news_pdf.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.news_pdf.FldCaption))

			' news_subject
			news_th.news_subject.EditCustomAttributes = ""
			news_th.news_subject.EditValue = ew_HtmlEncode(news_th.news_subject.CurrentValue)
			news_th.news_subject.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.news_subject.FldCaption))

			' news_subject_th
			news_th.news_subject_th.EditCustomAttributes = ""
			news_th.news_subject_th.EditValue = ew_HtmlEncode(news_th.news_subject_th.CurrentValue)
			news_th.news_subject_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.news_subject_th.FldCaption))

			' news_intro
			news_th.news_intro.EditCustomAttributes = ""
			news_th.news_intro.EditValue = news_th.news_intro.CurrentValue
			news_th.news_intro.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.news_intro.FldCaption))

			' news_intro_th
			news_th.news_intro_th.EditCustomAttributes = ""
			news_th.news_intro_th.EditValue = news_th.news_intro_th.CurrentValue
			news_th.news_intro_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.news_intro_th.FldCaption))

			' news_content
			news_th.news_content.EditCustomAttributes = ""
			news_th.news_content.EditValue = news_th.news_content.CurrentValue
			news_th.news_content.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.news_content.FldCaption))

			' news_content_th
			news_th.news_content_th.EditCustomAttributes = ""
			news_th.news_content_th.EditValue = news_th.news_content_th.CurrentValue
			news_th.news_content_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.news_content_th.FldCaption))

			' news_show_en
			news_th.news_show_en.EditCustomAttributes = ""
			news_th.news_show_en.EditValue = ew_HtmlEncode(news_th.news_show_en.CurrentValue)
			news_th.news_show_en.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.news_show_en.FldCaption))

			' news_show
			news_th.news_show.EditCustomAttributes = ""
			news_th.news_show.EditValue = ew_HtmlEncode(news_th.news_show.CurrentValue)
			news_th.news_show.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.news_show.FldCaption))

			' news_show_home
			news_th.news_show_home.EditCustomAttributes = ""
			news_th.news_show_home.EditValue = ew_HtmlEncode(news_th.news_show_home.CurrentValue)
			news_th.news_show_home.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.news_show_home.FldCaption))

			' news_create
			news_th.news_create.EditCustomAttributes = ""
			news_th.news_create.EditValue = ew_HtmlEncode(news_th.news_create.CurrentValue)
			news_th.news_create.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.news_create.FldCaption))

			' news_update
			news_th.news_update.EditCustomAttributes = ""
			news_th.news_update.EditValue = ew_HtmlEncode(news_th.news_update.CurrentValue)
			news_th.news_update.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_th.news_update.FldCaption))

			' Edit refer script
			' news_id

			news_th.news_id.HrefValue = ""

			' news_img
			news_th.news_img.HrefValue = ""

			' news_date
			news_th.news_date.HrefValue = ""

			' news_category
			news_th.news_category.HrefValue = ""

			' news_category_sub
			news_th.news_category_sub.HrefValue = ""

			' start_date
			news_th.start_date.HrefValue = ""

			' end_date
			news_th.end_date.HrefValue = ""

			' news_pdf
			news_th.news_pdf.HrefValue = ""

			' news_subject
			news_th.news_subject.HrefValue = ""

			' news_subject_th
			news_th.news_subject_th.HrefValue = ""

			' news_intro
			news_th.news_intro.HrefValue = ""

			' news_intro_th
			news_th.news_intro_th.HrefValue = ""

			' news_content
			news_th.news_content.HrefValue = ""

			' news_content_th
			news_th.news_content_th.HrefValue = ""

			' news_show_en
			news_th.news_show_en.HrefValue = ""

			' news_show
			news_th.news_show.HrefValue = ""

			' news_show_home
			news_th.news_show_home.HrefValue = ""

			' news_create
			news_th.news_create.HrefValue = ""

			' news_update
			news_th.news_update.HrefValue = ""
		End If
		If news_th.RowType = EW_ROWTYPE_ADD Or news_th.RowType = EW_ROWTYPE_EDIT Or news_th.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call news_th.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If news_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call news_th.Row_Rendered()
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
		If Not ew_CheckInteger(news_th.news_id.FormValue) Then
			Call ew_AddMessage(gsFormError, news_th.news_id.FldErrMsg)
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
		sFilter = news_th.KeyFilter
		news_th.CurrentFilter  = sFilter
		sSql = news_th.SQL
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

			' Field news_id
			' Field news_img

			Call news_th.news_img.SetDbValue(Rs, news_th.news_img.CurrentValue, Null, news_th.news_img.ReadOnly)

			' Field news_date
			Call news_th.news_date.SetDbValue(Rs, news_th.news_date.CurrentValue, Null, news_th.news_date.ReadOnly)

			' Field news_category
			Call news_th.news_category.SetDbValue(Rs, news_th.news_category.CurrentValue, Null, news_th.news_category.ReadOnly)

			' Field news_category_sub
			Call news_th.news_category_sub.SetDbValue(Rs, news_th.news_category_sub.CurrentValue, Null, news_th.news_category_sub.ReadOnly)

			' Field start_date
			Call news_th.start_date.SetDbValue(Rs, news_th.start_date.CurrentValue, Null, news_th.start_date.ReadOnly)

			' Field end_date
			Call news_th.end_date.SetDbValue(Rs, news_th.end_date.CurrentValue, Null, news_th.end_date.ReadOnly)

			' Field news_pdf
			Call news_th.news_pdf.SetDbValue(Rs, news_th.news_pdf.CurrentValue, Null, news_th.news_pdf.ReadOnly)

			' Field news_subject
			Call news_th.news_subject.SetDbValue(Rs, news_th.news_subject.CurrentValue, Null, news_th.news_subject.ReadOnly)

			' Field news_subject_th
			Call news_th.news_subject_th.SetDbValue(Rs, news_th.news_subject_th.CurrentValue, Null, news_th.news_subject_th.ReadOnly)

			' Field news_intro
			Call news_th.news_intro.SetDbValue(Rs, news_th.news_intro.CurrentValue, Null, news_th.news_intro.ReadOnly)

			' Field news_intro_th
			Call news_th.news_intro_th.SetDbValue(Rs, news_th.news_intro_th.CurrentValue, Null, news_th.news_intro_th.ReadOnly)

			' Field news_content
			Call news_th.news_content.SetDbValue(Rs, news_th.news_content.CurrentValue, Null, news_th.news_content.ReadOnly)

			' Field news_content_th
			Call news_th.news_content_th.SetDbValue(Rs, news_th.news_content_th.CurrentValue, Null, news_th.news_content_th.ReadOnly)

			' Field news_show_en
			Call news_th.news_show_en.SetDbValue(Rs, news_th.news_show_en.CurrentValue, Null, news_th.news_show_en.ReadOnly)

			' Field news_show
			Call news_th.news_show.SetDbValue(Rs, news_th.news_show.CurrentValue, Null, news_th.news_show.ReadOnly)

			' Field news_show_home
			Call news_th.news_show_home.SetDbValue(Rs, news_th.news_show_home.CurrentValue, Null, news_th.news_show_home.ReadOnly)

			' Field news_create
			Call news_th.news_create.SetDbValue(Rs, news_th.news_create.CurrentValue, Null, news_th.news_create.ReadOnly)

			' Field news_update
			Call news_th.news_update.SetDbValue(Rs, news_th.news_update.CurrentValue, Null, news_th.news_update.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = news_th.Row_Updating(RsOld, Rs)
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
				ElseIf news_th.CancelMessage <> "" Then
					FailureMessage = news_th.CancelMessage
					news_th.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call news_th.Row_Updated(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", news_th.TableVar, "pom_news_thlist.asp", news_th.TableVar, True)
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
