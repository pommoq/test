<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_newsinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim news_add
Set news_add = New cnews_add
Set Page = news_add

' Page init processing
news_add.Page_Init()

' Page main processing
news_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
news_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var news_add = new ew_Page("news_add");
news_add.PageID = "add"; // Page ID
var EW_PAGE_ID = news_add.PageID; // For backward compatibility
// Form object
var fnewsadd = new ew_Form("fnewsadd");
// Validate form
fnewsadd.Validate = function() {
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
				return this.OnError(elm, "<%= ew_JsEncode2(news.news_id.FldErrMsg) %>");
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
fnewsadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fnewsadd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fnewsadd.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If news.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% news_add.ShowPageHeader() %>
<% news_add.ShowMessage %>
<form name="fnewsadd" id="fnewsadd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="news">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_newsadd" class="table table-bordered table-striped">
<% If news.news_id.Visible Then ' news_id %>
	<tr id="r_news_id">
		<td><span id="elh_news_news_id"><%= news.news_id.FldCaption %></span></td>
		<td<%= news.news_id.CellAttributes %>>
<span id="el_news_news_id" class="control-group">
<input type="text" data-field="x_news_id" name="x_news_id" id="x_news_id" size="30" placeholder="<%= news.news_id.PlaceHolder %>" value="<%= news.news_id.EditValue %>"<%= news.news_id.EditAttributes %>>
</span>
<%= news.news_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.news_img.Visible Then ' news_img %>
	<tr id="r_news_img">
		<td><span id="elh_news_news_img"><%= news.news_img.FldCaption %></span></td>
		<td<%= news.news_img.CellAttributes %>>
<div id="el_news_news_img" class="control-group">
<span id="fd_x_news_img">
<span class="btn btn-small fileinput-button"<% If news.news_img.ReadOnly Or news.news_img.Disabled Then Response.Write " style=""display: none;""" %>>
	<span><%= Language.Phrase("ChooseFile") %></span>
	<input type="file" data-field="x_news_img" name="x_news_img" id="x_news_img">
</span>
<input type="hidden" name="fn_x_news_img" id= "fn_x_news_img" value="<%= news.news_img.Upload.FileName %>">
<input type="hidden" name="fa_x_news_img" id= "fa_x_news_img" value="0">
<input type="hidden" name="fs_x_news_img" id= "fs_x_news_img" value="255">
</span>
<table id="ft_x_news_img" class="table table-condensed pull-left ewUploadTable"><tbody class="files"></tbody></table>
</div>
<%= news.news_img.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.news_date.Visible Then ' news_date %>
	<tr id="r_news_date">
		<td><span id="elh_news_news_date"><%= news.news_date.FldCaption %></span></td>
		<td<%= news.news_date.CellAttributes %>>
<span id="el_news_news_date" class="control-group">
<input type="text" data-field="x_news_date" name="x_news_date" id="x_news_date" placeholder="<%= news.news_date.PlaceHolder %>" value="<%= news.news_date.EditValue %>"<%= news.news_date.EditAttributes %>>
</span>
<%= news.news_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.news_category.Visible Then ' news_category %>
	<tr id="r_news_category">
		<td><span id="elh_news_news_category"><%= news.news_category.FldCaption %></span></td>
		<td<%= news.news_category.CellAttributes %>>
<span id="el_news_news_category" class="control-group">
<input type="text" data-field="x_news_category" name="x_news_category" id="x_news_category" size="30" maxlength="255" placeholder="<%= news.news_category.PlaceHolder %>" value="<%= news.news_category.EditValue %>"<%= news.news_category.EditAttributes %>>
</span>
<%= news.news_category.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.news_category_sub.Visible Then ' news_category_sub %>
	<tr id="r_news_category_sub">
		<td><span id="elh_news_news_category_sub"><%= news.news_category_sub.FldCaption %></span></td>
		<td<%= news.news_category_sub.CellAttributes %>>
<span id="el_news_news_category_sub" class="control-group">
<input type="text" data-field="x_news_category_sub" name="x_news_category_sub" id="x_news_category_sub" size="30" maxlength="255" placeholder="<%= news.news_category_sub.PlaceHolder %>" value="<%= news.news_category_sub.EditValue %>"<%= news.news_category_sub.EditAttributes %>>
</span>
<%= news.news_category_sub.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.start_date.Visible Then ' start_date %>
	<tr id="r_start_date">
		<td><span id="elh_news_start_date"><%= news.start_date.FldCaption %></span></td>
		<td<%= news.start_date.CellAttributes %>>
<span id="el_news_start_date" class="control-group">
<input type="text" data-field="x_start_date" name="x_start_date" id="x_start_date" placeholder="<%= news.start_date.PlaceHolder %>" value="<%= news.start_date.EditValue %>"<%= news.start_date.EditAttributes %>>
</span>
<%= news.start_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.end_date.Visible Then ' end_date %>
	<tr id="r_end_date">
		<td><span id="elh_news_end_date"><%= news.end_date.FldCaption %></span></td>
		<td<%= news.end_date.CellAttributes %>>
<span id="el_news_end_date" class="control-group">
<input type="text" data-field="x_end_date" name="x_end_date" id="x_end_date" placeholder="<%= news.end_date.PlaceHolder %>" value="<%= news.end_date.EditValue %>"<%= news.end_date.EditAttributes %>>
</span>
<%= news.end_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.news_pdf.Visible Then ' news_pdf %>
	<tr id="r_news_pdf">
		<td><span id="elh_news_news_pdf"><%= news.news_pdf.FldCaption %></span></td>
		<td<%= news.news_pdf.CellAttributes %>>
<span id="el_news_news_pdf" class="control-group">
<input type="text" data-field="x_news_pdf" name="x_news_pdf" id="x_news_pdf" size="30" maxlength="255" placeholder="<%= news.news_pdf.PlaceHolder %>" value="<%= news.news_pdf.EditValue %>"<%= news.news_pdf.EditAttributes %>>
</span>
<%= news.news_pdf.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.news_subject.Visible Then ' news_subject %>
	<tr id="r_news_subject">
		<td><span id="elh_news_news_subject"><%= news.news_subject.FldCaption %></span></td>
		<td<%= news.news_subject.CellAttributes %>>
<span id="el_news_news_subject" class="control-group">
<input type="text" data-field="x_news_subject" name="x_news_subject" id="x_news_subject" size="30" maxlength="255" placeholder="<%= news.news_subject.PlaceHolder %>" value="<%= news.news_subject.EditValue %>"<%= news.news_subject.EditAttributes %>>
</span>
<%= news.news_subject.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.news_subject_th.Visible Then ' news_subject_th %>
	<tr id="r_news_subject_th">
		<td><span id="elh_news_news_subject_th"><%= news.news_subject_th.FldCaption %></span></td>
		<td<%= news.news_subject_th.CellAttributes %>>
<span id="el_news_news_subject_th" class="control-group">
<input type="text" data-field="x_news_subject_th" name="x_news_subject_th" id="x_news_subject_th" size="30" maxlength="255" placeholder="<%= news.news_subject_th.PlaceHolder %>" value="<%= news.news_subject_th.EditValue %>"<%= news.news_subject_th.EditAttributes %>>
</span>
<%= news.news_subject_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.news_intro.Visible Then ' news_intro %>
	<tr id="r_news_intro">
		<td><span id="elh_news_news_intro"><%= news.news_intro.FldCaption %></span></td>
		<td<%= news.news_intro.CellAttributes %>>
<span id="el_news_news_intro" class="control-group">
<textarea data-field="x_news_intro" name="x_news_intro" id="x_news_intro" cols="35" rows="4" placeholder="<%= news.news_intro.PlaceHolder %>"<%= news.news_intro.EditAttributes %>><%= news.news_intro.EditValue %></textarea>
</span>
<%= news.news_intro.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.news_intro_th.Visible Then ' news_intro_th %>
	<tr id="r_news_intro_th">
		<td><span id="elh_news_news_intro_th"><%= news.news_intro_th.FldCaption %></span></td>
		<td<%= news.news_intro_th.CellAttributes %>>
<span id="el_news_news_intro_th" class="control-group">
<textarea data-field="x_news_intro_th" name="x_news_intro_th" id="x_news_intro_th" cols="35" rows="4" placeholder="<%= news.news_intro_th.PlaceHolder %>"<%= news.news_intro_th.EditAttributes %>><%= news.news_intro_th.EditValue %></textarea>
</span>
<%= news.news_intro_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.news_content.Visible Then ' news_content %>
	<tr id="r_news_content">
		<td><span id="elh_news_news_content"><%= news.news_content.FldCaption %></span></td>
		<td<%= news.news_content.CellAttributes %>>
<span id="el_news_news_content" class="control-group">
<textarea data-field="x_news_content" name="x_news_content" id="x_news_content" cols="35" rows="4" placeholder="<%= news.news_content.PlaceHolder %>"<%= news.news_content.EditAttributes %>><%= news.news_content.EditValue %></textarea>
</span>
<%= news.news_content.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.news_content_th.Visible Then ' news_content_th %>
	<tr id="r_news_content_th">
		<td><span id="elh_news_news_content_th"><%= news.news_content_th.FldCaption %></span></td>
		<td<%= news.news_content_th.CellAttributes %>>
<span id="el_news_news_content_th" class="control-group">
<textarea data-field="x_news_content_th" name="x_news_content_th" id="x_news_content_th" cols="35" rows="4" placeholder="<%= news.news_content_th.PlaceHolder %>"<%= news.news_content_th.EditAttributes %>><%= news.news_content_th.EditValue %></textarea>
</span>
<%= news.news_content_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.news_show_en.Visible Then ' news_show_en %>
	<tr id="r_news_show_en">
		<td><span id="elh_news_news_show_en"><%= news.news_show_en.FldCaption %></span></td>
		<td<%= news.news_show_en.CellAttributes %>>
<span id="el_news_news_show_en" class="control-group">
<input type="text" data-field="x_news_show_en" name="x_news_show_en" id="x_news_show_en" size="30" maxlength="255" placeholder="<%= news.news_show_en.PlaceHolder %>" value="<%= news.news_show_en.EditValue %>"<%= news.news_show_en.EditAttributes %>>
</span>
<%= news.news_show_en.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.news_show.Visible Then ' news_show %>
	<tr id="r_news_show">
		<td><span id="elh_news_news_show"><%= news.news_show.FldCaption %></span></td>
		<td<%= news.news_show.CellAttributes %>>
<span id="el_news_news_show" class="control-group">
<input type="text" data-field="x_news_show" name="x_news_show" id="x_news_show" size="30" maxlength="255" placeholder="<%= news.news_show.PlaceHolder %>" value="<%= news.news_show.EditValue %>"<%= news.news_show.EditAttributes %>>
</span>
<%= news.news_show.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.news_show_home.Visible Then ' news_show_home %>
	<tr id="r_news_show_home">
		<td><span id="elh_news_news_show_home"><%= news.news_show_home.FldCaption %></span></td>
		<td<%= news.news_show_home.CellAttributes %>>
<span id="el_news_news_show_home" class="control-group">
<input type="text" data-field="x_news_show_home" name="x_news_show_home" id="x_news_show_home" size="30" maxlength="255" placeholder="<%= news.news_show_home.PlaceHolder %>" value="<%= news.news_show_home.EditValue %>"<%= news.news_show_home.EditAttributes %>>
</span>
<%= news.news_show_home.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.news_create.Visible Then ' news_create %>
	<tr id="r_news_create">
		<td><span id="elh_news_news_create"><%= news.news_create.FldCaption %></span></td>
		<td<%= news.news_create.CellAttributes %>>
<span id="el_news_news_create" class="control-group">
<input type="text" data-field="x_news_create" name="x_news_create" id="x_news_create" size="30" maxlength="255" placeholder="<%= news.news_create.PlaceHolder %>" value="<%= news.news_create.EditValue %>"<%= news.news_create.EditAttributes %>>
</span>
<%= news.news_create.CustomMsg %></td>
	</tr>
<% End If %>
<% If news.news_update.Visible Then ' news_update %>
	<tr id="r_news_update">
		<td><span id="elh_news_news_update"><%= news.news_update.FldCaption %></span></td>
		<td<%= news.news_update.CellAttributes %>>
<span id="el_news_news_update" class="control-group">
<input type="text" data-field="x_news_update" name="x_news_update" id="x_news_update" size="30" maxlength="255" placeholder="<%= news.news_update.PlaceHolder %>" value="<%= news.news_update.EditValue %>"<%= news.news_update.EditAttributes %>>
</span>
<%= news.news_update.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
fnewsadd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
news_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set news_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cnews_add

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
		TableName = "news"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "news_add"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If news.UseTokenInUrl Then PageUrl = PageUrl & "t=" & news.TableVar & "&" ' add page token
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
		If news.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (news.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (news.TableVar = Request.QueryString("t"))
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
		If IsEmpty(news) Then Set news = New cnews
		Set Table = news

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "add"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "news"

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

		news.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set news = Nothing
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
			news.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("news_id").Count > 0 Then
				news.news_id.QueryStringValue = Request.QueryString("news_id")
				Call news.SetKey("news_id", news.news_id.CurrentValue) ' Set up key
			Else
				Call news.SetKey("news_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				news.CurrentAction = "C" ' Copy Record
			Else
				news.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				news.CurrentAction = "I" ' Form error, reset action
				news.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case news.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_newslist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				news.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = news.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_newsview.asp" Then sReturnUrl = news.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					news.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		news.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call news.ResetAttrs()
		Call RenderRow()
	End Sub

	' -----------------------------------------------------------------
	' Function Get upload files
	'
	Function GetUploadFiles()

		' Get upload data
		news.news_img.Upload.Index = ObjForm.Index
		If news.news_img.Upload.UploadFile() Then

			' No action required
		Else
			Response.Write news.news_img.Upload.Message
			Page_Terminate("")
			Response.End
		End If
		news.news_img.CurrentValue = news.news_img.Upload.FileName
	End Function

	' -----------------------------------------------------------------
	' Load default values
	'
	Function LoadDefaultValues()
		news.news_id.CurrentValue = Null
		news.news_id.OldValue = news.news_id.CurrentValue
		news.news_img.Upload.DbValue = Null
		news.news_img.OldValue = news.news_img.Upload.DbValue
		news.news_img.CurrentValue = Null ' Clear file related field
		news.news_date.CurrentValue = Null
		news.news_date.OldValue = news.news_date.CurrentValue
		news.news_category.CurrentValue = Null
		news.news_category.OldValue = news.news_category.CurrentValue
		news.news_category_sub.CurrentValue = Null
		news.news_category_sub.OldValue = news.news_category_sub.CurrentValue
		news.start_date.CurrentValue = Null
		news.start_date.OldValue = news.start_date.CurrentValue
		news.end_date.CurrentValue = Null
		news.end_date.OldValue = news.end_date.CurrentValue
		news.news_pdf.CurrentValue = Null
		news.news_pdf.OldValue = news.news_pdf.CurrentValue
		news.news_subject.CurrentValue = Null
		news.news_subject.OldValue = news.news_subject.CurrentValue
		news.news_subject_th.CurrentValue = Null
		news.news_subject_th.OldValue = news.news_subject_th.CurrentValue
		news.news_intro.CurrentValue = Null
		news.news_intro.OldValue = news.news_intro.CurrentValue
		news.news_intro_th.CurrentValue = Null
		news.news_intro_th.OldValue = news.news_intro_th.CurrentValue
		news.news_content.CurrentValue = Null
		news.news_content.OldValue = news.news_content.CurrentValue
		news.news_content_th.CurrentValue = Null
		news.news_content_th.OldValue = news.news_content_th.CurrentValue
		news.news_show_en.CurrentValue = Null
		news.news_show_en.OldValue = news.news_show_en.CurrentValue
		news.news_show.CurrentValue = Null
		news.news_show.OldValue = news.news_show.CurrentValue
		news.news_show_home.CurrentValue = Null
		news.news_show_home.OldValue = news.news_show_home.CurrentValue
		news.news_create.CurrentValue = Null
		news.news_create.OldValue = news.news_create.CurrentValue
		news.news_update.CurrentValue = Null
		news.news_update.OldValue = news.news_update.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		Call GetUploadFiles() ' Get upload files
		If Not news.news_id.FldIsDetailKey Then news.news_id.FormValue = ObjForm.GetValue("x_news_id")
		If Not news.news_date.FldIsDetailKey Then news.news_date.FormValue = ObjForm.GetValue("x_news_date")
		If Not news.news_date.FldIsDetailKey Then news.news_date.CurrentValue = ew_UnFormatDateTime(news.news_date.CurrentValue, 8)
		If Not news.news_category.FldIsDetailKey Then news.news_category.FormValue = ObjForm.GetValue("x_news_category")
		If Not news.news_category_sub.FldIsDetailKey Then news.news_category_sub.FormValue = ObjForm.GetValue("x_news_category_sub")
		If Not news.start_date.FldIsDetailKey Then news.start_date.FormValue = ObjForm.GetValue("x_start_date")
		If Not news.start_date.FldIsDetailKey Then news.start_date.CurrentValue = ew_UnFormatDateTime(news.start_date.CurrentValue, 8)
		If Not news.end_date.FldIsDetailKey Then news.end_date.FormValue = ObjForm.GetValue("x_end_date")
		If Not news.end_date.FldIsDetailKey Then news.end_date.CurrentValue = ew_UnFormatDateTime(news.end_date.CurrentValue, 8)
		If Not news.news_pdf.FldIsDetailKey Then news.news_pdf.FormValue = ObjForm.GetValue("x_news_pdf")
		If Not news.news_subject.FldIsDetailKey Then news.news_subject.FormValue = ObjForm.GetValue("x_news_subject")
		If Not news.news_subject_th.FldIsDetailKey Then news.news_subject_th.FormValue = ObjForm.GetValue("x_news_subject_th")
		If Not news.news_intro.FldIsDetailKey Then news.news_intro.FormValue = ObjForm.GetValue("x_news_intro")
		If Not news.news_intro_th.FldIsDetailKey Then news.news_intro_th.FormValue = ObjForm.GetValue("x_news_intro_th")
		If Not news.news_content.FldIsDetailKey Then news.news_content.FormValue = ObjForm.GetValue("x_news_content")
		If Not news.news_content_th.FldIsDetailKey Then news.news_content_th.FormValue = ObjForm.GetValue("x_news_content_th")
		If Not news.news_show_en.FldIsDetailKey Then news.news_show_en.FormValue = ObjForm.GetValue("x_news_show_en")
		If Not news.news_show.FldIsDetailKey Then news.news_show.FormValue = ObjForm.GetValue("x_news_show")
		If Not news.news_show_home.FldIsDetailKey Then news.news_show_home.FormValue = ObjForm.GetValue("x_news_show_home")
		If Not news.news_create.FldIsDetailKey Then news.news_create.FormValue = ObjForm.GetValue("x_news_create")
		If Not news.news_update.FldIsDetailKey Then news.news_update.FormValue = ObjForm.GetValue("x_news_update")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		news.news_id.CurrentValue = news.news_id.FormValue
		news.news_date.CurrentValue = news.news_date.FormValue
		news.news_date.CurrentValue = ew_UnFormatDateTime(news.news_date.CurrentValue, 8)
		news.news_category.CurrentValue = news.news_category.FormValue
		news.news_category_sub.CurrentValue = news.news_category_sub.FormValue
		news.start_date.CurrentValue = news.start_date.FormValue
		news.start_date.CurrentValue = ew_UnFormatDateTime(news.start_date.CurrentValue, 8)
		news.end_date.CurrentValue = news.end_date.FormValue
		news.end_date.CurrentValue = ew_UnFormatDateTime(news.end_date.CurrentValue, 8)
		news.news_pdf.CurrentValue = news.news_pdf.FormValue
		news.news_subject.CurrentValue = news.news_subject.FormValue
		news.news_subject_th.CurrentValue = news.news_subject_th.FormValue
		news.news_intro.CurrentValue = news.news_intro.FormValue
		news.news_intro_th.CurrentValue = news.news_intro_th.FormValue
		news.news_content.CurrentValue = news.news_content.FormValue
		news.news_content_th.CurrentValue = news.news_content_th.FormValue
		news.news_show_en.CurrentValue = news.news_show_en.FormValue
		news.news_show.CurrentValue = news.news_show.FormValue
		news.news_show_home.CurrentValue = news.news_show_home.FormValue
		news.news_create.CurrentValue = news.news_create.FormValue
		news.news_update.CurrentValue = news.news_update.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = news.KeyFilter

		' Call Row Selecting event
		Call news.Row_Selecting(sFilter)

		' Load sql based on filter
		news.CurrentFilter = sFilter
		sSql = news.SQL
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
		Call news.Row_Selected(RsRow)
		news.news_id.DbValue = RsRow("news_id")
		news.news_img.Upload.DbValue = RsRow("news_img")
		news.news_img.CurrentValue = news.news_img.Upload.DbValue
		news.news_date.DbValue = RsRow("news_date")
		news.news_category.DbValue = RsRow("news_category")
		news.news_category_sub.DbValue = RsRow("news_category_sub")
		news.start_date.DbValue = RsRow("start_date")
		news.end_date.DbValue = RsRow("end_date")
		news.news_pdf.DbValue = RsRow("news_pdf")
		news.news_subject.DbValue = RsRow("news_subject")
		news.news_subject_th.DbValue = RsRow("news_subject_th")
		news.news_intro.DbValue = RsRow("news_intro")
		news.news_intro_th.DbValue = RsRow("news_intro_th")
		news.news_content.DbValue = RsRow("news_content")
		news.news_content_th.DbValue = RsRow("news_content_th")
		news.news_show_en.DbValue = RsRow("news_show_en")
		news.news_show.DbValue = RsRow("news_show")
		news.news_show_home.DbValue = RsRow("news_show_home")
		news.news_create.DbValue = RsRow("news_create")
		news.news_update.DbValue = RsRow("news_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		news.news_id.m_DbValue = Rs("news_id")
		news.news_img.Upload.DbValue = Rs("news_img")
		news.news_date.m_DbValue = Rs("news_date")
		news.news_category.m_DbValue = Rs("news_category")
		news.news_category_sub.m_DbValue = Rs("news_category_sub")
		news.start_date.m_DbValue = Rs("start_date")
		news.end_date.m_DbValue = Rs("end_date")
		news.news_pdf.m_DbValue = Rs("news_pdf")
		news.news_subject.m_DbValue = Rs("news_subject")
		news.news_subject_th.m_DbValue = Rs("news_subject_th")
		news.news_intro.m_DbValue = Rs("news_intro")
		news.news_intro_th.m_DbValue = Rs("news_intro_th")
		news.news_content.m_DbValue = Rs("news_content")
		news.news_content_th.m_DbValue = Rs("news_content_th")
		news.news_show_en.m_DbValue = Rs("news_show_en")
		news.news_show.m_DbValue = Rs("news_show")
		news.news_show_home.m_DbValue = Rs("news_show_home")
		news.news_create.m_DbValue = Rs("news_create")
		news.news_update.m_DbValue = Rs("news_update")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If news.GetKey("news_id")&"" <> "" Then
			news.news_id.CurrentValue = news.GetKey("news_id") ' news_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			news.CurrentFilter = news.KeyFilter
			Dim sSql
			sSql = news.SQL
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

		Call news.Row_Rendering()

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

		If news.RowType = EW_ROWTYPE_VIEW Then ' View row

			' news_id
			news.news_id.ViewValue = news.news_id.CurrentValue
			news.news_id.ViewCustomAttributes = ""

			' news_img
			news.news_img.UploadPath = "./Upload/news"
			If Not ew_Empty(news.news_img.Upload.DbValue) Then
				news.news_img.ViewValue = news.news_img.Upload.DbValue
				news.news_img.ImageAlt = news.news_img.FldAlt
				news.news_img.ViewValue = ew_UploadPathEx(False, news.news_img.UploadPath) & news.news_img.Upload.DbValue
			Else
				news.news_img.ViewValue = ""
			End If
			news.news_img.ViewCustomAttributes = ""

			' news_date
			news.news_date.ViewValue = news.news_date.CurrentValue
			news.news_date.ViewCustomAttributes = ""

			' news_category
			news.news_category.ViewValue = news.news_category.CurrentValue
			news.news_category.ViewCustomAttributes = ""

			' news_category_sub
			news.news_category_sub.ViewValue = news.news_category_sub.CurrentValue
			news.news_category_sub.ViewCustomAttributes = ""

			' start_date
			news.start_date.ViewValue = news.start_date.CurrentValue
			news.start_date.ViewCustomAttributes = ""

			' end_date
			news.end_date.ViewValue = news.end_date.CurrentValue
			news.end_date.ViewCustomAttributes = ""

			' news_pdf
			news.news_pdf.ViewValue = news.news_pdf.CurrentValue
			news.news_pdf.ViewCustomAttributes = ""

			' news_subject
			news.news_subject.ViewValue = news.news_subject.CurrentValue
			news.news_subject.ViewCustomAttributes = ""

			' news_subject_th
			news.news_subject_th.ViewValue = news.news_subject_th.CurrentValue
			news.news_subject_th.ViewCustomAttributes = ""

			' news_intro
			news.news_intro.ViewValue = news.news_intro.CurrentValue
			news.news_intro.ViewCustomAttributes = ""

			' news_intro_th
			news.news_intro_th.ViewValue = news.news_intro_th.CurrentValue
			news.news_intro_th.ViewCustomAttributes = ""

			' news_content
			news.news_content.ViewValue = news.news_content.CurrentValue
			news.news_content.ViewCustomAttributes = ""

			' news_content_th
			news.news_content_th.ViewValue = news.news_content_th.CurrentValue
			news.news_content_th.ViewCustomAttributes = ""

			' news_show_en
			news.news_show_en.ViewValue = news.news_show_en.CurrentValue
			news.news_show_en.ViewCustomAttributes = ""

			' news_show
			news.news_show.ViewValue = news.news_show.CurrentValue
			news.news_show.ViewCustomAttributes = ""

			' news_show_home
			news.news_show_home.ViewValue = news.news_show_home.CurrentValue
			news.news_show_home.ViewCustomAttributes = ""

			' news_create
			news.news_create.ViewValue = news.news_create.CurrentValue
			news.news_create.ViewCustomAttributes = ""

			' news_update
			news.news_update.ViewValue = news.news_update.CurrentValue
			news.news_update.ViewCustomAttributes = ""

			' View refer script
			' news_id

			news.news_id.LinkCustomAttributes = ""
			news.news_id.HrefValue = ""
			news.news_id.TooltipValue = ""

			' news_img
			news.news_img.LinkCustomAttributes = ""
			news.news_img.HrefValue = ""
			news.news_img.HrefValue2 = news.news_img.UploadPath & news.news_img.Upload.DbValue
			news.news_img.TooltipValue = ""

			' news_date
			news.news_date.LinkCustomAttributes = ""
			news.news_date.HrefValue = ""
			news.news_date.TooltipValue = ""

			' news_category
			news.news_category.LinkCustomAttributes = ""
			news.news_category.HrefValue = ""
			news.news_category.TooltipValue = ""

			' news_category_sub
			news.news_category_sub.LinkCustomAttributes = ""
			news.news_category_sub.HrefValue = ""
			news.news_category_sub.TooltipValue = ""

			' start_date
			news.start_date.LinkCustomAttributes = ""
			news.start_date.HrefValue = ""
			news.start_date.TooltipValue = ""

			' end_date
			news.end_date.LinkCustomAttributes = ""
			news.end_date.HrefValue = ""
			news.end_date.TooltipValue = ""

			' news_pdf
			news.news_pdf.LinkCustomAttributes = ""
			news.news_pdf.HrefValue = ""
			news.news_pdf.TooltipValue = ""

			' news_subject
			news.news_subject.LinkCustomAttributes = ""
			news.news_subject.HrefValue = ""
			news.news_subject.TooltipValue = ""

			' news_subject_th
			news.news_subject_th.LinkCustomAttributes = ""
			news.news_subject_th.HrefValue = ""
			news.news_subject_th.TooltipValue = ""

			' news_intro
			news.news_intro.LinkCustomAttributes = ""
			news.news_intro.HrefValue = ""
			news.news_intro.TooltipValue = ""

			' news_intro_th
			news.news_intro_th.LinkCustomAttributes = ""
			news.news_intro_th.HrefValue = ""
			news.news_intro_th.TooltipValue = ""

			' news_content
			news.news_content.LinkCustomAttributes = ""
			news.news_content.HrefValue = ""
			news.news_content.TooltipValue = ""

			' news_content_th
			news.news_content_th.LinkCustomAttributes = ""
			news.news_content_th.HrefValue = ""
			news.news_content_th.TooltipValue = ""

			' news_show_en
			news.news_show_en.LinkCustomAttributes = ""
			news.news_show_en.HrefValue = ""
			news.news_show_en.TooltipValue = ""

			' news_show
			news.news_show.LinkCustomAttributes = ""
			news.news_show.HrefValue = ""
			news.news_show.TooltipValue = ""

			' news_show_home
			news.news_show_home.LinkCustomAttributes = ""
			news.news_show_home.HrefValue = ""
			news.news_show_home.TooltipValue = ""

			' news_create
			news.news_create.LinkCustomAttributes = ""
			news.news_create.HrefValue = ""
			news.news_create.TooltipValue = ""

			' news_update
			news.news_update.LinkCustomAttributes = ""
			news.news_update.HrefValue = ""
			news.news_update.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf news.RowType = EW_ROWTYPE_ADD Then ' Add row

			' news_id
			news.news_id.EditCustomAttributes = ""
			news.news_id.EditValue = ew_HtmlEncode(news.news_id.CurrentValue)
			news.news_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.news_id.FldCaption))

			' news_img
			news.news_img.EditCustomAttributes = ""
			news.news_img.UploadPath = "./Upload/news"
			If Not ew_Empty(news.news_img.Upload.DbValue) Then
				news.news_img.EditValue = news.news_img.Upload.DbValue
				news.news_img.ImageAlt = news.news_img.FldAlt
				news.news_img.EditValue = ew_UploadPathEx(False, news.news_img.UploadPath) & news.news_img.Upload.DbValue
			Else
				news.news_img.EditValue = ""
			End If
			If (news.CurrentAction = "I" Or news.CurrentAction = "C") And Not news.EventCancelled Then Call ew_RenderUploadField(news.news_img, -1)

			' news_date
			news.news_date.EditCustomAttributes = ""
			news.news_date.EditValue = ew_HtmlEncode(news.news_date.CurrentValue)
			news.news_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.news_date.FldCaption))

			' news_category
			news.news_category.EditCustomAttributes = ""
			news.news_category.EditValue = ew_HtmlEncode(news.news_category.CurrentValue)
			news.news_category.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.news_category.FldCaption))

			' news_category_sub
			news.news_category_sub.EditCustomAttributes = ""
			news.news_category_sub.EditValue = ew_HtmlEncode(news.news_category_sub.CurrentValue)
			news.news_category_sub.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.news_category_sub.FldCaption))

			' start_date
			news.start_date.EditCustomAttributes = ""
			news.start_date.EditValue = ew_HtmlEncode(news.start_date.CurrentValue)
			news.start_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.start_date.FldCaption))

			' end_date
			news.end_date.EditCustomAttributes = ""
			news.end_date.EditValue = ew_HtmlEncode(news.end_date.CurrentValue)
			news.end_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.end_date.FldCaption))

			' news_pdf
			news.news_pdf.EditCustomAttributes = ""
			news.news_pdf.EditValue = ew_HtmlEncode(news.news_pdf.CurrentValue)
			news.news_pdf.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.news_pdf.FldCaption))

			' news_subject
			news.news_subject.EditCustomAttributes = ""
			news.news_subject.EditValue = ew_HtmlEncode(news.news_subject.CurrentValue)
			news.news_subject.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.news_subject.FldCaption))

			' news_subject_th
			news.news_subject_th.EditCustomAttributes = ""
			news.news_subject_th.EditValue = ew_HtmlEncode(news.news_subject_th.CurrentValue)
			news.news_subject_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.news_subject_th.FldCaption))

			' news_intro
			news.news_intro.EditCustomAttributes = ""
			news.news_intro.EditValue = news.news_intro.CurrentValue
			news.news_intro.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.news_intro.FldCaption))

			' news_intro_th
			news.news_intro_th.EditCustomAttributes = ""
			news.news_intro_th.EditValue = news.news_intro_th.CurrentValue
			news.news_intro_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.news_intro_th.FldCaption))

			' news_content
			news.news_content.EditCustomAttributes = ""
			news.news_content.EditValue = news.news_content.CurrentValue
			news.news_content.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.news_content.FldCaption))

			' news_content_th
			news.news_content_th.EditCustomAttributes = ""
			news.news_content_th.EditValue = news.news_content_th.CurrentValue
			news.news_content_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.news_content_th.FldCaption))

			' news_show_en
			news.news_show_en.EditCustomAttributes = ""
			news.news_show_en.EditValue = ew_HtmlEncode(news.news_show_en.CurrentValue)
			news.news_show_en.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.news_show_en.FldCaption))

			' news_show
			news.news_show.EditCustomAttributes = ""
			news.news_show.EditValue = ew_HtmlEncode(news.news_show.CurrentValue)
			news.news_show.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.news_show.FldCaption))

			' news_show_home
			news.news_show_home.EditCustomAttributes = ""
			news.news_show_home.EditValue = ew_HtmlEncode(news.news_show_home.CurrentValue)
			news.news_show_home.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.news_show_home.FldCaption))

			' news_create
			news.news_create.EditCustomAttributes = ""
			news.news_create.EditValue = ew_HtmlEncode(news.news_create.CurrentValue)
			news.news_create.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.news_create.FldCaption))

			' news_update
			news.news_update.EditCustomAttributes = ""
			news.news_update.EditValue = ew_HtmlEncode(news.news_update.CurrentValue)
			news.news_update.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news.news_update.FldCaption))

			' Edit refer script
			' news_id

			news.news_id.HrefValue = ""

			' news_img
			news.news_img.HrefValue = ""
			news.news_img.HrefValue2 = news.news_img.UploadPath & news.news_img.Upload.DbValue

			' news_date
			news.news_date.HrefValue = ""

			' news_category
			news.news_category.HrefValue = ""

			' news_category_sub
			news.news_category_sub.HrefValue = ""

			' start_date
			news.start_date.HrefValue = ""

			' end_date
			news.end_date.HrefValue = ""

			' news_pdf
			news.news_pdf.HrefValue = ""

			' news_subject
			news.news_subject.HrefValue = ""

			' news_subject_th
			news.news_subject_th.HrefValue = ""

			' news_intro
			news.news_intro.HrefValue = ""

			' news_intro_th
			news.news_intro_th.HrefValue = ""

			' news_content
			news.news_content.HrefValue = ""

			' news_content_th
			news.news_content_th.HrefValue = ""

			' news_show_en
			news.news_show_en.HrefValue = ""

			' news_show
			news.news_show.HrefValue = ""

			' news_show_home
			news.news_show_home.HrefValue = ""

			' news_create
			news.news_create.HrefValue = ""

			' news_update
			news.news_update.HrefValue = ""
		End If
		If news.RowType = EW_ROWTYPE_ADD Or news.RowType = EW_ROWTYPE_EDIT Or news.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call news.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If news.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call news.Row_Rendered()
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
		If Not ew_CheckInteger(news.news_id.FormValue) Then
			Call ew_AddMessage(gsFormError, news.news_id.FldErrMsg)
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
		If news.news_id.CurrentValue <> "" Then ' Check field with unique index
			sFilter = "([news_id] = " & ew_AdjustSql(news.news_id.CurrentValue) & ")"
			Set RsChk = news.LoadRs(sFilter)
			If Not (RsChk Is Nothing) Then
				sIdxErrMsg = Replace(Language.Phrase("DupIndex"), "%f", news.news_id.FldCaption)
				sIdxErrMsg = Replace(sIdxErrMsg, "%v", news.news_id.CurrentValue)
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
			news.news_img.OldUploadPath = "./Upload/news"
			news.news_img.UploadPath = news.news_img.OldUploadPath
		End If

		' Add new record
		sFilter = "(0 = 1)"
		news.CurrentFilter = sFilter
		sSql = news.SQL
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

		' Field news_id
		Call news.news_id.SetDbValue(Rs, news.news_id.CurrentValue, Null, False)

		' Field news_img
		If Not news.news_img.Upload.KeepFile Then
			If news.news_img.Upload.FileName & "" = "" Then
				Rs("news_img") = Null
			Else
				Rs("news_img") = news.news_img.Upload.FileName
			End If
		End If

		' Field news_date
		Call news.news_date.SetDbValue(Rs, news.news_date.CurrentValue, Null, False)

		' Field news_category
		Call news.news_category.SetDbValue(Rs, news.news_category.CurrentValue, Null, False)

		' Field news_category_sub
		Call news.news_category_sub.SetDbValue(Rs, news.news_category_sub.CurrentValue, Null, False)

		' Field start_date
		Call news.start_date.SetDbValue(Rs, news.start_date.CurrentValue, Null, False)

		' Field end_date
		Call news.end_date.SetDbValue(Rs, news.end_date.CurrentValue, Null, False)

		' Field news_pdf
		Call news.news_pdf.SetDbValue(Rs, news.news_pdf.CurrentValue, Null, False)

		' Field news_subject
		Call news.news_subject.SetDbValue(Rs, news.news_subject.CurrentValue, Null, False)

		' Field news_subject_th
		Call news.news_subject_th.SetDbValue(Rs, news.news_subject_th.CurrentValue, Null, False)

		' Field news_intro
		Call news.news_intro.SetDbValue(Rs, news.news_intro.CurrentValue, Null, False)

		' Field news_intro_th
		Call news.news_intro_th.SetDbValue(Rs, news.news_intro_th.CurrentValue, Null, False)

		' Field news_content
		Call news.news_content.SetDbValue(Rs, news.news_content.CurrentValue, Null, False)

		' Field news_content_th
		Call news.news_content_th.SetDbValue(Rs, news.news_content_th.CurrentValue, Null, False)

		' Field news_show_en
		Call news.news_show_en.SetDbValue(Rs, news.news_show_en.CurrentValue, Null, False)

		' Field news_show
		Call news.news_show.SetDbValue(Rs, news.news_show.CurrentValue, Null, False)

		' Field news_show_home
		Call news.news_show_home.SetDbValue(Rs, news.news_show_home.CurrentValue, Null, False)

		' Field news_create
		Call news.news_create.SetDbValue(Rs, news.news_create.CurrentValue, Null, False)

		' Field news_update
		Call news.news_update.SetDbValue(Rs, news.news_update.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If
		If Not news.news_img.Upload.KeepFile Then
			news.news_img.UploadPath = "./Upload/news"
			If Not ew_Empty(news.news_img.Upload.Value) Then
				If news.news_img.Upload.FileName = news.news_img.Upload.DbValue Then ' Overwrite if same file name
					news.news_img.Upload.DbValue = "" ' No need to delete any more
				Else
					Rs("news_img") = ew_UploadFileNameEx(ew_UploadPathEx(True, news.news_img.UploadPath), Rs("news_img")) ' Get new file name
				End If
			End If
		End If

		' Call Row Inserting event
		bInsertRow = news.Row_Inserting(RsOld, Rs)

		' Check if key value entered
		If bInsertRow And news.ValidateKey And news.news_id.CurrentValue = "" And news.news_id.SessionValue = "" Then
			FailureMessage = Language.Phrase("InvalidKeyValue")
			bInsertRow = False
		End If

		' Check for duplicate key
		Dim sKeyErrMsg
		If bInsertRow And news.ValidateKey Then
			sFilter = news.KeyFilter
			Set RsChk = news.LoadRs(sFilter)
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
				If Not news.news_img.Upload.KeepFile Then
					If Not ew_Empty(news.news_img.Upload.Value) Then
						news.news_img.Upload.SaveToFile news.news_img.UploadPath, Rs("news_img"), True
					End If
					If news.news_img.Upload.DbValue <> "" Then ew_DeleteFile ew_UploadPathEx(True, news.news_img.OldUploadPath) & news.news_img.Upload.DbValue
				End If
			End If
		Else
			Rs.CancelUpdate

			' Set up error message
			If SuccessMessage <> "" Or FailureMessage <> "" Then

				' Use the message, do nothing
			ElseIf news.CancelMessage <> "" Then
				FailureMessage = news.CancelMessage
				news.CancelMessage = ""
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
			Call news.Row_Inserted(RsOld, RsNew)
		End If

		' news_img
		Call ew_CleanUploadTempPath(news.news_img, news.news_img.Upload.Index)
		If IsObject(RsNew) Then
			RsNew.Close
			Set RsNew = Nothing
		End If
	End Function

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", news.TableVar, "pom_newslist.asp", news.TableVar, True)
		PageId = ew_IIf(news.CurrentAction = "C", "Copy", "Add")
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
