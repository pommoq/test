<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_eventcalendarinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim eventcalendar_edit
Set eventcalendar_edit = New ceventcalendar_edit
Set Page = eventcalendar_edit

' Page init processing
eventcalendar_edit.Page_Init()

' Page main processing
eventcalendar_edit.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
eventcalendar_edit.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var eventcalendar_edit = new ew_Page("eventcalendar_edit");
eventcalendar_edit.PageID = "edit"; // Page ID
var EW_PAGE_ID = eventcalendar_edit.PageID; // For backward compatibility
// Form object
var feventcalendaredit = new ew_Form("feventcalendaredit");
// Validate form
feventcalendaredit.Validate = function() {
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
			elm = this.GetElements("x" + infix + "_eventcalendar_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(eventcalendar.eventcalendar_id.FldErrMsg) %>");
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
feventcalendaredit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
feventcalendaredit.ValidateRequired = true; // Use JavaScript validation
<% Else %>
feventcalendaredit.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If eventcalendar.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% eventcalendar_edit.ShowPageHeader() %>
<% eventcalendar_edit.ShowMessage %>
<form name="feventcalendaredit" id="feventcalendaredit" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="a_table" id="a_table" value="eventcalendar">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table class="ewGrid"><tr><td>
<table id="tbl_eventcalendaredit" class="table table-bordered table-striped">
<% If eventcalendar.eventcalendar_id.Visible Then ' eventcalendar_id %>
	<tr id="r_eventcalendar_id">
		<td><span id="elh_eventcalendar_eventcalendar_id"><%= eventcalendar.eventcalendar_id.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_id.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_id" class="control-group">
<span<%= eventcalendar.eventcalendar_id.ViewAttributes %>>
<%= eventcalendar.eventcalendar_id.EditValue %>
</span>
</span>
<input type="hidden" data-field="x_eventcalendar_id" name="x_eventcalendar_id" id="x_eventcalendar_id" value="<%= Server.HTMLEncode(eventcalendar.eventcalendar_id.CurrentValue&"") %>">
<%= eventcalendar.eventcalendar_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.eventcalendar_img.Visible Then ' eventcalendar_img %>
	<tr id="r_eventcalendar_img">
		<td><span id="elh_eventcalendar_eventcalendar_img"><%= eventcalendar.eventcalendar_img.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_img.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_img" class="control-group">
<textarea data-field="x_eventcalendar_img" name="x_eventcalendar_img" id="x_eventcalendar_img" cols="35" rows="4" placeholder="<%= eventcalendar.eventcalendar_img.PlaceHolder %>"<%= eventcalendar.eventcalendar_img.EditAttributes %>><%= eventcalendar.eventcalendar_img.EditValue %></textarea>
</span>
<%= eventcalendar.eventcalendar_img.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.eventcalendar_date.Visible Then ' eventcalendar_date %>
	<tr id="r_eventcalendar_date">
		<td><span id="elh_eventcalendar_eventcalendar_date"><%= eventcalendar.eventcalendar_date.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_date.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_date" class="control-group">
<input type="text" data-field="x_eventcalendar_date" name="x_eventcalendar_date" id="x_eventcalendar_date" placeholder="<%= eventcalendar.eventcalendar_date.PlaceHolder %>" value="<%= eventcalendar.eventcalendar_date.EditValue %>"<%= eventcalendar.eventcalendar_date.EditAttributes %>>
</span>
<%= eventcalendar.eventcalendar_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.eventcalendar_category.Visible Then ' eventcalendar_category %>
	<tr id="r_eventcalendar_category">
		<td><span id="elh_eventcalendar_eventcalendar_category"><%= eventcalendar.eventcalendar_category.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_category.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_category" class="control-group">
<input type="text" data-field="x_eventcalendar_category" name="x_eventcalendar_category" id="x_eventcalendar_category" size="30" maxlength="255" placeholder="<%= eventcalendar.eventcalendar_category.PlaceHolder %>" value="<%= eventcalendar.eventcalendar_category.EditValue %>"<%= eventcalendar.eventcalendar_category.EditAttributes %>>
</span>
<%= eventcalendar.eventcalendar_category.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.eventcalendar_category_sub.Visible Then ' eventcalendar_category_sub %>
	<tr id="r_eventcalendar_category_sub">
		<td><span id="elh_eventcalendar_eventcalendar_category_sub"><%= eventcalendar.eventcalendar_category_sub.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_category_sub.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_category_sub" class="control-group">
<input type="text" data-field="x_eventcalendar_category_sub" name="x_eventcalendar_category_sub" id="x_eventcalendar_category_sub" size="30" maxlength="255" placeholder="<%= eventcalendar.eventcalendar_category_sub.PlaceHolder %>" value="<%= eventcalendar.eventcalendar_category_sub.EditValue %>"<%= eventcalendar.eventcalendar_category_sub.EditAttributes %>>
</span>
<%= eventcalendar.eventcalendar_category_sub.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.start_date.Visible Then ' start_date %>
	<tr id="r_start_date">
		<td><span id="elh_eventcalendar_start_date"><%= eventcalendar.start_date.FldCaption %></span></td>
		<td<%= eventcalendar.start_date.CellAttributes %>>
<span id="el_eventcalendar_start_date" class="control-group">
<input type="text" data-field="x_start_date" name="x_start_date" id="x_start_date" placeholder="<%= eventcalendar.start_date.PlaceHolder %>" value="<%= eventcalendar.start_date.EditValue %>"<%= eventcalendar.start_date.EditAttributes %>>
</span>
<%= eventcalendar.start_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.end_date.Visible Then ' end_date %>
	<tr id="r_end_date">
		<td><span id="elh_eventcalendar_end_date"><%= eventcalendar.end_date.FldCaption %></span></td>
		<td<%= eventcalendar.end_date.CellAttributes %>>
<span id="el_eventcalendar_end_date" class="control-group">
<input type="text" data-field="x_end_date" name="x_end_date" id="x_end_date" placeholder="<%= eventcalendar.end_date.PlaceHolder %>" value="<%= eventcalendar.end_date.EditValue %>"<%= eventcalendar.end_date.EditAttributes %>>
</span>
<%= eventcalendar.end_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.eventcalendar_pdf.Visible Then ' eventcalendar_pdf %>
	<tr id="r_eventcalendar_pdf">
		<td><span id="elh_eventcalendar_eventcalendar_pdf"><%= eventcalendar.eventcalendar_pdf.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_pdf.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_pdf" class="control-group">
<input type="text" data-field="x_eventcalendar_pdf" name="x_eventcalendar_pdf" id="x_eventcalendar_pdf" size="30" maxlength="255" placeholder="<%= eventcalendar.eventcalendar_pdf.PlaceHolder %>" value="<%= eventcalendar.eventcalendar_pdf.EditValue %>"<%= eventcalendar.eventcalendar_pdf.EditAttributes %>>
</span>
<%= eventcalendar.eventcalendar_pdf.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.eventcalendar_subject.Visible Then ' eventcalendar_subject %>
	<tr id="r_eventcalendar_subject">
		<td><span id="elh_eventcalendar_eventcalendar_subject"><%= eventcalendar.eventcalendar_subject.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_subject.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_subject" class="control-group">
<input type="text" data-field="x_eventcalendar_subject" name="x_eventcalendar_subject" id="x_eventcalendar_subject" size="30" maxlength="255" placeholder="<%= eventcalendar.eventcalendar_subject.PlaceHolder %>" value="<%= eventcalendar.eventcalendar_subject.EditValue %>"<%= eventcalendar.eventcalendar_subject.EditAttributes %>>
</span>
<%= eventcalendar.eventcalendar_subject.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.eventcalendar_subject_th.Visible Then ' eventcalendar_subject_th %>
	<tr id="r_eventcalendar_subject_th">
		<td><span id="elh_eventcalendar_eventcalendar_subject_th"><%= eventcalendar.eventcalendar_subject_th.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_subject_th.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_subject_th" class="control-group">
<input type="text" data-field="x_eventcalendar_subject_th" name="x_eventcalendar_subject_th" id="x_eventcalendar_subject_th" size="30" maxlength="255" placeholder="<%= eventcalendar.eventcalendar_subject_th.PlaceHolder %>" value="<%= eventcalendar.eventcalendar_subject_th.EditValue %>"<%= eventcalendar.eventcalendar_subject_th.EditAttributes %>>
</span>
<%= eventcalendar.eventcalendar_subject_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.eventcalendar_intro.Visible Then ' eventcalendar_intro %>
	<tr id="r_eventcalendar_intro">
		<td><span id="elh_eventcalendar_eventcalendar_intro"><%= eventcalendar.eventcalendar_intro.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_intro.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_intro" class="control-group">
<textarea data-field="x_eventcalendar_intro" name="x_eventcalendar_intro" id="x_eventcalendar_intro" cols="35" rows="4" placeholder="<%= eventcalendar.eventcalendar_intro.PlaceHolder %>"<%= eventcalendar.eventcalendar_intro.EditAttributes %>><%= eventcalendar.eventcalendar_intro.EditValue %></textarea>
</span>
<%= eventcalendar.eventcalendar_intro.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.eventcalendar_intro_th.Visible Then ' eventcalendar_intro_th %>
	<tr id="r_eventcalendar_intro_th">
		<td><span id="elh_eventcalendar_eventcalendar_intro_th"><%= eventcalendar.eventcalendar_intro_th.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_intro_th.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_intro_th" class="control-group">
<textarea data-field="x_eventcalendar_intro_th" name="x_eventcalendar_intro_th" id="x_eventcalendar_intro_th" cols="35" rows="4" placeholder="<%= eventcalendar.eventcalendar_intro_th.PlaceHolder %>"<%= eventcalendar.eventcalendar_intro_th.EditAttributes %>><%= eventcalendar.eventcalendar_intro_th.EditValue %></textarea>
</span>
<%= eventcalendar.eventcalendar_intro_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.eventcalendar_content.Visible Then ' eventcalendar_content %>
	<tr id="r_eventcalendar_content">
		<td><span id="elh_eventcalendar_eventcalendar_content"><%= eventcalendar.eventcalendar_content.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_content.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_content" class="control-group">
<textarea data-field="x_eventcalendar_content" name="x_eventcalendar_content" id="x_eventcalendar_content" cols="35" rows="4" placeholder="<%= eventcalendar.eventcalendar_content.PlaceHolder %>"<%= eventcalendar.eventcalendar_content.EditAttributes %>><%= eventcalendar.eventcalendar_content.EditValue %></textarea>
</span>
<%= eventcalendar.eventcalendar_content.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.eventcalendar_content_th.Visible Then ' eventcalendar_content_th %>
	<tr id="r_eventcalendar_content_th">
		<td><span id="elh_eventcalendar_eventcalendar_content_th"><%= eventcalendar.eventcalendar_content_th.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_content_th.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_content_th" class="control-group">
<textarea data-field="x_eventcalendar_content_th" name="x_eventcalendar_content_th" id="x_eventcalendar_content_th" cols="35" rows="4" placeholder="<%= eventcalendar.eventcalendar_content_th.PlaceHolder %>"<%= eventcalendar.eventcalendar_content_th.EditAttributes %>><%= eventcalendar.eventcalendar_content_th.EditValue %></textarea>
</span>
<%= eventcalendar.eventcalendar_content_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.eventcalendar_show_en.Visible Then ' eventcalendar_show_en %>
	<tr id="r_eventcalendar_show_en">
		<td><span id="elh_eventcalendar_eventcalendar_show_en"><%= eventcalendar.eventcalendar_show_en.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_show_en.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_show_en" class="control-group">
<input type="text" data-field="x_eventcalendar_show_en" name="x_eventcalendar_show_en" id="x_eventcalendar_show_en" size="30" maxlength="255" placeholder="<%= eventcalendar.eventcalendar_show_en.PlaceHolder %>" value="<%= eventcalendar.eventcalendar_show_en.EditValue %>"<%= eventcalendar.eventcalendar_show_en.EditAttributes %>>
</span>
<%= eventcalendar.eventcalendar_show_en.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.eventcalendar_show.Visible Then ' eventcalendar_show %>
	<tr id="r_eventcalendar_show">
		<td><span id="elh_eventcalendar_eventcalendar_show"><%= eventcalendar.eventcalendar_show.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_show.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_show" class="control-group">
<input type="text" data-field="x_eventcalendar_show" name="x_eventcalendar_show" id="x_eventcalendar_show" size="30" maxlength="255" placeholder="<%= eventcalendar.eventcalendar_show.PlaceHolder %>" value="<%= eventcalendar.eventcalendar_show.EditValue %>"<%= eventcalendar.eventcalendar_show.EditAttributes %>>
</span>
<%= eventcalendar.eventcalendar_show.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.eventcalendar_show_home.Visible Then ' eventcalendar_show_home %>
	<tr id="r_eventcalendar_show_home">
		<td><span id="elh_eventcalendar_eventcalendar_show_home"><%= eventcalendar.eventcalendar_show_home.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_show_home.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_show_home" class="control-group">
<input type="text" data-field="x_eventcalendar_show_home" name="x_eventcalendar_show_home" id="x_eventcalendar_show_home" size="30" maxlength="255" placeholder="<%= eventcalendar.eventcalendar_show_home.PlaceHolder %>" value="<%= eventcalendar.eventcalendar_show_home.EditValue %>"<%= eventcalendar.eventcalendar_show_home.EditAttributes %>>
</span>
<%= eventcalendar.eventcalendar_show_home.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.eventcalendar_create.Visible Then ' eventcalendar_create %>
	<tr id="r_eventcalendar_create">
		<td><span id="elh_eventcalendar_eventcalendar_create"><%= eventcalendar.eventcalendar_create.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_create.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_create" class="control-group">
<input type="text" data-field="x_eventcalendar_create" name="x_eventcalendar_create" id="x_eventcalendar_create" size="30" maxlength="255" placeholder="<%= eventcalendar.eventcalendar_create.PlaceHolder %>" value="<%= eventcalendar.eventcalendar_create.EditValue %>"<%= eventcalendar.eventcalendar_create.EditAttributes %>>
</span>
<%= eventcalendar.eventcalendar_create.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar.eventcalendar_update.Visible Then ' eventcalendar_update %>
	<tr id="r_eventcalendar_update">
		<td><span id="elh_eventcalendar_eventcalendar_update"><%= eventcalendar.eventcalendar_update.FldCaption %></span></td>
		<td<%= eventcalendar.eventcalendar_update.CellAttributes %>>
<span id="el_eventcalendar_eventcalendar_update" class="control-group">
<input type="text" data-field="x_eventcalendar_update" name="x_eventcalendar_update" id="x_eventcalendar_update" size="30" maxlength="255" placeholder="<%= eventcalendar.eventcalendar_update.PlaceHolder %>" value="<%= eventcalendar.eventcalendar_update.EditValue %>"<%= eventcalendar.eventcalendar_update.EditAttributes %>>
</span>
<%= eventcalendar.eventcalendar_update.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("EditBtn") %></button>
</form>
<script type="text/javascript">
feventcalendaredit.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
eventcalendar_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set eventcalendar_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class ceventcalendar_edit

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
		TableName = "eventcalendar"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "eventcalendar_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If eventcalendar.UseTokenInUrl Then PageUrl = PageUrl & "t=" & eventcalendar.TableVar & "&" ' add page token
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
		If eventcalendar.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (eventcalendar.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (eventcalendar.TableVar = Request.QueryString("t"))
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
		If IsEmpty(eventcalendar) Then Set eventcalendar = New ceventcalendar
		Set Table = eventcalendar

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "eventcalendar"

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

		eventcalendar.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set eventcalendar = Nothing
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
		If Request.QueryString("eventcalendar_id").Count > 0 Then
			eventcalendar.eventcalendar_id.QueryStringValue = Request.QueryString("eventcalendar_id")
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Process form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			eventcalendar.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values
		Else
			eventcalendar.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If eventcalendar.eventcalendar_id.CurrentValue = "" Then Call Page_Terminate("pom_eventcalendarlist.asp") ' Invalid key, return to list

		' Validate form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			If Not ValidateForm() Then
				eventcalendar.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				eventcalendar.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		End If
		Select Case eventcalendar.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_eventcalendarlist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				eventcalendar.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					sReturnUrl = eventcalendar.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					eventcalendar.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		eventcalendar.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call eventcalendar.ResetAttrs()
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
				eventcalendar.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					eventcalendar.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = eventcalendar.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			eventcalendar.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			eventcalendar.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			eventcalendar.StartRecordNumber = StartRec
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
		If Not eventcalendar.eventcalendar_id.FldIsDetailKey Then eventcalendar.eventcalendar_id.FormValue = ObjForm.GetValue("x_eventcalendar_id")
		If Not eventcalendar.eventcalendar_img.FldIsDetailKey Then eventcalendar.eventcalendar_img.FormValue = ObjForm.GetValue("x_eventcalendar_img")
		If Not eventcalendar.eventcalendar_date.FldIsDetailKey Then eventcalendar.eventcalendar_date.FormValue = ObjForm.GetValue("x_eventcalendar_date")
		If Not eventcalendar.eventcalendar_date.FldIsDetailKey Then eventcalendar.eventcalendar_date.CurrentValue = ew_UnFormatDateTime(eventcalendar.eventcalendar_date.CurrentValue, 8)
		If Not eventcalendar.eventcalendar_category.FldIsDetailKey Then eventcalendar.eventcalendar_category.FormValue = ObjForm.GetValue("x_eventcalendar_category")
		If Not eventcalendar.eventcalendar_category_sub.FldIsDetailKey Then eventcalendar.eventcalendar_category_sub.FormValue = ObjForm.GetValue("x_eventcalendar_category_sub")
		If Not eventcalendar.start_date.FldIsDetailKey Then eventcalendar.start_date.FormValue = ObjForm.GetValue("x_start_date")
		If Not eventcalendar.start_date.FldIsDetailKey Then eventcalendar.start_date.CurrentValue = ew_UnFormatDateTime(eventcalendar.start_date.CurrentValue, 8)
		If Not eventcalendar.end_date.FldIsDetailKey Then eventcalendar.end_date.FormValue = ObjForm.GetValue("x_end_date")
		If Not eventcalendar.end_date.FldIsDetailKey Then eventcalendar.end_date.CurrentValue = ew_UnFormatDateTime(eventcalendar.end_date.CurrentValue, 8)
		If Not eventcalendar.eventcalendar_pdf.FldIsDetailKey Then eventcalendar.eventcalendar_pdf.FormValue = ObjForm.GetValue("x_eventcalendar_pdf")
		If Not eventcalendar.eventcalendar_subject.FldIsDetailKey Then eventcalendar.eventcalendar_subject.FormValue = ObjForm.GetValue("x_eventcalendar_subject")
		If Not eventcalendar.eventcalendar_subject_th.FldIsDetailKey Then eventcalendar.eventcalendar_subject_th.FormValue = ObjForm.GetValue("x_eventcalendar_subject_th")
		If Not eventcalendar.eventcalendar_intro.FldIsDetailKey Then eventcalendar.eventcalendar_intro.FormValue = ObjForm.GetValue("x_eventcalendar_intro")
		If Not eventcalendar.eventcalendar_intro_th.FldIsDetailKey Then eventcalendar.eventcalendar_intro_th.FormValue = ObjForm.GetValue("x_eventcalendar_intro_th")
		If Not eventcalendar.eventcalendar_content.FldIsDetailKey Then eventcalendar.eventcalendar_content.FormValue = ObjForm.GetValue("x_eventcalendar_content")
		If Not eventcalendar.eventcalendar_content_th.FldIsDetailKey Then eventcalendar.eventcalendar_content_th.FormValue = ObjForm.GetValue("x_eventcalendar_content_th")
		If Not eventcalendar.eventcalendar_show_en.FldIsDetailKey Then eventcalendar.eventcalendar_show_en.FormValue = ObjForm.GetValue("x_eventcalendar_show_en")
		If Not eventcalendar.eventcalendar_show.FldIsDetailKey Then eventcalendar.eventcalendar_show.FormValue = ObjForm.GetValue("x_eventcalendar_show")
		If Not eventcalendar.eventcalendar_show_home.FldIsDetailKey Then eventcalendar.eventcalendar_show_home.FormValue = ObjForm.GetValue("x_eventcalendar_show_home")
		If Not eventcalendar.eventcalendar_create.FldIsDetailKey Then eventcalendar.eventcalendar_create.FormValue = ObjForm.GetValue("x_eventcalendar_create")
		If Not eventcalendar.eventcalendar_update.FldIsDetailKey Then eventcalendar.eventcalendar_update.FormValue = ObjForm.GetValue("x_eventcalendar_update")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		eventcalendar.eventcalendar_id.CurrentValue = eventcalendar.eventcalendar_id.FormValue
		eventcalendar.eventcalendar_img.CurrentValue = eventcalendar.eventcalendar_img.FormValue
		eventcalendar.eventcalendar_date.CurrentValue = eventcalendar.eventcalendar_date.FormValue
		eventcalendar.eventcalendar_date.CurrentValue = ew_UnFormatDateTime(eventcalendar.eventcalendar_date.CurrentValue, 8)
		eventcalendar.eventcalendar_category.CurrentValue = eventcalendar.eventcalendar_category.FormValue
		eventcalendar.eventcalendar_category_sub.CurrentValue = eventcalendar.eventcalendar_category_sub.FormValue
		eventcalendar.start_date.CurrentValue = eventcalendar.start_date.FormValue
		eventcalendar.start_date.CurrentValue = ew_UnFormatDateTime(eventcalendar.start_date.CurrentValue, 8)
		eventcalendar.end_date.CurrentValue = eventcalendar.end_date.FormValue
		eventcalendar.end_date.CurrentValue = ew_UnFormatDateTime(eventcalendar.end_date.CurrentValue, 8)
		eventcalendar.eventcalendar_pdf.CurrentValue = eventcalendar.eventcalendar_pdf.FormValue
		eventcalendar.eventcalendar_subject.CurrentValue = eventcalendar.eventcalendar_subject.FormValue
		eventcalendar.eventcalendar_subject_th.CurrentValue = eventcalendar.eventcalendar_subject_th.FormValue
		eventcalendar.eventcalendar_intro.CurrentValue = eventcalendar.eventcalendar_intro.FormValue
		eventcalendar.eventcalendar_intro_th.CurrentValue = eventcalendar.eventcalendar_intro_th.FormValue
		eventcalendar.eventcalendar_content.CurrentValue = eventcalendar.eventcalendar_content.FormValue
		eventcalendar.eventcalendar_content_th.CurrentValue = eventcalendar.eventcalendar_content_th.FormValue
		eventcalendar.eventcalendar_show_en.CurrentValue = eventcalendar.eventcalendar_show_en.FormValue
		eventcalendar.eventcalendar_show.CurrentValue = eventcalendar.eventcalendar_show.FormValue
		eventcalendar.eventcalendar_show_home.CurrentValue = eventcalendar.eventcalendar_show_home.FormValue
		eventcalendar.eventcalendar_create.CurrentValue = eventcalendar.eventcalendar_create.FormValue
		eventcalendar.eventcalendar_update.CurrentValue = eventcalendar.eventcalendar_update.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = eventcalendar.KeyFilter

		' Call Row Selecting event
		Call eventcalendar.Row_Selecting(sFilter)

		' Load sql based on filter
		eventcalendar.CurrentFilter = sFilter
		sSql = eventcalendar.SQL
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
		Call eventcalendar.Row_Selected(RsRow)
		eventcalendar.eventcalendar_id.DbValue = RsRow("eventcalendar_id")
		eventcalendar.eventcalendar_img.DbValue = RsRow("eventcalendar_img")
		eventcalendar.eventcalendar_date.DbValue = RsRow("eventcalendar_date")
		eventcalendar.eventcalendar_category.DbValue = RsRow("eventcalendar_category")
		eventcalendar.eventcalendar_category_sub.DbValue = RsRow("eventcalendar_category_sub")
		eventcalendar.start_date.DbValue = RsRow("start_date")
		eventcalendar.end_date.DbValue = RsRow("end_date")
		eventcalendar.eventcalendar_pdf.DbValue = RsRow("eventcalendar_pdf")
		eventcalendar.eventcalendar_subject.DbValue = RsRow("eventcalendar_subject")
		eventcalendar.eventcalendar_subject_th.DbValue = RsRow("eventcalendar_subject_th")
		eventcalendar.eventcalendar_intro.DbValue = RsRow("eventcalendar_intro")
		eventcalendar.eventcalendar_intro_th.DbValue = RsRow("eventcalendar_intro_th")
		eventcalendar.eventcalendar_content.DbValue = RsRow("eventcalendar_content")
		eventcalendar.eventcalendar_content_th.DbValue = RsRow("eventcalendar_content_th")
		eventcalendar.eventcalendar_show_en.DbValue = RsRow("eventcalendar_show_en")
		eventcalendar.eventcalendar_show.DbValue = RsRow("eventcalendar_show")
		eventcalendar.eventcalendar_show_home.DbValue = RsRow("eventcalendar_show_home")
		eventcalendar.eventcalendar_create.DbValue = RsRow("eventcalendar_create")
		eventcalendar.eventcalendar_update.DbValue = RsRow("eventcalendar_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		eventcalendar.eventcalendar_id.m_DbValue = Rs("eventcalendar_id")
		eventcalendar.eventcalendar_img.m_DbValue = Rs("eventcalendar_img")
		eventcalendar.eventcalendar_date.m_DbValue = Rs("eventcalendar_date")
		eventcalendar.eventcalendar_category.m_DbValue = Rs("eventcalendar_category")
		eventcalendar.eventcalendar_category_sub.m_DbValue = Rs("eventcalendar_category_sub")
		eventcalendar.start_date.m_DbValue = Rs("start_date")
		eventcalendar.end_date.m_DbValue = Rs("end_date")
		eventcalendar.eventcalendar_pdf.m_DbValue = Rs("eventcalendar_pdf")
		eventcalendar.eventcalendar_subject.m_DbValue = Rs("eventcalendar_subject")
		eventcalendar.eventcalendar_subject_th.m_DbValue = Rs("eventcalendar_subject_th")
		eventcalendar.eventcalendar_intro.m_DbValue = Rs("eventcalendar_intro")
		eventcalendar.eventcalendar_intro_th.m_DbValue = Rs("eventcalendar_intro_th")
		eventcalendar.eventcalendar_content.m_DbValue = Rs("eventcalendar_content")
		eventcalendar.eventcalendar_content_th.m_DbValue = Rs("eventcalendar_content_th")
		eventcalendar.eventcalendar_show_en.m_DbValue = Rs("eventcalendar_show_en")
		eventcalendar.eventcalendar_show.m_DbValue = Rs("eventcalendar_show")
		eventcalendar.eventcalendar_show_home.m_DbValue = Rs("eventcalendar_show_home")
		eventcalendar.eventcalendar_create.m_DbValue = Rs("eventcalendar_create")
		eventcalendar.eventcalendar_update.m_DbValue = Rs("eventcalendar_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call eventcalendar.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' eventcalendar_id
		' eventcalendar_img
		' eventcalendar_date
		' eventcalendar_category
		' eventcalendar_category_sub
		' start_date
		' end_date
		' eventcalendar_pdf
		' eventcalendar_subject
		' eventcalendar_subject_th
		' eventcalendar_intro
		' eventcalendar_intro_th
		' eventcalendar_content
		' eventcalendar_content_th
		' eventcalendar_show_en
		' eventcalendar_show
		' eventcalendar_show_home
		' eventcalendar_create
		' eventcalendar_update
		' -----------
		'  View  Row
		' -----------

		If eventcalendar.RowType = EW_ROWTYPE_VIEW Then ' View row

			' eventcalendar_id
			eventcalendar.eventcalendar_id.ViewValue = eventcalendar.eventcalendar_id.CurrentValue
			eventcalendar.eventcalendar_id.ViewCustomAttributes = ""

			' eventcalendar_img
			eventcalendar.eventcalendar_img.ViewValue = eventcalendar.eventcalendar_img.CurrentValue
			eventcalendar.eventcalendar_img.ViewCustomAttributes = ""

			' eventcalendar_date
			eventcalendar.eventcalendar_date.ViewValue = eventcalendar.eventcalendar_date.CurrentValue
			eventcalendar.eventcalendar_date.ViewCustomAttributes = ""

			' eventcalendar_category
			eventcalendar.eventcalendar_category.ViewValue = eventcalendar.eventcalendar_category.CurrentValue
			eventcalendar.eventcalendar_category.ViewCustomAttributes = ""

			' eventcalendar_category_sub
			eventcalendar.eventcalendar_category_sub.ViewValue = eventcalendar.eventcalendar_category_sub.CurrentValue
			eventcalendar.eventcalendar_category_sub.ViewCustomAttributes = ""

			' start_date
			eventcalendar.start_date.ViewValue = eventcalendar.start_date.CurrentValue
			eventcalendar.start_date.ViewCustomAttributes = ""

			' end_date
			eventcalendar.end_date.ViewValue = eventcalendar.end_date.CurrentValue
			eventcalendar.end_date.ViewCustomAttributes = ""

			' eventcalendar_pdf
			eventcalendar.eventcalendar_pdf.ViewValue = eventcalendar.eventcalendar_pdf.CurrentValue
			eventcalendar.eventcalendar_pdf.ViewCustomAttributes = ""

			' eventcalendar_subject
			eventcalendar.eventcalendar_subject.ViewValue = eventcalendar.eventcalendar_subject.CurrentValue
			eventcalendar.eventcalendar_subject.ViewCustomAttributes = ""

			' eventcalendar_subject_th
			eventcalendar.eventcalendar_subject_th.ViewValue = eventcalendar.eventcalendar_subject_th.CurrentValue
			eventcalendar.eventcalendar_subject_th.ViewCustomAttributes = ""

			' eventcalendar_intro
			eventcalendar.eventcalendar_intro.ViewValue = eventcalendar.eventcalendar_intro.CurrentValue
			eventcalendar.eventcalendar_intro.ViewCustomAttributes = ""

			' eventcalendar_intro_th
			eventcalendar.eventcalendar_intro_th.ViewValue = eventcalendar.eventcalendar_intro_th.CurrentValue
			eventcalendar.eventcalendar_intro_th.ViewCustomAttributes = ""

			' eventcalendar_content
			eventcalendar.eventcalendar_content.ViewValue = eventcalendar.eventcalendar_content.CurrentValue
			eventcalendar.eventcalendar_content.ViewCustomAttributes = ""

			' eventcalendar_content_th
			eventcalendar.eventcalendar_content_th.ViewValue = eventcalendar.eventcalendar_content_th.CurrentValue
			eventcalendar.eventcalendar_content_th.ViewCustomAttributes = ""

			' eventcalendar_show_en
			eventcalendar.eventcalendar_show_en.ViewValue = eventcalendar.eventcalendar_show_en.CurrentValue
			eventcalendar.eventcalendar_show_en.ViewCustomAttributes = ""

			' eventcalendar_show
			eventcalendar.eventcalendar_show.ViewValue = eventcalendar.eventcalendar_show.CurrentValue
			eventcalendar.eventcalendar_show.ViewCustomAttributes = ""

			' eventcalendar_show_home
			eventcalendar.eventcalendar_show_home.ViewValue = eventcalendar.eventcalendar_show_home.CurrentValue
			eventcalendar.eventcalendar_show_home.ViewCustomAttributes = ""

			' eventcalendar_create
			eventcalendar.eventcalendar_create.ViewValue = eventcalendar.eventcalendar_create.CurrentValue
			eventcalendar.eventcalendar_create.ViewCustomAttributes = ""

			' eventcalendar_update
			eventcalendar.eventcalendar_update.ViewValue = eventcalendar.eventcalendar_update.CurrentValue
			eventcalendar.eventcalendar_update.ViewCustomAttributes = ""

			' View refer script
			' eventcalendar_id

			eventcalendar.eventcalendar_id.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_id.HrefValue = ""
			eventcalendar.eventcalendar_id.TooltipValue = ""

			' eventcalendar_img
			eventcalendar.eventcalendar_img.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_img.HrefValue = ""
			eventcalendar.eventcalendar_img.TooltipValue = ""

			' eventcalendar_date
			eventcalendar.eventcalendar_date.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_date.HrefValue = ""
			eventcalendar.eventcalendar_date.TooltipValue = ""

			' eventcalendar_category
			eventcalendar.eventcalendar_category.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_category.HrefValue = ""
			eventcalendar.eventcalendar_category.TooltipValue = ""

			' eventcalendar_category_sub
			eventcalendar.eventcalendar_category_sub.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_category_sub.HrefValue = ""
			eventcalendar.eventcalendar_category_sub.TooltipValue = ""

			' start_date
			eventcalendar.start_date.LinkCustomAttributes = ""
			eventcalendar.start_date.HrefValue = ""
			eventcalendar.start_date.TooltipValue = ""

			' end_date
			eventcalendar.end_date.LinkCustomAttributes = ""
			eventcalendar.end_date.HrefValue = ""
			eventcalendar.end_date.TooltipValue = ""

			' eventcalendar_pdf
			eventcalendar.eventcalendar_pdf.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_pdf.HrefValue = ""
			eventcalendar.eventcalendar_pdf.TooltipValue = ""

			' eventcalendar_subject
			eventcalendar.eventcalendar_subject.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_subject.HrefValue = ""
			eventcalendar.eventcalendar_subject.TooltipValue = ""

			' eventcalendar_subject_th
			eventcalendar.eventcalendar_subject_th.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_subject_th.HrefValue = ""
			eventcalendar.eventcalendar_subject_th.TooltipValue = ""

			' eventcalendar_intro
			eventcalendar.eventcalendar_intro.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_intro.HrefValue = ""
			eventcalendar.eventcalendar_intro.TooltipValue = ""

			' eventcalendar_intro_th
			eventcalendar.eventcalendar_intro_th.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_intro_th.HrefValue = ""
			eventcalendar.eventcalendar_intro_th.TooltipValue = ""

			' eventcalendar_content
			eventcalendar.eventcalendar_content.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_content.HrefValue = ""
			eventcalendar.eventcalendar_content.TooltipValue = ""

			' eventcalendar_content_th
			eventcalendar.eventcalendar_content_th.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_content_th.HrefValue = ""
			eventcalendar.eventcalendar_content_th.TooltipValue = ""

			' eventcalendar_show_en
			eventcalendar.eventcalendar_show_en.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_show_en.HrefValue = ""
			eventcalendar.eventcalendar_show_en.TooltipValue = ""

			' eventcalendar_show
			eventcalendar.eventcalendar_show.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_show.HrefValue = ""
			eventcalendar.eventcalendar_show.TooltipValue = ""

			' eventcalendar_show_home
			eventcalendar.eventcalendar_show_home.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_show_home.HrefValue = ""
			eventcalendar.eventcalendar_show_home.TooltipValue = ""

			' eventcalendar_create
			eventcalendar.eventcalendar_create.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_create.HrefValue = ""
			eventcalendar.eventcalendar_create.TooltipValue = ""

			' eventcalendar_update
			eventcalendar.eventcalendar_update.LinkCustomAttributes = ""
			eventcalendar.eventcalendar_update.HrefValue = ""
			eventcalendar.eventcalendar_update.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf eventcalendar.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' eventcalendar_id
			eventcalendar.eventcalendar_id.EditCustomAttributes = ""
			eventcalendar.eventcalendar_id.EditValue = eventcalendar.eventcalendar_id.CurrentValue
			eventcalendar.eventcalendar_id.ViewCustomAttributes = ""

			' eventcalendar_img
			eventcalendar.eventcalendar_img.EditCustomAttributes = ""
			eventcalendar.eventcalendar_img.EditValue = eventcalendar.eventcalendar_img.CurrentValue
			eventcalendar.eventcalendar_img.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.eventcalendar_img.FldCaption))

			' eventcalendar_date
			eventcalendar.eventcalendar_date.EditCustomAttributes = ""
			eventcalendar.eventcalendar_date.EditValue = ew_HtmlEncode(eventcalendar.eventcalendar_date.CurrentValue)
			eventcalendar.eventcalendar_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.eventcalendar_date.FldCaption))

			' eventcalendar_category
			eventcalendar.eventcalendar_category.EditCustomAttributes = ""
			eventcalendar.eventcalendar_category.EditValue = ew_HtmlEncode(eventcalendar.eventcalendar_category.CurrentValue)
			eventcalendar.eventcalendar_category.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.eventcalendar_category.FldCaption))

			' eventcalendar_category_sub
			eventcalendar.eventcalendar_category_sub.EditCustomAttributes = ""
			eventcalendar.eventcalendar_category_sub.EditValue = ew_HtmlEncode(eventcalendar.eventcalendar_category_sub.CurrentValue)
			eventcalendar.eventcalendar_category_sub.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.eventcalendar_category_sub.FldCaption))

			' start_date
			eventcalendar.start_date.EditCustomAttributes = ""
			eventcalendar.start_date.EditValue = ew_HtmlEncode(eventcalendar.start_date.CurrentValue)
			eventcalendar.start_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.start_date.FldCaption))

			' end_date
			eventcalendar.end_date.EditCustomAttributes = ""
			eventcalendar.end_date.EditValue = ew_HtmlEncode(eventcalendar.end_date.CurrentValue)
			eventcalendar.end_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.end_date.FldCaption))

			' eventcalendar_pdf
			eventcalendar.eventcalendar_pdf.EditCustomAttributes = ""
			eventcalendar.eventcalendar_pdf.EditValue = ew_HtmlEncode(eventcalendar.eventcalendar_pdf.CurrentValue)
			eventcalendar.eventcalendar_pdf.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.eventcalendar_pdf.FldCaption))

			' eventcalendar_subject
			eventcalendar.eventcalendar_subject.EditCustomAttributes = ""
			eventcalendar.eventcalendar_subject.EditValue = ew_HtmlEncode(eventcalendar.eventcalendar_subject.CurrentValue)
			eventcalendar.eventcalendar_subject.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.eventcalendar_subject.FldCaption))

			' eventcalendar_subject_th
			eventcalendar.eventcalendar_subject_th.EditCustomAttributes = ""
			eventcalendar.eventcalendar_subject_th.EditValue = ew_HtmlEncode(eventcalendar.eventcalendar_subject_th.CurrentValue)
			eventcalendar.eventcalendar_subject_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.eventcalendar_subject_th.FldCaption))

			' eventcalendar_intro
			eventcalendar.eventcalendar_intro.EditCustomAttributes = ""
			eventcalendar.eventcalendar_intro.EditValue = eventcalendar.eventcalendar_intro.CurrentValue
			eventcalendar.eventcalendar_intro.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.eventcalendar_intro.FldCaption))

			' eventcalendar_intro_th
			eventcalendar.eventcalendar_intro_th.EditCustomAttributes = ""
			eventcalendar.eventcalendar_intro_th.EditValue = eventcalendar.eventcalendar_intro_th.CurrentValue
			eventcalendar.eventcalendar_intro_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.eventcalendar_intro_th.FldCaption))

			' eventcalendar_content
			eventcalendar.eventcalendar_content.EditCustomAttributes = ""
			eventcalendar.eventcalendar_content.EditValue = eventcalendar.eventcalendar_content.CurrentValue
			eventcalendar.eventcalendar_content.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.eventcalendar_content.FldCaption))

			' eventcalendar_content_th
			eventcalendar.eventcalendar_content_th.EditCustomAttributes = ""
			eventcalendar.eventcalendar_content_th.EditValue = eventcalendar.eventcalendar_content_th.CurrentValue
			eventcalendar.eventcalendar_content_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.eventcalendar_content_th.FldCaption))

			' eventcalendar_show_en
			eventcalendar.eventcalendar_show_en.EditCustomAttributes = ""
			eventcalendar.eventcalendar_show_en.EditValue = ew_HtmlEncode(eventcalendar.eventcalendar_show_en.CurrentValue)
			eventcalendar.eventcalendar_show_en.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.eventcalendar_show_en.FldCaption))

			' eventcalendar_show
			eventcalendar.eventcalendar_show.EditCustomAttributes = ""
			eventcalendar.eventcalendar_show.EditValue = ew_HtmlEncode(eventcalendar.eventcalendar_show.CurrentValue)
			eventcalendar.eventcalendar_show.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.eventcalendar_show.FldCaption))

			' eventcalendar_show_home
			eventcalendar.eventcalendar_show_home.EditCustomAttributes = ""
			eventcalendar.eventcalendar_show_home.EditValue = ew_HtmlEncode(eventcalendar.eventcalendar_show_home.CurrentValue)
			eventcalendar.eventcalendar_show_home.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.eventcalendar_show_home.FldCaption))

			' eventcalendar_create
			eventcalendar.eventcalendar_create.EditCustomAttributes = ""
			eventcalendar.eventcalendar_create.EditValue = ew_HtmlEncode(eventcalendar.eventcalendar_create.CurrentValue)
			eventcalendar.eventcalendar_create.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.eventcalendar_create.FldCaption))

			' eventcalendar_update
			eventcalendar.eventcalendar_update.EditCustomAttributes = ""
			eventcalendar.eventcalendar_update.EditValue = ew_HtmlEncode(eventcalendar.eventcalendar_update.CurrentValue)
			eventcalendar.eventcalendar_update.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar.eventcalendar_update.FldCaption))

			' Edit refer script
			' eventcalendar_id

			eventcalendar.eventcalendar_id.HrefValue = ""

			' eventcalendar_img
			eventcalendar.eventcalendar_img.HrefValue = ""

			' eventcalendar_date
			eventcalendar.eventcalendar_date.HrefValue = ""

			' eventcalendar_category
			eventcalendar.eventcalendar_category.HrefValue = ""

			' eventcalendar_category_sub
			eventcalendar.eventcalendar_category_sub.HrefValue = ""

			' start_date
			eventcalendar.start_date.HrefValue = ""

			' end_date
			eventcalendar.end_date.HrefValue = ""

			' eventcalendar_pdf
			eventcalendar.eventcalendar_pdf.HrefValue = ""

			' eventcalendar_subject
			eventcalendar.eventcalendar_subject.HrefValue = ""

			' eventcalendar_subject_th
			eventcalendar.eventcalendar_subject_th.HrefValue = ""

			' eventcalendar_intro
			eventcalendar.eventcalendar_intro.HrefValue = ""

			' eventcalendar_intro_th
			eventcalendar.eventcalendar_intro_th.HrefValue = ""

			' eventcalendar_content
			eventcalendar.eventcalendar_content.HrefValue = ""

			' eventcalendar_content_th
			eventcalendar.eventcalendar_content_th.HrefValue = ""

			' eventcalendar_show_en
			eventcalendar.eventcalendar_show_en.HrefValue = ""

			' eventcalendar_show
			eventcalendar.eventcalendar_show.HrefValue = ""

			' eventcalendar_show_home
			eventcalendar.eventcalendar_show_home.HrefValue = ""

			' eventcalendar_create
			eventcalendar.eventcalendar_create.HrefValue = ""

			' eventcalendar_update
			eventcalendar.eventcalendar_update.HrefValue = ""
		End If
		If eventcalendar.RowType = EW_ROWTYPE_ADD Or eventcalendar.RowType = EW_ROWTYPE_EDIT Or eventcalendar.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call eventcalendar.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If eventcalendar.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call eventcalendar.Row_Rendered()
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
		If Not ew_CheckInteger(eventcalendar.eventcalendar_id.FormValue) Then
			Call ew_AddMessage(gsFormError, eventcalendar.eventcalendar_id.FldErrMsg)
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
		sFilter = eventcalendar.KeyFilter
		eventcalendar.CurrentFilter  = sFilter
		sSql = eventcalendar.SQL
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

			' Field eventcalendar_id
			' Field eventcalendar_img

			Call eventcalendar.eventcalendar_img.SetDbValue(Rs, eventcalendar.eventcalendar_img.CurrentValue, Null, eventcalendar.eventcalendar_img.ReadOnly)

			' Field eventcalendar_date
			Call eventcalendar.eventcalendar_date.SetDbValue(Rs, eventcalendar.eventcalendar_date.CurrentValue, Null, eventcalendar.eventcalendar_date.ReadOnly)

			' Field eventcalendar_category
			Call eventcalendar.eventcalendar_category.SetDbValue(Rs, eventcalendar.eventcalendar_category.CurrentValue, Null, eventcalendar.eventcalendar_category.ReadOnly)

			' Field eventcalendar_category_sub
			Call eventcalendar.eventcalendar_category_sub.SetDbValue(Rs, eventcalendar.eventcalendar_category_sub.CurrentValue, Null, eventcalendar.eventcalendar_category_sub.ReadOnly)

			' Field start_date
			Call eventcalendar.start_date.SetDbValue(Rs, eventcalendar.start_date.CurrentValue, Null, eventcalendar.start_date.ReadOnly)

			' Field end_date
			Call eventcalendar.end_date.SetDbValue(Rs, eventcalendar.end_date.CurrentValue, Null, eventcalendar.end_date.ReadOnly)

			' Field eventcalendar_pdf
			Call eventcalendar.eventcalendar_pdf.SetDbValue(Rs, eventcalendar.eventcalendar_pdf.CurrentValue, Null, eventcalendar.eventcalendar_pdf.ReadOnly)

			' Field eventcalendar_subject
			Call eventcalendar.eventcalendar_subject.SetDbValue(Rs, eventcalendar.eventcalendar_subject.CurrentValue, Null, eventcalendar.eventcalendar_subject.ReadOnly)

			' Field eventcalendar_subject_th
			Call eventcalendar.eventcalendar_subject_th.SetDbValue(Rs, eventcalendar.eventcalendar_subject_th.CurrentValue, Null, eventcalendar.eventcalendar_subject_th.ReadOnly)

			' Field eventcalendar_intro
			Call eventcalendar.eventcalendar_intro.SetDbValue(Rs, eventcalendar.eventcalendar_intro.CurrentValue, Null, eventcalendar.eventcalendar_intro.ReadOnly)

			' Field eventcalendar_intro_th
			Call eventcalendar.eventcalendar_intro_th.SetDbValue(Rs, eventcalendar.eventcalendar_intro_th.CurrentValue, Null, eventcalendar.eventcalendar_intro_th.ReadOnly)

			' Field eventcalendar_content
			Call eventcalendar.eventcalendar_content.SetDbValue(Rs, eventcalendar.eventcalendar_content.CurrentValue, Null, eventcalendar.eventcalendar_content.ReadOnly)

			' Field eventcalendar_content_th
			Call eventcalendar.eventcalendar_content_th.SetDbValue(Rs, eventcalendar.eventcalendar_content_th.CurrentValue, Null, eventcalendar.eventcalendar_content_th.ReadOnly)

			' Field eventcalendar_show_en
			Call eventcalendar.eventcalendar_show_en.SetDbValue(Rs, eventcalendar.eventcalendar_show_en.CurrentValue, Null, eventcalendar.eventcalendar_show_en.ReadOnly)

			' Field eventcalendar_show
			Call eventcalendar.eventcalendar_show.SetDbValue(Rs, eventcalendar.eventcalendar_show.CurrentValue, Null, eventcalendar.eventcalendar_show.ReadOnly)

			' Field eventcalendar_show_home
			Call eventcalendar.eventcalendar_show_home.SetDbValue(Rs, eventcalendar.eventcalendar_show_home.CurrentValue, Null, eventcalendar.eventcalendar_show_home.ReadOnly)

			' Field eventcalendar_create
			Call eventcalendar.eventcalendar_create.SetDbValue(Rs, eventcalendar.eventcalendar_create.CurrentValue, Null, eventcalendar.eventcalendar_create.ReadOnly)

			' Field eventcalendar_update
			Call eventcalendar.eventcalendar_update.SetDbValue(Rs, eventcalendar.eventcalendar_update.CurrentValue, Null, eventcalendar.eventcalendar_update.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = eventcalendar.Row_Updating(RsOld, Rs)
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
				ElseIf eventcalendar.CancelMessage <> "" Then
					FailureMessage = eventcalendar.CancelMessage
					eventcalendar.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call eventcalendar.Row_Updated(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", eventcalendar.TableVar, "pom_eventcalendarlist.asp", eventcalendar.TableVar, True)
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
