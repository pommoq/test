<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_eventcalendar_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim eventcalendar_th_add
Set eventcalendar_th_add = New ceventcalendar_th_add
Set Page = eventcalendar_th_add

' Page init processing
eventcalendar_th_add.Page_Init()

' Page main processing
eventcalendar_th_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
eventcalendar_th_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var eventcalendar_th_add = new ew_Page("eventcalendar_th_add");
eventcalendar_th_add.PageID = "add"; // Page ID
var EW_PAGE_ID = eventcalendar_th_add.PageID; // For backward compatibility
// Form object
var feventcalendar_thadd = new ew_Form("feventcalendar_thadd");
// Validate form
feventcalendar_thadd.Validate = function() {
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
				return this.OnError(elm, "<%= ew_JsEncode2(eventcalendar_th.eventcalendar_id.FldErrMsg) %>");
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
feventcalendar_thadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
feventcalendar_thadd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
feventcalendar_thadd.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If eventcalendar_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% eventcalendar_th_add.ShowPageHeader() %>
<% eventcalendar_th_add.ShowMessage %>
<form name="feventcalendar_thadd" id="feventcalendar_thadd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="eventcalendar_th">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_eventcalendar_thadd" class="table table-bordered table-striped">
<% If eventcalendar_th.eventcalendar_id.Visible Then ' eventcalendar_id %>
	<tr id="r_eventcalendar_id">
		<td><span id="elh_eventcalendar_th_eventcalendar_id"><%= eventcalendar_th.eventcalendar_id.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_id.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_id" class="control-group">
<input type="text" data-field="x_eventcalendar_id" name="x_eventcalendar_id" id="x_eventcalendar_id" size="30" placeholder="<%= eventcalendar_th.eventcalendar_id.PlaceHolder %>" value="<%= eventcalendar_th.eventcalendar_id.EditValue %>"<%= eventcalendar_th.eventcalendar_id.EditAttributes %>>
</span>
<%= eventcalendar_th.eventcalendar_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_img.Visible Then ' eventcalendar_img %>
	<tr id="r_eventcalendar_img">
		<td><span id="elh_eventcalendar_th_eventcalendar_img"><%= eventcalendar_th.eventcalendar_img.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_img.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_img" class="control-group">
<input type="text" data-field="x_eventcalendar_img" name="x_eventcalendar_img" id="x_eventcalendar_img" size="30" maxlength="255" placeholder="<%= eventcalendar_th.eventcalendar_img.PlaceHolder %>" value="<%= eventcalendar_th.eventcalendar_img.EditValue %>"<%= eventcalendar_th.eventcalendar_img.EditAttributes %>>
</span>
<%= eventcalendar_th.eventcalendar_img.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_date.Visible Then ' eventcalendar_date %>
	<tr id="r_eventcalendar_date">
		<td><span id="elh_eventcalendar_th_eventcalendar_date"><%= eventcalendar_th.eventcalendar_date.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_date.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_date" class="control-group">
<input type="text" data-field="x_eventcalendar_date" name="x_eventcalendar_date" id="x_eventcalendar_date" placeholder="<%= eventcalendar_th.eventcalendar_date.PlaceHolder %>" value="<%= eventcalendar_th.eventcalendar_date.EditValue %>"<%= eventcalendar_th.eventcalendar_date.EditAttributes %>>
</span>
<%= eventcalendar_th.eventcalendar_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_category.Visible Then ' eventcalendar_category %>
	<tr id="r_eventcalendar_category">
		<td><span id="elh_eventcalendar_th_eventcalendar_category"><%= eventcalendar_th.eventcalendar_category.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_category.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_category" class="control-group">
<input type="text" data-field="x_eventcalendar_category" name="x_eventcalendar_category" id="x_eventcalendar_category" size="30" maxlength="255" placeholder="<%= eventcalendar_th.eventcalendar_category.PlaceHolder %>" value="<%= eventcalendar_th.eventcalendar_category.EditValue %>"<%= eventcalendar_th.eventcalendar_category.EditAttributes %>>
</span>
<%= eventcalendar_th.eventcalendar_category.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_category_sub.Visible Then ' eventcalendar_category_sub %>
	<tr id="r_eventcalendar_category_sub">
		<td><span id="elh_eventcalendar_th_eventcalendar_category_sub"><%= eventcalendar_th.eventcalendar_category_sub.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_category_sub.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_category_sub" class="control-group">
<input type="text" data-field="x_eventcalendar_category_sub" name="x_eventcalendar_category_sub" id="x_eventcalendar_category_sub" size="30" maxlength="255" placeholder="<%= eventcalendar_th.eventcalendar_category_sub.PlaceHolder %>" value="<%= eventcalendar_th.eventcalendar_category_sub.EditValue %>"<%= eventcalendar_th.eventcalendar_category_sub.EditAttributes %>>
</span>
<%= eventcalendar_th.eventcalendar_category_sub.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.start_date.Visible Then ' start_date %>
	<tr id="r_start_date">
		<td><span id="elh_eventcalendar_th_start_date"><%= eventcalendar_th.start_date.FldCaption %></span></td>
		<td<%= eventcalendar_th.start_date.CellAttributes %>>
<span id="el_eventcalendar_th_start_date" class="control-group">
<input type="text" data-field="x_start_date" name="x_start_date" id="x_start_date" placeholder="<%= eventcalendar_th.start_date.PlaceHolder %>" value="<%= eventcalendar_th.start_date.EditValue %>"<%= eventcalendar_th.start_date.EditAttributes %>>
</span>
<%= eventcalendar_th.start_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.end_date.Visible Then ' end_date %>
	<tr id="r_end_date">
		<td><span id="elh_eventcalendar_th_end_date"><%= eventcalendar_th.end_date.FldCaption %></span></td>
		<td<%= eventcalendar_th.end_date.CellAttributes %>>
<span id="el_eventcalendar_th_end_date" class="control-group">
<input type="text" data-field="x_end_date" name="x_end_date" id="x_end_date" placeholder="<%= eventcalendar_th.end_date.PlaceHolder %>" value="<%= eventcalendar_th.end_date.EditValue %>"<%= eventcalendar_th.end_date.EditAttributes %>>
</span>
<%= eventcalendar_th.end_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_pdf.Visible Then ' eventcalendar_pdf %>
	<tr id="r_eventcalendar_pdf">
		<td><span id="elh_eventcalendar_th_eventcalendar_pdf"><%= eventcalendar_th.eventcalendar_pdf.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_pdf.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_pdf" class="control-group">
<input type="text" data-field="x_eventcalendar_pdf" name="x_eventcalendar_pdf" id="x_eventcalendar_pdf" size="30" maxlength="255" placeholder="<%= eventcalendar_th.eventcalendar_pdf.PlaceHolder %>" value="<%= eventcalendar_th.eventcalendar_pdf.EditValue %>"<%= eventcalendar_th.eventcalendar_pdf.EditAttributes %>>
</span>
<%= eventcalendar_th.eventcalendar_pdf.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_subject.Visible Then ' eventcalendar_subject %>
	<tr id="r_eventcalendar_subject">
		<td><span id="elh_eventcalendar_th_eventcalendar_subject"><%= eventcalendar_th.eventcalendar_subject.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_subject.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_subject" class="control-group">
<input type="text" data-field="x_eventcalendar_subject" name="x_eventcalendar_subject" id="x_eventcalendar_subject" size="30" maxlength="255" placeholder="<%= eventcalendar_th.eventcalendar_subject.PlaceHolder %>" value="<%= eventcalendar_th.eventcalendar_subject.EditValue %>"<%= eventcalendar_th.eventcalendar_subject.EditAttributes %>>
</span>
<%= eventcalendar_th.eventcalendar_subject.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_subject_th.Visible Then ' eventcalendar_subject_th %>
	<tr id="r_eventcalendar_subject_th">
		<td><span id="elh_eventcalendar_th_eventcalendar_subject_th"><%= eventcalendar_th.eventcalendar_subject_th.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_subject_th.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_subject_th" class="control-group">
<input type="text" data-field="x_eventcalendar_subject_th" name="x_eventcalendar_subject_th" id="x_eventcalendar_subject_th" size="30" maxlength="255" placeholder="<%= eventcalendar_th.eventcalendar_subject_th.PlaceHolder %>" value="<%= eventcalendar_th.eventcalendar_subject_th.EditValue %>"<%= eventcalendar_th.eventcalendar_subject_th.EditAttributes %>>
</span>
<%= eventcalendar_th.eventcalendar_subject_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_intro.Visible Then ' eventcalendar_intro %>
	<tr id="r_eventcalendar_intro">
		<td><span id="elh_eventcalendar_th_eventcalendar_intro"><%= eventcalendar_th.eventcalendar_intro.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_intro.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_intro" class="control-group">
<textarea data-field="x_eventcalendar_intro" name="x_eventcalendar_intro" id="x_eventcalendar_intro" cols="35" rows="4" placeholder="<%= eventcalendar_th.eventcalendar_intro.PlaceHolder %>"<%= eventcalendar_th.eventcalendar_intro.EditAttributes %>><%= eventcalendar_th.eventcalendar_intro.EditValue %></textarea>
</span>
<%= eventcalendar_th.eventcalendar_intro.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_intro_th.Visible Then ' eventcalendar_intro_th %>
	<tr id="r_eventcalendar_intro_th">
		<td><span id="elh_eventcalendar_th_eventcalendar_intro_th"><%= eventcalendar_th.eventcalendar_intro_th.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_intro_th.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_intro_th" class="control-group">
<textarea data-field="x_eventcalendar_intro_th" name="x_eventcalendar_intro_th" id="x_eventcalendar_intro_th" cols="35" rows="4" placeholder="<%= eventcalendar_th.eventcalendar_intro_th.PlaceHolder %>"<%= eventcalendar_th.eventcalendar_intro_th.EditAttributes %>><%= eventcalendar_th.eventcalendar_intro_th.EditValue %></textarea>
</span>
<%= eventcalendar_th.eventcalendar_intro_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_content.Visible Then ' eventcalendar_content %>
	<tr id="r_eventcalendar_content">
		<td><span id="elh_eventcalendar_th_eventcalendar_content"><%= eventcalendar_th.eventcalendar_content.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_content.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_content" class="control-group">
<textarea data-field="x_eventcalendar_content" name="x_eventcalendar_content" id="x_eventcalendar_content" cols="35" rows="4" placeholder="<%= eventcalendar_th.eventcalendar_content.PlaceHolder %>"<%= eventcalendar_th.eventcalendar_content.EditAttributes %>><%= eventcalendar_th.eventcalendar_content.EditValue %></textarea>
</span>
<%= eventcalendar_th.eventcalendar_content.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_content_th.Visible Then ' eventcalendar_content_th %>
	<tr id="r_eventcalendar_content_th">
		<td><span id="elh_eventcalendar_th_eventcalendar_content_th"><%= eventcalendar_th.eventcalendar_content_th.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_content_th.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_content_th" class="control-group">
<textarea data-field="x_eventcalendar_content_th" name="x_eventcalendar_content_th" id="x_eventcalendar_content_th" cols="35" rows="4" placeholder="<%= eventcalendar_th.eventcalendar_content_th.PlaceHolder %>"<%= eventcalendar_th.eventcalendar_content_th.EditAttributes %>><%= eventcalendar_th.eventcalendar_content_th.EditValue %></textarea>
</span>
<%= eventcalendar_th.eventcalendar_content_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_show_en.Visible Then ' eventcalendar_show_en %>
	<tr id="r_eventcalendar_show_en">
		<td><span id="elh_eventcalendar_th_eventcalendar_show_en"><%= eventcalendar_th.eventcalendar_show_en.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_show_en.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_show_en" class="control-group">
<input type="text" data-field="x_eventcalendar_show_en" name="x_eventcalendar_show_en" id="x_eventcalendar_show_en" size="30" maxlength="255" placeholder="<%= eventcalendar_th.eventcalendar_show_en.PlaceHolder %>" value="<%= eventcalendar_th.eventcalendar_show_en.EditValue %>"<%= eventcalendar_th.eventcalendar_show_en.EditAttributes %>>
</span>
<%= eventcalendar_th.eventcalendar_show_en.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_show.Visible Then ' eventcalendar_show %>
	<tr id="r_eventcalendar_show">
		<td><span id="elh_eventcalendar_th_eventcalendar_show"><%= eventcalendar_th.eventcalendar_show.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_show.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_show" class="control-group">
<input type="text" data-field="x_eventcalendar_show" name="x_eventcalendar_show" id="x_eventcalendar_show" size="30" maxlength="255" placeholder="<%= eventcalendar_th.eventcalendar_show.PlaceHolder %>" value="<%= eventcalendar_th.eventcalendar_show.EditValue %>"<%= eventcalendar_th.eventcalendar_show.EditAttributes %>>
</span>
<%= eventcalendar_th.eventcalendar_show.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_show_home.Visible Then ' eventcalendar_show_home %>
	<tr id="r_eventcalendar_show_home">
		<td><span id="elh_eventcalendar_th_eventcalendar_show_home"><%= eventcalendar_th.eventcalendar_show_home.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_show_home.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_show_home" class="control-group">
<input type="text" data-field="x_eventcalendar_show_home" name="x_eventcalendar_show_home" id="x_eventcalendar_show_home" size="30" maxlength="255" placeholder="<%= eventcalendar_th.eventcalendar_show_home.PlaceHolder %>" value="<%= eventcalendar_th.eventcalendar_show_home.EditValue %>"<%= eventcalendar_th.eventcalendar_show_home.EditAttributes %>>
</span>
<%= eventcalendar_th.eventcalendar_show_home.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_create.Visible Then ' eventcalendar_create %>
	<tr id="r_eventcalendar_create">
		<td><span id="elh_eventcalendar_th_eventcalendar_create"><%= eventcalendar_th.eventcalendar_create.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_create.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_create" class="control-group">
<input type="text" data-field="x_eventcalendar_create" name="x_eventcalendar_create" id="x_eventcalendar_create" size="30" maxlength="255" placeholder="<%= eventcalendar_th.eventcalendar_create.PlaceHolder %>" value="<%= eventcalendar_th.eventcalendar_create.EditValue %>"<%= eventcalendar_th.eventcalendar_create.EditAttributes %>>
</span>
<%= eventcalendar_th.eventcalendar_create.CustomMsg %></td>
	</tr>
<% End If %>
<% If eventcalendar_th.eventcalendar_update.Visible Then ' eventcalendar_update %>
	<tr id="r_eventcalendar_update">
		<td><span id="elh_eventcalendar_th_eventcalendar_update"><%= eventcalendar_th.eventcalendar_update.FldCaption %></span></td>
		<td<%= eventcalendar_th.eventcalendar_update.CellAttributes %>>
<span id="el_eventcalendar_th_eventcalendar_update" class="control-group">
<input type="text" data-field="x_eventcalendar_update" name="x_eventcalendar_update" id="x_eventcalendar_update" size="30" maxlength="255" placeholder="<%= eventcalendar_th.eventcalendar_update.PlaceHolder %>" value="<%= eventcalendar_th.eventcalendar_update.EditValue %>"<%= eventcalendar_th.eventcalendar_update.EditAttributes %>>
</span>
<%= eventcalendar_th.eventcalendar_update.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
feventcalendar_thadd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
eventcalendar_th_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set eventcalendar_th_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class ceventcalendar_th_add

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
		TableName = "eventcalendar_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "eventcalendar_th_add"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If eventcalendar_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & eventcalendar_th.TableVar & "&" ' add page token
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
		If eventcalendar_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (eventcalendar_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (eventcalendar_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(eventcalendar_th) Then Set eventcalendar_th = New ceventcalendar_th
		Set Table = eventcalendar_th

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "add"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "eventcalendar_th"

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

		eventcalendar_th.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set eventcalendar_th = Nothing
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
			eventcalendar_th.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("eventcalendar_id").Count > 0 Then
				eventcalendar_th.eventcalendar_id.QueryStringValue = Request.QueryString("eventcalendar_id")
				Call eventcalendar_th.SetKey("eventcalendar_id", eventcalendar_th.eventcalendar_id.CurrentValue) ' Set up key
			Else
				Call eventcalendar_th.SetKey("eventcalendar_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				eventcalendar_th.CurrentAction = "C" ' Copy Record
			Else
				eventcalendar_th.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				eventcalendar_th.CurrentAction = "I" ' Form error, reset action
				eventcalendar_th.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case eventcalendar_th.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_eventcalendar_thlist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				eventcalendar_th.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = eventcalendar_th.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_eventcalendar_thview.asp" Then sReturnUrl = eventcalendar_th.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					eventcalendar_th.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		eventcalendar_th.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call eventcalendar_th.ResetAttrs()
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
		eventcalendar_th.eventcalendar_id.CurrentValue = Null
		eventcalendar_th.eventcalendar_id.OldValue = eventcalendar_th.eventcalendar_id.CurrentValue
		eventcalendar_th.eventcalendar_img.CurrentValue = Null
		eventcalendar_th.eventcalendar_img.OldValue = eventcalendar_th.eventcalendar_img.CurrentValue
		eventcalendar_th.eventcalendar_date.CurrentValue = Null
		eventcalendar_th.eventcalendar_date.OldValue = eventcalendar_th.eventcalendar_date.CurrentValue
		eventcalendar_th.eventcalendar_category.CurrentValue = Null
		eventcalendar_th.eventcalendar_category.OldValue = eventcalendar_th.eventcalendar_category.CurrentValue
		eventcalendar_th.eventcalendar_category_sub.CurrentValue = Null
		eventcalendar_th.eventcalendar_category_sub.OldValue = eventcalendar_th.eventcalendar_category_sub.CurrentValue
		eventcalendar_th.start_date.CurrentValue = Null
		eventcalendar_th.start_date.OldValue = eventcalendar_th.start_date.CurrentValue
		eventcalendar_th.end_date.CurrentValue = Null
		eventcalendar_th.end_date.OldValue = eventcalendar_th.end_date.CurrentValue
		eventcalendar_th.eventcalendar_pdf.CurrentValue = Null
		eventcalendar_th.eventcalendar_pdf.OldValue = eventcalendar_th.eventcalendar_pdf.CurrentValue
		eventcalendar_th.eventcalendar_subject.CurrentValue = Null
		eventcalendar_th.eventcalendar_subject.OldValue = eventcalendar_th.eventcalendar_subject.CurrentValue
		eventcalendar_th.eventcalendar_subject_th.CurrentValue = Null
		eventcalendar_th.eventcalendar_subject_th.OldValue = eventcalendar_th.eventcalendar_subject_th.CurrentValue
		eventcalendar_th.eventcalendar_intro.CurrentValue = Null
		eventcalendar_th.eventcalendar_intro.OldValue = eventcalendar_th.eventcalendar_intro.CurrentValue
		eventcalendar_th.eventcalendar_intro_th.CurrentValue = Null
		eventcalendar_th.eventcalendar_intro_th.OldValue = eventcalendar_th.eventcalendar_intro_th.CurrentValue
		eventcalendar_th.eventcalendar_content.CurrentValue = Null
		eventcalendar_th.eventcalendar_content.OldValue = eventcalendar_th.eventcalendar_content.CurrentValue
		eventcalendar_th.eventcalendar_content_th.CurrentValue = Null
		eventcalendar_th.eventcalendar_content_th.OldValue = eventcalendar_th.eventcalendar_content_th.CurrentValue
		eventcalendar_th.eventcalendar_show_en.CurrentValue = Null
		eventcalendar_th.eventcalendar_show_en.OldValue = eventcalendar_th.eventcalendar_show_en.CurrentValue
		eventcalendar_th.eventcalendar_show.CurrentValue = Null
		eventcalendar_th.eventcalendar_show.OldValue = eventcalendar_th.eventcalendar_show.CurrentValue
		eventcalendar_th.eventcalendar_show_home.CurrentValue = Null
		eventcalendar_th.eventcalendar_show_home.OldValue = eventcalendar_th.eventcalendar_show_home.CurrentValue
		eventcalendar_th.eventcalendar_create.CurrentValue = Null
		eventcalendar_th.eventcalendar_create.OldValue = eventcalendar_th.eventcalendar_create.CurrentValue
		eventcalendar_th.eventcalendar_update.CurrentValue = Null
		eventcalendar_th.eventcalendar_update.OldValue = eventcalendar_th.eventcalendar_update.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not eventcalendar_th.eventcalendar_id.FldIsDetailKey Then eventcalendar_th.eventcalendar_id.FormValue = ObjForm.GetValue("x_eventcalendar_id")
		If Not eventcalendar_th.eventcalendar_img.FldIsDetailKey Then eventcalendar_th.eventcalendar_img.FormValue = ObjForm.GetValue("x_eventcalendar_img")
		If Not eventcalendar_th.eventcalendar_date.FldIsDetailKey Then eventcalendar_th.eventcalendar_date.FormValue = ObjForm.GetValue("x_eventcalendar_date")
		If Not eventcalendar_th.eventcalendar_date.FldIsDetailKey Then eventcalendar_th.eventcalendar_date.CurrentValue = ew_UnFormatDateTime(eventcalendar_th.eventcalendar_date.CurrentValue, 8)
		If Not eventcalendar_th.eventcalendar_category.FldIsDetailKey Then eventcalendar_th.eventcalendar_category.FormValue = ObjForm.GetValue("x_eventcalendar_category")
		If Not eventcalendar_th.eventcalendar_category_sub.FldIsDetailKey Then eventcalendar_th.eventcalendar_category_sub.FormValue = ObjForm.GetValue("x_eventcalendar_category_sub")
		If Not eventcalendar_th.start_date.FldIsDetailKey Then eventcalendar_th.start_date.FormValue = ObjForm.GetValue("x_start_date")
		If Not eventcalendar_th.start_date.FldIsDetailKey Then eventcalendar_th.start_date.CurrentValue = ew_UnFormatDateTime(eventcalendar_th.start_date.CurrentValue, 8)
		If Not eventcalendar_th.end_date.FldIsDetailKey Then eventcalendar_th.end_date.FormValue = ObjForm.GetValue("x_end_date")
		If Not eventcalendar_th.end_date.FldIsDetailKey Then eventcalendar_th.end_date.CurrentValue = ew_UnFormatDateTime(eventcalendar_th.end_date.CurrentValue, 8)
		If Not eventcalendar_th.eventcalendar_pdf.FldIsDetailKey Then eventcalendar_th.eventcalendar_pdf.FormValue = ObjForm.GetValue("x_eventcalendar_pdf")
		If Not eventcalendar_th.eventcalendar_subject.FldIsDetailKey Then eventcalendar_th.eventcalendar_subject.FormValue = ObjForm.GetValue("x_eventcalendar_subject")
		If Not eventcalendar_th.eventcalendar_subject_th.FldIsDetailKey Then eventcalendar_th.eventcalendar_subject_th.FormValue = ObjForm.GetValue("x_eventcalendar_subject_th")
		If Not eventcalendar_th.eventcalendar_intro.FldIsDetailKey Then eventcalendar_th.eventcalendar_intro.FormValue = ObjForm.GetValue("x_eventcalendar_intro")
		If Not eventcalendar_th.eventcalendar_intro_th.FldIsDetailKey Then eventcalendar_th.eventcalendar_intro_th.FormValue = ObjForm.GetValue("x_eventcalendar_intro_th")
		If Not eventcalendar_th.eventcalendar_content.FldIsDetailKey Then eventcalendar_th.eventcalendar_content.FormValue = ObjForm.GetValue("x_eventcalendar_content")
		If Not eventcalendar_th.eventcalendar_content_th.FldIsDetailKey Then eventcalendar_th.eventcalendar_content_th.FormValue = ObjForm.GetValue("x_eventcalendar_content_th")
		If Not eventcalendar_th.eventcalendar_show_en.FldIsDetailKey Then eventcalendar_th.eventcalendar_show_en.FormValue = ObjForm.GetValue("x_eventcalendar_show_en")
		If Not eventcalendar_th.eventcalendar_show.FldIsDetailKey Then eventcalendar_th.eventcalendar_show.FormValue = ObjForm.GetValue("x_eventcalendar_show")
		If Not eventcalendar_th.eventcalendar_show_home.FldIsDetailKey Then eventcalendar_th.eventcalendar_show_home.FormValue = ObjForm.GetValue("x_eventcalendar_show_home")
		If Not eventcalendar_th.eventcalendar_create.FldIsDetailKey Then eventcalendar_th.eventcalendar_create.FormValue = ObjForm.GetValue("x_eventcalendar_create")
		If Not eventcalendar_th.eventcalendar_update.FldIsDetailKey Then eventcalendar_th.eventcalendar_update.FormValue = ObjForm.GetValue("x_eventcalendar_update")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		eventcalendar_th.eventcalendar_id.CurrentValue = eventcalendar_th.eventcalendar_id.FormValue
		eventcalendar_th.eventcalendar_img.CurrentValue = eventcalendar_th.eventcalendar_img.FormValue
		eventcalendar_th.eventcalendar_date.CurrentValue = eventcalendar_th.eventcalendar_date.FormValue
		eventcalendar_th.eventcalendar_date.CurrentValue = ew_UnFormatDateTime(eventcalendar_th.eventcalendar_date.CurrentValue, 8)
		eventcalendar_th.eventcalendar_category.CurrentValue = eventcalendar_th.eventcalendar_category.FormValue
		eventcalendar_th.eventcalendar_category_sub.CurrentValue = eventcalendar_th.eventcalendar_category_sub.FormValue
		eventcalendar_th.start_date.CurrentValue = eventcalendar_th.start_date.FormValue
		eventcalendar_th.start_date.CurrentValue = ew_UnFormatDateTime(eventcalendar_th.start_date.CurrentValue, 8)
		eventcalendar_th.end_date.CurrentValue = eventcalendar_th.end_date.FormValue
		eventcalendar_th.end_date.CurrentValue = ew_UnFormatDateTime(eventcalendar_th.end_date.CurrentValue, 8)
		eventcalendar_th.eventcalendar_pdf.CurrentValue = eventcalendar_th.eventcalendar_pdf.FormValue
		eventcalendar_th.eventcalendar_subject.CurrentValue = eventcalendar_th.eventcalendar_subject.FormValue
		eventcalendar_th.eventcalendar_subject_th.CurrentValue = eventcalendar_th.eventcalendar_subject_th.FormValue
		eventcalendar_th.eventcalendar_intro.CurrentValue = eventcalendar_th.eventcalendar_intro.FormValue
		eventcalendar_th.eventcalendar_intro_th.CurrentValue = eventcalendar_th.eventcalendar_intro_th.FormValue
		eventcalendar_th.eventcalendar_content.CurrentValue = eventcalendar_th.eventcalendar_content.FormValue
		eventcalendar_th.eventcalendar_content_th.CurrentValue = eventcalendar_th.eventcalendar_content_th.FormValue
		eventcalendar_th.eventcalendar_show_en.CurrentValue = eventcalendar_th.eventcalendar_show_en.FormValue
		eventcalendar_th.eventcalendar_show.CurrentValue = eventcalendar_th.eventcalendar_show.FormValue
		eventcalendar_th.eventcalendar_show_home.CurrentValue = eventcalendar_th.eventcalendar_show_home.FormValue
		eventcalendar_th.eventcalendar_create.CurrentValue = eventcalendar_th.eventcalendar_create.FormValue
		eventcalendar_th.eventcalendar_update.CurrentValue = eventcalendar_th.eventcalendar_update.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = eventcalendar_th.KeyFilter

		' Call Row Selecting event
		Call eventcalendar_th.Row_Selecting(sFilter)

		' Load sql based on filter
		eventcalendar_th.CurrentFilter = sFilter
		sSql = eventcalendar_th.SQL
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
		Call eventcalendar_th.Row_Selected(RsRow)
		eventcalendar_th.eventcalendar_id.DbValue = RsRow("eventcalendar_id")
		eventcalendar_th.eventcalendar_img.DbValue = RsRow("eventcalendar_img")
		eventcalendar_th.eventcalendar_date.DbValue = RsRow("eventcalendar_date")
		eventcalendar_th.eventcalendar_category.DbValue = RsRow("eventcalendar_category")
		eventcalendar_th.eventcalendar_category_sub.DbValue = RsRow("eventcalendar_category_sub")
		eventcalendar_th.start_date.DbValue = RsRow("start_date")
		eventcalendar_th.end_date.DbValue = RsRow("end_date")
		eventcalendar_th.eventcalendar_pdf.DbValue = RsRow("eventcalendar_pdf")
		eventcalendar_th.eventcalendar_subject.DbValue = RsRow("eventcalendar_subject")
		eventcalendar_th.eventcalendar_subject_th.DbValue = RsRow("eventcalendar_subject_th")
		eventcalendar_th.eventcalendar_intro.DbValue = RsRow("eventcalendar_intro")
		eventcalendar_th.eventcalendar_intro_th.DbValue = RsRow("eventcalendar_intro_th")
		eventcalendar_th.eventcalendar_content.DbValue = RsRow("eventcalendar_content")
		eventcalendar_th.eventcalendar_content_th.DbValue = RsRow("eventcalendar_content_th")
		eventcalendar_th.eventcalendar_show_en.DbValue = RsRow("eventcalendar_show_en")
		eventcalendar_th.eventcalendar_show.DbValue = RsRow("eventcalendar_show")
		eventcalendar_th.eventcalendar_show_home.DbValue = RsRow("eventcalendar_show_home")
		eventcalendar_th.eventcalendar_create.DbValue = RsRow("eventcalendar_create")
		eventcalendar_th.eventcalendar_update.DbValue = RsRow("eventcalendar_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		eventcalendar_th.eventcalendar_id.m_DbValue = Rs("eventcalendar_id")
		eventcalendar_th.eventcalendar_img.m_DbValue = Rs("eventcalendar_img")
		eventcalendar_th.eventcalendar_date.m_DbValue = Rs("eventcalendar_date")
		eventcalendar_th.eventcalendar_category.m_DbValue = Rs("eventcalendar_category")
		eventcalendar_th.eventcalendar_category_sub.m_DbValue = Rs("eventcalendar_category_sub")
		eventcalendar_th.start_date.m_DbValue = Rs("start_date")
		eventcalendar_th.end_date.m_DbValue = Rs("end_date")
		eventcalendar_th.eventcalendar_pdf.m_DbValue = Rs("eventcalendar_pdf")
		eventcalendar_th.eventcalendar_subject.m_DbValue = Rs("eventcalendar_subject")
		eventcalendar_th.eventcalendar_subject_th.m_DbValue = Rs("eventcalendar_subject_th")
		eventcalendar_th.eventcalendar_intro.m_DbValue = Rs("eventcalendar_intro")
		eventcalendar_th.eventcalendar_intro_th.m_DbValue = Rs("eventcalendar_intro_th")
		eventcalendar_th.eventcalendar_content.m_DbValue = Rs("eventcalendar_content")
		eventcalendar_th.eventcalendar_content_th.m_DbValue = Rs("eventcalendar_content_th")
		eventcalendar_th.eventcalendar_show_en.m_DbValue = Rs("eventcalendar_show_en")
		eventcalendar_th.eventcalendar_show.m_DbValue = Rs("eventcalendar_show")
		eventcalendar_th.eventcalendar_show_home.m_DbValue = Rs("eventcalendar_show_home")
		eventcalendar_th.eventcalendar_create.m_DbValue = Rs("eventcalendar_create")
		eventcalendar_th.eventcalendar_update.m_DbValue = Rs("eventcalendar_update")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If eventcalendar_th.GetKey("eventcalendar_id")&"" <> "" Then
			eventcalendar_th.eventcalendar_id.CurrentValue = eventcalendar_th.GetKey("eventcalendar_id") ' eventcalendar_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			eventcalendar_th.CurrentFilter = eventcalendar_th.KeyFilter
			Dim sSql
			sSql = eventcalendar_th.SQL
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

		Call eventcalendar_th.Row_Rendering()

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

		If eventcalendar_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' eventcalendar_id
			eventcalendar_th.eventcalendar_id.ViewValue = eventcalendar_th.eventcalendar_id.CurrentValue
			eventcalendar_th.eventcalendar_id.ViewCustomAttributes = ""

			' eventcalendar_img
			eventcalendar_th.eventcalendar_img.ViewValue = eventcalendar_th.eventcalendar_img.CurrentValue
			eventcalendar_th.eventcalendar_img.ViewCustomAttributes = ""

			' eventcalendar_date
			eventcalendar_th.eventcalendar_date.ViewValue = eventcalendar_th.eventcalendar_date.CurrentValue
			eventcalendar_th.eventcalendar_date.ViewCustomAttributes = ""

			' eventcalendar_category
			eventcalendar_th.eventcalendar_category.ViewValue = eventcalendar_th.eventcalendar_category.CurrentValue
			eventcalendar_th.eventcalendar_category.ViewCustomAttributes = ""

			' eventcalendar_category_sub
			eventcalendar_th.eventcalendar_category_sub.ViewValue = eventcalendar_th.eventcalendar_category_sub.CurrentValue
			eventcalendar_th.eventcalendar_category_sub.ViewCustomAttributes = ""

			' start_date
			eventcalendar_th.start_date.ViewValue = eventcalendar_th.start_date.CurrentValue
			eventcalendar_th.start_date.ViewCustomAttributes = ""

			' end_date
			eventcalendar_th.end_date.ViewValue = eventcalendar_th.end_date.CurrentValue
			eventcalendar_th.end_date.ViewCustomAttributes = ""

			' eventcalendar_pdf
			eventcalendar_th.eventcalendar_pdf.ViewValue = eventcalendar_th.eventcalendar_pdf.CurrentValue
			eventcalendar_th.eventcalendar_pdf.ViewCustomAttributes = ""

			' eventcalendar_subject
			eventcalendar_th.eventcalendar_subject.ViewValue = eventcalendar_th.eventcalendar_subject.CurrentValue
			eventcalendar_th.eventcalendar_subject.ViewCustomAttributes = ""

			' eventcalendar_subject_th
			eventcalendar_th.eventcalendar_subject_th.ViewValue = eventcalendar_th.eventcalendar_subject_th.CurrentValue
			eventcalendar_th.eventcalendar_subject_th.ViewCustomAttributes = ""

			' eventcalendar_intro
			eventcalendar_th.eventcalendar_intro.ViewValue = eventcalendar_th.eventcalendar_intro.CurrentValue
			eventcalendar_th.eventcalendar_intro.ViewCustomAttributes = ""

			' eventcalendar_intro_th
			eventcalendar_th.eventcalendar_intro_th.ViewValue = eventcalendar_th.eventcalendar_intro_th.CurrentValue
			eventcalendar_th.eventcalendar_intro_th.ViewCustomAttributes = ""

			' eventcalendar_content
			eventcalendar_th.eventcalendar_content.ViewValue = eventcalendar_th.eventcalendar_content.CurrentValue
			eventcalendar_th.eventcalendar_content.ViewCustomAttributes = ""

			' eventcalendar_content_th
			eventcalendar_th.eventcalendar_content_th.ViewValue = eventcalendar_th.eventcalendar_content_th.CurrentValue
			eventcalendar_th.eventcalendar_content_th.ViewCustomAttributes = ""

			' eventcalendar_show_en
			eventcalendar_th.eventcalendar_show_en.ViewValue = eventcalendar_th.eventcalendar_show_en.CurrentValue
			eventcalendar_th.eventcalendar_show_en.ViewCustomAttributes = ""

			' eventcalendar_show
			eventcalendar_th.eventcalendar_show.ViewValue = eventcalendar_th.eventcalendar_show.CurrentValue
			eventcalendar_th.eventcalendar_show.ViewCustomAttributes = ""

			' eventcalendar_show_home
			eventcalendar_th.eventcalendar_show_home.ViewValue = eventcalendar_th.eventcalendar_show_home.CurrentValue
			eventcalendar_th.eventcalendar_show_home.ViewCustomAttributes = ""

			' eventcalendar_create
			eventcalendar_th.eventcalendar_create.ViewValue = eventcalendar_th.eventcalendar_create.CurrentValue
			eventcalendar_th.eventcalendar_create.ViewCustomAttributes = ""

			' eventcalendar_update
			eventcalendar_th.eventcalendar_update.ViewValue = eventcalendar_th.eventcalendar_update.CurrentValue
			eventcalendar_th.eventcalendar_update.ViewCustomAttributes = ""

			' View refer script
			' eventcalendar_id

			eventcalendar_th.eventcalendar_id.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_id.HrefValue = ""
			eventcalendar_th.eventcalendar_id.TooltipValue = ""

			' eventcalendar_img
			eventcalendar_th.eventcalendar_img.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_img.HrefValue = ""
			eventcalendar_th.eventcalendar_img.TooltipValue = ""

			' eventcalendar_date
			eventcalendar_th.eventcalendar_date.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_date.HrefValue = ""
			eventcalendar_th.eventcalendar_date.TooltipValue = ""

			' eventcalendar_category
			eventcalendar_th.eventcalendar_category.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_category.HrefValue = ""
			eventcalendar_th.eventcalendar_category.TooltipValue = ""

			' eventcalendar_category_sub
			eventcalendar_th.eventcalendar_category_sub.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_category_sub.HrefValue = ""
			eventcalendar_th.eventcalendar_category_sub.TooltipValue = ""

			' start_date
			eventcalendar_th.start_date.LinkCustomAttributes = ""
			eventcalendar_th.start_date.HrefValue = ""
			eventcalendar_th.start_date.TooltipValue = ""

			' end_date
			eventcalendar_th.end_date.LinkCustomAttributes = ""
			eventcalendar_th.end_date.HrefValue = ""
			eventcalendar_th.end_date.TooltipValue = ""

			' eventcalendar_pdf
			eventcalendar_th.eventcalendar_pdf.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_pdf.HrefValue = ""
			eventcalendar_th.eventcalendar_pdf.TooltipValue = ""

			' eventcalendar_subject
			eventcalendar_th.eventcalendar_subject.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_subject.HrefValue = ""
			eventcalendar_th.eventcalendar_subject.TooltipValue = ""

			' eventcalendar_subject_th
			eventcalendar_th.eventcalendar_subject_th.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_subject_th.HrefValue = ""
			eventcalendar_th.eventcalendar_subject_th.TooltipValue = ""

			' eventcalendar_intro
			eventcalendar_th.eventcalendar_intro.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_intro.HrefValue = ""
			eventcalendar_th.eventcalendar_intro.TooltipValue = ""

			' eventcalendar_intro_th
			eventcalendar_th.eventcalendar_intro_th.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_intro_th.HrefValue = ""
			eventcalendar_th.eventcalendar_intro_th.TooltipValue = ""

			' eventcalendar_content
			eventcalendar_th.eventcalendar_content.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_content.HrefValue = ""
			eventcalendar_th.eventcalendar_content.TooltipValue = ""

			' eventcalendar_content_th
			eventcalendar_th.eventcalendar_content_th.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_content_th.HrefValue = ""
			eventcalendar_th.eventcalendar_content_th.TooltipValue = ""

			' eventcalendar_show_en
			eventcalendar_th.eventcalendar_show_en.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_show_en.HrefValue = ""
			eventcalendar_th.eventcalendar_show_en.TooltipValue = ""

			' eventcalendar_show
			eventcalendar_th.eventcalendar_show.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_show.HrefValue = ""
			eventcalendar_th.eventcalendar_show.TooltipValue = ""

			' eventcalendar_show_home
			eventcalendar_th.eventcalendar_show_home.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_show_home.HrefValue = ""
			eventcalendar_th.eventcalendar_show_home.TooltipValue = ""

			' eventcalendar_create
			eventcalendar_th.eventcalendar_create.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_create.HrefValue = ""
			eventcalendar_th.eventcalendar_create.TooltipValue = ""

			' eventcalendar_update
			eventcalendar_th.eventcalendar_update.LinkCustomAttributes = ""
			eventcalendar_th.eventcalendar_update.HrefValue = ""
			eventcalendar_th.eventcalendar_update.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf eventcalendar_th.RowType = EW_ROWTYPE_ADD Then ' Add row

			' eventcalendar_id
			eventcalendar_th.eventcalendar_id.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_id.EditValue = ew_HtmlEncode(eventcalendar_th.eventcalendar_id.CurrentValue)
			eventcalendar_th.eventcalendar_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_id.FldCaption))

			' eventcalendar_img
			eventcalendar_th.eventcalendar_img.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_img.EditValue = ew_HtmlEncode(eventcalendar_th.eventcalendar_img.CurrentValue)
			eventcalendar_th.eventcalendar_img.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_img.FldCaption))

			' eventcalendar_date
			eventcalendar_th.eventcalendar_date.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_date.EditValue = ew_HtmlEncode(eventcalendar_th.eventcalendar_date.CurrentValue)
			eventcalendar_th.eventcalendar_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_date.FldCaption))

			' eventcalendar_category
			eventcalendar_th.eventcalendar_category.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_category.EditValue = ew_HtmlEncode(eventcalendar_th.eventcalendar_category.CurrentValue)
			eventcalendar_th.eventcalendar_category.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_category.FldCaption))

			' eventcalendar_category_sub
			eventcalendar_th.eventcalendar_category_sub.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_category_sub.EditValue = ew_HtmlEncode(eventcalendar_th.eventcalendar_category_sub.CurrentValue)
			eventcalendar_th.eventcalendar_category_sub.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_category_sub.FldCaption))

			' start_date
			eventcalendar_th.start_date.EditCustomAttributes = ""
			eventcalendar_th.start_date.EditValue = ew_HtmlEncode(eventcalendar_th.start_date.CurrentValue)
			eventcalendar_th.start_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.start_date.FldCaption))

			' end_date
			eventcalendar_th.end_date.EditCustomAttributes = ""
			eventcalendar_th.end_date.EditValue = ew_HtmlEncode(eventcalendar_th.end_date.CurrentValue)
			eventcalendar_th.end_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.end_date.FldCaption))

			' eventcalendar_pdf
			eventcalendar_th.eventcalendar_pdf.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_pdf.EditValue = ew_HtmlEncode(eventcalendar_th.eventcalendar_pdf.CurrentValue)
			eventcalendar_th.eventcalendar_pdf.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_pdf.FldCaption))

			' eventcalendar_subject
			eventcalendar_th.eventcalendar_subject.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_subject.EditValue = ew_HtmlEncode(eventcalendar_th.eventcalendar_subject.CurrentValue)
			eventcalendar_th.eventcalendar_subject.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_subject.FldCaption))

			' eventcalendar_subject_th
			eventcalendar_th.eventcalendar_subject_th.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_subject_th.EditValue = ew_HtmlEncode(eventcalendar_th.eventcalendar_subject_th.CurrentValue)
			eventcalendar_th.eventcalendar_subject_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_subject_th.FldCaption))

			' eventcalendar_intro
			eventcalendar_th.eventcalendar_intro.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_intro.EditValue = eventcalendar_th.eventcalendar_intro.CurrentValue
			eventcalendar_th.eventcalendar_intro.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_intro.FldCaption))

			' eventcalendar_intro_th
			eventcalendar_th.eventcalendar_intro_th.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_intro_th.EditValue = eventcalendar_th.eventcalendar_intro_th.CurrentValue
			eventcalendar_th.eventcalendar_intro_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_intro_th.FldCaption))

			' eventcalendar_content
			eventcalendar_th.eventcalendar_content.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_content.EditValue = eventcalendar_th.eventcalendar_content.CurrentValue
			eventcalendar_th.eventcalendar_content.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_content.FldCaption))

			' eventcalendar_content_th
			eventcalendar_th.eventcalendar_content_th.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_content_th.EditValue = eventcalendar_th.eventcalendar_content_th.CurrentValue
			eventcalendar_th.eventcalendar_content_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_content_th.FldCaption))

			' eventcalendar_show_en
			eventcalendar_th.eventcalendar_show_en.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_show_en.EditValue = ew_HtmlEncode(eventcalendar_th.eventcalendar_show_en.CurrentValue)
			eventcalendar_th.eventcalendar_show_en.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_show_en.FldCaption))

			' eventcalendar_show
			eventcalendar_th.eventcalendar_show.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_show.EditValue = ew_HtmlEncode(eventcalendar_th.eventcalendar_show.CurrentValue)
			eventcalendar_th.eventcalendar_show.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_show.FldCaption))

			' eventcalendar_show_home
			eventcalendar_th.eventcalendar_show_home.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_show_home.EditValue = ew_HtmlEncode(eventcalendar_th.eventcalendar_show_home.CurrentValue)
			eventcalendar_th.eventcalendar_show_home.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_show_home.FldCaption))

			' eventcalendar_create
			eventcalendar_th.eventcalendar_create.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_create.EditValue = ew_HtmlEncode(eventcalendar_th.eventcalendar_create.CurrentValue)
			eventcalendar_th.eventcalendar_create.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_create.FldCaption))

			' eventcalendar_update
			eventcalendar_th.eventcalendar_update.EditCustomAttributes = ""
			eventcalendar_th.eventcalendar_update.EditValue = ew_HtmlEncode(eventcalendar_th.eventcalendar_update.CurrentValue)
			eventcalendar_th.eventcalendar_update.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(eventcalendar_th.eventcalendar_update.FldCaption))

			' Edit refer script
			' eventcalendar_id

			eventcalendar_th.eventcalendar_id.HrefValue = ""

			' eventcalendar_img
			eventcalendar_th.eventcalendar_img.HrefValue = ""

			' eventcalendar_date
			eventcalendar_th.eventcalendar_date.HrefValue = ""

			' eventcalendar_category
			eventcalendar_th.eventcalendar_category.HrefValue = ""

			' eventcalendar_category_sub
			eventcalendar_th.eventcalendar_category_sub.HrefValue = ""

			' start_date
			eventcalendar_th.start_date.HrefValue = ""

			' end_date
			eventcalendar_th.end_date.HrefValue = ""

			' eventcalendar_pdf
			eventcalendar_th.eventcalendar_pdf.HrefValue = ""

			' eventcalendar_subject
			eventcalendar_th.eventcalendar_subject.HrefValue = ""

			' eventcalendar_subject_th
			eventcalendar_th.eventcalendar_subject_th.HrefValue = ""

			' eventcalendar_intro
			eventcalendar_th.eventcalendar_intro.HrefValue = ""

			' eventcalendar_intro_th
			eventcalendar_th.eventcalendar_intro_th.HrefValue = ""

			' eventcalendar_content
			eventcalendar_th.eventcalendar_content.HrefValue = ""

			' eventcalendar_content_th
			eventcalendar_th.eventcalendar_content_th.HrefValue = ""

			' eventcalendar_show_en
			eventcalendar_th.eventcalendar_show_en.HrefValue = ""

			' eventcalendar_show
			eventcalendar_th.eventcalendar_show.HrefValue = ""

			' eventcalendar_show_home
			eventcalendar_th.eventcalendar_show_home.HrefValue = ""

			' eventcalendar_create
			eventcalendar_th.eventcalendar_create.HrefValue = ""

			' eventcalendar_update
			eventcalendar_th.eventcalendar_update.HrefValue = ""
		End If
		If eventcalendar_th.RowType = EW_ROWTYPE_ADD Or eventcalendar_th.RowType = EW_ROWTYPE_EDIT Or eventcalendar_th.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call eventcalendar_th.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If eventcalendar_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call eventcalendar_th.Row_Rendered()
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
		If Not ew_CheckInteger(eventcalendar_th.eventcalendar_id.FormValue) Then
			Call ew_AddMessage(gsFormError, eventcalendar_th.eventcalendar_id.FldErrMsg)
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
		If eventcalendar_th.eventcalendar_id.CurrentValue <> "" Then ' Check field with unique index
			sFilter = "([eventcalendar_id] = " & ew_AdjustSql(eventcalendar_th.eventcalendar_id.CurrentValue) & ")"
			Set RsChk = eventcalendar_th.LoadRs(sFilter)
			If Not (RsChk Is Nothing) Then
				sIdxErrMsg = Replace(Language.Phrase("DupIndex"), "%f", eventcalendar_th.eventcalendar_id.FldCaption)
				sIdxErrMsg = Replace(sIdxErrMsg, "%v", eventcalendar_th.eventcalendar_id.CurrentValue)
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
		eventcalendar_th.CurrentFilter = sFilter
		sSql = eventcalendar_th.SQL
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

		' Field eventcalendar_id
		Call eventcalendar_th.eventcalendar_id.SetDbValue(Rs, eventcalendar_th.eventcalendar_id.CurrentValue, Null, False)

		' Field eventcalendar_img
		Call eventcalendar_th.eventcalendar_img.SetDbValue(Rs, eventcalendar_th.eventcalendar_img.CurrentValue, Null, False)

		' Field eventcalendar_date
		Call eventcalendar_th.eventcalendar_date.SetDbValue(Rs, eventcalendar_th.eventcalendar_date.CurrentValue, Null, False)

		' Field eventcalendar_category
		Call eventcalendar_th.eventcalendar_category.SetDbValue(Rs, eventcalendar_th.eventcalendar_category.CurrentValue, Null, False)

		' Field eventcalendar_category_sub
		Call eventcalendar_th.eventcalendar_category_sub.SetDbValue(Rs, eventcalendar_th.eventcalendar_category_sub.CurrentValue, Null, False)

		' Field start_date
		Call eventcalendar_th.start_date.SetDbValue(Rs, eventcalendar_th.start_date.CurrentValue, Null, False)

		' Field end_date
		Call eventcalendar_th.end_date.SetDbValue(Rs, eventcalendar_th.end_date.CurrentValue, Null, False)

		' Field eventcalendar_pdf
		Call eventcalendar_th.eventcalendar_pdf.SetDbValue(Rs, eventcalendar_th.eventcalendar_pdf.CurrentValue, Null, False)

		' Field eventcalendar_subject
		Call eventcalendar_th.eventcalendar_subject.SetDbValue(Rs, eventcalendar_th.eventcalendar_subject.CurrentValue, Null, False)

		' Field eventcalendar_subject_th
		Call eventcalendar_th.eventcalendar_subject_th.SetDbValue(Rs, eventcalendar_th.eventcalendar_subject_th.CurrentValue, Null, False)

		' Field eventcalendar_intro
		Call eventcalendar_th.eventcalendar_intro.SetDbValue(Rs, eventcalendar_th.eventcalendar_intro.CurrentValue, Null, False)

		' Field eventcalendar_intro_th
		Call eventcalendar_th.eventcalendar_intro_th.SetDbValue(Rs, eventcalendar_th.eventcalendar_intro_th.CurrentValue, Null, False)

		' Field eventcalendar_content
		Call eventcalendar_th.eventcalendar_content.SetDbValue(Rs, eventcalendar_th.eventcalendar_content.CurrentValue, Null, False)

		' Field eventcalendar_content_th
		Call eventcalendar_th.eventcalendar_content_th.SetDbValue(Rs, eventcalendar_th.eventcalendar_content_th.CurrentValue, Null, False)

		' Field eventcalendar_show_en
		Call eventcalendar_th.eventcalendar_show_en.SetDbValue(Rs, eventcalendar_th.eventcalendar_show_en.CurrentValue, Null, False)

		' Field eventcalendar_show
		Call eventcalendar_th.eventcalendar_show.SetDbValue(Rs, eventcalendar_th.eventcalendar_show.CurrentValue, Null, False)

		' Field eventcalendar_show_home
		Call eventcalendar_th.eventcalendar_show_home.SetDbValue(Rs, eventcalendar_th.eventcalendar_show_home.CurrentValue, Null, False)

		' Field eventcalendar_create
		Call eventcalendar_th.eventcalendar_create.SetDbValue(Rs, eventcalendar_th.eventcalendar_create.CurrentValue, Null, False)

		' Field eventcalendar_update
		Call eventcalendar_th.eventcalendar_update.SetDbValue(Rs, eventcalendar_th.eventcalendar_update.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = eventcalendar_th.Row_Inserting(RsOld, Rs)

		' Check if key value entered
		If bInsertRow And eventcalendar_th.ValidateKey And eventcalendar_th.eventcalendar_id.CurrentValue = "" And eventcalendar_th.eventcalendar_id.SessionValue = "" Then
			FailureMessage = Language.Phrase("InvalidKeyValue")
			bInsertRow = False
		End If

		' Check for duplicate key
		Dim sKeyErrMsg
		If bInsertRow And eventcalendar_th.ValidateKey Then
			sFilter = eventcalendar_th.KeyFilter
			Set RsChk = eventcalendar_th.LoadRs(sFilter)
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
			ElseIf eventcalendar_th.CancelMessage <> "" Then
				FailureMessage = eventcalendar_th.CancelMessage
				eventcalendar_th.CancelMessage = ""
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
			Call eventcalendar_th.Row_Inserted(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", eventcalendar_th.TableVar, "pom_eventcalendar_thlist.asp", eventcalendar_th.TableVar, True)
		PageId = ew_IIf(eventcalendar_th.CurrentAction = "C", "Copy", "Add")
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
