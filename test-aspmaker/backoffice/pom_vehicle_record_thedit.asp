<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_vehicle_record_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim vehicle_record_th_edit
Set vehicle_record_th_edit = New cvehicle_record_th_edit
Set Page = vehicle_record_th_edit

' Page init processing
vehicle_record_th_edit.Page_Init()

' Page main processing
vehicle_record_th_edit.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
vehicle_record_th_edit.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var vehicle_record_th_edit = new ew_Page("vehicle_record_th_edit");
vehicle_record_th_edit.PageID = "edit"; // Page ID
var EW_PAGE_ID = vehicle_record_th_edit.PageID; // For backward compatibility
// Form object
var fvehicle_record_thedit = new ew_Form("fvehicle_record_thedit");
// Validate form
fvehicle_record_thedit.Validate = function() {
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
fvehicle_record_thedit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fvehicle_record_thedit.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fvehicle_record_thedit.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If vehicle_record_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% vehicle_record_th_edit.ShowPageHeader() %>
<% vehicle_record_th_edit.ShowMessage %>
<form name="fvehicle_record_thedit" id="fvehicle_record_thedit" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="a_table" id="a_table" value="vehicle_record_th">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table class="ewGrid"><tr><td>
<table id="tbl_vehicle_record_thedit" class="table table-bordered table-striped">
<% If vehicle_record_th.veh_id.Visible Then ' veh_id %>
	<tr id="r_veh_id">
		<td><span id="elh_vehicle_record_th_veh_id"><%= vehicle_record_th.veh_id.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_id.CellAttributes %>>
<span id="el_vehicle_record_th_veh_id" class="control-group">
<span<%= vehicle_record_th.veh_id.ViewAttributes %>>
<%= vehicle_record_th.veh_id.EditValue %>
</span>
</span>
<input type="hidden" data-field="x_veh_id" name="x_veh_id" id="x_veh_id" value="<%= Server.HTMLEncode(vehicle_record_th.veh_id.CurrentValue&"") %>">
<%= vehicle_record_th.veh_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.vch_month.Visible Then ' vch_month %>
	<tr id="r_vch_month">
		<td><span id="elh_vehicle_record_th_vch_month"><%= vehicle_record_th.vch_month.FldCaption %></span></td>
		<td<%= vehicle_record_th.vch_month.CellAttributes %>>
<span id="el_vehicle_record_th_vch_month" class="control-group">
<input type="text" data-field="x_vch_month" name="x_vch_month" id="x_vch_month" size="30" maxlength="255" placeholder="<%= vehicle_record_th.vch_month.PlaceHolder %>" value="<%= vehicle_record_th.vch_month.EditValue %>"<%= vehicle_record_th.vch_month.EditAttributes %>>
</span>
<%= vehicle_record_th.vch_month.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.vch_year.Visible Then ' vch_year %>
	<tr id="r_vch_year">
		<td><span id="elh_vehicle_record_th_vch_year"><%= vehicle_record_th.vch_year.FldCaption %></span></td>
		<td<%= vehicle_record_th.vch_year.CellAttributes %>>
<span id="el_vehicle_record_th_vch_year" class="control-group">
<input type="text" data-field="x_vch_year" name="x_vch_year" id="x_vch_year" size="30" maxlength="255" placeholder="<%= vehicle_record_th.vch_year.PlaceHolder %>" value="<%= vehicle_record_th.vch_year.EditValue %>"<%= vehicle_record_th.vch_year.EditAttributes %>>
</span>
<%= vehicle_record_th.vch_year.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_product_1.Visible Then ' veh_product_1 %>
	<tr id="r_veh_product_1">
		<td><span id="elh_vehicle_record_th_veh_product_1"><%= vehicle_record_th.veh_product_1.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_product_1.CellAttributes %>>
<span id="el_vehicle_record_th_veh_product_1" class="control-group">
<input type="text" data-field="x_veh_product_1" name="x_veh_product_1" id="x_veh_product_1" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_product_1.PlaceHolder %>" value="<%= vehicle_record_th.veh_product_1.EditValue %>"<%= vehicle_record_th.veh_product_1.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_product_1.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_product_2.Visible Then ' veh_product_2 %>
	<tr id="r_veh_product_2">
		<td><span id="elh_vehicle_record_th_veh_product_2"><%= vehicle_record_th.veh_product_2.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_product_2.CellAttributes %>>
<span id="el_vehicle_record_th_veh_product_2" class="control-group">
<input type="text" data-field="x_veh_product_2" name="x_veh_product_2" id="x_veh_product_2" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_product_2.PlaceHolder %>" value="<%= vehicle_record_th.veh_product_2.EditValue %>"<%= vehicle_record_th.veh_product_2.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_product_2.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_product_3.Visible Then ' veh_product_3 %>
	<tr id="r_veh_product_3">
		<td><span id="elh_vehicle_record_th_veh_product_3"><%= vehicle_record_th.veh_product_3.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_product_3.CellAttributes %>>
<span id="el_vehicle_record_th_veh_product_3" class="control-group">
<input type="text" data-field="x_veh_product_3" name="x_veh_product_3" id="x_veh_product_3" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_product_3.PlaceHolder %>" value="<%= vehicle_record_th.veh_product_3.EditValue %>"<%= vehicle_record_th.veh_product_3.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_product_3.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_product_4.Visible Then ' veh_product_4 %>
	<tr id="r_veh_product_4">
		<td><span id="elh_vehicle_record_th_veh_product_4"><%= vehicle_record_th.veh_product_4.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_product_4.CellAttributes %>>
<span id="el_vehicle_record_th_veh_product_4" class="control-group">
<input type="text" data-field="x_veh_product_4" name="x_veh_product_4" id="x_veh_product_4" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_product_4.PlaceHolder %>" value="<%= vehicle_record_th.veh_product_4.EditValue %>"<%= vehicle_record_th.veh_product_4.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_product_4.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_product_5.Visible Then ' veh_product_5 %>
	<tr id="r_veh_product_5">
		<td><span id="elh_vehicle_record_th_veh_product_5"><%= vehicle_record_th.veh_product_5.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_product_5.CellAttributes %>>
<span id="el_vehicle_record_th_veh_product_5" class="control-group">
<input type="text" data-field="x_veh_product_5" name="x_veh_product_5" id="x_veh_product_5" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_product_5.PlaceHolder %>" value="<%= vehicle_record_th.veh_product_5.EditValue %>"<%= vehicle_record_th.veh_product_5.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_product_5.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_product_6.Visible Then ' veh_product_6 %>
	<tr id="r_veh_product_6">
		<td><span id="elh_vehicle_record_th_veh_product_6"><%= vehicle_record_th.veh_product_6.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_product_6.CellAttributes %>>
<span id="el_vehicle_record_th_veh_product_6" class="control-group">
<input type="text" data-field="x_veh_product_6" name="x_veh_product_6" id="x_veh_product_6" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_product_6.PlaceHolder %>" value="<%= vehicle_record_th.veh_product_6.EditValue %>"<%= vehicle_record_th.veh_product_6.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_product_6.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_product_7.Visible Then ' veh_product_7 %>
	<tr id="r_veh_product_7">
		<td><span id="elh_vehicle_record_th_veh_product_7"><%= vehicle_record_th.veh_product_7.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_product_7.CellAttributes %>>
<span id="el_vehicle_record_th_veh_product_7" class="control-group">
<input type="text" data-field="x_veh_product_7" name="x_veh_product_7" id="x_veh_product_7" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_product_7.PlaceHolder %>" value="<%= vehicle_record_th.veh_product_7.EditValue %>"<%= vehicle_record_th.veh_product_7.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_product_7.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_product_8.Visible Then ' veh_product_8 %>
	<tr id="r_veh_product_8">
		<td><span id="elh_vehicle_record_th_veh_product_8"><%= vehicle_record_th.veh_product_8.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_product_8.CellAttributes %>>
<span id="el_vehicle_record_th_veh_product_8" class="control-group">
<input type="text" data-field="x_veh_product_8" name="x_veh_product_8" id="x_veh_product_8" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_product_8.PlaceHolder %>" value="<%= vehicle_record_th.veh_product_8.EditValue %>"<%= vehicle_record_th.veh_product_8.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_product_8.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_domes_1.Visible Then ' veh_domes_1 %>
	<tr id="r_veh_domes_1">
		<td><span id="elh_vehicle_record_th_veh_domes_1"><%= vehicle_record_th.veh_domes_1.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_domes_1.CellAttributes %>>
<span id="el_vehicle_record_th_veh_domes_1" class="control-group">
<input type="text" data-field="x_veh_domes_1" name="x_veh_domes_1" id="x_veh_domes_1" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_domes_1.PlaceHolder %>" value="<%= vehicle_record_th.veh_domes_1.EditValue %>"<%= vehicle_record_th.veh_domes_1.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_domes_1.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_domes_2.Visible Then ' veh_domes_2 %>
	<tr id="r_veh_domes_2">
		<td><span id="elh_vehicle_record_th_veh_domes_2"><%= vehicle_record_th.veh_domes_2.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_domes_2.CellAttributes %>>
<span id="el_vehicle_record_th_veh_domes_2" class="control-group">
<input type="text" data-field="x_veh_domes_2" name="x_veh_domes_2" id="x_veh_domes_2" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_domes_2.PlaceHolder %>" value="<%= vehicle_record_th.veh_domes_2.EditValue %>"<%= vehicle_record_th.veh_domes_2.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_domes_2.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_domes_3.Visible Then ' veh_domes_3 %>
	<tr id="r_veh_domes_3">
		<td><span id="elh_vehicle_record_th_veh_domes_3"><%= vehicle_record_th.veh_domes_3.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_domes_3.CellAttributes %>>
<span id="el_vehicle_record_th_veh_domes_3" class="control-group">
<input type="text" data-field="x_veh_domes_3" name="x_veh_domes_3" id="x_veh_domes_3" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_domes_3.PlaceHolder %>" value="<%= vehicle_record_th.veh_domes_3.EditValue %>"<%= vehicle_record_th.veh_domes_3.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_domes_3.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_domes_4.Visible Then ' veh_domes_4 %>
	<tr id="r_veh_domes_4">
		<td><span id="elh_vehicle_record_th_veh_domes_4"><%= vehicle_record_th.veh_domes_4.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_domes_4.CellAttributes %>>
<span id="el_vehicle_record_th_veh_domes_4" class="control-group">
<input type="text" data-field="x_veh_domes_4" name="x_veh_domes_4" id="x_veh_domes_4" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_domes_4.PlaceHolder %>" value="<%= vehicle_record_th.veh_domes_4.EditValue %>"<%= vehicle_record_th.veh_domes_4.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_domes_4.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_domes_5.Visible Then ' veh_domes_5 %>
	<tr id="r_veh_domes_5">
		<td><span id="elh_vehicle_record_th_veh_domes_5"><%= vehicle_record_th.veh_domes_5.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_domes_5.CellAttributes %>>
<span id="el_vehicle_record_th_veh_domes_5" class="control-group">
<input type="text" data-field="x_veh_domes_5" name="x_veh_domes_5" id="x_veh_domes_5" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_domes_5.PlaceHolder %>" value="<%= vehicle_record_th.veh_domes_5.EditValue %>"<%= vehicle_record_th.veh_domes_5.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_domes_5.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_domes_6.Visible Then ' veh_domes_6 %>
	<tr id="r_veh_domes_6">
		<td><span id="elh_vehicle_record_th_veh_domes_6"><%= vehicle_record_th.veh_domes_6.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_domes_6.CellAttributes %>>
<span id="el_vehicle_record_th_veh_domes_6" class="control-group">
<input type="text" data-field="x_veh_domes_6" name="x_veh_domes_6" id="x_veh_domes_6" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_domes_6.PlaceHolder %>" value="<%= vehicle_record_th.veh_domes_6.EditValue %>"<%= vehicle_record_th.veh_domes_6.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_domes_6.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_domes_7.Visible Then ' veh_domes_7 %>
	<tr id="r_veh_domes_7">
		<td><span id="elh_vehicle_record_th_veh_domes_7"><%= vehicle_record_th.veh_domes_7.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_domes_7.CellAttributes %>>
<span id="el_vehicle_record_th_veh_domes_7" class="control-group">
<input type="text" data-field="x_veh_domes_7" name="x_veh_domes_7" id="x_veh_domes_7" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_domes_7.PlaceHolder %>" value="<%= vehicle_record_th.veh_domes_7.EditValue %>"<%= vehicle_record_th.veh_domes_7.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_domes_7.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_domes_8.Visible Then ' veh_domes_8 %>
	<tr id="r_veh_domes_8">
		<td><span id="elh_vehicle_record_th_veh_domes_8"><%= vehicle_record_th.veh_domes_8.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_domes_8.CellAttributes %>>
<span id="el_vehicle_record_th_veh_domes_8" class="control-group">
<input type="text" data-field="x_veh_domes_8" name="x_veh_domes_8" id="x_veh_domes_8" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_domes_8.PlaceHolder %>" value="<%= vehicle_record_th.veh_domes_8.EditValue %>"<%= vehicle_record_th.veh_domes_8.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_domes_8.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_export_1.Visible Then ' veh_export_1 %>
	<tr id="r_veh_export_1">
		<td><span id="elh_vehicle_record_th_veh_export_1"><%= vehicle_record_th.veh_export_1.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_export_1.CellAttributes %>>
<span id="el_vehicle_record_th_veh_export_1" class="control-group">
<input type="text" data-field="x_veh_export_1" name="x_veh_export_1" id="x_veh_export_1" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_export_1.PlaceHolder %>" value="<%= vehicle_record_th.veh_export_1.EditValue %>"<%= vehicle_record_th.veh_export_1.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_export_1.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_export_2.Visible Then ' veh_export_2 %>
	<tr id="r_veh_export_2">
		<td><span id="elh_vehicle_record_th_veh_export_2"><%= vehicle_record_th.veh_export_2.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_export_2.CellAttributes %>>
<span id="el_vehicle_record_th_veh_export_2" class="control-group">
<input type="text" data-field="x_veh_export_2" name="x_veh_export_2" id="x_veh_export_2" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_export_2.PlaceHolder %>" value="<%= vehicle_record_th.veh_export_2.EditValue %>"<%= vehicle_record_th.veh_export_2.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_export_2.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_export_3.Visible Then ' veh_export_3 %>
	<tr id="r_veh_export_3">
		<td><span id="elh_vehicle_record_th_veh_export_3"><%= vehicle_record_th.veh_export_3.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_export_3.CellAttributes %>>
<span id="el_vehicle_record_th_veh_export_3" class="control-group">
<input type="text" data-field="x_veh_export_3" name="x_veh_export_3" id="x_veh_export_3" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_export_3.PlaceHolder %>" value="<%= vehicle_record_th.veh_export_3.EditValue %>"<%= vehicle_record_th.veh_export_3.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_export_3.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_export_4.Visible Then ' veh_export_4 %>
	<tr id="r_veh_export_4">
		<td><span id="elh_vehicle_record_th_veh_export_4"><%= vehicle_record_th.veh_export_4.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_export_4.CellAttributes %>>
<span id="el_vehicle_record_th_veh_export_4" class="control-group">
<input type="text" data-field="x_veh_export_4" name="x_veh_export_4" id="x_veh_export_4" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_export_4.PlaceHolder %>" value="<%= vehicle_record_th.veh_export_4.EditValue %>"<%= vehicle_record_th.veh_export_4.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_export_4.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_export_5.Visible Then ' veh_export_5 %>
	<tr id="r_veh_export_5">
		<td><span id="elh_vehicle_record_th_veh_export_5"><%= vehicle_record_th.veh_export_5.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_export_5.CellAttributes %>>
<span id="el_vehicle_record_th_veh_export_5" class="control-group">
<input type="text" data-field="x_veh_export_5" name="x_veh_export_5" id="x_veh_export_5" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_export_5.PlaceHolder %>" value="<%= vehicle_record_th.veh_export_5.EditValue %>"<%= vehicle_record_th.veh_export_5.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_export_5.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_export_6.Visible Then ' veh_export_6 %>
	<tr id="r_veh_export_6">
		<td><span id="elh_vehicle_record_th_veh_export_6"><%= vehicle_record_th.veh_export_6.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_export_6.CellAttributes %>>
<span id="el_vehicle_record_th_veh_export_6" class="control-group">
<input type="text" data-field="x_veh_export_6" name="x_veh_export_6" id="x_veh_export_6" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_export_6.PlaceHolder %>" value="<%= vehicle_record_th.veh_export_6.EditValue %>"<%= vehicle_record_th.veh_export_6.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_export_6.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_export_7.Visible Then ' veh_export_7 %>
	<tr id="r_veh_export_7">
		<td><span id="elh_vehicle_record_th_veh_export_7"><%= vehicle_record_th.veh_export_7.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_export_7.CellAttributes %>>
<span id="el_vehicle_record_th_veh_export_7" class="control-group">
<input type="text" data-field="x_veh_export_7" name="x_veh_export_7" id="x_veh_export_7" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_export_7.PlaceHolder %>" value="<%= vehicle_record_th.veh_export_7.EditValue %>"<%= vehicle_record_th.veh_export_7.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_export_7.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_export_8.Visible Then ' veh_export_8 %>
	<tr id="r_veh_export_8">
		<td><span id="elh_vehicle_record_th_veh_export_8"><%= vehicle_record_th.veh_export_8.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_export_8.CellAttributes %>>
<span id="el_vehicle_record_th_veh_export_8" class="control-group">
<input type="text" data-field="x_veh_export_8" name="x_veh_export_8" id="x_veh_export_8" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_export_8.PlaceHolder %>" value="<%= vehicle_record_th.veh_export_8.EditValue %>"<%= vehicle_record_th.veh_export_8.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_export_8.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_remark.Visible Then ' veh_remark %>
	<tr id="r_veh_remark">
		<td><span id="elh_vehicle_record_th_veh_remark"><%= vehicle_record_th.veh_remark.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_remark.CellAttributes %>>
<span id="el_vehicle_record_th_veh_remark" class="control-group">
<textarea data-field="x_veh_remark" name="x_veh_remark" id="x_veh_remark" cols="35" rows="4" placeholder="<%= vehicle_record_th.veh_remark.PlaceHolder %>"<%= vehicle_record_th.veh_remark.EditAttributes %>><%= vehicle_record_th.veh_remark.EditValue %></textarea>
</span>
<%= vehicle_record_th.veh_remark.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_month_title.Visible Then ' veh_month_title %>
	<tr id="r_veh_month_title">
		<td><span id="elh_vehicle_record_th_veh_month_title"><%= vehicle_record_th.veh_month_title.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_month_title.CellAttributes %>>
<span id="el_vehicle_record_th_veh_month_title" class="control-group">
<input type="text" data-field="x_veh_month_title" name="x_veh_month_title" id="x_veh_month_title" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_month_title.PlaceHolder %>" value="<%= vehicle_record_th.veh_month_title.EditValue %>"<%= vehicle_record_th.veh_month_title.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_month_title.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_range.Visible Then ' veh_range %>
	<tr id="r_veh_range">
		<td><span id="elh_vehicle_record_th_veh_range"><%= vehicle_record_th.veh_range.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_range.CellAttributes %>>
<span id="el_vehicle_record_th_veh_range" class="control-group">
<input type="text" data-field="x_veh_range" name="x_veh_range" id="x_veh_range" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_range.PlaceHolder %>" value="<%= vehicle_record_th.veh_range.EditValue %>"<%= vehicle_record_th.veh_range.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_range.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_month_title2.Visible Then ' veh_month_title2 %>
	<tr id="r_veh_month_title2">
		<td><span id="elh_vehicle_record_th_veh_month_title2"><%= vehicle_record_th.veh_month_title2.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_month_title2.CellAttributes %>>
<span id="el_vehicle_record_th_veh_month_title2" class="control-group">
<input type="text" data-field="x_veh_month_title2" name="x_veh_month_title2" id="x_veh_month_title2" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_month_title2.PlaceHolder %>" value="<%= vehicle_record_th.veh_month_title2.EditValue %>"<%= vehicle_record_th.veh_month_title2.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_month_title2.CustomMsg %></td>
	</tr>
<% End If %>
<% If vehicle_record_th.veh_range2.Visible Then ' veh_range2 %>
	<tr id="r_veh_range2">
		<td><span id="elh_vehicle_record_th_veh_range2"><%= vehicle_record_th.veh_range2.FldCaption %></span></td>
		<td<%= vehicle_record_th.veh_range2.CellAttributes %>>
<span id="el_vehicle_record_th_veh_range2" class="control-group">
<input type="text" data-field="x_veh_range2" name="x_veh_range2" id="x_veh_range2" size="30" maxlength="255" placeholder="<%= vehicle_record_th.veh_range2.PlaceHolder %>" value="<%= vehicle_record_th.veh_range2.EditValue %>"<%= vehicle_record_th.veh_range2.EditAttributes %>>
</span>
<%= vehicle_record_th.veh_range2.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("EditBtn") %></button>
</form>
<script type="text/javascript">
fvehicle_record_thedit.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
vehicle_record_th_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set vehicle_record_th_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cvehicle_record_th_edit

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
		TableName = "vehicle_record_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "vehicle_record_th_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If vehicle_record_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & vehicle_record_th.TableVar & "&" ' add page token
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
		If vehicle_record_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (vehicle_record_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (vehicle_record_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(vehicle_record_th) Then Set vehicle_record_th = New cvehicle_record_th
		Set Table = vehicle_record_th

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "vehicle_record_th"

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

		vehicle_record_th.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action
		vehicle_record_th.veh_id.Visible = Not vehicle_record_th.IsAdd() And Not vehicle_record_th.IsCopy() And Not vehicle_record_th.IsGridAdd()

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
		Set vehicle_record_th = Nothing
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
		If Request.QueryString("veh_id").Count > 0 Then
			vehicle_record_th.veh_id.QueryStringValue = Request.QueryString("veh_id")
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Process form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			vehicle_record_th.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values
		Else
			vehicle_record_th.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If vehicle_record_th.veh_id.CurrentValue = "" Then Call Page_Terminate("pom_vehicle_record_thlist.asp") ' Invalid key, return to list

		' Validate form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			If Not ValidateForm() Then
				vehicle_record_th.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				vehicle_record_th.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		End If
		Select Case vehicle_record_th.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_vehicle_record_thlist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				vehicle_record_th.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					sReturnUrl = vehicle_record_th.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					vehicle_record_th.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		vehicle_record_th.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call vehicle_record_th.ResetAttrs()
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
				vehicle_record_th.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					vehicle_record_th.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = vehicle_record_th.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			vehicle_record_th.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			vehicle_record_th.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			vehicle_record_th.StartRecordNumber = StartRec
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
		If Not vehicle_record_th.veh_id.FldIsDetailKey Then vehicle_record_th.veh_id.FormValue = ObjForm.GetValue("x_veh_id")
		If Not vehicle_record_th.vch_month.FldIsDetailKey Then vehicle_record_th.vch_month.FormValue = ObjForm.GetValue("x_vch_month")
		If Not vehicle_record_th.vch_year.FldIsDetailKey Then vehicle_record_th.vch_year.FormValue = ObjForm.GetValue("x_vch_year")
		If Not vehicle_record_th.veh_product_1.FldIsDetailKey Then vehicle_record_th.veh_product_1.FormValue = ObjForm.GetValue("x_veh_product_1")
		If Not vehicle_record_th.veh_product_2.FldIsDetailKey Then vehicle_record_th.veh_product_2.FormValue = ObjForm.GetValue("x_veh_product_2")
		If Not vehicle_record_th.veh_product_3.FldIsDetailKey Then vehicle_record_th.veh_product_3.FormValue = ObjForm.GetValue("x_veh_product_3")
		If Not vehicle_record_th.veh_product_4.FldIsDetailKey Then vehicle_record_th.veh_product_4.FormValue = ObjForm.GetValue("x_veh_product_4")
		If Not vehicle_record_th.veh_product_5.FldIsDetailKey Then vehicle_record_th.veh_product_5.FormValue = ObjForm.GetValue("x_veh_product_5")
		If Not vehicle_record_th.veh_product_6.FldIsDetailKey Then vehicle_record_th.veh_product_6.FormValue = ObjForm.GetValue("x_veh_product_6")
		If Not vehicle_record_th.veh_product_7.FldIsDetailKey Then vehicle_record_th.veh_product_7.FormValue = ObjForm.GetValue("x_veh_product_7")
		If Not vehicle_record_th.veh_product_8.FldIsDetailKey Then vehicle_record_th.veh_product_8.FormValue = ObjForm.GetValue("x_veh_product_8")
		If Not vehicle_record_th.veh_domes_1.FldIsDetailKey Then vehicle_record_th.veh_domes_1.FormValue = ObjForm.GetValue("x_veh_domes_1")
		If Not vehicle_record_th.veh_domes_2.FldIsDetailKey Then vehicle_record_th.veh_domes_2.FormValue = ObjForm.GetValue("x_veh_domes_2")
		If Not vehicle_record_th.veh_domes_3.FldIsDetailKey Then vehicle_record_th.veh_domes_3.FormValue = ObjForm.GetValue("x_veh_domes_3")
		If Not vehicle_record_th.veh_domes_4.FldIsDetailKey Then vehicle_record_th.veh_domes_4.FormValue = ObjForm.GetValue("x_veh_domes_4")
		If Not vehicle_record_th.veh_domes_5.FldIsDetailKey Then vehicle_record_th.veh_domes_5.FormValue = ObjForm.GetValue("x_veh_domes_5")
		If Not vehicle_record_th.veh_domes_6.FldIsDetailKey Then vehicle_record_th.veh_domes_6.FormValue = ObjForm.GetValue("x_veh_domes_6")
		If Not vehicle_record_th.veh_domes_7.FldIsDetailKey Then vehicle_record_th.veh_domes_7.FormValue = ObjForm.GetValue("x_veh_domes_7")
		If Not vehicle_record_th.veh_domes_8.FldIsDetailKey Then vehicle_record_th.veh_domes_8.FormValue = ObjForm.GetValue("x_veh_domes_8")
		If Not vehicle_record_th.veh_export_1.FldIsDetailKey Then vehicle_record_th.veh_export_1.FormValue = ObjForm.GetValue("x_veh_export_1")
		If Not vehicle_record_th.veh_export_2.FldIsDetailKey Then vehicle_record_th.veh_export_2.FormValue = ObjForm.GetValue("x_veh_export_2")
		If Not vehicle_record_th.veh_export_3.FldIsDetailKey Then vehicle_record_th.veh_export_3.FormValue = ObjForm.GetValue("x_veh_export_3")
		If Not vehicle_record_th.veh_export_4.FldIsDetailKey Then vehicle_record_th.veh_export_4.FormValue = ObjForm.GetValue("x_veh_export_4")
		If Not vehicle_record_th.veh_export_5.FldIsDetailKey Then vehicle_record_th.veh_export_5.FormValue = ObjForm.GetValue("x_veh_export_5")
		If Not vehicle_record_th.veh_export_6.FldIsDetailKey Then vehicle_record_th.veh_export_6.FormValue = ObjForm.GetValue("x_veh_export_6")
		If Not vehicle_record_th.veh_export_7.FldIsDetailKey Then vehicle_record_th.veh_export_7.FormValue = ObjForm.GetValue("x_veh_export_7")
		If Not vehicle_record_th.veh_export_8.FldIsDetailKey Then vehicle_record_th.veh_export_8.FormValue = ObjForm.GetValue("x_veh_export_8")
		If Not vehicle_record_th.veh_remark.FldIsDetailKey Then vehicle_record_th.veh_remark.FormValue = ObjForm.GetValue("x_veh_remark")
		If Not vehicle_record_th.veh_month_title.FldIsDetailKey Then vehicle_record_th.veh_month_title.FormValue = ObjForm.GetValue("x_veh_month_title")
		If Not vehicle_record_th.veh_range.FldIsDetailKey Then vehicle_record_th.veh_range.FormValue = ObjForm.GetValue("x_veh_range")
		If Not vehicle_record_th.veh_month_title2.FldIsDetailKey Then vehicle_record_th.veh_month_title2.FormValue = ObjForm.GetValue("x_veh_month_title2")
		If Not vehicle_record_th.veh_range2.FldIsDetailKey Then vehicle_record_th.veh_range2.FormValue = ObjForm.GetValue("x_veh_range2")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		vehicle_record_th.veh_id.CurrentValue = vehicle_record_th.veh_id.FormValue
		vehicle_record_th.vch_month.CurrentValue = vehicle_record_th.vch_month.FormValue
		vehicle_record_th.vch_year.CurrentValue = vehicle_record_th.vch_year.FormValue
		vehicle_record_th.veh_product_1.CurrentValue = vehicle_record_th.veh_product_1.FormValue
		vehicle_record_th.veh_product_2.CurrentValue = vehicle_record_th.veh_product_2.FormValue
		vehicle_record_th.veh_product_3.CurrentValue = vehicle_record_th.veh_product_3.FormValue
		vehicle_record_th.veh_product_4.CurrentValue = vehicle_record_th.veh_product_4.FormValue
		vehicle_record_th.veh_product_5.CurrentValue = vehicle_record_th.veh_product_5.FormValue
		vehicle_record_th.veh_product_6.CurrentValue = vehicle_record_th.veh_product_6.FormValue
		vehicle_record_th.veh_product_7.CurrentValue = vehicle_record_th.veh_product_7.FormValue
		vehicle_record_th.veh_product_8.CurrentValue = vehicle_record_th.veh_product_8.FormValue
		vehicle_record_th.veh_domes_1.CurrentValue = vehicle_record_th.veh_domes_1.FormValue
		vehicle_record_th.veh_domes_2.CurrentValue = vehicle_record_th.veh_domes_2.FormValue
		vehicle_record_th.veh_domes_3.CurrentValue = vehicle_record_th.veh_domes_3.FormValue
		vehicle_record_th.veh_domes_4.CurrentValue = vehicle_record_th.veh_domes_4.FormValue
		vehicle_record_th.veh_domes_5.CurrentValue = vehicle_record_th.veh_domes_5.FormValue
		vehicle_record_th.veh_domes_6.CurrentValue = vehicle_record_th.veh_domes_6.FormValue
		vehicle_record_th.veh_domes_7.CurrentValue = vehicle_record_th.veh_domes_7.FormValue
		vehicle_record_th.veh_domes_8.CurrentValue = vehicle_record_th.veh_domes_8.FormValue
		vehicle_record_th.veh_export_1.CurrentValue = vehicle_record_th.veh_export_1.FormValue
		vehicle_record_th.veh_export_2.CurrentValue = vehicle_record_th.veh_export_2.FormValue
		vehicle_record_th.veh_export_3.CurrentValue = vehicle_record_th.veh_export_3.FormValue
		vehicle_record_th.veh_export_4.CurrentValue = vehicle_record_th.veh_export_4.FormValue
		vehicle_record_th.veh_export_5.CurrentValue = vehicle_record_th.veh_export_5.FormValue
		vehicle_record_th.veh_export_6.CurrentValue = vehicle_record_th.veh_export_6.FormValue
		vehicle_record_th.veh_export_7.CurrentValue = vehicle_record_th.veh_export_7.FormValue
		vehicle_record_th.veh_export_8.CurrentValue = vehicle_record_th.veh_export_8.FormValue
		vehicle_record_th.veh_remark.CurrentValue = vehicle_record_th.veh_remark.FormValue
		vehicle_record_th.veh_month_title.CurrentValue = vehicle_record_th.veh_month_title.FormValue
		vehicle_record_th.veh_range.CurrentValue = vehicle_record_th.veh_range.FormValue
		vehicle_record_th.veh_month_title2.CurrentValue = vehicle_record_th.veh_month_title2.FormValue
		vehicle_record_th.veh_range2.CurrentValue = vehicle_record_th.veh_range2.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = vehicle_record_th.KeyFilter

		' Call Row Selecting event
		Call vehicle_record_th.Row_Selecting(sFilter)

		' Load sql based on filter
		vehicle_record_th.CurrentFilter = sFilter
		sSql = vehicle_record_th.SQL
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
		Call vehicle_record_th.Row_Selected(RsRow)
		vehicle_record_th.veh_id.DbValue = RsRow("veh_id")
		vehicle_record_th.vch_month.DbValue = RsRow("vch_month")
		vehicle_record_th.vch_year.DbValue = RsRow("vch_year")
		vehicle_record_th.veh_product_1.DbValue = RsRow("veh_product_1")
		vehicle_record_th.veh_product_2.DbValue = RsRow("veh_product_2")
		vehicle_record_th.veh_product_3.DbValue = RsRow("veh_product_3")
		vehicle_record_th.veh_product_4.DbValue = RsRow("veh_product_4")
		vehicle_record_th.veh_product_5.DbValue = RsRow("veh_product_5")
		vehicle_record_th.veh_product_6.DbValue = RsRow("veh_product_6")
		vehicle_record_th.veh_product_7.DbValue = RsRow("veh_product_7")
		vehicle_record_th.veh_product_8.DbValue = RsRow("veh_product_8")
		vehicle_record_th.veh_domes_1.DbValue = RsRow("veh_domes_1")
		vehicle_record_th.veh_domes_2.DbValue = RsRow("veh_domes_2")
		vehicle_record_th.veh_domes_3.DbValue = RsRow("veh_domes_3")
		vehicle_record_th.veh_domes_4.DbValue = RsRow("veh_domes_4")
		vehicle_record_th.veh_domes_5.DbValue = RsRow("veh_domes_5")
		vehicle_record_th.veh_domes_6.DbValue = RsRow("veh_domes_6")
		vehicle_record_th.veh_domes_7.DbValue = RsRow("veh_domes_7")
		vehicle_record_th.veh_domes_8.DbValue = RsRow("veh_domes_8")
		vehicle_record_th.veh_export_1.DbValue = RsRow("veh_export_1")
		vehicle_record_th.veh_export_2.DbValue = RsRow("veh_export_2")
		vehicle_record_th.veh_export_3.DbValue = RsRow("veh_export_3")
		vehicle_record_th.veh_export_4.DbValue = RsRow("veh_export_4")
		vehicle_record_th.veh_export_5.DbValue = RsRow("veh_export_5")
		vehicle_record_th.veh_export_6.DbValue = RsRow("veh_export_6")
		vehicle_record_th.veh_export_7.DbValue = RsRow("veh_export_7")
		vehicle_record_th.veh_export_8.DbValue = RsRow("veh_export_8")
		vehicle_record_th.veh_remark.DbValue = RsRow("veh_remark")
		vehicle_record_th.veh_month_title.DbValue = RsRow("veh_month_title")
		vehicle_record_th.veh_range.DbValue = RsRow("veh_range")
		vehicle_record_th.veh_month_title2.DbValue = RsRow("veh_month_title2")
		vehicle_record_th.veh_range2.DbValue = RsRow("veh_range2")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		vehicle_record_th.veh_id.m_DbValue = Rs("veh_id")
		vehicle_record_th.vch_month.m_DbValue = Rs("vch_month")
		vehicle_record_th.vch_year.m_DbValue = Rs("vch_year")
		vehicle_record_th.veh_product_1.m_DbValue = Rs("veh_product_1")
		vehicle_record_th.veh_product_2.m_DbValue = Rs("veh_product_2")
		vehicle_record_th.veh_product_3.m_DbValue = Rs("veh_product_3")
		vehicle_record_th.veh_product_4.m_DbValue = Rs("veh_product_4")
		vehicle_record_th.veh_product_5.m_DbValue = Rs("veh_product_5")
		vehicle_record_th.veh_product_6.m_DbValue = Rs("veh_product_6")
		vehicle_record_th.veh_product_7.m_DbValue = Rs("veh_product_7")
		vehicle_record_th.veh_product_8.m_DbValue = Rs("veh_product_8")
		vehicle_record_th.veh_domes_1.m_DbValue = Rs("veh_domes_1")
		vehicle_record_th.veh_domes_2.m_DbValue = Rs("veh_domes_2")
		vehicle_record_th.veh_domes_3.m_DbValue = Rs("veh_domes_3")
		vehicle_record_th.veh_domes_4.m_DbValue = Rs("veh_domes_4")
		vehicle_record_th.veh_domes_5.m_DbValue = Rs("veh_domes_5")
		vehicle_record_th.veh_domes_6.m_DbValue = Rs("veh_domes_6")
		vehicle_record_th.veh_domes_7.m_DbValue = Rs("veh_domes_7")
		vehicle_record_th.veh_domes_8.m_DbValue = Rs("veh_domes_8")
		vehicle_record_th.veh_export_1.m_DbValue = Rs("veh_export_1")
		vehicle_record_th.veh_export_2.m_DbValue = Rs("veh_export_2")
		vehicle_record_th.veh_export_3.m_DbValue = Rs("veh_export_3")
		vehicle_record_th.veh_export_4.m_DbValue = Rs("veh_export_4")
		vehicle_record_th.veh_export_5.m_DbValue = Rs("veh_export_5")
		vehicle_record_th.veh_export_6.m_DbValue = Rs("veh_export_6")
		vehicle_record_th.veh_export_7.m_DbValue = Rs("veh_export_7")
		vehicle_record_th.veh_export_8.m_DbValue = Rs("veh_export_8")
		vehicle_record_th.veh_remark.m_DbValue = Rs("veh_remark")
		vehicle_record_th.veh_month_title.m_DbValue = Rs("veh_month_title")
		vehicle_record_th.veh_range.m_DbValue = Rs("veh_range")
		vehicle_record_th.veh_month_title2.m_DbValue = Rs("veh_month_title2")
		vehicle_record_th.veh_range2.m_DbValue = Rs("veh_range2")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call vehicle_record_th.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' veh_id
		' vch_month
		' vch_year
		' veh_product_1
		' veh_product_2
		' veh_product_3
		' veh_product_4
		' veh_product_5
		' veh_product_6
		' veh_product_7
		' veh_product_8
		' veh_domes_1
		' veh_domes_2
		' veh_domes_3
		' veh_domes_4
		' veh_domes_5
		' veh_domes_6
		' veh_domes_7
		' veh_domes_8
		' veh_export_1
		' veh_export_2
		' veh_export_3
		' veh_export_4
		' veh_export_5
		' veh_export_6
		' veh_export_7
		' veh_export_8
		' veh_remark
		' veh_month_title
		' veh_range
		' veh_month_title2
		' veh_range2
		' -----------
		'  View  Row
		' -----------

		If vehicle_record_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' veh_id
			vehicle_record_th.veh_id.ViewValue = vehicle_record_th.veh_id.CurrentValue
			vehicle_record_th.veh_id.ViewCustomAttributes = ""

			' vch_month
			vehicle_record_th.vch_month.ViewValue = vehicle_record_th.vch_month.CurrentValue
			vehicle_record_th.vch_month.ViewCustomAttributes = ""

			' vch_year
			vehicle_record_th.vch_year.ViewValue = vehicle_record_th.vch_year.CurrentValue
			vehicle_record_th.vch_year.ViewCustomAttributes = ""

			' veh_product_1
			vehicle_record_th.veh_product_1.ViewValue = vehicle_record_th.veh_product_1.CurrentValue
			vehicle_record_th.veh_product_1.ViewCustomAttributes = ""

			' veh_product_2
			vehicle_record_th.veh_product_2.ViewValue = vehicle_record_th.veh_product_2.CurrentValue
			vehicle_record_th.veh_product_2.ViewCustomAttributes = ""

			' veh_product_3
			vehicle_record_th.veh_product_3.ViewValue = vehicle_record_th.veh_product_3.CurrentValue
			vehicle_record_th.veh_product_3.ViewCustomAttributes = ""

			' veh_product_4
			vehicle_record_th.veh_product_4.ViewValue = vehicle_record_th.veh_product_4.CurrentValue
			vehicle_record_th.veh_product_4.ViewCustomAttributes = ""

			' veh_product_5
			vehicle_record_th.veh_product_5.ViewValue = vehicle_record_th.veh_product_5.CurrentValue
			vehicle_record_th.veh_product_5.ViewCustomAttributes = ""

			' veh_product_6
			vehicle_record_th.veh_product_6.ViewValue = vehicle_record_th.veh_product_6.CurrentValue
			vehicle_record_th.veh_product_6.ViewCustomAttributes = ""

			' veh_product_7
			vehicle_record_th.veh_product_7.ViewValue = vehicle_record_th.veh_product_7.CurrentValue
			vehicle_record_th.veh_product_7.ViewCustomAttributes = ""

			' veh_product_8
			vehicle_record_th.veh_product_8.ViewValue = vehicle_record_th.veh_product_8.CurrentValue
			vehicle_record_th.veh_product_8.ViewCustomAttributes = ""

			' veh_domes_1
			vehicle_record_th.veh_domes_1.ViewValue = vehicle_record_th.veh_domes_1.CurrentValue
			vehicle_record_th.veh_domes_1.ViewCustomAttributes = ""

			' veh_domes_2
			vehicle_record_th.veh_domes_2.ViewValue = vehicle_record_th.veh_domes_2.CurrentValue
			vehicle_record_th.veh_domes_2.ViewCustomAttributes = ""

			' veh_domes_3
			vehicle_record_th.veh_domes_3.ViewValue = vehicle_record_th.veh_domes_3.CurrentValue
			vehicle_record_th.veh_domes_3.ViewCustomAttributes = ""

			' veh_domes_4
			vehicle_record_th.veh_domes_4.ViewValue = vehicle_record_th.veh_domes_4.CurrentValue
			vehicle_record_th.veh_domes_4.ViewCustomAttributes = ""

			' veh_domes_5
			vehicle_record_th.veh_domes_5.ViewValue = vehicle_record_th.veh_domes_5.CurrentValue
			vehicle_record_th.veh_domes_5.ViewCustomAttributes = ""

			' veh_domes_6
			vehicle_record_th.veh_domes_6.ViewValue = vehicle_record_th.veh_domes_6.CurrentValue
			vehicle_record_th.veh_domes_6.ViewCustomAttributes = ""

			' veh_domes_7
			vehicle_record_th.veh_domes_7.ViewValue = vehicle_record_th.veh_domes_7.CurrentValue
			vehicle_record_th.veh_domes_7.ViewCustomAttributes = ""

			' veh_domes_8
			vehicle_record_th.veh_domes_8.ViewValue = vehicle_record_th.veh_domes_8.CurrentValue
			vehicle_record_th.veh_domes_8.ViewCustomAttributes = ""

			' veh_export_1
			vehicle_record_th.veh_export_1.ViewValue = vehicle_record_th.veh_export_1.CurrentValue
			vehicle_record_th.veh_export_1.ViewCustomAttributes = ""

			' veh_export_2
			vehicle_record_th.veh_export_2.ViewValue = vehicle_record_th.veh_export_2.CurrentValue
			vehicle_record_th.veh_export_2.ViewCustomAttributes = ""

			' veh_export_3
			vehicle_record_th.veh_export_3.ViewValue = vehicle_record_th.veh_export_3.CurrentValue
			vehicle_record_th.veh_export_3.ViewCustomAttributes = ""

			' veh_export_4
			vehicle_record_th.veh_export_4.ViewValue = vehicle_record_th.veh_export_4.CurrentValue
			vehicle_record_th.veh_export_4.ViewCustomAttributes = ""

			' veh_export_5
			vehicle_record_th.veh_export_5.ViewValue = vehicle_record_th.veh_export_5.CurrentValue
			vehicle_record_th.veh_export_5.ViewCustomAttributes = ""

			' veh_export_6
			vehicle_record_th.veh_export_6.ViewValue = vehicle_record_th.veh_export_6.CurrentValue
			vehicle_record_th.veh_export_6.ViewCustomAttributes = ""

			' veh_export_7
			vehicle_record_th.veh_export_7.ViewValue = vehicle_record_th.veh_export_7.CurrentValue
			vehicle_record_th.veh_export_7.ViewCustomAttributes = ""

			' veh_export_8
			vehicle_record_th.veh_export_8.ViewValue = vehicle_record_th.veh_export_8.CurrentValue
			vehicle_record_th.veh_export_8.ViewCustomAttributes = ""

			' veh_remark
			vehicle_record_th.veh_remark.ViewValue = vehicle_record_th.veh_remark.CurrentValue
			vehicle_record_th.veh_remark.ViewCustomAttributes = ""

			' veh_month_title
			vehicle_record_th.veh_month_title.ViewValue = vehicle_record_th.veh_month_title.CurrentValue
			vehicle_record_th.veh_month_title.ViewCustomAttributes = ""

			' veh_range
			vehicle_record_th.veh_range.ViewValue = vehicle_record_th.veh_range.CurrentValue
			vehicle_record_th.veh_range.ViewCustomAttributes = ""

			' veh_month_title2
			vehicle_record_th.veh_month_title2.ViewValue = vehicle_record_th.veh_month_title2.CurrentValue
			vehicle_record_th.veh_month_title2.ViewCustomAttributes = ""

			' veh_range2
			vehicle_record_th.veh_range2.ViewValue = vehicle_record_th.veh_range2.CurrentValue
			vehicle_record_th.veh_range2.ViewCustomAttributes = ""

			' View refer script
			' veh_id

			vehicle_record_th.veh_id.LinkCustomAttributes = ""
			vehicle_record_th.veh_id.HrefValue = ""
			vehicle_record_th.veh_id.TooltipValue = ""

			' vch_month
			vehicle_record_th.vch_month.LinkCustomAttributes = ""
			vehicle_record_th.vch_month.HrefValue = ""
			vehicle_record_th.vch_month.TooltipValue = ""

			' vch_year
			vehicle_record_th.vch_year.LinkCustomAttributes = ""
			vehicle_record_th.vch_year.HrefValue = ""
			vehicle_record_th.vch_year.TooltipValue = ""

			' veh_product_1
			vehicle_record_th.veh_product_1.LinkCustomAttributes = ""
			vehicle_record_th.veh_product_1.HrefValue = ""
			vehicle_record_th.veh_product_1.TooltipValue = ""

			' veh_product_2
			vehicle_record_th.veh_product_2.LinkCustomAttributes = ""
			vehicle_record_th.veh_product_2.HrefValue = ""
			vehicle_record_th.veh_product_2.TooltipValue = ""

			' veh_product_3
			vehicle_record_th.veh_product_3.LinkCustomAttributes = ""
			vehicle_record_th.veh_product_3.HrefValue = ""
			vehicle_record_th.veh_product_3.TooltipValue = ""

			' veh_product_4
			vehicle_record_th.veh_product_4.LinkCustomAttributes = ""
			vehicle_record_th.veh_product_4.HrefValue = ""
			vehicle_record_th.veh_product_4.TooltipValue = ""

			' veh_product_5
			vehicle_record_th.veh_product_5.LinkCustomAttributes = ""
			vehicle_record_th.veh_product_5.HrefValue = ""
			vehicle_record_th.veh_product_5.TooltipValue = ""

			' veh_product_6
			vehicle_record_th.veh_product_6.LinkCustomAttributes = ""
			vehicle_record_th.veh_product_6.HrefValue = ""
			vehicle_record_th.veh_product_6.TooltipValue = ""

			' veh_product_7
			vehicle_record_th.veh_product_7.LinkCustomAttributes = ""
			vehicle_record_th.veh_product_7.HrefValue = ""
			vehicle_record_th.veh_product_7.TooltipValue = ""

			' veh_product_8
			vehicle_record_th.veh_product_8.LinkCustomAttributes = ""
			vehicle_record_th.veh_product_8.HrefValue = ""
			vehicle_record_th.veh_product_8.TooltipValue = ""

			' veh_domes_1
			vehicle_record_th.veh_domes_1.LinkCustomAttributes = ""
			vehicle_record_th.veh_domes_1.HrefValue = ""
			vehicle_record_th.veh_domes_1.TooltipValue = ""

			' veh_domes_2
			vehicle_record_th.veh_domes_2.LinkCustomAttributes = ""
			vehicle_record_th.veh_domes_2.HrefValue = ""
			vehicle_record_th.veh_domes_2.TooltipValue = ""

			' veh_domes_3
			vehicle_record_th.veh_domes_3.LinkCustomAttributes = ""
			vehicle_record_th.veh_domes_3.HrefValue = ""
			vehicle_record_th.veh_domes_3.TooltipValue = ""

			' veh_domes_4
			vehicle_record_th.veh_domes_4.LinkCustomAttributes = ""
			vehicle_record_th.veh_domes_4.HrefValue = ""
			vehicle_record_th.veh_domes_4.TooltipValue = ""

			' veh_domes_5
			vehicle_record_th.veh_domes_5.LinkCustomAttributes = ""
			vehicle_record_th.veh_domes_5.HrefValue = ""
			vehicle_record_th.veh_domes_5.TooltipValue = ""

			' veh_domes_6
			vehicle_record_th.veh_domes_6.LinkCustomAttributes = ""
			vehicle_record_th.veh_domes_6.HrefValue = ""
			vehicle_record_th.veh_domes_6.TooltipValue = ""

			' veh_domes_7
			vehicle_record_th.veh_domes_7.LinkCustomAttributes = ""
			vehicle_record_th.veh_domes_7.HrefValue = ""
			vehicle_record_th.veh_domes_7.TooltipValue = ""

			' veh_domes_8
			vehicle_record_th.veh_domes_8.LinkCustomAttributes = ""
			vehicle_record_th.veh_domes_8.HrefValue = ""
			vehicle_record_th.veh_domes_8.TooltipValue = ""

			' veh_export_1
			vehicle_record_th.veh_export_1.LinkCustomAttributes = ""
			vehicle_record_th.veh_export_1.HrefValue = ""
			vehicle_record_th.veh_export_1.TooltipValue = ""

			' veh_export_2
			vehicle_record_th.veh_export_2.LinkCustomAttributes = ""
			vehicle_record_th.veh_export_2.HrefValue = ""
			vehicle_record_th.veh_export_2.TooltipValue = ""

			' veh_export_3
			vehicle_record_th.veh_export_3.LinkCustomAttributes = ""
			vehicle_record_th.veh_export_3.HrefValue = ""
			vehicle_record_th.veh_export_3.TooltipValue = ""

			' veh_export_4
			vehicle_record_th.veh_export_4.LinkCustomAttributes = ""
			vehicle_record_th.veh_export_4.HrefValue = ""
			vehicle_record_th.veh_export_4.TooltipValue = ""

			' veh_export_5
			vehicle_record_th.veh_export_5.LinkCustomAttributes = ""
			vehicle_record_th.veh_export_5.HrefValue = ""
			vehicle_record_th.veh_export_5.TooltipValue = ""

			' veh_export_6
			vehicle_record_th.veh_export_6.LinkCustomAttributes = ""
			vehicle_record_th.veh_export_6.HrefValue = ""
			vehicle_record_th.veh_export_6.TooltipValue = ""

			' veh_export_7
			vehicle_record_th.veh_export_7.LinkCustomAttributes = ""
			vehicle_record_th.veh_export_7.HrefValue = ""
			vehicle_record_th.veh_export_7.TooltipValue = ""

			' veh_export_8
			vehicle_record_th.veh_export_8.LinkCustomAttributes = ""
			vehicle_record_th.veh_export_8.HrefValue = ""
			vehicle_record_th.veh_export_8.TooltipValue = ""

			' veh_remark
			vehicle_record_th.veh_remark.LinkCustomAttributes = ""
			vehicle_record_th.veh_remark.HrefValue = ""
			vehicle_record_th.veh_remark.TooltipValue = ""

			' veh_month_title
			vehicle_record_th.veh_month_title.LinkCustomAttributes = ""
			vehicle_record_th.veh_month_title.HrefValue = ""
			vehicle_record_th.veh_month_title.TooltipValue = ""

			' veh_range
			vehicle_record_th.veh_range.LinkCustomAttributes = ""
			vehicle_record_th.veh_range.HrefValue = ""
			vehicle_record_th.veh_range.TooltipValue = ""

			' veh_month_title2
			vehicle_record_th.veh_month_title2.LinkCustomAttributes = ""
			vehicle_record_th.veh_month_title2.HrefValue = ""
			vehicle_record_th.veh_month_title2.TooltipValue = ""

			' veh_range2
			vehicle_record_th.veh_range2.LinkCustomAttributes = ""
			vehicle_record_th.veh_range2.HrefValue = ""
			vehicle_record_th.veh_range2.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf vehicle_record_th.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' veh_id
			vehicle_record_th.veh_id.EditCustomAttributes = ""
			vehicle_record_th.veh_id.EditValue = vehicle_record_th.veh_id.CurrentValue
			vehicle_record_th.veh_id.ViewCustomAttributes = ""

			' vch_month
			vehicle_record_th.vch_month.EditCustomAttributes = ""
			vehicle_record_th.vch_month.EditValue = ew_HtmlEncode(vehicle_record_th.vch_month.CurrentValue)
			vehicle_record_th.vch_month.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.vch_month.FldCaption))

			' vch_year
			vehicle_record_th.vch_year.EditCustomAttributes = ""
			vehicle_record_th.vch_year.EditValue = ew_HtmlEncode(vehicle_record_th.vch_year.CurrentValue)
			vehicle_record_th.vch_year.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.vch_year.FldCaption))

			' veh_product_1
			vehicle_record_th.veh_product_1.EditCustomAttributes = ""
			vehicle_record_th.veh_product_1.EditValue = ew_HtmlEncode(vehicle_record_th.veh_product_1.CurrentValue)
			vehicle_record_th.veh_product_1.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_product_1.FldCaption))

			' veh_product_2
			vehicle_record_th.veh_product_2.EditCustomAttributes = ""
			vehicle_record_th.veh_product_2.EditValue = ew_HtmlEncode(vehicle_record_th.veh_product_2.CurrentValue)
			vehicle_record_th.veh_product_2.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_product_2.FldCaption))

			' veh_product_3
			vehicle_record_th.veh_product_3.EditCustomAttributes = ""
			vehicle_record_th.veh_product_3.EditValue = ew_HtmlEncode(vehicle_record_th.veh_product_3.CurrentValue)
			vehicle_record_th.veh_product_3.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_product_3.FldCaption))

			' veh_product_4
			vehicle_record_th.veh_product_4.EditCustomAttributes = ""
			vehicle_record_th.veh_product_4.EditValue = ew_HtmlEncode(vehicle_record_th.veh_product_4.CurrentValue)
			vehicle_record_th.veh_product_4.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_product_4.FldCaption))

			' veh_product_5
			vehicle_record_th.veh_product_5.EditCustomAttributes = ""
			vehicle_record_th.veh_product_5.EditValue = ew_HtmlEncode(vehicle_record_th.veh_product_5.CurrentValue)
			vehicle_record_th.veh_product_5.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_product_5.FldCaption))

			' veh_product_6
			vehicle_record_th.veh_product_6.EditCustomAttributes = ""
			vehicle_record_th.veh_product_6.EditValue = ew_HtmlEncode(vehicle_record_th.veh_product_6.CurrentValue)
			vehicle_record_th.veh_product_6.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_product_6.FldCaption))

			' veh_product_7
			vehicle_record_th.veh_product_7.EditCustomAttributes = ""
			vehicle_record_th.veh_product_7.EditValue = ew_HtmlEncode(vehicle_record_th.veh_product_7.CurrentValue)
			vehicle_record_th.veh_product_7.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_product_7.FldCaption))

			' veh_product_8
			vehicle_record_th.veh_product_8.EditCustomAttributes = ""
			vehicle_record_th.veh_product_8.EditValue = ew_HtmlEncode(vehicle_record_th.veh_product_8.CurrentValue)
			vehicle_record_th.veh_product_8.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_product_8.FldCaption))

			' veh_domes_1
			vehicle_record_th.veh_domes_1.EditCustomAttributes = ""
			vehicle_record_th.veh_domes_1.EditValue = ew_HtmlEncode(vehicle_record_th.veh_domes_1.CurrentValue)
			vehicle_record_th.veh_domes_1.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_domes_1.FldCaption))

			' veh_domes_2
			vehicle_record_th.veh_domes_2.EditCustomAttributes = ""
			vehicle_record_th.veh_domes_2.EditValue = ew_HtmlEncode(vehicle_record_th.veh_domes_2.CurrentValue)
			vehicle_record_th.veh_domes_2.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_domes_2.FldCaption))

			' veh_domes_3
			vehicle_record_th.veh_domes_3.EditCustomAttributes = ""
			vehicle_record_th.veh_domes_3.EditValue = ew_HtmlEncode(vehicle_record_th.veh_domes_3.CurrentValue)
			vehicle_record_th.veh_domes_3.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_domes_3.FldCaption))

			' veh_domes_4
			vehicle_record_th.veh_domes_4.EditCustomAttributes = ""
			vehicle_record_th.veh_domes_4.EditValue = ew_HtmlEncode(vehicle_record_th.veh_domes_4.CurrentValue)
			vehicle_record_th.veh_domes_4.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_domes_4.FldCaption))

			' veh_domes_5
			vehicle_record_th.veh_domes_5.EditCustomAttributes = ""
			vehicle_record_th.veh_domes_5.EditValue = ew_HtmlEncode(vehicle_record_th.veh_domes_5.CurrentValue)
			vehicle_record_th.veh_domes_5.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_domes_5.FldCaption))

			' veh_domes_6
			vehicle_record_th.veh_domes_6.EditCustomAttributes = ""
			vehicle_record_th.veh_domes_6.EditValue = ew_HtmlEncode(vehicle_record_th.veh_domes_6.CurrentValue)
			vehicle_record_th.veh_domes_6.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_domes_6.FldCaption))

			' veh_domes_7
			vehicle_record_th.veh_domes_7.EditCustomAttributes = ""
			vehicle_record_th.veh_domes_7.EditValue = ew_HtmlEncode(vehicle_record_th.veh_domes_7.CurrentValue)
			vehicle_record_th.veh_domes_7.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_domes_7.FldCaption))

			' veh_domes_8
			vehicle_record_th.veh_domes_8.EditCustomAttributes = ""
			vehicle_record_th.veh_domes_8.EditValue = ew_HtmlEncode(vehicle_record_th.veh_domes_8.CurrentValue)
			vehicle_record_th.veh_domes_8.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_domes_8.FldCaption))

			' veh_export_1
			vehicle_record_th.veh_export_1.EditCustomAttributes = ""
			vehicle_record_th.veh_export_1.EditValue = ew_HtmlEncode(vehicle_record_th.veh_export_1.CurrentValue)
			vehicle_record_th.veh_export_1.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_export_1.FldCaption))

			' veh_export_2
			vehicle_record_th.veh_export_2.EditCustomAttributes = ""
			vehicle_record_th.veh_export_2.EditValue = ew_HtmlEncode(vehicle_record_th.veh_export_2.CurrentValue)
			vehicle_record_th.veh_export_2.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_export_2.FldCaption))

			' veh_export_3
			vehicle_record_th.veh_export_3.EditCustomAttributes = ""
			vehicle_record_th.veh_export_3.EditValue = ew_HtmlEncode(vehicle_record_th.veh_export_3.CurrentValue)
			vehicle_record_th.veh_export_3.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_export_3.FldCaption))

			' veh_export_4
			vehicle_record_th.veh_export_4.EditCustomAttributes = ""
			vehicle_record_th.veh_export_4.EditValue = ew_HtmlEncode(vehicle_record_th.veh_export_4.CurrentValue)
			vehicle_record_th.veh_export_4.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_export_4.FldCaption))

			' veh_export_5
			vehicle_record_th.veh_export_5.EditCustomAttributes = ""
			vehicle_record_th.veh_export_5.EditValue = ew_HtmlEncode(vehicle_record_th.veh_export_5.CurrentValue)
			vehicle_record_th.veh_export_5.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_export_5.FldCaption))

			' veh_export_6
			vehicle_record_th.veh_export_6.EditCustomAttributes = ""
			vehicle_record_th.veh_export_6.EditValue = ew_HtmlEncode(vehicle_record_th.veh_export_6.CurrentValue)
			vehicle_record_th.veh_export_6.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_export_6.FldCaption))

			' veh_export_7
			vehicle_record_th.veh_export_7.EditCustomAttributes = ""
			vehicle_record_th.veh_export_7.EditValue = ew_HtmlEncode(vehicle_record_th.veh_export_7.CurrentValue)
			vehicle_record_th.veh_export_7.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_export_7.FldCaption))

			' veh_export_8
			vehicle_record_th.veh_export_8.EditCustomAttributes = ""
			vehicle_record_th.veh_export_8.EditValue = ew_HtmlEncode(vehicle_record_th.veh_export_8.CurrentValue)
			vehicle_record_th.veh_export_8.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_export_8.FldCaption))

			' veh_remark
			vehicle_record_th.veh_remark.EditCustomAttributes = ""
			vehicle_record_th.veh_remark.EditValue = vehicle_record_th.veh_remark.CurrentValue
			vehicle_record_th.veh_remark.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_remark.FldCaption))

			' veh_month_title
			vehicle_record_th.veh_month_title.EditCustomAttributes = ""
			vehicle_record_th.veh_month_title.EditValue = ew_HtmlEncode(vehicle_record_th.veh_month_title.CurrentValue)
			vehicle_record_th.veh_month_title.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_month_title.FldCaption))

			' veh_range
			vehicle_record_th.veh_range.EditCustomAttributes = ""
			vehicle_record_th.veh_range.EditValue = ew_HtmlEncode(vehicle_record_th.veh_range.CurrentValue)
			vehicle_record_th.veh_range.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_range.FldCaption))

			' veh_month_title2
			vehicle_record_th.veh_month_title2.EditCustomAttributes = ""
			vehicle_record_th.veh_month_title2.EditValue = ew_HtmlEncode(vehicle_record_th.veh_month_title2.CurrentValue)
			vehicle_record_th.veh_month_title2.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_month_title2.FldCaption))

			' veh_range2
			vehicle_record_th.veh_range2.EditCustomAttributes = ""
			vehicle_record_th.veh_range2.EditValue = ew_HtmlEncode(vehicle_record_th.veh_range2.CurrentValue)
			vehicle_record_th.veh_range2.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(vehicle_record_th.veh_range2.FldCaption))

			' Edit refer script
			' veh_id

			vehicle_record_th.veh_id.HrefValue = ""

			' vch_month
			vehicle_record_th.vch_month.HrefValue = ""

			' vch_year
			vehicle_record_th.vch_year.HrefValue = ""

			' veh_product_1
			vehicle_record_th.veh_product_1.HrefValue = ""

			' veh_product_2
			vehicle_record_th.veh_product_2.HrefValue = ""

			' veh_product_3
			vehicle_record_th.veh_product_3.HrefValue = ""

			' veh_product_4
			vehicle_record_th.veh_product_4.HrefValue = ""

			' veh_product_5
			vehicle_record_th.veh_product_5.HrefValue = ""

			' veh_product_6
			vehicle_record_th.veh_product_6.HrefValue = ""

			' veh_product_7
			vehicle_record_th.veh_product_7.HrefValue = ""

			' veh_product_8
			vehicle_record_th.veh_product_8.HrefValue = ""

			' veh_domes_1
			vehicle_record_th.veh_domes_1.HrefValue = ""

			' veh_domes_2
			vehicle_record_th.veh_domes_2.HrefValue = ""

			' veh_domes_3
			vehicle_record_th.veh_domes_3.HrefValue = ""

			' veh_domes_4
			vehicle_record_th.veh_domes_4.HrefValue = ""

			' veh_domes_5
			vehicle_record_th.veh_domes_5.HrefValue = ""

			' veh_domes_6
			vehicle_record_th.veh_domes_6.HrefValue = ""

			' veh_domes_7
			vehicle_record_th.veh_domes_7.HrefValue = ""

			' veh_domes_8
			vehicle_record_th.veh_domes_8.HrefValue = ""

			' veh_export_1
			vehicle_record_th.veh_export_1.HrefValue = ""

			' veh_export_2
			vehicle_record_th.veh_export_2.HrefValue = ""

			' veh_export_3
			vehicle_record_th.veh_export_3.HrefValue = ""

			' veh_export_4
			vehicle_record_th.veh_export_4.HrefValue = ""

			' veh_export_5
			vehicle_record_th.veh_export_5.HrefValue = ""

			' veh_export_6
			vehicle_record_th.veh_export_6.HrefValue = ""

			' veh_export_7
			vehicle_record_th.veh_export_7.HrefValue = ""

			' veh_export_8
			vehicle_record_th.veh_export_8.HrefValue = ""

			' veh_remark
			vehicle_record_th.veh_remark.HrefValue = ""

			' veh_month_title
			vehicle_record_th.veh_month_title.HrefValue = ""

			' veh_range
			vehicle_record_th.veh_range.HrefValue = ""

			' veh_month_title2
			vehicle_record_th.veh_month_title2.HrefValue = ""

			' veh_range2
			vehicle_record_th.veh_range2.HrefValue = ""
		End If
		If vehicle_record_th.RowType = EW_ROWTYPE_ADD Or vehicle_record_th.RowType = EW_ROWTYPE_EDIT Or vehicle_record_th.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call vehicle_record_th.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If vehicle_record_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call vehicle_record_th.Row_Rendered()
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
		sFilter = vehicle_record_th.KeyFilter
		vehicle_record_th.CurrentFilter  = sFilter
		sSql = vehicle_record_th.SQL
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

			' Field vch_month
			Call vehicle_record_th.vch_month.SetDbValue(Rs, vehicle_record_th.vch_month.CurrentValue, Null, vehicle_record_th.vch_month.ReadOnly)

			' Field vch_year
			Call vehicle_record_th.vch_year.SetDbValue(Rs, vehicle_record_th.vch_year.CurrentValue, Null, vehicle_record_th.vch_year.ReadOnly)

			' Field veh_product_1
			Call vehicle_record_th.veh_product_1.SetDbValue(Rs, vehicle_record_th.veh_product_1.CurrentValue, Null, vehicle_record_th.veh_product_1.ReadOnly)

			' Field veh_product_2
			Call vehicle_record_th.veh_product_2.SetDbValue(Rs, vehicle_record_th.veh_product_2.CurrentValue, Null, vehicle_record_th.veh_product_2.ReadOnly)

			' Field veh_product_3
			Call vehicle_record_th.veh_product_3.SetDbValue(Rs, vehicle_record_th.veh_product_3.CurrentValue, Null, vehicle_record_th.veh_product_3.ReadOnly)

			' Field veh_product_4
			Call vehicle_record_th.veh_product_4.SetDbValue(Rs, vehicle_record_th.veh_product_4.CurrentValue, Null, vehicle_record_th.veh_product_4.ReadOnly)

			' Field veh_product_5
			Call vehicle_record_th.veh_product_5.SetDbValue(Rs, vehicle_record_th.veh_product_5.CurrentValue, Null, vehicle_record_th.veh_product_5.ReadOnly)

			' Field veh_product_6
			Call vehicle_record_th.veh_product_6.SetDbValue(Rs, vehicle_record_th.veh_product_6.CurrentValue, Null, vehicle_record_th.veh_product_6.ReadOnly)

			' Field veh_product_7
			Call vehicle_record_th.veh_product_7.SetDbValue(Rs, vehicle_record_th.veh_product_7.CurrentValue, Null, vehicle_record_th.veh_product_7.ReadOnly)

			' Field veh_product_8
			Call vehicle_record_th.veh_product_8.SetDbValue(Rs, vehicle_record_th.veh_product_8.CurrentValue, Null, vehicle_record_th.veh_product_8.ReadOnly)

			' Field veh_domes_1
			Call vehicle_record_th.veh_domes_1.SetDbValue(Rs, vehicle_record_th.veh_domes_1.CurrentValue, Null, vehicle_record_th.veh_domes_1.ReadOnly)

			' Field veh_domes_2
			Call vehicle_record_th.veh_domes_2.SetDbValue(Rs, vehicle_record_th.veh_domes_2.CurrentValue, Null, vehicle_record_th.veh_domes_2.ReadOnly)

			' Field veh_domes_3
			Call vehicle_record_th.veh_domes_3.SetDbValue(Rs, vehicle_record_th.veh_domes_3.CurrentValue, Null, vehicle_record_th.veh_domes_3.ReadOnly)

			' Field veh_domes_4
			Call vehicle_record_th.veh_domes_4.SetDbValue(Rs, vehicle_record_th.veh_domes_4.CurrentValue, Null, vehicle_record_th.veh_domes_4.ReadOnly)

			' Field veh_domes_5
			Call vehicle_record_th.veh_domes_5.SetDbValue(Rs, vehicle_record_th.veh_domes_5.CurrentValue, Null, vehicle_record_th.veh_domes_5.ReadOnly)

			' Field veh_domes_6
			Call vehicle_record_th.veh_domes_6.SetDbValue(Rs, vehicle_record_th.veh_domes_6.CurrentValue, Null, vehicle_record_th.veh_domes_6.ReadOnly)

			' Field veh_domes_7
			Call vehicle_record_th.veh_domes_7.SetDbValue(Rs, vehicle_record_th.veh_domes_7.CurrentValue, Null, vehicle_record_th.veh_domes_7.ReadOnly)

			' Field veh_domes_8
			Call vehicle_record_th.veh_domes_8.SetDbValue(Rs, vehicle_record_th.veh_domes_8.CurrentValue, Null, vehicle_record_th.veh_domes_8.ReadOnly)

			' Field veh_export_1
			Call vehicle_record_th.veh_export_1.SetDbValue(Rs, vehicle_record_th.veh_export_1.CurrentValue, Null, vehicle_record_th.veh_export_1.ReadOnly)

			' Field veh_export_2
			Call vehicle_record_th.veh_export_2.SetDbValue(Rs, vehicle_record_th.veh_export_2.CurrentValue, Null, vehicle_record_th.veh_export_2.ReadOnly)

			' Field veh_export_3
			Call vehicle_record_th.veh_export_3.SetDbValue(Rs, vehicle_record_th.veh_export_3.CurrentValue, Null, vehicle_record_th.veh_export_3.ReadOnly)

			' Field veh_export_4
			Call vehicle_record_th.veh_export_4.SetDbValue(Rs, vehicle_record_th.veh_export_4.CurrentValue, Null, vehicle_record_th.veh_export_4.ReadOnly)

			' Field veh_export_5
			Call vehicle_record_th.veh_export_5.SetDbValue(Rs, vehicle_record_th.veh_export_5.CurrentValue, Null, vehicle_record_th.veh_export_5.ReadOnly)

			' Field veh_export_6
			Call vehicle_record_th.veh_export_6.SetDbValue(Rs, vehicle_record_th.veh_export_6.CurrentValue, Null, vehicle_record_th.veh_export_6.ReadOnly)

			' Field veh_export_7
			Call vehicle_record_th.veh_export_7.SetDbValue(Rs, vehicle_record_th.veh_export_7.CurrentValue, Null, vehicle_record_th.veh_export_7.ReadOnly)

			' Field veh_export_8
			Call vehicle_record_th.veh_export_8.SetDbValue(Rs, vehicle_record_th.veh_export_8.CurrentValue, Null, vehicle_record_th.veh_export_8.ReadOnly)

			' Field veh_remark
			Call vehicle_record_th.veh_remark.SetDbValue(Rs, vehicle_record_th.veh_remark.CurrentValue, Null, vehicle_record_th.veh_remark.ReadOnly)

			' Field veh_month_title
			Call vehicle_record_th.veh_month_title.SetDbValue(Rs, vehicle_record_th.veh_month_title.CurrentValue, Null, vehicle_record_th.veh_month_title.ReadOnly)

			' Field veh_range
			Call vehicle_record_th.veh_range.SetDbValue(Rs, vehicle_record_th.veh_range.CurrentValue, Null, vehicle_record_th.veh_range.ReadOnly)

			' Field veh_month_title2
			Call vehicle_record_th.veh_month_title2.SetDbValue(Rs, vehicle_record_th.veh_month_title2.CurrentValue, Null, vehicle_record_th.veh_month_title2.ReadOnly)

			' Field veh_range2
			Call vehicle_record_th.veh_range2.SetDbValue(Rs, vehicle_record_th.veh_range2.CurrentValue, Null, vehicle_record_th.veh_range2.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = vehicle_record_th.Row_Updating(RsOld, Rs)
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
				ElseIf vehicle_record_th.CancelMessage <> "" Then
					FailureMessage = vehicle_record_th.CancelMessage
					vehicle_record_th.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call vehicle_record_th.Row_Updated(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", vehicle_record_th.TableVar, "pom_vehicle_record_thlist.asp", vehicle_record_th.TableVar, True)
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
