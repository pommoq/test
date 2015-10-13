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
Dim vehicle_record_th_delete
Set vehicle_record_th_delete = New cvehicle_record_th_delete
Set Page = vehicle_record_th_delete

' Page init processing
vehicle_record_th_delete.Page_Init()

' Page main processing
vehicle_record_th_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
vehicle_record_th_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var vehicle_record_th_delete = new ew_Page("vehicle_record_th_delete");
vehicle_record_th_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = vehicle_record_th_delete.PageID; // For backward compatibility
// Form object
var fvehicle_record_thdelete = new ew_Form("fvehicle_record_thdelete");
// Form_CustomValidate event
fvehicle_record_thdelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fvehicle_record_thdelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fvehicle_record_thdelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set vehicle_record_th_delete.Recordset = vehicle_record_th_delete.LoadRecordset()
vehicle_record_th_delete.TotalRecs = vehicle_record_th_delete.Recordset.RecordCount ' Get record count
If vehicle_record_th_delete.TotalRecs <= 0 Then ' No record found, exit
	vehicle_record_th_delete.Recordset.Close
	Set vehicle_record_th_delete.Recordset = Nothing
	Call vehicle_record_th_delete.Page_Terminate("pom_vehicle_record_thlist.asp") ' Return to list
End If
%>
<% If vehicle_record_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% vehicle_record_th_delete.ShowPageHeader() %>
<% vehicle_record_th_delete.ShowMessage %>
<form name="fvehicle_record_thdelete" id="fvehicle_record_thdelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="vehicle_record_th">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(vehicle_record_th_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(vehicle_record_th_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_vehicle_record_thdelete" class="ewTable ewTableSeparate">
<%= vehicle_record_th.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If vehicle_record_th.veh_id.Visible Then ' veh_id %>
		<td><span id="elh_vehicle_record_th_veh_id" class="vehicle_record_th_veh_id"><%= vehicle_record_th.veh_id.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.vch_month.Visible Then ' vch_month %>
		<td><span id="elh_vehicle_record_th_vch_month" class="vehicle_record_th_vch_month"><%= vehicle_record_th.vch_month.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.vch_year.Visible Then ' vch_year %>
		<td><span id="elh_vehicle_record_th_vch_year" class="vehicle_record_th_vch_year"><%= vehicle_record_th.vch_year.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_product_1.Visible Then ' veh_product_1 %>
		<td><span id="elh_vehicle_record_th_veh_product_1" class="vehicle_record_th_veh_product_1"><%= vehicle_record_th.veh_product_1.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_product_2.Visible Then ' veh_product_2 %>
		<td><span id="elh_vehicle_record_th_veh_product_2" class="vehicle_record_th_veh_product_2"><%= vehicle_record_th.veh_product_2.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_product_3.Visible Then ' veh_product_3 %>
		<td><span id="elh_vehicle_record_th_veh_product_3" class="vehicle_record_th_veh_product_3"><%= vehicle_record_th.veh_product_3.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_product_4.Visible Then ' veh_product_4 %>
		<td><span id="elh_vehicle_record_th_veh_product_4" class="vehicle_record_th_veh_product_4"><%= vehicle_record_th.veh_product_4.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_product_5.Visible Then ' veh_product_5 %>
		<td><span id="elh_vehicle_record_th_veh_product_5" class="vehicle_record_th_veh_product_5"><%= vehicle_record_th.veh_product_5.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_product_6.Visible Then ' veh_product_6 %>
		<td><span id="elh_vehicle_record_th_veh_product_6" class="vehicle_record_th_veh_product_6"><%= vehicle_record_th.veh_product_6.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_product_7.Visible Then ' veh_product_7 %>
		<td><span id="elh_vehicle_record_th_veh_product_7" class="vehicle_record_th_veh_product_7"><%= vehicle_record_th.veh_product_7.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_product_8.Visible Then ' veh_product_8 %>
		<td><span id="elh_vehicle_record_th_veh_product_8" class="vehicle_record_th_veh_product_8"><%= vehicle_record_th.veh_product_8.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_domes_1.Visible Then ' veh_domes_1 %>
		<td><span id="elh_vehicle_record_th_veh_domes_1" class="vehicle_record_th_veh_domes_1"><%= vehicle_record_th.veh_domes_1.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_domes_2.Visible Then ' veh_domes_2 %>
		<td><span id="elh_vehicle_record_th_veh_domes_2" class="vehicle_record_th_veh_domes_2"><%= vehicle_record_th.veh_domes_2.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_domes_3.Visible Then ' veh_domes_3 %>
		<td><span id="elh_vehicle_record_th_veh_domes_3" class="vehicle_record_th_veh_domes_3"><%= vehicle_record_th.veh_domes_3.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_domes_4.Visible Then ' veh_domes_4 %>
		<td><span id="elh_vehicle_record_th_veh_domes_4" class="vehicle_record_th_veh_domes_4"><%= vehicle_record_th.veh_domes_4.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_domes_5.Visible Then ' veh_domes_5 %>
		<td><span id="elh_vehicle_record_th_veh_domes_5" class="vehicle_record_th_veh_domes_5"><%= vehicle_record_th.veh_domes_5.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_domes_6.Visible Then ' veh_domes_6 %>
		<td><span id="elh_vehicle_record_th_veh_domes_6" class="vehicle_record_th_veh_domes_6"><%= vehicle_record_th.veh_domes_6.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_domes_7.Visible Then ' veh_domes_7 %>
		<td><span id="elh_vehicle_record_th_veh_domes_7" class="vehicle_record_th_veh_domes_7"><%= vehicle_record_th.veh_domes_7.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_domes_8.Visible Then ' veh_domes_8 %>
		<td><span id="elh_vehicle_record_th_veh_domes_8" class="vehicle_record_th_veh_domes_8"><%= vehicle_record_th.veh_domes_8.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_export_1.Visible Then ' veh_export_1 %>
		<td><span id="elh_vehicle_record_th_veh_export_1" class="vehicle_record_th_veh_export_1"><%= vehicle_record_th.veh_export_1.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_export_2.Visible Then ' veh_export_2 %>
		<td><span id="elh_vehicle_record_th_veh_export_2" class="vehicle_record_th_veh_export_2"><%= vehicle_record_th.veh_export_2.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_export_3.Visible Then ' veh_export_3 %>
		<td><span id="elh_vehicle_record_th_veh_export_3" class="vehicle_record_th_veh_export_3"><%= vehicle_record_th.veh_export_3.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_export_4.Visible Then ' veh_export_4 %>
		<td><span id="elh_vehicle_record_th_veh_export_4" class="vehicle_record_th_veh_export_4"><%= vehicle_record_th.veh_export_4.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_export_5.Visible Then ' veh_export_5 %>
		<td><span id="elh_vehicle_record_th_veh_export_5" class="vehicle_record_th_veh_export_5"><%= vehicle_record_th.veh_export_5.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_export_6.Visible Then ' veh_export_6 %>
		<td><span id="elh_vehicle_record_th_veh_export_6" class="vehicle_record_th_veh_export_6"><%= vehicle_record_th.veh_export_6.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_export_7.Visible Then ' veh_export_7 %>
		<td><span id="elh_vehicle_record_th_veh_export_7" class="vehicle_record_th_veh_export_7"><%= vehicle_record_th.veh_export_7.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_export_8.Visible Then ' veh_export_8 %>
		<td><span id="elh_vehicle_record_th_veh_export_8" class="vehicle_record_th_veh_export_8"><%= vehicle_record_th.veh_export_8.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_month_title.Visible Then ' veh_month_title %>
		<td><span id="elh_vehicle_record_th_veh_month_title" class="vehicle_record_th_veh_month_title"><%= vehicle_record_th.veh_month_title.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_range.Visible Then ' veh_range %>
		<td><span id="elh_vehicle_record_th_veh_range" class="vehicle_record_th_veh_range"><%= vehicle_record_th.veh_range.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_month_title2.Visible Then ' veh_month_title2 %>
		<td><span id="elh_vehicle_record_th_veh_month_title2" class="vehicle_record_th_veh_month_title2"><%= vehicle_record_th.veh_month_title2.FldCaption %></span></td>
<% End If %>
<% If vehicle_record_th.veh_range2.Visible Then ' veh_range2 %>
		<td><span id="elh_vehicle_record_th_veh_range2" class="vehicle_record_th_veh_range2"><%= vehicle_record_th.veh_range2.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
vehicle_record_th_delete.RecCnt = 0
vehicle_record_th_delete.RowCnt = 0
Do While (Not vehicle_record_th_delete.Recordset.Eof)
	vehicle_record_th_delete.RecCnt = vehicle_record_th_delete.RecCnt + 1
	vehicle_record_th_delete.RowCnt = vehicle_record_th_delete.RowCnt + 1

	' Set row properties
	Call vehicle_record_th.ResetAttrs()
	vehicle_record_th.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call vehicle_record_th_delete.LoadRowValues(vehicle_record_th_delete.Recordset)

	' Render row
	Call vehicle_record_th_delete.RenderRow()
%>
	<tr<%= vehicle_record_th.RowAttributes %>>
<% If vehicle_record_th.veh_id.Visible Then ' veh_id %>
		<td<%= vehicle_record_th.veh_id.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_id" class="control-group vehicle_record_th_veh_id">
<span<%= vehicle_record_th.veh_id.ViewAttributes %>>
<%= vehicle_record_th.veh_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.vch_month.Visible Then ' vch_month %>
		<td<%= vehicle_record_th.vch_month.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_vch_month" class="control-group vehicle_record_th_vch_month">
<span<%= vehicle_record_th.vch_month.ViewAttributes %>>
<%= vehicle_record_th.vch_month.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.vch_year.Visible Then ' vch_year %>
		<td<%= vehicle_record_th.vch_year.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_vch_year" class="control-group vehicle_record_th_vch_year">
<span<%= vehicle_record_th.vch_year.ViewAttributes %>>
<%= vehicle_record_th.vch_year.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_product_1.Visible Then ' veh_product_1 %>
		<td<%= vehicle_record_th.veh_product_1.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_product_1" class="control-group vehicle_record_th_veh_product_1">
<span<%= vehicle_record_th.veh_product_1.ViewAttributes %>>
<%= vehicle_record_th.veh_product_1.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_product_2.Visible Then ' veh_product_2 %>
		<td<%= vehicle_record_th.veh_product_2.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_product_2" class="control-group vehicle_record_th_veh_product_2">
<span<%= vehicle_record_th.veh_product_2.ViewAttributes %>>
<%= vehicle_record_th.veh_product_2.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_product_3.Visible Then ' veh_product_3 %>
		<td<%= vehicle_record_th.veh_product_3.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_product_3" class="control-group vehicle_record_th_veh_product_3">
<span<%= vehicle_record_th.veh_product_3.ViewAttributes %>>
<%= vehicle_record_th.veh_product_3.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_product_4.Visible Then ' veh_product_4 %>
		<td<%= vehicle_record_th.veh_product_4.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_product_4" class="control-group vehicle_record_th_veh_product_4">
<span<%= vehicle_record_th.veh_product_4.ViewAttributes %>>
<%= vehicle_record_th.veh_product_4.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_product_5.Visible Then ' veh_product_5 %>
		<td<%= vehicle_record_th.veh_product_5.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_product_5" class="control-group vehicle_record_th_veh_product_5">
<span<%= vehicle_record_th.veh_product_5.ViewAttributes %>>
<%= vehicle_record_th.veh_product_5.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_product_6.Visible Then ' veh_product_6 %>
		<td<%= vehicle_record_th.veh_product_6.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_product_6" class="control-group vehicle_record_th_veh_product_6">
<span<%= vehicle_record_th.veh_product_6.ViewAttributes %>>
<%= vehicle_record_th.veh_product_6.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_product_7.Visible Then ' veh_product_7 %>
		<td<%= vehicle_record_th.veh_product_7.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_product_7" class="control-group vehicle_record_th_veh_product_7">
<span<%= vehicle_record_th.veh_product_7.ViewAttributes %>>
<%= vehicle_record_th.veh_product_7.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_product_8.Visible Then ' veh_product_8 %>
		<td<%= vehicle_record_th.veh_product_8.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_product_8" class="control-group vehicle_record_th_veh_product_8">
<span<%= vehicle_record_th.veh_product_8.ViewAttributes %>>
<%= vehicle_record_th.veh_product_8.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_domes_1.Visible Then ' veh_domes_1 %>
		<td<%= vehicle_record_th.veh_domes_1.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_domes_1" class="control-group vehicle_record_th_veh_domes_1">
<span<%= vehicle_record_th.veh_domes_1.ViewAttributes %>>
<%= vehicle_record_th.veh_domes_1.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_domes_2.Visible Then ' veh_domes_2 %>
		<td<%= vehicle_record_th.veh_domes_2.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_domes_2" class="control-group vehicle_record_th_veh_domes_2">
<span<%= vehicle_record_th.veh_domes_2.ViewAttributes %>>
<%= vehicle_record_th.veh_domes_2.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_domes_3.Visible Then ' veh_domes_3 %>
		<td<%= vehicle_record_th.veh_domes_3.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_domes_3" class="control-group vehicle_record_th_veh_domes_3">
<span<%= vehicle_record_th.veh_domes_3.ViewAttributes %>>
<%= vehicle_record_th.veh_domes_3.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_domes_4.Visible Then ' veh_domes_4 %>
		<td<%= vehicle_record_th.veh_domes_4.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_domes_4" class="control-group vehicle_record_th_veh_domes_4">
<span<%= vehicle_record_th.veh_domes_4.ViewAttributes %>>
<%= vehicle_record_th.veh_domes_4.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_domes_5.Visible Then ' veh_domes_5 %>
		<td<%= vehicle_record_th.veh_domes_5.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_domes_5" class="control-group vehicle_record_th_veh_domes_5">
<span<%= vehicle_record_th.veh_domes_5.ViewAttributes %>>
<%= vehicle_record_th.veh_domes_5.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_domes_6.Visible Then ' veh_domes_6 %>
		<td<%= vehicle_record_th.veh_domes_6.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_domes_6" class="control-group vehicle_record_th_veh_domes_6">
<span<%= vehicle_record_th.veh_domes_6.ViewAttributes %>>
<%= vehicle_record_th.veh_domes_6.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_domes_7.Visible Then ' veh_domes_7 %>
		<td<%= vehicle_record_th.veh_domes_7.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_domes_7" class="control-group vehicle_record_th_veh_domes_7">
<span<%= vehicle_record_th.veh_domes_7.ViewAttributes %>>
<%= vehicle_record_th.veh_domes_7.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_domes_8.Visible Then ' veh_domes_8 %>
		<td<%= vehicle_record_th.veh_domes_8.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_domes_8" class="control-group vehicle_record_th_veh_domes_8">
<span<%= vehicle_record_th.veh_domes_8.ViewAttributes %>>
<%= vehicle_record_th.veh_domes_8.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_export_1.Visible Then ' veh_export_1 %>
		<td<%= vehicle_record_th.veh_export_1.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_export_1" class="control-group vehicle_record_th_veh_export_1">
<span<%= vehicle_record_th.veh_export_1.ViewAttributes %>>
<%= vehicle_record_th.veh_export_1.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_export_2.Visible Then ' veh_export_2 %>
		<td<%= vehicle_record_th.veh_export_2.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_export_2" class="control-group vehicle_record_th_veh_export_2">
<span<%= vehicle_record_th.veh_export_2.ViewAttributes %>>
<%= vehicle_record_th.veh_export_2.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_export_3.Visible Then ' veh_export_3 %>
		<td<%= vehicle_record_th.veh_export_3.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_export_3" class="control-group vehicle_record_th_veh_export_3">
<span<%= vehicle_record_th.veh_export_3.ViewAttributes %>>
<%= vehicle_record_th.veh_export_3.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_export_4.Visible Then ' veh_export_4 %>
		<td<%= vehicle_record_th.veh_export_4.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_export_4" class="control-group vehicle_record_th_veh_export_4">
<span<%= vehicle_record_th.veh_export_4.ViewAttributes %>>
<%= vehicle_record_th.veh_export_4.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_export_5.Visible Then ' veh_export_5 %>
		<td<%= vehicle_record_th.veh_export_5.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_export_5" class="control-group vehicle_record_th_veh_export_5">
<span<%= vehicle_record_th.veh_export_5.ViewAttributes %>>
<%= vehicle_record_th.veh_export_5.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_export_6.Visible Then ' veh_export_6 %>
		<td<%= vehicle_record_th.veh_export_6.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_export_6" class="control-group vehicle_record_th_veh_export_6">
<span<%= vehicle_record_th.veh_export_6.ViewAttributes %>>
<%= vehicle_record_th.veh_export_6.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_export_7.Visible Then ' veh_export_7 %>
		<td<%= vehicle_record_th.veh_export_7.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_export_7" class="control-group vehicle_record_th_veh_export_7">
<span<%= vehicle_record_th.veh_export_7.ViewAttributes %>>
<%= vehicle_record_th.veh_export_7.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_export_8.Visible Then ' veh_export_8 %>
		<td<%= vehicle_record_th.veh_export_8.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_export_8" class="control-group vehicle_record_th_veh_export_8">
<span<%= vehicle_record_th.veh_export_8.ViewAttributes %>>
<%= vehicle_record_th.veh_export_8.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_month_title.Visible Then ' veh_month_title %>
		<td<%= vehicle_record_th.veh_month_title.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_month_title" class="control-group vehicle_record_th_veh_month_title">
<span<%= vehicle_record_th.veh_month_title.ViewAttributes %>>
<%= vehicle_record_th.veh_month_title.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_range.Visible Then ' veh_range %>
		<td<%= vehicle_record_th.veh_range.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_range" class="control-group vehicle_record_th_veh_range">
<span<%= vehicle_record_th.veh_range.ViewAttributes %>>
<%= vehicle_record_th.veh_range.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_month_title2.Visible Then ' veh_month_title2 %>
		<td<%= vehicle_record_th.veh_month_title2.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_month_title2" class="control-group vehicle_record_th_veh_month_title2">
<span<%= vehicle_record_th.veh_month_title2.ViewAttributes %>>
<%= vehicle_record_th.veh_month_title2.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If vehicle_record_th.veh_range2.Visible Then ' veh_range2 %>
		<td<%= vehicle_record_th.veh_range2.CellAttributes %>>
<span id="el<%= vehicle_record_th_delete.RowCnt %>_vehicle_record_th_veh_range2" class="control-group vehicle_record_th_veh_range2">
<span<%= vehicle_record_th.veh_range2.ViewAttributes %>>
<%= vehicle_record_th.veh_range2.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	vehicle_record_th_delete.Recordset.MoveNext
Loop
vehicle_record_th_delete.Recordset.Close
Set vehicle_record_th_delete.Recordset = Nothing
%>
	</tbody>
</table>
</div>
</td></tr></table>
<div class="btn-group ewButtonGroup">
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("DeleteBtn") %></button>
</div>
</form>
<script type="text/javascript">
fvehicle_record_thdelete.Init();
</script>
<%
vehicle_record_th_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set vehicle_record_th_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cvehicle_record_th_delete

	' Page ID
	Public Property Get PageID()
		PageID = "delete"
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
		PageObjName = "vehicle_record_th_delete"
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
		EW_PAGE_ID = "delete"

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

	Dim TotalRecs
	Dim RecCnt
	Dim RecKeys
	Dim Recordset
	Dim StartRowCnt
	Dim RowCnt

	' Page main processing
	Sub Page_Main()
		Dim sFilter
		StartRowCnt = 1

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Load Key Parameters
		RecKeys = vehicle_record_th.GetRecordKeys() ' Load record keys
		sFilter = vehicle_record_th.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_vehicle_record_thlist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in vehicle_record_th class, vehicle_record_thinfo.asp

		vehicle_record_th.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			vehicle_record_th.CurrentAction = Request.Form("a_delete")
		Else
			vehicle_record_th.CurrentAction = "I"	' Display record
		End If
		Select Case vehicle_record_th.CurrentAction
			Case "D" ' Delete
				vehicle_record_th.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(vehicle_record_th.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = vehicle_record_th.CurrentFilter
		Call vehicle_record_th.Recordset_Selecting(sFilter)
		vehicle_record_th.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = vehicle_record_th.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call vehicle_record_th.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
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
		End If

		' Call Row Rendered event
		If vehicle_record_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call vehicle_record_th.Row_Rendered()
		End If
	End Sub

	'
	' Delete records based on current filter
	'
	Function DeleteRows()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim sKey, sThisKey, sKeyFld, arKeyFlds
		Dim sSql, RsDelete
		Dim RsOld, RsDetail
		DeleteRows = True
		sSql = vehicle_record_th.SQL
		Conn.BeginTrans
		Set RsDelete = Server.CreateObject("ADODB.Recordset")
		RsDelete.CursorLocation = EW_CURSORLOCATION
		RsDelete.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			RsDelete.Close
			Set RsDelete = Nothing
			DeleteRows = False
			Exit Function
		ElseIf RsDelete.Eof Then
			FailureMessage = Language.Phrase("NoRecord") ' No record found
			RsDelete.Close
			Set RsDelete = Nothing
			DeleteRows = False
			Exit Function
		End If

		' Clone old recordset object
		Set RsOld = ew_CloneRs(RsDelete)

		' Call row deleting event
		If DeleteRows Then
			RsDelete.MoveFirst
			Do While Not RsDelete.Eof
				DeleteRows = vehicle_record_th.Row_Deleting(RsDelete)
				If Not DeleteRows Then Exit Do
				RsDelete.MoveNext
			Loop
			RsDelete.MoveFirst
		End If
		If DeleteRows Then
			sKey = ""
			RsDelete.MoveFirst
			Do While Not RsDelete.Eof
				sThisKey = ""
				If sThisKey <> "" Then sThisKey = sThisKey & EW_COMPOSITE_KEY_SEPARATOR
				sThisKey = sThisKey & RsDelete("veh_id")
				Call LoadDbValues(RsDelete)
				If DeleteRows Then
					RsDelete.Delete
				End If
				If Err.Number <> 0 Or Not DeleteRows Then
					If Err.Description <> "" Then FailureMessage = Err.Description ' Set up error message
					DeleteRows = False
					Exit Do
				End If
				If sKey <> "" Then sKey = sKey & ", "
				sKey = sKey & sThisKey
				RsDelete.MoveNext
			Loop
		Else

			' Set up error message
			If SuccessMessage <> "" Or FailureMessage <> "" Then

				' Use the message, do nothing
			ElseIf vehicle_record_th.CancelMessage <> "" Then
				FailureMessage = vehicle_record_th.CancelMessage
				vehicle_record_th.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("DeleteCancelled")
			End If
		End If
		If DeleteRows Then
			Conn.CommitTrans ' Commit the changes
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				DeleteRows = False ' Delete failed
			End If
		Else
			Conn.RollbackTrans ' Rollback changes
		End If
		RsDelete.Close
		Set RsDelete = Nothing

		' Call row deleting event
		If DeleteRows Then
			RsOld.MoveFirst
			Do While Not RsOld.Eof
				Call vehicle_record_th.Row_Deleted(RsOld)
				RsOld.MoveNext
			Loop
		End If
		RsOld.Close
		Set RsOld = Nothing
	End Function

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", vehicle_record_th.TableVar, "pom_vehicle_record_thlist.asp", vehicle_record_th.TableVar, True)
		PageId = "delete"
		Call Breadcrumb.Add("delete", PageId, ew_CurrentUrl, "", False)
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
End Class
%>
