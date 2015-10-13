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
Dim vehicle_record_th_list
Set vehicle_record_th_list = New cvehicle_record_th_list
Set Page = vehicle_record_th_list

' Page init processing
vehicle_record_th_list.Page_Init()

' Page main processing
vehicle_record_th_list.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
vehicle_record_th_list.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If vehicle_record_th.Export = "" Then %>
<script type="text/javascript">
// Page object
var vehicle_record_th_list = new ew_Page("vehicle_record_th_list");
vehicle_record_th_list.PageID = "list"; // Page ID
var EW_PAGE_ID = vehicle_record_th_list.PageID; // For backward compatibility
// Form object
var fvehicle_record_thlist = new ew_Form("fvehicle_record_thlist");
fvehicle_record_thlist.FormKeyCountName = '<%= vehicle_record_th_list.FormKeyCountName %>';
// Form_CustomValidate event
fvehicle_record_thlist.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fvehicle_record_thlist.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fvehicle_record_thlist.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
var fvehicle_record_thlistsrch = new ew_Form("fvehicle_record_thlistsrch");
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If vehicle_record_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If vehicle_record_th_list.ExportOptions.Visible Then %>
<div class="ewListExportOptions"><% vehicle_record_th_list.ExportOptions.Render "body", "", "", "", "", "" %></div>
<% End If %>
<% If (vehicle_record_th.Export = "") Or (EW_EXPORT_MASTER_RECORD And vehicle_record_th.Export = "print") Then %>
<% End If %>
<%

' Load recordset
Set vehicle_record_th_list.Recordset = vehicle_record_th_list.LoadRecordset()
	vehicle_record_th_list.TotalRecs = vehicle_record_th_list.Recordset.RecordCount
	vehicle_record_th_list.StartRec = 1
	If vehicle_record_th_list.DisplayRecs <= 0 Then ' Display all records
		vehicle_record_th_list.DisplayRecs = vehicle_record_th_list.TotalRecs
	End If
	If Not (vehicle_record_th.ExportAll And vehicle_record_th.Export <> "") Then
		vehicle_record_th_list.SetUpStartRec() ' Set up start record position
	End If
vehicle_record_th_list.RenderOtherOptions()
%>
<% If Security.IsLoggedIn() Then %>
<% If vehicle_record_th.Export = "" And vehicle_record_th.CurrentAction = "" Then %>
<form name="fvehicle_record_thlistsrch" id="fvehicle_record_thlistsrch" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewSearchTable"><tr><td>
<div class="accordion" id="fvehicle_record_thlistsrch_SearchGroup">
	<div class="accordion-group">
		<div class="accordion-heading">
<a class="accordion-toggle" data-toggle="collapse" data-parent="#fvehicle_record_thlistsrch_SearchGroup" href="#fvehicle_record_thlistsrch_SearchBody"><%= Language.Phrase("Search") %></a>
		</div>
		<div id="fvehicle_record_thlistsrch_SearchBody" class="accordion-body collapse in">
			<div class="accordion-inner">
<div id="fvehicle_record_thlistsrch_SearchPanel">
<input type="hidden" name="cmd" value="search">
<input type="hidden" name="t" value="vehicle_record_th">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewRow">
	<div class="btn-group ewButtonGroup">
	<div class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" class="input-large" value="<%= ew_HtmlEncode(vehicle_record_th.BasicSearch.getKeyword()) %>" placeholder="<%= ew_HtmlEncode(Language.Phrase("Search")) %>">
	<button class="btn btn-primary ewButton" name="btnsubmit" id="btnsubmit" type="submit"><%= Language.Phrase("QuickSearchBtn") %></button>
	</div>
	</div>
	<div class="btn-group ewButtonGroup">
	<a class="btn ewShowAll" href="<%= vehicle_record_th_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>
	</div>
</div>
<div id="xsr_2" class="ewRow">
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="="<% If vehicle_record_th.BasicSearch.getSearchType = "=" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If vehicle_record_th.BasicSearch.getSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If vehicle_record_th.BasicSearch.getSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
</div>
</div>
</div>
			</div>
		</div>
	</div>
</div>
</td></tr></table>
</form>
<% End If %>
<% End If %>
<% vehicle_record_th_list.ShowPageHeader() %>
<% vehicle_record_th_list.ShowMessage %>
<table class="ewGrid"><tr><td class="ewGridContent">
<form name="fvehicle_record_thlist" id="fvehicle_record_thlist" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="vehicle_record_th">
<div id="gmp_vehicle_record_th" class="ewGridMiddlePanel">
<% If vehicle_record_th_list.TotalRecs > 0 Then %>
<table id="tbl_vehicle_record_thlist" class="ewTable ewTableSeparate">
<%= vehicle_record_th.TableCustomInnerHtml %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call vehicle_record_th_list.RenderListOptions()

' Render list options (header, left)
vehicle_record_th_list.ListOptions.Render "header", "left", "", "", "", ""
%>
<% If vehicle_record_th.veh_id.Visible Then ' veh_id %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_id) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_id" class="vehicle_record_th_veh_id"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_id.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_id) %>',1);"><div id="elh_vehicle_record_th_veh_id" class="vehicle_record_th_veh_id">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_id.FldCaption %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_id.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_id.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.vch_month.Visible Then ' vch_month %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.vch_month) = "" Then %>
		<td><div id="elh_vehicle_record_th_vch_month" class="vehicle_record_th_vch_month"><div class="ewTableHeaderCaption"><%= vehicle_record_th.vch_month.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.vch_month) %>',1);"><div id="elh_vehicle_record_th_vch_month" class="vehicle_record_th_vch_month">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.vch_month.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.vch_month.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.vch_month.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.vch_year.Visible Then ' vch_year %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.vch_year) = "" Then %>
		<td><div id="elh_vehicle_record_th_vch_year" class="vehicle_record_th_vch_year"><div class="ewTableHeaderCaption"><%= vehicle_record_th.vch_year.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.vch_year) %>',1);"><div id="elh_vehicle_record_th_vch_year" class="vehicle_record_th_vch_year">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.vch_year.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.vch_year.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.vch_year.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_product_1.Visible Then ' veh_product_1 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_product_1) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_product_1" class="vehicle_record_th_veh_product_1"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_product_1.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_product_1) %>',1);"><div id="elh_vehicle_record_th_veh_product_1" class="vehicle_record_th_veh_product_1">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_product_1.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_product_1.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_product_1.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_product_2.Visible Then ' veh_product_2 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_product_2) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_product_2" class="vehicle_record_th_veh_product_2"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_product_2.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_product_2) %>',1);"><div id="elh_vehicle_record_th_veh_product_2" class="vehicle_record_th_veh_product_2">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_product_2.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_product_2.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_product_2.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_product_3.Visible Then ' veh_product_3 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_product_3) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_product_3" class="vehicle_record_th_veh_product_3"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_product_3.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_product_3) %>',1);"><div id="elh_vehicle_record_th_veh_product_3" class="vehicle_record_th_veh_product_3">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_product_3.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_product_3.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_product_3.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_product_4.Visible Then ' veh_product_4 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_product_4) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_product_4" class="vehicle_record_th_veh_product_4"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_product_4.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_product_4) %>',1);"><div id="elh_vehicle_record_th_veh_product_4" class="vehicle_record_th_veh_product_4">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_product_4.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_product_4.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_product_4.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_product_5.Visible Then ' veh_product_5 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_product_5) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_product_5" class="vehicle_record_th_veh_product_5"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_product_5.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_product_5) %>',1);"><div id="elh_vehicle_record_th_veh_product_5" class="vehicle_record_th_veh_product_5">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_product_5.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_product_5.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_product_5.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_product_6.Visible Then ' veh_product_6 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_product_6) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_product_6" class="vehicle_record_th_veh_product_6"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_product_6.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_product_6) %>',1);"><div id="elh_vehicle_record_th_veh_product_6" class="vehicle_record_th_veh_product_6">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_product_6.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_product_6.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_product_6.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_product_7.Visible Then ' veh_product_7 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_product_7) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_product_7" class="vehicle_record_th_veh_product_7"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_product_7.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_product_7) %>',1);"><div id="elh_vehicle_record_th_veh_product_7" class="vehicle_record_th_veh_product_7">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_product_7.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_product_7.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_product_7.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_product_8.Visible Then ' veh_product_8 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_product_8) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_product_8" class="vehicle_record_th_veh_product_8"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_product_8.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_product_8) %>',1);"><div id="elh_vehicle_record_th_veh_product_8" class="vehicle_record_th_veh_product_8">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_product_8.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_product_8.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_product_8.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_domes_1.Visible Then ' veh_domes_1 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_domes_1) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_domes_1" class="vehicle_record_th_veh_domes_1"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_domes_1.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_domes_1) %>',1);"><div id="elh_vehicle_record_th_veh_domes_1" class="vehicle_record_th_veh_domes_1">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_domes_1.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_domes_1.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_domes_1.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_domes_2.Visible Then ' veh_domes_2 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_domes_2) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_domes_2" class="vehicle_record_th_veh_domes_2"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_domes_2.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_domes_2) %>',1);"><div id="elh_vehicle_record_th_veh_domes_2" class="vehicle_record_th_veh_domes_2">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_domes_2.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_domes_2.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_domes_2.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_domes_3.Visible Then ' veh_domes_3 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_domes_3) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_domes_3" class="vehicle_record_th_veh_domes_3"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_domes_3.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_domes_3) %>',1);"><div id="elh_vehicle_record_th_veh_domes_3" class="vehicle_record_th_veh_domes_3">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_domes_3.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_domes_3.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_domes_3.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_domes_4.Visible Then ' veh_domes_4 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_domes_4) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_domes_4" class="vehicle_record_th_veh_domes_4"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_domes_4.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_domes_4) %>',1);"><div id="elh_vehicle_record_th_veh_domes_4" class="vehicle_record_th_veh_domes_4">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_domes_4.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_domes_4.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_domes_4.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_domes_5.Visible Then ' veh_domes_5 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_domes_5) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_domes_5" class="vehicle_record_th_veh_domes_5"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_domes_5.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_domes_5) %>',1);"><div id="elh_vehicle_record_th_veh_domes_5" class="vehicle_record_th_veh_domes_5">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_domes_5.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_domes_5.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_domes_5.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_domes_6.Visible Then ' veh_domes_6 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_domes_6) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_domes_6" class="vehicle_record_th_veh_domes_6"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_domes_6.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_domes_6) %>',1);"><div id="elh_vehicle_record_th_veh_domes_6" class="vehicle_record_th_veh_domes_6">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_domes_6.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_domes_6.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_domes_6.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_domes_7.Visible Then ' veh_domes_7 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_domes_7) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_domes_7" class="vehicle_record_th_veh_domes_7"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_domes_7.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_domes_7) %>',1);"><div id="elh_vehicle_record_th_veh_domes_7" class="vehicle_record_th_veh_domes_7">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_domes_7.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_domes_7.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_domes_7.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_domes_8.Visible Then ' veh_domes_8 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_domes_8) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_domes_8" class="vehicle_record_th_veh_domes_8"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_domes_8.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_domes_8) %>',1);"><div id="elh_vehicle_record_th_veh_domes_8" class="vehicle_record_th_veh_domes_8">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_domes_8.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_domes_8.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_domes_8.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_export_1.Visible Then ' veh_export_1 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_export_1) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_export_1" class="vehicle_record_th_veh_export_1"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_export_1.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_export_1) %>',1);"><div id="elh_vehicle_record_th_veh_export_1" class="vehicle_record_th_veh_export_1">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_export_1.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_export_1.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_export_1.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_export_2.Visible Then ' veh_export_2 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_export_2) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_export_2" class="vehicle_record_th_veh_export_2"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_export_2.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_export_2) %>',1);"><div id="elh_vehicle_record_th_veh_export_2" class="vehicle_record_th_veh_export_2">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_export_2.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_export_2.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_export_2.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_export_3.Visible Then ' veh_export_3 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_export_3) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_export_3" class="vehicle_record_th_veh_export_3"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_export_3.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_export_3) %>',1);"><div id="elh_vehicle_record_th_veh_export_3" class="vehicle_record_th_veh_export_3">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_export_3.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_export_3.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_export_3.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_export_4.Visible Then ' veh_export_4 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_export_4) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_export_4" class="vehicle_record_th_veh_export_4"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_export_4.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_export_4) %>',1);"><div id="elh_vehicle_record_th_veh_export_4" class="vehicle_record_th_veh_export_4">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_export_4.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_export_4.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_export_4.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_export_5.Visible Then ' veh_export_5 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_export_5) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_export_5" class="vehicle_record_th_veh_export_5"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_export_5.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_export_5) %>',1);"><div id="elh_vehicle_record_th_veh_export_5" class="vehicle_record_th_veh_export_5">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_export_5.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_export_5.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_export_5.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_export_6.Visible Then ' veh_export_6 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_export_6) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_export_6" class="vehicle_record_th_veh_export_6"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_export_6.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_export_6) %>',1);"><div id="elh_vehicle_record_th_veh_export_6" class="vehicle_record_th_veh_export_6">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_export_6.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_export_6.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_export_6.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_export_7.Visible Then ' veh_export_7 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_export_7) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_export_7" class="vehicle_record_th_veh_export_7"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_export_7.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_export_7) %>',1);"><div id="elh_vehicle_record_th_veh_export_7" class="vehicle_record_th_veh_export_7">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_export_7.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_export_7.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_export_7.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_export_8.Visible Then ' veh_export_8 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_export_8) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_export_8" class="vehicle_record_th_veh_export_8"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_export_8.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_export_8) %>',1);"><div id="elh_vehicle_record_th_veh_export_8" class="vehicle_record_th_veh_export_8">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_export_8.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_export_8.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_export_8.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_month_title.Visible Then ' veh_month_title %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_month_title) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_month_title" class="vehicle_record_th_veh_month_title"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_month_title.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_month_title) %>',1);"><div id="elh_vehicle_record_th_veh_month_title" class="vehicle_record_th_veh_month_title">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_month_title.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_month_title.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_month_title.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_range.Visible Then ' veh_range %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_range) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_range" class="vehicle_record_th_veh_range"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_range.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_range) %>',1);"><div id="elh_vehicle_record_th_veh_range" class="vehicle_record_th_veh_range">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_range.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_range.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_range.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_month_title2.Visible Then ' veh_month_title2 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_month_title2) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_month_title2" class="vehicle_record_th_veh_month_title2"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_month_title2.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_month_title2) %>',1);"><div id="elh_vehicle_record_th_veh_month_title2" class="vehicle_record_th_veh_month_title2">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_month_title2.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_month_title2.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_month_title2.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If vehicle_record_th.veh_range2.Visible Then ' veh_range2 %>
	<% If vehicle_record_th.SortUrl(vehicle_record_th.veh_range2) = "" Then %>
		<td><div id="elh_vehicle_record_th_veh_range2" class="vehicle_record_th_veh_range2"><div class="ewTableHeaderCaption"><%= vehicle_record_th.veh_range2.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= vehicle_record_th.SortUrl(vehicle_record_th.veh_range2) %>',1);"><div id="elh_vehicle_record_th_veh_range2" class="vehicle_record_th_veh_range2">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= vehicle_record_th.veh_range2.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If vehicle_record_th.veh_range2.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf vehicle_record_th.veh_range2.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
vehicle_record_th_list.ListOptions.Render "header", "right", "", "", "", ""
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (vehicle_record_th.ExportAll And vehicle_record_th.Export <> "") Then
	vehicle_record_th_list.StopRec = vehicle_record_th_list.TotalRecs
Else

	' Set the last record to display
	If vehicle_record_th_list.TotalRecs > vehicle_record_th_list.StartRec + vehicle_record_th_list.DisplayRecs - 1 Then
		vehicle_record_th_list.StopRec = vehicle_record_th_list.StartRec + vehicle_record_th_list.DisplayRecs - 1
	Else
		vehicle_record_th_list.StopRec = vehicle_record_th_list.TotalRecs
	End If
End If

' Move to first record
vehicle_record_th_list.RecCnt = vehicle_record_th_list.StartRec - 1
If Not vehicle_record_th_list.Recordset.Eof Then
	vehicle_record_th_list.Recordset.MoveFirst
	If vehicle_record_th_list.StartRec > 1 Then vehicle_record_th_list.Recordset.Move vehicle_record_th_list.StartRec - 1
ElseIf Not vehicle_record_th.AllowAddDeleteRow And vehicle_record_th_list.StopRec = 0 Then
	vehicle_record_th_list.StopRec = vehicle_record_th.GridAddRowCount
End If

' Initialize Aggregate
vehicle_record_th.RowType = EW_ROWTYPE_AGGREGATEINIT
Call vehicle_record_th.ResetAttrs()
Call vehicle_record_th_list.RenderRow()
vehicle_record_th_list.RowCnt = 0

' Output date rows
Do While CLng(vehicle_record_th_list.RecCnt) < CLng(vehicle_record_th_list.StopRec)
	vehicle_record_th_list.RecCnt = vehicle_record_th_list.RecCnt + 1
	If CLng(vehicle_record_th_list.RecCnt) >= CLng(vehicle_record_th_list.StartRec) Then
		vehicle_record_th_list.RowCnt = vehicle_record_th_list.RowCnt + 1

	' Set up key count
	vehicle_record_th_list.KeyCount = vehicle_record_th_list.RowIndex
	Call vehicle_record_th.ResetAttrs()
	vehicle_record_th.CssClass = ""
	If vehicle_record_th.CurrentAction = "gridadd" Then
	Else
		Call vehicle_record_th_list.LoadRowValues(vehicle_record_th_list.Recordset) ' Load row values
	End If
	vehicle_record_th.RowType = EW_ROWTYPE_VIEW ' Render view

	' Set up row id / data-rowindex
	vehicle_record_th.RowAttrs.AddAttributes Array(Array("data-rowindex", vehicle_record_th_list.RowCnt), Array("id", "r" & vehicle_record_th_list.RowCnt & "_vehicle_record_th"), Array("data-rowtype", vehicle_record_th.RowType))

	' Render row
	Call vehicle_record_th_list.RenderRow()

	' Render list options
	Call vehicle_record_th_list.RenderListOptions()
%>
	<tr<%= vehicle_record_th.RowAttributes %>>
<%

' Render list options (body, left)
vehicle_record_th_list.ListOptions.Render "body", "left", vehicle_record_th_list.RowCnt, "", "", ""
%>
	<% If vehicle_record_th.veh_id.Visible Then ' veh_id %>
		<td<%= vehicle_record_th.veh_id.CellAttributes %>>
<span<%= vehicle_record_th.veh_id.ViewAttributes %>>
<%= vehicle_record_th.veh_id.ListViewValue %>
</span>
<a id="<%= vehicle_record_th_list.PageObjName & "_row_" & vehicle_record_th_list.RowCnt %>"></a></td>
	<% End If %>
	<% If vehicle_record_th.vch_month.Visible Then ' vch_month %>
		<td<%= vehicle_record_th.vch_month.CellAttributes %>>
<span<%= vehicle_record_th.vch_month.ViewAttributes %>>
<%= vehicle_record_th.vch_month.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.vch_year.Visible Then ' vch_year %>
		<td<%= vehicle_record_th.vch_year.CellAttributes %>>
<span<%= vehicle_record_th.vch_year.ViewAttributes %>>
<%= vehicle_record_th.vch_year.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_product_1.Visible Then ' veh_product_1 %>
		<td<%= vehicle_record_th.veh_product_1.CellAttributes %>>
<span<%= vehicle_record_th.veh_product_1.ViewAttributes %>>
<%= vehicle_record_th.veh_product_1.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_product_2.Visible Then ' veh_product_2 %>
		<td<%= vehicle_record_th.veh_product_2.CellAttributes %>>
<span<%= vehicle_record_th.veh_product_2.ViewAttributes %>>
<%= vehicle_record_th.veh_product_2.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_product_3.Visible Then ' veh_product_3 %>
		<td<%= vehicle_record_th.veh_product_3.CellAttributes %>>
<span<%= vehicle_record_th.veh_product_3.ViewAttributes %>>
<%= vehicle_record_th.veh_product_3.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_product_4.Visible Then ' veh_product_4 %>
		<td<%= vehicle_record_th.veh_product_4.CellAttributes %>>
<span<%= vehicle_record_th.veh_product_4.ViewAttributes %>>
<%= vehicle_record_th.veh_product_4.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_product_5.Visible Then ' veh_product_5 %>
		<td<%= vehicle_record_th.veh_product_5.CellAttributes %>>
<span<%= vehicle_record_th.veh_product_5.ViewAttributes %>>
<%= vehicle_record_th.veh_product_5.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_product_6.Visible Then ' veh_product_6 %>
		<td<%= vehicle_record_th.veh_product_6.CellAttributes %>>
<span<%= vehicle_record_th.veh_product_6.ViewAttributes %>>
<%= vehicle_record_th.veh_product_6.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_product_7.Visible Then ' veh_product_7 %>
		<td<%= vehicle_record_th.veh_product_7.CellAttributes %>>
<span<%= vehicle_record_th.veh_product_7.ViewAttributes %>>
<%= vehicle_record_th.veh_product_7.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_product_8.Visible Then ' veh_product_8 %>
		<td<%= vehicle_record_th.veh_product_8.CellAttributes %>>
<span<%= vehicle_record_th.veh_product_8.ViewAttributes %>>
<%= vehicle_record_th.veh_product_8.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_domes_1.Visible Then ' veh_domes_1 %>
		<td<%= vehicle_record_th.veh_domes_1.CellAttributes %>>
<span<%= vehicle_record_th.veh_domes_1.ViewAttributes %>>
<%= vehicle_record_th.veh_domes_1.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_domes_2.Visible Then ' veh_domes_2 %>
		<td<%= vehicle_record_th.veh_domes_2.CellAttributes %>>
<span<%= vehicle_record_th.veh_domes_2.ViewAttributes %>>
<%= vehicle_record_th.veh_domes_2.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_domes_3.Visible Then ' veh_domes_3 %>
		<td<%= vehicle_record_th.veh_domes_3.CellAttributes %>>
<span<%= vehicle_record_th.veh_domes_3.ViewAttributes %>>
<%= vehicle_record_th.veh_domes_3.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_domes_4.Visible Then ' veh_domes_4 %>
		<td<%= vehicle_record_th.veh_domes_4.CellAttributes %>>
<span<%= vehicle_record_th.veh_domes_4.ViewAttributes %>>
<%= vehicle_record_th.veh_domes_4.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_domes_5.Visible Then ' veh_domes_5 %>
		<td<%= vehicle_record_th.veh_domes_5.CellAttributes %>>
<span<%= vehicle_record_th.veh_domes_5.ViewAttributes %>>
<%= vehicle_record_th.veh_domes_5.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_domes_6.Visible Then ' veh_domes_6 %>
		<td<%= vehicle_record_th.veh_domes_6.CellAttributes %>>
<span<%= vehicle_record_th.veh_domes_6.ViewAttributes %>>
<%= vehicle_record_th.veh_domes_6.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_domes_7.Visible Then ' veh_domes_7 %>
		<td<%= vehicle_record_th.veh_domes_7.CellAttributes %>>
<span<%= vehicle_record_th.veh_domes_7.ViewAttributes %>>
<%= vehicle_record_th.veh_domes_7.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_domes_8.Visible Then ' veh_domes_8 %>
		<td<%= vehicle_record_th.veh_domes_8.CellAttributes %>>
<span<%= vehicle_record_th.veh_domes_8.ViewAttributes %>>
<%= vehicle_record_th.veh_domes_8.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_export_1.Visible Then ' veh_export_1 %>
		<td<%= vehicle_record_th.veh_export_1.CellAttributes %>>
<span<%= vehicle_record_th.veh_export_1.ViewAttributes %>>
<%= vehicle_record_th.veh_export_1.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_export_2.Visible Then ' veh_export_2 %>
		<td<%= vehicle_record_th.veh_export_2.CellAttributes %>>
<span<%= vehicle_record_th.veh_export_2.ViewAttributes %>>
<%= vehicle_record_th.veh_export_2.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_export_3.Visible Then ' veh_export_3 %>
		<td<%= vehicle_record_th.veh_export_3.CellAttributes %>>
<span<%= vehicle_record_th.veh_export_3.ViewAttributes %>>
<%= vehicle_record_th.veh_export_3.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_export_4.Visible Then ' veh_export_4 %>
		<td<%= vehicle_record_th.veh_export_4.CellAttributes %>>
<span<%= vehicle_record_th.veh_export_4.ViewAttributes %>>
<%= vehicle_record_th.veh_export_4.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_export_5.Visible Then ' veh_export_5 %>
		<td<%= vehicle_record_th.veh_export_5.CellAttributes %>>
<span<%= vehicle_record_th.veh_export_5.ViewAttributes %>>
<%= vehicle_record_th.veh_export_5.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_export_6.Visible Then ' veh_export_6 %>
		<td<%= vehicle_record_th.veh_export_6.CellAttributes %>>
<span<%= vehicle_record_th.veh_export_6.ViewAttributes %>>
<%= vehicle_record_th.veh_export_6.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_export_7.Visible Then ' veh_export_7 %>
		<td<%= vehicle_record_th.veh_export_7.CellAttributes %>>
<span<%= vehicle_record_th.veh_export_7.ViewAttributes %>>
<%= vehicle_record_th.veh_export_7.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_export_8.Visible Then ' veh_export_8 %>
		<td<%= vehicle_record_th.veh_export_8.CellAttributes %>>
<span<%= vehicle_record_th.veh_export_8.ViewAttributes %>>
<%= vehicle_record_th.veh_export_8.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_month_title.Visible Then ' veh_month_title %>
		<td<%= vehicle_record_th.veh_month_title.CellAttributes %>>
<span<%= vehicle_record_th.veh_month_title.ViewAttributes %>>
<%= vehicle_record_th.veh_month_title.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_range.Visible Then ' veh_range %>
		<td<%= vehicle_record_th.veh_range.CellAttributes %>>
<span<%= vehicle_record_th.veh_range.ViewAttributes %>>
<%= vehicle_record_th.veh_range.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_month_title2.Visible Then ' veh_month_title2 %>
		<td<%= vehicle_record_th.veh_month_title2.CellAttributes %>>
<span<%= vehicle_record_th.veh_month_title2.ViewAttributes %>>
<%= vehicle_record_th.veh_month_title2.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If vehicle_record_th.veh_range2.Visible Then ' veh_range2 %>
		<td<%= vehicle_record_th.veh_range2.CellAttributes %>>
<span<%= vehicle_record_th.veh_range2.ViewAttributes %>>
<%= vehicle_record_th.veh_range2.ListViewValue %>
</span>
</td>
	<% End If %>
<%

' Render list options (body, right)
vehicle_record_th_list.ListOptions.Render "body", "right", vehicle_record_th_list.RowCnt, "", "", ""
%>
	</tr>
<%
	End If
	If vehicle_record_th.CurrentAction <> "gridadd" Then
		vehicle_record_th_list.Recordset.MoveNext()
	End If
Loop
%>
</tbody>
</table>
<% End If %>
<% If vehicle_record_th.CurrentAction = "" Then %>
<input type="hidden" name="a_list" id="a_list" value="">
<% End If %>
</div>
</form>
<%

' Close recordset and connection
vehicle_record_th_list.Recordset.Close
Set vehicle_record_th_list.Recordset = Nothing
%>
<% If vehicle_record_th.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If vehicle_record_th.CurrentAction <> "gridadd" And vehicle_record_th.CurrentAction <> "gridedit" Then %>
<form name="ewPagerForm" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewPager">
<tr><td>
<% If Not IsObject(vehicle_record_th_list.Pager) Then Set vehicle_record_th_list.Pager = ew_NewPrevNextPager(vehicle_record_th_list.StartRec, vehicle_record_th_list.DisplayRecs, vehicle_record_th_list.TotalRecs) %>
<% If vehicle_record_th_list.Pager.RecordCount > 0 Then %>
<table class="ewStdTable"><tbody><tr><td>
	<%= Language.Phrase("Page") %>&nbsp;
<div class="input-prepend input-append">
<!--first page button-->
	<% If vehicle_record_th_list.Pager.FirstButton.Enabled Then %>
	<a class="btn btn-small" href="<%= vehicle_record_th_list.PageUrl %>start=<%= vehicle_record_th_list.Pager.FirstButton.Start %>"><i class="icon-step-backward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-backward"></i></a>
	<% End If %>
<!--previous page button-->
	<% If vehicle_record_th_list.Pager.PrevButton.Enabled Then %>
	<a class="btn btn-small" href="<%= vehicle_record_th_list.PageUrl %>start=<%= vehicle_record_th_list.Pager.PrevButton.Start %>"><i class="icon-prev"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-prev"></i></a>
	<% End If %>
<!--current page number-->
	<input class="input-mini" type="text" name="<%= EW_TABLE_PAGE_NO %>" value="<%= vehicle_record_th_list.Pager.CurrentPage %>">
<!--next page button-->
	<% If vehicle_record_th_list.Pager.NextButton.Enabled Then %>
	<a class="btn btn-small" href="<%= vehicle_record_th_list.PageUrl %>start=<%= vehicle_record_th_list.Pager.NextButton.Start %>"><i class="icon-play"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-play"></i></a>
	<% End If %>
<!--last page button-->
	<% If vehicle_record_th_list.Pager.LastButton.Enabled Then %>
	<a class="btn btn-small" href="<%= vehicle_record_th_list.PageUrl %>start=<%= vehicle_record_th_list.Pager.LastButton.Start %>"><i class="icon-step-forward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-forward"></i></a>
	<% End If %>
</div>
	&nbsp;<%= Language.Phrase("of") %>&nbsp;<%= vehicle_record_th_list.Pager.PageCount %>
</td>
<td>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<%= Language.Phrase("Record") %>&nbsp;<%= vehicle_record_th_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= vehicle_record_th_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= vehicle_record_th_list.Pager.RecordCount %>
</td>
</tr></tbody></table>
<% Else %>
	<% If vehicle_record_th_list.SearchWhere = "0=101" Then %>
	<p><%= Language.Phrase("EnterSearchCriteria") %></p>
	<% Else %>
	<p><%= Language.Phrase("NoRecord") %></p>
	<% End If %>
<% End If %>
</td>
</tr></table>
</form>
<% End If %>
<div class="ewListOtherOptions">
<%
	vehicle_record_th_list.AddEditOptions.Render "body", "bottom", "", "", "", ""
	vehicle_record_th_list.DetailOptions.Render "body", "bottom", "", "", "", ""
	vehicle_record_th_list.ActionOptions.Render "body", "bottom", "", "", "", ""
%>
</div>
</div>
<% End If %>
</td></tr></table>
<% If vehicle_record_th.Export = "" Then %>
<script type="text/javascript">
fvehicle_record_thlistsrch.Init();
fvehicle_record_thlist.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<% End If %>
<%
vehicle_record_th_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If vehicle_record_th.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set vehicle_record_th_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cvehicle_record_th_list

	' Page ID
	Public Property Get PageID()
		PageID = "list"
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
		PageObjName = "vehicle_record_th_list"
	End Property

	' Grid form hidden field names
	Dim FormName
	Dim FormActionName
	Dim FormKeyName
	Dim FormOldKeyName
	Dim FormBlankRowName
	Dim FormKeyCountName

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If vehicle_record_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & vehicle_record_th.TableVar & "&" ' add page token
	End Property

	' Common urls
	Dim AddUrl
	Dim EditUrl
	Dim CopyUrl
	Dim DeleteUrl
	Dim ViewUrl
	Dim ListUrl

	' Export urls
	Dim ExportPrintUrl
	Dim ExportHtmlUrl
	Dim ExportExcelUrl
	Dim ExportWordUrl
	Dim ExportXmlUrl
	Dim ExportCsvUrl
	Dim ExportPdfUrl

	' Inline urls
	Dim InlineAddUrl
	Dim InlineCopyUrl
	Dim InlineEditUrl
	Dim GridAddUrl
	Dim GridEditUrl
	Dim MultiDeleteUrl
	Dim MultiUpdateUrl

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

		' Grid form hidden field names
		FormName = "fvehicle_record_thlist"
		FormActionName = "k_action"
		FormKeyName = "k_key"
		FormOldKeyName = "k_oldkey"
		FormBlankRowName = "k_blankrow"
		FormKeyCountName = "key_count"

		' Initialize language object
		If IsEmpty(Language) Then
			Set Language = New cLanguage
			Call Language.LoadPhrases()
		End If

		' Initialize table object
		If IsEmpty(vehicle_record_th) Then Set vehicle_record_th = New cvehicle_record_th
		Set Table = vehicle_record_th

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		ExportPdfUrl = PageUrl & "export=pdf"
		AddUrl = "pom_vehicle_record_thadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "pom_vehicle_record_thdelete.asp"
		MultiUpdateUrl = "pom_vehicle_record_thupdate.asp"

		' Initialize other table object
		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "vehicle_record_th"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' List options
		Set ListOptions = New cListOptions
		ListOptions.TableVar = vehicle_record_th.TableVar

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = vehicle_record_th.TableVar
		ExportOptions.Tag = "div"
		ExportOptions.TagClassName = "ewExportOption"

		' Other options
		Set AddEditOptions = New cListOptions
		AddEditOptions.Tag = "div"
		AddEditOptions.TagClassName = "ewAddEditOption"
		Set DetailOptions = New cListOptions
		DetailOptions.Tag = "div"
		DetailOptions.TagClassName = "ewDetailOption"
		Set ActionOptions = New cListOptions
		ActionOptions.Tag = "div"
		ActionOptions.TagClassName = "ewActionOption"
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

		' Get grid add count
		Dim gridaddcnt
		gridaddcnt = Request.QueryString(EW_TABLE_GRID_ADD_ROW_COUNT)
		If IsNumeric(gridaddcnt) Then
			If gridaddcnt > 0 Then
				vehicle_record_th.GridAddRowCount = gridaddcnt
			End If
		End If

		' Set up list options
		SetupListOptions()

		' Global page loading event (in userfn7.asp)
		Page_Loading()

		' Page load event, used in current page
		Page_Load()

		' Setup other options
		SetupOtherOptions()

		' Set "checkbox" visible
		If UBound(vehicle_record_th.CustomActions.CustomArray) >= 0 Then
			ListOptions.GetItem("checkbox").Visible = True
		End If
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
		Set ListOptions = Nothing
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

	Dim ListOptions ' List options
	Dim ExportOptions ' Export options
	Dim AddEditOptions ' Other options (add edit)
	Dim DetailOptions ' Other options (detail)
	Dim ActionOptions ' Other options (action)
	Dim DisplayRecs ' Number of display records
	Dim StartRec, StopRec, TotalRecs, RecRange
	Dim SearchWhere
	Dim RecCnt
	Dim EditRowCnt
	Dim StartRowCnt
	Dim RowCnt, RowIndex
	Dim Attrs
	Dim RecPerRow, ColCnt
	Dim KeyCount
	Dim RowAction
	Dim RowOldKey ' Row old key (for copy)
	Dim DbMasterFilter, DbDetailFilter
	Dim MasterRecordExists
	Dim MultiSelectKey
	Dim Command
	Dim RestoreSearch
	Dim Recordset, OldRecordset

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		DisplayRecs = 50
		RecRange = 10
		RecCnt = 0 ' Record count
		KeyCount = 0 ' Key count
		StartRowCnt = 1

		' Search filters
		Dim sSrchAdvanced, sSrchBasic, sFilter
		sSrchAdvanced = "" ' Advanced search filter
		sSrchBasic = "" ' Basic search filter
		SearchWhere = "" ' Search where clause
		sFilter = ""

		' Restore search
		RestoreSearch = False

		' Get command
		Command = LCase(Request.QueryString("cmd")&"")

		' Master/Detail
		DbMasterFilter = "" ' Master filter
		DbDetailFilter = "" ' Detail filter
		If IsPageRequest Then ' Validate request

			' Process custom action first
			ProcessCustomAction()

			' Handle reset command
			ResetCmd()

			' Set up Breadcrumb
			If vehicle_record_th.Export = "" Then
				SetupBreadcrumb()
			End If

			' Hide list options
			If vehicle_record_th.Export <> "" Then
				Call ListOptions.HideAllOptions(Array("sequence"))
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			ElseIf vehicle_record_th.CurrentAction = "gridadd" Or vehicle_record_th.CurrentAction = "gridedit" Then
				Call ListOptions.HideAllOptions(Array())
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			End If

			' Hide export options
			If vehicle_record_th.Export <> "" Or vehicle_record_th.CurrentAction <> "" Then
				ExportOptions.HideAllOptions(Array())
			End If

			' Hide other options
			If vehicle_record_th.Export <> "" Then
				AddEditOptions.HideAllOptions(Array())
				DetailOptions.HideAllOptions(Array())
				ActionOptions.HideAllOptions(Array())
			End If

			' Get basic search values
			Call LoadBasicSearchValues()

			' Restore search parms from Session if not searching / reset
			If Command <> "search" And Command <> "reset" And Command <> "resetall" And CheckSearchParms() Then
				Call RestoreSearchParms()
			End If

			' Call Recordset SearchValidated event
			Call vehicle_record_th.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If vehicle_record_th.RecordsPerPage <> "" Then
			DisplayRecs = vehicle_record_th.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Load search default if no existing search criteria
		If Not CheckSearchParms() Then

			' Load basic search from default
			vehicle_record_th.BasicSearch.Keyword = vehicle_record_th.BasicSearch.KeywordDefault
			vehicle_record_th.BasicSearch.SearchType = vehicle_record_th.BasicSearch.SearchTypeDefault
			vehicle_record_th.BasicSearch.setSearchType(vehicle_record_th.BasicSearch.SearchTypeDefault)
			If vehicle_record_th.BasicSearch.Keyword <> "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call vehicle_record_th.Recordset_Searching(SearchWhere)

		' Save search criteria
		If Command = "search" And Not RestoreSearch Then
			vehicle_record_th.SearchWhere = SearchWhere ' Save to Session
			StartRec = 1 ' Reset start record counter
			vehicle_record_th.StartRecordNumber = StartRec
		Else
			SearchWhere = vehicle_record_th.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		vehicle_record_th.SessionWhere = sFilter
		vehicle_record_th.CurrentFilter = ""
	End Sub

	' -----------------------------------------------------------------
	'  Build filter for all keys
	'
	Function BuildKeyFilter()
		Dim rowindex, sThisKey
		Dim sKey
		Dim sWrkFilter, sFilter
		sWrkFilter = ""

		' Update row index and get row key
		rowindex = 1
		ObjForm.Index = rowindex
		sThisKey = ObjForm.GetValue("k_key") & ""
		Do While (sThisKey <> "")
			If SetupKeyValues(sThisKey) Then
				sFilter = vehicle_record_th.KeyFilter
				If sWrkFilter <> "" Then sWrkFilter = sWrkFilter & " OR "
				sWrkFilter = sWrkFilter & sFilter
			Else
				sWrkFilter = "0=1"
				Exit Do
			End If

			' Update row index and get row key
			rowindex = rowindex + 1 ' Next row
			ObjForm.Index = rowindex
			sThisKey = ObjForm.GetValue("k_key") & ""
		Loop
		BuildKeyFilter = sWrkFilter
	End Function

	' -----------------------------------------------------------------
	' Set up key values
	'
	Function SetupKeyValues(key)
		Dim arrKeyFlds
		arrKeyFlds = Split(key&"", EW_COMPOSITE_KEY_SEPARATOR)
		If UBound(arrKeyFlds) >= 0 Then
			vehicle_record_th.veh_id.FormValue = arrKeyFlds(0)
			If Not IsNumeric(vehicle_record_th.veh_id.FormValue) Then
				SetupKeyValues = False
				Exit Function
			End If
		End If
		SetupKeyValues = True
	End Function

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.vch_month, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.vch_year, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_product_1, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_product_2, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_product_3, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_product_4, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_product_5, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_product_6, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_product_7, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_product_8, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_domes_1, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_domes_2, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_domes_3, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_domes_4, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_domes_5, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_domes_6, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_domes_7, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_domes_8, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_export_1, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_export_2, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_export_3, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_export_4, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_export_5, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_export_6, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_export_7, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_export_8, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_remark, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_month_title, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_range, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_month_title2, Keyword)
			Call BuildBasicSearchSQL(sWhere, vehicle_record_th.veh_range2, Keyword)
		BasicSearchSQL = sWhere
	End Function

	' -----------------------------------------------------------------
	' Build basic search sql
	'
	Sub BuildBasicSearchSql(Where, Fld, Keyword)
		Dim sFldExpression, lFldDataType
		Dim sWrk
		If Keyword = EW_NULL_VALUE Then
			sWrk = Fld.FldExpression & " IS NULL"
		ElseIf Keyword = EW_NOT_NULL_VALUE Then
			sWrk = Fld.FldExpression & " IS NOT NULL"
		Else
			If Fld.FldVirtualExpression <> Fld.FldExpression Then
				sFldExpression = Fld.FldVirtualExpression
			Else
				sFldExpression = Fld.FldBasicSearchExpression
			End If
			sWrk = sFldExpression & ew_Like(ew_QuotedValue("%" & Keyword & "%", EW_DATATYPE_STRING))
		End If
		If Where <> "" Then Where = Where & " OR "
		Where = Where & sWrk
	End Sub

	' -----------------------------------------------------------------
	' Return Basic Search Where based on search keyword and type
	'
	Function BasicSearchWhere()
		Dim sSearchStr, sSearchKeyword, sSearchType
		Dim sSearch, arKeyword, sKeyword
		sSearchStr = ""
		sSearchKeyword = vehicle_record_th.BasicSearch.Keyword
		sSearchType = vehicle_record_th.BasicSearch.SearchType
		If sSearchKeyword <> "" Then
			sSearch = Trim(sSearchKeyword)
			If sSearchType <> "=" Then
				While InStr(sSearch, "  ") > 0
					sSearch = Replace(sSearch, "  ", " ")
				Wend
				arKeyword = Split(Trim(sSearch), " ")
				For Each sKeyword In arKeyword
					If sSearchStr <> "" Then sSearchStr = sSearchStr & " " & sSearchType & " "
					sSearchStr = sSearchStr & "(" & BasicSearchSQL(sKeyword) & ")"
				Next
			Else
				sSearchStr = BasicSearchSQL(sSearch)
			End If
			Command = "search"
		End If
		If Command = "search" Then
			vehicle_record_th.BasicSearch.setKeyword(sSearchKeyword)
			vehicle_record_th.BasicSearch.setSearchType(sSearchType)
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' Check if search parm exists
	Function CheckSearchParms()

		' Check basic search
		If vehicle_record_th.BasicSearch.IssetSession() Then
			CheckSearchParms = True
			Exit Function
		End If
		CheckSearchParms = False
	End Function

	' Clear all search parameters
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		vehicle_record_th.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' Load advanced search default values
	Function LoadAdvancedSearchDefault()
		LoadAdvancedSearchDefault = False
	End Function

	' Clear all basic search parameters
	Sub ResetBasicSearchParms()
		vehicle_record_th.BasicSearch.UnsetSession()
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()

		' Restore search flag
		RestoreSearch = True

		' Restore basic search values
		Call vehicle_record_th.BasicSearch.Load()
	End Sub

	' -----------------------------------------------------------------
	' Set up Sort parameters based on Sort Links clicked
	'
	Sub SetUpSortOrder()
		Dim sOrderBy
		Dim sSortField, sLastSort, sThisSort
		Dim bCtrl

		' Check for an Order parameter
		If Request.QueryString("order").Count > 0 Then
			vehicle_record_th.CurrentOrder = Request.QueryString("order")
			vehicle_record_th.CurrentOrderType = Request.QueryString("ordertype")

			' Field veh_id
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_id)

			' Field vch_month
			Call vehicle_record_th.UpdateSort(vehicle_record_th.vch_month)

			' Field vch_year
			Call vehicle_record_th.UpdateSort(vehicle_record_th.vch_year)

			' Field veh_product_1
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_product_1)

			' Field veh_product_2
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_product_2)

			' Field veh_product_3
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_product_3)

			' Field veh_product_4
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_product_4)

			' Field veh_product_5
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_product_5)

			' Field veh_product_6
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_product_6)

			' Field veh_product_7
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_product_7)

			' Field veh_product_8
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_product_8)

			' Field veh_domes_1
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_domes_1)

			' Field veh_domes_2
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_domes_2)

			' Field veh_domes_3
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_domes_3)

			' Field veh_domes_4
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_domes_4)

			' Field veh_domes_5
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_domes_5)

			' Field veh_domes_6
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_domes_6)

			' Field veh_domes_7
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_domes_7)

			' Field veh_domes_8
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_domes_8)

			' Field veh_export_1
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_export_1)

			' Field veh_export_2
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_export_2)

			' Field veh_export_3
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_export_3)

			' Field veh_export_4
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_export_4)

			' Field veh_export_5
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_export_5)

			' Field veh_export_6
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_export_6)

			' Field veh_export_7
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_export_7)

			' Field veh_export_8
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_export_8)

			' Field veh_month_title
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_month_title)

			' Field veh_range
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_range)

			' Field veh_month_title2
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_month_title2)

			' Field veh_range2
			Call vehicle_record_th.UpdateSort(vehicle_record_th.veh_range2)
			vehicle_record_th.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = vehicle_record_th.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If vehicle_record_th.SqlOrderBy <> "" Then
				sOrderBy = vehicle_record_th.SqlOrderBy
				vehicle_record_th.SessionOrderBy = sOrderBy
			End If
		End If
	End Sub

	' -----------------------------------------------------------------
	' Reset command based on querystring parameter cmd=
	' - RESET: reset search parameters
	' - RESETALL: reset search & master/detail parameters
	' - RESETSORT: reset sort parameters
	'
	Sub ResetCmd()

		' Check if reset command
		If Left(Command,5) = "reset" Then

			' Reset search criteria
			If Command = "reset" Or Command = "resetall" Then
				Call ResetSearchParms()
			End If

			' Reset Sort Criteria
			If Command = "resetsort" Then
				Dim sOrderBy
				sOrderBy = ""
				vehicle_record_th.SessionOrderBy = sOrderBy
				vehicle_record_th.veh_id.Sort = ""
				vehicle_record_th.vch_month.Sort = ""
				vehicle_record_th.vch_year.Sort = ""
				vehicle_record_th.veh_product_1.Sort = ""
				vehicle_record_th.veh_product_2.Sort = ""
				vehicle_record_th.veh_product_3.Sort = ""
				vehicle_record_th.veh_product_4.Sort = ""
				vehicle_record_th.veh_product_5.Sort = ""
				vehicle_record_th.veh_product_6.Sort = ""
				vehicle_record_th.veh_product_7.Sort = ""
				vehicle_record_th.veh_product_8.Sort = ""
				vehicle_record_th.veh_domes_1.Sort = ""
				vehicle_record_th.veh_domes_2.Sort = ""
				vehicle_record_th.veh_domes_3.Sort = ""
				vehicle_record_th.veh_domes_4.Sort = ""
				vehicle_record_th.veh_domes_5.Sort = ""
				vehicle_record_th.veh_domes_6.Sort = ""
				vehicle_record_th.veh_domes_7.Sort = ""
				vehicle_record_th.veh_domes_8.Sort = ""
				vehicle_record_th.veh_export_1.Sort = ""
				vehicle_record_th.veh_export_2.Sort = ""
				vehicle_record_th.veh_export_3.Sort = ""
				vehicle_record_th.veh_export_4.Sort = ""
				vehicle_record_th.veh_export_5.Sort = ""
				vehicle_record_th.veh_export_6.Sort = ""
				vehicle_record_th.veh_export_7.Sort = ""
				vehicle_record_th.veh_export_8.Sort = ""
				vehicle_record_th.veh_month_title.Sort = ""
				vehicle_record_th.veh_range.Sort = ""
				vehicle_record_th.veh_month_title2.Sort = ""
				vehicle_record_th.veh_range2.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			vehicle_record_th.StartRecordNumber = StartRec
		End If
	End Sub

	' Set up list options
	Sub SetupListOptions()
		Dim item

		' Add group option item
		ListOptions.Add(ListOptions.GroupOptionName)
		Set item = ListOptions.GetItem(ListOptions.GroupOptionName)
		item.Body = ""
		item.OnLeft = False
		item.Visible = False

		' View
		ListOptions.Add("view")
		Set item = ListOptions.GetItem("view")
		item.CssStyle = "white-space: nowrap;"
		item.Visible = Security.IsLoggedIn()
		item.OnLeft = False

		' Edit
		ListOptions.Add("edit")
		Set item = ListOptions.GetItem("edit")
		item.CssStyle = "white-space: nowrap;"
		item.Visible = Security.IsLoggedIn()
		item.OnLeft = False

		' Copy
		ListOptions.Add("copy")
		Set item = ListOptions.GetItem("copy")
		item.CssStyle = "white-space: nowrap;"
		item.Visible = Security.IsLoggedIn()
		item.OnLeft = False

		' Delete
		ListOptions.Add("delete")
		Set item = ListOptions.GetItem("delete")
		item.CssStyle = "white-space: nowrap;"
		item.Visible = Security.IsLoggedIn()
		item.OnLeft = False

		' Checkbox
		ListOptions.Add("checkbox")
		Set item = ListOptions.GetItem("checkbox")
		item.Visible = False
		item.OnLeft = False
		item.Header = "<label class=""checkbox""><input type=""checkbox"" name=""key"" id=""key"" onclick=""ew_SelectAllKey(this);""></label>"
		item.ShowInDropDown = False
		item.ShowInButtonGroup = False

		' Drop down button for ListOptions
		ListOptions.UseDropDownButton = False
		ListOptions.DropDownButtonPhrase = Language.Phrase("ButtonListOptions")
		ListOptions.UseButtonGroup = False
		ListOptions.ButtonClass = "btn-small" ' Class for button group
		Call ListOptions_Load()

		' Set up group item visibility
		ListOptions.GetItem(ListOptions.GroupOptionName).Visible = ListOptions.GroupOptionVisible
	End Sub

	' Render list options
	Sub RenderListOptions()
		Dim item, links
		ListOptions.LoadDefault()
		If Security.IsLoggedIn() Then
			ListOptions.GetItem("view").Body = "<a class=""ewRowLink ewView"" data-caption=""" & ew_HtmlTitle(Language.Phrase("ViewLink")) & """ href=""" & ew_HtmlEncode(ViewUrl) & """>" & Language.Phrase("ViewLink") & "</a>"
		Else
			ListOptions.GetItem("view").Body = ""
		End If
		Set item = ListOptions.GetItem("edit")
		If Security.IsLoggedIn() Then
			item.Body = "<a class=""ewRowLink ewEdit"" data-caption=""" & ew_HtmlTitle(Language.Phrase("EditLink")) & """ href=""" & ew_HtmlEncode(EditUrl) & """>" & Language.Phrase("EditLink") & "</a>"
		Else
			item.Body = ""
		End If
		Set item = ListOptions.GetItem("copy")
		If Security.IsLoggedIn() Then
			item.Body = "<a class=""ewRowLink ewCopy"" data-caption=""" & ew_HtmlTitle(Language.Phrase("CopyLink")) & """ href=""" & ew_HtmlEncode(CopyUrl) & """>" & Language.Phrase("CopyLink") & "</a>"
		Else
			item.Body = ""
		End If
		If Security.IsLoggedIn() Then
			ListOptions.GetItem("delete").Body = "<a class=""ewRowLink ewDelete""" & "" & " data-caption=""" & ew_HtmlTitle(Language.Phrase("DeleteLink")) & """ href=""" & ew_HtmlEncode(DeleteUrl) & """>" & Language.Phrase("DeleteLink") & "</a>"
		Else
			ListOptions.GetItem("delete").Body = ""
		End If
		ListOptions.GetItem("checkbox").Body = "<label class=""checkbox""><input type=""checkbox"" name=""key_m"" value=""" & ew_HtmlEncode(vehicle_record_th.veh_id.CurrentValue) & """ onclick='ew_ClickMultiCheckbox(event, this);'></label>"
		Call RenderListOptionsExt()
		Call ListOptions_Rendered()
	End Sub

	' Set up other options
	Sub SetupOtherOptions()
		Dim opt, item, DetailTableLink, ar, i
		Set opt = AddEditOptions

		' Add
		Call opt.Add("add")
		Set item = opt.GetItem("add")
		item.Body = "<a class=""ewAddEdit ewAdd"" href=""" & ew_HtmlEncode(AddUrl) & """>" & Language.Phrase("AddLink") & "</a>"
		item.Visible = (AddUrl <> "" And Security.IsLoggedIn())
		Set opt = ActionOptions

		' Set up options default
		Set opt = AddEditOptions
		opt.DropDownButtonPhrase = Language.Phrase("ButtonAddEdit")
		opt.UseDropDownButton = False
		opt.UseButtonGroup = True
		opt.ButtonClass = "btn-small" ' Class for button group
		Call opt.Add(opt.GroupOptionName)
		Set item = opt.GetItem(opt.GroupOptionName)
		item.Body = ""
		item.Visible = False
		Set opt = DetailOptions
		opt.DropDownButtonPhrase = Language.Phrase("ButtonDetails")
		opt.UseDropDownButton = False
		opt.UseButtonGroup = True
		opt.ButtonClass = "btn-small" ' Class for button group
		Call opt.Add(opt.GroupOptionName)
		Set item = opt.GetItem(opt.GroupOptionName)
		item.Body = ""
		item.Visible = False
		Set opt = ActionOptions
		opt.DropDownButtonPhrase = Language.Phrase("ButtonActions")
		opt.UseDropDownButton = False
		opt.UseButtonGroup = True
		opt.ButtonClass = "btn-small" ' Class for button group
		Call opt.Add(opt.GroupOptionName)
		Set item = opt.GetItem(opt.GroupOptionName)
		item.Body = ""
		item.Visible = False
	End Sub

	' Render other options
	Sub RenderOtherOptions()
		Dim opt, item, i, Action, Name
			Set opt = ActionOptions
			For i = 0 to UBound(vehicle_record_th.CustomActions.CustomArray)
				Action = vehicle_record_th.CustomActions.CustomArray(i)(0)
				Name = vehicle_record_th.CustomActions.CustomArray(i)(1)

				' Add custom action
				Call opt.Add("custom_" & Action)
				Set item = opt.GetItem("custom_" & Action)
				item.Body = "<a class=""ewAction ewCustomAction"" href="""" onclick=""ew_SubmitSelected(document.fvehicle_record_thlist, '" & ew_CurrentUrl & "', null, '" & Action & "');return false;"">" & Name & "</a>"
			Next

			' Hide grid edit, multi-delete and multi-update
			If TotalRecs <= 0 Then
				Set opt = AddEditOptions
				Set item = opt.GetItem("gridedit")
				If (Not item Is Nothing) Then item.Visible = False
				Set opt = ActionOptions
				Set item = opt.GetItem("multidelete")
				If (Not item Is Nothing) Then item.Visible = False
				Set item = opt.GetItem("multiupdate")
				If (Not item Is Nothing) Then item.Visible = False
			End If
	End Sub

	' Process custom action
	Sub ProcessCustomAction()
		Dim sFilter, sSql, UserAction, Processed
		sFilter = vehicle_record_th.GetKeyFilter
		UserAction = Request.Form("useraction") & ""
		Processed = False
		If sFilter <> "" And UserAction <> "" Then
			vehicle_record_th.CurrentFilter = sFilter
			sSql = vehicle_record_th.SQL
			Conn.BeginTrans

			' Load recordset
			Dim Rs
			Set Rs = ew_LoadRecordset(sSql)
			If Not Rs.Eof Then Rs.MoveFirst

			' Call row custom action event
			Do While Not Rs.Eof
				Processed = Row_CustomAction(UserAction, Rs)
				If Not Processed Then
					Exit Do
				Else
					Rs.MoveNext
				End If
			Loop
			Rs.Close
			Set Rs = Nothing
			If Processed Then
				Conn.CommitTrans ' Commit the changes
				If SuccessMessage = "" Then
					SuccessMessage = Replace(Language.Phrase("CustomActionCompleted"), "%s", UserAction) ' Set up success message
				End If
			Else
				Conn.RollbackTrans ' Rollback transaction

				' Set up error message
				If SuccessMessage <> "" Or FailureMessage <> "" Then

					' Use the message, do nothing
				ElseIf vehicle_record_th.CancelMessage <> "" Then
					FailureMessage = vehicle_record_th.CancelMessage
					vehicle_record_th.CancelMessage = ""
				Else
					FailureMessage = Replace(Language.Phrase("CustomActionCancelled"), "%s", UserAction)
				End If
			End If
		End If
	End Sub

	Function RenderListOptionsExt()
	End Function
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
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		vehicle_record_th.BasicSearch.Keyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)&""
		If vehicle_record_th.BasicSearch.Keyword <> "" Then Command = "search"
		vehicle_record_th.BasicSearch.SearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)&""
	End Function

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

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If vehicle_record_th.GetKey("veh_id")&"" <> "" Then
			vehicle_record_th.veh_id.CurrentValue = vehicle_record_th.GetKey("veh_id") ' veh_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			vehicle_record_th.CurrentFilter = vehicle_record_th.KeyFilter
			Dim sSql
			sSql = vehicle_record_th.SQL
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
		ViewUrl = vehicle_record_th.ViewUrl("")
		EditUrl = vehicle_record_th.EditUrl("")
		InlineEditUrl = vehicle_record_th.InlineEditUrl
		CopyUrl = vehicle_record_th.CopyUrl("")
		InlineCopyUrl = vehicle_record_th.InlineCopyUrl
		DeleteUrl = vehicle_record_th.DeleteUrl

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

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		url = ew_CurrentUrl
		url = ew_RegExReplace("\?cmd=reset(all){0,1}$", url, "") ' Remove cmd=reset / cmd=resetall
		Call Breadcrumb.Add("list", vehicle_record_th.TableVar, url, vehicle_record_th.TableVar, True)
	End Sub

	Sub ExportPdf(html)
		Response.Write html
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

	' ListOptions Load event
	Sub ListOptions_Load()

		'Example: 
		' Dim opt
		' Set opt = ListOptions.Add("new")
		' opt.OnLeft = True ' Link on left
		' opt.MoveTo 0 ' Move to first column

	End Sub

	' ListOptions Rendered event
	Sub ListOptions_Rendered()

		'Example: 
		'ListOptions.GetItem("new").Body = "xxx"

	End Sub

	' Row Custom Action event
	Function Row_CustomAction(action, rs)

		' Return False to abort
		Row_CustomAction = True
	End Function
End Class
%>
