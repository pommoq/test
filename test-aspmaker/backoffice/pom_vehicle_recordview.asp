<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_vehicle_recordinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim vehicle_record_view
Set vehicle_record_view = New cvehicle_record_view
Set Page = vehicle_record_view

' Page init processing
vehicle_record_view.Page_Init()

' Page main processing
vehicle_record_view.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
vehicle_record_view.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If vehicle_record.Export = "" Then %>
<script type="text/javascript">
// Page object
var vehicle_record_view = new ew_Page("vehicle_record_view");
vehicle_record_view.PageID = "view"; // Page ID
var EW_PAGE_ID = vehicle_record_view.PageID; // For backward compatibility
// Form object
var fvehicle_recordview = new ew_Form("fvehicle_recordview");
// Form_CustomValidate event
fvehicle_recordview.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fvehicle_recordview.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fvehicle_recordview.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If vehicle_record.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If vehicle_record.Export = "" Then %>
<div class="ewViewExportOptions">
<% vehicle_record_view.ExportOptions.Render "body", "", "", "", "", "" %>
<% If Not vehicle_record_view.ExportOptions.UseDropDownButton Then %>
</div>
<div class="ewViewOtherOptions">
<% End If %>
<%
	vehicle_record_view.ActionOptions.Render "body", "", "", "", "", ""
	vehicle_record_view.DetailOptions.Render "body", "", "", "", "", ""
%>
</div>
<% End If %>
<% vehicle_record_view.ShowPageHeader() %>
<% vehicle_record_view.ShowMessage %>
<form name="fvehicle_recordview" id="fvehicle_recordview" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="vehicle_record">
<table class="ewGrid"><tr><td>
<table id="tbl_vehicle_recordview" class="table table-bordered table-striped">
<% If vehicle_record.veh_id.Visible Then ' veh_id %>
	<tr id="r_veh_id">
		<td><span id="elh_vehicle_record_veh_id"><%= vehicle_record.veh_id.FldCaption %></span></td>
		<td<%= vehicle_record.veh_id.CellAttributes %>>
<span id="el_vehicle_record_veh_id" class="control-group">
<span<%= vehicle_record.veh_id.ViewAttributes %>>
<%= vehicle_record.veh_id.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.vch_month.Visible Then ' vch_month %>
	<tr id="r_vch_month">
		<td><span id="elh_vehicle_record_vch_month"><%= vehicle_record.vch_month.FldCaption %></span></td>
		<td<%= vehicle_record.vch_month.CellAttributes %>>
<span id="el_vehicle_record_vch_month" class="control-group">
<span<%= vehicle_record.vch_month.ViewAttributes %>>
<%= vehicle_record.vch_month.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.vch_year.Visible Then ' vch_year %>
	<tr id="r_vch_year">
		<td><span id="elh_vehicle_record_vch_year"><%= vehicle_record.vch_year.FldCaption %></span></td>
		<td<%= vehicle_record.vch_year.CellAttributes %>>
<span id="el_vehicle_record_vch_year" class="control-group">
<span<%= vehicle_record.vch_year.ViewAttributes %>>
<%= vehicle_record.vch_year.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_product_1.Visible Then ' veh_product_1 %>
	<tr id="r_veh_product_1">
		<td><span id="elh_vehicle_record_veh_product_1"><%= vehicle_record.veh_product_1.FldCaption %></span></td>
		<td<%= vehicle_record.veh_product_1.CellAttributes %>>
<span id="el_vehicle_record_veh_product_1" class="control-group">
<span<%= vehicle_record.veh_product_1.ViewAttributes %>>
<%= vehicle_record.veh_product_1.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_product_2.Visible Then ' veh_product_2 %>
	<tr id="r_veh_product_2">
		<td><span id="elh_vehicle_record_veh_product_2"><%= vehicle_record.veh_product_2.FldCaption %></span></td>
		<td<%= vehicle_record.veh_product_2.CellAttributes %>>
<span id="el_vehicle_record_veh_product_2" class="control-group">
<span<%= vehicle_record.veh_product_2.ViewAttributes %>>
<%= vehicle_record.veh_product_2.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_product_3.Visible Then ' veh_product_3 %>
	<tr id="r_veh_product_3">
		<td><span id="elh_vehicle_record_veh_product_3"><%= vehicle_record.veh_product_3.FldCaption %></span></td>
		<td<%= vehicle_record.veh_product_3.CellAttributes %>>
<span id="el_vehicle_record_veh_product_3" class="control-group">
<span<%= vehicle_record.veh_product_3.ViewAttributes %>>
<%= vehicle_record.veh_product_3.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_product_4.Visible Then ' veh_product_4 %>
	<tr id="r_veh_product_4">
		<td><span id="elh_vehicle_record_veh_product_4"><%= vehicle_record.veh_product_4.FldCaption %></span></td>
		<td<%= vehicle_record.veh_product_4.CellAttributes %>>
<span id="el_vehicle_record_veh_product_4" class="control-group">
<span<%= vehicle_record.veh_product_4.ViewAttributes %>>
<%= vehicle_record.veh_product_4.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_product_5.Visible Then ' veh_product_5 %>
	<tr id="r_veh_product_5">
		<td><span id="elh_vehicle_record_veh_product_5"><%= vehicle_record.veh_product_5.FldCaption %></span></td>
		<td<%= vehicle_record.veh_product_5.CellAttributes %>>
<span id="el_vehicle_record_veh_product_5" class="control-group">
<span<%= vehicle_record.veh_product_5.ViewAttributes %>>
<%= vehicle_record.veh_product_5.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_product_6.Visible Then ' veh_product_6 %>
	<tr id="r_veh_product_6">
		<td><span id="elh_vehicle_record_veh_product_6"><%= vehicle_record.veh_product_6.FldCaption %></span></td>
		<td<%= vehicle_record.veh_product_6.CellAttributes %>>
<span id="el_vehicle_record_veh_product_6" class="control-group">
<span<%= vehicle_record.veh_product_6.ViewAttributes %>>
<%= vehicle_record.veh_product_6.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_product_7.Visible Then ' veh_product_7 %>
	<tr id="r_veh_product_7">
		<td><span id="elh_vehicle_record_veh_product_7"><%= vehicle_record.veh_product_7.FldCaption %></span></td>
		<td<%= vehicle_record.veh_product_7.CellAttributes %>>
<span id="el_vehicle_record_veh_product_7" class="control-group">
<span<%= vehicle_record.veh_product_7.ViewAttributes %>>
<%= vehicle_record.veh_product_7.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_product_8.Visible Then ' veh_product_8 %>
	<tr id="r_veh_product_8">
		<td><span id="elh_vehicle_record_veh_product_8"><%= vehicle_record.veh_product_8.FldCaption %></span></td>
		<td<%= vehicle_record.veh_product_8.CellAttributes %>>
<span id="el_vehicle_record_veh_product_8" class="control-group">
<span<%= vehicle_record.veh_product_8.ViewAttributes %>>
<%= vehicle_record.veh_product_8.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_domes_1.Visible Then ' veh_domes_1 %>
	<tr id="r_veh_domes_1">
		<td><span id="elh_vehicle_record_veh_domes_1"><%= vehicle_record.veh_domes_1.FldCaption %></span></td>
		<td<%= vehicle_record.veh_domes_1.CellAttributes %>>
<span id="el_vehicle_record_veh_domes_1" class="control-group">
<span<%= vehicle_record.veh_domes_1.ViewAttributes %>>
<%= vehicle_record.veh_domes_1.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_domes_2.Visible Then ' veh_domes_2 %>
	<tr id="r_veh_domes_2">
		<td><span id="elh_vehicle_record_veh_domes_2"><%= vehicle_record.veh_domes_2.FldCaption %></span></td>
		<td<%= vehicle_record.veh_domes_2.CellAttributes %>>
<span id="el_vehicle_record_veh_domes_2" class="control-group">
<span<%= vehicle_record.veh_domes_2.ViewAttributes %>>
<%= vehicle_record.veh_domes_2.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_domes_3.Visible Then ' veh_domes_3 %>
	<tr id="r_veh_domes_3">
		<td><span id="elh_vehicle_record_veh_domes_3"><%= vehicle_record.veh_domes_3.FldCaption %></span></td>
		<td<%= vehicle_record.veh_domes_3.CellAttributes %>>
<span id="el_vehicle_record_veh_domes_3" class="control-group">
<span<%= vehicle_record.veh_domes_3.ViewAttributes %>>
<%= vehicle_record.veh_domes_3.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_domes_4.Visible Then ' veh_domes_4 %>
	<tr id="r_veh_domes_4">
		<td><span id="elh_vehicle_record_veh_domes_4"><%= vehicle_record.veh_domes_4.FldCaption %></span></td>
		<td<%= vehicle_record.veh_domes_4.CellAttributes %>>
<span id="el_vehicle_record_veh_domes_4" class="control-group">
<span<%= vehicle_record.veh_domes_4.ViewAttributes %>>
<%= vehicle_record.veh_domes_4.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_domes_5.Visible Then ' veh_domes_5 %>
	<tr id="r_veh_domes_5">
		<td><span id="elh_vehicle_record_veh_domes_5"><%= vehicle_record.veh_domes_5.FldCaption %></span></td>
		<td<%= vehicle_record.veh_domes_5.CellAttributes %>>
<span id="el_vehicle_record_veh_domes_5" class="control-group">
<span<%= vehicle_record.veh_domes_5.ViewAttributes %>>
<%= vehicle_record.veh_domes_5.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_domes_6.Visible Then ' veh_domes_6 %>
	<tr id="r_veh_domes_6">
		<td><span id="elh_vehicle_record_veh_domes_6"><%= vehicle_record.veh_domes_6.FldCaption %></span></td>
		<td<%= vehicle_record.veh_domes_6.CellAttributes %>>
<span id="el_vehicle_record_veh_domes_6" class="control-group">
<span<%= vehicle_record.veh_domes_6.ViewAttributes %>>
<%= vehicle_record.veh_domes_6.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_domes_7.Visible Then ' veh_domes_7 %>
	<tr id="r_veh_domes_7">
		<td><span id="elh_vehicle_record_veh_domes_7"><%= vehicle_record.veh_domes_7.FldCaption %></span></td>
		<td<%= vehicle_record.veh_domes_7.CellAttributes %>>
<span id="el_vehicle_record_veh_domes_7" class="control-group">
<span<%= vehicle_record.veh_domes_7.ViewAttributes %>>
<%= vehicle_record.veh_domes_7.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_domes_8.Visible Then ' veh_domes_8 %>
	<tr id="r_veh_domes_8">
		<td><span id="elh_vehicle_record_veh_domes_8"><%= vehicle_record.veh_domes_8.FldCaption %></span></td>
		<td<%= vehicle_record.veh_domes_8.CellAttributes %>>
<span id="el_vehicle_record_veh_domes_8" class="control-group">
<span<%= vehicle_record.veh_domes_8.ViewAttributes %>>
<%= vehicle_record.veh_domes_8.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_export_1.Visible Then ' veh_export_1 %>
	<tr id="r_veh_export_1">
		<td><span id="elh_vehicle_record_veh_export_1"><%= vehicle_record.veh_export_1.FldCaption %></span></td>
		<td<%= vehicle_record.veh_export_1.CellAttributes %>>
<span id="el_vehicle_record_veh_export_1" class="control-group">
<span<%= vehicle_record.veh_export_1.ViewAttributes %>>
<%= vehicle_record.veh_export_1.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_export_2.Visible Then ' veh_export_2 %>
	<tr id="r_veh_export_2">
		<td><span id="elh_vehicle_record_veh_export_2"><%= vehicle_record.veh_export_2.FldCaption %></span></td>
		<td<%= vehicle_record.veh_export_2.CellAttributes %>>
<span id="el_vehicle_record_veh_export_2" class="control-group">
<span<%= vehicle_record.veh_export_2.ViewAttributes %>>
<%= vehicle_record.veh_export_2.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_export_3.Visible Then ' veh_export_3 %>
	<tr id="r_veh_export_3">
		<td><span id="elh_vehicle_record_veh_export_3"><%= vehicle_record.veh_export_3.FldCaption %></span></td>
		<td<%= vehicle_record.veh_export_3.CellAttributes %>>
<span id="el_vehicle_record_veh_export_3" class="control-group">
<span<%= vehicle_record.veh_export_3.ViewAttributes %>>
<%= vehicle_record.veh_export_3.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_export_4.Visible Then ' veh_export_4 %>
	<tr id="r_veh_export_4">
		<td><span id="elh_vehicle_record_veh_export_4"><%= vehicle_record.veh_export_4.FldCaption %></span></td>
		<td<%= vehicle_record.veh_export_4.CellAttributes %>>
<span id="el_vehicle_record_veh_export_4" class="control-group">
<span<%= vehicle_record.veh_export_4.ViewAttributes %>>
<%= vehicle_record.veh_export_4.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_export_5.Visible Then ' veh_export_5 %>
	<tr id="r_veh_export_5">
		<td><span id="elh_vehicle_record_veh_export_5"><%= vehicle_record.veh_export_5.FldCaption %></span></td>
		<td<%= vehicle_record.veh_export_5.CellAttributes %>>
<span id="el_vehicle_record_veh_export_5" class="control-group">
<span<%= vehicle_record.veh_export_5.ViewAttributes %>>
<%= vehicle_record.veh_export_5.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_export_6.Visible Then ' veh_export_6 %>
	<tr id="r_veh_export_6">
		<td><span id="elh_vehicle_record_veh_export_6"><%= vehicle_record.veh_export_6.FldCaption %></span></td>
		<td<%= vehicle_record.veh_export_6.CellAttributes %>>
<span id="el_vehicle_record_veh_export_6" class="control-group">
<span<%= vehicle_record.veh_export_6.ViewAttributes %>>
<%= vehicle_record.veh_export_6.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_export_7.Visible Then ' veh_export_7 %>
	<tr id="r_veh_export_7">
		<td><span id="elh_vehicle_record_veh_export_7"><%= vehicle_record.veh_export_7.FldCaption %></span></td>
		<td<%= vehicle_record.veh_export_7.CellAttributes %>>
<span id="el_vehicle_record_veh_export_7" class="control-group">
<span<%= vehicle_record.veh_export_7.ViewAttributes %>>
<%= vehicle_record.veh_export_7.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_export_8.Visible Then ' veh_export_8 %>
	<tr id="r_veh_export_8">
		<td><span id="elh_vehicle_record_veh_export_8"><%= vehicle_record.veh_export_8.FldCaption %></span></td>
		<td<%= vehicle_record.veh_export_8.CellAttributes %>>
<span id="el_vehicle_record_veh_export_8" class="control-group">
<span<%= vehicle_record.veh_export_8.ViewAttributes %>>
<%= vehicle_record.veh_export_8.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_remark.Visible Then ' veh_remark %>
	<tr id="r_veh_remark">
		<td><span id="elh_vehicle_record_veh_remark"><%= vehicle_record.veh_remark.FldCaption %></span></td>
		<td<%= vehicle_record.veh_remark.CellAttributes %>>
<span id="el_vehicle_record_veh_remark" class="control-group">
<span<%= vehicle_record.veh_remark.ViewAttributes %>>
<%= vehicle_record.veh_remark.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_month_title.Visible Then ' veh_month_title %>
	<tr id="r_veh_month_title">
		<td><span id="elh_vehicle_record_veh_month_title"><%= vehicle_record.veh_month_title.FldCaption %></span></td>
		<td<%= vehicle_record.veh_month_title.CellAttributes %>>
<span id="el_vehicle_record_veh_month_title" class="control-group">
<span<%= vehicle_record.veh_month_title.ViewAttributes %>>
<%= vehicle_record.veh_month_title.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_range.Visible Then ' veh_range %>
	<tr id="r_veh_range">
		<td><span id="elh_vehicle_record_veh_range"><%= vehicle_record.veh_range.FldCaption %></span></td>
		<td<%= vehicle_record.veh_range.CellAttributes %>>
<span id="el_vehicle_record_veh_range" class="control-group">
<span<%= vehicle_record.veh_range.ViewAttributes %>>
<%= vehicle_record.veh_range.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_month_title2.Visible Then ' veh_month_title2 %>
	<tr id="r_veh_month_title2">
		<td><span id="elh_vehicle_record_veh_month_title2"><%= vehicle_record.veh_month_title2.FldCaption %></span></td>
		<td<%= vehicle_record.veh_month_title2.CellAttributes %>>
<span id="el_vehicle_record_veh_month_title2" class="control-group">
<span<%= vehicle_record.veh_month_title2.ViewAttributes %>>
<%= vehicle_record.veh_month_title2.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
<% If vehicle_record.veh_range2.Visible Then ' veh_range2 %>
	<tr id="r_veh_range2">
		<td><span id="elh_vehicle_record_veh_range2"><%= vehicle_record.veh_range2.FldCaption %></span></td>
		<td<%= vehicle_record.veh_range2.CellAttributes %>>
<span id="el_vehicle_record_veh_range2" class="control-group">
<span<%= vehicle_record.veh_range2.ViewAttributes %>>
<%= vehicle_record.veh_range2.ViewValue %>
</span>
</span>
</td>
	</tr>
<% End If %>
</table>
</td></tr></table>
</form>
<script type="text/javascript">
fvehicle_recordview.Init();
</script>
<%
vehicle_record_view.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If vehicle_record.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set vehicle_record_view = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cvehicle_record_view

	' Page ID
	Public Property Get PageID()
		PageID = "view"
	End Property

	' Project ID
	Public Property Get ProjectID()
		ProjectID = "{324ED72D-DE20-46F7-B12E-7AF8CE8711A6}"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "vehicle_record"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "vehicle_record_view"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If vehicle_record.UseTokenInUrl Then PageUrl = PageUrl & "t=" & vehicle_record.TableVar & "&" ' add page token
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
		If vehicle_record.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (vehicle_record.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (vehicle_record.TableVar = Request.QueryString("t"))
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
		If IsEmpty(vehicle_record) Then Set vehicle_record = New cvehicle_record
		Set Table = vehicle_record

		' Initialize urls
		Dim KeyUrl
		KeyUrl = ""
		If Request.QueryString("veh_id").Count > 0 Then
			ew_AddKey RecKey, "veh_id", Request.QueryString("veh_id")
			KeyUrl = KeyUrl & "&amp;veh_id=" & Server.URLEncode(Request.QueryString("veh_id"))
		End If
		ExportPrintUrl = PageUrl & "export=print" & KeyUrl
		ExportHtmlUrl = PageUrl & "export=html" & KeyUrl
		ExportExcelUrl = PageUrl & "export=excel" & KeyUrl
		ExportWordUrl = PageUrl & "export=word" & KeyUrl
		ExportXmlUrl = PageUrl & "export=xml" & KeyUrl
		ExportCsvUrl = PageUrl & "export=csv" & KeyUrl
		ExportPdfUrl = PageUrl & "export=pdf" & KeyUrl

		' Initialize other table object
		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "view"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "vehicle_record"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = vehicle_record.TableVar
		ExportOptions.Tag = "div"
		ExportOptions.TagClassName = "ewExportOption"

		' Other options
		Set ActionOptions = New cListOptions
		ActionOptions.Tag = "div"
		ActionOptions.TagClassName = "ewActionOption"
		Set DetailOptions = New cListOptions
		DetailOptions.Tag = "div"
		DetailOptions.TagClassName = "ewDetailOption"
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
		Set vehicle_record = Nothing
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

	Dim DisplayRecs ' Number of display records
	Dim StartRec, StopRec, TotalRecs, RecRange
	Dim RecCnt
	Dim RecKey
	Dim ExportOptions ' Export options
	Dim DetailOptions ' Other options (detail)
	Dim ActionOptions ' Other options (action)
	Dim Recordset

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Paging variables
		DisplayRecs = 1
		RecRange = 10

		' Load current record
		Dim bLoadCurrentRecord
		bLoadCurrentRecord = False
		Dim sReturnUrl
		sReturnUrl = ""
		Dim bMatchRecord
		bMatchRecord = False

		' Set up Breadcrumb
		If vehicle_record.Export = "" Then
			SetupBreadcrumb()
		End If
		If IsPageRequest Then ' Validate request
			If Request.QueryString("veh_id").Count > 0 Then
				vehicle_record.veh_id.QueryStringValue = Request.QueryString("veh_id")
			Else
				sReturnUrl = "pom_vehicle_recordlist.asp" ' Return to list
			End If

			' Get action
			vehicle_record.CurrentAction = "I" ' Display form
			Select Case vehicle_record.CurrentAction
				Case "I" ' Get a record to display
					If Not LoadRow() Then ' Load record based on key
						If SuccessMessage = "" And FailureMessage = "" Then
							FailureMessage = Language.Phrase("NoRecord") ' Set no record message
						End If
						sReturnUrl = "pom_vehicle_recordlist.asp" ' No matching record, return to list
					End If
			End Select
		Else
			sReturnUrl = "pom_vehicle_recordlist.asp" ' Not page request, return to list
		End If
		If sReturnUrl <> "" Then Call Page_Terminate(sReturnUrl)

		' Render row
		vehicle_record.RowType = EW_ROWTYPE_VIEW
		Call vehicle_record.ResetAttrs()
		Call RenderRow()
	End Sub

	' Set up other options
	Sub SetupOtherOptions()
		Dim opt, item
		Set opt = ActionOptions

		' Add
		Call opt.Add("add")
		Set item = opt.GetItem("add")
		item.Body = "<a class=""ewAction ewAdd"" href=""" & ew_HtmlEncode(AddUrl) & """>" & Language.Phrase("ViewPageAddLink") & "</a>"
		item.Visible = (AddUrl <> "" And Security.IsLoggedIn())

		' Edit
		Call opt.Add("edit")
		Set item = opt.GetItem("edit")
		item.Body = "<a class=""ewAction ewEdit"" href=""" & ew_HtmlEncode(EditUrl) & """>" & Language.Phrase("ViewPageEditLink") & "</a>"
		item.Visible = (EditUrl <> "" And Security.IsLoggedIn())

		' Copy
		Call opt.Add("copy")
		Set item = opt.GetItem("copy")
		item.Body = "<a class=""ewAction ewCopy"" href=""" & ew_HtmlEncode(CopyUrl) & """>" & Language.Phrase("ViewPageCopyLink") & "</a>"
		item.Visible = (CopyUrl <> "" And Security.IsLoggedIn())

		' Delete
		Call opt.Add("delete")
		Set item = opt.GetItem("delete")
		item.Body = "<a class=""ewAction ewDelete"" href=""" & ew_HtmlEncode(DeleteUrl) & """>" & Language.Phrase("ViewPageDeleteLink") & "</a>"
		item.Visible = (DeleteUrl <> "" And Security.IsLoggedIn())

		' Set up options default
		Set opt = ActionOptions
		opt.DropDownButtonPhrase = Language.Phrase("ButtonActions")
		opt.UseDropDownButton = False
		opt.UseButtonGroup = True
		Call opt.Add(opt.GroupOptionName)
		Set item = opt.GetItem(opt.GroupOptionName)
		item.Body = ""
		item.Visible = False
		Set opt = DetailOptions
		opt.DropDownButtonPhrase = Language.Phrase("ButtonDetails")
		opt.UseDropDownButton = False
		opt.UseButtonGroup = True
		Call opt.Add(opt.GroupOptionName)
		Set item = opt.GetItem(opt.GroupOptionName)
		item.Body = ""
		item.Visible = False
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
				vehicle_record.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					vehicle_record.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = vehicle_record.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			vehicle_record.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			vehicle_record.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			vehicle_record.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = vehicle_record.KeyFilter

		' Call Row Selecting event
		Call vehicle_record.Row_Selecting(sFilter)

		' Load sql based on filter
		vehicle_record.CurrentFilter = sFilter
		sSql = vehicle_record.SQL
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
		Call vehicle_record.Row_Selected(RsRow)
		vehicle_record.veh_id.DbValue = RsRow("veh_id")
		vehicle_record.vch_month.DbValue = RsRow("vch_month")
		vehicle_record.vch_year.DbValue = RsRow("vch_year")
		vehicle_record.veh_product_1.DbValue = RsRow("veh_product_1")
		vehicle_record.veh_product_2.DbValue = RsRow("veh_product_2")
		vehicle_record.veh_product_3.DbValue = RsRow("veh_product_3")
		vehicle_record.veh_product_4.DbValue = RsRow("veh_product_4")
		vehicle_record.veh_product_5.DbValue = RsRow("veh_product_5")
		vehicle_record.veh_product_6.DbValue = RsRow("veh_product_6")
		vehicle_record.veh_product_7.DbValue = RsRow("veh_product_7")
		vehicle_record.veh_product_8.DbValue = RsRow("veh_product_8")
		vehicle_record.veh_domes_1.DbValue = RsRow("veh_domes_1")
		vehicle_record.veh_domes_2.DbValue = RsRow("veh_domes_2")
		vehicle_record.veh_domes_3.DbValue = RsRow("veh_domes_3")
		vehicle_record.veh_domes_4.DbValue = RsRow("veh_domes_4")
		vehicle_record.veh_domes_5.DbValue = RsRow("veh_domes_5")
		vehicle_record.veh_domes_6.DbValue = RsRow("veh_domes_6")
		vehicle_record.veh_domes_7.DbValue = RsRow("veh_domes_7")
		vehicle_record.veh_domes_8.DbValue = RsRow("veh_domes_8")
		vehicle_record.veh_export_1.DbValue = RsRow("veh_export_1")
		vehicle_record.veh_export_2.DbValue = RsRow("veh_export_2")
		vehicle_record.veh_export_3.DbValue = RsRow("veh_export_3")
		vehicle_record.veh_export_4.DbValue = RsRow("veh_export_4")
		vehicle_record.veh_export_5.DbValue = RsRow("veh_export_5")
		vehicle_record.veh_export_6.DbValue = RsRow("veh_export_6")
		vehicle_record.veh_export_7.DbValue = RsRow("veh_export_7")
		vehicle_record.veh_export_8.DbValue = RsRow("veh_export_8")
		vehicle_record.veh_remark.DbValue = RsRow("veh_remark")
		vehicle_record.veh_month_title.DbValue = RsRow("veh_month_title")
		vehicle_record.veh_range.DbValue = RsRow("veh_range")
		vehicle_record.veh_month_title2.DbValue = RsRow("veh_month_title2")
		vehicle_record.veh_range2.DbValue = RsRow("veh_range2")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		vehicle_record.veh_id.m_DbValue = Rs("veh_id")
		vehicle_record.vch_month.m_DbValue = Rs("vch_month")
		vehicle_record.vch_year.m_DbValue = Rs("vch_year")
		vehicle_record.veh_product_1.m_DbValue = Rs("veh_product_1")
		vehicle_record.veh_product_2.m_DbValue = Rs("veh_product_2")
		vehicle_record.veh_product_3.m_DbValue = Rs("veh_product_3")
		vehicle_record.veh_product_4.m_DbValue = Rs("veh_product_4")
		vehicle_record.veh_product_5.m_DbValue = Rs("veh_product_5")
		vehicle_record.veh_product_6.m_DbValue = Rs("veh_product_6")
		vehicle_record.veh_product_7.m_DbValue = Rs("veh_product_7")
		vehicle_record.veh_product_8.m_DbValue = Rs("veh_product_8")
		vehicle_record.veh_domes_1.m_DbValue = Rs("veh_domes_1")
		vehicle_record.veh_domes_2.m_DbValue = Rs("veh_domes_2")
		vehicle_record.veh_domes_3.m_DbValue = Rs("veh_domes_3")
		vehicle_record.veh_domes_4.m_DbValue = Rs("veh_domes_4")
		vehicle_record.veh_domes_5.m_DbValue = Rs("veh_domes_5")
		vehicle_record.veh_domes_6.m_DbValue = Rs("veh_domes_6")
		vehicle_record.veh_domes_7.m_DbValue = Rs("veh_domes_7")
		vehicle_record.veh_domes_8.m_DbValue = Rs("veh_domes_8")
		vehicle_record.veh_export_1.m_DbValue = Rs("veh_export_1")
		vehicle_record.veh_export_2.m_DbValue = Rs("veh_export_2")
		vehicle_record.veh_export_3.m_DbValue = Rs("veh_export_3")
		vehicle_record.veh_export_4.m_DbValue = Rs("veh_export_4")
		vehicle_record.veh_export_5.m_DbValue = Rs("veh_export_5")
		vehicle_record.veh_export_6.m_DbValue = Rs("veh_export_6")
		vehicle_record.veh_export_7.m_DbValue = Rs("veh_export_7")
		vehicle_record.veh_export_8.m_DbValue = Rs("veh_export_8")
		vehicle_record.veh_remark.m_DbValue = Rs("veh_remark")
		vehicle_record.veh_month_title.m_DbValue = Rs("veh_month_title")
		vehicle_record.veh_range.m_DbValue = Rs("veh_range")
		vehicle_record.veh_month_title2.m_DbValue = Rs("veh_month_title2")
		vehicle_record.veh_range2.m_DbValue = Rs("veh_range2")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		AddUrl = vehicle_record.AddUrl
		EditUrl = vehicle_record.EditUrl("")
		CopyUrl = vehicle_record.CopyUrl("")
		DeleteUrl = vehicle_record.DeleteUrl
		ListUrl = vehicle_record.ListUrl
		SetupOtherOptions()

		' Call Row Rendering event
		Call vehicle_record.Row_Rendering()

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

		If vehicle_record.RowType = EW_ROWTYPE_VIEW Then ' View row

			' veh_id
			vehicle_record.veh_id.ViewValue = vehicle_record.veh_id.CurrentValue
			vehicle_record.veh_id.ViewCustomAttributes = ""

			' vch_month
			vehicle_record.vch_month.ViewValue = vehicle_record.vch_month.CurrentValue
			vehicle_record.vch_month.ViewCustomAttributes = ""

			' vch_year
			vehicle_record.vch_year.ViewValue = vehicle_record.vch_year.CurrentValue
			vehicle_record.vch_year.ViewCustomAttributes = ""

			' veh_product_1
			vehicle_record.veh_product_1.ViewValue = vehicle_record.veh_product_1.CurrentValue
			vehicle_record.veh_product_1.ViewCustomAttributes = ""

			' veh_product_2
			vehicle_record.veh_product_2.ViewValue = vehicle_record.veh_product_2.CurrentValue
			vehicle_record.veh_product_2.ViewCustomAttributes = ""

			' veh_product_3
			vehicle_record.veh_product_3.ViewValue = vehicle_record.veh_product_3.CurrentValue
			vehicle_record.veh_product_3.ViewCustomAttributes = ""

			' veh_product_4
			vehicle_record.veh_product_4.ViewValue = vehicle_record.veh_product_4.CurrentValue
			vehicle_record.veh_product_4.ViewCustomAttributes = ""

			' veh_product_5
			vehicle_record.veh_product_5.ViewValue = vehicle_record.veh_product_5.CurrentValue
			vehicle_record.veh_product_5.ViewCustomAttributes = ""

			' veh_product_6
			vehicle_record.veh_product_6.ViewValue = vehicle_record.veh_product_6.CurrentValue
			vehicle_record.veh_product_6.ViewCustomAttributes = ""

			' veh_product_7
			vehicle_record.veh_product_7.ViewValue = vehicle_record.veh_product_7.CurrentValue
			vehicle_record.veh_product_7.ViewCustomAttributes = ""

			' veh_product_8
			vehicle_record.veh_product_8.ViewValue = vehicle_record.veh_product_8.CurrentValue
			vehicle_record.veh_product_8.ViewCustomAttributes = ""

			' veh_domes_1
			vehicle_record.veh_domes_1.ViewValue = vehicle_record.veh_domes_1.CurrentValue
			vehicle_record.veh_domes_1.ViewCustomAttributes = ""

			' veh_domes_2
			vehicle_record.veh_domes_2.ViewValue = vehicle_record.veh_domes_2.CurrentValue
			vehicle_record.veh_domes_2.ViewCustomAttributes = ""

			' veh_domes_3
			vehicle_record.veh_domes_3.ViewValue = vehicle_record.veh_domes_3.CurrentValue
			vehicle_record.veh_domes_3.ViewCustomAttributes = ""

			' veh_domes_4
			vehicle_record.veh_domes_4.ViewValue = vehicle_record.veh_domes_4.CurrentValue
			vehicle_record.veh_domes_4.ViewCustomAttributes = ""

			' veh_domes_5
			vehicle_record.veh_domes_5.ViewValue = vehicle_record.veh_domes_5.CurrentValue
			vehicle_record.veh_domes_5.ViewCustomAttributes = ""

			' veh_domes_6
			vehicle_record.veh_domes_6.ViewValue = vehicle_record.veh_domes_6.CurrentValue
			vehicle_record.veh_domes_6.ViewCustomAttributes = ""

			' veh_domes_7
			vehicle_record.veh_domes_7.ViewValue = vehicle_record.veh_domes_7.CurrentValue
			vehicle_record.veh_domes_7.ViewCustomAttributes = ""

			' veh_domes_8
			vehicle_record.veh_domes_8.ViewValue = vehicle_record.veh_domes_8.CurrentValue
			vehicle_record.veh_domes_8.ViewCustomAttributes = ""

			' veh_export_1
			vehicle_record.veh_export_1.ViewValue = vehicle_record.veh_export_1.CurrentValue
			vehicle_record.veh_export_1.ViewCustomAttributes = ""

			' veh_export_2
			vehicle_record.veh_export_2.ViewValue = vehicle_record.veh_export_2.CurrentValue
			vehicle_record.veh_export_2.ViewCustomAttributes = ""

			' veh_export_3
			vehicle_record.veh_export_3.ViewValue = vehicle_record.veh_export_3.CurrentValue
			vehicle_record.veh_export_3.ViewCustomAttributes = ""

			' veh_export_4
			vehicle_record.veh_export_4.ViewValue = vehicle_record.veh_export_4.CurrentValue
			vehicle_record.veh_export_4.ViewCustomAttributes = ""

			' veh_export_5
			vehicle_record.veh_export_5.ViewValue = vehicle_record.veh_export_5.CurrentValue
			vehicle_record.veh_export_5.ViewCustomAttributes = ""

			' veh_export_6
			vehicle_record.veh_export_6.ViewValue = vehicle_record.veh_export_6.CurrentValue
			vehicle_record.veh_export_6.ViewCustomAttributes = ""

			' veh_export_7
			vehicle_record.veh_export_7.ViewValue = vehicle_record.veh_export_7.CurrentValue
			vehicle_record.veh_export_7.ViewCustomAttributes = ""

			' veh_export_8
			vehicle_record.veh_export_8.ViewValue = vehicle_record.veh_export_8.CurrentValue
			vehicle_record.veh_export_8.ViewCustomAttributes = ""

			' veh_remark
			vehicle_record.veh_remark.ViewValue = vehicle_record.veh_remark.CurrentValue
			vehicle_record.veh_remark.ViewCustomAttributes = ""

			' veh_month_title
			vehicle_record.veh_month_title.ViewValue = vehicle_record.veh_month_title.CurrentValue
			vehicle_record.veh_month_title.ViewCustomAttributes = ""

			' veh_range
			vehicle_record.veh_range.ViewValue = vehicle_record.veh_range.CurrentValue
			vehicle_record.veh_range.ViewCustomAttributes = ""

			' veh_month_title2
			vehicle_record.veh_month_title2.ViewValue = vehicle_record.veh_month_title2.CurrentValue
			vehicle_record.veh_month_title2.ViewCustomAttributes = ""

			' veh_range2
			vehicle_record.veh_range2.ViewValue = vehicle_record.veh_range2.CurrentValue
			vehicle_record.veh_range2.ViewCustomAttributes = ""

			' View refer script
			' veh_id

			vehicle_record.veh_id.LinkCustomAttributes = ""
			vehicle_record.veh_id.HrefValue = ""
			vehicle_record.veh_id.TooltipValue = ""

			' vch_month
			vehicle_record.vch_month.LinkCustomAttributes = ""
			vehicle_record.vch_month.HrefValue = ""
			vehicle_record.vch_month.TooltipValue = ""

			' vch_year
			vehicle_record.vch_year.LinkCustomAttributes = ""
			vehicle_record.vch_year.HrefValue = ""
			vehicle_record.vch_year.TooltipValue = ""

			' veh_product_1
			vehicle_record.veh_product_1.LinkCustomAttributes = ""
			vehicle_record.veh_product_1.HrefValue = ""
			vehicle_record.veh_product_1.TooltipValue = ""

			' veh_product_2
			vehicle_record.veh_product_2.LinkCustomAttributes = ""
			vehicle_record.veh_product_2.HrefValue = ""
			vehicle_record.veh_product_2.TooltipValue = ""

			' veh_product_3
			vehicle_record.veh_product_3.LinkCustomAttributes = ""
			vehicle_record.veh_product_3.HrefValue = ""
			vehicle_record.veh_product_3.TooltipValue = ""

			' veh_product_4
			vehicle_record.veh_product_4.LinkCustomAttributes = ""
			vehicle_record.veh_product_4.HrefValue = ""
			vehicle_record.veh_product_4.TooltipValue = ""

			' veh_product_5
			vehicle_record.veh_product_5.LinkCustomAttributes = ""
			vehicle_record.veh_product_5.HrefValue = ""
			vehicle_record.veh_product_5.TooltipValue = ""

			' veh_product_6
			vehicle_record.veh_product_6.LinkCustomAttributes = ""
			vehicle_record.veh_product_6.HrefValue = ""
			vehicle_record.veh_product_6.TooltipValue = ""

			' veh_product_7
			vehicle_record.veh_product_7.LinkCustomAttributes = ""
			vehicle_record.veh_product_7.HrefValue = ""
			vehicle_record.veh_product_7.TooltipValue = ""

			' veh_product_8
			vehicle_record.veh_product_8.LinkCustomAttributes = ""
			vehicle_record.veh_product_8.HrefValue = ""
			vehicle_record.veh_product_8.TooltipValue = ""

			' veh_domes_1
			vehicle_record.veh_domes_1.LinkCustomAttributes = ""
			vehicle_record.veh_domes_1.HrefValue = ""
			vehicle_record.veh_domes_1.TooltipValue = ""

			' veh_domes_2
			vehicle_record.veh_domes_2.LinkCustomAttributes = ""
			vehicle_record.veh_domes_2.HrefValue = ""
			vehicle_record.veh_domes_2.TooltipValue = ""

			' veh_domes_3
			vehicle_record.veh_domes_3.LinkCustomAttributes = ""
			vehicle_record.veh_domes_3.HrefValue = ""
			vehicle_record.veh_domes_3.TooltipValue = ""

			' veh_domes_4
			vehicle_record.veh_domes_4.LinkCustomAttributes = ""
			vehicle_record.veh_domes_4.HrefValue = ""
			vehicle_record.veh_domes_4.TooltipValue = ""

			' veh_domes_5
			vehicle_record.veh_domes_5.LinkCustomAttributes = ""
			vehicle_record.veh_domes_5.HrefValue = ""
			vehicle_record.veh_domes_5.TooltipValue = ""

			' veh_domes_6
			vehicle_record.veh_domes_6.LinkCustomAttributes = ""
			vehicle_record.veh_domes_6.HrefValue = ""
			vehicle_record.veh_domes_6.TooltipValue = ""

			' veh_domes_7
			vehicle_record.veh_domes_7.LinkCustomAttributes = ""
			vehicle_record.veh_domes_7.HrefValue = ""
			vehicle_record.veh_domes_7.TooltipValue = ""

			' veh_domes_8
			vehicle_record.veh_domes_8.LinkCustomAttributes = ""
			vehicle_record.veh_domes_8.HrefValue = ""
			vehicle_record.veh_domes_8.TooltipValue = ""

			' veh_export_1
			vehicle_record.veh_export_1.LinkCustomAttributes = ""
			vehicle_record.veh_export_1.HrefValue = ""
			vehicle_record.veh_export_1.TooltipValue = ""

			' veh_export_2
			vehicle_record.veh_export_2.LinkCustomAttributes = ""
			vehicle_record.veh_export_2.HrefValue = ""
			vehicle_record.veh_export_2.TooltipValue = ""

			' veh_export_3
			vehicle_record.veh_export_3.LinkCustomAttributes = ""
			vehicle_record.veh_export_3.HrefValue = ""
			vehicle_record.veh_export_3.TooltipValue = ""

			' veh_export_4
			vehicle_record.veh_export_4.LinkCustomAttributes = ""
			vehicle_record.veh_export_4.HrefValue = ""
			vehicle_record.veh_export_4.TooltipValue = ""

			' veh_export_5
			vehicle_record.veh_export_5.LinkCustomAttributes = ""
			vehicle_record.veh_export_5.HrefValue = ""
			vehicle_record.veh_export_5.TooltipValue = ""

			' veh_export_6
			vehicle_record.veh_export_6.LinkCustomAttributes = ""
			vehicle_record.veh_export_6.HrefValue = ""
			vehicle_record.veh_export_6.TooltipValue = ""

			' veh_export_7
			vehicle_record.veh_export_7.LinkCustomAttributes = ""
			vehicle_record.veh_export_7.HrefValue = ""
			vehicle_record.veh_export_7.TooltipValue = ""

			' veh_export_8
			vehicle_record.veh_export_8.LinkCustomAttributes = ""
			vehicle_record.veh_export_8.HrefValue = ""
			vehicle_record.veh_export_8.TooltipValue = ""

			' veh_remark
			vehicle_record.veh_remark.LinkCustomAttributes = ""
			vehicle_record.veh_remark.HrefValue = ""
			vehicle_record.veh_remark.TooltipValue = ""

			' veh_month_title
			vehicle_record.veh_month_title.LinkCustomAttributes = ""
			vehicle_record.veh_month_title.HrefValue = ""
			vehicle_record.veh_month_title.TooltipValue = ""

			' veh_range
			vehicle_record.veh_range.LinkCustomAttributes = ""
			vehicle_record.veh_range.HrefValue = ""
			vehicle_record.veh_range.TooltipValue = ""

			' veh_month_title2
			vehicle_record.veh_month_title2.LinkCustomAttributes = ""
			vehicle_record.veh_month_title2.HrefValue = ""
			vehicle_record.veh_month_title2.TooltipValue = ""

			' veh_range2
			vehicle_record.veh_range2.LinkCustomAttributes = ""
			vehicle_record.veh_range2.HrefValue = ""
			vehicle_record.veh_range2.TooltipValue = ""
		End If

		' Call Row Rendered event
		If vehicle_record.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call vehicle_record.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", vehicle_record.TableVar, "pom_vehicle_recordlist.asp", vehicle_record.TableVar, True)
		PageId = "view"
		Call Breadcrumb.Add("view", PageId, ew_CurrentUrl, "", False)
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
End Class
%>
