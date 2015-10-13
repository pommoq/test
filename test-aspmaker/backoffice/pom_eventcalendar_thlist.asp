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
Dim eventcalendar_th_list
Set eventcalendar_th_list = New ceventcalendar_th_list
Set Page = eventcalendar_th_list

' Page init processing
eventcalendar_th_list.Page_Init()

' Page main processing
eventcalendar_th_list.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
eventcalendar_th_list.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If eventcalendar_th.Export = "" Then %>
<script type="text/javascript">
// Page object
var eventcalendar_th_list = new ew_Page("eventcalendar_th_list");
eventcalendar_th_list.PageID = "list"; // Page ID
var EW_PAGE_ID = eventcalendar_th_list.PageID; // For backward compatibility
// Form object
var feventcalendar_thlist = new ew_Form("feventcalendar_thlist");
feventcalendar_thlist.FormKeyCountName = '<%= eventcalendar_th_list.FormKeyCountName %>';
// Form_CustomValidate event
feventcalendar_thlist.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
feventcalendar_thlist.ValidateRequired = true; // Use JavaScript validation
<% Else %>
feventcalendar_thlist.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
var feventcalendar_thlistsrch = new ew_Form("feventcalendar_thlistsrch");
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If eventcalendar_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If eventcalendar_th_list.ExportOptions.Visible Then %>
<div class="ewListExportOptions"><% eventcalendar_th_list.ExportOptions.Render "body", "", "", "", "", "" %></div>
<% End If %>
<% If (eventcalendar_th.Export = "") Or (EW_EXPORT_MASTER_RECORD And eventcalendar_th.Export = "print") Then %>
<% End If %>
<%

' Load recordset
Set eventcalendar_th_list.Recordset = eventcalendar_th_list.LoadRecordset()
	eventcalendar_th_list.TotalRecs = eventcalendar_th_list.Recordset.RecordCount
	eventcalendar_th_list.StartRec = 1
	If eventcalendar_th_list.DisplayRecs <= 0 Then ' Display all records
		eventcalendar_th_list.DisplayRecs = eventcalendar_th_list.TotalRecs
	End If
	If Not (eventcalendar_th.ExportAll And eventcalendar_th.Export <> "") Then
		eventcalendar_th_list.SetUpStartRec() ' Set up start record position
	End If
eventcalendar_th_list.RenderOtherOptions()
%>
<% If Security.IsLoggedIn() Then %>
<% If eventcalendar_th.Export = "" And eventcalendar_th.CurrentAction = "" Then %>
<form name="feventcalendar_thlistsrch" id="feventcalendar_thlistsrch" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewSearchTable"><tr><td>
<div class="accordion" id="feventcalendar_thlistsrch_SearchGroup">
	<div class="accordion-group">
		<div class="accordion-heading">
<a class="accordion-toggle" data-toggle="collapse" data-parent="#feventcalendar_thlistsrch_SearchGroup" href="#feventcalendar_thlistsrch_SearchBody"><%= Language.Phrase("Search") %></a>
		</div>
		<div id="feventcalendar_thlistsrch_SearchBody" class="accordion-body collapse in">
			<div class="accordion-inner">
<div id="feventcalendar_thlistsrch_SearchPanel">
<input type="hidden" name="cmd" value="search">
<input type="hidden" name="t" value="eventcalendar_th">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewRow">
	<div class="btn-group ewButtonGroup">
	<div class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" class="input-large" value="<%= ew_HtmlEncode(eventcalendar_th.BasicSearch.getKeyword()) %>" placeholder="<%= ew_HtmlEncode(Language.Phrase("Search")) %>">
	<button class="btn btn-primary ewButton" name="btnsubmit" id="btnsubmit" type="submit"><%= Language.Phrase("QuickSearchBtn") %></button>
	</div>
	</div>
	<div class="btn-group ewButtonGroup">
	<a class="btn ewShowAll" href="<%= eventcalendar_th_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>
	</div>
</div>
<div id="xsr_2" class="ewRow">
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="="<% If eventcalendar_th.BasicSearch.getSearchType = "=" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If eventcalendar_th.BasicSearch.getSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If eventcalendar_th.BasicSearch.getSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
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
<% eventcalendar_th_list.ShowPageHeader() %>
<% eventcalendar_th_list.ShowMessage %>
<table class="ewGrid"><tr><td class="ewGridContent">
<form name="feventcalendar_thlist" id="feventcalendar_thlist" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="eventcalendar_th">
<div id="gmp_eventcalendar_th" class="ewGridMiddlePanel">
<% If eventcalendar_th_list.TotalRecs > 0 Then %>
<table id="tbl_eventcalendar_thlist" class="ewTable ewTableSeparate">
<%= eventcalendar_th.TableCustomInnerHtml %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call eventcalendar_th_list.RenderListOptions()

' Render list options (header, left)
eventcalendar_th_list.ListOptions.Render "header", "left", "", "", "", ""
%>
<% If eventcalendar_th.eventcalendar_id.Visible Then ' eventcalendar_id %>
	<% If eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_id) = "" Then %>
		<td><div id="elh_eventcalendar_th_eventcalendar_id" class="eventcalendar_th_eventcalendar_id"><div class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_id.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_id) %>',1);"><div id="elh_eventcalendar_th_eventcalendar_id" class="eventcalendar_th_eventcalendar_id">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_id.FldCaption %></span><span class="ewTableHeaderSort"><% If eventcalendar_th.eventcalendar_id.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf eventcalendar_th.eventcalendar_id.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If eventcalendar_th.eventcalendar_img.Visible Then ' eventcalendar_img %>
	<% If eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_img) = "" Then %>
		<td><div id="elh_eventcalendar_th_eventcalendar_img" class="eventcalendar_th_eventcalendar_img"><div class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_img.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_img) %>',1);"><div id="elh_eventcalendar_th_eventcalendar_img" class="eventcalendar_th_eventcalendar_img">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_img.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If eventcalendar_th.eventcalendar_img.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf eventcalendar_th.eventcalendar_img.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If eventcalendar_th.eventcalendar_date.Visible Then ' eventcalendar_date %>
	<% If eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_date) = "" Then %>
		<td><div id="elh_eventcalendar_th_eventcalendar_date" class="eventcalendar_th_eventcalendar_date"><div class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_date.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_date) %>',1);"><div id="elh_eventcalendar_th_eventcalendar_date" class="eventcalendar_th_eventcalendar_date">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_date.FldCaption %></span><span class="ewTableHeaderSort"><% If eventcalendar_th.eventcalendar_date.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf eventcalendar_th.eventcalendar_date.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If eventcalendar_th.eventcalendar_category.Visible Then ' eventcalendar_category %>
	<% If eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_category) = "" Then %>
		<td><div id="elh_eventcalendar_th_eventcalendar_category" class="eventcalendar_th_eventcalendar_category"><div class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_category.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_category) %>',1);"><div id="elh_eventcalendar_th_eventcalendar_category" class="eventcalendar_th_eventcalendar_category">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_category.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If eventcalendar_th.eventcalendar_category.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf eventcalendar_th.eventcalendar_category.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If eventcalendar_th.eventcalendar_category_sub.Visible Then ' eventcalendar_category_sub %>
	<% If eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_category_sub) = "" Then %>
		<td><div id="elh_eventcalendar_th_eventcalendar_category_sub" class="eventcalendar_th_eventcalendar_category_sub"><div class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_category_sub.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_category_sub) %>',1);"><div id="elh_eventcalendar_th_eventcalendar_category_sub" class="eventcalendar_th_eventcalendar_category_sub">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_category_sub.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If eventcalendar_th.eventcalendar_category_sub.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf eventcalendar_th.eventcalendar_category_sub.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If eventcalendar_th.start_date.Visible Then ' start_date %>
	<% If eventcalendar_th.SortUrl(eventcalendar_th.start_date) = "" Then %>
		<td><div id="elh_eventcalendar_th_start_date" class="eventcalendar_th_start_date"><div class="ewTableHeaderCaption"><%= eventcalendar_th.start_date.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= eventcalendar_th.SortUrl(eventcalendar_th.start_date) %>',1);"><div id="elh_eventcalendar_th_start_date" class="eventcalendar_th_start_date">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= eventcalendar_th.start_date.FldCaption %></span><span class="ewTableHeaderSort"><% If eventcalendar_th.start_date.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf eventcalendar_th.start_date.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If eventcalendar_th.end_date.Visible Then ' end_date %>
	<% If eventcalendar_th.SortUrl(eventcalendar_th.end_date) = "" Then %>
		<td><div id="elh_eventcalendar_th_end_date" class="eventcalendar_th_end_date"><div class="ewTableHeaderCaption"><%= eventcalendar_th.end_date.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= eventcalendar_th.SortUrl(eventcalendar_th.end_date) %>',1);"><div id="elh_eventcalendar_th_end_date" class="eventcalendar_th_end_date">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= eventcalendar_th.end_date.FldCaption %></span><span class="ewTableHeaderSort"><% If eventcalendar_th.end_date.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf eventcalendar_th.end_date.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If eventcalendar_th.eventcalendar_pdf.Visible Then ' eventcalendar_pdf %>
	<% If eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_pdf) = "" Then %>
		<td><div id="elh_eventcalendar_th_eventcalendar_pdf" class="eventcalendar_th_eventcalendar_pdf"><div class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_pdf.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_pdf) %>',1);"><div id="elh_eventcalendar_th_eventcalendar_pdf" class="eventcalendar_th_eventcalendar_pdf">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_pdf.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If eventcalendar_th.eventcalendar_pdf.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf eventcalendar_th.eventcalendar_pdf.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If eventcalendar_th.eventcalendar_subject.Visible Then ' eventcalendar_subject %>
	<% If eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_subject) = "" Then %>
		<td><div id="elh_eventcalendar_th_eventcalendar_subject" class="eventcalendar_th_eventcalendar_subject"><div class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_subject.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_subject) %>',1);"><div id="elh_eventcalendar_th_eventcalendar_subject" class="eventcalendar_th_eventcalendar_subject">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_subject.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If eventcalendar_th.eventcalendar_subject.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf eventcalendar_th.eventcalendar_subject.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If eventcalendar_th.eventcalendar_subject_th.Visible Then ' eventcalendar_subject_th %>
	<% If eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_subject_th) = "" Then %>
		<td><div id="elh_eventcalendar_th_eventcalendar_subject_th" class="eventcalendar_th_eventcalendar_subject_th"><div class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_subject_th.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_subject_th) %>',1);"><div id="elh_eventcalendar_th_eventcalendar_subject_th" class="eventcalendar_th_eventcalendar_subject_th">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_subject_th.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If eventcalendar_th.eventcalendar_subject_th.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf eventcalendar_th.eventcalendar_subject_th.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If eventcalendar_th.eventcalendar_show_en.Visible Then ' eventcalendar_show_en %>
	<% If eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_show_en) = "" Then %>
		<td><div id="elh_eventcalendar_th_eventcalendar_show_en" class="eventcalendar_th_eventcalendar_show_en"><div class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_show_en.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_show_en) %>',1);"><div id="elh_eventcalendar_th_eventcalendar_show_en" class="eventcalendar_th_eventcalendar_show_en">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_show_en.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If eventcalendar_th.eventcalendar_show_en.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf eventcalendar_th.eventcalendar_show_en.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If eventcalendar_th.eventcalendar_show.Visible Then ' eventcalendar_show %>
	<% If eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_show) = "" Then %>
		<td><div id="elh_eventcalendar_th_eventcalendar_show" class="eventcalendar_th_eventcalendar_show"><div class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_show.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_show) %>',1);"><div id="elh_eventcalendar_th_eventcalendar_show" class="eventcalendar_th_eventcalendar_show">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_show.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If eventcalendar_th.eventcalendar_show.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf eventcalendar_th.eventcalendar_show.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If eventcalendar_th.eventcalendar_show_home.Visible Then ' eventcalendar_show_home %>
	<% If eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_show_home) = "" Then %>
		<td><div id="elh_eventcalendar_th_eventcalendar_show_home" class="eventcalendar_th_eventcalendar_show_home"><div class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_show_home.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_show_home) %>',1);"><div id="elh_eventcalendar_th_eventcalendar_show_home" class="eventcalendar_th_eventcalendar_show_home">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_show_home.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If eventcalendar_th.eventcalendar_show_home.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf eventcalendar_th.eventcalendar_show_home.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If eventcalendar_th.eventcalendar_create.Visible Then ' eventcalendar_create %>
	<% If eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_create) = "" Then %>
		<td><div id="elh_eventcalendar_th_eventcalendar_create" class="eventcalendar_th_eventcalendar_create"><div class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_create.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_create) %>',1);"><div id="elh_eventcalendar_th_eventcalendar_create" class="eventcalendar_th_eventcalendar_create">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_create.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If eventcalendar_th.eventcalendar_create.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf eventcalendar_th.eventcalendar_create.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If eventcalendar_th.eventcalendar_update.Visible Then ' eventcalendar_update %>
	<% If eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_update) = "" Then %>
		<td><div id="elh_eventcalendar_th_eventcalendar_update" class="eventcalendar_th_eventcalendar_update"><div class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_update.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= eventcalendar_th.SortUrl(eventcalendar_th.eventcalendar_update) %>',1);"><div id="elh_eventcalendar_th_eventcalendar_update" class="eventcalendar_th_eventcalendar_update">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= eventcalendar_th.eventcalendar_update.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If eventcalendar_th.eventcalendar_update.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf eventcalendar_th.eventcalendar_update.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
eventcalendar_th_list.ListOptions.Render "header", "right", "", "", "", ""
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (eventcalendar_th.ExportAll And eventcalendar_th.Export <> "") Then
	eventcalendar_th_list.StopRec = eventcalendar_th_list.TotalRecs
Else

	' Set the last record to display
	If eventcalendar_th_list.TotalRecs > eventcalendar_th_list.StartRec + eventcalendar_th_list.DisplayRecs - 1 Then
		eventcalendar_th_list.StopRec = eventcalendar_th_list.StartRec + eventcalendar_th_list.DisplayRecs - 1
	Else
		eventcalendar_th_list.StopRec = eventcalendar_th_list.TotalRecs
	End If
End If

' Move to first record
eventcalendar_th_list.RecCnt = eventcalendar_th_list.StartRec - 1
If Not eventcalendar_th_list.Recordset.Eof Then
	eventcalendar_th_list.Recordset.MoveFirst
	If eventcalendar_th_list.StartRec > 1 Then eventcalendar_th_list.Recordset.Move eventcalendar_th_list.StartRec - 1
ElseIf Not eventcalendar_th.AllowAddDeleteRow And eventcalendar_th_list.StopRec = 0 Then
	eventcalendar_th_list.StopRec = eventcalendar_th.GridAddRowCount
End If

' Initialize Aggregate
eventcalendar_th.RowType = EW_ROWTYPE_AGGREGATEINIT
Call eventcalendar_th.ResetAttrs()
Call eventcalendar_th_list.RenderRow()
eventcalendar_th_list.RowCnt = 0

' Output date rows
Do While CLng(eventcalendar_th_list.RecCnt) < CLng(eventcalendar_th_list.StopRec)
	eventcalendar_th_list.RecCnt = eventcalendar_th_list.RecCnt + 1
	If CLng(eventcalendar_th_list.RecCnt) >= CLng(eventcalendar_th_list.StartRec) Then
		eventcalendar_th_list.RowCnt = eventcalendar_th_list.RowCnt + 1

	' Set up key count
	eventcalendar_th_list.KeyCount = eventcalendar_th_list.RowIndex
	Call eventcalendar_th.ResetAttrs()
	eventcalendar_th.CssClass = ""
	If eventcalendar_th.CurrentAction = "gridadd" Then
	Else
		Call eventcalendar_th_list.LoadRowValues(eventcalendar_th_list.Recordset) ' Load row values
	End If
	eventcalendar_th.RowType = EW_ROWTYPE_VIEW ' Render view

	' Set up row id / data-rowindex
	eventcalendar_th.RowAttrs.AddAttributes Array(Array("data-rowindex", eventcalendar_th_list.RowCnt), Array("id", "r" & eventcalendar_th_list.RowCnt & "_eventcalendar_th"), Array("data-rowtype", eventcalendar_th.RowType))

	' Render row
	Call eventcalendar_th_list.RenderRow()

	' Render list options
	Call eventcalendar_th_list.RenderListOptions()
%>
	<tr<%= eventcalendar_th.RowAttributes %>>
<%

' Render list options (body, left)
eventcalendar_th_list.ListOptions.Render "body", "left", eventcalendar_th_list.RowCnt, "", "", ""
%>
	<% If eventcalendar_th.eventcalendar_id.Visible Then ' eventcalendar_id %>
		<td<%= eventcalendar_th.eventcalendar_id.CellAttributes %>>
<span<%= eventcalendar_th.eventcalendar_id.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_id.ListViewValue %>
</span>
<a id="<%= eventcalendar_th_list.PageObjName & "_row_" & eventcalendar_th_list.RowCnt %>"></a></td>
	<% End If %>
	<% If eventcalendar_th.eventcalendar_img.Visible Then ' eventcalendar_img %>
		<td<%= eventcalendar_th.eventcalendar_img.CellAttributes %>>
<span<%= eventcalendar_th.eventcalendar_img.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_img.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If eventcalendar_th.eventcalendar_date.Visible Then ' eventcalendar_date %>
		<td<%= eventcalendar_th.eventcalendar_date.CellAttributes %>>
<span<%= eventcalendar_th.eventcalendar_date.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_date.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If eventcalendar_th.eventcalendar_category.Visible Then ' eventcalendar_category %>
		<td<%= eventcalendar_th.eventcalendar_category.CellAttributes %>>
<span<%= eventcalendar_th.eventcalendar_category.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_category.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If eventcalendar_th.eventcalendar_category_sub.Visible Then ' eventcalendar_category_sub %>
		<td<%= eventcalendar_th.eventcalendar_category_sub.CellAttributes %>>
<span<%= eventcalendar_th.eventcalendar_category_sub.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_category_sub.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If eventcalendar_th.start_date.Visible Then ' start_date %>
		<td<%= eventcalendar_th.start_date.CellAttributes %>>
<span<%= eventcalendar_th.start_date.ViewAttributes %>>
<%= eventcalendar_th.start_date.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If eventcalendar_th.end_date.Visible Then ' end_date %>
		<td<%= eventcalendar_th.end_date.CellAttributes %>>
<span<%= eventcalendar_th.end_date.ViewAttributes %>>
<%= eventcalendar_th.end_date.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If eventcalendar_th.eventcalendar_pdf.Visible Then ' eventcalendar_pdf %>
		<td<%= eventcalendar_th.eventcalendar_pdf.CellAttributes %>>
<span<%= eventcalendar_th.eventcalendar_pdf.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_pdf.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If eventcalendar_th.eventcalendar_subject.Visible Then ' eventcalendar_subject %>
		<td<%= eventcalendar_th.eventcalendar_subject.CellAttributes %>>
<span<%= eventcalendar_th.eventcalendar_subject.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_subject.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If eventcalendar_th.eventcalendar_subject_th.Visible Then ' eventcalendar_subject_th %>
		<td<%= eventcalendar_th.eventcalendar_subject_th.CellAttributes %>>
<span<%= eventcalendar_th.eventcalendar_subject_th.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_subject_th.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If eventcalendar_th.eventcalendar_show_en.Visible Then ' eventcalendar_show_en %>
		<td<%= eventcalendar_th.eventcalendar_show_en.CellAttributes %>>
<span<%= eventcalendar_th.eventcalendar_show_en.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_show_en.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If eventcalendar_th.eventcalendar_show.Visible Then ' eventcalendar_show %>
		<td<%= eventcalendar_th.eventcalendar_show.CellAttributes %>>
<span<%= eventcalendar_th.eventcalendar_show.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_show.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If eventcalendar_th.eventcalendar_show_home.Visible Then ' eventcalendar_show_home %>
		<td<%= eventcalendar_th.eventcalendar_show_home.CellAttributes %>>
<span<%= eventcalendar_th.eventcalendar_show_home.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_show_home.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If eventcalendar_th.eventcalendar_create.Visible Then ' eventcalendar_create %>
		<td<%= eventcalendar_th.eventcalendar_create.CellAttributes %>>
<span<%= eventcalendar_th.eventcalendar_create.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_create.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If eventcalendar_th.eventcalendar_update.Visible Then ' eventcalendar_update %>
		<td<%= eventcalendar_th.eventcalendar_update.CellAttributes %>>
<span<%= eventcalendar_th.eventcalendar_update.ViewAttributes %>>
<%= eventcalendar_th.eventcalendar_update.ListViewValue %>
</span>
</td>
	<% End If %>
<%

' Render list options (body, right)
eventcalendar_th_list.ListOptions.Render "body", "right", eventcalendar_th_list.RowCnt, "", "", ""
%>
	</tr>
<%
	End If
	If eventcalendar_th.CurrentAction <> "gridadd" Then
		eventcalendar_th_list.Recordset.MoveNext()
	End If
Loop
%>
</tbody>
</table>
<% End If %>
<% If eventcalendar_th.CurrentAction = "" Then %>
<input type="hidden" name="a_list" id="a_list" value="">
<% End If %>
</div>
</form>
<%

' Close recordset and connection
eventcalendar_th_list.Recordset.Close
Set eventcalendar_th_list.Recordset = Nothing
%>
<% If eventcalendar_th.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If eventcalendar_th.CurrentAction <> "gridadd" And eventcalendar_th.CurrentAction <> "gridedit" Then %>
<form name="ewPagerForm" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewPager">
<tr><td>
<% If Not IsObject(eventcalendar_th_list.Pager) Then Set eventcalendar_th_list.Pager = ew_NewPrevNextPager(eventcalendar_th_list.StartRec, eventcalendar_th_list.DisplayRecs, eventcalendar_th_list.TotalRecs) %>
<% If eventcalendar_th_list.Pager.RecordCount > 0 Then %>
<table class="ewStdTable"><tbody><tr><td>
	<%= Language.Phrase("Page") %>&nbsp;
<div class="input-prepend input-append">
<!--first page button-->
	<% If eventcalendar_th_list.Pager.FirstButton.Enabled Then %>
	<a class="btn btn-small" href="<%= eventcalendar_th_list.PageUrl %>start=<%= eventcalendar_th_list.Pager.FirstButton.Start %>"><i class="icon-step-backward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-backward"></i></a>
	<% End If %>
<!--previous page button-->
	<% If eventcalendar_th_list.Pager.PrevButton.Enabled Then %>
	<a class="btn btn-small" href="<%= eventcalendar_th_list.PageUrl %>start=<%= eventcalendar_th_list.Pager.PrevButton.Start %>"><i class="icon-prev"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-prev"></i></a>
	<% End If %>
<!--current page number-->
	<input class="input-mini" type="text" name="<%= EW_TABLE_PAGE_NO %>" value="<%= eventcalendar_th_list.Pager.CurrentPage %>">
<!--next page button-->
	<% If eventcalendar_th_list.Pager.NextButton.Enabled Then %>
	<a class="btn btn-small" href="<%= eventcalendar_th_list.PageUrl %>start=<%= eventcalendar_th_list.Pager.NextButton.Start %>"><i class="icon-play"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-play"></i></a>
	<% End If %>
<!--last page button-->
	<% If eventcalendar_th_list.Pager.LastButton.Enabled Then %>
	<a class="btn btn-small" href="<%= eventcalendar_th_list.PageUrl %>start=<%= eventcalendar_th_list.Pager.LastButton.Start %>"><i class="icon-step-forward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-forward"></i></a>
	<% End If %>
</div>
	&nbsp;<%= Language.Phrase("of") %>&nbsp;<%= eventcalendar_th_list.Pager.PageCount %>
</td>
<td>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<%= Language.Phrase("Record") %>&nbsp;<%= eventcalendar_th_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= eventcalendar_th_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= eventcalendar_th_list.Pager.RecordCount %>
</td>
</tr></tbody></table>
<% Else %>
	<% If eventcalendar_th_list.SearchWhere = "0=101" Then %>
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
	eventcalendar_th_list.AddEditOptions.Render "body", "bottom", "", "", "", ""
	eventcalendar_th_list.DetailOptions.Render "body", "bottom", "", "", "", ""
	eventcalendar_th_list.ActionOptions.Render "body", "bottom", "", "", "", ""
%>
</div>
</div>
<% End If %>
</td></tr></table>
<% If eventcalendar_th.Export = "" Then %>
<script type="text/javascript">
feventcalendar_thlistsrch.Init();
feventcalendar_thlist.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<% End If %>
<%
eventcalendar_th_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If eventcalendar_th.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set eventcalendar_th_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class ceventcalendar_th_list

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
		TableName = "eventcalendar_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "eventcalendar_th_list"
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
		If eventcalendar_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & eventcalendar_th.TableVar & "&" ' add page token
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

		' Grid form hidden field names
		FormName = "feventcalendar_thlist"
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
		If IsEmpty(eventcalendar_th) Then Set eventcalendar_th = New ceventcalendar_th
		Set Table = eventcalendar_th

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		ExportPdfUrl = PageUrl & "export=pdf"
		AddUrl = "pom_eventcalendar_thadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "pom_eventcalendar_thdelete.asp"
		MultiUpdateUrl = "pom_eventcalendar_thupdate.asp"

		' Initialize other table object
		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "eventcalendar_th"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' List options
		Set ListOptions = New cListOptions
		ListOptions.TableVar = eventcalendar_th.TableVar

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = eventcalendar_th.TableVar
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
				eventcalendar_th.GridAddRowCount = gridaddcnt
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
		If UBound(eventcalendar_th.CustomActions.CustomArray) >= 0 Then
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
		Set eventcalendar_th = Nothing
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
			If eventcalendar_th.Export = "" Then
				SetupBreadcrumb()
			End If

			' Hide list options
			If eventcalendar_th.Export <> "" Then
				Call ListOptions.HideAllOptions(Array("sequence"))
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			ElseIf eventcalendar_th.CurrentAction = "gridadd" Or eventcalendar_th.CurrentAction = "gridedit" Then
				Call ListOptions.HideAllOptions(Array())
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			End If

			' Hide export options
			If eventcalendar_th.Export <> "" Or eventcalendar_th.CurrentAction <> "" Then
				ExportOptions.HideAllOptions(Array())
			End If

			' Hide other options
			If eventcalendar_th.Export <> "" Then
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
			Call eventcalendar_th.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If eventcalendar_th.RecordsPerPage <> "" Then
			DisplayRecs = eventcalendar_th.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Load search default if no existing search criteria
		If Not CheckSearchParms() Then

			' Load basic search from default
			eventcalendar_th.BasicSearch.Keyword = eventcalendar_th.BasicSearch.KeywordDefault
			eventcalendar_th.BasicSearch.SearchType = eventcalendar_th.BasicSearch.SearchTypeDefault
			eventcalendar_th.BasicSearch.setSearchType(eventcalendar_th.BasicSearch.SearchTypeDefault)
			If eventcalendar_th.BasicSearch.Keyword <> "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call eventcalendar_th.Recordset_Searching(SearchWhere)

		' Save search criteria
		If Command = "search" And Not RestoreSearch Then
			eventcalendar_th.SearchWhere = SearchWhere ' Save to Session
			StartRec = 1 ' Reset start record counter
			eventcalendar_th.StartRecordNumber = StartRec
		Else
			SearchWhere = eventcalendar_th.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		eventcalendar_th.SessionWhere = sFilter
		eventcalendar_th.CurrentFilter = ""
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
				sFilter = eventcalendar_th.KeyFilter
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
			eventcalendar_th.eventcalendar_id.FormValue = arrKeyFlds(0)
			If Not IsNumeric(eventcalendar_th.eventcalendar_id.FormValue) Then
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
			Call BuildBasicSearchSQL(sWhere, eventcalendar_th.eventcalendar_img, Keyword)
			Call BuildBasicSearchSQL(sWhere, eventcalendar_th.eventcalendar_category, Keyword)
			Call BuildBasicSearchSQL(sWhere, eventcalendar_th.eventcalendar_category_sub, Keyword)
			Call BuildBasicSearchSQL(sWhere, eventcalendar_th.eventcalendar_pdf, Keyword)
			Call BuildBasicSearchSQL(sWhere, eventcalendar_th.eventcalendar_subject, Keyword)
			Call BuildBasicSearchSQL(sWhere, eventcalendar_th.eventcalendar_subject_th, Keyword)
			Call BuildBasicSearchSQL(sWhere, eventcalendar_th.eventcalendar_intro, Keyword)
			Call BuildBasicSearchSQL(sWhere, eventcalendar_th.eventcalendar_intro_th, Keyword)
			Call BuildBasicSearchSQL(sWhere, eventcalendar_th.eventcalendar_content, Keyword)
			Call BuildBasicSearchSQL(sWhere, eventcalendar_th.eventcalendar_content_th, Keyword)
			Call BuildBasicSearchSQL(sWhere, eventcalendar_th.eventcalendar_show_en, Keyword)
			Call BuildBasicSearchSQL(sWhere, eventcalendar_th.eventcalendar_show, Keyword)
			Call BuildBasicSearchSQL(sWhere, eventcalendar_th.eventcalendar_show_home, Keyword)
			Call BuildBasicSearchSQL(sWhere, eventcalendar_th.eventcalendar_create, Keyword)
			Call BuildBasicSearchSQL(sWhere, eventcalendar_th.eventcalendar_update, Keyword)
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
		sSearchKeyword = eventcalendar_th.BasicSearch.Keyword
		sSearchType = eventcalendar_th.BasicSearch.SearchType
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
			eventcalendar_th.BasicSearch.setKeyword(sSearchKeyword)
			eventcalendar_th.BasicSearch.setSearchType(sSearchType)
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' Check if search parm exists
	Function CheckSearchParms()

		' Check basic search
		If eventcalendar_th.BasicSearch.IssetSession() Then
			CheckSearchParms = True
			Exit Function
		End If
		CheckSearchParms = False
	End Function

	' Clear all search parameters
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		eventcalendar_th.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' Load advanced search default values
	Function LoadAdvancedSearchDefault()
		LoadAdvancedSearchDefault = False
	End Function

	' Clear all basic search parameters
	Sub ResetBasicSearchParms()
		eventcalendar_th.BasicSearch.UnsetSession()
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()

		' Restore search flag
		RestoreSearch = True

		' Restore basic search values
		Call eventcalendar_th.BasicSearch.Load()
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
			eventcalendar_th.CurrentOrder = Request.QueryString("order")
			eventcalendar_th.CurrentOrderType = Request.QueryString("ordertype")

			' Field eventcalendar_id
			Call eventcalendar_th.UpdateSort(eventcalendar_th.eventcalendar_id)

			' Field eventcalendar_img
			Call eventcalendar_th.UpdateSort(eventcalendar_th.eventcalendar_img)

			' Field eventcalendar_date
			Call eventcalendar_th.UpdateSort(eventcalendar_th.eventcalendar_date)

			' Field eventcalendar_category
			Call eventcalendar_th.UpdateSort(eventcalendar_th.eventcalendar_category)

			' Field eventcalendar_category_sub
			Call eventcalendar_th.UpdateSort(eventcalendar_th.eventcalendar_category_sub)

			' Field start_date
			Call eventcalendar_th.UpdateSort(eventcalendar_th.start_date)

			' Field end_date
			Call eventcalendar_th.UpdateSort(eventcalendar_th.end_date)

			' Field eventcalendar_pdf
			Call eventcalendar_th.UpdateSort(eventcalendar_th.eventcalendar_pdf)

			' Field eventcalendar_subject
			Call eventcalendar_th.UpdateSort(eventcalendar_th.eventcalendar_subject)

			' Field eventcalendar_subject_th
			Call eventcalendar_th.UpdateSort(eventcalendar_th.eventcalendar_subject_th)

			' Field eventcalendar_show_en
			Call eventcalendar_th.UpdateSort(eventcalendar_th.eventcalendar_show_en)

			' Field eventcalendar_show
			Call eventcalendar_th.UpdateSort(eventcalendar_th.eventcalendar_show)

			' Field eventcalendar_show_home
			Call eventcalendar_th.UpdateSort(eventcalendar_th.eventcalendar_show_home)

			' Field eventcalendar_create
			Call eventcalendar_th.UpdateSort(eventcalendar_th.eventcalendar_create)

			' Field eventcalendar_update
			Call eventcalendar_th.UpdateSort(eventcalendar_th.eventcalendar_update)
			eventcalendar_th.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = eventcalendar_th.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If eventcalendar_th.SqlOrderBy <> "" Then
				sOrderBy = eventcalendar_th.SqlOrderBy
				eventcalendar_th.SessionOrderBy = sOrderBy
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
				eventcalendar_th.SessionOrderBy = sOrderBy
				eventcalendar_th.eventcalendar_id.Sort = ""
				eventcalendar_th.eventcalendar_img.Sort = ""
				eventcalendar_th.eventcalendar_date.Sort = ""
				eventcalendar_th.eventcalendar_category.Sort = ""
				eventcalendar_th.eventcalendar_category_sub.Sort = ""
				eventcalendar_th.start_date.Sort = ""
				eventcalendar_th.end_date.Sort = ""
				eventcalendar_th.eventcalendar_pdf.Sort = ""
				eventcalendar_th.eventcalendar_subject.Sort = ""
				eventcalendar_th.eventcalendar_subject_th.Sort = ""
				eventcalendar_th.eventcalendar_show_en.Sort = ""
				eventcalendar_th.eventcalendar_show.Sort = ""
				eventcalendar_th.eventcalendar_show_home.Sort = ""
				eventcalendar_th.eventcalendar_create.Sort = ""
				eventcalendar_th.eventcalendar_update.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			eventcalendar_th.StartRecordNumber = StartRec
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
		ListOptions.GetItem("checkbox").Body = "<label class=""checkbox""><input type=""checkbox"" name=""key_m"" value=""" & ew_HtmlEncode(eventcalendar_th.eventcalendar_id.CurrentValue) & """ onclick='ew_ClickMultiCheckbox(event, this);'></label>"
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
			For i = 0 to UBound(eventcalendar_th.CustomActions.CustomArray)
				Action = eventcalendar_th.CustomActions.CustomArray(i)(0)
				Name = eventcalendar_th.CustomActions.CustomArray(i)(1)

				' Add custom action
				Call opt.Add("custom_" & Action)
				Set item = opt.GetItem("custom_" & Action)
				item.Body = "<a class=""ewAction ewCustomAction"" href="""" onclick=""ew_SubmitSelected(document.feventcalendar_thlist, '" & ew_CurrentUrl & "', null, '" & Action & "');return false;"">" & Name & "</a>"
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
		sFilter = eventcalendar_th.GetKeyFilter
		UserAction = Request.Form("useraction") & ""
		Processed = False
		If sFilter <> "" And UserAction <> "" Then
			eventcalendar_th.CurrentFilter = sFilter
			sSql = eventcalendar_th.SQL
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
				ElseIf eventcalendar_th.CancelMessage <> "" Then
					FailureMessage = eventcalendar_th.CancelMessage
					eventcalendar_th.CancelMessage = ""
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
				eventcalendar_th.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					eventcalendar_th.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = eventcalendar_th.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			eventcalendar_th.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			eventcalendar_th.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			eventcalendar_th.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		eventcalendar_th.BasicSearch.Keyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)&""
		If eventcalendar_th.BasicSearch.Keyword <> "" Then Command = "search"
		eventcalendar_th.BasicSearch.SearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)&""
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = eventcalendar_th.CurrentFilter
		Call eventcalendar_th.Recordset_Selecting(sFilter)
		eventcalendar_th.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = eventcalendar_th.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call eventcalendar_th.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
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
		ViewUrl = eventcalendar_th.ViewUrl("")
		EditUrl = eventcalendar_th.EditUrl("")
		InlineEditUrl = eventcalendar_th.InlineEditUrl
		CopyUrl = eventcalendar_th.CopyUrl("")
		InlineCopyUrl = eventcalendar_th.InlineCopyUrl
		DeleteUrl = eventcalendar_th.DeleteUrl

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
		End If

		' Call Row Rendered event
		If eventcalendar_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call eventcalendar_th.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		url = ew_CurrentUrl
		url = ew_RegExReplace("\?cmd=reset(all){0,1}$", url, "") ' Remove cmd=reset / cmd=resetall
		Call Breadcrumb.Add("list", eventcalendar_th.TableVar, url, eventcalendar_th.TableVar, True)
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
