<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_query2info.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim Query2_list
Set Query2_list = New cQuery2_list
Set Page = Query2_list

' Page init processing
Query2_list.Page_Init()

' Page main processing
Query2_list.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
Query2_list.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If Query2.Export = "" Then %>
<script type="text/javascript">
// Page object
var Query2_list = new ew_Page("Query2_list");
Query2_list.PageID = "list"; // Page ID
var EW_PAGE_ID = Query2_list.PageID; // For backward compatibility
// Form object
var fQuery2list = new ew_Form("fQuery2list");
fQuery2list.FormKeyCountName = '<%= Query2_list.FormKeyCountName %>';
// Form_CustomValidate event
fQuery2list.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fQuery2list.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fQuery2list.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
var fQuery2listsrch = new ew_Form("fQuery2listsrch");
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If Query2.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If Query2_list.ExportOptions.Visible Then %>
<div class="ewListExportOptions"><% Query2_list.ExportOptions.Render "body", "", "", "", "", "" %></div>
<% End If %>
<% If (Query2.Export = "") Or (EW_EXPORT_MASTER_RECORD And Query2.Export = "print") Then %>
<% End If %>
<%

' Load recordset
Set Query2_list.Recordset = Query2_list.LoadRecordset()
	Query2_list.TotalRecs = Query2_list.Recordset.RecordCount
	Query2_list.StartRec = 1
	If Query2_list.DisplayRecs <= 0 Then ' Display all records
		Query2_list.DisplayRecs = Query2_list.TotalRecs
	End If
	If Not (Query2.ExportAll And Query2.Export <> "") Then
		Query2_list.SetUpStartRec() ' Set up start record position
	End If
Query2_list.RenderOtherOptions()
%>
<% If Security.IsLoggedIn() Then %>
<% If Query2.Export = "" And Query2.CurrentAction = "" Then %>
<form name="fQuery2listsrch" id="fQuery2listsrch" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewSearchTable"><tr><td>
<div class="accordion" id="fQuery2listsrch_SearchGroup">
	<div class="accordion-group">
		<div class="accordion-heading">
<a class="accordion-toggle" data-toggle="collapse" data-parent="#fQuery2listsrch_SearchGroup" href="#fQuery2listsrch_SearchBody"><%= Language.Phrase("Search") %></a>
		</div>
		<div id="fQuery2listsrch_SearchBody" class="accordion-body collapse in">
			<div class="accordion-inner">
<div id="fQuery2listsrch_SearchPanel">
<input type="hidden" name="cmd" value="search">
<input type="hidden" name="t" value="Query2">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewRow">
	<div class="btn-group ewButtonGroup">
	<div class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" class="input-large" value="<%= ew_HtmlEncode(Query2.BasicSearch.getKeyword()) %>" placeholder="<%= ew_HtmlEncode(Language.Phrase("Search")) %>">
	<button class="btn btn-primary ewButton" name="btnsubmit" id="btnsubmit" type="submit"><%= Language.Phrase("QuickSearchBtn") %></button>
	</div>
	</div>
	<div class="btn-group ewButtonGroup">
	<a class="btn ewShowAll" href="<%= Query2_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>
	</div>
</div>
<div id="xsr_2" class="ewRow">
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="="<% If Query2.BasicSearch.getSearchType = "=" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If Query2.BasicSearch.getSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If Query2.BasicSearch.getSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
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
<% Query2_list.ShowPageHeader() %>
<% Query2_list.ShowMessage %>
<table class="ewGrid"><tr><td class="ewGridContent">
<form name="fQuery2list" id="fQuery2list" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="Query2">
<div id="gmp_Query2" class="ewGridMiddlePanel">
<% If Query2_list.TotalRecs > 0 Then %>
<table id="tbl_Query2list" class="ewTable ewTableSeparate">
<%= Query2.TableCustomInnerHtml %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call Query2_list.RenderListOptions()

' Render list options (header, left)
Query2_list.ListOptions.Render "header", "left", "", "", "", ""
%>
<% If Query2.Expr1.Visible Then ' Expr1 %>
	<% If Query2.SortUrl(Query2.Expr1) = "" Then %>
		<td><div id="elh_Query2_Expr1" class="Query2_Expr1"><div class="ewTableHeaderCaption"><%= Query2.Expr1.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr1) %>',1);"><div id="elh_Query2_Expr1" class="Query2_Expr1">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr1.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr1.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr1.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr2.Visible Then ' Expr2 %>
	<% If Query2.SortUrl(Query2.Expr2) = "" Then %>
		<td><div id="elh_Query2_Expr2" class="Query2_Expr2"><div class="ewTableHeaderCaption"><%= Query2.Expr2.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr2) %>',1);"><div id="elh_Query2_Expr2" class="Query2_Expr2">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr2.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr2.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr2.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr3.Visible Then ' Expr3 %>
	<% If Query2.SortUrl(Query2.Expr3) = "" Then %>
		<td><div id="elh_Query2_Expr3" class="Query2_Expr3"><div class="ewTableHeaderCaption"><%= Query2.Expr3.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr3) %>',1);"><div id="elh_Query2_Expr3" class="Query2_Expr3">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr3.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr3.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr3.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr4.Visible Then ' Expr4 %>
	<% If Query2.SortUrl(Query2.Expr4) = "" Then %>
		<td><div id="elh_Query2_Expr4" class="Query2_Expr4"><div class="ewTableHeaderCaption"><%= Query2.Expr4.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr4) %>',1);"><div id="elh_Query2_Expr4" class="Query2_Expr4">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr4.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr4.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr4.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr5.Visible Then ' Expr5 %>
	<% If Query2.SortUrl(Query2.Expr5) = "" Then %>
		<td><div id="elh_Query2_Expr5" class="Query2_Expr5"><div class="ewTableHeaderCaption"><%= Query2.Expr5.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr5) %>',1);"><div id="elh_Query2_Expr5" class="Query2_Expr5">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr5.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr5.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr5.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr6.Visible Then ' Expr6 %>
	<% If Query2.SortUrl(Query2.Expr6) = "" Then %>
		<td><div id="elh_Query2_Expr6" class="Query2_Expr6"><div class="ewTableHeaderCaption"><%= Query2.Expr6.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr6) %>',1);"><div id="elh_Query2_Expr6" class="Query2_Expr6">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr6.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr6.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr6.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr7.Visible Then ' Expr7 %>
	<% If Query2.SortUrl(Query2.Expr7) = "" Then %>
		<td><div id="elh_Query2_Expr7" class="Query2_Expr7"><div class="ewTableHeaderCaption"><%= Query2.Expr7.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr7) %>',1);"><div id="elh_Query2_Expr7" class="Query2_Expr7">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr7.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr7.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr7.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr8.Visible Then ' Expr8 %>
	<% If Query2.SortUrl(Query2.Expr8) = "" Then %>
		<td><div id="elh_Query2_Expr8" class="Query2_Expr8"><div class="ewTableHeaderCaption"><%= Query2.Expr8.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr8) %>',1);"><div id="elh_Query2_Expr8" class="Query2_Expr8">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr8.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr8.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr8.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr9.Visible Then ' Expr9 %>
	<% If Query2.SortUrl(Query2.Expr9) = "" Then %>
		<td><div id="elh_Query2_Expr9" class="Query2_Expr9"><div class="ewTableHeaderCaption"><%= Query2.Expr9.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr9) %>',1);"><div id="elh_Query2_Expr9" class="Query2_Expr9">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr9.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr9.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr9.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr10.Visible Then ' Expr10 %>
	<% If Query2.SortUrl(Query2.Expr10) = "" Then %>
		<td><div id="elh_Query2_Expr10" class="Query2_Expr10"><div class="ewTableHeaderCaption"><%= Query2.Expr10.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr10) %>',1);"><div id="elh_Query2_Expr10" class="Query2_Expr10">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr10.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr10.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr10.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr11.Visible Then ' Expr11 %>
	<% If Query2.SortUrl(Query2.Expr11) = "" Then %>
		<td><div id="elh_Query2_Expr11" class="Query2_Expr11"><div class="ewTableHeaderCaption"><%= Query2.Expr11.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr11) %>',1);"><div id="elh_Query2_Expr11" class="Query2_Expr11">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr11.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr11.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr11.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr12.Visible Then ' Expr12 %>
	<% If Query2.SortUrl(Query2.Expr12) = "" Then %>
		<td><div id="elh_Query2_Expr12" class="Query2_Expr12"><div class="ewTableHeaderCaption"><%= Query2.Expr12.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr12) %>',1);"><div id="elh_Query2_Expr12" class="Query2_Expr12">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr12.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr12.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr12.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr13.Visible Then ' Expr13 %>
	<% If Query2.SortUrl(Query2.Expr13) = "" Then %>
		<td><div id="elh_Query2_Expr13" class="Query2_Expr13"><div class="ewTableHeaderCaption"><%= Query2.Expr13.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr13) %>',1);"><div id="elh_Query2_Expr13" class="Query2_Expr13">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr13.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr13.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr13.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr14.Visible Then ' Expr14 %>
	<% If Query2.SortUrl(Query2.Expr14) = "" Then %>
		<td><div id="elh_Query2_Expr14" class="Query2_Expr14"><div class="ewTableHeaderCaption"><%= Query2.Expr14.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr14) %>',1);"><div id="elh_Query2_Expr14" class="Query2_Expr14">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr14.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr14.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr14.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr15.Visible Then ' Expr15 %>
	<% If Query2.SortUrl(Query2.Expr15) = "" Then %>
		<td><div id="elh_Query2_Expr15" class="Query2_Expr15"><div class="ewTableHeaderCaption"><%= Query2.Expr15.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr15) %>',1);"><div id="elh_Query2_Expr15" class="Query2_Expr15">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr15.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr15.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr15.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr16.Visible Then ' Expr16 %>
	<% If Query2.SortUrl(Query2.Expr16) = "" Then %>
		<td><div id="elh_Query2_Expr16" class="Query2_Expr16"><div class="ewTableHeaderCaption"><%= Query2.Expr16.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr16) %>',1);"><div id="elh_Query2_Expr16" class="Query2_Expr16">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr16.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr16.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr16.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr17.Visible Then ' Expr17 %>
	<% If Query2.SortUrl(Query2.Expr17) = "" Then %>
		<td><div id="elh_Query2_Expr17" class="Query2_Expr17"><div class="ewTableHeaderCaption"><%= Query2.Expr17.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr17) %>',1);"><div id="elh_Query2_Expr17" class="Query2_Expr17">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr17.FldCaption %></span><span class="ewTableHeaderSort"><% If Query2.Expr17.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr17.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr18.Visible Then ' Expr18 %>
	<% If Query2.SortUrl(Query2.Expr18) = "" Then %>
		<td><div id="elh_Query2_Expr18" class="Query2_Expr18"><div class="ewTableHeaderCaption"><%= Query2.Expr18.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr18) %>',1);"><div id="elh_Query2_Expr18" class="Query2_Expr18">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr18.FldCaption %></span><span class="ewTableHeaderSort"><% If Query2.Expr18.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr18.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr19.Visible Then ' Expr19 %>
	<% If Query2.SortUrl(Query2.Expr19) = "" Then %>
		<td><div id="elh_Query2_Expr19" class="Query2_Expr19"><div class="ewTableHeaderCaption"><%= Query2.Expr19.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr19) %>',1);"><div id="elh_Query2_Expr19" class="Query2_Expr19">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr19.FldCaption %></span><span class="ewTableHeaderSort"><% If Query2.Expr19.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr19.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr20.Visible Then ' Expr20 %>
	<% If Query2.SortUrl(Query2.Expr20) = "" Then %>
		<td><div id="elh_Query2_Expr20" class="Query2_Expr20"><div class="ewTableHeaderCaption"><%= Query2.Expr20.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr20) %>',1);"><div id="elh_Query2_Expr20" class="Query2_Expr20">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr20.FldCaption %></span><span class="ewTableHeaderSort"><% If Query2.Expr20.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr20.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr21.Visible Then ' Expr21 %>
	<% If Query2.SortUrl(Query2.Expr21) = "" Then %>
		<td><div id="elh_Query2_Expr21" class="Query2_Expr21"><div class="ewTableHeaderCaption"><%= Query2.Expr21.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr21) %>',1);"><div id="elh_Query2_Expr21" class="Query2_Expr21">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr21.FldCaption %></span><span class="ewTableHeaderSort"><% If Query2.Expr21.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr21.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr22.Visible Then ' Expr22 %>
	<% If Query2.SortUrl(Query2.Expr22) = "" Then %>
		<td><div id="elh_Query2_Expr22" class="Query2_Expr22"><div class="ewTableHeaderCaption"><%= Query2.Expr22.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr22) %>',1);"><div id="elh_Query2_Expr22" class="Query2_Expr22">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr22.FldCaption %></span><span class="ewTableHeaderSort"><% If Query2.Expr22.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr22.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr23.Visible Then ' Expr23 %>
	<% If Query2.SortUrl(Query2.Expr23) = "" Then %>
		<td><div id="elh_Query2_Expr23" class="Query2_Expr23"><div class="ewTableHeaderCaption"><%= Query2.Expr23.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr23) %>',1);"><div id="elh_Query2_Expr23" class="Query2_Expr23">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr23.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr23.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr23.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr24.Visible Then ' Expr24 %>
	<% If Query2.SortUrl(Query2.Expr24) = "" Then %>
		<td><div id="elh_Query2_Expr24" class="Query2_Expr24"><div class="ewTableHeaderCaption"><%= Query2.Expr24.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr24) %>',1);"><div id="elh_Query2_Expr24" class="Query2_Expr24">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr24.FldCaption %></span><span class="ewTableHeaderSort"><% If Query2.Expr24.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr24.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr25.Visible Then ' Expr25 %>
	<% If Query2.SortUrl(Query2.Expr25) = "" Then %>
		<td><div id="elh_Query2_Expr25" class="Query2_Expr25"><div class="ewTableHeaderCaption"><%= Query2.Expr25.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr25) %>',1);"><div id="elh_Query2_Expr25" class="Query2_Expr25">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr25.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr25.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr25.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr26.Visible Then ' Expr26 %>
	<% If Query2.SortUrl(Query2.Expr26) = "" Then %>
		<td><div id="elh_Query2_Expr26" class="Query2_Expr26"><div class="ewTableHeaderCaption"><%= Query2.Expr26.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr26) %>',1);"><div id="elh_Query2_Expr26" class="Query2_Expr26">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr26.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr26.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr26.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr27.Visible Then ' Expr27 %>
	<% If Query2.SortUrl(Query2.Expr27) = "" Then %>
		<td><div id="elh_Query2_Expr27" class="Query2_Expr27"><div class="ewTableHeaderCaption"><%= Query2.Expr27.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr27) %>',1);"><div id="elh_Query2_Expr27" class="Query2_Expr27">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr27.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr27.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr27.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr28.Visible Then ' Expr28 %>
	<% If Query2.SortUrl(Query2.Expr28) = "" Then %>
		<td><div id="elh_Query2_Expr28" class="Query2_Expr28"><div class="ewTableHeaderCaption"><%= Query2.Expr28.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr28) %>',1);"><div id="elh_Query2_Expr28" class="Query2_Expr28">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr28.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr28.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr28.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr29.Visible Then ' Expr29 %>
	<% If Query2.SortUrl(Query2.Expr29) = "" Then %>
		<td><div id="elh_Query2_Expr29" class="Query2_Expr29"><div class="ewTableHeaderCaption"><%= Query2.Expr29.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr29) %>',1);"><div id="elh_Query2_Expr29" class="Query2_Expr29">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr29.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr29.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr29.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr30.Visible Then ' Expr30 %>
	<% If Query2.SortUrl(Query2.Expr30) = "" Then %>
		<td><div id="elh_Query2_Expr30" class="Query2_Expr30"><div class="ewTableHeaderCaption"><%= Query2.Expr30.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr30) %>',1);"><div id="elh_Query2_Expr30" class="Query2_Expr30">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr30.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr30.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr30.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr31.Visible Then ' Expr31 %>
	<% If Query2.SortUrl(Query2.Expr31) = "" Then %>
		<td><div id="elh_Query2_Expr31" class="Query2_Expr31"><div class="ewTableHeaderCaption"><%= Query2.Expr31.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr31) %>',1);"><div id="elh_Query2_Expr31" class="Query2_Expr31">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr31.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr31.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr31.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr32.Visible Then ' Expr32 %>
	<% If Query2.SortUrl(Query2.Expr32) = "" Then %>
		<td><div id="elh_Query2_Expr32" class="Query2_Expr32"><div class="ewTableHeaderCaption"><%= Query2.Expr32.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr32) %>',1);"><div id="elh_Query2_Expr32" class="Query2_Expr32">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr32.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr32.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr32.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr33.Visible Then ' Expr33 %>
	<% If Query2.SortUrl(Query2.Expr33) = "" Then %>
		<td><div id="elh_Query2_Expr33" class="Query2_Expr33"><div class="ewTableHeaderCaption"><%= Query2.Expr33.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr33) %>',1);"><div id="elh_Query2_Expr33" class="Query2_Expr33">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr33.FldCaption %></span><span class="ewTableHeaderSort"><% If Query2.Expr33.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr33.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr34.Visible Then ' Expr34 %>
	<% If Query2.SortUrl(Query2.Expr34) = "" Then %>
		<td><div id="elh_Query2_Expr34" class="Query2_Expr34"><div class="ewTableHeaderCaption"><%= Query2.Expr34.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr34) %>',1);"><div id="elh_Query2_Expr34" class="Query2_Expr34">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr34.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr34.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr34.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr35.Visible Then ' Expr35 %>
	<% If Query2.SortUrl(Query2.Expr35) = "" Then %>
		<td><div id="elh_Query2_Expr35" class="Query2_Expr35"><div class="ewTableHeaderCaption"><%= Query2.Expr35.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr35) %>',1);"><div id="elh_Query2_Expr35" class="Query2_Expr35">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr35.FldCaption %></span><span class="ewTableHeaderSort"><% If Query2.Expr35.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr35.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr36.Visible Then ' Expr36 %>
	<% If Query2.SortUrl(Query2.Expr36) = "" Then %>
		<td><div id="elh_Query2_Expr36" class="Query2_Expr36"><div class="ewTableHeaderCaption"><%= Query2.Expr36.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr36) %>',1);"><div id="elh_Query2_Expr36" class="Query2_Expr36">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr36.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr36.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr36.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If Query2.Expr37.Visible Then ' Expr37 %>
	<% If Query2.SortUrl(Query2.Expr37) = "" Then %>
		<td><div id="elh_Query2_Expr37" class="Query2_Expr37"><div class="ewTableHeaderCaption"><%= Query2.Expr37.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= Query2.SortUrl(Query2.Expr37) %>',1);"><div id="elh_Query2_Expr37" class="Query2_Expr37">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= Query2.Expr37.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If Query2.Expr37.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf Query2.Expr37.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
Query2_list.ListOptions.Render "header", "right", "", "", "", ""
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (Query2.ExportAll And Query2.Export <> "") Then
	Query2_list.StopRec = Query2_list.TotalRecs
Else

	' Set the last record to display
	If Query2_list.TotalRecs > Query2_list.StartRec + Query2_list.DisplayRecs - 1 Then
		Query2_list.StopRec = Query2_list.StartRec + Query2_list.DisplayRecs - 1
	Else
		Query2_list.StopRec = Query2_list.TotalRecs
	End If
End If

' Move to first record
Query2_list.RecCnt = Query2_list.StartRec - 1
If Not Query2_list.Recordset.Eof Then
	Query2_list.Recordset.MoveFirst
	If Query2_list.StartRec > 1 Then Query2_list.Recordset.Move Query2_list.StartRec - 1
ElseIf Not Query2.AllowAddDeleteRow And Query2_list.StopRec = 0 Then
	Query2_list.StopRec = Query2.GridAddRowCount
End If

' Initialize Aggregate
Query2.RowType = EW_ROWTYPE_AGGREGATEINIT
Call Query2.ResetAttrs()
Call Query2_list.RenderRow()
Query2_list.RowCnt = 0

' Output date rows
Do While CLng(Query2_list.RecCnt) < CLng(Query2_list.StopRec)
	Query2_list.RecCnt = Query2_list.RecCnt + 1
	If CLng(Query2_list.RecCnt) >= CLng(Query2_list.StartRec) Then
		Query2_list.RowCnt = Query2_list.RowCnt + 1

	' Set up key count
	Query2_list.KeyCount = Query2_list.RowIndex
	Call Query2.ResetAttrs()
	Query2.CssClass = ""
	If Query2.CurrentAction = "gridadd" Then
	Else
		Call Query2_list.LoadRowValues(Query2_list.Recordset) ' Load row values
	End If
	Query2.RowType = EW_ROWTYPE_VIEW ' Render view

	' Set up row id / data-rowindex
	Query2.RowAttrs.AddAttributes Array(Array("data-rowindex", Query2_list.RowCnt), Array("id", "r" & Query2_list.RowCnt & "_Query2"), Array("data-rowtype", Query2.RowType))

	' Render row
	Call Query2_list.RenderRow()

	' Render list options
	Call Query2_list.RenderListOptions()
%>
	<tr<%= Query2.RowAttributes %>>
<%

' Render list options (body, left)
Query2_list.ListOptions.Render "body", "left", Query2_list.RowCnt, "", "", ""
%>
	<% If Query2.Expr1.Visible Then ' Expr1 %>
		<td<%= Query2.Expr1.CellAttributes %>>
<span<%= Query2.Expr1.ViewAttributes %>>
<%= Query2.Expr1.ListViewValue %>
</span>
<a id="<%= Query2_list.PageObjName & "_row_" & Query2_list.RowCnt %>"></a></td>
	<% End If %>
	<% If Query2.Expr2.Visible Then ' Expr2 %>
		<td<%= Query2.Expr2.CellAttributes %>>
<span<%= Query2.Expr2.ViewAttributes %>>
<%= Query2.Expr2.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr3.Visible Then ' Expr3 %>
		<td<%= Query2.Expr3.CellAttributes %>>
<span<%= Query2.Expr3.ViewAttributes %>>
<%= Query2.Expr3.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr4.Visible Then ' Expr4 %>
		<td<%= Query2.Expr4.CellAttributes %>>
<span<%= Query2.Expr4.ViewAttributes %>>
<%= Query2.Expr4.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr5.Visible Then ' Expr5 %>
		<td<%= Query2.Expr5.CellAttributes %>>
<span<%= Query2.Expr5.ViewAttributes %>>
<%= Query2.Expr5.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr6.Visible Then ' Expr6 %>
		<td<%= Query2.Expr6.CellAttributes %>>
<span<%= Query2.Expr6.ViewAttributes %>>
<%= Query2.Expr6.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr7.Visible Then ' Expr7 %>
		<td<%= Query2.Expr7.CellAttributes %>>
<span<%= Query2.Expr7.ViewAttributes %>>
<%= Query2.Expr7.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr8.Visible Then ' Expr8 %>
		<td<%= Query2.Expr8.CellAttributes %>>
<span<%= Query2.Expr8.ViewAttributes %>>
<%= Query2.Expr8.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr9.Visible Then ' Expr9 %>
		<td<%= Query2.Expr9.CellAttributes %>>
<span<%= Query2.Expr9.ViewAttributes %>>
<%= Query2.Expr9.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr10.Visible Then ' Expr10 %>
		<td<%= Query2.Expr10.CellAttributes %>>
<span<%= Query2.Expr10.ViewAttributes %>>
<%= Query2.Expr10.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr11.Visible Then ' Expr11 %>
		<td<%= Query2.Expr11.CellAttributes %>>
<span<%= Query2.Expr11.ViewAttributes %>>
<%= Query2.Expr11.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr12.Visible Then ' Expr12 %>
		<td<%= Query2.Expr12.CellAttributes %>>
<span<%= Query2.Expr12.ViewAttributes %>>
<%= Query2.Expr12.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr13.Visible Then ' Expr13 %>
		<td<%= Query2.Expr13.CellAttributes %>>
<span<%= Query2.Expr13.ViewAttributes %>>
<%= Query2.Expr13.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr14.Visible Then ' Expr14 %>
		<td<%= Query2.Expr14.CellAttributes %>>
<span<%= Query2.Expr14.ViewAttributes %>>
<%= Query2.Expr14.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr15.Visible Then ' Expr15 %>
		<td<%= Query2.Expr15.CellAttributes %>>
<span<%= Query2.Expr15.ViewAttributes %>>
<%= Query2.Expr15.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr16.Visible Then ' Expr16 %>
		<td<%= Query2.Expr16.CellAttributes %>>
<span<%= Query2.Expr16.ViewAttributes %>>
<%= Query2.Expr16.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr17.Visible Then ' Expr17 %>
		<td<%= Query2.Expr17.CellAttributes %>>
<span<%= Query2.Expr17.ViewAttributes %>>
<%= Query2.Expr17.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr18.Visible Then ' Expr18 %>
		<td<%= Query2.Expr18.CellAttributes %>>
<span<%= Query2.Expr18.ViewAttributes %>>
<%= Query2.Expr18.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr19.Visible Then ' Expr19 %>
		<td<%= Query2.Expr19.CellAttributes %>>
<span<%= Query2.Expr19.ViewAttributes %>>
<%= Query2.Expr19.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr20.Visible Then ' Expr20 %>
		<td<%= Query2.Expr20.CellAttributes %>>
<span<%= Query2.Expr20.ViewAttributes %>>
<%= Query2.Expr20.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr21.Visible Then ' Expr21 %>
		<td<%= Query2.Expr21.CellAttributes %>>
<span<%= Query2.Expr21.ViewAttributes %>>
<%= Query2.Expr21.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr22.Visible Then ' Expr22 %>
		<td<%= Query2.Expr22.CellAttributes %>>
<span<%= Query2.Expr22.ViewAttributes %>>
<%= Query2.Expr22.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr23.Visible Then ' Expr23 %>
		<td<%= Query2.Expr23.CellAttributes %>>
<span<%= Query2.Expr23.ViewAttributes %>>
<%= Query2.Expr23.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr24.Visible Then ' Expr24 %>
		<td<%= Query2.Expr24.CellAttributes %>>
<span<%= Query2.Expr24.ViewAttributes %>>
<%= Query2.Expr24.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr25.Visible Then ' Expr25 %>
		<td<%= Query2.Expr25.CellAttributes %>>
<span<%= Query2.Expr25.ViewAttributes %>>
<%= Query2.Expr25.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr26.Visible Then ' Expr26 %>
		<td<%= Query2.Expr26.CellAttributes %>>
<span<%= Query2.Expr26.ViewAttributes %>>
<%= Query2.Expr26.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr27.Visible Then ' Expr27 %>
		<td<%= Query2.Expr27.CellAttributes %>>
<span<%= Query2.Expr27.ViewAttributes %>>
<%= Query2.Expr27.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr28.Visible Then ' Expr28 %>
		<td<%= Query2.Expr28.CellAttributes %>>
<span<%= Query2.Expr28.ViewAttributes %>>
<%= Query2.Expr28.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr29.Visible Then ' Expr29 %>
		<td<%= Query2.Expr29.CellAttributes %>>
<span<%= Query2.Expr29.ViewAttributes %>>
<%= Query2.Expr29.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr30.Visible Then ' Expr30 %>
		<td<%= Query2.Expr30.CellAttributes %>>
<span<%= Query2.Expr30.ViewAttributes %>>
<%= Query2.Expr30.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr31.Visible Then ' Expr31 %>
		<td<%= Query2.Expr31.CellAttributes %>>
<span<%= Query2.Expr31.ViewAttributes %>>
<%= Query2.Expr31.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr32.Visible Then ' Expr32 %>
		<td<%= Query2.Expr32.CellAttributes %>>
<span<%= Query2.Expr32.ViewAttributes %>>
<%= Query2.Expr32.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr33.Visible Then ' Expr33 %>
		<td<%= Query2.Expr33.CellAttributes %>>
<span<%= Query2.Expr33.ViewAttributes %>>
<%= Query2.Expr33.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr34.Visible Then ' Expr34 %>
		<td<%= Query2.Expr34.CellAttributes %>>
<span<%= Query2.Expr34.ViewAttributes %>>
<%= Query2.Expr34.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr35.Visible Then ' Expr35 %>
		<td<%= Query2.Expr35.CellAttributes %>>
<span<%= Query2.Expr35.ViewAttributes %>>
<%= Query2.Expr35.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr36.Visible Then ' Expr36 %>
		<td<%= Query2.Expr36.CellAttributes %>>
<span<%= Query2.Expr36.ViewAttributes %>>
<%= Query2.Expr36.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If Query2.Expr37.Visible Then ' Expr37 %>
		<td<%= Query2.Expr37.CellAttributes %>>
<span<%= Query2.Expr37.ViewAttributes %>>
<%= Query2.Expr37.ListViewValue %>
</span>
</td>
	<% End If %>
<%

' Render list options (body, right)
Query2_list.ListOptions.Render "body", "right", Query2_list.RowCnt, "", "", ""
%>
	</tr>
<%
	End If
	If Query2.CurrentAction <> "gridadd" Then
		Query2_list.Recordset.MoveNext()
	End If
Loop
%>
</tbody>
</table>
<% End If %>
<% If Query2.CurrentAction = "" Then %>
<input type="hidden" name="a_list" id="a_list" value="">
<% End If %>
</div>
</form>
<%

' Close recordset and connection
Query2_list.Recordset.Close
Set Query2_list.Recordset = Nothing
%>
<% If Query2.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If Query2.CurrentAction <> "gridadd" And Query2.CurrentAction <> "gridedit" Then %>
<form name="ewPagerForm" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewPager">
<tr><td>
<% If Not IsObject(Query2_list.Pager) Then Set Query2_list.Pager = ew_NewPrevNextPager(Query2_list.StartRec, Query2_list.DisplayRecs, Query2_list.TotalRecs) %>
<% If Query2_list.Pager.RecordCount > 0 Then %>
<table class="ewStdTable"><tbody><tr><td>
	<%= Language.Phrase("Page") %>&nbsp;
<div class="input-prepend input-append">
<!--first page button-->
	<% If Query2_list.Pager.FirstButton.Enabled Then %>
	<a class="btn btn-small" href="<%= Query2_list.PageUrl %>start=<%= Query2_list.Pager.FirstButton.Start %>"><i class="icon-step-backward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-backward"></i></a>
	<% End If %>
<!--previous page button-->
	<% If Query2_list.Pager.PrevButton.Enabled Then %>
	<a class="btn btn-small" href="<%= Query2_list.PageUrl %>start=<%= Query2_list.Pager.PrevButton.Start %>"><i class="icon-prev"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-prev"></i></a>
	<% End If %>
<!--current page number-->
	<input class="input-mini" type="text" name="<%= EW_TABLE_PAGE_NO %>" value="<%= Query2_list.Pager.CurrentPage %>">
<!--next page button-->
	<% If Query2_list.Pager.NextButton.Enabled Then %>
	<a class="btn btn-small" href="<%= Query2_list.PageUrl %>start=<%= Query2_list.Pager.NextButton.Start %>"><i class="icon-play"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-play"></i></a>
	<% End If %>
<!--last page button-->
	<% If Query2_list.Pager.LastButton.Enabled Then %>
	<a class="btn btn-small" href="<%= Query2_list.PageUrl %>start=<%= Query2_list.Pager.LastButton.Start %>"><i class="icon-step-forward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-forward"></i></a>
	<% End If %>
</div>
	&nbsp;<%= Language.Phrase("of") %>&nbsp;<%= Query2_list.Pager.PageCount %>
</td>
<td>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<%= Language.Phrase("Record") %>&nbsp;<%= Query2_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= Query2_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= Query2_list.Pager.RecordCount %>
</td>
</tr></tbody></table>
<% Else %>
	<% If Query2_list.SearchWhere = "0=101" Then %>
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
	Query2_list.AddEditOptions.Render "body", "bottom", "", "", "", ""
	Query2_list.DetailOptions.Render "body", "bottom", "", "", "", ""
	Query2_list.ActionOptions.Render "body", "bottom", "", "", "", ""
%>
</div>
</div>
<% End If %>
</td></tr></table>
<% If Query2.Export = "" Then %>
<script type="text/javascript">
fQuery2listsrch.Init();
fQuery2list.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<% End If %>
<%
Query2_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If Query2.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set Query2_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cQuery2_list

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
		TableName = "Query2"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "Query2_list"
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
		If Query2.UseTokenInUrl Then PageUrl = PageUrl & "t=" & Query2.TableVar & "&" ' add page token
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
		If Query2.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (Query2.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (Query2.TableVar = Request.QueryString("t"))
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
		FormName = "fQuery2list"
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
		If IsEmpty(Query2) Then Set Query2 = New cQuery2
		Set Table = Query2

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		ExportPdfUrl = PageUrl & "export=pdf"
		AddUrl = "pom_query2add.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "pom_query2delete.asp"
		MultiUpdateUrl = "pom_query2update.asp"

		' Initialize other table object
		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "Query2"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' List options
		Set ListOptions = New cListOptions
		ListOptions.TableVar = Query2.TableVar

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = Query2.TableVar
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
				Query2.GridAddRowCount = gridaddcnt
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
		If UBound(Query2.CustomActions.CustomArray) >= 0 Then
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
		Set Query2 = Nothing
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
			If Query2.Export = "" Then
				SetupBreadcrumb()
			End If

			' Hide list options
			If Query2.Export <> "" Then
				Call ListOptions.HideAllOptions(Array("sequence"))
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			ElseIf Query2.CurrentAction = "gridadd" Or Query2.CurrentAction = "gridedit" Then
				Call ListOptions.HideAllOptions(Array())
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			End If

			' Hide export options
			If Query2.Export <> "" Or Query2.CurrentAction <> "" Then
				ExportOptions.HideAllOptions(Array())
			End If

			' Hide other options
			If Query2.Export <> "" Then
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
			Call Query2.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If Query2.RecordsPerPage <> "" Then
			DisplayRecs = Query2.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Load search default if no existing search criteria
		If Not CheckSearchParms() Then

			' Load basic search from default
			Query2.BasicSearch.Keyword = Query2.BasicSearch.KeywordDefault
			Query2.BasicSearch.SearchType = Query2.BasicSearch.SearchTypeDefault
			Query2.BasicSearch.setSearchType(Query2.BasicSearch.SearchTypeDefault)
			If Query2.BasicSearch.Keyword <> "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call Query2.Recordset_Searching(SearchWhere)

		' Save search criteria
		If Command = "search" And Not RestoreSearch Then
			Query2.SearchWhere = SearchWhere ' Save to Session
			StartRec = 1 ' Reset start record counter
			Query2.StartRecordNumber = StartRec
		Else
			SearchWhere = Query2.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		Query2.SessionWhere = sFilter
		Query2.CurrentFilter = ""
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
				sFilter = Query2.KeyFilter
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
		If UBound(arrKeyFlds) >= -1 Then
		End If
		SetupKeyValues = True
	End Function

	' -----------------------------------------------------------------
	' Return Basic Search sql
	'
	Function BasicSearchSQL(Keyword)
		Dim sWhere
		sWhere = ""
			Call BuildBasicSearchSQL(sWhere, Query2.Expr1, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr2, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr3, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr4, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr5, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr6, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr7, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr8, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr9, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr10, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr11, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr12, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr13, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr14, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr15, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr16, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr23, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr25, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr26, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr27, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr28, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr29, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr30, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr31, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr32, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr34, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr36, Keyword)
			Call BuildBasicSearchSQL(sWhere, Query2.Expr37, Keyword)
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
		sSearchKeyword = Query2.BasicSearch.Keyword
		sSearchType = Query2.BasicSearch.SearchType
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
			Query2.BasicSearch.setKeyword(sSearchKeyword)
			Query2.BasicSearch.setSearchType(sSearchType)
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' Check if search parm exists
	Function CheckSearchParms()

		' Check basic search
		If Query2.BasicSearch.IssetSession() Then
			CheckSearchParms = True
			Exit Function
		End If
		CheckSearchParms = False
	End Function

	' Clear all search parameters
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		Query2.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' Load advanced search default values
	Function LoadAdvancedSearchDefault()
		LoadAdvancedSearchDefault = False
	End Function

	' Clear all basic search parameters
	Sub ResetBasicSearchParms()
		Query2.BasicSearch.UnsetSession()
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()

		' Restore search flag
		RestoreSearch = True

		' Restore basic search values
		Call Query2.BasicSearch.Load()
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
			Query2.CurrentOrder = Request.QueryString("order")
			Query2.CurrentOrderType = Request.QueryString("ordertype")

			' Field Expr1
			Call Query2.UpdateSort(Query2.Expr1)

			' Field Expr2
			Call Query2.UpdateSort(Query2.Expr2)

			' Field Expr3
			Call Query2.UpdateSort(Query2.Expr3)

			' Field Expr4
			Call Query2.UpdateSort(Query2.Expr4)

			' Field Expr5
			Call Query2.UpdateSort(Query2.Expr5)

			' Field Expr6
			Call Query2.UpdateSort(Query2.Expr6)

			' Field Expr7
			Call Query2.UpdateSort(Query2.Expr7)

			' Field Expr8
			Call Query2.UpdateSort(Query2.Expr8)

			' Field Expr9
			Call Query2.UpdateSort(Query2.Expr9)

			' Field Expr10
			Call Query2.UpdateSort(Query2.Expr10)

			' Field Expr11
			Call Query2.UpdateSort(Query2.Expr11)

			' Field Expr12
			Call Query2.UpdateSort(Query2.Expr12)

			' Field Expr13
			Call Query2.UpdateSort(Query2.Expr13)

			' Field Expr14
			Call Query2.UpdateSort(Query2.Expr14)

			' Field Expr15
			Call Query2.UpdateSort(Query2.Expr15)

			' Field Expr16
			Call Query2.UpdateSort(Query2.Expr16)

			' Field Expr17
			Call Query2.UpdateSort(Query2.Expr17)

			' Field Expr18
			Call Query2.UpdateSort(Query2.Expr18)

			' Field Expr19
			Call Query2.UpdateSort(Query2.Expr19)

			' Field Expr20
			Call Query2.UpdateSort(Query2.Expr20)

			' Field Expr21
			Call Query2.UpdateSort(Query2.Expr21)

			' Field Expr22
			Call Query2.UpdateSort(Query2.Expr22)

			' Field Expr23
			Call Query2.UpdateSort(Query2.Expr23)

			' Field Expr24
			Call Query2.UpdateSort(Query2.Expr24)

			' Field Expr25
			Call Query2.UpdateSort(Query2.Expr25)

			' Field Expr26
			Call Query2.UpdateSort(Query2.Expr26)

			' Field Expr27
			Call Query2.UpdateSort(Query2.Expr27)

			' Field Expr28
			Call Query2.UpdateSort(Query2.Expr28)

			' Field Expr29
			Call Query2.UpdateSort(Query2.Expr29)

			' Field Expr30
			Call Query2.UpdateSort(Query2.Expr30)

			' Field Expr31
			Call Query2.UpdateSort(Query2.Expr31)

			' Field Expr32
			Call Query2.UpdateSort(Query2.Expr32)

			' Field Expr33
			Call Query2.UpdateSort(Query2.Expr33)

			' Field Expr34
			Call Query2.UpdateSort(Query2.Expr34)

			' Field Expr35
			Call Query2.UpdateSort(Query2.Expr35)

			' Field Expr36
			Call Query2.UpdateSort(Query2.Expr36)

			' Field Expr37
			Call Query2.UpdateSort(Query2.Expr37)
			Query2.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = Query2.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If Query2.SqlOrderBy <> "" Then
				sOrderBy = Query2.SqlOrderBy
				Query2.SessionOrderBy = sOrderBy
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
				Query2.SessionOrderBy = sOrderBy
				Query2.Expr1.Sort = ""
				Query2.Expr2.Sort = ""
				Query2.Expr3.Sort = ""
				Query2.Expr4.Sort = ""
				Query2.Expr5.Sort = ""
				Query2.Expr6.Sort = ""
				Query2.Expr7.Sort = ""
				Query2.Expr8.Sort = ""
				Query2.Expr9.Sort = ""
				Query2.Expr10.Sort = ""
				Query2.Expr11.Sort = ""
				Query2.Expr12.Sort = ""
				Query2.Expr13.Sort = ""
				Query2.Expr14.Sort = ""
				Query2.Expr15.Sort = ""
				Query2.Expr16.Sort = ""
				Query2.Expr17.Sort = ""
				Query2.Expr18.Sort = ""
				Query2.Expr19.Sort = ""
				Query2.Expr20.Sort = ""
				Query2.Expr21.Sort = ""
				Query2.Expr22.Sort = ""
				Query2.Expr23.Sort = ""
				Query2.Expr24.Sort = ""
				Query2.Expr25.Sort = ""
				Query2.Expr26.Sort = ""
				Query2.Expr27.Sort = ""
				Query2.Expr28.Sort = ""
				Query2.Expr29.Sort = ""
				Query2.Expr30.Sort = ""
				Query2.Expr31.Sort = ""
				Query2.Expr32.Sort = ""
				Query2.Expr33.Sort = ""
				Query2.Expr34.Sort = ""
				Query2.Expr35.Sort = ""
				Query2.Expr36.Sort = ""
				Query2.Expr37.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			Query2.StartRecordNumber = StartRec
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
		Call RenderListOptionsExt()
		Call ListOptions_Rendered()
	End Sub

	' Set up other options
	Sub SetupOtherOptions()
		Dim opt, item, DetailTableLink, ar, i
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
			For i = 0 to UBound(Query2.CustomActions.CustomArray)
				Action = Query2.CustomActions.CustomArray(i)(0)
				Name = Query2.CustomActions.CustomArray(i)(1)

				' Add custom action
				Call opt.Add("custom_" & Action)
				Set item = opt.GetItem("custom_" & Action)
				item.Body = "<a class=""ewAction ewCustomAction"" href="""" onclick=""ew_SubmitSelected(document.fQuery2list, '" & ew_CurrentUrl & "', null, '" & Action & "');return false;"">" & Name & "</a>"
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
		sFilter = Query2.GetKeyFilter
		UserAction = Request.Form("useraction") & ""
		Processed = False
		If sFilter <> "" And UserAction <> "" Then
			Query2.CurrentFilter = sFilter
			sSql = Query2.SQL
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
				ElseIf Query2.CancelMessage <> "" Then
					FailureMessage = Query2.CancelMessage
					Query2.CancelMessage = ""
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
				Query2.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					Query2.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = Query2.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			Query2.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			Query2.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			Query2.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		Query2.BasicSearch.Keyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)&""
		If Query2.BasicSearch.Keyword <> "" Then Command = "search"
		Query2.BasicSearch.SearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)&""
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = Query2.CurrentFilter
		Call Query2.Recordset_Selecting(sFilter)
		Query2.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = Query2.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call Query2.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = Query2.KeyFilter

		' Call Row Selecting event
		Call Query2.Row_Selecting(sFilter)

		' Load sql based on filter
		Query2.CurrentFilter = sFilter
		sSql = Query2.SQL
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
		Call Query2.Row_Selected(RsRow)
		Query2.Expr1.DbValue = RsRow("Expr1")
		Query2.Expr2.DbValue = RsRow("Expr2")
		Query2.Expr3.DbValue = RsRow("Expr3")
		Query2.Expr4.DbValue = RsRow("Expr4")
		Query2.Expr5.DbValue = RsRow("Expr5")
		Query2.Expr6.DbValue = RsRow("Expr6")
		Query2.Expr7.DbValue = RsRow("Expr7")
		Query2.Expr8.DbValue = RsRow("Expr8")
		Query2.Expr9.DbValue = RsRow("Expr9")
		Query2.Expr10.DbValue = RsRow("Expr10")
		Query2.Expr11.DbValue = RsRow("Expr11")
		Query2.Expr12.DbValue = RsRow("Expr12")
		Query2.Expr13.DbValue = RsRow("Expr13")
		Query2.Expr14.DbValue = RsRow("Expr14")
		Query2.Expr15.DbValue = RsRow("Expr15")
		Query2.Expr16.DbValue = RsRow("Expr16")
		Query2.Expr17.DbValue = RsRow("Expr17")
		Query2.Expr18.DbValue = RsRow("Expr18")
		Query2.Expr19.DbValue = RsRow("Expr19")
		Query2.Expr20.DbValue = RsRow("Expr20")
		Query2.Expr21.DbValue = RsRow("Expr21")
		Query2.Expr22.DbValue = RsRow("Expr22")
		Query2.Expr23.DbValue = RsRow("Expr23")
		Query2.Expr24.DbValue = RsRow("Expr24")
		Query2.Expr25.DbValue = RsRow("Expr25")
		Query2.Expr26.DbValue = RsRow("Expr26")
		Query2.Expr27.DbValue = RsRow("Expr27")
		Query2.Expr28.DbValue = RsRow("Expr28")
		Query2.Expr29.DbValue = RsRow("Expr29")
		Query2.Expr30.DbValue = RsRow("Expr30")
		Query2.Expr31.DbValue = RsRow("Expr31")
		Query2.Expr32.DbValue = RsRow("Expr32")
		Query2.Expr33.DbValue = RsRow("Expr33")
		Query2.Expr34.DbValue = RsRow("Expr34")
		Query2.Expr35.DbValue = RsRow("Expr35")
		Query2.Expr36.DbValue = RsRow("Expr36")
		Query2.Expr37.DbValue = RsRow("Expr37")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		Query2.Expr1.m_DbValue = Rs("Expr1")
		Query2.Expr2.m_DbValue = Rs("Expr2")
		Query2.Expr3.m_DbValue = Rs("Expr3")
		Query2.Expr4.m_DbValue = Rs("Expr4")
		Query2.Expr5.m_DbValue = Rs("Expr5")
		Query2.Expr6.m_DbValue = Rs("Expr6")
		Query2.Expr7.m_DbValue = Rs("Expr7")
		Query2.Expr8.m_DbValue = Rs("Expr8")
		Query2.Expr9.m_DbValue = Rs("Expr9")
		Query2.Expr10.m_DbValue = Rs("Expr10")
		Query2.Expr11.m_DbValue = Rs("Expr11")
		Query2.Expr12.m_DbValue = Rs("Expr12")
		Query2.Expr13.m_DbValue = Rs("Expr13")
		Query2.Expr14.m_DbValue = Rs("Expr14")
		Query2.Expr15.m_DbValue = Rs("Expr15")
		Query2.Expr16.m_DbValue = Rs("Expr16")
		Query2.Expr17.m_DbValue = Rs("Expr17")
		Query2.Expr18.m_DbValue = Rs("Expr18")
		Query2.Expr19.m_DbValue = Rs("Expr19")
		Query2.Expr20.m_DbValue = Rs("Expr20")
		Query2.Expr21.m_DbValue = Rs("Expr21")
		Query2.Expr22.m_DbValue = Rs("Expr22")
		Query2.Expr23.m_DbValue = Rs("Expr23")
		Query2.Expr24.m_DbValue = Rs("Expr24")
		Query2.Expr25.m_DbValue = Rs("Expr25")
		Query2.Expr26.m_DbValue = Rs("Expr26")
		Query2.Expr27.m_DbValue = Rs("Expr27")
		Query2.Expr28.m_DbValue = Rs("Expr28")
		Query2.Expr29.m_DbValue = Rs("Expr29")
		Query2.Expr30.m_DbValue = Rs("Expr30")
		Query2.Expr31.m_DbValue = Rs("Expr31")
		Query2.Expr32.m_DbValue = Rs("Expr32")
		Query2.Expr33.m_DbValue = Rs("Expr33")
		Query2.Expr34.m_DbValue = Rs("Expr34")
		Query2.Expr35.m_DbValue = Rs("Expr35")
		Query2.Expr36.m_DbValue = Rs("Expr36")
		Query2.Expr37.m_DbValue = Rs("Expr37")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True

		' Load old recordset
		If bValidKey Then
			Query2.CurrentFilter = Query2.KeyFilter
			Dim sSql
			sSql = Query2.SQL
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
		ViewUrl = Query2.ViewUrl("")
		EditUrl = Query2.EditUrl("")
		InlineEditUrl = Query2.InlineEditUrl
		CopyUrl = Query2.CopyUrl("")
		InlineCopyUrl = Query2.InlineCopyUrl
		DeleteUrl = Query2.DeleteUrl

		' Call Row Rendering event
		Call Query2.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' Expr1
		' Expr2
		' Expr3
		' Expr4
		' Expr5
		' Expr6
		' Expr7
		' Expr8
		' Expr9
		' Expr10
		' Expr11
		' Expr12
		' Expr13
		' Expr14
		' Expr15
		' Expr16
		' Expr17
		' Expr18
		' Expr19
		' Expr20
		' Expr21
		' Expr22
		' Expr23
		' Expr24
		' Expr25
		' Expr26
		' Expr27
		' Expr28
		' Expr29
		' Expr30
		' Expr31
		' Expr32
		' Expr33
		' Expr34
		' Expr35
		' Expr36
		' Expr37
		' -----------
		'  View  Row
		' -----------

		If Query2.RowType = EW_ROWTYPE_VIEW Then ' View row

			' Expr1
			Query2.Expr1.ViewValue = Query2.Expr1.CurrentValue
			Query2.Expr1.ViewCustomAttributes = ""

			' Expr2
			Query2.Expr2.ViewValue = Query2.Expr2.CurrentValue
			Query2.Expr2.ViewCustomAttributes = ""

			' Expr3
			Query2.Expr3.ViewValue = Query2.Expr3.CurrentValue
			Query2.Expr3.ViewCustomAttributes = ""

			' Expr4
			Query2.Expr4.ViewValue = Query2.Expr4.CurrentValue
			Query2.Expr4.ViewCustomAttributes = ""

			' Expr5
			Query2.Expr5.ViewValue = Query2.Expr5.CurrentValue
			Query2.Expr5.ViewCustomAttributes = ""

			' Expr6
			Query2.Expr6.ViewValue = Query2.Expr6.CurrentValue
			Query2.Expr6.ViewCustomAttributes = ""

			' Expr7
			Query2.Expr7.ViewValue = Query2.Expr7.CurrentValue
			Query2.Expr7.ViewCustomAttributes = ""

			' Expr8
			Query2.Expr8.ViewValue = Query2.Expr8.CurrentValue
			Query2.Expr8.ViewCustomAttributes = ""

			' Expr9
			Query2.Expr9.ViewValue = Query2.Expr9.CurrentValue
			Query2.Expr9.ViewCustomAttributes = ""

			' Expr10
			Query2.Expr10.ViewValue = Query2.Expr10.CurrentValue
			Query2.Expr10.ViewCustomAttributes = ""

			' Expr11
			Query2.Expr11.ViewValue = Query2.Expr11.CurrentValue
			Query2.Expr11.ViewCustomAttributes = ""

			' Expr12
			Query2.Expr12.ViewValue = Query2.Expr12.CurrentValue
			Query2.Expr12.ViewCustomAttributes = ""

			' Expr13
			Query2.Expr13.ViewValue = Query2.Expr13.CurrentValue
			Query2.Expr13.ViewCustomAttributes = ""

			' Expr14
			Query2.Expr14.ViewValue = Query2.Expr14.CurrentValue
			Query2.Expr14.ViewCustomAttributes = ""

			' Expr15
			Query2.Expr15.ViewValue = Query2.Expr15.CurrentValue
			Query2.Expr15.ViewCustomAttributes = ""

			' Expr16
			Query2.Expr16.ViewValue = Query2.Expr16.CurrentValue
			Query2.Expr16.ViewCustomAttributes = ""

			' Expr17
			Query2.Expr17.ViewValue = Query2.Expr17.CurrentValue
			Query2.Expr17.ViewCustomAttributes = ""

			' Expr18
			Query2.Expr18.ViewValue = Query2.Expr18.CurrentValue
			Query2.Expr18.ViewCustomAttributes = ""

			' Expr19
			Query2.Expr19.ViewValue = Query2.Expr19.CurrentValue
			Query2.Expr19.ViewCustomAttributes = ""

			' Expr20
			Query2.Expr20.ViewValue = Query2.Expr20.CurrentValue
			Query2.Expr20.ViewCustomAttributes = ""

			' Expr21
			Query2.Expr21.ViewValue = Query2.Expr21.CurrentValue
			Query2.Expr21.ViewCustomAttributes = ""

			' Expr22
			Query2.Expr22.ViewValue = Query2.Expr22.CurrentValue
			Query2.Expr22.ViewCustomAttributes = ""

			' Expr23
			Query2.Expr23.ViewValue = Query2.Expr23.CurrentValue
			Query2.Expr23.ViewCustomAttributes = ""

			' Expr24
			Query2.Expr24.ViewValue = Query2.Expr24.CurrentValue
			Query2.Expr24.ViewCustomAttributes = ""

			' Expr25
			Query2.Expr25.ViewValue = Query2.Expr25.CurrentValue
			Query2.Expr25.ViewCustomAttributes = ""

			' Expr26
			Query2.Expr26.ViewValue = Query2.Expr26.CurrentValue
			Query2.Expr26.ViewCustomAttributes = ""

			' Expr27
			Query2.Expr27.ViewValue = Query2.Expr27.CurrentValue
			Query2.Expr27.ViewCustomAttributes = ""

			' Expr28
			Query2.Expr28.ViewValue = Query2.Expr28.CurrentValue
			Query2.Expr28.ViewCustomAttributes = ""

			' Expr29
			Query2.Expr29.ViewValue = Query2.Expr29.CurrentValue
			Query2.Expr29.ViewCustomAttributes = ""

			' Expr30
			Query2.Expr30.ViewValue = Query2.Expr30.CurrentValue
			Query2.Expr30.ViewCustomAttributes = ""

			' Expr31
			Query2.Expr31.ViewValue = Query2.Expr31.CurrentValue
			Query2.Expr31.ViewCustomAttributes = ""

			' Expr32
			Query2.Expr32.ViewValue = Query2.Expr32.CurrentValue
			Query2.Expr32.ViewCustomAttributes = ""

			' Expr33
			Query2.Expr33.ViewValue = Query2.Expr33.CurrentValue
			Query2.Expr33.ViewCustomAttributes = ""

			' Expr34
			Query2.Expr34.ViewValue = Query2.Expr34.CurrentValue
			Query2.Expr34.ViewCustomAttributes = ""

			' Expr35
			Query2.Expr35.ViewValue = Query2.Expr35.CurrentValue
			Query2.Expr35.ViewCustomAttributes = ""

			' Expr36
			Query2.Expr36.ViewValue = Query2.Expr36.CurrentValue
			Query2.Expr36.ViewCustomAttributes = ""

			' Expr37
			Query2.Expr37.ViewValue = Query2.Expr37.CurrentValue
			Query2.Expr37.ViewCustomAttributes = ""

			' View refer script
			' Expr1

			Query2.Expr1.LinkCustomAttributes = ""
			Query2.Expr1.HrefValue = ""
			Query2.Expr1.TooltipValue = ""

			' Expr2
			Query2.Expr2.LinkCustomAttributes = ""
			Query2.Expr2.HrefValue = ""
			Query2.Expr2.TooltipValue = ""

			' Expr3
			Query2.Expr3.LinkCustomAttributes = ""
			Query2.Expr3.HrefValue = ""
			Query2.Expr3.TooltipValue = ""

			' Expr4
			Query2.Expr4.LinkCustomAttributes = ""
			Query2.Expr4.HrefValue = ""
			Query2.Expr4.TooltipValue = ""

			' Expr5
			Query2.Expr5.LinkCustomAttributes = ""
			Query2.Expr5.HrefValue = ""
			Query2.Expr5.TooltipValue = ""

			' Expr6
			Query2.Expr6.LinkCustomAttributes = ""
			Query2.Expr6.HrefValue = ""
			Query2.Expr6.TooltipValue = ""

			' Expr7
			Query2.Expr7.LinkCustomAttributes = ""
			Query2.Expr7.HrefValue = ""
			Query2.Expr7.TooltipValue = ""

			' Expr8
			Query2.Expr8.LinkCustomAttributes = ""
			Query2.Expr8.HrefValue = ""
			Query2.Expr8.TooltipValue = ""

			' Expr9
			Query2.Expr9.LinkCustomAttributes = ""
			Query2.Expr9.HrefValue = ""
			Query2.Expr9.TooltipValue = ""

			' Expr10
			Query2.Expr10.LinkCustomAttributes = ""
			Query2.Expr10.HrefValue = ""
			Query2.Expr10.TooltipValue = ""

			' Expr11
			Query2.Expr11.LinkCustomAttributes = ""
			Query2.Expr11.HrefValue = ""
			Query2.Expr11.TooltipValue = ""

			' Expr12
			Query2.Expr12.LinkCustomAttributes = ""
			Query2.Expr12.HrefValue = ""
			Query2.Expr12.TooltipValue = ""

			' Expr13
			Query2.Expr13.LinkCustomAttributes = ""
			Query2.Expr13.HrefValue = ""
			Query2.Expr13.TooltipValue = ""

			' Expr14
			Query2.Expr14.LinkCustomAttributes = ""
			Query2.Expr14.HrefValue = ""
			Query2.Expr14.TooltipValue = ""

			' Expr15
			Query2.Expr15.LinkCustomAttributes = ""
			Query2.Expr15.HrefValue = ""
			Query2.Expr15.TooltipValue = ""

			' Expr16
			Query2.Expr16.LinkCustomAttributes = ""
			Query2.Expr16.HrefValue = ""
			Query2.Expr16.TooltipValue = ""

			' Expr17
			Query2.Expr17.LinkCustomAttributes = ""
			Query2.Expr17.HrefValue = ""
			Query2.Expr17.TooltipValue = ""

			' Expr18
			Query2.Expr18.LinkCustomAttributes = ""
			Query2.Expr18.HrefValue = ""
			Query2.Expr18.TooltipValue = ""

			' Expr19
			Query2.Expr19.LinkCustomAttributes = ""
			Query2.Expr19.HrefValue = ""
			Query2.Expr19.TooltipValue = ""

			' Expr20
			Query2.Expr20.LinkCustomAttributes = ""
			Query2.Expr20.HrefValue = ""
			Query2.Expr20.TooltipValue = ""

			' Expr21
			Query2.Expr21.LinkCustomAttributes = ""
			Query2.Expr21.HrefValue = ""
			Query2.Expr21.TooltipValue = ""

			' Expr22
			Query2.Expr22.LinkCustomAttributes = ""
			Query2.Expr22.HrefValue = ""
			Query2.Expr22.TooltipValue = ""

			' Expr23
			Query2.Expr23.LinkCustomAttributes = ""
			Query2.Expr23.HrefValue = ""
			Query2.Expr23.TooltipValue = ""

			' Expr24
			Query2.Expr24.LinkCustomAttributes = ""
			Query2.Expr24.HrefValue = ""
			Query2.Expr24.TooltipValue = ""

			' Expr25
			Query2.Expr25.LinkCustomAttributes = ""
			Query2.Expr25.HrefValue = ""
			Query2.Expr25.TooltipValue = ""

			' Expr26
			Query2.Expr26.LinkCustomAttributes = ""
			Query2.Expr26.HrefValue = ""
			Query2.Expr26.TooltipValue = ""

			' Expr27
			Query2.Expr27.LinkCustomAttributes = ""
			Query2.Expr27.HrefValue = ""
			Query2.Expr27.TooltipValue = ""

			' Expr28
			Query2.Expr28.LinkCustomAttributes = ""
			Query2.Expr28.HrefValue = ""
			Query2.Expr28.TooltipValue = ""

			' Expr29
			Query2.Expr29.LinkCustomAttributes = ""
			Query2.Expr29.HrefValue = ""
			Query2.Expr29.TooltipValue = ""

			' Expr30
			Query2.Expr30.LinkCustomAttributes = ""
			Query2.Expr30.HrefValue = ""
			Query2.Expr30.TooltipValue = ""

			' Expr31
			Query2.Expr31.LinkCustomAttributes = ""
			Query2.Expr31.HrefValue = ""
			Query2.Expr31.TooltipValue = ""

			' Expr32
			Query2.Expr32.LinkCustomAttributes = ""
			Query2.Expr32.HrefValue = ""
			Query2.Expr32.TooltipValue = ""

			' Expr33
			Query2.Expr33.LinkCustomAttributes = ""
			Query2.Expr33.HrefValue = ""
			Query2.Expr33.TooltipValue = ""

			' Expr34
			Query2.Expr34.LinkCustomAttributes = ""
			Query2.Expr34.HrefValue = ""
			Query2.Expr34.TooltipValue = ""

			' Expr35
			Query2.Expr35.LinkCustomAttributes = ""
			Query2.Expr35.HrefValue = ""
			Query2.Expr35.TooltipValue = ""

			' Expr36
			Query2.Expr36.LinkCustomAttributes = ""
			Query2.Expr36.HrefValue = ""
			Query2.Expr36.TooltipValue = ""

			' Expr37
			Query2.Expr37.LinkCustomAttributes = ""
			Query2.Expr37.HrefValue = ""
			Query2.Expr37.TooltipValue = ""
		End If

		' Call Row Rendered event
		If Query2.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Query2.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		url = ew_CurrentUrl
		url = ew_RegExReplace("\?cmd=reset(all){0,1}$", url, "") ' Remove cmd=reset / cmd=resetall
		Call Breadcrumb.Add("list", Query2.TableVar, url, Query2.TableVar, True)
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
