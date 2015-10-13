<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_z40newsinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim z40news_list
Set z40news_list = New cz40news_list
Set Page = z40news_list

' Page init processing
z40news_list.Page_Init()

' Page main processing
z40news_list.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
z40news_list.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If z40news.Export = "" Then %>
<script type="text/javascript">
// Page object
var z40news_list = new ew_Page("z40news_list");
z40news_list.PageID = "list"; // Page ID
var EW_PAGE_ID = z40news_list.PageID; // For backward compatibility
// Form object
var fz40newslist = new ew_Form("fz40newslist");
fz40newslist.FormKeyCountName = '<%= z40news_list.FormKeyCountName %>';
// Form_CustomValidate event
fz40newslist.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fz40newslist.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fz40newslist.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
var fz40newslistsrch = new ew_Form("fz40newslistsrch");
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If z40news.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If z40news_list.ExportOptions.Visible Then %>
<div class="ewListExportOptions"><% z40news_list.ExportOptions.Render "body", "", "", "", "", "" %></div>
<% End If %>
<% If (z40news.Export = "") Or (EW_EXPORT_MASTER_RECORD And z40news.Export = "print") Then %>
<% End If %>
<%

' Load recordset
Set z40news_list.Recordset = z40news_list.LoadRecordset()
	z40news_list.TotalRecs = z40news_list.Recordset.RecordCount
	z40news_list.StartRec = 1
	If z40news_list.DisplayRecs <= 0 Then ' Display all records
		z40news_list.DisplayRecs = z40news_list.TotalRecs
	End If
	If Not (z40news.ExportAll And z40news.Export <> "") Then
		z40news_list.SetUpStartRec() ' Set up start record position
	End If
z40news_list.RenderOtherOptions()
%>
<% If Security.IsLoggedIn() Then %>
<% If z40news.Export = "" And z40news.CurrentAction = "" Then %>
<form name="fz40newslistsrch" id="fz40newslistsrch" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewSearchTable"><tr><td>
<div class="accordion" id="fz40newslistsrch_SearchGroup">
	<div class="accordion-group">
		<div class="accordion-heading">
<a class="accordion-toggle" data-toggle="collapse" data-parent="#fz40newslistsrch_SearchGroup" href="#fz40newslistsrch_SearchBody"><%= Language.Phrase("Search") %></a>
		</div>
		<div id="fz40newslistsrch_SearchBody" class="accordion-body collapse in">
			<div class="accordion-inner">
<div id="fz40newslistsrch_SearchPanel">
<input type="hidden" name="cmd" value="search">
<input type="hidden" name="t" value="z40news">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewRow">
	<div class="btn-group ewButtonGroup">
	<div class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" class="input-large" value="<%= ew_HtmlEncode(z40news.BasicSearch.getKeyword()) %>" placeholder="<%= ew_HtmlEncode(Language.Phrase("Search")) %>">
	<button class="btn btn-primary ewButton" name="btnsubmit" id="btnsubmit" type="submit"><%= Language.Phrase("QuickSearchBtn") %></button>
	</div>
	</div>
	<div class="btn-group ewButtonGroup">
	<a class="btn ewShowAll" href="<%= z40news_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>
	</div>
</div>
<div id="xsr_2" class="ewRow">
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="="<% If z40news.BasicSearch.getSearchType = "=" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If z40news.BasicSearch.getSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If z40news.BasicSearch.getSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
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
<% z40news_list.ShowPageHeader() %>
<% z40news_list.ShowMessage %>
<table class="ewGrid"><tr><td class="ewGridContent">
<form name="fz40newslist" id="fz40newslist" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="z40news">
<div id="gmp_z40news" class="ewGridMiddlePanel">
<% If z40news_list.TotalRecs > 0 Then %>
<table id="tbl_z40newslist" class="ewTable ewTableSeparate">
<%= z40news.TableCustomInnerHtml %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call z40news_list.RenderListOptions()

' Render list options (header, left)
z40news_list.ListOptions.Render "header", "left", "", "", "", ""
%>
<% If z40news.news_id.Visible Then ' news_id %>
	<% If z40news.SortUrl(z40news.news_id) = "" Then %>
		<td><div id="elh_z40news_news_id" class="z40news_news_id"><div class="ewTableHeaderCaption"><%= z40news.news_id.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= z40news.SortUrl(z40news.news_id) %>',1);"><div id="elh_z40news_news_id" class="z40news_news_id">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= z40news.news_id.FldCaption %></span><span class="ewTableHeaderSort"><% If z40news.news_id.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf z40news.news_id.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If z40news.news_img.Visible Then ' news_img %>
	<% If z40news.SortUrl(z40news.news_img) = "" Then %>
		<td><div id="elh_z40news_news_img" class="z40news_news_img"><div class="ewTableHeaderCaption"><%= z40news.news_img.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= z40news.SortUrl(z40news.news_img) %>',1);"><div id="elh_z40news_news_img" class="z40news_news_img">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= z40news.news_img.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If z40news.news_img.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf z40news.news_img.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If z40news.news_category.Visible Then ' news_category %>
	<% If z40news.SortUrl(z40news.news_category) = "" Then %>
		<td><div id="elh_z40news_news_category" class="z40news_news_category"><div class="ewTableHeaderCaption"><%= z40news.news_category.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= z40news.SortUrl(z40news.news_category) %>',1);"><div id="elh_z40news_news_category" class="z40news_news_category">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= z40news.news_category.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If z40news.news_category.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf z40news.news_category.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If z40news.news_subject.Visible Then ' news_subject %>
	<% If z40news.SortUrl(z40news.news_subject) = "" Then %>
		<td><div id="elh_z40news_news_subject" class="z40news_news_subject"><div class="ewTableHeaderCaption"><%= z40news.news_subject.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= z40news.SortUrl(z40news.news_subject) %>',1);"><div id="elh_z40news_news_subject" class="z40news_news_subject">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= z40news.news_subject.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If z40news.news_subject.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf z40news.news_subject.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If z40news.news_subject_th.Visible Then ' news_subject_th %>
	<% If z40news.SortUrl(z40news.news_subject_th) = "" Then %>
		<td><div id="elh_z40news_news_subject_th" class="z40news_news_subject_th"><div class="ewTableHeaderCaption"><%= z40news.news_subject_th.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= z40news.SortUrl(z40news.news_subject_th) %>',1);"><div id="elh_z40news_news_subject_th" class="z40news_news_subject_th">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= z40news.news_subject_th.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If z40news.news_subject_th.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf z40news.news_subject_th.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If z40news.news_show_en.Visible Then ' news_show_en %>
	<% If z40news.SortUrl(z40news.news_show_en) = "" Then %>
		<td><div id="elh_z40news_news_show_en" class="z40news_news_show_en"><div class="ewTableHeaderCaption"><%= z40news.news_show_en.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= z40news.SortUrl(z40news.news_show_en) %>',1);"><div id="elh_z40news_news_show_en" class="z40news_news_show_en">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= z40news.news_show_en.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If z40news.news_show_en.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf z40news.news_show_en.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If z40news.news_show.Visible Then ' news_show %>
	<% If z40news.SortUrl(z40news.news_show) = "" Then %>
		<td><div id="elh_z40news_news_show" class="z40news_news_show"><div class="ewTableHeaderCaption"><%= z40news.news_show.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= z40news.SortUrl(z40news.news_show) %>',1);"><div id="elh_z40news_news_show" class="z40news_news_show">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= z40news.news_show.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If z40news.news_show.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf z40news.news_show.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If z40news.news_create.Visible Then ' news_create %>
	<% If z40news.SortUrl(z40news.news_create) = "" Then %>
		<td><div id="elh_z40news_news_create" class="z40news_news_create"><div class="ewTableHeaderCaption"><%= z40news.news_create.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= z40news.SortUrl(z40news.news_create) %>',1);"><div id="elh_z40news_news_create" class="z40news_news_create">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= z40news.news_create.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If z40news.news_create.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf z40news.news_create.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If z40news.news_update.Visible Then ' news_update %>
	<% If z40news.SortUrl(z40news.news_update) = "" Then %>
		<td><div id="elh_z40news_news_update" class="z40news_news_update"><div class="ewTableHeaderCaption"><%= z40news.news_update.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= z40news.SortUrl(z40news.news_update) %>',1);"><div id="elh_z40news_news_update" class="z40news_news_update">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= z40news.news_update.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If z40news.news_update.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf z40news.news_update.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
z40news_list.ListOptions.Render "header", "right", "", "", "", ""
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (z40news.ExportAll And z40news.Export <> "") Then
	z40news_list.StopRec = z40news_list.TotalRecs
Else

	' Set the last record to display
	If z40news_list.TotalRecs > z40news_list.StartRec + z40news_list.DisplayRecs - 1 Then
		z40news_list.StopRec = z40news_list.StartRec + z40news_list.DisplayRecs - 1
	Else
		z40news_list.StopRec = z40news_list.TotalRecs
	End If
End If

' Move to first record
z40news_list.RecCnt = z40news_list.StartRec - 1
If Not z40news_list.Recordset.Eof Then
	z40news_list.Recordset.MoveFirst
	If z40news_list.StartRec > 1 Then z40news_list.Recordset.Move z40news_list.StartRec - 1
ElseIf Not z40news.AllowAddDeleteRow And z40news_list.StopRec = 0 Then
	z40news_list.StopRec = z40news.GridAddRowCount
End If

' Initialize Aggregate
z40news.RowType = EW_ROWTYPE_AGGREGATEINIT
Call z40news.ResetAttrs()
Call z40news_list.RenderRow()
z40news_list.RowCnt = 0

' Output date rows
Do While CLng(z40news_list.RecCnt) < CLng(z40news_list.StopRec)
	z40news_list.RecCnt = z40news_list.RecCnt + 1
	If CLng(z40news_list.RecCnt) >= CLng(z40news_list.StartRec) Then
		z40news_list.RowCnt = z40news_list.RowCnt + 1

	' Set up key count
	z40news_list.KeyCount = z40news_list.RowIndex
	Call z40news.ResetAttrs()
	z40news.CssClass = ""
	If z40news.CurrentAction = "gridadd" Then
	Else
		Call z40news_list.LoadRowValues(z40news_list.Recordset) ' Load row values
	End If
	z40news.RowType = EW_ROWTYPE_VIEW ' Render view

	' Set up row id / data-rowindex
	z40news.RowAttrs.AddAttributes Array(Array("data-rowindex", z40news_list.RowCnt), Array("id", "r" & z40news_list.RowCnt & "_z40news"), Array("data-rowtype", z40news.RowType))

	' Render row
	Call z40news_list.RenderRow()

	' Render list options
	Call z40news_list.RenderListOptions()
%>
	<tr<%= z40news.RowAttributes %>>
<%

' Render list options (body, left)
z40news_list.ListOptions.Render "body", "left", z40news_list.RowCnt, "", "", ""
%>
	<% If z40news.news_id.Visible Then ' news_id %>
		<td<%= z40news.news_id.CellAttributes %>>
<span<%= z40news.news_id.ViewAttributes %>>
<%= z40news.news_id.ListViewValue %>
</span>
<a id="<%= z40news_list.PageObjName & "_row_" & z40news_list.RowCnt %>"></a></td>
	<% End If %>
	<% If z40news.news_img.Visible Then ' news_img %>
		<td<%= z40news.news_img.CellAttributes %>>
<span<%= z40news.news_img.ViewAttributes %>>
<%= z40news.news_img.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If z40news.news_category.Visible Then ' news_category %>
		<td<%= z40news.news_category.CellAttributes %>>
<span<%= z40news.news_category.ViewAttributes %>>
<%= z40news.news_category.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If z40news.news_subject.Visible Then ' news_subject %>
		<td<%= z40news.news_subject.CellAttributes %>>
<span<%= z40news.news_subject.ViewAttributes %>>
<%= z40news.news_subject.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If z40news.news_subject_th.Visible Then ' news_subject_th %>
		<td<%= z40news.news_subject_th.CellAttributes %>>
<span<%= z40news.news_subject_th.ViewAttributes %>>
<%= z40news.news_subject_th.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If z40news.news_show_en.Visible Then ' news_show_en %>
		<td<%= z40news.news_show_en.CellAttributes %>>
<span<%= z40news.news_show_en.ViewAttributes %>>
<%= z40news.news_show_en.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If z40news.news_show.Visible Then ' news_show %>
		<td<%= z40news.news_show.CellAttributes %>>
<span<%= z40news.news_show.ViewAttributes %>>
<%= z40news.news_show.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If z40news.news_create.Visible Then ' news_create %>
		<td<%= z40news.news_create.CellAttributes %>>
<span<%= z40news.news_create.ViewAttributes %>>
<%= z40news.news_create.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If z40news.news_update.Visible Then ' news_update %>
		<td<%= z40news.news_update.CellAttributes %>>
<span<%= z40news.news_update.ViewAttributes %>>
<%= z40news.news_update.ListViewValue %>
</span>
</td>
	<% End If %>
<%

' Render list options (body, right)
z40news_list.ListOptions.Render "body", "right", z40news_list.RowCnt, "", "", ""
%>
	</tr>
<%
	End If
	If z40news.CurrentAction <> "gridadd" Then
		z40news_list.Recordset.MoveNext()
	End If
Loop
%>
</tbody>
</table>
<% End If %>
<% If z40news.CurrentAction = "" Then %>
<input type="hidden" name="a_list" id="a_list" value="">
<% End If %>
</div>
</form>
<%

' Close recordset and connection
z40news_list.Recordset.Close
Set z40news_list.Recordset = Nothing
%>
<% If z40news.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If z40news.CurrentAction <> "gridadd" And z40news.CurrentAction <> "gridedit" Then %>
<form name="ewPagerForm" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewPager">
<tr><td>
<% If Not IsObject(z40news_list.Pager) Then Set z40news_list.Pager = ew_NewPrevNextPager(z40news_list.StartRec, z40news_list.DisplayRecs, z40news_list.TotalRecs) %>
<% If z40news_list.Pager.RecordCount > 0 Then %>
<table class="ewStdTable"><tbody><tr><td>
	<%= Language.Phrase("Page") %>&nbsp;
<div class="input-prepend input-append">
<!--first page button-->
	<% If z40news_list.Pager.FirstButton.Enabled Then %>
	<a class="btn btn-small" href="<%= z40news_list.PageUrl %>start=<%= z40news_list.Pager.FirstButton.Start %>"><i class="icon-step-backward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-backward"></i></a>
	<% End If %>
<!--previous page button-->
	<% If z40news_list.Pager.PrevButton.Enabled Then %>
	<a class="btn btn-small" href="<%= z40news_list.PageUrl %>start=<%= z40news_list.Pager.PrevButton.Start %>"><i class="icon-prev"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-prev"></i></a>
	<% End If %>
<!--current page number-->
	<input class="input-mini" type="text" name="<%= EW_TABLE_PAGE_NO %>" value="<%= z40news_list.Pager.CurrentPage %>">
<!--next page button-->
	<% If z40news_list.Pager.NextButton.Enabled Then %>
	<a class="btn btn-small" href="<%= z40news_list.PageUrl %>start=<%= z40news_list.Pager.NextButton.Start %>"><i class="icon-play"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-play"></i></a>
	<% End If %>
<!--last page button-->
	<% If z40news_list.Pager.LastButton.Enabled Then %>
	<a class="btn btn-small" href="<%= z40news_list.PageUrl %>start=<%= z40news_list.Pager.LastButton.Start %>"><i class="icon-step-forward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-forward"></i></a>
	<% End If %>
</div>
	&nbsp;<%= Language.Phrase("of") %>&nbsp;<%= z40news_list.Pager.PageCount %>
</td>
<td>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<%= Language.Phrase("Record") %>&nbsp;<%= z40news_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= z40news_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= z40news_list.Pager.RecordCount %>
</td>
</tr></tbody></table>
<% Else %>
	<% If z40news_list.SearchWhere = "0=101" Then %>
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
	z40news_list.AddEditOptions.Render "body", "bottom", "", "", "", ""
	z40news_list.DetailOptions.Render "body", "bottom", "", "", "", ""
	z40news_list.ActionOptions.Render "body", "bottom", "", "", "", ""
%>
</div>
</div>
<% End If %>
</td></tr></table>
<% If z40news.Export = "" Then %>
<script type="text/javascript">
fz40newslistsrch.Init();
fz40newslist.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<% End If %>
<%
z40news_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If z40news.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set z40news_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cz40news_list

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
		TableName = "@news"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "z40news_list"
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
		If z40news.UseTokenInUrl Then PageUrl = PageUrl & "t=" & z40news.TableVar & "&" ' add page token
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
		If z40news.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (z40news.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (z40news.TableVar = Request.QueryString("t"))
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
		FormName = "fz40newslist"
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
		If IsEmpty(z40news) Then Set z40news = New cz40news
		Set Table = z40news

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		ExportPdfUrl = PageUrl & "export=pdf"
		AddUrl = "pom_z40newsadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "pom_z40newsdelete.asp"
		MultiUpdateUrl = "pom_z40newsupdate.asp"

		' Initialize other table object
		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "@news"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' List options
		Set ListOptions = New cListOptions
		ListOptions.TableVar = z40news.TableVar

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = z40news.TableVar
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
				z40news.GridAddRowCount = gridaddcnt
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
		If UBound(z40news.CustomActions.CustomArray) >= 0 Then
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
		Set z40news = Nothing
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
			If z40news.Export = "" Then
				SetupBreadcrumb()
			End If

			' Hide list options
			If z40news.Export <> "" Then
				Call ListOptions.HideAllOptions(Array("sequence"))
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			ElseIf z40news.CurrentAction = "gridadd" Or z40news.CurrentAction = "gridedit" Then
				Call ListOptions.HideAllOptions(Array())
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			End If

			' Hide export options
			If z40news.Export <> "" Or z40news.CurrentAction <> "" Then
				ExportOptions.HideAllOptions(Array())
			End If

			' Hide other options
			If z40news.Export <> "" Then
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
			Call z40news.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If z40news.RecordsPerPage <> "" Then
			DisplayRecs = z40news.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Load search default if no existing search criteria
		If Not CheckSearchParms() Then

			' Load basic search from default
			z40news.BasicSearch.Keyword = z40news.BasicSearch.KeywordDefault
			z40news.BasicSearch.SearchType = z40news.BasicSearch.SearchTypeDefault
			z40news.BasicSearch.setSearchType(z40news.BasicSearch.SearchTypeDefault)
			If z40news.BasicSearch.Keyword <> "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call z40news.Recordset_Searching(SearchWhere)

		' Save search criteria
		If Command = "search" And Not RestoreSearch Then
			z40news.SearchWhere = SearchWhere ' Save to Session
			StartRec = 1 ' Reset start record counter
			z40news.StartRecordNumber = StartRec
		Else
			SearchWhere = z40news.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		z40news.SessionWhere = sFilter
		z40news.CurrentFilter = ""
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
				sFilter = z40news.KeyFilter
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
			z40news.news_id.FormValue = arrKeyFlds(0)
			If Not IsNumeric(z40news.news_id.FormValue) Then
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
			Call BuildBasicSearchSQL(sWhere, z40news.news_img, Keyword)
			Call BuildBasicSearchSQL(sWhere, z40news.news_category, Keyword)
			Call BuildBasicSearchSQL(sWhere, z40news.news_subject, Keyword)
			Call BuildBasicSearchSQL(sWhere, z40news.news_subject_th, Keyword)
			Call BuildBasicSearchSQL(sWhere, z40news.news_intro, Keyword)
			Call BuildBasicSearchSQL(sWhere, z40news.news_intro_th, Keyword)
			Call BuildBasicSearchSQL(sWhere, z40news.news_content, Keyword)
			Call BuildBasicSearchSQL(sWhere, z40news.news_content_th, Keyword)
			Call BuildBasicSearchSQL(sWhere, z40news.news_show_en, Keyword)
			Call BuildBasicSearchSQL(sWhere, z40news.news_show, Keyword)
			Call BuildBasicSearchSQL(sWhere, z40news.news_create, Keyword)
			Call BuildBasicSearchSQL(sWhere, z40news.news_update, Keyword)
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
		sSearchKeyword = z40news.BasicSearch.Keyword
		sSearchType = z40news.BasicSearch.SearchType
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
			z40news.BasicSearch.setKeyword(sSearchKeyword)
			z40news.BasicSearch.setSearchType(sSearchType)
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' Check if search parm exists
	Function CheckSearchParms()

		' Check basic search
		If z40news.BasicSearch.IssetSession() Then
			CheckSearchParms = True
			Exit Function
		End If
		CheckSearchParms = False
	End Function

	' Clear all search parameters
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		z40news.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' Load advanced search default values
	Function LoadAdvancedSearchDefault()
		LoadAdvancedSearchDefault = False
	End Function

	' Clear all basic search parameters
	Sub ResetBasicSearchParms()
		z40news.BasicSearch.UnsetSession()
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()

		' Restore search flag
		RestoreSearch = True

		' Restore basic search values
		Call z40news.BasicSearch.Load()
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
			z40news.CurrentOrder = Request.QueryString("order")
			z40news.CurrentOrderType = Request.QueryString("ordertype")

			' Field news_id
			Call z40news.UpdateSort(z40news.news_id)

			' Field news_img
			Call z40news.UpdateSort(z40news.news_img)

			' Field news_category
			Call z40news.UpdateSort(z40news.news_category)

			' Field news_subject
			Call z40news.UpdateSort(z40news.news_subject)

			' Field news_subject_th
			Call z40news.UpdateSort(z40news.news_subject_th)

			' Field news_show_en
			Call z40news.UpdateSort(z40news.news_show_en)

			' Field news_show
			Call z40news.UpdateSort(z40news.news_show)

			' Field news_create
			Call z40news.UpdateSort(z40news.news_create)

			' Field news_update
			Call z40news.UpdateSort(z40news.news_update)
			z40news.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = z40news.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If z40news.SqlOrderBy <> "" Then
				sOrderBy = z40news.SqlOrderBy
				z40news.SessionOrderBy = sOrderBy
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
				z40news.SessionOrderBy = sOrderBy
				z40news.news_id.Sort = ""
				z40news.news_img.Sort = ""
				z40news.news_category.Sort = ""
				z40news.news_subject.Sort = ""
				z40news.news_subject_th.Sort = ""
				z40news.news_show_en.Sort = ""
				z40news.news_show.Sort = ""
				z40news.news_create.Sort = ""
				z40news.news_update.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			z40news.StartRecordNumber = StartRec
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
		ListOptions.GetItem("checkbox").Body = "<label class=""checkbox""><input type=""checkbox"" name=""key_m"" value=""" & ew_HtmlEncode(z40news.news_id.CurrentValue) & """ onclick='ew_ClickMultiCheckbox(event, this);'></label>"
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
			For i = 0 to UBound(z40news.CustomActions.CustomArray)
				Action = z40news.CustomActions.CustomArray(i)(0)
				Name = z40news.CustomActions.CustomArray(i)(1)

				' Add custom action
				Call opt.Add("custom_" & Action)
				Set item = opt.GetItem("custom_" & Action)
				item.Body = "<a class=""ewAction ewCustomAction"" href="""" onclick=""ew_SubmitSelected(document.fz40newslist, '" & ew_CurrentUrl & "', null, '" & Action & "');return false;"">" & Name & "</a>"
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
		sFilter = z40news.GetKeyFilter
		UserAction = Request.Form("useraction") & ""
		Processed = False
		If sFilter <> "" And UserAction <> "" Then
			z40news.CurrentFilter = sFilter
			sSql = z40news.SQL
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
				ElseIf z40news.CancelMessage <> "" Then
					FailureMessage = z40news.CancelMessage
					z40news.CancelMessage = ""
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
				z40news.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					z40news.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = z40news.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			z40news.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			z40news.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			z40news.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		z40news.BasicSearch.Keyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)&""
		If z40news.BasicSearch.Keyword <> "" Then Command = "search"
		z40news.BasicSearch.SearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)&""
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = z40news.CurrentFilter
		Call z40news.Recordset_Selecting(sFilter)
		z40news.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = z40news.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call z40news.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = z40news.KeyFilter

		' Call Row Selecting event
		Call z40news.Row_Selecting(sFilter)

		' Load sql based on filter
		z40news.CurrentFilter = sFilter
		sSql = z40news.SQL
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
		Call z40news.Row_Selected(RsRow)
		z40news.news_id.DbValue = RsRow("news_id")
		z40news.news_img.DbValue = RsRow("news_img")
		z40news.news_category.DbValue = RsRow("news_category")
		z40news.news_subject.DbValue = RsRow("news_subject")
		z40news.news_subject_th.DbValue = RsRow("news_subject_th")
		z40news.news_intro.DbValue = RsRow("news_intro")
		z40news.news_intro_th.DbValue = RsRow("news_intro_th")
		z40news.news_content.DbValue = RsRow("news_content")
		z40news.news_content_th.DbValue = RsRow("news_content_th")
		z40news.news_show_en.DbValue = RsRow("news_show_en")
		z40news.news_show.DbValue = RsRow("news_show")
		z40news.news_create.DbValue = RsRow("news_create")
		z40news.news_update.DbValue = RsRow("news_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		z40news.news_id.m_DbValue = Rs("news_id")
		z40news.news_img.m_DbValue = Rs("news_img")
		z40news.news_category.m_DbValue = Rs("news_category")
		z40news.news_subject.m_DbValue = Rs("news_subject")
		z40news.news_subject_th.m_DbValue = Rs("news_subject_th")
		z40news.news_intro.m_DbValue = Rs("news_intro")
		z40news.news_intro_th.m_DbValue = Rs("news_intro_th")
		z40news.news_content.m_DbValue = Rs("news_content")
		z40news.news_content_th.m_DbValue = Rs("news_content_th")
		z40news.news_show_en.m_DbValue = Rs("news_show_en")
		z40news.news_show.m_DbValue = Rs("news_show")
		z40news.news_create.m_DbValue = Rs("news_create")
		z40news.news_update.m_DbValue = Rs("news_update")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If z40news.GetKey("news_id")&"" <> "" Then
			z40news.news_id.CurrentValue = z40news.GetKey("news_id") ' news_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			z40news.CurrentFilter = z40news.KeyFilter
			Dim sSql
			sSql = z40news.SQL
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
		ViewUrl = z40news.ViewUrl("")
		EditUrl = z40news.EditUrl("")
		InlineEditUrl = z40news.InlineEditUrl
		CopyUrl = z40news.CopyUrl("")
		InlineCopyUrl = z40news.InlineCopyUrl
		DeleteUrl = z40news.DeleteUrl

		' Call Row Rendering event
		Call z40news.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' news_id
		' news_img
		' news_category
		' news_subject
		' news_subject_th
		' news_intro
		' news_intro_th
		' news_content
		' news_content_th
		' news_show_en
		' news_show
		' news_create
		' news_update
		' -----------
		'  View  Row
		' -----------

		If z40news.RowType = EW_ROWTYPE_VIEW Then ' View row

			' news_id
			z40news.news_id.ViewValue = z40news.news_id.CurrentValue
			z40news.news_id.ViewCustomAttributes = ""

			' news_img
			z40news.news_img.ViewValue = z40news.news_img.CurrentValue
			z40news.news_img.ViewCustomAttributes = ""

			' news_category
			z40news.news_category.ViewValue = z40news.news_category.CurrentValue
			z40news.news_category.ViewCustomAttributes = ""

			' news_subject
			z40news.news_subject.ViewValue = z40news.news_subject.CurrentValue
			z40news.news_subject.ViewCustomAttributes = ""

			' news_subject_th
			z40news.news_subject_th.ViewValue = z40news.news_subject_th.CurrentValue
			z40news.news_subject_th.ViewCustomAttributes = ""

			' news_show_en
			z40news.news_show_en.ViewValue = z40news.news_show_en.CurrentValue
			z40news.news_show_en.ViewCustomAttributes = ""

			' news_show
			z40news.news_show.ViewValue = z40news.news_show.CurrentValue
			z40news.news_show.ViewCustomAttributes = ""

			' news_create
			z40news.news_create.ViewValue = z40news.news_create.CurrentValue
			z40news.news_create.ViewCustomAttributes = ""

			' news_update
			z40news.news_update.ViewValue = z40news.news_update.CurrentValue
			z40news.news_update.ViewCustomAttributes = ""

			' View refer script
			' news_id

			z40news.news_id.LinkCustomAttributes = ""
			z40news.news_id.HrefValue = ""
			z40news.news_id.TooltipValue = ""

			' news_img
			z40news.news_img.LinkCustomAttributes = ""
			z40news.news_img.HrefValue = ""
			z40news.news_img.TooltipValue = ""

			' news_category
			z40news.news_category.LinkCustomAttributes = ""
			z40news.news_category.HrefValue = ""
			z40news.news_category.TooltipValue = ""

			' news_subject
			z40news.news_subject.LinkCustomAttributes = ""
			z40news.news_subject.HrefValue = ""
			z40news.news_subject.TooltipValue = ""

			' news_subject_th
			z40news.news_subject_th.LinkCustomAttributes = ""
			z40news.news_subject_th.HrefValue = ""
			z40news.news_subject_th.TooltipValue = ""

			' news_show_en
			z40news.news_show_en.LinkCustomAttributes = ""
			z40news.news_show_en.HrefValue = ""
			z40news.news_show_en.TooltipValue = ""

			' news_show
			z40news.news_show.LinkCustomAttributes = ""
			z40news.news_show.HrefValue = ""
			z40news.news_show.TooltipValue = ""

			' news_create
			z40news.news_create.LinkCustomAttributes = ""
			z40news.news_create.HrefValue = ""
			z40news.news_create.TooltipValue = ""

			' news_update
			z40news.news_update.LinkCustomAttributes = ""
			z40news.news_update.HrefValue = ""
			z40news.news_update.TooltipValue = ""
		End If

		' Call Row Rendered event
		If z40news.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call z40news.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		url = ew_CurrentUrl
		url = ew_RegExReplace("\?cmd=reset(all){0,1}$", url, "") ' Remove cmd=reset / cmd=resetall
		Call Breadcrumb.Add("list", z40news.TableVar, url, z40news.TableVar, True)
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
