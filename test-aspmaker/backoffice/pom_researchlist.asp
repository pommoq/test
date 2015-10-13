<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_researchinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim research_list
Set research_list = New cresearch_list
Set Page = research_list

' Page init processing
research_list.Page_Init()

' Page main processing
research_list.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
research_list.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If research.Export = "" Then %>
<script type="text/javascript">
// Page object
var research_list = new ew_Page("research_list");
research_list.PageID = "list"; // Page ID
var EW_PAGE_ID = research_list.PageID; // For backward compatibility
// Form object
var fresearchlist = new ew_Form("fresearchlist");
fresearchlist.FormKeyCountName = '<%= research_list.FormKeyCountName %>';
// Form_CustomValidate event
fresearchlist.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fresearchlist.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fresearchlist.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
var fresearchlistsrch = new ew_Form("fresearchlistsrch");
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If research.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If research_list.ExportOptions.Visible Then %>
<div class="ewListExportOptions"><% research_list.ExportOptions.Render "body", "", "", "", "", "" %></div>
<% End If %>
<% If (research.Export = "") Or (EW_EXPORT_MASTER_RECORD And research.Export = "print") Then %>
<% End If %>
<%

' Load recordset
Set research_list.Recordset = research_list.LoadRecordset()
	research_list.TotalRecs = research_list.Recordset.RecordCount
	research_list.StartRec = 1
	If research_list.DisplayRecs <= 0 Then ' Display all records
		research_list.DisplayRecs = research_list.TotalRecs
	End If
	If Not (research.ExportAll And research.Export <> "") Then
		research_list.SetUpStartRec() ' Set up start record position
	End If
research_list.RenderOtherOptions()
%>
<% If Security.IsLoggedIn() Then %>
<% If research.Export = "" And research.CurrentAction = "" Then %>
<form name="fresearchlistsrch" id="fresearchlistsrch" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewSearchTable"><tr><td>
<div class="accordion" id="fresearchlistsrch_SearchGroup">
	<div class="accordion-group">
		<div class="accordion-heading">
<a class="accordion-toggle" data-toggle="collapse" data-parent="#fresearchlistsrch_SearchGroup" href="#fresearchlistsrch_SearchBody"><%= Language.Phrase("Search") %></a>
		</div>
		<div id="fresearchlistsrch_SearchBody" class="accordion-body collapse in">
			<div class="accordion-inner">
<div id="fresearchlistsrch_SearchPanel">
<input type="hidden" name="cmd" value="search">
<input type="hidden" name="t" value="research">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewRow">
	<div class="btn-group ewButtonGroup">
	<div class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" class="input-large" value="<%= ew_HtmlEncode(research.BasicSearch.getKeyword()) %>" placeholder="<%= ew_HtmlEncode(Language.Phrase("Search")) %>">
	<button class="btn btn-primary ewButton" name="btnsubmit" id="btnsubmit" type="submit"><%= Language.Phrase("QuickSearchBtn") %></button>
	</div>
	</div>
	<div class="btn-group ewButtonGroup">
	<a class="btn ewShowAll" href="<%= research_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>
	</div>
</div>
<div id="xsr_2" class="ewRow">
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="="<% If research.BasicSearch.getSearchType = "=" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If research.BasicSearch.getSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If research.BasicSearch.getSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
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
<% research_list.ShowPageHeader() %>
<% research_list.ShowMessage %>
<table class="ewGrid"><tr><td class="ewGridContent">
<form name="fresearchlist" id="fresearchlist" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="research">
<div id="gmp_research" class="ewGridMiddlePanel">
<% If research_list.TotalRecs > 0 Then %>
<table id="tbl_researchlist" class="ewTable ewTableSeparate">
<%= research.TableCustomInnerHtml %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call research_list.RenderListOptions()

' Render list options (header, left)
research_list.ListOptions.Render "header", "left", "", "", "", ""
%>
<% If research.rsh_id.Visible Then ' rsh_id %>
	<% If research.SortUrl(research.rsh_id) = "" Then %>
		<td><div id="elh_research_rsh_id" class="research_rsh_id"><div class="ewTableHeaderCaption"><%= research.rsh_id.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research.SortUrl(research.rsh_id) %>',1);"><div id="elh_research_rsh_id" class="research_rsh_id">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research.rsh_id.FldCaption %></span><span class="ewTableHeaderSort"><% If research.rsh_id.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research.rsh_id.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research.rsh_img.Visible Then ' rsh_img %>
	<% If research.SortUrl(research.rsh_img) = "" Then %>
		<td><div id="elh_research_rsh_img" class="research_rsh_img"><div class="ewTableHeaderCaption"><%= research.rsh_img.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research.SortUrl(research.rsh_img) %>',1);"><div id="elh_research_rsh_img" class="research_rsh_img">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research.rsh_img.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research.rsh_img.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research.rsh_img.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research.rsh_date.Visible Then ' rsh_date %>
	<% If research.SortUrl(research.rsh_date) = "" Then %>
		<td><div id="elh_research_rsh_date" class="research_rsh_date"><div class="ewTableHeaderCaption"><%= research.rsh_date.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research.SortUrl(research.rsh_date) %>',1);"><div id="elh_research_rsh_date" class="research_rsh_date">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research.rsh_date.FldCaption %></span><span class="ewTableHeaderSort"><% If research.rsh_date.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research.rsh_date.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research.rsh_pdf.Visible Then ' rsh_pdf %>
	<% If research.SortUrl(research.rsh_pdf) = "" Then %>
		<td><div id="elh_research_rsh_pdf" class="research_rsh_pdf"><div class="ewTableHeaderCaption"><%= research.rsh_pdf.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research.SortUrl(research.rsh_pdf) %>',1);"><div id="elh_research_rsh_pdf" class="research_rsh_pdf">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research.rsh_pdf.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research.rsh_pdf.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research.rsh_pdf.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research.rsh_category.Visible Then ' rsh_category %>
	<% If research.SortUrl(research.rsh_category) = "" Then %>
		<td><div id="elh_research_rsh_category" class="research_rsh_category"><div class="ewTableHeaderCaption"><%= research.rsh_category.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research.SortUrl(research.rsh_category) %>',1);"><div id="elh_research_rsh_category" class="research_rsh_category">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research.rsh_category.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research.rsh_category.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research.rsh_category.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research.rsh_subject.Visible Then ' rsh_subject %>
	<% If research.SortUrl(research.rsh_subject) = "" Then %>
		<td><div id="elh_research_rsh_subject" class="research_rsh_subject"><div class="ewTableHeaderCaption"><%= research.rsh_subject.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research.SortUrl(research.rsh_subject) %>',1);"><div id="elh_research_rsh_subject" class="research_rsh_subject">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research.rsh_subject.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research.rsh_subject.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research.rsh_subject.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research.rsh_subject_th.Visible Then ' rsh_subject_th %>
	<% If research.SortUrl(research.rsh_subject_th) = "" Then %>
		<td><div id="elh_research_rsh_subject_th" class="research_rsh_subject_th"><div class="ewTableHeaderCaption"><%= research.rsh_subject_th.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research.SortUrl(research.rsh_subject_th) %>',1);"><div id="elh_research_rsh_subject_th" class="research_rsh_subject_th">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research.rsh_subject_th.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research.rsh_subject_th.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research.rsh_subject_th.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research.rsh_intro_th.Visible Then ' rsh_intro_th %>
	<% If research.SortUrl(research.rsh_intro_th) = "" Then %>
		<td><div id="elh_research_rsh_intro_th" class="research_rsh_intro_th"><div class="ewTableHeaderCaption"><%= research.rsh_intro_th.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research.SortUrl(research.rsh_intro_th) %>',1);"><div id="elh_research_rsh_intro_th" class="research_rsh_intro_th">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research.rsh_intro_th.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research.rsh_intro_th.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research.rsh_intro_th.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research.rsh_show.Visible Then ' rsh_show %>
	<% If research.SortUrl(research.rsh_show) = "" Then %>
		<td><div id="elh_research_rsh_show" class="research_rsh_show"><div class="ewTableHeaderCaption"><%= research.rsh_show.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research.SortUrl(research.rsh_show) %>',1);"><div id="elh_research_rsh_show" class="research_rsh_show">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research.rsh_show.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research.rsh_show.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research.rsh_show.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research.rsh_show_home.Visible Then ' rsh_show_home %>
	<% If research.SortUrl(research.rsh_show_home) = "" Then %>
		<td><div id="elh_research_rsh_show_home" class="research_rsh_show_home"><div class="ewTableHeaderCaption"><%= research.rsh_show_home.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research.SortUrl(research.rsh_show_home) %>',1);"><div id="elh_research_rsh_show_home" class="research_rsh_show_home">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research.rsh_show_home.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research.rsh_show_home.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research.rsh_show_home.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research.rsh_create.Visible Then ' rsh_create %>
	<% If research.SortUrl(research.rsh_create) = "" Then %>
		<td><div id="elh_research_rsh_create" class="research_rsh_create"><div class="ewTableHeaderCaption"><%= research.rsh_create.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research.SortUrl(research.rsh_create) %>',1);"><div id="elh_research_rsh_create" class="research_rsh_create">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research.rsh_create.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research.rsh_create.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research.rsh_create.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research.rsh_update.Visible Then ' rsh_update %>
	<% If research.SortUrl(research.rsh_update) = "" Then %>
		<td><div id="elh_research_rsh_update" class="research_rsh_update"><div class="ewTableHeaderCaption"><%= research.rsh_update.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research.SortUrl(research.rsh_update) %>',1);"><div id="elh_research_rsh_update" class="research_rsh_update">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research.rsh_update.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research.rsh_update.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research.rsh_update.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
research_list.ListOptions.Render "header", "right", "", "", "", ""
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (research.ExportAll And research.Export <> "") Then
	research_list.StopRec = research_list.TotalRecs
Else

	' Set the last record to display
	If research_list.TotalRecs > research_list.StartRec + research_list.DisplayRecs - 1 Then
		research_list.StopRec = research_list.StartRec + research_list.DisplayRecs - 1
	Else
		research_list.StopRec = research_list.TotalRecs
	End If
End If

' Move to first record
research_list.RecCnt = research_list.StartRec - 1
If Not research_list.Recordset.Eof Then
	research_list.Recordset.MoveFirst
	If research_list.StartRec > 1 Then research_list.Recordset.Move research_list.StartRec - 1
ElseIf Not research.AllowAddDeleteRow And research_list.StopRec = 0 Then
	research_list.StopRec = research.GridAddRowCount
End If

' Initialize Aggregate
research.RowType = EW_ROWTYPE_AGGREGATEINIT
Call research.ResetAttrs()
Call research_list.RenderRow()
research_list.RowCnt = 0

' Output date rows
Do While CLng(research_list.RecCnt) < CLng(research_list.StopRec)
	research_list.RecCnt = research_list.RecCnt + 1
	If CLng(research_list.RecCnt) >= CLng(research_list.StartRec) Then
		research_list.RowCnt = research_list.RowCnt + 1

	' Set up key count
	research_list.KeyCount = research_list.RowIndex
	Call research.ResetAttrs()
	research.CssClass = ""
	If research.CurrentAction = "gridadd" Then
	Else
		Call research_list.LoadRowValues(research_list.Recordset) ' Load row values
	End If
	research.RowType = EW_ROWTYPE_VIEW ' Render view

	' Set up row id / data-rowindex
	research.RowAttrs.AddAttributes Array(Array("data-rowindex", research_list.RowCnt), Array("id", "r" & research_list.RowCnt & "_research"), Array("data-rowtype", research.RowType))

	' Render row
	Call research_list.RenderRow()

	' Render list options
	Call research_list.RenderListOptions()
%>
	<tr<%= research.RowAttributes %>>
<%

' Render list options (body, left)
research_list.ListOptions.Render "body", "left", research_list.RowCnt, "", "", ""
%>
	<% If research.rsh_id.Visible Then ' rsh_id %>
		<td<%= research.rsh_id.CellAttributes %>>
<span<%= research.rsh_id.ViewAttributes %>>
<%= research.rsh_id.ListViewValue %>
</span>
<a id="<%= research_list.PageObjName & "_row_" & research_list.RowCnt %>"></a></td>
	<% End If %>
	<% If research.rsh_img.Visible Then ' rsh_img %>
		<td<%= research.rsh_img.CellAttributes %>>
<span<%= research.rsh_img.ViewAttributes %>>
<%= research.rsh_img.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research.rsh_date.Visible Then ' rsh_date %>
		<td<%= research.rsh_date.CellAttributes %>>
<span<%= research.rsh_date.ViewAttributes %>>
<%= research.rsh_date.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research.rsh_pdf.Visible Then ' rsh_pdf %>
		<td<%= research.rsh_pdf.CellAttributes %>>
<span<%= research.rsh_pdf.ViewAttributes %>>
<%= research.rsh_pdf.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research.rsh_category.Visible Then ' rsh_category %>
		<td<%= research.rsh_category.CellAttributes %>>
<span<%= research.rsh_category.ViewAttributes %>>
<%= research.rsh_category.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research.rsh_subject.Visible Then ' rsh_subject %>
		<td<%= research.rsh_subject.CellAttributes %>>
<span<%= research.rsh_subject.ViewAttributes %>>
<%= research.rsh_subject.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research.rsh_subject_th.Visible Then ' rsh_subject_th %>
		<td<%= research.rsh_subject_th.CellAttributes %>>
<span<%= research.rsh_subject_th.ViewAttributes %>>
<%= research.rsh_subject_th.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research.rsh_intro_th.Visible Then ' rsh_intro_th %>
		<td<%= research.rsh_intro_th.CellAttributes %>>
<span<%= research.rsh_intro_th.ViewAttributes %>>
<%= research.rsh_intro_th.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research.rsh_show.Visible Then ' rsh_show %>
		<td<%= research.rsh_show.CellAttributes %>>
<span<%= research.rsh_show.ViewAttributes %>>
<%= research.rsh_show.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research.rsh_show_home.Visible Then ' rsh_show_home %>
		<td<%= research.rsh_show_home.CellAttributes %>>
<span<%= research.rsh_show_home.ViewAttributes %>>
<%= research.rsh_show_home.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research.rsh_create.Visible Then ' rsh_create %>
		<td<%= research.rsh_create.CellAttributes %>>
<span<%= research.rsh_create.ViewAttributes %>>
<%= research.rsh_create.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research.rsh_update.Visible Then ' rsh_update %>
		<td<%= research.rsh_update.CellAttributes %>>
<span<%= research.rsh_update.ViewAttributes %>>
<%= research.rsh_update.ListViewValue %>
</span>
</td>
	<% End If %>
<%

' Render list options (body, right)
research_list.ListOptions.Render "body", "right", research_list.RowCnt, "", "", ""
%>
	</tr>
<%
	End If
	If research.CurrentAction <> "gridadd" Then
		research_list.Recordset.MoveNext()
	End If
Loop
%>
</tbody>
</table>
<% End If %>
<% If research.CurrentAction = "" Then %>
<input type="hidden" name="a_list" id="a_list" value="">
<% End If %>
</div>
</form>
<%

' Close recordset and connection
research_list.Recordset.Close
Set research_list.Recordset = Nothing
%>
<% If research.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If research.CurrentAction <> "gridadd" And research.CurrentAction <> "gridedit" Then %>
<form name="ewPagerForm" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewPager">
<tr><td>
<% If Not IsObject(research_list.Pager) Then Set research_list.Pager = ew_NewPrevNextPager(research_list.StartRec, research_list.DisplayRecs, research_list.TotalRecs) %>
<% If research_list.Pager.RecordCount > 0 Then %>
<table class="ewStdTable"><tbody><tr><td>
	<%= Language.Phrase("Page") %>&nbsp;
<div class="input-prepend input-append">
<!--first page button-->
	<% If research_list.Pager.FirstButton.Enabled Then %>
	<a class="btn btn-small" href="<%= research_list.PageUrl %>start=<%= research_list.Pager.FirstButton.Start %>"><i class="icon-step-backward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-backward"></i></a>
	<% End If %>
<!--previous page button-->
	<% If research_list.Pager.PrevButton.Enabled Then %>
	<a class="btn btn-small" href="<%= research_list.PageUrl %>start=<%= research_list.Pager.PrevButton.Start %>"><i class="icon-prev"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-prev"></i></a>
	<% End If %>
<!--current page number-->
	<input class="input-mini" type="text" name="<%= EW_TABLE_PAGE_NO %>" value="<%= research_list.Pager.CurrentPage %>">
<!--next page button-->
	<% If research_list.Pager.NextButton.Enabled Then %>
	<a class="btn btn-small" href="<%= research_list.PageUrl %>start=<%= research_list.Pager.NextButton.Start %>"><i class="icon-play"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-play"></i></a>
	<% End If %>
<!--last page button-->
	<% If research_list.Pager.LastButton.Enabled Then %>
	<a class="btn btn-small" href="<%= research_list.PageUrl %>start=<%= research_list.Pager.LastButton.Start %>"><i class="icon-step-forward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-forward"></i></a>
	<% End If %>
</div>
	&nbsp;<%= Language.Phrase("of") %>&nbsp;<%= research_list.Pager.PageCount %>
</td>
<td>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<%= Language.Phrase("Record") %>&nbsp;<%= research_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= research_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= research_list.Pager.RecordCount %>
</td>
</tr></tbody></table>
<% Else %>
	<% If research_list.SearchWhere = "0=101" Then %>
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
	research_list.AddEditOptions.Render "body", "bottom", "", "", "", ""
	research_list.DetailOptions.Render "body", "bottom", "", "", "", ""
	research_list.ActionOptions.Render "body", "bottom", "", "", "", ""
%>
</div>
</div>
<% End If %>
</td></tr></table>
<% If research.Export = "" Then %>
<script type="text/javascript">
fresearchlistsrch.Init();
fresearchlist.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<% End If %>
<%
research_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If research.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set research_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cresearch_list

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
		TableName = "research"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "research_list"
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
		If research.UseTokenInUrl Then PageUrl = PageUrl & "t=" & research.TableVar & "&" ' add page token
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
		If research.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (research.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (research.TableVar = Request.QueryString("t"))
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
		FormName = "fresearchlist"
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
		If IsEmpty(research) Then Set research = New cresearch
		Set Table = research

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		ExportPdfUrl = PageUrl & "export=pdf"
		AddUrl = "pom_researchadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "pom_researchdelete.asp"
		MultiUpdateUrl = "pom_researchupdate.asp"

		' Initialize other table object
		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "research"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' List options
		Set ListOptions = New cListOptions
		ListOptions.TableVar = research.TableVar

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = research.TableVar
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
				research.GridAddRowCount = gridaddcnt
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
		If UBound(research.CustomActions.CustomArray) >= 0 Then
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
		Set research = Nothing
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
			If research.Export = "" Then
				SetupBreadcrumb()
			End If

			' Hide list options
			If research.Export <> "" Then
				Call ListOptions.HideAllOptions(Array("sequence"))
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			ElseIf research.CurrentAction = "gridadd" Or research.CurrentAction = "gridedit" Then
				Call ListOptions.HideAllOptions(Array())
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			End If

			' Hide export options
			If research.Export <> "" Or research.CurrentAction <> "" Then
				ExportOptions.HideAllOptions(Array())
			End If

			' Hide other options
			If research.Export <> "" Then
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
			Call research.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If research.RecordsPerPage <> "" Then
			DisplayRecs = research.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Load search default if no existing search criteria
		If Not CheckSearchParms() Then

			' Load basic search from default
			research.BasicSearch.Keyword = research.BasicSearch.KeywordDefault
			research.BasicSearch.SearchType = research.BasicSearch.SearchTypeDefault
			research.BasicSearch.setSearchType(research.BasicSearch.SearchTypeDefault)
			If research.BasicSearch.Keyword <> "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call research.Recordset_Searching(SearchWhere)

		' Save search criteria
		If Command = "search" And Not RestoreSearch Then
			research.SearchWhere = SearchWhere ' Save to Session
			StartRec = 1 ' Reset start record counter
			research.StartRecordNumber = StartRec
		Else
			SearchWhere = research.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		research.SessionWhere = sFilter
		research.CurrentFilter = ""
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
				sFilter = research.KeyFilter
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
			research.rsh_id.FormValue = arrKeyFlds(0)
			If Not IsNumeric(research.rsh_id.FormValue) Then
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
			Call BuildBasicSearchSQL(sWhere, research.rsh_img, Keyword)
			Call BuildBasicSearchSQL(sWhere, research.rsh_pdf, Keyword)
			Call BuildBasicSearchSQL(sWhere, research.rsh_category, Keyword)
			Call BuildBasicSearchSQL(sWhere, research.rsh_subject, Keyword)
			Call BuildBasicSearchSQL(sWhere, research.rsh_subject_th, Keyword)
			Call BuildBasicSearchSQL(sWhere, research.rsh_intro, Keyword)
			Call BuildBasicSearchSQL(sWhere, research.rsh_intro_th, Keyword)
			Call BuildBasicSearchSQL(sWhere, research.rsh_content, Keyword)
			Call BuildBasicSearchSQL(sWhere, research.rsh_content_th, Keyword)
			Call BuildBasicSearchSQL(sWhere, research.rsh_show, Keyword)
			Call BuildBasicSearchSQL(sWhere, research.rsh_show_home, Keyword)
			Call BuildBasicSearchSQL(sWhere, research.rsh_create, Keyword)
			Call BuildBasicSearchSQL(sWhere, research.rsh_update, Keyword)
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
		sSearchKeyword = research.BasicSearch.Keyword
		sSearchType = research.BasicSearch.SearchType
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
			research.BasicSearch.setKeyword(sSearchKeyword)
			research.BasicSearch.setSearchType(sSearchType)
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' Check if search parm exists
	Function CheckSearchParms()

		' Check basic search
		If research.BasicSearch.IssetSession() Then
			CheckSearchParms = True
			Exit Function
		End If
		CheckSearchParms = False
	End Function

	' Clear all search parameters
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		research.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' Load advanced search default values
	Function LoadAdvancedSearchDefault()
		LoadAdvancedSearchDefault = False
	End Function

	' Clear all basic search parameters
	Sub ResetBasicSearchParms()
		research.BasicSearch.UnsetSession()
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()

		' Restore search flag
		RestoreSearch = True

		' Restore basic search values
		Call research.BasicSearch.Load()
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
			research.CurrentOrder = Request.QueryString("order")
			research.CurrentOrderType = Request.QueryString("ordertype")

			' Field rsh_id
			Call research.UpdateSort(research.rsh_id)

			' Field rsh_img
			Call research.UpdateSort(research.rsh_img)

			' Field rsh_date
			Call research.UpdateSort(research.rsh_date)

			' Field rsh_pdf
			Call research.UpdateSort(research.rsh_pdf)

			' Field rsh_category
			Call research.UpdateSort(research.rsh_category)

			' Field rsh_subject
			Call research.UpdateSort(research.rsh_subject)

			' Field rsh_subject_th
			Call research.UpdateSort(research.rsh_subject_th)

			' Field rsh_intro_th
			Call research.UpdateSort(research.rsh_intro_th)

			' Field rsh_show
			Call research.UpdateSort(research.rsh_show)

			' Field rsh_show_home
			Call research.UpdateSort(research.rsh_show_home)

			' Field rsh_create
			Call research.UpdateSort(research.rsh_create)

			' Field rsh_update
			Call research.UpdateSort(research.rsh_update)
			research.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = research.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If research.SqlOrderBy <> "" Then
				sOrderBy = research.SqlOrderBy
				research.SessionOrderBy = sOrderBy
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
				research.SessionOrderBy = sOrderBy
				research.rsh_id.Sort = ""
				research.rsh_img.Sort = ""
				research.rsh_date.Sort = ""
				research.rsh_pdf.Sort = ""
				research.rsh_category.Sort = ""
				research.rsh_subject.Sort = ""
				research.rsh_subject_th.Sort = ""
				research.rsh_intro_th.Sort = ""
				research.rsh_show.Sort = ""
				research.rsh_show_home.Sort = ""
				research.rsh_create.Sort = ""
				research.rsh_update.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			research.StartRecordNumber = StartRec
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
		ListOptions.GetItem("checkbox").Body = "<label class=""checkbox""><input type=""checkbox"" name=""key_m"" value=""" & ew_HtmlEncode(research.rsh_id.CurrentValue) & """ onclick='ew_ClickMultiCheckbox(event, this);'></label>"
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
			For i = 0 to UBound(research.CustomActions.CustomArray)
				Action = research.CustomActions.CustomArray(i)(0)
				Name = research.CustomActions.CustomArray(i)(1)

				' Add custom action
				Call opt.Add("custom_" & Action)
				Set item = opt.GetItem("custom_" & Action)
				item.Body = "<a class=""ewAction ewCustomAction"" href="""" onclick=""ew_SubmitSelected(document.fresearchlist, '" & ew_CurrentUrl & "', null, '" & Action & "');return false;"">" & Name & "</a>"
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
		sFilter = research.GetKeyFilter
		UserAction = Request.Form("useraction") & ""
		Processed = False
		If sFilter <> "" And UserAction <> "" Then
			research.CurrentFilter = sFilter
			sSql = research.SQL
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
				ElseIf research.CancelMessage <> "" Then
					FailureMessage = research.CancelMessage
					research.CancelMessage = ""
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
				research.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					research.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = research.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			research.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			research.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			research.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		research.BasicSearch.Keyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)&""
		If research.BasicSearch.Keyword <> "" Then Command = "search"
		research.BasicSearch.SearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)&""
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = research.CurrentFilter
		Call research.Recordset_Selecting(sFilter)
		research.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = research.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call research.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = research.KeyFilter

		' Call Row Selecting event
		Call research.Row_Selecting(sFilter)

		' Load sql based on filter
		research.CurrentFilter = sFilter
		sSql = research.SQL
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
		Call research.Row_Selected(RsRow)
		research.rsh_id.DbValue = RsRow("rsh_id")
		research.rsh_img.DbValue = RsRow("rsh_img")
		research.rsh_date.DbValue = RsRow("rsh_date")
		research.rsh_pdf.DbValue = RsRow("rsh_pdf")
		research.rsh_category.DbValue = RsRow("rsh_category")
		research.rsh_subject.DbValue = RsRow("rsh_subject")
		research.rsh_subject_th.DbValue = RsRow("rsh_subject_th")
		research.rsh_intro.DbValue = RsRow("rsh_intro")
		research.rsh_intro_th.DbValue = RsRow("rsh_intro_th")
		research.rsh_content.DbValue = RsRow("rsh_content")
		research.rsh_content_th.DbValue = RsRow("rsh_content_th")
		research.rsh_show.DbValue = RsRow("rsh_show")
		research.rsh_show_home.DbValue = RsRow("rsh_show_home")
		research.rsh_create.DbValue = RsRow("rsh_create")
		research.rsh_update.DbValue = RsRow("rsh_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		research.rsh_id.m_DbValue = Rs("rsh_id")
		research.rsh_img.m_DbValue = Rs("rsh_img")
		research.rsh_date.m_DbValue = Rs("rsh_date")
		research.rsh_pdf.m_DbValue = Rs("rsh_pdf")
		research.rsh_category.m_DbValue = Rs("rsh_category")
		research.rsh_subject.m_DbValue = Rs("rsh_subject")
		research.rsh_subject_th.m_DbValue = Rs("rsh_subject_th")
		research.rsh_intro.m_DbValue = Rs("rsh_intro")
		research.rsh_intro_th.m_DbValue = Rs("rsh_intro_th")
		research.rsh_content.m_DbValue = Rs("rsh_content")
		research.rsh_content_th.m_DbValue = Rs("rsh_content_th")
		research.rsh_show.m_DbValue = Rs("rsh_show")
		research.rsh_show_home.m_DbValue = Rs("rsh_show_home")
		research.rsh_create.m_DbValue = Rs("rsh_create")
		research.rsh_update.m_DbValue = Rs("rsh_update")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If research.GetKey("rsh_id")&"" <> "" Then
			research.rsh_id.CurrentValue = research.GetKey("rsh_id") ' rsh_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			research.CurrentFilter = research.KeyFilter
			Dim sSql
			sSql = research.SQL
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
		ViewUrl = research.ViewUrl("")
		EditUrl = research.EditUrl("")
		InlineEditUrl = research.InlineEditUrl
		CopyUrl = research.CopyUrl("")
		InlineCopyUrl = research.InlineCopyUrl
		DeleteUrl = research.DeleteUrl

		' Call Row Rendering event
		Call research.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' rsh_id
		' rsh_img
		' rsh_date
		' rsh_pdf
		' rsh_category
		' rsh_subject
		' rsh_subject_th
		' rsh_intro
		' rsh_intro_th
		' rsh_content
		' rsh_content_th
		' rsh_show
		' rsh_show_home
		' rsh_create
		' rsh_update
		' -----------
		'  View  Row
		' -----------

		If research.RowType = EW_ROWTYPE_VIEW Then ' View row

			' rsh_id
			research.rsh_id.ViewValue = research.rsh_id.CurrentValue
			research.rsh_id.ViewCustomAttributes = ""

			' rsh_img
			research.rsh_img.ViewValue = research.rsh_img.CurrentValue
			research.rsh_img.ViewCustomAttributes = ""

			' rsh_date
			research.rsh_date.ViewValue = research.rsh_date.CurrentValue
			research.rsh_date.ViewCustomAttributes = ""

			' rsh_pdf
			research.rsh_pdf.ViewValue = research.rsh_pdf.CurrentValue
			research.rsh_pdf.ViewCustomAttributes = ""

			' rsh_category
			research.rsh_category.ViewValue = research.rsh_category.CurrentValue
			research.rsh_category.ViewCustomAttributes = ""

			' rsh_subject
			research.rsh_subject.ViewValue = research.rsh_subject.CurrentValue
			research.rsh_subject.ViewCustomAttributes = ""

			' rsh_subject_th
			research.rsh_subject_th.ViewValue = research.rsh_subject_th.CurrentValue
			research.rsh_subject_th.ViewCustomAttributes = ""

			' rsh_intro_th
			research.rsh_intro_th.ViewValue = research.rsh_intro_th.CurrentValue
			research.rsh_intro_th.ViewCustomAttributes = ""

			' rsh_show
			research.rsh_show.ViewValue = research.rsh_show.CurrentValue
			research.rsh_show.ViewCustomAttributes = ""

			' rsh_show_home
			research.rsh_show_home.ViewValue = research.rsh_show_home.CurrentValue
			research.rsh_show_home.ViewCustomAttributes = ""

			' rsh_create
			research.rsh_create.ViewValue = research.rsh_create.CurrentValue
			research.rsh_create.ViewCustomAttributes = ""

			' rsh_update
			research.rsh_update.ViewValue = research.rsh_update.CurrentValue
			research.rsh_update.ViewCustomAttributes = ""

			' View refer script
			' rsh_id

			research.rsh_id.LinkCustomAttributes = ""
			research.rsh_id.HrefValue = ""
			research.rsh_id.TooltipValue = ""

			' rsh_img
			research.rsh_img.LinkCustomAttributes = ""
			research.rsh_img.HrefValue = ""
			research.rsh_img.TooltipValue = ""

			' rsh_date
			research.rsh_date.LinkCustomAttributes = ""
			research.rsh_date.HrefValue = ""
			research.rsh_date.TooltipValue = ""

			' rsh_pdf
			research.rsh_pdf.LinkCustomAttributes = ""
			research.rsh_pdf.HrefValue = ""
			research.rsh_pdf.TooltipValue = ""

			' rsh_category
			research.rsh_category.LinkCustomAttributes = ""
			research.rsh_category.HrefValue = ""
			research.rsh_category.TooltipValue = ""

			' rsh_subject
			research.rsh_subject.LinkCustomAttributes = ""
			research.rsh_subject.HrefValue = ""
			research.rsh_subject.TooltipValue = ""

			' rsh_subject_th
			research.rsh_subject_th.LinkCustomAttributes = ""
			research.rsh_subject_th.HrefValue = ""
			research.rsh_subject_th.TooltipValue = ""

			' rsh_intro_th
			research.rsh_intro_th.LinkCustomAttributes = ""
			research.rsh_intro_th.HrefValue = ""
			research.rsh_intro_th.TooltipValue = ""

			' rsh_show
			research.rsh_show.LinkCustomAttributes = ""
			research.rsh_show.HrefValue = ""
			research.rsh_show.TooltipValue = ""

			' rsh_show_home
			research.rsh_show_home.LinkCustomAttributes = ""
			research.rsh_show_home.HrefValue = ""
			research.rsh_show_home.TooltipValue = ""

			' rsh_create
			research.rsh_create.LinkCustomAttributes = ""
			research.rsh_create.HrefValue = ""
			research.rsh_create.TooltipValue = ""

			' rsh_update
			research.rsh_update.LinkCustomAttributes = ""
			research.rsh_update.HrefValue = ""
			research.rsh_update.TooltipValue = ""
		End If

		' Call Row Rendered event
		If research.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call research.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		url = ew_CurrentUrl
		url = ew_RegExReplace("\?cmd=reset(all){0,1}$", url, "") ' Remove cmd=reset / cmd=resetall
		Call Breadcrumb.Add("list", research.TableVar, url, research.TableVar, True)
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
