<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_research_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim research_th_list
Set research_th_list = New cresearch_th_list
Set Page = research_th_list

' Page init processing
research_th_list.Page_Init()

' Page main processing
research_th_list.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
research_th_list.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If research_th.Export = "" Then %>
<script type="text/javascript">
// Page object
var research_th_list = new ew_Page("research_th_list");
research_th_list.PageID = "list"; // Page ID
var EW_PAGE_ID = research_th_list.PageID; // For backward compatibility
// Form object
var fresearch_thlist = new ew_Form("fresearch_thlist");
fresearch_thlist.FormKeyCountName = '<%= research_th_list.FormKeyCountName %>';
// Form_CustomValidate event
fresearch_thlist.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fresearch_thlist.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fresearch_thlist.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
var fresearch_thlistsrch = new ew_Form("fresearch_thlistsrch");
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If research_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If research_th_list.ExportOptions.Visible Then %>
<div class="ewListExportOptions"><% research_th_list.ExportOptions.Render "body", "", "", "", "", "" %></div>
<% End If %>
<% If (research_th.Export = "") Or (EW_EXPORT_MASTER_RECORD And research_th.Export = "print") Then %>
<% End If %>
<%

' Load recordset
Set research_th_list.Recordset = research_th_list.LoadRecordset()
	research_th_list.TotalRecs = research_th_list.Recordset.RecordCount
	research_th_list.StartRec = 1
	If research_th_list.DisplayRecs <= 0 Then ' Display all records
		research_th_list.DisplayRecs = research_th_list.TotalRecs
	End If
	If Not (research_th.ExportAll And research_th.Export <> "") Then
		research_th_list.SetUpStartRec() ' Set up start record position
	End If
research_th_list.RenderOtherOptions()
%>
<% If Security.IsLoggedIn() Then %>
<% If research_th.Export = "" And research_th.CurrentAction = "" Then %>
<form name="fresearch_thlistsrch" id="fresearch_thlistsrch" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewSearchTable"><tr><td>
<div class="accordion" id="fresearch_thlistsrch_SearchGroup">
	<div class="accordion-group">
		<div class="accordion-heading">
<a class="accordion-toggle" data-toggle="collapse" data-parent="#fresearch_thlistsrch_SearchGroup" href="#fresearch_thlistsrch_SearchBody"><%= Language.Phrase("Search") %></a>
		</div>
		<div id="fresearch_thlistsrch_SearchBody" class="accordion-body collapse in">
			<div class="accordion-inner">
<div id="fresearch_thlistsrch_SearchPanel">
<input type="hidden" name="cmd" value="search">
<input type="hidden" name="t" value="research_th">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewRow">
	<div class="btn-group ewButtonGroup">
	<div class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" class="input-large" value="<%= ew_HtmlEncode(research_th.BasicSearch.getKeyword()) %>" placeholder="<%= ew_HtmlEncode(Language.Phrase("Search")) %>">
	<button class="btn btn-primary ewButton" name="btnsubmit" id="btnsubmit" type="submit"><%= Language.Phrase("QuickSearchBtn") %></button>
	</div>
	</div>
	<div class="btn-group ewButtonGroup">
	<a class="btn ewShowAll" href="<%= research_th_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>
	</div>
</div>
<div id="xsr_2" class="ewRow">
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="="<% If research_th.BasicSearch.getSearchType = "=" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If research_th.BasicSearch.getSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If research_th.BasicSearch.getSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
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
<% research_th_list.ShowPageHeader() %>
<% research_th_list.ShowMessage %>
<table class="ewGrid"><tr><td class="ewGridContent">
<form name="fresearch_thlist" id="fresearch_thlist" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="research_th">
<div id="gmp_research_th" class="ewGridMiddlePanel">
<% If research_th_list.TotalRecs > 0 Then %>
<table id="tbl_research_thlist" class="ewTable ewTableSeparate">
<%= research_th.TableCustomInnerHtml %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call research_th_list.RenderListOptions()

' Render list options (header, left)
research_th_list.ListOptions.Render "header", "left", "", "", "", ""
%>
<% If research_th.rsh_id.Visible Then ' rsh_id %>
	<% If research_th.SortUrl(research_th.rsh_id) = "" Then %>
		<td><div id="elh_research_th_rsh_id" class="research_th_rsh_id"><div class="ewTableHeaderCaption"><%= research_th.rsh_id.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research_th.SortUrl(research_th.rsh_id) %>',1);"><div id="elh_research_th_rsh_id" class="research_th_rsh_id">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research_th.rsh_id.FldCaption %></span><span class="ewTableHeaderSort"><% If research_th.rsh_id.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research_th.rsh_id.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research_th.rsh_img.Visible Then ' rsh_img %>
	<% If research_th.SortUrl(research_th.rsh_img) = "" Then %>
		<td><div id="elh_research_th_rsh_img" class="research_th_rsh_img"><div class="ewTableHeaderCaption"><%= research_th.rsh_img.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research_th.SortUrl(research_th.rsh_img) %>',1);"><div id="elh_research_th_rsh_img" class="research_th_rsh_img">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research_th.rsh_img.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research_th.rsh_img.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research_th.rsh_img.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research_th.rsh_date.Visible Then ' rsh_date %>
	<% If research_th.SortUrl(research_th.rsh_date) = "" Then %>
		<td><div id="elh_research_th_rsh_date" class="research_th_rsh_date"><div class="ewTableHeaderCaption"><%= research_th.rsh_date.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research_th.SortUrl(research_th.rsh_date) %>',1);"><div id="elh_research_th_rsh_date" class="research_th_rsh_date">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research_th.rsh_date.FldCaption %></span><span class="ewTableHeaderSort"><% If research_th.rsh_date.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research_th.rsh_date.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research_th.rsh_pdf.Visible Then ' rsh_pdf %>
	<% If research_th.SortUrl(research_th.rsh_pdf) = "" Then %>
		<td><div id="elh_research_th_rsh_pdf" class="research_th_rsh_pdf"><div class="ewTableHeaderCaption"><%= research_th.rsh_pdf.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research_th.SortUrl(research_th.rsh_pdf) %>',1);"><div id="elh_research_th_rsh_pdf" class="research_th_rsh_pdf">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research_th.rsh_pdf.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research_th.rsh_pdf.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research_th.rsh_pdf.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research_th.rsh_category.Visible Then ' rsh_category %>
	<% If research_th.SortUrl(research_th.rsh_category) = "" Then %>
		<td><div id="elh_research_th_rsh_category" class="research_th_rsh_category"><div class="ewTableHeaderCaption"><%= research_th.rsh_category.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research_th.SortUrl(research_th.rsh_category) %>',1);"><div id="elh_research_th_rsh_category" class="research_th_rsh_category">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research_th.rsh_category.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research_th.rsh_category.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research_th.rsh_category.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research_th.rsh_subject.Visible Then ' rsh_subject %>
	<% If research_th.SortUrl(research_th.rsh_subject) = "" Then %>
		<td><div id="elh_research_th_rsh_subject" class="research_th_rsh_subject"><div class="ewTableHeaderCaption"><%= research_th.rsh_subject.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research_th.SortUrl(research_th.rsh_subject) %>',1);"><div id="elh_research_th_rsh_subject" class="research_th_rsh_subject">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research_th.rsh_subject.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research_th.rsh_subject.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research_th.rsh_subject.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research_th.rsh_subject_th.Visible Then ' rsh_subject_th %>
	<% If research_th.SortUrl(research_th.rsh_subject_th) = "" Then %>
		<td><div id="elh_research_th_rsh_subject_th" class="research_th_rsh_subject_th"><div class="ewTableHeaderCaption"><%= research_th.rsh_subject_th.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research_th.SortUrl(research_th.rsh_subject_th) %>',1);"><div id="elh_research_th_rsh_subject_th" class="research_th_rsh_subject_th">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research_th.rsh_subject_th.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research_th.rsh_subject_th.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research_th.rsh_subject_th.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research_th.rsh_intro_th.Visible Then ' rsh_intro_th %>
	<% If research_th.SortUrl(research_th.rsh_intro_th) = "" Then %>
		<td><div id="elh_research_th_rsh_intro_th" class="research_th_rsh_intro_th"><div class="ewTableHeaderCaption"><%= research_th.rsh_intro_th.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research_th.SortUrl(research_th.rsh_intro_th) %>',1);"><div id="elh_research_th_rsh_intro_th" class="research_th_rsh_intro_th">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research_th.rsh_intro_th.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research_th.rsh_intro_th.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research_th.rsh_intro_th.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research_th.rsh_show.Visible Then ' rsh_show %>
	<% If research_th.SortUrl(research_th.rsh_show) = "" Then %>
		<td><div id="elh_research_th_rsh_show" class="research_th_rsh_show"><div class="ewTableHeaderCaption"><%= research_th.rsh_show.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research_th.SortUrl(research_th.rsh_show) %>',1);"><div id="elh_research_th_rsh_show" class="research_th_rsh_show">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research_th.rsh_show.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research_th.rsh_show.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research_th.rsh_show.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research_th.rsh_show_home.Visible Then ' rsh_show_home %>
	<% If research_th.SortUrl(research_th.rsh_show_home) = "" Then %>
		<td><div id="elh_research_th_rsh_show_home" class="research_th_rsh_show_home"><div class="ewTableHeaderCaption"><%= research_th.rsh_show_home.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research_th.SortUrl(research_th.rsh_show_home) %>',1);"><div id="elh_research_th_rsh_show_home" class="research_th_rsh_show_home">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research_th.rsh_show_home.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research_th.rsh_show_home.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research_th.rsh_show_home.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research_th.rsh_create.Visible Then ' rsh_create %>
	<% If research_th.SortUrl(research_th.rsh_create) = "" Then %>
		<td><div id="elh_research_th_rsh_create" class="research_th_rsh_create"><div class="ewTableHeaderCaption"><%= research_th.rsh_create.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research_th.SortUrl(research_th.rsh_create) %>',1);"><div id="elh_research_th_rsh_create" class="research_th_rsh_create">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research_th.rsh_create.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research_th.rsh_create.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research_th.rsh_create.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If research_th.rsh_update.Visible Then ' rsh_update %>
	<% If research_th.SortUrl(research_th.rsh_update) = "" Then %>
		<td><div id="elh_research_th_rsh_update" class="research_th_rsh_update"><div class="ewTableHeaderCaption"><%= research_th.rsh_update.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= research_th.SortUrl(research_th.rsh_update) %>',1);"><div id="elh_research_th_rsh_update" class="research_th_rsh_update">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= research_th.rsh_update.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If research_th.rsh_update.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf research_th.rsh_update.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
research_th_list.ListOptions.Render "header", "right", "", "", "", ""
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (research_th.ExportAll And research_th.Export <> "") Then
	research_th_list.StopRec = research_th_list.TotalRecs
Else

	' Set the last record to display
	If research_th_list.TotalRecs > research_th_list.StartRec + research_th_list.DisplayRecs - 1 Then
		research_th_list.StopRec = research_th_list.StartRec + research_th_list.DisplayRecs - 1
	Else
		research_th_list.StopRec = research_th_list.TotalRecs
	End If
End If

' Move to first record
research_th_list.RecCnt = research_th_list.StartRec - 1
If Not research_th_list.Recordset.Eof Then
	research_th_list.Recordset.MoveFirst
	If research_th_list.StartRec > 1 Then research_th_list.Recordset.Move research_th_list.StartRec - 1
ElseIf Not research_th.AllowAddDeleteRow And research_th_list.StopRec = 0 Then
	research_th_list.StopRec = research_th.GridAddRowCount
End If

' Initialize Aggregate
research_th.RowType = EW_ROWTYPE_AGGREGATEINIT
Call research_th.ResetAttrs()
Call research_th_list.RenderRow()
research_th_list.RowCnt = 0

' Output date rows
Do While CLng(research_th_list.RecCnt) < CLng(research_th_list.StopRec)
	research_th_list.RecCnt = research_th_list.RecCnt + 1
	If CLng(research_th_list.RecCnt) >= CLng(research_th_list.StartRec) Then
		research_th_list.RowCnt = research_th_list.RowCnt + 1

	' Set up key count
	research_th_list.KeyCount = research_th_list.RowIndex
	Call research_th.ResetAttrs()
	research_th.CssClass = ""
	If research_th.CurrentAction = "gridadd" Then
	Else
		Call research_th_list.LoadRowValues(research_th_list.Recordset) ' Load row values
	End If
	research_th.RowType = EW_ROWTYPE_VIEW ' Render view

	' Set up row id / data-rowindex
	research_th.RowAttrs.AddAttributes Array(Array("data-rowindex", research_th_list.RowCnt), Array("id", "r" & research_th_list.RowCnt & "_research_th"), Array("data-rowtype", research_th.RowType))

	' Render row
	Call research_th_list.RenderRow()

	' Render list options
	Call research_th_list.RenderListOptions()
%>
	<tr<%= research_th.RowAttributes %>>
<%

' Render list options (body, left)
research_th_list.ListOptions.Render "body", "left", research_th_list.RowCnt, "", "", ""
%>
	<% If research_th.rsh_id.Visible Then ' rsh_id %>
		<td<%= research_th.rsh_id.CellAttributes %>>
<span<%= research_th.rsh_id.ViewAttributes %>>
<%= research_th.rsh_id.ListViewValue %>
</span>
<a id="<%= research_th_list.PageObjName & "_row_" & research_th_list.RowCnt %>"></a></td>
	<% End If %>
	<% If research_th.rsh_img.Visible Then ' rsh_img %>
		<td<%= research_th.rsh_img.CellAttributes %>>
<span<%= research_th.rsh_img.ViewAttributes %>>
<%= research_th.rsh_img.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research_th.rsh_date.Visible Then ' rsh_date %>
		<td<%= research_th.rsh_date.CellAttributes %>>
<span<%= research_th.rsh_date.ViewAttributes %>>
<%= research_th.rsh_date.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research_th.rsh_pdf.Visible Then ' rsh_pdf %>
		<td<%= research_th.rsh_pdf.CellAttributes %>>
<span<%= research_th.rsh_pdf.ViewAttributes %>>
<%= research_th.rsh_pdf.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research_th.rsh_category.Visible Then ' rsh_category %>
		<td<%= research_th.rsh_category.CellAttributes %>>
<span<%= research_th.rsh_category.ViewAttributes %>>
<%= research_th.rsh_category.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research_th.rsh_subject.Visible Then ' rsh_subject %>
		<td<%= research_th.rsh_subject.CellAttributes %>>
<span<%= research_th.rsh_subject.ViewAttributes %>>
<%= research_th.rsh_subject.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research_th.rsh_subject_th.Visible Then ' rsh_subject_th %>
		<td<%= research_th.rsh_subject_th.CellAttributes %>>
<span<%= research_th.rsh_subject_th.ViewAttributes %>>
<%= research_th.rsh_subject_th.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research_th.rsh_intro_th.Visible Then ' rsh_intro_th %>
		<td<%= research_th.rsh_intro_th.CellAttributes %>>
<span<%= research_th.rsh_intro_th.ViewAttributes %>>
<%= research_th.rsh_intro_th.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research_th.rsh_show.Visible Then ' rsh_show %>
		<td<%= research_th.rsh_show.CellAttributes %>>
<span<%= research_th.rsh_show.ViewAttributes %>>
<%= research_th.rsh_show.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research_th.rsh_show_home.Visible Then ' rsh_show_home %>
		<td<%= research_th.rsh_show_home.CellAttributes %>>
<span<%= research_th.rsh_show_home.ViewAttributes %>>
<%= research_th.rsh_show_home.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research_th.rsh_create.Visible Then ' rsh_create %>
		<td<%= research_th.rsh_create.CellAttributes %>>
<span<%= research_th.rsh_create.ViewAttributes %>>
<%= research_th.rsh_create.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If research_th.rsh_update.Visible Then ' rsh_update %>
		<td<%= research_th.rsh_update.CellAttributes %>>
<span<%= research_th.rsh_update.ViewAttributes %>>
<%= research_th.rsh_update.ListViewValue %>
</span>
</td>
	<% End If %>
<%

' Render list options (body, right)
research_th_list.ListOptions.Render "body", "right", research_th_list.RowCnt, "", "", ""
%>
	</tr>
<%
	End If
	If research_th.CurrentAction <> "gridadd" Then
		research_th_list.Recordset.MoveNext()
	End If
Loop
%>
</tbody>
</table>
<% End If %>
<% If research_th.CurrentAction = "" Then %>
<input type="hidden" name="a_list" id="a_list" value="">
<% End If %>
</div>
</form>
<%

' Close recordset and connection
research_th_list.Recordset.Close
Set research_th_list.Recordset = Nothing
%>
<% If research_th.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If research_th.CurrentAction <> "gridadd" And research_th.CurrentAction <> "gridedit" Then %>
<form name="ewPagerForm" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewPager">
<tr><td>
<% If Not IsObject(research_th_list.Pager) Then Set research_th_list.Pager = ew_NewPrevNextPager(research_th_list.StartRec, research_th_list.DisplayRecs, research_th_list.TotalRecs) %>
<% If research_th_list.Pager.RecordCount > 0 Then %>
<table class="ewStdTable"><tbody><tr><td>
	<%= Language.Phrase("Page") %>&nbsp;
<div class="input-prepend input-append">
<!--first page button-->
	<% If research_th_list.Pager.FirstButton.Enabled Then %>
	<a class="btn btn-small" href="<%= research_th_list.PageUrl %>start=<%= research_th_list.Pager.FirstButton.Start %>"><i class="icon-step-backward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-backward"></i></a>
	<% End If %>
<!--previous page button-->
	<% If research_th_list.Pager.PrevButton.Enabled Then %>
	<a class="btn btn-small" href="<%= research_th_list.PageUrl %>start=<%= research_th_list.Pager.PrevButton.Start %>"><i class="icon-prev"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-prev"></i></a>
	<% End If %>
<!--current page number-->
	<input class="input-mini" type="text" name="<%= EW_TABLE_PAGE_NO %>" value="<%= research_th_list.Pager.CurrentPage %>">
<!--next page button-->
	<% If research_th_list.Pager.NextButton.Enabled Then %>
	<a class="btn btn-small" href="<%= research_th_list.PageUrl %>start=<%= research_th_list.Pager.NextButton.Start %>"><i class="icon-play"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-play"></i></a>
	<% End If %>
<!--last page button-->
	<% If research_th_list.Pager.LastButton.Enabled Then %>
	<a class="btn btn-small" href="<%= research_th_list.PageUrl %>start=<%= research_th_list.Pager.LastButton.Start %>"><i class="icon-step-forward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-forward"></i></a>
	<% End If %>
</div>
	&nbsp;<%= Language.Phrase("of") %>&nbsp;<%= research_th_list.Pager.PageCount %>
</td>
<td>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<%= Language.Phrase("Record") %>&nbsp;<%= research_th_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= research_th_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= research_th_list.Pager.RecordCount %>
</td>
</tr></tbody></table>
<% Else %>
	<% If research_th_list.SearchWhere = "0=101" Then %>
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
	research_th_list.AddEditOptions.Render "body", "bottom", "", "", "", ""
	research_th_list.DetailOptions.Render "body", "bottom", "", "", "", ""
	research_th_list.ActionOptions.Render "body", "bottom", "", "", "", ""
%>
</div>
</div>
<% End If %>
</td></tr></table>
<% If research_th.Export = "" Then %>
<script type="text/javascript">
fresearch_thlistsrch.Init();
fresearch_thlist.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<% End If %>
<%
research_th_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If research_th.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set research_th_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cresearch_th_list

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
		TableName = "research_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "research_th_list"
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
		If research_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & research_th.TableVar & "&" ' add page token
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
		If research_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (research_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (research_th.TableVar = Request.QueryString("t"))
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
		FormName = "fresearch_thlist"
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
		If IsEmpty(research_th) Then Set research_th = New cresearch_th
		Set Table = research_th

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		ExportPdfUrl = PageUrl & "export=pdf"
		AddUrl = "pom_research_thadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "pom_research_thdelete.asp"
		MultiUpdateUrl = "pom_research_thupdate.asp"

		' Initialize other table object
		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "research_th"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' List options
		Set ListOptions = New cListOptions
		ListOptions.TableVar = research_th.TableVar

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = research_th.TableVar
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
				research_th.GridAddRowCount = gridaddcnt
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
		If UBound(research_th.CustomActions.CustomArray) >= 0 Then
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
		Set research_th = Nothing
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
			If research_th.Export = "" Then
				SetupBreadcrumb()
			End If

			' Hide list options
			If research_th.Export <> "" Then
				Call ListOptions.HideAllOptions(Array("sequence"))
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			ElseIf research_th.CurrentAction = "gridadd" Or research_th.CurrentAction = "gridedit" Then
				Call ListOptions.HideAllOptions(Array())
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			End If

			' Hide export options
			If research_th.Export <> "" Or research_th.CurrentAction <> "" Then
				ExportOptions.HideAllOptions(Array())
			End If

			' Hide other options
			If research_th.Export <> "" Then
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
			Call research_th.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If research_th.RecordsPerPage <> "" Then
			DisplayRecs = research_th.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Load search default if no existing search criteria
		If Not CheckSearchParms() Then

			' Load basic search from default
			research_th.BasicSearch.Keyword = research_th.BasicSearch.KeywordDefault
			research_th.BasicSearch.SearchType = research_th.BasicSearch.SearchTypeDefault
			research_th.BasicSearch.setSearchType(research_th.BasicSearch.SearchTypeDefault)
			If research_th.BasicSearch.Keyword <> "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call research_th.Recordset_Searching(SearchWhere)

		' Save search criteria
		If Command = "search" And Not RestoreSearch Then
			research_th.SearchWhere = SearchWhere ' Save to Session
			StartRec = 1 ' Reset start record counter
			research_th.StartRecordNumber = StartRec
		Else
			SearchWhere = research_th.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		research_th.SessionWhere = sFilter
		research_th.CurrentFilter = ""
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
				sFilter = research_th.KeyFilter
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
			research_th.rsh_id.FormValue = arrKeyFlds(0)
			If Not IsNumeric(research_th.rsh_id.FormValue) Then
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
			Call BuildBasicSearchSQL(sWhere, research_th.rsh_img, Keyword)
			Call BuildBasicSearchSQL(sWhere, research_th.rsh_pdf, Keyword)
			Call BuildBasicSearchSQL(sWhere, research_th.rsh_category, Keyword)
			Call BuildBasicSearchSQL(sWhere, research_th.rsh_subject, Keyword)
			Call BuildBasicSearchSQL(sWhere, research_th.rsh_subject_th, Keyword)
			Call BuildBasicSearchSQL(sWhere, research_th.rsh_intro, Keyword)
			Call BuildBasicSearchSQL(sWhere, research_th.rsh_intro_th, Keyword)
			Call BuildBasicSearchSQL(sWhere, research_th.rsh_content, Keyword)
			Call BuildBasicSearchSQL(sWhere, research_th.rsh_content_th, Keyword)
			Call BuildBasicSearchSQL(sWhere, research_th.rsh_show, Keyword)
			Call BuildBasicSearchSQL(sWhere, research_th.rsh_show_home, Keyword)
			Call BuildBasicSearchSQL(sWhere, research_th.rsh_create, Keyword)
			Call BuildBasicSearchSQL(sWhere, research_th.rsh_update, Keyword)
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
		sSearchKeyword = research_th.BasicSearch.Keyword
		sSearchType = research_th.BasicSearch.SearchType
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
			research_th.BasicSearch.setKeyword(sSearchKeyword)
			research_th.BasicSearch.setSearchType(sSearchType)
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' Check if search parm exists
	Function CheckSearchParms()

		' Check basic search
		If research_th.BasicSearch.IssetSession() Then
			CheckSearchParms = True
			Exit Function
		End If
		CheckSearchParms = False
	End Function

	' Clear all search parameters
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		research_th.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' Load advanced search default values
	Function LoadAdvancedSearchDefault()
		LoadAdvancedSearchDefault = False
	End Function

	' Clear all basic search parameters
	Sub ResetBasicSearchParms()
		research_th.BasicSearch.UnsetSession()
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()

		' Restore search flag
		RestoreSearch = True

		' Restore basic search values
		Call research_th.BasicSearch.Load()
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
			research_th.CurrentOrder = Request.QueryString("order")
			research_th.CurrentOrderType = Request.QueryString("ordertype")

			' Field rsh_id
			Call research_th.UpdateSort(research_th.rsh_id)

			' Field rsh_img
			Call research_th.UpdateSort(research_th.rsh_img)

			' Field rsh_date
			Call research_th.UpdateSort(research_th.rsh_date)

			' Field rsh_pdf
			Call research_th.UpdateSort(research_th.rsh_pdf)

			' Field rsh_category
			Call research_th.UpdateSort(research_th.rsh_category)

			' Field rsh_subject
			Call research_th.UpdateSort(research_th.rsh_subject)

			' Field rsh_subject_th
			Call research_th.UpdateSort(research_th.rsh_subject_th)

			' Field rsh_intro_th
			Call research_th.UpdateSort(research_th.rsh_intro_th)

			' Field rsh_show
			Call research_th.UpdateSort(research_th.rsh_show)

			' Field rsh_show_home
			Call research_th.UpdateSort(research_th.rsh_show_home)

			' Field rsh_create
			Call research_th.UpdateSort(research_th.rsh_create)

			' Field rsh_update
			Call research_th.UpdateSort(research_th.rsh_update)
			research_th.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = research_th.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If research_th.SqlOrderBy <> "" Then
				sOrderBy = research_th.SqlOrderBy
				research_th.SessionOrderBy = sOrderBy
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
				research_th.SessionOrderBy = sOrderBy
				research_th.rsh_id.Sort = ""
				research_th.rsh_img.Sort = ""
				research_th.rsh_date.Sort = ""
				research_th.rsh_pdf.Sort = ""
				research_th.rsh_category.Sort = ""
				research_th.rsh_subject.Sort = ""
				research_th.rsh_subject_th.Sort = ""
				research_th.rsh_intro_th.Sort = ""
				research_th.rsh_show.Sort = ""
				research_th.rsh_show_home.Sort = ""
				research_th.rsh_create.Sort = ""
				research_th.rsh_update.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			research_th.StartRecordNumber = StartRec
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
		ListOptions.GetItem("checkbox").Body = "<label class=""checkbox""><input type=""checkbox"" name=""key_m"" value=""" & ew_HtmlEncode(research_th.rsh_id.CurrentValue) & """ onclick='ew_ClickMultiCheckbox(event, this);'></label>"
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
			For i = 0 to UBound(research_th.CustomActions.CustomArray)
				Action = research_th.CustomActions.CustomArray(i)(0)
				Name = research_th.CustomActions.CustomArray(i)(1)

				' Add custom action
				Call opt.Add("custom_" & Action)
				Set item = opt.GetItem("custom_" & Action)
				item.Body = "<a class=""ewAction ewCustomAction"" href="""" onclick=""ew_SubmitSelected(document.fresearch_thlist, '" & ew_CurrentUrl & "', null, '" & Action & "');return false;"">" & Name & "</a>"
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
		sFilter = research_th.GetKeyFilter
		UserAction = Request.Form("useraction") & ""
		Processed = False
		If sFilter <> "" And UserAction <> "" Then
			research_th.CurrentFilter = sFilter
			sSql = research_th.SQL
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
				ElseIf research_th.CancelMessage <> "" Then
					FailureMessage = research_th.CancelMessage
					research_th.CancelMessage = ""
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
				research_th.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					research_th.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = research_th.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			research_th.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			research_th.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			research_th.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		research_th.BasicSearch.Keyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)&""
		If research_th.BasicSearch.Keyword <> "" Then Command = "search"
		research_th.BasicSearch.SearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)&""
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = research_th.CurrentFilter
		Call research_th.Recordset_Selecting(sFilter)
		research_th.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = research_th.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call research_th.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = research_th.KeyFilter

		' Call Row Selecting event
		Call research_th.Row_Selecting(sFilter)

		' Load sql based on filter
		research_th.CurrentFilter = sFilter
		sSql = research_th.SQL
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
		Call research_th.Row_Selected(RsRow)
		research_th.rsh_id.DbValue = RsRow("rsh_id")
		research_th.rsh_img.DbValue = RsRow("rsh_img")
		research_th.rsh_date.DbValue = RsRow("rsh_date")
		research_th.rsh_pdf.DbValue = RsRow("rsh_pdf")
		research_th.rsh_category.DbValue = RsRow("rsh_category")
		research_th.rsh_subject.DbValue = RsRow("rsh_subject")
		research_th.rsh_subject_th.DbValue = RsRow("rsh_subject_th")
		research_th.rsh_intro.DbValue = RsRow("rsh_intro")
		research_th.rsh_intro_th.DbValue = RsRow("rsh_intro_th")
		research_th.rsh_content.DbValue = RsRow("rsh_content")
		research_th.rsh_content_th.DbValue = RsRow("rsh_content_th")
		research_th.rsh_show.DbValue = RsRow("rsh_show")
		research_th.rsh_show_home.DbValue = RsRow("rsh_show_home")
		research_th.rsh_create.DbValue = RsRow("rsh_create")
		research_th.rsh_update.DbValue = RsRow("rsh_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		research_th.rsh_id.m_DbValue = Rs("rsh_id")
		research_th.rsh_img.m_DbValue = Rs("rsh_img")
		research_th.rsh_date.m_DbValue = Rs("rsh_date")
		research_th.rsh_pdf.m_DbValue = Rs("rsh_pdf")
		research_th.rsh_category.m_DbValue = Rs("rsh_category")
		research_th.rsh_subject.m_DbValue = Rs("rsh_subject")
		research_th.rsh_subject_th.m_DbValue = Rs("rsh_subject_th")
		research_th.rsh_intro.m_DbValue = Rs("rsh_intro")
		research_th.rsh_intro_th.m_DbValue = Rs("rsh_intro_th")
		research_th.rsh_content.m_DbValue = Rs("rsh_content")
		research_th.rsh_content_th.m_DbValue = Rs("rsh_content_th")
		research_th.rsh_show.m_DbValue = Rs("rsh_show")
		research_th.rsh_show_home.m_DbValue = Rs("rsh_show_home")
		research_th.rsh_create.m_DbValue = Rs("rsh_create")
		research_th.rsh_update.m_DbValue = Rs("rsh_update")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If research_th.GetKey("rsh_id")&"" <> "" Then
			research_th.rsh_id.CurrentValue = research_th.GetKey("rsh_id") ' rsh_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			research_th.CurrentFilter = research_th.KeyFilter
			Dim sSql
			sSql = research_th.SQL
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
		ViewUrl = research_th.ViewUrl("")
		EditUrl = research_th.EditUrl("")
		InlineEditUrl = research_th.InlineEditUrl
		CopyUrl = research_th.CopyUrl("")
		InlineCopyUrl = research_th.InlineCopyUrl
		DeleteUrl = research_th.DeleteUrl

		' Call Row Rendering event
		Call research_th.Row_Rendering()

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

		If research_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' rsh_id
			research_th.rsh_id.ViewValue = research_th.rsh_id.CurrentValue
			research_th.rsh_id.ViewCustomAttributes = ""

			' rsh_img
			research_th.rsh_img.ViewValue = research_th.rsh_img.CurrentValue
			research_th.rsh_img.ViewCustomAttributes = ""

			' rsh_date
			research_th.rsh_date.ViewValue = research_th.rsh_date.CurrentValue
			research_th.rsh_date.ViewCustomAttributes = ""

			' rsh_pdf
			research_th.rsh_pdf.ViewValue = research_th.rsh_pdf.CurrentValue
			research_th.rsh_pdf.ViewCustomAttributes = ""

			' rsh_category
			research_th.rsh_category.ViewValue = research_th.rsh_category.CurrentValue
			research_th.rsh_category.ViewCustomAttributes = ""

			' rsh_subject
			research_th.rsh_subject.ViewValue = research_th.rsh_subject.CurrentValue
			research_th.rsh_subject.ViewCustomAttributes = ""

			' rsh_subject_th
			research_th.rsh_subject_th.ViewValue = research_th.rsh_subject_th.CurrentValue
			research_th.rsh_subject_th.ViewCustomAttributes = ""

			' rsh_intro_th
			research_th.rsh_intro_th.ViewValue = research_th.rsh_intro_th.CurrentValue
			research_th.rsh_intro_th.ViewCustomAttributes = ""

			' rsh_show
			research_th.rsh_show.ViewValue = research_th.rsh_show.CurrentValue
			research_th.rsh_show.ViewCustomAttributes = ""

			' rsh_show_home
			research_th.rsh_show_home.ViewValue = research_th.rsh_show_home.CurrentValue
			research_th.rsh_show_home.ViewCustomAttributes = ""

			' rsh_create
			research_th.rsh_create.ViewValue = research_th.rsh_create.CurrentValue
			research_th.rsh_create.ViewCustomAttributes = ""

			' rsh_update
			research_th.rsh_update.ViewValue = research_th.rsh_update.CurrentValue
			research_th.rsh_update.ViewCustomAttributes = ""

			' View refer script
			' rsh_id

			research_th.rsh_id.LinkCustomAttributes = ""
			research_th.rsh_id.HrefValue = ""
			research_th.rsh_id.TooltipValue = ""

			' rsh_img
			research_th.rsh_img.LinkCustomAttributes = ""
			research_th.rsh_img.HrefValue = ""
			research_th.rsh_img.TooltipValue = ""

			' rsh_date
			research_th.rsh_date.LinkCustomAttributes = ""
			research_th.rsh_date.HrefValue = ""
			research_th.rsh_date.TooltipValue = ""

			' rsh_pdf
			research_th.rsh_pdf.LinkCustomAttributes = ""
			research_th.rsh_pdf.HrefValue = ""
			research_th.rsh_pdf.TooltipValue = ""

			' rsh_category
			research_th.rsh_category.LinkCustomAttributes = ""
			research_th.rsh_category.HrefValue = ""
			research_th.rsh_category.TooltipValue = ""

			' rsh_subject
			research_th.rsh_subject.LinkCustomAttributes = ""
			research_th.rsh_subject.HrefValue = ""
			research_th.rsh_subject.TooltipValue = ""

			' rsh_subject_th
			research_th.rsh_subject_th.LinkCustomAttributes = ""
			research_th.rsh_subject_th.HrefValue = ""
			research_th.rsh_subject_th.TooltipValue = ""

			' rsh_intro_th
			research_th.rsh_intro_th.LinkCustomAttributes = ""
			research_th.rsh_intro_th.HrefValue = ""
			research_th.rsh_intro_th.TooltipValue = ""

			' rsh_show
			research_th.rsh_show.LinkCustomAttributes = ""
			research_th.rsh_show.HrefValue = ""
			research_th.rsh_show.TooltipValue = ""

			' rsh_show_home
			research_th.rsh_show_home.LinkCustomAttributes = ""
			research_th.rsh_show_home.HrefValue = ""
			research_th.rsh_show_home.TooltipValue = ""

			' rsh_create
			research_th.rsh_create.LinkCustomAttributes = ""
			research_th.rsh_create.HrefValue = ""
			research_th.rsh_create.TooltipValue = ""

			' rsh_update
			research_th.rsh_update.LinkCustomAttributes = ""
			research_th.rsh_update.HrefValue = ""
			research_th.rsh_update.TooltipValue = ""
		End If

		' Call Row Rendered event
		If research_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call research_th.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		url = ew_CurrentUrl
		url = ew_RegExReplace("\?cmd=reset(all){0,1}$", url, "") ' Remove cmd=reset / cmd=resetall
		Call Breadcrumb.Add("list", research_th.TableVar, url, research_th.TableVar, True)
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
