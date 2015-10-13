<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_journalinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim journal_list
Set journal_list = New cjournal_list
Set Page = journal_list

' Page init processing
journal_list.Page_Init()

' Page main processing
journal_list.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
journal_list.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If journal.Export = "" Then %>
<script type="text/javascript">
// Page object
var journal_list = new ew_Page("journal_list");
journal_list.PageID = "list"; // Page ID
var EW_PAGE_ID = journal_list.PageID; // For backward compatibility
// Form object
var fjournallist = new ew_Form("fjournallist");
fjournallist.FormKeyCountName = '<%= journal_list.FormKeyCountName %>';
// Form_CustomValidate event
fjournallist.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fjournallist.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fjournallist.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
var fjournallistsrch = new ew_Form("fjournallistsrch");
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If journal.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If journal_list.ExportOptions.Visible Then %>
<div class="ewListExportOptions"><% journal_list.ExportOptions.Render "body", "", "", "", "", "" %></div>
<% End If %>
<% If (journal.Export = "") Or (EW_EXPORT_MASTER_RECORD And journal.Export = "print") Then %>
<% End If %>
<%

' Load recordset
Set journal_list.Recordset = journal_list.LoadRecordset()
	journal_list.TotalRecs = journal_list.Recordset.RecordCount
	journal_list.StartRec = 1
	If journal_list.DisplayRecs <= 0 Then ' Display all records
		journal_list.DisplayRecs = journal_list.TotalRecs
	End If
	If Not (journal.ExportAll And journal.Export <> "") Then
		journal_list.SetUpStartRec() ' Set up start record position
	End If
journal_list.RenderOtherOptions()
%>
<% If Security.IsLoggedIn() Then %>
<% If journal.Export = "" And journal.CurrentAction = "" Then %>
<form name="fjournallistsrch" id="fjournallistsrch" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewSearchTable"><tr><td>
<div class="accordion" id="fjournallistsrch_SearchGroup">
	<div class="accordion-group">
		<div class="accordion-heading">
<a class="accordion-toggle" data-toggle="collapse" data-parent="#fjournallistsrch_SearchGroup" href="#fjournallistsrch_SearchBody"><%= Language.Phrase("Search") %></a>
		</div>
		<div id="fjournallistsrch_SearchBody" class="accordion-body collapse in">
			<div class="accordion-inner">
<div id="fjournallistsrch_SearchPanel">
<input type="hidden" name="cmd" value="search">
<input type="hidden" name="t" value="journal">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewRow">
	<div class="btn-group ewButtonGroup">
	<div class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" class="input-large" value="<%= ew_HtmlEncode(journal.BasicSearch.getKeyword()) %>" placeholder="<%= ew_HtmlEncode(Language.Phrase("Search")) %>">
	<button class="btn btn-primary ewButton" name="btnsubmit" id="btnsubmit" type="submit"><%= Language.Phrase("QuickSearchBtn") %></button>
	</div>
	</div>
	<div class="btn-group ewButtonGroup">
	<a class="btn ewShowAll" href="<%= journal_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>
	</div>
</div>
<div id="xsr_2" class="ewRow">
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="="<% If journal.BasicSearch.getSearchType = "=" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If journal.BasicSearch.getSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If journal.BasicSearch.getSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
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
<% journal_list.ShowPageHeader() %>
<% journal_list.ShowMessage %>
<table class="ewGrid"><tr><td class="ewGridContent">
<form name="fjournallist" id="fjournallist" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="journal">
<div id="gmp_journal" class="ewGridMiddlePanel">
<% If journal_list.TotalRecs > 0 Then %>
<table id="tbl_journallist" class="ewTable ewTableSeparate">
<%= journal.TableCustomInnerHtml %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call journal_list.RenderListOptions()

' Render list options (header, left)
journal_list.ListOptions.Render "header", "left", "", "", "", ""
%>
<% If journal.jrl_id.Visible Then ' jrl_id %>
	<% If journal.SortUrl(journal.jrl_id) = "" Then %>
		<td><div id="elh_journal_jrl_id" class="journal_jrl_id"><div class="ewTableHeaderCaption"><%= journal.jrl_id.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= journal.SortUrl(journal.jrl_id) %>',1);"><div id="elh_journal_jrl_id" class="journal_jrl_id">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= journal.jrl_id.FldCaption %></span><span class="ewTableHeaderSort"><% If journal.jrl_id.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf journal.jrl_id.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If journal.jrl_category.Visible Then ' jrl_category %>
	<% If journal.SortUrl(journal.jrl_category) = "" Then %>
		<td><div id="elh_journal_jrl_category" class="journal_jrl_category"><div class="ewTableHeaderCaption"><%= journal.jrl_category.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= journal.SortUrl(journal.jrl_category) %>',1);"><div id="elh_journal_jrl_category" class="journal_jrl_category">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= journal.jrl_category.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If journal.jrl_category.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf journal.jrl_category.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If journal.jrl_date.Visible Then ' jrl_date %>
	<% If journal.SortUrl(journal.jrl_date) = "" Then %>
		<td><div id="elh_journal_jrl_date" class="journal_jrl_date"><div class="ewTableHeaderCaption"><%= journal.jrl_date.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= journal.SortUrl(journal.jrl_date) %>',1);"><div id="elh_journal_jrl_date" class="journal_jrl_date">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= journal.jrl_date.FldCaption %></span><span class="ewTableHeaderSort"><% If journal.jrl_date.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf journal.jrl_date.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If journal.jrl_title.Visible Then ' jrl_title %>
	<% If journal.SortUrl(journal.jrl_title) = "" Then %>
		<td><div id="elh_journal_jrl_title" class="journal_jrl_title"><div class="ewTableHeaderCaption"><%= journal.jrl_title.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= journal.SortUrl(journal.jrl_title) %>',1);"><div id="elh_journal_jrl_title" class="journal_jrl_title">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= journal.jrl_title.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If journal.jrl_title.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf journal.jrl_title.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If journal.jrl_title_th.Visible Then ' jrl_title_th %>
	<% If journal.SortUrl(journal.jrl_title_th) = "" Then %>
		<td><div id="elh_journal_jrl_title_th" class="journal_jrl_title_th"><div class="ewTableHeaderCaption"><%= journal.jrl_title_th.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= journal.SortUrl(journal.jrl_title_th) %>',1);"><div id="elh_journal_jrl_title_th" class="journal_jrl_title_th">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= journal.jrl_title_th.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If journal.jrl_title_th.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf journal.jrl_title_th.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If journal.jrl_pdf.Visible Then ' jrl_pdf %>
	<% If journal.SortUrl(journal.jrl_pdf) = "" Then %>
		<td><div id="elh_journal_jrl_pdf" class="journal_jrl_pdf"><div class="ewTableHeaderCaption"><%= journal.jrl_pdf.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= journal.SortUrl(journal.jrl_pdf) %>',1);"><div id="elh_journal_jrl_pdf" class="journal_jrl_pdf">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= journal.jrl_pdf.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If journal.jrl_pdf.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf journal.jrl_pdf.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If journal.jrl_img.Visible Then ' jrl_img %>
	<% If journal.SortUrl(journal.jrl_img) = "" Then %>
		<td><div id="elh_journal_jrl_img" class="journal_jrl_img"><div class="ewTableHeaderCaption"><%= journal.jrl_img.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= journal.SortUrl(journal.jrl_img) %>',1);"><div id="elh_journal_jrl_img" class="journal_jrl_img">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= journal.jrl_img.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If journal.jrl_img.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf journal.jrl_img.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If journal.jrl_create.Visible Then ' jrl_create %>
	<% If journal.SortUrl(journal.jrl_create) = "" Then %>
		<td><div id="elh_journal_jrl_create" class="journal_jrl_create"><div class="ewTableHeaderCaption"><%= journal.jrl_create.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= journal.SortUrl(journal.jrl_create) %>',1);"><div id="elh_journal_jrl_create" class="journal_jrl_create">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= journal.jrl_create.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If journal.jrl_create.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf journal.jrl_create.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If journal.jrl_update.Visible Then ' jrl_update %>
	<% If journal.SortUrl(journal.jrl_update) = "" Then %>
		<td><div id="elh_journal_jrl_update" class="journal_jrl_update"><div class="ewTableHeaderCaption"><%= journal.jrl_update.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= journal.SortUrl(journal.jrl_update) %>',1);"><div id="elh_journal_jrl_update" class="journal_jrl_update">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= journal.jrl_update.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If journal.jrl_update.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf journal.jrl_update.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
journal_list.ListOptions.Render "header", "right", "", "", "", ""
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (journal.ExportAll And journal.Export <> "") Then
	journal_list.StopRec = journal_list.TotalRecs
Else

	' Set the last record to display
	If journal_list.TotalRecs > journal_list.StartRec + journal_list.DisplayRecs - 1 Then
		journal_list.StopRec = journal_list.StartRec + journal_list.DisplayRecs - 1
	Else
		journal_list.StopRec = journal_list.TotalRecs
	End If
End If

' Move to first record
journal_list.RecCnt = journal_list.StartRec - 1
If Not journal_list.Recordset.Eof Then
	journal_list.Recordset.MoveFirst
	If journal_list.StartRec > 1 Then journal_list.Recordset.Move journal_list.StartRec - 1
ElseIf Not journal.AllowAddDeleteRow And journal_list.StopRec = 0 Then
	journal_list.StopRec = journal.GridAddRowCount
End If

' Initialize Aggregate
journal.RowType = EW_ROWTYPE_AGGREGATEINIT
Call journal.ResetAttrs()
Call journal_list.RenderRow()
journal_list.RowCnt = 0

' Output date rows
Do While CLng(journal_list.RecCnt) < CLng(journal_list.StopRec)
	journal_list.RecCnt = journal_list.RecCnt + 1
	If CLng(journal_list.RecCnt) >= CLng(journal_list.StartRec) Then
		journal_list.RowCnt = journal_list.RowCnt + 1

	' Set up key count
	journal_list.KeyCount = journal_list.RowIndex
	Call journal.ResetAttrs()
	journal.CssClass = ""
	If journal.CurrentAction = "gridadd" Then
	Else
		Call journal_list.LoadRowValues(journal_list.Recordset) ' Load row values
	End If
	journal.RowType = EW_ROWTYPE_VIEW ' Render view

	' Set up row id / data-rowindex
	journal.RowAttrs.AddAttributes Array(Array("data-rowindex", journal_list.RowCnt), Array("id", "r" & journal_list.RowCnt & "_journal"), Array("data-rowtype", journal.RowType))

	' Render row
	Call journal_list.RenderRow()

	' Render list options
	Call journal_list.RenderListOptions()
%>
	<tr<%= journal.RowAttributes %>>
<%

' Render list options (body, left)
journal_list.ListOptions.Render "body", "left", journal_list.RowCnt, "", "", ""
%>
	<% If journal.jrl_id.Visible Then ' jrl_id %>
		<td<%= journal.jrl_id.CellAttributes %>>
<span<%= journal.jrl_id.ViewAttributes %>>
<%= journal.jrl_id.ListViewValue %>
</span>
<a id="<%= journal_list.PageObjName & "_row_" & journal_list.RowCnt %>"></a></td>
	<% End If %>
	<% If journal.jrl_category.Visible Then ' jrl_category %>
		<td<%= journal.jrl_category.CellAttributes %>>
<span<%= journal.jrl_category.ViewAttributes %>>
<%= journal.jrl_category.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If journal.jrl_date.Visible Then ' jrl_date %>
		<td<%= journal.jrl_date.CellAttributes %>>
<span<%= journal.jrl_date.ViewAttributes %>>
<%= journal.jrl_date.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If journal.jrl_title.Visible Then ' jrl_title %>
		<td<%= journal.jrl_title.CellAttributes %>>
<span<%= journal.jrl_title.ViewAttributes %>>
<%= journal.jrl_title.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If journal.jrl_title_th.Visible Then ' jrl_title_th %>
		<td<%= journal.jrl_title_th.CellAttributes %>>
<span<%= journal.jrl_title_th.ViewAttributes %>>
<%= journal.jrl_title_th.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If journal.jrl_pdf.Visible Then ' jrl_pdf %>
		<td<%= journal.jrl_pdf.CellAttributes %>>
<span<%= journal.jrl_pdf.ViewAttributes %>>
<%= journal.jrl_pdf.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If journal.jrl_img.Visible Then ' jrl_img %>
		<td<%= journal.jrl_img.CellAttributes %>>
<span<%= journal.jrl_img.ViewAttributes %>>
<%= journal.jrl_img.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If journal.jrl_create.Visible Then ' jrl_create %>
		<td<%= journal.jrl_create.CellAttributes %>>
<span<%= journal.jrl_create.ViewAttributes %>>
<%= journal.jrl_create.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If journal.jrl_update.Visible Then ' jrl_update %>
		<td<%= journal.jrl_update.CellAttributes %>>
<span<%= journal.jrl_update.ViewAttributes %>>
<%= journal.jrl_update.ListViewValue %>
</span>
</td>
	<% End If %>
<%

' Render list options (body, right)
journal_list.ListOptions.Render "body", "right", journal_list.RowCnt, "", "", ""
%>
	</tr>
<%
	End If
	If journal.CurrentAction <> "gridadd" Then
		journal_list.Recordset.MoveNext()
	End If
Loop
%>
</tbody>
</table>
<% End If %>
<% If journal.CurrentAction = "" Then %>
<input type="hidden" name="a_list" id="a_list" value="">
<% End If %>
</div>
</form>
<%

' Close recordset and connection
journal_list.Recordset.Close
Set journal_list.Recordset = Nothing
%>
<% If journal.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If journal.CurrentAction <> "gridadd" And journal.CurrentAction <> "gridedit" Then %>
<form name="ewPagerForm" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewPager">
<tr><td>
<% If Not IsObject(journal_list.Pager) Then Set journal_list.Pager = ew_NewPrevNextPager(journal_list.StartRec, journal_list.DisplayRecs, journal_list.TotalRecs) %>
<% If journal_list.Pager.RecordCount > 0 Then %>
<table class="ewStdTable"><tbody><tr><td>
	<%= Language.Phrase("Page") %>&nbsp;
<div class="input-prepend input-append">
<!--first page button-->
	<% If journal_list.Pager.FirstButton.Enabled Then %>
	<a class="btn btn-small" href="<%= journal_list.PageUrl %>start=<%= journal_list.Pager.FirstButton.Start %>"><i class="icon-step-backward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-backward"></i></a>
	<% End If %>
<!--previous page button-->
	<% If journal_list.Pager.PrevButton.Enabled Then %>
	<a class="btn btn-small" href="<%= journal_list.PageUrl %>start=<%= journal_list.Pager.PrevButton.Start %>"><i class="icon-prev"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-prev"></i></a>
	<% End If %>
<!--current page number-->
	<input class="input-mini" type="text" name="<%= EW_TABLE_PAGE_NO %>" value="<%= journal_list.Pager.CurrentPage %>">
<!--next page button-->
	<% If journal_list.Pager.NextButton.Enabled Then %>
	<a class="btn btn-small" href="<%= journal_list.PageUrl %>start=<%= journal_list.Pager.NextButton.Start %>"><i class="icon-play"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-play"></i></a>
	<% End If %>
<!--last page button-->
	<% If journal_list.Pager.LastButton.Enabled Then %>
	<a class="btn btn-small" href="<%= journal_list.PageUrl %>start=<%= journal_list.Pager.LastButton.Start %>"><i class="icon-step-forward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-forward"></i></a>
	<% End If %>
</div>
	&nbsp;<%= Language.Phrase("of") %>&nbsp;<%= journal_list.Pager.PageCount %>
</td>
<td>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<%= Language.Phrase("Record") %>&nbsp;<%= journal_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= journal_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= journal_list.Pager.RecordCount %>
</td>
</tr></tbody></table>
<% Else %>
	<% If journal_list.SearchWhere = "0=101" Then %>
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
	journal_list.AddEditOptions.Render "body", "bottom", "", "", "", ""
	journal_list.DetailOptions.Render "body", "bottom", "", "", "", ""
	journal_list.ActionOptions.Render "body", "bottom", "", "", "", ""
%>
</div>
</div>
<% End If %>
</td></tr></table>
<% If journal.Export = "" Then %>
<script type="text/javascript">
fjournallistsrch.Init();
fjournallist.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<% End If %>
<%
journal_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If journal.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set journal_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cjournal_list

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
		TableName = "journal"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "journal_list"
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
		If journal.UseTokenInUrl Then PageUrl = PageUrl & "t=" & journal.TableVar & "&" ' add page token
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
		If journal.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (journal.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (journal.TableVar = Request.QueryString("t"))
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
		FormName = "fjournallist"
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
		If IsEmpty(journal) Then Set journal = New cjournal
		Set Table = journal

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		ExportPdfUrl = PageUrl & "export=pdf"
		AddUrl = "pom_journaladd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "pom_journaldelete.asp"
		MultiUpdateUrl = "pom_journalupdate.asp"

		' Initialize other table object
		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "journal"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' List options
		Set ListOptions = New cListOptions
		ListOptions.TableVar = journal.TableVar

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = journal.TableVar
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
				journal.GridAddRowCount = gridaddcnt
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
		If UBound(journal.CustomActions.CustomArray) >= 0 Then
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
		Set journal = Nothing
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
			If journal.Export = "" Then
				SetupBreadcrumb()
			End If

			' Hide list options
			If journal.Export <> "" Then
				Call ListOptions.HideAllOptions(Array("sequence"))
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			ElseIf journal.CurrentAction = "gridadd" Or journal.CurrentAction = "gridedit" Then
				Call ListOptions.HideAllOptions(Array())
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			End If

			' Hide export options
			If journal.Export <> "" Or journal.CurrentAction <> "" Then
				ExportOptions.HideAllOptions(Array())
			End If

			' Hide other options
			If journal.Export <> "" Then
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
			Call journal.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If journal.RecordsPerPage <> "" Then
			DisplayRecs = journal.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Load search default if no existing search criteria
		If Not CheckSearchParms() Then

			' Load basic search from default
			journal.BasicSearch.Keyword = journal.BasicSearch.KeywordDefault
			journal.BasicSearch.SearchType = journal.BasicSearch.SearchTypeDefault
			journal.BasicSearch.setSearchType(journal.BasicSearch.SearchTypeDefault)
			If journal.BasicSearch.Keyword <> "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call journal.Recordset_Searching(SearchWhere)

		' Save search criteria
		If Command = "search" And Not RestoreSearch Then
			journal.SearchWhere = SearchWhere ' Save to Session
			StartRec = 1 ' Reset start record counter
			journal.StartRecordNumber = StartRec
		Else
			SearchWhere = journal.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		journal.SessionWhere = sFilter
		journal.CurrentFilter = ""
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
				sFilter = journal.KeyFilter
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
			journal.jrl_id.FormValue = arrKeyFlds(0)
			If Not IsNumeric(journal.jrl_id.FormValue) Then
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
			Call BuildBasicSearchSQL(sWhere, journal.jrl_category, Keyword)
			Call BuildBasicSearchSQL(sWhere, journal.jrl_title, Keyword)
			Call BuildBasicSearchSQL(sWhere, journal.jrl_title_th, Keyword)
			Call BuildBasicSearchSQL(sWhere, journal.jrl_pdf, Keyword)
			Call BuildBasicSearchSQL(sWhere, journal.jrl_img, Keyword)
			Call BuildBasicSearchSQL(sWhere, journal.jrl_create, Keyword)
			Call BuildBasicSearchSQL(sWhere, journal.jrl_update, Keyword)
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
		sSearchKeyword = journal.BasicSearch.Keyword
		sSearchType = journal.BasicSearch.SearchType
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
			journal.BasicSearch.setKeyword(sSearchKeyword)
			journal.BasicSearch.setSearchType(sSearchType)
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' Check if search parm exists
	Function CheckSearchParms()

		' Check basic search
		If journal.BasicSearch.IssetSession() Then
			CheckSearchParms = True
			Exit Function
		End If
		CheckSearchParms = False
	End Function

	' Clear all search parameters
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		journal.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' Load advanced search default values
	Function LoadAdvancedSearchDefault()
		LoadAdvancedSearchDefault = False
	End Function

	' Clear all basic search parameters
	Sub ResetBasicSearchParms()
		journal.BasicSearch.UnsetSession()
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()

		' Restore search flag
		RestoreSearch = True

		' Restore basic search values
		Call journal.BasicSearch.Load()
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
			journal.CurrentOrder = Request.QueryString("order")
			journal.CurrentOrderType = Request.QueryString("ordertype")

			' Field jrl_id
			Call journal.UpdateSort(journal.jrl_id)

			' Field jrl_category
			Call journal.UpdateSort(journal.jrl_category)

			' Field jrl_date
			Call journal.UpdateSort(journal.jrl_date)

			' Field jrl_title
			Call journal.UpdateSort(journal.jrl_title)

			' Field jrl_title_th
			Call journal.UpdateSort(journal.jrl_title_th)

			' Field jrl_pdf
			Call journal.UpdateSort(journal.jrl_pdf)

			' Field jrl_img
			Call journal.UpdateSort(journal.jrl_img)

			' Field jrl_create
			Call journal.UpdateSort(journal.jrl_create)

			' Field jrl_update
			Call journal.UpdateSort(journal.jrl_update)
			journal.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = journal.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If journal.SqlOrderBy <> "" Then
				sOrderBy = journal.SqlOrderBy
				journal.SessionOrderBy = sOrderBy
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
				journal.SessionOrderBy = sOrderBy
				journal.jrl_id.Sort = ""
				journal.jrl_category.Sort = ""
				journal.jrl_date.Sort = ""
				journal.jrl_title.Sort = ""
				journal.jrl_title_th.Sort = ""
				journal.jrl_pdf.Sort = ""
				journal.jrl_img.Sort = ""
				journal.jrl_create.Sort = ""
				journal.jrl_update.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			journal.StartRecordNumber = StartRec
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
		ListOptions.GetItem("checkbox").Body = "<label class=""checkbox""><input type=""checkbox"" name=""key_m"" value=""" & ew_HtmlEncode(journal.jrl_id.CurrentValue) & """ onclick='ew_ClickMultiCheckbox(event, this);'></label>"
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
			For i = 0 to UBound(journal.CustomActions.CustomArray)
				Action = journal.CustomActions.CustomArray(i)(0)
				Name = journal.CustomActions.CustomArray(i)(1)

				' Add custom action
				Call opt.Add("custom_" & Action)
				Set item = opt.GetItem("custom_" & Action)
				item.Body = "<a class=""ewAction ewCustomAction"" href="""" onclick=""ew_SubmitSelected(document.fjournallist, '" & ew_CurrentUrl & "', null, '" & Action & "');return false;"">" & Name & "</a>"
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
		sFilter = journal.GetKeyFilter
		UserAction = Request.Form("useraction") & ""
		Processed = False
		If sFilter <> "" And UserAction <> "" Then
			journal.CurrentFilter = sFilter
			sSql = journal.SQL
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
				ElseIf journal.CancelMessage <> "" Then
					FailureMessage = journal.CancelMessage
					journal.CancelMessage = ""
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
				journal.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					journal.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = journal.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			journal.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			journal.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			journal.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		journal.BasicSearch.Keyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)&""
		If journal.BasicSearch.Keyword <> "" Then Command = "search"
		journal.BasicSearch.SearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)&""
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = journal.CurrentFilter
		Call journal.Recordset_Selecting(sFilter)
		journal.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = journal.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call journal.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = journal.KeyFilter

		' Call Row Selecting event
		Call journal.Row_Selecting(sFilter)

		' Load sql based on filter
		journal.CurrentFilter = sFilter
		sSql = journal.SQL
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
		Call journal.Row_Selected(RsRow)
		journal.jrl_id.DbValue = RsRow("jrl_id")
		journal.jrl_category.DbValue = RsRow("jrl_category")
		journal.jrl_date.DbValue = RsRow("jrl_date")
		journal.jrl_title.DbValue = RsRow("jrl_title")
		journal.jrl_title_th.DbValue = RsRow("jrl_title_th")
		journal.jrl_pdf.DbValue = RsRow("jrl_pdf")
		journal.jrl_img.DbValue = RsRow("jrl_img")
		journal.jrl_create.DbValue = RsRow("jrl_create")
		journal.jrl_update.DbValue = RsRow("jrl_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		journal.jrl_id.m_DbValue = Rs("jrl_id")
		journal.jrl_category.m_DbValue = Rs("jrl_category")
		journal.jrl_date.m_DbValue = Rs("jrl_date")
		journal.jrl_title.m_DbValue = Rs("jrl_title")
		journal.jrl_title_th.m_DbValue = Rs("jrl_title_th")
		journal.jrl_pdf.m_DbValue = Rs("jrl_pdf")
		journal.jrl_img.m_DbValue = Rs("jrl_img")
		journal.jrl_create.m_DbValue = Rs("jrl_create")
		journal.jrl_update.m_DbValue = Rs("jrl_update")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If journal.GetKey("jrl_id")&"" <> "" Then
			journal.jrl_id.CurrentValue = journal.GetKey("jrl_id") ' jrl_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			journal.CurrentFilter = journal.KeyFilter
			Dim sSql
			sSql = journal.SQL
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
		ViewUrl = journal.ViewUrl("")
		EditUrl = journal.EditUrl("")
		InlineEditUrl = journal.InlineEditUrl
		CopyUrl = journal.CopyUrl("")
		InlineCopyUrl = journal.InlineCopyUrl
		DeleteUrl = journal.DeleteUrl

		' Call Row Rendering event
		Call journal.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' jrl_id
		' jrl_category
		' jrl_date
		' jrl_title
		' jrl_title_th
		' jrl_pdf
		' jrl_img
		' jrl_create
		' jrl_update
		' -----------
		'  View  Row
		' -----------

		If journal.RowType = EW_ROWTYPE_VIEW Then ' View row

			' jrl_id
			journal.jrl_id.ViewValue = journal.jrl_id.CurrentValue
			journal.jrl_id.ViewCustomAttributes = ""

			' jrl_category
			journal.jrl_category.ViewValue = journal.jrl_category.CurrentValue
			journal.jrl_category.ViewCustomAttributes = ""

			' jrl_date
			journal.jrl_date.ViewValue = journal.jrl_date.CurrentValue
			journal.jrl_date.ViewCustomAttributes = ""

			' jrl_title
			journal.jrl_title.ViewValue = journal.jrl_title.CurrentValue
			journal.jrl_title.ViewCustomAttributes = ""

			' jrl_title_th
			journal.jrl_title_th.ViewValue = journal.jrl_title_th.CurrentValue
			journal.jrl_title_th.ViewCustomAttributes = ""

			' jrl_pdf
			journal.jrl_pdf.ViewValue = journal.jrl_pdf.CurrentValue
			journal.jrl_pdf.ViewCustomAttributes = ""

			' jrl_img
			journal.jrl_img.ViewValue = journal.jrl_img.CurrentValue
			journal.jrl_img.ViewCustomAttributes = ""

			' jrl_create
			journal.jrl_create.ViewValue = journal.jrl_create.CurrentValue
			journal.jrl_create.ViewCustomAttributes = ""

			' jrl_update
			journal.jrl_update.ViewValue = journal.jrl_update.CurrentValue
			journal.jrl_update.ViewCustomAttributes = ""

			' View refer script
			' jrl_id

			journal.jrl_id.LinkCustomAttributes = ""
			journal.jrl_id.HrefValue = ""
			journal.jrl_id.TooltipValue = ""

			' jrl_category
			journal.jrl_category.LinkCustomAttributes = ""
			journal.jrl_category.HrefValue = ""
			journal.jrl_category.TooltipValue = ""

			' jrl_date
			journal.jrl_date.LinkCustomAttributes = ""
			journal.jrl_date.HrefValue = ""
			journal.jrl_date.TooltipValue = ""

			' jrl_title
			journal.jrl_title.LinkCustomAttributes = ""
			journal.jrl_title.HrefValue = ""
			journal.jrl_title.TooltipValue = ""

			' jrl_title_th
			journal.jrl_title_th.LinkCustomAttributes = ""
			journal.jrl_title_th.HrefValue = ""
			journal.jrl_title_th.TooltipValue = ""

			' jrl_pdf
			journal.jrl_pdf.LinkCustomAttributes = ""
			journal.jrl_pdf.HrefValue = ""
			journal.jrl_pdf.TooltipValue = ""

			' jrl_img
			journal.jrl_img.LinkCustomAttributes = ""
			journal.jrl_img.HrefValue = ""
			journal.jrl_img.TooltipValue = ""

			' jrl_create
			journal.jrl_create.LinkCustomAttributes = ""
			journal.jrl_create.HrefValue = ""
			journal.jrl_create.TooltipValue = ""

			' jrl_update
			journal.jrl_update.LinkCustomAttributes = ""
			journal.jrl_update.HrefValue = ""
			journal.jrl_update.TooltipValue = ""
		End If

		' Call Row Rendered event
		If journal.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call journal.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		url = ew_CurrentUrl
		url = ew_RegExReplace("\?cmd=reset(all){0,1}$", url, "") ' Remove cmd=reset / cmd=resetall
		Call Breadcrumb.Add("list", journal.TableVar, url, journal.TableVar, True)
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
