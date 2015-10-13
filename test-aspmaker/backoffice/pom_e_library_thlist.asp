<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_e_library_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim e_library_th_list
Set e_library_th_list = New ce_library_th_list
Set Page = e_library_th_list

' Page init processing
e_library_th_list.Page_Init()

' Page main processing
e_library_th_list.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
e_library_th_list.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If e_library_th.Export = "" Then %>
<script type="text/javascript">
// Page object
var e_library_th_list = new ew_Page("e_library_th_list");
e_library_th_list.PageID = "list"; // Page ID
var EW_PAGE_ID = e_library_th_list.PageID; // For backward compatibility
// Form object
var fe_library_thlist = new ew_Form("fe_library_thlist");
fe_library_thlist.FormKeyCountName = '<%= e_library_th_list.FormKeyCountName %>';
// Form_CustomValidate event
fe_library_thlist.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fe_library_thlist.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fe_library_thlist.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
var fe_library_thlistsrch = new ew_Form("fe_library_thlistsrch");
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If e_library_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If e_library_th_list.ExportOptions.Visible Then %>
<div class="ewListExportOptions"><% e_library_th_list.ExportOptions.Render "body", "", "", "", "", "" %></div>
<% End If %>
<% If (e_library_th.Export = "") Or (EW_EXPORT_MASTER_RECORD And e_library_th.Export = "print") Then %>
<% End If %>
<%

' Load recordset
Set e_library_th_list.Recordset = e_library_th_list.LoadRecordset()
	e_library_th_list.TotalRecs = e_library_th_list.Recordset.RecordCount
	e_library_th_list.StartRec = 1
	If e_library_th_list.DisplayRecs <= 0 Then ' Display all records
		e_library_th_list.DisplayRecs = e_library_th_list.TotalRecs
	End If
	If Not (e_library_th.ExportAll And e_library_th.Export <> "") Then
		e_library_th_list.SetUpStartRec() ' Set up start record position
	End If
e_library_th_list.RenderOtherOptions()
%>
<% If Security.IsLoggedIn() Then %>
<% If e_library_th.Export = "" And e_library_th.CurrentAction = "" Then %>
<form name="fe_library_thlistsrch" id="fe_library_thlistsrch" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewSearchTable"><tr><td>
<div class="accordion" id="fe_library_thlistsrch_SearchGroup">
	<div class="accordion-group">
		<div class="accordion-heading">
<a class="accordion-toggle" data-toggle="collapse" data-parent="#fe_library_thlistsrch_SearchGroup" href="#fe_library_thlistsrch_SearchBody"><%= Language.Phrase("Search") %></a>
		</div>
		<div id="fe_library_thlistsrch_SearchBody" class="accordion-body collapse in">
			<div class="accordion-inner">
<div id="fe_library_thlistsrch_SearchPanel">
<input type="hidden" name="cmd" value="search">
<input type="hidden" name="t" value="e_library_th">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewRow">
	<div class="btn-group ewButtonGroup">
	<div class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" class="input-large" value="<%= ew_HtmlEncode(e_library_th.BasicSearch.getKeyword()) %>" placeholder="<%= ew_HtmlEncode(Language.Phrase("Search")) %>">
	<button class="btn btn-primary ewButton" name="btnsubmit" id="btnsubmit" type="submit"><%= Language.Phrase("QuickSearchBtn") %></button>
	</div>
	</div>
	<div class="btn-group ewButtonGroup">
	<a class="btn ewShowAll" href="<%= e_library_th_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>
	</div>
</div>
<div id="xsr_2" class="ewRow">
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="="<% If e_library_th.BasicSearch.getSearchType = "=" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If e_library_th.BasicSearch.getSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If e_library_th.BasicSearch.getSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
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
<% e_library_th_list.ShowPageHeader() %>
<% e_library_th_list.ShowMessage %>
<table class="ewGrid"><tr><td class="ewGridContent">
<form name="fe_library_thlist" id="fe_library_thlist" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="e_library_th">
<div id="gmp_e_library_th" class="ewGridMiddlePanel">
<% If e_library_th_list.TotalRecs > 0 Then %>
<table id="tbl_e_library_thlist" class="ewTable ewTableSeparate">
<%= e_library_th.TableCustomInnerHtml %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call e_library_th_list.RenderListOptions()

' Render list options (header, left)
e_library_th_list.ListOptions.Render "header", "left", "", "", "", ""
%>
<% If e_library_th.el_id.Visible Then ' el_id %>
	<% If e_library_th.SortUrl(e_library_th.el_id) = "" Then %>
		<td><div id="elh_e_library_th_el_id" class="e_library_th_el_id"><div class="ewTableHeaderCaption"><%= e_library_th.el_id.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= e_library_th.SortUrl(e_library_th.el_id) %>',1);"><div id="elh_e_library_th_el_id" class="e_library_th_el_id">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= e_library_th.el_id.FldCaption %></span><span class="ewTableHeaderSort"><% If e_library_th.el_id.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf e_library_th.el_id.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If e_library_th.el_date.Visible Then ' el_date %>
	<% If e_library_th.SortUrl(e_library_th.el_date) = "" Then %>
		<td><div id="elh_e_library_th_el_date" class="e_library_th_el_date"><div class="ewTableHeaderCaption"><%= e_library_th.el_date.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= e_library_th.SortUrl(e_library_th.el_date) %>',1);"><div id="elh_e_library_th_el_date" class="e_library_th_el_date">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= e_library_th.el_date.FldCaption %></span><span class="ewTableHeaderSort"><% If e_library_th.el_date.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf e_library_th.el_date.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If e_library_th.el_title.Visible Then ' el_title %>
	<% If e_library_th.SortUrl(e_library_th.el_title) = "" Then %>
		<td><div id="elh_e_library_th_el_title" class="e_library_th_el_title"><div class="ewTableHeaderCaption"><%= e_library_th.el_title.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= e_library_th.SortUrl(e_library_th.el_title) %>',1);"><div id="elh_e_library_th_el_title" class="e_library_th_el_title">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= e_library_th.el_title.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If e_library_th.el_title.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf e_library_th.el_title.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If e_library_th.el_pdf.Visible Then ' el_pdf %>
	<% If e_library_th.SortUrl(e_library_th.el_pdf) = "" Then %>
		<td><div id="elh_e_library_th_el_pdf" class="e_library_th_el_pdf"><div class="ewTableHeaderCaption"><%= e_library_th.el_pdf.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= e_library_th.SortUrl(e_library_th.el_pdf) %>',1);"><div id="elh_e_library_th_el_pdf" class="e_library_th_el_pdf">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= e_library_th.el_pdf.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If e_library_th.el_pdf.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf e_library_th.el_pdf.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If e_library_th.el_img.Visible Then ' el_img %>
	<% If e_library_th.SortUrl(e_library_th.el_img) = "" Then %>
		<td><div id="elh_e_library_th_el_img" class="e_library_th_el_img"><div class="ewTableHeaderCaption"><%= e_library_th.el_img.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= e_library_th.SortUrl(e_library_th.el_img) %>',1);"><div id="elh_e_library_th_el_img" class="e_library_th_el_img">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= e_library_th.el_img.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If e_library_th.el_img.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf e_library_th.el_img.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If e_library_th.el_create.Visible Then ' el_create %>
	<% If e_library_th.SortUrl(e_library_th.el_create) = "" Then %>
		<td><div id="elh_e_library_th_el_create" class="e_library_th_el_create"><div class="ewTableHeaderCaption"><%= e_library_th.el_create.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= e_library_th.SortUrl(e_library_th.el_create) %>',1);"><div id="elh_e_library_th_el_create" class="e_library_th_el_create">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= e_library_th.el_create.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If e_library_th.el_create.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf e_library_th.el_create.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If e_library_th.el_update.Visible Then ' el_update %>
	<% If e_library_th.SortUrl(e_library_th.el_update) = "" Then %>
		<td><div id="elh_e_library_th_el_update" class="e_library_th_el_update"><div class="ewTableHeaderCaption"><%= e_library_th.el_update.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= e_library_th.SortUrl(e_library_th.el_update) %>',1);"><div id="elh_e_library_th_el_update" class="e_library_th_el_update">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= e_library_th.el_update.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If e_library_th.el_update.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf e_library_th.el_update.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
e_library_th_list.ListOptions.Render "header", "right", "", "", "", ""
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (e_library_th.ExportAll And e_library_th.Export <> "") Then
	e_library_th_list.StopRec = e_library_th_list.TotalRecs
Else

	' Set the last record to display
	If e_library_th_list.TotalRecs > e_library_th_list.StartRec + e_library_th_list.DisplayRecs - 1 Then
		e_library_th_list.StopRec = e_library_th_list.StartRec + e_library_th_list.DisplayRecs - 1
	Else
		e_library_th_list.StopRec = e_library_th_list.TotalRecs
	End If
End If

' Move to first record
e_library_th_list.RecCnt = e_library_th_list.StartRec - 1
If Not e_library_th_list.Recordset.Eof Then
	e_library_th_list.Recordset.MoveFirst
	If e_library_th_list.StartRec > 1 Then e_library_th_list.Recordset.Move e_library_th_list.StartRec - 1
ElseIf Not e_library_th.AllowAddDeleteRow And e_library_th_list.StopRec = 0 Then
	e_library_th_list.StopRec = e_library_th.GridAddRowCount
End If

' Initialize Aggregate
e_library_th.RowType = EW_ROWTYPE_AGGREGATEINIT
Call e_library_th.ResetAttrs()
Call e_library_th_list.RenderRow()
e_library_th_list.RowCnt = 0

' Output date rows
Do While CLng(e_library_th_list.RecCnt) < CLng(e_library_th_list.StopRec)
	e_library_th_list.RecCnt = e_library_th_list.RecCnt + 1
	If CLng(e_library_th_list.RecCnt) >= CLng(e_library_th_list.StartRec) Then
		e_library_th_list.RowCnt = e_library_th_list.RowCnt + 1

	' Set up key count
	e_library_th_list.KeyCount = e_library_th_list.RowIndex
	Call e_library_th.ResetAttrs()
	e_library_th.CssClass = ""
	If e_library_th.CurrentAction = "gridadd" Then
	Else
		Call e_library_th_list.LoadRowValues(e_library_th_list.Recordset) ' Load row values
	End If
	e_library_th.RowType = EW_ROWTYPE_VIEW ' Render view

	' Set up row id / data-rowindex
	e_library_th.RowAttrs.AddAttributes Array(Array("data-rowindex", e_library_th_list.RowCnt), Array("id", "r" & e_library_th_list.RowCnt & "_e_library_th"), Array("data-rowtype", e_library_th.RowType))

	' Render row
	Call e_library_th_list.RenderRow()

	' Render list options
	Call e_library_th_list.RenderListOptions()
%>
	<tr<%= e_library_th.RowAttributes %>>
<%

' Render list options (body, left)
e_library_th_list.ListOptions.Render "body", "left", e_library_th_list.RowCnt, "", "", ""
%>
	<% If e_library_th.el_id.Visible Then ' el_id %>
		<td<%= e_library_th.el_id.CellAttributes %>>
<span<%= e_library_th.el_id.ViewAttributes %>>
<%= e_library_th.el_id.ListViewValue %>
</span>
<a id="<%= e_library_th_list.PageObjName & "_row_" & e_library_th_list.RowCnt %>"></a></td>
	<% End If %>
	<% If e_library_th.el_date.Visible Then ' el_date %>
		<td<%= e_library_th.el_date.CellAttributes %>>
<span<%= e_library_th.el_date.ViewAttributes %>>
<%= e_library_th.el_date.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If e_library_th.el_title.Visible Then ' el_title %>
		<td<%= e_library_th.el_title.CellAttributes %>>
<span<%= e_library_th.el_title.ViewAttributes %>>
<%= e_library_th.el_title.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If e_library_th.el_pdf.Visible Then ' el_pdf %>
		<td<%= e_library_th.el_pdf.CellAttributes %>>
<span<%= e_library_th.el_pdf.ViewAttributes %>>
<%= e_library_th.el_pdf.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If e_library_th.el_img.Visible Then ' el_img %>
		<td<%= e_library_th.el_img.CellAttributes %>>
<span<%= e_library_th.el_img.ViewAttributes %>>
<%= e_library_th.el_img.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If e_library_th.el_create.Visible Then ' el_create %>
		<td<%= e_library_th.el_create.CellAttributes %>>
<span<%= e_library_th.el_create.ViewAttributes %>>
<%= e_library_th.el_create.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If e_library_th.el_update.Visible Then ' el_update %>
		<td<%= e_library_th.el_update.CellAttributes %>>
<span<%= e_library_th.el_update.ViewAttributes %>>
<%= e_library_th.el_update.ListViewValue %>
</span>
</td>
	<% End If %>
<%

' Render list options (body, right)
e_library_th_list.ListOptions.Render "body", "right", e_library_th_list.RowCnt, "", "", ""
%>
	</tr>
<%
	End If
	If e_library_th.CurrentAction <> "gridadd" Then
		e_library_th_list.Recordset.MoveNext()
	End If
Loop
%>
</tbody>
</table>
<% End If %>
<% If e_library_th.CurrentAction = "" Then %>
<input type="hidden" name="a_list" id="a_list" value="">
<% End If %>
</div>
</form>
<%

' Close recordset and connection
e_library_th_list.Recordset.Close
Set e_library_th_list.Recordset = Nothing
%>
<% If e_library_th.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If e_library_th.CurrentAction <> "gridadd" And e_library_th.CurrentAction <> "gridedit" Then %>
<form name="ewPagerForm" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewPager">
<tr><td>
<% If Not IsObject(e_library_th_list.Pager) Then Set e_library_th_list.Pager = ew_NewPrevNextPager(e_library_th_list.StartRec, e_library_th_list.DisplayRecs, e_library_th_list.TotalRecs) %>
<% If e_library_th_list.Pager.RecordCount > 0 Then %>
<table class="ewStdTable"><tbody><tr><td>
	<%= Language.Phrase("Page") %>&nbsp;
<div class="input-prepend input-append">
<!--first page button-->
	<% If e_library_th_list.Pager.FirstButton.Enabled Then %>
	<a class="btn btn-small" href="<%= e_library_th_list.PageUrl %>start=<%= e_library_th_list.Pager.FirstButton.Start %>"><i class="icon-step-backward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-backward"></i></a>
	<% End If %>
<!--previous page button-->
	<% If e_library_th_list.Pager.PrevButton.Enabled Then %>
	<a class="btn btn-small" href="<%= e_library_th_list.PageUrl %>start=<%= e_library_th_list.Pager.PrevButton.Start %>"><i class="icon-prev"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-prev"></i></a>
	<% End If %>
<!--current page number-->
	<input class="input-mini" type="text" name="<%= EW_TABLE_PAGE_NO %>" value="<%= e_library_th_list.Pager.CurrentPage %>">
<!--next page button-->
	<% If e_library_th_list.Pager.NextButton.Enabled Then %>
	<a class="btn btn-small" href="<%= e_library_th_list.PageUrl %>start=<%= e_library_th_list.Pager.NextButton.Start %>"><i class="icon-play"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-play"></i></a>
	<% End If %>
<!--last page button-->
	<% If e_library_th_list.Pager.LastButton.Enabled Then %>
	<a class="btn btn-small" href="<%= e_library_th_list.PageUrl %>start=<%= e_library_th_list.Pager.LastButton.Start %>"><i class="icon-step-forward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-forward"></i></a>
	<% End If %>
</div>
	&nbsp;<%= Language.Phrase("of") %>&nbsp;<%= e_library_th_list.Pager.PageCount %>
</td>
<td>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<%= Language.Phrase("Record") %>&nbsp;<%= e_library_th_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= e_library_th_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= e_library_th_list.Pager.RecordCount %>
</td>
</tr></tbody></table>
<% Else %>
	<% If e_library_th_list.SearchWhere = "0=101" Then %>
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
	e_library_th_list.AddEditOptions.Render "body", "bottom", "", "", "", ""
	e_library_th_list.DetailOptions.Render "body", "bottom", "", "", "", ""
	e_library_th_list.ActionOptions.Render "body", "bottom", "", "", "", ""
%>
</div>
</div>
<% End If %>
</td></tr></table>
<% If e_library_th.Export = "" Then %>
<script type="text/javascript">
fe_library_thlistsrch.Init();
fe_library_thlist.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<% End If %>
<%
e_library_th_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If e_library_th.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set e_library_th_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class ce_library_th_list

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
		TableName = "e_library_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "e_library_th_list"
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
		If e_library_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & e_library_th.TableVar & "&" ' add page token
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
		If e_library_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (e_library_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (e_library_th.TableVar = Request.QueryString("t"))
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
		FormName = "fe_library_thlist"
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
		If IsEmpty(e_library_th) Then Set e_library_th = New ce_library_th
		Set Table = e_library_th

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		ExportPdfUrl = PageUrl & "export=pdf"
		AddUrl = "pom_e_library_thadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "pom_e_library_thdelete.asp"
		MultiUpdateUrl = "pom_e_library_thupdate.asp"

		' Initialize other table object
		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "e_library_th"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' List options
		Set ListOptions = New cListOptions
		ListOptions.TableVar = e_library_th.TableVar

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = e_library_th.TableVar
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
				e_library_th.GridAddRowCount = gridaddcnt
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
		If UBound(e_library_th.CustomActions.CustomArray) >= 0 Then
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
		Set e_library_th = Nothing
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
			If e_library_th.Export = "" Then
				SetupBreadcrumb()
			End If

			' Hide list options
			If e_library_th.Export <> "" Then
				Call ListOptions.HideAllOptions(Array("sequence"))
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			ElseIf e_library_th.CurrentAction = "gridadd" Or e_library_th.CurrentAction = "gridedit" Then
				Call ListOptions.HideAllOptions(Array())
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			End If

			' Hide export options
			If e_library_th.Export <> "" Or e_library_th.CurrentAction <> "" Then
				ExportOptions.HideAllOptions(Array())
			End If

			' Hide other options
			If e_library_th.Export <> "" Then
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
			Call e_library_th.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If e_library_th.RecordsPerPage <> "" Then
			DisplayRecs = e_library_th.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Load search default if no existing search criteria
		If Not CheckSearchParms() Then

			' Load basic search from default
			e_library_th.BasicSearch.Keyword = e_library_th.BasicSearch.KeywordDefault
			e_library_th.BasicSearch.SearchType = e_library_th.BasicSearch.SearchTypeDefault
			e_library_th.BasicSearch.setSearchType(e_library_th.BasicSearch.SearchTypeDefault)
			If e_library_th.BasicSearch.Keyword <> "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call e_library_th.Recordset_Searching(SearchWhere)

		' Save search criteria
		If Command = "search" And Not RestoreSearch Then
			e_library_th.SearchWhere = SearchWhere ' Save to Session
			StartRec = 1 ' Reset start record counter
			e_library_th.StartRecordNumber = StartRec
		Else
			SearchWhere = e_library_th.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		e_library_th.SessionWhere = sFilter
		e_library_th.CurrentFilter = ""
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
				sFilter = e_library_th.KeyFilter
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
			e_library_th.el_id.FormValue = arrKeyFlds(0)
			If Not IsNumeric(e_library_th.el_id.FormValue) Then
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
			Call BuildBasicSearchSQL(sWhere, e_library_th.el_title, Keyword)
			Call BuildBasicSearchSQL(sWhere, e_library_th.el_pdf, Keyword)
			Call BuildBasicSearchSQL(sWhere, e_library_th.el_img, Keyword)
			Call BuildBasicSearchSQL(sWhere, e_library_th.el_detail, Keyword)
			Call BuildBasicSearchSQL(sWhere, e_library_th.el_create, Keyword)
			Call BuildBasicSearchSQL(sWhere, e_library_th.el_update, Keyword)
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
		sSearchKeyword = e_library_th.BasicSearch.Keyword
		sSearchType = e_library_th.BasicSearch.SearchType
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
			e_library_th.BasicSearch.setKeyword(sSearchKeyword)
			e_library_th.BasicSearch.setSearchType(sSearchType)
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' Check if search parm exists
	Function CheckSearchParms()

		' Check basic search
		If e_library_th.BasicSearch.IssetSession() Then
			CheckSearchParms = True
			Exit Function
		End If
		CheckSearchParms = False
	End Function

	' Clear all search parameters
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		e_library_th.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' Load advanced search default values
	Function LoadAdvancedSearchDefault()
		LoadAdvancedSearchDefault = False
	End Function

	' Clear all basic search parameters
	Sub ResetBasicSearchParms()
		e_library_th.BasicSearch.UnsetSession()
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()

		' Restore search flag
		RestoreSearch = True

		' Restore basic search values
		Call e_library_th.BasicSearch.Load()
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
			e_library_th.CurrentOrder = Request.QueryString("order")
			e_library_th.CurrentOrderType = Request.QueryString("ordertype")

			' Field el_id
			Call e_library_th.UpdateSort(e_library_th.el_id)

			' Field el_date
			Call e_library_th.UpdateSort(e_library_th.el_date)

			' Field el_title
			Call e_library_th.UpdateSort(e_library_th.el_title)

			' Field el_pdf
			Call e_library_th.UpdateSort(e_library_th.el_pdf)

			' Field el_img
			Call e_library_th.UpdateSort(e_library_th.el_img)

			' Field el_create
			Call e_library_th.UpdateSort(e_library_th.el_create)

			' Field el_update
			Call e_library_th.UpdateSort(e_library_th.el_update)
			e_library_th.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = e_library_th.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If e_library_th.SqlOrderBy <> "" Then
				sOrderBy = e_library_th.SqlOrderBy
				e_library_th.SessionOrderBy = sOrderBy
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
				e_library_th.SessionOrderBy = sOrderBy
				e_library_th.el_id.Sort = ""
				e_library_th.el_date.Sort = ""
				e_library_th.el_title.Sort = ""
				e_library_th.el_pdf.Sort = ""
				e_library_th.el_img.Sort = ""
				e_library_th.el_create.Sort = ""
				e_library_th.el_update.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			e_library_th.StartRecordNumber = StartRec
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
		ListOptions.GetItem("checkbox").Body = "<label class=""checkbox""><input type=""checkbox"" name=""key_m"" value=""" & ew_HtmlEncode(e_library_th.el_id.CurrentValue) & """ onclick='ew_ClickMultiCheckbox(event, this);'></label>"
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
			For i = 0 to UBound(e_library_th.CustomActions.CustomArray)
				Action = e_library_th.CustomActions.CustomArray(i)(0)
				Name = e_library_th.CustomActions.CustomArray(i)(1)

				' Add custom action
				Call opt.Add("custom_" & Action)
				Set item = opt.GetItem("custom_" & Action)
				item.Body = "<a class=""ewAction ewCustomAction"" href="""" onclick=""ew_SubmitSelected(document.fe_library_thlist, '" & ew_CurrentUrl & "', null, '" & Action & "');return false;"">" & Name & "</a>"
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
		sFilter = e_library_th.GetKeyFilter
		UserAction = Request.Form("useraction") & ""
		Processed = False
		If sFilter <> "" And UserAction <> "" Then
			e_library_th.CurrentFilter = sFilter
			sSql = e_library_th.SQL
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
				ElseIf e_library_th.CancelMessage <> "" Then
					FailureMessage = e_library_th.CancelMessage
					e_library_th.CancelMessage = ""
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
				e_library_th.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					e_library_th.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = e_library_th.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			e_library_th.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			e_library_th.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			e_library_th.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		e_library_th.BasicSearch.Keyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)&""
		If e_library_th.BasicSearch.Keyword <> "" Then Command = "search"
		e_library_th.BasicSearch.SearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)&""
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = e_library_th.CurrentFilter
		Call e_library_th.Recordset_Selecting(sFilter)
		e_library_th.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = e_library_th.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call e_library_th.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = e_library_th.KeyFilter

		' Call Row Selecting event
		Call e_library_th.Row_Selecting(sFilter)

		' Load sql based on filter
		e_library_th.CurrentFilter = sFilter
		sSql = e_library_th.SQL
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
		Call e_library_th.Row_Selected(RsRow)
		e_library_th.el_id.DbValue = RsRow("el_id")
		e_library_th.el_date.DbValue = RsRow("el_date")
		e_library_th.el_title.DbValue = RsRow("el_title")
		e_library_th.el_pdf.DbValue = RsRow("el_pdf")
		e_library_th.el_img.DbValue = RsRow("el_img")
		e_library_th.el_detail.DbValue = RsRow("el_detail")
		e_library_th.el_create.DbValue = RsRow("el_create")
		e_library_th.el_update.DbValue = RsRow("el_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		e_library_th.el_id.m_DbValue = Rs("el_id")
		e_library_th.el_date.m_DbValue = Rs("el_date")
		e_library_th.el_title.m_DbValue = Rs("el_title")
		e_library_th.el_pdf.m_DbValue = Rs("el_pdf")
		e_library_th.el_img.m_DbValue = Rs("el_img")
		e_library_th.el_detail.m_DbValue = Rs("el_detail")
		e_library_th.el_create.m_DbValue = Rs("el_create")
		e_library_th.el_update.m_DbValue = Rs("el_update")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If e_library_th.GetKey("el_id")&"" <> "" Then
			e_library_th.el_id.CurrentValue = e_library_th.GetKey("el_id") ' el_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			e_library_th.CurrentFilter = e_library_th.KeyFilter
			Dim sSql
			sSql = e_library_th.SQL
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
		ViewUrl = e_library_th.ViewUrl("")
		EditUrl = e_library_th.EditUrl("")
		InlineEditUrl = e_library_th.InlineEditUrl
		CopyUrl = e_library_th.CopyUrl("")
		InlineCopyUrl = e_library_th.InlineCopyUrl
		DeleteUrl = e_library_th.DeleteUrl

		' Call Row Rendering event
		Call e_library_th.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' el_id
		' el_date
		' el_title
		' el_pdf
		' el_img
		' el_detail
		' el_create
		' el_update
		' -----------
		'  View  Row
		' -----------

		If e_library_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' el_id
			e_library_th.el_id.ViewValue = e_library_th.el_id.CurrentValue
			e_library_th.el_id.ViewCustomAttributes = ""

			' el_date
			e_library_th.el_date.ViewValue = e_library_th.el_date.CurrentValue
			e_library_th.el_date.ViewCustomAttributes = ""

			' el_title
			e_library_th.el_title.ViewValue = e_library_th.el_title.CurrentValue
			e_library_th.el_title.ViewCustomAttributes = ""

			' el_pdf
			e_library_th.el_pdf.ViewValue = e_library_th.el_pdf.CurrentValue
			e_library_th.el_pdf.ViewCustomAttributes = ""

			' el_img
			e_library_th.el_img.ViewValue = e_library_th.el_img.CurrentValue
			e_library_th.el_img.ViewCustomAttributes = ""

			' el_create
			e_library_th.el_create.ViewValue = e_library_th.el_create.CurrentValue
			e_library_th.el_create.ViewCustomAttributes = ""

			' el_update
			e_library_th.el_update.ViewValue = e_library_th.el_update.CurrentValue
			e_library_th.el_update.ViewCustomAttributes = ""

			' View refer script
			' el_id

			e_library_th.el_id.LinkCustomAttributes = ""
			e_library_th.el_id.HrefValue = ""
			e_library_th.el_id.TooltipValue = ""

			' el_date
			e_library_th.el_date.LinkCustomAttributes = ""
			e_library_th.el_date.HrefValue = ""
			e_library_th.el_date.TooltipValue = ""

			' el_title
			e_library_th.el_title.LinkCustomAttributes = ""
			e_library_th.el_title.HrefValue = ""
			e_library_th.el_title.TooltipValue = ""

			' el_pdf
			e_library_th.el_pdf.LinkCustomAttributes = ""
			e_library_th.el_pdf.HrefValue = ""
			e_library_th.el_pdf.TooltipValue = ""

			' el_img
			e_library_th.el_img.LinkCustomAttributes = ""
			e_library_th.el_img.HrefValue = ""
			e_library_th.el_img.TooltipValue = ""

			' el_create
			e_library_th.el_create.LinkCustomAttributes = ""
			e_library_th.el_create.HrefValue = ""
			e_library_th.el_create.TooltipValue = ""

			' el_update
			e_library_th.el_update.LinkCustomAttributes = ""
			e_library_th.el_update.HrefValue = ""
			e_library_th.el_update.TooltipValue = ""
		End If

		' Call Row Rendered event
		If e_library_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call e_library_th.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		url = ew_CurrentUrl
		url = ew_RegExReplace("\?cmd=reset(all){0,1}$", url, "") ' Remove cmd=reset / cmd=resetall
		Call Breadcrumb.Add("list", e_library_th.TableVar, url, e_library_th.TableVar, True)
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
