<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_banner_logo_02info.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim banner_logo_02_list
Set banner_logo_02_list = New cbanner_logo_02_list
Set Page = banner_logo_02_list

' Page init processing
banner_logo_02_list.Page_Init()

' Page main processing
banner_logo_02_list.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
banner_logo_02_list.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If banner_logo_02.Export = "" Then %>
<script type="text/javascript">
// Page object
var banner_logo_02_list = new ew_Page("banner_logo_02_list");
banner_logo_02_list.PageID = "list"; // Page ID
var EW_PAGE_ID = banner_logo_02_list.PageID; // For backward compatibility
// Form object
var fbanner_logo_02list = new ew_Form("fbanner_logo_02list");
fbanner_logo_02list.FormKeyCountName = '<%= banner_logo_02_list.FormKeyCountName %>';
// Form_CustomValidate event
fbanner_logo_02list.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fbanner_logo_02list.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fbanner_logo_02list.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
var fbanner_logo_02listsrch = new ew_Form("fbanner_logo_02listsrch");
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If banner_logo_02.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If banner_logo_02_list.ExportOptions.Visible Then %>
<div class="ewListExportOptions"><% banner_logo_02_list.ExportOptions.Render "body", "", "", "", "", "" %></div>
<% End If %>
<% If (banner_logo_02.Export = "") Or (EW_EXPORT_MASTER_RECORD And banner_logo_02.Export = "print") Then %>
<% End If %>
<%

' Load recordset
Set banner_logo_02_list.Recordset = banner_logo_02_list.LoadRecordset()
	banner_logo_02_list.TotalRecs = banner_logo_02_list.Recordset.RecordCount
	banner_logo_02_list.StartRec = 1
	If banner_logo_02_list.DisplayRecs <= 0 Then ' Display all records
		banner_logo_02_list.DisplayRecs = banner_logo_02_list.TotalRecs
	End If
	If Not (banner_logo_02.ExportAll And banner_logo_02.Export <> "") Then
		banner_logo_02_list.SetUpStartRec() ' Set up start record position
	End If
banner_logo_02_list.RenderOtherOptions()
%>
<% If Security.IsLoggedIn() Then %>
<% If banner_logo_02.Export = "" And banner_logo_02.CurrentAction = "" Then %>
<form name="fbanner_logo_02listsrch" id="fbanner_logo_02listsrch" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewSearchTable"><tr><td>
<div class="accordion" id="fbanner_logo_02listsrch_SearchGroup">
	<div class="accordion-group">
		<div class="accordion-heading">
<a class="accordion-toggle" data-toggle="collapse" data-parent="#fbanner_logo_02listsrch_SearchGroup" href="#fbanner_logo_02listsrch_SearchBody"><%= Language.Phrase("Search") %></a>
		</div>
		<div id="fbanner_logo_02listsrch_SearchBody" class="accordion-body collapse in">
			<div class="accordion-inner">
<div id="fbanner_logo_02listsrch_SearchPanel">
<input type="hidden" name="cmd" value="search">
<input type="hidden" name="t" value="banner_logo_02">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewRow">
	<div class="btn-group ewButtonGroup">
	<div class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" class="input-large" value="<%= ew_HtmlEncode(banner_logo_02.BasicSearch.getKeyword()) %>" placeholder="<%= ew_HtmlEncode(Language.Phrase("Search")) %>">
	<button class="btn btn-primary ewButton" name="btnsubmit" id="btnsubmit" type="submit"><%= Language.Phrase("QuickSearchBtn") %></button>
	</div>
	</div>
	<div class="btn-group ewButtonGroup">
	<a class="btn ewShowAll" href="<%= banner_logo_02_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>
	</div>
</div>
<div id="xsr_2" class="ewRow">
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="="<% If banner_logo_02.BasicSearch.getSearchType = "=" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If banner_logo_02.BasicSearch.getSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If banner_logo_02.BasicSearch.getSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
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
<% banner_logo_02_list.ShowPageHeader() %>
<% banner_logo_02_list.ShowMessage %>
<table class="ewGrid"><tr><td class="ewGridContent">
<form name="fbanner_logo_02list" id="fbanner_logo_02list" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="banner_logo_02">
<div id="gmp_banner_logo_02" class="ewGridMiddlePanel">
<% If banner_logo_02_list.TotalRecs > 0 Then %>
<table id="tbl_banner_logo_02list" class="ewTable ewTableSeparate">
<%= banner_logo_02.TableCustomInnerHtml %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call banner_logo_02_list.RenderListOptions()

' Render list options (header, left)
banner_logo_02_list.ListOptions.Render "header", "left", "", "", "", ""
%>
<% If banner_logo_02.banner_id.Visible Then ' banner_id %>
	<% If banner_logo_02.SortUrl(banner_logo_02.banner_id) = "" Then %>
		<td><div id="elh_banner_logo_02_banner_id" class="banner_logo_02_banner_id"><div class="ewTableHeaderCaption"><%= banner_logo_02.banner_id.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= banner_logo_02.SortUrl(banner_logo_02.banner_id) %>',1);"><div id="elh_banner_logo_02_banner_id" class="banner_logo_02_banner_id">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= banner_logo_02.banner_id.FldCaption %></span><span class="ewTableHeaderSort"><% If banner_logo_02.banner_id.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf banner_logo_02.banner_id.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If banner_logo_02.banner_img.Visible Then ' banner_img %>
	<% If banner_logo_02.SortUrl(banner_logo_02.banner_img) = "" Then %>
		<td><div id="elh_banner_logo_02_banner_img" class="banner_logo_02_banner_img"><div class="ewTableHeaderCaption"><%= banner_logo_02.banner_img.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= banner_logo_02.SortUrl(banner_logo_02.banner_img) %>',1);"><div id="elh_banner_logo_02_banner_img" class="banner_logo_02_banner_img">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= banner_logo_02.banner_img.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If banner_logo_02.banner_img.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf banner_logo_02.banner_img.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If banner_logo_02.banner_link.Visible Then ' banner_link %>
	<% If banner_logo_02.SortUrl(banner_logo_02.banner_link) = "" Then %>
		<td><div id="elh_banner_logo_02_banner_link" class="banner_logo_02_banner_link"><div class="ewTableHeaderCaption"><%= banner_logo_02.banner_link.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= banner_logo_02.SortUrl(banner_logo_02.banner_link) %>',1);"><div id="elh_banner_logo_02_banner_link" class="banner_logo_02_banner_link">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= banner_logo_02.banner_link.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If banner_logo_02.banner_link.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf banner_logo_02.banner_link.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If banner_logo_02.banner_sort.Visible Then ' banner_sort %>
	<% If banner_logo_02.SortUrl(banner_logo_02.banner_sort) = "" Then %>
		<td><div id="elh_banner_logo_02_banner_sort" class="banner_logo_02_banner_sort"><div class="ewTableHeaderCaption"><%= banner_logo_02.banner_sort.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= banner_logo_02.SortUrl(banner_logo_02.banner_sort) %>',1);"><div id="elh_banner_logo_02_banner_sort" class="banner_logo_02_banner_sort">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= banner_logo_02.banner_sort.FldCaption %></span><span class="ewTableHeaderSort"><% If banner_logo_02.banner_sort.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf banner_logo_02.banner_sort.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If banner_logo_02.start_date.Visible Then ' start_date %>
	<% If banner_logo_02.SortUrl(banner_logo_02.start_date) = "" Then %>
		<td><div id="elh_banner_logo_02_start_date" class="banner_logo_02_start_date"><div class="ewTableHeaderCaption"><%= banner_logo_02.start_date.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= banner_logo_02.SortUrl(banner_logo_02.start_date) %>',1);"><div id="elh_banner_logo_02_start_date" class="banner_logo_02_start_date">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= banner_logo_02.start_date.FldCaption %></span><span class="ewTableHeaderSort"><% If banner_logo_02.start_date.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf banner_logo_02.start_date.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If banner_logo_02.end_date.Visible Then ' end_date %>
	<% If banner_logo_02.SortUrl(banner_logo_02.end_date) = "" Then %>
		<td><div id="elh_banner_logo_02_end_date" class="banner_logo_02_end_date"><div class="ewTableHeaderCaption"><%= banner_logo_02.end_date.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= banner_logo_02.SortUrl(banner_logo_02.end_date) %>',1);"><div id="elh_banner_logo_02_end_date" class="banner_logo_02_end_date">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= banner_logo_02.end_date.FldCaption %></span><span class="ewTableHeaderSort"><% If banner_logo_02.end_date.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf banner_logo_02.end_date.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
banner_logo_02_list.ListOptions.Render "header", "right", "", "", "", ""
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (banner_logo_02.ExportAll And banner_logo_02.Export <> "") Then
	banner_logo_02_list.StopRec = banner_logo_02_list.TotalRecs
Else

	' Set the last record to display
	If banner_logo_02_list.TotalRecs > banner_logo_02_list.StartRec + banner_logo_02_list.DisplayRecs - 1 Then
		banner_logo_02_list.StopRec = banner_logo_02_list.StartRec + banner_logo_02_list.DisplayRecs - 1
	Else
		banner_logo_02_list.StopRec = banner_logo_02_list.TotalRecs
	End If
End If

' Move to first record
banner_logo_02_list.RecCnt = banner_logo_02_list.StartRec - 1
If Not banner_logo_02_list.Recordset.Eof Then
	banner_logo_02_list.Recordset.MoveFirst
	If banner_logo_02_list.StartRec > 1 Then banner_logo_02_list.Recordset.Move banner_logo_02_list.StartRec - 1
ElseIf Not banner_logo_02.AllowAddDeleteRow And banner_logo_02_list.StopRec = 0 Then
	banner_logo_02_list.StopRec = banner_logo_02.GridAddRowCount
End If

' Initialize Aggregate
banner_logo_02.RowType = EW_ROWTYPE_AGGREGATEINIT
Call banner_logo_02.ResetAttrs()
Call banner_logo_02_list.RenderRow()
banner_logo_02_list.RowCnt = 0

' Output date rows
Do While CLng(banner_logo_02_list.RecCnt) < CLng(banner_logo_02_list.StopRec)
	banner_logo_02_list.RecCnt = banner_logo_02_list.RecCnt + 1
	If CLng(banner_logo_02_list.RecCnt) >= CLng(banner_logo_02_list.StartRec) Then
		banner_logo_02_list.RowCnt = banner_logo_02_list.RowCnt + 1

	' Set up key count
	banner_logo_02_list.KeyCount = banner_logo_02_list.RowIndex
	Call banner_logo_02.ResetAttrs()
	banner_logo_02.CssClass = ""
	If banner_logo_02.CurrentAction = "gridadd" Then
	Else
		Call banner_logo_02_list.LoadRowValues(banner_logo_02_list.Recordset) ' Load row values
	End If
	banner_logo_02.RowType = EW_ROWTYPE_VIEW ' Render view

	' Set up row id / data-rowindex
	banner_logo_02.RowAttrs.AddAttributes Array(Array("data-rowindex", banner_logo_02_list.RowCnt), Array("id", "r" & banner_logo_02_list.RowCnt & "_banner_logo_02"), Array("data-rowtype", banner_logo_02.RowType))

	' Render row
	Call banner_logo_02_list.RenderRow()

	' Render list options
	Call banner_logo_02_list.RenderListOptions()
%>
	<tr<%= banner_logo_02.RowAttributes %>>
<%

' Render list options (body, left)
banner_logo_02_list.ListOptions.Render "body", "left", banner_logo_02_list.RowCnt, "", "", ""
%>
	<% If banner_logo_02.banner_id.Visible Then ' banner_id %>
		<td<%= banner_logo_02.banner_id.CellAttributes %>>
<span<%= banner_logo_02.banner_id.ViewAttributes %>>
<%= banner_logo_02.banner_id.ListViewValue %>
</span>
<a id="<%= banner_logo_02_list.PageObjName & "_row_" & banner_logo_02_list.RowCnt %>"></a></td>
	<% End If %>
	<% If banner_logo_02.banner_img.Visible Then ' banner_img %>
		<td<%= banner_logo_02.banner_img.CellAttributes %>>
<span<%= banner_logo_02.banner_img.ViewAttributes %>>
<%= banner_logo_02.banner_img.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If banner_logo_02.banner_link.Visible Then ' banner_link %>
		<td<%= banner_logo_02.banner_link.CellAttributes %>>
<span<%= banner_logo_02.banner_link.ViewAttributes %>>
<%= banner_logo_02.banner_link.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If banner_logo_02.banner_sort.Visible Then ' banner_sort %>
		<td<%= banner_logo_02.banner_sort.CellAttributes %>>
<span<%= banner_logo_02.banner_sort.ViewAttributes %>>
<%= banner_logo_02.banner_sort.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If banner_logo_02.start_date.Visible Then ' start_date %>
		<td<%= banner_logo_02.start_date.CellAttributes %>>
<span<%= banner_logo_02.start_date.ViewAttributes %>>
<%= banner_logo_02.start_date.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If banner_logo_02.end_date.Visible Then ' end_date %>
		<td<%= banner_logo_02.end_date.CellAttributes %>>
<span<%= banner_logo_02.end_date.ViewAttributes %>>
<%= banner_logo_02.end_date.ListViewValue %>
</span>
</td>
	<% End If %>
<%

' Render list options (body, right)
banner_logo_02_list.ListOptions.Render "body", "right", banner_logo_02_list.RowCnt, "", "", ""
%>
	</tr>
<%
	End If
	If banner_logo_02.CurrentAction <> "gridadd" Then
		banner_logo_02_list.Recordset.MoveNext()
	End If
Loop
%>
</tbody>
</table>
<% End If %>
<% If banner_logo_02.CurrentAction = "" Then %>
<input type="hidden" name="a_list" id="a_list" value="">
<% End If %>
</div>
</form>
<%

' Close recordset and connection
banner_logo_02_list.Recordset.Close
Set banner_logo_02_list.Recordset = Nothing
%>
<% If banner_logo_02.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If banner_logo_02.CurrentAction <> "gridadd" And banner_logo_02.CurrentAction <> "gridedit" Then %>
<form name="ewPagerForm" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewPager">
<tr><td>
<% If Not IsObject(banner_logo_02_list.Pager) Then Set banner_logo_02_list.Pager = ew_NewPrevNextPager(banner_logo_02_list.StartRec, banner_logo_02_list.DisplayRecs, banner_logo_02_list.TotalRecs) %>
<% If banner_logo_02_list.Pager.RecordCount > 0 Then %>
<table class="ewStdTable"><tbody><tr><td>
	<%= Language.Phrase("Page") %>&nbsp;
<div class="input-prepend input-append">
<!--first page button-->
	<% If banner_logo_02_list.Pager.FirstButton.Enabled Then %>
	<a class="btn btn-small" href="<%= banner_logo_02_list.PageUrl %>start=<%= banner_logo_02_list.Pager.FirstButton.Start %>"><i class="icon-step-backward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-backward"></i></a>
	<% End If %>
<!--previous page button-->
	<% If banner_logo_02_list.Pager.PrevButton.Enabled Then %>
	<a class="btn btn-small" href="<%= banner_logo_02_list.PageUrl %>start=<%= banner_logo_02_list.Pager.PrevButton.Start %>"><i class="icon-prev"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-prev"></i></a>
	<% End If %>
<!--current page number-->
	<input class="input-mini" type="text" name="<%= EW_TABLE_PAGE_NO %>" value="<%= banner_logo_02_list.Pager.CurrentPage %>">
<!--next page button-->
	<% If banner_logo_02_list.Pager.NextButton.Enabled Then %>
	<a class="btn btn-small" href="<%= banner_logo_02_list.PageUrl %>start=<%= banner_logo_02_list.Pager.NextButton.Start %>"><i class="icon-play"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-play"></i></a>
	<% End If %>
<!--last page button-->
	<% If banner_logo_02_list.Pager.LastButton.Enabled Then %>
	<a class="btn btn-small" href="<%= banner_logo_02_list.PageUrl %>start=<%= banner_logo_02_list.Pager.LastButton.Start %>"><i class="icon-step-forward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-forward"></i></a>
	<% End If %>
</div>
	&nbsp;<%= Language.Phrase("of") %>&nbsp;<%= banner_logo_02_list.Pager.PageCount %>
</td>
<td>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<%= Language.Phrase("Record") %>&nbsp;<%= banner_logo_02_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= banner_logo_02_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= banner_logo_02_list.Pager.RecordCount %>
</td>
</tr></tbody></table>
<% Else %>
	<% If banner_logo_02_list.SearchWhere = "0=101" Then %>
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
	banner_logo_02_list.AddEditOptions.Render "body", "bottom", "", "", "", ""
	banner_logo_02_list.DetailOptions.Render "body", "bottom", "", "", "", ""
	banner_logo_02_list.ActionOptions.Render "body", "bottom", "", "", "", ""
%>
</div>
</div>
<% End If %>
</td></tr></table>
<% If banner_logo_02.Export = "" Then %>
<script type="text/javascript">
fbanner_logo_02listsrch.Init();
fbanner_logo_02list.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<% End If %>
<%
banner_logo_02_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If banner_logo_02.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set banner_logo_02_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cbanner_logo_02_list

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
		TableName = "banner_logo_02"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "banner_logo_02_list"
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
		If banner_logo_02.UseTokenInUrl Then PageUrl = PageUrl & "t=" & banner_logo_02.TableVar & "&" ' add page token
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
		If banner_logo_02.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (banner_logo_02.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (banner_logo_02.TableVar = Request.QueryString("t"))
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
		FormName = "fbanner_logo_02list"
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
		If IsEmpty(banner_logo_02) Then Set banner_logo_02 = New cbanner_logo_02
		Set Table = banner_logo_02

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		ExportPdfUrl = PageUrl & "export=pdf"
		AddUrl = "pom_banner_logo_02add.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "pom_banner_logo_02delete.asp"
		MultiUpdateUrl = "pom_banner_logo_02update.asp"

		' Initialize other table object
		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "banner_logo_02"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' List options
		Set ListOptions = New cListOptions
		ListOptions.TableVar = banner_logo_02.TableVar

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = banner_logo_02.TableVar
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
				banner_logo_02.GridAddRowCount = gridaddcnt
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
		If UBound(banner_logo_02.CustomActions.CustomArray) >= 0 Then
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
		Set banner_logo_02 = Nothing
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
			If banner_logo_02.Export = "" Then
				SetupBreadcrumb()
			End If

			' Hide list options
			If banner_logo_02.Export <> "" Then
				Call ListOptions.HideAllOptions(Array("sequence"))
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			ElseIf banner_logo_02.CurrentAction = "gridadd" Or banner_logo_02.CurrentAction = "gridedit" Then
				Call ListOptions.HideAllOptions(Array())
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			End If

			' Hide export options
			If banner_logo_02.Export <> "" Or banner_logo_02.CurrentAction <> "" Then
				ExportOptions.HideAllOptions(Array())
			End If

			' Hide other options
			If banner_logo_02.Export <> "" Then
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
			Call banner_logo_02.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If banner_logo_02.RecordsPerPage <> "" Then
			DisplayRecs = banner_logo_02.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Load search default if no existing search criteria
		If Not CheckSearchParms() Then

			' Load basic search from default
			banner_logo_02.BasicSearch.Keyword = banner_logo_02.BasicSearch.KeywordDefault
			banner_logo_02.BasicSearch.SearchType = banner_logo_02.BasicSearch.SearchTypeDefault
			banner_logo_02.BasicSearch.setSearchType(banner_logo_02.BasicSearch.SearchTypeDefault)
			If banner_logo_02.BasicSearch.Keyword <> "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call banner_logo_02.Recordset_Searching(SearchWhere)

		' Save search criteria
		If Command = "search" And Not RestoreSearch Then
			banner_logo_02.SearchWhere = SearchWhere ' Save to Session
			StartRec = 1 ' Reset start record counter
			banner_logo_02.StartRecordNumber = StartRec
		Else
			SearchWhere = banner_logo_02.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		banner_logo_02.SessionWhere = sFilter
		banner_logo_02.CurrentFilter = ""
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
				sFilter = banner_logo_02.KeyFilter
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
			banner_logo_02.banner_id.FormValue = arrKeyFlds(0)
			If Not IsNumeric(banner_logo_02.banner_id.FormValue) Then
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
			Call BuildBasicSearchSQL(sWhere, banner_logo_02.banner_img, Keyword)
			Call BuildBasicSearchSQL(sWhere, banner_logo_02.banner_link, Keyword)
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
		sSearchKeyword = banner_logo_02.BasicSearch.Keyword
		sSearchType = banner_logo_02.BasicSearch.SearchType
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
			banner_logo_02.BasicSearch.setKeyword(sSearchKeyword)
			banner_logo_02.BasicSearch.setSearchType(sSearchType)
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' Check if search parm exists
	Function CheckSearchParms()

		' Check basic search
		If banner_logo_02.BasicSearch.IssetSession() Then
			CheckSearchParms = True
			Exit Function
		End If
		CheckSearchParms = False
	End Function

	' Clear all search parameters
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		banner_logo_02.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' Load advanced search default values
	Function LoadAdvancedSearchDefault()
		LoadAdvancedSearchDefault = False
	End Function

	' Clear all basic search parameters
	Sub ResetBasicSearchParms()
		banner_logo_02.BasicSearch.UnsetSession()
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()

		' Restore search flag
		RestoreSearch = True

		' Restore basic search values
		Call banner_logo_02.BasicSearch.Load()
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
			banner_logo_02.CurrentOrder = Request.QueryString("order")
			banner_logo_02.CurrentOrderType = Request.QueryString("ordertype")

			' Field banner_id
			Call banner_logo_02.UpdateSort(banner_logo_02.banner_id)

			' Field banner_img
			Call banner_logo_02.UpdateSort(banner_logo_02.banner_img)

			' Field banner_link
			Call banner_logo_02.UpdateSort(banner_logo_02.banner_link)

			' Field banner_sort
			Call banner_logo_02.UpdateSort(banner_logo_02.banner_sort)

			' Field start_date
			Call banner_logo_02.UpdateSort(banner_logo_02.start_date)

			' Field end_date
			Call banner_logo_02.UpdateSort(banner_logo_02.end_date)
			banner_logo_02.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = banner_logo_02.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If banner_logo_02.SqlOrderBy <> "" Then
				sOrderBy = banner_logo_02.SqlOrderBy
				banner_logo_02.SessionOrderBy = sOrderBy
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
				banner_logo_02.SessionOrderBy = sOrderBy
				banner_logo_02.banner_id.Sort = ""
				banner_logo_02.banner_img.Sort = ""
				banner_logo_02.banner_link.Sort = ""
				banner_logo_02.banner_sort.Sort = ""
				banner_logo_02.start_date.Sort = ""
				banner_logo_02.end_date.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			banner_logo_02.StartRecordNumber = StartRec
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
		ListOptions.GetItem("checkbox").Body = "<label class=""checkbox""><input type=""checkbox"" name=""key_m"" value=""" & ew_HtmlEncode(banner_logo_02.banner_id.CurrentValue) & """ onclick='ew_ClickMultiCheckbox(event, this);'></label>"
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
			For i = 0 to UBound(banner_logo_02.CustomActions.CustomArray)
				Action = banner_logo_02.CustomActions.CustomArray(i)(0)
				Name = banner_logo_02.CustomActions.CustomArray(i)(1)

				' Add custom action
				Call opt.Add("custom_" & Action)
				Set item = opt.GetItem("custom_" & Action)
				item.Body = "<a class=""ewAction ewCustomAction"" href="""" onclick=""ew_SubmitSelected(document.fbanner_logo_02list, '" & ew_CurrentUrl & "', null, '" & Action & "');return false;"">" & Name & "</a>"
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
		sFilter = banner_logo_02.GetKeyFilter
		UserAction = Request.Form("useraction") & ""
		Processed = False
		If sFilter <> "" And UserAction <> "" Then
			banner_logo_02.CurrentFilter = sFilter
			sSql = banner_logo_02.SQL
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
				ElseIf banner_logo_02.CancelMessage <> "" Then
					FailureMessage = banner_logo_02.CancelMessage
					banner_logo_02.CancelMessage = ""
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
				banner_logo_02.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					banner_logo_02.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = banner_logo_02.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			banner_logo_02.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			banner_logo_02.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			banner_logo_02.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		banner_logo_02.BasicSearch.Keyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)&""
		If banner_logo_02.BasicSearch.Keyword <> "" Then Command = "search"
		banner_logo_02.BasicSearch.SearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)&""
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = banner_logo_02.CurrentFilter
		Call banner_logo_02.Recordset_Selecting(sFilter)
		banner_logo_02.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = banner_logo_02.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call banner_logo_02.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = banner_logo_02.KeyFilter

		' Call Row Selecting event
		Call banner_logo_02.Row_Selecting(sFilter)

		' Load sql based on filter
		banner_logo_02.CurrentFilter = sFilter
		sSql = banner_logo_02.SQL
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
		Call banner_logo_02.Row_Selected(RsRow)
		banner_logo_02.banner_id.DbValue = RsRow("banner_id")
		banner_logo_02.banner_img.DbValue = RsRow("banner_img")
		banner_logo_02.banner_link.DbValue = RsRow("banner_link")
		banner_logo_02.banner_sort.DbValue = RsRow("banner_sort")
		banner_logo_02.start_date.DbValue = RsRow("start_date")
		banner_logo_02.end_date.DbValue = RsRow("end_date")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		banner_logo_02.banner_id.m_DbValue = Rs("banner_id")
		banner_logo_02.banner_img.m_DbValue = Rs("banner_img")
		banner_logo_02.banner_link.m_DbValue = Rs("banner_link")
		banner_logo_02.banner_sort.m_DbValue = Rs("banner_sort")
		banner_logo_02.start_date.m_DbValue = Rs("start_date")
		banner_logo_02.end_date.m_DbValue = Rs("end_date")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If banner_logo_02.GetKey("banner_id")&"" <> "" Then
			banner_logo_02.banner_id.CurrentValue = banner_logo_02.GetKey("banner_id") ' banner_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			banner_logo_02.CurrentFilter = banner_logo_02.KeyFilter
			Dim sSql
			sSql = banner_logo_02.SQL
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
		ViewUrl = banner_logo_02.ViewUrl("")
		EditUrl = banner_logo_02.EditUrl("")
		InlineEditUrl = banner_logo_02.InlineEditUrl
		CopyUrl = banner_logo_02.CopyUrl("")
		InlineCopyUrl = banner_logo_02.InlineCopyUrl
		DeleteUrl = banner_logo_02.DeleteUrl

		' Call Row Rendering event
		Call banner_logo_02.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' banner_id
		' banner_img
		' banner_link
		' banner_sort
		' start_date
		' end_date
		' -----------
		'  View  Row
		' -----------

		If banner_logo_02.RowType = EW_ROWTYPE_VIEW Then ' View row

			' banner_id
			banner_logo_02.banner_id.ViewValue = banner_logo_02.banner_id.CurrentValue
			banner_logo_02.banner_id.ViewCustomAttributes = ""

			' banner_img
			banner_logo_02.banner_img.ViewValue = banner_logo_02.banner_img.CurrentValue
			banner_logo_02.banner_img.ViewCustomAttributes = ""

			' banner_link
			banner_logo_02.banner_link.ViewValue = banner_logo_02.banner_link.CurrentValue
			banner_logo_02.banner_link.ViewCustomAttributes = ""

			' banner_sort
			banner_logo_02.banner_sort.ViewValue = banner_logo_02.banner_sort.CurrentValue
			banner_logo_02.banner_sort.ViewCustomAttributes = ""

			' start_date
			banner_logo_02.start_date.ViewValue = banner_logo_02.start_date.CurrentValue
			banner_logo_02.start_date.ViewCustomAttributes = ""

			' end_date
			banner_logo_02.end_date.ViewValue = banner_logo_02.end_date.CurrentValue
			banner_logo_02.end_date.ViewCustomAttributes = ""

			' View refer script
			' banner_id

			banner_logo_02.banner_id.LinkCustomAttributes = ""
			banner_logo_02.banner_id.HrefValue = ""
			banner_logo_02.banner_id.TooltipValue = ""

			' banner_img
			banner_logo_02.banner_img.LinkCustomAttributes = ""
			banner_logo_02.banner_img.HrefValue = ""
			banner_logo_02.banner_img.TooltipValue = ""

			' banner_link
			banner_logo_02.banner_link.LinkCustomAttributes = ""
			banner_logo_02.banner_link.HrefValue = ""
			banner_logo_02.banner_link.TooltipValue = ""

			' banner_sort
			banner_logo_02.banner_sort.LinkCustomAttributes = ""
			banner_logo_02.banner_sort.HrefValue = ""
			banner_logo_02.banner_sort.TooltipValue = ""

			' start_date
			banner_logo_02.start_date.LinkCustomAttributes = ""
			banner_logo_02.start_date.HrefValue = ""
			banner_logo_02.start_date.TooltipValue = ""

			' end_date
			banner_logo_02.end_date.LinkCustomAttributes = ""
			banner_logo_02.end_date.HrefValue = ""
			banner_logo_02.end_date.TooltipValue = ""
		End If

		' Call Row Rendered event
		If banner_logo_02.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call banner_logo_02.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		url = ew_CurrentUrl
		url = ew_RegExReplace("\?cmd=reset(all){0,1}$", url, "") ' Remove cmd=reset / cmd=resetall
		Call Breadcrumb.Add("list", banner_logo_02.TableVar, url, banner_logo_02.TableVar, True)
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