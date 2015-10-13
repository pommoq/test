<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_sys_menuinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim sys_menu_list
Set sys_menu_list = New csys_menu_list
Set Page = sys_menu_list

' Page init processing
sys_menu_list.Page_Init()

' Page main processing
sys_menu_list.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
sys_menu_list.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If sys_menu.Export = "" Then %>
<script type="text/javascript">
// Page object
var sys_menu_list = new ew_Page("sys_menu_list");
sys_menu_list.PageID = "list"; // Page ID
var EW_PAGE_ID = sys_menu_list.PageID; // For backward compatibility
// Form object
var fsys_menulist = new ew_Form("fsys_menulist");
fsys_menulist.FormKeyCountName = '<%= sys_menu_list.FormKeyCountName %>';
// Form_CustomValidate event
fsys_menulist.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fsys_menulist.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fsys_menulist.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
var fsys_menulistsrch = new ew_Form("fsys_menulistsrch");
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If sys_menu.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If sys_menu_list.ExportOptions.Visible Then %>
<div class="ewListExportOptions"><% sys_menu_list.ExportOptions.Render "body", "", "", "", "", "" %></div>
<% End If %>
<% If (sys_menu.Export = "") Or (EW_EXPORT_MASTER_RECORD And sys_menu.Export = "print") Then %>
<% End If %>
<%

' Load recordset
Set sys_menu_list.Recordset = sys_menu_list.LoadRecordset()
	sys_menu_list.TotalRecs = sys_menu_list.Recordset.RecordCount
	sys_menu_list.StartRec = 1
	If sys_menu_list.DisplayRecs <= 0 Then ' Display all records
		sys_menu_list.DisplayRecs = sys_menu_list.TotalRecs
	End If
	If Not (sys_menu.ExportAll And sys_menu.Export <> "") Then
		sys_menu_list.SetUpStartRec() ' Set up start record position
	End If
sys_menu_list.RenderOtherOptions()
%>
<% If Security.IsLoggedIn() Then %>
<% If sys_menu.Export = "" And sys_menu.CurrentAction = "" Then %>
<form name="fsys_menulistsrch" id="fsys_menulistsrch" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewSearchTable"><tr><td>
<div class="accordion" id="fsys_menulistsrch_SearchGroup">
	<div class="accordion-group">
		<div class="accordion-heading">
<a class="accordion-toggle" data-toggle="collapse" data-parent="#fsys_menulistsrch_SearchGroup" href="#fsys_menulistsrch_SearchBody"><%= Language.Phrase("Search") %></a>
		</div>
		<div id="fsys_menulistsrch_SearchBody" class="accordion-body collapse in">
			<div class="accordion-inner">
<div id="fsys_menulistsrch_SearchPanel">
<input type="hidden" name="cmd" value="search">
<input type="hidden" name="t" value="sys_menu">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewRow">
	<div class="btn-group ewButtonGroup">
	<div class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" class="input-large" value="<%= ew_HtmlEncode(sys_menu.BasicSearch.getKeyword()) %>" placeholder="<%= ew_HtmlEncode(Language.Phrase("Search")) %>">
	<button class="btn btn-primary ewButton" name="btnsubmit" id="btnsubmit" type="submit"><%= Language.Phrase("QuickSearchBtn") %></button>
	</div>
	</div>
	<div class="btn-group ewButtonGroup">
	<a class="btn ewShowAll" href="<%= sys_menu_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>
	</div>
</div>
<div id="xsr_2" class="ewRow">
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="="<% If sys_menu.BasicSearch.getSearchType = "=" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If sys_menu.BasicSearch.getSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If sys_menu.BasicSearch.getSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
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
<% sys_menu_list.ShowPageHeader() %>
<% sys_menu_list.ShowMessage %>
<table class="ewGrid"><tr><td class="ewGridContent">
<form name="fsys_menulist" id="fsys_menulist" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="sys_menu">
<div id="gmp_sys_menu" class="ewGridMiddlePanel">
<% If sys_menu_list.TotalRecs > 0 Then %>
<table id="tbl_sys_menulist" class="ewTable ewTableSeparate">
<%= sys_menu.TableCustomInnerHtml %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call sys_menu_list.RenderListOptions()

' Render list options (header, left)
sys_menu_list.ListOptions.Render "header", "left", "", "", "", ""
%>
<% If sys_menu.menu_id.Visible Then ' menu_id %>
	<% If sys_menu.SortUrl(sys_menu.menu_id) = "" Then %>
		<td><div id="elh_sys_menu_menu_id" class="sys_menu_menu_id"><div class="ewTableHeaderCaption"><%= sys_menu.menu_id.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= sys_menu.SortUrl(sys_menu.menu_id) %>',1);"><div id="elh_sys_menu_menu_id" class="sys_menu_menu_id">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= sys_menu.menu_id.FldCaption %></span><span class="ewTableHeaderSort"><% If sys_menu.menu_id.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf sys_menu.menu_id.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If sys_menu.menu_name.Visible Then ' menu_name %>
	<% If sys_menu.SortUrl(sys_menu.menu_name) = "" Then %>
		<td><div id="elh_sys_menu_menu_name" class="sys_menu_menu_name"><div class="ewTableHeaderCaption"><%= sys_menu.menu_name.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= sys_menu.SortUrl(sys_menu.menu_name) %>',1);"><div id="elh_sys_menu_menu_name" class="sys_menu_menu_name">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= sys_menu.menu_name.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If sys_menu.menu_name.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf sys_menu.menu_name.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If sys_menu.menu_parent_id.Visible Then ' menu_parent_id %>
	<% If sys_menu.SortUrl(sys_menu.menu_parent_id) = "" Then %>
		<td><div id="elh_sys_menu_menu_parent_id" class="sys_menu_menu_parent_id"><div class="ewTableHeaderCaption"><%= sys_menu.menu_parent_id.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= sys_menu.SortUrl(sys_menu.menu_parent_id) %>',1);"><div id="elh_sys_menu_menu_parent_id" class="sys_menu_menu_parent_id">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= sys_menu.menu_parent_id.FldCaption %></span><span class="ewTableHeaderSort"><% If sys_menu.menu_parent_id.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf sys_menu.menu_parent_id.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If sys_menu.menu_thai.Visible Then ' menu_thai %>
	<% If sys_menu.SortUrl(sys_menu.menu_thai) = "" Then %>
		<td><div id="elh_sys_menu_menu_thai" class="sys_menu_menu_thai"><div class="ewTableHeaderCaption"><%= sys_menu.menu_thai.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= sys_menu.SortUrl(sys_menu.menu_thai) %>',1);"><div id="elh_sys_menu_menu_thai" class="sys_menu_menu_thai">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= sys_menu.menu_thai.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If sys_menu.menu_thai.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf sys_menu.menu_thai.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If sys_menu.menu_idname.Visible Then ' menu_idname %>
	<% If sys_menu.SortUrl(sys_menu.menu_idname) = "" Then %>
		<td><div id="elh_sys_menu_menu_idname" class="sys_menu_menu_idname"><div class="ewTableHeaderCaption"><%= sys_menu.menu_idname.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= sys_menu.SortUrl(sys_menu.menu_idname) %>',1);"><div id="elh_sys_menu_menu_idname" class="sys_menu_menu_idname">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= sys_menu.menu_idname.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If sys_menu.menu_idname.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf sys_menu.menu_idname.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If sys_menu.menu_filename.Visible Then ' menu_filename %>
	<% If sys_menu.SortUrl(sys_menu.menu_filename) = "" Then %>
		<td><div id="elh_sys_menu_menu_filename" class="sys_menu_menu_filename"><div class="ewTableHeaderCaption"><%= sys_menu.menu_filename.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= sys_menu.SortUrl(sys_menu.menu_filename) %>',1);"><div id="elh_sys_menu_menu_filename" class="sys_menu_menu_filename">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= sys_menu.menu_filename.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If sys_menu.menu_filename.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf sys_menu.menu_filename.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If sys_menu.target.Visible Then ' target %>
	<% If sys_menu.SortUrl(sys_menu.target) = "" Then %>
		<td><div id="elh_sys_menu_target" class="sys_menu_target"><div class="ewTableHeaderCaption"><%= sys_menu.target.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= sys_menu.SortUrl(sys_menu.target) %>',1);"><div id="elh_sys_menu_target" class="sys_menu_target">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= sys_menu.target.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If sys_menu.target.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf sys_menu.target.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If sys_menu.OrderList.Visible Then ' OrderList %>
	<% If sys_menu.SortUrl(sys_menu.OrderList) = "" Then %>
		<td><div id="elh_sys_menu_OrderList" class="sys_menu_OrderList"><div class="ewTableHeaderCaption"><%= sys_menu.OrderList.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= sys_menu.SortUrl(sys_menu.OrderList) %>',1);"><div id="elh_sys_menu_OrderList" class="sys_menu_OrderList">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= sys_menu.OrderList.FldCaption %></span><span class="ewTableHeaderSort"><% If sys_menu.OrderList.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf sys_menu.OrderList.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
sys_menu_list.ListOptions.Render "header", "right", "", "", "", ""
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (sys_menu.ExportAll And sys_menu.Export <> "") Then
	sys_menu_list.StopRec = sys_menu_list.TotalRecs
Else

	' Set the last record to display
	If sys_menu_list.TotalRecs > sys_menu_list.StartRec + sys_menu_list.DisplayRecs - 1 Then
		sys_menu_list.StopRec = sys_menu_list.StartRec + sys_menu_list.DisplayRecs - 1
	Else
		sys_menu_list.StopRec = sys_menu_list.TotalRecs
	End If
End If

' Move to first record
sys_menu_list.RecCnt = sys_menu_list.StartRec - 1
If Not sys_menu_list.Recordset.Eof Then
	sys_menu_list.Recordset.MoveFirst
	If sys_menu_list.StartRec > 1 Then sys_menu_list.Recordset.Move sys_menu_list.StartRec - 1
ElseIf Not sys_menu.AllowAddDeleteRow And sys_menu_list.StopRec = 0 Then
	sys_menu_list.StopRec = sys_menu.GridAddRowCount
End If

' Initialize Aggregate
sys_menu.RowType = EW_ROWTYPE_AGGREGATEINIT
Call sys_menu.ResetAttrs()
Call sys_menu_list.RenderRow()
sys_menu_list.RowCnt = 0

' Output date rows
Do While CLng(sys_menu_list.RecCnt) < CLng(sys_menu_list.StopRec)
	sys_menu_list.RecCnt = sys_menu_list.RecCnt + 1
	If CLng(sys_menu_list.RecCnt) >= CLng(sys_menu_list.StartRec) Then
		sys_menu_list.RowCnt = sys_menu_list.RowCnt + 1

	' Set up key count
	sys_menu_list.KeyCount = sys_menu_list.RowIndex
	Call sys_menu.ResetAttrs()
	sys_menu.CssClass = ""
	If sys_menu.CurrentAction = "gridadd" Then
	Else
		Call sys_menu_list.LoadRowValues(sys_menu_list.Recordset) ' Load row values
	End If
	sys_menu.RowType = EW_ROWTYPE_VIEW ' Render view

	' Set up row id / data-rowindex
	sys_menu.RowAttrs.AddAttributes Array(Array("data-rowindex", sys_menu_list.RowCnt), Array("id", "r" & sys_menu_list.RowCnt & "_sys_menu"), Array("data-rowtype", sys_menu.RowType))

	' Render row
	Call sys_menu_list.RenderRow()

	' Render list options
	Call sys_menu_list.RenderListOptions()
%>
	<tr<%= sys_menu.RowAttributes %>>
<%

' Render list options (body, left)
sys_menu_list.ListOptions.Render "body", "left", sys_menu_list.RowCnt, "", "", ""
%>
	<% If sys_menu.menu_id.Visible Then ' menu_id %>
		<td<%= sys_menu.menu_id.CellAttributes %>>
<span<%= sys_menu.menu_id.ViewAttributes %>>
<%= sys_menu.menu_id.ListViewValue %>
</span>
<a id="<%= sys_menu_list.PageObjName & "_row_" & sys_menu_list.RowCnt %>"></a></td>
	<% End If %>
	<% If sys_menu.menu_name.Visible Then ' menu_name %>
		<td<%= sys_menu.menu_name.CellAttributes %>>
<span<%= sys_menu.menu_name.ViewAttributes %>>
<%= sys_menu.menu_name.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If sys_menu.menu_parent_id.Visible Then ' menu_parent_id %>
		<td<%= sys_menu.menu_parent_id.CellAttributes %>>
<span<%= sys_menu.menu_parent_id.ViewAttributes %>>
<%= sys_menu.menu_parent_id.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If sys_menu.menu_thai.Visible Then ' menu_thai %>
		<td<%= sys_menu.menu_thai.CellAttributes %>>
<span<%= sys_menu.menu_thai.ViewAttributes %>>
<%= sys_menu.menu_thai.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If sys_menu.menu_idname.Visible Then ' menu_idname %>
		<td<%= sys_menu.menu_idname.CellAttributes %>>
<span<%= sys_menu.menu_idname.ViewAttributes %>>
<%= sys_menu.menu_idname.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If sys_menu.menu_filename.Visible Then ' menu_filename %>
		<td<%= sys_menu.menu_filename.CellAttributes %>>
<span<%= sys_menu.menu_filename.ViewAttributes %>>
<%= sys_menu.menu_filename.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If sys_menu.target.Visible Then ' target %>
		<td<%= sys_menu.target.CellAttributes %>>
<span<%= sys_menu.target.ViewAttributes %>>
<%= sys_menu.target.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If sys_menu.OrderList.Visible Then ' OrderList %>
		<td<%= sys_menu.OrderList.CellAttributes %>>
<span<%= sys_menu.OrderList.ViewAttributes %>>
<%= sys_menu.OrderList.ListViewValue %>
</span>
</td>
	<% End If %>
<%

' Render list options (body, right)
sys_menu_list.ListOptions.Render "body", "right", sys_menu_list.RowCnt, "", "", ""
%>
	</tr>
<%
	End If
	If sys_menu.CurrentAction <> "gridadd" Then
		sys_menu_list.Recordset.MoveNext()
	End If
Loop
%>
</tbody>
</table>
<% End If %>
<% If sys_menu.CurrentAction = "" Then %>
<input type="hidden" name="a_list" id="a_list" value="">
<% End If %>
</div>
</form>
<%

' Close recordset and connection
sys_menu_list.Recordset.Close
Set sys_menu_list.Recordset = Nothing
%>
<% If sys_menu.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If sys_menu.CurrentAction <> "gridadd" And sys_menu.CurrentAction <> "gridedit" Then %>
<form name="ewPagerForm" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewPager">
<tr><td>
<% If Not IsObject(sys_menu_list.Pager) Then Set sys_menu_list.Pager = ew_NewPrevNextPager(sys_menu_list.StartRec, sys_menu_list.DisplayRecs, sys_menu_list.TotalRecs) %>
<% If sys_menu_list.Pager.RecordCount > 0 Then %>
<table class="ewStdTable"><tbody><tr><td>
	<%= Language.Phrase("Page") %>&nbsp;
<div class="input-prepend input-append">
<!--first page button-->
	<% If sys_menu_list.Pager.FirstButton.Enabled Then %>
	<a class="btn btn-small" href="<%= sys_menu_list.PageUrl %>start=<%= sys_menu_list.Pager.FirstButton.Start %>"><i class="icon-step-backward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-backward"></i></a>
	<% End If %>
<!--previous page button-->
	<% If sys_menu_list.Pager.PrevButton.Enabled Then %>
	<a class="btn btn-small" href="<%= sys_menu_list.PageUrl %>start=<%= sys_menu_list.Pager.PrevButton.Start %>"><i class="icon-prev"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-prev"></i></a>
	<% End If %>
<!--current page number-->
	<input class="input-mini" type="text" name="<%= EW_TABLE_PAGE_NO %>" value="<%= sys_menu_list.Pager.CurrentPage %>">
<!--next page button-->
	<% If sys_menu_list.Pager.NextButton.Enabled Then %>
	<a class="btn btn-small" href="<%= sys_menu_list.PageUrl %>start=<%= sys_menu_list.Pager.NextButton.Start %>"><i class="icon-play"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-play"></i></a>
	<% End If %>
<!--last page button-->
	<% If sys_menu_list.Pager.LastButton.Enabled Then %>
	<a class="btn btn-small" href="<%= sys_menu_list.PageUrl %>start=<%= sys_menu_list.Pager.LastButton.Start %>"><i class="icon-step-forward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-forward"></i></a>
	<% End If %>
</div>
	&nbsp;<%= Language.Phrase("of") %>&nbsp;<%= sys_menu_list.Pager.PageCount %>
</td>
<td>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<%= Language.Phrase("Record") %>&nbsp;<%= sys_menu_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= sys_menu_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= sys_menu_list.Pager.RecordCount %>
</td>
</tr></tbody></table>
<% Else %>
	<% If sys_menu_list.SearchWhere = "0=101" Then %>
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
	sys_menu_list.AddEditOptions.Render "body", "bottom", "", "", "", ""
	sys_menu_list.DetailOptions.Render "body", "bottom", "", "", "", ""
	sys_menu_list.ActionOptions.Render "body", "bottom", "", "", "", ""
%>
</div>
</div>
<% End If %>
</td></tr></table>
<% If sys_menu.Export = "" Then %>
<script type="text/javascript">
fsys_menulistsrch.Init();
fsys_menulist.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<% End If %>
<%
sys_menu_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If sys_menu.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set sys_menu_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class csys_menu_list

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
		TableName = "sys_menu"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "sys_menu_list"
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
		If sys_menu.UseTokenInUrl Then PageUrl = PageUrl & "t=" & sys_menu.TableVar & "&" ' add page token
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
		If sys_menu.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (sys_menu.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (sys_menu.TableVar = Request.QueryString("t"))
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
		FormName = "fsys_menulist"
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
		If IsEmpty(sys_menu) Then Set sys_menu = New csys_menu
		Set Table = sys_menu

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		ExportPdfUrl = PageUrl & "export=pdf"
		AddUrl = "pom_sys_menuadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "pom_sys_menudelete.asp"
		MultiUpdateUrl = "pom_sys_menuupdate.asp"

		' Initialize other table object
		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "sys_menu"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' List options
		Set ListOptions = New cListOptions
		ListOptions.TableVar = sys_menu.TableVar

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = sys_menu.TableVar
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
				sys_menu.GridAddRowCount = gridaddcnt
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
		If UBound(sys_menu.CustomActions.CustomArray) >= 0 Then
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
		Set sys_menu = Nothing
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
			If sys_menu.Export = "" Then
				SetupBreadcrumb()
			End If

			' Hide list options
			If sys_menu.Export <> "" Then
				Call ListOptions.HideAllOptions(Array("sequence"))
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			ElseIf sys_menu.CurrentAction = "gridadd" Or sys_menu.CurrentAction = "gridedit" Then
				Call ListOptions.HideAllOptions(Array())
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			End If

			' Hide export options
			If sys_menu.Export <> "" Or sys_menu.CurrentAction <> "" Then
				ExportOptions.HideAllOptions(Array())
			End If

			' Hide other options
			If sys_menu.Export <> "" Then
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
			Call sys_menu.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If sys_menu.RecordsPerPage <> "" Then
			DisplayRecs = sys_menu.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Load search default if no existing search criteria
		If Not CheckSearchParms() Then

			' Load basic search from default
			sys_menu.BasicSearch.Keyword = sys_menu.BasicSearch.KeywordDefault
			sys_menu.BasicSearch.SearchType = sys_menu.BasicSearch.SearchTypeDefault
			sys_menu.BasicSearch.setSearchType(sys_menu.BasicSearch.SearchTypeDefault)
			If sys_menu.BasicSearch.Keyword <> "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call sys_menu.Recordset_Searching(SearchWhere)

		' Save search criteria
		If Command = "search" And Not RestoreSearch Then
			sys_menu.SearchWhere = SearchWhere ' Save to Session
			StartRec = 1 ' Reset start record counter
			sys_menu.StartRecordNumber = StartRec
		Else
			SearchWhere = sys_menu.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		sys_menu.SessionWhere = sFilter
		sys_menu.CurrentFilter = ""
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
				sFilter = sys_menu.KeyFilter
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
			sys_menu.menu_id.FormValue = arrKeyFlds(0)
			If Not IsNumeric(sys_menu.menu_id.FormValue) Then
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
			Call BuildBasicSearchSQL(sWhere, sys_menu.menu_name, Keyword)
			Call BuildBasicSearchSQL(sWhere, sys_menu.menu_thai, Keyword)
			Call BuildBasicSearchSQL(sWhere, sys_menu.menu_idname, Keyword)
			Call BuildBasicSearchSQL(sWhere, sys_menu.menu_filename, Keyword)
			Call BuildBasicSearchSQL(sWhere, sys_menu.target, Keyword)
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
		sSearchKeyword = sys_menu.BasicSearch.Keyword
		sSearchType = sys_menu.BasicSearch.SearchType
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
			sys_menu.BasicSearch.setKeyword(sSearchKeyword)
			sys_menu.BasicSearch.setSearchType(sSearchType)
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' Check if search parm exists
	Function CheckSearchParms()

		' Check basic search
		If sys_menu.BasicSearch.IssetSession() Then
			CheckSearchParms = True
			Exit Function
		End If
		CheckSearchParms = False
	End Function

	' Clear all search parameters
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		sys_menu.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' Load advanced search default values
	Function LoadAdvancedSearchDefault()
		LoadAdvancedSearchDefault = False
	End Function

	' Clear all basic search parameters
	Sub ResetBasicSearchParms()
		sys_menu.BasicSearch.UnsetSession()
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()

		' Restore search flag
		RestoreSearch = True

		' Restore basic search values
		Call sys_menu.BasicSearch.Load()
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
			sys_menu.CurrentOrder = Request.QueryString("order")
			sys_menu.CurrentOrderType = Request.QueryString("ordertype")

			' Field menu_id
			Call sys_menu.UpdateSort(sys_menu.menu_id)

			' Field menu_name
			Call sys_menu.UpdateSort(sys_menu.menu_name)

			' Field menu_parent_id
			Call sys_menu.UpdateSort(sys_menu.menu_parent_id)

			' Field menu_thai
			Call sys_menu.UpdateSort(sys_menu.menu_thai)

			' Field menu_idname
			Call sys_menu.UpdateSort(sys_menu.menu_idname)

			' Field menu_filename
			Call sys_menu.UpdateSort(sys_menu.menu_filename)

			' Field target
			Call sys_menu.UpdateSort(sys_menu.target)

			' Field OrderList
			Call sys_menu.UpdateSort(sys_menu.OrderList)
			sys_menu.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = sys_menu.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If sys_menu.SqlOrderBy <> "" Then
				sOrderBy = sys_menu.SqlOrderBy
				sys_menu.SessionOrderBy = sOrderBy
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
				sys_menu.SessionOrderBy = sOrderBy
				sys_menu.menu_id.Sort = ""
				sys_menu.menu_name.Sort = ""
				sys_menu.menu_parent_id.Sort = ""
				sys_menu.menu_thai.Sort = ""
				sys_menu.menu_idname.Sort = ""
				sys_menu.menu_filename.Sort = ""
				sys_menu.target.Sort = ""
				sys_menu.OrderList.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			sys_menu.StartRecordNumber = StartRec
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
		ListOptions.GetItem("checkbox").Body = "<label class=""checkbox""><input type=""checkbox"" name=""key_m"" value=""" & ew_HtmlEncode(sys_menu.menu_id.CurrentValue) & """ onclick='ew_ClickMultiCheckbox(event, this);'></label>"
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
			For i = 0 to UBound(sys_menu.CustomActions.CustomArray)
				Action = sys_menu.CustomActions.CustomArray(i)(0)
				Name = sys_menu.CustomActions.CustomArray(i)(1)

				' Add custom action
				Call opt.Add("custom_" & Action)
				Set item = opt.GetItem("custom_" & Action)
				item.Body = "<a class=""ewAction ewCustomAction"" href="""" onclick=""ew_SubmitSelected(document.fsys_menulist, '" & ew_CurrentUrl & "', null, '" & Action & "');return false;"">" & Name & "</a>"
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
		sFilter = sys_menu.GetKeyFilter
		UserAction = Request.Form("useraction") & ""
		Processed = False
		If sFilter <> "" And UserAction <> "" Then
			sys_menu.CurrentFilter = sFilter
			sSql = sys_menu.SQL
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
				ElseIf sys_menu.CancelMessage <> "" Then
					FailureMessage = sys_menu.CancelMessage
					sys_menu.CancelMessage = ""
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
				sys_menu.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					sys_menu.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = sys_menu.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			sys_menu.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			sys_menu.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			sys_menu.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		sys_menu.BasicSearch.Keyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)&""
		If sys_menu.BasicSearch.Keyword <> "" Then Command = "search"
		sys_menu.BasicSearch.SearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)&""
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = sys_menu.CurrentFilter
		Call sys_menu.Recordset_Selecting(sFilter)
		sys_menu.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = sys_menu.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call sys_menu.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = sys_menu.KeyFilter

		' Call Row Selecting event
		Call sys_menu.Row_Selecting(sFilter)

		' Load sql based on filter
		sys_menu.CurrentFilter = sFilter
		sSql = sys_menu.SQL
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
		Call sys_menu.Row_Selected(RsRow)
		sys_menu.menu_id.DbValue = RsRow("menu_id")
		sys_menu.menu_name.DbValue = RsRow("menu_name")
		sys_menu.menu_parent_id.DbValue = RsRow("menu_parent_id")
		sys_menu.menu_thai.DbValue = RsRow("menu_thai")
		sys_menu.menu_idname.DbValue = RsRow("menu_idname")
		sys_menu.menu_filename.DbValue = RsRow("menu_filename")
		sys_menu.target.DbValue = RsRow("target")
		sys_menu.OrderList.DbValue = RsRow("OrderList")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		sys_menu.menu_id.m_DbValue = Rs("menu_id")
		sys_menu.menu_name.m_DbValue = Rs("menu_name")
		sys_menu.menu_parent_id.m_DbValue = Rs("menu_parent_id")
		sys_menu.menu_thai.m_DbValue = Rs("menu_thai")
		sys_menu.menu_idname.m_DbValue = Rs("menu_idname")
		sys_menu.menu_filename.m_DbValue = Rs("menu_filename")
		sys_menu.target.m_DbValue = Rs("target")
		sys_menu.OrderList.m_DbValue = Rs("OrderList")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If sys_menu.GetKey("menu_id")&"" <> "" Then
			sys_menu.menu_id.CurrentValue = sys_menu.GetKey("menu_id") ' menu_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			sys_menu.CurrentFilter = sys_menu.KeyFilter
			Dim sSql
			sSql = sys_menu.SQL
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
		ViewUrl = sys_menu.ViewUrl("")
		EditUrl = sys_menu.EditUrl("")
		InlineEditUrl = sys_menu.InlineEditUrl
		CopyUrl = sys_menu.CopyUrl("")
		InlineCopyUrl = sys_menu.InlineCopyUrl
		DeleteUrl = sys_menu.DeleteUrl

		' Convert decimal values if posted back
		If sys_menu.OrderList.FormValue = sys_menu.OrderList.CurrentValue And IsNumeric(sys_menu.OrderList.CurrentValue) Then
			sys_menu.OrderList.CurrentValue = ew_StrToFloat(sys_menu.OrderList.CurrentValue)
		End If

		' Call Row Rendering event
		Call sys_menu.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' menu_id
		' menu_name
		' menu_parent_id
		' menu_thai
		' menu_idname
		' menu_filename
		' target
		' OrderList
		' -----------
		'  View  Row
		' -----------

		If sys_menu.RowType = EW_ROWTYPE_VIEW Then ' View row

			' menu_id
			sys_menu.menu_id.ViewValue = sys_menu.menu_id.CurrentValue
			sys_menu.menu_id.ViewCustomAttributes = ""

			' menu_name
			sys_menu.menu_name.ViewValue = sys_menu.menu_name.CurrentValue
			sys_menu.menu_name.ViewCustomAttributes = ""

			' menu_parent_id
			sys_menu.menu_parent_id.ViewValue = sys_menu.menu_parent_id.CurrentValue
			sys_menu.menu_parent_id.ViewCustomAttributes = ""

			' menu_thai
			sys_menu.menu_thai.ViewValue = sys_menu.menu_thai.CurrentValue
			sys_menu.menu_thai.ViewCustomAttributes = ""

			' menu_idname
			sys_menu.menu_idname.ViewValue = sys_menu.menu_idname.CurrentValue
			sys_menu.menu_idname.ViewCustomAttributes = ""

			' menu_filename
			sys_menu.menu_filename.ViewValue = sys_menu.menu_filename.CurrentValue
			sys_menu.menu_filename.ViewCustomAttributes = ""

			' target
			sys_menu.target.ViewValue = sys_menu.target.CurrentValue
			sys_menu.target.ViewCustomAttributes = ""

			' OrderList
			sys_menu.OrderList.ViewValue = sys_menu.OrderList.CurrentValue
			sys_menu.OrderList.ViewCustomAttributes = ""

			' View refer script
			' menu_id

			sys_menu.menu_id.LinkCustomAttributes = ""
			sys_menu.menu_id.HrefValue = ""
			sys_menu.menu_id.TooltipValue = ""

			' menu_name
			sys_menu.menu_name.LinkCustomAttributes = ""
			sys_menu.menu_name.HrefValue = ""
			sys_menu.menu_name.TooltipValue = ""

			' menu_parent_id
			sys_menu.menu_parent_id.LinkCustomAttributes = ""
			sys_menu.menu_parent_id.HrefValue = ""
			sys_menu.menu_parent_id.TooltipValue = ""

			' menu_thai
			sys_menu.menu_thai.LinkCustomAttributes = ""
			sys_menu.menu_thai.HrefValue = ""
			sys_menu.menu_thai.TooltipValue = ""

			' menu_idname
			sys_menu.menu_idname.LinkCustomAttributes = ""
			sys_menu.menu_idname.HrefValue = ""
			sys_menu.menu_idname.TooltipValue = ""

			' menu_filename
			sys_menu.menu_filename.LinkCustomAttributes = ""
			sys_menu.menu_filename.HrefValue = ""
			sys_menu.menu_filename.TooltipValue = ""

			' target
			sys_menu.target.LinkCustomAttributes = ""
			sys_menu.target.HrefValue = ""
			sys_menu.target.TooltipValue = ""

			' OrderList
			sys_menu.OrderList.LinkCustomAttributes = ""
			sys_menu.OrderList.HrefValue = ""
			sys_menu.OrderList.TooltipValue = ""
		End If

		' Call Row Rendered event
		If sys_menu.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call sys_menu.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		url = ew_CurrentUrl
		url = ew_RegExReplace("\?cmd=reset(all){0,1}$", url, "") ' Remove cmd=reset / cmd=resetall
		Call Breadcrumb.Add("list", sys_menu.TableVar, url, sys_menu.TableVar, True)
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
