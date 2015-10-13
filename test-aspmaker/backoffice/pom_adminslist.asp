<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim admins_list
Set admins_list = New cadmins_list
Set Page = admins_list

' Page init processing
admins_list.Page_Init()

' Page main processing
admins_list.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
admins_list.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If admins.Export = "" Then %>
<script type="text/javascript">
// Page object
var admins_list = new ew_Page("admins_list");
admins_list.PageID = "list"; // Page ID
var EW_PAGE_ID = admins_list.PageID; // For backward compatibility
// Form object
var fadminslist = new ew_Form("fadminslist");
fadminslist.FormKeyCountName = '<%= admins_list.FormKeyCountName %>';
// Form_CustomValidate event
fadminslist.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fadminslist.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fadminslist.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
var fadminslistsrch = new ew_Form("fadminslistsrch");
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If admins.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If admins_list.ExportOptions.Visible Then %>
<div class="ewListExportOptions"><% admins_list.ExportOptions.Render "body", "", "", "", "", "" %></div>
<% End If %>
<% If (admins.Export = "") Or (EW_EXPORT_MASTER_RECORD And admins.Export = "print") Then %>
<% End If %>
<%

' Load recordset
Set admins_list.Recordset = admins_list.LoadRecordset()
	admins_list.TotalRecs = admins_list.Recordset.RecordCount
	admins_list.StartRec = 1
	If admins_list.DisplayRecs <= 0 Then ' Display all records
		admins_list.DisplayRecs = admins_list.TotalRecs
	End If
	If Not (admins.ExportAll And admins.Export <> "") Then
		admins_list.SetUpStartRec() ' Set up start record position
	End If
admins_list.RenderOtherOptions()
%>
<% If Security.IsLoggedIn() Then %>
<% If admins.Export = "" And admins.CurrentAction = "" Then %>
<form name="fadminslistsrch" id="fadminslistsrch" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewSearchTable"><tr><td>
<div class="accordion" id="fadminslistsrch_SearchGroup">
	<div class="accordion-group">
		<div class="accordion-heading">
<a class="accordion-toggle" data-toggle="collapse" data-parent="#fadminslistsrch_SearchGroup" href="#fadminslistsrch_SearchBody"><%= Language.Phrase("Search") %></a>
		</div>
		<div id="fadminslistsrch_SearchBody" class="accordion-body collapse in">
			<div class="accordion-inner">
<div id="fadminslistsrch_SearchPanel">
<input type="hidden" name="cmd" value="search">
<input type="hidden" name="t" value="admins">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewRow">
	<div class="btn-group ewButtonGroup">
	<div class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" class="input-large" value="<%= ew_HtmlEncode(admins.BasicSearch.getKeyword()) %>" placeholder="<%= ew_HtmlEncode(Language.Phrase("Search")) %>">
	<button class="btn btn-primary ewButton" name="btnsubmit" id="btnsubmit" type="submit"><%= Language.Phrase("QuickSearchBtn") %></button>
	</div>
	</div>
	<div class="btn-group ewButtonGroup">
	<a class="btn ewShowAll" href="<%= admins_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>
	</div>
</div>
<div id="xsr_2" class="ewRow">
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="="<% If admins.BasicSearch.getSearchType = "=" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If admins.BasicSearch.getSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If admins.BasicSearch.getSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
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
<% admins_list.ShowPageHeader() %>
<% admins_list.ShowMessage %>
<table class="ewGrid"><tr><td class="ewGridContent">
<form name="fadminslist" id="fadminslist" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="admins">
<div id="gmp_admins" class="ewGridMiddlePanel">
<% If admins_list.TotalRecs > 0 Then %>
<table id="tbl_adminslist" class="ewTable ewTableSeparate">
<%= admins.TableCustomInnerHtml %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call admins_list.RenderListOptions()

' Render list options (header, left)
admins_list.ListOptions.Render "header", "left", "", "", "", ""
%>
<% If admins.admin_id.Visible Then ' admin_id %>
	<% If admins.SortUrl(admins.admin_id) = "" Then %>
		<td><div id="elh_admins_admin_id" class="admins_admin_id"><div class="ewTableHeaderCaption"><%= admins.admin_id.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= admins.SortUrl(admins.admin_id) %>',1);"><div id="elh_admins_admin_id" class="admins_admin_id">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= admins.admin_id.FldCaption %></span><span class="ewTableHeaderSort"><% If admins.admin_id.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf admins.admin_id.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If admins.admin_username.Visible Then ' admin_username %>
	<% If admins.SortUrl(admins.admin_username) = "" Then %>
		<td><div id="elh_admins_admin_username" class="admins_admin_username"><div class="ewTableHeaderCaption"><%= admins.admin_username.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= admins.SortUrl(admins.admin_username) %>',1);"><div id="elh_admins_admin_username" class="admins_admin_username">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= admins.admin_username.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If admins.admin_username.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf admins.admin_username.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If admins.admin_password.Visible Then ' admin_password %>
	<% If admins.SortUrl(admins.admin_password) = "" Then %>
		<td><div id="elh_admins_admin_password" class="admins_admin_password"><div class="ewTableHeaderCaption"><%= admins.admin_password.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= admins.SortUrl(admins.admin_password) %>',1);"><div id="elh_admins_admin_password" class="admins_admin_password">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= admins.admin_password.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If admins.admin_password.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf admins.admin_password.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If admins.admin_name.Visible Then ' admin_name %>
	<% If admins.SortUrl(admins.admin_name) = "" Then %>
		<td><div id="elh_admins_admin_name" class="admins_admin_name"><div class="ewTableHeaderCaption"><%= admins.admin_name.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= admins.SortUrl(admins.admin_name) %>',1);"><div id="elh_admins_admin_name" class="admins_admin_name">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= admins.admin_name.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If admins.admin_name.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf admins.admin_name.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If admins.admin_email.Visible Then ' admin_email %>
	<% If admins.SortUrl(admins.admin_email) = "" Then %>
		<td><div id="elh_admins_admin_email" class="admins_admin_email"><div class="ewTableHeaderCaption"><%= admins.admin_email.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= admins.SortUrl(admins.admin_email) %>',1);"><div id="elh_admins_admin_email" class="admins_admin_email">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= admins.admin_email.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If admins.admin_email.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf admins.admin_email.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If admins.admin_tel.Visible Then ' admin_tel %>
	<% If admins.SortUrl(admins.admin_tel) = "" Then %>
		<td><div id="elh_admins_admin_tel" class="admins_admin_tel"><div class="ewTableHeaderCaption"><%= admins.admin_tel.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= admins.SortUrl(admins.admin_tel) %>',1);"><div id="elh_admins_admin_tel" class="admins_admin_tel">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= admins.admin_tel.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If admins.admin_tel.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf admins.admin_tel.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If admins.admin_permis.Visible Then ' admin_permis %>
	<% If admins.SortUrl(admins.admin_permis) = "" Then %>
		<td><div id="elh_admins_admin_permis" class="admins_admin_permis"><div class="ewTableHeaderCaption"><%= admins.admin_permis.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= admins.SortUrl(admins.admin_permis) %>',1);"><div id="elh_admins_admin_permis" class="admins_admin_permis">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= admins.admin_permis.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If admins.admin_permis.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf admins.admin_permis.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If admins.admin_create.Visible Then ' admin_create %>
	<% If admins.SortUrl(admins.admin_create) = "" Then %>
		<td><div id="elh_admins_admin_create" class="admins_admin_create"><div class="ewTableHeaderCaption"><%= admins.admin_create.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= admins.SortUrl(admins.admin_create) %>',1);"><div id="elh_admins_admin_create" class="admins_admin_create">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= admins.admin_create.FldCaption %></span><span class="ewTableHeaderSort"><% If admins.admin_create.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf admins.admin_create.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If admins.admin_update.Visible Then ' admin_update %>
	<% If admins.SortUrl(admins.admin_update) = "" Then %>
		<td><div id="elh_admins_admin_update" class="admins_admin_update"><div class="ewTableHeaderCaption"><%= admins.admin_update.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= admins.SortUrl(admins.admin_update) %>',1);"><div id="elh_admins_admin_update" class="admins_admin_update">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= admins.admin_update.FldCaption %></span><span class="ewTableHeaderSort"><% If admins.admin_update.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf admins.admin_update.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If admins.last_online.Visible Then ' last_online %>
	<% If admins.SortUrl(admins.last_online) = "" Then %>
		<td><div id="elh_admins_last_online" class="admins_last_online"><div class="ewTableHeaderCaption"><%= admins.last_online.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= admins.SortUrl(admins.last_online) %>',1);"><div id="elh_admins_last_online" class="admins_last_online">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= admins.last_online.FldCaption %></span><span class="ewTableHeaderSort"><% If admins.last_online.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf admins.last_online.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
admins_list.ListOptions.Render "header", "right", "", "", "", ""
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (admins.ExportAll And admins.Export <> "") Then
	admins_list.StopRec = admins_list.TotalRecs
Else

	' Set the last record to display
	If admins_list.TotalRecs > admins_list.StartRec + admins_list.DisplayRecs - 1 Then
		admins_list.StopRec = admins_list.StartRec + admins_list.DisplayRecs - 1
	Else
		admins_list.StopRec = admins_list.TotalRecs
	End If
End If

' Move to first record
admins_list.RecCnt = admins_list.StartRec - 1
If Not admins_list.Recordset.Eof Then
	admins_list.Recordset.MoveFirst
	If admins_list.StartRec > 1 Then admins_list.Recordset.Move admins_list.StartRec - 1
ElseIf Not admins.AllowAddDeleteRow And admins_list.StopRec = 0 Then
	admins_list.StopRec = admins.GridAddRowCount
End If

' Initialize Aggregate
admins.RowType = EW_ROWTYPE_AGGREGATEINIT
Call admins.ResetAttrs()
Call admins_list.RenderRow()
admins_list.RowCnt = 0

' Output date rows
Do While CLng(admins_list.RecCnt) < CLng(admins_list.StopRec)
	admins_list.RecCnt = admins_list.RecCnt + 1
	If CLng(admins_list.RecCnt) >= CLng(admins_list.StartRec) Then
		admins_list.RowCnt = admins_list.RowCnt + 1

	' Set up key count
	admins_list.KeyCount = admins_list.RowIndex
	Call admins.ResetAttrs()
	admins.CssClass = ""
	If admins.CurrentAction = "gridadd" Then
	Else
		Call admins_list.LoadRowValues(admins_list.Recordset) ' Load row values
	End If
	admins.RowType = EW_ROWTYPE_VIEW ' Render view

	' Set up row id / data-rowindex
	admins.RowAttrs.AddAttributes Array(Array("data-rowindex", admins_list.RowCnt), Array("id", "r" & admins_list.RowCnt & "_admins"), Array("data-rowtype", admins.RowType))

	' Render row
	Call admins_list.RenderRow()

	' Render list options
	Call admins_list.RenderListOptions()
%>
	<tr<%= admins.RowAttributes %>>
<%

' Render list options (body, left)
admins_list.ListOptions.Render "body", "left", admins_list.RowCnt, "", "", ""
%>
	<% If admins.admin_id.Visible Then ' admin_id %>
		<td<%= admins.admin_id.CellAttributes %>>
<span<%= admins.admin_id.ViewAttributes %>>
<%= admins.admin_id.ListViewValue %>
</span>
<a id="<%= admins_list.PageObjName & "_row_" & admins_list.RowCnt %>"></a></td>
	<% End If %>
	<% If admins.admin_username.Visible Then ' admin_username %>
		<td<%= admins.admin_username.CellAttributes %>>
<span<%= admins.admin_username.ViewAttributes %>>
<%= admins.admin_username.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If admins.admin_password.Visible Then ' admin_password %>
		<td<%= admins.admin_password.CellAttributes %>>
<span<%= admins.admin_password.ViewAttributes %>>
<%= admins.admin_password.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If admins.admin_name.Visible Then ' admin_name %>
		<td<%= admins.admin_name.CellAttributes %>>
<span<%= admins.admin_name.ViewAttributes %>>
<%= admins.admin_name.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If admins.admin_email.Visible Then ' admin_email %>
		<td<%= admins.admin_email.CellAttributes %>>
<span<%= admins.admin_email.ViewAttributes %>>
<%= admins.admin_email.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If admins.admin_tel.Visible Then ' admin_tel %>
		<td<%= admins.admin_tel.CellAttributes %>>
<span<%= admins.admin_tel.ViewAttributes %>>
<%= admins.admin_tel.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If admins.admin_permis.Visible Then ' admin_permis %>
		<td<%= admins.admin_permis.CellAttributes %>>
<span<%= admins.admin_permis.ViewAttributes %>>
<%= admins.admin_permis.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If admins.admin_create.Visible Then ' admin_create %>
		<td<%= admins.admin_create.CellAttributes %>>
<span<%= admins.admin_create.ViewAttributes %>>
<%= admins.admin_create.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If admins.admin_update.Visible Then ' admin_update %>
		<td<%= admins.admin_update.CellAttributes %>>
<span<%= admins.admin_update.ViewAttributes %>>
<%= admins.admin_update.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If admins.last_online.Visible Then ' last_online %>
		<td<%= admins.last_online.CellAttributes %>>
<span<%= admins.last_online.ViewAttributes %>>
<%= admins.last_online.ListViewValue %>
</span>
</td>
	<% End If %>
<%

' Render list options (body, right)
admins_list.ListOptions.Render "body", "right", admins_list.RowCnt, "", "", ""
%>
	</tr>
<%
	End If
	If admins.CurrentAction <> "gridadd" Then
		admins_list.Recordset.MoveNext()
	End If
Loop
%>
</tbody>
</table>
<% End If %>
<% If admins.CurrentAction = "" Then %>
<input type="hidden" name="a_list" id="a_list" value="">
<% End If %>
</div>
</form>
<%

' Close recordset and connection
admins_list.Recordset.Close
Set admins_list.Recordset = Nothing
%>
<% If admins.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If admins.CurrentAction <> "gridadd" And admins.CurrentAction <> "gridedit" Then %>
<form name="ewPagerForm" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewPager">
<tr><td>
<% If Not IsObject(admins_list.Pager) Then Set admins_list.Pager = ew_NewPrevNextPager(admins_list.StartRec, admins_list.DisplayRecs, admins_list.TotalRecs) %>
<% If admins_list.Pager.RecordCount > 0 Then %>
<table class="ewStdTable"><tbody><tr><td>
	<%= Language.Phrase("Page") %>&nbsp;
<div class="input-prepend input-append">
<!--first page button-->
	<% If admins_list.Pager.FirstButton.Enabled Then %>
	<a class="btn btn-small" href="<%= admins_list.PageUrl %>start=<%= admins_list.Pager.FirstButton.Start %>"><i class="icon-step-backward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-backward"></i></a>
	<% End If %>
<!--previous page button-->
	<% If admins_list.Pager.PrevButton.Enabled Then %>
	<a class="btn btn-small" href="<%= admins_list.PageUrl %>start=<%= admins_list.Pager.PrevButton.Start %>"><i class="icon-prev"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-prev"></i></a>
	<% End If %>
<!--current page number-->
	<input class="input-mini" type="text" name="<%= EW_TABLE_PAGE_NO %>" value="<%= admins_list.Pager.CurrentPage %>">
<!--next page button-->
	<% If admins_list.Pager.NextButton.Enabled Then %>
	<a class="btn btn-small" href="<%= admins_list.PageUrl %>start=<%= admins_list.Pager.NextButton.Start %>"><i class="icon-play"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-play"></i></a>
	<% End If %>
<!--last page button-->
	<% If admins_list.Pager.LastButton.Enabled Then %>
	<a class="btn btn-small" href="<%= admins_list.PageUrl %>start=<%= admins_list.Pager.LastButton.Start %>"><i class="icon-step-forward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-forward"></i></a>
	<% End If %>
</div>
	&nbsp;<%= Language.Phrase("of") %>&nbsp;<%= admins_list.Pager.PageCount %>
</td>
<td>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<%= Language.Phrase("Record") %>&nbsp;<%= admins_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= admins_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= admins_list.Pager.RecordCount %>
</td>
</tr></tbody></table>
<% Else %>
	<% If admins_list.SearchWhere = "0=101" Then %>
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
	admins_list.AddEditOptions.Render "body", "bottom", "", "", "", ""
	admins_list.DetailOptions.Render "body", "bottom", "", "", "", ""
	admins_list.ActionOptions.Render "body", "bottom", "", "", "", ""
%>
</div>
</div>
<% End If %>
</td></tr></table>
<% If admins.Export = "" Then %>
<script type="text/javascript">
fadminslistsrch.Init();
fadminslist.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<% End If %>
<%
admins_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If admins.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set admins_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cadmins_list

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
		TableName = "admins"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "admins_list"
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
		If admins.UseTokenInUrl Then PageUrl = PageUrl & "t=" & admins.TableVar & "&" ' add page token
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
		If admins.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (admins.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (admins.TableVar = Request.QueryString("t"))
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
		FormName = "fadminslist"
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
		If IsEmpty(admins) Then Set admins = New cadmins
		Set Table = admins

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		ExportPdfUrl = PageUrl & "export=pdf"
		AddUrl = "pom_adminsadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "pom_adminsdelete.asp"
		MultiUpdateUrl = "pom_adminsupdate.asp"

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "admins"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' List options
		Set ListOptions = New cListOptions
		ListOptions.TableVar = admins.TableVar

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = admins.TableVar
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
				admins.GridAddRowCount = gridaddcnt
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
		If UBound(admins.CustomActions.CustomArray) >= 0 Then
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
		Set admins = Nothing
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
			If admins.Export = "" Then
				SetupBreadcrumb()
			End If

			' Hide list options
			If admins.Export <> "" Then
				Call ListOptions.HideAllOptions(Array("sequence"))
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			ElseIf admins.CurrentAction = "gridadd" Or admins.CurrentAction = "gridedit" Then
				Call ListOptions.HideAllOptions(Array())
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			End If

			' Hide export options
			If admins.Export <> "" Or admins.CurrentAction <> "" Then
				ExportOptions.HideAllOptions(Array())
			End If

			' Hide other options
			If admins.Export <> "" Then
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
			Call admins.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If admins.RecordsPerPage <> "" Then
			DisplayRecs = admins.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Load search default if no existing search criteria
		If Not CheckSearchParms() Then

			' Load basic search from default
			admins.BasicSearch.Keyword = admins.BasicSearch.KeywordDefault
			admins.BasicSearch.SearchType = admins.BasicSearch.SearchTypeDefault
			admins.BasicSearch.setSearchType(admins.BasicSearch.SearchTypeDefault)
			If admins.BasicSearch.Keyword <> "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call admins.Recordset_Searching(SearchWhere)

		' Save search criteria
		If Command = "search" And Not RestoreSearch Then
			admins.SearchWhere = SearchWhere ' Save to Session
			StartRec = 1 ' Reset start record counter
			admins.StartRecordNumber = StartRec
		Else
			SearchWhere = admins.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		admins.SessionWhere = sFilter
		admins.CurrentFilter = ""
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
				sFilter = admins.KeyFilter
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
			admins.admin_id.FormValue = arrKeyFlds(0)
			If Not IsNumeric(admins.admin_id.FormValue) Then
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
			Call BuildBasicSearchSQL(sWhere, admins.admin_username, Keyword)
			Call BuildBasicSearchSQL(sWhere, admins.admin_password, Keyword)
			Call BuildBasicSearchSQL(sWhere, admins.admin_name, Keyword)
			Call BuildBasicSearchSQL(sWhere, admins.admin_email, Keyword)
			Call BuildBasicSearchSQL(sWhere, admins.admin_tel, Keyword)
			Call BuildBasicSearchSQL(sWhere, admins.admin_permis, Keyword)
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
		sSearchKeyword = admins.BasicSearch.Keyword
		sSearchType = admins.BasicSearch.SearchType
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
			admins.BasicSearch.setKeyword(sSearchKeyword)
			admins.BasicSearch.setSearchType(sSearchType)
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' Check if search parm exists
	Function CheckSearchParms()

		' Check basic search
		If admins.BasicSearch.IssetSession() Then
			CheckSearchParms = True
			Exit Function
		End If
		CheckSearchParms = False
	End Function

	' Clear all search parameters
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		admins.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' Load advanced search default values
	Function LoadAdvancedSearchDefault()
		LoadAdvancedSearchDefault = False
	End Function

	' Clear all basic search parameters
	Sub ResetBasicSearchParms()
		admins.BasicSearch.UnsetSession()
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()

		' Restore search flag
		RestoreSearch = True

		' Restore basic search values
		Call admins.BasicSearch.Load()
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
			admins.CurrentOrder = Request.QueryString("order")
			admins.CurrentOrderType = Request.QueryString("ordertype")

			' Field admin_id
			Call admins.UpdateSort(admins.admin_id)

			' Field admin_username
			Call admins.UpdateSort(admins.admin_username)

			' Field admin_password
			Call admins.UpdateSort(admins.admin_password)

			' Field admin_name
			Call admins.UpdateSort(admins.admin_name)

			' Field admin_email
			Call admins.UpdateSort(admins.admin_email)

			' Field admin_tel
			Call admins.UpdateSort(admins.admin_tel)

			' Field admin_permis
			Call admins.UpdateSort(admins.admin_permis)

			' Field admin_create
			Call admins.UpdateSort(admins.admin_create)

			' Field admin_update
			Call admins.UpdateSort(admins.admin_update)

			' Field last_online
			Call admins.UpdateSort(admins.last_online)
			admins.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = admins.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If admins.SqlOrderBy <> "" Then
				sOrderBy = admins.SqlOrderBy
				admins.SessionOrderBy = sOrderBy
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
				admins.SessionOrderBy = sOrderBy
				admins.admin_id.Sort = ""
				admins.admin_username.Sort = ""
				admins.admin_password.Sort = ""
				admins.admin_name.Sort = ""
				admins.admin_email.Sort = ""
				admins.admin_tel.Sort = ""
				admins.admin_permis.Sort = ""
				admins.admin_create.Sort = ""
				admins.admin_update.Sort = ""
				admins.last_online.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			admins.StartRecordNumber = StartRec
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
		ListOptions.GetItem("checkbox").Body = "<label class=""checkbox""><input type=""checkbox"" name=""key_m"" value=""" & ew_HtmlEncode(admins.admin_id.CurrentValue) & """ onclick='ew_ClickMultiCheckbox(event, this);'></label>"
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
			For i = 0 to UBound(admins.CustomActions.CustomArray)
				Action = admins.CustomActions.CustomArray(i)(0)
				Name = admins.CustomActions.CustomArray(i)(1)

				' Add custom action
				Call opt.Add("custom_" & Action)
				Set item = opt.GetItem("custom_" & Action)
				item.Body = "<a class=""ewAction ewCustomAction"" href="""" onclick=""ew_SubmitSelected(document.fadminslist, '" & ew_CurrentUrl & "', null, '" & Action & "');return false;"">" & Name & "</a>"
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
		sFilter = admins.GetKeyFilter
		UserAction = Request.Form("useraction") & ""
		Processed = False
		If sFilter <> "" And UserAction <> "" Then
			admins.CurrentFilter = sFilter
			sSql = admins.SQL
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
				ElseIf admins.CancelMessage <> "" Then
					FailureMessage = admins.CancelMessage
					admins.CancelMessage = ""
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
				admins.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					admins.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = admins.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			admins.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			admins.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			admins.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		admins.BasicSearch.Keyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)&""
		If admins.BasicSearch.Keyword <> "" Then Command = "search"
		admins.BasicSearch.SearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)&""
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = admins.CurrentFilter
		Call admins.Recordset_Selecting(sFilter)
		admins.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = admins.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call admins.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = admins.KeyFilter

		' Call Row Selecting event
		Call admins.Row_Selecting(sFilter)

		' Load sql based on filter
		admins.CurrentFilter = sFilter
		sSql = admins.SQL
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
		Call admins.Row_Selected(RsRow)
		admins.admin_id.DbValue = RsRow("admin_id")
		admins.admin_username.DbValue = RsRow("admin_username")
		admins.admin_password.DbValue = RsRow("admin_password")
		admins.admin_name.DbValue = RsRow("admin_name")
		admins.admin_email.DbValue = RsRow("admin_email")
		admins.admin_tel.DbValue = RsRow("admin_tel")
		admins.admin_permis.DbValue = RsRow("admin_permis")
		admins.admin_create.DbValue = RsRow("admin_create")
		admins.admin_update.DbValue = RsRow("admin_update")
		admins.last_online.DbValue = RsRow("last_online")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		admins.admin_id.m_DbValue = Rs("admin_id")
		admins.admin_username.m_DbValue = Rs("admin_username")
		admins.admin_password.m_DbValue = Rs("admin_password")
		admins.admin_name.m_DbValue = Rs("admin_name")
		admins.admin_email.m_DbValue = Rs("admin_email")
		admins.admin_tel.m_DbValue = Rs("admin_tel")
		admins.admin_permis.m_DbValue = Rs("admin_permis")
		admins.admin_create.m_DbValue = Rs("admin_create")
		admins.admin_update.m_DbValue = Rs("admin_update")
		admins.last_online.m_DbValue = Rs("last_online")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If admins.GetKey("admin_id")&"" <> "" Then
			admins.admin_id.CurrentValue = admins.GetKey("admin_id") ' admin_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			admins.CurrentFilter = admins.KeyFilter
			Dim sSql
			sSql = admins.SQL
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
		ViewUrl = admins.ViewUrl("")
		EditUrl = admins.EditUrl("")
		InlineEditUrl = admins.InlineEditUrl
		CopyUrl = admins.CopyUrl("")
		InlineCopyUrl = admins.InlineCopyUrl
		DeleteUrl = admins.DeleteUrl

		' Call Row Rendering event
		Call admins.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' admin_id
		' admin_username
		' admin_password
		' admin_name
		' admin_email
		' admin_tel
		' admin_permis
		' admin_create
		' admin_update
		' last_online
		' -----------
		'  View  Row
		' -----------

		If admins.RowType = EW_ROWTYPE_VIEW Then ' View row

			' admin_id
			admins.admin_id.ViewValue = admins.admin_id.CurrentValue
			admins.admin_id.ViewCustomAttributes = ""

			' admin_username
			admins.admin_username.ViewValue = admins.admin_username.CurrentValue
			admins.admin_username.ViewCustomAttributes = ""

			' admin_password
			admins.admin_password.ViewValue = admins.admin_password.CurrentValue
			admins.admin_password.ViewCustomAttributes = ""

			' admin_name
			admins.admin_name.ViewValue = admins.admin_name.CurrentValue
			admins.admin_name.ViewCustomAttributes = ""

			' admin_email
			admins.admin_email.ViewValue = admins.admin_email.CurrentValue
			admins.admin_email.ViewCustomAttributes = ""

			' admin_tel
			admins.admin_tel.ViewValue = admins.admin_tel.CurrentValue
			admins.admin_tel.ViewCustomAttributes = ""

			' admin_permis
			admins.admin_permis.ViewValue = admins.admin_permis.CurrentValue
			admins.admin_permis.ViewCustomAttributes = ""

			' admin_create
			admins.admin_create.ViewValue = admins.admin_create.CurrentValue
			admins.admin_create.ViewCustomAttributes = ""

			' admin_update
			admins.admin_update.ViewValue = admins.admin_update.CurrentValue
			admins.admin_update.ViewCustomAttributes = ""

			' last_online
			admins.last_online.ViewValue = admins.last_online.CurrentValue
			admins.last_online.ViewCustomAttributes = ""

			' View refer script
			' admin_id

			admins.admin_id.LinkCustomAttributes = ""
			admins.admin_id.HrefValue = ""
			admins.admin_id.TooltipValue = ""

			' admin_username
			admins.admin_username.LinkCustomAttributes = ""
			admins.admin_username.HrefValue = ""
			admins.admin_username.TooltipValue = ""

			' admin_password
			admins.admin_password.LinkCustomAttributes = ""
			admins.admin_password.HrefValue = ""
			admins.admin_password.TooltipValue = ""

			' admin_name
			admins.admin_name.LinkCustomAttributes = ""
			admins.admin_name.HrefValue = ""
			admins.admin_name.TooltipValue = ""

			' admin_email
			admins.admin_email.LinkCustomAttributes = ""
			admins.admin_email.HrefValue = ""
			admins.admin_email.TooltipValue = ""

			' admin_tel
			admins.admin_tel.LinkCustomAttributes = ""
			admins.admin_tel.HrefValue = ""
			admins.admin_tel.TooltipValue = ""

			' admin_permis
			admins.admin_permis.LinkCustomAttributes = ""
			admins.admin_permis.HrefValue = ""
			admins.admin_permis.TooltipValue = ""

			' admin_create
			admins.admin_create.LinkCustomAttributes = ""
			admins.admin_create.HrefValue = ""
			admins.admin_create.TooltipValue = ""

			' admin_update
			admins.admin_update.LinkCustomAttributes = ""
			admins.admin_update.HrefValue = ""
			admins.admin_update.TooltipValue = ""

			' last_online
			admins.last_online.LinkCustomAttributes = ""
			admins.last_online.HrefValue = ""
			admins.last_online.TooltipValue = ""
		End If

		' Call Row Rendered event
		If admins.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call admins.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		url = ew_CurrentUrl
		url = ew_RegExReplace("\?cmd=reset(all){0,1}$", url, "") ' Remove cmd=reset / cmd=resetall
		Call Breadcrumb.Add("list", admins.TableVar, url, admins.TableVar, True)
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
