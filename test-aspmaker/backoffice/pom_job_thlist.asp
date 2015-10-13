<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_job_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim job_th_list
Set job_th_list = New cjob_th_list
Set Page = job_th_list

' Page init processing
job_th_list.Page_Init()

' Page main processing
job_th_list.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
job_th_list.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If job_th.Export = "" Then %>
<script type="text/javascript">
// Page object
var job_th_list = new ew_Page("job_th_list");
job_th_list.PageID = "list"; // Page ID
var EW_PAGE_ID = job_th_list.PageID; // For backward compatibility
// Form object
var fjob_thlist = new ew_Form("fjob_thlist");
fjob_thlist.FormKeyCountName = '<%= job_th_list.FormKeyCountName %>';
// Form_CustomValidate event
fjob_thlist.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fjob_thlist.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fjob_thlist.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
var fjob_thlistsrch = new ew_Form("fjob_thlistsrch");
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If job_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If job_th_list.ExportOptions.Visible Then %>
<div class="ewListExportOptions"><% job_th_list.ExportOptions.Render "body", "", "", "", "", "" %></div>
<% End If %>
<% If (job_th.Export = "") Or (EW_EXPORT_MASTER_RECORD And job_th.Export = "print") Then %>
<% End If %>
<%

' Load recordset
Set job_th_list.Recordset = job_th_list.LoadRecordset()
	job_th_list.TotalRecs = job_th_list.Recordset.RecordCount
	job_th_list.StartRec = 1
	If job_th_list.DisplayRecs <= 0 Then ' Display all records
		job_th_list.DisplayRecs = job_th_list.TotalRecs
	End If
	If Not (job_th.ExportAll And job_th.Export <> "") Then
		job_th_list.SetUpStartRec() ' Set up start record position
	End If
job_th_list.RenderOtherOptions()
%>
<% If Security.IsLoggedIn() Then %>
<% If job_th.Export = "" And job_th.CurrentAction = "" Then %>
<form name="fjob_thlistsrch" id="fjob_thlistsrch" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewSearchTable"><tr><td>
<div class="accordion" id="fjob_thlistsrch_SearchGroup">
	<div class="accordion-group">
		<div class="accordion-heading">
<a class="accordion-toggle" data-toggle="collapse" data-parent="#fjob_thlistsrch_SearchGroup" href="#fjob_thlistsrch_SearchBody"><%= Language.Phrase("Search") %></a>
		</div>
		<div id="fjob_thlistsrch_SearchBody" class="accordion-body collapse in">
			<div class="accordion-inner">
<div id="fjob_thlistsrch_SearchPanel">
<input type="hidden" name="cmd" value="search">
<input type="hidden" name="t" value="job_th">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewRow">
	<div class="btn-group ewButtonGroup">
	<div class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" class="input-large" value="<%= ew_HtmlEncode(job_th.BasicSearch.getKeyword()) %>" placeholder="<%= ew_HtmlEncode(Language.Phrase("Search")) %>">
	<button class="btn btn-primary ewButton" name="btnsubmit" id="btnsubmit" type="submit"><%= Language.Phrase("QuickSearchBtn") %></button>
	</div>
	</div>
	<div class="btn-group ewButtonGroup">
	<a class="btn ewShowAll" href="<%= job_th_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>
	</div>
</div>
<div id="xsr_2" class="ewRow">
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="="<% If job_th.BasicSearch.getSearchType = "=" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If job_th.BasicSearch.getSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If job_th.BasicSearch.getSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
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
<% job_th_list.ShowPageHeader() %>
<% job_th_list.ShowMessage %>
<table class="ewGrid"><tr><td class="ewGridContent">
<form name="fjob_thlist" id="fjob_thlist" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="job_th">
<div id="gmp_job_th" class="ewGridMiddlePanel">
<% If job_th_list.TotalRecs > 0 Then %>
<table id="tbl_job_thlist" class="ewTable ewTableSeparate">
<%= job_th.TableCustomInnerHtml %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call job_th_list.RenderListOptions()

' Render list options (header, left)
job_th_list.ListOptions.Render "header", "left", "", "", "", ""
%>
<% If job_th.job_id.Visible Then ' job_id %>
	<% If job_th.SortUrl(job_th.job_id) = "" Then %>
		<td><div id="elh_job_th_job_id" class="job_th_job_id"><div class="ewTableHeaderCaption"><%= job_th.job_id.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= job_th.SortUrl(job_th.job_id) %>',1);"><div id="elh_job_th_job_id" class="job_th_job_id">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= job_th.job_id.FldCaption %></span><span class="ewTableHeaderSort"><% If job_th.job_id.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf job_th.job_id.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If job_th.company_id.Visible Then ' company_id %>
	<% If job_th.SortUrl(job_th.company_id) = "" Then %>
		<td><div id="elh_job_th_company_id" class="job_th_company_id"><div class="ewTableHeaderCaption"><%= job_th.company_id.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= job_th.SortUrl(job_th.company_id) %>',1);"><div id="elh_job_th_company_id" class="job_th_company_id">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= job_th.company_id.FldCaption %></span><span class="ewTableHeaderSort"><% If job_th.company_id.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf job_th.company_id.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If job_th.job_date.Visible Then ' job_date %>
	<% If job_th.SortUrl(job_th.job_date) = "" Then %>
		<td><div id="elh_job_th_job_date" class="job_th_job_date"><div class="ewTableHeaderCaption"><%= job_th.job_date.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= job_th.SortUrl(job_th.job_date) %>',1);"><div id="elh_job_th_job_date" class="job_th_job_date">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= job_th.job_date.FldCaption %></span><span class="ewTableHeaderSort"><% If job_th.job_date.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf job_th.job_date.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If job_th.job_title.Visible Then ' job_title %>
	<% If job_th.SortUrl(job_th.job_title) = "" Then %>
		<td><div id="elh_job_th_job_title" class="job_th_job_title"><div class="ewTableHeaderCaption"><%= job_th.job_title.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= job_th.SortUrl(job_th.job_title) %>',1);"><div id="elh_job_th_job_title" class="job_th_job_title">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= job_th.job_title.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If job_th.job_title.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf job_th.job_title.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If job_th.job_create.Visible Then ' job_create %>
	<% If job_th.SortUrl(job_th.job_create) = "" Then %>
		<td><div id="elh_job_th_job_create" class="job_th_job_create"><div class="ewTableHeaderCaption"><%= job_th.job_create.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= job_th.SortUrl(job_th.job_create) %>',1);"><div id="elh_job_th_job_create" class="job_th_job_create">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= job_th.job_create.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If job_th.job_create.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf job_th.job_create.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If job_th.job_update.Visible Then ' job_update %>
	<% If job_th.SortUrl(job_th.job_update) = "" Then %>
		<td><div id="elh_job_th_job_update" class="job_th_job_update"><div class="ewTableHeaderCaption"><%= job_th.job_update.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= job_th.SortUrl(job_th.job_update) %>',1);"><div id="elh_job_th_job_update" class="job_th_job_update">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= job_th.job_update.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If job_th.job_update.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf job_th.job_update.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If job_th.job_show.Visible Then ' job_show %>
	<% If job_th.SortUrl(job_th.job_show) = "" Then %>
		<td><div id="elh_job_th_job_show" class="job_th_job_show"><div class="ewTableHeaderCaption"><%= job_th.job_show.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= job_th.SortUrl(job_th.job_show) %>',1);"><div id="elh_job_th_job_show" class="job_th_job_show">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= job_th.job_show.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If job_th.job_show.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf job_th.job_show.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
job_th_list.ListOptions.Render "header", "right", "", "", "", ""
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (job_th.ExportAll And job_th.Export <> "") Then
	job_th_list.StopRec = job_th_list.TotalRecs
Else

	' Set the last record to display
	If job_th_list.TotalRecs > job_th_list.StartRec + job_th_list.DisplayRecs - 1 Then
		job_th_list.StopRec = job_th_list.StartRec + job_th_list.DisplayRecs - 1
	Else
		job_th_list.StopRec = job_th_list.TotalRecs
	End If
End If

' Move to first record
job_th_list.RecCnt = job_th_list.StartRec - 1
If Not job_th_list.Recordset.Eof Then
	job_th_list.Recordset.MoveFirst
	If job_th_list.StartRec > 1 Then job_th_list.Recordset.Move job_th_list.StartRec - 1
ElseIf Not job_th.AllowAddDeleteRow And job_th_list.StopRec = 0 Then
	job_th_list.StopRec = job_th.GridAddRowCount
End If

' Initialize Aggregate
job_th.RowType = EW_ROWTYPE_AGGREGATEINIT
Call job_th.ResetAttrs()
Call job_th_list.RenderRow()
job_th_list.RowCnt = 0

' Output date rows
Do While CLng(job_th_list.RecCnt) < CLng(job_th_list.StopRec)
	job_th_list.RecCnt = job_th_list.RecCnt + 1
	If CLng(job_th_list.RecCnt) >= CLng(job_th_list.StartRec) Then
		job_th_list.RowCnt = job_th_list.RowCnt + 1

	' Set up key count
	job_th_list.KeyCount = job_th_list.RowIndex
	Call job_th.ResetAttrs()
	job_th.CssClass = ""
	If job_th.CurrentAction = "gridadd" Then
	Else
		Call job_th_list.LoadRowValues(job_th_list.Recordset) ' Load row values
	End If
	job_th.RowType = EW_ROWTYPE_VIEW ' Render view

	' Set up row id / data-rowindex
	job_th.RowAttrs.AddAttributes Array(Array("data-rowindex", job_th_list.RowCnt), Array("id", "r" & job_th_list.RowCnt & "_job_th"), Array("data-rowtype", job_th.RowType))

	' Render row
	Call job_th_list.RenderRow()

	' Render list options
	Call job_th_list.RenderListOptions()
%>
	<tr<%= job_th.RowAttributes %>>
<%

' Render list options (body, left)
job_th_list.ListOptions.Render "body", "left", job_th_list.RowCnt, "", "", ""
%>
	<% If job_th.job_id.Visible Then ' job_id %>
		<td<%= job_th.job_id.CellAttributes %>>
<span<%= job_th.job_id.ViewAttributes %>>
<%= job_th.job_id.ListViewValue %>
</span>
<a id="<%= job_th_list.PageObjName & "_row_" & job_th_list.RowCnt %>"></a></td>
	<% End If %>
	<% If job_th.company_id.Visible Then ' company_id %>
		<td<%= job_th.company_id.CellAttributes %>>
<span<%= job_th.company_id.ViewAttributes %>>
<%= job_th.company_id.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If job_th.job_date.Visible Then ' job_date %>
		<td<%= job_th.job_date.CellAttributes %>>
<span<%= job_th.job_date.ViewAttributes %>>
<%= job_th.job_date.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If job_th.job_title.Visible Then ' job_title %>
		<td<%= job_th.job_title.CellAttributes %>>
<span<%= job_th.job_title.ViewAttributes %>>
<%= job_th.job_title.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If job_th.job_create.Visible Then ' job_create %>
		<td<%= job_th.job_create.CellAttributes %>>
<span<%= job_th.job_create.ViewAttributes %>>
<%= job_th.job_create.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If job_th.job_update.Visible Then ' job_update %>
		<td<%= job_th.job_update.CellAttributes %>>
<span<%= job_th.job_update.ViewAttributes %>>
<%= job_th.job_update.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If job_th.job_show.Visible Then ' job_show %>
		<td<%= job_th.job_show.CellAttributes %>>
<span<%= job_th.job_show.ViewAttributes %>>
<%= job_th.job_show.ListViewValue %>
</span>
</td>
	<% End If %>
<%

' Render list options (body, right)
job_th_list.ListOptions.Render "body", "right", job_th_list.RowCnt, "", "", ""
%>
	</tr>
<%
	End If
	If job_th.CurrentAction <> "gridadd" Then
		job_th_list.Recordset.MoveNext()
	End If
Loop
%>
</tbody>
</table>
<% End If %>
<% If job_th.CurrentAction = "" Then %>
<input type="hidden" name="a_list" id="a_list" value="">
<% End If %>
</div>
</form>
<%

' Close recordset and connection
job_th_list.Recordset.Close
Set job_th_list.Recordset = Nothing
%>
<% If job_th.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If job_th.CurrentAction <> "gridadd" And job_th.CurrentAction <> "gridedit" Then %>
<form name="ewPagerForm" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewPager">
<tr><td>
<% If Not IsObject(job_th_list.Pager) Then Set job_th_list.Pager = ew_NewPrevNextPager(job_th_list.StartRec, job_th_list.DisplayRecs, job_th_list.TotalRecs) %>
<% If job_th_list.Pager.RecordCount > 0 Then %>
<table class="ewStdTable"><tbody><tr><td>
	<%= Language.Phrase("Page") %>&nbsp;
<div class="input-prepend input-append">
<!--first page button-->
	<% If job_th_list.Pager.FirstButton.Enabled Then %>
	<a class="btn btn-small" href="<%= job_th_list.PageUrl %>start=<%= job_th_list.Pager.FirstButton.Start %>"><i class="icon-step-backward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-backward"></i></a>
	<% End If %>
<!--previous page button-->
	<% If job_th_list.Pager.PrevButton.Enabled Then %>
	<a class="btn btn-small" href="<%= job_th_list.PageUrl %>start=<%= job_th_list.Pager.PrevButton.Start %>"><i class="icon-prev"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-prev"></i></a>
	<% End If %>
<!--current page number-->
	<input class="input-mini" type="text" name="<%= EW_TABLE_PAGE_NO %>" value="<%= job_th_list.Pager.CurrentPage %>">
<!--next page button-->
	<% If job_th_list.Pager.NextButton.Enabled Then %>
	<a class="btn btn-small" href="<%= job_th_list.PageUrl %>start=<%= job_th_list.Pager.NextButton.Start %>"><i class="icon-play"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-play"></i></a>
	<% End If %>
<!--last page button-->
	<% If job_th_list.Pager.LastButton.Enabled Then %>
	<a class="btn btn-small" href="<%= job_th_list.PageUrl %>start=<%= job_th_list.Pager.LastButton.Start %>"><i class="icon-step-forward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-forward"></i></a>
	<% End If %>
</div>
	&nbsp;<%= Language.Phrase("of") %>&nbsp;<%= job_th_list.Pager.PageCount %>
</td>
<td>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<%= Language.Phrase("Record") %>&nbsp;<%= job_th_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= job_th_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= job_th_list.Pager.RecordCount %>
</td>
</tr></tbody></table>
<% Else %>
	<% If job_th_list.SearchWhere = "0=101" Then %>
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
	job_th_list.AddEditOptions.Render "body", "bottom", "", "", "", ""
	job_th_list.DetailOptions.Render "body", "bottom", "", "", "", ""
	job_th_list.ActionOptions.Render "body", "bottom", "", "", "", ""
%>
</div>
</div>
<% End If %>
</td></tr></table>
<% If job_th.Export = "" Then %>
<script type="text/javascript">
fjob_thlistsrch.Init();
fjob_thlist.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<% End If %>
<%
job_th_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If job_th.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set job_th_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cjob_th_list

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
		TableName = "job_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "job_th_list"
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
		If job_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & job_th.TableVar & "&" ' add page token
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
		If job_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (job_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (job_th.TableVar = Request.QueryString("t"))
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
		FormName = "fjob_thlist"
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
		If IsEmpty(job_th) Then Set job_th = New cjob_th
		Set Table = job_th

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		ExportPdfUrl = PageUrl & "export=pdf"
		AddUrl = "pom_job_thadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "pom_job_thdelete.asp"
		MultiUpdateUrl = "pom_job_thupdate.asp"

		' Initialize other table object
		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "job_th"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' List options
		Set ListOptions = New cListOptions
		ListOptions.TableVar = job_th.TableVar

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = job_th.TableVar
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
				job_th.GridAddRowCount = gridaddcnt
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
		If UBound(job_th.CustomActions.CustomArray) >= 0 Then
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
		Set job_th = Nothing
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
			If job_th.Export = "" Then
				SetupBreadcrumb()
			End If

			' Hide list options
			If job_th.Export <> "" Then
				Call ListOptions.HideAllOptions(Array("sequence"))
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			ElseIf job_th.CurrentAction = "gridadd" Or job_th.CurrentAction = "gridedit" Then
				Call ListOptions.HideAllOptions(Array())
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			End If

			' Hide export options
			If job_th.Export <> "" Or job_th.CurrentAction <> "" Then
				ExportOptions.HideAllOptions(Array())
			End If

			' Hide other options
			If job_th.Export <> "" Then
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
			Call job_th.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If job_th.RecordsPerPage <> "" Then
			DisplayRecs = job_th.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Load search default if no existing search criteria
		If Not CheckSearchParms() Then

			' Load basic search from default
			job_th.BasicSearch.Keyword = job_th.BasicSearch.KeywordDefault
			job_th.BasicSearch.SearchType = job_th.BasicSearch.SearchTypeDefault
			job_th.BasicSearch.setSearchType(job_th.BasicSearch.SearchTypeDefault)
			If job_th.BasicSearch.Keyword <> "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call job_th.Recordset_Searching(SearchWhere)

		' Save search criteria
		If Command = "search" And Not RestoreSearch Then
			job_th.SearchWhere = SearchWhere ' Save to Session
			StartRec = 1 ' Reset start record counter
			job_th.StartRecordNumber = StartRec
		Else
			SearchWhere = job_th.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		job_th.SessionWhere = sFilter
		job_th.CurrentFilter = ""
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
				sFilter = job_th.KeyFilter
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
			job_th.job_id.FormValue = arrKeyFlds(0)
			If Not IsNumeric(job_th.job_id.FormValue) Then
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
			Call BuildBasicSearchSQL(sWhere, job_th.job_title, Keyword)
			Call BuildBasicSearchSQL(sWhere, job_th.job_intro, Keyword)
			Call BuildBasicSearchSQL(sWhere, job_th.job_detail, Keyword)
			Call BuildBasicSearchSQL(sWhere, job_th.job_create, Keyword)
			Call BuildBasicSearchSQL(sWhere, job_th.job_update, Keyword)
			Call BuildBasicSearchSQL(sWhere, job_th.job_show, Keyword)
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
		sSearchKeyword = job_th.BasicSearch.Keyword
		sSearchType = job_th.BasicSearch.SearchType
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
			job_th.BasicSearch.setKeyword(sSearchKeyword)
			job_th.BasicSearch.setSearchType(sSearchType)
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' Check if search parm exists
	Function CheckSearchParms()

		' Check basic search
		If job_th.BasicSearch.IssetSession() Then
			CheckSearchParms = True
			Exit Function
		End If
		CheckSearchParms = False
	End Function

	' Clear all search parameters
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		job_th.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' Load advanced search default values
	Function LoadAdvancedSearchDefault()
		LoadAdvancedSearchDefault = False
	End Function

	' Clear all basic search parameters
	Sub ResetBasicSearchParms()
		job_th.BasicSearch.UnsetSession()
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()

		' Restore search flag
		RestoreSearch = True

		' Restore basic search values
		Call job_th.BasicSearch.Load()
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
			job_th.CurrentOrder = Request.QueryString("order")
			job_th.CurrentOrderType = Request.QueryString("ordertype")

			' Field job_id
			Call job_th.UpdateSort(job_th.job_id)

			' Field company_id
			Call job_th.UpdateSort(job_th.company_id)

			' Field job_date
			Call job_th.UpdateSort(job_th.job_date)

			' Field job_title
			Call job_th.UpdateSort(job_th.job_title)

			' Field job_create
			Call job_th.UpdateSort(job_th.job_create)

			' Field job_update
			Call job_th.UpdateSort(job_th.job_update)

			' Field job_show
			Call job_th.UpdateSort(job_th.job_show)
			job_th.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = job_th.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If job_th.SqlOrderBy <> "" Then
				sOrderBy = job_th.SqlOrderBy
				job_th.SessionOrderBy = sOrderBy
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
				job_th.SessionOrderBy = sOrderBy
				job_th.job_id.Sort = ""
				job_th.company_id.Sort = ""
				job_th.job_date.Sort = ""
				job_th.job_title.Sort = ""
				job_th.job_create.Sort = ""
				job_th.job_update.Sort = ""
				job_th.job_show.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			job_th.StartRecordNumber = StartRec
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
		ListOptions.GetItem("checkbox").Body = "<label class=""checkbox""><input type=""checkbox"" name=""key_m"" value=""" & ew_HtmlEncode(job_th.job_id.CurrentValue) & """ onclick='ew_ClickMultiCheckbox(event, this);'></label>"
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
			For i = 0 to UBound(job_th.CustomActions.CustomArray)
				Action = job_th.CustomActions.CustomArray(i)(0)
				Name = job_th.CustomActions.CustomArray(i)(1)

				' Add custom action
				Call opt.Add("custom_" & Action)
				Set item = opt.GetItem("custom_" & Action)
				item.Body = "<a class=""ewAction ewCustomAction"" href="""" onclick=""ew_SubmitSelected(document.fjob_thlist, '" & ew_CurrentUrl & "', null, '" & Action & "');return false;"">" & Name & "</a>"
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
		sFilter = job_th.GetKeyFilter
		UserAction = Request.Form("useraction") & ""
		Processed = False
		If sFilter <> "" And UserAction <> "" Then
			job_th.CurrentFilter = sFilter
			sSql = job_th.SQL
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
				ElseIf job_th.CancelMessage <> "" Then
					FailureMessage = job_th.CancelMessage
					job_th.CancelMessage = ""
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
				job_th.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					job_th.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = job_th.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			job_th.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			job_th.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			job_th.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		job_th.BasicSearch.Keyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)&""
		If job_th.BasicSearch.Keyword <> "" Then Command = "search"
		job_th.BasicSearch.SearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)&""
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = job_th.CurrentFilter
		Call job_th.Recordset_Selecting(sFilter)
		job_th.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = job_th.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call job_th.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = job_th.KeyFilter

		' Call Row Selecting event
		Call job_th.Row_Selecting(sFilter)

		' Load sql based on filter
		job_th.CurrentFilter = sFilter
		sSql = job_th.SQL
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
		Call job_th.Row_Selected(RsRow)
		job_th.job_id.DbValue = RsRow("job_id")
		job_th.company_id.DbValue = RsRow("company_id")
		job_th.job_date.DbValue = RsRow("job_date")
		job_th.job_title.DbValue = RsRow("job_title")
		job_th.job_intro.DbValue = RsRow("job_intro")
		job_th.job_detail.DbValue = RsRow("job_detail")
		job_th.job_create.DbValue = RsRow("job_create")
		job_th.job_update.DbValue = RsRow("job_update")
		job_th.job_show.DbValue = RsRow("job_show")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		job_th.job_id.m_DbValue = Rs("job_id")
		job_th.company_id.m_DbValue = Rs("company_id")
		job_th.job_date.m_DbValue = Rs("job_date")
		job_th.job_title.m_DbValue = Rs("job_title")
		job_th.job_intro.m_DbValue = Rs("job_intro")
		job_th.job_detail.m_DbValue = Rs("job_detail")
		job_th.job_create.m_DbValue = Rs("job_create")
		job_th.job_update.m_DbValue = Rs("job_update")
		job_th.job_show.m_DbValue = Rs("job_show")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If job_th.GetKey("job_id")&"" <> "" Then
			job_th.job_id.CurrentValue = job_th.GetKey("job_id") ' job_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			job_th.CurrentFilter = job_th.KeyFilter
			Dim sSql
			sSql = job_th.SQL
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
		ViewUrl = job_th.ViewUrl("")
		EditUrl = job_th.EditUrl("")
		InlineEditUrl = job_th.InlineEditUrl
		CopyUrl = job_th.CopyUrl("")
		InlineCopyUrl = job_th.InlineCopyUrl
		DeleteUrl = job_th.DeleteUrl

		' Call Row Rendering event
		Call job_th.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' job_id
		' company_id
		' job_date
		' job_title
		' job_intro
		' job_detail
		' job_create
		' job_update
		' job_show
		' -----------
		'  View  Row
		' -----------

		If job_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' job_id
			job_th.job_id.ViewValue = job_th.job_id.CurrentValue
			job_th.job_id.ViewCustomAttributes = ""

			' company_id
			job_th.company_id.ViewValue = job_th.company_id.CurrentValue
			job_th.company_id.ViewCustomAttributes = ""

			' job_date
			job_th.job_date.ViewValue = job_th.job_date.CurrentValue
			job_th.job_date.ViewCustomAttributes = ""

			' job_title
			job_th.job_title.ViewValue = job_th.job_title.CurrentValue
			job_th.job_title.ViewCustomAttributes = ""

			' job_create
			job_th.job_create.ViewValue = job_th.job_create.CurrentValue
			job_th.job_create.ViewCustomAttributes = ""

			' job_update
			job_th.job_update.ViewValue = job_th.job_update.CurrentValue
			job_th.job_update.ViewCustomAttributes = ""

			' job_show
			job_th.job_show.ViewValue = job_th.job_show.CurrentValue
			job_th.job_show.ViewCustomAttributes = ""

			' View refer script
			' job_id

			job_th.job_id.LinkCustomAttributes = ""
			job_th.job_id.HrefValue = ""
			job_th.job_id.TooltipValue = ""

			' company_id
			job_th.company_id.LinkCustomAttributes = ""
			job_th.company_id.HrefValue = ""
			job_th.company_id.TooltipValue = ""

			' job_date
			job_th.job_date.LinkCustomAttributes = ""
			job_th.job_date.HrefValue = ""
			job_th.job_date.TooltipValue = ""

			' job_title
			job_th.job_title.LinkCustomAttributes = ""
			job_th.job_title.HrefValue = ""
			job_th.job_title.TooltipValue = ""

			' job_create
			job_th.job_create.LinkCustomAttributes = ""
			job_th.job_create.HrefValue = ""
			job_th.job_create.TooltipValue = ""

			' job_update
			job_th.job_update.LinkCustomAttributes = ""
			job_th.job_update.HrefValue = ""
			job_th.job_update.TooltipValue = ""

			' job_show
			job_th.job_show.LinkCustomAttributes = ""
			job_th.job_show.HrefValue = ""
			job_th.job_show.TooltipValue = ""
		End If

		' Call Row Rendered event
		If job_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call job_th.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		url = ew_CurrentUrl
		url = ew_RegExReplace("\?cmd=reset(all){0,1}$", url, "") ' Remove cmd=reset / cmd=resetall
		Call Breadcrumb.Add("list", job_th.TableVar, url, job_th.TableVar, True)
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
