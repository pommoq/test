<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_person_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim person_th_list
Set person_th_list = New cperson_th_list
Set Page = person_th_list

' Page init processing
person_th_list.Page_Init()

' Page main processing
person_th_list.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
person_th_list.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<% If person_th.Export = "" Then %>
<script type="text/javascript">
// Page object
var person_th_list = new ew_Page("person_th_list");
person_th_list.PageID = "list"; // Page ID
var EW_PAGE_ID = person_th_list.PageID; // For backward compatibility
// Form object
var fperson_thlist = new ew_Form("fperson_thlist");
fperson_thlist.FormKeyCountName = '<%= person_th_list.FormKeyCountName %>';
// Form_CustomValidate event
fperson_thlist.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fperson_thlist.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fperson_thlist.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
var fperson_thlistsrch = new ew_Form("fperson_thlistsrch");
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% End If %>
<% If person_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% If person_th_list.ExportOptions.Visible Then %>
<div class="ewListExportOptions"><% person_th_list.ExportOptions.Render "body", "", "", "", "", "" %></div>
<% End If %>
<% If (person_th.Export = "") Or (EW_EXPORT_MASTER_RECORD And person_th.Export = "print") Then %>
<% End If %>
<%

' Load recordset
Set person_th_list.Recordset = person_th_list.LoadRecordset()
	person_th_list.TotalRecs = person_th_list.Recordset.RecordCount
	person_th_list.StartRec = 1
	If person_th_list.DisplayRecs <= 0 Then ' Display all records
		person_th_list.DisplayRecs = person_th_list.TotalRecs
	End If
	If Not (person_th.ExportAll And person_th.Export <> "") Then
		person_th_list.SetUpStartRec() ' Set up start record position
	End If
person_th_list.RenderOtherOptions()
%>
<% If Security.IsLoggedIn() Then %>
<% If person_th.Export = "" And person_th.CurrentAction = "" Then %>
<form name="fperson_thlistsrch" id="fperson_thlistsrch" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewSearchTable"><tr><td>
<div class="accordion" id="fperson_thlistsrch_SearchGroup">
	<div class="accordion-group">
		<div class="accordion-heading">
<a class="accordion-toggle" data-toggle="collapse" data-parent="#fperson_thlistsrch_SearchGroup" href="#fperson_thlistsrch_SearchBody"><%= Language.Phrase("Search") %></a>
		</div>
		<div id="fperson_thlistsrch_SearchBody" class="accordion-body collapse in">
			<div class="accordion-inner">
<div id="fperson_thlistsrch_SearchPanel">
<input type="hidden" name="cmd" value="search">
<input type="hidden" name="t" value="person_th">
<div class="ewBasicSearch">
<div id="xsr_1" class="ewRow">
	<div class="btn-group ewButtonGroup">
	<div class="input-append">
	<input type="text" name="<%= EW_TABLE_BASIC_SEARCH %>" id="<%= EW_TABLE_BASIC_SEARCH %>" class="input-large" value="<%= ew_HtmlEncode(person_th.BasicSearch.getKeyword()) %>" placeholder="<%= ew_HtmlEncode(Language.Phrase("Search")) %>">
	<button class="btn btn-primary ewButton" name="btnsubmit" id="btnsubmit" type="submit"><%= Language.Phrase("QuickSearchBtn") %></button>
	</div>
	</div>
	<div class="btn-group ewButtonGroup">
	<a class="btn ewShowAll" href="<%= person_th_list.PageUrl %>cmd=reset"><%= Language.Phrase("ShowAll") %></a>
	</div>
</div>
<div id="xsr_2" class="ewRow">
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="="<% If person_th.BasicSearch.getSearchType = "=" Then %> checked="checked"<% End If %>><%= Language.Phrase("ExactPhrase") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="AND"<% If person_th.BasicSearch.getSearchType = "AND" Then %> checked="checked"<% End If %>><%= Language.Phrase("AllWord") %></label>
	<label class="inline radio ewRadio" style="white-space: nowrap;"><input type="radio" name="<%= EW_TABLE_BASIC_SEARCH_TYPE %>" value="OR"<% If person_th.BasicSearch.getSearchType = "OR" Then %> checked="checked"<% End If %>><%= Language.Phrase("AnyWord") %></label>
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
<% person_th_list.ShowPageHeader() %>
<% person_th_list.ShowMessage %>
<table class="ewGrid"><tr><td class="ewGridContent">
<form name="fperson_thlist" id="fperson_thlist" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="t" value="person_th">
<div id="gmp_person_th" class="ewGridMiddlePanel">
<% If person_th_list.TotalRecs > 0 Then %>
<table id="tbl_person_thlist" class="ewTable ewTableSeparate">
<%= person_th.TableCustomInnerHtml %>
<thead><!-- Table header -->
	<tr class="ewTableHeader">
<%
Call person_th_list.RenderListOptions()

' Render list options (header, left)
person_th_list.ListOptions.Render "header", "left", "", "", "", ""
%>
<% If person_th.per_id.Visible Then ' per_id %>
	<% If person_th.SortUrl(person_th.per_id) = "" Then %>
		<td><div id="elh_person_th_per_id" class="person_th_per_id"><div class="ewTableHeaderCaption"><%= person_th.per_id.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.per_id) %>',1);"><div id="elh_person_th_per_id" class="person_th_per_id">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.per_id.FldCaption %></span><span class="ewTableHeaderSort"><% If person_th.per_id.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.per_id.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If person_th.dept_id.Visible Then ' dept_id %>
	<% If person_th.SortUrl(person_th.dept_id) = "" Then %>
		<td><div id="elh_person_th_dept_id" class="person_th_dept_id"><div class="ewTableHeaderCaption"><%= person_th.dept_id.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.dept_id) %>',1);"><div id="elh_person_th_dept_id" class="person_th_dept_id">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.dept_id.FldCaption %></span><span class="ewTableHeaderSort"><% If person_th.dept_id.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.dept_id.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If person_th.office_id.Visible Then ' office_id %>
	<% If person_th.SortUrl(person_th.office_id) = "" Then %>
		<td><div id="elh_person_th_office_id" class="person_th_office_id"><div class="ewTableHeaderCaption"><%= person_th.office_id.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.office_id) %>',1);"><div id="elh_person_th_office_id" class="person_th_office_id">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.office_id.FldCaption %></span><span class="ewTableHeaderSort"><% If person_th.office_id.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.office_id.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If person_th.per_img.Visible Then ' per_img %>
	<% If person_th.SortUrl(person_th.per_img) = "" Then %>
		<td><div id="elh_person_th_per_img" class="person_th_per_img"><div class="ewTableHeaderCaption"><%= person_th.per_img.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.per_img) %>',1);"><div id="elh_person_th_per_img" class="person_th_per_img">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.per_img.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If person_th.per_img.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.per_img.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If person_th.per_en_name.Visible Then ' per_en_name %>
	<% If person_th.SortUrl(person_th.per_en_name) = "" Then %>
		<td><div id="elh_person_th_per_en_name" class="person_th_per_en_name"><div class="ewTableHeaderCaption"><%= person_th.per_en_name.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.per_en_name) %>',1);"><div id="elh_person_th_per_en_name" class="person_th_per_en_name">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.per_en_name.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If person_th.per_en_name.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.per_en_name.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If person_th.per_th_name.Visible Then ' per_th_name %>
	<% If person_th.SortUrl(person_th.per_th_name) = "" Then %>
		<td><div id="elh_person_th_per_th_name" class="person_th_per_th_name"><div class="ewTableHeaderCaption"><%= person_th.per_th_name.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.per_th_name) %>',1);"><div id="elh_person_th_per_th_name" class="person_th_per_th_name">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.per_th_name.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If person_th.per_th_name.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.per_th_name.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If person_th.per_position.Visible Then ' per_position %>
	<% If person_th.SortUrl(person_th.per_position) = "" Then %>
		<td><div id="elh_person_th_per_position" class="person_th_per_position"><div class="ewTableHeaderCaption"><%= person_th.per_position.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.per_position) %>',1);"><div id="elh_person_th_per_position" class="person_th_per_position">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.per_position.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If person_th.per_position.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.per_position.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If person_th.per_mobile.Visible Then ' per_mobile %>
	<% If person_th.SortUrl(person_th.per_mobile) = "" Then %>
		<td><div id="elh_person_th_per_mobile" class="person_th_per_mobile"><div class="ewTableHeaderCaption"><%= person_th.per_mobile.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.per_mobile) %>',1);"><div id="elh_person_th_per_mobile" class="person_th_per_mobile">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.per_mobile.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If person_th.per_mobile.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.per_mobile.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If person_th.per_tel.Visible Then ' per_tel %>
	<% If person_th.SortUrl(person_th.per_tel) = "" Then %>
		<td><div id="elh_person_th_per_tel" class="person_th_per_tel"><div class="ewTableHeaderCaption"><%= person_th.per_tel.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.per_tel) %>',1);"><div id="elh_person_th_per_tel" class="person_th_per_tel">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.per_tel.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If person_th.per_tel.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.per_tel.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If person_th.per_fax.Visible Then ' per_fax %>
	<% If person_th.SortUrl(person_th.per_fax) = "" Then %>
		<td><div id="elh_person_th_per_fax" class="person_th_per_fax"><div class="ewTableHeaderCaption"><%= person_th.per_fax.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.per_fax) %>',1);"><div id="elh_person_th_per_fax" class="person_th_per_fax">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.per_fax.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If person_th.per_fax.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.per_fax.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If person_th.per_email.Visible Then ' per_email %>
	<% If person_th.SortUrl(person_th.per_email) = "" Then %>
		<td><div id="elh_person_th_per_email" class="person_th_per_email"><div class="ewTableHeaderCaption"><%= person_th.per_email.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.per_email) %>',1);"><div id="elh_person_th_per_email" class="person_th_per_email">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.per_email.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If person_th.per_email.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.per_email.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If person_th.per_address.Visible Then ' per_address %>
	<% If person_th.SortUrl(person_th.per_address) = "" Then %>
		<td><div id="elh_person_th_per_address" class="person_th_per_address"><div class="ewTableHeaderCaption"><%= person_th.per_address.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.per_address) %>',1);"><div id="elh_person_th_per_address" class="person_th_per_address">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.per_address.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If person_th.per_address.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.per_address.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If person_th.per_show.Visible Then ' per_show %>
	<% If person_th.SortUrl(person_th.per_show) = "" Then %>
		<td><div id="elh_person_th_per_show" class="person_th_per_show"><div class="ewTableHeaderCaption"><%= person_th.per_show.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.per_show) %>',1);"><div id="elh_person_th_per_show" class="person_th_per_show">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.per_show.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If person_th.per_show.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.per_show.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If person_th.per_create.Visible Then ' per_create %>
	<% If person_th.SortUrl(person_th.per_create) = "" Then %>
		<td><div id="elh_person_th_per_create" class="person_th_per_create"><div class="ewTableHeaderCaption"><%= person_th.per_create.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.per_create) %>',1);"><div id="elh_person_th_per_create" class="person_th_per_create">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.per_create.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If person_th.per_create.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.per_create.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If person_th.per_update.Visible Then ' per_update %>
	<% If person_th.SortUrl(person_th.per_update) = "" Then %>
		<td><div id="elh_person_th_per_update" class="person_th_per_update"><div class="ewTableHeaderCaption"><%= person_th.per_update.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.per_update) %>',1);"><div id="elh_person_th_per_update" class="person_th_per_update">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.per_update.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If person_th.per_update.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.per_update.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If person_th.per_sort.Visible Then ' per_sort %>
	<% If person_th.SortUrl(person_th.per_sort) = "" Then %>
		<td><div id="elh_person_th_per_sort" class="person_th_per_sort"><div class="ewTableHeaderCaption"><%= person_th.per_sort.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.per_sort) %>',1);"><div id="elh_person_th_per_sort" class="person_th_per_sort">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.per_sort.FldCaption %></span><span class="ewTableHeaderSort"><% If person_th.per_sort.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.per_sort.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<% If person_th.per_department.Visible Then ' per_department %>
	<% If person_th.SortUrl(person_th.per_department) = "" Then %>
		<td><div id="elh_person_th_per_department" class="person_th_per_department"><div class="ewTableHeaderCaption"><%= person_th.per_department.FldCaption %></div></div></td>
	<% Else %>
		<td><div class="ewPointer" onclick="ew_Sort(event,'<%= person_th.SortUrl(person_th.per_department) %>',1);"><div id="elh_person_th_per_department" class="person_th_per_department">
			<div class="ewTableHeaderBtn"><span class="ewTableHeaderCaption"><%= person_th.per_department.FldCaption %><%= Language.Phrase("SrchLegend") %></span><span class="ewTableHeaderSort"><% If person_th.per_department.Sort = "ASC" Then %><span class="caret ewSortUp"></span><% ElseIf person_th.per_department.Sort = "DESC" Then %><span class="caret"></span><% End If %></span></div>
        </div></div></td>
	<% End If %>
<% End If %>		
<%

' Render list options (header, right)
person_th_list.ListOptions.Render "header", "right", "", "", "", ""
%>
	</tr>
</thead>
<tbody><!-- Table body -->
<%
If (person_th.ExportAll And person_th.Export <> "") Then
	person_th_list.StopRec = person_th_list.TotalRecs
Else

	' Set the last record to display
	If person_th_list.TotalRecs > person_th_list.StartRec + person_th_list.DisplayRecs - 1 Then
		person_th_list.StopRec = person_th_list.StartRec + person_th_list.DisplayRecs - 1
	Else
		person_th_list.StopRec = person_th_list.TotalRecs
	End If
End If

' Move to first record
person_th_list.RecCnt = person_th_list.StartRec - 1
If Not person_th_list.Recordset.Eof Then
	person_th_list.Recordset.MoveFirst
	If person_th_list.StartRec > 1 Then person_th_list.Recordset.Move person_th_list.StartRec - 1
ElseIf Not person_th.AllowAddDeleteRow And person_th_list.StopRec = 0 Then
	person_th_list.StopRec = person_th.GridAddRowCount
End If

' Initialize Aggregate
person_th.RowType = EW_ROWTYPE_AGGREGATEINIT
Call person_th.ResetAttrs()
Call person_th_list.RenderRow()
person_th_list.RowCnt = 0

' Output date rows
Do While CLng(person_th_list.RecCnt) < CLng(person_th_list.StopRec)
	person_th_list.RecCnt = person_th_list.RecCnt + 1
	If CLng(person_th_list.RecCnt) >= CLng(person_th_list.StartRec) Then
		person_th_list.RowCnt = person_th_list.RowCnt + 1

	' Set up key count
	person_th_list.KeyCount = person_th_list.RowIndex
	Call person_th.ResetAttrs()
	person_th.CssClass = ""
	If person_th.CurrentAction = "gridadd" Then
	Else
		Call person_th_list.LoadRowValues(person_th_list.Recordset) ' Load row values
	End If
	person_th.RowType = EW_ROWTYPE_VIEW ' Render view

	' Set up row id / data-rowindex
	person_th.RowAttrs.AddAttributes Array(Array("data-rowindex", person_th_list.RowCnt), Array("id", "r" & person_th_list.RowCnt & "_person_th"), Array("data-rowtype", person_th.RowType))

	' Render row
	Call person_th_list.RenderRow()

	' Render list options
	Call person_th_list.RenderListOptions()
%>
	<tr<%= person_th.RowAttributes %>>
<%

' Render list options (body, left)
person_th_list.ListOptions.Render "body", "left", person_th_list.RowCnt, "", "", ""
%>
	<% If person_th.per_id.Visible Then ' per_id %>
		<td<%= person_th.per_id.CellAttributes %>>
<span<%= person_th.per_id.ViewAttributes %>>
<%= person_th.per_id.ListViewValue %>
</span>
<a id="<%= person_th_list.PageObjName & "_row_" & person_th_list.RowCnt %>"></a></td>
	<% End If %>
	<% If person_th.dept_id.Visible Then ' dept_id %>
		<td<%= person_th.dept_id.CellAttributes %>>
<span<%= person_th.dept_id.ViewAttributes %>>
<%= person_th.dept_id.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If person_th.office_id.Visible Then ' office_id %>
		<td<%= person_th.office_id.CellAttributes %>>
<span<%= person_th.office_id.ViewAttributes %>>
<%= person_th.office_id.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If person_th.per_img.Visible Then ' per_img %>
		<td<%= person_th.per_img.CellAttributes %>>
<span<%= person_th.per_img.ViewAttributes %>>
<%= person_th.per_img.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If person_th.per_en_name.Visible Then ' per_en_name %>
		<td<%= person_th.per_en_name.CellAttributes %>>
<span<%= person_th.per_en_name.ViewAttributes %>>
<%= person_th.per_en_name.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If person_th.per_th_name.Visible Then ' per_th_name %>
		<td<%= person_th.per_th_name.CellAttributes %>>
<span<%= person_th.per_th_name.ViewAttributes %>>
<%= person_th.per_th_name.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If person_th.per_position.Visible Then ' per_position %>
		<td<%= person_th.per_position.CellAttributes %>>
<span<%= person_th.per_position.ViewAttributes %>>
<%= person_th.per_position.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If person_th.per_mobile.Visible Then ' per_mobile %>
		<td<%= person_th.per_mobile.CellAttributes %>>
<span<%= person_th.per_mobile.ViewAttributes %>>
<%= person_th.per_mobile.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If person_th.per_tel.Visible Then ' per_tel %>
		<td<%= person_th.per_tel.CellAttributes %>>
<span<%= person_th.per_tel.ViewAttributes %>>
<%= person_th.per_tel.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If person_th.per_fax.Visible Then ' per_fax %>
		<td<%= person_th.per_fax.CellAttributes %>>
<span<%= person_th.per_fax.ViewAttributes %>>
<%= person_th.per_fax.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If person_th.per_email.Visible Then ' per_email %>
		<td<%= person_th.per_email.CellAttributes %>>
<span<%= person_th.per_email.ViewAttributes %>>
<%= person_th.per_email.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If person_th.per_address.Visible Then ' per_address %>
		<td<%= person_th.per_address.CellAttributes %>>
<span<%= person_th.per_address.ViewAttributes %>>
<%= person_th.per_address.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If person_th.per_show.Visible Then ' per_show %>
		<td<%= person_th.per_show.CellAttributes %>>
<span<%= person_th.per_show.ViewAttributes %>>
<%= person_th.per_show.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If person_th.per_create.Visible Then ' per_create %>
		<td<%= person_th.per_create.CellAttributes %>>
<span<%= person_th.per_create.ViewAttributes %>>
<%= person_th.per_create.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If person_th.per_update.Visible Then ' per_update %>
		<td<%= person_th.per_update.CellAttributes %>>
<span<%= person_th.per_update.ViewAttributes %>>
<%= person_th.per_update.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If person_th.per_sort.Visible Then ' per_sort %>
		<td<%= person_th.per_sort.CellAttributes %>>
<span<%= person_th.per_sort.ViewAttributes %>>
<%= person_th.per_sort.ListViewValue %>
</span>
</td>
	<% End If %>
	<% If person_th.per_department.Visible Then ' per_department %>
		<td<%= person_th.per_department.CellAttributes %>>
<span<%= person_th.per_department.ViewAttributes %>>
<%= person_th.per_department.ListViewValue %>
</span>
</td>
	<% End If %>
<%

' Render list options (body, right)
person_th_list.ListOptions.Render "body", "right", person_th_list.RowCnt, "", "", ""
%>
	</tr>
<%
	End If
	If person_th.CurrentAction <> "gridadd" Then
		person_th_list.Recordset.MoveNext()
	End If
Loop
%>
</tbody>
</table>
<% End If %>
<% If person_th.CurrentAction = "" Then %>
<input type="hidden" name="a_list" id="a_list" value="">
<% End If %>
</div>
</form>
<%

' Close recordset and connection
person_th_list.Recordset.Close
Set person_th_list.Recordset = Nothing
%>
<% If person_th.Export = "" Then %>
<div class="ewGridLowerPanel">
<% If person_th.CurrentAction <> "gridadd" And person_th.CurrentAction <> "gridedit" Then %>
<form name="ewPagerForm" class="ewForm form-inline" action="<%= ew_CurrentPage %>">
<table class="ewPager">
<tr><td>
<% If Not IsObject(person_th_list.Pager) Then Set person_th_list.Pager = ew_NewPrevNextPager(person_th_list.StartRec, person_th_list.DisplayRecs, person_th_list.TotalRecs) %>
<% If person_th_list.Pager.RecordCount > 0 Then %>
<table class="ewStdTable"><tbody><tr><td>
	<%= Language.Phrase("Page") %>&nbsp;
<div class="input-prepend input-append">
<!--first page button-->
	<% If person_th_list.Pager.FirstButton.Enabled Then %>
	<a class="btn btn-small" href="<%= person_th_list.PageUrl %>start=<%= person_th_list.Pager.FirstButton.Start %>"><i class="icon-step-backward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-backward"></i></a>
	<% End If %>
<!--previous page button-->
	<% If person_th_list.Pager.PrevButton.Enabled Then %>
	<a class="btn btn-small" href="<%= person_th_list.PageUrl %>start=<%= person_th_list.Pager.PrevButton.Start %>"><i class="icon-prev"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-prev"></i></a>
	<% End If %>
<!--current page number-->
	<input class="input-mini" type="text" name="<%= EW_TABLE_PAGE_NO %>" value="<%= person_th_list.Pager.CurrentPage %>">
<!--next page button-->
	<% If person_th_list.Pager.NextButton.Enabled Then %>
	<a class="btn btn-small" href="<%= person_th_list.PageUrl %>start=<%= person_th_list.Pager.NextButton.Start %>"><i class="icon-play"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-play"></i></a>
	<% End If %>
<!--last page button-->
	<% If person_th_list.Pager.LastButton.Enabled Then %>
	<a class="btn btn-small" href="<%= person_th_list.PageUrl %>start=<%= person_th_list.Pager.LastButton.Start %>"><i class="icon-step-forward"></i></a>
	<% Else %>
	<a class="btn btn-small disabled"><i class="icon-step-forward"></i></a>
	<% End If %>
</div>
	&nbsp;<%= Language.Phrase("of") %>&nbsp;<%= person_th_list.Pager.PageCount %>
</td>
<td>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<%= Language.Phrase("Record") %>&nbsp;<%= person_th_list.Pager.FromIndex %>&nbsp;<%= Language.Phrase("To") %>&nbsp;<%= person_th_list.Pager.ToIndex %>&nbsp;<%= Language.Phrase("Of") %>&nbsp;<%= person_th_list.Pager.RecordCount %>
</td>
</tr></tbody></table>
<% Else %>
	<% If person_th_list.SearchWhere = "0=101" Then %>
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
	person_th_list.AddEditOptions.Render "body", "bottom", "", "", "", ""
	person_th_list.DetailOptions.Render "body", "bottom", "", "", "", ""
	person_th_list.ActionOptions.Render "body", "bottom", "", "", "", ""
%>
</div>
</div>
<% End If %>
</td></tr></table>
<% If person_th.Export = "" Then %>
<script type="text/javascript">
fperson_thlistsrch.Init();
fperson_thlist.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<% End If %>
<%
person_th_list.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<% If person_th.Export = "" Then %>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<% End If %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set person_th_list = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cperson_th_list

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
		TableName = "person_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "person_th_list"
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
		If person_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & person_th.TableVar & "&" ' add page token
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
		If person_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (person_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (person_th.TableVar = Request.QueryString("t"))
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
		FormName = "fperson_thlist"
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
		If IsEmpty(person_th) Then Set person_th = New cperson_th
		Set Table = person_th

		' Initialize urls
		ExportPrintUrl = PageUrl & "export=print"
		ExportExcelUrl = PageUrl & "export=excel"
		ExportWordUrl = PageUrl & "export=word"
		ExportHtmlUrl = PageUrl & "export=html"
		ExportXmlUrl = PageUrl & "export=xml"
		ExportCsvUrl = PageUrl & "export=csv"
		ExportPdfUrl = PageUrl & "export=pdf"
		AddUrl = "pom_person_thadd.asp"
		InlineAddUrl = PageUrl & "a=add"
		GridAddUrl = PageUrl & "a=gridadd"
		GridEditUrl = PageUrl & "a=gridedit"
		MultiDeleteUrl = "pom_person_thdelete.asp"
		MultiUpdateUrl = "pom_person_thupdate.asp"

		' Initialize other table object
		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "list"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "person_th"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()

		' List options
		Set ListOptions = New cListOptions
		ListOptions.TableVar = person_th.TableVar

		' Export options
		Set ExportOptions = New cListOptions
		ExportOptions.TableVar = person_th.TableVar
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
				person_th.GridAddRowCount = gridaddcnt
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
		If UBound(person_th.CustomActions.CustomArray) >= 0 Then
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
		Set person_th = Nothing
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
			If person_th.Export = "" Then
				SetupBreadcrumb()
			End If

			' Hide list options
			If person_th.Export <> "" Then
				Call ListOptions.HideAllOptions(Array("sequence"))
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			ElseIf person_th.CurrentAction = "gridadd" Or person_th.CurrentAction = "gridedit" Then
				Call ListOptions.HideAllOptions(Array())
				ListOptions.UseDropDownButton = False ' Disable drop down button
				ListOptions.UseButtonGroup = False ' Disable button group
			End If

			' Hide export options
			If person_th.Export <> "" Or person_th.CurrentAction <> "" Then
				ExportOptions.HideAllOptions(Array())
			End If

			' Hide other options
			If person_th.Export <> "" Then
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
			Call person_th.Recordset_SearchValidated()

			' Set Up Sorting Order
			SetUpSortOrder()

			' Get basic search criteria
			If gsSearchError = "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If ' End Validate Request

		' Restore display records
		If person_th.RecordsPerPage <> "" Then
			DisplayRecs = person_th.RecordsPerPage ' Restore from Session
		Else
			DisplayRecs = 50 ' Load default
		End If

		' Load Sorting Order
		LoadSortOrder()

		' Load search default if no existing search criteria
		If Not CheckSearchParms() Then

			' Load basic search from default
			person_th.BasicSearch.Keyword = person_th.BasicSearch.KeywordDefault
			person_th.BasicSearch.SearchType = person_th.BasicSearch.SearchTypeDefault
			person_th.BasicSearch.setSearchType(person_th.BasicSearch.SearchTypeDefault)
			If person_th.BasicSearch.Keyword <> "" Then
				sSrchBasic = BasicSearchWhere()
			End If
		End If

		' Build search criteria
		Call ew_AddFilter(SearchWhere, sSrchAdvanced)
		Call ew_AddFilter(SearchWhere, sSrchBasic)

		' Call Recordset Searching event
		Call person_th.Recordset_Searching(SearchWhere)

		' Save search criteria
		If Command = "search" And Not RestoreSearch Then
			person_th.SearchWhere = SearchWhere ' Save to Session
			StartRec = 1 ' Reset start record counter
			person_th.StartRecordNumber = StartRec
		Else
			SearchWhere = person_th.SearchWhere
		End If
		sFilter = ""
		Call ew_AddFilter(sFilter, DbDetailFilter)
		Call ew_AddFilter(sFilter, SearchWhere)

		' Set up filter in Session
		person_th.SessionWhere = sFilter
		person_th.CurrentFilter = ""
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
				sFilter = person_th.KeyFilter
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
			person_th.per_id.FormValue = arrKeyFlds(0)
			If Not IsNumeric(person_th.per_id.FormValue) Then
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
			Call BuildBasicSearchSQL(sWhere, person_th.per_img, Keyword)
			Call BuildBasicSearchSQL(sWhere, person_th.per_en_name, Keyword)
			Call BuildBasicSearchSQL(sWhere, person_th.per_th_name, Keyword)
			Call BuildBasicSearchSQL(sWhere, person_th.per_position, Keyword)
			Call BuildBasicSearchSQL(sWhere, person_th.per_mobile, Keyword)
			Call BuildBasicSearchSQL(sWhere, person_th.per_tel, Keyword)
			Call BuildBasicSearchSQL(sWhere, person_th.per_fax, Keyword)
			Call BuildBasicSearchSQL(sWhere, person_th.per_email, Keyword)
			Call BuildBasicSearchSQL(sWhere, person_th.per_address, Keyword)
			Call BuildBasicSearchSQL(sWhere, person_th.per_show, Keyword)
			Call BuildBasicSearchSQL(sWhere, person_th.per_create, Keyword)
			Call BuildBasicSearchSQL(sWhere, person_th.per_update, Keyword)
			Call BuildBasicSearchSQL(sWhere, person_th.per_department, Keyword)
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
		sSearchKeyword = person_th.BasicSearch.Keyword
		sSearchType = person_th.BasicSearch.SearchType
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
			person_th.BasicSearch.setKeyword(sSearchKeyword)
			person_th.BasicSearch.setSearchType(sSearchType)
		End If
		BasicSearchWhere = sSearchStr
	End Function

	' Check if search parm exists
	Function CheckSearchParms()

		' Check basic search
		If person_th.BasicSearch.IssetSession() Then
			CheckSearchParms = True
			Exit Function
		End If
		CheckSearchParms = False
	End Function

	' Clear all search parameters
	Sub ResetSearchParms()

		' Clear search where
		SearchWhere = ""
		person_th.SearchWhere = SearchWhere

		' Clear basic search parameters
		Call ResetBasicSearchParms()
	End Sub

	' Load advanced search default values
	Function LoadAdvancedSearchDefault()
		LoadAdvancedSearchDefault = False
	End Function

	' Clear all basic search parameters
	Sub ResetBasicSearchParms()
		person_th.BasicSearch.UnsetSession()
	End Sub

	' -----------------------------------------------------------------
	' Restore all search parameters
	'
	Sub RestoreSearchParms()

		' Restore search flag
		RestoreSearch = True

		' Restore basic search values
		Call person_th.BasicSearch.Load()
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
			person_th.CurrentOrder = Request.QueryString("order")
			person_th.CurrentOrderType = Request.QueryString("ordertype")

			' Field per_id
			Call person_th.UpdateSort(person_th.per_id)

			' Field dept_id
			Call person_th.UpdateSort(person_th.dept_id)

			' Field office_id
			Call person_th.UpdateSort(person_th.office_id)

			' Field per_img
			Call person_th.UpdateSort(person_th.per_img)

			' Field per_en_name
			Call person_th.UpdateSort(person_th.per_en_name)

			' Field per_th_name
			Call person_th.UpdateSort(person_th.per_th_name)

			' Field per_position
			Call person_th.UpdateSort(person_th.per_position)

			' Field per_mobile
			Call person_th.UpdateSort(person_th.per_mobile)

			' Field per_tel
			Call person_th.UpdateSort(person_th.per_tel)

			' Field per_fax
			Call person_th.UpdateSort(person_th.per_fax)

			' Field per_email
			Call person_th.UpdateSort(person_th.per_email)

			' Field per_address
			Call person_th.UpdateSort(person_th.per_address)

			' Field per_show
			Call person_th.UpdateSort(person_th.per_show)

			' Field per_create
			Call person_th.UpdateSort(person_th.per_create)

			' Field per_update
			Call person_th.UpdateSort(person_th.per_update)

			' Field per_sort
			Call person_th.UpdateSort(person_th.per_sort)

			' Field per_department
			Call person_th.UpdateSort(person_th.per_department)
			person_th.StartRecordNumber = 1 ' Reset start position
		End If
	End Sub

	' -----------------------------------------------------------------
	' Load Sort Order parameters
	'
	Sub LoadSortOrder()
		Dim sOrderBy
		sOrderBy = person_th.SessionOrderBy ' Get order by from Session
		If sOrderBy = "" Then
			If person_th.SqlOrderBy <> "" Then
				sOrderBy = person_th.SqlOrderBy
				person_th.SessionOrderBy = sOrderBy
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
				person_th.SessionOrderBy = sOrderBy
				person_th.per_id.Sort = ""
				person_th.dept_id.Sort = ""
				person_th.office_id.Sort = ""
				person_th.per_img.Sort = ""
				person_th.per_en_name.Sort = ""
				person_th.per_th_name.Sort = ""
				person_th.per_position.Sort = ""
				person_th.per_mobile.Sort = ""
				person_th.per_tel.Sort = ""
				person_th.per_fax.Sort = ""
				person_th.per_email.Sort = ""
				person_th.per_address.Sort = ""
				person_th.per_show.Sort = ""
				person_th.per_create.Sort = ""
				person_th.per_update.Sort = ""
				person_th.per_sort.Sort = ""
				person_th.per_department.Sort = ""
			End If

			' Reset start position
			StartRec = 1
			person_th.StartRecordNumber = StartRec
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
		ListOptions.GetItem("checkbox").Body = "<label class=""checkbox""><input type=""checkbox"" name=""key_m"" value=""" & ew_HtmlEncode(person_th.per_id.CurrentValue) & """ onclick='ew_ClickMultiCheckbox(event, this);'></label>"
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
			For i = 0 to UBound(person_th.CustomActions.CustomArray)
				Action = person_th.CustomActions.CustomArray(i)(0)
				Name = person_th.CustomActions.CustomArray(i)(1)

				' Add custom action
				Call opt.Add("custom_" & Action)
				Set item = opt.GetItem("custom_" & Action)
				item.Body = "<a class=""ewAction ewCustomAction"" href="""" onclick=""ew_SubmitSelected(document.fperson_thlist, '" & ew_CurrentUrl & "', null, '" & Action & "');return false;"">" & Name & "</a>"
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
		sFilter = person_th.GetKeyFilter
		UserAction = Request.Form("useraction") & ""
		Processed = False
		If sFilter <> "" And UserAction <> "" Then
			person_th.CurrentFilter = sFilter
			sSql = person_th.SQL
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
				ElseIf person_th.CancelMessage <> "" Then
					FailureMessage = person_th.CancelMessage
					person_th.CancelMessage = ""
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
				person_th.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					person_th.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = person_th.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			person_th.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			person_th.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			person_th.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	'  Load basic search values
	'
	Function LoadBasicSearchValues()
		person_th.BasicSearch.Keyword = Request.QueryString(EW_TABLE_BASIC_SEARCH)&""
		If person_th.BasicSearch.Keyword <> "" Then Command = "search"
		person_th.BasicSearch.SearchType = Request.QueryString(EW_TABLE_BASIC_SEARCH_TYPE)&""
	End Function

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = person_th.CurrentFilter
		Call person_th.Recordset_Selecting(sFilter)
		person_th.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = person_th.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call person_th.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = person_th.KeyFilter

		' Call Row Selecting event
		Call person_th.Row_Selecting(sFilter)

		' Load sql based on filter
		person_th.CurrentFilter = sFilter
		sSql = person_th.SQL
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
		Call person_th.Row_Selected(RsRow)
		person_th.per_id.DbValue = RsRow("per_id")
		person_th.dept_id.DbValue = RsRow("dept_id")
		person_th.office_id.DbValue = RsRow("office_id")
		person_th.per_img.DbValue = RsRow("per_img")
		person_th.per_en_name.DbValue = RsRow("per_en_name")
		person_th.per_th_name.DbValue = RsRow("per_th_name")
		person_th.per_position.DbValue = RsRow("per_position")
		person_th.per_mobile.DbValue = RsRow("per_mobile")
		person_th.per_tel.DbValue = RsRow("per_tel")
		person_th.per_fax.DbValue = RsRow("per_fax")
		person_th.per_email.DbValue = RsRow("per_email")
		person_th.per_address.DbValue = RsRow("per_address")
		person_th.per_show.DbValue = RsRow("per_show")
		person_th.per_create.DbValue = RsRow("per_create")
		person_th.per_update.DbValue = RsRow("per_update")
		person_th.per_sort.DbValue = RsRow("per_sort")
		person_th.per_department.DbValue = RsRow("per_department")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		person_th.per_id.m_DbValue = Rs("per_id")
		person_th.dept_id.m_DbValue = Rs("dept_id")
		person_th.office_id.m_DbValue = Rs("office_id")
		person_th.per_img.m_DbValue = Rs("per_img")
		person_th.per_en_name.m_DbValue = Rs("per_en_name")
		person_th.per_th_name.m_DbValue = Rs("per_th_name")
		person_th.per_position.m_DbValue = Rs("per_position")
		person_th.per_mobile.m_DbValue = Rs("per_mobile")
		person_th.per_tel.m_DbValue = Rs("per_tel")
		person_th.per_fax.m_DbValue = Rs("per_fax")
		person_th.per_email.m_DbValue = Rs("per_email")
		person_th.per_address.m_DbValue = Rs("per_address")
		person_th.per_show.m_DbValue = Rs("per_show")
		person_th.per_create.m_DbValue = Rs("per_create")
		person_th.per_update.m_DbValue = Rs("per_update")
		person_th.per_sort.m_DbValue = Rs("per_sort")
		person_th.per_department.m_DbValue = Rs("per_department")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If person_th.GetKey("per_id")&"" <> "" Then
			person_th.per_id.CurrentValue = person_th.GetKey("per_id") ' per_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			person_th.CurrentFilter = person_th.KeyFilter
			Dim sSql
			sSql = person_th.SQL
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
		ViewUrl = person_th.ViewUrl("")
		EditUrl = person_th.EditUrl("")
		InlineEditUrl = person_th.InlineEditUrl
		CopyUrl = person_th.CopyUrl("")
		InlineCopyUrl = person_th.InlineCopyUrl
		DeleteUrl = person_th.DeleteUrl

		' Call Row Rendering event
		Call person_th.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' per_id
		' dept_id
		' office_id
		' per_img
		' per_en_name
		' per_th_name
		' per_position
		' per_mobile
		' per_tel
		' per_fax
		' per_email
		' per_address
		' per_show
		' per_create
		' per_update
		' per_sort
		' per_department
		' -----------
		'  View  Row
		' -----------

		If person_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' per_id
			person_th.per_id.ViewValue = person_th.per_id.CurrentValue
			person_th.per_id.ViewCustomAttributes = ""

			' dept_id
			person_th.dept_id.ViewValue = person_th.dept_id.CurrentValue
			person_th.dept_id.ViewCustomAttributes = ""

			' office_id
			person_th.office_id.ViewValue = person_th.office_id.CurrentValue
			person_th.office_id.ViewCustomAttributes = ""

			' per_img
			person_th.per_img.ViewValue = person_th.per_img.CurrentValue
			person_th.per_img.ViewCustomAttributes = ""

			' per_en_name
			person_th.per_en_name.ViewValue = person_th.per_en_name.CurrentValue
			person_th.per_en_name.ViewCustomAttributes = ""

			' per_th_name
			person_th.per_th_name.ViewValue = person_th.per_th_name.CurrentValue
			person_th.per_th_name.ViewCustomAttributes = ""

			' per_position
			person_th.per_position.ViewValue = person_th.per_position.CurrentValue
			person_th.per_position.ViewCustomAttributes = ""

			' per_mobile
			person_th.per_mobile.ViewValue = person_th.per_mobile.CurrentValue
			person_th.per_mobile.ViewCustomAttributes = ""

			' per_tel
			person_th.per_tel.ViewValue = person_th.per_tel.CurrentValue
			person_th.per_tel.ViewCustomAttributes = ""

			' per_fax
			person_th.per_fax.ViewValue = person_th.per_fax.CurrentValue
			person_th.per_fax.ViewCustomAttributes = ""

			' per_email
			person_th.per_email.ViewValue = person_th.per_email.CurrentValue
			person_th.per_email.ViewCustomAttributes = ""

			' per_address
			person_th.per_address.ViewValue = person_th.per_address.CurrentValue
			person_th.per_address.ViewCustomAttributes = ""

			' per_show
			person_th.per_show.ViewValue = person_th.per_show.CurrentValue
			person_th.per_show.ViewCustomAttributes = ""

			' per_create
			person_th.per_create.ViewValue = person_th.per_create.CurrentValue
			person_th.per_create.ViewCustomAttributes = ""

			' per_update
			person_th.per_update.ViewValue = person_th.per_update.CurrentValue
			person_th.per_update.ViewCustomAttributes = ""

			' per_sort
			person_th.per_sort.ViewValue = person_th.per_sort.CurrentValue
			person_th.per_sort.ViewCustomAttributes = ""

			' per_department
			person_th.per_department.ViewValue = person_th.per_department.CurrentValue
			person_th.per_department.ViewCustomAttributes = ""

			' View refer script
			' per_id

			person_th.per_id.LinkCustomAttributes = ""
			person_th.per_id.HrefValue = ""
			person_th.per_id.TooltipValue = ""

			' dept_id
			person_th.dept_id.LinkCustomAttributes = ""
			person_th.dept_id.HrefValue = ""
			person_th.dept_id.TooltipValue = ""

			' office_id
			person_th.office_id.LinkCustomAttributes = ""
			person_th.office_id.HrefValue = ""
			person_th.office_id.TooltipValue = ""

			' per_img
			person_th.per_img.LinkCustomAttributes = ""
			person_th.per_img.HrefValue = ""
			person_th.per_img.TooltipValue = ""

			' per_en_name
			person_th.per_en_name.LinkCustomAttributes = ""
			person_th.per_en_name.HrefValue = ""
			person_th.per_en_name.TooltipValue = ""

			' per_th_name
			person_th.per_th_name.LinkCustomAttributes = ""
			person_th.per_th_name.HrefValue = ""
			person_th.per_th_name.TooltipValue = ""

			' per_position
			person_th.per_position.LinkCustomAttributes = ""
			person_th.per_position.HrefValue = ""
			person_th.per_position.TooltipValue = ""

			' per_mobile
			person_th.per_mobile.LinkCustomAttributes = ""
			person_th.per_mobile.HrefValue = ""
			person_th.per_mobile.TooltipValue = ""

			' per_tel
			person_th.per_tel.LinkCustomAttributes = ""
			person_th.per_tel.HrefValue = ""
			person_th.per_tel.TooltipValue = ""

			' per_fax
			person_th.per_fax.LinkCustomAttributes = ""
			person_th.per_fax.HrefValue = ""
			person_th.per_fax.TooltipValue = ""

			' per_email
			person_th.per_email.LinkCustomAttributes = ""
			person_th.per_email.HrefValue = ""
			person_th.per_email.TooltipValue = ""

			' per_address
			person_th.per_address.LinkCustomAttributes = ""
			person_th.per_address.HrefValue = ""
			person_th.per_address.TooltipValue = ""

			' per_show
			person_th.per_show.LinkCustomAttributes = ""
			person_th.per_show.HrefValue = ""
			person_th.per_show.TooltipValue = ""

			' per_create
			person_th.per_create.LinkCustomAttributes = ""
			person_th.per_create.HrefValue = ""
			person_th.per_create.TooltipValue = ""

			' per_update
			person_th.per_update.LinkCustomAttributes = ""
			person_th.per_update.HrefValue = ""
			person_th.per_update.TooltipValue = ""

			' per_sort
			person_th.per_sort.LinkCustomAttributes = ""
			person_th.per_sort.HrefValue = ""
			person_th.per_sort.TooltipValue = ""

			' per_department
			person_th.per_department.LinkCustomAttributes = ""
			person_th.per_department.HrefValue = ""
			person_th.per_department.TooltipValue = ""
		End If

		' Call Row Rendered event
		If person_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call person_th.Row_Rendered()
		End If
	End Sub

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		url = ew_CurrentUrl
		url = ew_RegExReplace("\?cmd=reset(all){0,1}$", url, "") ' Remove cmd=reset / cmd=resetall
		Call Breadcrumb.Add("list", person_th.TableVar, url, person_th.TableVar, True)
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
