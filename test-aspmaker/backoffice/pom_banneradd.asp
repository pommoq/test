<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_bannerinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim banner_add
Set banner_add = New cbanner_add
Set Page = banner_add

' Page init processing
banner_add.Page_Init()

' Page main processing
banner_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
banner_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var banner_add = new ew_Page("banner_add");
banner_add.PageID = "add"; // Page ID
var EW_PAGE_ID = banner_add.PageID; // For backward compatibility
// Form object
var fbanneradd = new ew_Form("fbanneradd");
// Validate form
fbanneradd.Validate = function() {
	if (!this.ValidateRequired)
		return true; // Ignore validation
	var $ = jQuery, fobj = this.GetForm(), $fobj = $(fobj);
	this.PostAutoSuggest();
	if ($fobj.find("#a_confirm").val() == "F")
		return true;
	var elm, felm, uelm, addcnt = 0;
	var $k = $fobj.find("#" + this.FormKeyCountName); // Get key_count
	var rowcnt = ($k[0]) ? parseInt($k.val(), 10) : 1;
	var startcnt = (rowcnt == 0) ? 0 : 1; // Check rowcnt == 0 => Inline-Add
	var gridinsert = $fobj.find("#a_list").val() == "gridinsert";
	for (var i = startcnt; i <= rowcnt; i++) {
		var infix = ($k[0]) ? String(i) : "";
		$fobj.data("rowindex", infix);
			elm = this.GetElements("x" + infix + "_banner_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(banner.banner_id.FldErrMsg) %>");
			elm = this.GetElements("x" + infix + "_banner_sort");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(banner.banner_sort.FldErrMsg) %>");
			// Set up row object
			ew_ElementsToRow(fobj);
			// Fire Form_CustomValidate event
			if (!this.Form_CustomValidate(fobj))
				return false;
	}
	// Process detail forms
	var dfs = $fobj.find("input[name='detailpage']").get();
	for (var i = 0; i < dfs.length; i++) {
		var df = dfs[i], val = df.value;
		if (val && ewForms[val])
			if (!ewForms[val].Validate())
				return false;
	}
	return true;
}
// Form_CustomValidate event
fbanneradd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fbanneradd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fbanneradd.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If banner.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% banner_add.ShowPageHeader() %>
<% banner_add.ShowMessage %>
<form name="fbanneradd" id="fbanneradd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="banner">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_banneradd" class="table table-bordered table-striped">
<% If banner.banner_id.Visible Then ' banner_id %>
	<tr id="r_banner_id">
		<td><span id="elh_banner_banner_id"><%= banner.banner_id.FldCaption %></span></td>
		<td<%= banner.banner_id.CellAttributes %>>
<span id="el_banner_banner_id" class="control-group">
<input type="text" data-field="x_banner_id" name="x_banner_id" id="x_banner_id" size="30" placeholder="<%= banner.banner_id.PlaceHolder %>" value="<%= banner.banner_id.EditValue %>"<%= banner.banner_id.EditAttributes %>>
</span>
<%= banner.banner_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If banner.banner_img.Visible Then ' banner_img %>
	<tr id="r_banner_img">
		<td><span id="elh_banner_banner_img"><%= banner.banner_img.FldCaption %></span></td>
		<td<%= banner.banner_img.CellAttributes %>>
<span id="el_banner_banner_img" class="control-group">
<input type="text" data-field="x_banner_img" name="x_banner_img" id="x_banner_img" size="30" maxlength="255" placeholder="<%= banner.banner_img.PlaceHolder %>" value="<%= banner.banner_img.EditValue %>"<%= banner.banner_img.EditAttributes %>>
</span>
<%= banner.banner_img.CustomMsg %></td>
	</tr>
<% End If %>
<% If banner.banner_link.Visible Then ' banner_link %>
	<tr id="r_banner_link">
		<td><span id="elh_banner_banner_link"><%= banner.banner_link.FldCaption %></span></td>
		<td<%= banner.banner_link.CellAttributes %>>
<span id="el_banner_banner_link" class="control-group">
<input type="text" data-field="x_banner_link" name="x_banner_link" id="x_banner_link" size="30" maxlength="255" placeholder="<%= banner.banner_link.PlaceHolder %>" value="<%= banner.banner_link.EditValue %>"<%= banner.banner_link.EditAttributes %>>
</span>
<%= banner.banner_link.CustomMsg %></td>
	</tr>
<% End If %>
<% If banner.banner_sort.Visible Then ' banner_sort %>
	<tr id="r_banner_sort">
		<td><span id="elh_banner_banner_sort"><%= banner.banner_sort.FldCaption %></span></td>
		<td<%= banner.banner_sort.CellAttributes %>>
<span id="el_banner_banner_sort" class="control-group">
<input type="text" data-field="x_banner_sort" name="x_banner_sort" id="x_banner_sort" size="30" placeholder="<%= banner.banner_sort.PlaceHolder %>" value="<%= banner.banner_sort.EditValue %>"<%= banner.banner_sort.EditAttributes %>>
</span>
<%= banner.banner_sort.CustomMsg %></td>
	</tr>
<% End If %>
<% If banner.start_date.Visible Then ' start_date %>
	<tr id="r_start_date">
		<td><span id="elh_banner_start_date"><%= banner.start_date.FldCaption %></span></td>
		<td<%= banner.start_date.CellAttributes %>>
<span id="el_banner_start_date" class="control-group">
<input type="text" data-field="x_start_date" name="x_start_date" id="x_start_date" placeholder="<%= banner.start_date.PlaceHolder %>" value="<%= banner.start_date.EditValue %>"<%= banner.start_date.EditAttributes %>>
</span>
<%= banner.start_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If banner.end_date.Visible Then ' end_date %>
	<tr id="r_end_date">
		<td><span id="elh_banner_end_date"><%= banner.end_date.FldCaption %></span></td>
		<td<%= banner.end_date.CellAttributes %>>
<span id="el_banner_end_date" class="control-group">
<input type="text" data-field="x_end_date" name="x_end_date" id="x_end_date" placeholder="<%= banner.end_date.PlaceHolder %>" value="<%= banner.end_date.EditValue %>"<%= banner.end_date.EditAttributes %>>
</span>
<%= banner.end_date.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
fbanneradd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
banner_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set banner_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cbanner_add

	' Page ID
	Public Property Get PageID()
		PageID = "add"
	End Property

	' Project ID
	Public Property Get ProjectID()
		ProjectID = "{324ED72D-DE20-46F7-B12E-7AF8CE8711A6}"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "banner"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "banner_add"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If banner.UseTokenInUrl Then PageUrl = PageUrl & "t=" & banner.TableVar & "&" ' add page token
	End Property

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
		If banner.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (banner.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (banner.TableVar = Request.QueryString("t"))
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

		' Initialize language object
		If IsEmpty(Language) Then
			Set Language = New cLanguage
			Call Language.LoadPhrases()
		End If

		' Initialize table object
		If IsEmpty(banner) Then Set banner = New cbanner
		Set Table = banner

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "add"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "banner"

		' Open connection to the database
		If IsEmpty(Conn) Then Call ew_Connect()
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

		' Create form object
		'If Request.ServerVariables("HTTP_CONTENT_TYPE") = "application/x-www-form-urlencoded" Then

			Set ObjForm = New cFormObj

		'Else
		'	Set ObjForm = ew_GetUploadObj()
		'End If

		banner.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

		' Global page loading event (in userfn7.asp)
		Page_Loading()

		' Page load event, used in current page
		Page_Load()
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
		Set banner = Nothing
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

	Dim DbMasterFilter, DbDetailFilter
	Dim Priv
	Dim OldRecordset
	Dim CopyRecord

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Process form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			banner.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("banner_id").Count > 0 Then
				banner.banner_id.QueryStringValue = Request.QueryString("banner_id")
				Call banner.SetKey("banner_id", banner.banner_id.CurrentValue) ' Set up key
			Else
				Call banner.SetKey("banner_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				banner.CurrentAction = "C" ' Copy Record
			Else
				banner.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				banner.CurrentAction = "I" ' Form error, reset action
				banner.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case banner.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_bannerlist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				banner.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = banner.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_bannerview.asp" Then sReturnUrl = banner.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					banner.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		banner.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call banner.ResetAttrs()
		Call RenderRow()
	End Sub

	' -----------------------------------------------------------------
	' Function Get upload files
	'
	Function GetUploadFiles()

		' Get upload data
	End Function

	' -----------------------------------------------------------------
	' Load default values
	'
	Function LoadDefaultValues()
		banner.banner_id.CurrentValue = Null
		banner.banner_id.OldValue = banner.banner_id.CurrentValue
		banner.banner_img.CurrentValue = Null
		banner.banner_img.OldValue = banner.banner_img.CurrentValue
		banner.banner_link.CurrentValue = Null
		banner.banner_link.OldValue = banner.banner_link.CurrentValue
		banner.banner_sort.CurrentValue = Null
		banner.banner_sort.OldValue = banner.banner_sort.CurrentValue
		banner.start_date.CurrentValue = Null
		banner.start_date.OldValue = banner.start_date.CurrentValue
		banner.end_date.CurrentValue = Null
		banner.end_date.OldValue = banner.end_date.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not banner.banner_id.FldIsDetailKey Then banner.banner_id.FormValue = ObjForm.GetValue("x_banner_id")
		If Not banner.banner_img.FldIsDetailKey Then banner.banner_img.FormValue = ObjForm.GetValue("x_banner_img")
		If Not banner.banner_link.FldIsDetailKey Then banner.banner_link.FormValue = ObjForm.GetValue("x_banner_link")
		If Not banner.banner_sort.FldIsDetailKey Then banner.banner_sort.FormValue = ObjForm.GetValue("x_banner_sort")
		If Not banner.start_date.FldIsDetailKey Then banner.start_date.FormValue = ObjForm.GetValue("x_start_date")
		If Not banner.start_date.FldIsDetailKey Then banner.start_date.CurrentValue = ew_UnFormatDateTime(banner.start_date.CurrentValue, 8)
		If Not banner.end_date.FldIsDetailKey Then banner.end_date.FormValue = ObjForm.GetValue("x_end_date")
		If Not banner.end_date.FldIsDetailKey Then banner.end_date.CurrentValue = ew_UnFormatDateTime(banner.end_date.CurrentValue, 8)
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		banner.banner_id.CurrentValue = banner.banner_id.FormValue
		banner.banner_img.CurrentValue = banner.banner_img.FormValue
		banner.banner_link.CurrentValue = banner.banner_link.FormValue
		banner.banner_sort.CurrentValue = banner.banner_sort.FormValue
		banner.start_date.CurrentValue = banner.start_date.FormValue
		banner.start_date.CurrentValue = ew_UnFormatDateTime(banner.start_date.CurrentValue, 8)
		banner.end_date.CurrentValue = banner.end_date.FormValue
		banner.end_date.CurrentValue = ew_UnFormatDateTime(banner.end_date.CurrentValue, 8)
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = banner.KeyFilter

		' Call Row Selecting event
		Call banner.Row_Selecting(sFilter)

		' Load sql based on filter
		banner.CurrentFilter = sFilter
		sSql = banner.SQL
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
		Call banner.Row_Selected(RsRow)
		banner.banner_id.DbValue = RsRow("banner_id")
		banner.banner_img.DbValue = RsRow("banner_img")
		banner.banner_link.DbValue = RsRow("banner_link")
		banner.banner_sort.DbValue = RsRow("banner_sort")
		banner.start_date.DbValue = RsRow("start_date")
		banner.end_date.DbValue = RsRow("end_date")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		banner.banner_id.m_DbValue = Rs("banner_id")
		banner.banner_img.m_DbValue = Rs("banner_img")
		banner.banner_link.m_DbValue = Rs("banner_link")
		banner.banner_sort.m_DbValue = Rs("banner_sort")
		banner.start_date.m_DbValue = Rs("start_date")
		banner.end_date.m_DbValue = Rs("end_date")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If banner.GetKey("banner_id")&"" <> "" Then
			banner.banner_id.CurrentValue = banner.GetKey("banner_id") ' banner_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			banner.CurrentFilter = banner.KeyFilter
			Dim sSql
			sSql = banner.SQL
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
		' Call Row Rendering event

		Call banner.Row_Rendering()

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

		If banner.RowType = EW_ROWTYPE_VIEW Then ' View row

			' banner_id
			banner.banner_id.ViewValue = banner.banner_id.CurrentValue
			banner.banner_id.ViewCustomAttributes = ""

			' banner_img
			banner.banner_img.ViewValue = banner.banner_img.CurrentValue
			banner.banner_img.ViewCustomAttributes = ""

			' banner_link
			banner.banner_link.ViewValue = banner.banner_link.CurrentValue
			banner.banner_link.ViewCustomAttributes = ""

			' banner_sort
			banner.banner_sort.ViewValue = banner.banner_sort.CurrentValue
			banner.banner_sort.ViewCustomAttributes = ""

			' start_date
			banner.start_date.ViewValue = banner.start_date.CurrentValue
			banner.start_date.ViewCustomAttributes = ""

			' end_date
			banner.end_date.ViewValue = banner.end_date.CurrentValue
			banner.end_date.ViewCustomAttributes = ""

			' View refer script
			' banner_id

			banner.banner_id.LinkCustomAttributes = ""
			banner.banner_id.HrefValue = ""
			banner.banner_id.TooltipValue = ""

			' banner_img
			banner.banner_img.LinkCustomAttributes = ""
			banner.banner_img.HrefValue = ""
			banner.banner_img.TooltipValue = ""

			' banner_link
			banner.banner_link.LinkCustomAttributes = ""
			banner.banner_link.HrefValue = ""
			banner.banner_link.TooltipValue = ""

			' banner_sort
			banner.banner_sort.LinkCustomAttributes = ""
			banner.banner_sort.HrefValue = ""
			banner.banner_sort.TooltipValue = ""

			' start_date
			banner.start_date.LinkCustomAttributes = ""
			banner.start_date.HrefValue = ""
			banner.start_date.TooltipValue = ""

			' end_date
			banner.end_date.LinkCustomAttributes = ""
			banner.end_date.HrefValue = ""
			banner.end_date.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf banner.RowType = EW_ROWTYPE_ADD Then ' Add row

			' banner_id
			banner.banner_id.EditCustomAttributes = ""
			banner.banner_id.EditValue = ew_HtmlEncode(banner.banner_id.CurrentValue)
			banner.banner_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(banner.banner_id.FldCaption))

			' banner_img
			banner.banner_img.EditCustomAttributes = ""
			banner.banner_img.EditValue = ew_HtmlEncode(banner.banner_img.CurrentValue)
			banner.banner_img.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(banner.banner_img.FldCaption))

			' banner_link
			banner.banner_link.EditCustomAttributes = ""
			banner.banner_link.EditValue = ew_HtmlEncode(banner.banner_link.CurrentValue)
			banner.banner_link.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(banner.banner_link.FldCaption))

			' banner_sort
			banner.banner_sort.EditCustomAttributes = ""
			banner.banner_sort.EditValue = ew_HtmlEncode(banner.banner_sort.CurrentValue)
			banner.banner_sort.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(banner.banner_sort.FldCaption))

			' start_date
			banner.start_date.EditCustomAttributes = ""
			banner.start_date.EditValue = ew_HtmlEncode(banner.start_date.CurrentValue)
			banner.start_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(banner.start_date.FldCaption))

			' end_date
			banner.end_date.EditCustomAttributes = ""
			banner.end_date.EditValue = ew_HtmlEncode(banner.end_date.CurrentValue)
			banner.end_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(banner.end_date.FldCaption))

			' Edit refer script
			' banner_id

			banner.banner_id.HrefValue = ""

			' banner_img
			banner.banner_img.HrefValue = ""

			' banner_link
			banner.banner_link.HrefValue = ""

			' banner_sort
			banner.banner_sort.HrefValue = ""

			' start_date
			banner.start_date.HrefValue = ""

			' end_date
			banner.end_date.HrefValue = ""
		End If
		If banner.RowType = EW_ROWTYPE_ADD Or banner.RowType = EW_ROWTYPE_EDIT Or banner.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call banner.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If banner.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call banner.Row_Rendered()
		End If
	End Sub

	' -----------------------------------------------------------------
	' Validate form
	'
	Function ValidateForm()

		' Initialize
		gsFormError = ""

		' Check if validation required
		If Not EW_SERVER_VALIDATE Then
			ValidateForm = (gsFormError = "")
			Exit Function
		End If
		If Not ew_CheckInteger(banner.banner_id.FormValue) Then
			Call ew_AddMessage(gsFormError, banner.banner_id.FldErrMsg)
		End If
		If Not ew_CheckInteger(banner.banner_sort.FormValue) Then
			Call ew_AddMessage(gsFormError, banner.banner_sort.FldErrMsg)
		End If

		' Return validate result
		ValidateForm = (gsFormError = "")

		' Call Form Custom Validate event
		Dim sFormCustomError
		sFormCustomError = ""
		ValidateForm = ValidateForm And Form_CustomValidate(sFormCustomError)
		If sFormCustomError <> "" Then
			Call ew_AddMessage(gsFormError, sFormCustomError)
		End If
	End Function

	' -----------------------------------------------------------------
	' Add record
	'
	Function AddRow(RsOld)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim Rs, sSql, sFilter
		Dim RsNew
		Dim bInsertRow
		Dim RsChk
		Dim sIdxErrMsg

		' Clear any previous errors
		Err.Clear
		Dim RsMaster, sMasterUserIdMsg, sMasterFilter, bCheckMasterRecord
		If banner.banner_id.CurrentValue <> "" Then ' Check field with unique index
			sFilter = "([banner_id] = " & ew_AdjustSql(banner.banner_id.CurrentValue) & ")"
			Set RsChk = banner.LoadRs(sFilter)
			If Not (RsChk Is Nothing) Then
				sIdxErrMsg = Replace(Language.Phrase("DupIndex"), "%f", banner.banner_id.FldCaption)
				sIdxErrMsg = Replace(sIdxErrMsg, "%v", banner.banner_id.CurrentValue)
				FailureMessage = sIdxErrMsg
				RsChk.Close
				Set RsChk = Nothing
				AddRow = False
				Exit Function
			End If
		End If

		' Load db values from rsold
		If Not IsNull(RsOld) Then
			Call LoadDbValues(RsOld)
		End If

		' Add new record
		sFilter = "(0 = 1)"
		banner.CurrentFilter = sFilter
		sSql = banner.SQL
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = EW_CURSORLOCATION
		Rs.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		Rs.AddNew
		If Err.Number <> 0 Then
			Message = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Field banner_id
		Call banner.banner_id.SetDbValue(Rs, banner.banner_id.CurrentValue, Null, False)

		' Field banner_img
		Call banner.banner_img.SetDbValue(Rs, banner.banner_img.CurrentValue, Null, False)

		' Field banner_link
		Call banner.banner_link.SetDbValue(Rs, banner.banner_link.CurrentValue, Null, False)

		' Field banner_sort
		Call banner.banner_sort.SetDbValue(Rs, banner.banner_sort.CurrentValue, Null, False)

		' Field start_date
		Call banner.start_date.SetDbValue(Rs, banner.start_date.CurrentValue, Null, False)

		' Field end_date
		Call banner.end_date.SetDbValue(Rs, banner.end_date.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = banner.Row_Inserting(RsOld, Rs)

		' Check if key value entered
		If bInsertRow And banner.ValidateKey And banner.banner_id.CurrentValue = "" And banner.banner_id.SessionValue = "" Then
			FailureMessage = Language.Phrase("InvalidKeyValue")
			bInsertRow = False
		End If

		' Check for duplicate key
		Dim sKeyErrMsg
		If bInsertRow And banner.ValidateKey Then
			sFilter = banner.KeyFilter
			Set RsChk = banner.LoadRs(sFilter)
			If Not (RsChk Is Nothing) Then
				sKeyErrMsg = Replace(Language.Phrase("DupKey"), "%f", sFilter)
				FailureMessage = sKeyErrMsg
				RsChk.Close
				Set RsChk = Nothing
				bInsertRow = False
			End If
		End If
		If bInsertRow Then

			' Clone new recordset object
			Set RsNew = ew_CloneRs(Rs)
			Rs.Update
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				AddRow = False
			Else
				AddRow = True
			End If
			If AddRow Then
			End If
		Else
			Rs.CancelUpdate

			' Set up error message
			If SuccessMessage <> "" Or FailureMessage <> "" Then

				' Use the message, do nothing
			ElseIf banner.CancelMessage <> "" Then
				FailureMessage = banner.CancelMessage
				banner.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("InsertCancelled")
			End If
			AddRow = False
		End If
		Rs.Close
		Set Rs = Nothing
		If AddRow Then
		End If
		If AddRow Then

			' Call Row Inserted event
			Call banner.Row_Inserted(RsOld, RsNew)
		End If
		If IsObject(RsNew) Then
			RsNew.Close
			Set RsNew = Nothing
		End If
	End Function

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", banner.TableVar, "pom_bannerlist.asp", banner.TableVar, True)
		PageId = ew_IIf(banner.CurrentAction = "C", "Copy", "Add")
		Call Breadcrumb.Add("add", PageId, ew_CurrentUrl, "", False)
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
End Class
%>
