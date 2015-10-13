<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_banner_logo_02_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim banner_logo_02_th_add
Set banner_logo_02_th_add = New cbanner_logo_02_th_add
Set Page = banner_logo_02_th_add

' Page init processing
banner_logo_02_th_add.Page_Init()

' Page main processing
banner_logo_02_th_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
banner_logo_02_th_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var banner_logo_02_th_add = new ew_Page("banner_logo_02_th_add");
banner_logo_02_th_add.PageID = "add"; // Page ID
var EW_PAGE_ID = banner_logo_02_th_add.PageID; // For backward compatibility
// Form object
var fbanner_logo_02_thadd = new ew_Form("fbanner_logo_02_thadd");
// Validate form
fbanner_logo_02_thadd.Validate = function() {
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
				return this.OnError(elm, "<%= ew_JsEncode2(banner_logo_02_th.banner_id.FldErrMsg) %>");
			elm = this.GetElements("x" + infix + "_banner_sort");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(banner_logo_02_th.banner_sort.FldErrMsg) %>");
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
fbanner_logo_02_thadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fbanner_logo_02_thadd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fbanner_logo_02_thadd.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If banner_logo_02_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% banner_logo_02_th_add.ShowPageHeader() %>
<% banner_logo_02_th_add.ShowMessage %>
<form name="fbanner_logo_02_thadd" id="fbanner_logo_02_thadd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="banner_logo_02_th">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_banner_logo_02_thadd" class="table table-bordered table-striped">
<% If banner_logo_02_th.banner_id.Visible Then ' banner_id %>
	<tr id="r_banner_id">
		<td><span id="elh_banner_logo_02_th_banner_id"><%= banner_logo_02_th.banner_id.FldCaption %></span></td>
		<td<%= banner_logo_02_th.banner_id.CellAttributes %>>
<span id="el_banner_logo_02_th_banner_id" class="control-group">
<input type="text" data-field="x_banner_id" name="x_banner_id" id="x_banner_id" size="30" placeholder="<%= banner_logo_02_th.banner_id.PlaceHolder %>" value="<%= banner_logo_02_th.banner_id.EditValue %>"<%= banner_logo_02_th.banner_id.EditAttributes %>>
</span>
<%= banner_logo_02_th.banner_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If banner_logo_02_th.banner_img.Visible Then ' banner_img %>
	<tr id="r_banner_img">
		<td><span id="elh_banner_logo_02_th_banner_img"><%= banner_logo_02_th.banner_img.FldCaption %></span></td>
		<td<%= banner_logo_02_th.banner_img.CellAttributes %>>
<span id="el_banner_logo_02_th_banner_img" class="control-group">
<input type="text" data-field="x_banner_img" name="x_banner_img" id="x_banner_img" size="30" maxlength="255" placeholder="<%= banner_logo_02_th.banner_img.PlaceHolder %>" value="<%= banner_logo_02_th.banner_img.EditValue %>"<%= banner_logo_02_th.banner_img.EditAttributes %>>
</span>
<%= banner_logo_02_th.banner_img.CustomMsg %></td>
	</tr>
<% End If %>
<% If banner_logo_02_th.banner_link.Visible Then ' banner_link %>
	<tr id="r_banner_link">
		<td><span id="elh_banner_logo_02_th_banner_link"><%= banner_logo_02_th.banner_link.FldCaption %></span></td>
		<td<%= banner_logo_02_th.banner_link.CellAttributes %>>
<span id="el_banner_logo_02_th_banner_link" class="control-group">
<input type="text" data-field="x_banner_link" name="x_banner_link" id="x_banner_link" size="30" maxlength="255" placeholder="<%= banner_logo_02_th.banner_link.PlaceHolder %>" value="<%= banner_logo_02_th.banner_link.EditValue %>"<%= banner_logo_02_th.banner_link.EditAttributes %>>
</span>
<%= banner_logo_02_th.banner_link.CustomMsg %></td>
	</tr>
<% End If %>
<% If banner_logo_02_th.banner_sort.Visible Then ' banner_sort %>
	<tr id="r_banner_sort">
		<td><span id="elh_banner_logo_02_th_banner_sort"><%= banner_logo_02_th.banner_sort.FldCaption %></span></td>
		<td<%= banner_logo_02_th.banner_sort.CellAttributes %>>
<span id="el_banner_logo_02_th_banner_sort" class="control-group">
<input type="text" data-field="x_banner_sort" name="x_banner_sort" id="x_banner_sort" size="30" placeholder="<%= banner_logo_02_th.banner_sort.PlaceHolder %>" value="<%= banner_logo_02_th.banner_sort.EditValue %>"<%= banner_logo_02_th.banner_sort.EditAttributes %>>
</span>
<%= banner_logo_02_th.banner_sort.CustomMsg %></td>
	</tr>
<% End If %>
<% If banner_logo_02_th.start_date.Visible Then ' start_date %>
	<tr id="r_start_date">
		<td><span id="elh_banner_logo_02_th_start_date"><%= banner_logo_02_th.start_date.FldCaption %></span></td>
		<td<%= banner_logo_02_th.start_date.CellAttributes %>>
<span id="el_banner_logo_02_th_start_date" class="control-group">
<input type="text" data-field="x_start_date" name="x_start_date" id="x_start_date" placeholder="<%= banner_logo_02_th.start_date.PlaceHolder %>" value="<%= banner_logo_02_th.start_date.EditValue %>"<%= banner_logo_02_th.start_date.EditAttributes %>>
</span>
<%= banner_logo_02_th.start_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If banner_logo_02_th.end_date.Visible Then ' end_date %>
	<tr id="r_end_date">
		<td><span id="elh_banner_logo_02_th_end_date"><%= banner_logo_02_th.end_date.FldCaption %></span></td>
		<td<%= banner_logo_02_th.end_date.CellAttributes %>>
<span id="el_banner_logo_02_th_end_date" class="control-group">
<input type="text" data-field="x_end_date" name="x_end_date" id="x_end_date" placeholder="<%= banner_logo_02_th.end_date.PlaceHolder %>" value="<%= banner_logo_02_th.end_date.EditValue %>"<%= banner_logo_02_th.end_date.EditAttributes %>>
</span>
<%= banner_logo_02_th.end_date.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
fbanner_logo_02_thadd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
banner_logo_02_th_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set banner_logo_02_th_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cbanner_logo_02_th_add

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
		TableName = "banner_logo_02_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "banner_logo_02_th_add"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If banner_logo_02_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & banner_logo_02_th.TableVar & "&" ' add page token
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
		If banner_logo_02_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (banner_logo_02_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (banner_logo_02_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(banner_logo_02_th) Then Set banner_logo_02_th = New cbanner_logo_02_th
		Set Table = banner_logo_02_th

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "add"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "banner_logo_02_th"

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

		banner_logo_02_th.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set banner_logo_02_th = Nothing
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
			banner_logo_02_th.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("banner_id").Count > 0 Then
				banner_logo_02_th.banner_id.QueryStringValue = Request.QueryString("banner_id")
				Call banner_logo_02_th.SetKey("banner_id", banner_logo_02_th.banner_id.CurrentValue) ' Set up key
			Else
				Call banner_logo_02_th.SetKey("banner_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				banner_logo_02_th.CurrentAction = "C" ' Copy Record
			Else
				banner_logo_02_th.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				banner_logo_02_th.CurrentAction = "I" ' Form error, reset action
				banner_logo_02_th.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case banner_logo_02_th.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_banner_logo_02_thlist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				banner_logo_02_th.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = banner_logo_02_th.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_banner_logo_02_thview.asp" Then sReturnUrl = banner_logo_02_th.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					banner_logo_02_th.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		banner_logo_02_th.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call banner_logo_02_th.ResetAttrs()
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
		banner_logo_02_th.banner_id.CurrentValue = Null
		banner_logo_02_th.banner_id.OldValue = banner_logo_02_th.banner_id.CurrentValue
		banner_logo_02_th.banner_img.CurrentValue = Null
		banner_logo_02_th.banner_img.OldValue = banner_logo_02_th.banner_img.CurrentValue
		banner_logo_02_th.banner_link.CurrentValue = Null
		banner_logo_02_th.banner_link.OldValue = banner_logo_02_th.banner_link.CurrentValue
		banner_logo_02_th.banner_sort.CurrentValue = Null
		banner_logo_02_th.banner_sort.OldValue = banner_logo_02_th.banner_sort.CurrentValue
		banner_logo_02_th.start_date.CurrentValue = Null
		banner_logo_02_th.start_date.OldValue = banner_logo_02_th.start_date.CurrentValue
		banner_logo_02_th.end_date.CurrentValue = Null
		banner_logo_02_th.end_date.OldValue = banner_logo_02_th.end_date.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not banner_logo_02_th.banner_id.FldIsDetailKey Then banner_logo_02_th.banner_id.FormValue = ObjForm.GetValue("x_banner_id")
		If Not banner_logo_02_th.banner_img.FldIsDetailKey Then banner_logo_02_th.banner_img.FormValue = ObjForm.GetValue("x_banner_img")
		If Not banner_logo_02_th.banner_link.FldIsDetailKey Then banner_logo_02_th.banner_link.FormValue = ObjForm.GetValue("x_banner_link")
		If Not banner_logo_02_th.banner_sort.FldIsDetailKey Then banner_logo_02_th.banner_sort.FormValue = ObjForm.GetValue("x_banner_sort")
		If Not banner_logo_02_th.start_date.FldIsDetailKey Then banner_logo_02_th.start_date.FormValue = ObjForm.GetValue("x_start_date")
		If Not banner_logo_02_th.start_date.FldIsDetailKey Then banner_logo_02_th.start_date.CurrentValue = ew_UnFormatDateTime(banner_logo_02_th.start_date.CurrentValue, 8)
		If Not banner_logo_02_th.end_date.FldIsDetailKey Then banner_logo_02_th.end_date.FormValue = ObjForm.GetValue("x_end_date")
		If Not banner_logo_02_th.end_date.FldIsDetailKey Then banner_logo_02_th.end_date.CurrentValue = ew_UnFormatDateTime(banner_logo_02_th.end_date.CurrentValue, 8)
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		banner_logo_02_th.banner_id.CurrentValue = banner_logo_02_th.banner_id.FormValue
		banner_logo_02_th.banner_img.CurrentValue = banner_logo_02_th.banner_img.FormValue
		banner_logo_02_th.banner_link.CurrentValue = banner_logo_02_th.banner_link.FormValue
		banner_logo_02_th.banner_sort.CurrentValue = banner_logo_02_th.banner_sort.FormValue
		banner_logo_02_th.start_date.CurrentValue = banner_logo_02_th.start_date.FormValue
		banner_logo_02_th.start_date.CurrentValue = ew_UnFormatDateTime(banner_logo_02_th.start_date.CurrentValue, 8)
		banner_logo_02_th.end_date.CurrentValue = banner_logo_02_th.end_date.FormValue
		banner_logo_02_th.end_date.CurrentValue = ew_UnFormatDateTime(banner_logo_02_th.end_date.CurrentValue, 8)
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = banner_logo_02_th.KeyFilter

		' Call Row Selecting event
		Call banner_logo_02_th.Row_Selecting(sFilter)

		' Load sql based on filter
		banner_logo_02_th.CurrentFilter = sFilter
		sSql = banner_logo_02_th.SQL
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
		Call banner_logo_02_th.Row_Selected(RsRow)
		banner_logo_02_th.banner_id.DbValue = RsRow("banner_id")
		banner_logo_02_th.banner_img.DbValue = RsRow("banner_img")
		banner_logo_02_th.banner_link.DbValue = RsRow("banner_link")
		banner_logo_02_th.banner_sort.DbValue = RsRow("banner_sort")
		banner_logo_02_th.start_date.DbValue = RsRow("start_date")
		banner_logo_02_th.end_date.DbValue = RsRow("end_date")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		banner_logo_02_th.banner_id.m_DbValue = Rs("banner_id")
		banner_logo_02_th.banner_img.m_DbValue = Rs("banner_img")
		banner_logo_02_th.banner_link.m_DbValue = Rs("banner_link")
		banner_logo_02_th.banner_sort.m_DbValue = Rs("banner_sort")
		banner_logo_02_th.start_date.m_DbValue = Rs("start_date")
		banner_logo_02_th.end_date.m_DbValue = Rs("end_date")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If banner_logo_02_th.GetKey("banner_id")&"" <> "" Then
			banner_logo_02_th.banner_id.CurrentValue = banner_logo_02_th.GetKey("banner_id") ' banner_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			banner_logo_02_th.CurrentFilter = banner_logo_02_th.KeyFilter
			Dim sSql
			sSql = banner_logo_02_th.SQL
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

		Call banner_logo_02_th.Row_Rendering()

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

		If banner_logo_02_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' banner_id
			banner_logo_02_th.banner_id.ViewValue = banner_logo_02_th.banner_id.CurrentValue
			banner_logo_02_th.banner_id.ViewCustomAttributes = ""

			' banner_img
			banner_logo_02_th.banner_img.ViewValue = banner_logo_02_th.banner_img.CurrentValue
			banner_logo_02_th.banner_img.ViewCustomAttributes = ""

			' banner_link
			banner_logo_02_th.banner_link.ViewValue = banner_logo_02_th.banner_link.CurrentValue
			banner_logo_02_th.banner_link.ViewCustomAttributes = ""

			' banner_sort
			banner_logo_02_th.banner_sort.ViewValue = banner_logo_02_th.banner_sort.CurrentValue
			banner_logo_02_th.banner_sort.ViewCustomAttributes = ""

			' start_date
			banner_logo_02_th.start_date.ViewValue = banner_logo_02_th.start_date.CurrentValue
			banner_logo_02_th.start_date.ViewCustomAttributes = ""

			' end_date
			banner_logo_02_th.end_date.ViewValue = banner_logo_02_th.end_date.CurrentValue
			banner_logo_02_th.end_date.ViewCustomAttributes = ""

			' View refer script
			' banner_id

			banner_logo_02_th.banner_id.LinkCustomAttributes = ""
			banner_logo_02_th.banner_id.HrefValue = ""
			banner_logo_02_th.banner_id.TooltipValue = ""

			' banner_img
			banner_logo_02_th.banner_img.LinkCustomAttributes = ""
			banner_logo_02_th.banner_img.HrefValue = ""
			banner_logo_02_th.banner_img.TooltipValue = ""

			' banner_link
			banner_logo_02_th.banner_link.LinkCustomAttributes = ""
			banner_logo_02_th.banner_link.HrefValue = ""
			banner_logo_02_th.banner_link.TooltipValue = ""

			' banner_sort
			banner_logo_02_th.banner_sort.LinkCustomAttributes = ""
			banner_logo_02_th.banner_sort.HrefValue = ""
			banner_logo_02_th.banner_sort.TooltipValue = ""

			' start_date
			banner_logo_02_th.start_date.LinkCustomAttributes = ""
			banner_logo_02_th.start_date.HrefValue = ""
			banner_logo_02_th.start_date.TooltipValue = ""

			' end_date
			banner_logo_02_th.end_date.LinkCustomAttributes = ""
			banner_logo_02_th.end_date.HrefValue = ""
			banner_logo_02_th.end_date.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf banner_logo_02_th.RowType = EW_ROWTYPE_ADD Then ' Add row

			' banner_id
			banner_logo_02_th.banner_id.EditCustomAttributes = ""
			banner_logo_02_th.banner_id.EditValue = ew_HtmlEncode(banner_logo_02_th.banner_id.CurrentValue)
			banner_logo_02_th.banner_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(banner_logo_02_th.banner_id.FldCaption))

			' banner_img
			banner_logo_02_th.banner_img.EditCustomAttributes = ""
			banner_logo_02_th.banner_img.EditValue = ew_HtmlEncode(banner_logo_02_th.banner_img.CurrentValue)
			banner_logo_02_th.banner_img.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(banner_logo_02_th.banner_img.FldCaption))

			' banner_link
			banner_logo_02_th.banner_link.EditCustomAttributes = ""
			banner_logo_02_th.banner_link.EditValue = ew_HtmlEncode(banner_logo_02_th.banner_link.CurrentValue)
			banner_logo_02_th.banner_link.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(banner_logo_02_th.banner_link.FldCaption))

			' banner_sort
			banner_logo_02_th.banner_sort.EditCustomAttributes = ""
			banner_logo_02_th.banner_sort.EditValue = ew_HtmlEncode(banner_logo_02_th.banner_sort.CurrentValue)
			banner_logo_02_th.banner_sort.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(banner_logo_02_th.banner_sort.FldCaption))

			' start_date
			banner_logo_02_th.start_date.EditCustomAttributes = ""
			banner_logo_02_th.start_date.EditValue = ew_HtmlEncode(banner_logo_02_th.start_date.CurrentValue)
			banner_logo_02_th.start_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(banner_logo_02_th.start_date.FldCaption))

			' end_date
			banner_logo_02_th.end_date.EditCustomAttributes = ""
			banner_logo_02_th.end_date.EditValue = ew_HtmlEncode(banner_logo_02_th.end_date.CurrentValue)
			banner_logo_02_th.end_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(banner_logo_02_th.end_date.FldCaption))

			' Edit refer script
			' banner_id

			banner_logo_02_th.banner_id.HrefValue = ""

			' banner_img
			banner_logo_02_th.banner_img.HrefValue = ""

			' banner_link
			banner_logo_02_th.banner_link.HrefValue = ""

			' banner_sort
			banner_logo_02_th.banner_sort.HrefValue = ""

			' start_date
			banner_logo_02_th.start_date.HrefValue = ""

			' end_date
			banner_logo_02_th.end_date.HrefValue = ""
		End If
		If banner_logo_02_th.RowType = EW_ROWTYPE_ADD Or banner_logo_02_th.RowType = EW_ROWTYPE_EDIT Or banner_logo_02_th.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call banner_logo_02_th.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If banner_logo_02_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call banner_logo_02_th.Row_Rendered()
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
		If Not ew_CheckInteger(banner_logo_02_th.banner_id.FormValue) Then
			Call ew_AddMessage(gsFormError, banner_logo_02_th.banner_id.FldErrMsg)
		End If
		If Not ew_CheckInteger(banner_logo_02_th.banner_sort.FormValue) Then
			Call ew_AddMessage(gsFormError, banner_logo_02_th.banner_sort.FldErrMsg)
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
		If banner_logo_02_th.banner_id.CurrentValue <> "" Then ' Check field with unique index
			sFilter = "([banner_id] = " & ew_AdjustSql(banner_logo_02_th.banner_id.CurrentValue) & ")"
			Set RsChk = banner_logo_02_th.LoadRs(sFilter)
			If Not (RsChk Is Nothing) Then
				sIdxErrMsg = Replace(Language.Phrase("DupIndex"), "%f", banner_logo_02_th.banner_id.FldCaption)
				sIdxErrMsg = Replace(sIdxErrMsg, "%v", banner_logo_02_th.banner_id.CurrentValue)
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
		banner_logo_02_th.CurrentFilter = sFilter
		sSql = banner_logo_02_th.SQL
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
		Call banner_logo_02_th.banner_id.SetDbValue(Rs, banner_logo_02_th.banner_id.CurrentValue, Null, False)

		' Field banner_img
		Call banner_logo_02_th.banner_img.SetDbValue(Rs, banner_logo_02_th.banner_img.CurrentValue, Null, False)

		' Field banner_link
		Call banner_logo_02_th.banner_link.SetDbValue(Rs, banner_logo_02_th.banner_link.CurrentValue, Null, False)

		' Field banner_sort
		Call banner_logo_02_th.banner_sort.SetDbValue(Rs, banner_logo_02_th.banner_sort.CurrentValue, Null, False)

		' Field start_date
		Call banner_logo_02_th.start_date.SetDbValue(Rs, banner_logo_02_th.start_date.CurrentValue, Null, False)

		' Field end_date
		Call banner_logo_02_th.end_date.SetDbValue(Rs, banner_logo_02_th.end_date.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = banner_logo_02_th.Row_Inserting(RsOld, Rs)

		' Check if key value entered
		If bInsertRow And banner_logo_02_th.ValidateKey And banner_logo_02_th.banner_id.CurrentValue = "" And banner_logo_02_th.banner_id.SessionValue = "" Then
			FailureMessage = Language.Phrase("InvalidKeyValue")
			bInsertRow = False
		End If

		' Check for duplicate key
		Dim sKeyErrMsg
		If bInsertRow And banner_logo_02_th.ValidateKey Then
			sFilter = banner_logo_02_th.KeyFilter
			Set RsChk = banner_logo_02_th.LoadRs(sFilter)
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
			ElseIf banner_logo_02_th.CancelMessage <> "" Then
				FailureMessage = banner_logo_02_th.CancelMessage
				banner_logo_02_th.CancelMessage = ""
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
			Call banner_logo_02_th.Row_Inserted(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", banner_logo_02_th.TableVar, "pom_banner_logo_02_thlist.asp", banner_logo_02_th.TableVar, True)
		PageId = ew_IIf(banner_logo_02_th.CurrentAction = "C", "Copy", "Add")
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
