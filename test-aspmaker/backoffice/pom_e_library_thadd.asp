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
Dim e_library_th_add
Set e_library_th_add = New ce_library_th_add
Set Page = e_library_th_add

' Page init processing
e_library_th_add.Page_Init()

' Page main processing
e_library_th_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
e_library_th_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var e_library_th_add = new ew_Page("e_library_th_add");
e_library_th_add.PageID = "add"; // Page ID
var EW_PAGE_ID = e_library_th_add.PageID; // For backward compatibility
// Form object
var fe_library_thadd = new ew_Form("fe_library_thadd");
// Validate form
fe_library_thadd.Validate = function() {
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
			elm = this.GetElements("x" + infix + "_el_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(e_library_th.el_id.FldErrMsg) %>");
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
fe_library_thadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fe_library_thadd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fe_library_thadd.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If e_library_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% e_library_th_add.ShowPageHeader() %>
<% e_library_th_add.ShowMessage %>
<form name="fe_library_thadd" id="fe_library_thadd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="e_library_th">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_e_library_thadd" class="table table-bordered table-striped">
<% If e_library_th.el_id.Visible Then ' el_id %>
	<tr id="r_el_id">
		<td><span id="elh_e_library_th_el_id"><%= e_library_th.el_id.FldCaption %></span></td>
		<td<%= e_library_th.el_id.CellAttributes %>>
<span id="el_e_library_th_el_id" class="control-group">
<input type="text" data-field="x_el_id" name="x_el_id" id="x_el_id" size="30" placeholder="<%= e_library_th.el_id.PlaceHolder %>" value="<%= e_library_th.el_id.EditValue %>"<%= e_library_th.el_id.EditAttributes %>>
</span>
<%= e_library_th.el_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If e_library_th.el_date.Visible Then ' el_date %>
	<tr id="r_el_date">
		<td><span id="elh_e_library_th_el_date"><%= e_library_th.el_date.FldCaption %></span></td>
		<td<%= e_library_th.el_date.CellAttributes %>>
<span id="el_e_library_th_el_date" class="control-group">
<input type="text" data-field="x_el_date" name="x_el_date" id="x_el_date" placeholder="<%= e_library_th.el_date.PlaceHolder %>" value="<%= e_library_th.el_date.EditValue %>"<%= e_library_th.el_date.EditAttributes %>>
</span>
<%= e_library_th.el_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If e_library_th.el_title.Visible Then ' el_title %>
	<tr id="r_el_title">
		<td><span id="elh_e_library_th_el_title"><%= e_library_th.el_title.FldCaption %></span></td>
		<td<%= e_library_th.el_title.CellAttributes %>>
<span id="el_e_library_th_el_title" class="control-group">
<input type="text" data-field="x_el_title" name="x_el_title" id="x_el_title" size="30" maxlength="255" placeholder="<%= e_library_th.el_title.PlaceHolder %>" value="<%= e_library_th.el_title.EditValue %>"<%= e_library_th.el_title.EditAttributes %>>
</span>
<%= e_library_th.el_title.CustomMsg %></td>
	</tr>
<% End If %>
<% If e_library_th.el_pdf.Visible Then ' el_pdf %>
	<tr id="r_el_pdf">
		<td><span id="elh_e_library_th_el_pdf"><%= e_library_th.el_pdf.FldCaption %></span></td>
		<td<%= e_library_th.el_pdf.CellAttributes %>>
<span id="el_e_library_th_el_pdf" class="control-group">
<input type="text" data-field="x_el_pdf" name="x_el_pdf" id="x_el_pdf" size="30" maxlength="255" placeholder="<%= e_library_th.el_pdf.PlaceHolder %>" value="<%= e_library_th.el_pdf.EditValue %>"<%= e_library_th.el_pdf.EditAttributes %>>
</span>
<%= e_library_th.el_pdf.CustomMsg %></td>
	</tr>
<% End If %>
<% If e_library_th.el_img.Visible Then ' el_img %>
	<tr id="r_el_img">
		<td><span id="elh_e_library_th_el_img"><%= e_library_th.el_img.FldCaption %></span></td>
		<td<%= e_library_th.el_img.CellAttributes %>>
<span id="el_e_library_th_el_img" class="control-group">
<input type="text" data-field="x_el_img" name="x_el_img" id="x_el_img" size="30" maxlength="255" placeholder="<%= e_library_th.el_img.PlaceHolder %>" value="<%= e_library_th.el_img.EditValue %>"<%= e_library_th.el_img.EditAttributes %>>
</span>
<%= e_library_th.el_img.CustomMsg %></td>
	</tr>
<% End If %>
<% If e_library_th.el_detail.Visible Then ' el_detail %>
	<tr id="r_el_detail">
		<td><span id="elh_e_library_th_el_detail"><%= e_library_th.el_detail.FldCaption %></span></td>
		<td<%= e_library_th.el_detail.CellAttributes %>>
<span id="el_e_library_th_el_detail" class="control-group">
<textarea data-field="x_el_detail" name="x_el_detail" id="x_el_detail" cols="35" rows="4" placeholder="<%= e_library_th.el_detail.PlaceHolder %>"<%= e_library_th.el_detail.EditAttributes %>><%= e_library_th.el_detail.EditValue %></textarea>
</span>
<%= e_library_th.el_detail.CustomMsg %></td>
	</tr>
<% End If %>
<% If e_library_th.el_create.Visible Then ' el_create %>
	<tr id="r_el_create">
		<td><span id="elh_e_library_th_el_create"><%= e_library_th.el_create.FldCaption %></span></td>
		<td<%= e_library_th.el_create.CellAttributes %>>
<span id="el_e_library_th_el_create" class="control-group">
<input type="text" data-field="x_el_create" name="x_el_create" id="x_el_create" size="30" maxlength="255" placeholder="<%= e_library_th.el_create.PlaceHolder %>" value="<%= e_library_th.el_create.EditValue %>"<%= e_library_th.el_create.EditAttributes %>>
</span>
<%= e_library_th.el_create.CustomMsg %></td>
	</tr>
<% End If %>
<% If e_library_th.el_update.Visible Then ' el_update %>
	<tr id="r_el_update">
		<td><span id="elh_e_library_th_el_update"><%= e_library_th.el_update.FldCaption %></span></td>
		<td<%= e_library_th.el_update.CellAttributes %>>
<span id="el_e_library_th_el_update" class="control-group">
<input type="text" data-field="x_el_update" name="x_el_update" id="x_el_update" size="30" maxlength="255" placeholder="<%= e_library_th.el_update.PlaceHolder %>" value="<%= e_library_th.el_update.EditValue %>"<%= e_library_th.el_update.EditAttributes %>>
</span>
<%= e_library_th.el_update.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
fe_library_thadd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
e_library_th_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set e_library_th_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class ce_library_th_add

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
		TableName = "e_library_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "e_library_th_add"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If e_library_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & e_library_th.TableVar & "&" ' add page token
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

		' Initialize language object
		If IsEmpty(Language) Then
			Set Language = New cLanguage
			Call Language.LoadPhrases()
		End If

		' Initialize table object
		If IsEmpty(e_library_th) Then Set e_library_th = New ce_library_th
		Set Table = e_library_th

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "add"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "e_library_th"

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

		e_library_th.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set e_library_th = Nothing
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
			e_library_th.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("el_id").Count > 0 Then
				e_library_th.el_id.QueryStringValue = Request.QueryString("el_id")
				Call e_library_th.SetKey("el_id", e_library_th.el_id.CurrentValue) ' Set up key
			Else
				Call e_library_th.SetKey("el_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				e_library_th.CurrentAction = "C" ' Copy Record
			Else
				e_library_th.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				e_library_th.CurrentAction = "I" ' Form error, reset action
				e_library_th.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case e_library_th.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_e_library_thlist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				e_library_th.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = e_library_th.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_e_library_thview.asp" Then sReturnUrl = e_library_th.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					e_library_th.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		e_library_th.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call e_library_th.ResetAttrs()
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
		e_library_th.el_id.CurrentValue = Null
		e_library_th.el_id.OldValue = e_library_th.el_id.CurrentValue
		e_library_th.el_date.CurrentValue = Null
		e_library_th.el_date.OldValue = e_library_th.el_date.CurrentValue
		e_library_th.el_title.CurrentValue = Null
		e_library_th.el_title.OldValue = e_library_th.el_title.CurrentValue
		e_library_th.el_pdf.CurrentValue = Null
		e_library_th.el_pdf.OldValue = e_library_th.el_pdf.CurrentValue
		e_library_th.el_img.CurrentValue = Null
		e_library_th.el_img.OldValue = e_library_th.el_img.CurrentValue
		e_library_th.el_detail.CurrentValue = Null
		e_library_th.el_detail.OldValue = e_library_th.el_detail.CurrentValue
		e_library_th.el_create.CurrentValue = Null
		e_library_th.el_create.OldValue = e_library_th.el_create.CurrentValue
		e_library_th.el_update.CurrentValue = Null
		e_library_th.el_update.OldValue = e_library_th.el_update.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not e_library_th.el_id.FldIsDetailKey Then e_library_th.el_id.FormValue = ObjForm.GetValue("x_el_id")
		If Not e_library_th.el_date.FldIsDetailKey Then e_library_th.el_date.FormValue = ObjForm.GetValue("x_el_date")
		If Not e_library_th.el_date.FldIsDetailKey Then e_library_th.el_date.CurrentValue = ew_UnFormatDateTime(e_library_th.el_date.CurrentValue, 8)
		If Not e_library_th.el_title.FldIsDetailKey Then e_library_th.el_title.FormValue = ObjForm.GetValue("x_el_title")
		If Not e_library_th.el_pdf.FldIsDetailKey Then e_library_th.el_pdf.FormValue = ObjForm.GetValue("x_el_pdf")
		If Not e_library_th.el_img.FldIsDetailKey Then e_library_th.el_img.FormValue = ObjForm.GetValue("x_el_img")
		If Not e_library_th.el_detail.FldIsDetailKey Then e_library_th.el_detail.FormValue = ObjForm.GetValue("x_el_detail")
		If Not e_library_th.el_create.FldIsDetailKey Then e_library_th.el_create.FormValue = ObjForm.GetValue("x_el_create")
		If Not e_library_th.el_update.FldIsDetailKey Then e_library_th.el_update.FormValue = ObjForm.GetValue("x_el_update")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		e_library_th.el_id.CurrentValue = e_library_th.el_id.FormValue
		e_library_th.el_date.CurrentValue = e_library_th.el_date.FormValue
		e_library_th.el_date.CurrentValue = ew_UnFormatDateTime(e_library_th.el_date.CurrentValue, 8)
		e_library_th.el_title.CurrentValue = e_library_th.el_title.FormValue
		e_library_th.el_pdf.CurrentValue = e_library_th.el_pdf.FormValue
		e_library_th.el_img.CurrentValue = e_library_th.el_img.FormValue
		e_library_th.el_detail.CurrentValue = e_library_th.el_detail.FormValue
		e_library_th.el_create.CurrentValue = e_library_th.el_create.FormValue
		e_library_th.el_update.CurrentValue = e_library_th.el_update.FormValue
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

			' el_detail
			e_library_th.el_detail.ViewValue = e_library_th.el_detail.CurrentValue
			e_library_th.el_detail.ViewCustomAttributes = ""

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

			' el_detail
			e_library_th.el_detail.LinkCustomAttributes = ""
			e_library_th.el_detail.HrefValue = ""
			e_library_th.el_detail.TooltipValue = ""

			' el_create
			e_library_th.el_create.LinkCustomAttributes = ""
			e_library_th.el_create.HrefValue = ""
			e_library_th.el_create.TooltipValue = ""

			' el_update
			e_library_th.el_update.LinkCustomAttributes = ""
			e_library_th.el_update.HrefValue = ""
			e_library_th.el_update.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf e_library_th.RowType = EW_ROWTYPE_ADD Then ' Add row

			' el_id
			e_library_th.el_id.EditCustomAttributes = ""
			e_library_th.el_id.EditValue = ew_HtmlEncode(e_library_th.el_id.CurrentValue)
			e_library_th.el_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(e_library_th.el_id.FldCaption))

			' el_date
			e_library_th.el_date.EditCustomAttributes = ""
			e_library_th.el_date.EditValue = ew_HtmlEncode(e_library_th.el_date.CurrentValue)
			e_library_th.el_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(e_library_th.el_date.FldCaption))

			' el_title
			e_library_th.el_title.EditCustomAttributes = ""
			e_library_th.el_title.EditValue = ew_HtmlEncode(e_library_th.el_title.CurrentValue)
			e_library_th.el_title.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(e_library_th.el_title.FldCaption))

			' el_pdf
			e_library_th.el_pdf.EditCustomAttributes = ""
			e_library_th.el_pdf.EditValue = ew_HtmlEncode(e_library_th.el_pdf.CurrentValue)
			e_library_th.el_pdf.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(e_library_th.el_pdf.FldCaption))

			' el_img
			e_library_th.el_img.EditCustomAttributes = ""
			e_library_th.el_img.EditValue = ew_HtmlEncode(e_library_th.el_img.CurrentValue)
			e_library_th.el_img.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(e_library_th.el_img.FldCaption))

			' el_detail
			e_library_th.el_detail.EditCustomAttributes = ""
			e_library_th.el_detail.EditValue = e_library_th.el_detail.CurrentValue
			e_library_th.el_detail.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(e_library_th.el_detail.FldCaption))

			' el_create
			e_library_th.el_create.EditCustomAttributes = ""
			e_library_th.el_create.EditValue = ew_HtmlEncode(e_library_th.el_create.CurrentValue)
			e_library_th.el_create.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(e_library_th.el_create.FldCaption))

			' el_update
			e_library_th.el_update.EditCustomAttributes = ""
			e_library_th.el_update.EditValue = ew_HtmlEncode(e_library_th.el_update.CurrentValue)
			e_library_th.el_update.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(e_library_th.el_update.FldCaption))

			' Edit refer script
			' el_id

			e_library_th.el_id.HrefValue = ""

			' el_date
			e_library_th.el_date.HrefValue = ""

			' el_title
			e_library_th.el_title.HrefValue = ""

			' el_pdf
			e_library_th.el_pdf.HrefValue = ""

			' el_img
			e_library_th.el_img.HrefValue = ""

			' el_detail
			e_library_th.el_detail.HrefValue = ""

			' el_create
			e_library_th.el_create.HrefValue = ""

			' el_update
			e_library_th.el_update.HrefValue = ""
		End If
		If e_library_th.RowType = EW_ROWTYPE_ADD Or e_library_th.RowType = EW_ROWTYPE_EDIT Or e_library_th.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call e_library_th.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If e_library_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call e_library_th.Row_Rendered()
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
		If Not ew_CheckInteger(e_library_th.el_id.FormValue) Then
			Call ew_AddMessage(gsFormError, e_library_th.el_id.FldErrMsg)
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
		If e_library_th.el_id.CurrentValue <> "" Then ' Check field with unique index
			sFilter = "([el_id] = " & ew_AdjustSql(e_library_th.el_id.CurrentValue) & ")"
			Set RsChk = e_library_th.LoadRs(sFilter)
			If Not (RsChk Is Nothing) Then
				sIdxErrMsg = Replace(Language.Phrase("DupIndex"), "%f", e_library_th.el_id.FldCaption)
				sIdxErrMsg = Replace(sIdxErrMsg, "%v", e_library_th.el_id.CurrentValue)
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
		e_library_th.CurrentFilter = sFilter
		sSql = e_library_th.SQL
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

		' Field el_id
		Call e_library_th.el_id.SetDbValue(Rs, e_library_th.el_id.CurrentValue, Null, False)

		' Field el_date
		Call e_library_th.el_date.SetDbValue(Rs, e_library_th.el_date.CurrentValue, Null, False)

		' Field el_title
		Call e_library_th.el_title.SetDbValue(Rs, e_library_th.el_title.CurrentValue, Null, False)

		' Field el_pdf
		Call e_library_th.el_pdf.SetDbValue(Rs, e_library_th.el_pdf.CurrentValue, Null, False)

		' Field el_img
		Call e_library_th.el_img.SetDbValue(Rs, e_library_th.el_img.CurrentValue, Null, False)

		' Field el_detail
		Call e_library_th.el_detail.SetDbValue(Rs, e_library_th.el_detail.CurrentValue, Null, False)

		' Field el_create
		Call e_library_th.el_create.SetDbValue(Rs, e_library_th.el_create.CurrentValue, Null, False)

		' Field el_update
		Call e_library_th.el_update.SetDbValue(Rs, e_library_th.el_update.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = e_library_th.Row_Inserting(RsOld, Rs)

		' Check if key value entered
		If bInsertRow And e_library_th.ValidateKey And e_library_th.el_id.CurrentValue = "" And e_library_th.el_id.SessionValue = "" Then
			FailureMessage = Language.Phrase("InvalidKeyValue")
			bInsertRow = False
		End If

		' Check for duplicate key
		Dim sKeyErrMsg
		If bInsertRow And e_library_th.ValidateKey Then
			sFilter = e_library_th.KeyFilter
			Set RsChk = e_library_th.LoadRs(sFilter)
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
			ElseIf e_library_th.CancelMessage <> "" Then
				FailureMessage = e_library_th.CancelMessage
				e_library_th.CancelMessage = ""
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
			Call e_library_th.Row_Inserted(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", e_library_th.TableVar, "pom_e_library_thlist.asp", e_library_th.TableVar, True)
		PageId = ew_IIf(e_library_th.CurrentAction = "C", "Copy", "Add")
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
