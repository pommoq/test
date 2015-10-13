<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_e_libraryinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim e_library_edit
Set e_library_edit = New ce_library_edit
Set Page = e_library_edit

' Page init processing
e_library_edit.Page_Init()

' Page main processing
e_library_edit.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
e_library_edit.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var e_library_edit = new ew_Page("e_library_edit");
e_library_edit.PageID = "edit"; // Page ID
var EW_PAGE_ID = e_library_edit.PageID; // For backward compatibility
// Form object
var fe_libraryedit = new ew_Form("fe_libraryedit");
// Validate form
fe_libraryedit.Validate = function() {
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
				return this.OnError(elm, "<%= ew_JsEncode2(e_library.el_id.FldErrMsg) %>");
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
fe_libraryedit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fe_libraryedit.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fe_libraryedit.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If e_library.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% e_library_edit.ShowPageHeader() %>
<% e_library_edit.ShowMessage %>
<form name="fe_libraryedit" id="fe_libraryedit" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="a_table" id="a_table" value="e_library">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table class="ewGrid"><tr><td>
<table id="tbl_e_libraryedit" class="table table-bordered table-striped">
<% If e_library.el_id.Visible Then ' el_id %>
	<tr id="r_el_id">
		<td><span id="elh_e_library_el_id"><%= e_library.el_id.FldCaption %></span></td>
		<td<%= e_library.el_id.CellAttributes %>>
<span id="el_e_library_el_id" class="control-group">
<span<%= e_library.el_id.ViewAttributes %>>
<%= e_library.el_id.EditValue %>
</span>
</span>
<input type="hidden" data-field="x_el_id" name="x_el_id" id="x_el_id" value="<%= Server.HTMLEncode(e_library.el_id.CurrentValue&"") %>">
<%= e_library.el_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If e_library.el_date.Visible Then ' el_date %>
	<tr id="r_el_date">
		<td><span id="elh_e_library_el_date"><%= e_library.el_date.FldCaption %></span></td>
		<td<%= e_library.el_date.CellAttributes %>>
<span id="el_e_library_el_date" class="control-group">
<input type="text" data-field="x_el_date" name="x_el_date" id="x_el_date" placeholder="<%= e_library.el_date.PlaceHolder %>" value="<%= e_library.el_date.EditValue %>"<%= e_library.el_date.EditAttributes %>>
</span>
<%= e_library.el_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If e_library.el_title.Visible Then ' el_title %>
	<tr id="r_el_title">
		<td><span id="elh_e_library_el_title"><%= e_library.el_title.FldCaption %></span></td>
		<td<%= e_library.el_title.CellAttributes %>>
<span id="el_e_library_el_title" class="control-group">
<input type="text" data-field="x_el_title" name="x_el_title" id="x_el_title" size="30" maxlength="255" placeholder="<%= e_library.el_title.PlaceHolder %>" value="<%= e_library.el_title.EditValue %>"<%= e_library.el_title.EditAttributes %>>
</span>
<%= e_library.el_title.CustomMsg %></td>
	</tr>
<% End If %>
<% If e_library.el_pdf.Visible Then ' el_pdf %>
	<tr id="r_el_pdf">
		<td><span id="elh_e_library_el_pdf"><%= e_library.el_pdf.FldCaption %></span></td>
		<td<%= e_library.el_pdf.CellAttributes %>>
<span id="el_e_library_el_pdf" class="control-group">
<input type="text" data-field="x_el_pdf" name="x_el_pdf" id="x_el_pdf" size="30" maxlength="255" placeholder="<%= e_library.el_pdf.PlaceHolder %>" value="<%= e_library.el_pdf.EditValue %>"<%= e_library.el_pdf.EditAttributes %>>
</span>
<%= e_library.el_pdf.CustomMsg %></td>
	</tr>
<% End If %>
<% If e_library.el_img.Visible Then ' el_img %>
	<tr id="r_el_img">
		<td><span id="elh_e_library_el_img"><%= e_library.el_img.FldCaption %></span></td>
		<td<%= e_library.el_img.CellAttributes %>>
<span id="el_e_library_el_img" class="control-group">
<input type="text" data-field="x_el_img" name="x_el_img" id="x_el_img" size="30" maxlength="255" placeholder="<%= e_library.el_img.PlaceHolder %>" value="<%= e_library.el_img.EditValue %>"<%= e_library.el_img.EditAttributes %>>
</span>
<%= e_library.el_img.CustomMsg %></td>
	</tr>
<% End If %>
<% If e_library.el_detail.Visible Then ' el_detail %>
	<tr id="r_el_detail">
		<td><span id="elh_e_library_el_detail"><%= e_library.el_detail.FldCaption %></span></td>
		<td<%= e_library.el_detail.CellAttributes %>>
<span id="el_e_library_el_detail" class="control-group">
<textarea data-field="x_el_detail" name="x_el_detail" id="x_el_detail" cols="35" rows="4" placeholder="<%= e_library.el_detail.PlaceHolder %>"<%= e_library.el_detail.EditAttributes %>><%= e_library.el_detail.EditValue %></textarea>
</span>
<%= e_library.el_detail.CustomMsg %></td>
	</tr>
<% End If %>
<% If e_library.el_create.Visible Then ' el_create %>
	<tr id="r_el_create">
		<td><span id="elh_e_library_el_create"><%= e_library.el_create.FldCaption %></span></td>
		<td<%= e_library.el_create.CellAttributes %>>
<span id="el_e_library_el_create" class="control-group">
<input type="text" data-field="x_el_create" name="x_el_create" id="x_el_create" size="30" maxlength="255" placeholder="<%= e_library.el_create.PlaceHolder %>" value="<%= e_library.el_create.EditValue %>"<%= e_library.el_create.EditAttributes %>>
</span>
<%= e_library.el_create.CustomMsg %></td>
	</tr>
<% End If %>
<% If e_library.el_update.Visible Then ' el_update %>
	<tr id="r_el_update">
		<td><span id="elh_e_library_el_update"><%= e_library.el_update.FldCaption %></span></td>
		<td<%= e_library.el_update.CellAttributes %>>
<span id="el_e_library_el_update" class="control-group">
<input type="text" data-field="x_el_update" name="x_el_update" id="x_el_update" size="30" maxlength="255" placeholder="<%= e_library.el_update.PlaceHolder %>" value="<%= e_library.el_update.EditValue %>"<%= e_library.el_update.EditAttributes %>>
</span>
<%= e_library.el_update.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("EditBtn") %></button>
</form>
<script type="text/javascript">
fe_libraryedit.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
e_library_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set e_library_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class ce_library_edit

	' Page ID
	Public Property Get PageID()
		PageID = "edit"
	End Property

	' Project ID
	Public Property Get ProjectID()
		ProjectID = "{324ED72D-DE20-46F7-B12E-7AF8CE8711A6}"
	End Property

	' Table Name
	Public Property Get TableName()
		TableName = "e_library"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "e_library_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If e_library.UseTokenInUrl Then PageUrl = PageUrl & "t=" & e_library.TableVar & "&" ' add page token
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
		If e_library.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (e_library.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (e_library.TableVar = Request.QueryString("t"))
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
		If IsEmpty(e_library) Then Set e_library = New ce_library
		Set Table = e_library

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "e_library"

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

		e_library.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set e_library = Nothing
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

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		Dim sReturnUrl
		sReturnUrl = ""

		' Load key from QueryString
		If Request.QueryString("el_id").Count > 0 Then
			e_library.el_id.QueryStringValue = Request.QueryString("el_id")
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Process form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			e_library.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values
		Else
			e_library.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If e_library.el_id.CurrentValue = "" Then Call Page_Terminate("pom_e_librarylist.asp") ' Invalid key, return to list

		' Validate form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			If Not ValidateForm() Then
				e_library.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				e_library.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		End If
		Select Case e_library.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_e_librarylist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				e_library.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					sReturnUrl = e_library.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					e_library.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		e_library.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call e_library.ResetAttrs()
		Call RenderRow()
	End Sub
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
				e_library.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					e_library.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = e_library.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			e_library.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			e_library.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			e_library.StartRecordNumber = StartRec
		End If
	End Sub

	' -----------------------------------------------------------------
	' Function Get upload files
	'
	Function GetUploadFiles()

		' Get upload data
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not e_library.el_id.FldIsDetailKey Then e_library.el_id.FormValue = ObjForm.GetValue("x_el_id")
		If Not e_library.el_date.FldIsDetailKey Then e_library.el_date.FormValue = ObjForm.GetValue("x_el_date")
		If Not e_library.el_date.FldIsDetailKey Then e_library.el_date.CurrentValue = ew_UnFormatDateTime(e_library.el_date.CurrentValue, 8)
		If Not e_library.el_title.FldIsDetailKey Then e_library.el_title.FormValue = ObjForm.GetValue("x_el_title")
		If Not e_library.el_pdf.FldIsDetailKey Then e_library.el_pdf.FormValue = ObjForm.GetValue("x_el_pdf")
		If Not e_library.el_img.FldIsDetailKey Then e_library.el_img.FormValue = ObjForm.GetValue("x_el_img")
		If Not e_library.el_detail.FldIsDetailKey Then e_library.el_detail.FormValue = ObjForm.GetValue("x_el_detail")
		If Not e_library.el_create.FldIsDetailKey Then e_library.el_create.FormValue = ObjForm.GetValue("x_el_create")
		If Not e_library.el_update.FldIsDetailKey Then e_library.el_update.FormValue = ObjForm.GetValue("x_el_update")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		e_library.el_id.CurrentValue = e_library.el_id.FormValue
		e_library.el_date.CurrentValue = e_library.el_date.FormValue
		e_library.el_date.CurrentValue = ew_UnFormatDateTime(e_library.el_date.CurrentValue, 8)
		e_library.el_title.CurrentValue = e_library.el_title.FormValue
		e_library.el_pdf.CurrentValue = e_library.el_pdf.FormValue
		e_library.el_img.CurrentValue = e_library.el_img.FormValue
		e_library.el_detail.CurrentValue = e_library.el_detail.FormValue
		e_library.el_create.CurrentValue = e_library.el_create.FormValue
		e_library.el_update.CurrentValue = e_library.el_update.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = e_library.KeyFilter

		' Call Row Selecting event
		Call e_library.Row_Selecting(sFilter)

		' Load sql based on filter
		e_library.CurrentFilter = sFilter
		sSql = e_library.SQL
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
		Call e_library.Row_Selected(RsRow)
		e_library.el_id.DbValue = RsRow("el_id")
		e_library.el_date.DbValue = RsRow("el_date")
		e_library.el_title.DbValue = RsRow("el_title")
		e_library.el_pdf.DbValue = RsRow("el_pdf")
		e_library.el_img.DbValue = RsRow("el_img")
		e_library.el_detail.DbValue = RsRow("el_detail")
		e_library.el_create.DbValue = RsRow("el_create")
		e_library.el_update.DbValue = RsRow("el_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		e_library.el_id.m_DbValue = Rs("el_id")
		e_library.el_date.m_DbValue = Rs("el_date")
		e_library.el_title.m_DbValue = Rs("el_title")
		e_library.el_pdf.m_DbValue = Rs("el_pdf")
		e_library.el_img.m_DbValue = Rs("el_img")
		e_library.el_detail.m_DbValue = Rs("el_detail")
		e_library.el_create.m_DbValue = Rs("el_create")
		e_library.el_update.m_DbValue = Rs("el_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call e_library.Row_Rendering()

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

		If e_library.RowType = EW_ROWTYPE_VIEW Then ' View row

			' el_id
			e_library.el_id.ViewValue = e_library.el_id.CurrentValue
			e_library.el_id.ViewCustomAttributes = ""

			' el_date
			e_library.el_date.ViewValue = e_library.el_date.CurrentValue
			e_library.el_date.ViewCustomAttributes = ""

			' el_title
			e_library.el_title.ViewValue = e_library.el_title.CurrentValue
			e_library.el_title.ViewCustomAttributes = ""

			' el_pdf
			e_library.el_pdf.ViewValue = e_library.el_pdf.CurrentValue
			e_library.el_pdf.ViewCustomAttributes = ""

			' el_img
			e_library.el_img.ViewValue = e_library.el_img.CurrentValue
			e_library.el_img.ViewCustomAttributes = ""

			' el_detail
			e_library.el_detail.ViewValue = e_library.el_detail.CurrentValue
			e_library.el_detail.ViewCustomAttributes = ""

			' el_create
			e_library.el_create.ViewValue = e_library.el_create.CurrentValue
			e_library.el_create.ViewCustomAttributes = ""

			' el_update
			e_library.el_update.ViewValue = e_library.el_update.CurrentValue
			e_library.el_update.ViewCustomAttributes = ""

			' View refer script
			' el_id

			e_library.el_id.LinkCustomAttributes = ""
			e_library.el_id.HrefValue = ""
			e_library.el_id.TooltipValue = ""

			' el_date
			e_library.el_date.LinkCustomAttributes = ""
			e_library.el_date.HrefValue = ""
			e_library.el_date.TooltipValue = ""

			' el_title
			e_library.el_title.LinkCustomAttributes = ""
			e_library.el_title.HrefValue = ""
			e_library.el_title.TooltipValue = ""

			' el_pdf
			e_library.el_pdf.LinkCustomAttributes = ""
			e_library.el_pdf.HrefValue = ""
			e_library.el_pdf.TooltipValue = ""

			' el_img
			e_library.el_img.LinkCustomAttributes = ""
			e_library.el_img.HrefValue = ""
			e_library.el_img.TooltipValue = ""

			' el_detail
			e_library.el_detail.LinkCustomAttributes = ""
			e_library.el_detail.HrefValue = ""
			e_library.el_detail.TooltipValue = ""

			' el_create
			e_library.el_create.LinkCustomAttributes = ""
			e_library.el_create.HrefValue = ""
			e_library.el_create.TooltipValue = ""

			' el_update
			e_library.el_update.LinkCustomAttributes = ""
			e_library.el_update.HrefValue = ""
			e_library.el_update.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf e_library.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' el_id
			e_library.el_id.EditCustomAttributes = ""
			e_library.el_id.EditValue = e_library.el_id.CurrentValue
			e_library.el_id.ViewCustomAttributes = ""

			' el_date
			e_library.el_date.EditCustomAttributes = ""
			e_library.el_date.EditValue = ew_HtmlEncode(e_library.el_date.CurrentValue)
			e_library.el_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(e_library.el_date.FldCaption))

			' el_title
			e_library.el_title.EditCustomAttributes = ""
			e_library.el_title.EditValue = ew_HtmlEncode(e_library.el_title.CurrentValue)
			e_library.el_title.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(e_library.el_title.FldCaption))

			' el_pdf
			e_library.el_pdf.EditCustomAttributes = ""
			e_library.el_pdf.EditValue = ew_HtmlEncode(e_library.el_pdf.CurrentValue)
			e_library.el_pdf.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(e_library.el_pdf.FldCaption))

			' el_img
			e_library.el_img.EditCustomAttributes = ""
			e_library.el_img.EditValue = ew_HtmlEncode(e_library.el_img.CurrentValue)
			e_library.el_img.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(e_library.el_img.FldCaption))

			' el_detail
			e_library.el_detail.EditCustomAttributes = ""
			e_library.el_detail.EditValue = e_library.el_detail.CurrentValue
			e_library.el_detail.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(e_library.el_detail.FldCaption))

			' el_create
			e_library.el_create.EditCustomAttributes = ""
			e_library.el_create.EditValue = ew_HtmlEncode(e_library.el_create.CurrentValue)
			e_library.el_create.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(e_library.el_create.FldCaption))

			' el_update
			e_library.el_update.EditCustomAttributes = ""
			e_library.el_update.EditValue = ew_HtmlEncode(e_library.el_update.CurrentValue)
			e_library.el_update.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(e_library.el_update.FldCaption))

			' Edit refer script
			' el_id

			e_library.el_id.HrefValue = ""

			' el_date
			e_library.el_date.HrefValue = ""

			' el_title
			e_library.el_title.HrefValue = ""

			' el_pdf
			e_library.el_pdf.HrefValue = ""

			' el_img
			e_library.el_img.HrefValue = ""

			' el_detail
			e_library.el_detail.HrefValue = ""

			' el_create
			e_library.el_create.HrefValue = ""

			' el_update
			e_library.el_update.HrefValue = ""
		End If
		If e_library.RowType = EW_ROWTYPE_ADD Or e_library.RowType = EW_ROWTYPE_EDIT Or e_library.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call e_library.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If e_library.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call e_library.Row_Rendered()
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
		If Not ew_CheckInteger(e_library.el_id.FormValue) Then
			Call ew_AddMessage(gsFormError, e_library.el_id.FldErrMsg)
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
	' Update record based on key values
	'
	Function EditRow()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim Rs, sSql, sFilter
		Dim RsChk, sSqlChk, sFilterChk
		Dim bUpdateRow
		Dim RsOld, RsNew
		Dim sIdxErrMsg

		' Clear any previous errors
		Err.Clear
		sFilter = e_library.KeyFilter
		e_library.CurrentFilter  = sFilter
		sSql = e_library.SQL
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.CursorLocation = EW_CURSORLOCATION
		Rs.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		If Err.Number <> 0 Then
			Message = Err.Description
			Rs.Close
			Set Rs = Nothing
			EditRow = False
			Exit Function
		End If

		' Clone old recordset object
		Set RsOld = ew_CloneRs(Rs)
		Call LoadDbValues(RsOld)
		If Rs.Eof Then
			EditRow = False ' Update Failed
		Else

			' Field el_id
			' Field el_date

			Call e_library.el_date.SetDbValue(Rs, e_library.el_date.CurrentValue, Null, e_library.el_date.ReadOnly)

			' Field el_title
			Call e_library.el_title.SetDbValue(Rs, e_library.el_title.CurrentValue, Null, e_library.el_title.ReadOnly)

			' Field el_pdf
			Call e_library.el_pdf.SetDbValue(Rs, e_library.el_pdf.CurrentValue, Null, e_library.el_pdf.ReadOnly)

			' Field el_img
			Call e_library.el_img.SetDbValue(Rs, e_library.el_img.CurrentValue, Null, e_library.el_img.ReadOnly)

			' Field el_detail
			Call e_library.el_detail.SetDbValue(Rs, e_library.el_detail.CurrentValue, Null, e_library.el_detail.ReadOnly)

			' Field el_create
			Call e_library.el_create.SetDbValue(Rs, e_library.el_create.CurrentValue, Null, e_library.el_create.ReadOnly)

			' Field el_update
			Call e_library.el_update.SetDbValue(Rs, e_library.el_update.CurrentValue, Null, e_library.el_update.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = e_library.Row_Updating(RsOld, Rs)
			If bUpdateRow Then

				' Clone new recordset object
				Set RsNew = ew_CloneRs(Rs)
				EditRow = True
				If EditRow Then
					Rs.Update
				End If
				If Err.Number <> 0 Or Not EditRow Then
					If Err.Description <> "" Then FailureMessage = Err.Description
					EditRow = False
				Else
					EditRow = True
				End If
				If EditRow Then
				End If
			Else
				Rs.CancelUpdate

				' Set up error message
				If SuccessMessage <> "" Or FailureMessage <> "" Then

					' Use the message, do nothing
				ElseIf e_library.CancelMessage <> "" Then
					FailureMessage = e_library.CancelMessage
					e_library.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call e_library.Row_Updated(RsOld, RsNew)
		End If
		Rs.Close
		Set Rs = Nothing
		If IsObject(RsOld) Then
			RsOld.Close
			Set RsOld = Nothing
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
		Call Breadcrumb.Add("list", e_library.TableVar, "pom_e_librarylist.asp", e_library.TableVar, True)
		PageId = "edit"
		Call Breadcrumb.Add("edit", PageId, ew_CurrentUrl, "", False)
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
