<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_officeinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim office_add
Set office_add = New coffice_add
Set Page = office_add

' Page init processing
office_add.Page_Init()

' Page main processing
office_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
office_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var office_add = new ew_Page("office_add");
office_add.PageID = "add"; // Page ID
var EW_PAGE_ID = office_add.PageID; // For backward compatibility
// Form object
var fofficeadd = new ew_Form("fofficeadd");
// Validate form
fofficeadd.Validate = function() {
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
fofficeadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fofficeadd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fofficeadd.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If office.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% office_add.ShowPageHeader() %>
<% office_add.ShowMessage %>
<form name="fofficeadd" id="fofficeadd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="office">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_officeadd" class="table table-bordered table-striped">
<% If office.office_name_en.Visible Then ' office_name_en %>
	<tr id="r_office_name_en">
		<td><span id="elh_office_office_name_en"><%= office.office_name_en.FldCaption %></span></td>
		<td<%= office.office_name_en.CellAttributes %>>
<span id="el_office_office_name_en" class="control-group">
<input type="text" data-field="x_office_name_en" name="x_office_name_en" id="x_office_name_en" size="30" maxlength="255" placeholder="<%= office.office_name_en.PlaceHolder %>" value="<%= office.office_name_en.EditValue %>"<%= office.office_name_en.EditAttributes %>>
</span>
<%= office.office_name_en.CustomMsg %></td>
	</tr>
<% End If %>
<% If office.office_name_th.Visible Then ' office_name_th %>
	<tr id="r_office_name_th">
		<td><span id="elh_office_office_name_th"><%= office.office_name_th.FldCaption %></span></td>
		<td<%= office.office_name_th.CellAttributes %>>
<span id="el_office_office_name_th" class="control-group">
<input type="text" data-field="x_office_name_th" name="x_office_name_th" id="x_office_name_th" size="30" maxlength="255" placeholder="<%= office.office_name_th.PlaceHolder %>" value="<%= office.office_name_th.EditValue %>"<%= office.office_name_th.EditAttributes %>>
</span>
<%= office.office_name_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If office.office_sort.Visible Then ' office_sort %>
	<tr id="r_office_sort">
		<td><span id="elh_office_office_sort"><%= office.office_sort.FldCaption %></span></td>
		<td<%= office.office_sort.CellAttributes %>>
<span id="el_office_office_sort" class="control-group">
<input type="text" data-field="x_office_sort" name="x_office_sort" id="x_office_sort" size="30" maxlength="255" placeholder="<%= office.office_sort.PlaceHolder %>" value="<%= office.office_sort.EditValue %>"<%= office.office_sort.EditAttributes %>>
</span>
<%= office.office_sort.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
fofficeadd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
office_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set office_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class coffice_add

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
		TableName = "office"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "office_add"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If office.UseTokenInUrl Then PageUrl = PageUrl & "t=" & office.TableVar & "&" ' add page token
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
		If office.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (office.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (office.TableVar = Request.QueryString("t"))
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
		If IsEmpty(office) Then Set office = New coffice
		Set Table = office

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "add"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "office"

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

		office.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set office = Nothing
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
			office.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("office_id").Count > 0 Then
				office.office_id.QueryStringValue = Request.QueryString("office_id")
				Call office.SetKey("office_id", office.office_id.CurrentValue) ' Set up key
			Else
				Call office.SetKey("office_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				office.CurrentAction = "C" ' Copy Record
			Else
				office.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				office.CurrentAction = "I" ' Form error, reset action
				office.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case office.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_officelist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				office.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = office.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_officeview.asp" Then sReturnUrl = office.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					office.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		office.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call office.ResetAttrs()
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
		office.office_name_en.CurrentValue = Null
		office.office_name_en.OldValue = office.office_name_en.CurrentValue
		office.office_name_th.CurrentValue = Null
		office.office_name_th.OldValue = office.office_name_th.CurrentValue
		office.office_sort.CurrentValue = Null
		office.office_sort.OldValue = office.office_sort.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not office.office_name_en.FldIsDetailKey Then office.office_name_en.FormValue = ObjForm.GetValue("x_office_name_en")
		If Not office.office_name_th.FldIsDetailKey Then office.office_name_th.FormValue = ObjForm.GetValue("x_office_name_th")
		If Not office.office_sort.FldIsDetailKey Then office.office_sort.FormValue = ObjForm.GetValue("x_office_sort")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		office.office_name_en.CurrentValue = office.office_name_en.FormValue
		office.office_name_th.CurrentValue = office.office_name_th.FormValue
		office.office_sort.CurrentValue = office.office_sort.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = office.KeyFilter

		' Call Row Selecting event
		Call office.Row_Selecting(sFilter)

		' Load sql based on filter
		office.CurrentFilter = sFilter
		sSql = office.SQL
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
		Call office.Row_Selected(RsRow)
		office.office_id.DbValue = RsRow("office_id")
		office.office_name_en.DbValue = RsRow("office_name_en")
		office.office_name_th.DbValue = RsRow("office_name_th")
		office.office_sort.DbValue = RsRow("office_sort")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		office.office_id.m_DbValue = Rs("office_id")
		office.office_name_en.m_DbValue = Rs("office_name_en")
		office.office_name_th.m_DbValue = Rs("office_name_th")
		office.office_sort.m_DbValue = Rs("office_sort")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If office.GetKey("office_id")&"" <> "" Then
			office.office_id.CurrentValue = office.GetKey("office_id") ' office_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			office.CurrentFilter = office.KeyFilter
			Dim sSql
			sSql = office.SQL
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

		Call office.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' office_id
		' office_name_en
		' office_name_th
		' office_sort
		' -----------
		'  View  Row
		' -----------

		If office.RowType = EW_ROWTYPE_VIEW Then ' View row

			' office_id
			office.office_id.ViewValue = office.office_id.CurrentValue
			office.office_id.ViewCustomAttributes = ""

			' office_name_en
			office.office_name_en.ViewValue = office.office_name_en.CurrentValue
			office.office_name_en.ViewCustomAttributes = ""

			' office_name_th
			office.office_name_th.ViewValue = office.office_name_th.CurrentValue
			office.office_name_th.ViewCustomAttributes = ""

			' office_sort
			office.office_sort.ViewValue = office.office_sort.CurrentValue
			office.office_sort.ViewCustomAttributes = ""

			' View refer script
			' office_name_en

			office.office_name_en.LinkCustomAttributes = ""
			office.office_name_en.HrefValue = ""
			office.office_name_en.TooltipValue = ""

			' office_name_th
			office.office_name_th.LinkCustomAttributes = ""
			office.office_name_th.HrefValue = ""
			office.office_name_th.TooltipValue = ""

			' office_sort
			office.office_sort.LinkCustomAttributes = ""
			office.office_sort.HrefValue = ""
			office.office_sort.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf office.RowType = EW_ROWTYPE_ADD Then ' Add row

			' office_name_en
			office.office_name_en.EditCustomAttributes = ""
			office.office_name_en.EditValue = ew_HtmlEncode(office.office_name_en.CurrentValue)
			office.office_name_en.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(office.office_name_en.FldCaption))

			' office_name_th
			office.office_name_th.EditCustomAttributes = ""
			office.office_name_th.EditValue = ew_HtmlEncode(office.office_name_th.CurrentValue)
			office.office_name_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(office.office_name_th.FldCaption))

			' office_sort
			office.office_sort.EditCustomAttributes = ""
			office.office_sort.EditValue = ew_HtmlEncode(office.office_sort.CurrentValue)
			office.office_sort.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(office.office_sort.FldCaption))

			' Edit refer script
			' office_name_en

			office.office_name_en.HrefValue = ""

			' office_name_th
			office.office_name_th.HrefValue = ""

			' office_sort
			office.office_sort.HrefValue = ""
		End If
		If office.RowType = EW_ROWTYPE_ADD Or office.RowType = EW_ROWTYPE_EDIT Or office.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call office.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If office.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call office.Row_Rendered()
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

		' Load db values from rsold
		If Not IsNull(RsOld) Then
			Call LoadDbValues(RsOld)
		End If

		' Add new record
		sFilter = "(0 = 1)"
		office.CurrentFilter = sFilter
		sSql = office.SQL
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

		' Field office_name_en
		Call office.office_name_en.SetDbValue(Rs, office.office_name_en.CurrentValue, Null, False)

		' Field office_name_th
		Call office.office_name_th.SetDbValue(Rs, office.office_name_th.CurrentValue, Null, False)

		' Field office_sort
		Call office.office_sort.SetDbValue(Rs, office.office_sort.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = office.Row_Inserting(RsOld, Rs)
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
			ElseIf office.CancelMessage <> "" Then
				FailureMessage = office.CancelMessage
				office.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("InsertCancelled")
			End If
			AddRow = False
		End If
		Rs.Close
		Set Rs = Nothing
		If AddRow Then
			office.office_id.DbValue = RsNew("office_id")
		End If
		If AddRow Then

			' Call Row Inserted event
			Call office.Row_Inserted(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", office.TableVar, "pom_officelist.asp", office.TableVar, True)
		PageId = ew_IIf(office.CurrentAction = "C", "Copy", "Add")
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
