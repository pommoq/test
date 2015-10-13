<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_research_pdf_file_thinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim research_pdf_file_th_edit
Set research_pdf_file_th_edit = New cresearch_pdf_file_th_edit
Set Page = research_pdf_file_th_edit

' Page init processing
research_pdf_file_th_edit.Page_Init()

' Page main processing
research_pdf_file_th_edit.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
research_pdf_file_th_edit.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var research_pdf_file_th_edit = new ew_Page("research_pdf_file_th_edit");
research_pdf_file_th_edit.PageID = "edit"; // Page ID
var EW_PAGE_ID = research_pdf_file_th_edit.PageID; // For backward compatibility
// Form object
var fresearch_pdf_file_thedit = new ew_Form("fresearch_pdf_file_thedit");
// Validate form
fresearch_pdf_file_thedit.Validate = function() {
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
			elm = this.GetElements("x" + infix + "_rsh_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(research_pdf_file_th.rsh_id.FldErrMsg) %>");
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
fresearch_pdf_file_thedit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fresearch_pdf_file_thedit.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fresearch_pdf_file_thedit.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If research_pdf_file_th.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% research_pdf_file_th_edit.ShowPageHeader() %>
<% research_pdf_file_th_edit.ShowMessage %>
<form name="fresearch_pdf_file_thedit" id="fresearch_pdf_file_thedit" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="a_table" id="a_table" value="research_pdf_file_th">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table class="ewGrid"><tr><td>
<table id="tbl_research_pdf_file_thedit" class="table table-bordered table-striped">
<% If research_pdf_file_th.rsh_pdf_id.Visible Then ' rsh_pdf_id %>
	<tr id="r_rsh_pdf_id">
		<td><span id="elh_research_pdf_file_th_rsh_pdf_id"><%= research_pdf_file_th.rsh_pdf_id.FldCaption %></span></td>
		<td<%= research_pdf_file_th.rsh_pdf_id.CellAttributes %>>
<span id="el_research_pdf_file_th_rsh_pdf_id" class="control-group">
<span<%= research_pdf_file_th.rsh_pdf_id.ViewAttributes %>>
<%= research_pdf_file_th.rsh_pdf_id.EditValue %>
</span>
</span>
<input type="hidden" data-field="x_rsh_pdf_id" name="x_rsh_pdf_id" id="x_rsh_pdf_id" value="<%= Server.HTMLEncode(research_pdf_file_th.rsh_pdf_id.CurrentValue&"") %>">
<%= research_pdf_file_th.rsh_pdf_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_pdf_file_th.rsh_id.Visible Then ' rsh_id %>
	<tr id="r_rsh_id">
		<td><span id="elh_research_pdf_file_th_rsh_id"><%= research_pdf_file_th.rsh_id.FldCaption %></span></td>
		<td<%= research_pdf_file_th.rsh_id.CellAttributes %>>
<span id="el_research_pdf_file_th_rsh_id" class="control-group">
<input type="text" data-field="x_rsh_id" name="x_rsh_id" id="x_rsh_id" size="30" placeholder="<%= research_pdf_file_th.rsh_id.PlaceHolder %>" value="<%= research_pdf_file_th.rsh_id.EditValue %>"<%= research_pdf_file_th.rsh_id.EditAttributes %>>
</span>
<%= research_pdf_file_th.rsh_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_pdf_file_th.rsh_pdf_file.Visible Then ' rsh_pdf_file %>
	<tr id="r_rsh_pdf_file">
		<td><span id="elh_research_pdf_file_th_rsh_pdf_file"><%= research_pdf_file_th.rsh_pdf_file.FldCaption %></span></td>
		<td<%= research_pdf_file_th.rsh_pdf_file.CellAttributes %>>
<span id="el_research_pdf_file_th_rsh_pdf_file" class="control-group">
<input type="text" data-field="x_rsh_pdf_file" name="x_rsh_pdf_file" id="x_rsh_pdf_file" size="30" maxlength="255" placeholder="<%= research_pdf_file_th.rsh_pdf_file.PlaceHolder %>" value="<%= research_pdf_file_th.rsh_pdf_file.EditValue %>"<%= research_pdf_file_th.rsh_pdf_file.EditAttributes %>>
</span>
<%= research_pdf_file_th.rsh_pdf_file.CustomMsg %></td>
	</tr>
<% End If %>
<% If research_pdf_file_th.rsh_pdf_title.Visible Then ' rsh_pdf_title %>
	<tr id="r_rsh_pdf_title">
		<td><span id="elh_research_pdf_file_th_rsh_pdf_title"><%= research_pdf_file_th.rsh_pdf_title.FldCaption %></span></td>
		<td<%= research_pdf_file_th.rsh_pdf_title.CellAttributes %>>
<span id="el_research_pdf_file_th_rsh_pdf_title" class="control-group">
<input type="text" data-field="x_rsh_pdf_title" name="x_rsh_pdf_title" id="x_rsh_pdf_title" size="30" maxlength="255" placeholder="<%= research_pdf_file_th.rsh_pdf_title.PlaceHolder %>" value="<%= research_pdf_file_th.rsh_pdf_title.EditValue %>"<%= research_pdf_file_th.rsh_pdf_title.EditAttributes %>>
</span>
<%= research_pdf_file_th.rsh_pdf_title.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("EditBtn") %></button>
</form>
<script type="text/javascript">
fresearch_pdf_file_thedit.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
research_pdf_file_th_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set research_pdf_file_th_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cresearch_pdf_file_th_edit

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
		TableName = "research_pdf_file_th"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "research_pdf_file_th_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If research_pdf_file_th.UseTokenInUrl Then PageUrl = PageUrl & "t=" & research_pdf_file_th.TableVar & "&" ' add page token
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
		If research_pdf_file_th.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (research_pdf_file_th.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (research_pdf_file_th.TableVar = Request.QueryString("t"))
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
		If IsEmpty(research_pdf_file_th) Then Set research_pdf_file_th = New cresearch_pdf_file_th
		Set Table = research_pdf_file_th

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "research_pdf_file_th"

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

		research_pdf_file_th.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action
		research_pdf_file_th.rsh_pdf_id.Visible = Not research_pdf_file_th.IsAdd() And Not research_pdf_file_th.IsCopy() And Not research_pdf_file_th.IsGridAdd()

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
		Set research_pdf_file_th = Nothing
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
		If Request.QueryString("rsh_pdf_id").Count > 0 Then
			research_pdf_file_th.rsh_pdf_id.QueryStringValue = Request.QueryString("rsh_pdf_id")
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Process form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			research_pdf_file_th.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values
		Else
			research_pdf_file_th.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If research_pdf_file_th.rsh_pdf_id.CurrentValue = "" Then Call Page_Terminate("pom_research_pdf_file_thlist.asp") ' Invalid key, return to list

		' Validate form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			If Not ValidateForm() Then
				research_pdf_file_th.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				research_pdf_file_th.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		End If
		Select Case research_pdf_file_th.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_research_pdf_file_thlist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				research_pdf_file_th.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					sReturnUrl = research_pdf_file_th.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					research_pdf_file_th.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		research_pdf_file_th.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call research_pdf_file_th.ResetAttrs()
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
				research_pdf_file_th.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					research_pdf_file_th.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = research_pdf_file_th.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			research_pdf_file_th.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			research_pdf_file_th.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			research_pdf_file_th.StartRecordNumber = StartRec
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
		If Not research_pdf_file_th.rsh_pdf_id.FldIsDetailKey Then research_pdf_file_th.rsh_pdf_id.FormValue = ObjForm.GetValue("x_rsh_pdf_id")
		If Not research_pdf_file_th.rsh_id.FldIsDetailKey Then research_pdf_file_th.rsh_id.FormValue = ObjForm.GetValue("x_rsh_id")
		If Not research_pdf_file_th.rsh_pdf_file.FldIsDetailKey Then research_pdf_file_th.rsh_pdf_file.FormValue = ObjForm.GetValue("x_rsh_pdf_file")
		If Not research_pdf_file_th.rsh_pdf_title.FldIsDetailKey Then research_pdf_file_th.rsh_pdf_title.FormValue = ObjForm.GetValue("x_rsh_pdf_title")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		research_pdf_file_th.rsh_pdf_id.CurrentValue = research_pdf_file_th.rsh_pdf_id.FormValue
		research_pdf_file_th.rsh_id.CurrentValue = research_pdf_file_th.rsh_id.FormValue
		research_pdf_file_th.rsh_pdf_file.CurrentValue = research_pdf_file_th.rsh_pdf_file.FormValue
		research_pdf_file_th.rsh_pdf_title.CurrentValue = research_pdf_file_th.rsh_pdf_title.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = research_pdf_file_th.KeyFilter

		' Call Row Selecting event
		Call research_pdf_file_th.Row_Selecting(sFilter)

		' Load sql based on filter
		research_pdf_file_th.CurrentFilter = sFilter
		sSql = research_pdf_file_th.SQL
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
		Call research_pdf_file_th.Row_Selected(RsRow)
		research_pdf_file_th.rsh_pdf_id.DbValue = RsRow("rsh_pdf_id")
		research_pdf_file_th.rsh_id.DbValue = RsRow("rsh_id")
		research_pdf_file_th.rsh_pdf_file.DbValue = RsRow("rsh_pdf_file")
		research_pdf_file_th.rsh_pdf_title.DbValue = RsRow("rsh_pdf_title")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		research_pdf_file_th.rsh_pdf_id.m_DbValue = Rs("rsh_pdf_id")
		research_pdf_file_th.rsh_id.m_DbValue = Rs("rsh_id")
		research_pdf_file_th.rsh_pdf_file.m_DbValue = Rs("rsh_pdf_file")
		research_pdf_file_th.rsh_pdf_title.m_DbValue = Rs("rsh_pdf_title")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call research_pdf_file_th.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' rsh_pdf_id
		' rsh_id
		' rsh_pdf_file
		' rsh_pdf_title
		' -----------
		'  View  Row
		' -----------

		If research_pdf_file_th.RowType = EW_ROWTYPE_VIEW Then ' View row

			' rsh_pdf_id
			research_pdf_file_th.rsh_pdf_id.ViewValue = research_pdf_file_th.rsh_pdf_id.CurrentValue
			research_pdf_file_th.rsh_pdf_id.ViewCustomAttributes = ""

			' rsh_id
			research_pdf_file_th.rsh_id.ViewValue = research_pdf_file_th.rsh_id.CurrentValue
			research_pdf_file_th.rsh_id.ViewCustomAttributes = ""

			' rsh_pdf_file
			research_pdf_file_th.rsh_pdf_file.ViewValue = research_pdf_file_th.rsh_pdf_file.CurrentValue
			research_pdf_file_th.rsh_pdf_file.ViewCustomAttributes = ""

			' rsh_pdf_title
			research_pdf_file_th.rsh_pdf_title.ViewValue = research_pdf_file_th.rsh_pdf_title.CurrentValue
			research_pdf_file_th.rsh_pdf_title.ViewCustomAttributes = ""

			' View refer script
			' rsh_pdf_id

			research_pdf_file_th.rsh_pdf_id.LinkCustomAttributes = ""
			research_pdf_file_th.rsh_pdf_id.HrefValue = ""
			research_pdf_file_th.rsh_pdf_id.TooltipValue = ""

			' rsh_id
			research_pdf_file_th.rsh_id.LinkCustomAttributes = ""
			research_pdf_file_th.rsh_id.HrefValue = ""
			research_pdf_file_th.rsh_id.TooltipValue = ""

			' rsh_pdf_file
			research_pdf_file_th.rsh_pdf_file.LinkCustomAttributes = ""
			research_pdf_file_th.rsh_pdf_file.HrefValue = ""
			research_pdf_file_th.rsh_pdf_file.TooltipValue = ""

			' rsh_pdf_title
			research_pdf_file_th.rsh_pdf_title.LinkCustomAttributes = ""
			research_pdf_file_th.rsh_pdf_title.HrefValue = ""
			research_pdf_file_th.rsh_pdf_title.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf research_pdf_file_th.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' rsh_pdf_id
			research_pdf_file_th.rsh_pdf_id.EditCustomAttributes = ""
			research_pdf_file_th.rsh_pdf_id.EditValue = research_pdf_file_th.rsh_pdf_id.CurrentValue
			research_pdf_file_th.rsh_pdf_id.ViewCustomAttributes = ""

			' rsh_id
			research_pdf_file_th.rsh_id.EditCustomAttributes = ""
			research_pdf_file_th.rsh_id.EditValue = ew_HtmlEncode(research_pdf_file_th.rsh_id.CurrentValue)
			research_pdf_file_th.rsh_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_pdf_file_th.rsh_id.FldCaption))

			' rsh_pdf_file
			research_pdf_file_th.rsh_pdf_file.EditCustomAttributes = ""
			research_pdf_file_th.rsh_pdf_file.EditValue = ew_HtmlEncode(research_pdf_file_th.rsh_pdf_file.CurrentValue)
			research_pdf_file_th.rsh_pdf_file.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_pdf_file_th.rsh_pdf_file.FldCaption))

			' rsh_pdf_title
			research_pdf_file_th.rsh_pdf_title.EditCustomAttributes = ""
			research_pdf_file_th.rsh_pdf_title.EditValue = ew_HtmlEncode(research_pdf_file_th.rsh_pdf_title.CurrentValue)
			research_pdf_file_th.rsh_pdf_title.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(research_pdf_file_th.rsh_pdf_title.FldCaption))

			' Edit refer script
			' rsh_pdf_id

			research_pdf_file_th.rsh_pdf_id.HrefValue = ""

			' rsh_id
			research_pdf_file_th.rsh_id.HrefValue = ""

			' rsh_pdf_file
			research_pdf_file_th.rsh_pdf_file.HrefValue = ""

			' rsh_pdf_title
			research_pdf_file_th.rsh_pdf_title.HrefValue = ""
		End If
		If research_pdf_file_th.RowType = EW_ROWTYPE_ADD Or research_pdf_file_th.RowType = EW_ROWTYPE_EDIT Or research_pdf_file_th.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call research_pdf_file_th.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If research_pdf_file_th.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call research_pdf_file_th.Row_Rendered()
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
		If Not ew_CheckInteger(research_pdf_file_th.rsh_id.FormValue) Then
			Call ew_AddMessage(gsFormError, research_pdf_file_th.rsh_id.FldErrMsg)
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
		sFilter = research_pdf_file_th.KeyFilter
		research_pdf_file_th.CurrentFilter  = sFilter
		sSql = research_pdf_file_th.SQL
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

			' Field rsh_id
			Call research_pdf_file_th.rsh_id.SetDbValue(Rs, research_pdf_file_th.rsh_id.CurrentValue, Null, research_pdf_file_th.rsh_id.ReadOnly)

			' Field rsh_pdf_file
			Call research_pdf_file_th.rsh_pdf_file.SetDbValue(Rs, research_pdf_file_th.rsh_pdf_file.CurrentValue, Null, research_pdf_file_th.rsh_pdf_file.ReadOnly)

			' Field rsh_pdf_title
			Call research_pdf_file_th.rsh_pdf_title.SetDbValue(Rs, research_pdf_file_th.rsh_pdf_title.CurrentValue, Null, research_pdf_file_th.rsh_pdf_title.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = research_pdf_file_th.Row_Updating(RsOld, Rs)
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
				ElseIf research_pdf_file_th.CancelMessage <> "" Then
					FailureMessage = research_pdf_file_th.CancelMessage
					research_pdf_file_th.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call research_pdf_file_th.Row_Updated(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", research_pdf_file_th.TableVar, "pom_research_pdf_file_thlist.asp", research_pdf_file_th.TableVar, True)
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
