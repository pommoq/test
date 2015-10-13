<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_news_pdf_fileinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim news_pdf_file_edit
Set news_pdf_file_edit = New cnews_pdf_file_edit
Set Page = news_pdf_file_edit

' Page init processing
news_pdf_file_edit.Page_Init()

' Page main processing
news_pdf_file_edit.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
news_pdf_file_edit.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var news_pdf_file_edit = new ew_Page("news_pdf_file_edit");
news_pdf_file_edit.PageID = "edit"; // Page ID
var EW_PAGE_ID = news_pdf_file_edit.PageID; // For backward compatibility
// Form object
var fnews_pdf_fileedit = new ew_Form("fnews_pdf_fileedit");
// Validate form
fnews_pdf_fileedit.Validate = function() {
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
			elm = this.GetElements("x" + infix + "_news_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(news_pdf_file.news_id.FldErrMsg) %>");
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
fnews_pdf_fileedit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fnews_pdf_fileedit.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fnews_pdf_fileedit.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If news_pdf_file.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% news_pdf_file_edit.ShowPageHeader() %>
<% news_pdf_file_edit.ShowMessage %>
<form name="fnews_pdf_fileedit" id="fnews_pdf_fileedit" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="a_table" id="a_table" value="news_pdf_file">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table class="ewGrid"><tr><td>
<table id="tbl_news_pdf_fileedit" class="table table-bordered table-striped">
<% If news_pdf_file.news_pdf_id.Visible Then ' news_pdf_id %>
	<tr id="r_news_pdf_id">
		<td><span id="elh_news_pdf_file_news_pdf_id"><%= news_pdf_file.news_pdf_id.FldCaption %></span></td>
		<td<%= news_pdf_file.news_pdf_id.CellAttributes %>>
<span id="el_news_pdf_file_news_pdf_id" class="control-group">
<span<%= news_pdf_file.news_pdf_id.ViewAttributes %>>
<%= news_pdf_file.news_pdf_id.EditValue %>
</span>
</span>
<input type="hidden" data-field="x_news_pdf_id" name="x_news_pdf_id" id="x_news_pdf_id" value="<%= Server.HTMLEncode(news_pdf_file.news_pdf_id.CurrentValue&"") %>">
<%= news_pdf_file.news_pdf_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_pdf_file.news_id.Visible Then ' news_id %>
	<tr id="r_news_id">
		<td><span id="elh_news_pdf_file_news_id"><%= news_pdf_file.news_id.FldCaption %></span></td>
		<td<%= news_pdf_file.news_id.CellAttributes %>>
<span id="el_news_pdf_file_news_id" class="control-group">
<input type="text" data-field="x_news_id" name="x_news_id" id="x_news_id" size="30" placeholder="<%= news_pdf_file.news_id.PlaceHolder %>" value="<%= news_pdf_file.news_id.EditValue %>"<%= news_pdf_file.news_id.EditAttributes %>>
</span>
<%= news_pdf_file.news_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_pdf_file.news_pdf_file_1.Visible Then ' news_pdf_file %>
	<tr id="r_news_pdf_file_1">
		<td><span id="elh_news_pdf_file_news_pdf_file_1"><%= news_pdf_file.news_pdf_file_1.FldCaption %></span></td>
		<td<%= news_pdf_file.news_pdf_file_1.CellAttributes %>>
<span id="el_news_pdf_file_news_pdf_file_1" class="control-group">
<input type="text" data-field="x_news_pdf_file_1" name="x_news_pdf_file_1" id="x_news_pdf_file_1" size="30" maxlength="255" placeholder="<%= news_pdf_file.news_pdf_file_1.PlaceHolder %>" value="<%= news_pdf_file.news_pdf_file_1.EditValue %>"<%= news_pdf_file.news_pdf_file_1.EditAttributes %>>
</span>
<%= news_pdf_file.news_pdf_file_1.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_pdf_file.news_pdf_title.Visible Then ' news_pdf_title %>
	<tr id="r_news_pdf_title">
		<td><span id="elh_news_pdf_file_news_pdf_title"><%= news_pdf_file.news_pdf_title.FldCaption %></span></td>
		<td<%= news_pdf_file.news_pdf_title.CellAttributes %>>
<span id="el_news_pdf_file_news_pdf_title" class="control-group">
<input type="text" data-field="x_news_pdf_title" name="x_news_pdf_title" id="x_news_pdf_title" size="30" maxlength="255" placeholder="<%= news_pdf_file.news_pdf_title.PlaceHolder %>" value="<%= news_pdf_file.news_pdf_title.EditValue %>"<%= news_pdf_file.news_pdf_title.EditAttributes %>>
</span>
<%= news_pdf_file.news_pdf_title.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("EditBtn") %></button>
</form>
<script type="text/javascript">
fnews_pdf_fileedit.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
news_pdf_file_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set news_pdf_file_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cnews_pdf_file_edit

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
		TableName = "news_pdf_file"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "news_pdf_file_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If news_pdf_file.UseTokenInUrl Then PageUrl = PageUrl & "t=" & news_pdf_file.TableVar & "&" ' add page token
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
		If news_pdf_file.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (news_pdf_file.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (news_pdf_file.TableVar = Request.QueryString("t"))
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
		If IsEmpty(news_pdf_file) Then Set news_pdf_file = New cnews_pdf_file
		Set Table = news_pdf_file

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "news_pdf_file"

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

		news_pdf_file.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action
		news_pdf_file.news_pdf_id.Visible = Not news_pdf_file.IsAdd() And Not news_pdf_file.IsCopy() And Not news_pdf_file.IsGridAdd()

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
		Set news_pdf_file = Nothing
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
		If Request.QueryString("news_pdf_id").Count > 0 Then
			news_pdf_file.news_pdf_id.QueryStringValue = Request.QueryString("news_pdf_id")
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Process form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			news_pdf_file.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values
		Else
			news_pdf_file.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If news_pdf_file.news_pdf_id.CurrentValue = "" Then Call Page_Terminate("pom_news_pdf_filelist.asp") ' Invalid key, return to list

		' Validate form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			If Not ValidateForm() Then
				news_pdf_file.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				news_pdf_file.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		End If
		Select Case news_pdf_file.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_news_pdf_filelist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				news_pdf_file.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					sReturnUrl = news_pdf_file.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					news_pdf_file.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		news_pdf_file.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call news_pdf_file.ResetAttrs()
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
				news_pdf_file.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					news_pdf_file.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = news_pdf_file.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			news_pdf_file.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			news_pdf_file.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			news_pdf_file.StartRecordNumber = StartRec
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
		If Not news_pdf_file.news_pdf_id.FldIsDetailKey Then news_pdf_file.news_pdf_id.FormValue = ObjForm.GetValue("x_news_pdf_id")
		If Not news_pdf_file.news_id.FldIsDetailKey Then news_pdf_file.news_id.FormValue = ObjForm.GetValue("x_news_id")
		If Not news_pdf_file.news_pdf_file_1.FldIsDetailKey Then news_pdf_file.news_pdf_file_1.FormValue = ObjForm.GetValue("x_news_pdf_file_1")
		If Not news_pdf_file.news_pdf_title.FldIsDetailKey Then news_pdf_file.news_pdf_title.FormValue = ObjForm.GetValue("x_news_pdf_title")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		news_pdf_file.news_pdf_id.CurrentValue = news_pdf_file.news_pdf_id.FormValue
		news_pdf_file.news_id.CurrentValue = news_pdf_file.news_id.FormValue
		news_pdf_file.news_pdf_file_1.CurrentValue = news_pdf_file.news_pdf_file_1.FormValue
		news_pdf_file.news_pdf_title.CurrentValue = news_pdf_file.news_pdf_title.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = news_pdf_file.KeyFilter

		' Call Row Selecting event
		Call news_pdf_file.Row_Selecting(sFilter)

		' Load sql based on filter
		news_pdf_file.CurrentFilter = sFilter
		sSql = news_pdf_file.SQL
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
		Call news_pdf_file.Row_Selected(RsRow)
		news_pdf_file.news_pdf_id.DbValue = RsRow("news_pdf_id")
		news_pdf_file.news_id.DbValue = RsRow("news_id")
		news_pdf_file.news_pdf_file_1.DbValue = RsRow("news_pdf_file")
		news_pdf_file.news_pdf_title.DbValue = RsRow("news_pdf_title")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		news_pdf_file.news_pdf_id.m_DbValue = Rs("news_pdf_id")
		news_pdf_file.news_id.m_DbValue = Rs("news_id")
		news_pdf_file.news_pdf_file_1.m_DbValue = Rs("news_pdf_file")
		news_pdf_file.news_pdf_title.m_DbValue = Rs("news_pdf_title")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call news_pdf_file.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' news_pdf_id
		' news_id
		' news_pdf_file
		' news_pdf_title
		' -----------
		'  View  Row
		' -----------

		If news_pdf_file.RowType = EW_ROWTYPE_VIEW Then ' View row

			' news_pdf_id
			news_pdf_file.news_pdf_id.ViewValue = news_pdf_file.news_pdf_id.CurrentValue
			news_pdf_file.news_pdf_id.ViewCustomAttributes = ""

			' news_id
			news_pdf_file.news_id.ViewValue = news_pdf_file.news_id.CurrentValue
			news_pdf_file.news_id.ViewCustomAttributes = ""

			' news_pdf_file
			news_pdf_file.news_pdf_file_1.ViewValue = news_pdf_file.news_pdf_file_1.CurrentValue
			news_pdf_file.news_pdf_file_1.ViewCustomAttributes = ""

			' news_pdf_title
			news_pdf_file.news_pdf_title.ViewValue = news_pdf_file.news_pdf_title.CurrentValue
			news_pdf_file.news_pdf_title.ViewCustomAttributes = ""

			' View refer script
			' news_pdf_id

			news_pdf_file.news_pdf_id.LinkCustomAttributes = ""
			news_pdf_file.news_pdf_id.HrefValue = ""
			news_pdf_file.news_pdf_id.TooltipValue = ""

			' news_id
			news_pdf_file.news_id.LinkCustomAttributes = ""
			news_pdf_file.news_id.HrefValue = ""
			news_pdf_file.news_id.TooltipValue = ""

			' news_pdf_file
			news_pdf_file.news_pdf_file_1.LinkCustomAttributes = ""
			news_pdf_file.news_pdf_file_1.HrefValue = ""
			news_pdf_file.news_pdf_file_1.TooltipValue = ""

			' news_pdf_title
			news_pdf_file.news_pdf_title.LinkCustomAttributes = ""
			news_pdf_file.news_pdf_title.HrefValue = ""
			news_pdf_file.news_pdf_title.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf news_pdf_file.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' news_pdf_id
			news_pdf_file.news_pdf_id.EditCustomAttributes = ""
			news_pdf_file.news_pdf_id.EditValue = news_pdf_file.news_pdf_id.CurrentValue
			news_pdf_file.news_pdf_id.ViewCustomAttributes = ""

			' news_id
			news_pdf_file.news_id.EditCustomAttributes = ""
			news_pdf_file.news_id.EditValue = ew_HtmlEncode(news_pdf_file.news_id.CurrentValue)
			news_pdf_file.news_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_pdf_file.news_id.FldCaption))

			' news_pdf_file
			news_pdf_file.news_pdf_file_1.EditCustomAttributes = ""
			news_pdf_file.news_pdf_file_1.EditValue = ew_HtmlEncode(news_pdf_file.news_pdf_file_1.CurrentValue)
			news_pdf_file.news_pdf_file_1.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_pdf_file.news_pdf_file_1.FldCaption))

			' news_pdf_title
			news_pdf_file.news_pdf_title.EditCustomAttributes = ""
			news_pdf_file.news_pdf_title.EditValue = ew_HtmlEncode(news_pdf_file.news_pdf_title.CurrentValue)
			news_pdf_file.news_pdf_title.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_pdf_file.news_pdf_title.FldCaption))

			' Edit refer script
			' news_pdf_id

			news_pdf_file.news_pdf_id.HrefValue = ""

			' news_id
			news_pdf_file.news_id.HrefValue = ""

			' news_pdf_file
			news_pdf_file.news_pdf_file_1.HrefValue = ""

			' news_pdf_title
			news_pdf_file.news_pdf_title.HrefValue = ""
		End If
		If news_pdf_file.RowType = EW_ROWTYPE_ADD Or news_pdf_file.RowType = EW_ROWTYPE_EDIT Or news_pdf_file.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call news_pdf_file.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If news_pdf_file.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call news_pdf_file.Row_Rendered()
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
		If Not ew_CheckInteger(news_pdf_file.news_id.FormValue) Then
			Call ew_AddMessage(gsFormError, news_pdf_file.news_id.FldErrMsg)
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
		sFilter = news_pdf_file.KeyFilter
		news_pdf_file.CurrentFilter  = sFilter
		sSql = news_pdf_file.SQL
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

			' Field news_id
			Call news_pdf_file.news_id.SetDbValue(Rs, news_pdf_file.news_id.CurrentValue, Null, news_pdf_file.news_id.ReadOnly)

			' Field news_pdf_file
			Call news_pdf_file.news_pdf_file_1.SetDbValue(Rs, news_pdf_file.news_pdf_file_1.CurrentValue, Null, news_pdf_file.news_pdf_file_1.ReadOnly)

			' Field news_pdf_title
			Call news_pdf_file.news_pdf_title.SetDbValue(Rs, news_pdf_file.news_pdf_title.CurrentValue, Null, news_pdf_file.news_pdf_title.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = news_pdf_file.Row_Updating(RsOld, Rs)
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
				ElseIf news_pdf_file.CancelMessage <> "" Then
					FailureMessage = news_pdf_file.CancelMessage
					news_pdf_file.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call news_pdf_file.Row_Updated(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", news_pdf_file.TableVar, "pom_news_pdf_filelist.asp", news_pdf_file.TableVar, True)
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
