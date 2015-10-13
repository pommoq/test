<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_videoinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim video_edit
Set video_edit = New cvideo_edit
Set Page = video_edit

' Page init processing
video_edit.Page_Init()

' Page main processing
video_edit.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
video_edit.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var video_edit = new ew_Page("video_edit");
video_edit.PageID = "edit"; // Page ID
var EW_PAGE_ID = video_edit.PageID; // For backward compatibility
// Form object
var fvideoedit = new ew_Form("fvideoedit");
// Validate form
fvideoedit.Validate = function() {
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
			elm = this.GetElements("x" + infix + "_video_id");
			if (elm && !ew_CheckInteger(elm.value))
				return this.OnError(elm, "<%= ew_JsEncode2(video.video_id.FldErrMsg) %>");
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
fvideoedit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fvideoedit.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fvideoedit.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If video.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% video_edit.ShowPageHeader() %>
<% video_edit.ShowMessage %>
<form name="fvideoedit" id="fvideoedit" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="a_table" id="a_table" value="video">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table class="ewGrid"><tr><td>
<table id="tbl_videoedit" class="table table-bordered table-striped">
<% If video.video_id.Visible Then ' video_id %>
	<tr id="r_video_id">
		<td><span id="elh_video_video_id"><%= video.video_id.FldCaption %></span></td>
		<td<%= video.video_id.CellAttributes %>>
<span id="el_video_video_id" class="control-group">
<span<%= video.video_id.ViewAttributes %>>
<%= video.video_id.EditValue %>
</span>
</span>
<input type="hidden" data-field="x_video_id" name="x_video_id" id="x_video_id" value="<%= Server.HTMLEncode(video.video_id.CurrentValue&"") %>">
<%= video.video_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If video.video_title.Visible Then ' video_title %>
	<tr id="r_video_title">
		<td><span id="elh_video_video_title"><%= video.video_title.FldCaption %></span></td>
		<td<%= video.video_title.CellAttributes %>>
<span id="el_video_video_title" class="control-group">
<input type="text" data-field="x_video_title" name="x_video_title" id="x_video_title" size="30" maxlength="255" placeholder="<%= video.video_title.PlaceHolder %>" value="<%= video.video_title.EditValue %>"<%= video.video_title.EditAttributes %>>
</span>
<%= video.video_title.CustomMsg %></td>
	</tr>
<% End If %>
<% If video.video_link.Visible Then ' video_link %>
	<tr id="r_video_link">
		<td><span id="elh_video_video_link"><%= video.video_link.FldCaption %></span></td>
		<td<%= video.video_link.CellAttributes %>>
<span id="el_video_video_link" class="control-group">
<input type="text" data-field="x_video_link" name="x_video_link" id="x_video_link" size="30" maxlength="255" placeholder="<%= video.video_link.PlaceHolder %>" value="<%= video.video_link.EditValue %>"<%= video.video_link.EditAttributes %>>
</span>
<%= video.video_link.CustomMsg %></td>
	</tr>
<% End If %>
<% If video.video_detail.Visible Then ' video_detail %>
	<tr id="r_video_detail">
		<td><span id="elh_video_video_detail"><%= video.video_detail.FldCaption %></span></td>
		<td<%= video.video_detail.CellAttributes %>>
<span id="el_video_video_detail" class="control-group">
<textarea data-field="x_video_detail" name="x_video_detail" id="x_video_detail" cols="35" rows="4" placeholder="<%= video.video_detail.PlaceHolder %>"<%= video.video_detail.EditAttributes %>><%= video.video_detail.EditValue %></textarea>
</span>
<%= video.video_detail.CustomMsg %></td>
	</tr>
<% End If %>
<% If video.video_create.Visible Then ' video_create %>
	<tr id="r_video_create">
		<td><span id="elh_video_video_create"><%= video.video_create.FldCaption %></span></td>
		<td<%= video.video_create.CellAttributes %>>
<span id="el_video_video_create" class="control-group">
<input type="text" data-field="x_video_create" name="x_video_create" id="x_video_create" placeholder="<%= video.video_create.PlaceHolder %>" value="<%= video.video_create.EditValue %>"<%= video.video_create.EditAttributes %>>
</span>
<%= video.video_create.CustomMsg %></td>
	</tr>
<% End If %>
<% If video.video_update.Visible Then ' video_update %>
	<tr id="r_video_update">
		<td><span id="elh_video_video_update"><%= video.video_update.FldCaption %></span></td>
		<td<%= video.video_update.CellAttributes %>>
<span id="el_video_video_update" class="control-group">
<input type="text" data-field="x_video_update" name="x_video_update" id="x_video_update" placeholder="<%= video.video_update.PlaceHolder %>" value="<%= video.video_update.EditValue %>"<%= video.video_update.EditAttributes %>>
</span>
<%= video.video_update.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("EditBtn") %></button>
</form>
<script type="text/javascript">
fvideoedit.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
video_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set video_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cvideo_edit

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
		TableName = "video"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "video_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If video.UseTokenInUrl Then PageUrl = PageUrl & "t=" & video.TableVar & "&" ' add page token
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
		If video.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (video.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (video.TableVar = Request.QueryString("t"))
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
		If IsEmpty(video) Then Set video = New cvideo
		Set Table = video

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "video"

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

		video.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set video = Nothing
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
		If Request.QueryString("video_id").Count > 0 Then
			video.video_id.QueryStringValue = Request.QueryString("video_id")
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Process form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			video.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values
		Else
			video.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If video.video_id.CurrentValue = "" Then Call Page_Terminate("pom_videolist.asp") ' Invalid key, return to list

		' Validate form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			If Not ValidateForm() Then
				video.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				video.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		End If
		Select Case video.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_videolist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				video.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					sReturnUrl = video.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					video.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		video.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call video.ResetAttrs()
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
				video.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					video.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = video.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			video.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			video.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			video.StartRecordNumber = StartRec
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
		If Not video.video_id.FldIsDetailKey Then video.video_id.FormValue = ObjForm.GetValue("x_video_id")
		If Not video.video_title.FldIsDetailKey Then video.video_title.FormValue = ObjForm.GetValue("x_video_title")
		If Not video.video_link.FldIsDetailKey Then video.video_link.FormValue = ObjForm.GetValue("x_video_link")
		If Not video.video_detail.FldIsDetailKey Then video.video_detail.FormValue = ObjForm.GetValue("x_video_detail")
		If Not video.video_create.FldIsDetailKey Then video.video_create.FormValue = ObjForm.GetValue("x_video_create")
		If Not video.video_create.FldIsDetailKey Then video.video_create.CurrentValue = ew_UnFormatDateTime(video.video_create.CurrentValue, 8)
		If Not video.video_update.FldIsDetailKey Then video.video_update.FormValue = ObjForm.GetValue("x_video_update")
		If Not video.video_update.FldIsDetailKey Then video.video_update.CurrentValue = ew_UnFormatDateTime(video.video_update.CurrentValue, 8)
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		video.video_id.CurrentValue = video.video_id.FormValue
		video.video_title.CurrentValue = video.video_title.FormValue
		video.video_link.CurrentValue = video.video_link.FormValue
		video.video_detail.CurrentValue = video.video_detail.FormValue
		video.video_create.CurrentValue = video.video_create.FormValue
		video.video_create.CurrentValue = ew_UnFormatDateTime(video.video_create.CurrentValue, 8)
		video.video_update.CurrentValue = video.video_update.FormValue
		video.video_update.CurrentValue = ew_UnFormatDateTime(video.video_update.CurrentValue, 8)
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = video.KeyFilter

		' Call Row Selecting event
		Call video.Row_Selecting(sFilter)

		' Load sql based on filter
		video.CurrentFilter = sFilter
		sSql = video.SQL
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
		Call video.Row_Selected(RsRow)
		video.video_id.DbValue = RsRow("video_id")
		video.video_title.DbValue = RsRow("video_title")
		video.video_link.DbValue = RsRow("video_link")
		video.video_detail.DbValue = RsRow("video_detail")
		video.video_create.DbValue = RsRow("video_create")
		video.video_update.DbValue = RsRow("video_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		video.video_id.m_DbValue = Rs("video_id")
		video.video_title.m_DbValue = Rs("video_title")
		video.video_link.m_DbValue = Rs("video_link")
		video.video_detail.m_DbValue = Rs("video_detail")
		video.video_create.m_DbValue = Rs("video_create")
		video.video_update.m_DbValue = Rs("video_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call video.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' video_id
		' video_title
		' video_link
		' video_detail
		' video_create
		' video_update
		' -----------
		'  View  Row
		' -----------

		If video.RowType = EW_ROWTYPE_VIEW Then ' View row

			' video_id
			video.video_id.ViewValue = video.video_id.CurrentValue
			video.video_id.ViewCustomAttributes = ""

			' video_title
			video.video_title.ViewValue = video.video_title.CurrentValue
			video.video_title.ViewCustomAttributes = ""

			' video_link
			video.video_link.ViewValue = video.video_link.CurrentValue
			video.video_link.ViewCustomAttributes = ""

			' video_detail
			video.video_detail.ViewValue = video.video_detail.CurrentValue
			video.video_detail.ViewCustomAttributes = ""

			' video_create
			video.video_create.ViewValue = video.video_create.CurrentValue
			video.video_create.ViewCustomAttributes = ""

			' video_update
			video.video_update.ViewValue = video.video_update.CurrentValue
			video.video_update.ViewCustomAttributes = ""

			' View refer script
			' video_id

			video.video_id.LinkCustomAttributes = ""
			video.video_id.HrefValue = ""
			video.video_id.TooltipValue = ""

			' video_title
			video.video_title.LinkCustomAttributes = ""
			video.video_title.HrefValue = ""
			video.video_title.TooltipValue = ""

			' video_link
			video.video_link.LinkCustomAttributes = ""
			video.video_link.HrefValue = ""
			video.video_link.TooltipValue = ""

			' video_detail
			video.video_detail.LinkCustomAttributes = ""
			video.video_detail.HrefValue = ""
			video.video_detail.TooltipValue = ""

			' video_create
			video.video_create.LinkCustomAttributes = ""
			video.video_create.HrefValue = ""
			video.video_create.TooltipValue = ""

			' video_update
			video.video_update.LinkCustomAttributes = ""
			video.video_update.HrefValue = ""
			video.video_update.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf video.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' video_id
			video.video_id.EditCustomAttributes = ""
			video.video_id.EditValue = video.video_id.CurrentValue
			video.video_id.ViewCustomAttributes = ""

			' video_title
			video.video_title.EditCustomAttributes = ""
			video.video_title.EditValue = ew_HtmlEncode(video.video_title.CurrentValue)
			video.video_title.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(video.video_title.FldCaption))

			' video_link
			video.video_link.EditCustomAttributes = ""
			video.video_link.EditValue = ew_HtmlEncode(video.video_link.CurrentValue)
			video.video_link.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(video.video_link.FldCaption))

			' video_detail
			video.video_detail.EditCustomAttributes = ""
			video.video_detail.EditValue = video.video_detail.CurrentValue
			video.video_detail.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(video.video_detail.FldCaption))

			' video_create
			video.video_create.EditCustomAttributes = ""
			video.video_create.EditValue = ew_HtmlEncode(video.video_create.CurrentValue)
			video.video_create.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(video.video_create.FldCaption))

			' video_update
			video.video_update.EditCustomAttributes = ""
			video.video_update.EditValue = ew_HtmlEncode(video.video_update.CurrentValue)
			video.video_update.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(video.video_update.FldCaption))

			' Edit refer script
			' video_id

			video.video_id.HrefValue = ""

			' video_title
			video.video_title.HrefValue = ""

			' video_link
			video.video_link.HrefValue = ""

			' video_detail
			video.video_detail.HrefValue = ""

			' video_create
			video.video_create.HrefValue = ""

			' video_update
			video.video_update.HrefValue = ""
		End If
		If video.RowType = EW_ROWTYPE_ADD Or video.RowType = EW_ROWTYPE_EDIT Or video.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call video.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If video.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call video.Row_Rendered()
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
		If Not ew_CheckInteger(video.video_id.FormValue) Then
			Call ew_AddMessage(gsFormError, video.video_id.FldErrMsg)
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
		sFilter = video.KeyFilter
		video.CurrentFilter  = sFilter
		sSql = video.SQL
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

			' Field video_id
			' Field video_title

			Call video.video_title.SetDbValue(Rs, video.video_title.CurrentValue, Null, video.video_title.ReadOnly)

			' Field video_link
			Call video.video_link.SetDbValue(Rs, video.video_link.CurrentValue, Null, video.video_link.ReadOnly)

			' Field video_detail
			Call video.video_detail.SetDbValue(Rs, video.video_detail.CurrentValue, Null, video.video_detail.ReadOnly)

			' Field video_create
			Call video.video_create.SetDbValue(Rs, video.video_create.CurrentValue, Null, video.video_create.ReadOnly)

			' Field video_update
			Call video.video_update.SetDbValue(Rs, video.video_update.CurrentValue, Null, video.video_update.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = video.Row_Updating(RsOld, Rs)
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
				ElseIf video.CancelMessage <> "" Then
					FailureMessage = video.CancelMessage
					video.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call video.Row_Updated(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", video.TableVar, "pom_videolist.asp", video.TableVar, True)
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
