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
Dim video_add
Set video_add = New cvideo_add
Set Page = video_add

' Page init processing
video_add.Page_Init()

' Page main processing
video_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
video_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var video_add = new ew_Page("video_add");
video_add.PageID = "add"; // Page ID
var EW_PAGE_ID = video_add.PageID; // For backward compatibility
// Form object
var fvideoadd = new ew_Form("fvideoadd");
// Validate form
fvideoadd.Validate = function() {
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
fvideoadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fvideoadd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fvideoadd.ValidateRequired = false; // No JavaScript validation
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
<% video_add.ShowPageHeader() %>
<% video_add.ShowMessage %>
<form name="fvideoadd" id="fvideoadd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="video">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_videoadd" class="table table-bordered table-striped">
<% If video.video_id.Visible Then ' video_id %>
	<tr id="r_video_id">
		<td><span id="elh_video_video_id"><%= video.video_id.FldCaption %></span></td>
		<td<%= video.video_id.CellAttributes %>>
<span id="el_video_video_id" class="control-group">
<input type="text" data-field="x_video_id" name="x_video_id" id="x_video_id" size="30" placeholder="<%= video.video_id.PlaceHolder %>" value="<%= video.video_id.EditValue %>"<%= video.video_id.EditAttributes %>>
</span>
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
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
fvideoadd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
video_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set video_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cvideo_add

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
		TableName = "video"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "video_add"
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
		EW_PAGE_ID = "add"

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
	Dim Priv
	Dim OldRecordset
	Dim CopyRecord

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()

		' Process form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			video.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("video_id").Count > 0 Then
				video.video_id.QueryStringValue = Request.QueryString("video_id")
				Call video.SetKey("video_id", video.video_id.CurrentValue) ' Set up key
			Else
				Call video.SetKey("video_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				video.CurrentAction = "C" ' Copy Record
			Else
				video.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				video.CurrentAction = "I" ' Form error, reset action
				video.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case video.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_videolist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				video.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = video.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_videoview.asp" Then sReturnUrl = video.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					video.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		video.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call video.ResetAttrs()
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
		video.video_id.CurrentValue = Null
		video.video_id.OldValue = video.video_id.CurrentValue
		video.video_title.CurrentValue = Null
		video.video_title.OldValue = video.video_title.CurrentValue
		video.video_link.CurrentValue = Null
		video.video_link.OldValue = video.video_link.CurrentValue
		video.video_detail.CurrentValue = Null
		video.video_detail.OldValue = video.video_detail.CurrentValue
		video.video_create.CurrentValue = Null
		video.video_create.OldValue = video.video_create.CurrentValue
		video.video_update.CurrentValue = Null
		video.video_update.OldValue = video.video_update.CurrentValue
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
		Call LoadOldRecord()
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

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If video.GetKey("video_id")&"" <> "" Then
			video.video_id.CurrentValue = video.GetKey("video_id") ' video_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			video.CurrentFilter = video.KeyFilter
			Dim sSql
			sSql = video.SQL
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

		' ---------
		'  Add Row
		' ---------

		ElseIf video.RowType = EW_ROWTYPE_ADD Then ' Add row

			' video_id
			video.video_id.EditCustomAttributes = ""
			video.video_id.EditValue = ew_HtmlEncode(video.video_id.CurrentValue)
			video.video_id.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(video.video_id.FldCaption))

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
		If video.video_id.CurrentValue <> "" Then ' Check field with unique index
			sFilter = "([video_id] = " & ew_AdjustSql(video.video_id.CurrentValue) & ")"
			Set RsChk = video.LoadRs(sFilter)
			If Not (RsChk Is Nothing) Then
				sIdxErrMsg = Replace(Language.Phrase("DupIndex"), "%f", video.video_id.FldCaption)
				sIdxErrMsg = Replace(sIdxErrMsg, "%v", video.video_id.CurrentValue)
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
		video.CurrentFilter = sFilter
		sSql = video.SQL
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

		' Field video_id
		Call video.video_id.SetDbValue(Rs, video.video_id.CurrentValue, Null, False)

		' Field video_title
		Call video.video_title.SetDbValue(Rs, video.video_title.CurrentValue, Null, False)

		' Field video_link
		Call video.video_link.SetDbValue(Rs, video.video_link.CurrentValue, Null, False)

		' Field video_detail
		Call video.video_detail.SetDbValue(Rs, video.video_detail.CurrentValue, Null, False)

		' Field video_create
		Call video.video_create.SetDbValue(Rs, video.video_create.CurrentValue, Null, False)

		' Field video_update
		Call video.video_update.SetDbValue(Rs, video.video_update.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = video.Row_Inserting(RsOld, Rs)

		' Check if key value entered
		If bInsertRow And video.ValidateKey And video.video_id.CurrentValue = "" And video.video_id.SessionValue = "" Then
			FailureMessage = Language.Phrase("InvalidKeyValue")
			bInsertRow = False
		End If

		' Check for duplicate key
		Dim sKeyErrMsg
		If bInsertRow And video.ValidateKey Then
			sFilter = video.KeyFilter
			Set RsChk = video.LoadRs(sFilter)
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
			ElseIf video.CancelMessage <> "" Then
				FailureMessage = video.CancelMessage
				video.CancelMessage = ""
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
			Call video.Row_Inserted(RsOld, RsNew)
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
		PageId = ew_IIf(video.CurrentAction = "C", "Copy", "Add")
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
