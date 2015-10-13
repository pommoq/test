<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_z40newsinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim z40news_edit
Set z40news_edit = New cz40news_edit
Set Page = z40news_edit

' Page init processing
z40news_edit.Page_Init()

' Page main processing
z40news_edit.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
z40news_edit.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var z40news_edit = new ew_Page("z40news_edit");
z40news_edit.PageID = "edit"; // Page ID
var EW_PAGE_ID = z40news_edit.PageID; // For backward compatibility
// Form object
var fz40newsedit = new ew_Form("fz40newsedit");
// Validate form
fz40newsedit.Validate = function() {
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
fz40newsedit.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fz40newsedit.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fz40newsedit.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If z40news.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% z40news_edit.ShowPageHeader() %>
<% z40news_edit.ShowMessage %>
<form name="fz40newsedit" id="fz40newsedit" class="ewForm form-inline" action="<%= ew_CurrentPage %>" method="post">
<input type="hidden" name="a_table" id="a_table" value="z40news">
<input type="hidden" name="a_edit" id="a_edit" value="U">
<table class="ewGrid"><tr><td>
<table id="tbl_z40newsedit" class="table table-bordered table-striped">
<% If z40news.news_id.Visible Then ' news_id %>
	<tr id="r_news_id">
		<td><span id="elh_z40news_news_id"><%= z40news.news_id.FldCaption %></span></td>
		<td<%= z40news.news_id.CellAttributes %>>
<span id="el_z40news_news_id" class="control-group">
<span<%= z40news.news_id.ViewAttributes %>>
<%= z40news.news_id.EditValue %>
</span>
</span>
<input type="hidden" data-field="x_news_id" name="x_news_id" id="x_news_id" value="<%= Server.HTMLEncode(z40news.news_id.CurrentValue&"") %>">
<%= z40news.news_id.CustomMsg %></td>
	</tr>
<% End If %>
<% If z40news.news_img.Visible Then ' news_img %>
	<tr id="r_news_img">
		<td><span id="elh_z40news_news_img"><%= z40news.news_img.FldCaption %></span></td>
		<td<%= z40news.news_img.CellAttributes %>>
<span id="el_z40news_news_img" class="control-group">
<input type="text" data-field="x_news_img" name="x_news_img" id="x_news_img" size="30" maxlength="255" placeholder="<%= z40news.news_img.PlaceHolder %>" value="<%= z40news.news_img.EditValue %>"<%= z40news.news_img.EditAttributes %>>
</span>
<%= z40news.news_img.CustomMsg %></td>
	</tr>
<% End If %>
<% If z40news.news_category.Visible Then ' news_category %>
	<tr id="r_news_category">
		<td><span id="elh_z40news_news_category"><%= z40news.news_category.FldCaption %></span></td>
		<td<%= z40news.news_category.CellAttributes %>>
<span id="el_z40news_news_category" class="control-group">
<input type="text" data-field="x_news_category" name="x_news_category" id="x_news_category" size="30" maxlength="255" placeholder="<%= z40news.news_category.PlaceHolder %>" value="<%= z40news.news_category.EditValue %>"<%= z40news.news_category.EditAttributes %>>
</span>
<%= z40news.news_category.CustomMsg %></td>
	</tr>
<% End If %>
<% If z40news.news_subject.Visible Then ' news_subject %>
	<tr id="r_news_subject">
		<td><span id="elh_z40news_news_subject"><%= z40news.news_subject.FldCaption %></span></td>
		<td<%= z40news.news_subject.CellAttributes %>>
<span id="el_z40news_news_subject" class="control-group">
<input type="text" data-field="x_news_subject" name="x_news_subject" id="x_news_subject" size="30" maxlength="255" placeholder="<%= z40news.news_subject.PlaceHolder %>" value="<%= z40news.news_subject.EditValue %>"<%= z40news.news_subject.EditAttributes %>>
</span>
<%= z40news.news_subject.CustomMsg %></td>
	</tr>
<% End If %>
<% If z40news.news_subject_th.Visible Then ' news_subject_th %>
	<tr id="r_news_subject_th">
		<td><span id="elh_z40news_news_subject_th"><%= z40news.news_subject_th.FldCaption %></span></td>
		<td<%= z40news.news_subject_th.CellAttributes %>>
<span id="el_z40news_news_subject_th" class="control-group">
<input type="text" data-field="x_news_subject_th" name="x_news_subject_th" id="x_news_subject_th" size="30" maxlength="255" placeholder="<%= z40news.news_subject_th.PlaceHolder %>" value="<%= z40news.news_subject_th.EditValue %>"<%= z40news.news_subject_th.EditAttributes %>>
</span>
<%= z40news.news_subject_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If z40news.news_intro.Visible Then ' news_intro %>
	<tr id="r_news_intro">
		<td><span id="elh_z40news_news_intro"><%= z40news.news_intro.FldCaption %></span></td>
		<td<%= z40news.news_intro.CellAttributes %>>
<span id="el_z40news_news_intro" class="control-group">
<textarea data-field="x_news_intro" name="x_news_intro" id="x_news_intro" cols="35" rows="4" placeholder="<%= z40news.news_intro.PlaceHolder %>"<%= z40news.news_intro.EditAttributes %>><%= z40news.news_intro.EditValue %></textarea>
</span>
<%= z40news.news_intro.CustomMsg %></td>
	</tr>
<% End If %>
<% If z40news.news_intro_th.Visible Then ' news_intro_th %>
	<tr id="r_news_intro_th">
		<td><span id="elh_z40news_news_intro_th"><%= z40news.news_intro_th.FldCaption %></span></td>
		<td<%= z40news.news_intro_th.CellAttributes %>>
<span id="el_z40news_news_intro_th" class="control-group">
<textarea data-field="x_news_intro_th" name="x_news_intro_th" id="x_news_intro_th" cols="35" rows="4" placeholder="<%= z40news.news_intro_th.PlaceHolder %>"<%= z40news.news_intro_th.EditAttributes %>><%= z40news.news_intro_th.EditValue %></textarea>
</span>
<%= z40news.news_intro_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If z40news.news_content.Visible Then ' news_content %>
	<tr id="r_news_content">
		<td><span id="elh_z40news_news_content"><%= z40news.news_content.FldCaption %></span></td>
		<td<%= z40news.news_content.CellAttributes %>>
<span id="el_z40news_news_content" class="control-group">
<textarea data-field="x_news_content" name="x_news_content" id="x_news_content" cols="35" rows="4" placeholder="<%= z40news.news_content.PlaceHolder %>"<%= z40news.news_content.EditAttributes %>><%= z40news.news_content.EditValue %></textarea>
</span>
<%= z40news.news_content.CustomMsg %></td>
	</tr>
<% End If %>
<% If z40news.news_content_th.Visible Then ' news_content_th %>
	<tr id="r_news_content_th">
		<td><span id="elh_z40news_news_content_th"><%= z40news.news_content_th.FldCaption %></span></td>
		<td<%= z40news.news_content_th.CellAttributes %>>
<span id="el_z40news_news_content_th" class="control-group">
<textarea data-field="x_news_content_th" name="x_news_content_th" id="x_news_content_th" cols="35" rows="4" placeholder="<%= z40news.news_content_th.PlaceHolder %>"<%= z40news.news_content_th.EditAttributes %>><%= z40news.news_content_th.EditValue %></textarea>
</span>
<%= z40news.news_content_th.CustomMsg %></td>
	</tr>
<% End If %>
<% If z40news.news_show_en.Visible Then ' news_show_en %>
	<tr id="r_news_show_en">
		<td><span id="elh_z40news_news_show_en"><%= z40news.news_show_en.FldCaption %></span></td>
		<td<%= z40news.news_show_en.CellAttributes %>>
<span id="el_z40news_news_show_en" class="control-group">
<input type="text" data-field="x_news_show_en" name="x_news_show_en" id="x_news_show_en" size="30" maxlength="255" placeholder="<%= z40news.news_show_en.PlaceHolder %>" value="<%= z40news.news_show_en.EditValue %>"<%= z40news.news_show_en.EditAttributes %>>
</span>
<%= z40news.news_show_en.CustomMsg %></td>
	</tr>
<% End If %>
<% If z40news.news_show.Visible Then ' news_show %>
	<tr id="r_news_show">
		<td><span id="elh_z40news_news_show"><%= z40news.news_show.FldCaption %></span></td>
		<td<%= z40news.news_show.CellAttributes %>>
<span id="el_z40news_news_show" class="control-group">
<input type="text" data-field="x_news_show" name="x_news_show" id="x_news_show" size="30" maxlength="255" placeholder="<%= z40news.news_show.PlaceHolder %>" value="<%= z40news.news_show.EditValue %>"<%= z40news.news_show.EditAttributes %>>
</span>
<%= z40news.news_show.CustomMsg %></td>
	</tr>
<% End If %>
<% If z40news.news_create.Visible Then ' news_create %>
	<tr id="r_news_create">
		<td><span id="elh_z40news_news_create"><%= z40news.news_create.FldCaption %></span></td>
		<td<%= z40news.news_create.CellAttributes %>>
<span id="el_z40news_news_create" class="control-group">
<input type="text" data-field="x_news_create" name="x_news_create" id="x_news_create" size="30" maxlength="255" placeholder="<%= z40news.news_create.PlaceHolder %>" value="<%= z40news.news_create.EditValue %>"<%= z40news.news_create.EditAttributes %>>
</span>
<%= z40news.news_create.CustomMsg %></td>
	</tr>
<% End If %>
<% If z40news.news_update.Visible Then ' news_update %>
	<tr id="r_news_update">
		<td><span id="elh_z40news_news_update"><%= z40news.news_update.FldCaption %></span></td>
		<td<%= z40news.news_update.CellAttributes %>>
<span id="el_z40news_news_update" class="control-group">
<input type="text" data-field="x_news_update" name="x_news_update" id="x_news_update" size="30" maxlength="255" placeholder="<%= z40news.news_update.PlaceHolder %>" value="<%= z40news.news_update.EditValue %>"<%= z40news.news_update.EditAttributes %>>
</span>
<%= z40news.news_update.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("EditBtn") %></button>
</form>
<script type="text/javascript">
fz40newsedit.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
z40news_edit.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set z40news_edit = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cz40news_edit

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
		TableName = "@news"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "z40news_edit"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If z40news.UseTokenInUrl Then PageUrl = PageUrl & "t=" & z40news.TableVar & "&" ' add page token
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
		If z40news.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (z40news.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (z40news.TableVar = Request.QueryString("t"))
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
		If IsEmpty(z40news) Then Set z40news = New cz40news
		Set Table = z40news

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "edit"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "@news"

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

		z40news.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action
		z40news.news_id.Visible = Not z40news.IsAdd() And Not z40news.IsCopy() And Not z40news.IsGridAdd()

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
		Set z40news = Nothing
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
		If Request.QueryString("news_id").Count > 0 Then
			z40news.news_id.QueryStringValue = Request.QueryString("news_id")
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Process form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			z40news.CurrentAction = ObjForm.GetValue("a_edit") ' Get action code
			Call LoadFormValues() ' Get form values
		Else
			z40news.CurrentAction = "I" ' Default action is display
		End If

		' Check if valid key
		If z40news.news_id.CurrentValue = "" Then Call Page_Terminate("pom_z40newslist.asp") ' Invalid key, return to list

		' Validate form if post back
		If ObjForm.GetValue("a_edit")&"" <> "" Then
			If Not ValidateForm() Then
				z40news.CurrentAction = "" ' Form error, reset action
				FailureMessage = gsFormError
				z40news.EventCancelled = True ' Event cancelled
				Call LoadRow() ' Restore row
				Call RestoreFormValues() ' Restore form values if validate failed
			End If
		End If
		Select Case z40news.CurrentAction
			Case "I" ' Get a record to display
				If Not LoadRow() Then ' Load Record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_z40newslist.asp") ' No matching record, return to list
				End If
			Case "U" ' Update
				z40news.SendEmail = True ' Send email on update success
				If EditRow() Then ' Update Record based on key
					SuccessMessage = Language.Phrase("UpdateSuccess") ' Update success
					sReturnUrl = z40news.ReturnUrl
					Call Page_Terminate(sReturnUrl) ' Return to caller
				Else
					z40news.EventCancelled = True ' Event cancelled
					Call LoadRow() ' Restore row
					Call RestoreFormValues() ' Restore form values if update failed
				End If
		End Select

		' Render the record
		z40news.RowType = EW_ROWTYPE_EDIT ' Render as edit

		' Render row
		Call z40news.ResetAttrs()
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
				z40news.StartRecordNumber = StartRec
			ElseIf Request.QueryString(EW_TABLE_PAGE_NO).Count > 0 Then
				PageNo = Request.QueryString(EW_TABLE_PAGE_NO)
				If IsNumeric(PageNo) Then
					StartRec = (PageNo-1)*DisplayRecs+1
					If StartRec <= 0 Then
						StartRec = 1
					ElseIf StartRec >= ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 Then
						StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1
					End If
					z40news.StartRecordNumber = StartRec
				End If
			End If
		End If
		StartRec = z40news.StartRecordNumber

		' Check if correct start record counter
		If Not IsNumeric(StartRec) Or StartRec = "" Then ' Avoid invalid start record counter
			StartRec = 1 ' Reset start record counter
			z40news.StartRecordNumber = StartRec
		ElseIf CLng(StartRec) > CLng(TotalRecs) Then ' Avoid starting record > total records
			StartRec = ((TotalRecs-1)\DisplayRecs)*DisplayRecs+1 ' Point to last page first record
			z40news.StartRecordNumber = StartRec
		ElseIf (StartRec-1) Mod DisplayRecs <> 0 Then
			StartRec = ((StartRec-1)\DisplayRecs)*DisplayRecs+1 ' Point to page boundary
			z40news.StartRecordNumber = StartRec
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
		If Not z40news.news_id.FldIsDetailKey Then z40news.news_id.FormValue = ObjForm.GetValue("x_news_id")
		If Not z40news.news_img.FldIsDetailKey Then z40news.news_img.FormValue = ObjForm.GetValue("x_news_img")
		If Not z40news.news_category.FldIsDetailKey Then z40news.news_category.FormValue = ObjForm.GetValue("x_news_category")
		If Not z40news.news_subject.FldIsDetailKey Then z40news.news_subject.FormValue = ObjForm.GetValue("x_news_subject")
		If Not z40news.news_subject_th.FldIsDetailKey Then z40news.news_subject_th.FormValue = ObjForm.GetValue("x_news_subject_th")
		If Not z40news.news_intro.FldIsDetailKey Then z40news.news_intro.FormValue = ObjForm.GetValue("x_news_intro")
		If Not z40news.news_intro_th.FldIsDetailKey Then z40news.news_intro_th.FormValue = ObjForm.GetValue("x_news_intro_th")
		If Not z40news.news_content.FldIsDetailKey Then z40news.news_content.FormValue = ObjForm.GetValue("x_news_content")
		If Not z40news.news_content_th.FldIsDetailKey Then z40news.news_content_th.FormValue = ObjForm.GetValue("x_news_content_th")
		If Not z40news.news_show_en.FldIsDetailKey Then z40news.news_show_en.FormValue = ObjForm.GetValue("x_news_show_en")
		If Not z40news.news_show.FldIsDetailKey Then z40news.news_show.FormValue = ObjForm.GetValue("x_news_show")
		If Not z40news.news_create.FldIsDetailKey Then z40news.news_create.FormValue = ObjForm.GetValue("x_news_create")
		If Not z40news.news_update.FldIsDetailKey Then z40news.news_update.FormValue = ObjForm.GetValue("x_news_update")
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadRow()
		z40news.news_id.CurrentValue = z40news.news_id.FormValue
		z40news.news_img.CurrentValue = z40news.news_img.FormValue
		z40news.news_category.CurrentValue = z40news.news_category.FormValue
		z40news.news_subject.CurrentValue = z40news.news_subject.FormValue
		z40news.news_subject_th.CurrentValue = z40news.news_subject_th.FormValue
		z40news.news_intro.CurrentValue = z40news.news_intro.FormValue
		z40news.news_intro_th.CurrentValue = z40news.news_intro_th.FormValue
		z40news.news_content.CurrentValue = z40news.news_content.FormValue
		z40news.news_content_th.CurrentValue = z40news.news_content_th.FormValue
		z40news.news_show_en.CurrentValue = z40news.news_show_en.FormValue
		z40news.news_show.CurrentValue = z40news.news_show.FormValue
		z40news.news_create.CurrentValue = z40news.news_create.FormValue
		z40news.news_update.CurrentValue = z40news.news_update.FormValue
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = z40news.KeyFilter

		' Call Row Selecting event
		Call z40news.Row_Selecting(sFilter)

		' Load sql based on filter
		z40news.CurrentFilter = sFilter
		sSql = z40news.SQL
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
		Call z40news.Row_Selected(RsRow)
		z40news.news_id.DbValue = RsRow("news_id")
		z40news.news_img.DbValue = RsRow("news_img")
		z40news.news_category.DbValue = RsRow("news_category")
		z40news.news_subject.DbValue = RsRow("news_subject")
		z40news.news_subject_th.DbValue = RsRow("news_subject_th")
		z40news.news_intro.DbValue = RsRow("news_intro")
		z40news.news_intro_th.DbValue = RsRow("news_intro_th")
		z40news.news_content.DbValue = RsRow("news_content")
		z40news.news_content_th.DbValue = RsRow("news_content_th")
		z40news.news_show_en.DbValue = RsRow("news_show_en")
		z40news.news_show.DbValue = RsRow("news_show")
		z40news.news_create.DbValue = RsRow("news_create")
		z40news.news_update.DbValue = RsRow("news_update")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		z40news.news_id.m_DbValue = Rs("news_id")
		z40news.news_img.m_DbValue = Rs("news_img")
		z40news.news_category.m_DbValue = Rs("news_category")
		z40news.news_subject.m_DbValue = Rs("news_subject")
		z40news.news_subject_th.m_DbValue = Rs("news_subject_th")
		z40news.news_intro.m_DbValue = Rs("news_intro")
		z40news.news_intro_th.m_DbValue = Rs("news_intro_th")
		z40news.news_content.m_DbValue = Rs("news_content")
		z40news.news_content_th.m_DbValue = Rs("news_content_th")
		z40news.news_show_en.m_DbValue = Rs("news_show_en")
		z40news.news_show.m_DbValue = Rs("news_show")
		z40news.news_create.m_DbValue = Rs("news_create")
		z40news.news_update.m_DbValue = Rs("news_update")
	End Sub

	' -----------------------------------------------------------------
	' Render row values based on field settings
	'
	Sub RenderRow()

		' Initialize urls
		' Call Row Rendering event

		Call z40news.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' news_id
		' news_img
		' news_category
		' news_subject
		' news_subject_th
		' news_intro
		' news_intro_th
		' news_content
		' news_content_th
		' news_show_en
		' news_show
		' news_create
		' news_update
		' -----------
		'  View  Row
		' -----------

		If z40news.RowType = EW_ROWTYPE_VIEW Then ' View row

			' news_id
			z40news.news_id.ViewValue = z40news.news_id.CurrentValue
			z40news.news_id.ViewCustomAttributes = ""

			' news_img
			z40news.news_img.ViewValue = z40news.news_img.CurrentValue
			z40news.news_img.ViewCustomAttributes = ""

			' news_category
			z40news.news_category.ViewValue = z40news.news_category.CurrentValue
			z40news.news_category.ViewCustomAttributes = ""

			' news_subject
			z40news.news_subject.ViewValue = z40news.news_subject.CurrentValue
			z40news.news_subject.ViewCustomAttributes = ""

			' news_subject_th
			z40news.news_subject_th.ViewValue = z40news.news_subject_th.CurrentValue
			z40news.news_subject_th.ViewCustomAttributes = ""

			' news_intro
			z40news.news_intro.ViewValue = z40news.news_intro.CurrentValue
			z40news.news_intro.ViewCustomAttributes = ""

			' news_intro_th
			z40news.news_intro_th.ViewValue = z40news.news_intro_th.CurrentValue
			z40news.news_intro_th.ViewCustomAttributes = ""

			' news_content
			z40news.news_content.ViewValue = z40news.news_content.CurrentValue
			z40news.news_content.ViewCustomAttributes = ""

			' news_content_th
			z40news.news_content_th.ViewValue = z40news.news_content_th.CurrentValue
			z40news.news_content_th.ViewCustomAttributes = ""

			' news_show_en
			z40news.news_show_en.ViewValue = z40news.news_show_en.CurrentValue
			z40news.news_show_en.ViewCustomAttributes = ""

			' news_show
			z40news.news_show.ViewValue = z40news.news_show.CurrentValue
			z40news.news_show.ViewCustomAttributes = ""

			' news_create
			z40news.news_create.ViewValue = z40news.news_create.CurrentValue
			z40news.news_create.ViewCustomAttributes = ""

			' news_update
			z40news.news_update.ViewValue = z40news.news_update.CurrentValue
			z40news.news_update.ViewCustomAttributes = ""

			' View refer script
			' news_id

			z40news.news_id.LinkCustomAttributes = ""
			z40news.news_id.HrefValue = ""
			z40news.news_id.TooltipValue = ""

			' news_img
			z40news.news_img.LinkCustomAttributes = ""
			z40news.news_img.HrefValue = ""
			z40news.news_img.TooltipValue = ""

			' news_category
			z40news.news_category.LinkCustomAttributes = ""
			z40news.news_category.HrefValue = ""
			z40news.news_category.TooltipValue = ""

			' news_subject
			z40news.news_subject.LinkCustomAttributes = ""
			z40news.news_subject.HrefValue = ""
			z40news.news_subject.TooltipValue = ""

			' news_subject_th
			z40news.news_subject_th.LinkCustomAttributes = ""
			z40news.news_subject_th.HrefValue = ""
			z40news.news_subject_th.TooltipValue = ""

			' news_intro
			z40news.news_intro.LinkCustomAttributes = ""
			z40news.news_intro.HrefValue = ""
			z40news.news_intro.TooltipValue = ""

			' news_intro_th
			z40news.news_intro_th.LinkCustomAttributes = ""
			z40news.news_intro_th.HrefValue = ""
			z40news.news_intro_th.TooltipValue = ""

			' news_content
			z40news.news_content.LinkCustomAttributes = ""
			z40news.news_content.HrefValue = ""
			z40news.news_content.TooltipValue = ""

			' news_content_th
			z40news.news_content_th.LinkCustomAttributes = ""
			z40news.news_content_th.HrefValue = ""
			z40news.news_content_th.TooltipValue = ""

			' news_show_en
			z40news.news_show_en.LinkCustomAttributes = ""
			z40news.news_show_en.HrefValue = ""
			z40news.news_show_en.TooltipValue = ""

			' news_show
			z40news.news_show.LinkCustomAttributes = ""
			z40news.news_show.HrefValue = ""
			z40news.news_show.TooltipValue = ""

			' news_create
			z40news.news_create.LinkCustomAttributes = ""
			z40news.news_create.HrefValue = ""
			z40news.news_create.TooltipValue = ""

			' news_update
			z40news.news_update.LinkCustomAttributes = ""
			z40news.news_update.HrefValue = ""
			z40news.news_update.TooltipValue = ""

		' ----------
		'  Edit Row
		' ----------

		ElseIf z40news.RowType = EW_ROWTYPE_EDIT Then ' Edit row

			' news_id
			z40news.news_id.EditCustomAttributes = ""
			z40news.news_id.EditValue = z40news.news_id.CurrentValue
			z40news.news_id.ViewCustomAttributes = ""

			' news_img
			z40news.news_img.EditCustomAttributes = ""
			z40news.news_img.EditValue = ew_HtmlEncode(z40news.news_img.CurrentValue)
			z40news.news_img.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(z40news.news_img.FldCaption))

			' news_category
			z40news.news_category.EditCustomAttributes = ""
			z40news.news_category.EditValue = ew_HtmlEncode(z40news.news_category.CurrentValue)
			z40news.news_category.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(z40news.news_category.FldCaption))

			' news_subject
			z40news.news_subject.EditCustomAttributes = ""
			z40news.news_subject.EditValue = ew_HtmlEncode(z40news.news_subject.CurrentValue)
			z40news.news_subject.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(z40news.news_subject.FldCaption))

			' news_subject_th
			z40news.news_subject_th.EditCustomAttributes = ""
			z40news.news_subject_th.EditValue = ew_HtmlEncode(z40news.news_subject_th.CurrentValue)
			z40news.news_subject_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(z40news.news_subject_th.FldCaption))

			' news_intro
			z40news.news_intro.EditCustomAttributes = ""
			z40news.news_intro.EditValue = z40news.news_intro.CurrentValue
			z40news.news_intro.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(z40news.news_intro.FldCaption))

			' news_intro_th
			z40news.news_intro_th.EditCustomAttributes = ""
			z40news.news_intro_th.EditValue = z40news.news_intro_th.CurrentValue
			z40news.news_intro_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(z40news.news_intro_th.FldCaption))

			' news_content
			z40news.news_content.EditCustomAttributes = ""
			z40news.news_content.EditValue = z40news.news_content.CurrentValue
			z40news.news_content.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(z40news.news_content.FldCaption))

			' news_content_th
			z40news.news_content_th.EditCustomAttributes = ""
			z40news.news_content_th.EditValue = z40news.news_content_th.CurrentValue
			z40news.news_content_th.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(z40news.news_content_th.FldCaption))

			' news_show_en
			z40news.news_show_en.EditCustomAttributes = ""
			z40news.news_show_en.EditValue = ew_HtmlEncode(z40news.news_show_en.CurrentValue)
			z40news.news_show_en.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(z40news.news_show_en.FldCaption))

			' news_show
			z40news.news_show.EditCustomAttributes = ""
			z40news.news_show.EditValue = ew_HtmlEncode(z40news.news_show.CurrentValue)
			z40news.news_show.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(z40news.news_show.FldCaption))

			' news_create
			z40news.news_create.EditCustomAttributes = ""
			z40news.news_create.EditValue = ew_HtmlEncode(z40news.news_create.CurrentValue)
			z40news.news_create.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(z40news.news_create.FldCaption))

			' news_update
			z40news.news_update.EditCustomAttributes = ""
			z40news.news_update.EditValue = ew_HtmlEncode(z40news.news_update.CurrentValue)
			z40news.news_update.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(z40news.news_update.FldCaption))

			' Edit refer script
			' news_id

			z40news.news_id.HrefValue = ""

			' news_img
			z40news.news_img.HrefValue = ""

			' news_category
			z40news.news_category.HrefValue = ""

			' news_subject
			z40news.news_subject.HrefValue = ""

			' news_subject_th
			z40news.news_subject_th.HrefValue = ""

			' news_intro
			z40news.news_intro.HrefValue = ""

			' news_intro_th
			z40news.news_intro_th.HrefValue = ""

			' news_content
			z40news.news_content.HrefValue = ""

			' news_content_th
			z40news.news_content_th.HrefValue = ""

			' news_show_en
			z40news.news_show_en.HrefValue = ""

			' news_show
			z40news.news_show.HrefValue = ""

			' news_create
			z40news.news_create.HrefValue = ""

			' news_update
			z40news.news_update.HrefValue = ""
		End If
		If z40news.RowType = EW_ROWTYPE_ADD Or z40news.RowType = EW_ROWTYPE_EDIT Or z40news.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call z40news.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If z40news.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call z40news.Row_Rendered()
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
		sFilter = z40news.KeyFilter
		z40news.CurrentFilter  = sFilter
		sSql = z40news.SQL
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

			' Field news_img
			Call z40news.news_img.SetDbValue(Rs, z40news.news_img.CurrentValue, Null, z40news.news_img.ReadOnly)

			' Field news_category
			Call z40news.news_category.SetDbValue(Rs, z40news.news_category.CurrentValue, Null, z40news.news_category.ReadOnly)

			' Field news_subject
			Call z40news.news_subject.SetDbValue(Rs, z40news.news_subject.CurrentValue, Null, z40news.news_subject.ReadOnly)

			' Field news_subject_th
			Call z40news.news_subject_th.SetDbValue(Rs, z40news.news_subject_th.CurrentValue, Null, z40news.news_subject_th.ReadOnly)

			' Field news_intro
			Call z40news.news_intro.SetDbValue(Rs, z40news.news_intro.CurrentValue, Null, z40news.news_intro.ReadOnly)

			' Field news_intro_th
			Call z40news.news_intro_th.SetDbValue(Rs, z40news.news_intro_th.CurrentValue, Null, z40news.news_intro_th.ReadOnly)

			' Field news_content
			Call z40news.news_content.SetDbValue(Rs, z40news.news_content.CurrentValue, Null, z40news.news_content.ReadOnly)

			' Field news_content_th
			Call z40news.news_content_th.SetDbValue(Rs, z40news.news_content_th.CurrentValue, Null, z40news.news_content_th.ReadOnly)

			' Field news_show_en
			Call z40news.news_show_en.SetDbValue(Rs, z40news.news_show_en.CurrentValue, Null, z40news.news_show_en.ReadOnly)

			' Field news_show
			Call z40news.news_show.SetDbValue(Rs, z40news.news_show.CurrentValue, Null, z40news.news_show.ReadOnly)

			' Field news_create
			Call z40news.news_create.SetDbValue(Rs, z40news.news_create.CurrentValue, Null, z40news.news_create.ReadOnly)

			' Field news_update
			Call z40news.news_update.SetDbValue(Rs, z40news.news_update.CurrentValue, Null, z40news.news_update.ReadOnly)

			' Check recordset update error
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				Rs.Close
				Set Rs = Nothing
				EditRow = False
				Exit Function
			End If

			' Call Row Updating event
			bUpdateRow = z40news.Row_Updating(RsOld, Rs)
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
				ElseIf z40news.CancelMessage <> "" Then
					FailureMessage = z40news.CancelMessage
					z40news.CancelMessage = ""
				Else
					FailureMessage = Language.Phrase("UpdateCancelled")
				End If
				EditRow = False
			End If
		End If

		' Call Row Updated event
		If EditRow Then
			Call z40news.Row_Updated(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", z40news.TableVar, "pom_z40newslist.asp", z40news.TableVar, True)
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
