<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_news_saleinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim news_sale_add
Set news_sale_add = New cnews_sale_add
Set Page = news_sale_add

' Page init processing
news_sale_add.Page_Init()

' Page main processing
news_sale_add.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
news_sale_add.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var news_sale_add = new ew_Page("news_sale_add");
news_sale_add.PageID = "add"; // Page ID
var EW_PAGE_ID = news_sale_add.PageID; // For backward compatibility
// Form object
var fnews_saleadd = new ew_Form("fnews_saleadd");
// Validate form
fnews_saleadd.Validate = function() {
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
fnews_saleadd.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fnews_saleadd.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fnews_saleadd.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<% If news_sale.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% news_sale_add.ShowPageHeader() %>
<% news_sale_add.ShowMessage %>
<form name="fnews_saleadd" id="fnews_saleadd" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="news_sale">
<input type="hidden" name="a_add" id="a_add" value="A">
<table class="ewGrid"><tr><td>
<table id="tbl_news_saleadd" class="table table-bordered table-striped">
<% If news_sale.news_sale_pdf.Visible Then ' news_sale_pdf %>
	<tr id="r_news_sale_pdf">
		<td><span id="elh_news_sale_news_sale_pdf"><%= news_sale.news_sale_pdf.FldCaption %></span></td>
		<td<%= news_sale.news_sale_pdf.CellAttributes %>>
<span id="el_news_sale_news_sale_pdf" class="control-group">
<input type="text" data-field="x_news_sale_pdf" name="x_news_sale_pdf" id="x_news_sale_pdf" size="30" maxlength="255" placeholder="<%= news_sale.news_sale_pdf.PlaceHolder %>" value="<%= news_sale.news_sale_pdf.EditValue %>"<%= news_sale.news_sale_pdf.EditAttributes %>>
</span>
<%= news_sale.news_sale_pdf.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_sale.news_sale_title.Visible Then ' news_sale_title %>
	<tr id="r_news_sale_title">
		<td><span id="elh_news_sale_news_sale_title"><%= news_sale.news_sale_title.FldCaption %></span></td>
		<td<%= news_sale.news_sale_title.CellAttributes %>>
<span id="el_news_sale_news_sale_title" class="control-group">
<input type="text" data-field="x_news_sale_title" name="x_news_sale_title" id="x_news_sale_title" size="30" maxlength="255" placeholder="<%= news_sale.news_sale_title.PlaceHolder %>" value="<%= news_sale.news_sale_title.EditValue %>"<%= news_sale.news_sale_title.EditAttributes %>>
</span>
<%= news_sale.news_sale_title.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_sale.start_date.Visible Then ' start_date %>
	<tr id="r_start_date">
		<td><span id="elh_news_sale_start_date"><%= news_sale.start_date.FldCaption %></span></td>
		<td<%= news_sale.start_date.CellAttributes %>>
<span id="el_news_sale_start_date" class="control-group">
<input type="text" data-field="x_start_date" name="x_start_date" id="x_start_date" placeholder="<%= news_sale.start_date.PlaceHolder %>" value="<%= news_sale.start_date.EditValue %>"<%= news_sale.start_date.EditAttributes %>>
</span>
<%= news_sale.start_date.CustomMsg %></td>
	</tr>
<% End If %>
<% If news_sale.end_date.Visible Then ' end_date %>
	<tr id="r_end_date">
		<td><span id="elh_news_sale_end_date"><%= news_sale.end_date.FldCaption %></span></td>
		<td<%= news_sale.end_date.CellAttributes %>>
<span id="el_news_sale_end_date" class="control-group">
<input type="text" data-field="x_end_date" name="x_end_date" id="x_end_date" placeholder="<%= news_sale.end_date.PlaceHolder %>" value="<%= news_sale.end_date.EditValue %>"<%= news_sale.end_date.EditAttributes %>>
</span>
<%= news_sale.end_date.CustomMsg %></td>
	</tr>
<% End If %>
</table>
</td></tr></table>
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("AddBtn") %></button>
</form>
<script type="text/javascript">
fnews_saleadd.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
news_sale_add.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set news_sale_add = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cnews_sale_add

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
		TableName = "news_sale"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "news_sale_add"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
		If news_sale.UseTokenInUrl Then PageUrl = PageUrl & "t=" & news_sale.TableVar & "&" ' add page token
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
		If news_sale.UseTokenInUrl Then
			IsPageRequest = False
			If Not (ObjForm Is Nothing) Then
				IsPageRequest = (news_sale.TableVar = ObjForm.GetValue("t"))
			End If
			If Request.QueryString("t").Count > 0 Then
				IsPageRequest = (news_sale.TableVar = Request.QueryString("t"))
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
		If IsEmpty(news_sale) Then Set news_sale = New cnews_sale
		Set Table = news_sale

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "add"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "news_sale"

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

		news_sale.CurrentAction = ew_IIf(Request.QueryString("a").Count > 0, Request.QueryString("a") & "", ObjForm.GetValue("a_list") & "") ' Set up current action

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
		Set news_sale = Nothing
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
			news_sale.CurrentAction = ObjForm.GetValue("a_add") ' Get form action
			CopyRecord = LoadOldRecord() ' Load old recordset
			Call LoadFormValues() ' Load form values

		' Not post back
		Else

			' Load key values from QueryString
			CopyRecord = True
			If Request.QueryString("news_sale_id").Count > 0 Then
				news_sale.news_sale_id.QueryStringValue = Request.QueryString("news_sale_id")
				Call news_sale.SetKey("news_sale_id", news_sale.news_sale_id.CurrentValue) ' Set up key
			Else
				Call news_sale.SetKey("news_sale_id", "") ' Clear key
				CopyRecord = False
			End If
			If CopyRecord Then
				news_sale.CurrentAction = "C" ' Copy Record
			Else
				news_sale.CurrentAction = "I" ' Display Blank Record
				Call LoadDefaultValues() ' Load default values
			End If
		End If

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Validate form if post back
		If ObjForm.GetValue("a_add")&"" <> "" Then
			If Not ValidateForm() Then
				news_sale.CurrentAction = "I" ' Form error, reset action
				news_sale.EventCancelled = True ' Event cancelled
				Call RestoreFormValues() ' Restore form values
				FailureMessage = gsFormError
			End If
		End If

		' Perform action based on action code
		Select Case news_sale.CurrentAction
			Case "I" ' Blank record, no action required
			Case "C" ' Copy an existing record
				If Not LoadRow() Then ' Load record based on key
					If FailureMessage = "" Then FailureMessage = Language.Phrase("NoRecord") ' No record found
					Call Page_Terminate("pom_news_salelist.asp") ' No matching record, return to list
				End If
			Case "A" ' Add new record
				news_sale.SendEmail = True ' Send email on add success
				If AddRow(OldRecordset) Then ' Add successful
					SuccessMessage = Language.Phrase("AddSuccess") ' Set up success message
					Dim sReturnUrl
					sReturnUrl = news_sale.ReturnUrl
					If ew_GetPageName(sReturnUrl) = "pom_news_saleview.asp" Then sReturnUrl = news_sale.ViewUrl("") ' View paging, return to view page with keyurl directly
					Call Page_Terminate(sReturnUrl) ' Clean up and return
				Else
					news_sale.EventCancelled = True ' Event cancelled
					Call RestoreFormValues() ' Add failed, restore form values
				End If
		End Select

		' Render row based on row type
		news_sale.RowType = EW_ROWTYPE_ADD ' Render add type

		' Render row
		Call news_sale.ResetAttrs()
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
		news_sale.news_sale_pdf.CurrentValue = Null
		news_sale.news_sale_pdf.OldValue = news_sale.news_sale_pdf.CurrentValue
		news_sale.news_sale_title.CurrentValue = Null
		news_sale.news_sale_title.OldValue = news_sale.news_sale_title.CurrentValue
		news_sale.start_date.CurrentValue = Null
		news_sale.start_date.OldValue = news_sale.start_date.CurrentValue
		news_sale.end_date.CurrentValue = Null
		news_sale.end_date.OldValue = news_sale.end_date.CurrentValue
	End Function

	' -----------------------------------------------------------------
	' Load form values
	'
	Function LoadFormValues()

		' Load values from form
		If Not news_sale.news_sale_pdf.FldIsDetailKey Then news_sale.news_sale_pdf.FormValue = ObjForm.GetValue("x_news_sale_pdf")
		If Not news_sale.news_sale_title.FldIsDetailKey Then news_sale.news_sale_title.FormValue = ObjForm.GetValue("x_news_sale_title")
		If Not news_sale.start_date.FldIsDetailKey Then news_sale.start_date.FormValue = ObjForm.GetValue("x_start_date")
		If Not news_sale.start_date.FldIsDetailKey Then news_sale.start_date.CurrentValue = ew_UnFormatDateTime(news_sale.start_date.CurrentValue, 8)
		If Not news_sale.end_date.FldIsDetailKey Then news_sale.end_date.FormValue = ObjForm.GetValue("x_end_date")
		If Not news_sale.end_date.FldIsDetailKey Then news_sale.end_date.CurrentValue = ew_UnFormatDateTime(news_sale.end_date.CurrentValue, 8)
	End Function

	' -----------------------------------------------------------------
	' Restore form values
	'
	Function RestoreFormValues()
		Call LoadOldRecord()
		news_sale.news_sale_pdf.CurrentValue = news_sale.news_sale_pdf.FormValue
		news_sale.news_sale_title.CurrentValue = news_sale.news_sale_title.FormValue
		news_sale.start_date.CurrentValue = news_sale.start_date.FormValue
		news_sale.start_date.CurrentValue = ew_UnFormatDateTime(news_sale.start_date.CurrentValue, 8)
		news_sale.end_date.CurrentValue = news_sale.end_date.FormValue
		news_sale.end_date.CurrentValue = ew_UnFormatDateTime(news_sale.end_date.CurrentValue, 8)
	End Function

	' -----------------------------------------------------------------
	' Load row based on key values
	'
	Function LoadRow()
		Dim RsRow, sSql, sFilter
		sFilter = news_sale.KeyFilter

		' Call Row Selecting event
		Call news_sale.Row_Selecting(sFilter)

		' Load sql based on filter
		news_sale.CurrentFilter = sFilter
		sSql = news_sale.SQL
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
		Call news_sale.Row_Selected(RsRow)
		news_sale.news_sale_id.DbValue = RsRow("news_sale_id")
		news_sale.news_sale_pdf.DbValue = RsRow("news_sale_pdf")
		news_sale.news_sale_title.DbValue = RsRow("news_sale_title")
		news_sale.start_date.DbValue = RsRow("start_date")
		news_sale.end_date.DbValue = RsRow("end_date")
	End Sub

	' Load DbValue from recordset
	Sub LoadDbValues(Rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		If Rs.Eof Then Exit Sub
		news_sale.news_sale_id.m_DbValue = Rs("news_sale_id")
		news_sale.news_sale_pdf.m_DbValue = Rs("news_sale_pdf")
		news_sale.news_sale_title.m_DbValue = Rs("news_sale_title")
		news_sale.start_date.m_DbValue = Rs("start_date")
		news_sale.end_date.m_DbValue = Rs("end_date")
	End Sub

	' Load old record
	Function LoadOldRecord()

		' Load key values from Session
		Dim bValidKey
		bValidKey = True
		If news_sale.GetKey("news_sale_id")&"" <> "" Then
			news_sale.news_sale_id.CurrentValue = news_sale.GetKey("news_sale_id") ' news_sale_id
		Else
			bValidKey = False
		End If

		' Load old recordset
		If bValidKey Then
			news_sale.CurrentFilter = news_sale.KeyFilter
			Dim sSql
			sSql = news_sale.SQL
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

		Call news_sale.Row_Rendering()

		' ---------------------------------------
		'  Common render codes for all row types
		' ---------------------------------------
		' news_sale_id
		' news_sale_pdf
		' news_sale_title
		' start_date
		' end_date
		' -----------
		'  View  Row
		' -----------

		If news_sale.RowType = EW_ROWTYPE_VIEW Then ' View row

			' news_sale_id
			news_sale.news_sale_id.ViewValue = news_sale.news_sale_id.CurrentValue
			news_sale.news_sale_id.ViewCustomAttributes = ""

			' news_sale_pdf
			news_sale.news_sale_pdf.ViewValue = news_sale.news_sale_pdf.CurrentValue
			news_sale.news_sale_pdf.ViewCustomAttributes = ""

			' news_sale_title
			news_sale.news_sale_title.ViewValue = news_sale.news_sale_title.CurrentValue
			news_sale.news_sale_title.ViewCustomAttributes = ""

			' start_date
			news_sale.start_date.ViewValue = news_sale.start_date.CurrentValue
			news_sale.start_date.ViewCustomAttributes = ""

			' end_date
			news_sale.end_date.ViewValue = news_sale.end_date.CurrentValue
			news_sale.end_date.ViewCustomAttributes = ""

			' View refer script
			' news_sale_pdf

			news_sale.news_sale_pdf.LinkCustomAttributes = ""
			news_sale.news_sale_pdf.HrefValue = ""
			news_sale.news_sale_pdf.TooltipValue = ""

			' news_sale_title
			news_sale.news_sale_title.LinkCustomAttributes = ""
			news_sale.news_sale_title.HrefValue = ""
			news_sale.news_sale_title.TooltipValue = ""

			' start_date
			news_sale.start_date.LinkCustomAttributes = ""
			news_sale.start_date.HrefValue = ""
			news_sale.start_date.TooltipValue = ""

			' end_date
			news_sale.end_date.LinkCustomAttributes = ""
			news_sale.end_date.HrefValue = ""
			news_sale.end_date.TooltipValue = ""

		' ---------
		'  Add Row
		' ---------

		ElseIf news_sale.RowType = EW_ROWTYPE_ADD Then ' Add row

			' news_sale_pdf
			news_sale.news_sale_pdf.EditCustomAttributes = ""
			news_sale.news_sale_pdf.EditValue = ew_HtmlEncode(news_sale.news_sale_pdf.CurrentValue)
			news_sale.news_sale_pdf.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_sale.news_sale_pdf.FldCaption))

			' news_sale_title
			news_sale.news_sale_title.EditCustomAttributes = ""
			news_sale.news_sale_title.EditValue = ew_HtmlEncode(news_sale.news_sale_title.CurrentValue)
			news_sale.news_sale_title.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_sale.news_sale_title.FldCaption))

			' start_date
			news_sale.start_date.EditCustomAttributes = ""
			news_sale.start_date.EditValue = ew_HtmlEncode(news_sale.start_date.CurrentValue)
			news_sale.start_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_sale.start_date.FldCaption))

			' end_date
			news_sale.end_date.EditCustomAttributes = ""
			news_sale.end_date.EditValue = ew_HtmlEncode(news_sale.end_date.CurrentValue)
			news_sale.end_date.PlaceHolder = ew_HtmlEncode(ew_RemoveHtml(news_sale.end_date.FldCaption))

			' Edit refer script
			' news_sale_pdf

			news_sale.news_sale_pdf.HrefValue = ""

			' news_sale_title
			news_sale.news_sale_title.HrefValue = ""

			' start_date
			news_sale.start_date.HrefValue = ""

			' end_date
			news_sale.end_date.HrefValue = ""
		End If
		If news_sale.RowType = EW_ROWTYPE_ADD Or news_sale.RowType = EW_ROWTYPE_EDIT Or news_sale.RowType = EW_ROWTYPE_SEARCH Then ' Add / Edit / Search row
			Call news_sale.SetupFieldTitles()
		End If

		' Call Row Rendered event
		If news_sale.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call news_sale.Row_Rendered()
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
		news_sale.CurrentFilter = sFilter
		sSql = news_sale.SQL
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

		' Field news_sale_pdf
		Call news_sale.news_sale_pdf.SetDbValue(Rs, news_sale.news_sale_pdf.CurrentValue, Null, False)

		' Field news_sale_title
		Call news_sale.news_sale_title.SetDbValue(Rs, news_sale.news_sale_title.CurrentValue, Null, False)

		' Field start_date
		Call news_sale.start_date.SetDbValue(Rs, news_sale.start_date.CurrentValue, Null, False)

		' Field end_date
		Call news_sale.end_date.SetDbValue(Rs, news_sale.end_date.CurrentValue, Null, False)

		' Check recordset update error
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			Rs.Close
			Set Rs = Nothing
			AddRow = False
			Exit Function
		End If

		' Call Row Inserting event
		bInsertRow = news_sale.Row_Inserting(RsOld, Rs)
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
			ElseIf news_sale.CancelMessage <> "" Then
				FailureMessage = news_sale.CancelMessage
				news_sale.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("InsertCancelled")
			End If
			AddRow = False
		End If
		Rs.Close
		Set Rs = Nothing
		If AddRow Then
			news_sale.news_sale_id.DbValue = RsNew("news_sale_id")
		End If
		If AddRow Then

			' Call Row Inserted event
			Call news_sale.Row_Inserted(RsOld, RsNew)
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
		Call Breadcrumb.Add("list", news_sale.TableVar, "pom_news_salelist.asp", news_sale.TableVar, True)
		PageId = ew_IIf(news_sale.CurrentAction = "C", "Copy", "Add")
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
