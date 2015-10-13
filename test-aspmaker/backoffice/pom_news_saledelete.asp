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
Dim news_sale_delete
Set news_sale_delete = New cnews_sale_delete
Set Page = news_sale_delete

' Page init processing
news_sale_delete.Page_Init()

' Page main processing
news_sale_delete.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
news_sale_delete.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Page object
var news_sale_delete = new ew_Page("news_sale_delete");
news_sale_delete.PageID = "delete"; // Page ID
var EW_PAGE_ID = news_sale_delete.PageID; // For backward compatibility
// Form object
var fnews_saledelete = new ew_Form("fnews_saledelete");
// Form_CustomValidate event
fnews_saledelete.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Use JavaScript validation or not
<% If EW_CLIENT_VALIDATE Then %>
fnews_saledelete.ValidateRequired = true; // Use JavaScript validation
<% Else %>
fnews_saledelete.ValidateRequired = false; // No JavaScript validation
<% End If %>
// Dynamic selection lists
// Form object for search
</script>
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<%

' Load records for display
Set news_sale_delete.Recordset = news_sale_delete.LoadRecordset()
news_sale_delete.TotalRecs = news_sale_delete.Recordset.RecordCount ' Get record count
If news_sale_delete.TotalRecs <= 0 Then ' No record found, exit
	news_sale_delete.Recordset.Close
	Set news_sale_delete.Recordset = Nothing
	Call news_sale_delete.Page_Terminate("pom_news_salelist.asp") ' Return to list
End If
%>
<% If news_sale.Export = "" Then %>
<% Breadcrumb.Render() %>
<% End If %>
<% news_sale_delete.ShowPageHeader() %>
<% news_sale_delete.ShowMessage %>
<form name="fnews_saledelete" id="fnews_saledelete" class="ewForm form-inline" action="<%= ew_CurrentPage() %>" method="post">
<input type="hidden" name="t" value="news_sale">
<input type="hidden" name="a_delete" id="a_delete" value="D">
<% For i = 0 to UBound(news_sale_delete.RecKeys) %>
<input type="hidden" name="key_m" id="key_m" value="<%= ew_HtmlEncode(ew_GetKeyValue(news_sale_delete.RecKeys(i))) %>">
<% Next %>
<table class="ewGrid"><tr><td class="ewGridContent">
<div class="ewGridMiddlePanel">
<table id="tbl_news_saledelete" class="ewTable ewTableSeparate">
<%= news_sale.TableCustomInnerHtml %>
	<thead>
	<tr class="ewTableHeader">
<% If news_sale.news_sale_id.Visible Then ' news_sale_id %>
		<td><span id="elh_news_sale_news_sale_id" class="news_sale_news_sale_id"><%= news_sale.news_sale_id.FldCaption %></span></td>
<% End If %>
<% If news_sale.news_sale_pdf.Visible Then ' news_sale_pdf %>
		<td><span id="elh_news_sale_news_sale_pdf" class="news_sale_news_sale_pdf"><%= news_sale.news_sale_pdf.FldCaption %></span></td>
<% End If %>
<% If news_sale.news_sale_title.Visible Then ' news_sale_title %>
		<td><span id="elh_news_sale_news_sale_title" class="news_sale_news_sale_title"><%= news_sale.news_sale_title.FldCaption %></span></td>
<% End If %>
<% If news_sale.start_date.Visible Then ' start_date %>
		<td><span id="elh_news_sale_start_date" class="news_sale_start_date"><%= news_sale.start_date.FldCaption %></span></td>
<% End If %>
<% If news_sale.end_date.Visible Then ' end_date %>
		<td><span id="elh_news_sale_end_date" class="news_sale_end_date"><%= news_sale.end_date.FldCaption %></span></td>
<% End If %>
	</tr>
	</thead>
	<tbody>
<%
news_sale_delete.RecCnt = 0
news_sale_delete.RowCnt = 0
Do While (Not news_sale_delete.Recordset.Eof)
	news_sale_delete.RecCnt = news_sale_delete.RecCnt + 1
	news_sale_delete.RowCnt = news_sale_delete.RowCnt + 1

	' Set row properties
	Call news_sale.ResetAttrs()
	news_sale.RowType = EW_ROWTYPE_VIEW ' view

	' Get the field contents
	Call news_sale_delete.LoadRowValues(news_sale_delete.Recordset)

	' Render row
	Call news_sale_delete.RenderRow()
%>
	<tr<%= news_sale.RowAttributes %>>
<% If news_sale.news_sale_id.Visible Then ' news_sale_id %>
		<td<%= news_sale.news_sale_id.CellAttributes %>>
<span id="el<%= news_sale_delete.RowCnt %>_news_sale_news_sale_id" class="control-group news_sale_news_sale_id">
<span<%= news_sale.news_sale_id.ViewAttributes %>>
<%= news_sale.news_sale_id.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_sale.news_sale_pdf.Visible Then ' news_sale_pdf %>
		<td<%= news_sale.news_sale_pdf.CellAttributes %>>
<span id="el<%= news_sale_delete.RowCnt %>_news_sale_news_sale_pdf" class="control-group news_sale_news_sale_pdf">
<span<%= news_sale.news_sale_pdf.ViewAttributes %>>
<%= news_sale.news_sale_pdf.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_sale.news_sale_title.Visible Then ' news_sale_title %>
		<td<%= news_sale.news_sale_title.CellAttributes %>>
<span id="el<%= news_sale_delete.RowCnt %>_news_sale_news_sale_title" class="control-group news_sale_news_sale_title">
<span<%= news_sale.news_sale_title.ViewAttributes %>>
<%= news_sale.news_sale_title.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_sale.start_date.Visible Then ' start_date %>
		<td<%= news_sale.start_date.CellAttributes %>>
<span id="el<%= news_sale_delete.RowCnt %>_news_sale_start_date" class="control-group news_sale_start_date">
<span<%= news_sale.start_date.ViewAttributes %>>
<%= news_sale.start_date.ListViewValue %>
</span>
</span>
</td>
<% End If %>
<% If news_sale.end_date.Visible Then ' end_date %>
		<td<%= news_sale.end_date.CellAttributes %>>
<span id="el<%= news_sale_delete.RowCnt %>_news_sale_end_date" class="control-group news_sale_end_date">
<span<%= news_sale.end_date.ViewAttributes %>>
<%= news_sale.end_date.ListViewValue %>
</span>
</span>
</td>
<% End If %>
	</tr>
<%
	news_sale_delete.Recordset.MoveNext
Loop
news_sale_delete.Recordset.Close
Set news_sale_delete.Recordset = Nothing
%>
	</tbody>
</table>
</div>
</td></tr></table>
<div class="btn-group ewButtonGroup">
<button class="btn btn-primary ewButton" name="btnAction" id="btnAction" type="submit"><%= Language.Phrase("DeleteBtn") %></button>
</div>
</form>
<script type="text/javascript">
fnews_saledelete.Init();
</script>
<%
news_sale_delete.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your table-specific startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set news_sale_delete = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cnews_sale_delete

	' Page ID
	Public Property Get PageID()
		PageID = "delete"
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
		PageObjName = "news_sale_delete"
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
		EW_PAGE_ID = "delete"

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

	Dim TotalRecs
	Dim RecCnt
	Dim RecKeys
	Dim Recordset
	Dim StartRowCnt
	Dim RowCnt

	' Page main processing
	Sub Page_Main()
		Dim sFilter
		StartRowCnt = 1

		' Set up Breadcrumb
		SetupBreadcrumb()

		' Load Key Parameters
		RecKeys = news_sale.GetRecordKeys() ' Load record keys
		sFilter = news_sale.GetKeyFilter()
		If sFilter = "" Then
			Call Page_Terminate("pom_news_salelist.asp") ' Prevent SQL injection, return to list
		End If

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in news_sale class, news_saleinfo.asp

		news_sale.CurrentFilter = sFilter

		' Get action
		If Request.Form("a_delete").Count > 0 Then
			news_sale.CurrentAction = Request.Form("a_delete")
		Else
			news_sale.CurrentAction = "I"	' Display record
		End If
		Select Case news_sale.CurrentAction
			Case "D" ' Delete
				news_sale.SendEmail = True ' Send email on delete success
				If DeleteRows() Then ' Delete rows
					SuccessMessage = Language.Phrase("DeleteSuccess") ' Set up success message
					Call Page_Terminate(news_sale.ReturnUrl) ' Return to caller
				End If
		End Select
	End Sub

	' -----------------------------------------------------------------
	' Load recordset
	'
	Function LoadRecordset()

		' Call Recordset Selecting event
		Dim sFilter
		sFilter = news_sale.CurrentFilter
		Call news_sale.Recordset_Selecting(sFilter)
		news_sale.CurrentFilter = sFilter

		' Load list page sql
		Dim sSql
		sSql = news_sale.ListSQL
		Call ew_SetDebugMsg("LoadRecordset: " & sSql) ' Show SQL for debugging

		' Load recordset
		Dim RsRecordset
		Set RsRecordset = ew_LoadRecordset(sSql)

		' Call Recordset Selected event
		Call news_sale.Recordset_Selected(RsRecordset)
		Set LoadRecordset = RsRecordset
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
			' news_sale_id

			news_sale.news_sale_id.LinkCustomAttributes = ""
			news_sale.news_sale_id.HrefValue = ""
			news_sale.news_sale_id.TooltipValue = ""

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
		End If

		' Call Row Rendered event
		If news_sale.RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call news_sale.Row_Rendered()
		End If
	End Sub

	'
	' Delete records based on current filter
	'
	Function DeleteRows()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim sKey, sThisKey, sKeyFld, arKeyFlds
		Dim sSql, RsDelete
		Dim RsOld, RsDetail
		DeleteRows = True
		sSql = news_sale.SQL
		Conn.BeginTrans
		Set RsDelete = Server.CreateObject("ADODB.Recordset")
		RsDelete.CursorLocation = EW_CURSORLOCATION
		RsDelete.Open sSql, Conn, 1, EW_RECORDSET_LOCKTYPE
		If Err.Number <> 0 Then
			FailureMessage = Err.Description
			RsDelete.Close
			Set RsDelete = Nothing
			DeleteRows = False
			Exit Function
		ElseIf RsDelete.Eof Then
			FailureMessage = Language.Phrase("NoRecord") ' No record found
			RsDelete.Close
			Set RsDelete = Nothing
			DeleteRows = False
			Exit Function
		End If

		' Clone old recordset object
		Set RsOld = ew_CloneRs(RsDelete)

		' Call row deleting event
		If DeleteRows Then
			RsDelete.MoveFirst
			Do While Not RsDelete.Eof
				DeleteRows = news_sale.Row_Deleting(RsDelete)
				If Not DeleteRows Then Exit Do
				RsDelete.MoveNext
			Loop
			RsDelete.MoveFirst
		End If
		If DeleteRows Then
			sKey = ""
			RsDelete.MoveFirst
			Do While Not RsDelete.Eof
				sThisKey = ""
				If sThisKey <> "" Then sThisKey = sThisKey & EW_COMPOSITE_KEY_SEPARATOR
				sThisKey = sThisKey & RsDelete("news_sale_id")
				Call LoadDbValues(RsDelete)
				If DeleteRows Then
					RsDelete.Delete
				End If
				If Err.Number <> 0 Or Not DeleteRows Then
					If Err.Description <> "" Then FailureMessage = Err.Description ' Set up error message
					DeleteRows = False
					Exit Do
				End If
				If sKey <> "" Then sKey = sKey & ", "
				sKey = sKey & sThisKey
				RsDelete.MoveNext
			Loop
		Else

			' Set up error message
			If SuccessMessage <> "" Or FailureMessage <> "" Then

				' Use the message, do nothing
			ElseIf news_sale.CancelMessage <> "" Then
				FailureMessage = news_sale.CancelMessage
				news_sale.CancelMessage = ""
			Else
				FailureMessage = Language.Phrase("DeleteCancelled")
			End If
		End If
		If DeleteRows Then
			Conn.CommitTrans ' Commit the changes
			If Err.Number <> 0 Then
				FailureMessage = Err.Description
				DeleteRows = False ' Delete failed
			End If
		Else
			Conn.RollbackTrans ' Rollback changes
		End If
		RsDelete.Close
		Set RsDelete = Nothing

		' Call row deleting event
		If DeleteRows Then
			RsOld.MoveFirst
			Do While Not RsOld.Eof
				Call news_sale.Row_Deleted(RsOld)
				RsOld.MoveNext
			Loop
		End If
		RsOld.Close
		Set RsOld = Nothing
	End Function

	' Set up Breadcrumb
	Sub SetupBreadcrumb()
		Dim PageId, url
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("list", news_sale.TableVar, "pom_news_salelist.asp", news_sale.TableVar, True)
		PageId = "delete"
		Call Breadcrumb.Add("delete", PageId, ew_CurrentUrl, "", False)
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
End Class
%>
