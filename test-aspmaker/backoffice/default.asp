<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim default
Set default = New cdefault
Set Page = default

' Page init processing
default.Page_Init()

' Page main processing
default.Page_Main()
%>
<!--#include file="pom_header.asp"-->
<% default.ShowMessage %>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set default = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cdefault

	' Page ID
	Public Property Get PageID()
		PageID = "default"
	End Property

	' Project ID
	Public Property Get ProjectID()
		ProjectID = "{324ED72D-DE20-46F7-B12E-7AF8CE8711A6}"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "default"
	End Property

	' Page Name
	Public Property Get PageName()
		PageName = ew_CurrentPage()
	End Property

	' Page Url
	Public Property Get PageUrl()
		PageUrl = ew_CurrentPage() & "?"
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

		' Initialize user table object
		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "default"

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
	' Page main processing
	Sub Page_Main()
		If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_z40newslist.asp") ' Exit and go to default page
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_adminslist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_bannerlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_banner_logo_01list.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_banner_logo_01_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_banner_logo_02list.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_banner_logo_02_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_banner_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_companylist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_company_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_departmentlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_department_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_e_librarylist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_e_library_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_eventcalendarlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_eventcalendar_pdf_filelist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_eventcalendar_pdf_file_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_eventcalendar_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_homepagelist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_joblist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_job_filelist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_job_file_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_job_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_journallist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_newslist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_news_pdf_filelist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_news_pdf_file_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_news_salelist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_news_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_officelist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_office_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_personlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_person_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_researchlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_research_pdf_filelist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_research_pdf_file_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_research_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_sys_admin_menulist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_sys_menulist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_vehicle_recordlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_vehicle_record_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_videolist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_video_thlist.asp")
		End If
		If Security.IsLoggedIn() Then
			Call Page_Terminate("pom_query2list.asp")
		End If
		If Security.IsLoggedIn() Then
			FailureMessage = Language.Phrase("NoPermission") & "<br><br><a href=""pom_logout.asp"">" & Language.Phrase("BackToLogin") & "</a>"
		Else
			Call Page_Terminate("pom_login.asp") ' Exit and go to login page
		End If
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
	' typ = ""|"success"|"failure"
	Sub Message_Showing(msg, typ)

		' Example:
		'If typ = "success" Then msg = "your success message"

	End Sub
End Class
%>
