<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_newsinfo.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% If ew_IsHttps() Then Call ew_Header(True, EW_CHARSET) Else Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim news_news_img_blobview
Set news_news_img_blobview = New cnews_news_img_blobview
Set Page = news_news_img_blobview

' Page init processing
news_news_img_blobview.Page_Init()

' Page main processing
news_news_img_blobview.Page_Main()
%>
<%

' Drop page object
Set news_news_img_blobview = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class cnews_news_img_blobview

	' Page ID
	Public Property Get PageID()
		PageID = "blobview"
	End Property

	' Project ID
	Public Property Get ProjectID()
		ProjectID = "{324ED72D-DE20-46F7-B12E-7AF8CE8711A6}"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "news_news_img_blobview"
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

		' Initialize table object
		If IsEmpty(news) Then Set news = New cnews
		Set Table = news

		' Initialize urls
		' Initialize other table object

		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "blobview"

		' Initialize table name (for backward compatibility)
		EW_TABLE_NAME = "news"

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
		Dim sRedirectUrl
		sReDirectUrl = url
		Call Page_Redirecting(sReDirectUrl)
		If Not (Conn Is Nothing) Then Conn.Close ' Close Connection
		Set Conn = Nothing
		Set Security = Nothing
		Set news = Nothing
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
	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		Dim sFilter, objBinary

		' Get key
		If Request.QueryString("news_id").Count > 0 Then
			news.news_id.QueryStringValue = Request.QueryString("news_id")
		Else
			Call Page_Terminate("") ' Clean up
			Response.End ' Exit
		End If
		Set objBinary = New cUpload

		' Show thumbnail
		Dim bShowThumbnail, iThumbnailWidth, iThumbnailHeight, iInterpolation
		bShowThumbnail = (Request.QueryString("showthumbnail") = "1")
		If Request.QueryString("thumbnailwidth").Count <= 0 And Request.QueryString("thumbnailheight").Count <= 0 Then
			iThumbnailWidth = EW_THUMBNAIL_DEFAULT_WIDTH ' Set default width
			iThumbnailHeight = EW_THUMBNAIL_DEFAULT_HEIGHT ' Set default height
		Else
			If Request.QueryString("thumbnailwidth").Count > 0 Then
				iThumbnailWidth = Request.QueryString("thumbnailwidth")
				If Not IsNumeric(iThumbnailWidth) Or iThumbnailWidth < 0 Then iThumbnailWidth = 0
			End If
			If Request.QueryString("thumbnailheight").Count > 0 Then
				iThumbnailHeight = Request.QueryString("thumbnailheight")
				If Not IsNumeric(iThumbnailHeight) Or iThumbnailHeight < 0 Then iThumbnailHeight = 0
			End If
		End If
		If Request.QueryString("interpolation").Count > 0 Then
			iInterpolation = Request.QueryString("interpolation")
			If Not IsNumeric(iInterpolation) Or iInterpolation < 0 Or iInterpolation > 2 Then iInterpolation = EW_THUMBNAIL_DEFAULT_INTERPOLATION ' Set Default
		Else
			iInterpolation = EW_THUMBNAIL_DEFAULT_INTERPOLATION
		End If
		sFilter = news.KeyFilter

		' Set up filter (Sql Where Clause) and get Return Sql
		' Sql constructor in news class, newsinfo.asp

		Dim sSql, Recordset
		news.CurrentFilter = sFilter
		sSql = news.SQL
		Set Recordset = Server.CreateObject("ADODB.Recordset")
		Recordset.Open sSql, Conn
		If Not Recordset.Eof Then
			If Response.Buffer Then Response.Clear
			Dim filesize, filetype, filename
			filename = Recordset("news_img")
			objBinary.Value = Recordset("news_img").Value
			If bShowThumbnail Then
				Call objBinary.Resize(iThumbnailWidth, iThumbnailHeight, iInterpolation)
			End If

			'Response.ContentType = "image/bmp"
			Response.ContentType = ew_ContentType(LeftB(objBinary.Value,11), "")
			If filename&"" <> "" Then
				Response.AddHeader "Content-Disposition", "attachment; filename=""" & filename & """"
			End If
			Response.BinaryWrite objBinary.Value
		End If
		Recordset.Close
		Set Recordset = Nothing
		Set objBinary = Nothing
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
