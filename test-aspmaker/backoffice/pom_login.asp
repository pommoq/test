<%@ CodePage="65001" LCID="1054" %>
<%
Response.Buffer = EW_RESPONSE_BUFFER
%>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="pom_adminsinfo.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="pom_userfn11.asp"-->
<% Session.Timeout = 20 %>
<% Call ew_Header(False, EW_CHARSET) %>
<%

' Define page object
Dim login
Set login = New clogin
Set Page = login

' Page init processing
login.Page_Init()

' Page main processing
login.Page_Main()

' Global Page Rendering event (in userfn*.asp)
Page_Rendering()

' Page Rendering event
login.Page_Render()
%>
<!--#include file="pom_header.asp"-->
<script type="text/javascript">
// Write your client script here, no need to add script tags.
</script>
<script type="text/javascript">
var flogin = new ew_Form("flogin");
// Validate function
flogin.Validate = function()
{
	var fobj = this.Form;
	if (!this.ValidateRequired)
		return true; // Ignore validation
	if (!ew_HasValue(fobj.username))
		return this.OnError(fobj.username, ewLanguage.Phrase("EnterUid"));
	if (!ew_HasValue(fobj.password))
		return this.OnError(fobj.password, ewLanguage.Phrase("EnterPwd"));
	// Call Form Custom Validate event
	if (!this.Form_CustomValidate(fobj)) return false;
	return true;
}
// Form_CustomValidate function
flogin.Form_CustomValidate = 
 function(fobj) { // DO NOT CHANGE THIS LINE!
 	// Your custom validation code here, return false if invalid. 
 	return true;
 }
// Requires js validation
<% If EW_CLIENT_VALIDATE Then %>
flogin.ValidateRequired = true;
<% Else %>
flogin.ValidateRequired = false;
<% End If %>
</script>
<% Breadcrumb.Render() %>
<% login.ShowPageHeader() %>
<% login.ShowMessage %>
<form name="flogin" id="flogin" class="ewForm form-horizontal" action="<%= ew_CurrentPage %>" method="post">
<div class="ewLoginContent">
	<div class="control-group">
		<label class="control-label" for="username"><%= Language.Phrase("Username") %></label>
		<div class="controls"><input type="text" name="username" id="username" class="input-large" value="<%= ew_HtmlEncode(login.Username) %>" placeholder="<%= ew_HtmlEncode(Language.Phrase("Username")) %>"></div>
	</div>
	<div class="control-group">
		<label class="control-label" for="password"><%= Language.Phrase("Password") %></label>
		<div class="controls"><input type="password" name="password" id="password" class="input-large" placeholder="<%= ew_HtmlEncode(Language.Phrase("Password")) %>"></div>
	</div>
	<div class="control-group">
		<div class="controls">
		<label class="radio ewRadio" style="white-space: nowrap;"><input type="radio" name="type" value="a"<% If login.LoginType = "a" Then %> checked="checked"<% End If %>><%= Language.Phrase("AutoLogin") %></label>
		<label class="radio ewRadio" style="white-space: nowrap;"><input type="radio" name="type" value="u"<% If login.LoginType = "u" Then %>  checked="checked"<% End If %>><%= Language.Phrase("SaveUserName") %></label>
		<label class="radio ewRadio" style="white-space: nowrap;"><input type="radio" name="type" value=""<% If login.LoginType = "" Then %> checked="checked"<% End If %>><%= Language.Phrase("AlwaysAsk") %></label>
		</div>
	</div>
	<div class="control-group">
		<div class="controls">
			<button class="btn btn-primary ewButton" name="btnsubmit" id="btnsubmit" type="submit"><%= Language.Phrase("Login") %></button>
		</div>
	</div>
</div>
</form>
<br>
<script type="text/javascript">
flogin.Init();
<% If EW_MOBILE_REFLOW And ew_IsMobile() Then %>
ew_Reflow();
<% End If %>
</script>
<%
login.ShowPageFooter()
If EW_DEBUG_ENABLED Then Response.Write ew_DebugMsg()
%>
<script type="text/javascript">
// Write your startup script here
// document.write("page loaded");
</script>
<!--#include file="pom_footer.asp"-->
<%

' Drop page object
Set login = Nothing
%>
<%

' -----------------------------------------------------------------
' Page Class
'
Class clogin

	' Page ID
	Public Property Get PageID()
		PageID = "login"
	End Property

	' Project ID
	Public Property Get ProjectID()
		ProjectID = "{324ED72D-DE20-46F7-B12E-7AF8CE8711A6}"
	End Property

	' Page Object Name
	Public Property Get PageObjName()
		PageObjName = "login"
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
		IsPageRequest = True
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
		If IsEmpty(admins) Then Set admins = New cadmins

		' Initialize form object
		Set ObjForm = Nothing

		' Intialize page id (for backward compatibility)
		EW_PAGE_ID = "login"

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

	Dim Username
	Dim LoginType

	' -----------------------------------------------------------------
	' Page main processing
	'
	Sub Page_Main()
		Dim bValidate, bValidPwd
		Dim sPassword
		Dim sLastUrl
		Set Breadcrumb = New cBreadcrumb
		Call Breadcrumb.Add("login", "LoginPage", ew_CurrentUrl, "", True)
		sLastUrl = Security.LastUrl ' Get Last Url
		If sLastUrl = "" Then sLastUrl = "default.asp"
		If IsLoggingIn() Then
			Username = Session(EW_SESSION_USER_PROFILE_USER_NAME)
			sPassword = Session(EW_SESSION_USER_PROFILE_PASSWORD)
			LoginType = Session(EW_SESSION_USER_PROFILE_LOGIN_TYPE)
			bValidPwd = Security.ValidateUser(Username, sPassword, False)
			If bValidPwd Then
				Session(EW_SESSION_USER_PROFILE_USER_NAME) = ""
				Session(EW_SESSION_USER_PROFILE_PASSWORD) = ""
				Session(EW_SESSION_USER_PROFILE_LOGIN_TYPE) = ""
			End If
		Else
			If Not Security.IsLoggedIn() Then Call Security.AutoLogin()
			Username = "" ' Initialize
			If Request.Form <> "" Then

				' Setup variables
				Username = ew_RemoveXSS(Request.Form("username"))
				sPassword = ew_RemoveXSS(Request.Form("password"))
				LoginType = LCase(ew_RemoveXSS(Request.Form("type")))
			End If
			If Username <> "" Then
				bValidate = ValidateForm(Username, sPassword)
				If Not bValidate Then
					FailureMessage = gsFormError
				End If
				Session(EW_SESSION_USER_PROFILE_USER_NAME) = Username ' Save login user name
				Session(EW_SESSION_USER_PROFILE_LOGIN_TYPE) = LoginType ' Save login type
			Else
				If Security.IsLoggedIn() Then
					If FailureMessage = "" Then Page_Terminate(sLastUrl) ' Return to last accessed page
				End If
				bValidate = False

				' Restore settings
				Username = Request.Cookies(EW_PROJECT_NAME)("username")
				If Request.Cookies(EW_PROJECT_NAME)("autologin") = "autologin" Then
					LoginType = "a"
				ElseIf Request.Cookies(EW_PROJECT_NAME)("autologin") = "rememberusername" Then
					LoginType = "u"
				Else
					LoginType = ""
				End If
			End If
			bValidPwd = False
			If bValidate Then

				' Call logging in event
				bValidate = User_LoggingIn(Username, sPassword)
				If bValidate Then
					bValidPwd = Security.ValidateUser(Username, sPassword, False) ' Manual login
					If Not bValidPwd Then
						If FailureMessage = "" Then FailureMessage = Language.Phrase("InvalidUidPwd") ' Invalid user id/password
					End If
				Else
					If FailureMessage = "" Then FailureMessage = Language.Phrase("LoginCancelled") ' Login cancelled
				End If
			End If
		End If
		If bValidPwd Then

			' Write cookies
			If LoginType = "a" Then ' Auto login
				Response.Cookies(EW_PROJECT_NAME)("autologin") = "autologin" ' Set up autologin cookies
				Response.Cookies(EW_PROJECT_NAME)("username") = Username ' Set up user name cookies
				Response.Cookies(EW_PROJECT_NAME)("password") = ew_Encode(ew_Encrypt(sPassword, EW_RANDOM_KEY)) ' Set up password cookies
			ElseIf LoginType = "u" Then ' Remember user name
				Response.Cookies(EW_PROJECT_NAME)("autologin") = "rememberusername" ' Set up remember user name cookies
				Response.Cookies(EW_PROJECT_NAME)("username") = Username ' Set up user name cookies
			Else
				Response.Cookies(EW_PROJECT_NAME)("autologin") = "" ' Clear autologin cookies
			End If
			Response.Cookies(EW_PROJECT_NAME).Expires = DateAdd("d", EW_COOKIE_EXPIRY_TIME, Now)

			' Call loggedin event
			Call User_LoggedIn(Username)
			Call Page_Terminate(sLastUrl) ' Return to last accessed url
		ElseIf Username <> "" And sPassword <> "" Then

			' Call user login error event
			Call User_LoginError(Username, sPassword)
		End If
	End Sub

	' -----------------------------------------------------------------
	' Validate form
	'
	Function ValidateForm(usr, pwd)

		' Initialize
		gsFormError = ""

		' Check if validation required
		If Not EW_SERVER_VALIDATE Then
			ValidateForm = True
			Exit Function
		End If
		If usr = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterUid"))
		End If
		If pwd = "" Then
			Call ew_AddMessage(gsFormError, Language.Phrase("EnterPwd"))
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

	' User Logging In event
	Function User_LoggingIn(usr, pwd)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
		' To cancel, set return value to False

		User_LoggingIn = True
	End Function

	' User Logged In event
	Sub User_LoggedIn(usr)

		' Response.Write "User Logged In"
	End Sub

	' User Login Error event
	Sub User_LoginError(usr, pwd)

		' Response.Write "User Login Error"
	End Sub

	' Form Custom Validate event
	Function Form_CustomValidate(CustomError)

		'Return error message in CustomError
		Form_CustomValidate = True
	End Function
End Class
%>
