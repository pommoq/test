' Compatibility codes for ASP Report Maker
Const EW_PROJECT_NAME = "project1" ' Project Name
Dim EW_CONFIG_FILE_FOLDER
EW_CONFIG_FILE_FOLDER = EW_PROJECT_NAME & "" ' Config file name
Const EW_PROJECT_ID = "{324ED72D-DE20-46F7-B12E-7AF8CE8711A6}" ' Project ID (GUID)
Dim EW_RELATED_PROJECT_ID
Dim EW_RELATED_LANGUAGE_FOLDER
Const EW_MAX_EMAIL_RECIPIENT = 3
' Auto suggest max entries
Const EW_AUTO_SUGGEST_MAX_ENTRIES = 10
' Upload max file size / thumbnail width and height
Const EW_MAX_FILE_SIZE = 2000000 ' Max file size
Const EW_UPLOAD_THUMBNAIL_WIDTH = 200 ' Temporary thumbnail max width
Const EW_UPLOAD_THUMBNAIL_HEIGHT = 200 ' Temporary thumbnail max height
' Language settings
Const EW_LANGUAGE_FOLDER = "lang/"
Dim EW_LANGUAGE_FILE(0)
EW_LANGUAGE_FILE(0) = Array("en", "", "english.xml")
Const EW_LANGUAGE_DEFAULT_ID = "en"
Dim EW_SESSION_LANGUAGE_FILE_CACHE
EW_SESSION_LANGUAGE_FILE_CACHE = EW_PROJECT_NAME & "_LanguageFile_9hPcPj4spse46Spe" ' Language File Cache
Dim EW_SESSION_LANGUAGE_CACHE
EW_SESSION_LANGUAGE_CACHE = EW_PROJECT_NAME & "_Language_9hPcPj4spse46Spe" ' Language Cache
Dim EW_SESSION_LANGUAGE_ID
EW_SESSION_LANGUAGE_ID = EW_PROJECT_NAME & "_LanguageId" ' Language ID
' Css file name
Const EW_PROJECT_STYLESHEET_FILENAME = "css/project1.css"
' Relative path of app root
Const EW_ROOT_RELATIVE_PATH = ".."
'
' *** DO NOT CHANGE BELOW
'
' Init language object
Set Language = new cLanguage
Call Language.LoadPhrases()
' ------------------------
'  Language class (begin)
'
Class cLanguage
	Dim LanguageId
	Dim objDOM
	Dim objDict
	Dim LanguageFolder
	Dim Key
	' Class initialize
	Private Sub Class_Initialize
		LanguageFolder = EW_LANGUAGE_FOLDER
	End Sub
	' Load phrases
	Public Sub LoadPhrases()
		' Set up file list
		LoadFileList()
		' Set up language id
		If Request.QueryString("language") <> "" Then
			LanguageId = Request.QueryString("language")
			Session(EW_SESSION_LANGUAGE_ID) = LanguageId
		ElseIf Session(EW_SESSION_LANGUAGE_ID) <> "" Then
			LanguageId = Session(EW_SESSION_LANGUAGE_ID)
		Else
			LanguageId = EW_LANGUAGE_DEFAULT_ID
		End If
		gsLanguage = LanguageId
		If EW_USE_DOM_XML Then
			Set objDOM = ew_CreateXmlDom()
			objDOM.async = False
		Else
			Set objDict = Server.CreateObject("Scripting.Dictionary")
		End If
		' Load current language
		Load(LanguageId)
	End Sub
	' Terminate
	Private Sub Class_Terminate()
		If EW_USE_DOM_XML Then
			Set objDOM = Nothing
		Else
			Set objDict = Nothing
		End If
	End Sub
	' Load language file list
	Private Sub LoadFileList()
		If IsArray(EW_LANGUAGE_FILE) Then
			For i = 0 to UBound(EW_LANGUAGE_FILE)
				EW_LANGUAGE_FILE(i)(1) = LoadFileDesc(Server.MapPath(LanguageFolder & EW_LANGUAGE_FILE(i)(2)))
			Next
		End If
	End Sub
	' Load language file description
	Private Function LoadFileDesc(File)
		LoadFileDesc = ""
		Set objDOM = ew_CreateXmlDom()
		objDOM.async = False
		objDOM.Load(File)
		If objDOM.ParseError.ErrorCode = 0 Then
			LoadFileDesc = GetNodeAtt(objDOM.documentElement, "desc")
		End If
	End Function
	' Load language file
	Private Sub Load(id)
		Dim sFileName
		sFileName = GetFileName(id)
		If sFileName = "" Then
			sFileName = GetFileName(EW_LANGUAGE_DEFAULT_ID)
		End If
		If sFileName = "" Then Exit Sub
		If EW_USE_DOM_XML Then
			objDOM.Load(sFileName)
			If objDOM.ParseError.ErrorCode = 0 Then
				objDOM.setProperty "SelectionLanguage", "XPath"
			End If
		Else
			XmlToCollection(sFileName)
		End If
		' Set up LCID from language file
		Dim langLCID
		If LocalePhrase("use_system_locale") = "1" Then
			langLCID = LocalePhrase("LCID")
			If langLCID <> "0" Then
				SetLocale(langLCID)
				EW_DECIMAL_POINT = Mid(FormatNumber(0.0,1,0,0,0),1,1) ' Get decimal point
				EW_THOUSANDS_SEP = Mid(FormatNumber(1000,0,0,0,-2),2,1) ' Get thousands sep
				If IsNumeric(EW_THOUSANDS_SEP) Then EW_THOUSANDS_SEP = ""
			End If
		Else
			EW_DECIMAL_POINT = LocalePhrase("decimal_point") ' Get decimal point
			EW_THOUSANDS_SEP = LocalePhrase("thousands_sep") ' Get thousands sep
			EW_CURRENCY_SYMBOL = LocalePhrase("currency_symbol") ' Get thousands sep
		End If
	End Sub
	Private Sub IterateNodes(Node)
		If Node.baseName = vbNullString Then Exit Sub
		Dim Index, Id, Client, ImageUrl, ImageWidth, ImageHeight
		If Node.nodeType = 1 And Node.baseName <> "ew-language" Then ' NODE_ELEMENT
			Id = ""
			If Node.attributes.length > 0 Then
				Id = Node.getAttribute("id")
			End If
			If Node.hasChildNodes Then
				Key = Key & Node.baseName & "/"
				If Id <> "" Then Key = Key & Id & "/"
			End If
			If Id <> "" And Not Node.hasChildNodes Then ' phrase
				Id = Node.baseName & "/" & Id
				Client = Node.getAttribute("client") & ""
				ImageUrl = Node.getAttribute("imageurl") & ""
				ImageWidth = Node.getAttribute("imagewidth") & ""
				ImageHeight = Node.getAttribute("imageheight") & ""
				If Id <> "" Then 
					objDict(Key & Id & "/attr/value") = Node.getAttribute("value") & ""
					If Client <> "" Then objDict(Key & Id & "/attr/client") = Client
					If ImageUrl <> "" Then objDict(Key & Id & "/attr/imageurl") = ImageUrl
					If ImageWidth <> "" Then objDict(Key & Id & "/attr/imagewidth") = ImageWidth
					If ImageHeight <> "" Then objDict(Key & Id & "/attr/imageheight") = ImageHeight
				End If
			End If
		End If
		If Node.hasChildNodes Then
			For Index = 0 To Node.childNodes.length - 1
				IterateNodes Node.childNodes(Index)
			Next
			Index	=	InStrRev(Key, "/"	&	Node.baseName & "/")
			If Index > 0	Then Key = Left(Key, Index)
		End If
	End Sub
	' Convert XML to Collection
	Private Sub XmlToCollection(File)
		Dim I, xmlr
		Key = "/"
		Set xmlr = ew_CreateXmlDom()
		xmlr.async = False
		xmlr.Load(File)
		For I = 0 To xmlr.childNodes.length - 1
			IterateNodes xmlr.childNodes(I)
		Next
		Set xmlr = Nothing
	End Sub
	' Get language file name
	Private Function GetFileName(Id)
		GetFileName = ""
		If IsArray(EW_LANGUAGE_FILE) Then
			For i = 0 to UBound(EW_LANGUAGE_FILE)
				If EW_LANGUAGE_FILE(i)(0) = Id Then
					GetFileName = Server.MapPath(LanguageFolder & EW_LANGUAGE_FILE(i)(2))
					Exit For
				End If
			Next
		End If
	End Function
	' Get node attribute
	Private Function GetNodeAtt(Node, Att)
		If Not (Node Is Nothing) Then
			GetNodeAtt = Node.getAttribute(Att)
		Else
			GetNodeAtt = ""
		End If
	End Function
	' Get dictionary attribute
	Private Function GetDictAtt(Att)
		If objDict.Exists(Att) Then
			GetDictAtt = objDict(Att)
		Else
			GetDictAtt = ""
		End If
	End Function
	' Get locale phrase
	Public Function LocalePhrase(Id)
		If EW_USE_DOM_XML Then
			LocalePhrase = GetNodeAtt(objDOM.SelectSingleNode("//locale/phrase[@id='" & LCase(Id) & "']"), "value")
		Else
			LocalePhrase = GetDictAtt("/locale/phrase/" & LCase(Id) & "/attr/value")
		End If  
	End Function
	' Set locale phrase
	Public Sub SetLocalePhrase(Id, Value)
		If Not EW_USE_DOM_XML Then
			objDict("/locale/phrase/" & LCase(Id) & "/attr/value") = Value
		End If
	End Sub
	' Get phrase
	Public Function Phrase(Id)
		Dim Text, ImageUrl, ImageWidth, ImageHeight, Style
		If EW_USE_DOM_XML Then
			ImageUrl = GetNodeAtt(objDOM.SelectSingleNode("//global/phrase[@id='" & LCase(Id) & "']"), "imageurl")
			ImageWidth = GetNodeAtt(objDOM.SelectSingleNode("//global/phrase[@id='" & LCase(Id) & "']"), "imagewidth")
			ImageHeight = GetNodeAtt(objDOM.SelectSingleNode("//global/phrase[@id='" & LCase(Id) & "']"), "imageheight")
			Text = GetNodeAtt(objDOM.SelectSingleNode("//global/phrase[@id='" & LCase(Id) & "']"), "value")
		Else
			ImageUrl = GetDictAtt("/global/phrase/" & LCase(Id) & "/attr/imageurl")
			ImageWidth = GetDictAtt("/global/phrase/" & LCase(Id) & "/attr/imagewidth")
			ImageHeight = GetDictAtt("/global/phrase/" & LCase(Id) & "/attr/imageheight")
			Text = GetDictAtt("/global/phrase/" & LCase(Id) & "/attr/value")
		End If
		If ImageUrl <> "" Then
			Style = ew_IIf(ImageWidth <> "", " width: " & ImageWidth & "px;", "")
			Style = Style & ew_IIf(ImageHeight <> "", " height: " & ImageHeight & "px;", "")
			Phrase = "<img data-phrase=""" & Id & """ src=""" & ew_HtmlEncode(ImageUrl) & """ style=""border: 0;" & Style & """ alt=""" & ew_HtmlEncode(Text) & """ title=""" & ew_HtmlEncode(Text) & """>"
		Else
			Phrase = Text
		End If
	End Function
	' Set phrase
	Public Sub SetPhrase(Id, Value)
		If Not EW_USE_DOM_XML Then
			objDict("/global/phrase/" & LCase(Id) & "/attr/value") = Value
		End If
	End Sub
	' Get project phrase
	Public Function ProjectPhrase(Id)
		If EW_USE_DOM_XML Then
			ProjectPhrase = GetNodeAtt(objDOM.SelectSingleNode("//project/phrase[@id='" & LCase(Id) & "']"), "value")
		Else
			ProjectPhrase = GetDictAtt("/project/phrase/" & LCase(Id) & "/attr/value")
		End If
	End Function
	' Set project phrase
	Public Sub SetProjectPhrase(Id, Value)
		If Not EW_USE_DOM_XML Then
			objDict("/project/phrase/" & LCase(Id) & "/attr/value") = Value
		End If
	End Sub
	' Get menu phrase
	Public Function MenuPhrase(MenuId, Id)
		If EW_USE_DOM_XML Then
			MenuPhrase = GetNodeAtt(objDOM.SelectSingleNode("//project/menu[@id='" & MenuId & "']/phrase[@id='" & LCase(Id) & "']"), "value")
		Else
			MenuPhrase = GetDictAtt("/project/menu/" & MenuId & "/phrase/" & LCase(Id) & "/attr/value")
		End If
	End Function
	' Set menu phrase
	Public Sub SetMenuPhrase(MenuId, Id, Value)
		If Not EW_USE_DOM_XML Then
			objDict("/project/menu/" & MenuId & "/phrase/" & LCase(Id) & "/attr/value") = Value
		End If
	End Sub
	' Get table phrase
	Public Function TablePhrase(TblVar, Id)
		If EW_USE_DOM_XML Then
			TablePhrase = GetNodeAtt(objDOM.SelectSingleNode("//project/table[@id='" & LCase(TblVar) & "']/phrase[@id='" & LCase(Id) & "']"), "value")
		Else
			TablePhrase = GetDictAtt("/project/table/" & LCase(TblVar) & "/phrase/" & LCase(Id) & "/attr/value")
		End If
	End Function
	' Set table phrase
	Public Sub SetTablePhrase(TblVar, Id, Value)
		If Not EW_USE_DOM_XML Then
			objDict("/project/table/" & LCase(TblVar) & "/phrase/" & LCase(Id) & "/attr/value") = Value
		End If
	End Sub
	' Get field phrase
	Public Function FieldPhrase(TblVar, FldVar, Id)
		If EW_USE_DOM_XML Then
			FieldPhrase = GetNodeAtt(objDOM.SelectSingleNode("//project/table[@id='" & LCase(TblVar) & "']/field[@id='" & LCase(FldVar) & "']/phrase[@id='" & LCase(Id) & "']"), "value")
		Else
			FieldPhrase = GetDictAtt("/project/table/" & LCase(TblVar) & "/field/" & LCase(FldVar) & "/phrase/" & LCase(Id) & "/attr/value")
		End If
	End Function
	' Set field phrase
	Public Sub SetFieldPhrase(TblVar, FldVar, Id, Value)
		If Not EW_USE_DOM_XML Then
			objDict("/project/table/" & LCase(TblVar) & "/field/" & LCase(FldVar) & "/phrase/" & LCase(Id) & "/attr/value") = Value
		End If
	End Sub
	' Output XML as JSON
	Public Function XmlToJSON(XPath)
		Dim Node, NodeList, Id, Value, Str
		Set NodeList = objDOM.selectNodes(XPath)
		Str = "{"
		For Each Node In NodeList
			Id = GetNodeAtt(Node, "id")
			Value = GetNodeAtt(Node, "value")
			Str = Str & """" & ew_JsEncode2(Id) & """:""" & ew_JsEncode2(Value) & ""","
		Next
		If Right(Str, 1) = "," Then Str = Left(Str, Len(Str)-1)
		Str = Str & "}"
		XmlToJSON = Str
	End Function
	' Output collection as JSON
	Public Function CollectionToJSON(Prefix, Client)
		Dim Name, Id, Str, Pos, Keys, I
		Dim Suffix, IsClient
		Suffix = "/attr/value"
		Str = "{"
		Keys = objDict.Keys
		For I = 0 To Ubound(Keys)
			Name = Keys(I)
			If Left(Name, Len(Prefix)) = Prefix And Right(Name, Len(Suffix)) = Suffix Then
				Pos = InStrRev(Name, Suffix)
				Id = Mid(Name, Len(Prefix) + 1, Pos - Len(Prefix) - 1)
				IsClient = (GetDictAtt(Prefix & Id & "/attr/client") = "1")
				If Not Client Or Client And IsClient Then
					Str = Str & """" & ew_JsEncode2(Id) & """:""" & ew_JsEncode2(GetDictAtt(Name)) & ""","
				End If
			End If
		Next  
		If Right(Str, 1) = "," Then Str = Left(Str, Len(Str)-1)
		Str = Str & "}"
		CollectionToJSON = Str
	End Function
	' Output all phrases as JSON
	Public Function AllToJSON()
		If EW_USE_DOM_XML Then
			AllToJSON ="var ewLanguage = new ew_Language(" & XmlToJSON("//global/phrase") & ");"
		Else
			AllToJSON = "var ewLanguage = new ew_Language(" & CollectionToJSON("/global/phrase/", False) & ");"
		End If
	End Function
	' Output client phrases as JSON
	Public Function ToJSON()
		If EW_USE_DOM_XML Then
			ToJSON = "var ewLanguage = new ew_Language(" & XmlToJSON("//global/phrase[@client='1']") & ");"
		Else
			ToJSON = "var ewLanguage = new ew_Language(" & CollectionToJSON("/global/phrase/", True) & ");"
		End If
	End Function
End Class
'
'  Language class (end)
' ----------------------
' Format sequence number
Function ew_FormatSeqNo(seq)
	ew_FormatSeqNo =  Replace(Language.Phrase("SequenceNumber"), "%s", seq)
End Function
' Encode value for single-quoted JavaScript string
Function ew_JsEncode(val)
	val = Replace(val & "", "\", "\\")
	val = Replace(val, "'", "\'")
'	val = Replace(val, vbCrLf, "\r\n")
'	val = Replace(val, vbCr, "\r")
'	val = Replace(val, vbLf, "\n")
	val = Replace(val, vbCrLf, "<br>")
	val = Replace(val, vbCr, "<br>")
	val = Replace(val, vbLf, "<br>")
	ew_JsEncode = val
End Function
' Encode value for double-quoted Javascript string
Function ew_JsEncode2(val)
	val = Replace(val & "", "\", "\\")
	val = Replace(val, """", "\""")
'	val = Replace(val, vbCrLf, "\r\n")
'	val = Replace(val, vbCr, "\r")
'	val = Replace(val, vbLf, "\n")
	val = Replace(val, vbCrLf, "<br>")
	val = Replace(val, vbCr, "<br>")
	val = Replace(val, vbLf, "<br>")
	ew_JsEncode2 = val
End Function
' Encode value to single-quoted Javascript string for HTML attributes
Function ew_JsEncode3(val)
	val = Replace(val & "", "\", "\\")
	val = Replace(val, "'", "\'")
	val = Replace(val, """", "&quot;")
	ew_JsEncode3 = val
End Function
' Get full url
Function ew_FullUrl()
	ew_FullUrl = ew_DomainUrl() & ew_ScriptName()
End Function 
' Get current script name
Function ew_ScriptName()
	ew_ScriptName = Request.ServerVariables("SCRIPT_NAME")
End Function
' Get current page name
Function ew_CurrentPage()
	ew_CurrentPage = ew_GetPageName(ew_ScriptName())
End Function
' Get page name
Function ew_GetPageName(url)
	If url <> "" Then
		ew_GetPageName = url
		If InStr(ew_GetPageName, "?") > 0 Then
			ew_GetPageName = Mid(ew_GetPageName, 1, InStr(ew_GetPageName, "?")-1) ' Remove querystring first
		End If
		ew_GetPageName = Mid(ew_GetPageName, InStrRev(ew_GetPageName, "/")+1) ' Remove path
	Else
		ew_GetPageName = ""
	End If
End Function
' Get domain url
Function ew_DomainUrl()
	Dim sUrl, bSSL, sPort, defPort
	sUrl = "http"
	bSSL = ew_IsHttps()
	sPort = Request.ServerVariables("SERVER_PORT")
	If bSSL Then defPort = "443" Else defPort = "80"
	If sPort = defPort Then sPort = "" Else sPort = ":" & sPort
	If bSSL Then sUrl = sUrl & "s"
	sUrl = sUrl & "://"
	sUrl = sUrl & Request.ServerVariables("SERVER_NAME") & sPort
	ew_DomainUrl = sUrl
End Function 
' Get jQuery files host
Function ew_jQueryHost(mobile)
	ew_jQueryHost = "jquery/" ' Use local files
End Function
' jQuery version
Function ew_jQueryFile(f)
	Dim ver, mver, m, v
	ver = "1.11.0" ' jQuery version
	mver = "1.3.2" ' jquery.mobile version
	m = (InStr(f, "mobile") > 0)
	v = ew_IIf(m, mver, ver)
	ew_jQueryFile = Replace(ew_jQueryHost(m) & f, "%v", v)
End Function
' IIf function
Function ew_IIf(cond, v1, v2)
	On Error Resume Next
	If cond & "" = "" Then
		ew_IIf = v2
	ElseIf CBool(cond) Then
		ew_IIf = v1
	Else
		ew_IIf = v2
	End If
End Function
' Check if HTTPS
Function ew_IsHttps()
	ew_IsHttps = (Request.ServerVariables("HTTPS") <> "" And Request.ServerVariables("HTTPS") <> "off")
End Function
' Get current url
Function ew_CurrentUrl()
	Dim s, q
	s = ew_ScriptName()
	q = Request.ServerVariables("QUERY_STRING")
	If q <> "" Then s = s & "?" & q
	ew_CurrentUrl = s
End Function
' Convert to full url
Function ew_ConvertFullUrl(url)
	Dim sUrl
	If url = "" Then
		ew_ConvertFullUrl = ""
	ElseIf Instr(url, "://") > 0 Then
		ew_ConvertFullUrl = url
	Else
		sUrl = ew_FullUrl
		ew_ConvertFullUrl = Mid(sUrl, 1, InStrRev(sUrl, "/")) & url
	End If
End Function
Function ew_RegExMatch(expr, src, m)
	Dim RE
	Set RE = New RegExp
	RE.IgnoreCase = True
	RE.Global = True
	RE.Pattern = expr
	Set m = RE.Execute(src)
	ew_RegExMatch = (m.Count > 0)
	Set RE = Nothing
End Function
' Create XML Dom object
Function ew_CreateXmlDom()
	On Error Resume Next
	Dim ProgId
	ProgId = Array("MSXML2.DOMDocument", "Microsoft.XMLDOM") ' Add other ProgID here
	Dim i
	For i = 0 To UBound(ProgId)
		Set ew_CreateXmlDom = Server.CreateObject(ProgId(i))
		If Err.Number = 0 Then Exit For
	Next
End Function
' Check if mobile device
Function ew_IsMobile()
	ew_IsMobile = ewr_IsMobile()
End Function
' *** DO NOT CHANGE
