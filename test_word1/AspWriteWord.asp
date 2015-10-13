<% Option Explicit %>
<html>
<head>
<title>ThaiCreate.Com ASP & Word.Application</title>
</head>
<body>
<%
	Dim Wrd,DocName
	Set Wrd = CreateObject("Word.Application")
	Wrd.Application.Visible = False
	call Wrd.Documents.Add()
	DocName = "MyDoc/MyWord.doc"
	With Wrd
		.Selection.TypeText "Welcome To www.ThaiCreate.Com"
	End With

	Wrd.ActiveDocument.SaveAs(Server.MapPath(DocName))
	Wrd.Application.Quit
	Set Wrd = Nothing
	
%>
Word Created <a href="<%=DocName%>">Click here</a> to Download.
</body>
</html>