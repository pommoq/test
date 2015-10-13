<%@LANGUAGE="VBSCRIPT" CODEPAGE="874"%>
<% Option Explicit %>
<html>
<head>
<title>ThaiCreate.Com ASP & Word.Application</title>
</head>
<body>
<%
	Dim Wrd,WrdDoc,DocName,WTable
	Set Wrd = CreateObject("Word.Application")
	Wrd.Application.Visible = False
	DocName = "MyDoc/MyWord.doc"
	
	Set WrdDoc = Wrd.Documents.Add()

	Set WTable = WrdDoc.Tables.Add(Wrd.Selection.Range, 3, 3) ' Colums, Rows

	WTable.Cell(1, 1).Range.Font.Name = "Times New Roman"
	WTable.Cell(1, 1).Range.Text = "ThaiCreate.Com 1"
	WTable.Cell(1, 2).Range.Font.Size = 18
	WTable.Cell(1, 2).Range.Bold = True
	WTable.Cell(1, 2).Range.Font.Italic = True
	WTable.Cell(1, 2).Range.Text = "ThaiCreate.Com 2"	
	WTable.Cell(2, 1).Range.ParagraphFormat.Alignment = 1 ' 0= Left, 1=Center, 2=Right

	WTable.Cell(2, 1).Range.Font.Name = "Arial"
	WTable.Cell(2, 1).Range.Font.Size = 12
	WTable.Cell(2, 1).Range.Bold = False
	WTable.Cell(2, 1).Range.ParagraphFormat.Alignment = 2	

	WTable.Cell(3, 3).Range.Font.Name = "Times New Roman"
	WTable.Cell(3, 3).Range.Font.Size = 14
	WTable.Cell(3, 3).Range.Bold = True
	WTable.Cell(3, 3).Range.Font.Underline = True
	WTable.Cell(3, 3).Range.ParagraphFormat.Alignment = 0
	WTable.Cell(3, 2).Range.Text = "ThaiCreate.Com 3"



	Wrd.ActiveDocument.SaveAs(Server.MapPath(DocName))
	Wrd.Application.Quit
	Set Wrd = Nothing
	
%>
Word Created <a href="<%=DocName%>">Click here</a> to Download.
</body>
</html>
