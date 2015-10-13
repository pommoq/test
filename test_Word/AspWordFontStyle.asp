<% Option Explicit %>
<html>
<head>
<title>ThaiCreate.Com ASP & Word.Application</title>
</head>
<body>
<%
	Const wdColorDarkBlue = &H800000

	Dim Wrd,DocName
	Set Wrd = CreateObject("Word.Application")
	Wrd.Application.Visible = False
	Wrd.Documents.Add()
	DocName = "MyDoc/MyWord.doc"
	With Wrd		
		.Selection.PageSetup.LeftMargin = "0.25"
		.Selection.PageSetup.RightMargin = "0.25"
		.Selection.PageSetup.TopMargin = "0.25"
		.Selection.PageSetup.BottomMargin = "0.25"
		'*** Font Properties ***'
		.Selection.Font.Name = "Verdana"
		.Selection.Font.Size = "25"		
		.Selection.Font.Bold = True
		.Selection.Font.Italic = True
		.Selection.Font.Color = wdColorDarkBlue		
		'.Selection.Font.Emboss = True
		'.Selection.Font.Engrave = True
		'.Selection.Font.Kerning = True
		'.Selection.Font.Shadow = True
		'.Selection.Font.StrikeThrough = True
		'.Selection.Font.Subscript = True
		'.Selection.Font.Superscript = True
		.Selection.TypeText "Welcome To www.ThaiCreate.Com"
	End With

	Wrd.ActiveDocument.SaveAs(Server.MapPath(DocName))
	Wrd.Application.Quit
	Set Wrd = Nothing	


	'*** Font Color ***'
	'Const wdColorAqua = &HCCCC33
	'Const wdColorAutomatic = &HFF000000
	'Const wdColorBlack = &H0
	'Const wdColorBlue = &HFF0000
	'Const wdColorBlueGray = &H996666
	'Const wdColorBrightGreen = &HFF00
	'Const wdColorBrown = &H3399
	'Const wdColorDarkBlue = &H800000
	'Const wdColorDarkGreen = &H3300
	'Const wdColorDarkRed = &H80
	'Const wdColorDarkTeal = &H663300
	'Const wdColorDarkYellow = &H8080
	'Const wdColorGold = &HCCFF
	'Const wdColorGray05 = &HF3F3F3
	'Const wdColorGray10 = &HE6E6E6
	'Const wdColorGray125 = &HE0E0E0
	'Const wdColorGray15 = &HD9D9D9
	'Const wdColorGray20 = &HCCCCCC
	'Const wdColorGray25 = &HC0C0C0
	'Const wdColorGray30 = &HB3B3B3
	'Const wdColorGray35 = &HA6A6A6
	'Const wdColorGray375 = &HA0A0A0
	'Const wdColorGray40 = &H999999
	'Const wdColorGray45 = &H8C8C8C
	'Const wdColorGray50 = &H808080
	'Const wdColorGray55 = &H737373
	'Const wdColorGray60 = &H666666
	'Const wdColorGray625 = &H606060
	'Const wdColorGray65 = &H595959
	'Const wdColorGray70 = &H4C4C4C
	'Const wdColorGray75 = &H404040
	'Const wdColorGray80 = &H333333
	'Const wdColorGray85 = &H262626
	'Const wdColorGray875 = &H202020
	'Const wdColorGray90 = &H191919
	'Const wdColorGray95 = &HC0C0C
	'Const wdColorGreen = &H8000
	'Const wdColorIndigo = &H993333
	'Const wdColorLavender = &HFF99CC
	'Const wdColorLightBlue = &HFF6633
	'Const wdColorLightGreen = &HCCFFCC
	'Const wdColorLightOrange = &H99FF
	'Const wdColorLightTurquoise = &HFFFFCC
	'Const wdColorLightYellow = &H99FFFF
	'Const wdColorLime = &HCC99
	'Const wdColorOliveGreen = &H3333
	'Const wdColorOrange = &H66FF
	'Const wdColorPaleBlue = &HFFCC99
	'Const wdColorPink = &HFF00FF
	'Const wdColorPlum = &H663399
	'Const wdColorRed = &HFF
	'Const wdColorRose = &HCC99FF
	'Const wdColorSeaGreen = &H669933
	'Const wdColorSkyBlue = &HFFCC00
	'Const wdColorTan = &H99CCFF
	'Const wdColorTeal = &H808000
	'Const wdColorTurquoise = &HFFFF00
	'Const wdColorViolet = &H800080
	'Const wdColorWhite = &HFFFFFF
	'Const wdColorYellow = &HFFFF


%>
Word Created <a href="<%=DocName%>">Click here</a> to Download.
</body>
</html>