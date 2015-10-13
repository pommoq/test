<%

' ASPMaker configuration for Table Query2
Dim Query2

' Define table class
Class cQuery2

	' Class Initialize
	Private Sub Class_Initialize()
		UseTokenInUrl = EW_USE_TOKEN_IN_URL
		ExportAll = True
		ExportPageBreakCount = 0 ' Page break per every n record (PDF only)
		ExportPageOrientation = "portrait" ' Page orientation (PDF only)
		ExportPageSize = "a4" ' Page size (PDF only)
		Set RowAttrs = New cAttributes ' Row attributes
		Set CustomActions = New cCustomArray
		PrinterFriendlyForPdf = False
		AllowAddDeleteRow = ew_AllowAddDeleteRow() ' Allow add/delete row
		DetailAdd = False ' Allow detail add
		DetailEdit = False ' Allow detail edit
		DetailView = False ' Allow detail view
		ShowMultipleDetails = False ' Show multiple details
		GridAddRowCount = 5 ' Grid add row count
		ValidateKey = True ' Validate key
		Visible = True
		BasicSearch.TblVar = TableVar
		BasicSearch.KeywordDefault = ""
		BasicSearch.SearchTypeDefault = "="
		UserIDAllowSecurity = 0 ' User ID Allow
		Call ew_SetArObj(Fields, "Expr1", Expr1)
		Call ew_SetArObj(Fields, "Expr2", Expr2)
		Call ew_SetArObj(Fields, "Expr3", Expr3)
		Call ew_SetArObj(Fields, "Expr4", Expr4)
		Call ew_SetArObj(Fields, "Expr5", Expr5)
		Call ew_SetArObj(Fields, "Expr6", Expr6)
		Call ew_SetArObj(Fields, "Expr7", Expr7)
		Call ew_SetArObj(Fields, "Expr8", Expr8)
		Call ew_SetArObj(Fields, "Expr9", Expr9)
		Call ew_SetArObj(Fields, "Expr10", Expr10)
		Call ew_SetArObj(Fields, "Expr11", Expr11)
		Call ew_SetArObj(Fields, "Expr12", Expr12)
		Call ew_SetArObj(Fields, "Expr13", Expr13)
		Call ew_SetArObj(Fields, "Expr14", Expr14)
		Call ew_SetArObj(Fields, "Expr15", Expr15)
		Call ew_SetArObj(Fields, "Expr16", Expr16)
		Call ew_SetArObj(Fields, "Expr17", Expr17)
		Call ew_SetArObj(Fields, "Expr18", Expr18)
		Call ew_SetArObj(Fields, "Expr19", Expr19)
		Call ew_SetArObj(Fields, "Expr20", Expr20)
		Call ew_SetArObj(Fields, "Expr21", Expr21)
		Call ew_SetArObj(Fields, "Expr22", Expr22)
		Call ew_SetArObj(Fields, "Expr23", Expr23)
		Call ew_SetArObj(Fields, "Expr24", Expr24)
		Call ew_SetArObj(Fields, "Expr25", Expr25)
		Call ew_SetArObj(Fields, "Expr26", Expr26)
		Call ew_SetArObj(Fields, "Expr27", Expr27)
		Call ew_SetArObj(Fields, "Expr28", Expr28)
		Call ew_SetArObj(Fields, "Expr29", Expr29)
		Call ew_SetArObj(Fields, "Expr30", Expr30)
		Call ew_SetArObj(Fields, "Expr31", Expr31)
		Call ew_SetArObj(Fields, "Expr32", Expr32)
		Call ew_SetArObj(Fields, "Expr33", Expr33)
		Call ew_SetArObj(Fields, "Expr34", Expr34)
		Call ew_SetArObj(Fields, "Expr35", Expr35)
		Call ew_SetArObj(Fields, "Expr36", Expr36)
		Call ew_SetArObj(Fields, "Expr37", Expr37)
	End Sub

	' Reset attributes for table object
	Public Sub ResetAttrs()
		CssClass = ""
		CssStyle = ""
		RowAttrs.Clear()
		Dim i, fld
		If IsArray(Fields) Then
			For i = 0 to UBound(Fields,2)
				Set fld = Fields(1,i)
				Call fld.ResetAttrs()
			Next
		End If
	End Sub

	' Setup field titles
	Public Sub SetupFieldTitles()
		Dim i, fld
		If IsArray(Fields) Then
			For i = 0 to UBound(Fields,2)
				Set fld = Fields(1,i)
				If fld.FldTitle <> "" Then
					fld.EditAttrs.UpdateAttribute "data-toggle", "tooltip"
					fld.EditAttrs.UpdateAttribute "title", ew_HtmlEncode(fld.FldTitle)
				End If
			Next
		End If
	End Sub

	' Define table level constants
	' Use table token in Url

	Dim UseTokenInUrl

	' Table variable
	Public Property Get TableVar()
		TableVar = "Query2"
	End Property

	' Table name
	Public Property Get TableName()
		TableName = "Query2"
	End Property

	' Table type
	Public Property Get TableType()
		TableType = "VIEW"
	End Property

	' Table caption
	Dim Caption

	Public Property Let TableCaption(v)
		Caption = v
	End Property

	Public Property Get TableCaption()
		If Caption & "" <> "" Then
			TableCaption = Caption
		Else
			TableCaption = Language.TablePhrase(TableVar, "TblCaption")
		End If
	End Property

	' Page caption
	Dim PgCaption

	Public Property Let PageCaption(Page, v)
		If Not IsArray(PgCaption) Then
			ReDim PgCaption(Page)
		ElseIf Page > UBound(PgCaption) Then
			ReDim Preserve PgCaption(Page)
		End If
		PgCaption(Page) = v
	End Property

	Public Property Get PageCaption(Page)
		PageCaption = ""
		If IsArray(PgCaption) Then
			If Page <= UBound(PgCaption) Then
				PageCaption = PgCaption(Page)
			End If
		End If
		If PageCaption = "" Then PageCaption = Language.TablePhrase(TableVar, "TblPageCaption" & Page)
		If PageCaption = "" Then PageCaption = "Page " & Page
	End Property
	Dim Visible

	' Export Return Page
	Public Property Get ExportReturnUrl()
		If Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_EXPORT_RETURN_URL) <> "" Then
			ExportReturnUrl = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_EXPORT_RETURN_URL)
		Else
			ExportReturnUrl = ew_CurrentPage
		End If
	End Property

	Public Property Let ExportReturnUrl(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_EXPORT_RETURN_URL) = v
	End Property

	' Records per page
	Public Property Get RecordsPerPage()
		RecordsPerPage = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_REC_PER_PAGE)
	End Property

	Public Property Let RecordsPerPage(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_REC_PER_PAGE) = v
	End Property

	' Start record number
	Public Property Get StartRecordNumber()
		StartRecordNumber = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_START_REC)
	End Property

	Public Property Let StartRecordNumber(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_START_REC) = v
	End Property

	' Search Highlight Name
	Public Property Get HighlightName()
		HighlightName = "Query2_Highlight"
	End Property

	' Search where clause
	Public Property Get SearchWhere()
		SearchWhere = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_SEARCH_WHERE)
	End Property

	Public Property Let SearchWhere(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_SEARCH_WHERE) = v
	End Property

	' Single column sort
	Public Sub UpdateSort(ofld)
		Dim sSortField, sLastSort, sThisSort
		If CurrentOrder = ofld.FldName Then
			sSortField = ofld.FldExpression
			sLastSort = ofld.Sort
			If CurrentOrderType = "ASC" Or CurrentOrderType = "DESC" Then
				sThisSort = CurrentOrderType
			Else
				If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			End If
			ofld.Sort = sThisSort
			SessionOrderBy = sSortField & " " & sThisSort ' Save to Session
		Else
			ofld.Sort = ""
		End If
	End Sub

	' BasicSearch Object
	Private m_BasicSearch

	Public Property Get BasicSearch()
		If Not IsObject(m_BasicSearch) Then
			Set m_BasicSearch = New cBasicSearch
		End If
		Set BasicSearch = m_BasicSearch
	End Property

	' Session WHERE Clause
	Public Property Get SessionWhere()
		SessionWhere = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_WHERE)
	End Property

	Public Property Let SessionWhere(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_WHERE) = v
	End Property

	' Session ORDER BY
	Public Property Get SessionOrderBy()
		SessionOrderBy = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_ORDER_BY)
	End Property

	Public Property Let SessionOrderBy(v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_ORDER_BY) = v
	End Property

	' Session Key
	Public Function GetKey(fld)
		GetKey = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_KEY & "_" & fld)
	End Function

	Public Function SetKey(fld, v)
		Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_KEY & "_" & fld) = v
	End Function

	' Table level SQL
	Public Property Get SqlSelect() ' Select
		SqlSelect = "SELECT * FROM [Query2]"
	End Property

	Private Property Get TableFilter()
		TableFilter = ""
	End Property

	Public Property Get SqlWhere() ' Where
		Dim sWhere
		sWhere = ""
		Call ew_AddFilter(sWhere, TableFilter)
		SqlWhere = sWhere
	End Property

	Public Property Get SqlGroupBy() ' Group By
		SqlGroupBy = ""
	End Property

	Public Property Get SqlHaving() ' Having
		SqlHaving = ""
	End Property

	Public Property Get SqlOrderBy() ' Order By
		SqlOrderBy = ""
	End Property

	' SQL variables
	Dim CurrentFilter ' Current filter
	Dim CurrentOrder ' Current order
	Dim CurrentOrderType ' Current order type

	' Get sql
	Public Function GetSQL(where, orderby)
		GetSQL = ew_BuildSelectSql(SqlSelect, SqlWhere, SqlGroupBy, SqlHaving, SqlOrderBy, where, orderby)
	End Function

	' Table sql
	Public Property Get SQL()
		Dim sFilter, sSort
		sFilter = CurrentFilter
		sSort = SessionOrderBy
		SQL = ew_BuildSelectSql(SqlSelect, SqlWhere, SqlGroupBy, SqlHaving, SqlOrderBy, sFilter, sSort)
	End Property

	' Return table sql with list page filter
	Public Property Get ListSQL()
		Dim sFilter, sSort
		sFilter = SessionWhere
		Call ew_AddFilter(sFilter, CurrentFilter)
		sSort = SessionOrderBy
		ListSQL = ew_BuildSelectSql(SqlSelect, SqlWhere, SqlGroupBy, SqlHaving, SqlOrderBy, sFilter, sSort)
	End Property

	' Key filter for table
	Private Property Get SqlKeyFilter()
		SqlKeyFilter = ""
	End Property

	' Return Key filter for table
	Public Property Get KeyFilter()
		Dim sKeyFilter
		sKeyFilter = SqlKeyFilter
		KeyFilter = sKeyFilter
	End Property

	' Return url
	Public Property Get ReturnUrl()

		' Get referer url automatically
		If Request.ServerVariables("HTTP_REFERER") <> "" Then
			If ew_ReferPage <> ew_CurrentPage And ew_ReferPage <> "pom_login.asp" Then ' Referer not same page or login page
				Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL) = Request.ServerVariables("HTTP_REFERER") ' Save to Session
			End If
		End If
		If Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL) <> "" Then
			ReturnUrl = Session(EW_PROJECT_NAME & "_" & TableVar & "_" & EW_TABLE_RETURN_URL)
		Else
			ReturnUrl = "pom_query2list.asp"
		End If
	End Property

	' List url
	Public Function ListUrl()
		ListUrl = "pom_query2list.asp"
	End Function

	' View url
	Public Function ViewUrl(parm)
		If parm <> "" Then
			ViewUrl = KeyUrl("pom_query2view.asp", UrlParm(parm))
		Else
			ViewUrl = KeyUrl("pom_query2view.asp", UrlParm(EW_TABLE_SHOW_DETAIL & "="))
		End If
	End Function

	' Add url
	Public Function AddUrl()
		AddUrl = "pom_query2add.asp"

'		Dim sUrlParm
'		sUrlParm = UrlParm("")
'		If sUrlParm <> "" Then AddUrl = AddUrl & "?" & sUrlParm

	End Function

	' Edit url
	Public Function EditUrl(parm)
		EditUrl = KeyUrl("pom_query2edit.asp", UrlParm(parm))
	End Function

	' Inline edit url
	Public Function InlineEditUrl()
		InlineEditUrl = KeyUrl(ew_CurrentPage, UrlParm("a=edit"))
	End Function

	' Copy url
	Public Function CopyUrl(parm)
		CopyUrl = KeyUrl("pom_query2add.asp", UrlParm(parm))
	End Function

	' Inline copy url
	Public Function InlineCopyUrl()
		InlineCopyUrl = KeyUrl(ew_CurrentPage, UrlParm("a=copy"))
	End Function

	' Delete url
	Public Function DeleteUrl()
		DeleteUrl = KeyUrl("pom_query2delete.asp", UrlParm(""))
	End Function

	' Key url
	Public Function KeyUrl(url, parm)
		Dim sUrl: sUrl = url & "?"
		If parm <> "" Then sUrl = sUrl & parm & "&"
		KeyUrl = sUrl
	End Function

	' Sort Url
	Public Property Get SortUrl(fld)
		If CurrentAction <> "" Or Export <> "" Or (fld.FldType = 201 Or fld.FldType = 203 Or fld.FldType = 205 Or fld.FldType = 141) Then
			SortUrl = ""
		ElseIf fld.Sortable Then
			SortUrl = ew_CurrentPage
			Dim sUrlParm
			sUrlParm = UrlParm("order=" & Server.URLEncode(fld.FldName) & "&amp;ordertype=" & fld.ReverseSort)
			SortUrl = SortUrl & "?" & sUrlParm
		Else
			SortUrl = ""
		End If
	End Property

	' Url parm
	Function UrlParm(parm)
		If UseTokenInUrl Then
			UrlParm = "t=Query2"
		Else
			UrlParm = ""
		End If
		If parm <> "" Then
			If UrlParm <> "" Then UrlParm = UrlParm & "&"
			UrlParm = UrlParm & parm
		End If
	End Function

	' Get record keys from Form/QueryString/Session
	Public Function GetRecordKeys()
		Dim arKeys, arKey, cnt, i, bHasKey
		bHasKey = False

		' Check ObjForm first
		If IsObject(ObjForm) And Not (ObjForm Is Nothing) Then
			ObjForm.Index = -1
			If ObjForm.HasValue("key_m") Then
				arKeys = ObjForm.GetValue("key_m")
				If Not IsArray(arKeys) Then
					arKeys = Array(arKeys)
				End If
				bHasKey = True
			End If
		End If

		' Check Form/QueryString
		If Not bHasKey Then
			If Request.Form("key_m").Count > 0 Then
				cnt = Request.Form("key_m").Count
				ReDim arKeys(cnt-1)
				For i = 1 to cnt ' Set up keys
					arKeys(i-1) = Request.Form("key_m")(i)
				Next
			ElseIf Request.QueryString("key_m").Count > 0 Then
				cnt = Request.QueryString("key_m").Count
				ReDim arKeys(cnt-1)
				For i = 1 to cnt ' Set up keys
					arKeys(i-1) = Request.QueryString("key_m")(i)
				Next
			ElseIf Request.QueryString <> "" Then
				ReDim arKeys(0)

				'GetRecordKeys = arKeys ' Do not return yet, so the values will also be checked by the following code
			End If
		End If

		' Check keys
		Dim ar, key
		If IsArray(arKeys) Then
			For i = 0 to UBound(arKeys)
				key = arKeys(i)
						Dim skip
						skip = False
						If Not skip Then
							If IsArray(ar) Then
								ReDim Preserve ar(UBound(ar)+1)
							Else
								ReDim ar(0)
							End If
							ar(UBound(ar)) = key
						End If
			Next
		End If
		GetRecordKeys = ar
	End Function

	' Get key filter
	Public Function GetKeyFilter()
		Dim arKeys, sKeyFilter, i, key
		arKeys = GetRecordKeys()
		sKeyFilter = ""
		If IsArray(arKeys) Then
			For i = 0 to UBound(arKeys)
				key = arKeys(i)
				If sKeyFilter <> "" Then sKeyFilter = sKeyFilter & " OR "
				sKeyFilter = sKeyFilter & "(" & KeyFilter & ")"
			Next
		End If
		GetKeyFilter = sKeyFilter
	End Function

	' Function LoadRecordCount
	' - Load record count based on filter
	Public Function LoadRecordCount(sFilter)
		Dim wrkrs
		Set wrkrs = LoadRs(sFilter)
		If Not wrkrs Is Nothing Then
			LoadRecordCount = wrkrs.RecordCount
		Else
			LoadRecordCount = 0
		End If
		Set wrkrs = Nothing
	End Function

	' Function LoadRs
	' - Load Rows based on filter
	Public Function LoadRs(sFilter)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next
		Dim RsRows, sSql

		' Set up filter (Sql Where Clause) and get Return Sql
		'CurrentFilter = sFilter
		'sSql = SQL

		sSql = GetSQL(sFilter, "")
		Err.Clear
		Set RsRows = Server.CreateObject("ADODB.Recordset")
		RsRows.CursorLocation = EW_CURSORLOCATION
		RsRows.Open sSql, Conn, 3, 1, 1 ' adOpenStatic, adLockReadOnly, adCmdText
		If Err.Number <> 0 Then
			Err.Clear
			Set LoadRs = Nothing
			RsRows.Close
			Set RsRows = Nothing
		ElseIf RsRows.Eof Then
			Set LoadRs = Nothing
			RsRows.Close
			Set RsRows = Nothing
		Else
			Set LoadRs = RsRows
		End If
	End Function

	' Load row values from recordset
	Public Sub LoadListRowValues(RsRow)
		Expr1.DbValue = RsRow("Expr1")
		Expr2.DbValue = RsRow("Expr2")
		Expr3.DbValue = RsRow("Expr3")
		Expr4.DbValue = RsRow("Expr4")
		Expr5.DbValue = RsRow("Expr5")
		Expr6.DbValue = RsRow("Expr6")
		Expr7.DbValue = RsRow("Expr7")
		Expr8.DbValue = RsRow("Expr8")
		Expr9.DbValue = RsRow("Expr9")
		Expr10.DbValue = RsRow("Expr10")
		Expr11.DbValue = RsRow("Expr11")
		Expr12.DbValue = RsRow("Expr12")
		Expr13.DbValue = RsRow("Expr13")
		Expr14.DbValue = RsRow("Expr14")
		Expr15.DbValue = RsRow("Expr15")
		Expr16.DbValue = RsRow("Expr16")
		Expr17.DbValue = RsRow("Expr17")
		Expr18.DbValue = RsRow("Expr18")
		Expr19.DbValue = RsRow("Expr19")
		Expr20.DbValue = RsRow("Expr20")
		Expr21.DbValue = RsRow("Expr21")
		Expr22.DbValue = RsRow("Expr22")
		Expr23.DbValue = RsRow("Expr23")
		Expr24.DbValue = RsRow("Expr24")
		Expr25.DbValue = RsRow("Expr25")
		Expr26.DbValue = RsRow("Expr26")
		Expr27.DbValue = RsRow("Expr27")
		Expr28.DbValue = RsRow("Expr28")
		Expr29.DbValue = RsRow("Expr29")
		Expr30.DbValue = RsRow("Expr30")
		Expr31.DbValue = RsRow("Expr31")
		Expr32.DbValue = RsRow("Expr32")
		Expr33.DbValue = RsRow("Expr33")
		Expr34.DbValue = RsRow("Expr34")
		Expr35.DbValue = RsRow("Expr35")
		Expr36.DbValue = RsRow("Expr36")
		Expr37.DbValue = RsRow("Expr37")
	End Sub

	' Render list row values
	Sub RenderListRow()

		'
		'  Common render codes
		'
		' Expr1
		' Expr2
		' Expr3
		' Expr4
		' Expr5
		' Expr6
		' Expr7
		' Expr8
		' Expr9
		' Expr10
		' Expr11
		' Expr12
		' Expr13
		' Expr14
		' Expr15
		' Expr16
		' Expr17
		' Expr18
		' Expr19
		' Expr20
		' Expr21
		' Expr22
		' Expr23
		' Expr24
		' Expr25
		' Expr26
		' Expr27
		' Expr28
		' Expr29
		' Expr30
		' Expr31
		' Expr32
		' Expr33
		' Expr34
		' Expr35
		' Expr36
		' Expr37
		' Call Row Rendering event

		Call Row_Rendering()

		'
		'  Render for View
		'
		' Expr1

		Expr1.ViewValue = Expr1.CurrentValue
		Expr1.ViewCustomAttributes = ""

		' Expr2
		Expr2.ViewValue = Expr2.CurrentValue
		Expr2.ViewCustomAttributes = ""

		' Expr3
		Expr3.ViewValue = Expr3.CurrentValue
		Expr3.ViewCustomAttributes = ""

		' Expr4
		Expr4.ViewValue = Expr4.CurrentValue
		Expr4.ViewCustomAttributes = ""

		' Expr5
		Expr5.ViewValue = Expr5.CurrentValue
		Expr5.ViewCustomAttributes = ""

		' Expr6
		Expr6.ViewValue = Expr6.CurrentValue
		Expr6.ViewCustomAttributes = ""

		' Expr7
		Expr7.ViewValue = Expr7.CurrentValue
		Expr7.ViewCustomAttributes = ""

		' Expr8
		Expr8.ViewValue = Expr8.CurrentValue
		Expr8.ViewCustomAttributes = ""

		' Expr9
		Expr9.ViewValue = Expr9.CurrentValue
		Expr9.ViewCustomAttributes = ""

		' Expr10
		Expr10.ViewValue = Expr10.CurrentValue
		Expr10.ViewCustomAttributes = ""

		' Expr11
		Expr11.ViewValue = Expr11.CurrentValue
		Expr11.ViewCustomAttributes = ""

		' Expr12
		Expr12.ViewValue = Expr12.CurrentValue
		Expr12.ViewCustomAttributes = ""

		' Expr13
		Expr13.ViewValue = Expr13.CurrentValue
		Expr13.ViewCustomAttributes = ""

		' Expr14
		Expr14.ViewValue = Expr14.CurrentValue
		Expr14.ViewCustomAttributes = ""

		' Expr15
		Expr15.ViewValue = Expr15.CurrentValue
		Expr15.ViewCustomAttributes = ""

		' Expr16
		Expr16.ViewValue = Expr16.CurrentValue
		Expr16.ViewCustomAttributes = ""

		' Expr17
		Expr17.ViewValue = Expr17.CurrentValue
		Expr17.ViewCustomAttributes = ""

		' Expr18
		Expr18.ViewValue = Expr18.CurrentValue
		Expr18.ViewCustomAttributes = ""

		' Expr19
		Expr19.ViewValue = Expr19.CurrentValue
		Expr19.ViewCustomAttributes = ""

		' Expr20
		Expr20.ViewValue = Expr20.CurrentValue
		Expr20.ViewCustomAttributes = ""

		' Expr21
		Expr21.ViewValue = Expr21.CurrentValue
		Expr21.ViewCustomAttributes = ""

		' Expr22
		Expr22.ViewValue = Expr22.CurrentValue
		Expr22.ViewCustomAttributes = ""

		' Expr23
		Expr23.ViewValue = Expr23.CurrentValue
		Expr23.ViewCustomAttributes = ""

		' Expr24
		Expr24.ViewValue = Expr24.CurrentValue
		Expr24.ViewCustomAttributes = ""

		' Expr25
		Expr25.ViewValue = Expr25.CurrentValue
		Expr25.ViewCustomAttributes = ""

		' Expr26
		Expr26.ViewValue = Expr26.CurrentValue
		Expr26.ViewCustomAttributes = ""

		' Expr27
		Expr27.ViewValue = Expr27.CurrentValue
		Expr27.ViewCustomAttributes = ""

		' Expr28
		Expr28.ViewValue = Expr28.CurrentValue
		Expr28.ViewCustomAttributes = ""

		' Expr29
		Expr29.ViewValue = Expr29.CurrentValue
		Expr29.ViewCustomAttributes = ""

		' Expr30
		Expr30.ViewValue = Expr30.CurrentValue
		Expr30.ViewCustomAttributes = ""

		' Expr31
		Expr31.ViewValue = Expr31.CurrentValue
		Expr31.ViewCustomAttributes = ""

		' Expr32
		Expr32.ViewValue = Expr32.CurrentValue
		Expr32.ViewCustomAttributes = ""

		' Expr33
		Expr33.ViewValue = Expr33.CurrentValue
		Expr33.ViewCustomAttributes = ""

		' Expr34
		Expr34.ViewValue = Expr34.CurrentValue
		Expr34.ViewCustomAttributes = ""

		' Expr35
		Expr35.ViewValue = Expr35.CurrentValue
		Expr35.ViewCustomAttributes = ""

		' Expr36
		Expr36.ViewValue = Expr36.CurrentValue
		Expr36.ViewCustomAttributes = ""

		' Expr37
		Expr37.ViewValue = Expr37.CurrentValue
		Expr37.ViewCustomAttributes = ""

		' Expr1
		Expr1.LinkCustomAttributes = ""
		Expr1.HrefValue = ""
		Expr1.TooltipValue = ""

		' Expr2
		Expr2.LinkCustomAttributes = ""
		Expr2.HrefValue = ""
		Expr2.TooltipValue = ""

		' Expr3
		Expr3.LinkCustomAttributes = ""
		Expr3.HrefValue = ""
		Expr3.TooltipValue = ""

		' Expr4
		Expr4.LinkCustomAttributes = ""
		Expr4.HrefValue = ""
		Expr4.TooltipValue = ""

		' Expr5
		Expr5.LinkCustomAttributes = ""
		Expr5.HrefValue = ""
		Expr5.TooltipValue = ""

		' Expr6
		Expr6.LinkCustomAttributes = ""
		Expr6.HrefValue = ""
		Expr6.TooltipValue = ""

		' Expr7
		Expr7.LinkCustomAttributes = ""
		Expr7.HrefValue = ""
		Expr7.TooltipValue = ""

		' Expr8
		Expr8.LinkCustomAttributes = ""
		Expr8.HrefValue = ""
		Expr8.TooltipValue = ""

		' Expr9
		Expr9.LinkCustomAttributes = ""
		Expr9.HrefValue = ""
		Expr9.TooltipValue = ""

		' Expr10
		Expr10.LinkCustomAttributes = ""
		Expr10.HrefValue = ""
		Expr10.TooltipValue = ""

		' Expr11
		Expr11.LinkCustomAttributes = ""
		Expr11.HrefValue = ""
		Expr11.TooltipValue = ""

		' Expr12
		Expr12.LinkCustomAttributes = ""
		Expr12.HrefValue = ""
		Expr12.TooltipValue = ""

		' Expr13
		Expr13.LinkCustomAttributes = ""
		Expr13.HrefValue = ""
		Expr13.TooltipValue = ""

		' Expr14
		Expr14.LinkCustomAttributes = ""
		Expr14.HrefValue = ""
		Expr14.TooltipValue = ""

		' Expr15
		Expr15.LinkCustomAttributes = ""
		Expr15.HrefValue = ""
		Expr15.TooltipValue = ""

		' Expr16
		Expr16.LinkCustomAttributes = ""
		Expr16.HrefValue = ""
		Expr16.TooltipValue = ""

		' Expr17
		Expr17.LinkCustomAttributes = ""
		Expr17.HrefValue = ""
		Expr17.TooltipValue = ""

		' Expr18
		Expr18.LinkCustomAttributes = ""
		Expr18.HrefValue = ""
		Expr18.TooltipValue = ""

		' Expr19
		Expr19.LinkCustomAttributes = ""
		Expr19.HrefValue = ""
		Expr19.TooltipValue = ""

		' Expr20
		Expr20.LinkCustomAttributes = ""
		Expr20.HrefValue = ""
		Expr20.TooltipValue = ""

		' Expr21
		Expr21.LinkCustomAttributes = ""
		Expr21.HrefValue = ""
		Expr21.TooltipValue = ""

		' Expr22
		Expr22.LinkCustomAttributes = ""
		Expr22.HrefValue = ""
		Expr22.TooltipValue = ""

		' Expr23
		Expr23.LinkCustomAttributes = ""
		Expr23.HrefValue = ""
		Expr23.TooltipValue = ""

		' Expr24
		Expr24.LinkCustomAttributes = ""
		Expr24.HrefValue = ""
		Expr24.TooltipValue = ""

		' Expr25
		Expr25.LinkCustomAttributes = ""
		Expr25.HrefValue = ""
		Expr25.TooltipValue = ""

		' Expr26
		Expr26.LinkCustomAttributes = ""
		Expr26.HrefValue = ""
		Expr26.TooltipValue = ""

		' Expr27
		Expr27.LinkCustomAttributes = ""
		Expr27.HrefValue = ""
		Expr27.TooltipValue = ""

		' Expr28
		Expr28.LinkCustomAttributes = ""
		Expr28.HrefValue = ""
		Expr28.TooltipValue = ""

		' Expr29
		Expr29.LinkCustomAttributes = ""
		Expr29.HrefValue = ""
		Expr29.TooltipValue = ""

		' Expr30
		Expr30.LinkCustomAttributes = ""
		Expr30.HrefValue = ""
		Expr30.TooltipValue = ""

		' Expr31
		Expr31.LinkCustomAttributes = ""
		Expr31.HrefValue = ""
		Expr31.TooltipValue = ""

		' Expr32
		Expr32.LinkCustomAttributes = ""
		Expr32.HrefValue = ""
		Expr32.TooltipValue = ""

		' Expr33
		Expr33.LinkCustomAttributes = ""
		Expr33.HrefValue = ""
		Expr33.TooltipValue = ""

		' Expr34
		Expr34.LinkCustomAttributes = ""
		Expr34.HrefValue = ""
		Expr34.TooltipValue = ""

		' Expr35
		Expr35.LinkCustomAttributes = ""
		Expr35.HrefValue = ""
		Expr35.TooltipValue = ""

		' Expr36
		Expr36.LinkCustomAttributes = ""
		Expr36.HrefValue = ""
		Expr36.TooltipValue = ""

		' Expr37
		Expr37.LinkCustomAttributes = ""
		Expr37.HrefValue = ""
		Expr37.TooltipValue = ""

		' Call Row Rendered event
		If RowType <> EW_ROWTYPE_AGGREGATEINIT Then
			Call Row_Rendered()
		End If
	End Sub

	' Aggregate list row values
	Public Sub AggregateListRowValues()
	End Sub

	' Aggregate list row (for rendering)
	Sub AggregateListRow()
	End Sub

	' Update detail records
	Function UpdateDetailRecords(RsOld, RsNew)
		Dim bUpdate, sFieldList, sWhereList, sSql
		On Error Resume Next
		UpdateDetailRecords = True
	End Function

	' Delete detail records
	Function DeleteDetailRecords(Rs, Where)
		Dim sWhereList, sSql
		On Error Resume Next
		DeleteDetailRecords = True
		sWhereList = Where

		' Delete upload files if necessary
		If IsNull(Rs) Then
			Dim rsfile
			sSql = "SELECT * FROM [Query2] WHERE " & sWhereList
			Set rsfile = Conn.Execute(sSql)
			If Not rsfile.Eof Then rsfile.MoveFirst
			Do While Not rsfile.Eof
				Call LoadListRowValues(rsfile)
				rsfile.MoveNext
			Loop
			rsfile.Close
			Set rsfile = Nothing
		End If
	End Function

	' Export data in Xml Format
	Public Sub ExportXmlDocument(XmlDoc, HasParent, Recordset, StartRec, StopRec, ExportPageType)
		If Not IsObject(Recordset) Or Not IsObject(XmlDoc) Then
			Exit Sub
		End If
		If Not HasParent Then
			Call XmlDoc.AddRoot(TableVar)
		End If

		' Move to first record
		Dim RecCnt, RowCnt
		RecCnt = StartRec - 1
		If Not Recordset.Eof Then
			Recordset.MoveFirst()
			If StartRec > 1 Then Recordset.Move(StartRec - 1)
		End If
		Do While Not Recordset.Eof And RecCnt < StopRec
			RecCnt = RecCnt + 1
			If CLng(RecCnt) >= CLng(StartRec) Then
				RowCnt = CLng(RecCnt) - CLng(StartRec) + 1
				Call LoadListRowValues(Recordset)

				' Render row
				RowType = EW_ROWTYPE_VIEW ' Render view
				Call ResetAttrs()
				Call RenderListRow()
				If HasParent Then
					Call XmlDoc.AddRow(TableVar, "")
				Else
					Call XmlDoc.AddRow("", "")
				End If
				If ExportPageType = "view" Then
					Call XmlDoc.AddField("Expr1", Expr1.ExportValue(Export))
					Call XmlDoc.AddField("Expr2", Expr2.ExportValue(Export))
					Call XmlDoc.AddField("Expr3", Expr3.ExportValue(Export))
					Call XmlDoc.AddField("Expr4", Expr4.ExportValue(Export))
					Call XmlDoc.AddField("Expr5", Expr5.ExportValue(Export))
					Call XmlDoc.AddField("Expr6", Expr6.ExportValue(Export))
					Call XmlDoc.AddField("Expr7", Expr7.ExportValue(Export))
					Call XmlDoc.AddField("Expr8", Expr8.ExportValue(Export))
					Call XmlDoc.AddField("Expr9", Expr9.ExportValue(Export))
					Call XmlDoc.AddField("Expr10", Expr10.ExportValue(Export))
					Call XmlDoc.AddField("Expr11", Expr11.ExportValue(Export))
					Call XmlDoc.AddField("Expr12", Expr12.ExportValue(Export))
					Call XmlDoc.AddField("Expr13", Expr13.ExportValue(Export))
					Call XmlDoc.AddField("Expr14", Expr14.ExportValue(Export))
					Call XmlDoc.AddField("Expr15", Expr15.ExportValue(Export))
					Call XmlDoc.AddField("Expr16", Expr16.ExportValue(Export))
					Call XmlDoc.AddField("Expr17", Expr17.ExportValue(Export))
					Call XmlDoc.AddField("Expr18", Expr18.ExportValue(Export))
					Call XmlDoc.AddField("Expr19", Expr19.ExportValue(Export))
					Call XmlDoc.AddField("Expr20", Expr20.ExportValue(Export))
					Call XmlDoc.AddField("Expr21", Expr21.ExportValue(Export))
					Call XmlDoc.AddField("Expr22", Expr22.ExportValue(Export))
					Call XmlDoc.AddField("Expr23", Expr23.ExportValue(Export))
					Call XmlDoc.AddField("Expr24", Expr24.ExportValue(Export))
					Call XmlDoc.AddField("Expr25", Expr25.ExportValue(Export))
					Call XmlDoc.AddField("Expr26", Expr26.ExportValue(Export))
					Call XmlDoc.AddField("Expr27", Expr27.ExportValue(Export))
					Call XmlDoc.AddField("Expr28", Expr28.ExportValue(Export))
					Call XmlDoc.AddField("Expr29", Expr29.ExportValue(Export))
					Call XmlDoc.AddField("Expr30", Expr30.ExportValue(Export))
					Call XmlDoc.AddField("Expr31", Expr31.ExportValue(Export))
					Call XmlDoc.AddField("Expr32", Expr32.ExportValue(Export))
					Call XmlDoc.AddField("Expr33", Expr33.ExportValue(Export))
					Call XmlDoc.AddField("Expr34", Expr34.ExportValue(Export))
					Call XmlDoc.AddField("Expr35", Expr35.ExportValue(Export))
					Call XmlDoc.AddField("Expr36", Expr36.ExportValue(Export))
					Call XmlDoc.AddField("Expr37", Expr37.ExportValue(Export))
				Else
					Call XmlDoc.AddField("Expr1", Expr1.ExportValue(Export))
					Call XmlDoc.AddField("Expr2", Expr2.ExportValue(Export))
					Call XmlDoc.AddField("Expr3", Expr3.ExportValue(Export))
					Call XmlDoc.AddField("Expr4", Expr4.ExportValue(Export))
					Call XmlDoc.AddField("Expr5", Expr5.ExportValue(Export))
					Call XmlDoc.AddField("Expr6", Expr6.ExportValue(Export))
					Call XmlDoc.AddField("Expr7", Expr7.ExportValue(Export))
					Call XmlDoc.AddField("Expr8", Expr8.ExportValue(Export))
					Call XmlDoc.AddField("Expr9", Expr9.ExportValue(Export))
					Call XmlDoc.AddField("Expr10", Expr10.ExportValue(Export))
					Call XmlDoc.AddField("Expr11", Expr11.ExportValue(Export))
					Call XmlDoc.AddField("Expr12", Expr12.ExportValue(Export))
					Call XmlDoc.AddField("Expr13", Expr13.ExportValue(Export))
					Call XmlDoc.AddField("Expr14", Expr14.ExportValue(Export))
					Call XmlDoc.AddField("Expr15", Expr15.ExportValue(Export))
					Call XmlDoc.AddField("Expr16", Expr16.ExportValue(Export))
					Call XmlDoc.AddField("Expr17", Expr17.ExportValue(Export))
					Call XmlDoc.AddField("Expr18", Expr18.ExportValue(Export))
					Call XmlDoc.AddField("Expr19", Expr19.ExportValue(Export))
					Call XmlDoc.AddField("Expr20", Expr20.ExportValue(Export))
					Call XmlDoc.AddField("Expr21", Expr21.ExportValue(Export))
					Call XmlDoc.AddField("Expr22", Expr22.ExportValue(Export))
					Call XmlDoc.AddField("Expr23", Expr23.ExportValue(Export))
					Call XmlDoc.AddField("Expr24", Expr24.ExportValue(Export))
					Call XmlDoc.AddField("Expr25", Expr25.ExportValue(Export))
					Call XmlDoc.AddField("Expr26", Expr26.ExportValue(Export))
					Call XmlDoc.AddField("Expr27", Expr27.ExportValue(Export))
					Call XmlDoc.AddField("Expr28", Expr28.ExportValue(Export))
					Call XmlDoc.AddField("Expr29", Expr29.ExportValue(Export))
					Call XmlDoc.AddField("Expr30", Expr30.ExportValue(Export))
					Call XmlDoc.AddField("Expr31", Expr31.ExportValue(Export))
					Call XmlDoc.AddField("Expr32", Expr32.ExportValue(Export))
					Call XmlDoc.AddField("Expr33", Expr33.ExportValue(Export))
					Call XmlDoc.AddField("Expr34", Expr34.ExportValue(Export))
					Call XmlDoc.AddField("Expr35", Expr35.ExportValue(Export))
					Call XmlDoc.AddField("Expr36", Expr36.ExportValue(Export))
					Call XmlDoc.AddField("Expr37", Expr37.ExportValue(Export))
				End If
			End If
			Recordset.MoveNext()
		Loop
	End Sub

	' Export data in HTML/CSV/Word/Excel/Email format
	Public Sub ExportDocument(Doc, Recordset, StartRec, StopRec, ExportPageType)
		If Not IsObject(Recordset) Or Not IsObject(Doc) Then
			Exit Sub
		End If

		' Write header
		Call Doc.ExportTableHeader()
		If Doc.Horizontal Then ' Horizontal format, write header
			Call Doc.BeginExportRow(0)
			If ExportPageType = "view" Then
				If Expr1.Exportable Then Call Doc.ExportCaption(Expr1)
				If Expr2.Exportable Then Call Doc.ExportCaption(Expr2)
				If Expr3.Exportable Then Call Doc.ExportCaption(Expr3)
				If Expr4.Exportable Then Call Doc.ExportCaption(Expr4)
				If Expr5.Exportable Then Call Doc.ExportCaption(Expr5)
				If Expr6.Exportable Then Call Doc.ExportCaption(Expr6)
				If Expr7.Exportable Then Call Doc.ExportCaption(Expr7)
				If Expr8.Exportable Then Call Doc.ExportCaption(Expr8)
				If Expr9.Exportable Then Call Doc.ExportCaption(Expr9)
				If Expr10.Exportable Then Call Doc.ExportCaption(Expr10)
				If Expr11.Exportable Then Call Doc.ExportCaption(Expr11)
				If Expr12.Exportable Then Call Doc.ExportCaption(Expr12)
				If Expr13.Exportable Then Call Doc.ExportCaption(Expr13)
				If Expr14.Exportable Then Call Doc.ExportCaption(Expr14)
				If Expr15.Exportable Then Call Doc.ExportCaption(Expr15)
				If Expr16.Exportable Then Call Doc.ExportCaption(Expr16)
				If Expr17.Exportable Then Call Doc.ExportCaption(Expr17)
				If Expr18.Exportable Then Call Doc.ExportCaption(Expr18)
				If Expr19.Exportable Then Call Doc.ExportCaption(Expr19)
				If Expr20.Exportable Then Call Doc.ExportCaption(Expr20)
				If Expr21.Exportable Then Call Doc.ExportCaption(Expr21)
				If Expr22.Exportable Then Call Doc.ExportCaption(Expr22)
				If Expr23.Exportable Then Call Doc.ExportCaption(Expr23)
				If Expr24.Exportable Then Call Doc.ExportCaption(Expr24)
				If Expr25.Exportable Then Call Doc.ExportCaption(Expr25)
				If Expr26.Exportable Then Call Doc.ExportCaption(Expr26)
				If Expr27.Exportable Then Call Doc.ExportCaption(Expr27)
				If Expr28.Exportable Then Call Doc.ExportCaption(Expr28)
				If Expr29.Exportable Then Call Doc.ExportCaption(Expr29)
				If Expr30.Exportable Then Call Doc.ExportCaption(Expr30)
				If Expr31.Exportable Then Call Doc.ExportCaption(Expr31)
				If Expr32.Exportable Then Call Doc.ExportCaption(Expr32)
				If Expr33.Exportable Then Call Doc.ExportCaption(Expr33)
				If Expr34.Exportable Then Call Doc.ExportCaption(Expr34)
				If Expr35.Exportable Then Call Doc.ExportCaption(Expr35)
				If Expr36.Exportable Then Call Doc.ExportCaption(Expr36)
				If Expr37.Exportable Then Call Doc.ExportCaption(Expr37)
			Else
				If Expr1.Exportable Then Call Doc.ExportCaption(Expr1)
				If Expr2.Exportable Then Call Doc.ExportCaption(Expr2)
				If Expr3.Exportable Then Call Doc.ExportCaption(Expr3)
				If Expr4.Exportable Then Call Doc.ExportCaption(Expr4)
				If Expr5.Exportable Then Call Doc.ExportCaption(Expr5)
				If Expr6.Exportable Then Call Doc.ExportCaption(Expr6)
				If Expr7.Exportable Then Call Doc.ExportCaption(Expr7)
				If Expr8.Exportable Then Call Doc.ExportCaption(Expr8)
				If Expr9.Exportable Then Call Doc.ExportCaption(Expr9)
				If Expr10.Exportable Then Call Doc.ExportCaption(Expr10)
				If Expr11.Exportable Then Call Doc.ExportCaption(Expr11)
				If Expr12.Exportable Then Call Doc.ExportCaption(Expr12)
				If Expr13.Exportable Then Call Doc.ExportCaption(Expr13)
				If Expr14.Exportable Then Call Doc.ExportCaption(Expr14)
				If Expr15.Exportable Then Call Doc.ExportCaption(Expr15)
				If Expr16.Exportable Then Call Doc.ExportCaption(Expr16)
				If Expr17.Exportable Then Call Doc.ExportCaption(Expr17)
				If Expr18.Exportable Then Call Doc.ExportCaption(Expr18)
				If Expr19.Exportable Then Call Doc.ExportCaption(Expr19)
				If Expr20.Exportable Then Call Doc.ExportCaption(Expr20)
				If Expr21.Exportable Then Call Doc.ExportCaption(Expr21)
				If Expr22.Exportable Then Call Doc.ExportCaption(Expr22)
				If Expr23.Exportable Then Call Doc.ExportCaption(Expr23)
				If Expr24.Exportable Then Call Doc.ExportCaption(Expr24)
				If Expr25.Exportable Then Call Doc.ExportCaption(Expr25)
				If Expr26.Exportable Then Call Doc.ExportCaption(Expr26)
				If Expr27.Exportable Then Call Doc.ExportCaption(Expr27)
				If Expr28.Exportable Then Call Doc.ExportCaption(Expr28)
				If Expr29.Exportable Then Call Doc.ExportCaption(Expr29)
				If Expr30.Exportable Then Call Doc.ExportCaption(Expr30)
				If Expr31.Exportable Then Call Doc.ExportCaption(Expr31)
				If Expr32.Exportable Then Call Doc.ExportCaption(Expr32)
				If Expr33.Exportable Then Call Doc.ExportCaption(Expr33)
				If Expr34.Exportable Then Call Doc.ExportCaption(Expr34)
				If Expr35.Exportable Then Call Doc.ExportCaption(Expr35)
				If Expr36.Exportable Then Call Doc.ExportCaption(Expr36)
				If Expr37.Exportable Then Call Doc.ExportCaption(Expr37)
			End If
			Call Doc.EndExportRow()
		End If

		' Move to first record
		Dim RecCnt, RowCnt
		RecCnt = StartRec - 1
		If Not Recordset.Eof Then
			Recordset.MoveFirst()
			If StartRec > 1 Then Recordset.Move(StartRec - 1)
		End If
		Do While Not Recordset.Eof And CLng(RecCnt) < CLng(StopRec)
			RecCnt = RecCnt + 1
			If CLng(RecCnt) >= CLng(StartRec) Then
				RowCnt = CLng(RecCnt) - CLng(StartRec) + 1

				' Page break
				If ExportPageBreakCount > 0 Then
					If RowCnt > 1 And ((RowCnt - 1) Mod ExportPageBreakCount = 0) Then
						Call Doc.ExportPageBreak()
					End If
				End If
				Call LoadListRowValues(Recordset)

				' Render row
				RowType = EW_ROWTYPE_VIEW ' Render view
				Call ResetAttrs()
				Call RenderListRow()
				Call Doc.BeginExportRow(RowCnt)
				If ExportPageType = "view" Then
					If Expr1.Exportable Then Call Doc.ExportField(Expr1)
					If Expr2.Exportable Then Call Doc.ExportField(Expr2)
					If Expr3.Exportable Then Call Doc.ExportField(Expr3)
					If Expr4.Exportable Then Call Doc.ExportField(Expr4)
					If Expr5.Exportable Then Call Doc.ExportField(Expr5)
					If Expr6.Exportable Then Call Doc.ExportField(Expr6)
					If Expr7.Exportable Then Call Doc.ExportField(Expr7)
					If Expr8.Exportable Then Call Doc.ExportField(Expr8)
					If Expr9.Exportable Then Call Doc.ExportField(Expr9)
					If Expr10.Exportable Then Call Doc.ExportField(Expr10)
					If Expr11.Exportable Then Call Doc.ExportField(Expr11)
					If Expr12.Exportable Then Call Doc.ExportField(Expr12)
					If Expr13.Exportable Then Call Doc.ExportField(Expr13)
					If Expr14.Exportable Then Call Doc.ExportField(Expr14)
					If Expr15.Exportable Then Call Doc.ExportField(Expr15)
					If Expr16.Exportable Then Call Doc.ExportField(Expr16)
					If Expr17.Exportable Then Call Doc.ExportField(Expr17)
					If Expr18.Exportable Then Call Doc.ExportField(Expr18)
					If Expr19.Exportable Then Call Doc.ExportField(Expr19)
					If Expr20.Exportable Then Call Doc.ExportField(Expr20)
					If Expr21.Exportable Then Call Doc.ExportField(Expr21)
					If Expr22.Exportable Then Call Doc.ExportField(Expr22)
					If Expr23.Exportable Then Call Doc.ExportField(Expr23)
					If Expr24.Exportable Then Call Doc.ExportField(Expr24)
					If Expr25.Exportable Then Call Doc.ExportField(Expr25)
					If Expr26.Exportable Then Call Doc.ExportField(Expr26)
					If Expr27.Exportable Then Call Doc.ExportField(Expr27)
					If Expr28.Exportable Then Call Doc.ExportField(Expr28)
					If Expr29.Exportable Then Call Doc.ExportField(Expr29)
					If Expr30.Exportable Then Call Doc.ExportField(Expr30)
					If Expr31.Exportable Then Call Doc.ExportField(Expr31)
					If Expr32.Exportable Then Call Doc.ExportField(Expr32)
					If Expr33.Exportable Then Call Doc.ExportField(Expr33)
					If Expr34.Exportable Then Call Doc.ExportField(Expr34)
					If Expr35.Exportable Then Call Doc.ExportField(Expr35)
					If Expr36.Exportable Then Call Doc.ExportField(Expr36)
					If Expr37.Exportable Then Call Doc.ExportField(Expr37)
				Else
					If Expr1.Exportable Then Call Doc.ExportField(Expr1)
					If Expr2.Exportable Then Call Doc.ExportField(Expr2)
					If Expr3.Exportable Then Call Doc.ExportField(Expr3)
					If Expr4.Exportable Then Call Doc.ExportField(Expr4)
					If Expr5.Exportable Then Call Doc.ExportField(Expr5)
					If Expr6.Exportable Then Call Doc.ExportField(Expr6)
					If Expr7.Exportable Then Call Doc.ExportField(Expr7)
					If Expr8.Exportable Then Call Doc.ExportField(Expr8)
					If Expr9.Exportable Then Call Doc.ExportField(Expr9)
					If Expr10.Exportable Then Call Doc.ExportField(Expr10)
					If Expr11.Exportable Then Call Doc.ExportField(Expr11)
					If Expr12.Exportable Then Call Doc.ExportField(Expr12)
					If Expr13.Exportable Then Call Doc.ExportField(Expr13)
					If Expr14.Exportable Then Call Doc.ExportField(Expr14)
					If Expr15.Exportable Then Call Doc.ExportField(Expr15)
					If Expr16.Exportable Then Call Doc.ExportField(Expr16)
					If Expr17.Exportable Then Call Doc.ExportField(Expr17)
					If Expr18.Exportable Then Call Doc.ExportField(Expr18)
					If Expr19.Exportable Then Call Doc.ExportField(Expr19)
					If Expr20.Exportable Then Call Doc.ExportField(Expr20)
					If Expr21.Exportable Then Call Doc.ExportField(Expr21)
					If Expr22.Exportable Then Call Doc.ExportField(Expr22)
					If Expr23.Exportable Then Call Doc.ExportField(Expr23)
					If Expr24.Exportable Then Call Doc.ExportField(Expr24)
					If Expr25.Exportable Then Call Doc.ExportField(Expr25)
					If Expr26.Exportable Then Call Doc.ExportField(Expr26)
					If Expr27.Exportable Then Call Doc.ExportField(Expr27)
					If Expr28.Exportable Then Call Doc.ExportField(Expr28)
					If Expr29.Exportable Then Call Doc.ExportField(Expr29)
					If Expr30.Exportable Then Call Doc.ExportField(Expr30)
					If Expr31.Exportable Then Call Doc.ExportField(Expr31)
					If Expr32.Exportable Then Call Doc.ExportField(Expr32)
					If Expr33.Exportable Then Call Doc.ExportField(Expr33)
					If Expr34.Exportable Then Call Doc.ExportField(Expr34)
					If Expr35.Exportable Then Call Doc.ExportField(Expr35)
					If Expr36.Exportable Then Call Doc.ExportField(Expr36)
					If Expr37.Exportable Then Call Doc.ExportField(Expr37)
				End If
				Call Doc.EndExportRow()
			End If
			Recordset.MoveNext()
		Loop
		Call Doc.ExportTableFooter()
	End Sub

	' Check if Anonymous User is allowed
    Private Function AllowAnonymousUser()
		Select Case EW_PAGE_ID
			Case "add", "register", "addopt"
				AllowAnonymousUser = False
			Case "edit", "update"
				AllowAnonymousUser = False
			Case "delete"
				AllowAnonymousUser = False
			Case "view"
				AllowAnonymousUser = False
			Case "search"
				AllowAnonymousUser = False
			Case Else
				AllowAnonymousUser = False
		End Select
	End Function

	Public Function ApplyUserIDFilters(Filter)

		' Add user id filter
		Dim sFilter
		sFilter = Filter
		ApplyUserIDFilters = sFilter
	End Function

	' Check if User ID security allows view all
	Dim UserIDAllowSecurity

	Function UserIDAllow(id)
		Dim allow
		allow = EW_USER_ID_ALLOW
		Select Case id
			Case "add", "copy", "gridadd", "register", "addopt"
				UserIDAllow = ((allow And EW_ALLOW_ADD) = EW_ALLOW_ADD)
			Case "edit", "gridedit", "update", "changepwd", "forgotpwd"
				UserIDAllow = ((allow And EW_ALLOW_EDIT) = EW_ALLOW_EDIT)
			Case "delete"
				UserIDAllow = ((allow And EW_ALLOW_DELETE) = EW_ALLOW_DELETE)
			Case "view"
				UserIDAllow = ((allow And EW_ALLOW_VIEW) = EW_ALLOW_VIEW)
			Case "search"
				UserIDAllow = ((allow And EW_ALLOW_SEARCH) = EW_ALLOW_SEARCH)
			Case Else
				UserIDAllow = ((allow And EW_ALLOW_LIST) = EW_ALLOW_LIST)
		End Select
	End Function
	Dim CurrentAction ' Current action
	Dim LastAction ' Last action
	Dim CurrentMode ' Current mode
	Dim UpdateConflict ' Update conflict
	Dim EventName ' Event name
	Dim EventCancelled ' Event cancelled
	Dim CancelMessage ' Cancel message
	Dim AllowAddDeleteRow ' Allow add/delete row
	Dim ValidateKey ' Validate key
	Dim DetailAdd ' Allow detail add
	Dim DetailEdit ' Allow detail edit
	Dim DetailView ' Allow detail view
	Dim ShowMultipleDetails ' Show multiple details
	Dim GridAddRowCount ' Grid add row count
	Dim CustomActions ' Custom action array

	' Check current action
	' - Add
	Public Function IsAdd()
		IsAdd = (CurrentAction = "add")
	End Function

	' - Copy
	Public Function IsCopy()
		IsCopy = (CurrentAction = "copy" Or CurrentAction = "C")
	End Function

	' - Edit
	Public Function IsEdit()
		IsEdit = (CurrentAction = "edit")
	End Function

	' - Delete
	Public Function IsDelete()
		IsDelete = (CurrentAction = "D")
	End Function

	' - Confirm
	Public Function IsConfirm()
		IsConfirm = (CurrentAction = "F")
	End Function

	' - Overwrite
	Public Function IsOverwrite()
		IsOverwrite = (CurrentAction = "overwrite")
	End Function

	' - Cancel
	Public Function IsCancel()
		IsCancel = (CurrentAction = "cancel")
	End Function

	' - Grid add
	Public Function IsGridAdd()
		IsGridAdd = (CurrentAction = "gridadd")
	End Function

	' - Grid edit
	Public Function IsGridEdit()
		IsGridEdit = (CurrentAction = "gridedit")
	End Function

	' - Add/Copy/Edit/GridAdd/GridEdit
	Public Function IsAddOrEdit()
		IsAddOrEdit = IsAdd() Or IsCopy() Or IsEdit() Or IsGridAdd() Or IsGridEdit()
	End Function

	' - Insert
	Public Function IsInsert()
		IsInsert = (CurrentAction = "insert" Or CurrentAction = "A")
	End Function

	' - Update
	Public Function IsUpdate()
		IsUpdate = (CurrentAction = "update" Or CurrentAction = "U")
	End Function

	' - Grid update
	Public Function IsGridUpdate()
		IsGridUpdate = (CurrentAction = "gridupdate")
	End Function

	' - Grid insert
	Public Function IsGridInsert()
		IsGridInsert = (CurrentAction = "gridinsert")
	End Function

	' - Grid overwrite
	Public Function IsGridOverwrite()
		IsGridOverwrite = (CurrentAction = "gridoverwrite")
	End Function

	' Check last action
	' - Cancelled
	Public Function IsCancelled()
		IsCancelled = (LastAction = "cancel" And CurrentAction = "")
	End Function

	' - Inline inserted
	Public Function IsInlineInserted()
		IsInlineInserted = (LastAction = "insert" And CurrentAction = "")
	End Function

	' - Inline updated
	Public Function IsInlineUpdated()
		IsInlineUpdated = (LastAction = "update" And CurrentAction = "")
	End Function

	' - Grid updated
	Public Function IsGridUpdated()
		IsGridUpdated = (LastAction = "gridupdate" And CurrentAction = "")
	End Function

	' - Grid inserted
	Public Function IsGridInserted()
		IsGridInserted = (LastAction = "gridinsert" And CurrentAction = "")
	End Function

	' Row Type
	Private m_RowType

	Public Property Get RowType()
		RowType = m_RowType
	End Property

	Public Property Let RowType(v)
		m_RowType = v
	End Property
	Dim CssClass ' Css class
	Dim CssStyle' Css style

'	Dim RowClientEvents ' Row client events
	Dim RowAttrs ' Row attributes

	' Row Styles
	Public Property Get RowStyles()
		Dim sAtt, Value
		Dim sStyle, sClass
		sAtt = ""
		sStyle = CssStyle
		If RowAttrs.Exists("style") Then
			Value = RowAttrs.Item("style")
			If Trim(Value) <> "" Then
				sStyle = sStyle & " " & Value
			End If
		End If
		sClass = CssClass
		If RowAttrs.Exists("class") Then
			Value = RowAttrs.Item("class")
			If Trim(Value) <> "" Then
				sClass = sClass & " " & Value
			End If
		End If
		If Trim(sStyle) <> "" Then
			sAtt = sAtt & " style=""" & Trim(sStyle) & """" 
		End If
		If Trim(sClass) <> "" Then
			sAtt = sAtt & " class=""" & Trim(sClass) & """" 
		End If
		RowStyles = sAtt
	End Property

	' Row Attribute
	Public Property Get RowAttributes()
		Dim sAtt, Attr, Value, i
		sAtt = RowStyles
		If m_Export = "" Then

'			If Trim(RowClientEvents) <> "" Then
'				sAtt = sAtt & " " & Trim(RowClientEvents)
'			End If

			For i = 0 to UBound(RowAttrs.Attributes)
				Attr = RowAttrs.Attributes(i)(0)
				Value = RowAttrs.Attributes(i)(1)
				If Attr <> "style" And Attr <> "class" And Attr <> "" And Value <> "" Then
					sAtt = sAtt & " " & Attr & "=""" & Value & """"
				End If
			Next
		End If
		RowAttributes = sAtt
	End Property

	' Export
	Private m_Export

	Public Property Get Export()
		Export = m_Export
	End Property

	Public Property Let Export(v)
		m_Export = v
	End Property

	' Export All
	Dim ExportAll
	Dim ExportPageBreakCount ' Page break per every n record (PDF only)
	Dim ExportPageOrientation ' Page orientation (PDF only)
	Dim ExportPageSize ' Page size (PDF only)
	Dim PrinterFriendlyForPdf ' Use printer friendly layout for PDF '???

	' Send Email
	Dim SendEmail

	' Custom Inner Html
	Dim TableCustomInnerHtml

	' ----------------
	'  Field objects
	' ----------------
	' Field Expr1
	Private m_Expr1

	Public Property Get Expr1()
		If Not IsObject(m_Expr1) Then
			Set m_Expr1 = NewFldObj("Query2", "Query2", "x_Expr1", "Expr1", "[Expr1]", "[Expr1]", 202, 0, "[Expr1]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr1 = m_Expr1
	End Property

	' Field Expr2
	Private m_Expr2

	Public Property Get Expr2()
		If Not IsObject(m_Expr2) Then
			Set m_Expr2 = NewFldObj("Query2", "Query2", "x_Expr2", "Expr2", "[Expr2]", "[Expr2]", 202, 0, "[Expr2]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr2 = m_Expr2
	End Property

	' Field Expr3
	Private m_Expr3

	Public Property Get Expr3()
		If Not IsObject(m_Expr3) Then
			Set m_Expr3 = NewFldObj("Query2", "Query2", "x_Expr3", "Expr3", "[Expr3]", "[Expr3]", 202, 0, "[Expr3]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr3 = m_Expr3
	End Property

	' Field Expr4
	Private m_Expr4

	Public Property Get Expr4()
		If Not IsObject(m_Expr4) Then
			Set m_Expr4 = NewFldObj("Query2", "Query2", "x_Expr4", "Expr4", "[Expr4]", "[Expr4]", 202, 0, "[Expr4]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr4 = m_Expr4
	End Property

	' Field Expr5
	Private m_Expr5

	Public Property Get Expr5()
		If Not IsObject(m_Expr5) Then
			Set m_Expr5 = NewFldObj("Query2", "Query2", "x_Expr5", "Expr5", "[Expr5]", "[Expr5]", 202, 0, "[Expr5]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr5 = m_Expr5
	End Property

	' Field Expr6
	Private m_Expr6

	Public Property Get Expr6()
		If Not IsObject(m_Expr6) Then
			Set m_Expr6 = NewFldObj("Query2", "Query2", "x_Expr6", "Expr6", "[Expr6]", "[Expr6]", 202, 0, "[Expr6]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr6 = m_Expr6
	End Property

	' Field Expr7
	Private m_Expr7

	Public Property Get Expr7()
		If Not IsObject(m_Expr7) Then
			Set m_Expr7 = NewFldObj("Query2", "Query2", "x_Expr7", "Expr7", "[Expr7]", "[Expr7]", 202, 0, "[Expr7]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr7 = m_Expr7
	End Property

	' Field Expr8
	Private m_Expr8

	Public Property Get Expr8()
		If Not IsObject(m_Expr8) Then
			Set m_Expr8 = NewFldObj("Query2", "Query2", "x_Expr8", "Expr8", "[Expr8]", "[Expr8]", 202, 0, "[Expr8]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr8 = m_Expr8
	End Property

	' Field Expr9
	Private m_Expr9

	Public Property Get Expr9()
		If Not IsObject(m_Expr9) Then
			Set m_Expr9 = NewFldObj("Query2", "Query2", "x_Expr9", "Expr9", "[Expr9]", "[Expr9]", 202, 0, "[Expr9]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr9 = m_Expr9
	End Property

	' Field Expr10
	Private m_Expr10

	Public Property Get Expr10()
		If Not IsObject(m_Expr10) Then
			Set m_Expr10 = NewFldObj("Query2", "Query2", "x_Expr10", "Expr10", "[Expr10]", "[Expr10]", 202, 0, "[Expr10]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr10 = m_Expr10
	End Property

	' Field Expr11
	Private m_Expr11

	Public Property Get Expr11()
		If Not IsObject(m_Expr11) Then
			Set m_Expr11 = NewFldObj("Query2", "Query2", "x_Expr11", "Expr11", "[Expr11]", "[Expr11]", 202, 0, "[Expr11]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr11 = m_Expr11
	End Property

	' Field Expr12
	Private m_Expr12

	Public Property Get Expr12()
		If Not IsObject(m_Expr12) Then
			Set m_Expr12 = NewFldObj("Query2", "Query2", "x_Expr12", "Expr12", "[Expr12]", "[Expr12]", 202, 0, "[Expr12]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr12 = m_Expr12
	End Property

	' Field Expr13
	Private m_Expr13

	Public Property Get Expr13()
		If Not IsObject(m_Expr13) Then
			Set m_Expr13 = NewFldObj("Query2", "Query2", "x_Expr13", "Expr13", "[Expr13]", "[Expr13]", 202, 0, "[Expr13]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr13 = m_Expr13
	End Property

	' Field Expr14
	Private m_Expr14

	Public Property Get Expr14()
		If Not IsObject(m_Expr14) Then
			Set m_Expr14 = NewFldObj("Query2", "Query2", "x_Expr14", "Expr14", "[Expr14]", "[Expr14]", 202, 0, "[Expr14]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr14 = m_Expr14
	End Property

	' Field Expr15
	Private m_Expr15

	Public Property Get Expr15()
		If Not IsObject(m_Expr15) Then
			Set m_Expr15 = NewFldObj("Query2", "Query2", "x_Expr15", "Expr15", "[Expr15]", "[Expr15]", 202, 0, "[Expr15]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr15 = m_Expr15
	End Property

	' Field Expr16
	Private m_Expr16

	Public Property Get Expr16()
		If Not IsObject(m_Expr16) Then
			Set m_Expr16 = NewFldObj("Query2", "Query2", "x_Expr16", "Expr16", "[Expr16]", "[Expr16]", 202, 0, "[Expr16]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr16 = m_Expr16
	End Property

	' Field Expr17
	Private m_Expr17

	Public Property Get Expr17()
		If Not IsObject(m_Expr17) Then
			Set m_Expr17 = NewFldObj("Query2", "Query2", "x_Expr17", "Expr17", "[Expr17]", "[Expr17]", 3, 0, "[Expr17]", False, False, FALSE, "FORMATTED TEXT")
			m_Expr17.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set Expr17 = m_Expr17
	End Property

	' Field Expr18
	Private m_Expr18

	Public Property Get Expr18()
		If Not IsObject(m_Expr18) Then
			Set m_Expr18 = NewFldObj("Query2", "Query2", "x_Expr18", "Expr18", "[Expr18]", "[Expr18]", 3, 0, "[Expr18]", False, False, FALSE, "FORMATTED TEXT")
			m_Expr18.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set Expr18 = m_Expr18
	End Property

	' Field Expr19
	Private m_Expr19

	Public Property Get Expr19()
		If Not IsObject(m_Expr19) Then
			Set m_Expr19 = NewFldObj("Query2", "Query2", "x_Expr19", "Expr19", "[Expr19]", "[Expr19]", 3, 0, "[Expr19]", False, False, FALSE, "FORMATTED TEXT")
			m_Expr19.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set Expr19 = m_Expr19
	End Property

	' Field Expr20
	Private m_Expr20

	Public Property Get Expr20()
		If Not IsObject(m_Expr20) Then
			Set m_Expr20 = NewFldObj("Query2", "Query2", "x_Expr20", "Expr20", "[Expr20]", "[Expr20]", 3, 0, "[Expr20]", False, False, FALSE, "FORMATTED TEXT")
			m_Expr20.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set Expr20 = m_Expr20
	End Property

	' Field Expr21
	Private m_Expr21

	Public Property Get Expr21()
		If Not IsObject(m_Expr21) Then
			Set m_Expr21 = NewFldObj("Query2", "Query2", "x_Expr21", "Expr21", "[Expr21]", "[Expr21]", 3, 0, "[Expr21]", False, False, FALSE, "FORMATTED TEXT")
			m_Expr21.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set Expr21 = m_Expr21
	End Property

	' Field Expr22
	Private m_Expr22

	Public Property Get Expr22()
		If Not IsObject(m_Expr22) Then
			Set m_Expr22 = NewFldObj("Query2", "Query2", "x_Expr22", "Expr22", "[Expr22]", "[Expr22]", 3, 0, "[Expr22]", False, False, FALSE, "FORMATTED TEXT")
			m_Expr22.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set Expr22 = m_Expr22
	End Property

	' Field Expr23
	Private m_Expr23

	Public Property Get Expr23()
		If Not IsObject(m_Expr23) Then
			Set m_Expr23 = NewFldObj("Query2", "Query2", "x_Expr23", "Expr23", "[Expr23]", "[Expr23]", 202, 0, "[Expr23]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr23 = m_Expr23
	End Property

	' Field Expr24
	Private m_Expr24

	Public Property Get Expr24()
		If Not IsObject(m_Expr24) Then
			Set m_Expr24 = NewFldObj("Query2", "Query2", "x_Expr24", "Expr24", "[Expr24]", "[Expr24]", 3, 0, "[Expr24]", False, False, FALSE, "FORMATTED TEXT")
			m_Expr24.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set Expr24 = m_Expr24
	End Property

	' Field Expr25
	Private m_Expr25

	Public Property Get Expr25()
		If Not IsObject(m_Expr25) Then
			Set m_Expr25 = NewFldObj("Query2", "Query2", "x_Expr25", "Expr25", "[Expr25]", "[Expr25]", 202, 0, "[Expr25]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr25 = m_Expr25
	End Property

	' Field Expr26
	Private m_Expr26

	Public Property Get Expr26()
		If Not IsObject(m_Expr26) Then
			Set m_Expr26 = NewFldObj("Query2", "Query2", "x_Expr26", "Expr26", "[Expr26]", "[Expr26]", 202, 0, "[Expr26]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr26 = m_Expr26
	End Property

	' Field Expr27
	Private m_Expr27

	Public Property Get Expr27()
		If Not IsObject(m_Expr27) Then
			Set m_Expr27 = NewFldObj("Query2", "Query2", "x_Expr27", "Expr27", "[Expr27]", "[Expr27]", 202, 0, "[Expr27]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr27 = m_Expr27
	End Property

	' Field Expr28
	Private m_Expr28

	Public Property Get Expr28()
		If Not IsObject(m_Expr28) Then
			Set m_Expr28 = NewFldObj("Query2", "Query2", "x_Expr28", "Expr28", "[Expr28]", "[Expr28]", 202, 0, "[Expr28]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr28 = m_Expr28
	End Property

	' Field Expr29
	Private m_Expr29

	Public Property Get Expr29()
		If Not IsObject(m_Expr29) Then
			Set m_Expr29 = NewFldObj("Query2", "Query2", "x_Expr29", "Expr29", "[Expr29]", "[Expr29]", 202, 0, "[Expr29]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr29 = m_Expr29
	End Property

	' Field Expr30
	Private m_Expr30

	Public Property Get Expr30()
		If Not IsObject(m_Expr30) Then
			Set m_Expr30 = NewFldObj("Query2", "Query2", "x_Expr30", "Expr30", "[Expr30]", "[Expr30]", 202, 0, "[Expr30]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr30 = m_Expr30
	End Property

	' Field Expr31
	Private m_Expr31

	Public Property Get Expr31()
		If Not IsObject(m_Expr31) Then
			Set m_Expr31 = NewFldObj("Query2", "Query2", "x_Expr31", "Expr31", "[Expr31]", "[Expr31]", 202, 0, "[Expr31]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr31 = m_Expr31
	End Property

	' Field Expr32
	Private m_Expr32

	Public Property Get Expr32()
		If Not IsObject(m_Expr32) Then
			Set m_Expr32 = NewFldObj("Query2", "Query2", "x_Expr32", "Expr32", "[Expr32]", "[Expr32]", 202, 0, "[Expr32]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr32 = m_Expr32
	End Property

	' Field Expr33
	Private m_Expr33

	Public Property Get Expr33()
		If Not IsObject(m_Expr33) Then
			Set m_Expr33 = NewFldObj("Query2", "Query2", "x_Expr33", "Expr33", "[Expr33]", "FORMAT([Expr33], '')", 135, 8, "[Expr33]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr33 = m_Expr33
	End Property

	' Field Expr34
	Private m_Expr34

	Public Property Get Expr34()
		If Not IsObject(m_Expr34) Then
			Set m_Expr34 = NewFldObj("Query2", "Query2", "x_Expr34", "Expr34", "[Expr34]", "[Expr34]", 202, 0, "[Expr34]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr34 = m_Expr34
	End Property

	' Field Expr35
	Private m_Expr35

	Public Property Get Expr35()
		If Not IsObject(m_Expr35) Then
			Set m_Expr35 = NewFldObj("Query2", "Query2", "x_Expr35", "Expr35", "[Expr35]", "FORMAT([Expr35], '')", 135, 8, "[Expr35]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr35 = m_Expr35
	End Property

	' Field Expr36
	Private m_Expr36

	Public Property Get Expr36()
		If Not IsObject(m_Expr36) Then
			Set m_Expr36 = NewFldObj("Query2", "Query2", "x_Expr36", "Expr36", "[Expr36]", "[Expr36]", 202, 0, "[Expr36]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr36 = m_Expr36
	End Property

	' Field Expr37
	Private m_Expr37

	Public Property Get Expr37()
		If Not IsObject(m_Expr37) Then
			Set m_Expr37 = NewFldObj("Query2", "Query2", "x_Expr37", "Expr37", "[Expr37]", "[Expr37]", 202, 0, "[Expr37]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set Expr37 = m_Expr37
	End Property
	Dim Fields ' Fields

	' Get field object by name
	Public Function GetField(Name)
		Dim fld, i
		Set fld = Nothing
		For i = 0 to UBound(Fields,2)
			If Fields(0,i) = Name Then
				Set fld = Fields(1,i)
				Exit For
			End If
		Next
		Set GetField = fld
	End Function

	' Create new field object
	Private Function NewFldObj(TblVar, TblName, FldVar, FldName, FldExpression, FldBasicSearchExpression, FldType, FldDtFormat, FldVirtualExp, FldVirtual, FldForceSelect, FldVirtualSearch, FldViewTag)
		Dim fld
		Set fld = New cField
		fld.TblVar = TblVar
		fld.TblName = TblName
		fld.FldVar = FldVar
		fld.FldName = FldName
		fld.FldExpression = FldExpression
		fld.FldBasicSearchExpression = FldBasicSearchExpression
		fld.FldType = FldType
		fld.FldDataType = ew_FieldDataType(FldType)
		fld.FldDateTimeFormat = FldDtFormat
		fld.FldVirtualExpression = FldVirtualExp
		fld.FldIsVirtual = FldVirtual
		fld.FldForceSelection = FldForceSelect
		fld.FldVirtualSearch = FldVirtualSearch
		fld.FldViewTag = FldViewTag
		fld.AdvancedSearch.TblVar = TblVar
		fld.AdvancedSearch.FldVar = FldVar
		Set NewFldObj = fld
	End Function

	' Table level events
	' Recordset Selecting event
	Sub Recordset_Selecting(filter)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here	
	End Sub

	' Recordset Selected event
	Sub Recordset_Selected(rs)

		'Response.Write "Recordset Selected"
	End Sub

	' Recordset Search Validated event
	Sub Recordset_SearchValidated()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
	End Sub

	' Recordset Searching event
	Sub Recordset_Searching(filter)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here	
	End Sub

	' Row_Selecting event
	Sub Row_Selecting(filter)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here	
	End Sub

	' Row Selected event
	Sub Row_Selected(rs)

		'Response.Write "Row Selected"
	End Sub

	' Row Inserting event
	Function Row_Inserting(rsold, rsnew)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
		' To cancel, set return value to False

		Row_Inserting = True
	End Function

	' Row Inserted event
	Sub Row_Inserted(rsold, rsnew)

		' Response.Write "Row Inserted"
	End Sub

	' Row Updating event
	Function Row_Updating(rsold, rsnew)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
		' To cancel, set return value to False

		Row_Updating = True
	End Function

	' Row Updated event
	Sub Row_Updated(rsold, rsnew)

		' Response.Write "Row Updated"
	End Sub

	' Row Update Conflict event
	Function Row_UpdateConflict(rsold, rsnew)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
		' To ignore conflict, set return value to False

		Row_UpdateConflict = True
	End Function

	' Row Deleting event
	Function Row_Deleting(rs)
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here
		' To cancel, set return value to False

		Row_Deleting = True
	End Function

	' Row Deleted event
	Sub Row_Deleted(rs)

		' Response.Write "Row Deleted"
	End Sub

	' Email Sending event
	Function Email_Sending(Email, Args)

		'Response.Write Email.AsString
		'Response.Write "Keys of Args: " & Join(Args.Keys, ", ")
		'Response.End

		Email_Sending = True
	End Function

	' Lookup Selecting event
	Sub Lookup_Selecting(fld, filter)

		' Enter your code here
	End Sub

	' Row Rendering event
	Sub Row_Rendering()
		If Not EW_DEBUG_ENABLED Then On Error Resume Next

		' Enter your code here	
	End Sub

	' Row Rendered event
	Sub Row_Rendered()

		' To view properties of field class, use:
		' Response.Write <FieldName>.AsString() 

	End Sub

	' User ID Filtering event
	Sub UserID_Filtering(filter)

		' Enter your code here
	End Sub

	' Class terminate
	Private Sub Class_Terminate
		If IsObject(m_Expr1) Then Set m_Expr1 = Nothing
		If IsObject(m_Expr2) Then Set m_Expr2 = Nothing
		If IsObject(m_Expr3) Then Set m_Expr3 = Nothing
		If IsObject(m_Expr4) Then Set m_Expr4 = Nothing
		If IsObject(m_Expr5) Then Set m_Expr5 = Nothing
		If IsObject(m_Expr6) Then Set m_Expr6 = Nothing
		If IsObject(m_Expr7) Then Set m_Expr7 = Nothing
		If IsObject(m_Expr8) Then Set m_Expr8 = Nothing
		If IsObject(m_Expr9) Then Set m_Expr9 = Nothing
		If IsObject(m_Expr10) Then Set m_Expr10 = Nothing
		If IsObject(m_Expr11) Then Set m_Expr11 = Nothing
		If IsObject(m_Expr12) Then Set m_Expr12 = Nothing
		If IsObject(m_Expr13) Then Set m_Expr13 = Nothing
		If IsObject(m_Expr14) Then Set m_Expr14 = Nothing
		If IsObject(m_Expr15) Then Set m_Expr15 = Nothing
		If IsObject(m_Expr16) Then Set m_Expr16 = Nothing
		If IsObject(m_Expr17) Then Set m_Expr17 = Nothing
		If IsObject(m_Expr18) Then Set m_Expr18 = Nothing
		If IsObject(m_Expr19) Then Set m_Expr19 = Nothing
		If IsObject(m_Expr20) Then Set m_Expr20 = Nothing
		If IsObject(m_Expr21) Then Set m_Expr21 = Nothing
		If IsObject(m_Expr22) Then Set m_Expr22 = Nothing
		If IsObject(m_Expr23) Then Set m_Expr23 = Nothing
		If IsObject(m_Expr24) Then Set m_Expr24 = Nothing
		If IsObject(m_Expr25) Then Set m_Expr25 = Nothing
		If IsObject(m_Expr26) Then Set m_Expr26 = Nothing
		If IsObject(m_Expr27) Then Set m_Expr27 = Nothing
		If IsObject(m_Expr28) Then Set m_Expr28 = Nothing
		If IsObject(m_Expr29) Then Set m_Expr29 = Nothing
		If IsObject(m_Expr30) Then Set m_Expr30 = Nothing
		If IsObject(m_Expr31) Then Set m_Expr31 = Nothing
		If IsObject(m_Expr32) Then Set m_Expr32 = Nothing
		If IsObject(m_Expr33) Then Set m_Expr33 = Nothing
		If IsObject(m_Expr34) Then Set m_Expr34 = Nothing
		If IsObject(m_Expr35) Then Set m_Expr35 = Nothing
		If IsObject(m_Expr36) Then Set m_Expr36 = Nothing
		If IsObject(m_Expr37) Then Set m_Expr37 = Nothing
		Set RowAttrs = Nothing
	End Sub
End Class
%>
