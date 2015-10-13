<%

' ASPMaker configuration for Table journal
Dim journal

' Define table class
Class cjournal

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
		Call ew_SetArObj(Fields, "jrl_id", jrl_id)
		Call ew_SetArObj(Fields, "jrl_category", jrl_category)
		Call ew_SetArObj(Fields, "jrl_date", jrl_date)
		Call ew_SetArObj(Fields, "jrl_title", jrl_title)
		Call ew_SetArObj(Fields, "jrl_title_th", jrl_title_th)
		Call ew_SetArObj(Fields, "jrl_pdf", jrl_pdf)
		Call ew_SetArObj(Fields, "jrl_img", jrl_img)
		Call ew_SetArObj(Fields, "jrl_create", jrl_create)
		Call ew_SetArObj(Fields, "jrl_update", jrl_update)
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
		TableVar = "journal"
	End Property

	' Table name
	Public Property Get TableName()
		TableName = "journal"
	End Property

	' Table type
	Public Property Get TableType()
		TableType = "TABLE"
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
		HighlightName = "journal_Highlight"
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
		SqlSelect = "SELECT * FROM [journal]"
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
		SqlKeyFilter = "[jrl_id] = @jrl_id@"
	End Property

	' Return Key filter for table
	Public Property Get KeyFilter()
		Dim sKeyFilter
		sKeyFilter = SqlKeyFilter
		If Not IsNumeric(jrl_id.CurrentValue) Then
			sKeyFilter = "0=1" ' Invalid key
		End If
		sKeyFilter = Replace(sKeyFilter, "@jrl_id@", ew_AdjustSql(jrl_id.CurrentValue)) ' Replace key value
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
			ReturnUrl = "pom_journallist.asp"
		End If
	End Property

	' List url
	Public Function ListUrl()
		ListUrl = "pom_journallist.asp"
	End Function

	' View url
	Public Function ViewUrl(parm)
		If parm <> "" Then
			ViewUrl = KeyUrl("pom_journalview.asp", UrlParm(parm))
		Else
			ViewUrl = KeyUrl("pom_journalview.asp", UrlParm(EW_TABLE_SHOW_DETAIL & "="))
		End If
	End Function

	' Add url
	Public Function AddUrl()
		AddUrl = "pom_journaladd.asp"

'		Dim sUrlParm
'		sUrlParm = UrlParm("")
'		If sUrlParm <> "" Then AddUrl = AddUrl & "?" & sUrlParm

	End Function

	' Edit url
	Public Function EditUrl(parm)
		EditUrl = KeyUrl("pom_journaledit.asp", UrlParm(parm))
	End Function

	' Inline edit url
	Public Function InlineEditUrl()
		InlineEditUrl = KeyUrl(ew_CurrentPage, UrlParm("a=edit"))
	End Function

	' Copy url
	Public Function CopyUrl(parm)
		CopyUrl = KeyUrl("pom_journaladd.asp", UrlParm(parm))
	End Function

	' Inline copy url
	Public Function InlineCopyUrl()
		InlineCopyUrl = KeyUrl(ew_CurrentPage, UrlParm("a=copy"))
	End Function

	' Delete url
	Public Function DeleteUrl()
		DeleteUrl = KeyUrl("pom_journaldelete.asp", UrlParm(""))
	End Function

	' Key url
	Public Function KeyUrl(url, parm)
		Dim sUrl: sUrl = url & "?"
		If parm <> "" Then sUrl = sUrl & parm & "&"
		If Not IsNull(jrl_id.CurrentValue) Then
			sUrl = sUrl & "jrl_id=" & jrl_id.CurrentValue
		Else
			KeyUrl = "javascript:alert(ewLanguage.Phrase('InvalidRecord'));"
			Exit Function
		End If
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
			UrlParm = "t=journal"
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
				arKeys(0) = Request.QueryString("jrl_id") ' jrl_id

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
						If Not IsNumeric(key) Then skip = True
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
				jrl_id.CurrentValue = key
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
		jrl_id.DbValue = RsRow("jrl_id")
		jrl_category.DbValue = RsRow("jrl_category")
		jrl_date.DbValue = RsRow("jrl_date")
		jrl_title.DbValue = RsRow("jrl_title")
		jrl_title_th.DbValue = RsRow("jrl_title_th")
		jrl_pdf.DbValue = RsRow("jrl_pdf")
		jrl_img.DbValue = RsRow("jrl_img")
		jrl_create.DbValue = RsRow("jrl_create")
		jrl_update.DbValue = RsRow("jrl_update")
	End Sub

	' Render list row values
	Sub RenderListRow()

		'
		'  Common render codes
		'
		' jrl_id
		' jrl_category
		' jrl_date
		' jrl_title
		' jrl_title_th
		' jrl_pdf
		' jrl_img
		' jrl_create
		' jrl_update
		' Call Row Rendering event

		Call Row_Rendering()

		'
		'  Render for View
		'
		' jrl_id

		jrl_id.ViewValue = jrl_id.CurrentValue
		jrl_id.ViewCustomAttributes = ""

		' jrl_category
		jrl_category.ViewValue = jrl_category.CurrentValue
		jrl_category.ViewCustomAttributes = ""

		' jrl_date
		jrl_date.ViewValue = jrl_date.CurrentValue
		jrl_date.ViewCustomAttributes = ""

		' jrl_title
		jrl_title.ViewValue = jrl_title.CurrentValue
		jrl_title.ViewCustomAttributes = ""

		' jrl_title_th
		jrl_title_th.ViewValue = jrl_title_th.CurrentValue
		jrl_title_th.ViewCustomAttributes = ""

		' jrl_pdf
		jrl_pdf.ViewValue = jrl_pdf.CurrentValue
		jrl_pdf.ViewCustomAttributes = ""

		' jrl_img
		jrl_img.ViewValue = jrl_img.CurrentValue
		jrl_img.ViewCustomAttributes = ""

		' jrl_create
		jrl_create.ViewValue = jrl_create.CurrentValue
		jrl_create.ViewCustomAttributes = ""

		' jrl_update
		jrl_update.ViewValue = jrl_update.CurrentValue
		jrl_update.ViewCustomAttributes = ""

		' jrl_id
		jrl_id.LinkCustomAttributes = ""
		jrl_id.HrefValue = ""
		jrl_id.TooltipValue = ""

		' jrl_category
		jrl_category.LinkCustomAttributes = ""
		jrl_category.HrefValue = ""
		jrl_category.TooltipValue = ""

		' jrl_date
		jrl_date.LinkCustomAttributes = ""
		jrl_date.HrefValue = ""
		jrl_date.TooltipValue = ""

		' jrl_title
		jrl_title.LinkCustomAttributes = ""
		jrl_title.HrefValue = ""
		jrl_title.TooltipValue = ""

		' jrl_title_th
		jrl_title_th.LinkCustomAttributes = ""
		jrl_title_th.HrefValue = ""
		jrl_title_th.TooltipValue = ""

		' jrl_pdf
		jrl_pdf.LinkCustomAttributes = ""
		jrl_pdf.HrefValue = ""
		jrl_pdf.TooltipValue = ""

		' jrl_img
		jrl_img.LinkCustomAttributes = ""
		jrl_img.HrefValue = ""
		jrl_img.TooltipValue = ""

		' jrl_create
		jrl_create.LinkCustomAttributes = ""
		jrl_create.HrefValue = ""
		jrl_create.TooltipValue = ""

		' jrl_update
		jrl_update.LinkCustomAttributes = ""
		jrl_update.HrefValue = ""
		jrl_update.TooltipValue = ""

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
			sSql = "SELECT * FROM [journal] WHERE " & sWhereList
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
					Call XmlDoc.AddField("jrl_id", jrl_id.ExportValue(Export))
					Call XmlDoc.AddField("jrl_category", jrl_category.ExportValue(Export))
					Call XmlDoc.AddField("jrl_date", jrl_date.ExportValue(Export))
					Call XmlDoc.AddField("jrl_title", jrl_title.ExportValue(Export))
					Call XmlDoc.AddField("jrl_title_th", jrl_title_th.ExportValue(Export))
					Call XmlDoc.AddField("jrl_pdf", jrl_pdf.ExportValue(Export))
					Call XmlDoc.AddField("jrl_img", jrl_img.ExportValue(Export))
					Call XmlDoc.AddField("jrl_create", jrl_create.ExportValue(Export))
					Call XmlDoc.AddField("jrl_update", jrl_update.ExportValue(Export))
				Else
					Call XmlDoc.AddField("jrl_id", jrl_id.ExportValue(Export))
					Call XmlDoc.AddField("jrl_category", jrl_category.ExportValue(Export))
					Call XmlDoc.AddField("jrl_date", jrl_date.ExportValue(Export))
					Call XmlDoc.AddField("jrl_title", jrl_title.ExportValue(Export))
					Call XmlDoc.AddField("jrl_title_th", jrl_title_th.ExportValue(Export))
					Call XmlDoc.AddField("jrl_pdf", jrl_pdf.ExportValue(Export))
					Call XmlDoc.AddField("jrl_img", jrl_img.ExportValue(Export))
					Call XmlDoc.AddField("jrl_create", jrl_create.ExportValue(Export))
					Call XmlDoc.AddField("jrl_update", jrl_update.ExportValue(Export))
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
				If jrl_id.Exportable Then Call Doc.ExportCaption(jrl_id)
				If jrl_category.Exportable Then Call Doc.ExportCaption(jrl_category)
				If jrl_date.Exportable Then Call Doc.ExportCaption(jrl_date)
				If jrl_title.Exportable Then Call Doc.ExportCaption(jrl_title)
				If jrl_title_th.Exportable Then Call Doc.ExportCaption(jrl_title_th)
				If jrl_pdf.Exportable Then Call Doc.ExportCaption(jrl_pdf)
				If jrl_img.Exportable Then Call Doc.ExportCaption(jrl_img)
				If jrl_create.Exportable Then Call Doc.ExportCaption(jrl_create)
				If jrl_update.Exportable Then Call Doc.ExportCaption(jrl_update)
			Else
				If jrl_id.Exportable Then Call Doc.ExportCaption(jrl_id)
				If jrl_category.Exportable Then Call Doc.ExportCaption(jrl_category)
				If jrl_date.Exportable Then Call Doc.ExportCaption(jrl_date)
				If jrl_title.Exportable Then Call Doc.ExportCaption(jrl_title)
				If jrl_title_th.Exportable Then Call Doc.ExportCaption(jrl_title_th)
				If jrl_pdf.Exportable Then Call Doc.ExportCaption(jrl_pdf)
				If jrl_img.Exportable Then Call Doc.ExportCaption(jrl_img)
				If jrl_create.Exportable Then Call Doc.ExportCaption(jrl_create)
				If jrl_update.Exportable Then Call Doc.ExportCaption(jrl_update)
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
					If jrl_id.Exportable Then Call Doc.ExportField(jrl_id)
					If jrl_category.Exportable Then Call Doc.ExportField(jrl_category)
					If jrl_date.Exportable Then Call Doc.ExportField(jrl_date)
					If jrl_title.Exportable Then Call Doc.ExportField(jrl_title)
					If jrl_title_th.Exportable Then Call Doc.ExportField(jrl_title_th)
					If jrl_pdf.Exportable Then Call Doc.ExportField(jrl_pdf)
					If jrl_img.Exportable Then Call Doc.ExportField(jrl_img)
					If jrl_create.Exportable Then Call Doc.ExportField(jrl_create)
					If jrl_update.Exportable Then Call Doc.ExportField(jrl_update)
				Else
					If jrl_id.Exportable Then Call Doc.ExportField(jrl_id)
					If jrl_category.Exportable Then Call Doc.ExportField(jrl_category)
					If jrl_date.Exportable Then Call Doc.ExportField(jrl_date)
					If jrl_title.Exportable Then Call Doc.ExportField(jrl_title)
					If jrl_title_th.Exportable Then Call Doc.ExportField(jrl_title_th)
					If jrl_pdf.Exportable Then Call Doc.ExportField(jrl_pdf)
					If jrl_img.Exportable Then Call Doc.ExportField(jrl_img)
					If jrl_create.Exportable Then Call Doc.ExportField(jrl_create)
					If jrl_update.Exportable Then Call Doc.ExportField(jrl_update)
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
	' Field jrl_id
	Private m_jrl_id

	Public Property Get jrl_id()
		If Not IsObject(m_jrl_id) Then
			Set m_jrl_id = NewFldObj("journal", "journal", "x_jrl_id", "jrl_id", "[jrl_id]", "[jrl_id]", 3, 0, "[jrl_id]", False, False, FALSE, "FORMATTED TEXT")
			m_jrl_id.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set jrl_id = m_jrl_id
	End Property

	' Field jrl_category
	Private m_jrl_category

	Public Property Get jrl_category()
		If Not IsObject(m_jrl_category) Then
			Set m_jrl_category = NewFldObj("journal", "journal", "x_jrl_category", "jrl_category", "[jrl_category]", "[jrl_category]", 202, 0, "[jrl_category]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set jrl_category = m_jrl_category
	End Property

	' Field jrl_date
	Private m_jrl_date

	Public Property Get jrl_date()
		If Not IsObject(m_jrl_date) Then
			Set m_jrl_date = NewFldObj("journal", "journal", "x_jrl_date", "jrl_date", "[jrl_date]", "FORMAT([jrl_date], '')", 135, 8, "[jrl_date]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set jrl_date = m_jrl_date
	End Property

	' Field jrl_title
	Private m_jrl_title

	Public Property Get jrl_title()
		If Not IsObject(m_jrl_title) Then
			Set m_jrl_title = NewFldObj("journal", "journal", "x_jrl_title", "jrl_title", "[jrl_title]", "[jrl_title]", 202, 0, "[jrl_title]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set jrl_title = m_jrl_title
	End Property

	' Field jrl_title_th
	Private m_jrl_title_th

	Public Property Get jrl_title_th()
		If Not IsObject(m_jrl_title_th) Then
			Set m_jrl_title_th = NewFldObj("journal", "journal", "x_jrl_title_th", "jrl_title_th", "[jrl_title_th]", "[jrl_title_th]", 202, 0, "[jrl_title_th]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set jrl_title_th = m_jrl_title_th
	End Property

	' Field jrl_pdf
	Private m_jrl_pdf

	Public Property Get jrl_pdf()
		If Not IsObject(m_jrl_pdf) Then
			Set m_jrl_pdf = NewFldObj("journal", "journal", "x_jrl_pdf", "jrl_pdf", "[jrl_pdf]", "[jrl_pdf]", 202, 0, "[jrl_pdf]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set jrl_pdf = m_jrl_pdf
	End Property

	' Field jrl_img
	Private m_jrl_img

	Public Property Get jrl_img()
		If Not IsObject(m_jrl_img) Then
			Set m_jrl_img = NewFldObj("journal", "journal", "x_jrl_img", "jrl_img", "[jrl_img]", "[jrl_img]", 202, 0, "[jrl_img]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set jrl_img = m_jrl_img
	End Property

	' Field jrl_create
	Private m_jrl_create

	Public Property Get jrl_create()
		If Not IsObject(m_jrl_create) Then
			Set m_jrl_create = NewFldObj("journal", "journal", "x_jrl_create", "jrl_create", "[jrl_create]", "[jrl_create]", 202, 0, "[jrl_create]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set jrl_create = m_jrl_create
	End Property

	' Field jrl_update
	Private m_jrl_update

	Public Property Get jrl_update()
		If Not IsObject(m_jrl_update) Then
			Set m_jrl_update = NewFldObj("journal", "journal", "x_jrl_update", "jrl_update", "[jrl_update]", "[jrl_update]", 202, 0, "[jrl_update]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set jrl_update = m_jrl_update
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
		If IsObject(m_jrl_id) Then Set m_jrl_id = Nothing
		If IsObject(m_jrl_category) Then Set m_jrl_category = Nothing
		If IsObject(m_jrl_date) Then Set m_jrl_date = Nothing
		If IsObject(m_jrl_title) Then Set m_jrl_title = Nothing
		If IsObject(m_jrl_title_th) Then Set m_jrl_title_th = Nothing
		If IsObject(m_jrl_pdf) Then Set m_jrl_pdf = Nothing
		If IsObject(m_jrl_img) Then Set m_jrl_img = Nothing
		If IsObject(m_jrl_create) Then Set m_jrl_create = Nothing
		If IsObject(m_jrl_update) Then Set m_jrl_update = Nothing
		Set RowAttrs = Nothing
	End Sub
End Class
%>
