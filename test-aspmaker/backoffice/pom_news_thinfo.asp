<%

' ASPMaker configuration for Table news_th
Dim news_th

' Define table class
Class cnews_th

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
		Call ew_SetArObj(Fields, "news_id", news_id)
		Call ew_SetArObj(Fields, "news_img", news_img)
		Call ew_SetArObj(Fields, "news_date", news_date)
		Call ew_SetArObj(Fields, "news_category", news_category)
		Call ew_SetArObj(Fields, "news_category_sub", news_category_sub)
		Call ew_SetArObj(Fields, "start_date", start_date)
		Call ew_SetArObj(Fields, "end_date", end_date)
		Call ew_SetArObj(Fields, "news_pdf", news_pdf)
		Call ew_SetArObj(Fields, "news_subject", news_subject)
		Call ew_SetArObj(Fields, "news_subject_th", news_subject_th)
		Call ew_SetArObj(Fields, "news_intro", news_intro)
		Call ew_SetArObj(Fields, "news_intro_th", news_intro_th)
		Call ew_SetArObj(Fields, "news_content", news_content)
		Call ew_SetArObj(Fields, "news_content_th", news_content_th)
		Call ew_SetArObj(Fields, "news_show_en", news_show_en)
		Call ew_SetArObj(Fields, "news_show", news_show)
		Call ew_SetArObj(Fields, "news_show_home", news_show_home)
		Call ew_SetArObj(Fields, "news_create", news_create)
		Call ew_SetArObj(Fields, "news_update", news_update)
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
		TableVar = "news_th"
	End Property

	' Table name
	Public Property Get TableName()
		TableName = "news_th"
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
		HighlightName = "news_th_Highlight"
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
		SqlSelect = "SELECT * FROM [news_th]"
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
		SqlKeyFilter = "[news_id] = @news_id@"
	End Property

	' Return Key filter for table
	Public Property Get KeyFilter()
		Dim sKeyFilter
		sKeyFilter = SqlKeyFilter
		If Not IsNumeric(news_id.CurrentValue) Then
			sKeyFilter = "0=1" ' Invalid key
		End If
		sKeyFilter = Replace(sKeyFilter, "@news_id@", ew_AdjustSql(news_id.CurrentValue)) ' Replace key value
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
			ReturnUrl = "pom_news_thlist.asp"
		End If
	End Property

	' List url
	Public Function ListUrl()
		ListUrl = "pom_news_thlist.asp"
	End Function

	' View url
	Public Function ViewUrl(parm)
		If parm <> "" Then
			ViewUrl = KeyUrl("pom_news_thview.asp", UrlParm(parm))
		Else
			ViewUrl = KeyUrl("pom_news_thview.asp", UrlParm(EW_TABLE_SHOW_DETAIL & "="))
		End If
	End Function

	' Add url
	Public Function AddUrl()
		AddUrl = "pom_news_thadd.asp"

'		Dim sUrlParm
'		sUrlParm = UrlParm("")
'		If sUrlParm <> "" Then AddUrl = AddUrl & "?" & sUrlParm

	End Function

	' Edit url
	Public Function EditUrl(parm)
		EditUrl = KeyUrl("pom_news_thedit.asp", UrlParm(parm))
	End Function

	' Inline edit url
	Public Function InlineEditUrl()
		InlineEditUrl = KeyUrl(ew_CurrentPage, UrlParm("a=edit"))
	End Function

	' Copy url
	Public Function CopyUrl(parm)
		CopyUrl = KeyUrl("pom_news_thadd.asp", UrlParm(parm))
	End Function

	' Inline copy url
	Public Function InlineCopyUrl()
		InlineCopyUrl = KeyUrl(ew_CurrentPage, UrlParm("a=copy"))
	End Function

	' Delete url
	Public Function DeleteUrl()
		DeleteUrl = KeyUrl("pom_news_thdelete.asp", UrlParm(""))
	End Function

	' Key url
	Public Function KeyUrl(url, parm)
		Dim sUrl: sUrl = url & "?"
		If parm <> "" Then sUrl = sUrl & parm & "&"
		If Not IsNull(news_id.CurrentValue) Then
			sUrl = sUrl & "news_id=" & news_id.CurrentValue
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
			UrlParm = "t=news_th"
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
				arKeys(0) = Request.QueryString("news_id") ' news_id

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
				news_id.CurrentValue = key
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
		news_id.DbValue = RsRow("news_id")
		news_img.DbValue = RsRow("news_img")
		news_date.DbValue = RsRow("news_date")
		news_category.DbValue = RsRow("news_category")
		news_category_sub.DbValue = RsRow("news_category_sub")
		start_date.DbValue = RsRow("start_date")
		end_date.DbValue = RsRow("end_date")
		news_pdf.DbValue = RsRow("news_pdf")
		news_subject.DbValue = RsRow("news_subject")
		news_subject_th.DbValue = RsRow("news_subject_th")
		news_intro.DbValue = RsRow("news_intro")
		news_intro_th.DbValue = RsRow("news_intro_th")
		news_content.DbValue = RsRow("news_content")
		news_content_th.DbValue = RsRow("news_content_th")
		news_show_en.DbValue = RsRow("news_show_en")
		news_show.DbValue = RsRow("news_show")
		news_show_home.DbValue = RsRow("news_show_home")
		news_create.DbValue = RsRow("news_create")
		news_update.DbValue = RsRow("news_update")
	End Sub

	' Render list row values
	Sub RenderListRow()

		'
		'  Common render codes
		'
		' news_id
		' news_img
		' news_date
		' news_category
		' news_category_sub
		' start_date
		' end_date
		' news_pdf
		' news_subject
		' news_subject_th
		' news_intro
		' news_intro_th
		' news_content
		' news_content_th
		' news_show_en
		' news_show
		' news_show_home
		' news_create
		' news_update
		' Call Row Rendering event

		Call Row_Rendering()

		'
		'  Render for View
		'
		' news_id

		news_id.ViewValue = news_id.CurrentValue
		news_id.ViewCustomAttributes = ""

		' news_img
		news_img.ViewValue = news_img.CurrentValue
		news_img.ViewCustomAttributes = ""

		' news_date
		news_date.ViewValue = news_date.CurrentValue
		news_date.ViewCustomAttributes = ""

		' news_category
		news_category.ViewValue = news_category.CurrentValue
		news_category.ViewCustomAttributes = ""

		' news_category_sub
		news_category_sub.ViewValue = news_category_sub.CurrentValue
		news_category_sub.ViewCustomAttributes = ""

		' start_date
		start_date.ViewValue = start_date.CurrentValue
		start_date.ViewCustomAttributes = ""

		' end_date
		end_date.ViewValue = end_date.CurrentValue
		end_date.ViewCustomAttributes = ""

		' news_pdf
		news_pdf.ViewValue = news_pdf.CurrentValue
		news_pdf.ViewCustomAttributes = ""

		' news_subject
		news_subject.ViewValue = news_subject.CurrentValue
		news_subject.ViewCustomAttributes = ""

		' news_subject_th
		news_subject_th.ViewValue = news_subject_th.CurrentValue
		news_subject_th.ViewCustomAttributes = ""

		' news_intro
		news_intro.ViewValue = news_intro.CurrentValue
		news_intro.ViewCustomAttributes = ""

		' news_intro_th
		news_intro_th.ViewValue = news_intro_th.CurrentValue
		news_intro_th.ViewCustomAttributes = ""

		' news_content
		news_content.ViewValue = news_content.CurrentValue
		news_content.ViewCustomAttributes = ""

		' news_content_th
		news_content_th.ViewValue = news_content_th.CurrentValue
		news_content_th.ViewCustomAttributes = ""

		' news_show_en
		news_show_en.ViewValue = news_show_en.CurrentValue
		news_show_en.ViewCustomAttributes = ""

		' news_show
		news_show.ViewValue = news_show.CurrentValue
		news_show.ViewCustomAttributes = ""

		' news_show_home
		news_show_home.ViewValue = news_show_home.CurrentValue
		news_show_home.ViewCustomAttributes = ""

		' news_create
		news_create.ViewValue = news_create.CurrentValue
		news_create.ViewCustomAttributes = ""

		' news_update
		news_update.ViewValue = news_update.CurrentValue
		news_update.ViewCustomAttributes = ""

		' news_id
		news_id.LinkCustomAttributes = ""
		news_id.HrefValue = ""
		news_id.TooltipValue = ""

		' news_img
		news_img.LinkCustomAttributes = ""
		news_img.HrefValue = ""
		news_img.TooltipValue = ""

		' news_date
		news_date.LinkCustomAttributes = ""
		news_date.HrefValue = ""
		news_date.TooltipValue = ""

		' news_category
		news_category.LinkCustomAttributes = ""
		news_category.HrefValue = ""
		news_category.TooltipValue = ""

		' news_category_sub
		news_category_sub.LinkCustomAttributes = ""
		news_category_sub.HrefValue = ""
		news_category_sub.TooltipValue = ""

		' start_date
		start_date.LinkCustomAttributes = ""
		start_date.HrefValue = ""
		start_date.TooltipValue = ""

		' end_date
		end_date.LinkCustomAttributes = ""
		end_date.HrefValue = ""
		end_date.TooltipValue = ""

		' news_pdf
		news_pdf.LinkCustomAttributes = ""
		news_pdf.HrefValue = ""
		news_pdf.TooltipValue = ""

		' news_subject
		news_subject.LinkCustomAttributes = ""
		news_subject.HrefValue = ""
		news_subject.TooltipValue = ""

		' news_subject_th
		news_subject_th.LinkCustomAttributes = ""
		news_subject_th.HrefValue = ""
		news_subject_th.TooltipValue = ""

		' news_intro
		news_intro.LinkCustomAttributes = ""
		news_intro.HrefValue = ""
		news_intro.TooltipValue = ""

		' news_intro_th
		news_intro_th.LinkCustomAttributes = ""
		news_intro_th.HrefValue = ""
		news_intro_th.TooltipValue = ""

		' news_content
		news_content.LinkCustomAttributes = ""
		news_content.HrefValue = ""
		news_content.TooltipValue = ""

		' news_content_th
		news_content_th.LinkCustomAttributes = ""
		news_content_th.HrefValue = ""
		news_content_th.TooltipValue = ""

		' news_show_en
		news_show_en.LinkCustomAttributes = ""
		news_show_en.HrefValue = ""
		news_show_en.TooltipValue = ""

		' news_show
		news_show.LinkCustomAttributes = ""
		news_show.HrefValue = ""
		news_show.TooltipValue = ""

		' news_show_home
		news_show_home.LinkCustomAttributes = ""
		news_show_home.HrefValue = ""
		news_show_home.TooltipValue = ""

		' news_create
		news_create.LinkCustomAttributes = ""
		news_create.HrefValue = ""
		news_create.TooltipValue = ""

		' news_update
		news_update.LinkCustomAttributes = ""
		news_update.HrefValue = ""
		news_update.TooltipValue = ""

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
			sSql = "SELECT * FROM [news_th] WHERE " & sWhereList
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
					Call XmlDoc.AddField("news_id", news_id.ExportValue(Export))
					Call XmlDoc.AddField("news_img", news_img.ExportValue(Export))
					Call XmlDoc.AddField("news_date", news_date.ExportValue(Export))
					Call XmlDoc.AddField("news_category", news_category.ExportValue(Export))
					Call XmlDoc.AddField("news_category_sub", news_category_sub.ExportValue(Export))
					Call XmlDoc.AddField("start_date", start_date.ExportValue(Export))
					Call XmlDoc.AddField("end_date", end_date.ExportValue(Export))
					Call XmlDoc.AddField("news_pdf", news_pdf.ExportValue(Export))
					Call XmlDoc.AddField("news_subject", news_subject.ExportValue(Export))
					Call XmlDoc.AddField("news_subject_th", news_subject_th.ExportValue(Export))
					Call XmlDoc.AddField("news_intro", news_intro.ExportValue(Export))
					Call XmlDoc.AddField("news_intro_th", news_intro_th.ExportValue(Export))
					Call XmlDoc.AddField("news_content", news_content.ExportValue(Export))
					Call XmlDoc.AddField("news_content_th", news_content_th.ExportValue(Export))
					Call XmlDoc.AddField("news_show_en", news_show_en.ExportValue(Export))
					Call XmlDoc.AddField("news_show", news_show.ExportValue(Export))
					Call XmlDoc.AddField("news_show_home", news_show_home.ExportValue(Export))
					Call XmlDoc.AddField("news_create", news_create.ExportValue(Export))
					Call XmlDoc.AddField("news_update", news_update.ExportValue(Export))
				Else
					Call XmlDoc.AddField("news_id", news_id.ExportValue(Export))
					Call XmlDoc.AddField("news_img", news_img.ExportValue(Export))
					Call XmlDoc.AddField("news_date", news_date.ExportValue(Export))
					Call XmlDoc.AddField("news_category", news_category.ExportValue(Export))
					Call XmlDoc.AddField("news_category_sub", news_category_sub.ExportValue(Export))
					Call XmlDoc.AddField("start_date", start_date.ExportValue(Export))
					Call XmlDoc.AddField("end_date", end_date.ExportValue(Export))
					Call XmlDoc.AddField("news_pdf", news_pdf.ExportValue(Export))
					Call XmlDoc.AddField("news_subject", news_subject.ExportValue(Export))
					Call XmlDoc.AddField("news_subject_th", news_subject_th.ExportValue(Export))
					Call XmlDoc.AddField("news_show_en", news_show_en.ExportValue(Export))
					Call XmlDoc.AddField("news_show", news_show.ExportValue(Export))
					Call XmlDoc.AddField("news_show_home", news_show_home.ExportValue(Export))
					Call XmlDoc.AddField("news_create", news_create.ExportValue(Export))
					Call XmlDoc.AddField("news_update", news_update.ExportValue(Export))
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
				If news_id.Exportable Then Call Doc.ExportCaption(news_id)
				If news_img.Exportable Then Call Doc.ExportCaption(news_img)
				If news_date.Exportable Then Call Doc.ExportCaption(news_date)
				If news_category.Exportable Then Call Doc.ExportCaption(news_category)
				If news_category_sub.Exportable Then Call Doc.ExportCaption(news_category_sub)
				If start_date.Exportable Then Call Doc.ExportCaption(start_date)
				If end_date.Exportable Then Call Doc.ExportCaption(end_date)
				If news_pdf.Exportable Then Call Doc.ExportCaption(news_pdf)
				If news_subject.Exportable Then Call Doc.ExportCaption(news_subject)
				If news_subject_th.Exportable Then Call Doc.ExportCaption(news_subject_th)
				If news_intro.Exportable Then Call Doc.ExportCaption(news_intro)
				If news_intro_th.Exportable Then Call Doc.ExportCaption(news_intro_th)
				If news_content.Exportable Then Call Doc.ExportCaption(news_content)
				If news_content_th.Exportable Then Call Doc.ExportCaption(news_content_th)
				If news_show_en.Exportable Then Call Doc.ExportCaption(news_show_en)
				If news_show.Exportable Then Call Doc.ExportCaption(news_show)
				If news_show_home.Exportable Then Call Doc.ExportCaption(news_show_home)
				If news_create.Exportable Then Call Doc.ExportCaption(news_create)
				If news_update.Exportable Then Call Doc.ExportCaption(news_update)
			Else
				If news_id.Exportable Then Call Doc.ExportCaption(news_id)
				If news_img.Exportable Then Call Doc.ExportCaption(news_img)
				If news_date.Exportable Then Call Doc.ExportCaption(news_date)
				If news_category.Exportable Then Call Doc.ExportCaption(news_category)
				If news_category_sub.Exportable Then Call Doc.ExportCaption(news_category_sub)
				If start_date.Exportable Then Call Doc.ExportCaption(start_date)
				If end_date.Exportable Then Call Doc.ExportCaption(end_date)
				If news_pdf.Exportable Then Call Doc.ExportCaption(news_pdf)
				If news_subject.Exportable Then Call Doc.ExportCaption(news_subject)
				If news_subject_th.Exportable Then Call Doc.ExportCaption(news_subject_th)
				If news_show_en.Exportable Then Call Doc.ExportCaption(news_show_en)
				If news_show.Exportable Then Call Doc.ExportCaption(news_show)
				If news_show_home.Exportable Then Call Doc.ExportCaption(news_show_home)
				If news_create.Exportable Then Call Doc.ExportCaption(news_create)
				If news_update.Exportable Then Call Doc.ExportCaption(news_update)
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
					If news_id.Exportable Then Call Doc.ExportField(news_id)
					If news_img.Exportable Then Call Doc.ExportField(news_img)
					If news_date.Exportable Then Call Doc.ExportField(news_date)
					If news_category.Exportable Then Call Doc.ExportField(news_category)
					If news_category_sub.Exportable Then Call Doc.ExportField(news_category_sub)
					If start_date.Exportable Then Call Doc.ExportField(start_date)
					If end_date.Exportable Then Call Doc.ExportField(end_date)
					If news_pdf.Exportable Then Call Doc.ExportField(news_pdf)
					If news_subject.Exportable Then Call Doc.ExportField(news_subject)
					If news_subject_th.Exportable Then Call Doc.ExportField(news_subject_th)
					If news_intro.Exportable Then Call Doc.ExportField(news_intro)
					If news_intro_th.Exportable Then Call Doc.ExportField(news_intro_th)
					If news_content.Exportable Then Call Doc.ExportField(news_content)
					If news_content_th.Exportable Then Call Doc.ExportField(news_content_th)
					If news_show_en.Exportable Then Call Doc.ExportField(news_show_en)
					If news_show.Exportable Then Call Doc.ExportField(news_show)
					If news_show_home.Exportable Then Call Doc.ExportField(news_show_home)
					If news_create.Exportable Then Call Doc.ExportField(news_create)
					If news_update.Exportable Then Call Doc.ExportField(news_update)
				Else
					If news_id.Exportable Then Call Doc.ExportField(news_id)
					If news_img.Exportable Then Call Doc.ExportField(news_img)
					If news_date.Exportable Then Call Doc.ExportField(news_date)
					If news_category.Exportable Then Call Doc.ExportField(news_category)
					If news_category_sub.Exportable Then Call Doc.ExportField(news_category_sub)
					If start_date.Exportable Then Call Doc.ExportField(start_date)
					If end_date.Exportable Then Call Doc.ExportField(end_date)
					If news_pdf.Exportable Then Call Doc.ExportField(news_pdf)
					If news_subject.Exportable Then Call Doc.ExportField(news_subject)
					If news_subject_th.Exportable Then Call Doc.ExportField(news_subject_th)
					If news_show_en.Exportable Then Call Doc.ExportField(news_show_en)
					If news_show.Exportable Then Call Doc.ExportField(news_show)
					If news_show_home.Exportable Then Call Doc.ExportField(news_show_home)
					If news_create.Exportable Then Call Doc.ExportField(news_create)
					If news_update.Exportable Then Call Doc.ExportField(news_update)
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
	' Field news_id
	Private m_news_id

	Public Property Get news_id()
		If Not IsObject(m_news_id) Then
			Set m_news_id = NewFldObj("news_th", "news_th", "x_news_id", "news_id", "[news_id]", "[news_id]", 3, 0, "[news_id]", False, False, FALSE, "FORMATTED TEXT")
			m_news_id.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set news_id = m_news_id
	End Property

	' Field news_img
	Private m_news_img

	Public Property Get news_img()
		If Not IsObject(m_news_img) Then
			Set m_news_img = NewFldObj("news_th", "news_th", "x_news_img", "news_img", "[news_img]", "[news_img]", 202, 0, "[news_img]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set news_img = m_news_img
	End Property

	' Field news_date
	Private m_news_date

	Public Property Get news_date()
		If Not IsObject(m_news_date) Then
			Set m_news_date = NewFldObj("news_th", "news_th", "x_news_date", "news_date", "[news_date]", "FORMAT([news_date], '')", 135, 8, "[news_date]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set news_date = m_news_date
	End Property

	' Field news_category
	Private m_news_category

	Public Property Get news_category()
		If Not IsObject(m_news_category) Then
			Set m_news_category = NewFldObj("news_th", "news_th", "x_news_category", "news_category", "[news_category]", "[news_category]", 202, 0, "[news_category]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set news_category = m_news_category
	End Property

	' Field news_category_sub
	Private m_news_category_sub

	Public Property Get news_category_sub()
		If Not IsObject(m_news_category_sub) Then
			Set m_news_category_sub = NewFldObj("news_th", "news_th", "x_news_category_sub", "news_category_sub", "[news_category_sub]", "[news_category_sub]", 202, 0, "[news_category_sub]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set news_category_sub = m_news_category_sub
	End Property

	' Field start_date
	Private m_start_date

	Public Property Get start_date()
		If Not IsObject(m_start_date) Then
			Set m_start_date = NewFldObj("news_th", "news_th", "x_start_date", "start_date", "[start_date]", "FORMAT([start_date], '')", 135, 8, "[start_date]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set start_date = m_start_date
	End Property

	' Field end_date
	Private m_end_date

	Public Property Get end_date()
		If Not IsObject(m_end_date) Then
			Set m_end_date = NewFldObj("news_th", "news_th", "x_end_date", "end_date", "[end_date]", "FORMAT([end_date], '')", 135, 8, "[end_date]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set end_date = m_end_date
	End Property

	' Field news_pdf
	Private m_news_pdf

	Public Property Get news_pdf()
		If Not IsObject(m_news_pdf) Then
			Set m_news_pdf = NewFldObj("news_th", "news_th", "x_news_pdf", "news_pdf", "[news_pdf]", "[news_pdf]", 202, 0, "[news_pdf]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set news_pdf = m_news_pdf
	End Property

	' Field news_subject
	Private m_news_subject

	Public Property Get news_subject()
		If Not IsObject(m_news_subject) Then
			Set m_news_subject = NewFldObj("news_th", "news_th", "x_news_subject", "news_subject", "[news_subject]", "[news_subject]", 202, 0, "[news_subject]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set news_subject = m_news_subject
	End Property

	' Field news_subject_th
	Private m_news_subject_th

	Public Property Get news_subject_th()
		If Not IsObject(m_news_subject_th) Then
			Set m_news_subject_th = NewFldObj("news_th", "news_th", "x_news_subject_th", "news_subject_th", "[news_subject_th]", "[news_subject_th]", 202, 0, "[news_subject_th]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set news_subject_th = m_news_subject_th
	End Property

	' Field news_intro
	Private m_news_intro

	Public Property Get news_intro()
		If Not IsObject(m_news_intro) Then
			Set m_news_intro = NewFldObj("news_th", "news_th", "x_news_intro", "news_intro", "[news_intro]", "[news_intro]", 203, 0, "[news_intro]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set news_intro = m_news_intro
	End Property

	' Field news_intro_th
	Private m_news_intro_th

	Public Property Get news_intro_th()
		If Not IsObject(m_news_intro_th) Then
			Set m_news_intro_th = NewFldObj("news_th", "news_th", "x_news_intro_th", "news_intro_th", "[news_intro_th]", "[news_intro_th]", 203, 0, "[news_intro_th]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set news_intro_th = m_news_intro_th
	End Property

	' Field news_content
	Private m_news_content

	Public Property Get news_content()
		If Not IsObject(m_news_content) Then
			Set m_news_content = NewFldObj("news_th", "news_th", "x_news_content", "news_content", "[news_content]", "[news_content]", 203, 0, "[news_content]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set news_content = m_news_content
	End Property

	' Field news_content_th
	Private m_news_content_th

	Public Property Get news_content_th()
		If Not IsObject(m_news_content_th) Then
			Set m_news_content_th = NewFldObj("news_th", "news_th", "x_news_content_th", "news_content_th", "[news_content_th]", "[news_content_th]", 203, 0, "[news_content_th]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set news_content_th = m_news_content_th
	End Property

	' Field news_show_en
	Private m_news_show_en

	Public Property Get news_show_en()
		If Not IsObject(m_news_show_en) Then
			Set m_news_show_en = NewFldObj("news_th", "news_th", "x_news_show_en", "news_show_en", "[news_show_en]", "[news_show_en]", 202, 0, "[news_show_en]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set news_show_en = m_news_show_en
	End Property

	' Field news_show
	Private m_news_show

	Public Property Get news_show()
		If Not IsObject(m_news_show) Then
			Set m_news_show = NewFldObj("news_th", "news_th", "x_news_show", "news_show", "[news_show]", "[news_show]", 202, 0, "[news_show]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set news_show = m_news_show
	End Property

	' Field news_show_home
	Private m_news_show_home

	Public Property Get news_show_home()
		If Not IsObject(m_news_show_home) Then
			Set m_news_show_home = NewFldObj("news_th", "news_th", "x_news_show_home", "news_show_home", "[news_show_home]", "[news_show_home]", 202, 0, "[news_show_home]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set news_show_home = m_news_show_home
	End Property

	' Field news_create
	Private m_news_create

	Public Property Get news_create()
		If Not IsObject(m_news_create) Then
			Set m_news_create = NewFldObj("news_th", "news_th", "x_news_create", "news_create", "[news_create]", "[news_create]", 202, 0, "[news_create]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set news_create = m_news_create
	End Property

	' Field news_update
	Private m_news_update

	Public Property Get news_update()
		If Not IsObject(m_news_update) Then
			Set m_news_update = NewFldObj("news_th", "news_th", "x_news_update", "news_update", "[news_update]", "[news_update]", 202, 0, "[news_update]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set news_update = m_news_update
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
		If IsObject(m_news_id) Then Set m_news_id = Nothing
		If IsObject(m_news_img) Then Set m_news_img = Nothing
		If IsObject(m_news_date) Then Set m_news_date = Nothing
		If IsObject(m_news_category) Then Set m_news_category = Nothing
		If IsObject(m_news_category_sub) Then Set m_news_category_sub = Nothing
		If IsObject(m_start_date) Then Set m_start_date = Nothing
		If IsObject(m_end_date) Then Set m_end_date = Nothing
		If IsObject(m_news_pdf) Then Set m_news_pdf = Nothing
		If IsObject(m_news_subject) Then Set m_news_subject = Nothing
		If IsObject(m_news_subject_th) Then Set m_news_subject_th = Nothing
		If IsObject(m_news_intro) Then Set m_news_intro = Nothing
		If IsObject(m_news_intro_th) Then Set m_news_intro_th = Nothing
		If IsObject(m_news_content) Then Set m_news_content = Nothing
		If IsObject(m_news_content_th) Then Set m_news_content_th = Nothing
		If IsObject(m_news_show_en) Then Set m_news_show_en = Nothing
		If IsObject(m_news_show) Then Set m_news_show = Nothing
		If IsObject(m_news_show_home) Then Set m_news_show_home = Nothing
		If IsObject(m_news_create) Then Set m_news_create = Nothing
		If IsObject(m_news_update) Then Set m_news_update = Nothing
		Set RowAttrs = Nothing
	End Sub
End Class
%>
