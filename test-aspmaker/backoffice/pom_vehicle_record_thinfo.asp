<%

' ASPMaker configuration for Table vehicle_record_th
Dim vehicle_record_th

' Define table class
Class cvehicle_record_th

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
		Call ew_SetArObj(Fields, "veh_id", veh_id)
		Call ew_SetArObj(Fields, "vch_month", vch_month)
		Call ew_SetArObj(Fields, "vch_year", vch_year)
		Call ew_SetArObj(Fields, "veh_product_1", veh_product_1)
		Call ew_SetArObj(Fields, "veh_product_2", veh_product_2)
		Call ew_SetArObj(Fields, "veh_product_3", veh_product_3)
		Call ew_SetArObj(Fields, "veh_product_4", veh_product_4)
		Call ew_SetArObj(Fields, "veh_product_5", veh_product_5)
		Call ew_SetArObj(Fields, "veh_product_6", veh_product_6)
		Call ew_SetArObj(Fields, "veh_product_7", veh_product_7)
		Call ew_SetArObj(Fields, "veh_product_8", veh_product_8)
		Call ew_SetArObj(Fields, "veh_domes_1", veh_domes_1)
		Call ew_SetArObj(Fields, "veh_domes_2", veh_domes_2)
		Call ew_SetArObj(Fields, "veh_domes_3", veh_domes_3)
		Call ew_SetArObj(Fields, "veh_domes_4", veh_domes_4)
		Call ew_SetArObj(Fields, "veh_domes_5", veh_domes_5)
		Call ew_SetArObj(Fields, "veh_domes_6", veh_domes_6)
		Call ew_SetArObj(Fields, "veh_domes_7", veh_domes_7)
		Call ew_SetArObj(Fields, "veh_domes_8", veh_domes_8)
		Call ew_SetArObj(Fields, "veh_export_1", veh_export_1)
		Call ew_SetArObj(Fields, "veh_export_2", veh_export_2)
		Call ew_SetArObj(Fields, "veh_export_3", veh_export_3)
		Call ew_SetArObj(Fields, "veh_export_4", veh_export_4)
		Call ew_SetArObj(Fields, "veh_export_5", veh_export_5)
		Call ew_SetArObj(Fields, "veh_export_6", veh_export_6)
		Call ew_SetArObj(Fields, "veh_export_7", veh_export_7)
		Call ew_SetArObj(Fields, "veh_export_8", veh_export_8)
		Call ew_SetArObj(Fields, "veh_remark", veh_remark)
		Call ew_SetArObj(Fields, "veh_month_title", veh_month_title)
		Call ew_SetArObj(Fields, "veh_range", veh_range)
		Call ew_SetArObj(Fields, "veh_month_title2", veh_month_title2)
		Call ew_SetArObj(Fields, "veh_range2", veh_range2)
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
		TableVar = "vehicle_record_th"
	End Property

	' Table name
	Public Property Get TableName()
		TableName = "vehicle_record_th"
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
		HighlightName = "vehicle_record_th_Highlight"
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
		SqlSelect = "SELECT * FROM [vehicle_record_th]"
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
		SqlKeyFilter = "[veh_id] = @veh_id@"
	End Property

	' Return Key filter for table
	Public Property Get KeyFilter()
		Dim sKeyFilter
		sKeyFilter = SqlKeyFilter
		If Not IsNumeric(veh_id.CurrentValue) Then
			sKeyFilter = "0=1" ' Invalid key
		End If
		sKeyFilter = Replace(sKeyFilter, "@veh_id@", ew_AdjustSql(veh_id.CurrentValue)) ' Replace key value
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
			ReturnUrl = "pom_vehicle_record_thlist.asp"
		End If
	End Property

	' List url
	Public Function ListUrl()
		ListUrl = "pom_vehicle_record_thlist.asp"
	End Function

	' View url
	Public Function ViewUrl(parm)
		If parm <> "" Then
			ViewUrl = KeyUrl("pom_vehicle_record_thview.asp", UrlParm(parm))
		Else
			ViewUrl = KeyUrl("pom_vehicle_record_thview.asp", UrlParm(EW_TABLE_SHOW_DETAIL & "="))
		End If
	End Function

	' Add url
	Public Function AddUrl()
		AddUrl = "pom_vehicle_record_thadd.asp"

'		Dim sUrlParm
'		sUrlParm = UrlParm("")
'		If sUrlParm <> "" Then AddUrl = AddUrl & "?" & sUrlParm

	End Function

	' Edit url
	Public Function EditUrl(parm)
		EditUrl = KeyUrl("pom_vehicle_record_thedit.asp", UrlParm(parm))
	End Function

	' Inline edit url
	Public Function InlineEditUrl()
		InlineEditUrl = KeyUrl(ew_CurrentPage, UrlParm("a=edit"))
	End Function

	' Copy url
	Public Function CopyUrl(parm)
		CopyUrl = KeyUrl("pom_vehicle_record_thadd.asp", UrlParm(parm))
	End Function

	' Inline copy url
	Public Function InlineCopyUrl()
		InlineCopyUrl = KeyUrl(ew_CurrentPage, UrlParm("a=copy"))
	End Function

	' Delete url
	Public Function DeleteUrl()
		DeleteUrl = KeyUrl("pom_vehicle_record_thdelete.asp", UrlParm(""))
	End Function

	' Key url
	Public Function KeyUrl(url, parm)
		Dim sUrl: sUrl = url & "?"
		If parm <> "" Then sUrl = sUrl & parm & "&"
		If Not IsNull(veh_id.CurrentValue) Then
			sUrl = sUrl & "veh_id=" & veh_id.CurrentValue
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
			UrlParm = "t=vehicle_record_th"
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
				arKeys(0) = Request.QueryString("veh_id") ' veh_id

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
				veh_id.CurrentValue = key
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
		veh_id.DbValue = RsRow("veh_id")
		vch_month.DbValue = RsRow("vch_month")
		vch_year.DbValue = RsRow("vch_year")
		veh_product_1.DbValue = RsRow("veh_product_1")
		veh_product_2.DbValue = RsRow("veh_product_2")
		veh_product_3.DbValue = RsRow("veh_product_3")
		veh_product_4.DbValue = RsRow("veh_product_4")
		veh_product_5.DbValue = RsRow("veh_product_5")
		veh_product_6.DbValue = RsRow("veh_product_6")
		veh_product_7.DbValue = RsRow("veh_product_7")
		veh_product_8.DbValue = RsRow("veh_product_8")
		veh_domes_1.DbValue = RsRow("veh_domes_1")
		veh_domes_2.DbValue = RsRow("veh_domes_2")
		veh_domes_3.DbValue = RsRow("veh_domes_3")
		veh_domes_4.DbValue = RsRow("veh_domes_4")
		veh_domes_5.DbValue = RsRow("veh_domes_5")
		veh_domes_6.DbValue = RsRow("veh_domes_6")
		veh_domes_7.DbValue = RsRow("veh_domes_7")
		veh_domes_8.DbValue = RsRow("veh_domes_8")
		veh_export_1.DbValue = RsRow("veh_export_1")
		veh_export_2.DbValue = RsRow("veh_export_2")
		veh_export_3.DbValue = RsRow("veh_export_3")
		veh_export_4.DbValue = RsRow("veh_export_4")
		veh_export_5.DbValue = RsRow("veh_export_5")
		veh_export_6.DbValue = RsRow("veh_export_6")
		veh_export_7.DbValue = RsRow("veh_export_7")
		veh_export_8.DbValue = RsRow("veh_export_8")
		veh_remark.DbValue = RsRow("veh_remark")
		veh_month_title.DbValue = RsRow("veh_month_title")
		veh_range.DbValue = RsRow("veh_range")
		veh_month_title2.DbValue = RsRow("veh_month_title2")
		veh_range2.DbValue = RsRow("veh_range2")
	End Sub

	' Render list row values
	Sub RenderListRow()

		'
		'  Common render codes
		'
		' veh_id
		' vch_month
		' vch_year
		' veh_product_1
		' veh_product_2
		' veh_product_3
		' veh_product_4
		' veh_product_5
		' veh_product_6
		' veh_product_7
		' veh_product_8
		' veh_domes_1
		' veh_domes_2
		' veh_domes_3
		' veh_domes_4
		' veh_domes_5
		' veh_domes_6
		' veh_domes_7
		' veh_domes_8
		' veh_export_1
		' veh_export_2
		' veh_export_3
		' veh_export_4
		' veh_export_5
		' veh_export_6
		' veh_export_7
		' veh_export_8
		' veh_remark
		' veh_month_title
		' veh_range
		' veh_month_title2
		' veh_range2
		' Call Row Rendering event

		Call Row_Rendering()

		'
		'  Render for View
		'
		' veh_id

		veh_id.ViewValue = veh_id.CurrentValue
		veh_id.ViewCustomAttributes = ""

		' vch_month
		vch_month.ViewValue = vch_month.CurrentValue
		vch_month.ViewCustomAttributes = ""

		' vch_year
		vch_year.ViewValue = vch_year.CurrentValue
		vch_year.ViewCustomAttributes = ""

		' veh_product_1
		veh_product_1.ViewValue = veh_product_1.CurrentValue
		veh_product_1.ViewCustomAttributes = ""

		' veh_product_2
		veh_product_2.ViewValue = veh_product_2.CurrentValue
		veh_product_2.ViewCustomAttributes = ""

		' veh_product_3
		veh_product_3.ViewValue = veh_product_3.CurrentValue
		veh_product_3.ViewCustomAttributes = ""

		' veh_product_4
		veh_product_4.ViewValue = veh_product_4.CurrentValue
		veh_product_4.ViewCustomAttributes = ""

		' veh_product_5
		veh_product_5.ViewValue = veh_product_5.CurrentValue
		veh_product_5.ViewCustomAttributes = ""

		' veh_product_6
		veh_product_6.ViewValue = veh_product_6.CurrentValue
		veh_product_6.ViewCustomAttributes = ""

		' veh_product_7
		veh_product_7.ViewValue = veh_product_7.CurrentValue
		veh_product_7.ViewCustomAttributes = ""

		' veh_product_8
		veh_product_8.ViewValue = veh_product_8.CurrentValue
		veh_product_8.ViewCustomAttributes = ""

		' veh_domes_1
		veh_domes_1.ViewValue = veh_domes_1.CurrentValue
		veh_domes_1.ViewCustomAttributes = ""

		' veh_domes_2
		veh_domes_2.ViewValue = veh_domes_2.CurrentValue
		veh_domes_2.ViewCustomAttributes = ""

		' veh_domes_3
		veh_domes_3.ViewValue = veh_domes_3.CurrentValue
		veh_domes_3.ViewCustomAttributes = ""

		' veh_domes_4
		veh_domes_4.ViewValue = veh_domes_4.CurrentValue
		veh_domes_4.ViewCustomAttributes = ""

		' veh_domes_5
		veh_domes_5.ViewValue = veh_domes_5.CurrentValue
		veh_domes_5.ViewCustomAttributes = ""

		' veh_domes_6
		veh_domes_6.ViewValue = veh_domes_6.CurrentValue
		veh_domes_6.ViewCustomAttributes = ""

		' veh_domes_7
		veh_domes_7.ViewValue = veh_domes_7.CurrentValue
		veh_domes_7.ViewCustomAttributes = ""

		' veh_domes_8
		veh_domes_8.ViewValue = veh_domes_8.CurrentValue
		veh_domes_8.ViewCustomAttributes = ""

		' veh_export_1
		veh_export_1.ViewValue = veh_export_1.CurrentValue
		veh_export_1.ViewCustomAttributes = ""

		' veh_export_2
		veh_export_2.ViewValue = veh_export_2.CurrentValue
		veh_export_2.ViewCustomAttributes = ""

		' veh_export_3
		veh_export_3.ViewValue = veh_export_3.CurrentValue
		veh_export_3.ViewCustomAttributes = ""

		' veh_export_4
		veh_export_4.ViewValue = veh_export_4.CurrentValue
		veh_export_4.ViewCustomAttributes = ""

		' veh_export_5
		veh_export_5.ViewValue = veh_export_5.CurrentValue
		veh_export_5.ViewCustomAttributes = ""

		' veh_export_6
		veh_export_6.ViewValue = veh_export_6.CurrentValue
		veh_export_6.ViewCustomAttributes = ""

		' veh_export_7
		veh_export_7.ViewValue = veh_export_7.CurrentValue
		veh_export_7.ViewCustomAttributes = ""

		' veh_export_8
		veh_export_8.ViewValue = veh_export_8.CurrentValue
		veh_export_8.ViewCustomAttributes = ""

		' veh_remark
		veh_remark.ViewValue = veh_remark.CurrentValue
		veh_remark.ViewCustomAttributes = ""

		' veh_month_title
		veh_month_title.ViewValue = veh_month_title.CurrentValue
		veh_month_title.ViewCustomAttributes = ""

		' veh_range
		veh_range.ViewValue = veh_range.CurrentValue
		veh_range.ViewCustomAttributes = ""

		' veh_month_title2
		veh_month_title2.ViewValue = veh_month_title2.CurrentValue
		veh_month_title2.ViewCustomAttributes = ""

		' veh_range2
		veh_range2.ViewValue = veh_range2.CurrentValue
		veh_range2.ViewCustomAttributes = ""

		' veh_id
		veh_id.LinkCustomAttributes = ""
		veh_id.HrefValue = ""
		veh_id.TooltipValue = ""

		' vch_month
		vch_month.LinkCustomAttributes = ""
		vch_month.HrefValue = ""
		vch_month.TooltipValue = ""

		' vch_year
		vch_year.LinkCustomAttributes = ""
		vch_year.HrefValue = ""
		vch_year.TooltipValue = ""

		' veh_product_1
		veh_product_1.LinkCustomAttributes = ""
		veh_product_1.HrefValue = ""
		veh_product_1.TooltipValue = ""

		' veh_product_2
		veh_product_2.LinkCustomAttributes = ""
		veh_product_2.HrefValue = ""
		veh_product_2.TooltipValue = ""

		' veh_product_3
		veh_product_3.LinkCustomAttributes = ""
		veh_product_3.HrefValue = ""
		veh_product_3.TooltipValue = ""

		' veh_product_4
		veh_product_4.LinkCustomAttributes = ""
		veh_product_4.HrefValue = ""
		veh_product_4.TooltipValue = ""

		' veh_product_5
		veh_product_5.LinkCustomAttributes = ""
		veh_product_5.HrefValue = ""
		veh_product_5.TooltipValue = ""

		' veh_product_6
		veh_product_6.LinkCustomAttributes = ""
		veh_product_6.HrefValue = ""
		veh_product_6.TooltipValue = ""

		' veh_product_7
		veh_product_7.LinkCustomAttributes = ""
		veh_product_7.HrefValue = ""
		veh_product_7.TooltipValue = ""

		' veh_product_8
		veh_product_8.LinkCustomAttributes = ""
		veh_product_8.HrefValue = ""
		veh_product_8.TooltipValue = ""

		' veh_domes_1
		veh_domes_1.LinkCustomAttributes = ""
		veh_domes_1.HrefValue = ""
		veh_domes_1.TooltipValue = ""

		' veh_domes_2
		veh_domes_2.LinkCustomAttributes = ""
		veh_domes_2.HrefValue = ""
		veh_domes_2.TooltipValue = ""

		' veh_domes_3
		veh_domes_3.LinkCustomAttributes = ""
		veh_domes_3.HrefValue = ""
		veh_domes_3.TooltipValue = ""

		' veh_domes_4
		veh_domes_4.LinkCustomAttributes = ""
		veh_domes_4.HrefValue = ""
		veh_domes_4.TooltipValue = ""

		' veh_domes_5
		veh_domes_5.LinkCustomAttributes = ""
		veh_domes_5.HrefValue = ""
		veh_domes_5.TooltipValue = ""

		' veh_domes_6
		veh_domes_6.LinkCustomAttributes = ""
		veh_domes_6.HrefValue = ""
		veh_domes_6.TooltipValue = ""

		' veh_domes_7
		veh_domes_7.LinkCustomAttributes = ""
		veh_domes_7.HrefValue = ""
		veh_domes_7.TooltipValue = ""

		' veh_domes_8
		veh_domes_8.LinkCustomAttributes = ""
		veh_domes_8.HrefValue = ""
		veh_domes_8.TooltipValue = ""

		' veh_export_1
		veh_export_1.LinkCustomAttributes = ""
		veh_export_1.HrefValue = ""
		veh_export_1.TooltipValue = ""

		' veh_export_2
		veh_export_2.LinkCustomAttributes = ""
		veh_export_2.HrefValue = ""
		veh_export_2.TooltipValue = ""

		' veh_export_3
		veh_export_3.LinkCustomAttributes = ""
		veh_export_3.HrefValue = ""
		veh_export_3.TooltipValue = ""

		' veh_export_4
		veh_export_4.LinkCustomAttributes = ""
		veh_export_4.HrefValue = ""
		veh_export_4.TooltipValue = ""

		' veh_export_5
		veh_export_5.LinkCustomAttributes = ""
		veh_export_5.HrefValue = ""
		veh_export_5.TooltipValue = ""

		' veh_export_6
		veh_export_6.LinkCustomAttributes = ""
		veh_export_6.HrefValue = ""
		veh_export_6.TooltipValue = ""

		' veh_export_7
		veh_export_7.LinkCustomAttributes = ""
		veh_export_7.HrefValue = ""
		veh_export_7.TooltipValue = ""

		' veh_export_8
		veh_export_8.LinkCustomAttributes = ""
		veh_export_8.HrefValue = ""
		veh_export_8.TooltipValue = ""

		' veh_remark
		veh_remark.LinkCustomAttributes = ""
		veh_remark.HrefValue = ""
		veh_remark.TooltipValue = ""

		' veh_month_title
		veh_month_title.LinkCustomAttributes = ""
		veh_month_title.HrefValue = ""
		veh_month_title.TooltipValue = ""

		' veh_range
		veh_range.LinkCustomAttributes = ""
		veh_range.HrefValue = ""
		veh_range.TooltipValue = ""

		' veh_month_title2
		veh_month_title2.LinkCustomAttributes = ""
		veh_month_title2.HrefValue = ""
		veh_month_title2.TooltipValue = ""

		' veh_range2
		veh_range2.LinkCustomAttributes = ""
		veh_range2.HrefValue = ""
		veh_range2.TooltipValue = ""

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
			sSql = "SELECT * FROM [vehicle_record_th] WHERE " & sWhereList
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
					Call XmlDoc.AddField("veh_id", veh_id.ExportValue(Export))
					Call XmlDoc.AddField("vch_month", vch_month.ExportValue(Export))
					Call XmlDoc.AddField("vch_year", vch_year.ExportValue(Export))
					Call XmlDoc.AddField("veh_product_1", veh_product_1.ExportValue(Export))
					Call XmlDoc.AddField("veh_product_2", veh_product_2.ExportValue(Export))
					Call XmlDoc.AddField("veh_product_3", veh_product_3.ExportValue(Export))
					Call XmlDoc.AddField("veh_product_4", veh_product_4.ExportValue(Export))
					Call XmlDoc.AddField("veh_product_5", veh_product_5.ExportValue(Export))
					Call XmlDoc.AddField("veh_product_6", veh_product_6.ExportValue(Export))
					Call XmlDoc.AddField("veh_product_7", veh_product_7.ExportValue(Export))
					Call XmlDoc.AddField("veh_product_8", veh_product_8.ExportValue(Export))
					Call XmlDoc.AddField("veh_domes_1", veh_domes_1.ExportValue(Export))
					Call XmlDoc.AddField("veh_domes_2", veh_domes_2.ExportValue(Export))
					Call XmlDoc.AddField("veh_domes_3", veh_domes_3.ExportValue(Export))
					Call XmlDoc.AddField("veh_domes_4", veh_domes_4.ExportValue(Export))
					Call XmlDoc.AddField("veh_domes_5", veh_domes_5.ExportValue(Export))
					Call XmlDoc.AddField("veh_domes_6", veh_domes_6.ExportValue(Export))
					Call XmlDoc.AddField("veh_domes_7", veh_domes_7.ExportValue(Export))
					Call XmlDoc.AddField("veh_domes_8", veh_domes_8.ExportValue(Export))
					Call XmlDoc.AddField("veh_export_1", veh_export_1.ExportValue(Export))
					Call XmlDoc.AddField("veh_export_2", veh_export_2.ExportValue(Export))
					Call XmlDoc.AddField("veh_export_3", veh_export_3.ExportValue(Export))
					Call XmlDoc.AddField("veh_export_4", veh_export_4.ExportValue(Export))
					Call XmlDoc.AddField("veh_export_5", veh_export_5.ExportValue(Export))
					Call XmlDoc.AddField("veh_export_6", veh_export_6.ExportValue(Export))
					Call XmlDoc.AddField("veh_export_7", veh_export_7.ExportValue(Export))
					Call XmlDoc.AddField("veh_export_8", veh_export_8.ExportValue(Export))
					Call XmlDoc.AddField("veh_remark", veh_remark.ExportValue(Export))
					Call XmlDoc.AddField("veh_month_title", veh_month_title.ExportValue(Export))
					Call XmlDoc.AddField("veh_range", veh_range.ExportValue(Export))
					Call XmlDoc.AddField("veh_month_title2", veh_month_title2.ExportValue(Export))
					Call XmlDoc.AddField("veh_range2", veh_range2.ExportValue(Export))
				Else
					Call XmlDoc.AddField("veh_id", veh_id.ExportValue(Export))
					Call XmlDoc.AddField("vch_month", vch_month.ExportValue(Export))
					Call XmlDoc.AddField("vch_year", vch_year.ExportValue(Export))
					Call XmlDoc.AddField("veh_product_1", veh_product_1.ExportValue(Export))
					Call XmlDoc.AddField("veh_product_2", veh_product_2.ExportValue(Export))
					Call XmlDoc.AddField("veh_product_3", veh_product_3.ExportValue(Export))
					Call XmlDoc.AddField("veh_product_4", veh_product_4.ExportValue(Export))
					Call XmlDoc.AddField("veh_product_5", veh_product_5.ExportValue(Export))
					Call XmlDoc.AddField("veh_product_6", veh_product_6.ExportValue(Export))
					Call XmlDoc.AddField("veh_product_7", veh_product_7.ExportValue(Export))
					Call XmlDoc.AddField("veh_product_8", veh_product_8.ExportValue(Export))
					Call XmlDoc.AddField("veh_domes_1", veh_domes_1.ExportValue(Export))
					Call XmlDoc.AddField("veh_domes_2", veh_domes_2.ExportValue(Export))
					Call XmlDoc.AddField("veh_domes_3", veh_domes_3.ExportValue(Export))
					Call XmlDoc.AddField("veh_domes_4", veh_domes_4.ExportValue(Export))
					Call XmlDoc.AddField("veh_domes_5", veh_domes_5.ExportValue(Export))
					Call XmlDoc.AddField("veh_domes_6", veh_domes_6.ExportValue(Export))
					Call XmlDoc.AddField("veh_domes_7", veh_domes_7.ExportValue(Export))
					Call XmlDoc.AddField("veh_domes_8", veh_domes_8.ExportValue(Export))
					Call XmlDoc.AddField("veh_export_1", veh_export_1.ExportValue(Export))
					Call XmlDoc.AddField("veh_export_2", veh_export_2.ExportValue(Export))
					Call XmlDoc.AddField("veh_export_3", veh_export_3.ExportValue(Export))
					Call XmlDoc.AddField("veh_export_4", veh_export_4.ExportValue(Export))
					Call XmlDoc.AddField("veh_export_5", veh_export_5.ExportValue(Export))
					Call XmlDoc.AddField("veh_export_6", veh_export_6.ExportValue(Export))
					Call XmlDoc.AddField("veh_export_7", veh_export_7.ExportValue(Export))
					Call XmlDoc.AddField("veh_export_8", veh_export_8.ExportValue(Export))
					Call XmlDoc.AddField("veh_month_title", veh_month_title.ExportValue(Export))
					Call XmlDoc.AddField("veh_range", veh_range.ExportValue(Export))
					Call XmlDoc.AddField("veh_month_title2", veh_month_title2.ExportValue(Export))
					Call XmlDoc.AddField("veh_range2", veh_range2.ExportValue(Export))
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
				If veh_id.Exportable Then Call Doc.ExportCaption(veh_id)
				If vch_month.Exportable Then Call Doc.ExportCaption(vch_month)
				If vch_year.Exportable Then Call Doc.ExportCaption(vch_year)
				If veh_product_1.Exportable Then Call Doc.ExportCaption(veh_product_1)
				If veh_product_2.Exportable Then Call Doc.ExportCaption(veh_product_2)
				If veh_product_3.Exportable Then Call Doc.ExportCaption(veh_product_3)
				If veh_product_4.Exportable Then Call Doc.ExportCaption(veh_product_4)
				If veh_product_5.Exportable Then Call Doc.ExportCaption(veh_product_5)
				If veh_product_6.Exportable Then Call Doc.ExportCaption(veh_product_6)
				If veh_product_7.Exportable Then Call Doc.ExportCaption(veh_product_7)
				If veh_product_8.Exportable Then Call Doc.ExportCaption(veh_product_8)
				If veh_domes_1.Exportable Then Call Doc.ExportCaption(veh_domes_1)
				If veh_domes_2.Exportable Then Call Doc.ExportCaption(veh_domes_2)
				If veh_domes_3.Exportable Then Call Doc.ExportCaption(veh_domes_3)
				If veh_domes_4.Exportable Then Call Doc.ExportCaption(veh_domes_4)
				If veh_domes_5.Exportable Then Call Doc.ExportCaption(veh_domes_5)
				If veh_domes_6.Exportable Then Call Doc.ExportCaption(veh_domes_6)
				If veh_domes_7.Exportable Then Call Doc.ExportCaption(veh_domes_7)
				If veh_domes_8.Exportable Then Call Doc.ExportCaption(veh_domes_8)
				If veh_export_1.Exportable Then Call Doc.ExportCaption(veh_export_1)
				If veh_export_2.Exportable Then Call Doc.ExportCaption(veh_export_2)
				If veh_export_3.Exportable Then Call Doc.ExportCaption(veh_export_3)
				If veh_export_4.Exportable Then Call Doc.ExportCaption(veh_export_4)
				If veh_export_5.Exportable Then Call Doc.ExportCaption(veh_export_5)
				If veh_export_6.Exportable Then Call Doc.ExportCaption(veh_export_6)
				If veh_export_7.Exportable Then Call Doc.ExportCaption(veh_export_7)
				If veh_export_8.Exportable Then Call Doc.ExportCaption(veh_export_8)
				If veh_remark.Exportable Then Call Doc.ExportCaption(veh_remark)
				If veh_month_title.Exportable Then Call Doc.ExportCaption(veh_month_title)
				If veh_range.Exportable Then Call Doc.ExportCaption(veh_range)
				If veh_month_title2.Exportable Then Call Doc.ExportCaption(veh_month_title2)
				If veh_range2.Exportable Then Call Doc.ExportCaption(veh_range2)
			Else
				If veh_id.Exportable Then Call Doc.ExportCaption(veh_id)
				If vch_month.Exportable Then Call Doc.ExportCaption(vch_month)
				If vch_year.Exportable Then Call Doc.ExportCaption(vch_year)
				If veh_product_1.Exportable Then Call Doc.ExportCaption(veh_product_1)
				If veh_product_2.Exportable Then Call Doc.ExportCaption(veh_product_2)
				If veh_product_3.Exportable Then Call Doc.ExportCaption(veh_product_3)
				If veh_product_4.Exportable Then Call Doc.ExportCaption(veh_product_4)
				If veh_product_5.Exportable Then Call Doc.ExportCaption(veh_product_5)
				If veh_product_6.Exportable Then Call Doc.ExportCaption(veh_product_6)
				If veh_product_7.Exportable Then Call Doc.ExportCaption(veh_product_7)
				If veh_product_8.Exportable Then Call Doc.ExportCaption(veh_product_8)
				If veh_domes_1.Exportable Then Call Doc.ExportCaption(veh_domes_1)
				If veh_domes_2.Exportable Then Call Doc.ExportCaption(veh_domes_2)
				If veh_domes_3.Exportable Then Call Doc.ExportCaption(veh_domes_3)
				If veh_domes_4.Exportable Then Call Doc.ExportCaption(veh_domes_4)
				If veh_domes_5.Exportable Then Call Doc.ExportCaption(veh_domes_5)
				If veh_domes_6.Exportable Then Call Doc.ExportCaption(veh_domes_6)
				If veh_domes_7.Exportable Then Call Doc.ExportCaption(veh_domes_7)
				If veh_domes_8.Exportable Then Call Doc.ExportCaption(veh_domes_8)
				If veh_export_1.Exportable Then Call Doc.ExportCaption(veh_export_1)
				If veh_export_2.Exportable Then Call Doc.ExportCaption(veh_export_2)
				If veh_export_3.Exportable Then Call Doc.ExportCaption(veh_export_3)
				If veh_export_4.Exportable Then Call Doc.ExportCaption(veh_export_4)
				If veh_export_5.Exportable Then Call Doc.ExportCaption(veh_export_5)
				If veh_export_6.Exportable Then Call Doc.ExportCaption(veh_export_6)
				If veh_export_7.Exportable Then Call Doc.ExportCaption(veh_export_7)
				If veh_export_8.Exportable Then Call Doc.ExportCaption(veh_export_8)
				If veh_month_title.Exportable Then Call Doc.ExportCaption(veh_month_title)
				If veh_range.Exportable Then Call Doc.ExportCaption(veh_range)
				If veh_month_title2.Exportable Then Call Doc.ExportCaption(veh_month_title2)
				If veh_range2.Exportable Then Call Doc.ExportCaption(veh_range2)
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
					If veh_id.Exportable Then Call Doc.ExportField(veh_id)
					If vch_month.Exportable Then Call Doc.ExportField(vch_month)
					If vch_year.Exportable Then Call Doc.ExportField(vch_year)
					If veh_product_1.Exportable Then Call Doc.ExportField(veh_product_1)
					If veh_product_2.Exportable Then Call Doc.ExportField(veh_product_2)
					If veh_product_3.Exportable Then Call Doc.ExportField(veh_product_3)
					If veh_product_4.Exportable Then Call Doc.ExportField(veh_product_4)
					If veh_product_5.Exportable Then Call Doc.ExportField(veh_product_5)
					If veh_product_6.Exportable Then Call Doc.ExportField(veh_product_6)
					If veh_product_7.Exportable Then Call Doc.ExportField(veh_product_7)
					If veh_product_8.Exportable Then Call Doc.ExportField(veh_product_8)
					If veh_domes_1.Exportable Then Call Doc.ExportField(veh_domes_1)
					If veh_domes_2.Exportable Then Call Doc.ExportField(veh_domes_2)
					If veh_domes_3.Exportable Then Call Doc.ExportField(veh_domes_3)
					If veh_domes_4.Exportable Then Call Doc.ExportField(veh_domes_4)
					If veh_domes_5.Exportable Then Call Doc.ExportField(veh_domes_5)
					If veh_domes_6.Exportable Then Call Doc.ExportField(veh_domes_6)
					If veh_domes_7.Exportable Then Call Doc.ExportField(veh_domes_7)
					If veh_domes_8.Exportable Then Call Doc.ExportField(veh_domes_8)
					If veh_export_1.Exportable Then Call Doc.ExportField(veh_export_1)
					If veh_export_2.Exportable Then Call Doc.ExportField(veh_export_2)
					If veh_export_3.Exportable Then Call Doc.ExportField(veh_export_3)
					If veh_export_4.Exportable Then Call Doc.ExportField(veh_export_4)
					If veh_export_5.Exportable Then Call Doc.ExportField(veh_export_5)
					If veh_export_6.Exportable Then Call Doc.ExportField(veh_export_6)
					If veh_export_7.Exportable Then Call Doc.ExportField(veh_export_7)
					If veh_export_8.Exportable Then Call Doc.ExportField(veh_export_8)
					If veh_remark.Exportable Then Call Doc.ExportField(veh_remark)
					If veh_month_title.Exportable Then Call Doc.ExportField(veh_month_title)
					If veh_range.Exportable Then Call Doc.ExportField(veh_range)
					If veh_month_title2.Exportable Then Call Doc.ExportField(veh_month_title2)
					If veh_range2.Exportable Then Call Doc.ExportField(veh_range2)
				Else
					If veh_id.Exportable Then Call Doc.ExportField(veh_id)
					If vch_month.Exportable Then Call Doc.ExportField(vch_month)
					If vch_year.Exportable Then Call Doc.ExportField(vch_year)
					If veh_product_1.Exportable Then Call Doc.ExportField(veh_product_1)
					If veh_product_2.Exportable Then Call Doc.ExportField(veh_product_2)
					If veh_product_3.Exportable Then Call Doc.ExportField(veh_product_3)
					If veh_product_4.Exportable Then Call Doc.ExportField(veh_product_4)
					If veh_product_5.Exportable Then Call Doc.ExportField(veh_product_5)
					If veh_product_6.Exportable Then Call Doc.ExportField(veh_product_6)
					If veh_product_7.Exportable Then Call Doc.ExportField(veh_product_7)
					If veh_product_8.Exportable Then Call Doc.ExportField(veh_product_8)
					If veh_domes_1.Exportable Then Call Doc.ExportField(veh_domes_1)
					If veh_domes_2.Exportable Then Call Doc.ExportField(veh_domes_2)
					If veh_domes_3.Exportable Then Call Doc.ExportField(veh_domes_3)
					If veh_domes_4.Exportable Then Call Doc.ExportField(veh_domes_4)
					If veh_domes_5.Exportable Then Call Doc.ExportField(veh_domes_5)
					If veh_domes_6.Exportable Then Call Doc.ExportField(veh_domes_6)
					If veh_domes_7.Exportable Then Call Doc.ExportField(veh_domes_7)
					If veh_domes_8.Exportable Then Call Doc.ExportField(veh_domes_8)
					If veh_export_1.Exportable Then Call Doc.ExportField(veh_export_1)
					If veh_export_2.Exportable Then Call Doc.ExportField(veh_export_2)
					If veh_export_3.Exportable Then Call Doc.ExportField(veh_export_3)
					If veh_export_4.Exportable Then Call Doc.ExportField(veh_export_4)
					If veh_export_5.Exportable Then Call Doc.ExportField(veh_export_5)
					If veh_export_6.Exportable Then Call Doc.ExportField(veh_export_6)
					If veh_export_7.Exportable Then Call Doc.ExportField(veh_export_7)
					If veh_export_8.Exportable Then Call Doc.ExportField(veh_export_8)
					If veh_month_title.Exportable Then Call Doc.ExportField(veh_month_title)
					If veh_range.Exportable Then Call Doc.ExportField(veh_range)
					If veh_month_title2.Exportable Then Call Doc.ExportField(veh_month_title2)
					If veh_range2.Exportable Then Call Doc.ExportField(veh_range2)
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
	' Field veh_id
	Private m_veh_id

	Public Property Get veh_id()
		If Not IsObject(m_veh_id) Then
			Set m_veh_id = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_id", "veh_id", "[veh_id]", "[veh_id]", 3, 0, "[veh_id]", False, False, FALSE, "FORMATTED TEXT")
			m_veh_id.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set veh_id = m_veh_id
	End Property

	' Field vch_month
	Private m_vch_month

	Public Property Get vch_month()
		If Not IsObject(m_vch_month) Then
			Set m_vch_month = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_vch_month", "vch_month", "[vch_month]", "[vch_month]", 202, 0, "[vch_month]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set vch_month = m_vch_month
	End Property

	' Field vch_year
	Private m_vch_year

	Public Property Get vch_year()
		If Not IsObject(m_vch_year) Then
			Set m_vch_year = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_vch_year", "vch_year", "[vch_year]", "[vch_year]", 202, 0, "[vch_year]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set vch_year = m_vch_year
	End Property

	' Field veh_product_1
	Private m_veh_product_1

	Public Property Get veh_product_1()
		If Not IsObject(m_veh_product_1) Then
			Set m_veh_product_1 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_product_1", "veh_product_1", "[veh_product_1]", "[veh_product_1]", 202, 0, "[veh_product_1]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_product_1 = m_veh_product_1
	End Property

	' Field veh_product_2
	Private m_veh_product_2

	Public Property Get veh_product_2()
		If Not IsObject(m_veh_product_2) Then
			Set m_veh_product_2 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_product_2", "veh_product_2", "[veh_product_2]", "[veh_product_2]", 202, 0, "[veh_product_2]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_product_2 = m_veh_product_2
	End Property

	' Field veh_product_3
	Private m_veh_product_3

	Public Property Get veh_product_3()
		If Not IsObject(m_veh_product_3) Then
			Set m_veh_product_3 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_product_3", "veh_product_3", "[veh_product_3]", "[veh_product_3]", 202, 0, "[veh_product_3]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_product_3 = m_veh_product_3
	End Property

	' Field veh_product_4
	Private m_veh_product_4

	Public Property Get veh_product_4()
		If Not IsObject(m_veh_product_4) Then
			Set m_veh_product_4 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_product_4", "veh_product_4", "[veh_product_4]", "[veh_product_4]", 202, 0, "[veh_product_4]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_product_4 = m_veh_product_4
	End Property

	' Field veh_product_5
	Private m_veh_product_5

	Public Property Get veh_product_5()
		If Not IsObject(m_veh_product_5) Then
			Set m_veh_product_5 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_product_5", "veh_product_5", "[veh_product_5]", "[veh_product_5]", 202, 0, "[veh_product_5]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_product_5 = m_veh_product_5
	End Property

	' Field veh_product_6
	Private m_veh_product_6

	Public Property Get veh_product_6()
		If Not IsObject(m_veh_product_6) Then
			Set m_veh_product_6 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_product_6", "veh_product_6", "[veh_product_6]", "[veh_product_6]", 202, 0, "[veh_product_6]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_product_6 = m_veh_product_6
	End Property

	' Field veh_product_7
	Private m_veh_product_7

	Public Property Get veh_product_7()
		If Not IsObject(m_veh_product_7) Then
			Set m_veh_product_7 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_product_7", "veh_product_7", "[veh_product_7]", "[veh_product_7]", 202, 0, "[veh_product_7]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_product_7 = m_veh_product_7
	End Property

	' Field veh_product_8
	Private m_veh_product_8

	Public Property Get veh_product_8()
		If Not IsObject(m_veh_product_8) Then
			Set m_veh_product_8 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_product_8", "veh_product_8", "[veh_product_8]", "[veh_product_8]", 202, 0, "[veh_product_8]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_product_8 = m_veh_product_8
	End Property

	' Field veh_domes_1
	Private m_veh_domes_1

	Public Property Get veh_domes_1()
		If Not IsObject(m_veh_domes_1) Then
			Set m_veh_domes_1 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_domes_1", "veh_domes_1", "[veh_domes_1]", "[veh_domes_1]", 202, 0, "[veh_domes_1]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_domes_1 = m_veh_domes_1
	End Property

	' Field veh_domes_2
	Private m_veh_domes_2

	Public Property Get veh_domes_2()
		If Not IsObject(m_veh_domes_2) Then
			Set m_veh_domes_2 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_domes_2", "veh_domes_2", "[veh_domes_2]", "[veh_domes_2]", 202, 0, "[veh_domes_2]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_domes_2 = m_veh_domes_2
	End Property

	' Field veh_domes_3
	Private m_veh_domes_3

	Public Property Get veh_domes_3()
		If Not IsObject(m_veh_domes_3) Then
			Set m_veh_domes_3 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_domes_3", "veh_domes_3", "[veh_domes_3]", "[veh_domes_3]", 202, 0, "[veh_domes_3]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_domes_3 = m_veh_domes_3
	End Property

	' Field veh_domes_4
	Private m_veh_domes_4

	Public Property Get veh_domes_4()
		If Not IsObject(m_veh_domes_4) Then
			Set m_veh_domes_4 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_domes_4", "veh_domes_4", "[veh_domes_4]", "[veh_domes_4]", 202, 0, "[veh_domes_4]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_domes_4 = m_veh_domes_4
	End Property

	' Field veh_domes_5
	Private m_veh_domes_5

	Public Property Get veh_domes_5()
		If Not IsObject(m_veh_domes_5) Then
			Set m_veh_domes_5 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_domes_5", "veh_domes_5", "[veh_domes_5]", "[veh_domes_5]", 202, 0, "[veh_domes_5]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_domes_5 = m_veh_domes_5
	End Property

	' Field veh_domes_6
	Private m_veh_domes_6

	Public Property Get veh_domes_6()
		If Not IsObject(m_veh_domes_6) Then
			Set m_veh_domes_6 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_domes_6", "veh_domes_6", "[veh_domes_6]", "[veh_domes_6]", 202, 0, "[veh_domes_6]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_domes_6 = m_veh_domes_6
	End Property

	' Field veh_domes_7
	Private m_veh_domes_7

	Public Property Get veh_domes_7()
		If Not IsObject(m_veh_domes_7) Then
			Set m_veh_domes_7 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_domes_7", "veh_domes_7", "[veh_domes_7]", "[veh_domes_7]", 202, 0, "[veh_domes_7]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_domes_7 = m_veh_domes_7
	End Property

	' Field veh_domes_8
	Private m_veh_domes_8

	Public Property Get veh_domes_8()
		If Not IsObject(m_veh_domes_8) Then
			Set m_veh_domes_8 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_domes_8", "veh_domes_8", "[veh_domes_8]", "[veh_domes_8]", 202, 0, "[veh_domes_8]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_domes_8 = m_veh_domes_8
	End Property

	' Field veh_export_1
	Private m_veh_export_1

	Public Property Get veh_export_1()
		If Not IsObject(m_veh_export_1) Then
			Set m_veh_export_1 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_export_1", "veh_export_1", "[veh_export_1]", "[veh_export_1]", 202, 0, "[veh_export_1]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_export_1 = m_veh_export_1
	End Property

	' Field veh_export_2
	Private m_veh_export_2

	Public Property Get veh_export_2()
		If Not IsObject(m_veh_export_2) Then
			Set m_veh_export_2 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_export_2", "veh_export_2", "[veh_export_2]", "[veh_export_2]", 202, 0, "[veh_export_2]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_export_2 = m_veh_export_2
	End Property

	' Field veh_export_3
	Private m_veh_export_3

	Public Property Get veh_export_3()
		If Not IsObject(m_veh_export_3) Then
			Set m_veh_export_3 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_export_3", "veh_export_3", "[veh_export_3]", "[veh_export_3]", 202, 0, "[veh_export_3]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_export_3 = m_veh_export_3
	End Property

	' Field veh_export_4
	Private m_veh_export_4

	Public Property Get veh_export_4()
		If Not IsObject(m_veh_export_4) Then
			Set m_veh_export_4 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_export_4", "veh_export_4", "[veh_export_4]", "[veh_export_4]", 202, 0, "[veh_export_4]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_export_4 = m_veh_export_4
	End Property

	' Field veh_export_5
	Private m_veh_export_5

	Public Property Get veh_export_5()
		If Not IsObject(m_veh_export_5) Then
			Set m_veh_export_5 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_export_5", "veh_export_5", "[veh_export_5]", "[veh_export_5]", 202, 0, "[veh_export_5]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_export_5 = m_veh_export_5
	End Property

	' Field veh_export_6
	Private m_veh_export_6

	Public Property Get veh_export_6()
		If Not IsObject(m_veh_export_6) Then
			Set m_veh_export_6 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_export_6", "veh_export_6", "[veh_export_6]", "[veh_export_6]", 202, 0, "[veh_export_6]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_export_6 = m_veh_export_6
	End Property

	' Field veh_export_7
	Private m_veh_export_7

	Public Property Get veh_export_7()
		If Not IsObject(m_veh_export_7) Then
			Set m_veh_export_7 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_export_7", "veh_export_7", "[veh_export_7]", "[veh_export_7]", 202, 0, "[veh_export_7]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_export_7 = m_veh_export_7
	End Property

	' Field veh_export_8
	Private m_veh_export_8

	Public Property Get veh_export_8()
		If Not IsObject(m_veh_export_8) Then
			Set m_veh_export_8 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_export_8", "veh_export_8", "[veh_export_8]", "[veh_export_8]", 202, 0, "[veh_export_8]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_export_8 = m_veh_export_8
	End Property

	' Field veh_remark
	Private m_veh_remark

	Public Property Get veh_remark()
		If Not IsObject(m_veh_remark) Then
			Set m_veh_remark = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_remark", "veh_remark", "[veh_remark]", "[veh_remark]", 203, 0, "[veh_remark]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_remark = m_veh_remark
	End Property

	' Field veh_month_title
	Private m_veh_month_title

	Public Property Get veh_month_title()
		If Not IsObject(m_veh_month_title) Then
			Set m_veh_month_title = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_month_title", "veh_month_title", "[veh_month_title]", "[veh_month_title]", 202, 0, "[veh_month_title]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_month_title = m_veh_month_title
	End Property

	' Field veh_range
	Private m_veh_range

	Public Property Get veh_range()
		If Not IsObject(m_veh_range) Then
			Set m_veh_range = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_range", "veh_range", "[veh_range]", "[veh_range]", 202, 0, "[veh_range]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_range = m_veh_range
	End Property

	' Field veh_month_title2
	Private m_veh_month_title2

	Public Property Get veh_month_title2()
		If Not IsObject(m_veh_month_title2) Then
			Set m_veh_month_title2 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_month_title2", "veh_month_title2", "[veh_month_title2]", "[veh_month_title2]", 202, 0, "[veh_month_title2]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_month_title2 = m_veh_month_title2
	End Property

	' Field veh_range2
	Private m_veh_range2

	Public Property Get veh_range2()
		If Not IsObject(m_veh_range2) Then
			Set m_veh_range2 = NewFldObj("vehicle_record_th", "vehicle_record_th", "x_veh_range2", "veh_range2", "[veh_range2]", "[veh_range2]", 202, 0, "[veh_range2]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set veh_range2 = m_veh_range2
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
		If IsObject(m_veh_id) Then Set m_veh_id = Nothing
		If IsObject(m_vch_month) Then Set m_vch_month = Nothing
		If IsObject(m_vch_year) Then Set m_vch_year = Nothing
		If IsObject(m_veh_product_1) Then Set m_veh_product_1 = Nothing
		If IsObject(m_veh_product_2) Then Set m_veh_product_2 = Nothing
		If IsObject(m_veh_product_3) Then Set m_veh_product_3 = Nothing
		If IsObject(m_veh_product_4) Then Set m_veh_product_4 = Nothing
		If IsObject(m_veh_product_5) Then Set m_veh_product_5 = Nothing
		If IsObject(m_veh_product_6) Then Set m_veh_product_6 = Nothing
		If IsObject(m_veh_product_7) Then Set m_veh_product_7 = Nothing
		If IsObject(m_veh_product_8) Then Set m_veh_product_8 = Nothing
		If IsObject(m_veh_domes_1) Then Set m_veh_domes_1 = Nothing
		If IsObject(m_veh_domes_2) Then Set m_veh_domes_2 = Nothing
		If IsObject(m_veh_domes_3) Then Set m_veh_domes_3 = Nothing
		If IsObject(m_veh_domes_4) Then Set m_veh_domes_4 = Nothing
		If IsObject(m_veh_domes_5) Then Set m_veh_domes_5 = Nothing
		If IsObject(m_veh_domes_6) Then Set m_veh_domes_6 = Nothing
		If IsObject(m_veh_domes_7) Then Set m_veh_domes_7 = Nothing
		If IsObject(m_veh_domes_8) Then Set m_veh_domes_8 = Nothing
		If IsObject(m_veh_export_1) Then Set m_veh_export_1 = Nothing
		If IsObject(m_veh_export_2) Then Set m_veh_export_2 = Nothing
		If IsObject(m_veh_export_3) Then Set m_veh_export_3 = Nothing
		If IsObject(m_veh_export_4) Then Set m_veh_export_4 = Nothing
		If IsObject(m_veh_export_5) Then Set m_veh_export_5 = Nothing
		If IsObject(m_veh_export_6) Then Set m_veh_export_6 = Nothing
		If IsObject(m_veh_export_7) Then Set m_veh_export_7 = Nothing
		If IsObject(m_veh_export_8) Then Set m_veh_export_8 = Nothing
		If IsObject(m_veh_remark) Then Set m_veh_remark = Nothing
		If IsObject(m_veh_month_title) Then Set m_veh_month_title = Nothing
		If IsObject(m_veh_range) Then Set m_veh_range = Nothing
		If IsObject(m_veh_month_title2) Then Set m_veh_month_title2 = Nothing
		If IsObject(m_veh_range2) Then Set m_veh_range2 = Nothing
		Set RowAttrs = Nothing
	End Sub
End Class
%>
