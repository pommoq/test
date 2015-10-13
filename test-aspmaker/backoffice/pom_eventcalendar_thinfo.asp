<%

' ASPMaker configuration for Table eventcalendar_th
Dim eventcalendar_th

' Define table class
Class ceventcalendar_th

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
		Call ew_SetArObj(Fields, "eventcalendar_id", eventcalendar_id)
		Call ew_SetArObj(Fields, "eventcalendar_img", eventcalendar_img)
		Call ew_SetArObj(Fields, "eventcalendar_date", eventcalendar_date)
		Call ew_SetArObj(Fields, "eventcalendar_category", eventcalendar_category)
		Call ew_SetArObj(Fields, "eventcalendar_category_sub", eventcalendar_category_sub)
		Call ew_SetArObj(Fields, "start_date", start_date)
		Call ew_SetArObj(Fields, "end_date", end_date)
		Call ew_SetArObj(Fields, "eventcalendar_pdf", eventcalendar_pdf)
		Call ew_SetArObj(Fields, "eventcalendar_subject", eventcalendar_subject)
		Call ew_SetArObj(Fields, "eventcalendar_subject_th", eventcalendar_subject_th)
		Call ew_SetArObj(Fields, "eventcalendar_intro", eventcalendar_intro)
		Call ew_SetArObj(Fields, "eventcalendar_intro_th", eventcalendar_intro_th)
		Call ew_SetArObj(Fields, "eventcalendar_content", eventcalendar_content)
		Call ew_SetArObj(Fields, "eventcalendar_content_th", eventcalendar_content_th)
		Call ew_SetArObj(Fields, "eventcalendar_show_en", eventcalendar_show_en)
		Call ew_SetArObj(Fields, "eventcalendar_show", eventcalendar_show)
		Call ew_SetArObj(Fields, "eventcalendar_show_home", eventcalendar_show_home)
		Call ew_SetArObj(Fields, "eventcalendar_create", eventcalendar_create)
		Call ew_SetArObj(Fields, "eventcalendar_update", eventcalendar_update)
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
		TableVar = "eventcalendar_th"
	End Property

	' Table name
	Public Property Get TableName()
		TableName = "eventcalendar_th"
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
		HighlightName = "eventcalendar_th_Highlight"
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
		SqlSelect = "SELECT * FROM [eventcalendar_th]"
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
		SqlKeyFilter = "[eventcalendar_id] = @eventcalendar_id@"
	End Property

	' Return Key filter for table
	Public Property Get KeyFilter()
		Dim sKeyFilter
		sKeyFilter = SqlKeyFilter
		If Not IsNumeric(eventcalendar_id.CurrentValue) Then
			sKeyFilter = "0=1" ' Invalid key
		End If
		sKeyFilter = Replace(sKeyFilter, "@eventcalendar_id@", ew_AdjustSql(eventcalendar_id.CurrentValue)) ' Replace key value
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
			ReturnUrl = "pom_eventcalendar_thlist.asp"
		End If
	End Property

	' List url
	Public Function ListUrl()
		ListUrl = "pom_eventcalendar_thlist.asp"
	End Function

	' View url
	Public Function ViewUrl(parm)
		If parm <> "" Then
			ViewUrl = KeyUrl("pom_eventcalendar_thview.asp", UrlParm(parm))
		Else
			ViewUrl = KeyUrl("pom_eventcalendar_thview.asp", UrlParm(EW_TABLE_SHOW_DETAIL & "="))
		End If
	End Function

	' Add url
	Public Function AddUrl()
		AddUrl = "pom_eventcalendar_thadd.asp"

'		Dim sUrlParm
'		sUrlParm = UrlParm("")
'		If sUrlParm <> "" Then AddUrl = AddUrl & "?" & sUrlParm

	End Function

	' Edit url
	Public Function EditUrl(parm)
		EditUrl = KeyUrl("pom_eventcalendar_thedit.asp", UrlParm(parm))
	End Function

	' Inline edit url
	Public Function InlineEditUrl()
		InlineEditUrl = KeyUrl(ew_CurrentPage, UrlParm("a=edit"))
	End Function

	' Copy url
	Public Function CopyUrl(parm)
		CopyUrl = KeyUrl("pom_eventcalendar_thadd.asp", UrlParm(parm))
	End Function

	' Inline copy url
	Public Function InlineCopyUrl()
		InlineCopyUrl = KeyUrl(ew_CurrentPage, UrlParm("a=copy"))
	End Function

	' Delete url
	Public Function DeleteUrl()
		DeleteUrl = KeyUrl("pom_eventcalendar_thdelete.asp", UrlParm(""))
	End Function

	' Key url
	Public Function KeyUrl(url, parm)
		Dim sUrl: sUrl = url & "?"
		If parm <> "" Then sUrl = sUrl & parm & "&"
		If Not IsNull(eventcalendar_id.CurrentValue) Then
			sUrl = sUrl & "eventcalendar_id=" & eventcalendar_id.CurrentValue
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
			UrlParm = "t=eventcalendar_th"
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
				arKeys(0) = Request.QueryString("eventcalendar_id") ' eventcalendar_id

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
				eventcalendar_id.CurrentValue = key
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
		eventcalendar_id.DbValue = RsRow("eventcalendar_id")
		eventcalendar_img.DbValue = RsRow("eventcalendar_img")
		eventcalendar_date.DbValue = RsRow("eventcalendar_date")
		eventcalendar_category.DbValue = RsRow("eventcalendar_category")
		eventcalendar_category_sub.DbValue = RsRow("eventcalendar_category_sub")
		start_date.DbValue = RsRow("start_date")
		end_date.DbValue = RsRow("end_date")
		eventcalendar_pdf.DbValue = RsRow("eventcalendar_pdf")
		eventcalendar_subject.DbValue = RsRow("eventcalendar_subject")
		eventcalendar_subject_th.DbValue = RsRow("eventcalendar_subject_th")
		eventcalendar_intro.DbValue = RsRow("eventcalendar_intro")
		eventcalendar_intro_th.DbValue = RsRow("eventcalendar_intro_th")
		eventcalendar_content.DbValue = RsRow("eventcalendar_content")
		eventcalendar_content_th.DbValue = RsRow("eventcalendar_content_th")
		eventcalendar_show_en.DbValue = RsRow("eventcalendar_show_en")
		eventcalendar_show.DbValue = RsRow("eventcalendar_show")
		eventcalendar_show_home.DbValue = RsRow("eventcalendar_show_home")
		eventcalendar_create.DbValue = RsRow("eventcalendar_create")
		eventcalendar_update.DbValue = RsRow("eventcalendar_update")
	End Sub

	' Render list row values
	Sub RenderListRow()

		'
		'  Common render codes
		'
		' eventcalendar_id
		' eventcalendar_img
		' eventcalendar_date
		' eventcalendar_category
		' eventcalendar_category_sub
		' start_date
		' end_date
		' eventcalendar_pdf
		' eventcalendar_subject
		' eventcalendar_subject_th
		' eventcalendar_intro
		' eventcalendar_intro_th
		' eventcalendar_content
		' eventcalendar_content_th
		' eventcalendar_show_en
		' eventcalendar_show
		' eventcalendar_show_home
		' eventcalendar_create
		' eventcalendar_update
		' Call Row Rendering event

		Call Row_Rendering()

		'
		'  Render for View
		'
		' eventcalendar_id

		eventcalendar_id.ViewValue = eventcalendar_id.CurrentValue
		eventcalendar_id.ViewCustomAttributes = ""

		' eventcalendar_img
		eventcalendar_img.ViewValue = eventcalendar_img.CurrentValue
		eventcalendar_img.ViewCustomAttributes = ""

		' eventcalendar_date
		eventcalendar_date.ViewValue = eventcalendar_date.CurrentValue
		eventcalendar_date.ViewCustomAttributes = ""

		' eventcalendar_category
		eventcalendar_category.ViewValue = eventcalendar_category.CurrentValue
		eventcalendar_category.ViewCustomAttributes = ""

		' eventcalendar_category_sub
		eventcalendar_category_sub.ViewValue = eventcalendar_category_sub.CurrentValue
		eventcalendar_category_sub.ViewCustomAttributes = ""

		' start_date
		start_date.ViewValue = start_date.CurrentValue
		start_date.ViewCustomAttributes = ""

		' end_date
		end_date.ViewValue = end_date.CurrentValue
		end_date.ViewCustomAttributes = ""

		' eventcalendar_pdf
		eventcalendar_pdf.ViewValue = eventcalendar_pdf.CurrentValue
		eventcalendar_pdf.ViewCustomAttributes = ""

		' eventcalendar_subject
		eventcalendar_subject.ViewValue = eventcalendar_subject.CurrentValue
		eventcalendar_subject.ViewCustomAttributes = ""

		' eventcalendar_subject_th
		eventcalendar_subject_th.ViewValue = eventcalendar_subject_th.CurrentValue
		eventcalendar_subject_th.ViewCustomAttributes = ""

		' eventcalendar_intro
		eventcalendar_intro.ViewValue = eventcalendar_intro.CurrentValue
		eventcalendar_intro.ViewCustomAttributes = ""

		' eventcalendar_intro_th
		eventcalendar_intro_th.ViewValue = eventcalendar_intro_th.CurrentValue
		eventcalendar_intro_th.ViewCustomAttributes = ""

		' eventcalendar_content
		eventcalendar_content.ViewValue = eventcalendar_content.CurrentValue
		eventcalendar_content.ViewCustomAttributes = ""

		' eventcalendar_content_th
		eventcalendar_content_th.ViewValue = eventcalendar_content_th.CurrentValue
		eventcalendar_content_th.ViewCustomAttributes = ""

		' eventcalendar_show_en
		eventcalendar_show_en.ViewValue = eventcalendar_show_en.CurrentValue
		eventcalendar_show_en.ViewCustomAttributes = ""

		' eventcalendar_show
		eventcalendar_show.ViewValue = eventcalendar_show.CurrentValue
		eventcalendar_show.ViewCustomAttributes = ""

		' eventcalendar_show_home
		eventcalendar_show_home.ViewValue = eventcalendar_show_home.CurrentValue
		eventcalendar_show_home.ViewCustomAttributes = ""

		' eventcalendar_create
		eventcalendar_create.ViewValue = eventcalendar_create.CurrentValue
		eventcalendar_create.ViewCustomAttributes = ""

		' eventcalendar_update
		eventcalendar_update.ViewValue = eventcalendar_update.CurrentValue
		eventcalendar_update.ViewCustomAttributes = ""

		' eventcalendar_id
		eventcalendar_id.LinkCustomAttributes = ""
		eventcalendar_id.HrefValue = ""
		eventcalendar_id.TooltipValue = ""

		' eventcalendar_img
		eventcalendar_img.LinkCustomAttributes = ""
		eventcalendar_img.HrefValue = ""
		eventcalendar_img.TooltipValue = ""

		' eventcalendar_date
		eventcalendar_date.LinkCustomAttributes = ""
		eventcalendar_date.HrefValue = ""
		eventcalendar_date.TooltipValue = ""

		' eventcalendar_category
		eventcalendar_category.LinkCustomAttributes = ""
		eventcalendar_category.HrefValue = ""
		eventcalendar_category.TooltipValue = ""

		' eventcalendar_category_sub
		eventcalendar_category_sub.LinkCustomAttributes = ""
		eventcalendar_category_sub.HrefValue = ""
		eventcalendar_category_sub.TooltipValue = ""

		' start_date
		start_date.LinkCustomAttributes = ""
		start_date.HrefValue = ""
		start_date.TooltipValue = ""

		' end_date
		end_date.LinkCustomAttributes = ""
		end_date.HrefValue = ""
		end_date.TooltipValue = ""

		' eventcalendar_pdf
		eventcalendar_pdf.LinkCustomAttributes = ""
		eventcalendar_pdf.HrefValue = ""
		eventcalendar_pdf.TooltipValue = ""

		' eventcalendar_subject
		eventcalendar_subject.LinkCustomAttributes = ""
		eventcalendar_subject.HrefValue = ""
		eventcalendar_subject.TooltipValue = ""

		' eventcalendar_subject_th
		eventcalendar_subject_th.LinkCustomAttributes = ""
		eventcalendar_subject_th.HrefValue = ""
		eventcalendar_subject_th.TooltipValue = ""

		' eventcalendar_intro
		eventcalendar_intro.LinkCustomAttributes = ""
		eventcalendar_intro.HrefValue = ""
		eventcalendar_intro.TooltipValue = ""

		' eventcalendar_intro_th
		eventcalendar_intro_th.LinkCustomAttributes = ""
		eventcalendar_intro_th.HrefValue = ""
		eventcalendar_intro_th.TooltipValue = ""

		' eventcalendar_content
		eventcalendar_content.LinkCustomAttributes = ""
		eventcalendar_content.HrefValue = ""
		eventcalendar_content.TooltipValue = ""

		' eventcalendar_content_th
		eventcalendar_content_th.LinkCustomAttributes = ""
		eventcalendar_content_th.HrefValue = ""
		eventcalendar_content_th.TooltipValue = ""

		' eventcalendar_show_en
		eventcalendar_show_en.LinkCustomAttributes = ""
		eventcalendar_show_en.HrefValue = ""
		eventcalendar_show_en.TooltipValue = ""

		' eventcalendar_show
		eventcalendar_show.LinkCustomAttributes = ""
		eventcalendar_show.HrefValue = ""
		eventcalendar_show.TooltipValue = ""

		' eventcalendar_show_home
		eventcalendar_show_home.LinkCustomAttributes = ""
		eventcalendar_show_home.HrefValue = ""
		eventcalendar_show_home.TooltipValue = ""

		' eventcalendar_create
		eventcalendar_create.LinkCustomAttributes = ""
		eventcalendar_create.HrefValue = ""
		eventcalendar_create.TooltipValue = ""

		' eventcalendar_update
		eventcalendar_update.LinkCustomAttributes = ""
		eventcalendar_update.HrefValue = ""
		eventcalendar_update.TooltipValue = ""

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
			sSql = "SELECT * FROM [eventcalendar_th] WHERE " & sWhereList
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
					Call XmlDoc.AddField("eventcalendar_id", eventcalendar_id.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_img", eventcalendar_img.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_date", eventcalendar_date.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_category", eventcalendar_category.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_category_sub", eventcalendar_category_sub.ExportValue(Export))
					Call XmlDoc.AddField("start_date", start_date.ExportValue(Export))
					Call XmlDoc.AddField("end_date", end_date.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_pdf", eventcalendar_pdf.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_subject", eventcalendar_subject.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_subject_th", eventcalendar_subject_th.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_intro", eventcalendar_intro.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_intro_th", eventcalendar_intro_th.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_content", eventcalendar_content.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_content_th", eventcalendar_content_th.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_show_en", eventcalendar_show_en.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_show", eventcalendar_show.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_show_home", eventcalendar_show_home.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_create", eventcalendar_create.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_update", eventcalendar_update.ExportValue(Export))
				Else
					Call XmlDoc.AddField("eventcalendar_id", eventcalendar_id.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_img", eventcalendar_img.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_date", eventcalendar_date.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_category", eventcalendar_category.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_category_sub", eventcalendar_category_sub.ExportValue(Export))
					Call XmlDoc.AddField("start_date", start_date.ExportValue(Export))
					Call XmlDoc.AddField("end_date", end_date.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_pdf", eventcalendar_pdf.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_subject", eventcalendar_subject.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_subject_th", eventcalendar_subject_th.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_show_en", eventcalendar_show_en.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_show", eventcalendar_show.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_show_home", eventcalendar_show_home.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_create", eventcalendar_create.ExportValue(Export))
					Call XmlDoc.AddField("eventcalendar_update", eventcalendar_update.ExportValue(Export))
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
				If eventcalendar_id.Exportable Then Call Doc.ExportCaption(eventcalendar_id)
				If eventcalendar_img.Exportable Then Call Doc.ExportCaption(eventcalendar_img)
				If eventcalendar_date.Exportable Then Call Doc.ExportCaption(eventcalendar_date)
				If eventcalendar_category.Exportable Then Call Doc.ExportCaption(eventcalendar_category)
				If eventcalendar_category_sub.Exportable Then Call Doc.ExportCaption(eventcalendar_category_sub)
				If start_date.Exportable Then Call Doc.ExportCaption(start_date)
				If end_date.Exportable Then Call Doc.ExportCaption(end_date)
				If eventcalendar_pdf.Exportable Then Call Doc.ExportCaption(eventcalendar_pdf)
				If eventcalendar_subject.Exportable Then Call Doc.ExportCaption(eventcalendar_subject)
				If eventcalendar_subject_th.Exportable Then Call Doc.ExportCaption(eventcalendar_subject_th)
				If eventcalendar_intro.Exportable Then Call Doc.ExportCaption(eventcalendar_intro)
				If eventcalendar_intro_th.Exportable Then Call Doc.ExportCaption(eventcalendar_intro_th)
				If eventcalendar_content.Exportable Then Call Doc.ExportCaption(eventcalendar_content)
				If eventcalendar_content_th.Exportable Then Call Doc.ExportCaption(eventcalendar_content_th)
				If eventcalendar_show_en.Exportable Then Call Doc.ExportCaption(eventcalendar_show_en)
				If eventcalendar_show.Exportable Then Call Doc.ExportCaption(eventcalendar_show)
				If eventcalendar_show_home.Exportable Then Call Doc.ExportCaption(eventcalendar_show_home)
				If eventcalendar_create.Exportable Then Call Doc.ExportCaption(eventcalendar_create)
				If eventcalendar_update.Exportable Then Call Doc.ExportCaption(eventcalendar_update)
			Else
				If eventcalendar_id.Exportable Then Call Doc.ExportCaption(eventcalendar_id)
				If eventcalendar_img.Exportable Then Call Doc.ExportCaption(eventcalendar_img)
				If eventcalendar_date.Exportable Then Call Doc.ExportCaption(eventcalendar_date)
				If eventcalendar_category.Exportable Then Call Doc.ExportCaption(eventcalendar_category)
				If eventcalendar_category_sub.Exportable Then Call Doc.ExportCaption(eventcalendar_category_sub)
				If start_date.Exportable Then Call Doc.ExportCaption(start_date)
				If end_date.Exportable Then Call Doc.ExportCaption(end_date)
				If eventcalendar_pdf.Exportable Then Call Doc.ExportCaption(eventcalendar_pdf)
				If eventcalendar_subject.Exportable Then Call Doc.ExportCaption(eventcalendar_subject)
				If eventcalendar_subject_th.Exportable Then Call Doc.ExportCaption(eventcalendar_subject_th)
				If eventcalendar_show_en.Exportable Then Call Doc.ExportCaption(eventcalendar_show_en)
				If eventcalendar_show.Exportable Then Call Doc.ExportCaption(eventcalendar_show)
				If eventcalendar_show_home.Exportable Then Call Doc.ExportCaption(eventcalendar_show_home)
				If eventcalendar_create.Exportable Then Call Doc.ExportCaption(eventcalendar_create)
				If eventcalendar_update.Exportable Then Call Doc.ExportCaption(eventcalendar_update)
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
					If eventcalendar_id.Exportable Then Call Doc.ExportField(eventcalendar_id)
					If eventcalendar_img.Exportable Then Call Doc.ExportField(eventcalendar_img)
					If eventcalendar_date.Exportable Then Call Doc.ExportField(eventcalendar_date)
					If eventcalendar_category.Exportable Then Call Doc.ExportField(eventcalendar_category)
					If eventcalendar_category_sub.Exportable Then Call Doc.ExportField(eventcalendar_category_sub)
					If start_date.Exportable Then Call Doc.ExportField(start_date)
					If end_date.Exportable Then Call Doc.ExportField(end_date)
					If eventcalendar_pdf.Exportable Then Call Doc.ExportField(eventcalendar_pdf)
					If eventcalendar_subject.Exportable Then Call Doc.ExportField(eventcalendar_subject)
					If eventcalendar_subject_th.Exportable Then Call Doc.ExportField(eventcalendar_subject_th)
					If eventcalendar_intro.Exportable Then Call Doc.ExportField(eventcalendar_intro)
					If eventcalendar_intro_th.Exportable Then Call Doc.ExportField(eventcalendar_intro_th)
					If eventcalendar_content.Exportable Then Call Doc.ExportField(eventcalendar_content)
					If eventcalendar_content_th.Exportable Then Call Doc.ExportField(eventcalendar_content_th)
					If eventcalendar_show_en.Exportable Then Call Doc.ExportField(eventcalendar_show_en)
					If eventcalendar_show.Exportable Then Call Doc.ExportField(eventcalendar_show)
					If eventcalendar_show_home.Exportable Then Call Doc.ExportField(eventcalendar_show_home)
					If eventcalendar_create.Exportable Then Call Doc.ExportField(eventcalendar_create)
					If eventcalendar_update.Exportable Then Call Doc.ExportField(eventcalendar_update)
				Else
					If eventcalendar_id.Exportable Then Call Doc.ExportField(eventcalendar_id)
					If eventcalendar_img.Exportable Then Call Doc.ExportField(eventcalendar_img)
					If eventcalendar_date.Exportable Then Call Doc.ExportField(eventcalendar_date)
					If eventcalendar_category.Exportable Then Call Doc.ExportField(eventcalendar_category)
					If eventcalendar_category_sub.Exportable Then Call Doc.ExportField(eventcalendar_category_sub)
					If start_date.Exportable Then Call Doc.ExportField(start_date)
					If end_date.Exportable Then Call Doc.ExportField(end_date)
					If eventcalendar_pdf.Exportable Then Call Doc.ExportField(eventcalendar_pdf)
					If eventcalendar_subject.Exportable Then Call Doc.ExportField(eventcalendar_subject)
					If eventcalendar_subject_th.Exportable Then Call Doc.ExportField(eventcalendar_subject_th)
					If eventcalendar_show_en.Exportable Then Call Doc.ExportField(eventcalendar_show_en)
					If eventcalendar_show.Exportable Then Call Doc.ExportField(eventcalendar_show)
					If eventcalendar_show_home.Exportable Then Call Doc.ExportField(eventcalendar_show_home)
					If eventcalendar_create.Exportable Then Call Doc.ExportField(eventcalendar_create)
					If eventcalendar_update.Exportable Then Call Doc.ExportField(eventcalendar_update)
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
	' Field eventcalendar_id
	Private m_eventcalendar_id

	Public Property Get eventcalendar_id()
		If Not IsObject(m_eventcalendar_id) Then
			Set m_eventcalendar_id = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_id", "eventcalendar_id", "[eventcalendar_id]", "[eventcalendar_id]", 3, 0, "[eventcalendar_id]", False, False, FALSE, "FORMATTED TEXT")
			m_eventcalendar_id.FldDefaultErrMsg = Language.Phrase("IncorrectInteger")
		End If
		Set eventcalendar_id = m_eventcalendar_id
	End Property

	' Field eventcalendar_img
	Private m_eventcalendar_img

	Public Property Get eventcalendar_img()
		If Not IsObject(m_eventcalendar_img) Then
			Set m_eventcalendar_img = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_img", "eventcalendar_img", "[eventcalendar_img]", "[eventcalendar_img]", 202, 0, "[eventcalendar_img]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set eventcalendar_img = m_eventcalendar_img
	End Property

	' Field eventcalendar_date
	Private m_eventcalendar_date

	Public Property Get eventcalendar_date()
		If Not IsObject(m_eventcalendar_date) Then
			Set m_eventcalendar_date = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_date", "eventcalendar_date", "[eventcalendar_date]", "FORMAT([eventcalendar_date], '')", 135, 8, "[eventcalendar_date]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set eventcalendar_date = m_eventcalendar_date
	End Property

	' Field eventcalendar_category
	Private m_eventcalendar_category

	Public Property Get eventcalendar_category()
		If Not IsObject(m_eventcalendar_category) Then
			Set m_eventcalendar_category = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_category", "eventcalendar_category", "[eventcalendar_category]", "[eventcalendar_category]", 202, 0, "[eventcalendar_category]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set eventcalendar_category = m_eventcalendar_category
	End Property

	' Field eventcalendar_category_sub
	Private m_eventcalendar_category_sub

	Public Property Get eventcalendar_category_sub()
		If Not IsObject(m_eventcalendar_category_sub) Then
			Set m_eventcalendar_category_sub = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_category_sub", "eventcalendar_category_sub", "[eventcalendar_category_sub]", "[eventcalendar_category_sub]", 202, 0, "[eventcalendar_category_sub]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set eventcalendar_category_sub = m_eventcalendar_category_sub
	End Property

	' Field start_date
	Private m_start_date

	Public Property Get start_date()
		If Not IsObject(m_start_date) Then
			Set m_start_date = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_start_date", "start_date", "[start_date]", "FORMAT([start_date], '')", 135, 8, "[start_date]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set start_date = m_start_date
	End Property

	' Field end_date
	Private m_end_date

	Public Property Get end_date()
		If Not IsObject(m_end_date) Then
			Set m_end_date = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_end_date", "end_date", "[end_date]", "FORMAT([end_date], '')", 135, 8, "[end_date]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set end_date = m_end_date
	End Property

	' Field eventcalendar_pdf
	Private m_eventcalendar_pdf

	Public Property Get eventcalendar_pdf()
		If Not IsObject(m_eventcalendar_pdf) Then
			Set m_eventcalendar_pdf = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_pdf", "eventcalendar_pdf", "[eventcalendar_pdf]", "[eventcalendar_pdf]", 202, 0, "[eventcalendar_pdf]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set eventcalendar_pdf = m_eventcalendar_pdf
	End Property

	' Field eventcalendar_subject
	Private m_eventcalendar_subject

	Public Property Get eventcalendar_subject()
		If Not IsObject(m_eventcalendar_subject) Then
			Set m_eventcalendar_subject = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_subject", "eventcalendar_subject", "[eventcalendar_subject]", "[eventcalendar_subject]", 202, 0, "[eventcalendar_subject]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set eventcalendar_subject = m_eventcalendar_subject
	End Property

	' Field eventcalendar_subject_th
	Private m_eventcalendar_subject_th

	Public Property Get eventcalendar_subject_th()
		If Not IsObject(m_eventcalendar_subject_th) Then
			Set m_eventcalendar_subject_th = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_subject_th", "eventcalendar_subject_th", "[eventcalendar_subject_th]", "[eventcalendar_subject_th]", 202, 0, "[eventcalendar_subject_th]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set eventcalendar_subject_th = m_eventcalendar_subject_th
	End Property

	' Field eventcalendar_intro
	Private m_eventcalendar_intro

	Public Property Get eventcalendar_intro()
		If Not IsObject(m_eventcalendar_intro) Then
			Set m_eventcalendar_intro = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_intro", "eventcalendar_intro", "[eventcalendar_intro]", "[eventcalendar_intro]", 203, 0, "[eventcalendar_intro]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set eventcalendar_intro = m_eventcalendar_intro
	End Property

	' Field eventcalendar_intro_th
	Private m_eventcalendar_intro_th

	Public Property Get eventcalendar_intro_th()
		If Not IsObject(m_eventcalendar_intro_th) Then
			Set m_eventcalendar_intro_th = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_intro_th", "eventcalendar_intro_th", "[eventcalendar_intro_th]", "[eventcalendar_intro_th]", 203, 0, "[eventcalendar_intro_th]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set eventcalendar_intro_th = m_eventcalendar_intro_th
	End Property

	' Field eventcalendar_content
	Private m_eventcalendar_content

	Public Property Get eventcalendar_content()
		If Not IsObject(m_eventcalendar_content) Then
			Set m_eventcalendar_content = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_content", "eventcalendar_content", "[eventcalendar_content]", "[eventcalendar_content]", 203, 0, "[eventcalendar_content]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set eventcalendar_content = m_eventcalendar_content
	End Property

	' Field eventcalendar_content_th
	Private m_eventcalendar_content_th

	Public Property Get eventcalendar_content_th()
		If Not IsObject(m_eventcalendar_content_th) Then
			Set m_eventcalendar_content_th = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_content_th", "eventcalendar_content_th", "[eventcalendar_content_th]", "[eventcalendar_content_th]", 203, 0, "[eventcalendar_content_th]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set eventcalendar_content_th = m_eventcalendar_content_th
	End Property

	' Field eventcalendar_show_en
	Private m_eventcalendar_show_en

	Public Property Get eventcalendar_show_en()
		If Not IsObject(m_eventcalendar_show_en) Then
			Set m_eventcalendar_show_en = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_show_en", "eventcalendar_show_en", "[eventcalendar_show_en]", "[eventcalendar_show_en]", 202, 0, "[eventcalendar_show_en]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set eventcalendar_show_en = m_eventcalendar_show_en
	End Property

	' Field eventcalendar_show
	Private m_eventcalendar_show

	Public Property Get eventcalendar_show()
		If Not IsObject(m_eventcalendar_show) Then
			Set m_eventcalendar_show = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_show", "eventcalendar_show", "[eventcalendar_show]", "[eventcalendar_show]", 202, 0, "[eventcalendar_show]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set eventcalendar_show = m_eventcalendar_show
	End Property

	' Field eventcalendar_show_home
	Private m_eventcalendar_show_home

	Public Property Get eventcalendar_show_home()
		If Not IsObject(m_eventcalendar_show_home) Then
			Set m_eventcalendar_show_home = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_show_home", "eventcalendar_show_home", "[eventcalendar_show_home]", "[eventcalendar_show_home]", 202, 0, "[eventcalendar_show_home]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set eventcalendar_show_home = m_eventcalendar_show_home
	End Property

	' Field eventcalendar_create
	Private m_eventcalendar_create

	Public Property Get eventcalendar_create()
		If Not IsObject(m_eventcalendar_create) Then
			Set m_eventcalendar_create = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_create", "eventcalendar_create", "[eventcalendar_create]", "[eventcalendar_create]", 202, 0, "[eventcalendar_create]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set eventcalendar_create = m_eventcalendar_create
	End Property

	' Field eventcalendar_update
	Private m_eventcalendar_update

	Public Property Get eventcalendar_update()
		If Not IsObject(m_eventcalendar_update) Then
			Set m_eventcalendar_update = NewFldObj("eventcalendar_th", "eventcalendar_th", "x_eventcalendar_update", "eventcalendar_update", "[eventcalendar_update]", "[eventcalendar_update]", 202, 0, "[eventcalendar_update]", False, False, FALSE, "FORMATTED TEXT")
		End If
		Set eventcalendar_update = m_eventcalendar_update
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
		If IsObject(m_eventcalendar_id) Then Set m_eventcalendar_id = Nothing
		If IsObject(m_eventcalendar_img) Then Set m_eventcalendar_img = Nothing
		If IsObject(m_eventcalendar_date) Then Set m_eventcalendar_date = Nothing
		If IsObject(m_eventcalendar_category) Then Set m_eventcalendar_category = Nothing
		If IsObject(m_eventcalendar_category_sub) Then Set m_eventcalendar_category_sub = Nothing
		If IsObject(m_start_date) Then Set m_start_date = Nothing
		If IsObject(m_end_date) Then Set m_end_date = Nothing
		If IsObject(m_eventcalendar_pdf) Then Set m_eventcalendar_pdf = Nothing
		If IsObject(m_eventcalendar_subject) Then Set m_eventcalendar_subject = Nothing
		If IsObject(m_eventcalendar_subject_th) Then Set m_eventcalendar_subject_th = Nothing
		If IsObject(m_eventcalendar_intro) Then Set m_eventcalendar_intro = Nothing
		If IsObject(m_eventcalendar_intro_th) Then Set m_eventcalendar_intro_th = Nothing
		If IsObject(m_eventcalendar_content) Then Set m_eventcalendar_content = Nothing
		If IsObject(m_eventcalendar_content_th) Then Set m_eventcalendar_content_th = Nothing
		If IsObject(m_eventcalendar_show_en) Then Set m_eventcalendar_show_en = Nothing
		If IsObject(m_eventcalendar_show) Then Set m_eventcalendar_show = Nothing
		If IsObject(m_eventcalendar_show_home) Then Set m_eventcalendar_show_home = Nothing
		If IsObject(m_eventcalendar_create) Then Set m_eventcalendar_create = Nothing
		If IsObject(m_eventcalendar_update) Then Set m_eventcalendar_update = Nothing
		Set RowAttrs = Nothing
	End Sub
End Class
%>
