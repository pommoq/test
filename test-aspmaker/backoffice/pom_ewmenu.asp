<%@ CodePage="65001" LCID="1054" %>
<!--#include file="pom_ewcfg11.asp"-->
<!--#include file="pom_aspfn11.asp"-->
<!--#include file="pom_userfn11.asp"-->
<!-- Begin Main Menu -->
<div class="ewMenu">
<%

' Initialize language object
If IsEmpty(Language) Then
	Set Language = New cLanguage
	Call Language.LoadPhrases()
End If
If IsEmpty(Conn) Then Call ew_Connect()

' Generate all menu items
Dim RootMenu
Set RootMenu = new cMenu
RootMenu.Id = EW_MENUBAR_ID
%>
<%

' Get Menu Text
Function GetMenuText(Id, Text)
	GetMenuText = Language.MenuPhrase(Id, "MenuText")
	If GetMenuText = "" Then GetMenuText = Text
End Function
%>
<%

' Generate all menu items
RootMenu.Id = EW_MENUBAR_ID
RootMenu.IsRoot = True
RootMenu.AddMenuItem 1, GetMenuText("1", "@news"), "pom_z40newslist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 2, GetMenuText("2", "admins"), "pom_adminslist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 3, GetMenuText("3", "banner"), "pom_bannerlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 4, GetMenuText("4", "banner logo 01"), "pom_banner_logo_01list.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 5, GetMenuText("5", "banner logo 01 th"), "pom_banner_logo_01_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 6, GetMenuText("6", "banner logo 02"), "pom_banner_logo_02list.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 7, GetMenuText("7", "banner logo 02 th"), "pom_banner_logo_02_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 8, GetMenuText("8", "banner th"), "pom_banner_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 9, GetMenuText("9", "company"), "pom_companylist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 10, GetMenuText("10", "company th"), "pom_company_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 11, GetMenuText("11", "department"), "pom_departmentlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 12, GetMenuText("12", "department th"), "pom_department_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 13, GetMenuText("13", "e library"), "pom_e_librarylist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 14, GetMenuText("14", "e library th"), "pom_e_library_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 15, GetMenuText("15", "eventcalendar"), "pom_eventcalendarlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 16, GetMenuText("16", "eventcalendar pdf file"), "pom_eventcalendar_pdf_filelist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 17, GetMenuText("17", "eventcalendar pdf file th"), "pom_eventcalendar_pdf_file_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 18, GetMenuText("18", "eventcalendar th"), "pom_eventcalendar_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 19, GetMenuText("19", "homepage"), "pom_homepagelist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 20, GetMenuText("20", "job"), "pom_joblist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 21, GetMenuText("21", "job file"), "pom_job_filelist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 22, GetMenuText("22", "job file th"), "pom_job_file_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 23, GetMenuText("23", "job th"), "pom_job_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 24, GetMenuText("24", "journal"), "pom_journallist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 25, GetMenuText("25", "news"), "pom_newslist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 26, GetMenuText("26", "news pdf file"), "pom_news_pdf_filelist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 27, GetMenuText("27", "news pdf file th"), "pom_news_pdf_file_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 28, GetMenuText("28", "news sale"), "pom_news_salelist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 29, GetMenuText("29", "news th"), "pom_news_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 30, GetMenuText("30", "office"), "pom_officelist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 31, GetMenuText("31", "office th"), "pom_office_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 32, GetMenuText("32", "person"), "pom_personlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 33, GetMenuText("33", "person th"), "pom_person_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 34, GetMenuText("34", "research"), "pom_researchlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 35, GetMenuText("35", "research pdf file"), "pom_research_pdf_filelist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 36, GetMenuText("36", "research pdf file th"), "pom_research_pdf_file_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 37, GetMenuText("37", "research th"), "pom_research_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 38, GetMenuText("38", "sys admin menu"), "pom_sys_admin_menulist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 39, GetMenuText("39", "sys menu"), "pom_sys_menulist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 40, GetMenuText("40", "vehicle record"), "pom_vehicle_recordlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 41, GetMenuText("41", "vehicle record th"), "pom_vehicle_record_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 42, GetMenuText("42", "video"), "pom_videolist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 43, GetMenuText("43", "video th"), "pom_video_thlist.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem 44, GetMenuText("44", "Query 2"), "pom_query2list.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem &HFFFFFFFF, Language.Phrase("Logout"), "pom_logout.asp", -1, "", "", IsLoggedIn(), False, False
RootMenu.AddMenuItem &HFFFFFFFF, Language.Phrase("Login"), "pom_login.asp", -1, "", "", (Not IsLoggedIn() And Right(Request.ServerVariables("URL"), Len("pom_login.asp")) <> "pom_login.asp"), False, False
RootMenu.Render(False)
Set RootMenu = Nothing
%>
</div>
<!-- End Main Menu -->
