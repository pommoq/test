<HTML>
<HEAD>
	<TITLE>JavaScript Toolbox - Calendar Popup To Select Date</TITLE>
<SCRIPT LANGUAGE="JavaScript" SRC="CalendarPopup.js"></SCRIPT>
 
<!-- This javascript is only used for the show/hide source on my example page.
     It is not used by the Calendar Popup script -->
<SCRIPT LANGUAGE="JavaScript" SRC="common.js"></SCRIPT>
 
<!-- This prints out the default stylehseets used by the DIV style calendar.
     Only needed if you are using the DIV style popup -->
<SCRIPT LANGUAGE="JavaScript">document.write(getCalendarStyles());</SCRIPT>
 
<!-- These styles are here only as an example of how you can over-ride the default
     styles that are included in the script itself. -->
<STYLE> 
	.TESTcpYearNavigation,
	.TESTcpMonthNavigation
			{
			background-color:#6677DD;
			text-align:center;
			vertical-align:center;
			text-decoration:none;
			color:#FFFFFF;
			font-weight:bold;
			}
	.TESTcpDayColumnHeader,
	.TESTcpYearNavigation,
	.TESTcpMonthNavigation,
	.TESTcpCurrentMonthDate,
	.TESTcpCurrentMonthDateDisabled,
	.TESTcpOtherMonthDate,
	.TESTcpOtherMonthDateDisabled,
	.TESTcpCurrentDate,
	.TESTcpCurrentDateDisabled,
	.TESTcpTodayText,
	.TESTcpTodayTextDisabled,
	.TESTcpText
			{
			font-family:arial;
			font-size:8pt;
			}
	TD.TESTcpDayColumnHeader
			{
			text-align:right;
			border:solid thin #6677DD;
			border-width:0 0 1 0;
			}
	.TESTcpCurrentMonthDate,
	.TESTcpOtherMonthDate,
	.TESTcpCurrentDate
			{
			text-align:right;
			text-decoration:none;
			}
	.TESTcpCurrentMonthDateDisabled,
	.TESTcpOtherMonthDateDisabled,
	.TESTcpCurrentDateDisabled
			{
			color:#D0D0D0;
			text-align:right;
			text-decoration:line-through;
			}
	.TESTcpCurrentMonthDate
			{
			color:#6677DD;
			font-weight:bold;
			}
	.TESTcpCurrentDate
			{
			color: #FFFFFF;
			font-weight:bold;
			}
	.TESTcpOtherMonthDate
			{
			color:#808080;
			}
	TD.TESTcpCurrentDate
			{
			color:#FFFFFF;
			background-color: #6677DD;
			border-width:1;
			border:solid thin #000000;
			}
	TD.TESTcpCurrentDateDisabled
			{
			border-width:1;
			border:solid thin #FFAAAA;
			}
	TD.TESTcpTodayText,
	TD.TESTcpTodayTextDisabled
			{
			border:solid thin #6677DD;
			border-width:1 0 0 0;
			}
	A.TESTcpTodayText,
	SPAN.TESTcpTodayTextDisabled
			{
			height:20px;
			}
	A.TESTcpTodayText
			{
			color:#6677DD;
			font-weight:bold;
			}
	SPAN.TESTcpTodayTextDisabled
			{
			color:#D0D0D0;
			}
	.TESTcpBorder
			{
			border:solid thin #6677DD;
			}
</STYLE>
</HEAD>
<BODY BGCOLOR=#FFFFFF LINK="#00615F" VLINK="#00615F" ALINK="#00615F">
<div style="border:50px solid red;padding:10px;">
<h1>JavascriptToolbox.com Now Available!</h1>
Almost all of my javascript code has been moved over to its new home at <a href="http://www.JavascriptToolbox.com/">The Javascript Toolbox</a>. Please go there to find the latest scripts, information, etc. These pages will remain here for a while for historical purposes in case anyone needs a production copy of old code.
</div>
 
<TABLE WIDTH=600 CELLPADDING=5>
<TR>
	<TD><IMG SRC="../icon.gif" WIDTH="130" HEIGHT="107" ALT="" BORDER="0"></TD>
	<TD VALIGN=MIDDLE>
		<FONT SIZE="+3">Calendar Popup</FONT>
		<HR NOSHADE WIDTH=300 ALIGN=LEFT COLOR="black">
		[<A HREF="../">Javascript Toolbox</A>]&nbsp;&nbsp;[<SPAN STYLE="background-color:yellow;">Example</SPAN>]&nbsp;&nbsp;[<A HREF="source.html">Source</A>]
	</TD>
</TR>
</TABLE>
 
<TABLE WIDTH="100%" BORDER="0"><TR><TD WIDTH="100%" ALIGN="LEFT" VALIGN="TOP">
 
<U><B>Description:</B></U>
<BR>
 
This script uses DHTML or Popup windows to display a calendar for the user to select a date. It was designed to look and behave like Microsoft Outlook.<BR>
It can be implemented in only a few lines of code, yet also provides customization options to make it work correctly in any country's display format, etc.
<P>
<B>Note:</B> Why are form elements or &lt;SELECT&gt; boxes showing
over top of the DIV-style calendar popup? It's not a bug in the javascript -
it's a bug/feature of browsers. See this
<A href="http://www.webreference.com/dhtml/diner/seethru/">explanation</A>
by WebReference.
<BR><BR>
<U><B>Example:</B></U>
<BR>
Below are multiple examples of the CalendarPopup in use. Each is slightly different to show different capabilities of the script.
<BR>
Click the "Show Source" links for each example to see how it was done, and hover over the "Select" links to see how they are coded. Or view the source of the entire page!
<BR>
<BR>
 
I have also provided a <b><a href="simple.html" target="_blank">The Simplest Possible Implementation Of A Calendar Popup</a></b> in case you're overwhelmed by this page :)
<BR><BR>
<FORM>
 
<!-- ================================================================================== --><HR>
 
Default calendar.<BR>
<SCRIPT LANGUAGE="JavaScript" ID="js1"> 
var cal1 = new CalendarPopup();
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js1");</SCRIPT>
<INPUT TYPE="text" NAME="date1" VALUE="" SIZE=25>
<A HREF="#" onClick="cal1.select(document.forms[0].date1,'anchor1','MM/dd/yyyy'); return false;" TITLE="cal1.select(document.forms[0].date1,'anchor1','MM/dd/yyyy'); return false;" NAME="anchor1" ID="anchor1">select</A>
 
<!-- ================================================================================== --><HR>
 
Default calendar using the DIV-style display.<BR>
<SCRIPT LANGUAGE="JavaScript" ID="jscal1x"> 
var cal1x = new CalendarPopup("testdiv1");
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("jscal1x");</SCRIPT>
<INPUT TYPE="text" NAME="date1x" VALUE="" SIZE=25>
<A HREF="#" onClick="cal1x.select(document.forms[0].date1x,'anchor1x','MM/dd/yyyy'); return false;" TITLE="cal1x.select(document.forms[0].date1x,'anchor1x','MM/dd/yyyy'); return false;" NAME="anchor1x" ID="anchor1x">select</A>
 
<!-- ================================================================================== --><HR>
 
Default calendar using the DIV-style display, with navigation drop-downs enabled.<BR>
<SCRIPT LANGUAGE="JavaScript" ID="jscal1xx"> 
var cal1xx = new CalendarPopup("testdiv1");
cal1xx.showNavigationDropdowns();
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("jscal1xx");</SCRIPT>
<INPUT TYPE="text" NAME="date1xx" VALUE="" SIZE=25>
<A HREF="#" onClick="cal1xx.select(document.forms[0].date1xx,'anchor1xx','MM/dd/yyyy'); return false;" TITLE="cal1xx.select(document.forms[0].date1xx,'anchor1xx','MM/dd/yyyy'); return false;" NAME="anchor1xx" ID="anchor1xx">select</A>
 
<!-- ================================================================================== --><HR>
 
DIV-style calendar using a CSS prefix and different styles define in this HTML page (view source to see the defined styles)<BR>
<SCRIPT LANGUAGE="JavaScript" ID="js18"> 
var cal18 = new CalendarPopup("testdiv1");
cal18.setCssPrefix("TEST");
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js18");</SCRIPT>
<INPUT TYPE="text" NAME="date18" VALUE="" SIZE=25>
<A HREF="#" onClick="cal18.select(document.forms[0].date18,'anchor18','MM/dd/yyyy'); return false;" TITLE="cal18.select(document.forms[0].date18,'anchor1x','MM/dd/yyyy'); return false;" NAME="anchor18" ID="anchor18">select</A>
 
<!-- ================================================================================== --><HR>
 
Some dates manually disabled from selection.<BR>
Dates disabled: Anything up to today, December 25, 2007, and anything after January 1, 2008.<BR>
<SCRIPT LANGUAGE="JavaScript" ID="js17"> 
var now = new Date();
var cal17 = new CalendarPopup("testdiv1");
cal17.addDisabledDates(null,formatDate(now,"yyyy-MM-dd"));
cal17.addDisabledDates("12/25/2007");
cal17.addDisabledDates("Jan 1, 2008",null);
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js17");</SCRIPT>
<INPUT TYPE="text" NAME="date17" VALUE="" SIZE=25>
<A HREF="#" onClick="cal17.select(document.forms[0].date17,'anchor17','MM/dd/yyyy'); return false;" TITLE="cal17.select(document.forms[0].date17,'anchor17','MM/dd/yyyy'); return false;" NAME="anchor17" ID="anchor17">select</A>
 
<!-- ================================================================================== --><HR>
 
Week-end select.<BR>
<SCRIPT LANGUAGE="JavaScript" ID="js8"> 
var cal8 = new CalendarPopup();
cal8.setDisplayType("week-end");
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js8");</SCRIPT>
<INPUT TYPE="text" NAME="date8" VALUE="" SIZE=25>
<A HREF="#"
onClick="cal8.select(document.forms[0].date8,'anchor8','MM/dd/yyyy'); return false;" TITLE="cal8.select(document.forms[0].date8,'anchor8','MM/dd/yyyy'); return false;" NAME="anchor8" ID="anchor8">select</A>
 
<!-- ================================================================================== --><HR>
 
Calendar with showYearNavigation() enabled.<BR>
<SCRIPT LANGUAGE="JavaScript" ID="js2"> 
var cal2 = new CalendarPopup();
cal2.showYearNavigation();
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js2");</SCRIPT>
<INPUT TYPE="text" NAME="date2" VALUE="" SIZE=25>
<A HREF="#" onClick="cal2.select(document.forms[0].date2,'anchor2','MM/dd/yyyy'); return false;" TITLE="cal2.select(document.forms[0].date2,'anchor2','MM/dd/yyyy'); return false;" NAME="anchor2" ID="anchor2">select</A>
 
<!-- ================================================================================== --><HR>
 
Calendar with showYearNavigation() enabled and showYearNavigationInput() enabled, to allow manual entering of years<BR>
<SCRIPT LANGUAGE="JavaScript" ID="js19"> 
var cal19 = new CalendarPopup();
cal19.showYearNavigation();
cal19.showYearNavigationInput();
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js19");</SCRIPT>
<INPUT TYPE="text" NAME="date19" VALUE="" SIZE=25>
<A HREF="#" onClick="cal19.select(document.forms[0].date19,'anchor19','MM/dd/yyyy'); return false;" TITLE="cal19.select(document.forms[0].date19,'anchor19','MM/dd/yyyy'); return false;" NAME="anchor19" ID="anchor19">select</A>
 
<!-- ================================================================================== --><HR>
 
Calendar with only Saturdays allowed to be selected enabled.<BR>
<SCRIPT LANGUAGE="JavaScript" ID="js3"> 
var cal3 = new CalendarPopup();
cal3.setDisabledWeekDays(0,1,2,3,4,5);
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js3");</SCRIPT>
<INPUT TYPE="text" NAME="date3" VALUE="" SIZE=25>
<A HREF="#" onClick="cal3.select(document.forms[0].date3,'anchor3','MM/dd/yyyy'); return false;" TITLE="cal3.select(document.forms[0].date3,'anchor3','MM/dd/yyyy'); return false;" NAME="anchor3" ID="anchor3">select</A>
 
<!-- ================================================================================== --><HR>
 
German Calendar, with modified month names, day names, and week starting on Monday. Date format changed to dd/MM/yyyy<BR>
<SCRIPT LANGUAGE="JavaScript" ID="js4"> 
var cal4 = new CalendarPopup();
cal4.setMonthNames('Januar','Februar','M�rz','April','Mag','Juni','Juli','August','September','Oktober','November','Dezember');
cal4.setDayHeaders('S','M','D','M','D','F','S');
cal4.setWeekStartDay(1);
cal4.setTodayText("Heute");
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js4");</SCRIPT>
<INPUT TYPE="text" NAME="date4" VALUE="" SIZE=25>
<A HREF="#" onClick="cal4.select(document.forms[0].date4,'anchor4','dd/MM/yyyy'); return false;" TITLE="cal4.select(document.forms[0].date4,'anchor4','dd/MM/yyyy'); return false;" NAME="anchor4" ID="anchor4">select</A>
 
<!-- ================================================================================== --><HR>
 
Month-select calendar<BR>
<SCRIPT LANGUAGE="JavaScript" ID="js5"> 
var cal5 = new CalendarPopup();
cal5.setDisplayType("month");
cal5.setReturnMonthFunction("monthReturn");
cal5.showYearNavigation();
function monthReturn(y,m) {
	document.forms[0].date5.value=m+"/"+y;
	}
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js5");</SCRIPT>
<INPUT TYPE="text" NAME="date5" VALUE="" SIZE=25>
<A HREF="#" onClick="cal5.showCalendar('anchor5'); return false;" TITLE="cal5.showCalendar('anchor5'); return false;" NAME="anchor5" ID="anchor5">select</A>
 
<!-- ================================================================================== --><HR>
 
Quarter-select calendar<BR>
<SCRIPT LANGUAGE="JavaScript" ID="js6"> 
var cal6 = new CalendarPopup();
cal6.setDisplayType("quarter");
cal6.setReturnQuarterFunction("quarterReturn");
cal6.showYearNavigation();
function quarterReturn(y,q) {
	document.forms[0].date6.value="Quarter "+q+", "+y;
	}
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js6");</SCRIPT>
<INPUT TYPE="text" NAME="date6" VALUE="" SIZE=25>
<A HREF="#" onClick="cal6.showCalendar('anchor6'); return false;" TITLE="cal6.showCalendar('anchor6'); return false;" NAME="anchor6" ID="anchor6">select</A>
 
<!-- ================================================================================== --><HR>
 
Year-select calendar<BR>
<SCRIPT LANGUAGE="JavaScript" ID="js7"> 
var cal7 = new CalendarPopup();
cal7.setDisplayType("year");
cal7.setReturnYearFunction("yearReturn");
cal7.showYearNavigation();
function yearReturn(y) {
	document.forms[0].date7.value=y;
	}
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js7");</SCRIPT>
<INPUT TYPE="text" NAME="date7" VALUE="" SIZE=25>
<A HREF="#" onClick="cal7.showCalendar('anchor7'); return false;" TITLE="cal7.showCalendar('anchor7'); return false;" NAME="anchor7" ID="anchor7">select</A>
 
<!-- ================================================================================== --><HR>
 1111
Default calendar, but results are split into multiple fields.<BR>
<SCRIPT LANGUAGE="JavaScript" ID="js9"> 
var cal9 = new CalendarPopup();
cal9.setReturnFunction("setMultipleValues111");
function setMultipleValues111(y,m,d) {
	document.forms[0].date9_year.value=y;
	document.forms[0].date9_month.value=m;
	document.forms[0].date9_date.value=d;
	}
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js9");</SCRIPT>
<INPUT TYPE="text" NAME="date9_month" VALUE="" SIZE=3> /
<INPUT TYPE="text" NAME="date9_date" VALUE="" SIZE=3> /
<INPUT TYPE="text" NAME="date9_year" VALUE="" SIZE=5> (m/d/y)
<A HREF="#" onClick="cal9.showCalendar('anchor9'); return false;" TITLE="cal9.showCalendar('anchor9'); return false;" NAME="anchor9" ID="anchor9">select</A>
 
<!-- ================================================================================== --><HR>
 
<SCRIPT LANGUAGE="JavaScript" ID="js10"> 
var cal10 = new CalendarPopup();
cal10.setReturnFunction("setMultipleValues2");
function setMultipleValues2(y,m,d) {
	document.forms[0].date10_year.value=y;
	document.forms[0].date10_month.value=LZ(m);
	document.forms[0].date10_date.value=LZ(d);
	}
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js10");</SCRIPT>
<INPUT TYPE="text" NAME="date10_month" VALUE="" SIZE=3> /
<INPUT TYPE="text" NAME="date10_date" VALUE="" SIZE=3> /
<INPUT TYPE="text" NAME="date10_year" VALUE="" SIZE=5> (mm/dd/yyyy)
<A HREF="#" onClick="cal10.showCalendar('anchor10'); return false;" TITLE="cal10.showCalendar('anchor10'); return false;" NAME="anchor10" ID="anchor10">select</A>
 
<!-- ================================================================================== --><HR>
 
<SCRIPT LANGUAGE="JavaScript" ID="js11"> 
var cal11 = new CalendarPopup();
cal11.setReturnFunction("setMultipleValues3");
function setMultipleValues3(y,m,d) {
	document.forms[0].date11_year.value=y;
	document.forms[0].date11_month.selectedIndex=m;
	document.forms[0].date11_date.selectedIndex=d;
	}
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js11");</SCRIPT>
<SELECT NAME="date11_month">
	<OPTION>
	<OPTION VALUE="Jan">January
	<OPTION VALUE="Feb">February
	<OPTION VALUE="Mar">March
	<OPTION VALUE="Apr">April
	<OPTION VALUE="May">May
	<OPTION VALUE="Jun">June
	<OPTION VALUE="Jul">July
	<OPTION VALUE="Aug">August
	<OPTION VALUE="Sep">September
	<OPTION VALUE="Oct">October
	<OPTION VALUE="Nov">November
	<OPTION VALUE="Dec">December
</SELECT>
 /
<SELECT NAME="date11_date">
	<OPTION>
	<OPTION VALUE="1">1
	<OPTION VALUE="2">2
	<OPTION VALUE="3">3
	<OPTION VALUE="4">4
	<OPTION VALUE="5">5
	<OPTION VALUE="6">6
	<OPTION VALUE="7">7
	<OPTION VALUE="8">8
	<OPTION VALUE="9">9
	<OPTION VALUE="10">10
	<OPTION VALUE="11">11
	<OPTION VALUE="12">12
	<OPTION VALUE="13">13
	<OPTION VALUE="14">14
	<OPTION VALUE="15">15
	<OPTION VALUE="16">16
	<OPTION VALUE="17">17
	<OPTION VALUE="18">18
	<OPTION VALUE="19">19
	<OPTION VALUE="20">20
	<OPTION VALUE="21">21
	<OPTION VALUE="22">22
	<OPTION VALUE="23">23
	<OPTION VALUE="24">24
	<OPTION VALUE="25">25
	<OPTION VALUE="26">26
	<OPTION VALUE="27">27
	<OPTION VALUE="28">28
	<OPTION VALUE="29">29
	<OPTION VALUE="30">30
	<OPTION VALUE="31">31
</SELECT>
<INPUT TYPE="text" NAME="date11_year" VALUE="" SIZE=5>
<A HREF="#" onClick="cal11.showCalendar('anchor11'); return false;" TITLE="cal11.showCalendar('anchor11'); return false;" NAME="anchor11" ID="anchor11">select</A>
 
<!-- ================================================================================== --><HR>
 
Calendar with popup pre-selected to be January 29, 1974 (my birthday!)<BR>
<SCRIPT LANGUAGE="JavaScript" ID="js12"> 
var cal12 = new CalendarPopup();
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js12");</SCRIPT>
<INPUT TYPE="text" NAME="date12" VALUE="" SIZE=25>
<A HREF="#" onClick="cal12.select(document.forms[0].date12,'anchor12','MM/dd/yyyy','01/29/1974'); return false;" TITLE="cal12.select(document.forms[0].date12,'anchor12','MM/dd/yyyy','01/29/1974'); return false;" NAME="anchor12" ID="anchor12">select</A>
 
<!-- ================================================================================== --><HR>
 
Start date and end date, with end date popup defaulting to same date as start date<BR>
<SCRIPT LANGUAGE="JavaScript" ID="js13"> 
var cal13 = new CalendarPopup();
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js13");</SCRIPT>
Start: <INPUT TYPE="text" NAME="date13" VALUE="" SIZE=25>
<A HREF="#" onClick="cal13.select(document.forms[0].date13,'anchor13','MM/dd/yyyy'); return false;" TITLE="cal13.select(document.forms[0].date13,'anchor13','MM/dd/yyyy'); return false;" NAME="anchor13" ID="anchor13">select</A>
&nbsp;&nbsp;&nbsp;
End: <INPUT TYPE="text" NAME="date14" VALUE="" SIZE=25>
<A HREF="#" onClick="cal13.select(document.forms[0].date14,'anchor14','MM/dd/yyyy',(document.forms[0].date14.value=='')?document.forms[0].date13.value:null); return false;" TITLE="cal13.select(document.forms[0].date14,'anchor14','MM/dd/yyyy',(document.forms[0].date14.value=='')?document.forms[0].date13.value:null); return false;" NAME="anchor14" ID="anchor14">select</A>
 
<!-- ================================================================================== --><HR>
 
<SCRIPT LANGUAGE="JavaScript" ID="js15"> 
var cal15 = new CalendarPopup();
cal15.setReturnFunction("setMultipleValues4");
function setMultipleValues4(y,m,d) {
	document.forms[0].date15_year.value=y;
	document.forms[0].date15_month.selectedIndex=m;
	for (var i=0; i<document.forms[0].date15_date.options.length; i++) {
		if (document.forms[0].date15_date.options[i].value==d) {
			document.forms[0].date15_date.selectedIndex=i;
			}
		}
	}
var cal16 = new CalendarPopup();
cal16.setReturnFunction("setMultipleValues5");
function setMultipleValues5(y,m,d) {
	document.forms[0].date16_year.value=y;
	document.forms[0].date16_month.selectedIndex=m;
	for (var i=0; i<document.forms[0].date16_date.options.length; i++) {
		if (document.forms[0].date16_date.options[i].value==d) {
			document.forms[0].date16_date.selectedIndex=i;
			}
		}
	}
function getDateString(y_obj,m_obj,d_obj) {
	var y = y_obj.options[y_obj.selectedIndex].value;
	var m = m_obj.options[m_obj.selectedIndex].value;
	var d = d_obj.options[d_obj.selectedIndex].value;
	if (y=="" || m=="") { return null; }
	if (d=="") { d=1; }
	return str= y+'-'+m+'-'+d;
	}
</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->
<SCRIPT LANGUAGE="JavaScript">writeSource("js15");</SCRIPT>
Start: <SELECT NAME="date15_month">
	<OPTION>
	<OPTION VALUE="Jan">January
	<OPTION VALUE="Feb">February
	<OPTION VALUE="Mar">March
	<OPTION VALUE="Apr">April
	<OPTION VALUE="May">May
	<OPTION VALUE="Jun">June
	<OPTION VALUE="Jul">July
	<OPTION VALUE="Aug">August
	<OPTION VALUE="Sep">September
	<OPTION VALUE="Oct">October
	<OPTION VALUE="Nov">November
	<OPTION VALUE="Dec">December
</SELECT>
 /
<SELECT NAME="date15_date">
	<OPTION>
	<OPTION VALUE="1">1
	<OPTION VALUE="2">2
	<OPTION VALUE="3">3
	<OPTION VALUE="4">4
	<OPTION VALUE="5">5
	<OPTION VALUE="6">6
	<OPTION VALUE="7">7
	<OPTION VALUE="8">8
	<OPTION VALUE="9">9
	<OPTION VALUE="10">10
	<OPTION VALUE="11">11
	<OPTION VALUE="12">12
	<OPTION VALUE="13">13
	<OPTION VALUE="14">14
	<OPTION VALUE="15">15
	<OPTION VALUE="16">16
	<OPTION VALUE="17">17
	<OPTION VALUE="18">18
	<OPTION VALUE="19">19
	<OPTION VALUE="20">20
	<OPTION VALUE="21">21
	<OPTION VALUE="22">22
	<OPTION VALUE="23">23
	<OPTION VALUE="24">24
	<OPTION VALUE="25">25
	<OPTION VALUE="26">26
	<OPTION VALUE="27">27
	<OPTION VALUE="28">28
	<OPTION VALUE="29">29
	<OPTION VALUE="30">30
	<OPTION VALUE="31">31
</SELECT>
<SELECT NAME="date15_year">
	<OPTION>
	<OPTION VALUE="2000">2000
	<OPTION VALUE="2001">2001
	<OPTION VALUE="2002">2002
	<OPTION VALUE="2003">2003
	<OPTION VALUE="2004">2004
	<OPTION VALUE="2005">2005
	<OPTION VALUE="2006">2006
</SELECT>
<A HREF="#" onClick="cal15.showCalendar('anchor15',getDateString(document.forms[0].date15_year,document.forms[0].date15_month,document.forms[0].date15_date)); return false;" TITLE="cal15.showCalendar('anchor15',getDateString(document.forms[0].date15_year,document.forms[0].date15_month,document.forms[0].date15_date)); return false;" NAME="anchor15" ID="anchor15">select</A>
&nbsp;&nbsp;&nbsp;
End: <SELECT NAME="date16_month">
	<OPTION>
	<OPTION VALUE="Jan">January
	<OPTION VALUE="Feb">February
	<OPTION VALUE="Mar">March
	<OPTION VALUE="Apr">April
	<OPTION VALUE="May">May
	<OPTION VALUE="Jun">June
	<OPTION VALUE="Jul">July
	<OPTION VALUE="Aug">August
	<OPTION VALUE="Sep">September
	<OPTION VALUE="Oct">October
	<OPTION VALUE="Nov">November
	<OPTION VALUE="Dec">December
</SELECT>
 /
<SELECT NAME="date16_date">
	<OPTION>
	<OPTION VALUE="1">1
	<OPTION VALUE="2">2
	<OPTION VALUE="3">3
	<OPTION VALUE="4">4
	<OPTION VALUE="5">5
	<OPTION VALUE="6">6
	<OPTION VALUE="7">7
	<OPTION VALUE="8">8
	<OPTION VALUE="9">9
	<OPTION VALUE="10">10
	<OPTION VALUE="11">11
	<OPTION VALUE="12">12
	<OPTION VALUE="13">13
	<OPTION VALUE="14">14
	<OPTION VALUE="15">15
	<OPTION VALUE="16">16
	<OPTION VALUE="17">17
	<OPTION VALUE="18">18
	<OPTION VALUE="19">19
	<OPTION VALUE="20">20
	<OPTION VALUE="21">21
	<OPTION VALUE="22">22
	<OPTION VALUE="23">23
	<OPTION VALUE="24">24
	<OPTION VALUE="25">25
	<OPTION VALUE="26">26
	<OPTION VALUE="27">27
	<OPTION VALUE="28">28
	<OPTION VALUE="29">29
	<OPTION VALUE="30">30
	<OPTION VALUE="31">31
</SELECT>
<SELECT NAME="date16_year">
	<OPTION>
	<OPTION VALUE="2000">2000
	<OPTION VALUE="2001">2001
	<OPTION VALUE="2002">2002
	<OPTION VALUE="2003">2003
	<OPTION VALUE="2004">2004
	<OPTION VALUE="2005">2005
	<OPTION VALUE="2006">2006
</SELECT>
<A HREF="#" onClick="var d=getDateString(document.forms[0].date16_year,document.forms[0].date16_month,document.forms[0].date16_date); cal16.showCalendar('anchor16',(d==null)?getDateString(document.forms[0].date15_year,document.forms[0].date15_month,document.forms[0].date15_date):d); return false;" TITLE="var d=getDateString(document.forms[0].date16_year,document.forms[0].date16_month,document.forms[0].date16_date); cal16.showCalendar('anchor16',(d==null)?getDateString(document.forms[0].date15_year,document.forms[0].date15_month,document.forms[0].date15_date):d); return false;" NAME="anchor16" ID="anchor16">select</A>
 
<!-- ================================================================================== --><HR>
 
</FORM>
<DIV ID="testdiv1" STYLE="position:absolute;visibility:hidden;background-color:white;layer-background-color:white;"></DIV>
 
</TD><TD VALIGN="TOP">
<script type="text/javascript"> 
google_ad_client = "pub-9155030588311591";
google_ad_width = 120;
google_ad_height = 600;
google_ad_format = "120x600_as";
google_color_border = "006666";
google_color_bg = "FFFFFF";
google_color_link = "006666";
google_color_url = "006666";
google_color_text = "000000";
</script>
<script type="text/javascript"
  src="http://pagead2.googlesyndication.com/pagead/show_ads.js"> 
</script>
</TD></TR></TABLE>
 
</BODY>
</HTML>

