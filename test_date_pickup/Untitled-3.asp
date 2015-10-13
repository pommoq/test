<%@LANGUAGE="VBSCRIPT" CODEPAGE="874"%>
<HTML>
<HEAD>
	<TITLE>JavaScript Toolbox - Calendar Popup To Select Date</TITLE>
<SCRIPT LANGUAGE="JavaScript" SRC="CalendarPopup.js"></SCRIPT>
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

 

 
<TABLE WIDTH="100%" BORDER="0"><TR><TD WIDTH="100%" ALIGN="LEFT" VALIGN="TOP">
<FORM name="f" id="f">
 

 
Default calendar using the DIV-style display.<BR>
<SCRIPT LANGUAGE="JavaScript" ID="jscal1x"> 
var cal1x = new CalendarPopup("testdiv1");
cal1x.showYearNavigation(); 
cal1x.showNavigationDropdowns();
cal1x.setThaiYear(543);

function convert_mmdd2ddmm(obj){
	obj_display = document.getElementsById(obj.id+'_display')
	if(obj_display){
		obj_display.value = obj.value
	}
}

</SCRIPT>
<!-- The next line prints out the source in this example page. It should not be included when you actually use the calendar popup code -->

<INPUT style="display:; background-color:#CCCCCC;" TYPE="text" NAME="date1x"  id="date1x" VALUE="" SIZE=25 >
<INPUT TYPE="text" NAME="date1x_display"  id="date1x_display" VALUE="" SIZE=25>
<A HREF="#" onClick="cal1x.select(document.forms[0].date1x,'anchor1x','MM/dd/yyyy');  return false;"  NAME="anchor1x" ID="anchor1x">select</A>
<SCRIPT LANGUAGE="JavaScript">
function date1x_setdate(y,m,d) {
	document.forms[0].date1x.value= m + "/" + d + "/" + y;
	document.forms[0].date1x_display.value= d + "/" + m + "/" + y;
	}
</SCRIPT>
<!-- ================================================================================== --><HR>
<INPUT style="display:;background-color:#CCCCCC;" TYPE="text" NAME="date2x"  id="date2x" VALUE="" SIZE=25 >
<INPUT TYPE="text" NAME="date2x_display"  id="date2x_display" VALUE="" SIZE=25>
<A HREF="#" onClick="cal1x.setReturnFunction('date2x_setdate'); cal1x.select(document.forms[0].date2x,'anchor2x','MM/dd/yyyy');  return false;"  NAME="anchor2x" ID="anchor2x">select</A>
<SCRIPT LANGUAGE="JavaScript">
function date2x_setdate(y,m,d) {
	document.forms[0].date2x.value= m + "/" + d + "/" + y;
	document.forms[0].date2x_display.value= d + "/" + m + "/" + y;
	}
</SCRIPT>
<!-- ================================================================================== --><HR>
<INPUT style='display:; background-color:#CCCCCC;' TYPE='text' NAME='S1_ReqDate'  id='S1_ReqDate'  SIZE=25 value="'1/1/1900'" >
<INPUT readonly='true' style='width:130px' TYPE='text' NAME='S1_ReqDate_display'  id='S1_ReqDate_display' value="1/1/2443">
<A HREF='#' onClick="cal1x.setReturnFunction('S1_ReqDate_setdate'); cal1x.select(document.form_trans.S1_ReqDate,'S1_ReqDate_anchor','MM/dd/yyyy');  return false;"  NAME='S1_ReqDate_anchor' ID='S1_ReqDate_anchor'><img  align='absmiddle'  border='0' src='/thaiauto/tainet/db/ITISO/../pic/cal_icon.gif' width='16' height='16'></A>
<A HREF='#' onClick="S1_ReqDate_setdate(2009,12,21  ); return false;"  NAME='S1_ReqDate_anchor' ID='S1_ReqDate_anchor'>วันนี้</A>
<SCRIPT LANGUAGE='JavaScript'> 
function S1_ReqDate_setdate(y,m,d) {
	if (y == 0) {
	    document.form_trans.S1_ReqDate.value= 'NULL';
	    document.form_trans.S1_ReqDate_display.value= '';
	} else {
	     var now = new  Date();
	     var t_hour = now.getHours();
	     var t_min = now.getMinutes();
	     var t_sec = now.getSeconds();
	document.form_trans.S1_ReqDate.value= '\'' + m + '/' + d + '/' + y  + ' ' + t_hour + ':' + t_min + ':'  + t_sec + '\'' ;
	document.form_trans.S1_ReqDate_display.value=   d + '/' + m + '/' + (parseInt(y) + cal1x.thaiYear)  ;
	}
	}
</SCRIPT>


 
</TD></TR></TABLE>
 
</BODY>
</HTML>

<DIV ID="testdiv1" STYLE="position:absolute;visibility:hidden;background-color:white;layer-background-color:white;"></DIV>