<%@LANGUAGE="VBSCRIPT" CODEPAGE="874"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<title>Untitled Document</title>
</head>

<BODY BGCOLOR=#FFFFFF LINK="#00615F" VLINK="#00615F" ALINK="#00615F" onLoad="">
<form>
<TABLE BORDER=0 align="center" width="100%">
<TR>
	<TD width="48%">
	<SELECT NAME="list1" MULTIPLE SIZE=10 onDblClick="opt.transferRight()" style="width:100%">
		<OPTION VALUE="Matt">Matt</OPTION>
		<OPTION VALUE="Matt2">Matt2</OPTION>
		<OPTION VALUE="Bill">Bill</OPTION>
		<OPTION VALUE="Bob">Bob</OPTION>
		<OPTION VALUE="Jane">Jane</OPTION>
		<OPTION VALUE="Mary">Mary</OPTION>
		<OPTION VALUE="George">George</OPTION>
		<OPTION VALUE="Fred">Fred</OPTION>
		<OPTION VALUE="Ryan">Ryan</OPTION>
		<OPTION VALUE="Angela">Angela</OPTION>
		<OPTION VALUE="Jill">Jill</OPTION>		
	</SELECT>
	</TD>
	<TD VALIGN=MIDDLE ALIGN=CENTER width="4%">
		<INPUT TYPE="button" NAME="right" VALUE="&gt;" ONCLICK="opt.transferRight()" style="width:100%"><BR><BR>
		<INPUT TYPE="button" NAME="right" VALUE="&gt;&gt;" ONCLICK="opt.transferAllRight()" style="width:100%"><BR><BR>
		<INPUT TYPE="button" NAME="left" VALUE="&lt;" ONCLICK="opt.transferLeft()" style="width:100%"><BR><BR>
		<INPUT TYPE="button" NAME="left" VALUE="&lt;&lt;" ONCLICK="opt.transferAllLeft()" style="width:100%">
	</TD>
	<TD width="48%">
	<SELECT NAME="list2" MULTIPLE SIZE=10 onDblClick="opt.transferLeft()" style="width:100%">
	</SELECT>
	</TD>
</TR>
</TABLE>
Prefix:
<INPUT TYPE="text" NAME="prefix" VALUE="(" SIZE=2 MAXLENGTH=10 onChange="opt.setPrefix(this.value);opt.update()">
<br>
Delimiter:
<INPUT TYPE="text" NAME="delimiter" VALUE="," SIZE=2 MAXLENGTH=10 onChange="opt.setDelimiter(this.value);opt.update()">
<br>
Post:
<INPUT TYPE="text" NAME="post" VALUE=")" SIZE=2 MAXLENGTH=10 onChange="opt.setPostfix(this.value);opt.update()">
<br>
AutoSort:
<SELECT NAME="autosort" onChange="opt.setAutoSort(this.selectedIndex==0?true:false);opt.update()">
  <OPTION VALUE="Y">Yes
  <OPTION VALUE="N">No
</SELECT>
<br>
Options that Match this regular expression cannot be moved:
<INPUT TYPE="text" NAME="regex" SIZE="15" VALUE="^(Bill|Bob|Matt)$" onChange="opt.setStaticOptionRegex(this.value);opt.update()">
<br>
Removed from Left: <INPUT TYPE="text" NAME="removedLeft" VALUE="" SIZE=70><BR>
Removed from Right: <INPUT TYPE="text" NAME="removedRight" VALUE="" SIZE=70><BR>
Added to Left: <INPUT TYPE="text" NAME="addedLeft" VALUE="" SIZE=70><BR>
Added to Right: <INPUT TYPE="text" NAME="addedRight" VALUE="" SIZE=70><BR>
Left list contents: <INPUT TYPE="text" NAME="newLeft" VALUE="" SIZE=70><BR>
Right list contents: <INPUT TYPE="text" NAME="newRight" VALUE="" SIZE=70><BR>
</form>
</body>
</html>
<SCRIPT LANGUAGE="JavaScript" SRC="OptionTransfer.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
var opt = new OptionTransfer("list1","list2");
opt.setAutoSort(true);
opt.setPrefix(",");
opt.setPostfix("");
opt.setDelimiter("");
//opt.setStaticOptionRegex(""); //^(Bill|Bob|Matt)$
//opt.saveRemovedLeftOptions("removedLeft");
//opt.saveRemovedRightOptions("removedRight");
//opt.saveAddedLeftOptions("addedLeft");
opt.saveAddedRightOptions("addedRight");
//opt.saveNewLeftOptions("newLeft");
//opt.saveNewRightOptions("newRight");
opt.init(document.forms[0])
</SCRIPT>