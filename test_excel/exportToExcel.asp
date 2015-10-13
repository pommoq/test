<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
<%
    Response.ContentType = "application/vnd.ms-excel"
    Response.AddHeader "Content-Disposition","attachment;filename=ExportToExcel-"&now()&".xls"

if Request.Form("excel") <> "" then
	Session("ExportToExcel") =  Request.Form("excel")
end if	
%>
</head>

<body>
<%=Session("ExportToExcel")%>
</body>
</html>
