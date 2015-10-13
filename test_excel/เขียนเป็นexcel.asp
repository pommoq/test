<%@LANGUAGE="VBSCRIPT" CODEPAGE="874"%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>Untitled Document</title>
<%
    Response.ContentType = "application/vnd.ms-excel"
    Response.AddHeader "Content-Disposition","attachment;filename=ExportToExcel-"&now()&".xls"
%>
</head>

<body>
<%=Request("excel")%><%=Request.Form("excel")%>
</body>
</html>
