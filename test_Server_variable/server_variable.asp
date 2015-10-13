<%on Error Resume Next%>
<table width="70%" align="center" border="1">
<tr bgcolor="#FF9900"> 
<td width="25%"> 
<div align="center"><font size="2"><b><font face="Verdana, Arial, Helvetica, sans-serif">ServerVariable</font></b></font></div></td>
<td width="75%"> 
<div align="center"><font size="2"><b><font face="Verdana, Arial, Helvetica, sans-serif">Value</font></b></font></div></td>
</tr>
<%
For Each Item in Session.Contents %>
<tr> 
<td bgcolor=#FF9900 width="25%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%= "Session(""" & Item & """)"%></font></td>
<td width="75%"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=Session(Item)%></font>&nbsp;</td>
</tr>
<%
Next
%>
<%
For Each Item in Request.ServerVariables %>
<tr> 
<td bgcolor=#FF9900 width="25%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%= Item %></font></td>
<td width="75%"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=Request.ServerVariables(Item)%></font>&nbsp;</td>
</tr>
<%
Next
PATH_INFO =  Request.ServerVariables("PATH_INFO")
%>
<tr>
<td bgcolor=#FF9900 width="25%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">PATH_INFO =  Request.ServerVariables("PATH_INFO")<br />
Left(PATH_INFO ,instrRev(PATH_INFO, "/" ))</font></td><td> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%= Left(PATH_INFO ,instrRev(PATH_INFO, "/" )) %></font></td>
</tr>
<tr>
  <td bgcolor=#FF9900></td>
  <td>
<% 
    prot = "http" 
    https = lcase(request.ServerVariables("HTTPS")) 
    if https <> "off" then prot = "https" 
    domainname = Request.ServerVariables("SERVER_NAME") 
    filename = Request.ServerVariables("SCRIPT_NAME") 
	path = Left(filename ,instrRev(filename, "/" ))
	filename = mid(filename ,instrRev(filename, "/" ) +1, len(filename))
    querystring = Request.ServerVariables("QUERY_STRING") 
    response.write prot & "://" & domainname  & path & filename & "?" & querystring 
%>
</td>
</tr>

</table> 

