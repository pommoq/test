<%
if Session("LED") = "" then 
	Session("LED") = "01010"
end if
%>
<%=Session("LED")%>