<%
p = Request("p")
s = Request("s")

if p <> "" then 

	pos = left(p,1)
	v = right(p,1)
	temp = session("LED")
	l = left(temp,pos)
	r = mid(temp,pos+2,10)
	session("LED") = l&v&r
	response.write(p)
end if


%>