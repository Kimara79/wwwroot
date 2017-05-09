<%
Function charpr(str,num)
a=len(str)
if a<=num then
 response.Write str
else
 response.Write left(str,num)&"..."
end if
End Function
%>