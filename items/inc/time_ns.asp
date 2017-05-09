<%
function time_ns(fString)
if not isnull(fString) then
 fString1=split(fString,",")
 for z=0 to ubound(fString1)
  time_ns=time_ns&"<span style='white-space:nowrap;'>"&fString1(z)
  if z<ubound(fString1) then
   time_ns=time_ns&"，"
  end if
  time_ns=time_ns&"</span>"
 next
end if
end function 
%>