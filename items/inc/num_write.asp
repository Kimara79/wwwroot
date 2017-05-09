<%
function num_write(fString,num)
	n=int(num)
	if fString<>"" then
	  if len(fString)>n then
	  fString=left(fString,n)&"..."
	  else
	  fString=fString
	  end if
	else
	  fString="&nbsp;"
	end if
	num_write=fString
end function
%>