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

function num_write2(fString)
	if len(fString)>=2 then
		num_write2=left(fString,1)&"****"&right(fString,1)
	end if
end function

function num_write3(fString)
	if len(fString)>=1 then
		num_write3=left(fString,1)&"**"
	end if
end function
%>