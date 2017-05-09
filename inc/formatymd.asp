<%
function formatymd(fString)
	if not isnull(fString) then
		formatymd=year(fString)&"-"&Right("0"&Month(fString),2)&"-"&Right("0"&Day(fString),2)
	end if
end function

function formatmd(fString)
	if not isnull(fString) then
		formatmd=Right("0"&Month(fString),2)&"-"&Right("0"&Day(fString),2)
	end if
end function 
%> 