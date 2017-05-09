<%
'年-月-日 (月日 2位)
function formatymd3(fString)
	if not isnull(fString) then
		formatymd3=year(fString)&"-"&Right("0"&Month(fString),2)&"-"&Right("0"&Day(fString),2)
	end if
end function 

'年月日 (月日 2位)
function formatymd3a(fString)
	if not isnull(fString) then
		formatymd3a=year(fString)&Right("0"&Month(fString),2)&Right("0"&Day(fString),2)
	end if
end function

'年-月-日
function formatymd3b(fString)
	if not isnull(fString) then
		formatymd3b=year(fString)&"-"&Month(fString)&"-"&Day(fString)
	end if
end function 

'年-月-日 (中文)
function formatymd3c(fString)
	if not isnull(fString) then
		formatymd3c=year(fString)&"年"&Month(fString)&"月"&Day(fString)&"日"
	end if
end function 
%> 