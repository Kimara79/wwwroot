<%
'年-月-日 时:分:秒 (月日时分秒 2位)
function formatymd(fString)
	if not isnull(fString) then
		fString1=year(fString)&"-"&Right("0"&Month(fString),2)&"-"&Right("0"&Day(fString),2)
		fString2=Right("0"&hour(fString),2)&":"&Right("0"&minute(fString),2)&":"&Right("0"&second(fString),2)
		formatymd = fString1&" "&fString2
	end if
end function 

'年月日时分秒 (月日时分秒 2位)
function formatymda(fString)
	if not isnull(fString) then
		fString1=year(fString)&Right("0"&Month(fString),2)&Right("0"&Day(fString),2)
		fString2=Right("0"&hour(fString),2)&Right("0"&minute(fString),2)&Right("0"&second(fString),2)
		formatymda = fString1&fString2
	end if
end function 
%>