<%
'年月日时分 (月日时分 2位)
function formatymd2(fString)
	if not isnull(fString) then
		fString1=year(fString)&Right("0"&Month(fString),2)&Right("0"&Day(fString),2)
		fString2=Right("0"&hour(fString),2)&Right("0"&minute(fString),2)
		formatymd2 = fString1&fString2
	end if
end function

'年-月-日 时:分 (月日时分 2位)
function formatymd2a(fString)
	if not isnull(fString) then
		fString1=year(fString)&"-"&Right("0"&Month(fString),2)&"-"&Right("0"&Day(fString),2)
		fString2=Right("0"&hour(fString),2)&":"&Right("0"&minute(fString),2)
		formatymd2a = fString1&" "&fString2
	end if
end function 

'年-月-日 时:分 (中文 时分2位)
function formatymd2b(fString)
	if not isnull(fString) then
		fString1=year(fString)&"年"&Month(fString)&"月"&Day(fString)&"日"
		fString2=Right("0"&hour(fString),2)&":"&Right("0"&minute(fString),2)
		formatymd2b = fString1&" "&fString2
	end if
end function 

'年-月-日 时:分 (中文 时分2位 无空格)
function formatymd2c(fString)
	if not isnull(fString) then
		fString1=year(fString)&"年"&Month(fString)&"月"&Day(fString)&"日"
		fString2=Right("0"&hour(fString),2)&":"&Right("0"&minute(fString),2)
		formatymd2c = fString1&fString2
	end if
end function 
%> 