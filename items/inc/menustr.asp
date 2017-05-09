<%
'like时
function menustr(fString)
	if fString<>"" then
		menustr="_"&fString&"_"
	end if
end function

'表单提交时处理
function menustr1(fString)
	temp=trim(""&fString)
	if temp<>"" then
		temp=replace(temp," ","")
		temp=replace(temp,",","")
		temp=replace(temp,"__","_")
		menustr1=temp
	end if
end function

'列表页时处理
function menustr2(fString)
	temp=trim(""&fString)
	if temp<>"" then
		temp=replace(temp,"_",",")
		temp=left(temp,len(temp)-1)
		temp=right(temp,len(temp)-1)
		menustr2=temp
	end if
end function 
%>