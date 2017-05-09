<%
Function charnum(TheString,TheChar)
	charnum=0
	if inStr(TheString,TheChar) then
		for n=1 to Len(TheString)
		if Mid(TheString,n,Len(TheChar))=TheChar then 
		 charnum=charnum+1
		End if
		Next
	end if
End Function
%>