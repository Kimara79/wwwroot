<%
'检查输入的Email是否正确
Function chkMail(strEmail)
	Dim objRegExp
	Set objRegExp = New RegExp
	objRegExp.Pattern = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
	chkMail= objRegExp.Test(strEmail)
End Function
%>