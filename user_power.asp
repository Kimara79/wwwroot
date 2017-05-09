<%
if Request.Cookies("userqt")("tel")="" then
	Response.Write "<script>alert('请先登录！');location.href='UserLogin.asp';</script>"
	Response.End
end if
%>