<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>up</title>
</head>

<!--#include file="inc/power.asp"-->
<!--#include file="inc/cokk.asp"-->
<%
if request("id")<>"" then
	zd=split(request("zd"),",")
	for z=0 to ubound(zd)
		sql="update "&request("tb")&" set "&zd(z)&"='"&request(zd(z))&"' where id="&request("id")&""
		conn.execute sql,1,1
	next
end if
%>

<body>
</body>
</html>