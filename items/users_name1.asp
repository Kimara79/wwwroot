<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>会员姓名</title>
<style type="text/css">
<!--
body {
	margin: 0px;
}
body,td,th {
	font-family: 宋体;
	font-size: 12px;
	line-height:200%;
}
-->
</style>
</head>

<!--#include file="inc/cokk.asp"-->

<body>
<%
On Error Resume Next
if request.QueryString("users_card")<>"" then
	sql="select name1 from users where card="&request.QueryString("users_card")&" "
	rsstr=trim(conn.execute(sql).getstring)
	if rsstr<>"" then
		response.Write "（ "
		rsstr=replace(""&rsstr,chr(13),"")
		response.Write "姓名："&rsstr
		sql="select score from users where card="&request.QueryString("users_card")&" "
		rsstr=trim(conn.execute(sql).getstring)
		rsstr=replace(""&rsstr,chr(13),"")
		response.Write "&nbsp;&nbsp;积分："&rsstr
		sql="select id from users where card="&request.QueryString("users_card")&" "
		id=int(conn.execute(sql).getstring)
		sql="select count(id) from xiaofei where usersid="&id&" and score>0 "
		rsnum=int(conn.execute(sql).getstring)
		response.Write "&nbsp;&nbsp;记录：<span>"&rsnum&"</span>"
		response.Write "）"
	else
		response.Write "（查无此卡号）"
	end if
end if
%>
</body>
</html>