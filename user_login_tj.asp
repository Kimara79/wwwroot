<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title>会员登录</title>
</head>

<!--#include file="cokk.asp"-->
<!--#include file="items/inc/md5.asp"-->

<body>
<%
'登录
if Request("tj")="in" then
	tel=replace(trim(request.form("tel")),"'","")
	password=replace(trim(request.form("password")),"'","")
	
	'出错判断
	if tel="" then 
		Response.Write "<script language=javascript>alert('请输入手机号！');history.back();</script>"
		Response.End
	end if
	if password="" then 
		Response.Write "<script language=javascript>alert('请输入密码！');history.back();</script>"
		Response.End
	end if
	password=md5(password)
	sql="select count(id) from users where tel='"&tel&"' and password='"&password&"' "
	rsnum=int(conn.execute(sql).getstring)
	if rsnum=0 then 
		Response.Write("<script language=javascript>alert('卡号或密码错误！');history.back();</script>")
		Response.End
	end if
	
	sql="select * from users where tel='"&tel&"' and password='"&password&"' "
	rs.open sql,conn,1,3
	Response.Cookies("userqt")("id")=rs("id")
	Response.Cookies("userqt")("tel")=rs("tel")
	Response.Cookies("userqt")("name1")=rs("name1")
	Response.Cookies("userqt")("logintime")=""&rs("logintime")
'	if request.form("autologin")="yes" then
'		Response.Cookies("userqt").expires=Date()+365
'	end if
	rs("logintime")=now()
	rs.update
	rs.close
	conn.close
	set conn=nothing
	if request.Cookies("userqt")("logintime")<>"" then
		response.Write "<script>location.href='UserIndex.asp';</script>"
	else
		Response.Write("<script language=javascript>alert('您的当前的帐号密码为默认密码，为了安全考虑，请尽快修改密码！');location.href='UserPassword.asp';</script>")
	end if
	response.end
end if

'退出
if Request("tj")="out" then
	Response.Cookies("userqt")=""
	response.Write "<script>location.href='Index1.asp';</script>"
	response.end
end if
%>
</body>
</html>