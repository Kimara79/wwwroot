<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>登录验证</title>
</head>

<body>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/md5.asp"-->

<%
Response.AddHeader "P3P","CP=CAO PSA OUR"

if Request("tj")="in" then

user=replace(trim(request.form("admin")),"'","")
password=replace(trim(request.form("password")),"'","")

if user="" then 
 Response.Write("<script language=javascript>alert('请输入用户名！');window.location.href='"&request.servervariables("http_referer")&"';</script>")
 Response.End
end if
 
if password="" then 
 Response.Write("<script language=javascript>alert('请输入密码！');window.location.href='"&request.servervariables("http_referer")&"';</script>")
 Response.End
end if

if Request.Form("s")="" then 
 Response.Write("<script language=javascript>alert('请输入验证码！');window.location.href='"&request.servervariables("http_referer")&"';</script>")
 Response.End
end if

If Request.Form("s")<>Request.Form("s2") Then
response.Write "<script language=javascript>alert('请输入正确的验证码！');window.location.href='"&request.servervariables("http_referer")&"';</script>"
Response.End
end if

if instr(user,"%") or instr(user,"#") or instr(user,"?") or instr(user,"|") then
response.Write "<script language=javascript>alert('用户名含有非法字符！');window.location.href='"&request.servervariables("http_referer")&"';</script>"
Response.End
end if

if instr(password,"%") or instr(password,"#") or instr(password,"?") or instr(password,"|") then
response.Write "<script language=javascript>alert('密码含有非法字符！');window.location.href='"&request.servervariables("http_referer")&"';</script>"
Response.End
end if

password=md5(password)

sql="select * from admin where admin='"&user&"' and password='"&password&"'"
set rs=conn.execute(sql)
if rs.eof then
 response.Write "<script language=javascript>alert('用户名或密码错误！');window.location.href='"&request.servervariables("http_referer")&"';</script>"
else 
 Session("step")=rs("step")
 Session("admin")=rs("admin")
 Session("truename")=rs("truename")
 Session.Timeout=600
 rs.close

  		rsrz.open "book",conn,1,3                                   '添加日志              
		rsrz.addnew
		  rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
		  rsrz("caozuo")="登录"
		rsrz.update
		rsrz.close

 Response.Redirect("index1.asp")
 Response.End
end if

conn.close
set conn=nothing

end if
%>

<%
if Request("tj")="out" then
 if session("admin")<>"" and Session("truename")<>"" then
  		rsrz.open "book",conn,1,3                                   '添加日志              
		rsrz.addnew
		  rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
		  rsrz("caozuo")="退出"
		rsrz.update
		rsrz.close
 end if
 Response.Redirect("login.asp")
 Response.End
end if
%>
</body>
</html>