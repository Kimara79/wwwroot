<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="会员管理" %>
<% menunamel="会员" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="js/fbjs.asp"-->

<%
'默认密码
sql="select info from system where id=14"
rsstr=trim(conn.execute(sql).getstring)
password=replace(rsstr,chr(13),"")
%>

<%
'--------------------------------------------------------出错判断
if request("tj")="Add" or request("tj")="Modify" then
	'空判断
	if request.Form("card")="" then 
		Response.Write("<script language=javascript>alert('卡号不能为空！');history.back();</script>")
		Response.End
	end if
	if request.Form("name1")="" then 
		Response.Write("<script language=javascript>alert('姓名不能为空！');history.back();</script>")
		Response.End
	end if
	if request.Form("score")="" then 
		Response.Write("<script language=javascript>alert('积分不能为空！');history.back();</script>")
		Response.End
	end if

	'卡号重复判断
	if request("tj")="Add" then
		sql="select count(id) from users where card="&request.Form("card")&" "
	end if
	if request("tj")="Modify" then
		sql="select count(id) from users where card="&request.Form("card")&" and id<>"&request("id")&""
	end if
	rsnum=int(conn.execute(sql).getstring)
	if rsnum>0 then
		Response.Write("<script language=javascript>alert('已存在此卡号！');history.back()</script>")
		Response.End
	end if

	'手机号重复判断
	if request("tj")="Add" then
		sql="select count(id) from users where tel='"&request.Form("tel")&"' "
	end if
	if request("tj")="Modify" then
		sql="select count(id) from users where tel='"&request.Form("tel")&"' and id<>"&request("id")&""
	end if
	rsnum=int(conn.execute(sql).getstring)
	if rsnum>0 then
		Response.Write("<script language=javascript>alert('已存在此手机号！');history.back()</script>")
		Response.End
	end if
	
	'密码判断
	if request("tj")="Add" then
		if request.Form("password")="" then 
			Response.Write("<script language=javascript>alert('密码不能为空！');history.back()</script>")
			Response.End
		end if
		if request.Form("re_password")="" then 
			Response.Write("<script language=javascript>alert('确认密码不能为空！');history.back()</script>")
			Response.End
		end if
		if request.Form("password")<>request.Form("re_password") then 
			Response.Write("<script language=javascript>alert('确认密码错误！');history.back()</script>")
			Response.End
		end if
	end if 
end if

'--------------------------------------------------------提交操作判断-添加
if request("tj")="Add" then
	rs.Open "users",conn,1,3
	rs.addnew
	On Error Resume Next
	For each obj in request.Form
		rs(obj)=request.Form(obj)
	Next
	if request.Form("password")<>"" then 
		rs("password")=md5(request.Form("password"))
	end if
	rs("addtime")=now()
	rs.update
	rs.close
	
	rsrz.open "book",conn,1,3                                   '添加日志              
	rsrz.addnew
	rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
	rsrz("caozuo")="添加"
	rsrz("menu")=menuname
	rsrz("duixian")=request.Form("name1")
	rsrz.update
	rsrz.close
	response.redirect "usersManage.asp"
	Response.End
end if
%>

<% '--------------------------------------------------------提交操作判断-修改
if request("tj")="Modify" then
	sql="select * from users where id="&request("id")
	rs.Open sql,conn,1,3
	On Error Resume Next
	For each obj in request.Form
		rs(obj)=request.Form(obj)
	Next
	if request.Form("password")<>"" then 
		rs("password")=md5(request.Form("password"))
	end if
	rs.update
	rs.close
	
	rsrz.open "book",conn,1,3                                   '添加日志              
	rsrz.addnew
	rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
	rsrz("caozuo")="修改"
	rsrz("menu")=menuname
	rsrz("duixian")=request.Form("name1")
	rsrz.update
	rsrz.close
	
	response.redirect "usersManage.asp"
	Response.End
end if
%>

<body>
<!--#include file="top.asp"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23" align="center" background="images/am_mn2.gif" bgcolor="cccccc"><strong><%=menuname%></strong></td>
  </tr>
</table>
<div align="center">
<br />
<!---------------------------------------------添加界面------------------------------------------->
<%if  request("cz")="Add" then %>
<strong>添加<%=menunamel%></strong><br><br>
<form action="?tj=Add" method="post" name="form2" id="form2">
<table border="1" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr>
    <td width="100" align="center" bgcolor="#F0F0F0"><strong>卡号：</strong></td>
    <td align="left"><input name="card" type="text" id="card" class="input1" onKeyUp="<%=fbjs_rnum1%>" />
      <span class="red">*</span></td>
  </tr>
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong>姓名：</strong></td>
    <td align="left"><input name="name1" type="text" id="name1" class="input1" />
      <span class="red">*</span></td>
  </tr>
  <tr>
    <td width="100" align="center" bgcolor="#F0F0F0"><strong>手机号：</strong></td>
    <td align="left"><input name="tel" type="text" id="tel" class="input1" onKeyUp="<%=fbjs_rnum1%>" />
      <span class="red">*</span></td>
  </tr>
  <tr>
    <td height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">密码：</span></td>
    <td align="left"><input name="password" type="password" id="password" class="input1" value="<%=password%>" />
        <span class="red">*</span> (默认为：<%=password%>)</td>
  </tr>
  <tr>
    <td height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">确认密码：</span></td>
    <td align="left"><input name="re_password" type="password" id="re_password" class="input1" value="<%=password%>" />
        <span class="red">*</span> (默认为：<%=password%>)</td>
  </tr>
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong>积分：</strong></td>
    <td align="left"><input name="score" type="text" id="score" value="0" class="input1" onKeyUp="<%=fbjs_rnum1%>" />
      <span class="red">*</span></td>
  </tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong>备注：</strong></td>
	  <td align="left"><input name="info" type="text" id="info" class="input1" /></td>
	</tr>
</table>
<!--#include file="inc/button.asp"-->
</form>
      <%end if%>
<!---------------------------------------------修改界面------------------------------------------->
<% if  request("cz")="Modify" then %>
<%
 sql="select * from users where id="&request("id")
 rs.Open sql,conn,1,3
%>
<strong>修改<%=menunamel%></strong><br /><br />
<form action="?id=<%=rs("id")%>&amp;tj=Modify" method="post" name="form3" id="form3">
  <table border="1" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
    <tr>
      <td width="100" align="center" bgcolor="#F0F0F0"><strong>卡号：</strong></td>
      <td align="left"><input name="card" type="text" id="card" value="<%=rs("card")%>" class="input1" onKeyUp="<%=fbjs_rnum1%>" />
        <span class="red">*</span></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>姓名：</strong></td>
      <td align="left"><input name="name1" type="text" id="name1" value="<%=rs("name1")%>" class="input1" />
        <span class="red">*</span></td>
    </tr>
  <tr>
    <td width="100" align="center" bgcolor="#F0F0F0"><strong>手机号：</strong></td>
    <td align="left"><input name="tel" type="text" id="tel" class="input1" value="<%=rs("tel")%>" onKeyUp="<%=fbjs_rnum1%>" />
      <span class="red">*</span></td>
  </tr>
    <tr>
      <td height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">修改密码：</span></td>
      <td align="left"><input name="password" type="password" id="password" class="input1" /></td>
    </tr>
    <tr>
      <td height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">确认密码：</span></td>
      <td align="left"><input name="re_password" type="password" id="re_password" class="input1" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>积分：</strong></td>
      <td align="left"><input name="score" type="text" id="score" value="<%=rs("score")%>" class="input1" onKeyUp="<%=fbjs_rnum1%>" />
        <span class="red">*</span></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>备注：</strong></td>
      <td align="left"><input name="info" type="text" id="info" value="<%=rs("info")%>" class="input1" /></td>
    </tr>
  </table>
  <!--#include file="inc/button.asp"-->
</form>
<%rs.close%>
<%end if%>
</div>
<!--#include file="foot.asp"-->
</body>
</html>