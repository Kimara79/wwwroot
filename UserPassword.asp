<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<!--#include file="cokk.asp"-->
<%menu1_name1="修改密码"%>
<title><%=menu1_name1%> - <%=site_title%></title>
</head>

<!--#include file="user_power.asp"-->
<!--#include file="items/inc/md5.asp"-->

<% 
'----------------------------------------------------------------------提交操作判断-修改
if request.QueryString("tj")="Modify" then
	'空判断
	if request.Form("a_password")="" then 
		Response.Write("<script language=javascript>alert('请输入原密码！');history.back();</script>")
		Response.End
	end if
	if request.Form("password")="" then 
		Response.Write("<script language=javascript>alert('请输入新密码！');history.back();</script>")
		Response.End
	end if
	if request.Form("re_password")="" then 
		Response.Write("<script language=javascript>alert('请输入确认新密码！');history.back();</script>")
		Response.End
	end if
	
	'出错判断
	sql="select count(id) from users where id="&Request.Cookies("userqt")("id")&" and password='"&md5(request.Form("a_password"))&"' "
	rsnum=int(conn.execute(sql).getstring)
	if rsnum=0 then 
		Response.Write("<script language=javascript>alert('原密码错误！');history.back();</script>")
		Response.End
	end if
	if request.Form("password")<>request.Form("re_password") then 
		Response.Write("<script language=javascript>alert('新密码与确认新密码不符！');history.back();</script>")
		Response.End
	end if
	
	'更新数据
	sql="update users set password='"&md5(request.Form("password"))&"' where id="&Request.Cookies("userqt")("id")&" "
	conn.execute sql,1,1

	'返回
	response.write "<script>alert('修改成功！');location='?';</script>"
	response.end
end if
%>

<body>
<!--#include file="top.asp"-->

<center>
<table align="center" border="0" cellspacing="0" cellpadding="0" class="bodyw">
  <tr>
    <td valign="top" class="bodyw_left" align="center"><!--#include file="user_left.asp"--></td>
	<td width="10"></td>
    <td align="center" valign="top">
	<div class="bgb_01"><%=menu1_name1%></div>
	<div class="bgb_02" align="center">
		<div style="width:94%;">
		<div style="height:10px;"></div>
		<br />
		<form id="form1" name="form1" method="post" action="?tj=Modify">
		<table border="0" cellspacing="0" cellpadding="5">
		  <tr>
			<td align="center">原密码：</td>
			<td align="left"><input type="password" name="a_password" id="a_password" class="input2" maxlength="20" /></td>
		  </tr>
		  <tr>
			<td align="center">新密码：</td>
			<td align="left"><input type="password" name="password" id="password" class="input2" maxlength="20" /></td>
		  </tr>
		  <tr>
			<td align="center">确认新密码：</td>
			<td align="left"><input type="password" name="re_password" id="re_password" class="input2" maxlength="20" /></td>
		  </tr>
		  <tr>
			<td colspan="2" align="center" style="padding-top:10px;">
			  <input type="submit" name="Submit" value="提交" class="bt1" />&nbsp;&nbsp;
			  <input type="reset" name="Submit" value="重置" class="bt1" />
			  </td>
		  </tr>
		</table>
		</form>
		<br /><br />
		</div>
	</div>
	<div class="bgb_03"></div>
	</td>
  </tr>
</table>
</center>

<!--#include file="foot.asp"-->
</body>
</html>