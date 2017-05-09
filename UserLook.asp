<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<!--#include file="cokk.asp"-->
<%menu1_name1="查看资料"%>
<title><%=menu1_name1%> - <%=site_title%></title>
</head>

<!--#include file="user_power.asp"-->

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
		<%
		sql="select * from users where id="&Request.Cookies("userqt")("id")&" "
		rs.open sql,conn,1,1
		sql="select count(id) from xiaofei where usersid="&Request.Cookies("userqt")("id")&" "
		rsnum=int(conn.execute(sql).getstring)
		%>
		<table width="400" border="0" cellpadding="5" cellspacing="1" bgcolor="#D7A743">
		  <tr>
			<td width="100" align="center" class="color_white">卡号：</td>
			<td align="left" bgcolor="#FFFFFF" class="color_hs">&nbsp;&nbsp;<%=rs("card")%></td>
		  </tr>
		  <tr>
			<td align="center" class="color_white">姓名：</td>
			<td align="left" bgcolor="#FFFFFF" class="color_hs">&nbsp;&nbsp;<%=rs("name1")%></td>
		  </tr>
		  <tr>
			<td align="center" class="color_white">手机号：</td>
			<td align="left" bgcolor="#FFFFFF" class="color_hs">&nbsp;&nbsp;<%=rs("tel")%></td>
		  </tr>
		  <tr>
			<td align="center" class="color_white">积分：</td>
			<td align="left" bgcolor="#FFFFFF" class="color_hs">&nbsp;&nbsp;<%=rs("score")%></td>
		  </tr>
		  <tr>
			<td align="center" class="color_white">消费记录：</td>
			<td align="left" bgcolor="#FFFFFF" class="color_hs">&nbsp;&nbsp;<a href="UserXiaofei.asp" class="a_hs"><%=rsnum%></a></td>
		  </tr>
		</table>
		<%rs.close%>
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