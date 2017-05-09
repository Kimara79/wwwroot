<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<!--#include file="cokk.asp"-->
<%menu1_name1="中心首页"%>
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
		登录成功！<br />
		上次登录时间为：<%=request.Cookies("userqt")("logintime")%>
		<br /><br />
		<%
		rsstr=""
		sql="select info from menu1 where id=31"
		rsstr=trim(conn.execute(sql).getstring)
		response.Write rsstr
		%>
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