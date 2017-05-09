<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<!--#include file="cokk.asp"-->
<%menu1_name1="会员登录"%>
<title><%=menu1_name1%> - <%=site_title%></title>
</head>

<%
if request.Cookies("userqt")("tel")<>"" then
	response.redirect "UserIndex.asp"
	response.end
end if
%>

<body>
<!--#include file="top.asp"-->

<center>
<table align="center" border="0" cellspacing="0" cellpadding="0" class="bodyw">
  <tr>
    <td align="center" valign="top">
	<div class="bga_01"><%=menu1_name1%></div>
	<div class="bga_02" align="center">
		<div style="width:94%;">
		<div style="height:10px;"></div>
		<br />
		<%
		fontclass=""
		cellpadding="8"
		input="input2"
		bt="bt"
		%>
		<!--#include file="user_login.asp"-->
		<br /><br />
		</div>
	</div>
	<div class="bga_03"></div>
	</td>
  </tr>
</table>
</center>

<!--#include file="foot.asp"-->
</body>
</html>