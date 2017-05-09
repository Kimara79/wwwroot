<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title>空间占用</title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="inc/mofei.asp"-->
<!--#include file="inc/flodersize.asp"-->

<%
'空间大小
sql="select info from system where id=22"
rsstr=trim(conn.execute(sql).getstring)
rsstr=replace(rsstr,chr(13),"")
spa=int(rsstr)
%>

<body>
<div align="center">
<!--#include file="top.asp"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
	  <td height="23" align="center" background="images/am_mn2.gif" bgcolor="cccccc"><strong>空间占用</strong></td>
	</tr>
</table>

<br><br>
<div align="center">
	<%
	spa2=flodersize("../")
	spa2=replace(spa2,"M","")
	width=400
	%>
	总共：<span style="color:#FF0000;font-weight:bold;"><%=spa%>M</span>，
	&nbsp;&nbsp;已用：<span style="color:#FF0000;font-weight:bold;"><%=spa2%>M</span>，
	&nbsp;&nbsp;剩余：<span style="color:#FF0000;font-weight:bold;"><%=spa-spa2%>M</span>
	
	<div style="background:url(images/space1.gif);height:10px; width:<%=width%>px;margin-top:15px;" align="left">
		<div style="background:url(images/space2.gif);height:10px;width:<%=int(spa2*width/spa)%>px;"></div>
	</div>
</div>
<br><br>

<!--#include file="foot.asp"-->
</div>
</body>
</html>