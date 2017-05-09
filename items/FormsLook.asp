<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="表单管理" %>
<% menunamel="表单" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->

<body>
<!--#include file="top.asp"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23" align="center" background="images/am_mn2.gif" bgcolor="cccccc"><strong><%=menuname%></strong></td>
  </tr>
</table>
<div align="center">
<br />
<%
sql="select * from forms where id="&request("id")
rs.Open sql,conn,1,1

sqladd=""
if rs("menu1")<>"" then
	sqladd=sqladd&" and menu1="&rs("menu1")&" "
end if
if rs("menu2")<>"" and rs("menu2")<>0 then
	sqladd=sqladd&" and menu2="&rs("menu2")&" "
end if
sql="select * from ziduan where id>0 "&sqladd&" order by no1"
rs1.open sql,conn,1,1
%>
<strong>查看<%=menunamel%></strong><br />
更新时间：<%=rs("lasttime")%>&nbsp;&nbsp;&nbsp;&nbsp;更新人：<%=rs("laster")%><br /><br />
<form action="?id=<%=rs("id")%>&amp;tj=Modify" method="post" name="form3" id="form3">
  <table border="1" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong>ID：</strong></td>
	  <td align="left">&nbsp;<%=rs("id")%>&nbsp;</td>
	</tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0">&nbsp;&nbsp;<strong>一级栏目：</strong>&nbsp;&nbsp;</td>
	  <td align="left">&nbsp;
		<%
		sql="select name1 from menu1 where id="&rs("menu1")&" "
		rsstr=trim(conn.execute(sql).getstring)
		rsstr=replace(rsstr,chr(13),"")
		response.Write rsstr
		%>&nbsp;
	  </td>
	</tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0">&nbsp;&nbsp;<strong>二级栏目：</strong>&nbsp;&nbsp;</td>
	  <td align="left">&nbsp;
		<%
		if rs("menu2")>0 then
			sql="select name1 from menu2 where id="&rs("menu2")&" "
			rsstr=trim(conn.execute(sql).getstring)
			rsstr=replace(rsstr,chr(13),"")
			response.Write rsstr
		end if
		%>&nbsp;
	  </td>
	</tr>
	<%Do while not rs1.Eof%>
	<tr>
	  <td align="center" bgcolor="#F0F0F0" valign="top">&nbsp;<strong><%=rs1("name1")%>：</strong>&nbsp;</td>
	  <td align="left" width="700"><div align="left">&nbsp;<%=rs("zd_"&rs1("id"))%>&nbsp;</div></td>
	</tr>
	<%
	rs1.MoveNext
	loop
	%>
  </table>
  <br>
  <input type="button" name="history11" value=" 返回 " onclick="history.back()" />
  <br><br>
</form>
<%
rs1.close
rs.close
%>
</div>
<!--#include file="foot.asp"-->
</body>
</html>