<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<!--#include file="cokk.asp"-->
<%tb="picture"%>
<!--#include file="title.asp"-->
<title><%=title1%> - <%=site_title%></title>
</head>

<!--#include file="js/pic.js"-->

<body>
<!--#include file="top.asp"-->

<center>
<table align="center" border="0" cellspacing="0" cellpadding="0" class="bodyw">
  <tr>
    <td align="center" valign="top">
	<div class="bga_01"><%if menu2_name1<>"" then%><%=menu2_name1%><%else%><%=menu1_name1%><%end if%></div>
	<div class="bga_02" align="center">
		<div style="width:94%;">
		<div style="height:10px;"></div>
	
<%
sql="select * from picture where id="&request("id")&" "
rs.open sql,conn,1,3
rs("click")=rs("click")+1
rs.update
%>
<div align="center" style="padding:10px;"><img src="items/pic/picture/<%=rs("pic")%>" class="border_pic" /></div>
<%if rs("name1")<>"" then%>
	<div class="title1"><%=rs("name1")%></div>
<%end if%>
<%if rs("info")<>"" then%>
	<br style="line-height:20px;">
	<div align="left" style="width:94%;line-height:200%;"><%=rs("info")%></div>
<%end if%>
<br style="line-height:20px;">
<%rs.close%>

<%
sql="select * from picture where "&sqladd&" order by shezhi desc,no1"
rs.Open sql,conn,1,1
Do while not rs.eof
	if int(rs("id"))=int(request("id")) then
		rs.moveprevious
		if rs.bof then
			p_id=0
			p_name1=""
		else
			p_id=rs("id")
			p_name1=rs("name1")
		end if
		rs.movenext
		rs.movenext
		if rs.eof then
			n_id=0
			n_name1=""
		else
			n_id=rs("id")
			n_name1=rs("name1")
		end if
		Exit Do
	end if
rs.movenext
loop
rs.close
%>
<table width="50%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="left">上一页：<%if p_id>0 then%><a href="?id=<%=p_id%>"><%=p_name1%></a><%else%>无<%end if%></td>
    <td align="right">下一页：<%if n_id>0 then%><a href="?id=<%=n_id%>"><%=n_name1%></a><%else%>无<%end if%></td>
  </tr>
</table>
<br style="line-height:20px;">
	
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