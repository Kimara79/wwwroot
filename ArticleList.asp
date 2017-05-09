<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<!--#include file="cokk.asp"-->
<%tb="article"%>
<!--#include file="title.asp"-->
<title><%=title1%> - <%=site_title%></title>
</head>

<!--#include file="inc/formatymd.asp"-->

<body>
<!--#include file="top.asp"-->

<table align="center" border="0" cellspacing="0" cellpadding="0" class="bodyw">
  <tr>
    <td align="center" valign="top">
	<div class="bga_01"><%if menu2_name1<>"" then%><%=menu2_name1%><%else%><%=menu1_name1%><%end if%></div>
	<div class="bga_02" align="center">
		<div style="width:94%;">
		<div style="height:10px;"></div>
		<%
		sql="select * from article where "&sqladd&" order by shezhi desc,no1"
		rs.open sql,conn,1,1
		if rs.eof then
			response.Write "<br><br>暂无信息<br><br><br><br><br>"
		else
		%>
		<table width="90%" border="0" cellpadding="0" cellspacing="0">
		<%
		rs.pagesize=10  '设置每页显示数目
		currentpage=Clng(request("page"))
		if currentpage<1 then currentpage=1
		if currentpage>rs.pagecount then currentpage=rs.pagecount
		rs.absolutepage=currentpage
		Do while not rs.Eof
		%>
			<tr height="32">
			  <td align="left" class="border_b">
			  <img src="images/list3.gif" align="absmiddle" />&nbsp;&nbsp;&nbsp;<a href="ArticleShow.asp?id=<%=rs("id")%>"><%=rs("name1")%></a>
			  </td>
			  <td width="80" align="center" class="border_b"><%=formatymd(rs("lasttime"))%></td>
			</tr>
		<%
		i=i+1
		rs.MoveNext
		If i>=rs.pagesize Then Exit Do
		Loop
		%>
		</table>
		<%danwei="条"%>
		<!--#include file="inc/page.asp"-->
		<%
		end if
		rs.close
		%>
		</div>
	</div>
	<div class="bga_03"></div>
	</td>
  </tr>
</table>

<!--#include file="foot.asp"-->
</body>
</html>