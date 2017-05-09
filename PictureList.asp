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
		sql="select * from picture where "&sqladd&" order by shezhi desc,no1"
		rs.open sql,conn,1,1
		if rs.eof then
			response.Write "<br><br>暂无信息<br><br><br><br><br>"
		else
		%>
			<%
			rs.pagesize=15  '设置每页显示数目
			currentpage=Clng(request("page"))
			if currentpage<1 then currentpage=1
			if currentpage>rs.pagecount then currentpage=rs.pagecount
			rs.absolutepage=currentpage
			%>
			<table border="0" cellspacing="10" cellpadding="0" width="100%">
			<%Do while not rs.eof%>
			  <tr>
			   <%for z=1 to 5%>
				<td align="center" valign="top">
				<%if not rs.eof then%>
				<a href="PictureShow.asp?id=<%=rs("id")%>"><img src="items/pic/s_picture/<%=rs("pic")%>" title="<%=rs("name1")%>" width="100" height="130" border="0"  class="border_pic" /></a>
				<%if rs("name1")<>"" then%><div align="center" style="margin-top:5px;"><%=rs("name1")%></div><%end if%>
				<%
				rs.movenext
				end if
				%>
				</td>
			   <%next%>
			  </tr>
			<%loop%>
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
</center>

<!--#include file="foot.asp"-->
</body>
</html>