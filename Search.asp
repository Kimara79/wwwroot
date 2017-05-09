<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<!--#include file="cokk.asp"-->
<%menu1_name1="站内搜索"%>
<title><%=menu1_name1%> - <%=site_title%></title>
</head>

<!--#include file="inc/formatymd.asp"-->

<%
if request.QueryString("keyword")<>"" then
	keyword=request.QueryString("keyword")
	keyword=trim(keyword)
	keyword=replace(keyword,"'","")
else
	response.redirect "Index1.asp"
	response.end
end if
%>

<body>
<!--#include file="top.asp"-->

<table align="center" border="0" cellspacing="0" cellpadding="0" class="bodyw">
  <tr>
    <td align="center" valign="top">
	<div class="bga_01"><%=menu1_name1%></div>
	<div class="bga_02" align="center">
		<div style="width:98%;">
			<div style="height:10px;"></div>
			
			<div style="width:90%;">
				<div align="left">关键字：<span style="color:#FF0000;"><%=keyword%></span></div>
				
				<br /><div align="left" class="border_b" style="font-weight:bold;line-height:40px;">符合条件的文章：</div>
				<%
				sql="select * from article where name1 like '%"&keyword&"%' order by id desc"
				rs.open sql,conn,1,1
				if rs.eof then
					response.Write "<br><br>暂无信息<br><br><br>"
				else
				%>
				<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<%Do while not rs.Eof%>
					<tr height="32">
					  <td align="left" class="border_b">
					  <img src="images/list3.gif" align="absmiddle" />&nbsp;&nbsp;&nbsp;<a href="ArticleShow.asp?id=<%=rs("id")%>"><%=rs("name1")%></a>
					  </td>
					  <td width="80" align="center" class="border_b"><%=formatymd(rs("lasttime"))%></td>
					</tr>
				<%
				rs.MoveNext
				Loop
				%>
				</table>
				<%
				end if
				rs.close
				%>
				
				<br /><div align="left" class="border_b" style="font-weight:bold;line-height:40px;">符合条件的图片：</div>
				<%
				sql="select * from Picture where name1 like '%"&keyword&"%' and menu1<>26 order by id desc"
				rs.open sql,conn,1,1
				if rs.eof then
					response.Write "<br><br>暂无信息<br><br><br>"
				else
				%>
				<div align="left">
				<table border="0" cellspacing="15" cellpadding="0">
				<%Do while not rs.Eof%>
					<tr>
					   <%for z=1 to 6%>
						<td align="center" valign="top">
						<%if not rs.eof then%>
						<a href="PictureShow.asp?id=<%=rs("id")%>"><img src="items/pic/s_picture/<%=rs("pic")%>" title="<%=rs("name1")%>" width="100" height="130" border="0"  class="border_pic" /></a>
						<%if rs("name1")<>"" then%><div align="center" style="margin-top:5px;"><%=num_write(rs("name1"),6)%></div><%end if%>
						<%
						rs.movenext
						end if
						%>
						</td>
					   <%next%>
					</tr>
				<%Loop%>
				</table>
				</div>
				<%
				end if
				rs.close
				%>
			</div>
			
		</div>
	</div>
	<div class="bga_03"></div>
	</td>
  </tr>
</table>

<!--#include file="foot.asp"-->
</body>
</html>