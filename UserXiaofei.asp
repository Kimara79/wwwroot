<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<!--#include file="cokk.asp"-->
<%menu1_name1="消费记录"%>
<title><%=menu1_name1%> - <%=site_title%></title>
</head>

<!--#include file="user_power.asp"-->
<!--#include file="items/inc/formatymd3.asp"-->

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
		
<%
sql="select * from xiaofei where usersid="&Request.Cookies("userqt")("id")&" "
rs.open sql,conn,1,1
if rs.bof then
%>
	<br /><br />暂无信息<br /><br /><br />
<%
else
rs.pagesize=10  '设置每页显示数目
currentpage=Clng(request("page"))
if currentpage<1 then currentpage=1
if currentpage>rs.pagecount then currentpage=rs.pagecount
rs.absolutepage=currentpage
%>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#D7A743">
		  <tr>
			<td align="center" class="color_white">消费日期</td>
			<td align="center" class="color_white">商品名称</td>
			<td align="center" class="color_white">商品单价</td>
			<td align="center" class="color_white">购买数量</td>
			<td align="center" class="color_white">合计金额</td>
			<td align="center" class="color_white">获得积分</td>
		  </tr>
<%Do while not rs.Eof%>
		  <tr>
			<td align="center" bgcolor="#FFFFFF" class="color_hs"><%=formatymd3b(rs("time1"))%></td>
			<td align="left" bgcolor="#FFFFFF" class="color_hs"><%=rs("name1")%></td>
			<td align="right" bgcolor="#FFFFFF" class="color_hs"><%=rs("price")%></td>
			<td align="center" bgcolor="#FFFFFF" class="color_hs"><%=rs("num")%></td>
			<td align="right" bgcolor="#FFFFFF" class="color_hs"><%=rs("priceall")%></td>
			<td align="right" bgcolor="#FFFFFF" class="color_hs"><%=rs("score")%></td>
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
	<div class="bgb_03"></div>
	</td>
  </tr>
</table>
</center>

<!--#include file="foot.asp"-->
</body>
</html>