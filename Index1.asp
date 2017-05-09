<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<!--#include file="cokk.asp"-->
<title><%=site_title%></title>
</head>

<body>
<!--#include file="top.asp"-->

<table border="0" cellspacing="0" cellpadding="0" class="bodyw margin1">
  <tr>
    <td align="left" valign="top" class="bodyw_left">
	<div class="tba_1">Welcome to zhx2</div>
	<div class="tba_2">
		<div style="height:413px;">
			<div style="height:10px;"></div>
			<%
			for z=0 to 1
			a=2-z
			sql="select * from menu"&a&" where shezhi like '%首页左侧%' order by no1"
			rs1.open sql,conn,1,1
			Do while not rs1.Eof

			if rs1("url")<>"" then
				tempurl=rs1("url")
			else
				tempurl=menuurl(rs1("leixing"),a,rs1("id"))
			end if
			%>
			<div class="idxpic">
				<div class="idxpic_ct1"><a href="<%=tempurl%>"><img src="items/pic/picture/<%=rs1("pic")%>" /></a></div>
				<div class="idxpic_ct2"><h3><a href="<%=tempurl%>" class="a_black"><%=rs1("name1")%></a></h3></div>
			</div>
			<div class="pic_line"><img src="images/pic_line.gif" /></div>
			<%
			rs1.MoveNext
			Loop
			rs1.close
			next
			%>
			<div style="height:3px;"></div>
			<div align="center"><!--#include file="ic_80w25h.asp"--></div>
		</div>
	</div>
	<div class="tba_3"></div>
	</td>
    <td align="center" valign="top"><!--#include file="index_adv.asp"--></td>
    <td align="right" valign="top" class="bodyw_left">
		<!--#include file="index_score.asp"-->

		<div class="tba_1 margin1">会员登录</div>
		<div class="tba_2" style="height:92px;">
			<div align="center" style="padding-top:9px;" class="color_hs">
				<%if request.Cookies("userqt")("tel")<>"" then%>
					欢迎您！<br />
					<%=request.Cookies("userqt")("name1")%> (<%=request.Cookies("userqt")("tel")%>)
					<div align="center" style="margin-top:10px;">
					<a href="UserIndex.asp" class="a_hs">进入会员中心</a>&nbsp;&nbsp;&nbsp;&nbsp;
					<a href="user_login_tj.asp?tj=out" class="a_hs">退出</a>
					</div>
				<%else%>
					<%
					fontclass="color_hs"
					cellpadding="3"
					input="input0"
					bt="bt2"
					%>
					<!--#include file="user_login.asp"-->
				<%end if%>
			</div>
		</div>
		<div class="tba_3"></div>
	</td>
  </tr>
</table>

<!--#include file="foot.asp"-->
</body>
</html>