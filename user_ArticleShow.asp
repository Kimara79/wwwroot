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

<body>
<!--#include file="top.asp"-->

<center>
<table align="center" border="0" cellspacing="0" cellpadding="0" class="bodyw">
  <tr>
    <td valign="top" class="bodyw_left" align="center"><!--#include file="user_left.asp"--></td>
	<td width="10"></td>
    <td align="center" valign="top">
	<div class="bgb_01"><%if menu2_name1<>"" then%><%=menu2_name1%><%else%><%=menu1_name1%><%end if%></div>
	<div class="bgb_02" align="center">
		<div style="width:94%;">
		<div style="height:30px;"></div>
	
		<%
		sql="select * from "&tb&" where id="&id&" "
		rs.open sql,conn,1,3
		%>
		
		<%if tb="article" then%>
			<%
			rs("click")=rs("click")+1
			rs.update
			%>
			<div class="title1" style="margin-bottom:5px;"><%=rs("name1")%></div>
			更新时间：<%=rs("lasttime")%>&nbsp;&nbsp;&nbsp;点击次数：<%=rs("click")%>
			<br><br style="line-height:10px;">
		<%else%>
			<br />
		<%end if%>
		
		<div align="left" style="width:98%;line-height:200%;">
			<%if menu2_name1="联系方式" then%>
				<!--#include file="ic_80w25h.asp"-->
			<%end if%>
			<%
			if rs("info")<>"" then
				if instr(rs("info"),"page-break-after") then
					info=rs("info")
					info=replace(info,"page-break-after: always;","xxxxx")
					info=replace(info,"page-break-after: always","xxxxx")
					info=split(info,"<div style=""xxxxx"">"&chr(13)&chr(10)&chr(9)&"<span style=""display: none;"">&nbsp;</span></div>")
					page=0
					if request.QueryString("page")<>"" then
						page=int(request.QueryString("page"))
					end if
					response.Write info(page)
					%>
					<div align="center">
					<%for z=0 to ubound(info)%>
					&nbsp;&nbsp;<a href="?<%=requesttrp("page")&"page="&z%>" <%if z=page then%>style="color:#FF0000;"<%end if%>>分页<%=z+1%></a>&nbsp;&nbsp;
					<%next%>
					</div>
					<%					
				else
					response.Write rs("info")
				end if
			end if
			%>
		</div>
		<br style="line-height:20px;">
		<%rs.close%>
		
		
		<%if tb="article" then%>
		<%
		sql="select * from article where "&sqladd&" order by shezhi desc,no1"
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
		<table width="90%" border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td align="left">上一篇：<%if p_id>0 then%><a href="?id=<%=p_id%>"><%=p_name1%></a><%else%>无<%end if%></td>
			<td align="right">下一篇：<%if n_id>0 then%><a href="?id=<%=n_id%>"><%=n_name1%></a><%else%>无<%end if%></td>
		  </tr>
		</table>
		<br style="line-height:20px;">
		<%end if%>

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