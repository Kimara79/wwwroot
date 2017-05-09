<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title>在线客服</title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->

<%  
if request("cz")="del" then    '----------------------------------------------------删除判断	
  sql="delete * from cserver where id="&request("id")               '删除数据
  conn.execute sql,1,1 
   
  		rsrz.open "book",conn,1,3                                  '添加日志              
		rsrz.addnew
		  rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
		  rsrz("caozuo")="删除"
		  rsrz("menu")="在线客服"  
		  rsrz("duixian")=request("name1")
		rsrz.update
		rsrz.close
		
  conn.close
  response.redirect ("?")
end if
%>

<body>
<!--#include file="top.asp"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23" align="center" background="images/am_mn2.gif" bgcolor="cccccc"><strong>在线客服</strong></td>
  </tr>
</table>
<div align="center"><br />
<%
sql="select * from cserver order by sort1,no1"
rs.open sql,conn,1,1
%>
      <table border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" width="900">
        <tr class="STYLE1">
          <td width="100" align="center" bgcolor="#F0F0F0">类别</td>
          <td width="100" align="center" bgcolor="#F0F0F0">序号</td>
          <td width="300" align="center" bgcolor="#F0F0F0">帐号</td>
          <td align="center" bgcolor="#F0F0F0">说明</td>
          <td width="100" align="center" bgcolor="#F0F0F0">操作</td>
        </tr>
<% Do While Not rs.Eof %>
          <tr onmouseover="bgColor='#D2F5FF';" onmouseout="bgColor='#FFFFFF';">
            <td align="center"><%=rs("sort1")%></td>
            <td align="center"><%=rs("no1")%></td>
            <td align="left">&nbsp;<%=rs("name1")%></td>
            <td align="left">&nbsp;<%=rs("info")%></td>
            <td align="center">
			<a href="?id=<%=rs("id")%>&amp;cz=del&amp;name1=<%=rs("name1")%>" onclick="{if(confirm('确定要删除么?')){this.document.album.submit();return true;}return false;}">删除</a>&nbsp;&nbsp;
			<a href="CserverAM.asp?id=<%=rs("id")%>&amp;cz=Modify">修改</a>
			</td>
          </tr>
<%
rs.MoveNext
Loop
rs.close
%>
      </table>
      <br />
      <input type="button" name="Submit2" onclick="location.href='CserverAM.asp?cz=Add'" value=" 添 加 " />
      <br /><br />
</div>
<!--#include file="foot.asp"-->
</body>
</html>