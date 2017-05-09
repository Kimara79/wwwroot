<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title>帐号管理</title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->

<%  
if request("cz")="del" then    '----------------------------------------------------删除判断	
  sql="delete * from admin where id="&request("id")               '删除数据
  conn.execute sql,1,1 
   
  		rsrz.open "book",conn,1,3                                  '添加日志              
		rsrz.addnew
		  rsrz("zhanghao")=session("admin")&"("&session("truename")&")"
		  rsrz("caozuo")="删除"
		  rsrz("menu")="帐号管理"  
		  rsrz("duixian")="帐号："&request("admin")&"("&request("truename")&")"
		rsrz.update
		rsrz.close
		
  response.redirect "?"
  response.end
end if
%>
<%
if request("tj")="search" then
session("admin_search")="search"
session("admin_search_type")=request.form("search_type")
session("admin_search_key")=request.form("search_key")
end if
if request("xs")="all" then
session("admin_search")=""
end if
%>

<body>
<!--#include file="top.asp"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23" align="center" background="images/am_mn2.gif" bgcolor="cccccc"><strong>帐号管理</strong></td>
  </tr>
</table>
<div align="center">
<br>
<form id="form1" name="form1" method="post" action="AdminManage.asp?tj=search">
  <table width="450" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td class="STYLE1">搜索：</td>
      <td><select name="search_type" id="search_type">
        <option value="step">级别</option>
        <option value="admin">用户名</option>
        <option value="truename">真实姓名</option>
        <option value="info">备注</option>
      </select>
      </td>
      <td><input name="search_key" type="text" id="search_key" /></td>
      <td width="60" align="right"><input type="submit" name="Submit22" value=" 搜索 " /></td>
      <td width="80" align="right"><input type="button" name="Submit2" onclick="location.href='AdminAM.asp?cz=Add'" value=" 添加 " /></td>
    </tr>
  </table>
</form>
<br>

<%
if session("admin_search")<>"" then
	sql="select * from admin where "&session("admin_search_type")&" like '%"&session("admin_search_key")&"%' order by step,admin"
else
	sql="select * from admin order by step,admin"
end if
rs.open sql,conn,1,1
%>
        <%if rs.bof then%>
        <p align="center">暂无信息</p>
      <%else %>
        <table border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" width="90%">
          <tr class="STYLE1">
            <td width="10%" height="20" align="center" bgcolor="#F0F0F0">级别</td>
            <td width="20%" align="center" bgcolor="#F0F0F0">用户名</td>
            <td width="20%" align="center" bgcolor="#F0F0F0">真实姓名</td>
            <td align="center" bgcolor="#F0F0F0">备注</td>
            <td width="20%" align="center" bgcolor="#F0F0F0">操作</td>
          </tr>
          <% Do While Not rs.Eof %>
		  <%if rs("admin")<>"sosovipp" then%>
          <tr onmouseover="bgColor='#D2F5FF';" onmouseout="bgColor='#FFFFFF';">
            <td height="15" align="center"><%=rs("step")%></td>
            <td align="center"><%=rs("admin")%></td>
            <td align="center"><%=rs("truename")%></td>
            <td align="left"><%=rs("info")%>&nbsp;</td>
            <td align="center"><a href="AdminManage.asp?id=<%=rs("id")%>&amp;cz=del&amp;admin=<%=rs("admin")%>&amp;truename=<%=rs("truename")%>" onclick="{if(confirm('确定要删除么?')){this.document.album.submit();return true;}return false;}">删除</a>&nbsp;&nbsp;&nbsp;<a href="AdminAM.asp?id=<%=rs("id")%>&amp;cz=Modify">修改</a></td>
          </tr>
		  <%end if%>
          <%
	i=i+1
	rs.MoveNext
	Loop  %>
        </table>
<%
end if
rs.close
%>
        <br>
</div>
<!--#include file="foot.asp"-->
</body>
</html>