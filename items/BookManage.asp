<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title>日志管理</title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="js/maskEdit.js"-->

<%
'-------------------------------------------------------------------------清空判断
if request("cz")="del" then    	
  sql="delete * from book"
  conn.execute sql,1,1 
  
  		rsrz.open "book",conn,1,3                                   '添加日志              
		rsrz.addnew
		  rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
		  rsrz("caozuo")="清空"
		  rsrz("menu")="日志管理"
		rsrz.update
		rsrz.close
  
  conn.close 
  response.redirect ("BookManage.asp")
  response.end
end if
%>
<%
if request("order")<>"" then
session("book_order")=request("order")
end if
if session("book_order")="" then
session("book_order")="id desc"
end if

if request("tj")="search" then
session("book_search")="search"
session("book_search_type")=request.form("search_type")
session("book_search_key")=request.form("search_key")
end if
if request("xs")="all" then
session("book_search")=""
end if
%>

<body>
<!--#include file="top.asp"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23" align="center" background="images/am_mn2.gif" bgcolor="cccccc"><strong>日志管理</strong></td>
  </tr>
</table>
<div align="center">
<br>
<form id="form1" name="form1" method="post" action="BookManage.asp?tj=search">
  <table border="0" cellspacing="0" cellpadding="0">
	<tr>
	  <td class="STYLE1">搜索：</td>
	  <td><select name="search_type" id="search_type">
		  <option value="zhanghao">帐号</option>
		  <option value="caozuo">操作</option>
		  <option value="menu">栏目</option>
		  <option value="duixian">对象</option>
		  <option value="time1">时间</option>
	  </select>              </td>
	  <td>&nbsp;<input name="search_key" type="text" id="search_key" />&nbsp;&nbsp;</td>
	  <td align="right"><input type="submit" name="Submit22" value=" 搜索 " />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
      <%if session("step")=1 then%>
	  <td align="right"><input type="button" name="Submit" value="清空日志" onclick="if(confirm('确定要清空么?')){location.href='?cz=del';return true;}return false;"/></td>
      <%end if%>
	</tr>
  </table>
</form>
<br>

<%
if session("book_search")="search" then
sql="select * from book where "&session("book_search_type")&" like '%"&session("book_search_key")&"%' order by "&session("book_order")&""
else
sql="select * from book order by "&session("book_order")&""
end if
rs.open sql,conn,1,1
%>
          <%if rs.bof then%>
      <p align="center">暂无信息</p>
      <%else %>
        <table border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" width="90%">
          <tr  align="center" bgcolor="#F0F0F0"class="STYLE1">
            <td height="20"><a href="BookManage.asp?order=zhanghao">帐号</a></td>
            <td><a href="BookManage.asp?order=caozuo">操作</a></td>
            <td><a href="BookManage.asp?order=menu">栏目</a></td>
            <td><a href="BookManage.asp?order=duixian">对象</a></td>
            <td><a href="BookManage.asp?order=id desc">时间</a></td>
          </tr>
          <%
     	rs.pagesize=20  '设置每页显示数目
		currentpage=Clng(request("page"))
		if currentpage<1 then currentpage=1
		if currentpage>rs.pagecount then currentpage=rs.pagecount
		rs.absolutepage=currentpage
		
		 Do While Not rs.Eof %>
          <tr onmouseover="bgColor='#D2F5FF';" onmouseout="bgColor='#FFFFFF';">
            <td height="15" align="center"><%=rs("zhanghao")%></td>
            <td align="center"><%=rs("caozuo")%></td>
            <td align="center">&nbsp;<%=rs("menu")%>&nbsp;</td>
            <td align="left">&nbsp;<%=rs("duixian")%>&nbsp;</td>
            <td align="center"><%=rs("time1")%>&nbsp;</td>
          </tr>
          <%
i=i+1
rs.MoveNext
If i>=rs.pagesize Then Exit Do
Loop  %>
        </table>
<!--#include file="inc/page.asp"-->
  <%end if%>
</div>
<!--#include file="foot.asp"-->
</body>
</html>