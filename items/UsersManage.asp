<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="会员管理" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="js/red.js"-->

<%
'删除操作
if request("cz")="del" then
	sql="delete * from xiaofei where usersid="&request("id")
	conn.execute sql,1,1 
	sql="delete * from jifen where usersid="&request("id")
	conn.execute sql,1,1 
	sql="delete * from users where id="&request("id")
	conn.execute sql,1,1 
	
	rsrz.open "book",conn,1,3      
	rsrz.addnew
	rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
	rsrz("caozuo")="删除"
	rsrz("menu")=menuname
	rsrz("duixian")=request("name1")
	rsrz.update
	rsrz.close
	
	response.redirect "?"
	Response.End
end if

if request("xs")="all" then
	session("users_all")="1"
	session("users_search_type")=""
	session("users_search_key")=""
end if

if request("search_key")<>"" then
	session("users_all")="0"
	session("users_search_type")=request("search_type")
	session("users_search_key")=request("search_key")
end if

if request("page")<>"" then
	session("users_page")=request("page")
end if
if session("users_page")="" then
	session("users_page")=1
end if

sqladd=""
if session("users_search_key")<>"" then
	sqladd=sqladd&" and "&session("users_search_type")&" like '%"&session("users_search_key")&"%' "
end if
%>

<body>
<!--#include file="top.asp"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23" align="center" background="images/am_mn2.gif" bgcolor="cccccc"><strong><%=menuname%></strong></td>
  </tr>
</table>
<br>

<center>
  <form id="form1" name="form1" method="post" action="?xs=all">
    <table border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td class="STYLE1">搜索：&nbsp;&nbsp;</td>
        <td>
		<select name="search_type" id="search_type">
            <option value="card" <%if session("users_search_type")="card" then%>selected="selected"<%end if%>>卡号</option>
            <option value="name1" <%if session("users_search_type")="name1" then%>selected="selected"<%end if%>>姓名</option>
			<option value="info" <%if session("users_search_type")="info" then%>selected="selected"<%end if%>>备注</option>
        </select>
	    </td>
        <td>&nbsp;&nbsp;<input name="search_key" type="text" id="search_key" value="<%=session("users_search_key")%>" onKeyPress="maskEdit(/^[\w]*$/)" /></td>
        <td align="right">&nbsp;&nbsp;<input type="submit" name="Submit22" value=" 搜索 " /></td>
        <td align="right">&nbsp;&nbsp;<input type="button" name="Submit2x2" onClick="location.href='?xs=all'" value="清除搜索" /></td>
        <td align="right">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="Submit2" onClick="location.href='usersAM.asp?cz=Add'" value=" 添加会员 " /></td>
     </tr>
    </table>
  </form>
</center>
<br>

<center>
<%
sql="select * from users where id>0 "&sqladd&" order by id desc"
rs.open sql,conn,1,1
%>
  <%if rs.bof then%>
  <p align="center">暂无信息</p>
  <%
  else
	rs.pagesize=10  '设置每页显示数目
	currentpage=Clng(session("users_page"))
	if currentpage<1 then currentpage=1
	if currentpage>rs.pagecount then currentpage=rs.pagecount
	rs.absolutepage=currentpage
  %>

  <table width="900" border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
	<tr align="center" bgcolor="#F0F0F0" class="STYLE1">
	  <td><strong>卡号</strong></td>
	  <td><strong>姓名</strong></td>
	  <td><strong>手机号</strong></td>
	  <td><strong>积分</strong></td>
	  <td><strong>备注</strong></td>
	  <td width="200"><strong>操作</strong></td>
	</tr>
<%Do while not rs.Eof%>
        <tr onMouseOver="bgColor='#D2F5FF';" onMouseOut="bgColor='#FFFFFF';">
          <td align="center">&nbsp;<%=rs("card")%>&nbsp;</td>
          <td align="left">&nbsp;<%=rs("name1")%>&nbsp;</td>
          <td align="left">&nbsp;<%=rs("tel")%>&nbsp;</td>
          <td align="right">&nbsp;<%=rs("score")%>&nbsp;</td>
          <td align="left">&nbsp;<%=rs("info")%>&nbsp;</td>
          <td align="center">
			<a href="?id=<%=rs("id")%>&cz=del&name1=<%=rs("name1")%>" onClick="{if(confirm('该会员的消费记录、兑换记录也将全部删除！确定要删除么?')){this.document.album.submit();return true;}return false;}">删除</a>&nbsp;&nbsp;
			<a href="usersAM.asp?id=<%=rs("id")%>&amp;cz=Modify">修改</a>&nbsp;&nbsp;
			<a href="XiaofeiAM.asp?cz=Add&usersid=<%=rs("id")%>">添加消费记录</a>&nbsp;&nbsp;
		  </td>
        </tr>
<%
i=i+1
rs.MoveNext
If i>=rs.pagesize Then Exit Do
Loop
%>
</table>
<!--#include file="inc/page.asp"-->
<%
end if
rs.close
%>
</center>
<!--#include file="foot.asp"-->
</body>
</html>