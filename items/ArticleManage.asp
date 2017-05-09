<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="文章管理" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="js/red.js"-->

<%
'其他设置
sql="select info from system where id=4"
rsstr=trim(conn.execute(sql).getstring)
rsstr=replace(rsstr,chr(13),"")
shezhi=split(rsstr,",")

'删除操作
if request("cz")="del" then
	sql="delete * from article where id="&request("id")               '删除数据
	conn.execute sql,1,1 
	
	rsrz.open "book",conn,1,3                                   '添加日志              
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
	session("article_all")="1"
	session("article_search_type")=""
	session("article_search_key")=""
	session("article_sz")=""
	session("article_menu1")=""
	session("article_menu2")=""
end if

if request("search_key")<>"" then
	session("article_all")="0"
	session("article_search_type")=request("search_type")
	session("article_search_key")=request("search_key")
end if
if request("sz")<>"" then
	session("article_all")="0"
	session("article_sz")=request("sz")
end if
if request("menu1")<>"" then
	session("article_all")="0"
	session("article_menu1")=request("menu1")
end if
if request("menu2")<>"" then
	session("article_all")="0"
	session("article_menu2")=request("menu2")
end if

if request("page")<>"" then
	session("article_page")=request("page")
end if
if session("article_page")="" then
	session("article_page")=1
end if

sqladd=""
if session("article_search_key")<>"" then
	sqladd=sqladd&" and "&session("article_search_type")&" like '%"&session("article_search_key")&"%' "
end if
if session("article_sz")<>"" then
	sqladd=sqladd&" and shezhi like '%"&session("article_sz")&"%' "
end if
if session("article_menu1")<>"" then
	sqladd=sqladd&" and menu1="&session("article_menu1")&" "
end if
if session("article_menu2")<>"" then
	sqladd=sqladd&" and menu2="&session("article_menu2")&" "
end if

'预设排序
sql="update article set no1=id where (no1 is null or no=null or no1=0)"
conn.execute sql,1,1 

'接收排序
if request("tj")="order" and int(request("id1"))>0 and int(request("id2"))>0 then
	sql="update article set no1="&request("no2")&" where id="&request("id1")&" "
	conn.execute sql,1,1
	sql="update article set no1="&request("no1")&" where id="&request("id2")&" "
	conn.execute sql,1,1
	response.redirect "?"
	response.end
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
            <option value="name1" <%if session("article_search_type")="name1" then%>selected="selected"<%end if%>>标题</option>
			<option value="info" <%if session("article_search_type")="info" then%>selected="selected"<%end if%>>内容</option>
        </select>
	    </td>
        <td>&nbsp;&nbsp;<input name="search_key" type="text" id="search_key" value="<%=session("article_search_key")%>" onKeyPress="maskEdit(/^[\w]*$/)" /></td>
        <td align="right">&nbsp;&nbsp;<input type="submit" name="Submit22" value=" 搜索 " /></td>
        <td align="right">&nbsp;&nbsp;<input type="button" name="Submit2x2" onClick="location.href='?xs=all'" value="清除搜索" /></td>
        <td align="right">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="Submit2" onClick="location.href='articleAM.asp?cz=Add'" value=" 添加文章 " /></td>
     </tr>
    </table>
  </form>
</center>
<br>

<%
toptb="article"
likes="多篇文章"
%>
<!--#include file="top2.asp"-->

<center>
<%
sql="select * from article where id>0 "&sqladd&" order by shezhi desc,no1"
rs.open sql,conn,1,1
%>
  <%if rs.bof then%>
  <p align="center">暂无信息</p>
  <%
  else
	rs.pagesize=20  '设置每页显示数目
	currentpage=Clng(session("article_page"))
	if currentpage<1 then currentpage=1
	if currentpage>rs.pagecount then currentpage=rs.pagecount
	rs.absolutepage=currentpage
  %>

  <table border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
	<tr align="center" bgcolor="#F0F0F0" class="STYLE1">
	  <td><strong>ID</strong></td>
	  <td><strong>一级栏目</strong></td>
	  <td><strong>二级栏目</strong></td>
	  <td><strong>标题</strong></td>
	  <td><strong>特别链接</strong></td>
	  <td><strong>点击数</strong></td>
	  <td><strong>设置</strong></td>
	  <td><strong>操作</strong></td>
	</tr>
<%
Do while not rs.Eof
  rs.moveprevious
  if not rs.bof then
   id2a=rs("id")
   no2a=rs("no1")
  else
   id2a=0
   no2a=0
  end if
  rs.movenext
  rs.movenext
  if not rs.eof then
   id2b=rs("id")
   no2b=rs("no1")
  else
   id2b=0
   no2b=0
  end if
  rs.moveprevious
%>
<form id="change<%=rs("id")%>" name="change<%=rs("id")%>" method="post" action="update.asp?id=<%=rs("id")%>&tb=article&zd=shezhi" target="update">
        <tr onMouseOver="bgColor='#D2F5FF';" onMouseOut="bgColor='#FFFFFF';">
          <td align="center">&nbsp;
              <%=rs("id")%>
  &nbsp;</td>
          <td align="center">&nbsp;
<%
sql="select name1 from menu1 where id="&rs("menu1")&" "
rsstr=trim(conn.execute(sql).getstring)
rsstr=replace(rsstr,chr(13),"")
response.Write rsstr
%>
  &nbsp;</td>
          <td align="center">&nbsp;
<%
if rs("menu2")>0 then
	sql="select name1 from menu2 where id="&rs("menu2")&" "
	rsstr=trim(conn.execute(sql).getstring)
	rsstr=replace(rsstr,chr(13),"")
	response.Write rsstr
end if
%>
  &nbsp;</td>
          <td align="left">&nbsp;<%=rs("name1")%>&nbsp;</td>
          <td align="left">&nbsp;<%=rs("url")%>&nbsp;</td>
          <td align="center">&nbsp;<%=rs("click")%>&nbsp;</td>
          <td align="center">
			<%for z=0 to ubound(shezhi)%>
			<label <%if instr(rs("shezhi"),shezhi(z)) then%>style="color:red;"<%end if%>>
			<input name="shezhi" type="checkbox" class="dq" value="<%=shezhi(z)%>" <%if instr(""&rs("shezhi"),shezhi(z)) then%> checked="checked" <%end if%> onClick="red(this);document.getElementById('change<%=rs("id")%>').submit();"/> <%=shezhi(z)%></label>
			<%next%>		  </td>
          <td align="center">&nbsp;&nbsp;
			<a href="?id=<%=rs("id")%>&cz=del&name1=<%=rs("name1")%>" onClick="{if(confirm('确定要删除么?')){this.document.album.submit();return true;}return false;}">删除</a>&nbsp;&nbsp;
			<a href="articleAM.asp?id=<%=rs("id")%>&amp;cz=Modify">修改</a>&nbsp;&nbsp;
		<a href="?tj=order&id1=<%=rs("id")%>&no1=<%=rs("no1")%>&id2=<%=id2a%>&no2=<%=no2a%>">上移</a>&nbsp;&nbsp;
		<a href="?tj=order&id1=<%=rs("id")%>&no1=<%=rs("no1")%>&id2=<%=id2b%>&no2=<%=no2b%>">下移</a>&nbsp;&nbsp;
		  </td>
        </tr>
</form>
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
<iframe id="update" name="update" style="display:none;"></iframe>
<!--#include file="foot.asp"-->
</body>
</html>