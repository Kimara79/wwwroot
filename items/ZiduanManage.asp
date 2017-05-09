<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="字段管理" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="ziduan_str.asp"-->
<!--#include file="js/red.js"-->

<%
'删除操作
if request("cz")="del" then
	On Error Resume Next
	sql="delete * from ziduan where id="&request("id")               '删除数据
	conn.execute sql,1,1
	sql="ALTER TABLE forms DROP COLUMN zd_"&request("id")&" "        '删除字段
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
%>

<%
if request("xs")="all" then
	session("ziduan_all")="1"
	session("ziduan_search_type")=""
	session("ziduan_search_key")=""
	session("ziduan_sz")=""
	session("ziduan_menu1")=""
	session("ziduan_menu2")=""
end if

if request("search_key")<>"" then
	session("ziduan_all")="0"
	session("ziduan_search_type")=request("search_type")
	session("ziduan_search_key")=request("search_key")
end if
if request("sz")<>"" then
	session("ziduan_all")="0"
	session("ziduan_sz")=request("sz")
end if
if request("menu1")<>"" then
	session("ziduan_all")="0"
	session("ziduan_menu1")=request("menu1")
end if
if request("menu2")<>"" then
	session("ziduan_all")="0"
	session("ziduan_menu2")=request("menu2")
end if

if request("page")<>"" then
	session("ziduan_page")=request("page")
end if
if session("ziduan_page")="" then
	session("ziduan_page")=1
end if

sqladd=""
if session("ziduan_search_key")<>"" then
	sqladd=sqladd&" and "&session("ziduan_search_type")&" like '%"&session("ziduan_search_key")&"%' "
end if
if session("ziduan_sz")<>"" then
	sqladd=sqladd&" and shezhi like '%"&session("ziduan_sz")&"%' "
end if
if session("ziduan_menu1")<>"" then
	sqladd=sqladd&" and menu1="&session("ziduan_menu1")&" "
end if
if session("ziduan_menu2")<>"" then
	sqladd=sqladd&" and menu2="&session("ziduan_menu2")&" "
end if

'预设排序
sql="update ziduan set no1=id where (no1 is null or no=null or no1=0)"
conn.execute sql,1,1 

'接收排序
if request("tj")="order" and int(request("id1"))>0 and int(request("id2"))>0 then
	sql="update ziduan set no1="&request("no2")&" where id="&request("id1")&" "
	conn.execute sql,1,1
	sql="update ziduan set no1="&request("no1")&" where id="&request("id2")&" "
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
            <option value="name1" <%if session("ziduan_search_type")="name1" then%>selected="selected"<%end if%>>名称</option>
			<option value="info" <%if session("ziduan_search_type")="info" then%>selected="selected"<%end if%>>说明</option>
        </select>
	    </td>
        <td>&nbsp;&nbsp;<input name="search_key" type="text" id="search_key" value="<%=session("ziduan_search_key")%>" onKeyPress="maskEdit(/^[\w]*$/)" /></td>
        <td align="right">&nbsp;&nbsp;<input type="submit" name="Submit22" value=" 搜索 " /></td>
        <td align="right">&nbsp;&nbsp;<input type="button" name="Submit2x2" onClick="location.href='?xs=all'" value="清除搜索" /></td>
        <td align="right">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="Submit2" onClick="location.href='ziduanAM.asp?cz=Add'" value=" 添加字段 " /></td>
     </tr>
    </table>
  </form>
</center>
<br>

<%
toptb="ziduan"
likes="表单提交"
%>
<!--#include file="top2.asp"-->

<center>
<%
sql="select * from ziduan where id>0 "&sqladd&" order by no1 "
rs.open sql,conn,1,1
%>
  <%if rs.bof then%>
  <p align="center">暂无信息</p>
  <%
  else
	rs.pagesize=200  '设置每页显示数目
	currentpage=Clng(session("ziduan_page"))
	if currentpage<1 then currentpage=1
	if currentpage>rs.pagecount then currentpage=rs.pagecount
	rs.absolutepage=currentpage
  %>

  <table border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
	<tr align="center" bgcolor="#F0F0F0" class="STYLE1">
	  <td><strong>ID</strong></td>
	  <td><strong>一级栏目</strong></td>
	  <td><strong>二级栏目</strong></td>
	  <td>字段名称</td>
	  <td>数据类型</td>
	  <td>控件类型</td>
	  <td>控件宽度</td>
	  <td>控件注释</td>
	  <td>控件选项</td>
	  <td><strong>其他设置</strong></td>
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
<form id="change<%=rs("id")%>" name="change<%=rs("id")%>" method="post" action="update.asp?id=<%=rs("id")%>&tb=ziduan&zd=shezhi" target="update">
	<tr onMouseOver="bgColor='#D2F5FF';" onMouseOut="bgColor='#FFFFFF';">
	  <td align="center">&nbsp;<%=rs("id")%>&nbsp;</td>
	  <td align="left">&nbsp;
		<%
		sql="select name1 from menu1 where id="&rs("menu1")&" "
		rsstr=trim(conn.execute(sql).getstring)
		rsstr=replace(rsstr,chr(13),",")
		response.Write rsstr
		%>
		&nbsp;</td>
      <td align="left">&nbsp;
		<%
		if rs("menu2")<>"" then
			sql="select name1 from menu2 where id="&rs("menu2")&" "
			rsstr=trim(conn.execute(sql).getstring)
			rsstr=replace(rsstr,chr(13),",")
			response.Write rsstr
		end if
		%>
		&nbsp;</td>
	  <td align="left">&nbsp;<%=rs("name1")%>&nbsp;</td>
	  <td align="left">&nbsp;<%=rs("shuju")%><%if rs("length_sj")<>"" then%>(<%=rs("length_sj")%>)<%end if%>&nbsp;</td>
	  <td align="left">&nbsp;<%=rs("kongjian")%>&nbsp;</td>
	  <td align="left">&nbsp;<%=rs("length_kj")%>&nbsp;</td>
	  <td align="left">&nbsp;<%=rs("zhushi")%>&nbsp;</td>
	  <td align="left" <%if rs("info")<>"" then%>title="<%=rs("info")%>"<%end if%>>&nbsp;<%if rs("info")<>"" then%><%=left(rs("info"),5)%>..<%end if%>&nbsp;</td>
	  <td align="center"><%for z=0 to ubound(shezhi)%>
		  <label <%if instr(rs("shezhi"),shezhi(z)) then%>style="color:red;"<%end if%>>
		  <input name="shezhi" type="checkbox" class="dq" value="<%=shezhi(z)%>" <%if instr(""&rs("shezhi"),shezhi(z)) then%> checked="checked" <%end if%> onClick="red(this);document.change<%=rs("id")%>.submit();"/>
		  <%=shezhi(z)%></label>
		  <%next%>	  </td>
	  <td align="center">&nbsp;&nbsp;
		<a href="?id=<%=rs("id")%>&cz=del&name1=<%=rs("name1")%>" onClick="{if(confirm('确定要删除么?')){this.document.album.submit();return true;}return false;}">删除</a>&nbsp;&nbsp;
		<a href="ziduanAM.asp?id=<%=rs("id")%>&amp;cz=Modify">修改</a>&nbsp;&nbsp;
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