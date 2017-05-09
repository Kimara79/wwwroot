<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="表单管理" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="inc/num_write.asp"-->
<!--#include file="js/red.js"-->

<%
'其他设置
sql="select info from system where id=21"
rsstr=trim(conn.execute(sql).getstring)
rsstr=replace(rsstr,chr(13),"")
shezhi=split(rsstr,",")

'删除操作
if request("cz")="del" then
	sql="delete * from forms where id="&request("id")               '删除数据
	conn.execute sql,1,1 
	
	rsrz.open "book",conn,1,3                                   '添加日志              
	rsrz.addnew
	rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
	rsrz("caozuo")="删除"
	rsrz("menu")=menuname
	rsrz("duixian")=request("id")
	rsrz.update
	rsrz.close
	
	response.redirect "?"
	Response.End
end if
%>

<%
'接收
if request("xs")="all" then
	session("forms_all")="1"
	session("forms_search_type")=""
	session("forms_search_key")=""
	session("forms_sz")=""
	session("forms_menu1")=""
	session("forms_menu2")=""
end if
if request("search_key")<>"" then
	session("forms_all")="0"
	session("forms_search_type")=request("search_type")
	session("forms_search_key")=request("search_key")
end if
if request("sz")<>"" then
	session("forms_all")="0"
	session("forms_sz")=request("sz")
end if
if request("menu1")<>"" then
	session("forms_all")="0"
	session("forms_menu1")=request("menu1")
end if
if request("menu2")<>"" then
	session("forms_all")="0"
	session("forms_menu2")=request("menu2")
end if

if request("page")<>"" then
	session("forms_page")=request("page")
end if
if session("forms_page")="" then
	session("forms_page")=1
end if

'默认
On Error Resume Next
if session("forms_menu1")="" and session("forms_menu2")="" then
	rsstr=""
	sql="select top 1 id from menu1 where leixing like '%表单提交%' and id not in (select menu1 from menu2) order by no1"
	rsstr=trim(conn.execute(sql).getstring)
	rsstr=replace(""&rsstr,chr(13),"")
	if rsstr<>"" then
		session("forms_menu1")=rsstr
	else
		rsstr=""
		sql="select top 1 id from menu2 where leixing like '%表单提交%' order by no1"
		rsstr=trim(conn.execute(sql).getstring)
		rsstr=replace(""&rsstr,chr(13),"")
		session("forms_menu2")=rsstr
	end if
end if

'SQL
sqladd=""
sqladd2=""
if session("forms_search_key")<>"" then
	sqladd=sqladd&" and "&session("forms_search_type")&" like '%"&session("forms_search_key")&"%' "
end if
if session("forms_sz")<>"" then
	sqladd=sqladd&" and shezhi like '%"&session("forms_sz")&"%' "
end if
if session("forms_menu1")<>"" then
	sqladd=sqladd&" and menu1="&session("forms_menu1")&" "
	sqladd2=sqladd2&" and menu1="&session("forms_menu1")&" "
end if
if session("forms_menu2")<>"" then
	sqladd=sqladd&" and menu2="&session("forms_menu2")&" "
	sqladd2=sqladd2&" and menu2="&session("forms_menu2")&" "
end if
sqladd2=sqladd2&" and shezhi like '%后台列表页显示%' "
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
  <form id="form1" name="form1" method="post" action="?xs=all&menu1=<%=session("forms_menu1")%>&menu2=<%=session("forms_menu2")%>">
    <table border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td class="STYLE1">搜索：&nbsp;&nbsp;</td>
        <td>
		<select name="search_type" id="search_type">
		<%
		sql="select * from ziduan where id>0 "&sqladd2&" order by no1"
		rs1.open sql,conn,1,1
		Do while not rs1.Eof
		%>
            <option value="zd_<%=rs1("id")%>" <%if session("forms_search_type")="zd_"&rs1("id") then%>selected="selected"<%end if%>><%=rs1("name1")%></option>
		<%
		rs1.MoveNext
		Loop
		rs1.close
		%>
        </select>
	    </td>
        <td>&nbsp;&nbsp;<input name="search_key" type="text" id="search_key" value="<%=session("forms_search_key")%>" onKeyPress="maskEdit(/^[\w]*$/)" /></td>
        <td align="right">&nbsp;&nbsp;<input type="submit" name="Submit22" value=" 搜索 " /></td>
        <td align="right">&nbsp;&nbsp;<input type="button" name="Submit2x2" onClick="location.href='?xs=all'" value="清除搜索" /></td>
     </tr>
    </table>
  </form>
</center>
<br>

<%
toptb="forms"
likes="表单提交"
%>
<!--#include file="top2.asp"-->

<center>
<%
sql="select * from forms where id>0 "&sqladd&" order by id desc"
rs.open sql,conn,1,1
sql="select * from ziduan where id>0 "&sqladd2&" order by no1"
rs1.open sql,conn,1,1
%>
<%if rs.bof then%>
	<p align="center">暂无信息</p>
<%
else
	rs.pagesize=20  '设置每页显示数目
	currentpage=Clng(session("forms_page"))
	if currentpage<1 then currentpage=1
	if currentpage>rs.pagecount then currentpage=rs.pagecount
	rs.absolutepage=currentpage
  %>
  <table border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
	<tr align="center" bgcolor="#F0F0F0" class="STYLE1">
	  <td><strong>ID</strong></td>
	  <td><strong>一级栏目</strong></td>
	  <td><strong>二级栏目</strong></td>
	  <%
	  rs1.movefirst
	  Do while not rs1.Eof
	  %>
	  <td><strong><%=rs1("name1")%></strong></td>
	  <%
	  rs1.MoveNext
	  loop
	  %>
	  <td><strong>设置</strong></td>
	  <td><strong>操作</strong></td>
	</tr>
	<%Do while not rs.Eof%>
	<form id="change<%=rs("id")%>" name="change<%=rs("id")%>" method="post" action="update.asp?id=<%=rs("id")%>&tb=forms&zd=shezhi" target="update">
		<tr onMouseOver="bgColor='#D2F5FF';" onMouseOut="bgColor='#FFFFFF';">
		  <td align="center">&nbsp;<%=rs("id")%>&nbsp;</td>
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
		  <%
		  rs1.movefirst
		  Do while not rs1.Eof
		  %>
		  <td align="left">&nbsp;<%if rs1("kongjian")="段落" then%><%=num_write(rs("zd_"&rs1("id")),10)%><%else%><%=rs("zd_"&rs1("id"))%><%end if%>&nbsp;</td>
		  <%
		  rs1.MoveNext
		  loop
		  %>
          <td align="center">
			<%for z=0 to ubound(shezhi)%>
			<label <%if instr(""&rs("shezhi"),shezhi(z)) then%>style="color:red;"<%end if%>>
			<input name="shezhi" type="checkbox" class="dq" value="<%=shezhi(z)%>" <%if instr(""&rs("shezhi"),shezhi(z)) then%> checked="checked" <%end if%> onClick="red(this);document.getElementById('change<%=rs("id")%>').submit();"/> <%=shezhi(z)%></label>
			<%next%>		  </td>
          <td align="center">
			<a href="?cz=del&id=<%=rs("id")%>" onClick="{if(confirm('确定要删除么?')){this.document.album.submit();return true;}return false;}">删除</a>&nbsp;&nbsp;
			<a href="FormsLook.asp?id=<%=rs("id")%>">查看</a>
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
rs1.close
rs.close
%>
</center>
<iframe id="update" name="update" style="display:none;"></iframe>
<!--#include file="foot.asp"-->
</body>
</html>