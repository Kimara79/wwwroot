<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="1css/style.css" rel="stylesheet" type="text/css" />
<title>ACCESS数据库管理</title>
<style type="text/css">
<!--
/*-----------------------------------------主体样式*/
body {
	margin-top: 10px;
	margin-bottom: 10px;
	margin-left: 10px;
	margin-right: 10px;
}
body,td,th {
	font-size: 12px;
	line-height:120%;
	white-space: nowrap;
}
form{
	margin:0px;
	margin-top:5px;
}

a:link {
	text-decoration: none;
	color: #18a795;
	font-weight:bold;
}
a:visited {
	text-decoration: none;
	color: #18a795;
	font-weight:bold;
}
a:hover {
	text-decoration: underline;
	color: #fda249;
	font-weight:bold;
}
a:active {
	text-decoration: none;
	color: #fda249;
	font-weight:bold;
}

a.a_red:link {
	text-decoration: none;
	color: #FF0000;
	font-weight:bold;
}
a.a_red:visited {
	text-decoration: none;
	color: #FF0000;
	font-weight:bold;
}
a.a_red:hover {
	text-decoration: underline;
	color: #fda249;
	font-weight:bold;
}
a.a_red:active {
	text-decoration: none;
	color: #fda249;
	font-weight:bold;
}

/*------------------------------------------其他样式*/
.textarea1{
	height:expression(this.scrollHeight);
	padding:10px;
	width:900px;
}
-->
</style>
</head>

<!--#include file="inc/power.asp"-->

<%
'on error resume next
'Session.Timeout=60
'Server.ScriptTimeOut=600
'Session.abandon
%>
<%
'字段类型名称   
Function columns_typename(id)
 id2=int(id)
 select case id
  case 2
   columns_typename="整型"
  case 3
   columns_typename="长整型"
  case 4
   columns_typename="单精度"
  case 5
   columns_typename="双精度"
  case 6
   columns_typename="货币"
  case 11
   columns_typename="布尔"
  case 17
   columns_typename="字节"
  case 72
   columns_typename="同步复制ID"
  case 131
   columns_typename="小数"
  case 134
   columns_typename="时间"
  case 135
   columns_typename="日期时间"
  case 202
   columns_typename="文本"
  case 203
   columns_typename="备注"
  case 205
   columns_typename="OLE对象"
 end select
End Function
%>

<%
'排序
if request("order")<>"" then
session("admin_order")=request("order")
end if
if session("admin_order")="" then
session("admin_order")="id desc"
end if
'搜索
if request("tj")="search" then
session("admin_search")="search"
session("admin_search_type")=request.form("search_type")
session("admin_search_key")=request.form("search_key")
end if
if request("xs")="all" then
session("admin_search")=""
end if
'换页
if request("page")<>"" then
session("admin_page")=request("page")
end if
if session("admin_page")="" then
session("admin_page")=1
end if
%>

<%
'----------------------------------------------------读取数据库

database="#data_XSDKJ2323#8.asp"

set conn=server.createobject("ADODB.CONNECTION")
connstr="DBQ="+server.mappath("databases/"&database)+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};pwd=1981#TOKING&cheny918"
conn.open connstr
set rstb=conn.openschema(20)  '表名
set rs=server.createobject("adodb.recordset")
set rs1=server.createobject("adodb.recordset")
'字段说明
Set cat=server.CreateObject("ADOX.Catalog")
cat.ActiveConnection="PROVIDER=Microsoft.jet.OLEDB.4.0;data source="&server.mappath("databases/"&database)&";Jet OLEDB:Database Password='1981#TOKING&cheny918'"

table=""
Do while not rstb.eof
 if rstb("table_type")="TABLE" then
  table=table&rstb("table_name")&","
 end if
 rstb.movenext
loop
rstb.close
table1=table
table=split(table,",")

if request("db")<>"" then
 if instr(table1,session("admin_tb")) then
 else
  session("admin_tb")=table(0)
 end if
 session("admin_page")=1
end if
if request("tb")<>"" then
 session("admin_tb")=request("tb")
 session("admin_page")=1
end if
if session("admin_tb")="" then
 session("admin_tb")=table(0)
end if
%>

<%
'----------------------------------------------------删除处理
if request("tj")="del" then
  sql="delete * from "&session("admin_tb")&" where id="&request("id")
  conn.execute sql,1,1
  response.redirect "?"
  response.End
end if

'----------------------------------------------------替换处理
if request("tj")="replace" then
  sql="select "&request.Form("replace_type")&" from "&session("admin_tb")&" where InStr(1,LCase(["&request.Form("replace_type")&"]),LCase('"&request.Form("replace_key1")&"'),0)<>0 "
  rs.open sql,conn,1,3
  Do while not rs.eof
   rs(request.Form("replace_type"))=replace(rs(request.Form("replace_type")),request.Form("replace_key1"),request.Form("replace_key2"))
   rs.update
   rs.movenext
  loop
  rs.close
  response.redirect "?"
  response.End
end if

'----------------------------------------------------修改处理
if request("tj")="modify" then
  sql="select * from "&session("admin_tb")&" where id="&request("id") 
  rs.open sql,conn,1,3
  For each obj in Request.Form
   rs(obj)=Request.Form(obj)
  next
  rs.update
  rs.close
  response.redirect "?"
  response.End
end if

'----------------------------------------------------添加处理
if request("tj")="add" then
  rs.open session("admin_tb"),conn,1,3
  rs.addnew
  For each obj in Request.Form
   rs(obj)=Request.Form(obj)
  next
  rs.update
  rs.close
  response.redirect "?"
  response.End
end if

'----------------------------------------------------导入处理
if request("tj")="dr" then
  fileall=request.Form("info_dr")
  fileall=replace(fileall,CHR(13),"")
  fileall=fileall&CHR(10)
  fileall=replace(fileall,CHR(10)&CHR(10),"")
  fileall=split(fileall,CHR(10))
  rs.open session("admin_tb"),conn,1,3
  for z=0 to ubound(fileall)
   if fileall(z)<>"" then
   fileline=fileall(z)
   fileline=split(fileline,CHR(9))
   rs.addnew
   for y=0 to ubound(fileline)
   rs(rs.Fields(y+1).name)=fileline(y)
   next
   rs.update
   end if
  next
  rs.close
  response.redirect "?"
  response.End
end if

'----------------------------------------------------执行处理
if request("tj")="run" then
  sql=trim(request.Form("info_run"))
  sql=replace(sql,CHR(13),"")
  sql=split(sql,CHR(10))
  for z=0 to ubound(sql)
   conn.execute sql(z),1,1
  next
  session("sql_run")=request.form("info_run")
  response.redirect "?"
  response.End
end if
%>

<body>
<center>
<!---检索---->
<table border="1" align="center" cellpadding="5" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr>
	<td align="center" bgcolor="#F0F0F0" class="STYLE1">表：</td>
	<td align="left">
  <%for z=0 to ubound(table)%>
	&nbsp;<a href="?tb=<%=table(z)%>" <%if table(z)=session("admin_tb") then%>class="a_red"<%end if%>><%=table(z)%></a>&nbsp;
  <%next%>
	</td>
  </tr>
</table>

<%
if session("admin_search")="search" then
 sql="select * from "&session("admin_tb")&" where "&session("admin_search_type")&" like '%"&session("admin_search_key")&"%' order by "&session("admin_order")&" "
else
 sql="select * from "&session("admin_tb")&" order by "&session("admin_order")&" "
end if
rs.open sql,conn,1,1
%>

<!---搜索---->
<form id="form_search" name="form_search" method="post" action="?tj=search" class="form1">
  <table border="0" cellspacing="0" cellpadding="5">
	<tr>
	  <td class="STYLE1">搜索：</td>
	  <td>
	  <select name="search_type" id="search_type">
	  <%for z=0 to rs.Fields.count-1%>
		  <option value="<%=rs.Fields(z).name%>"><%=rs.Fields(z).name%>(<%=cat.Tables(""&session("admin_tb")).Columns(""&rs.Fields(z).name).Properties("Description").Value%>)(<%=columns_typename(rs.Fields(z).type)%>)</option>
	  <%next%>
	  </select>
	  </td>
	  <td><input name="search_key" type="text" id="search_key"/></td>
	  <td align="right"><input type="submit" name="search_submit1" value=" 搜索 " /></td>
	  <td align="right"><input type="button" name="search_submit2" onclick="location.href='?xs=all'" value="清除搜索" />
	  </td>
	</tr>
  </table>
</form>

<!---替换---->
<form id="form_replace" name="form_replace" method="post" action="?tj=replace" class="form1">
  <table border="0" cellspacing="0" cellpadding="5">
	<tr>
	  <td class="STYLE1">替换：</td>
	  <td>
	  <select name="replace_type" id="replace_type">
	  <%for z=0 to rs.Fields.count-1%>
		  <option value="<%=rs.Fields(z).name%>"><%=rs.Fields(z).name%>(<%=cat.Tables(""&session("admin_tb")).Columns(""&rs.Fields(z).name).Properties("Description").Value%>)(<%=columns_typename(rs.Fields(z).type)%>)</option>
	  <%next%>
	  </select>
	  </td>
	  <td><input name="replace_key1" type="text" id="replace_key1"/> 替换为 </td>
	  <td><input name="replace_key2" type="text" id="replace_key2"/></td>
	  <td align="right"><input type="submit" name="replace_submit" value=" 替换 " /></td>
	</tr>
  </table>
</form>
<br>

<!---表内容---->
<table border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr bgcolor="#F0F0F0">
  <!---表头---->
  <%for z=0 to rs.Fields.count-1%>
	<td style="cursor:default;">
	&nbsp;<a href="?order=<%=rs.Fields(z).name%>" title="按此排序"><%=cat.Tables(""&session("admin_tb")).Columns(""&rs.Fields(z).name).Properties("Description").Value%></a>&nbsp;
	</td>
  <%next%>
    <td align="center" class="STYLE1" style="cursor:default;" rowspan="2">&nbsp;操作&nbsp;</td>
  </tr>
  <tr bgcolor="#F0F0F0">
  <!---表头---->
  <%for z=0 to rs.Fields.count-1%>
	<td style="cursor:default;">
	<a href="?order=<%=rs.Fields(z).name%>" title="按此排序"><%=rs.Fields(z).name%></a>
	</td>
  <%next%>
  </tr>
  <!---添加---->
  <form id="form_add" name="form_add" method="post" action="?tj=add">
  <tr>
    <td align="center">自动</td>
    <%for z=1 to rs.Fields.count-1%>
	<td align="left"><input name="<%=rs.Fields(z).name%>" type="text" style="width:90%;" value=""/></td>
    <%next%>
    <td align="center"><input type="submit" name="add_submit" value=" 添加 " /></td>
  </tr>
  </form>
<%
rs.pagesize=20  '设置每页显示数目
currentpage=Clng(session("admin_page"))
if currentpage<1 then currentpage=1
if currentpage>rs.pagecount then currentpage=rs.pagecount
if rs.recordcount>0 then
rs.absolutepage=currentpage
end if
Do While Not rs.Eof
%>
  <!---修改---->
  <%if request("cz")="modify" and int(rs("id"))=int(request("id")) then%>
  <form id="form_modify" name="form_modify" method="post" action="?tj=modify&id=<%=request("id")%>">
  <tr>
    <%for z=0 to rs.Fields.count-1%>
	<td align="left"><input name="<%=rs.Fields(z).name%>" type="text" value="<%=rs(rs.Fields(z).name)%>" onkeydown="if(event.keyCode==13){this.parentElement.parentElement.parentElement.submit();}" title="回车可提交"/></td>
    <%next%>
    <td align="center"><input type="button" name="historynack" value="返回" onclick="location.href='javascript:onclick=history.go(-1)'" />&nbsp;<input type="submit" name="modify_submit" value="修改" /></td>
  </tr>
  </form>
  <%else%>
  <!---常规内容---->
  <tr style="cursor:pointer;" onmouseover="bgColor='#D2F5FF';" onmouseout="bgColor='#FFFFFF';" onClick="location.href='?cz=modify&id=<%=rs("id")%>'">
    <%for z=0 to rs.Fields.count-1%>
	<td align="left" title="点击进行修改">&nbsp;<%=rs(rs.Fields(z).name)%>&nbsp;</td>
    <%next%>
    <td align="center">
	<a href="?tj=del&id=<%=rs("id")%>" onclick="{if(confirm('确定删除么？')){this.document.album.submit();return true;}return false;}">删除</a>&nbsp;
	<a href="?cz=modify&id=<%=rs("id")%>">修改</a>
	</td>
  </tr>
  <%end if%>
<%
i=i+1
rs.MoveNext
If i>=rs.pagesize Then Exit Do
Loop
%>
</table>

<!---换页---->
<form action="?" method="post" name="form_page" id="form_page" class="form1">
共有<font color="red"><%=rs.recordcount%></font>条信息&nbsp;&nbsp;
每页<font color="red"><%=rs.pagesize%></font>条&nbsp;&nbsp;&nbsp;&lt;&lt;&nbsp;
<%if currentpage>1 then%>
<a href="?page=1">首页</a>
<%else%>
首页
<%end if%>
<%if currentpage>1 then%>
<a href="?page=<%=currentpage-1%>"> 上一页</a>&nbsp;
<%else%>
上一页
<%end if%>
第<font color="red"><%=currentpage%></font>/<font color="red"><%=rs.pagecount%></font>页
<%if currentpage<rs.pagecount then%>
<a href="?page=<%=currentpage+1%>">下一页</a>&nbsp;
<%else%>
下一页
<%end if%>
<%if currentpage<rs.pagecount then%>
<a href="?page=<%=rs.pagecount%>">尾页</a>
<%else%>
尾页
<%end if%>
&gt;&gt;
&nbsp;&nbsp;&nbsp;转到
<input name="page" type="text" id="page" value="<%=currentpage%>" size="4" />
&nbsp;
<input type="submit" name="Submit" value=" GO " />
</form>
<br>
  
<!---执行---->
  <form id="form_run" name="form_run" method="post" action="?tj=run">
<table border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="left" valign="top">
	增加字段：ALTER TABLE <%=session("admin_tb")%> ADD COLUMN a char(50)<br>
	修改字段：ALTER TABLE <%=session("admin_tb")%> ALTER COLUMN a Short<br>
	删除字段：ALTER TABLE <%=session("admin_tb")%> DROP COLUMN a<br>
	增加表：Create TABLE <%=session("admin_tb")%><br>
	删除表：DROP TABLE <%=session("admin_tb")%><br>
	添加记录：insert into <%=session("admin_tb")%> (a,b) values (1,2)<br>
	添加记录：insert into <%=session("admin_tb")%> (name1,info,laster,lasttime) values ('','','admin',#2010-3-9 11:11:11#)<br>
	删除记录：delect * from <%=session("admin_tb")%> where id=<br>
	更新记录：update <%=session("admin_tb")%> set a=1,b=2 where id=1<br>
Byte 数字[字节]，Long 数字[长整型]，Short 数字[整型]，Single 数字[单精度]，Double 数字[双精度],Currency 货币<br>
Char 文本，Text(n) 文本，其中n表示字段大小，Binary 二进制，Counter 自动编号，Memo 备注，Time 日期/时间
	</td>
  </tr>
</table>
    <textarea name="info_run" class="textarea1"><%=session("sql_run")%></textarea><br>
    <input type="submit" name="run_submit" value=" 执行 " />
  </form>
<br>
<br>

<!---导入---->
  <form id="form_dr" name="form_dr" method="post" action="?tj=dr">
    注意：第1列的id为自动。导入内容从第2列开始。可从电子表格直接粘贴过来，或者手动用Tab做分隔符。<br>
    <textarea name="info_dr" class="textarea1"></textarea>
    <br>
    <input type="submit" name="dr_submit" value=" 导入 " />
  </form>

<%rs.close%>
</center>
</body>
</html>