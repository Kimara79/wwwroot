<%
'On Error Resume Next
'定义数据库连接
dim conn,connstr
connstr="DBQ="+server.mappath("items/databases/#data_XSDKJ2323#8.asp")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)}"
set conn=server.createobject("ADODB.CONNECTION")
conn.open connstr
set rs=server.createobject("adodb.recordset")
set rs1=server.createobject("adodb.recordset")
set rs2=server.createobject("adodb.recordset")
set rsrz=server.createobject("adodb.recordset")
%>

<%
'网站域名
sql="select info from system where id=11"
rsstr=trim(conn.execute(sql).getstring)
rsstr=replace(rsstr,chr(13),"")
site_url=rsstr
site_url2="http://"&rsstr
'网站名称
sql="select info from system where id=7"
rsstr=trim(conn.execute(sql).getstring)
rsstr=replace(rsstr,chr(13),"")
site_name=rsstr
'网站标题
sql="select info from system where id=8"
rsstr=trim(conn.execute(sql).getstring)
rsstr=replace(rsstr,chr(13),"")
site_title=rsstr
'网站关键字
sql="select info from system where id=9"
rsstr=trim(conn.execute(sql).getstring)
rsstr=replace(rsstr,chr(13),"")
site_keywords=rsstr
'网站描述
sql="select info from system where id=10"
rsstr=trim(conn.execute(sql).getstring)
rsstr=replace(rsstr,chr(13),"")
site_description=rsstr

'日期检测
'outdate="2014-8-1"
'if DateDiff("d",date(),outdate)<0 then
'	Response.End
'end if
%>
<meta name="keywords" content="<%=site_keywords%>" />
<meta name="description" content="<%=site_description%>" />
<meta name="author" content="<%=site_name%>,<%=site_url2%>" />
<meta name="owner" content="<%=site_name%>,<%=site_url2%>" />
<meta name="Copyright" content="<%=site_name%>,<%=site_url2%>" />

<!--#include file="inc/requesths.asp"-->
<!--#include file="inc/menuurl.asp"-->
<!--#include file="inc/num_write.asp"-->
<!--#include file="items/js/fbjs.asp"-->
