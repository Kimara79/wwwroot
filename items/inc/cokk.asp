<%
On Error Resume Next
dim conn,connstr
connstr="DBQ="+server.mappath("databases/#data_XSDKJ2323#8.asp")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
conn.open connstr

set rs=server.createobject("adodb.recordset")
set rs1=server.createobject("adodb.recordset")
set rs2=server.createobject("adodb.recordset")
set rssm=server.createobject("adodb.recordset")
set rsrz=server.createobject("adodb.recordset")
%>