<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title>数据库管理</title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="inc/mofei.asp"-->

<%
'----------------------------------------------压缩操作
if request("cz")="ys" then
oldDB = server.mappath("databases/#data_XSDKJ2323#8.asp") '更改数据库地址
newDB = server.mappath("databases/#data_XSDKJ2323#8_new.asp") '生成临时文件
Set FSO = Server.CreateObject("Scripting.FileSystemObject")
Set Engine = Server.CreateObject("JRO.JetEngine")
prov = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Engine.CompactDatabase prov & OldDB, prov & newDB
set Engine = nothing
FSO.DeleteFile oldDB '删除临时文件
FSO.MoveFile newDB, oldDB
set FSO = Nothing
response.redirect "DataManage.asp"
response.end
end if

'----------------------------------------------备份操作
if request("cz")="bf" then 
oldname=server.mappath("databases/#data_XSDKJ2323#8.asp")
newname=server.mappath("databases/"&"backup_"&request("i")&".xxx")
Set fso=server.CreateObject("scripting.FileSystemObject")
on error resume next
fso.copyfile oldname,newname
set fso=nothing
response.redirect "DataManage.asp"
response.end
end if

'----------------------------------------------还原操作
if request("cz")="hy" then
oldname=server.mappath("databases/"&"backup_"&request("i")&".xxx")
newname=server.mappath("databases/#data_XSDKJ2323#8.asp")
Set fso=server.CreateObject("scripting.FileSystemObject")
on error resume next
fso.copyfile oldname,newname
set fso=nothing
response.redirect "DataManage.asp"
response.end
end if
%>

<body>
<div align="center">
<!--#include file="top.asp"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
	  <td height="23" align="center" background="images/am_mn2.gif" bgcolor="cccccc"><strong>数据库管理</strong></td>
	</tr>
</table>

<br>
<table width="800" border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr class="STYLE1">
    <td height="20" align="center" bgcolor="#F0F0F0">
	当前数据库：&nbsp;&nbsp;
	<%=mofei2("databases/#data_XSDKJ2323#8.asp")%>&nbsp;&nbsp;&nbsp;&nbsp;
	<%=mofei3("databases/#data_XSDKJ2323#8.asp")%>&nbsp;&nbsp;&nbsp;&nbsp;
	<a href="DataManage.asp?cz=ys">压缩数据库</a>
	</td>
  </tr>
</table>
<br />
<table width="800" border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
<tr class="STYLE1">
  <td align="center" bgcolor="#F0F0F0">文件名称</td>
  <td align="center" bgcolor="#F0F0F0">文件大小</td>
  <td align="center" bgcolor="#F0F0F0">文件时间</td>
  <td align="center" bgcolor="#F0F0F0">操作</td>
</tr>
<%for i=1 to 3%>
<tr onmouseover="bgColor='#D2F5FF';" onmouseout="bgColor='#FFFFFF';">
  <td align="center">&nbsp;<%=mofei2("databases/backup_"&i&".xxx")%>&nbsp;</td>
  <td align="center">&nbsp;<%=mofei3("databases/backup_"&i&".xxx")%>&nbsp;</td>
  <td align="center">&nbsp;<%=mofei1("databases/backup_"&i&".xxx")%>&nbsp;</td>
  <td align="center"><input type="button" name="Submit2" onclick="{if(confirm('一旦还原，新更新的数据将全部丢失！确定么？\n\n建议先做备份后再还原！')){location.href='DataManage.asp?cz=hy&i=<%=i%>';return true;}return false;}" value=" 还 原 " />
  &nbsp;&nbsp;
  <input type="button" name="Submit2" onclick="location.href='DataManage.asp?cz=bf&i=<%=i%>'" value=" 备 份 " /></td>
</tr>
<%next%>
</table>

<br>
<iframe id="daochu" name="daochu" style="display:none;"></iframe>
<!--#include file="foot.asp"-->
</div>
</body>
</html>