<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="背景音乐" %>
<% menunamel="音乐" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="inc/UpLoadClass.asp"-->

<%
picfld="music/"
chicun_sm="可直接上传文件(5M以内)，或填上外链地址(末尾必须是：.mp3)<br />上传的文件格式必须为mp3<br />如上传跟外链同时存在，将优先调用上传的文件"
%>

<%
tj=request.QueryString("tj")

'添加/修改
if (tj="Add" or tj="Modify") then
	Server.ScriptTimeOut=600   '上传时间限制,单位为秒
	
	'初始化类
	set requesta=New UpLoadClass
	requesta.Charset="UTF-8"
	requesta.TotalSize=0
	requesta.MaxSize=0
	requesta.FileType=""
	requesta.AutoSave=2
	requesta.SavePath=picfld
	requesta.Open()
	
	'读取文件
	filea="file1"
	fileext=lcase(requesta.form(filea&"_Ext"))   '文件格式
	filemaxsize=5  '上传大小限制,单位M

	'出错判断
	if requesta.Form("name1")="" then
		Response.Write "<script>alert('请填上"&menunamel&"名称！');history.back();</script>"
		Response.End
	end if
	if (tj="Add" and fileext="" and requesta.Form("url")="") then
		Response.Write "<script>alert('请上传或外链文件！');history.back();</script>"
		Response.End
	end if
	if fileext<>"" then
		if (fileext<>"mp3") then 
			Response.Write "<script>alert('文件格式错误！请上传 mp3 格式的文件！');history.back();</script>"
			Response.End
		end if
		filesize=FormatNumber(requesta.form(filea&"_Size")/(1024*1024),2,-1)
		if filesize>FormatNumber(filemaxsize,2,-1) then 
			Response.Write "<script>alert('文件大小"&filesize&"M，超出限制！最大为"&filemaxsize&"M！');history.back();</script>"
			Response.End
		end if
	end if
	
	'上传文件
	filename2=requesta.Form("a_file1")
	if fileext<>"" then
		'删除原来的文件
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		objFSO.DeleteFile(Server.MapPath(picfld&filename2))
		set objFSO=nothing
		'上传新文件
		filename2=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"."&fileext
		call requesta.Save(filea,filename2)
	end if
	
	'添加
	if tj="Add" then
		rs.Open "music",conn,1,3
		rs.addnew
		rs("name1")=requesta.Form("name1")
		rs("url")=requesta.Form("url")
		rs("info")=requesta.Form("info")
		rs("file1")=filename2
		rs("laster")=session("admin")&"("&Session("truename")&")"
		rs("lasttime")=now()
		rs.update
		rs.close
		rsrz.open "book",conn,1,3                                   '添加日志              
		rsrz.addnew
		rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
		rsrz("caozuo")="添加"
		rsrz("menu")=menuname
		rsrz("duixian")=requesta.Form("name1")
		rsrz.update
		rsrz.close
	end if
	
	'修改
	if tj="Modify" then
		sql="select * from music where id="&requesta.Form("id")
		rs.Open sql,conn,1,3
		rs("name1")=requesta.Form("name1")
		rs("url")=requesta.Form("url")
		rs("info")=requesta.Form("info")
		rs("file1")=filename2
		rs("laster")=session("admin")&"("&Session("truename")&")"
		rs("lasttime")=now()
		rs.update
		rs.close
		rsrz.open "book",conn,1,3                                   '添加日志              
		rsrz.addnew
		rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
		rsrz("caozuo")="修改"
		rsrz("menu")=menuname
		rsrz("duixian")=requesta.Form("name1")
		rsrz.update
		rsrz.close
	end if
	%>
	
	<!--#include file="music_inc.asp"-->	
	
	<%	
	'返回
	response.redirect "musicManage.asp"
	Response.End
end if
%>

<body>
<!--#include file="top.asp"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23" align="center" background="images/am_mn2.gif" bgcolor="cccccc"><strong><%=menuname%></strong></td>
  </tr>
</table>
<div align="center">
<br />
<!---------------------------------------------添加界面------------------------------------------->
<%
if request("cz")="Add" then
%>
<strong>添加<%=menunamel%></strong><br>
<span class="red"><%=chicun_sm%></span><br><br>
<form action="?tj=Add" method="post" enctype="multipart/form-data" name="form2" id="form2">
<table border="1" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr>
    <td width="100" align="center" bgcolor="#F0F0F0"><strong><%=menunamel%>名称：</strong></td>
    <td align="left"><input name="name1" type="text" id="name1" class="input1" />
        <span class="red">*</span></td>
  </tr>
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong>上传文件：</strong></td>
    <td align="left"><input name="file1" type="file" /></td>
  </tr>
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong>外链文件：</strong></td>
    <td align="left"><input name="url" id="url" type="text" class="input1" /></td>
  </tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong>备注：</strong></td>
	  <td align="left"><input name="info" type="text" id="info" class="input1" /></td>
	</tr>
</table>
<!--#include file="inc/button.asp"-->
</form>
<%end if%>

<!---------------------------------------------修改界面------------------------------------------->
<% if  request("cz")="Modify" then %>
<%
sql="select * from music where id="&request("id")
rs.Open sql,conn,1,1
%>
<strong>修改<%=menunamel%></strong><br />
更新时间：<%=rs("lasttime")%>&nbsp;&nbsp;&nbsp;&nbsp;更新人：<%=rs("laster")%><br />
<span class="red"><%=chicun_sm%></span><br><br>
<form action="?tj=Modify" method="post" enctype="multipart/form-data" name="form3" id="form3">
  <input name="id" type="hidden" value="<%=request("id")%>" />
  <table border="1" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
    <tr>
      <td width="100" align="center" bgcolor="#F0F0F0"><strong><%=menunamel%>名称：</strong></td>
      <td align="left"><input name="name1" type="text" class="input1" id="name1" value="<%=rs("name1")%>" />
          <span class="red">*</span></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>已传文件：</strong></td>
      <td align="left">&nbsp;<%=rs("file1")%><input name="a_file1" type="hidden" value="<%=rs("file1")%>" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>重传文件：</strong></td>
      <td align="left"><input name="file1" type="file" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>外链文件：</strong></td>
      <td align="left"><input name="url" id="url" type="text" value="<%=rs("url")%>" class="input1" /></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>备注：</strong></td>
      <td align="left"><input name="info" type="text" class="input1" id="info" value="<%=rs("info")%>" /></td>
    </tr>
  </table>
  <!--#include file="inc/button.asp"-->
</form>
<%rs.close%>
<%end if%>
</div>
<!--#include file="foot.asp"-->
</body>
</html>