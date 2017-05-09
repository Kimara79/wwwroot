<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title>图片上传</title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="inc/inc_upload.asp"-->
<%
typeall="jpg,gif"
fzmax=1*1024*1024
%>
<!--#include file="inc/uploadcheck.asp"-->
<!--#include file="js/SetCwinHeight.js"-->

<%
schicun=split(session("schicun"),",")
sw=int(schicun(0))
sh=int(schicun(1))

On Error Resume Next
'添加
if request("tj")="Add" then
  set upload=new upload_5xSoft
  set filea=upload.file("pic")
	  if filea.FileSize<>0 then
	    'hztemp=split(filea.filename,".")
		'hz=lcase(hztemp(ubound(hztemp)))
		hz=right(filea.filename,3)
		hz=lcase(hz)
	    pic=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"."&hz
		filea.SaveAs Server.mappath("pic/"&session("folder")&pic)
		rsrz.open "book",conn,1,3             
		rsrz.addnew
		  rsrz("zhanghao")=session("admin")&"("&session("truename")&")"
		  rsrz("caozuo")="上传"
		  rsrz("menu")="图片上传"
		  rsrz("duixian")=pic
		  rsrz("time1")=now()
		rsrz.update
		rsrz.close
	  else
		 Response.Write("<script language=javascript>alert('请选择文件！');history.back()</script>")
		 Response.End
	  end if
  if session("pic")<>"" then
  	session("pic")=session("pic")&","&pic
  else
  	session("pic")=session("pic")&pic
  end if

	'生成缩略图
	Set Jpeg=Server.CreateObject("Persits.Jpeg")
	path1=Server.mappath("pic/"&session("folder")&pic)
	path2=Server.mappath("pic/s_"&session("folder")&pic)
	Jpeg.Open path1
	if sw>0 and sh>0 then
		Jpeg.Width=sw
		Jpeg.Height=sh
	end if
	if sw>0 and sh=0 then
		Jpeg.Width=sw
		Jpeg.Height=int((Jpeg.Width*sh)/sw)
	end if
	if sw=0 and sh>0 then
		Jpeg.Width=int((Jpeg.Height*sw)/sh)
		Jpeg.Height=sh
	end if
 	Jpeg.Save path2
 
  response.Write "<script>SetCwinHeight(top.picif);</script>"
  response.redirect "?"
end if

'删除
if request("tj")="Del" then
	session("pic")=replace(session("pic"),request("pic")&",","")
	session("pic")=replace(session("pic"),","&request("pic"),"")
	session("pic")=replace(session("pic"),request("pic"),"")
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")    '删除文件
	objFSO.DeleteFile Server.MapPath("pic/"&session("folder")&request("pic"))
	objFSO.DeleteFile Server.MapPath("pic/s_"&session("folder")&request("pic"))
	response.Write "<script>SetCwinHeight(top.picif);</script>"
	response.redirect "?"
end if
%>
<body>
<%if session("pic")<>"" then%>
	<%
	spic=split(session("pic"),",")
	for s=0 to ubound(spic)
	%>
	<a href="<%="pic/"&session("folder")&spic(s)%>" target="_blank"><img src="<%="pic/s_"&session("folder")&spic(s)%>" width="<%=int(sw/2)%>" height="<%=int(sh/2)%>" class="border1" /></a>&nbsp;&nbsp;
	<a href="?tj=Del&pic=<%=spic(s)%>" onClick="{if(confirm('确定要删除么?')){this.document.album.submit();return true;}return false;}">删除</a>&nbsp;&nbsp;
	<%next%>
<%end if%>
<%if int(ubound(spic)+1)<int(session("picnummax")) then%>
<form action="?tj=Add" method="post" enctype="multipart/form-data" name="formpic" id="formpic" style="margin:0px;" >
<input name="pic" type="file" size="60" onchange="filecheck(this)"/>&nbsp;&nbsp;&nbsp;&nbsp;
<input name="Submitpic" type="submit" id="Submitpic" value=" 添加 " />
</form>
<%end if%>
</body>
</html>