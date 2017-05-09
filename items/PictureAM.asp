<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="图片管理" %>
<% menunamel="图片" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<script type="text/javascript" src="../ckeditor/ckeditor.js"></script>
<script type="text/javascript" src="../ckfinder/ckfinder.js"></script>

<link rel="stylesheet" href="js_color_picker_v2.css" media="screen">
<script type="text/javascript" src="color_functions.js"></script>		
<script type="text/javascript" src="js_color_picker_v2.js"></script>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="js/red.js"-->
<!--#include file="js/SetCwinHeight.js"-->

<%
chicun_sm="(首页大图的最佳尺寸是：524*375)"
'传图提醒
sql="select info from system where id=19"
rsstr=trim(conn.execute(sql).getstring)
rsstr=replace(rsstr,chr(13),"")
rsstr=replace(rsstr,chr(10),"<br/>")
chicun_sm=chicun_sm&"<br/>"&rsstr

'--------------------------------------------------------出错判断
if request("tj")="Add" or request("tj")="Modify" then
	'空判断
	if session("pic")="" then 
		Response.Write("<script language=javascript>alert('请上传图片！');history.back();</script>")
		Response.End
	end if
end if

'--------------------------------------------------------提交操作判断-添加
if request("tj")="Add" then
	rs.Open "picture",conn,1,3
	rs.addnew
	On Error Resume Next
	For each obj in request.Form
		rs(obj)=request.Form(obj)
	Next
	rs("pic")=session("pic")
	if instr(request.Form("menu"),"_") then
		temp=split(request.Form("menu"),"_")
		rs("menu1")=temp(0)
		rs("menu2")=temp(1)
	else
		rs("menu1")=request.Form("menu")
		rs("menu2")=0
	end if
	rs("laster")=session("admin")&"("&Session("truename")&")"
	rs("lasttime")=now()
	rs.update
	rs.close
	
	rsrz.open "book",conn,1,3                                   '添加日志              
	rsrz.addnew
	rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
	rsrz("caozuo")="添加"
	rsrz("menu")=menuname
	rsrz("duixian")=request.Form("name1")
	rsrz.update
	rsrz.close
	response.redirect "pictureManage.asp"
	Response.End
end if
%>

<% '--------------------------------------------------------提交操作判断-修改
if request("tj")="Modify" then
	sql="select * from picture where id="&request("id")
	rs.Open sql,conn,1,3
	On Error Resume Next
	For each obj in request.Form
		rs(obj)=request.Form(obj)
	Next
	rs("pic")=session("pic")
	if instr(request.Form("menu"),"_") then
		temp=split(request.Form("menu"),"_")
		rs("menu1")=temp(0)
		rs("menu2")=temp(1)
	else
		rs("menu1")=request.Form("menu")
		rs("menu2")=0
	end if
	rs("laster")=session("admin")&"("&Session("truename")&")"
	rs("lasttime")=now()
	rs.update
	rs.close
	
	rsrz.open "book",conn,1,3                                   '添加日志              
	rsrz.addnew
	rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
	rsrz("caozuo")="修改"
	rsrz("menu")=menuname
	rsrz("duixian")=request.Form("name1")
	rsrz.update
	rsrz.close
	
	response.redirect "pictureManage.asp"
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
<%if  request("cz")="Add" then %>
<strong>添加<%=menunamel%></strong><br>
<span class="red"><%=chicun_sm%></span><br><br>
<form action="?tj=Add" method="post" name="form2" id="form2">
<table border="1" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
	<tr>
	  <td width="120" align="center" bgcolor="#F0F0F0"><strong>所属栏目：</strong></td>
	  <td width="900" align="left">
		<select name="menu" id="menu">
		<%
		sql="select * from menu1 where (leixing like '%多张图片%' or id in (select menu1 from menu2 where leixing like '%多张图片%')) order by no1 "
		rs1.open sql,conn,1,1
		Do while not rs1.Eof
		%>
			<%if instr(rs1("leixing"),"多张图片") then%><option value="<%=rs1("id")%>" <%if ""&rs1("id")=session("picture_menu1") then%>selected="selected"<%end if%>><%=rs1("name1")%></option><%end if%>
			<%
			sql="select * from menu2 where leixing like '%多张图片%' and menu1="&rs1("id")&" order by no1"
			rs2.open sql,conn,1,1
			Do while not rs2.Eof
				sql="select name1 from menu1 where id="&rs2("menu1")&" "
				rsstr=trim(conn.execute(sql).getstring)
				rsstr=replace(""&rsstr,chr(13),"")
			%>
			<option value="<%=rs1("id")%>_<%=rs2("id")%>" <%if ""&rs2("id")=session("picture_menu2") then%>selected="selected"<%end if%>><%=rsstr%>--<%=rs2("name1")%></option>
			<%
			rs2.MoveNext
			Loop
			rs2.close
			%>
		<%
		rs1.MoveNext
		Loop
		rs1.close
		%>
		</select>
		<span class="red">*</span>	  </td>
	</tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong>图片名称：</strong></td>
	  <td align="left"><input name="name1" type="text" id="name1" size="60"/></td>
	</tr>
	<tr>
      <td align="center" bgcolor="#F0F0F0"><strong>名称颜色：</strong></td>
	  <td align="left"><input name="color" type="text" id="color" size="60"/>&nbsp;&nbsp;<input type="button" value="选择颜色" onclick="showColorPicker(this,color)"></td>
	  </tr>
	<tr>
      <td align="center" bgcolor="#F0F0F0"><strong>点击次数：</strong></td>
	  <td align="left"><input name="click" type="text" id="click" value="80" size="60"/></td>
	</tr>
	<tr>
      <td align="center" bgcolor="#F0F0F0" class="STYLE1"><strong>特别链接：</strong></td>
	  <td align="left"><input name="url" type="text" id="url" size="80" /></td>
	  </tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong>上传图片：</strong></td>
	  <td align="left">
	  <%
	   session("folder")="picture/"
	   session("schicun")="100,130"
	   session("picnummax")=1
	  %>
		  <iframe id="picif" src="Upload.asp" onload="SetCwinHeight(this)" name="main" width="100%" marginwidth="0" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
	</tr>
	<tr>
      <td align="center" bgcolor="#F0F0F0" class="STYLE1"><strong>图片简介：</strong></td>
	  <td align="left"><textarea name="jianjie" cols="80" rows="4" id="jianjie"></textarea></td>
	  </tr>
	<tr>
      <td align="center" valign="top" bgcolor="#F0F0F0" class="STYLE1"><strong>图片说明：</strong></td>
	  <td align="left"><textarea name="info" class="ckeditor" id="info"></textarea></td>
	</tr>
</table>
<!--#include file="inc/button.asp"-->
</form>
      <%end if%>
<!---------------------------------------------修改界面------------------------------------------->
<% if  request("cz")="Modify" then %>
<%
 sql="select * from picture where id="&request("id")
 rs.Open sql,conn,1,3
%>
<strong>修改<%=menunamel%></strong><br />
更新时间：<%=rs("lasttime")%>&nbsp;&nbsp;&nbsp;&nbsp;更新人：<%=rs("laster")%>
<br>
<span class="red"><%=chicun_sm%></span><br />
<br />
<form action="?id=<%=rs("id")%>&amp;tj=Modify" method="post" name="form3" id="form3">
  <table border="1" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
	<tr>
	  <td width="120" align="center" bgcolor="#F0F0F0"><strong>所属栏目：</strong></td>
	  <td width="900" align="left">
		<select name="menu" id="menu">
		<%
		sql="select * from menu1 where (leixing like '%多张图片%' or id in (select menu1 from menu2 where leixing like '%多张图片%')) order by no1 "
		rs1.open sql,conn,1,1
		Do while not rs1.Eof
		%>
			<%if instr(rs1("leixing"),"多张图片") then%><option value="<%=rs1("id")%>" <%if rs1("id")=rs("menu1") then%>selected="selected"<%end if%>><%=rs1("name1")%></option><%end if%>
			<%
			sql="select * from menu2 where leixing like '%多张图片%' and menu1="&rs1("id")&" order by no1"
			rs2.open sql,conn,1,1
			Do while not rs2.Eof
				sql="select name1 from menu1 where id="&rs2("menu1")&" "
				rsstr=trim(conn.execute(sql).getstring)
				rsstr=replace(""&rsstr,chr(13),"")
			%>
			<option value="<%=rs1("id")%>_<%=rs2("id")%>" <%if rs2("id")=rs("menu2") then%>selected="selected"<%end if%>><%=rsstr%>--<%=rs2("name1")%></option>
			<%
			rs2.MoveNext
			Loop
			rs2.close
			%>
		<%
		rs1.MoveNext
		Loop
		rs1.close
		%>
		</select>
		<span class="red">*</span>	  </td>
	</tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong>图片名称：</strong></td>
	  <td align="left"><input name="name1" type="text" id="name1" value="<%=rs("name1")%>" size="60"/></td>
	</tr>
	<tr>
      <td align="center" bgcolor="#F0F0F0"><strong>名称颜色：</strong></td>
	  <td align="left"><input name="color" type="text" id="color" value="<%=rs("color")%>" size="60"/>&nbsp;&nbsp;<input type="button" value="选择颜色" onclick="showColorPicker(this,color)"></td>
	  </tr>
	<tr>
      <td align="center" bgcolor="#F0F0F0"><strong>点击次数：</strong></td>
	  <td align="left"><input name="click" type="text" id="click" value="<%=rs("click")%>" size="60"/></td>
	  </tr>
	<tr>
      <td align="center" bgcolor="#F0F0F0" class="STYLE1"><strong>特别链接：</strong></td>
	  <td align="left"><input name="url" type="text" id="url" value="<%=rs("url")%>" size="80" /></td>
	  </tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong>上传图片：</strong></td>
	  <td align="left">
	  <%
	   session("folder")="picture/"
	   session("pic")=rs("pic")       
	   session("schicun")="100,130"
	   session("picnummax")=1
	  %>
		  <iframe id="picif" src="Upload.asp" onload="SetCwinHeight(this)" name="main" width="100%" marginwidth="0" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
	</tr>
	<tr>
      <td align="center" bgcolor="#F0F0F0" class="STYLE1"><strong>图片简介：</strong></td>
	  <td align="left"><textarea name="jianjie" cols="80" rows="4" id="jianjie"><%=rs("jianjie")%></textarea></td>
	  </tr>
	<tr>
	  <td align="center" valign="top" bgcolor="#F0F0F0" class="STYLE1"><strong>图片说明：</strong></td>
	  <td align="left"><textarea name="info" class="ckeditor" id="info"><%=rs("info")%></textarea></td>
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