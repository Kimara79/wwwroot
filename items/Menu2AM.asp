<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="栏目管理" %>
<% menunamel="栏目" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<script type="text/javascript" src="../ckeditor/ckeditor.js"></script>
<script type="text/javascript" src="../ckfinder/ckfinder.js"></script>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="js/red.js"-->
<!--#include file="js/SetCwinHeight.js"-->

<%
chicun_sm="(栏目图标的宽度为：49，背景色为：fefef6)"
'传图提醒
sql="select info from system where id=19"
rsstr=trim(conn.execute(sql).getstring)
rsstr=replace(rsstr,chr(13),"")
rsstr=replace(rsstr,chr(10),"<br/>")
chicun_sm=chicun_sm&"<br/>"&rsstr

'显示位置
sql="select info from system where id=1"
rsstr=trim(conn.execute(sql).getstring)
rsstr=replace(rsstr,chr(13),"")
weizhi=split(rsstr,",")

'栏目类型
sql="select info from system where id=2"
rsstr=trim(conn.execute(sql).getstring)
rsstr=replace(rsstr,chr(13),"")
leixing=split(rsstr,",")

'其他设置
sql="select info from system where id=5"
rsstr=trim(conn.execute(sql).getstring)
rsstr=replace(rsstr,chr(13),"")
shezhi=split(rsstr,",")
%>

<%
'--------------------------------------------------------出错判断
if request("tj")="Add" or request("tj")="Modify" then
	'空判断
	if request.Form("name1")="" then 
		Response.Write("<script language=javascript>alert('栏目名称不能为空！');history.back();</script>")
		Response.End
	end if
 	'重复判断
    if request("tj")="Add" then
		sql="select count(id) from menu2 where name1='"&request.Form("name1")&"' "
	else
		sql="select count(id) from menu2 where name1='"&request.Form("name1")&"' and id<>"&request("id")&" "
	end if
	rsnum=int(conn.execute(sql).getstring)
	if rsnum>0 then
		Response.Write("<script language=javascript>alert('已存在相同的栏目名称！');history.back();</script>")
		Response.End
	end if
end if
%>

<% '--------------------------------------------------------添加
if request("tj")="Add" then
	rs.Open "menu2",conn,1,3
	rs.addnew
	On Error Resume Next
	For each obj in Request.Form
		rs(obj)=trim(Request.Form(obj))
	Next
	rs("pic")=session("pic")
	rs("lasttime")=now()
	rs("laster")=session("admin")&"("&session("truename")&")"
	rs.update
	rs.close
	
	rsrz.open "book",conn,1,3                                   '添加日志
	rsrz.addnew
	rsrz("zhanghao")=session("admin")&"("&session("truename")&")"
	rsrz("caozuo")="添加"
	rsrz("menu")=menuname
	rsrz("duixian")="二级栏目："&Request("name1")
	rsrz.update
	rsrz.close

	response.redirect "menuManage.asp"
	response.end
end if
%>

<% '--------------------------------------------------------修改
if request("tj")="Modify" then
	sql="select * from menu2 where id="&request("id")
	rs.Open sql,conn,1,3
	On Error Resume Next
	For each obj in Request.Form
		rs(obj)=trim(Request.Form(obj))
	Next
	rs("weizhi")=trim(Request.Form("weizhi"))
	rs("leixing")=trim(Request.Form("leixing"))
	rs("shezhi")=trim(Request.Form("shezhi"))
	rs("pic")=session("pic")
	rs("lasttime")=now()
	rs("laster")=session("admin")&"("&Session("truename")&")"
	rs.update
	rs.close
	
	rsrz.open "book",conn,1,3                                   '添加日志
	rsrz.addnew
	rsrz("zhanghao")=session("admin")&"("&session("truename")&")"
	rsrz("caozuo")="修改"
	rsrz("menu")=menuname
	rsrz("duixian")="二级栏目："&Request("name1")
	rsrz.update
	rsrz.close

	'改一级栏目
	if request.form("menu1")<>request.form("a_menu1") then
		sql="update article set menu1="&request.form("menu1")&" where menu2="&request("id")&" "
		conn.execute sql,1,1
		sql="update picture set menu1="&request.form("menu1")&" where menu2="&request("id")&" "
		conn.execute sql,1,1
		sql="update form1 set menu1="&request.form("menu1")&" where menu2="&request("id")&" "
		conn.execute sql,1,1
		sql="update ziduan set menu1="&request.form("menu1")&" where menu2="&request("id")&" "
		conn.execute sql,1,1
	end if
	
	response.redirect "menuManage.asp"
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
<div align="center">
<br />
<!---------------------------------------------添加界面------------------------------------>  
<%if  request("cz")="Add" then %>
<strong>添加二级<%=menunamel%></strong><br />
<span class="red"><%=chicun_sm%></span><br><br>
<form action="?tj=Add" method="post" name="form" id="form">
<table  border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr>
    <td align="center" bgcolor="#F0F0F0" class="STYLE1"><strong>上级<%=menunamel%>：</strong></td>
    <td align="left"><select name="menu1" id="menu1">
        <%
		sql="select * from menu1 order by no1"
		rs1.open sql,conn,1,1
		Do while not rs1.Eof
		%>
        <option value="<%=rs1("id")%>" <%if ""&rs1("id")=request("menu1") then%>selected="selected"<%end if%>><%=rs1("name1")%></option>
        <%
		rs1.MoveNext
		Loop
		rs1.close
		%>
      </select>
        <span class="red">*</span> </td>
  </tr>
	<tr>
		<td width="100" align="center" bgcolor="#F0F0F0" class="STYLE1"><strong><%=menunamel%>名称：</strong></td>
		<td width="900" align="left"><input name="name1" type="text" id="name1" size="30" /> <span class="red">*</span></td>
	</tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong>栏目类型：</strong></td>
	  <td align="left">
		<%for z=0 to ubound(leixing)%>
		 <label style="white-space:nowrap;<%if z=0 then%>color:#FF0000;<%end if%>">
		 <input type="radio" class="dq" name="leixing" id="leixing" value="<%=leixing(z)%>" onpropertychange="red(this)" <%if z=0 then%>checked="checked"<%end if%> />
		 <%=leixing(z)%></label>
		<%next%>
		<span class="red">*</span>		</td>
	</tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong>其他设置：</strong></td>
	  <td align="left">
		<%for z=0 to ubound(shezhi)%>
		 <label style="white-space:nowrap;">
		 <input type="checkbox" class="dq" name="shezhi" id="shezhi" value="<%=shezhi(z)%>" onpropertychange="red(this)" />
		 <%=shezhi(z)%></label>
		<%next%>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0F0" class="STYLE1"><strong>特别链接：</strong></td>
		<td align="left"><input name="url" type="text" id="url" size="80" /></td>
	</tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong><%=menunamel%>图标：</strong></td>
	  <td align="left">
	  <%
	   session("folder")="picture/"
	   session("schicun")="100,100"
	   session("picnummax")=1
	  %>
		  <iframe id="picif" src="Upload.asp" onload="SetCwinHeight(this)" name="main" width="100%" marginwidth="0" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
	</tr>
	<tr>
		<td align="center" valign="top" bgcolor="#F0F0F0" class="STYLE1"><strong><%=menunamel%>内容：</strong></td>
		<td align="left"><textarea name="info" class="ckeditor" id="info"></textarea></td>
	</tr>
</table>
<!--#include file="inc/button.asp"-->
</form>
<%end if%>

<!---------------------------------------------修改界面------------------------------------>			  
<%
if  request("cz")="Modify" then
sql="select * from menu2 where id="&request("id")
rs.Open sql,conn,1,1
%>
<strong>修改二级<%=menunamel%></strong><br />
<span class="red"><%=chicun_sm%></span><br><br>
<form action="?tj=Modify&id=<%=rs("id")%>" method="post" name="form" id="form">
<table border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr>
    <td align="center" bgcolor="#F0F0F0" class="STYLE1"><strong>上级<%=menunamel%>：</strong></td>
    <td align="left"><select name="menu1" id="menu1">
        <%
		sql="select * from menu1 order by no1"
		rs1.open sql,conn,1,1
		Do while not rs1.Eof
		%>
        <option value="<%=rs1("id")%>" <%if rs("menu1")=rs1("id") then%>selected="selected"<%end if%>><%=rs1("name1")%></option>
        <%
		rs1.MoveNext
		Loop
		rs1.close
		%>
      </select><input name="a_menu1" type="hidden" id="a_menu1" value="<%=rs("name1")%>" />
        <span class="red">*</span> </td>
  </tr>
	<tr>
		<td width="100" align="center" bgcolor="#F0F0F0" class="STYLE1"><strong><%=menunamel%>名称：</strong></td>
		<td width="900" align="left"><input name="name1" type="text" id="name1" value="<%=rs("name1")%>" size="30" /> <span class="red">*</span></td>
	</tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong>栏目类型：</strong></td>
	  <td align="left">
		<%for z=0 to ubound(leixing)%>
		 <label style="white-space:nowrap;<%if instr(""&rs("leixing"),leixing(z)) then%>color:#FF0000;<%end if%>">
		 <input type="radio" class="dq" name="leixing" id="leixing" value="<%=leixing(z)%>" onpropertychange="red(this)" <%if instr(""&rs("leixing"),leixing(z)) then%>checked="checked"<%end if%> />
		 <%=leixing(z)%></label>
		<%next%>
		<span class="red">*</span>		</td>
	</tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong>其他设置：</strong></td>
	  <td align="left">
		<%for z=0 to ubound(shezhi)%>
		 <label style="white-space:nowrap;<%if instr(""&rs("shezhi"),shezhi(z)) then%>color:#FF0000;<%end if%>">
		 <input type="checkbox" class="dq" name="shezhi" id="shezhi" value="<%=shezhi(z)%>" onpropertychange="red(this)" <%if instr(""&rs("shezhi"),shezhi(z)) then%>checked="checked"<%end if%> />
		 <%=shezhi(z)%></label>
		<%next%>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="#F0F0F0" class="STYLE1"><strong>特别链接：</strong></td>
		<td align="left"><input name="url" type="text" id="url" value="<%=rs("url")%>" size="80" /></td>
	</tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong><%=menunamel%>图标：</strong></td>
	  <td align="left">
	  <%
	   session("folder")="picture/"
	   session("pic")=rs("pic")       
	   session("schicun")="100,100"
	   session("picnummax")=1
	  %>
		  <iframe id="picif" src="Upload.asp" onload="SetCwinHeight(this)" name="main" width="100%" marginwidth="0" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
	</tr>
	<tr>
		<td align="center" valign="top" bgcolor="#F0F0F0" class="STYLE1"><strong><%=menunamel%>内容：</strong></td>
		<td align="left"><textarea name="info" class="ckeditor" id="info"><%=rs("info")%></textarea></td>
	</tr>
</table>
<!--#include file="inc/button.asp"-->
</form>
<%
rs.close
end if
%>
</div>
<!--#include file="foot.asp"-->
</body>
</html>