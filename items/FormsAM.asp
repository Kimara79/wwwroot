<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="表单管理" %>
<% menunamel="表单" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="js/red.js"-->

<%
'--------------------------------------------------------出错判断
if request("tj")="Add" or request("tj")="Modify" then
	'空判断
	if request.Form("name1")="" then 
		Response.Write("<script language=javascript>alert('请填表单标题！');history.back();</script>")
		Response.End
	end if
end if

'--------------------------------------------------------提交操作判断-添加
if request("tj")="Add" then
	rs.Open "forms",conn,1,3
	rs.addnew
	On Error Resume Next
	For each obj in request.Form
		rs(obj)=request.Form(obj)
	Next
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
	rsrz("duixian")=request.form("menu")
	rsrz.update
	rsrz.close
	response.redirect "formsManage.asp"
	Response.End
end if
%>

<% '--------------------------------------------------------提交操作判断-修改
if request("tj")="Modify" then
	sql="select * from forms where id="&request("id")
	rs.Open sql,conn,1,3
	On Error Resume Next
	For each obj in request.Form
		rs(obj)=request.Form(obj)
	Next
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
	rsrz("duixian")=request.form("menu")
	rsrz.update
	rsrz.close
	
	response.redirect "formsManage.asp"
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
<strong>添加<%=menunamel%></strong>
<form action="?tj=Add" method="post" name="form2" id="form2">
<table border="1" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
	<tr>
	  <td width="100" align="center" bgcolor="#F0F0F0"><strong>所属栏目：</strong></td>
	  <td width="900" align="left">
		<select name="menu" id="menu">
		<%
		sql="select * from menu1 where (leixing like '%表单提交%' or id in (select menu1 from menu2 where leixing like '%表单提交%')) order by no1 "
		rs1.open sql,conn,1,1
		Do while not rs1.Eof
		%>
			<%if instr(rs1("leixing"),"表单提交") then%><option value="<%=rs1("id")%>" <%if ""&rs1("id")=session("forms_menu1") then%>selected="selected"<%end if%>><%=rs1("name1")%></option><%end if%>
			<%
			sql="select * from menu2 where leixing like '%表单提交%' and menu1="&rs1("id")&" order by no1"
			rs2.open sql,conn,1,1
			Do while not rs2.Eof
			%>
			<option value="<%=rs1("id")%>_<%=rs2("id")%>" <%if ""&rs2("id")=session("forms_menu2") then%>selected="selected"<%end if%>>--<%=rs2("name1")%></option>
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
	  <td align="center" bgcolor="#F0F0F0"><strong>表单标题：</strong></td>
	  <td align="left"><input name="name1" type="text" id="name1" size="60"/></td>
	</tr>
</table>
<!--#include file="inc/button.asp"-->
</form>
      <%end if%>
<!---------------------------------------------修改界面------------------------------------------->
<% if  request("cz")="Modify" then %>
<%
 sql="select * from forms where id="&request("id")
 rs.Open sql,conn,1,3
%>
<strong>修改<%=menunamel%></strong><br />
更新时间：<%=rs("lasttime")%>&nbsp;&nbsp;&nbsp;&nbsp;更新人：<%=rs("laster")%><br /><br />
<form action="?id=<%=rs("id")%>&amp;tj=Modify" method="post" name="form3" id="form3">
  <table border="1" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
	<tr>
	  <td width="100" align="center" bgcolor="#F0F0F0"><strong>所属栏目：</strong></td>
	  <td width="900" align="left">
		<select name="menu" id="menu">
		<%
		sql="select * from menu1 where (leixing like '%表单提交%' or id in (select menu1 from menu2 where leixing like '%表单提交%')) order by no1 "
		rs1.open sql,conn,1,1
		Do while not rs1.Eof
		%>
			<%if instr(rs1("leixing"),"表单提交") then%><option value="<%=rs1("id")%>" <%if rs1("id")=rs("menu1") then%>selected="selected"<%end if%>><%=rs1("name1")%></option><%end if%>
			<%
			sql="select * from menu2 where leixing like '%表单提交%' and menu1="&rs1("id")&" order by no1"
			rs2.open sql,conn,1,1
			Do while not rs2.Eof
			%>
			<option value="<%=rs1("id")%>_<%=rs2("id")%>" <%if rs2("id")=rs("menu2") then%>selected="selected"<%end if%>>--<%=rs2("name1")%></option>
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
	  <td align="center" bgcolor="#F0F0F0"><strong>表单标题：</strong></td>
	  <td align="left"><input name="name1" type="text" id="name1" value="<%=rs("name1")%>" size="60"/></td>
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