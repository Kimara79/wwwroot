<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="字段管理" %>
<% menunamel="字段" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="ziduan_str.asp"-->
<!--#include file="js/red.js"-->

<%
'--------------------------------------------------------出错判断
if request("tj")="Add" or request("tj")="Modify" then
	'空判断
	if request.Form("name1")="" then 
		Response.Write("<script language=javascript>alert('请填字段名称！');history.back();</script>")
		Response.End
	end if
	'上限判断
	sql="select count(id) from ziduan"
	rsnum=int(conn.execute(sql).getstring)
	if rsnum>=250 then 
		Response.Write("<script language=javascript>alert('字段数量已达上限！');history.back();</script>")
		Response.End
	end if
end if

'--------------------------------------------------------提交操作判断-添加
if request("tj")="Add" then
	rs.Open "ziduan",conn,1,3
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
	
	'id
	sql="select top 1 id from ziduan order by id desc "
	rsnum=int(conn.execute(sql).getstring)
	
	'关联添加字段
	select case request.Form("shuju")
		case "文本"
			sql="ALTER TABLE forms ADD COLUMN zd_"&rsnum&" Text("&request.Form("length_sj")&") "
		case "长整型"
			sql="ALTER TABLE forms ADD COLUMN zd_"&rsnum&" Long "
		case "双精度"
			sql="ALTER TABLE forms ADD COLUMN zd_"&rsnum&" Double "
		case "日期时间"
			sql="ALTER TABLE forms ADD COLUMN zd_"&rsnum&" Time "
		case "备注"
			sql="ALTER TABLE forms ADD COLUMN zd_"&rsnum&" Memo "
	end select
	conn.execute sql,1,1
	
	rsrz.open "book",conn,1,3                                   '添加日志              
	rsrz.addnew
	rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
	rsrz("caozuo")="添加"
	rsrz("menu")=menuname
	rsrz("duixian")=request.Form("name1")
	rsrz.update
	rsrz.close
	response.redirect "ziduanManage.asp"
	Response.End
end if
%>

<% '--------------------------------------------------------提交操作判断-修改
if request("tj")="Modify" then
	sql="select * from ziduan where id="&request("id")
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
	
	'关联修改字段
	if request.Form("shuju")<>request.Form("a_shuju") then
		select case request.Form("shuju")
			case "文本"
				sql="ALTER TABLE forms ALTER COLUMN zd_"&request("id")&" Text(50) "
			case "长整型"
				sql="ALTER TABLE forms ALTER COLUMN zd_"&request("id")&" Long "
			case "双精度"
				sql="ALTER TABLE forms ALTER COLUMN zd_"&request("id")&" Double "
			case "日期时间"
				sql="ALTER TABLE forms ALTER COLUMN zd_"&request("id")&" Time "
			case "备注"
				sql="ALTER TABLE forms ALTER COLUMN zd_"&request("id")&" Memo "
		end select
		conn.execute sql,1,1
	end if
	
	rsrz.open "book",conn,1,3                                   '添加日志              
	rsrz.addnew
	rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
	rsrz("caozuo")="修改"
	rsrz("menu")=menuname
	rsrz("duixian")=request.Form("name1")
	rsrz.update
	rsrz.close
	
	response.redirect "ziduanManage.asp"
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
<strong>添加<%=menunamel%></strong><br /><br />
<form action="?tj=Add" method="post" name="form2" id="form2">
<table border="1" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
	<tr>
	  <td width="100" align="center" bgcolor="#F0F0F0"><strong>所属栏目：</strong></td>
	  <td width="700" align="left">
		<select name="menu" id="menu">
		<%
		sql="select * from menu1 where (leixing like '%表单提交%' or id in (select menu1 from menu2 where leixing like '%表单提交%')) order by no1 "
		rs1.open sql,conn,1,1
		Do while not rs1.Eof
		%>
			<%if instr(rs1("leixing"),"表单提交") then%><option value="<%=rs1("id")%>" <%if ""&rs1("id")=session("ziduan_menu1") then%>selected="selected"<%end if%>><%=rs1("name1")%></option><%end if%>
			<%
			sql="select * from menu2 where leixing like '%表单提交%' and menu1="&rs1("id")&" order by no1"
			rs2.open sql,conn,1,1
			Do while not rs2.Eof
			%>
			<option value="<%=rs1("id")%>_<%=rs2("id")%>" <%if ""&rs2("id")=session("ziduan_menu2") then%>selected="selected"<%end if%>>--<%=rs2("name1")%></option>
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
      <td align="center" bgcolor="#F0F0F0"><strong>字段名称：</strong></td>
	  <td align="left"><input name="name1" type="text" id="name1" size="60" /> <span class="red">*</span></td>
	</tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong>数据类型：</strong></td>
	  <td align="left">
		<%for z=0 to ubound(shuju)%>
		 <label style="white-space:nowrap;<%if z=0 then%>color:#FF0000;<%end if%>">
		 <input type="radio" class="dq" name="shuju" id="shuju" value="<%=shuju(z)%>" onclick="red(this);" <%if z=0 then%>checked="checked"<%end if%> />
		 <%=shuju(z)%></label>
		<%next%>
		<span class="red">*</span>
	  </td>
	</tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong>控件类型：</strong></td>
	  <td align="left">
		<%for z=0 to ubound(kongjian)%>
		 <label style="white-space:nowrap;<%if z=0 then%>color:#FF0000;<%end if%>">
		 <input type="radio" class="dq" name="kongjian" id="kongjian" value="<%=kongjian(z)%>" onclick="red(this);" <%if z=0 then%>checked="checked"<%end if%> />
		 <%=kongjian(z)%></label>
		<%next%>
		<span class="red">*</span>
	  </td>
	</tr>
	<tr>
      <td align="center" bgcolor="#F0F0F0"><strong>数据长度：</strong></td>
	  <td align="left"><input name="length_sj" type="text" id="length_sj" value="50" size="60" /> 
	  (文本控件时)</td>
	</tr>
	<tr>
      <td align="center" bgcolor="#F0F0F0"><strong>控件宽度：</strong></td>
	  <td align="left"><input name="length_kj" type="text" id="length_kj" size="60" /> 
	    (页面显示的宽度，单位px) </td>
	</tr>
	<tr>
      <td align="center" bgcolor="#F0F0F0"><strong>控件注释：</strong></td>
	  <td align="left"><input name="zhushi" type="text" id="zhushi" size="60" /> 
	  (如数据的单位、格式等) </td>
	</tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong>其他设置：</strong></td>
	  <td align="left">
		<%for z=0 to ubound(shezhi)%>
		 <label style="white-space:nowrap;">
		 <input type="checkbox" class="dq" name="shezhi" id="shezhi" value="<%=shezhi(z)%>" onpropertychange="red(this)" />
		 <%=shezhi(z)%></label>
		<%next%></td>
	</tr>
	<tr>
      <td align="center" bgcolor="#F0F0F0"><strong>控件选项：</strong><br>(每行一个)</td>
	  <td align="left">
	    <textarea name="info" cols="60" id="info" style="height:expression(this.scrollHeight);padding:10px;"></textarea></td>
	</tr>
</table>
<!--#include file="inc/button.asp"-->
</form>
      <%end if%>
<!---------------------------------------------修改界面------------------------------------------->
<% if  request("cz")="Modify" then %>
<%
 sql="select * from ziduan where id="&request("id")
 rs.Open sql,conn,1,3
%>
<strong>修改<%=menunamel%></strong><br />
更新时间：<%=rs("lasttime")%>&nbsp;&nbsp;&nbsp;&nbsp;更新人：<%=rs("laster")%><br /><br />
<form action="?id=<%=rs("id")%>&amp;tj=Modify" method="post" name="form3" id="form3">
  <table border="1" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
	<tr>
	  <td width="100" align="center" bgcolor="#F0F0F0"><strong>所属栏目：</strong></td>
	  <td width="700" align="left">
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
	  <td align="center" bgcolor="#F0F0F0"><strong>字段名称：</strong></td>
	  <td align="left"><input name="name1" type="text" id="name1" value="<%=rs("name1")%>" size="60" /> <span class="red">*</span></td>
	</tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong>数据类型：</strong></td>
	  <td align="left">
		<%for z=0 to ubound(shuju)%>
		 <label style="white-space:nowrap;<%if rs("shuju")=shuju(z) then%>color:#FF0000;<%end if%>">
		 <input type="radio" class="dq" name="shuju" id="shuju" value="<%=shuju(z)%>" onpropertychange="red(this)" <%if rs("shuju")=shuju(z) then%>checked="checked"<%end if%> />
		 <%=shuju(z)%></label>
		<%next%>
		<input name="a_shuju" type="hidden" id="a_shuju" value="<%=rs("shuju")%>" />
		<span class="red">*</span>
	  </td>
	</tr>
	<tr>
	  <td align="center" bgcolor="#F0F0F0"><strong>控件类型：</strong></td>
	  <td align="left">
		<%for z=0 to ubound(kongjian)%>
		 <label style="white-space:nowrap;<%if rs("kongjian")=kongjian(z) then%>color:#FF0000;<%end if%>">
		 <input type="radio" class="dq" name="kongjian" id="kongjian" value="<%=kongjian(z)%>" onpropertychange="red(this)" <%if rs("kongjian")=kongjian(z) then%>checked="checked"<%end if%> />
		 <%=kongjian(z)%></label>
		<%next%>
		<span class="red">*</span>
	  </td>
	</tr>
	<tr>
      <td align="center" bgcolor="#F0F0F0"><strong>数据长度：</strong></td>
	  <td align="left"><input name="length_sj" type="text" id="length_sj" size="60" value="<%=rs("length_sj")%>" /> (文本控件时填)</td>
	</tr>
	<tr>
      <td align="center" bgcolor="#F0F0F0"><strong>控件宽度：</strong></td>
	  <td align="left"><input name="length_kj" type="text" id="length_kj" size="60" value="<%=rs("length_kj")%>" /> 
	    (页面显示的宽度，单位px) </td>
	</tr>
	<tr>
      <td align="center" bgcolor="#F0F0F0"><strong>控件注释：</strong></td>
	  <td align="left"><input name="zhushi" type="text" id="zhushi" size="60" value="<%=rs("zhushi")%>" /> 
	  (如数据的单位、格式等) </td>
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
      <td align="center" bgcolor="#F0F0F0"><strong>控件选项：</strong><br>(每行一个)</td>
	  <td align="left">
	    <textarea name="info" cols="60" id="info" style="height:expression(this.scrollHeight);padding:10px;"><%=rs("info")%></textarea></td>
	</tr>
  </table>
  <!--#include file="inc/button.asp"-->
  <input type="button" name="Submit3d2" value="添加为新记录" onclick="javascript:document.form3.action='?tj=Add';document.form3.submit();" /><br><br>
</form>
<%rs.close%>
<%end if%>
</div>
<!--#include file="foot.asp"-->
</body>
</html>