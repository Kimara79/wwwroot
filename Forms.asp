<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<!--#include file="cokk.asp"-->
<%tb="forms"%>
<!--#include file="title.asp"-->
<title><%=title1%> - <%=site_title%></title>
</head>

<!--#include file="inc/chkMail.asp"-->
<!--#include file="items/js/red.js"-->

<%
if request("tj")="Add" then
	'必填判断
	sql="select * from ziduan where "&tb&"="&id&" and shezhi like '%必填%' order by no1 "
	rs.open sql,conn,1,1
	Do while not rs.Eof
		if request.Form("zd_"&rs("id"))="" then
			Response.Write("<script language=javascript>alert('请填“"&rs("name1")&"”！');history.back();</script>")
			Response.End
		end if
	rs.MoveNext
	Loop
	rs.close
	
	'邮箱判断
	sql="select * from ziduan where "&tb&"="&id&" and shezhi like '%邮箱%' order by no1 "
	rs.open sql,conn,1,1
	Do while not rs.Eof
		if chkMail(request.Form("zd_"&rs("id"))) then
		else
			Response.Write("<script language=javascript>alert('“"&rs("name1")&"”格式错误！');history.back();</script>")
			Response.End
		end if
	rs.MoveNext
	Loop
	rs.close
	
	'验证码判断
	if request.Form("code")="" then 
		Response.Write("<script language=javascript>alert('请填验证码！');history.back()</script>")
		Response.End
	end if
	if request.Form("code")<>request.Form("code2") then 
		Response.Write("<script language=javascript>alert('验证码错误！');history.back()</script>")
		Response.End
	end if
	
	'入库
	rs.Open "forms",conn,1,3
	rs.addnew
	On Error Resume Next
	For each obj in request.Form
		rs(obj)=trim(request.Form(obj))
	Next
	rs("menu1")=menu1_id
	if menu2_id<>"" then
		rs("menu2")=menu2_id
	else
		rs("menu2")=0
	end if
	rs("laster")="前台"
	rs("lasttime")=now()
	rs.update
	rs.close
	
	Response.Write "<script>alert('提交成功！通过审核后将会显示。');location='?menu1="&request("menu1")&"&menu2="&request("menu2")&"';</script>"
	Response.End
end if
%>

<body>
<!--#include file="top.asp"-->

<center>
<table align="center" border="0" cellspacing="0" cellpadding="0" class="bodyw">
  <tr>
    <td align="center" valign="top">
	<div class="bga_01"><%if menu2_name1<>"" then%><%=menu2_name1%><%else%><%=menu1_name1%><%end if%></div>
	<div class="bga_02" align="center">
		<div style="width:94%;">
		<div style="height:30px;"></div>
	  
	<form id="forms" name="forms" method="post" action="?tj=Add&menu1=<%=request("menu1")%>&menu2=<%=request("menu2")%>">
		  <table border="0" cellpadding="5" cellspacing="0">
			<%
			sql="select * from ziduan where "&tb&"="&id&" order by no1 "
			rs.open sql,conn,1,1
			Do while not rs.Eof
				if rs("info")<>"" then
					tempstr=rs("info")
					tempstr=replace(tempstr,chr(10),"")
					tempstr=split(tempstr,chr(13))
				end if
			%>
			  <tr>
				<td align="right" valign="top"><%=rs("name1")%>：</td>
				<td align="left" class="text1">
				<%select case rs("kongjian")%>
					<%case "填空"%>
						<input name="zd_<%=rs("id")%>" id="zd_<%=rs("id")%>" type="text" class="fdbw1" />
					<%case "单选"%>
						<%for z=0 to ubound(tempstr)%>
						<label><input name="zd_<%=rs("id")%>" id="zd_<%=rs("id")%>" type="radio" value="<%=tempstr(z)%>" onpropertychange="red(this)" /> <%=tempstr(z)%></label>
						<%next%>
					<%case "多选"%>
						<%for z=0 to ubound(tempstr)%>
						<label><input name="zd_<%=rs("id")%>" id="zd_<%=rs("id")%>" type="checkbox" value="<%=tempstr(z)%>" onpropertychange="red(this)" /> <%=tempstr(z)%></label>
						<%next%>
					<%case "下拉"%>
						<select name="zd_<%=rs("id")%>" id="zd_<%=rs("id")%>">
						<% for z=0 to ubound(tempstr) %>
						<option value="<%=tempstr(z)%>"><%=tempstr(z)%></option>
						<%next%>
						</select>
					<%case "段落"%>
						<textarea name="zd_<%=rs("id")%>" id="zd_<%=rs("id")%>" class="fdbw2 textarea_none"></textarea>
				<%end select%>
				</td>
				<td align="left"><%if instr(""&rs("shezhi"),"必填") then%><span style="color:#FF0000">*</span><%end if%></td>
			  </tr>
			<%
			rs.MoveNext
			Loop
			rs.close
			%>
			  <tr>
				<td align="right">验证码：</td>
				<td align="left"><input name="code" type="text" class="xh" id="code" size="4" onKeyUp="if(event.keyCode !=37 && event.keyCode != 39) value=value.replace(/\D/g,'');" onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/\D/g,''))" maxlength="4" onKeyDown="if (event.keyCode==13) {document.form.submit()}" />
				<%
				randomize timer
				code=Int((8999)*Rnd +1000)
				%>
				&nbsp;&nbsp;<b><%=code%></b>
				<input name="code2" type="hidden" id="code2" value="<%=code%>" />
				&nbsp;&nbsp;<span style="color:#FF0000">*</span></td>
				<td align="left"></td>
			  </tr>
		  </table>
		  <br />
		  <input name="Submit" type="submit" class="bt1" value="提交" />
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		  <input name="Submit" type="reset" class="bt1" value="重置" />
		  <br /><br />
	</form>
	
	
	<!----留言列表----->
	<%
	sql="select count(id) from "&tb&" where shezhi like '%显示留言%'"
	rsnum=int(conn.execute(sql).getstring)
	if rsnum>0 then
	%>
		<%
		i=0
		sql="select * from forms where "&tb&"="&id&" and shezhi like '%已审核%' order by id desc"
		rs.open sql,conn,1,1
		if rs.recordcount>0 then
			rs.pagesize=10  '设置每页显示数目
			currentpage=Clng(request("page"))
			if currentpage<1 then currentpage=1
			if currentpage>rs.pagecount then currentpage=rs.pagecount
			rs.absolutepage=currentpage
		%>
			<br />
			<div style="width:80%;">
			<div class="border_b"></div>
			<%
			sql="select * from ziduan where "&tb&"="&id&" and shezhi like '%前台列表页显示%' order by no1"
			rs1.open sql,conn,1,1
			Do while not rs.Eof
			%>
				<div style="padding:10px;" align="left">
					<img src="images/book.gif" align="absmiddle" />&nbsp;<%=rs("lasttime")%><br />
					<%
					rs1.movefirst
					Do while not rs1.Eof
					%>
					<span style="color:#ff6600;font-weight:bold;"><%=rs1("name1")%></span>：&nbsp;<%=rs("zd_"&rs1("id"))%><br />
					<%
					rs1.MoveNext
					loop
					%>
				</div>
				<div class="border_b"></div>
			<%
			i=i+1
			rs.MoveNext
			If i>=rs.pagesize Then Exit Do
			Loop
			rs1.close
			%>
			<%danwei="条"%>
			<!--#include file="inc/page.asp"-->
			</div>
		<%
		end if
		rs.close
		%>
	<%end if%>
	

		</div>
	</div>
	<div class="bga_03"></div>
	</td>
  </tr>
</table>
</center>

<!--#include file="foot.asp"-->
</body>
</html>