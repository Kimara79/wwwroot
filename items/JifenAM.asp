<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="积分兑换" %>
<% menunamel="记录" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="js/fbjs.asp"-->

<%
'积分转换比例
sql="select info from system where id=25"
rsstr=trim(conn.execute(sql).getstring)
zhh=replace(rsstr,chr(13),"")
zhh=split(zhh,":")
%>

<script language="JavaScript">
function jifenjs1(){
	document.getElementById("users_name1").src="users_name1.asp?users_card="+document.getElementById("users_card").value;
}
function jifenjs2(){
	document.getElementById("priceall").value=document.getElementById("price").value*document.getElementById("num").value;
}
function jifenjs3(){
	document.getElementById("score").value=parseInt(document.getElementById("priceall").value*<%=zhh(1)%>/<%=zhh(0)%>,10);
}
</script>

<%
'--------------------------------------------------------出错判断
if request("tj")="Add" or request("tj")="Modify" then
	'空判断
	if request.Form("users_card")="" then 
		Response.Write("<script language=javascript>alert('会员卡号不能为空！');history.back();</script>")
		Response.End
	end if
	if request.Form("name1")="" then 
		Response.Write("<script language=javascript>alert('兑换商品不能为空！');history.back();</script>")
		Response.End
	end if
	if request.Form("price")="" then 
		Response.Write("<script language=javascript>alert('商品价值不能为空！');history.back();</script>")
		Response.End
	end if
	if request.Form("num")="" then 
		Response.Write("<script language=javascript>alert('兑换数量不能为空！');history.back();</script>")
		Response.End
	end if
	if request.Form("priceall")="" then 
		Response.Write("<script language=javascript>alert('合计价值不能为空！');history.back();</script>")
		Response.End
	end if
	if request.Form("score")="" then 
		Response.Write("<script language=javascript>alert('所需积分不能为空！');history.back();</script>")
		Response.End
	end if
	'积分判断
	sql="select score from users where card="&request.Form("users_card")&" "
	score=int(conn.execute(sql).getstring)
	if score<int(request.Form("score")) then
		Response.Write("<script language=javascript>alert('积分不足！');history.back();</script>")
		Response.End
	end if
	'卡号判断
	On Error Resume Next
	sql="select id from users where card="&request.Form("users_card")&" "
	usersid=int(conn.execute(sql).getstring)
	if request("tj")="Modify" then
		sql="select id from users where card="&request.Form("a_users_card")&" "
		a_usersid=int(conn.execute(sql).getstring)
	end if
end if

'--------------------------------------------------------提交操作判断-添加
if request("tj")="Add" then
	'会员减积分
	sql="update users set score=score-"&request.Form("score")&" where id="&usersid&" "
	conn.execute sql,1,1
	'剩余积分
	sql="select score from users where id="&usersid&" "
	score2=int(conn.execute(sql).getstring)

	rs.Open "jifen",conn,1,3
	rs.addnew
	On Error Resume Next
	For each obj in request.Form
		rs(obj)=request.Form(obj)
	Next
	rs("time1")=request.Form("time1_y")&"-"&request.Form("time1_m")&"-"&request.Form("time1_d")
	rs("usersid")=usersid
	rs("score2")=score2
	rs("addtime")=now()
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
	response.redirect "jifenManage.asp"
	Response.End
end if
%>

<% '--------------------------------------------------------提交操作判断-修改
if request("tj")="Modify" then
	'修改卡号或积分
	sql="update users set score=score+"&request.Form("a_score")&" where id="&a_usersid&" "
	conn.execute sql,1,1 
	sql="update users set score=score-"&request.Form("score")&" where id="&usersid&" "
	conn.execute sql,1,1
	
	sql="select * from jifen where id="&request("id")
	rs.Open sql,conn,1,3
	On Error Resume Next
	For each obj in request.Form
		rs(obj)=request.Form(obj)
	Next
	rs("time1")=request.Form("time1_y")&"-"&request.Form("time1_m")&"-"&request.Form("time1_d")
	rs("usersid")=usersid
	rs("score2")=rs("score2")+request.Form("a_score")-request.Form("score")
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
	
	response.redirect "jifenManage.asp"
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
	rsstr=""
	if request.QueryString("usersid")<>"" then
		sql="select card from users where id="&request.QueryString("usersid")&" "
		rsstr=trim(conn.execute(sql).getstring)
		rsstr=replace(""&rsstr,chr(13),"")
	end if
%>
<strong>添加<%=menunamel%></strong><br><br>
<form action="?tj=Add" method="post" name="form2" id="form2">
<table border="1" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr>
    <td width="100" align="center" bgcolor="#F0F0F0"><strong>兑换日期：</strong></td>
    <td align="left">
	<select name="time1_y" id="time1_y">
	  <%for i=2013 to year(now)+1%>
      <option value="<%=i%>" <%if i=year(now) then%>selected="selected"<%end if%>><%=i%></option>
	  <%next%>
    </select>
	年
	<select name="time1_m" id="time1_m">
	  <%for i=1 to 12%>
      <option value="<%=i%>" <%if i=month(now) then%>selected="selected"<%end if%>><%=i%></option>
	  <%next%>
    </select>
	月
	<select name="time1_d" id="time1_d">
	  <%for i=1 to 31%>
      <option value="<%=i%>" <%if i=day(now) then%>selected="selected"<%end if%>><%=i%></option>
	  <%next%>
    </select>
	日
    <span class="red">*</span></td>
  </tr>
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong>会员卡号：</strong></td>
    <td align="left">
	<table border="0" cellspacing="0" cellpadding="0">
	  <tr>
		<td>
		<input name="users_card" id="users_card" type="text" size="20" value="<%=rsstr%>" onKeyUp="<%=fbjs_rnum1%>jifenjs1();" />
		<span class="red">*</span>&nbsp;&nbsp;&nbsp;&nbsp;
		</td>
		<td valign="middle">
		<iframe id="users_name1" name="users_name1" frameborder="0" width="230" height="23" scrolling="no" <%if rsstr<>"" then%>src="users_name1.asp?users_card=<%=rsstr%>"<%end if%>></iframe>
		</td>
	  </tr>
	</table>
	  </td>
  </tr>
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong>兑换商品：</strong></td>
    <td align="left"><input name="name1" type="text" id="name1" class="input1" />
        <span class="red">*</span></td>
  </tr>
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong>商品价值：</strong></td>
    <td align="left"><input name="price" id="price" type="text" class="input1" onKeyUp="<%=fbjs_rnum2%>jifenjs2();jifenjs3();" />
        <span class="red">*</span></td>
  </tr>
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong>兑换数量：</strong></td>
    <td align="left"><input name="num" id="num" type="text" value="1" class="input1" onKeyUp="<%=fbjs_rnum1%>jifenjs2();jifenjs3();" />
        <span class="red">*</span></td>
  </tr>
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong>合计价值：</strong></td>
    <td align="left"><input name="priceall" id="priceall" type="text" class="input1" onKeyUp="<%=fbjs_rnum2%>jifenjs3();" />
        <span class="red">*</span></td>
  </tr>
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong>所需积分：</strong></td>
    <td align="left"><input name="score" id="score" type="text" class="input1" onKeyUp="<%=fbjs_rnum1%>" />
      <span class="red">*</span></td>
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
 sql="select * from jifen where id="&request("id")
 rs.Open sql,conn,1,3
	sql="select card from users where id="&rs("usersid")&" "
	rsstr=trim(conn.execute(sql).getstring)
	rsstr=replace(""&rsstr,chr(13),"")
%>
<strong>修改<%=menunamel%></strong><br /><br />
<form action="?id=<%=rs("id")%>&amp;tj=Modify" method="post" name="form3" id="form3">
  <table border="1" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
    <tr>
      <td width="100" align="center" bgcolor="#F0F0F0"><strong>兑换日期：</strong></td>
      <td align="left"><select name="time1_y" id="time1_y">
          <%for i=2013 to year(now)+1%>
          <option value="<%=i%>" <%if i=rs("time1_y") then%>selected="selected"<%end if%>><%=i%></option>
          <%next%>
        </select>
        年
        <select name="time1_m" id="time1_m">
          <%for i=1 to 12%>
          <option value="<%=i%>" <%if i=rs("time1_m") then%>selected="selected"<%end if%>><%=i%></option>
          <%next%>
        </select>
        月
        <select name="time1_d" id="time1_d">
          <%for i=1 to 31%>
          <option value="<%=i%>" <%if i=rs("time1_d") then%>selected="selected"<%end if%>><%=i%></option>
          <%next%>
        </select>
        日 <span class="red">*</span></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>会员卡号：</strong></td>
      <td align="left"><table border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><input name="users_card" id="users_card" type="text" size="20" value="<%=rsstr%>" onkeyup="<%=fbjs_rnum1%>jifenjs1();" />
			<input name="a_users_card" id="a_users_card" type="hidden" value="<%=rsstr%>" />
                <span class="red">*</span>&nbsp;&nbsp;&nbsp;&nbsp; </td>
            <td valign="middle"><iframe id="users_name1" name="users_name1" frameborder="0" width="230" height="23" scrolling="No" src="users_name1.asp?users_card=<%=rsstr%>"></iframe></td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>兑换商品：</strong></td>
      <td align="left"><input name="name1" type="text" class="input1" id="name1" value="<%=rs("name1")%>" />
          <span class="red">*</span></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>商品价值：</strong></td>
      <td align="left"><input name="price" type="text" class="input1" id="price" onkeyup="<%=fbjs_rnum2%>jifenjs2();jifenjs3();" value="<%=rs("price")%>" />
          <span class="red">*</span></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>兑换数量：</strong></td>
      <td align="left"><input name="num" id="num" type="text" value="<%=rs("num")%>" class="input1" onkeyup="<%=fbjs_rnum1%>jifenjs2();jifenjs3();" />
          <span class="red">*</span></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>合计价值：</strong></td>
      <td align="left"><input name="priceall" type="text" class="input1" id="priceall" onkeyup="<%=fbjs_rnum2%>jifenjs3();" value="<%=rs("priceall")%>" />
          <span class="red">*</span></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>所需积分：</strong></td>
      <td align="left"><input name="score" id="score" type="text" class="input1" onkeyup="<%=fbjs_rnum1%>" value="<%=rs("score")%>" />
		  <input name="a_score" id="a_score" type="hidden" value="<%=rs("score")%>" />
          <span class="red">*</span>
		  </td>
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