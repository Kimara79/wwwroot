<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="消费记录" %>
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
sql="select info from system where id=15"
rsstr=trim(conn.execute(sql).getstring)
zhh=replace(rsstr,chr(13),"")
zhh=split(zhh,":")
%>

<script language="JavaScript">
function xiaofeijs1(){
	document.getElementById("users_name1").src="users_name1.asp?users_card="+document.getElementById("users_card").value;
}
function xiaofeijs2(){
	document.getElementById("priceall").value=document.getElementById("price").value*document.getElementById("num").value;
}
function xiaofeijs3(){
	//最开始的100元只算1积分
	var jf;
	if(parent.frames['users_name1'].document.getElementsByTagName('span')){
		jf=parent.frames['users_name1'].document.getElementsByTagName('span')[0].innerHTML;
		jf=jf.replace(" ","");
		jf=parseInt(jf,10);
		if(document.getElementById("priceall").value){
			var a;
			a=parseInt(document.getElementById("priceall").value,10);
			if(jf>0){
				document.getElementById("score").value=parseInt(a*<%=zhh(1)%>/<%=zhh(0)%>,10);
			}
			else{
				if(a<100){
					document.getElementById("score").value=0;
				}
				else{
					a=a-100;
					document.getElementById("score").value=parseInt(1+(a*<%=zhh(1)%>/<%=zhh(0)%>),10);
				}
			}
		}
	}
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
		Response.Write("<script language=javascript>alert('商品名称不能为空！');history.back();</script>")
		Response.End
	end if
	if request.Form("price")="" then 
		Response.Write("<script language=javascript>alert('商品单价不能为空！');history.back();</script>")
		Response.End
	end if
	if request.Form("num")="" then 
		Response.Write("<script language=javascript>alert('商品数量不能为空！');history.back();</script>")
		Response.End
	end if
	if request.Form("priceall")="" then 
		Response.Write("<script language=javascript>alert('合计金额不能为空！');history.back();</script>")
		Response.End
	end if
	if request.Form("score")="" then 
		Response.Write("<script language=javascript>alert('获得积分不能为空！');history.back();</script>")
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
	rs.Open "xiaofei",conn,1,3
	rs.addnew
	On Error Resume Next
	For each obj in request.Form
		rs(obj)=request.Form(obj)
	Next
	rs("time1")=request.Form("time1_y")&"-"&request.Form("time1_m")&"-"&request.Form("time1_d")
	rs("usersid")=usersid
	rs("addtime")=now()
	rs.update
	rs.close
	
	'会员加积分
	sql="update users set score=score+"&request.Form("score")&" where id="&usersid&" "
	conn.execute sql,1,1 
	
	rsrz.open "book",conn,1,3                                   '添加日志              
	rsrz.addnew
	rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
	rsrz("caozuo")="添加"
	rsrz("menu")=menuname
	rsrz("duixian")=request.Form("name1")
	rsrz.update
	rsrz.close
	response.redirect "xiaofeiManage.asp"
	Response.End
end if
%>

<% '--------------------------------------------------------提交操作判断-修改
if request("tj")="Modify" then
	sql="select * from xiaofei where id="&request("id")
	rs.Open sql,conn,1,3
	On Error Resume Next
	For each obj in request.Form
		rs(obj)=request.Form(obj)
	Next
	rs("time1")=request.Form("time1_y")&"-"&request.Form("time1_m")&"-"&request.Form("time1_d")
	rs("usersid")=usersid
	rs.update
	rs.close
	
	'修改卡号或积分
	sql="update users set score=score-"&request.Form("a_score")&" where id="&a_usersid&" "
	conn.execute sql,1,1 
	sql="update users set score=score+"&request.Form("score")&" where id="&usersid&" "
	conn.execute sql,1,1 	
	
	rsrz.open "book",conn,1,3                                   '添加日志              
	rsrz.addnew
	rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
	rsrz("caozuo")="修改"
	rsrz("menu")=menuname
	rsrz("duixian")=request.Form("name1")
	rsrz.update
	rsrz.close
	
	response.redirect "xiaofeiManage.asp"
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
    <td width="100" align="center" bgcolor="#F0F0F0"><strong>消费日期：</strong></td>
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
		<input name="users_card" id="users_card" type="text" size="20" value="<%=rsstr%>" onKeyUp="<%=fbjs_rnum1%>xiaofeijs1();xiaofeijs3();" />
		<span class="red">*</span>&nbsp;&nbsp;&nbsp;&nbsp;
		</td>
		<td valign="middle">
		<iframe id="users_name1" name="users_name1" frameborder="0" width="270" height="23" scrolling="no" <%if rsstr<>"" then%>src="users_name1.asp?users_card=<%=rsstr%>"<%end if%>></iframe>
		</td>
	  </tr>
	</table>
	  </td>
  </tr>
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong>商品名称：</strong></td>
    <td align="left"><input name="name1" type="text" id="name1" class="input1" />
        <span class="red">*</span></td>
  </tr>
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong>商品单价：</strong></td>
    <td align="left"><input name="price" id="price" type="text" class="input1" onKeyUp="<%=fbjs_rnum2%>xiaofeijs2();xiaofeijs3();" />
        <span class="red">*</span></td>
  </tr>
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong>购买数量：</strong></td>
    <td align="left"><input name="num" id="num" type="text" value="1" class="input1" onKeyUp="<%=fbjs_rnum1%>xiaofeijs2();xiaofeijs3();" />
        <span class="red">*</span></td>
  </tr>
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong>合计金额：</strong></td>
    <td align="left"><input name="priceall" id="priceall" type="text" class="input1" onKeyUp="<%=fbjs_rnum2%>xiaofeijs3();" />
        <span class="red">*</span></td>
  </tr>
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong>获得积分：</strong></td>
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
 sql="select * from xiaofei where id="&request("id")
 rs.Open sql,conn,1,3
	sql="select card from users where id="&rs("usersid")&" "
	rsstr=trim(conn.execute(sql).getstring)
	rsstr=replace(""&rsstr,chr(13),"")
%>
<strong>修改<%=menunamel%></strong><br /><br />
<form action="?id=<%=rs("id")%>&amp;tj=Modify" method="post" name="form3" id="form3">
  <table border="1" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
    <tr>
      <td width="100" align="center" bgcolor="#F0F0F0"><strong>消费日期：</strong></td>
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
            <td><input name="users_card" id="users_card" type="text" size="20" value="<%=rsstr%>" onkeyup="<%=fbjs_rnum1%>xiaofeijs1();xiaofeijs3();" />
			<input name="a_users_card" id="a_users_card" type="hidden" value="<%=rsstr%>" />
                <span class="red">*</span>&nbsp;&nbsp;&nbsp;&nbsp; </td>
            <td valign="middle"><iframe id="users_name1" name="users_name1" frameborder="0" width="270" height="23" scrolling="No" src="users_name1.asp?users_card=<%=rsstr%>"></iframe></td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>商品名称：</strong></td>
      <td align="left"><input name="name1" type="text" class="input1" id="name1" value="<%=rs("name1")%>" />
          <span class="red">*</span></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>商品单价：</strong></td>
      <td align="left"><input name="price" type="text" class="input1" id="price" onkeyup="<%=fbjs_rnum2%>xiaofeijs2();xiaofeijs3();" value="<%=rs("price")%>" />
          <span class="red">*</span></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>购买数量：</strong></td>
      <td align="left"><input name="num" id="num" type="text" value="<%=rs("num")%>" class="input1" onkeyup="<%=fbjs_rnum1%>xiaofeijs2();xiaofeijs3();" />
          <span class="red">*</span></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>合计金额：</strong></td>
      <td align="left"><input name="priceall" type="text" class="input1" id="priceall" onkeyup="<%=fbjs_rnum2%>xiaofeijs3();" value="<%=rs("priceall")%>" />
          <span class="red">*</span></td>
    </tr>
    <tr>
      <td align="center" bgcolor="#F0F0F0"><strong>获得积分：</strong></td>
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