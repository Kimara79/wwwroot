<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="积分兑换" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="inc/formatymd3.asp"-->

<%
'删除操作
if request("cz")="del" then
	sql="delete * from jifen where id="&request("id")               '删除数据
	conn.execute sql,1,1 
	
	rsrz.open "book",conn,1,3                                   '添加日志              
	rsrz.addnew
	rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
	rsrz("caozuo")="删除"
	rsrz("menu")=menuname
	rsrz("duixian")=request("name1")
	rsrz.update
	rsrz.close
	
	'会员减积分
	sql="update users set score=score+"&request.QueryString("score")&" where id="&request.QueryString("usersid")&" "
	conn.execute sql,1,1 
	
	response.redirect "?"
	Response.End
end if

'出错判断
if request.Form("search_key")<>"" then
	if request.Form("search_type")="users_card" then
		if IsNumeric(request.Form("search_key")) then
		else
			Response.Write("<script language=javascript>alert('卡号有误！');history.back();</script>")
			Response.End
		end if
	end if
end if

if request("xs")="all" then
	session("jifen_all")="1"
	session("jifen_search_type")=""
	session("jifen_search_key")=""
end if

if request("search_key")<>"" then
	session("jifen_all")="0"
	session("jifen_search_type")=request("search_type")
	session("jifen_search_key")=request("search_key")
end if

if request("page")<>"" then
	session("jifen_page")=request("page")
end if
if session("jifen_page")="" then
	session("jifen_page")=1
end if

sqladd=""
if session("jifen_search_key")<>"" then
	select case session("jifen_search_type")
		case "users_card"
			sqladd=sqladd&" and usersid in (select id from users where card="&session("jifen_search_key")&" ) "
		case "users_name1"
			sqladd=sqladd&" and usersid in (select id from users where name1 like '%"&session("jifen_search_key")&"%' ) "
		case else
			sqladd=sqladd&" and "&session("jifen_search_type")&" like '%"&session("jifen_search_key")&"%' "
	end select
end if
%>

<body>
<!--#include file="top.asp"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23" align="center" background="images/am_mn2.gif" bgcolor="cccccc"><strong><%=menuname%></strong></td>
  </tr>
</table>
<br>

<center>
  <form id="form1" name="form1" method="post" action="?xs=all">
    <table border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td class="STYLE1">搜索：&nbsp;&nbsp;</td>
        <td>
		<select name="search_type" id="search_type">
            <option value="users_card" <%if session("jifen_search_type")="users_card" then%>selected="selected"<%end if%>>会员卡号</option>
            <option value="users_name1" <%if session("jifen_search_type")="users_name1" then%>selected="selected"<%end if%>>会员姓名</option>
            <option value="name1" <%if session("jifen_search_type")="name1" then%>selected="selected"<%end if%>>兑换商品</option>
			<option value="info" <%if session("jifen_search_type")="info" then%>selected="selected"<%end if%>>备注</option>
        </select>
	    </td>
        <td>&nbsp;&nbsp;<input name="search_key" type="text" id="search_key" value="<%=session("jifen_search_key")%>" onKeyPress="maskEdit(/^[\w]*$/)" /></td>
        <td align="right">&nbsp;&nbsp;<input type="submit" name="Submit22" value=" 搜索 " /></td>
        <td align="right">&nbsp;&nbsp;<input type="button" name="Submit2x2" onClick="location.href='?xs=all'" value="清除搜索" /></td>
        <td align="right">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="Submit2" onClick="location.href='jifenAM.asp?cz=Add'" value=" 添加记录 " /></td>
     </tr>
    </table>
  </form>
</center>
<br>

<center>
<%
sql="select * from jifen where id>0 "&sqladd&" order by time1 desc"
rs.open sql,conn,1,1
%>
  <%if rs.bof then%>
  <p align="center">暂无信息</p>
  <%
  else
	rs.pagesize=20  '设置每页显示数目
	currentpage=Clng(session("jifen_page"))
	if currentpage<1 then currentpage=1
	if currentpage>rs.pagecount then currentpage=rs.pagecount
	rs.absolutepage=currentpage
  %>

  <table border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
	<tr align="center" bgcolor="#F0F0F0" class="STYLE1">
	  <td><strong>兑换日期</strong></td>
	  <td><strong>会员卡号</strong></td>
	  <td><strong>会员姓名</strong></td>
	  <td><strong>兑换商品</strong></td>
	  <td><strong>商品价值</strong></td>
	  <td><strong>兑换数量</strong></td>
	  <td><strong>合计价值</strong></td>
	  <td><strong>所需积分</strong></td>
	  <td><strong>剩余积分</strong></td>
	  <td><strong>备注</strong></td>
	  <td><strong>操作</strong></td>
	</tr>
<%Do while not rs.Eof%>
        <tr onMouseOver="bgColor='#D2F5FF';" onMouseOut="bgColor='#FFFFFF';">
          <td align="center">&nbsp;<%=formatymd3b(rs("time1"))%>&nbsp;</td>
          <td align="center">&nbsp;
		  <%
			sql="select card from users where id="&rs("usersid")&" "
			rsstr=trim(conn.execute(sql).getstring)
			rsstr=replace(""&rsstr,chr(13),"")
			response.Write rsstr
		  %>
		  &nbsp;</td>
          <td align="center">&nbsp;
		  <%
			sql="select name1 from users where id="&rs("usersid")&" "
			rsstr=trim(conn.execute(sql).getstring)
			rsstr=replace(""&rsstr,chr(13),"")
			response.Write rsstr
		  %>
		  &nbsp;</td>
          <td align="left">&nbsp;<%=rs("name1")%>&nbsp;</td>
          <td align="right">&nbsp;<%=rs("price")%>&nbsp;</td>
          <td align="center">&nbsp;<%=rs("num")%>&nbsp;</td>
          <td align="right">&nbsp;<%=rs("priceall")%>&nbsp;</td>
          <td align="center">&nbsp;<%=rs("score")%>&nbsp;</td>
          <td align="center">&nbsp;<%=rs("score2")%>&nbsp;</td>
          <td align="left">&nbsp;<%=rs("info")%>&nbsp;</td>
          <td align="center">
			<a href="?id=<%=rs("id")%>&cz=del&name1=<%=rs("name1")%>&score=<%=rs("score")%>&usersid=<%=rs("usersid")%>" onClick="{if(confirm('确定要删除么?该会员将被扣除此次记录的积分。')){this.document.album.submit();return true;}return false;}">删除</a>&nbsp;&nbsp;
			<a href="jifenAM.asp?id=<%=rs("id")%>&amp;cz=Modify">修改</a>
		  </td>
        </tr>
<%
i=i+1
rs.MoveNext
If i>=rs.pagesize Then Exit Do
Loop
%>
</table>
<!--#include file="inc/page.asp"-->
<%
end if
rs.close
%>
</center>
<!--#include file="foot.asp"-->
</body>
</html>