<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title>其他设置</title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="js/SetCwinHeight.js"-->

<% 
'----------------------------------------------------------------------添加
if request("tj")="Add" then

rs.open "system",conn,1,3
rs.addnew
On Error Resume Next
  For each obj in Request.Form
   rs(obj)=Request.Form(obj)
  Next
   rs("pic")=session("pic")
	rs("laster")=session("admin")&"("&session("truename")&")"
	rs("lasttime")=now()
rs.update
rs.close

		rsrz.open "book",conn,1,3                                  '添加日志              
		rsrz.addnew
		  rsrz("zhanghao")=session("admin")&"("&session("truename")&")"
		  rsrz("caozuo")="添加"
		  rsrz("menu")="其他设置"
		  rsrz("duixian")=rs("menu_en")&"("&rs("menu")&")"
		  rsrz("time")=now()
		rsrz.update
		rsrz.close

response.redirect "systemManage.asp"
Response.End
end if

'----------------------------------------------------------------------修改
if request("tj")="Modify" then

sql="select * from system where id="&request("id")                 '修改数据  
rs.open sql,conn,1,3
On Error Resume Next
  For each obj in Request.Form
   rs(obj)=Request.Form(obj)
  Next
   rs("pic")=session("pic")
	rs("laster")=session("admin")&"("&session("truename")&")"
	rs("lasttime")=now()
rs.update
rs.close

		rsrz.open "book",conn,1,3                                  '添加日志              
		rsrz.addnew
		  rsrz("zhanghao")=session("admin")&"("&session("truename")&")"
		  rsrz("caozuo")="修改"
		  rsrz("menu")="其他设置"
		  rsrz("duixian")=rs("menu_en")&"("&rs("menu")&")"
		  rsrz("time")=now()
		rsrz.update
		rsrz.close

response.redirect "systemManage.asp"
Response.End
end if
%>

<body>
<!--#include file="top.asp"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23" align="center" background="images/am_mn2.gif" bgcolor="cccccc"><strong>其他设置</strong></td>
  </tr>
</table>
<div align="center"><br />
<%if  request("cz")="Add" then  '------------------------------------------添加界面%>
<strong>添加设置</strong><br /><br />
<form id="form" name="form" method="post" action="systemAM.asp?tj=Add">
          <table border="1" cellspacing="0" cellpadding="5" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
            <tr>
              <td width="100" align="center" bgcolor="#F0F0F0" class="STYLE1">名称：</td>
              <td align="left"><input name="name1" type="text" id="name1" size="40" /></td>
            </tr>
            <tr>
              <td align="center" bgcolor="#F0F0F0" class="STYLE1">内容：</td>
              <td align="left"><textarea name="info" cols="60" id="info" style="height:expression(this.scrollHeight);padding:10px;"></textarea></td>
            </tr>
    </table>
         <!--#include file="inc/button.asp"-->
  </form>
 <%end if%>

<%if  request("cz")="Modify" then  '------------------------------------------修改界面%>
<%
sql="select * from system where id="&request("id")
rs.open sql,conn,1,1
%>
      <strong>修改设置</strong><br />
      最后更新时间：<%=rs("lasttime")%>　　最后更新人：<%=rs("laster")%><br /><br />
  <form id="form" name="form" method="post" action="systemAM.asp?id=<%=rs("id")%>&amp;tj=Modify">
          <table border="1" cellspacing="0" cellpadding="5" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
            <tr>
              <td width="100" align="center" bgcolor="#F0F0F0" class="STYLE1">名称：</td>
              <td align="left"><input name="name1" type="text" id="name1" value="<%=rs("name1")%>" size="40" /></td>
            </tr>
            <tr>
              <td align="center" bgcolor="#F0F0F0" class="STYLE1">内容：</td>
              <td align="left"><textarea name="info" cols="60" id="info" style="height:expression(this.scrollHeight);padding:10px;"><%=rs("info")%></textarea>              </td>
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