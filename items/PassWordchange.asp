<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title>修改密码</title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="inc/md5.asp"-->

<%
if request("no")="change" then

sql="select  * from  admin where admin='"&session("admin")&"'"
rs.open sql,conn,1,3
 
 if MD5(request.Form("password1")) <> rs("password") then 
 Response.Write("<script language=javascript>alert('原密码错误！');history.back()</script>")
 Response.End
 end if
 
 if request.Form("password21")="" then 
 Response.Write("<script language=javascript>alert('新密码不能为空！');history.back()</script>")
 Response.End
 end if
 
 if request.Form("password22")="" then 
 Response.Write("<script language=javascript>alert('确认码不能为空！');history.back()</script>")
 Response.End
 end if
 
 if request.Form("password21") <> request.Form("password22") then 
 Response.Write("<script language=javascript>alert('确认码错误！');history.back()</script>")
 Response.End
 end if
 
 rs("password")=MD5(request.Form("password21"))
 rs.update
 rs.close
 
		rsrz.open "book",conn,1,3                                  '添加日志              
		rsrz.addnew
		  rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
		  rsrz("caozuo")="修改"
		  rsrz("menu")="修改密码"
		  rsrz("duixian")="该帐号密码"
		rsrz.update
		rsrz.close
 
 Response.Write("<script language=javascript>alert('修改成功！');location='PassWordchange.asp'</script>")
 Response.End

end if
%>

<body>
<!--#include file="top.asp"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23" align="center" background="images/am_mn2.gif" bgcolor="cccccc"><strong>修改密码</strong>
	</td>
  </tr>
</table>
<div align="center"><br />
  <form id="form1" name="form1" method="post" action="PassWordchange.asp?no=change">
    <table width="420" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
            <tr>
              <td width="100" height="30" align="center" bgcolor="#F0F0F0"><span class="STYLE1">用户名：</span></td>
              <td align="left" >&nbsp;&nbsp;<b><%=session("admin")%></b></td>
            </tr>
            <tr>
              <td height="30" align="center" bgcolor="#F0F0F0"><span class="STYLE1">真实姓名：</span></td>
              <td align="left" >&nbsp;&nbsp;<b><%=Session("truename")%></b></td>
            </tr>
            <tr>
              <td height="30" align="center" bgcolor="#F0F0F0"><span class="STYLE1">原密码：</span></td>
              <td align="left"><input name="password1" type="password" class="table2" id="password1" size="40" />              </td>
            </tr>
            <tr>
              <td height="30" align="center" bgcolor="#F0F0F0"><span class="STYLE1">新密码：</span></td>
              <td align="left"><input name="password21" type="password" class="table2" id="password21" size="40" /></td>
            </tr>
            <tr>
              <td height="30" align="center" bgcolor="#F0F0F0"><span class="STYLE1">确　认：</span></td>
              <td align="left"><input name="password22" type="password" class="table2" id="password22" size="40" /></td>
            </tr>
    </table>
         <!--#include file="inc/button.asp"-->
  </form>
</div>
<!--#include file="foot.asp"-->
</body>
</html>