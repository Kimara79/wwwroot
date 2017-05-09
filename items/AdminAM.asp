<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title>帐号管理</title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="inc/md5.asp"-->

<%                                                           '出错判断1
if request("tj")="add" or request("tj")="Modify" then

 if request.Form("step")="" then 
 Response.Write("<script language=javascript>alert('级别不能为空！');history.back()</script>")
 Response.End
 end if

 if request.Form("admin")="" then 
 Response.Write("<script language=javascript>alert('用户名不能为空！');history.back()</script>")
 Response.End
 end if
 
 '用户名重复判断
 if request("tj")="add" then
 sql="select * from admin where admin='"&request.Form("admin")&"'"
 end if
 if request("tj")="Modify" then
 sql="select * from admin where admin='"&request.Form("admin")&"' and id<>"&request("id")&""
 end if
 rs.open sql,conn,1,1
 if not rs.eof then
 Response.Write("<script language=javascript>alert('已存在此用户名！');history.back()</script>")
 Response.End
 end if
 rs.close
 
end if 
%>
            <%                                                           '出错判断2
if request("tj")="add" then

 if request.Form("password")="" then 
 Response.Write("<script language=javascript>alert('密码不能为空！');history.back()</script>")
 Response.End
 end if
 
 if request.Form("re_password")="" then 
 Response.Write("<script language=javascript>alert('确认码不能为空！');history.back()</script>")
 Response.End
 end if
 
 if request.Form("password") <> request.Form("re_password") then 
 Response.Write("<script language=javascript>alert('确认码错误！');history.back()</script>")
 Response.End
 end if
 
end if 
%>
            <% '----------------------------------------------------------------------提交操作判断-添加
if request("tj")="add" then

rs.Open "admin",conn,1,3                                             '添加数据
rs.addnew
	rs("step")=request.Form("step")
    rs("admin")=request.Form("admin")
  if request.Form("password")<>"" then 
    rs("password")=md5(request.Form("password"))
  end if
    rs("info")=request.Form("info")
	rs("truename")=request.Form("truename")
rs.update

		rsrz.open "book",conn,1,3                                  '添加日志              
		rsrz.addnew
		  rsrz("zhanghao")=session("admin")&"("&session("truename")&")"
		  rsrz("caozuo")="添加"
		  rsrz("menu")="帐号管理"
		  rsrz("duixian")="帐号："&request.Form("admin")&"("&request.Form("truename")&")"
		rsrz.update
		rsrz.close
		
response.redirect "AdminManage.asp"
end if
%>
            <% 
'----------------------------------------------------------------------提交操作判断-修改
if request("tj")="Modify" then

sql="select * from admin where id="&request("id")                 '修改数据  
rs.open sql,conn,1,3
    rs("admin")=request.Form("admin")
  if request.Form("password")<>"" then 
    rs("password")=md5(request.Form("password"))
  end if
    rs("info")=request.Form("info")
	rs("step")=request.Form("step")
	rs("truename")=request.Form("truename")
rs.update

		rsrz.open "book",conn,1,3                                  '添加日志              
		rsrz.addnew
		  rsrz("zhanghao")=session("admin")&"("&session("truename")&")"
		  rsrz("caozuo")="修改"
		  rsrz("menu")="帐号管理"
		  rsrz("duixian")="帐号："&request.Form("admin")&"("&request.Form("truename")&")"
		rsrz.update
		rsrz.close

response.redirect "AdminManage.asp"
end if
%>

<body>    
<!--#include file="top.asp"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23" align="center" background="images/am_mn2.gif" bgcolor="cccccc"><strong>帐号管理</strong></td>
  </tr>
</table>
<div align="center">
	<br />
        <%if  request("cz")="Add" then  '------------------------------------------添加界面%>
        <form id="form" name="form" method="post" action="AdminAM.asp?tj=add">
          <label><strong>添加帐号</strong><br />
          <br />
          </label>
          <table width="500"  border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" >
            <tr>
              <td height="40" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">级　　别:</span></td>
              <td align="left"><p>
                  <label>
                  <input name="step" type="text" id="step" value="2" size="10" />
                    ( 1级才有“帐号管理、日志管理、数据管理”等)</label>
                  <br />
              </p></td>
            </tr>
            <tr>
              <td width="25%" height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">用 户 名:</span></td>
              <td align="left"><input name="admin" type="text" id="admin" size="41" />
                  <span class="red">*</span></td>
            </tr>
            <tr>
              <td height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">真实姓名:</span></td>
              <td align="left"><input name="truename" type="text" id="truename" size="41" />
                  <span class="red">*</span></td>
            </tr>
            <tr>
              <td height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">密　　码:</span></td>
              <td align="left"><input name="password" type="password" id="password" value="" size="45" />
                  <span class="red">*</span></td>
            </tr>
            <tr>
              <td height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">确认密码:</span></td>
              <td align="left"><input name="re_password" type="password" id="re_password" size="45" />
                  <span class="red">*</span></td>
            </tr>
            <tr>
              <td height="60" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">备　　注:</span></td>
              <td align="left"><textarea name="info" cols="40" rows="7" id="info"></textarea></td>
            </tr>
          </table>
         <!--#include file="inc/button.asp"-->
        </form>
      <%end if%>
        <%
if  request("cz")="Modify" then      '------------------------------------------------修改界面
 sql="select * from admin where id="&request("id")
 rs.Open sql,conn,1,3
 %>
        <form id="form" name="form" method="post" action="AdminAM.asp?id=<%=rs("id")%>&amp;tj=Modify">
          <label><strong>修改帐号</strong><br />
          <br />
          </label>
          <table width="500" border="1" cellspacing="0" cellpadding="3"  bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" >
            <tr>
              <td height="40" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">级　　别:</span></td>
              <td align="left"><p>
                  <label>
                  <input name="step" type="text" id="step" value="<%=rs("step")%>" size="10" />
                  ( 1级才有“帐号管理、日志管理、数据管理”等)</label>
<br />
              </p></td>
            </tr>
            <tr>
              <td width="25%" height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">用 户 名:</span></td>
              <td align="left"><input name="admin" type="text" id="admin" value="<%=rs("admin")%>" size="41" />
                  <span class="red">*</span></td>
            </tr>
            <tr>
              <td height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">真实姓名:</span></td>
              <td align="left"><input name="truename" type="text" id="truename" value="<%=rs("truename")%>" size="41" />
                  <span class="red">*</span></td>
            </tr>
            <tr>
              <td height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">密　　码:</span></td>
              <td align="left"><input name="password" type="password" id="password" size="45" /></td>
            </tr>
            <tr>
              <td height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">确认密码:</span></td>
              <td align="left"><input name="re_password" type="password" id="re_password" size="45" /></td>
            </tr>
            <tr>
              <td height="60" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">备　　注:</span></td>
              <td align="left"><textarea name="info" cols="40" rows="7" id="info"><%=rs("info")%></textarea></td>
            </tr>
          </table>
         <!--#include file="inc/button.asp"-->
        </form>
      <%end if%>
</div>
<!--#include file="foot.asp"-->
</body>
</html>