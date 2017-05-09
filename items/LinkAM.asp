<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="友情链接" %>
<% menuname1="链接" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="js/red.js"-->

<%
On Error Resume Next                                                          '出错判断
if request("tj")="Add" or request("tj")="Modify" then

 if request.Form("name1")="" then 
 Response.Write("<script language=javascript>alert('名称不能为空！');history.back()</script>")
 Response.End
 end if

 if request.Form("link")="" then 
 Response.Write("<script language=javascript>alert('网址不能为空！');history.back()</script>")
 Response.End
 end if
 
end if 
%>

<% '----------------------------------------------------------------------提交操作判断-添加
if request("tj")="Add" then

rs.Open "link",conn,1,3                                             '添加数据
rs.addnew
  For each obj in Request.Form
   rs(obj)=trim(Request.Form(obj))
  Next
rs.update
rs.close

		rsrz.open "book",conn,1,3                                   '添加日志              
		rsrz.addnew
		  rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
		  rsrz("caozuo")="添加"
		  rsrz("menu")=menuname
		  rsrz("duixian")=request.Form("link")
		rsrz.update
		rsrz.close
		
response.redirect "linkManage.asp"
end if
%>
<% 
'----------------------------------------------------------------------提交操作判断-修改
if request("tj")="Modify" then

sql="select * from link where id="&request("id")                 '修改数据  
rs.open sql,conn,1,3
  For each obj in Request.Form
   rs(obj)=trim(Request.Form(obj))
  Next
rs.update
rs.close

		rsrz.open "book",conn,1,3                               '添加日志              
		rsrz.addnew
		  rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
		  rsrz("caozuo")="修改"
		  rsrz("menu")=menuname
		  rsrz("duixian")=request.Form("link")
		rsrz.update
		rsrz.close

response.redirect "linkManage.asp"
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
        <%if  request("cz")="Add" then  '------------------------------------------添加界面
		rs.Open "link",conn,1,3
		%>
        <form id="form" name="form" method="post" action="?tj=Add">
          <label><strong>添加<%=menuname1%></strong><br />
          <br />
          </label>
          <table width="600"  border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" >
            <tr>
              <td height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">审核:</span></td>
              <td align="left">
			  <label style="color:#FF0000"><input name="shenghe" type="radio" value="通过" checked="checked" onpropertychange="red(this)"/>
                通过</label>
              <label><input type="radio" name="shenghe" value="未通过" onpropertychange="red(this)"/>
                未通过</label>
			  &nbsp;&nbsp;
              <span class="red">*</span></td>
            </tr>
            <tr>
              <td height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">序号:</span></td>
              <td align="left"><input name="no1" type="text" id="no1" value="<%=rs.recordcount+1%>" size="10" />
                  <span class="red">*</span></td>
            </tr>
            <tr>
              <td width="25%" height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">名称:</span></td>
              <td align="left"><input name="name1" type="text" id="name1" size="41" />
                  <span class="red">*</span></td>
            </tr>
            <tr>
              <td height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">网址:</span></td>
              <td align="left"><input name="link" type="text" id="link" size="41" />
                  <span class="red">*</span></td>
            </tr>
            <tr>
              <td height="60" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">备注:</span></td>
              <td align="left"><textarea name="info" cols="40" rows="7" id="info"></textarea></td>
            </tr>
          </table>
         <!--#include file="inc/button.asp"-->
        </form>
      <%
	  rs.close
	  end if
	  %>
        <%
if  request("cz")="Modify" then      '------------------------------------------------修改界面
 sql="select * from link where id="&request("id")
 rs.Open sql,conn,1,3
 %>
        <form id="form" name="form" method="post" action="?id=<%=rs("id")%>&amp;tj=Modify">
          <strong>修改<%=menuname1%></strong><br />
         最后更新时间：<%=rs("time1")%> <br /><br />
          <table width="600" border="1" cellspacing="0" cellpadding="3"  bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" >
            <tr>
              <td height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">审核:</span></td>
              <td align="left">
			  <label <%if rs("shenghe")="通过" then%>style="color:#FF0000;"<%end if%>>
			  <input name="shenghe" type="radio" value="通过" <%if rs("shenghe")="通过" then%>checked="checked"<%end if%> onpropertychange="red(this)"/> 通过</label>
			  <label <%if rs("shenghe")="未通过" then%>style="color:#FF0000;"<%end if%>>
              <input type="radio" name="shenghe" value="未通过" <%if rs("shenghe")="未通过" then%>checked="checked"<%end if%> onpropertychange="red(this)"/> 未通过</label>
			  &nbsp;&nbsp;
			  <span class="red">*</span>			  </td>
            </tr>
            <tr>
              <td height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">序号:</span></td>
              <td align="left"><input name="no1" type="text" id="no1" value="<%=rs("no1")%>" size="10" />
                <span class="red">*</span></td>
            </tr>
            <tr>
              <td width="25%" height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">名称:</span></td>
              <td align="left"><input name="name1" type="text" id="name1" value="<%=rs("name1")%>" size="41" />
                <span class="red">*</span></td>
            </tr>
            <tr>
              <td height="30" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">网址:</span></td>
              <td align="left"><input name="link" type="text" id="link" value="<%=rs("link")%>" size="41" />
                <span class="red">*</span></td>
            </tr>
            <tr>
              <td height="60" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">备注:</span></td>
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