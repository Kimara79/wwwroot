<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title>在线客服</title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="js/red.js"-->
<!--#include file="CserverAM_inc.asp"-->

<%
On Error Resume Next                                                          '出错判断
if request("tj")="Add" or request("tj")="Modify" then

 if request.Form("no1")="" then 
 Response.Write("<script language=javascript>alert('序号不能为空！');history.back()</script>")
 Response.End
 end if

 if request.Form("name1")="" then 
 Response.Write("<script language=javascript>alert('帐号不能为空！');history.back()</script>")
 Response.End
 end if
 
end if 
%>

<% '----------------------------------------------------------------------提交操作判断-添加
if request("tj")="Add" then

rs.Open "cserver",conn,1,3                                             '添加数据
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
		  rsrz("menu")="在线客服"
		  rsrz("duixian")=request.Form("name")
		rsrz.update
		rsrz.close
		
response.redirect "CserverManage.asp"
end if
%>
<% 
'----------------------------------------------------------------------提交操作判断-修改
if request("tj")="Modify" then

sql="select * from cserver where id="&request("id")                 '修改数据  
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
		  rsrz("menu")="在线客服"
		  rsrz("duixian")=request.Form("name")
		rsrz.update
		rsrz.close

response.redirect "CserverManage.asp"
end if
%>

<body>
<!--#include file="top.asp"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23" align="center" background="images/am_mn2.gif" bgcolor="cccccc"><strong>在线客服</strong></td>
  </tr>
</table>
<div align="center">
	<br />
        <%if  request("cz")="Add" then  '------------------------------------------添加界面
		rs.Open "cserver",conn,1,3
		%>
        <form id="form" name="form" method="post" action="?tj=Add">
          <strong>添加客服</strong><br /><br />
          <table width="600"  border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" >
            <tr>
              <td align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">类别:</span></td>
              <td align="left">
				<%for z=0 to ubound(sort1)%>
				 <label style="white-space:nowrap;<%if z=0 then%>color:#FF0000;<%end if%>">
				 <input type="radio" class="dq" name="sort1" id="sort1<%=z%>" value="<%=sort1(z)%>" onpropertychange="red(this)" <%if z=0 then%>checked="checked"<%end if%>/>
				 <%=sort1(z)%></label>
				<%next%>
			  <span class="red">*</span>				</td>
            </tr>
            <tr>
              <td align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">序号:</span></td>
              <td align="left"><input name="no1" type="text" id="no1" value="<%=rs.recordcount+1%>" size="40" />
                  <span class="red">*</span></td>
            </tr>
            <tr>
              <td width="25%" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">帐号:</span></td>
              <td align="left"><input name="name1" type="text" id="name1" size="40" />
                  <span class="red">*</span></td>
            </tr>
            <tr>
              <td align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">说明:</span></td>
              <td align="left"><input name="info" type="text" id="info" value="" size="40" /></td>
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
 sql="select * from cserver where id="&request("id")
 rs.Open sql,conn,1,3
 %>
        <form id="form" name="form" method="post" action="?id=<%=rs("id")%>&amp;tj=Modify">
          <strong>修改客服</strong><br /><br />
          <table width="600" border="1" cellspacing="0" cellpadding="3"  bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" >
            <tr>
              <td align="center" bgcolor="#F0F0F0" class="STYLE1">类别:</td>
              <td align="left">
				<%for z=0 to ubound(sort1)%>
				 <label style="white-space:nowrap;<%if rs("sort1")=sort1(z) then%>color:#FF0000;<%end if%>">
				 <input type="radio" class="dq" name="sort1" id="sort1<%=z%>" value="<%=sort1(z)%>" onpropertychange="red(this)" <%if rs("sort1")=sort1(z) then%>checked="checked"<%end if%>/>
				 <%=sort1(z)%></label>
				<%next%>
			  <span class="red">*</span>			  </td>
            </tr>
            <tr>
              <td align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">序号:</span></td>
              <td align="left"><input name="no1" type="text" id="no1" value="<%=rs("no1")%>" size="40" />
                <span class="red">*</span></td>
            </tr>
            <tr>
              <td width="25%" align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">帐号:</span></td>
              <td align="left"><input name="name1" type="text" id="name1" value="<%=rs("name1")%>" size="40" />
                <span class="red">*</span></td>
            </tr>
            <tr>
              <td align="center" bgcolor="#F0F0F0" class="STYLE1"><span class="STYLE2">说明:</span></td>
              <td align="left"><input name="info" type="text" id="info" value="<%=rs("info")%>" size="40" /></td>
            </tr>
          </table>
         <!--#include file="inc/button.asp"-->
        </form>
      <%end if%>
</div>
<!--#include file="foot.asp"-->
</body>
</html>