<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="友情链接" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->

<%  
if request("cz")="del" then    '----------------------------------------------------删除判断	
  sql="delete * from link where id="&request("id")               '删除数据
  conn.execute sql,1,1 
   
  		rsrz.open "book",conn,1,3                                  '添加日志              
		rsrz.addnew
		  rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
		  rsrz("caozuo")="删除"
		  rsrz("menu")=menuname
		  rsrz("duixian")=request("link")
		rsrz.update
		rsrz.close
		
  conn.close
  response.redirect ("?")
  response.end
end if
%>
<%
if request("cz")="tg" then
  sql="select * from link where id="&request("id")
  rs.open sql,conn,1,3
   if request.form("shenghe")="通过" then
    rs("shenghe")="通过"
   else
    rs("shenghe")="未通过"
   end if
  rs.update
  rs.close

  response.redirect ("?")
  response.end
end if
%>

<body>
<!--#include file="top.asp"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23" align="center" background="images/am_mn2.gif" bgcolor="cccccc"><strong><%=menuname%></strong></td>
  </tr>
</table>
<div align="center"><br />

<%
sql="select * from link order by no1"
rs.open sql,conn,1,1
%>
      <table border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" width="900">
        <tr class="STYLE1">
          <td width="7%" align="center" bgcolor="#F0F0F0">序号</td>
          <td height="20" align="center" bgcolor="#F0F0F0">名称</td>
          <td align="center" bgcolor="#F0F0F0">网址</td>
          <td width="18%" align="center" bgcolor="#F0F0F0">操作</td>
        </tr>
        <% Do While Not rs.Eof %>
        <form id="shenghe<%=rs("id")%>" name="shenghe<%=rs("id")%>" method="post" action="?cz=tg&amp;id=<%=rs("id")%>">
          <tr onmouseover="bgColor='#D2F5FF';" onmouseout="bgColor='#FFFFFF';">
            <td align="center"><%=rs("no1")%></td>
            <td align="left">&nbsp;<%=rs("name1")%></td>
            <td align="left">&nbsp;<a href="<%=rs("link")%>" target="_blank"><%=rs("link")%></a></td>
            <td align="center"><a href="?id=<%=rs("id")%>&amp;cz=del&amp;link=<%=rs("link")%>" onclick="{if(confirm('确定要删除么?')){this.document.album.submit();return true;}return false;}">删除</a>&nbsp;&nbsp;<a href="linkAM.asp?id=<%=rs("id")%>&amp;cz=Modify">修改</a>&nbsp;&nbsp;
                <label <%if rs("shenghe")="通过" then%>style="color:#FF0000;"<%end if%>><input name="shenghe" type="checkbox" id="shenghe" value="通过" onclick="document.shenghe<%=rs("id")%>.submit();" <%if rs("shenghe")="通过" then%>checked="checked"<%end if%>/>
              通过</label>			  </td>
          </tr>
        </form>
        <%
rs.MoveNext
Loop
rs.close
%>
      </table>
      <br />
      <input type="button" name="Submit2" onclick="location.href='LinkAM.asp?cz=Add'" value=" 添 加 " />
      <br /><br />
</div>
<!--#include file="foot.asp"-->
</body>
</html>