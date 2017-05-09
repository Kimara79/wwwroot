<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title>其他设置</title></head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="inc/num_write.asp"-->

<%
'预设排序
sql="update system set no1=id where (no1 is null or no=null or no1=0)"
conn.execute sql,1,1 

if request("cz")="del" then    '----------------------------------------------------删除判断	
  sql="delete * from system where id="&request("id")               '删除数据
  conn.execute sql,1,1 
   
  		rsrz.open "book",conn,1,3                                  '添加日志              
		rsrz.addnew
		  rsrz("zhanghao")=session("admin")&"("&session("truename")&")"
		  rsrz("caozuo")="删除"
		  rsrz("menu")="其他设置"  
		  rsrz("duixian")=request("name1")
		rsrz.update
		rsrz.close
		
  response.redirect "?"
  response.end
end if

'排序
if request("tj")="order1" and int(request("no1a"))>0 then
	sql="update system set no1="&request("no1a")&" where id="&request("id")&" "
	conn.execute sql,1,1
	sql="update system set no1="&request("no1")&" where id="&request("id1")&" "
	conn.execute sql,1,1
	response.redirect "?"
	response.end
end if
if request("tj")="order2" and int(request("no1b"))>0 then
	sql="update system set no1="&request("no1b")&" where id="&request("id")&" "
	conn.execute sql,1,1
	sql="update system set no1="&request("no1")&" where id="&request("id2")&" "
	conn.execute sql,1,1
	response.redirect "?"
	response.end
end if
%>

<body>
<!--#include file="top.asp"-->
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23" align="center" background="images/am_mn2.gif" bgcolor="cccccc"><strong>其他设置</strong></td>
  </tr>
</table>
<div align="center">
	<br />
<%
sql="select * from system order by no1"
rs.open sql,conn,1,1
%>
        <%if rs.bof then%>
        <p align="center">暂无信息</p>
      <%else %>
        <table border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
          <tr class="STYLE1">
            <td align="center" bgcolor="#F0F0F0" style="color:#FF0000; font-weight:bold;">&nbsp;ID&nbsp;</td>
            <td align="center" bgcolor="#F0F0F0">名称</td>
            <td align="center" bgcolor="#F0F0F0">内容</td>
            <td align="center" bgcolor="#F0F0F0">最后更新时间</td>
            <td align="center" bgcolor="#F0F0F0">最后更新人</td>
            <td align="center" bgcolor="#F0F0F0">操作</td>
          </tr>
<%
Do While Not rs.Eof
  rs.moveprevious
  if not rs.bof then
   id1=rs("id")
   no1a=rs("no1")
  else
   id1=0
   no1a=0
  end if
  rs.movenext
  rs.movenext
  if not rs.eof then
   id2=rs("id")
   no1b=rs("no1")
  else
   id2=0
   no1b=0
  end if
  rs.moveprevious
%>
          <tr align="center" onmouseover="bgColor='#D2F5FF';" onmouseout="bgColor='#FFFFFF';">
            <td style="color:#FF0000; font-weight:bold;"><%=rs("id")%></td>
            <td align="left">&nbsp;<%=rs("name1")%>&nbsp;</td>
            <td align="left">&nbsp;<%=num_write(rs("info"),25)%>&nbsp;</td>
            <td align="left">&nbsp;<%=rs("lasttime")%>&nbsp;</td>
            <td align="left">&nbsp;<%=rs("laster")%>&nbsp;</td>
            <td>&nbsp;&nbsp;
			<%if Session("admin")="sosovipp" then%><a href="?id=<%=rs("id")%>&amp;cz=del&amp;name1=<%=rs("name1")%>" onclick="{if(confirm('确定要删除么?')){this.document.album.submit();return true;}return false;}">删除</a>&nbsp;&nbsp;
			<a href="?tj=order1&id=<%=rs("id")%>&no1=<%=rs("no1")%>&id1=<%=id1%>&no1a=<%=no1a%>">上移</a>&nbsp;&nbsp;
			<a href="?tj=order2&id=<%=rs("id")%>&no1=<%=rs("no1")%>&id2=<%=id2%>&no1b=<%=no1b%>">下移</a>&nbsp;&nbsp;<%end if%>
			<a href="systemAM.asp?id=<%=rs("id")%>&amp;cz=Modify">修改</a>
			&nbsp;&nbsp;</td>
          </tr>
          <%
rs.MoveNext
Loop %>
  </table>
<%
end if
rs.close
%>
      <br />
      <%if Session("admin")="sosovipp" then%><input type="button" name="Submit2" onclick="location.href='systemAM.asp?cz=Add'" value=" 添 加 " /><br /><%end if%>
      <br />
</div>
<!--#include file="foot.asp"-->
</body>
</html>