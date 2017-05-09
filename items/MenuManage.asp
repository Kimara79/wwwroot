<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="栏目管理" %>
<% menunamel="栏目" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="js/red.js"-->
<!--#include file="inc/FilterHtml.asp"-->

<%
'其他设置
sql="select info from system where id=5"
rsstr=trim(conn.execute(sql).getstring)
rsstr=replace(rsstr,chr(13),"")
shezhi=split(rsstr,",")

'删除操作
if request("cz")="del" then
	On Error Resume Next
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	'删除栏目图标
	if request("pic")<>"" then
		objFSO.DeleteFile(Server.MapPath("pic/picture/"&request("pic")))
		objFSO.DeleteFile(Server.MapPath("pic/s_picture/"&request("pic")))
	end if
	'删除二级栏目图标
	if request("tb")="menu1" then
		sql="select pic from menu2 where menu1="&request("id")&" and (pic is not null and pic<>null and pic<>'') "
		rs.open sql,conn,1,1
		Do while not rs.Eof
			objFSO.DeleteFile(Server.MapPath("pic/picture/"&rs("pic")))
			objFSO.DeleteFile(Server.MapPath("pic/s_picture/"&rs("pic")))
		rs.MoveNext
		Loop
		rs.close
	end if
	'删除栏目图片
	sql="select pic from picture where "&request("tb")&"="&request("id")&" "
	rs.open sql,conn,1,1
	Do while not rs.Eof
		objFSO.DeleteFile(Server.MapPath("pic/picture/"&rs("pic")))
		objFSO.DeleteFile(Server.MapPath("pic/s_picture/"&rs("pic")))
	rs.MoveNext
	Loop
	rs.close
	sql="delete * from "&request("tb")&" where id="&request("id")
	conn.execute sql,1,1
	sql="delete * from article where "&request("tb")&"="&request("id")&" "
	conn.execute sql,1,1
	sql="delete * from picture where "&request("tb")&"="&request("id")&" "
	conn.execute sql,1,1
	sql="delete * from forms where "&request("tb")&"='_"&request("id")&"_' "
	conn.execute sql,1,1
	Set objFSO = nothing
	
	rsrz.open "book",conn,1,3
	rsrz.addnew
	rsrz("zhanghao")=session("admin")&"("&session("truename")&")"
	rsrz("caozuo")="删除"
	rsrz("menu")=menuname
	if request("tb")="menu1" then
		rsrz("duixian")="一级"&menunamel&"【"&request("name1")&"】及其子"&menunamel&"，以及相应的文章、图片"
	else
		rsrz("duixian")="二级"&menunamel&"【"&request("name1")&"】，以及相应的文章、图片"
	end if
	rsrz.update
	rsrz.close

	response.redirect "?"
	response.end
end if

'预设排序
sql="update menu1 set no1=id where (no1 is null or no=null or no1=0)"
conn.execute sql,1,1 
sql="update menu2 set no1=id where (no1 is null or no=null or no1=0)"
conn.execute sql,1,1 

'接收排序
if request("tj")="order" and int(request("id1"))>0 and int(request("id2"))>0 then
	sql="update "&request("tb")&" set no1="&request("no2")&" where id="&request("id1")&" "
	conn.execute sql,1,1
	sql="update "&request("tb")&" set no1="&request("no1")&" where id="&request("id2")&" "
	conn.execute sql,1,1
	response.redirect "?"
	response.end
end if

session("pic")=""
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

<table border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr>
	<td align="center" bgcolor="#F0F0F0" class="STYLE1">ID</td>
	<td align="center" bgcolor="#F0F0F0" class="STYLE1">显示位置</td>
	<td align="center" bgcolor="#F0F0F0" class="STYLE1">&nbsp;<%=menunamel%>图标&nbsp;</td>
	<td align="center" bgcolor="#F0F0F0" class="STYLE1">&nbsp;<%=menunamel%>名称&nbsp;</td>
	<td align="center" bgcolor="#F0F0F0" class="STYLE1">&nbsp;<%=menunamel%>类型&nbsp;</td>
	<td align="center" bgcolor="#F0F0F0" class="STYLE1">&nbsp;栏目内容&nbsp;</td>
	<td align="center" bgcolor="#F0F0F0" class="STYLE1">&nbsp;其他设置&nbsp;</td>
	<td align="center" bgcolor="#F0F0F0" class="STYLE1">&nbsp;特别链接&nbsp;</td>
	<td align="center" bgcolor="#F0F0F0" class="STYLE1">更新时间</td>
	<td align="center" bgcolor="#F0F0F0" class="STYLE1">更新人</td>
	<td align="center" bgcolor="#F0F0F0" class="STYLE1">操作</td>
  </tr>
<%
sql="select * from menu1 order by no1 "
rs.open sql,conn,1,1
Do while not rs.Eof
  rs.moveprevious
  if not rs.bof then
   id2a=rs("id")
   no2a=rs("no1")
  else
   id2a=0
   no2a=0
  end if
  rs.movenext
  rs.movenext
  if not rs.eof then
   id2b=rs("id")
   no2b=rs("no1")
  else
   id2b=0
   no2b=0
  end if
  rs.moveprevious
  if instr(rs("weizhi"),"导航") then
  	namestyle="font-size:14px;color:#ff0000;font-weight:bold;"
  else
  	namestyle="font-size:12px;color:#0000ff;font-weight:bold;"
  end if
%>    
<form id="change<%=rs("id")%>" name="change<%=rs("id")%>" method="post" action="update.asp?id=<%=rs("id")%>&tb=menu1&zd=shezhi" target="update">
  <tr onmouseover="bgColor='#D2F5FF';" onmouseout="bgColor='#FFFFFF';">
	<td align="center">&nbsp;<%=rs("id")%>&nbsp;</td>
	<td align="center">&nbsp;<%=rs("weizhi")%>&nbsp;</td>
	<td align="center">&nbsp;<%if rs("pic")<>"" then%><a href="pic/picture/<%=rs("pic")%>" target="_blank"><img src="pic/s_picture/<%=rs("pic")%>" border="0" width="25" height="25" /></a><%end if%>&nbsp;</td>
	<td align="left" style="<%=namestyle%>">&nbsp;<%=rs("name1")%>&nbsp;</td>
	<td align="left">&nbsp;<%=rs("leixing")%>&nbsp;</td>
	<td align="left" <%if rs("info")<>"" then%>title="<%=left(FilterHtml(rs("info")),100)%>.."<%end if%>>&nbsp;<%if rs("info")<>"" then%><%=left(FilterHtml2(rs("info")),5)%>..<%end if%>&nbsp;</td>
	<td align="center">
	<%for z=0 to ubound(shezhi)%>
	<label <%if instr(rs("shezhi"),shezhi(z)) then%>style="color:red;"<%end if%>>
	<input name="shezhi" type="checkbox" class="dq" value="<%=shezhi(z)%>" <%if instr(""&rs("shezhi"),shezhi(z)) then%> checked="checked" <%end if%> onclick="red(this);document.getElementById('change<%=rs("id")%>').submit();"/> <%=shezhi(z)%></label>
	<%next%>
	</td>
	<td align="left">&nbsp;<%=rs("url")%>&nbsp;</td>
	<td align="center">&nbsp;<%=rs("lasttime")%>&nbsp;</td>
	<td align="center">&nbsp;<%=rs("laster")%>&nbsp;</td>
	<td align="left">&nbsp;&nbsp;
	<a href="?cz=del&tb=menu1&id=<%=rs("id")%>&name1=<%=rs("name1")%>&pic=<%=rs("pic")%>" onclick="{if(confirm('该栏目及子栏目其相应的文章、图片将全部被删除！')){this.document.album.submit();return true;}return false;}">删除</a>&nbsp;&nbsp;
	<a href="Menu1AM.asp?cz=Modify&id=<%=rs("id")%>">修改</a>&nbsp;&nbsp;
	<a href="?tj=order&tb=menu1&id1=<%=rs("id")%>&no1=<%=rs("no1")%>&id2=<%=id2a%>&no2=<%=no2a%>">上移</a>&nbsp;&nbsp;
	<a href="?tj=order&tb=menu1&id1=<%=rs("id")%>&no1=<%=rs("no1")%>&id2=<%=id2b%>&no2=<%=no2b%>">下移</a>&nbsp;&nbsp;	
	<a href="menu2AM.asp?cz=Add&menu1=<%=rs("id")%>">添加下级栏目</a>&nbsp;&nbsp;
	</td>
  </tr>
</form>
<%
sql="select * from menu2 where menu1="&rs("id")&" order by no1 "
rs2.open sql,conn,1,1
Do while not rs2.Eof
  rs2.moveprevious
  if not rs2.bof then
   id2a=rs2("id")
   no2a=rs2("no1")
  else
   id2a=0
   no2a=0
  end if
  rs2.movenext
  rs2.movenext
  if not rs2.eof then
   id2b=rs2("id")
   no2b=rs2("no1")
  else
   id2b=0
   no2b=0
  end if
  rs2.moveprevious
%>
<form id="change<%=rs2("id")%>" name="change<%=rs2("id")%>" method="post" action="update.asp?id=<%=rs2("id")%>&tb=menu2&zd=shezhi" target="update">
  <tr onmouseover="bgColor='#D2F5FF';" onmouseout="bgColor='#FFFFFF';">
	<td align="center">&nbsp;<%=rs2("id")%>&nbsp;</td>
	<td align="center">&nbsp;&nbsp;</td>
	<td align="center">&nbsp;<%if rs2("pic")<>"" then%><a href="pic/picture/<%=rs2("pic")%>" target="_blank"><img src="pic/s_picture/<%=rs2("pic")%>" border="0" width="25" height="25" /></a><%end if%>&nbsp;</td>
	<td align="left">&nbsp;--&nbsp;<%=rs2("name1")%>&nbsp;</td>
	<td align="center">&nbsp;<%=rs2("leixing")%>&nbsp;</td>
	<td align="left" <%if rs2("info")<>"" then%>title="<%=left(FilterHtml(rs2("info")),100)%>.."<%end if%>>&nbsp;<%if rs2("info")<>"" then%><%=left(FilterHtml2(rs2("info")),5)%>..<%end if%>&nbsp;</td>
	<td align="center">
	<%for z=0 to ubound(shezhi)%>
	<label <%if instr(rs2("shezhi"),shezhi(z)) then%>style="color:red;"<%end if%>>
	<input name="shezhi" type="checkbox" class="dq" value="<%=shezhi(z)%>" <%if instr(""&rs2("shezhi"),shezhi(z)) then%> checked="checked" <%end if%> onclick="red(this);document.getElementById('change<%=rs2("id")%>').submit();"/> <%=shezhi(z)%></label>
	<%next%>
	</td>
	<td align="left">&nbsp;<%=rs2("url")%>&nbsp;</td>
	<td align="center">&nbsp;<%=rs2("lasttime")%>&nbsp;</td>
	<td align="center">&nbsp;<%=rs2("laster")%>&nbsp;</td>
	<td align="left">&nbsp;&nbsp;
	<a href="?cz=del&tb=menu2&id=<%=rs2("id")%>&name1=<%=rs2("name1")%>&pic=<%=rs2("pic")%>" onclick="{if(confirm('该栏目及相应的文章、图片将全部删除?')){this.document.album.submit();return true;}return false;}">删除</a>&nbsp;&nbsp;
	<a href="menu2AM.asp?cz=Modify&id=<%=rs2("id")%>">修改</a>&nbsp;&nbsp;
	<a href="?tj=order&tb=menu2&id1=<%=rs2("id")%>&no1=<%=rs2("no1")%>&id2=<%=id2a%>&no2=<%=no2a%>">上移</a>&nbsp;&nbsp;
	<a href="?tj=order&tb=menu2&id1=<%=rs2("id")%>&no1=<%=rs2("no1")%>&id2=<%=id2b%>&no2=<%=no2b%>">下移</a>&nbsp;&nbsp;	
	</td>
  </tr>
</form>
<%
rs2.MoveNext
Loop
rs2.close
%>
<%
rs.MoveNext
Loop
rs.close
%>
</table>
<br />
<input type="button" name="Submit2" onclick="location.href='Menu1AM.asp?cz=Add'" value=" 添加一级栏目 " />
<br /><br />
</div>
<iframe id="update" name="update" style="display:none;"></iframe>
<!--#include file="foot.asp"-->
</body>
</html>