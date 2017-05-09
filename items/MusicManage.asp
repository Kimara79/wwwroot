<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<% menuname="背景音乐" %>
<% menunamel="音乐" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<title><%=menuname%></title>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="inc/power.asp"-->
<!--#include file="js/red.js"-->

<%
'其他设置
sql="select info from system where id=23"
rsstr=trim(conn.execute(sql).getstring)
rsstr=replace(rsstr,chr(13),"")
shezhi=split(rsstr,",")

'删除操作
if request("cz")="del" then
	On Error Resume Next
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	objFSO.DeleteFile(Server.MapPath("music/"&request("file1")))
	set objFSO=nothing

	sql="delete * from music where id="&request("id")               '删除数据
	conn.execute sql,1,1 
	
	rsrz.open "book",conn,1,3                                   '添加日志              
	rsrz.addnew
	rsrz("zhanghao")=session("admin")&"("&Session("truename")&")"
	rsrz("caozuo")="删除"
	rsrz("menu")=menuname
	rsrz("duixian")=request("name1")
	rsrz.update
	rsrz.close
	%>
	
	<!--#include file="music_inc.asp"-->	
	
	<%	
	response.redirect "?"
	Response.End
end if
%>

<%
'重新生成播放列表
if request("cz")="list" then
%>
	<!--#include file="music_inc.asp"-->	
<%
	Response.Write("<script>alert('生成完毕！');location.href='?';</script>")
	Response.End
end if
%>

<%
if request("xs")="all" then
	session("music_all")="1"
	session("music_search_type")=""
	session("music_search_key")=""
end if

if request("search_key")<>"" then
	session("music_all")="0"
	session("music_search_type")=request("search_type")
	session("music_search_key")=request("search_key")
end if

if request("page")<>"" then
	session("music_page")=request("page")
end if
if session("music_page")="" then
	session("music_page")=1
end if

sqladd=""
if session("music_search_key")<>"" then
	sqladd=sqladd&" and "&session("music_search_type")&" like '%"&session("music_search_key")&"%' "
end if

'预设排序
sql="update music set no1=id where (no1 is null or no=null or no1=0)"
conn.execute sql,1,1 

'接收排序
if request("tj")="order" and int(request("id1"))>0 and int(request("id2"))>0 then
	sql="update music set no1="&request("no2")&" where id="&request("id1")&" "
	conn.execute sql,1,1
	sql="update music set no1="&request("no1")&" where id="&request("id2")&" "
	conn.execute sql,1,1
	response.redirect "?"
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
<br>

<center>
  <form id="form1" name="form1" method="post" action="?xs=all">
    <table border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td class="STYLE1">搜索：&nbsp;&nbsp;</td>
        <td>
		<select name="search_type" id="search_type">
            <option value="name1" <%if session("music_search_type")="name1" then%>selected="selected"<%end if%>><%=menunamel%>名称</option>
			<option value="info" <%if session("music_search_type")="info" then%>selected="selected"<%end if%>>备注</option>
        </select>
	    </td>
        <td>&nbsp;&nbsp;<input name="search_key" type="text" id="search_key" value="<%=session("music_search_key")%>" onKeyPress="maskEdit(/^[\w]*$/)" /></td>
        <td align="right">&nbsp;&nbsp;<input type="submit" name="Submit22" value=" 搜索 " /></td>
        <td align="right">&nbsp;&nbsp;<input type="button" name="Submit2x2" onClick="location.href='?xs=all'" value="清除搜索" /></td>
        <td align="right">&nbsp;&nbsp;<input type="button" name="Submit2x2" onClick="location.href='?cz=list'" value="重新生成播放列表" /></td>
        <td align="right">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="Submit2" onClick="location.href='musicAM.asp?cz=Add'" value=" 添加<%=menunamel%> " /></td>
     </tr>
    </table>
  </form>
</center>
<br>

<center>
<%
sql="select * from music where id>0 "&sqladd&" order by shezhi desc,no1"
rs.open sql,conn,1,1
%>
  <%if rs.bof then%>
  <p align="center">暂无信息</p>
  <%
  else
	rs.pagesize=20  '设置每页显示数目
	currentpage=Clng(session("music_page"))
	if currentpage<1 then currentpage=1
	if currentpage>rs.pagecount then currentpage=rs.pagecount
	rs.absolutepage=currentpage
  %>

  <table border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
	<tr align="center" bgcolor="#F0F0F0" class="STYLE1">
	  <td><strong><%=menunamel%>名称</strong></td>
	  <td><strong>上传文件</strong></td>
	  <td><strong>外链文件</strong></td>
	  <td><strong>备注</strong></td>
	  <td><strong>设置</strong></td>
	  <td><strong>操作</strong></td>
	</tr>
<%
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
%>
<form id="change<%=rs("id")%>" name="change<%=rs("id")%>" method="post" action="update.asp?id=<%=rs("id")%>&tb=music&zd=shezhi" target="update">
        <tr onMouseOver="bgColor='#D2F5FF';" onMouseOut="bgColor='#FFFFFF';">
          <td align="left">&nbsp;<%=rs("name1")%>&nbsp;</td>
          <td align="left">&nbsp;<%=rs("file1")%>&nbsp;</td>
          <td align="left">&nbsp;<%=rs("url")%>&nbsp;</td>
          <td align="left">&nbsp;<%=rs("info")%>&nbsp;</td>
	  <td align="center">
		<%for z=0 to ubound(shezhi)%>
		<label <%if instr(rs("shezhi"),shezhi(z)) then%>style="color:red;"<%end if%>>
		<input name="shezhi" type="checkbox" class="dq" value="<%=shezhi(z)%>" <%if instr(""&rs("shezhi"),shezhi(z)) then%> checked="checked" <%end if%> onClick="red(this);document.getElementById('change<%=rs("id")%>').submit();"/> <%=shezhi(z)%></label>
		<%next%>		  </td>
          <td align="center">&nbsp;&nbsp;
			<a href="?id=<%=rs("id")%>&cz=del&name1=<%=rs("name1")%>&file1=<%=rs("file1")%>" onClick="{if(confirm('确定要删除么?')){this.document.album.submit();return true;}return false;}">删除</a>&nbsp;&nbsp;
			<a href="musicAM.asp?id=<%=rs("id")%>&amp;cz=Modify">修改</a>&nbsp;&nbsp;
			<a href="?tj=order&id1=<%=rs("id")%>&no1=<%=rs("no1")%>&id2=<%=id2a%>&no2=<%=no2a%>">上移</a>&nbsp;&nbsp;
			<a href="?tj=order&id1=<%=rs("id")%>&no1=<%=rs("no1")%>&id2=<%=id2b%>&no2=<%=no2b%>">下移</a>&nbsp;&nbsp;
		  </td>
        </tr>
</form>
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
<iframe id="update" name="update" style="display:none;"></iframe>
<!--#include file="foot.asp"-->
</body>
</html>