<script type="text/javascript" src="js/jquery-1.3.min.js"></script>
<!--[if IE 6]>
<script src="js/DD_belatedPNG_0.0.8a.js"></script>
<script>DD_belatedPNG.fix('*');</script>
<![endif]-->
<!--#include file="js/showmenu.js"-->
<!--#include file="js/favorite.js"-->
<!--#include file="ic_inc.asp"-->

<%
response.Cookies("qturl")=request.ServerVariables("URL")&"?"&Request.ServerVariables("QUERY_STRING")
%>

<img src="images/menua.png" style="display:none;" />
<img src="images/leftli2.gif" style="display:none;" />
<img src="images/menusub3.gif" style="display:none;" />
<img src="images/bt1a.png" style="display:none;" />
<img src="images/bt2a.png" style="display:none;" />

<center>
<div align="right" class="bodyw" style="padding:10px;padding-bottom:0px;">
	<form id="form_search" name="form_search" method="post" action="search.asp?">
	<%
	wordx="站内搜索"
	searchjs="if(document.getElementById('keyword').value!='' && document.getElementById('keyword').value!='"&wordx&"'){document.getElementById('form_search').action=document.getElementById('form_search').action+'keyword='+document.getElementById('keyword').value;document.getElementById('form_search').submit();}"
	%>
	<table border="0" cellspacing="0" cellpadding="0">
	  <tr>
		<td align="left" width="100"><img src="images/list1.gif" align="absmiddle" />&nbsp;&nbsp;<a href="#" onclick="AddFavorite()" class="a_small">加入收藏</a></td>
		<td><input type="text" name="keyword" id="keyword" class="search" maxlength="30" value="<%if keyword<>"" then%><%=keyword%><%else%><%=wordx%><%end if%>" onfocus="if(this.value=='<%=wordx%>'){this.value='';}" onblur="if(this.value.length===0){this.value='<%=wordx%>';}" onkeydown="if(event.keyCode==13){<%=searchjs%>}" /></td>
		<td><a href="#" onclick="<%=searchjs%>"><img src="images/searchbt.png" style="margin-left:5px;cursor:pointer;" border="0" /></a></td>
	  </tr>
	</table>
	</form>
</div>
<table border="0" cellspacing="0" cellpadding="0" class="bodywa top">
  <tr>
    <td align="left" valign="bottom"><img src="images/logo.png" style="margin-left:0px;margin-bottom:-3px;" /></td>
    <td align="right" valign="bottom">
	<table border="0" cellspacing="0" cellpadding="0" height="49" style="margin-bottom:4px;margin-right:20px;">
	  <tr>
		<td align="center" valign="middle" class="menutd" onclick="location.href='index1.asp'"><div class="menutdadd"></div><a href="index1.asp" class="a_menu">首页</a></td>
		<%
		sql="select * from menu1 where weizhi like '%导航%' order by no1"
		rs1.open sql,conn,1,1
		Do while not rs1.Eof
		%>
		<%
		sql="select * from menu2 where menu1="&rs1("id")&" order by no1"
		rs2.open sql,conn,1,1
		num=rs2.recordcount
			if rs1("url")<>"" then
				tempurl=rs1("url")
			else
				if rs2.eof then
					tempurl=menuurl(rs1("leixing"),1,rs1("id"))
				else
					tempurl=menuurl(rs2("leixing"),2,rs2("id"))
				end if
			end if
		%>
		<td align="left" valign="middle" class="menutd" onmouseover="dd_<%=rs1("id")%>.style.display='block';" onmouseout="dd_<%=rs1("id")%>.style.display='none';" onclick="location.href='<%=tempurl%>'">
		<div class="menutdadd"></div><a href="<%=tempurl%>" class="a_menu"><%=rs1("name1")%></a><br/>
		<%if not rs2.eof then%>
			<div id="dd_<%=rs1("id")%>" class="menusub">
				<div class="menusub12">
				<%
				Do while not rs2.eof
				if rs2("url")<>"" then
					tempurl=rs2("url")
				else
					tempurl=menuurl(rs2("leixing"),2,rs2("id"))
				end if
				%>
					<a href="<%=tempurl%>" class="a_menu2"><%=rs2("name1")%></a>
				<%
				rs2.movenext
				loop
				%>
				</div>
				<div class="menusub3"></div>
			</div>
		<%
		end if
		rs2.close
		%>
		<%if num=0 then%>
			<div id="dd_<%=rs1("id")%>" class="menusub">
				<div class="menusub12">
				<%
				sql="select * from Picture where menu1="&rs1("id")&" order by id"
				rs2.open sql,conn,1,1
				Do while not rs2.eof
					tempurl="PictureShow.asp?id="&rs2("id")
				%>
					<a href="<%=tempurl%>" class="a_menu2"><%=rs2("name1")%></a>
				<%
				rs2.movenext
				loop
				rs2.close
				%>
				</div>
				<div class="menusub3"></div>
			</div>
		<%end if%>
		</td>
		<%
		rs1.MoveNext
		Loop
		rs1.close
		%>
	  </tr>
	</table>
	</td>
  </tr>
</table>
