<%
sql="select count(id) from menu2 where menu1="&menu1_id&" "
rsnum=int(conn.execute(sql).getstring)
leftmenu1id=menu1_id
leftname1=menu1_name1
%>
<div class="tba_1"><%=leftname1%></div>
<div class="tba_2" align="center">
	<div style="height:1px;"></div>
	<div style="width:90%;">
	<%
	sql="select * from menu2 where menu1="&leftmenu1id&" order by no1"
	rs.open sql,conn,1,1
	Do while not rs.Eof
	tempurl=menuurl(rs("leixing"),2,rs("id"))
	%>
	<li class="leftli" onclick="location.href='<%=tempurl%>'"><a href="<%=tempurl%>" style="color:#FFFFFF;"><%=rs("name1")%></a></li>
	<%
	rs.movenext
	Loop
	rs.close
	%>
	<%if rsnum=0 then%>
		<%
		sql="select * from picture where menu1="&leftmenu1id&" order by id desc"
		rs.open sql,conn,1,1
		Do while not rs.Eof
		tempurl="PictureShow.asp?id="&rs("id")
		%>
		<li class="leftli" onclick="location.href='<%=tempurl%>'"><a href="<%=tempurl%>" style="color:#FFFFFF;"><%=rs("name1")%></a></li>
		<%
		rs.movenext
		Loop
		rs.close
		%>
	<%end if%>
	</div>
	<div style="height:3px;"></div>
</div>
<div class="tba_3"></div>
<div style="clear:both;"></div>
