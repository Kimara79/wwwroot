<center>
<table border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr>
    <td align="center" bgcolor="#F0F0F0" width="100"><strong>基本检索：</strong></td>
    <td align="left">
	&nbsp;&nbsp;
	<%
	'全部
	 sql="select count(id) from "&toptb&" "
	 rsnum=int(conn.execute(sql).getstring)
	%>  
	 <a href="?xs=all"><span <%if session(toptb&"_all")="1" then%>style="color:#FF0000;"<%end if%>>全部</span></a>(<span style="color:#FF0000"><%=rsnum%></span>)
	<%
	'特别设置
	for z=0 to ubound(shezhi)
	   sql="select count(id) from "&toptb&" where shezhi like '%"&shezhi(z)&"%' "
	   rsnum=int(conn.execute(sql).getstring)
	%>
	&nbsp;|&nbsp; <a href="?xs=all&sz=<%=shezhi(z)%>"><span <%if session(toptb&"_sz")=shezhi(z) then%>style="color:#FF0000;"<%end if%>><%=shezhi(z)%></span></a>(<span style="color:#FF0000"><%=rsnum%></span>)
	<%next%>
	&nbsp;&nbsp;
	</td>
  </tr>
<%
sql="select * from menu1 where leixing like '%"&likes&"%' order by no1"
rs.open sql,conn,1,1
if rs.recordcount>0 then
%>
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong>一级栏目：</strong></td>
    <td align="left">
	&nbsp;&nbsp;
	  <%
	  Do while not rs.Eof
		sql="select count(id) from "&toptb&" where menu1="&rs("id")&" "
		rsnum=int(conn.execute(sql).getstring)
	  %>
	  <a href="?xs=all&menu1=<%=rs("id")%>"><span <%if session(toptb&"_menu1")<>"" then%><%if rs("id")=int(session(toptb&"_menu1")) then%>style="color:#FF0000;"<%end if%><%end if%>><%=rs("name1")%></span></a>(<span style="color:#FF0000"><%=rsnum%></span>)&nbsp;&nbsp;|&nbsp;&nbsp;
	  <%
	   rs.MoveNext
	   Loop
	  %>
	&nbsp;&nbsp;
	</td>
  </tr>
<%
end if
rs.close
%>
<%
sql="select * from menu1 order by no1 "
rs1.open sql,conn,1,1
Do while not rs1.Eof
  sql="select count(id) from menu2 where leixing like '%"&likes&"%' and menu1="&rs1("id")&" "
  rsnum=int(conn.execute(sql).getstring)
  if rsnum>0 then
%> 
  <tr>
    <td align="center" bgcolor="#F0F0F0"><strong><%=rs1("name1")%>：</strong></td>
    <td align="left">
	&nbsp;&nbsp;
	  <%
	  sql="select * from menu2 where leixing like '%"&likes&"%' and menu1="&rs1("id")&" order by no1"
	  rs.open sql,conn,1,1
	  Do while not rs.Eof
		sql="select count(id) from "&toptb&" where menu2="&rs("id")&" "
		rsnum=int(conn.execute(sql).getstring)
	  %>
	  <a href="?xs=all&menu2=<%=rs("id")%>"><span <%if session(toptb&"_menu2")<>"" then%><%if rs("id")=int(session(toptb&"_menu2")) then%>style="color:#FF0000;"<%end if%><%end if%>><%=rs("name1")%></span></a>(<span style="color:#FF0000"><%=rsnum%></span>)&nbsp;&nbsp;|&nbsp;&nbsp;
	  <%
	   rs.MoveNext
	   Loop
	   rs.close
	  %>
	&nbsp;&nbsp;
	</td>
  </tr>
<%
  end if
rs1.MoveNext
Loop
rs1.close
%>
</table>
</center>
<br>
