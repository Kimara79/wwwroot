<div align="center" id="div4" style="width:100%;">
	<div style="margin-top:40px;width:1000px;" align="right">
	<table width="38" border="0" cellspacing="0" cellpadding="0">
	  <tr>
		<td background="images/ic/24w_c1<%=Request.Cookies("lg")%>.gif" height="<%if Request.Cookies("lg")="" then%>43<%else%>37<%end if%>"></td>
	  </tr>
	  <tr>
		<td background="images/ic/24w_c2.gif" align="center">
		<%
		sql="select * from cserver order by sort1,no1"
		rs.open sql,conn,1,1
		icname="24w"
		Do While Not rs.Eof
		  for z=0 to ubound(sort1)
			if sort1(z)=rs("sort1") then
			  exit for
			end if
		  next
		%>
		<a onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('ImageIC_<%=icname%>_<%=rs("id")%>','','images/ic/<%=icname%>_a<%=z%>.gif',1)" href="<%=replace(sort1_lk(z),"_____",rs("name1"))%>" style="white-space:nowrap;" title="<%if rs("info")<>"" then%><%=rs("info")%><%else%><%=rs("name1")%><%end if%>"><IMG src="images/ic/<%=icname%>_<%=z%>.gif" name="ImageIC_<%=icname%>_<%=rs("id")%>" border="0" style="margin:1px;" /></a><br />
		<img src="images/ic/<%=icname%>_a<%=z%>.gif" style="display:none;" />
		<%
		rs.MoveNext
		Loop
		rs.close
		%>
		</td>
	  </tr>
	  <tr>
		<td background="images/ic/24w_c3.gif" height="12"></td>
	  </tr>
	</table>
	</div>
</div>