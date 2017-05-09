<%
sql="select * from cserver order by sort1,no1"
rs1.open sql,conn,1,1
icname="24w"
Do While Not rs1.Eof
  for z=0 to ubound(sort1)
	if sort1(z)=rs1("sort1") then
	  exit for
	end if
  next
%><a onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('ImageIC_<%=icname%>_<%=rs1("id")%>','','images/ic/<%=icname%>_a<%=z%><%=Request.Cookies("lg")%>.gif',1)" href="<%=replace(sort1_lk(z),"_____",rs1("name1"))%>" style="white-space:nowrap;" title="<%if rs1("info")<>"" then%><%=rs1("info")%><%else%><%=rs1("name1")%><%end if%>"><IMG src="images/ic/<%=icname%>_<%=z%><%=Request.Cookies("lg")%>.gif" name="ImageIC_<%=icname%>_<%=rs1("id")%>" border="0" style="margin-left:5px;margin-right:5px;" /></a><img src="images/ic/<%=icname%>_a<%=z%><%=Request.Cookies("lg")%>.gif" style="display:none;" /><%
rs1.MoveNext
Loop
rs1.close
%>