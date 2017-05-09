<div style="width:96%;" align="center">
Total&nbsp;<%=rs.recordcount%>&nbsp;&nbsp;&nbsp;&nbsp;
<%=rs.pagesize%> per page&nbsp;&nbsp;&nbsp;&nbsp;

<%if currentpage>1 then%><a href="?<%=requesttrp("page")%>page=1"><strong>First</strong></a><%else%>First<%end if%>
&nbsp;
<%if currentpage>1 then%><a href="?<%=requesttrp("page")%>page=<%=currentpage-1%>"><strong>Previous</strong></a><%else%>Previous<%end if%>
&nbsp;&nbsp;
<font color="red"><%=currentpage%></font>/<font color="red"><%=rs.pagecount%></font>
&nbsp;&nbsp;
<%if currentpage<rs.pagecount then%><a href="?<%=requesttrp("page")%>page=<%=currentpage+1%>"><strong>Next</strong></a><%else%>Next<%end if%>
&nbsp;
<%if currentpage<rs.pagecount then%><a href="?<%=requesttrp("page")%>page=<%=rs.pagecount%>"><strong>Last</strong></a><%else%>Last<%end if%>

&nbsp;&nbsp;&nbsp;Turn to
<input name="page" type="text" id="page" value="<%=currentpage%>" style="width:30px;" maxlength="3" onkeyup="value=value.replace(/\D/g,'');" />&nbsp;&nbsp;
<input type="submit" name="Submit" value=" GO " style="font-size:10px;" onclick="if(document.getElementById('page').value!=''){location.href='?<%=requesttrp("page")%>'+'page='+document.getElementById('page').value;}" />
</div>