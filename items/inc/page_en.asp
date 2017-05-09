<br>
<form action="?" method="post" name="turn_page" id="turn_page" style="width:96%;">
Total&nbsp;<font color="red"><%=rs.recordcount%></font>&nbsp;&nbsp;
&nbsp;&nbsp;<font color="red"><%=rs.pagesize%></font>/Page&nbsp;&nbsp;&nbsp;&nbsp;

&lt;&lt;
<%if currentpage>1 then%>
<a href="?page=1"><strong>First</strong></a>
<%else%>
First
<%end if%>
&nbsp;&nbsp;
<%if currentpage>1 then%>
<a href="?page=<%=currentpage-1%>"><strong>Previous</strong></a>
<%else%>
Previous
<%end if%>
&nbsp;&nbsp;
<font color="red"><%=currentpage%></font>/<font color="red"><%=rs.pagecount%></font>
&nbsp;&nbsp;
<%if currentpage<rs.pagecount then%>
<a href="?page=<%=currentpage+1%>"><strong>Next</strong></a>
<%else%>
Next
<%end if%>
&nbsp;&nbsp;
<%if currentpage<rs.pagecount then%>
<a href="?page=<%=rs.pagecount%>"><strong>Last</strong></a>
<%else%>
Last
<%end if%>
&gt;&gt;

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Turn to
<input name="page" type="text" id="page" value="<%=currentpage%>" style="width:40px;" maxlength="4" />&nbsp;&nbsp;
<input type="submit" name="Submit" value="GO" />
</form>