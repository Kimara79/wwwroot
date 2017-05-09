<br>
<form action="?" method="post" name="turn_page" id="turn_page" style="width:96%;">
共有<font color="red"><%=rs.recordcount%></font><%=pagedw%><%=danwei%>条&nbsp;&nbsp;
每页<font color="red"><%=rs.pagesize%></font><%=pagedw%><%=danwei%>条&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

&lt;&lt;
<%if currentpage>1 then%>
<a href="?page=1"><strong>首页</strong></a>
<%else%>
首页
<%end if%>
&nbsp;&nbsp;
<%if currentpage>1 then%>
<a href="?page=<%=currentpage-1%>"><strong>上一页</strong></a>
<%else%>
上一页
<%end if%>
&nbsp;&nbsp;
第<font color="red"><%=currentpage%></font>/<font color="red"><%=rs.pagecount%></font>页
&nbsp;&nbsp;
<%if currentpage<rs.pagecount then%>
<a href="?page=<%=currentpage+1%>"><strong>下一页</strong></a>
<%else%>
下一页
<%end if%>
&nbsp;&nbsp;
<%if currentpage<rs.pagecount then%>
<a href="?page=<%=rs.pagecount%>"><strong>尾页</strong></a>
<%else%>
尾页
<%end if%>
&gt;&gt;

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;转到
<input name="page" type="text" id="page" value="<%=currentpage%>" style="width:40px;" maxlength="4" />&nbsp;&nbsp;
<input type="submit" name="Submit" value="GO" />
</form>
<br>