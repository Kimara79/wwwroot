<div style="width:96%;margin-top:15px;" align="center">
共有<%=rs.recordcount%><%=pagedw%><%=danwei%>&nbsp;&nbsp;&nbsp;&nbsp;
每页<%=rs.pagesize%><%=danwei%>&nbsp;&nbsp;&nbsp;&nbsp;

<%if currentpage>1 then%><a href="?<%=requesttrp("page")%>page=1"><strong>首页</strong></a><%else%>首页<%end if%>
&nbsp;
<%if currentpage>1 then%><a href="?<%=requesttrp("page")%>page=<%=currentpage-1%>"><strong>上一页</strong></a><%else%>上一页<%end if%>
&nbsp;
第<font color="red"><%=currentpage%></font>/<font color="red"><%=rs.pagecount%></font>页
&nbsp;
<%if currentpage<rs.pagecount then%><a href="?<%=requesttrp("page")%>page=<%=currentpage+1%>"><strong>下一页</strong></a><%else%>下一页<%end if%>
&nbsp;
<%if currentpage<rs.pagecount then%><a href="?<%=requesttrp("page")%>page=<%=rs.pagecount%>"><strong>尾页</strong></a><%else%>尾页<%end if%>


&nbsp;&nbsp;&nbsp;转到
<input name="page" type="text" id="page" value="<%=currentpage%>" style="width:30px;" maxlength="3" onkeyup="value=value.replace(/\D/g,'');" onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/\D/g,''));" />&nbsp;
<input type="submit" name="Submit" value=" GO " style="font-size:11px;" onclick="if(document.getElementById('page').value!=''){location.href='?<%=requesttrp("page")%>'+'page='+document.getElementById('page').value;}" />
</div>
<br />