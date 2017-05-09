<!----页底----->
<div align="center">
<table border="0" cellspacing="0" cellpadding="0" class="bodyw" style="margin-top:15px;">
  <tr>
    <td class="footbga_1" width="200" valign="middle" align="center">
		<%
		rsstr=""
		sql="select info from system where id=18"
		rsstr=trim(conn.execute(sql).getstring)
		rsstr=replace(""&rsstr,chr(13),"")
		%>
		<a href="<%=rsstr%>" target="_blank"><img src="images/taobao.gif" border="0" /></a>
	</td>
    <td class="footbga_2" valign="middle" width="1"><img src="images/foot_line.gif" /></td>
    <td class="footbga_2" valign="middle" align="center">
		<table border="0" cellspacing="0" cellpadding="0" style="cursor:pointer;" class="footad">
		  <tr>
			<td height="55" colspan="3" onclick="location.href='http://www.xzx2.com/ArticleShow.asp?id=95'"></td>
		  </tr>
		  <tr>
			<td height="45" width="163" onclick="location.href='#'"></td>
			<td width="165" onclick="location.href='#'"></td>
			<td onclick="location.href='#'"></td>
		  </tr>
		  <tr>
			<td align="center" colspan="3">
				<%
				sql="select * from link where shenghe='通过' order by no1"
				rs.open sql,conn,1,1
				i=0
				Do While Not rs.Eof
				i=i+1
				%>
				<a href="<%=rs("link")%>" class="a_hs" target="_blank"><%=rs("name1")%></a>&nbsp;&nbsp;<%if (i mod 5)=0 then%><br /><%end if%>
				<%
				rs.MoveNext
				Loop
				rs.close
				%>
			</td>
		  </tr>
		</table>
	</td>
    <td class="footbga_2" valign="middle" width="1"><img src="images/foot_line.gif" /></td>
    <td class="footbga_3" width="200" style="color:#000000;">
		<iframe width="190" height="115" class="share_self" frameborder="0" scrolling="no" src="http://widget.weibo.com/weiboshow/index.php?language=&width=190&height=115&fansRow=2&ptype=1&speed=0&skin=1&isTitle=0&noborder=1&isWeibo=0&isFans=1&uid=2112224175&verifier=8feb7de9&dpc=1"></iframe>		
	</td>
  </tr>
</table>
</div>

<div align="center" class="margin1 color_foot">
	<%
	i=0
	for z=0 to 1
	a=2-z
	sql="select * from menu"&a&" where shezhi like '%页面底部%' order by no1"
	rs1.open sql,conn,1,1
	Do while not rs1.Eof
	i=i+1
	if rs1("url")<>"" then
		tempurl=rs1("url")
	else
		tempurl=menuurl(rs1("leixing"),a,rs1("id"))
	end if
	%>
	<%if i>1 then%>&nbsp;&nbsp;┊&nbsp;&nbsp;<%end if%><a href="<%=tempurl%>" class="a_gy"><%=rs1("name1")%></a>
	<%
	rs1.MoveNext
	Loop
	rs1.close
	next
	%>
</div>

<table border="0" cellspacing="0" cellpadding="0" class="margin1">
  <tr>
	<td><script type="text/javascript">var cnzz_protocol = (("https:" == document.location.protocol) ? " https://" : " http://");document.write(unescape("%3Cspan id='cnzz_stat_icon_5940447'%3E%3C/span%3E%3Cscript src='" + cnzz_protocol + "s4.cnzz.com/stat.php%3Fid%3D5940447%26show%3Dpic1' type='text/javascript'%3E%3C/script%3E"));</script></td>
	<td class="color_foot">
	&nbsp;&nbsp;版权所有&nbsp;&nbsp;&copy;
	<%=year(now())%>&nbsp;&nbsp;
	<a href="<%=site_url2%>" class="a_gy"><%=site_name%></a>&nbsp;&nbsp;
	<%
	rsstr=""
	sql="select info from menu1 where id=25"
	rsstr=trim(conn.execute(sql).getstring)
	rsstr=replace(""&rsstr,chr(13),"")
	response.Write rsstr
	%>&nbsp;&nbsp;
	<%
	rsstr=""
	sql="select info from system where id=13"
	rsstr=trim(conn.execute(sql).getstring)
	rsstr=replace(""&rsstr,chr(13),"")
	response.Write rsstr
	%>
	</td>
  </tr>
</table>

<div align="center" style="margin-top:15px;"><img src="images/footlogo.png" /></div>
</center>
<%
conn.close
set conn=nothing
%>