<div class="tba_1">最新消息</div>
<div class="tba_2">
	<div style="height:88px;">
	<div align="left" style="padding:7px;line-height:180%;padding-left:15px;">
	<%
	sql="select top 4 * from Article where menu2=29 order by shezhi desc,no1"
	rs1.open sql,conn,1,1
	Do while not rs1.eof
	%>
	<img src="images/list5.gif" align="absmiddle" />&nbsp;&nbsp;<a href="ArticleShow.asp?id=<%=rs1("id")%>" class="a_hs" title="<%=rs1("name1")%>"><%=num_write(rs1("name1"),10)%></a><br />
	<%
	rs1.movenext
	loop
	rs1.close
	%>
	</div>
	</div>
	
	<div class="pic_line"><img src="images/pic_line.gif" /></div>
	
	<div align="center">
	<%
	rsstr=""
	sql="select info from system where id=17"
	rsstr=trim(conn.execute(sql).getstring)
	rsstr=replace(""&rsstr,chr(13),"")
	%>
	<table border="0" cellspacing="0" cellpadding="0">
	  <tr>
		<td valign="middle"><a href="<%=rsstr%>" target="_blank" class="a_hs">西二博客</a></td>
		<td valign="middle"><a href="<%=rsstr%>" target="_blank"><img src="images/blog.jpg" border="0" style="margin-left:10px;" /></a></td>
	  </tr>
	</table>
	</div>
</div>
<div class="tba_3"></div>
<div style="clear:both;"></div>

<div class="tba_1 margin1">积分排行</div>
<div class="tba_2">
	<div style="height:5px;"></div>
	<%pic_menu1="b"%>
	<table border="0" cellspacing="0" cellpadding="0" align="center">
	  <tr>
		<td width="50" align="left" class="color_bl">姓名</td>
		<td width="50" align="left" class="color_bl">卡号</td>
		<td width="50" align="right" class="color_bl">积分</td>
	  </tr>
	</table>
	<div class="pic_line"><img src="images/pic_line.gif" /></div>
	<div align="center" id="demo_<%=pic_menu1%>_" style="overflow:hidden;height:60px;">
		<table border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td align="center">
			</td>
		  </tr>
		  <tr>
			<td id="demo_<%=pic_menu1%>_1" align="center">
				<table border="0" cellspacing="0" cellpadding="0">
					<%
					sql="select top 6 * from users order by score desc"
					rs1.open sql,conn,1,1
					Do while not rs1.eof
					%>
				  <tr>
					<td width="50" align="left" class="color_hs"><%=num_write3(rs1("name1"))%></td>
					<td width="50" align="left" class="color_hs"><%=num_write2(rs1("card"))%></td>
					<td width="50" align="right" class="color_hs"><%=rs1("score")%></td>
				  </tr>
					<%
					rs1.movenext
					loop
					rs1.close
					%>
				</table>
			</td>
		  </tr>
		  <tr>
			<td id="demo_<%=pic_menu1%>_2" align="left"></td>
		  </tr>
		</table>
	</div>
	<SCRIPT>
	var speed_<%=pic_menu1%>=20;
	document.getElementById("demo_<%=pic_menu1%>_2").innerHTML=document.getElementById("demo_<%=pic_menu1%>_1").innerHTML;
	function Marquee_<%=pic_menu1%>(){
		if(document.getElementById("demo_<%=pic_menu1%>_2").offsetTop-document.getElementById("demo_<%=pic_menu1%>_").scrollTop<=0)
			document.getElementById("demo_<%=pic_menu1%>_").scrollTop-=document.getElementById("demo_<%=pic_menu1%>_1").offsetHeight;
		else{
			document.getElementById("demo_<%=pic_menu1%>_").scrollTop++;
		}
	}
	var MyMar_<%=pic_menu1%>=setInterval(Marquee_<%=pic_menu1%>,speed_<%=pic_menu1%>)
	document.getElementById("demo_<%=pic_menu1%>_").onmouseover=function() {clearInterval(MyMar_<%=pic_menu1%>)}
	document.getElementById("demo_<%=pic_menu1%>_").onmouseout=function() {MyMar_<%=pic_menu1%>=setInterval(Marquee_<%=pic_menu1%>,speed_<%=pic_menu1%>)}
	</SCRIPT>
	
</div>
<div class="tba_3"></div>
<div style="clear:both;"></div>