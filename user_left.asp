<div class="tba_1">会员中心</div>
<div class="tba_2" align="center">
	<div style="height:1px;"></div>
	
	<%if request.Cookies("userqt")("tel")<>"" then%>
		<div style="width:90%;">
		<li class="leftli" onclick="location.href='UserIndex.asp'">中心首页</li>
		<li class="leftli" onclick="location.href='UserLook.asp'">查看资料</li>
		<li class="leftli" onclick="location.href='UserXiaofei.asp'">消费记录</li>
		<li class="leftli" onclick="location.href='UserJifen.asp'">积分兑换</li>
		<%
		i=0
		for z=0 to 1
		a=2-z
		sql="select * from menu"&a&" where shezhi like '%会员中心%' order by no1"
		rs1.open sql,conn,1,1
		Do while not rs1.Eof
		i=i+1
		if rs1("url")<>"" then
			tempurl=rs1("url")
		else
			tempurl=menuurl(rs1("leixing"),a,rs1("id"))
		end if
		%>
		<li class="leftli" onclick="location.href='User_<%=tempurl%>'"><%=rs1("name1")%></li>
		<%
		rs1.MoveNext
		Loop
		rs1.close
		next
		%>
		<li class="leftli" onclick="location.href='UserPassword.asp'">修改密码</li>
		<li class="leftli" onclick="location.href='user_login_tj.asp?tj=out'">退出</li>
		</div>
	<%else%>
		<div align="center"><br />请先登录！<br /><br /></div>
	<%end if%>
	
	<div style="height:3px;"></div>
</div>
<div class="tba_3"></div>
<div style="clear:both;"></div>
