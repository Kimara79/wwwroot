<table width="100%" height="95" border="0" align="center" cellpadding="0" cellspacing="0" background="images/am_top1_bg.jpg">
    <tr>
      <td align="left"><img src="images/am_top1.jpg" /></td>
    </tr>
</table>
<table width="100%" height="38" border="0" align="center" cellpadding="0" cellspacing="0" background="images/am_top2_bg.jpg">
    <tr>
      <td align="left"><img src="images/am_top2.jpg" /></td>
      <td align="right" class="font_foot">
	  <table width="80%" border="0" cellpadding="0" cellspacing="0">
		<tr>
          <td align="right" style="color:#FFFFFF;padding-top:12px;">欢迎您！&nbsp;&nbsp;<%=Session("step")%>级管理员&nbsp;&nbsp;<%=Session("admin")%>(<%=Session("truename")%>)　今天是<%=year(now())%>年<%=month(now())%>月<%=day(now())%>日&nbsp;&nbsp;<%=weekdayname(weekday(now()))%>&nbsp;&nbsp;&nbsp;
		  </td>
        </tr>
        <tr>
          <td height="5"></td>
        </tr>
      </table>
	  </td>
    </tr>
</table>

<%'Session("step")=1%>
<table width="100%" border="1" cellspacing="2" cellpadding="3">
  <tr>
    <td align="left" background="images/menu_bg2.gif">
	&nbsp;<img src="images/list1.gif" align="absmiddle" />&nbsp;
	<strong>常用管理：</strong>
    <%if Session("step")=1 then%>
	<a href="MenuManage.asp">栏目管理</a>&nbsp;&nbsp;|&nbsp;
    <%end if%>
	<a href="ArticleManage.asp?xs=all&order=id desc">文章管理</a>&nbsp;&nbsp;|&nbsp;
	<a href="PictureManage.asp?xs=all&order=id desc">图片管理</a>&nbsp;&nbsp;|&nbsp;
	<a href="FormsManage.asp?xs=all&order=id desc">表单管理</a>&nbsp;&nbsp;|&nbsp;
	<a href="usersManage.asp?xs=all&order=id desc">会员管理</a>&nbsp;&nbsp;|&nbsp;
	<a href="xiaofeiManage.asp?xs=all&order=id desc">消费记录</a>&nbsp;&nbsp;|&nbsp;
	<a href="JifenManage.asp?xs=all&order=id desc">积分兑换</a>&nbsp;&nbsp;|&nbsp;
    <%if Session("step")=1 then%>
	<a href="MusicManage.asp?xs=all&order=id desc">背景音乐</a>&nbsp;&nbsp;|&nbsp;
	<a href="LinkManage.asp?xs=all&order=id desc">友情链接</a>&nbsp;&nbsp;|&nbsp;
	<a href="CserverManage.asp">在线客服</a>&nbsp;&nbsp;|&nbsp;
	<a href="SystemManage.asp">其他设置</a>
    <%end if%>
	<%if Session("step")=1 then%><br><%end if%>
	
    <%if Session("step")=1 then%>
    &nbsp;<img src="images/list4.gif" align="absmiddle" />&nbsp;
	<strong>系统管理：</strong>
	<a href="AdminManage.asp?xs=all">帐号管理</a>&nbsp;&nbsp;|&nbsp;
	<a href="BookManage.asp?xs=all&amp;order=id desc">日志管理</a>&nbsp;&nbsp;|&nbsp;
	<a href="DataManage.asp">数据管理</a>&nbsp;&nbsp;|&nbsp;
	<a href="SpaceManage.asp">空间占用</a>&nbsp;&nbsp;|&nbsp;
	<a href="PassWordchange.asp">修改密码</a>&nbsp;&nbsp;|&nbsp;
	<%if Session("admin")="sosovipp" then%>
		<a href="ZiduanManage.asp?xs=all&order=id desc">字段管理</a>&nbsp;&nbsp;|&nbsp;
		<a href="Adm_acs.asp" target="_blank">高级模式</a>&nbsp;&nbsp;|&nbsp;
	<%end if%>
    <%end if%>
	<a href="loginchk.asp?tj=out">退出</a>
	</td>
  </tr>
</table>