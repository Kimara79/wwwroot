<form id="user_login" name="user_login" method="post" action="user_login_tj.asp?tj=in">
<table border="0" cellspacing="0" cellpadding="<%=cellpadding%>">
  <tr>
	<td align="center" class="<%=fontclass%>">手机：</td>
	<td align="left"><input type="text" name="tel" id="tel" class="<%=input%>" maxlength="20" onkeyup="<%=fbjs_rnum1%>" /></td>
  </tr>
  <tr>
	<td align="center" class="<%=fontclass%>">密码：</td>
	<td align="left"><input type="password" name="password" id="password" class="<%=input%>" maxlength="20" onkeydown="if (event.keyCode==13) {document.user_login.submit()}" /></td>
  </tr>
  <tr>
	<td colspan="2" align="center">
	  <input type="submit" name="Submit" value="登录" class="<%=bt%>" />&nbsp;&nbsp;
	  <input type="reset" name="Submit" value="重置" class="<%=bt%>" />
	  </td>
	</tr>
</table>
</form>
