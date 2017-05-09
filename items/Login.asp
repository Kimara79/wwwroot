<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="Shortcut Icon" href="../favicon.ico">
<title>后台管理</title>

<style type="text/css">
<!--
body {
	margin-top: 0px;
	margin-bottom: 0px;
	margin-left: 0px;
	margin-right: 0px;
	background-color: #c5c5c5;
}
body,td,th {
	font-size: 12px;
	color: #268fc0;
	font-weight: bold;
}
.xh {
	border-top-style: none;
	border-right-style: none;
	border-bottom-style: solid;
	border-left-style: none;
	border-bottom-width: 1px;
	border-bottom-color: #268fc0;
	background:transparent;
	
	font-size: 12px;
	color: #268fc0;
	font-weight: bold;
}
-->
</style>
</head>

<!--#include file="inc/cokk.asp"-->
<!--#include file="js/maskEdit.js"-->

<%
'  response.Write "<center><br><br><br><br><font size='+3'>系统升级中，请稍后访问……</font></center>"
'  response.end
  session("admin")=""
%>

<body>
<div align="center">
  <br>
    <br>
    <br>
  <br>
  <form action="loginchk.asp?tj=in" method="post" name="form" id="form">
  <table width="600" height="400" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td align="left" valign="bottom" background="images/lg_bg.jpg"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="7%" height="60" align="right" valign="bottom">&nbsp;</td>
          <td valign="bottom"><label>用户名:<input name="admin" type="text" class="xh" id="admin" onKeyPress="maskEdit(/^[\w]*$/)" size="10" /></label></td>
          <td valign="bottom"><label>密码:<input name="password" type="password" class="xh" id="password" onKeyPress="maskEdit(/^[\w]*$/)" onkeydown="if (event.keyCode==13) {document.form.submit()}" size="10" /></label></td>
          <td valign="bottom"><label>验证码:<input name="s" type="text" class="xh" id="s"  style="font-size: 12px" size="4" onkeyup="if(event.keyCode !=37 && event.keyCode != 39) value=value.replace(/\D/g,'');" onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/\D/g,''))" maxlength="4" onkeydown="if (event.keyCode==13) {document.form.submit()}" /></label>
<%
randomize timer
s=Int((8999)*Rnd +1000)
%>
<b><%=s%></b>
<input name="s2" type="hidden" id="s2" value="<%=s%>" />
           </td>
          <td align="right" valign="bottom"><img src="images/lg_bt1.gif" width="99" height="25" onClick="javascript:document.form.submit()" style="CURSOR: hand;"></td>
          <td width="7%" align="right" valign="bottom">&nbsp;</td>
        </tr>
        <tr>
          <td height="65" colspan="6" align="right" valign="bottom">&nbsp;</td>
        </tr>
      </table></td>
    </tr>
  </table>
</form>
</div>
<script language="javascript">
document.form.admin.focus(); 
</script>
</body>
</html>