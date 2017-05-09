<% 
Response.Write "<b>cookies变量: </b>" 
Response.Write "<br>" 
for each item in Request.Cookies
	Response.Write  "<FONT COLOR=red>"&item&":</FONT>"&Request.Cookies(item) &"<br>"
next 

Response.Write "<br>" 
Response.Write "<br>" 

Response.write "<b>session变量：</b>"
Response.Write "<br>" 
for each item in session.contents 
	Response.Write"<FONT COLOR=red>"& item &":</FONT>"&session(item) &"<br>"
next 

Response.Write "<br>" 
Response.Write "<br>" 

Response.write "<b>application变量：</b>"
Response.Write "<br>" 
for each item in application.contents 
	Response.Write "<FONT COLOR=red>"& item & ": </FONT>"&application(item) &"<br>"
next

Response.Write "<br>" 
Response.Write "<br>" 

Response.Write "<b>服务器变量：</b>"
Response.Write "<br>" 
for each item in request.servervariables
	Response.Write"<FONT COLOR=red>"&item&":</FONT>"&Request.ServerVariables(item) &"<br>"
next 

Response.Write "<br>"
Response.Write "<br>"

Response.Write "<b>脚本引擎: </b>"
Response.Write "<br>"
Response.Write scriptengine()
Response.Write "<br>"
Response.Write scriptenginemajorversion()
Response.Write "<br>" 
Response.Write scriptengineminorversion()
Response.Write "<br>" 
Response.Write scriptenginebuildversion()
%>