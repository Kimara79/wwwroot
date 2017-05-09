<%
'SQL通用防注入程序
'--------定义部份------------------
Dim Fy_Post,Fy_Get,Fy_In,Fy_Inf,Fy_Xh,Fy_db,Fy_dbstr,alerts
'自定义需要过滤的字串,用 "|" 分隔
Fy_In = "and|exec|insert|select|delete|update|count|chr|mid|master|truncate|char|declare|get|post|script|cokkie"
Fy_Inf = split(Fy_In,"|")
'----------------------------------

Sub e_alert(a)
 alerts = "<"&"Script Language=JavaScript"&">"
 alerts = alerts & "alert('非法关键字："&a&"。请不要试图攻击本站,我们已经记录下你的信息！');window.opener=null; window.close();"
 alerts = alerts & "<"&"/Script"&">" 
Response.Write alerts  
Response.End
end Sub
%>

<%
'--------POST部份------------------
If Request.Form<>"" Then
  For Each Fy_Post In Request.Form
    For Fy_Xh=0 To Ubound(Fy_Inf)
      If Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh))<>0 Then
  		call e_alert(Fy_Inf(Fy_Xh))
      End If
    Next
  Next
End If

'--------GET部份-------------------
If Request.QueryString<>"" Then
  For Each Fy_Get In Request.QueryString
    For Fy_Xh=0 To Ubound(Fy_Inf)
      If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh))<>0 Then
		  call e_alert(Fy_Inf(Fy_Xh))
      End If
    Next
  Next
End If

'--------cookies部份-------------------
If Request.Cookies<>"" Then
  For Each Fy_Get In Request.Cookies
    For Fy_Xh=0 To Ubound(Fy_Inf)
      If Instr(LCase(Request.Cookies(Fy_Get)),Fy_Inf(Fy_Xh))<>0 Then
		  call e_alert(Fy_Inf(Fy_Xh))
      End If
    Next
  Next
End If
%>