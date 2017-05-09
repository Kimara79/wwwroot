<%
dim sql_leach,sql_leach_0,Sql_DATA,SQL_Get,Sql_Post 
sql_leach = "',and,exec,insert,select,delete,update,count,*,%,chr,mid,master,truncate,char,declare"
sql_leach_0 = split(sql_leach,",")

If Request.QueryString<>"" Then
For Each SQL_Get In Request.QueryString
For SQL_Data=0 To Ubound(sql_leach_0)
if instr(Request.QueryString(SQL_Get),sql_leach_0(Sql_DATA))>0 Then
Response.Write("<script language=javascript>alert('存在非法字符！');history.back()</script>")
Response.end
end if
next
Next
End If


If Request.Form<>"" Then
For Each Sql_Post In Request.Form
For SQL_Data=0 To Ubound(sql_leach_0)
if instr(Request.Form(Sql_Post),sql_leach_0(Sql_DATA))>0 Then
Response.Write("<script language=javascript>alert('存在非法字符！');history.back()</script>")
Response.end
end if
next
next
end if
%>