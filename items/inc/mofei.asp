<%
Function ReportFileStatus(FileName) 
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Dim msg 
msg = -1 
If (objFSO.FileExists(FileName)) Then 
msg = 1 
Else 
msg = -1 
End If 
ReportFileStatus = msg 
End Function 
%>

<%
Function mofei(sfile)
sfilere=""
Set fso = Server.CreateObject("Scripting.FileSystemObject")
if sfile="" or isnull(sfile) then sfile=Request.ServerVariables("SCRIPT_NAME")
if not fso.FileExists(server.Mappath(sfile)) then sfile=Request.ServerVariables("SCRIPT_NAME")
sfile=server.Mappath(sfile)
Set fsofile=fso.getfile(sfile)
sfilere=sfilere&fsofile.DateLastModified '文件时间
sfilere=sfilere&chr(9)&"&nbsp;&nbsp;"&"<font color=green>"&fsofile.type&"</font>" '类型
sfilere=sfilere&chr(9)&"&nbsp;&nbsp;"&round(fsofile.size*1000/1024/1000,2)&" KB" '大小
set fso=nothing
mofei=sfilere
End Function
%>

<%
Function mofei1(sfile)
sfilere=""
Set fso = Server.CreateObject("Scripting.FileSystemObject")
if sfile="" or isnull(sfile) then sfile=Request.ServerVariables("SCRIPT_NAME")
if not fso.FileExists(server.Mappath(sfile)) then sfile=Request.ServerVariables("SCRIPT_NAME")
sfile=server.Mappath(sfile)
Set fsofile=fso.getfile(sfile)
sfilere=sfilere&fsofile.DateLastModified '文件时间
set fso=nothing
mofei1=sfilere
End Function
%>

<%
Function mofei2(sfile)
sfilere=""
Set fso = Server.CreateObject("Scripting.FileSystemObject")
if sfile="" or isnull(sfile) then sfile=Request.ServerVariables("SCRIPT_NAME")
if not fso.FileExists(server.Mappath(sfile)) then sfile=Request.ServerVariables("SCRIPT_NAME")
sfile=server.Mappath(sfile)
Set fsofile=fso.getfile(sfile)
sfilere=sfilere&"<font color=green>"&fsofile.name&"</font>" '类型
set fso=nothing
mofei2=sfilere
End Function
%>
<%
Function mofei3(sfile)
sfilere=""
Set fso = Server.CreateObject("Scripting.FileSystemObject")
if sfile="" or isnull(sfile) then sfile=Request.ServerVariables("SCRIPT_NAME")
if not fso.FileExists(server.Mappath(sfile)) then sfile=Request.ServerVariables("SCRIPT_NAME")
sfile=server.Mappath(sfile)
Set fsofile=fso.getfile(sfile)
sfilere=sfilere&round(fsofile.size*1000/1024/1000,2)&" KB" '大小
set fso=nothing
mofei3=sfilere
End Function
%>