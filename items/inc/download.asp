<%
 if Session("admin")="" then
 response.redirect ("login.asp")
 end if
%>
<%
if request("filename")<>"" then

'空判断
Dim Stream
Dim Contents
Dim FileName
Dim FileExt
Const adTypeBinary = 1
FileName = Request.QueryString("FileName")
FileName1 = split(FileName,"/")
FileName2 = FileName1(1)
if FileName = "" Then
Response.Write "无效文件名."
Response.End
End if

' 下面是不希望下载的文件
FileExt = Mid(FileName, InStrRev(FileName, ".") + 1)
Select Case UCase(FileExt)
Case "ASP", "ASA", "ASPX", "ASAX", "MDB"
Response.Write "受保护文件,不能下载."
Response.End
End Select

' 下载这个文件
Response.Clear
Response.ContentType = "application/octet-stream"
Response.AddHeader "content-disposition", "attachment; filename=" & FileName2
Set Stream = server.CreateObject("ADODB.Stream")
Stream.Type = adTypeBinary
Stream.Open
Stream.LoadFromFile Server.MapPath(FileName)
While Not Stream.EOS
Response.BinaryWrite Stream.Read(1024 * 64)
Wend

'关闭
Stream.Close
Set Stream = Nothing
Response.Flush

end if
%>