<%
 if Session("admin")="" then
 response.redirect ("login.asp")
 end if
%>
<%
if request("filename")<>"" then

'���ж�
Dim Stream
Dim Contents
Dim FileName
Dim FileExt
Const adTypeBinary = 1
FileName = Request.QueryString("FileName")
FileName1 = split(FileName,"/")
FileName2 = FileName1(1)
if FileName = "" Then
Response.Write "��Ч�ļ���."
Response.End
End if

' �����ǲ�ϣ�����ص��ļ�
FileExt = Mid(FileName, InStrRev(FileName, ".") + 1)
Select Case UCase(FileExt)
Case "ASP", "ASA", "ASPX", "ASAX", "MDB"
Response.Write "�ܱ����ļ�,��������."
Response.End
End Select

' ��������ļ�
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

'�ر�
Stream.Close
Set Stream = Nothing
Response.Flush

end if
%>