<%
Function IsObjInstalled(strClassString)
 On Error Resume Next
 IsObjInstalled = False
 Err = 0
 Dim xTestObj
 Set xTestObj = Server.CreateObject(strClassString)
 If 0 = Err Then IsObjInstalled = True
 Set xTestObj = Nothing
 Err = 0
End Function
%>

<% if IsObjInstalled("Scripting.FileSystemObject") = False Then %>

不支持FSO

<% Else %>

支持FSO

<% End If %>