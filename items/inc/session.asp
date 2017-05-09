<%
Dim strName, iLoop
For Each strName in Session.Contents
Response.Write strName & " - " & Session.Contents(strName)& "<BR>"
Next
%>
<br>Session.Timeout=<%=Session.Timeout%>