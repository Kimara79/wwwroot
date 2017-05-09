<%
Function Sort(ary)
Dim KeepChecking,I,FirstValue,SecondValue
 KeepChecking = TRUE 
Do Until KeepChecking = FALSE 
 KeepChecking = FALSE 
 For I = 0 to UBound(ary) 
  If I = UBound(ary) Then Exit For 
   If ary(I) > ary(I+1) Then 
    FirstValue = ary(I) 
    SecondValue = ary(I+1) 
    ary(I) = SecondValue 
    ary(I+1) = FirstValue 
    KeepChecking = TRUE 
   End If 
 Next 
Loop 
 Sort = ary 
End Function 
%>