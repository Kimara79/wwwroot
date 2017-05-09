<%
'数字加小写字母
function randomx(maxLen)
	Dim strNewPass
	Dim whatsNext, upper, lower, intCounter
	Randomize
	For intCounter = 1 To maxLen
	whatsNext = Int((1 - 0 + 1) * Rnd + 0)
	If whatsNext = 0 Then 
		lower = 97
		upper = 122
	Else
		lower = 48
		upper = 57
	End If
	strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	Next
	randomx = strNewPass
end function

'数字
function randomx1(maxLen)
	Dim strNewPass
	Dim whatsNext, upper, lower, intCounter
	Randomize
	For intCounter = 1 To maxLen
	lower = 48
	upper = 57
	strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	Next
	randomx1 = strNewPass
end function

'小写字母
function randomx2(maxLen)
	Dim strNewPass
	Dim whatsNext, upper, lower, intCounter
	Randomize
	For intCounter = 1 To maxLen
	lower = 97
	upper = 122
	strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	Next
	randomx2 = strNewPass
end function
%>