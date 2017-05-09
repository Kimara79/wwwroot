<%
'html转文本
Function FilterHtml(str)
	Dim regEx
	Set regEx=New RegExp
	regEx.IgnoreCase=True
	regEx.Global=True
	regEx.MultiLine=True
	regEx.Pattern="<.+?>"
	temp=regEx.Replace(str,"")
	temp=Replace(temp,"&nbsp;"," ")
	FilterHtml=temp
	Set regEx=Nothing
End Function

'html转文本-去掉换行及空格
Function FilterHtml2(str)
	Dim regEx
	Set regEx=New RegExp
	regEx.IgnoreCase=True
	regEx.Global=True
	regEx.MultiLine=True
	regEx.Pattern="<.+?>"
	temp=regEx.Replace(str,"")
	temp=Replace(temp,"&nbsp;","")
	temp=Replace(temp,chr(10),"")
	temp=Replace(temp,chr(13),"")
	temp=Replace(temp,chr(9),"")
	temp=Replace(temp," ","")
	temp=Replace(temp,"　","")
	FilterHtml2=temp
	Set regEx=Nothing
End Function

'显示特殊符号编码
Function charcode(TheString)
 temp1=""&TheString
 for n=0 to 32
  temp1=replace(temp1,chr(n),"chr("&n&")")
 next
 charcode=temp1
End Function

'显示特殊符号编码2
Function charcode2(TheString)
 temp1=""&TheString
 temp1=replace(temp1,chr(9),"chr(9)")
 temp1=replace(temp1,chr(10),"chr(10)")
 temp1=replace(temp1,chr(13),"chr(13)")
 charcode2=temp1
End Function
%>