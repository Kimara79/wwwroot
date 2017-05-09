<%
'获取IP所在地
Function ipadr(ip)
	body=getHTTPPage("http://www.ip138.com/ips.asp?action=2&ip="&ip,"gb2312")
	temp=split(body,"本站主数据：")
  	temp2=split(temp(1),"</li>")
  	ipadr=temp2(0)
End Function

Function ipadr2(ip)
	body=getHTTPPage("http://www.123cha.com/ip/?q="&ip,"utf-8")
	temp=split(body,"本站主数据:")
  	temp2=split(temp(1),chr(13)&chr(10))
  	ipadr2=replace(temp2(1),"</li>","")
End Function

Function ipadr3(ip)
	body=getHTTPPage("http://www.ip38.com/index.php?ip="&ip,"gb2312")
	temp=split(body,"查询结果：")
  	temp2=split(temp(1),"</font>")
  	ipadr3=temp2(0)
End Function

Function ipadr4(ip)
	body=getHTTPPage("http://www.1234i.com/?ip="&ip,"gb2312")
	temp=split(body,"face="&chr(34)&"黑体"&chr(34)&">")
  	temp2=split(temp(1),"<br>")
  	ipadr4=temp2(0)
End Function
%>