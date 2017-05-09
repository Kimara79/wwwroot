<%
Function getbody(url,code)
	set http = Server.CreateObject("Microsoft.XMLHTTP")
	http.Open "GET",url,false
	http.setRequestHeader "Content-type:", "text/xml;charset="&code
	http.Send
	getbody=http.ResponseText
	set	http = nothing
End Function

Function BytesToBstr(body,Cset) 
	dim objstream 
	set objstream=Server.CreateObject("adodb.stream") 
	objstream.Type=1 
	objstream.Mode=3 
	objstream.Open 
	objstream.Write body 
	objstream.Position=0 
	objstream.Type=2 
	objstream.Charset=Cset 
	BytesToBstr=objstream.ReadText   
	objstream.Close 
	set objstream=nothing 
End Function

function getHTTPPage(url,code)
	dim Http
	set Http=server.createobject("MsXml2.XmlHttp")
	Http.open "GET",url,false
	Http.send()
	if Http.readystate<>4 then   
	exit function 
	end if 
	getHTTPPage=BytesToBstr(Http.responseBody,code) 
	set http=nothing 
	if err.number<>0 then   
	err.Clear 
	end if 
end function
%>