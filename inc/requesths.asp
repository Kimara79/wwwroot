<%
function requesttrp(cs)
	requesttrp=""
	if request.QueryString<>"" then
		for each item in request.QueryString
			if request.QueryString(item)<>"" then
				if lcase(item)<>lcase(cs) then
					requesttrp=requesttrp&item&"="&request.QueryString(item)&"&"
				end if
			end if
		next
	end if
end function

function requestformtoqs()
	requestformtoqs=""
	for each item in request.form
		if request.form(item)<>"" then
			if (lcase(item)<>"submit" and lcase(item)<>"reset") then
				requestformtoqs=requestformtoqs&item&"="&request.form(item)&"&"
			end if
		end if
	next
	if requestformtoqs<>"" then
		requestformtoqs=left(requestformtoqs,len(requestformtoqs)-1)
	end if
end function
%>