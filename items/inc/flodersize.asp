<%
Function flodersize(drvpath)
	dim fso,d,size,showsize
	set fso=server.createobject("scripting.filesystemobject")
	set d=fso.getfolder(server.mappath(drvpath))                 
	size=d.size
	showsize=size&"Byte"
	if size>1024 then
	   size=(size\1024)
	   showsize=size&"K"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2)&"M"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2)&"G"
	end if
	flodersize=showsize
end Function                 
%>