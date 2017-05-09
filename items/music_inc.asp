<%
''网站域名
'sql="select info from system where id=11"
'rsstr=trim(conn.execute(sql).getstring)
'rsstr=replace(rsstr,chr(13),"")
'site_url=rsstr

''写入文件
filepath="../Player/Haopylist.php"
site_url=Request.ServerVariables("HTTP_HOST")
set fso = Server.CreateObject("Scripting.FileSystemObject")
set fout = fso.CreateTextFile(server.mappath(filepath))
sql="select * from music where shezhi like '%播放%' order by no1"
rs.open sql,conn,1,1
fout.WriteLine "<data>"
a=0
Do while a<16
	i=0
	rs.movefirst
	Do while not rs.eof
		i=i+1
		a=a+1
		if rs("file1")<>"" then
			url=site_url&"/items/music/"&rs("file1")
		else
			url=rs("url")
		end if
		url=lcase(url)
		url=replace(url,"http://","")
		url=replace(url,".mp3","")
		fout.WriteLine "<song>"
		fout.WriteLine "<title>"&rs("name1")&"</title>"
		fout.WriteLine "<name>"&url&"</name>"
		fout.WriteLine "</song>"
		rs.movenext
	loop
loop
fout.WriteLine "</data>"
rs.close
fout.close
set fout=nothing
set fso=nothing
%>