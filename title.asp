<%
if request("menu1")<>"" then
	tb="menu1"
	id=request("menu1")
	sql="select * from menu1 where id="&id
	rs.open sql,conn,1,1
	menu1_id=rs("id")
	menu1_name1=rs("name1")
	rs.close
	title1=menu1_name1
end if
if request("menu2")<>"" then
	tb="menu2"
	id=request("menu2")
	sql="select * from menu2 where id="&id
	rs.open sql,conn,1,1
	menu1_id=rs("menu1")
	menu2_id=rs("id")
	menu2_name1=rs("name1")
	title1=menu2_name1
	rs.close
	rsstr=""
	sql="select name1 from menu1 where id="&menu1_id
	rsstr=trim(conn.execute(sql).getstring)
	rsstr=replace(rsstr,chr(13),"")
	menu1_name1=rsstr
	title1=title1&" - "&rsstr
end if
if request("id")<>"" then
	id=request("id")
	sql="select name1,menu1,menu2 from "&tb&" where id="&id
	rs.open sql,conn,1,1
	title1=rs("name1")
	menu1_id=rs("menu1")
	menu2_id=rs("menu2")
	rs.close
	if int(menu2_id)>0 then
		rsstr=""
		sql="select name1 from menu2 where id="&menu2_id
		rsstr=trim(conn.execute(sql).getstring)
		rsstr=replace(rsstr,chr(13),"")
		menu2_name1=rsstr
		title1=title1&" - "&rsstr
	end if
	rsstr=""
	sql="select name1 from menu1 where id="&menu1_id
	rsstr=trim(conn.execute(sql).getstring)
	rsstr=replace(rsstr,chr(13),"")
	menu1_name1=rsstr
	title1=title1&" - "&rsstr
end if
if request("menu1")<>"" then
	sqladd=" menu1="&menu1_id&" "
end if
if request("menu2")<>"" then
	sqladd=" menu2="&menu2_id&" "
end if
if (request("menu1")="" and request("menu2")="") then
	if menu2_id<>"" then
		sqladd=" menu2="&menu2_id&" "
	else
		sqladd=" menu1="&menu1_id&" "
	end if
end if
%>