<!----inc------>
<link type="text/css" rel="stylesheet" href="Gallery/galleryview.css" />
<script type="text/javascript" src="Gallery/jquery-1.8.0.min.js"></script>
<script type="text/javascript" src="Gallery/jquery-ui-1.8.23.custom.min.js"></script>
<script type="text/javascript" src="Gallery/jquery.galleryview-2.0-pack.js"></script>
<script type="text/javascript" src="Gallery/jquery.timers-1.1.2.js"></script>
<script type="text/javascript" src="Gallery/jquery.easing.1.2.js"></script>

<!----set------>
<script>
	$(document).ready(function()
	{$('#gallery').galleryView({
	panel_width:524,
	panel_height:375,
	frame_width:50,
	frame_height:50,
	transition_speed:350,
	easing:'easeInOutQuad',
	transition_interval:5000
	})
	});
</script>

<!----htm------>
<center>
<%
sql="select * from picture where menu1=26 order by shezhi desc,no1"
rs.open sql,conn,1,1
text1=rs("name1")
%>
<div align="center"><img src="images/topline1.png" /></div>
<ul id="gallery">
	<%Do while not rs.eof%>
	<li>
	<div class="panel-content">
		<a href="<%=rs("url")%>"><img src="items/pic/picture/<%=rs("pic")%>" border="0" /></a>
		<span class="panel-overlay">
			<span class="panel-overlay-h1" style="color:<%=rs("color")%>;"><%=rs("name1")%></span><br />
			<span class="panel-overlay-sub"><%=rs("jianjie")%></span>&nbsp;&nbsp;<a href="<%=rs("url")%>">查看详情 &raquo;</a>
		</span>
	</div>
	<img src="items/pic/s_picture/<%=rs("pic")%>" />
	</li> 
	<%
	rs.movenext
	loop
	%>
</ul>
<%rs.close%>
</center>
