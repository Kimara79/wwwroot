<script type="text/javascript">　　
function filecheck(_obj)　　
{　　
   var type1="<%=typeall%>";
   var hz=_obj.value.split(".");
   var hz1=hz[hz.length-1];
   if (type1.indexOf(hz1)<0){
	_obj.focus(); 
	document.execCommand("selectall"); 
	document.execCommand("Delete"); 
	alert("允许上传的文件格式：<%=typeall%>");
	location.reload();
	return false;
   }
	
   var fso = new ActiveXObject("Scripting.FileSystemObject"); 
   var f,s,file; 
   if ("object" != typeof(fso)) return; 
   file = _obj.value; 
   f = fso.GetFile(file); 
   if (f.size><%=fzmax%>)　{
	_obj.focus(); 
	document.execCommand("selectall"); 
	document.execCommand("Delete"); 
	alert("允许上传的大小：<%=left(""&fzmax,1)%>M");
	location.reload();
	return false;
   }
}　　
</script>