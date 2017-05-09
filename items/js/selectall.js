<SCRIPT>
function selectall()   
{  
   if (document.form.selectButton.checked=="1")  
   {  
	 for (var i=0;i<document.form.dst.length;i++) { 
	 document.form.dst[i].checked="checked";
     }    
   }  
   else  
   {  
	 for (var i=0;i<document.form.dst.length;i++) { 
	 document.form.dst[i].checked="";
	 } 
   }  
}

function selectall1()   
{  
   if (document.form1.selectButton.checked=="1")  
   {  
	 for (var i=0;i<document.form1.dst.length;i++) { 
	 document.form1.dst[i].checked="checked";
     }    
   }  
   else  
   {  
	 for (var i=0;i<document.form1.dst.length;i++) { 
	 document.form1.dst[i].checked="";
	 } 
   }  
}

function selectall2()   
{  
   if (document.form2.selectButton.checked=="1")  
   {  
	 for (var i=0;i<document.form2.dst.length;i++) { 
	 document.form2.dst[i].checked="checked";
     }    
   }  
   else  
   {  
	 for (var i=0;i<document.form2.dst.length;i++) { 
	 document.form2.dst[i].checked="";
	 } 
   }  
}
</SCRIPT>