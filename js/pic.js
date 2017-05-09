<script language="javascript"> 

//定制最大尺寸
function changeImg(mypic,w,h){ 
     var xw=w; 
     var xl=h; 
     var width = mypic.width; 
     var height = mypic.height; 
     if (width > xw ) mypic.width = xw; 
     if (height > xl ) mypic.height = xl; 
} 

//按比例缩放2
function changeImg1(mypic,w,h){ 
     var xw=w; 
     var xl=h; 
     var width = mypic.width; 
     var height = mypic.height; 
     var bili = width/height;             
     var A=xw/width; 
     var B=xl/height; 
     if(A<1||B<1) 
     { 
             if(A<B) 
             { 
                 mypic.width=xw; 
                 mypic.height=xw/bili; 
             } 
             if(A>B) 
             { 
                 mypic.width=xl*bili; 
                 mypic.height=xl; 
             } 
     } 
} 

//按比例缩放1
function changeImg2(obj,picW,picH){
 if(obj.width>picW || obj.height>picH ){
  if(obj.width/obj.height>picW/picH  )
   obj.width=picW;
  else 
   obj.height=picH;
 }
}

</script>