<script>
if(window.name != "Bencalie"){
//如果页面的 name 属性不是指定的名称就刷新它
location.reload();
window.name = "Bencalie";
}
else{
//如果页面的 name 属性是指定的名称就什么都不做，避免不断的刷新
window.name = "";
}
</script>