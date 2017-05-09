<script language="JavaScript">
//duplex
//paper_size
//paper_source
//printer
var hkey_root,hkey_path,hkey_key
hkey_root="HKEY_CURRENT_USER"
hkey_path="\\Software\\Microsoft\\Internet Explorer\\PageSetup\\"
//A4纸张
function pagesetup_papersize2(){
try{
var RegWsh = new ActiveXObject("WScript.Shell")
hkey_key="papersize" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"A5")
}catch(e){}
}
//页眉页脚为默认值
function pagesetup_hf(){
try{
var RegWsh = new ActiveXObject("WScript.Shell")
hkey_key="header" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"&w&b页码，&p/&P")
hkey_key="footer"
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"&u&b&d")
}catch(e){}
}
//页眉页脚为空
function pagesetup_hf2(){
try{
var RegWsh = new ActiveXObject("WScript.Shell")
hkey_key="header" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"")
hkey_key="footer"
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"")
}catch(e){}
}
//纵向打印
function pagesetup_orientation(){
try{
var RegWsh = new ActiveXObject("WScript.Shell")
hkey_key="orientation" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"landscape")
}catch(e){}
}
//横向打印
function pagesetup_orientation2(){
try{
var RegWsh = new ActiveXObject("WScript.Shell")
hkey_key="orientation" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"portrait")
}catch(e){}
}
//页边距为19.5
function pagesetup_margin(){
try{
var RegWsh = new ActiveXObject("WScript.Shell")
hkey_key="margin_bottom" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"0.76772")
hkey_key="margin_left" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"0.76772")
hkey_key="margin_right" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"0.76772")
hkey_key="margin_top" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"0.76772")
}catch(e){}
}
//页边距为10
function pagesetup_margin2(){
try{
var RegWsh = new ActiveXObject("WScript.Shell")
hkey_key="margin_bottom" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"0.39370")
hkey_key="margin_left" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"0.39370")
hkey_key="margin_right" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"0.39370")
hkey_key="margin_top" 
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"0.39370")
}catch(e){}
}
</script>