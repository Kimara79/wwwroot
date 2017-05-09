<script language="JavaScript">
var idTmr = ""; 

function copy(tabid)
{ 
var oControlRange = document.body.createControlRange();
oControlRange.add(tabid,0);
oControlRange.select();
document.execCommand("Copy");
} 

function toExcel(tabid){
copy(tabid);
try
{
var xls = new ActiveXObject( "Excel.Application" );
xls.Visible = true;
}
catch(e) 
{ 
alert( "Excel没有安装或浏览器设置不正确，请启用所有Active控件和插件。\n工具->Internet选项->安全->Internet->自定义级别->ActiveX 控件和插件->启用“对没有标记为可安全的 ActiveX 控件进行初始化和脚本运行”");
return false; 
} 
xls.visible = true; 
var xlBook = xls.Workbooks.Add;
var xlsheet = xlBook.Worksheets(1);
xlBook.Worksheets(1).Activate;
for(var i=0;i<tabid.rows(0).cells.length;i++){
xlsheet.Columns(i+1).ColumnWidth=15; 
}
xlsheet.Paste;
xls=null;
}
</script>