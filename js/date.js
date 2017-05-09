<SCRIPT>
var isnMonths=new initArray("1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月");
var isnDays=new initArray("星期日","星期一","星期二","星期三","星期四","星期五","星期六","星期日");
//var isnDays2=new initArray("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday");
var isnDays2=new initArray("Sun","Mon","Tue","Wed","Thu","Fri","Sat","Sun");
today=new Date();
hrs=today.getHours();
min=today.getMinutes();
sec=today.getSeconds();
clckh=""+((hrs>12)?hrs:hrs);
clckm=((min<10)?"0":"")+min;
clcks=((sec<10)?"0":"")+sec;
clck=(hrs>=12)?"下午":"上午";
var stnr="";
var ns="0123456789";
var a="";

function initArray(){
	for(i=0;i<initArray.arguments.length;i++)
	this[i]=initArray.arguments[i];
}

function getFullYear(d){
	yr=d.getYear();if(yr<1000)
	yr+=1900;
	return yr;
}
</SCRIPT>