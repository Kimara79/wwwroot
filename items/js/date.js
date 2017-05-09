<script language="JavaScript">
function date(_obj)
{
	date1 = new Date(); 
	year = date1.getYear(); 
	month = date1.getMonth();
	day = date1.getDay();
	if (_obj.value)
	{
	}
	else
	{
		_obj.value=year+'年'+month+'月'+day+'日';
	}
}
function date_ym(a,b)
{
	date1 = new Date(); 
	year = date1.getYear(); 
	month = date1.getMonth();
	if (a.value)
	{
	}
	else
	{
		a.value=year;
	}
	if (b.value)
	{
	}
	else
	{
		b.value=month;
	}
}
function date_cg(a,b,c)
{
	if (c.value.length>14)
	{
	}
	else
	{
		c.value=a.value+'年'+b.value+'月';
	}
}
</script>