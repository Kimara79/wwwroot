<script language="JavaScript">
function date(_obj)
{
	if (_obj.value)
	{
	}
	else
	{
		_obj.value=<%=year(now())%>+'年'+<%=month(now())%>+'月'+<%=day(now())%>+'日';
	}
}
function date_ym(a,b)
{
	if (a.value)
	{
	}
	else
	{
		a.value=<%=year(now())%>;
	}
	if (b.value)
	{
	}
	else
	{
		b.value=<%=month(now())%>;
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