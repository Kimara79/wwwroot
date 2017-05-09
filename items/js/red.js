<script language="JavaScript">
function red(_obj)
{
	var _a=_obj.name;
	var _b=document.getElementsByName(_a);
	for (var i=0;i<_b.length;i++) { 
		if (_b[i].checked){
				_b[i].parentElement.style.color='#ff0000';
			}
			else{
				_b[i].parentElement.style.color='';
			}
	}
}

function red2(_obj){
	var _b=document.getElementsByName(_obj);
	for (var i=0;i<_b.length;i++) { 
		if (_b[i].checked){
				_b[i].parentElement.style.background='#ff0000';
			}
			else{
				_b[i].parentElement.style.background='';
			}
	}
}

function red3(_obj){
	var _b=document.getElementsByName(_obj);
	for (var i=0;i<_b.length;i++) { 
		if (_b[i].checked){
				_b[i].parentElement.style.background='#ff0000';
			}
			else{
				_b[i].parentElement.style.background='';
			}
	}
}
</script>