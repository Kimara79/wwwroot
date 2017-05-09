<%
'禁止半角、全角空格
fbjs_space="this.value=this.value.replace(/[\s\'\042　]/g,'');" 

'只能输入数字
fbjs_rnum1="this.value=this.value.replace(/[^\d]/g,'');"

'只能输入数字及小数点
fbjs_rnum2="this.value=this.value.replace(/[^\d\.]/g,'');"

'只能输入数字及横杆
fbjs_rnum3="this.value=this.value.replace(/[^\d\-]/g,'');"

'只能输入数字及冒号
fbjs_rnum4="this.value=this.value.replace(/[^\d\:]/g,'');"

'只能输入字母
fbjs_rletter1="this.value=this.value.replace(/[^A-Za-z]/g,'');"

'只能输入字母、数字
fbjs_rletter2="this.value=this.value.replace(/[^\w\d]/g,'');"

'只能输入字母、数字、小数点、下划线
fbjs_rletter3="this.value=this.value.replace(/[^\w\.\040]/g,'');"

'自动转换为大写
fbjs_else1="this.value=this.value.toUpperCase();"
%>