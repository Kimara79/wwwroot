function focuss(pics1,links1,texts1){
var pics=pics1
var links=links1
var texts=texts1
var show_text=0; //�Ƿ���ʾ���ֱ�ǩ 1��ʾ 0����ʾ
var interval_time=5 //ͼƬͣ��ʱ��,��λΪ��,Ϊ0��ֹͣ�Զ��л�
var focus_width=693 //���
var focus_height=380 //�߶�
var text_height=0 //����߶�
var text_align= 'center' //�������ֶ��뷽ʽ(left��center��right)
var swf_height = focus_height+text_height //���֮�������ż��,�������ֻ����ģ��ʧ�������

document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">');
document.write('<param name="movie" value="js/focus.swf"><param name="quality" value="high"><param name="bgcolor" value="#eced6d"><param name="fcolor" value="#ffffff">');
document.write('<param name="menu" value="false"><param name=wmode value="transparent">');
document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'&text_align='+text_align+'&interval_time='+interval_time+'">');
document.write('<embed src="js/focus.swf" wmode="transparent" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'&text_align='+text_align+'&interval_time='+interval_time+'" menu="false" bgcolor="#eced6d" fcolor="#ffffff" quality="high" width="'+ focus_width +'" height="'+ swf_height +'" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');
document.write('</object>');
}