/**
* followDiv plugin (跟随屏幕滚动的DIV)
* Copyright (c) 2008 CssRain (CssRain.cn)
* (Use)使用方法：
*   $('#div1').followDiv();
*   $('#div1').followDiv("rightbottom");
*	 $('#div2').followDiv("leftbottom");
*	 $('#div3').followDiv("lefttop");
*	 $('#div4').followDiv("righttop");
* 兼容 IE6+ ，FF2+ .
*/
(function($) {
$.fn.extend({
	"followDiv":function(str){
		var _self = this;
		var pos; //层的绝对定位位置
		switch(str){
			case("rightbottom")://右下角
				pos={"right":"0px","bottom":"0px"};
				break;
			case("leftbottom")://左下角
				pos={"left":"0px","bottom":"0px"};
				break;	
			case("lefttop"): //左上角
				pos={"left":"0px","top":"0px"};
				break;
			case("righttop")://右上角
				pos={"right":"0px","top":"0px"};
				break;
			default :   //默认为右下角
				pos={"right":"0px","bottom":"0px"};
				break;
		}
		/*FF和IE7可以通过position:fixed来定位，*/
		_self.css({"position":"fixed","z-index":"0"}).css(pos);
		/*ie6需要动态设置距顶端高度top.*/
		if($.browser.msie && $.browser.version == 6) {
				_self.css('position','absolute');						
				$(window).scroll(function(){
					var topIE6;
					if(str=="rightbottom"||str=="leftbottom"){
						topIE6=$(window).scrollTop() + $(window).height() - _self.outerWidth();
					}else if(str=="lefttop"||str=="righttop"){
						topIE6=$(window).scrollTop();
					}else{
						topIE6=$(window).scrollTop() + $(window).height() - _self.outerWidth();
					}
					_self.css( 'top' , topIE6 );
				});
		}
		return _self;  //返回this，使方法可链。
	}
});
})(jQuery);

$(document).ready(function(){
   $('#div0').followDiv();
   $('#div1').followDiv("rightbottom");
   $('#div2').followDiv("leftbottom");
   $('#div3').followDiv("lefttop");
   $('#div4').followDiv("righttop");
});