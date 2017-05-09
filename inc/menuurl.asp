<%
function menuurl(leixing,step,id)
	select case leixing
		case "单篇文章"
			menuurl="ArticleShow.asp?menu"&step&"="&id
		case "多篇文章"
			menuurl="ArticleList.asp?menu"&step&"="&id
		case "多张图片"
			menuurl="PictureList.asp?menu"&step&"="&id
		case "表单提交"
			menuurl="Forms.asp?menu"&step&"="&id
	end select
end function
%>