// JavaScript Document
//last modify by xhg 2006-07-31
var MessTop,MessLeft,MessWidth,MessHeight,docHeight,docWidth,MessTimer,MessAccount = 0;
function initMess()
{
	try{
		MessTop = parseInt(document.getElementById("MessDiv").style.top,10);
		MessLeft = parseInt(document.getElementById("MessDiv").style.left,10);
		MessHeight = parseInt(document.getElementById("MessDiv").offsetHeight,10);
		MessWidth = parseInt(document.getElementById("MessDiv").offsetWidth,10);
		docWidth = document.body.clientWidth;
		docHeight = document.body.clientHeight;
		document.getElementById("MessDiv").style.top = parseInt(document.body.scrollTop,10) + docHeight + 10;
		document.getElementById("MessDiv").style.left = parseInt(document.body.scrollLeft,10) + docWidth - MessWidth;
		document.getElementById("MessDiv").style.visibility="visible";
		MessTimer = window.setInterval("moveMess()",10);
	}
	catch(e){}
}

function resizeMess()
{
	//MessAccount+=1
	//if(MessAccount>5000) closeMess()
	try{
		MessHeight = parseInt(document.getElementById("MessDiv").offsetHeight,10);
		MessWidth = parseInt(document.getElementById("MessDiv").offsetWidth,10);
		docWidth = document.body.clientWidth;
		docHeight = document.body.clientHeight;
		document.getElementById("MessDiv").style.top = docHeight - MessHeight + parseInt(document.body.scrollTop,10);
		document.getElementById("MessDiv").style.left = docWidth - MessWidth + parseInt(document.body.scrollLeft,10);
	}
	catch(e){}
}

function moveMess()
{
	try
	{
		if(parseInt(document.getElementById("MessDiv").style.top,10) <= (docHeight - MessHeight + parseInt(document.body.scrollTop,10)))
		{
			window.clearInterval(MessTimer);
			MessTimer = window.setInterval("resizeMess()",100);
		}
		MessTop = parseInt(document.getElementById("MessDiv").style.top,10);
		document.getElementById("MessDiv").style.top = MessTop - 1;
	}
	catch(e){}
}
function closeMess()
{
	document.getElementById('MessDiv').style.visibility='hidden';
	if(MessTimer) window.clearInterval(MessTimer);
}
