var xmlhttp;

window.onload = onload;

if (document.addEventListener) {
   document.addEventListener("DOMContentLoaded", onload, false);
}


function getlatestquote()
{
	xmlhttp=GetXmlHttpObject();
	if (xmlhttp==null)
	{
		return;
	}
	
	var url="clean.html";
	xmlhttp.onreadystatechange=stateChanged;
	xmlhttp.open("GET",url,true);
	xmlhttp.send(null);
}

function stateChanged()
{
	if (xmlhttp.readyState==4)
	{
		document.getElementById("content").innerHTML=xmlhttp.responseText;
		setTimeout("getlatestquote()", 5000);
	}
}

function GetXmlHttpObject()
{
	if (window.XMLHttpRequest)
	{
		// code for IE7+, Firefox, Chrome, Opera, Safari
		return new XMLHttpRequest();
	}
	if (window.ActiveXObject)
	{
		// code for IE6, IE5
		return new ActiveXObject("Microsoft.XMLHTTP");
	}
	return null;
}

function onload()
{
	hideform();
	setTimeout("getlatestquote()", 5000);
}

function hideform() 
{
	document.getElementById("submitform").style.display = 'none';
	document.getElementById("moreawesome").style.visibility = 'visible';
	document.getElementById("content").style.visibility = 'visible';
}

function showform() 
{
	document.getElementById("submitform").style.display = 'block';
	document.getElementById("moreawesome").style.visibility = 'hidden';
	document.getElementById("content").style.visibility = 'hidden';
	
	document.getElementById("wintextinput").focus();
	document.getElementById("wintextinput").select();
	
}