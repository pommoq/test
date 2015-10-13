
var toggle;
var popUpWin=0;
var windowFeatures = "toolbar=no,dependent=yes, location=1, status=1, menubar=no, scrollbars=no, resizable=no, height=" + screen.height/2 + ", width=" + screen.width/2 + ", top=" + ((screen.height - (screen.height/2))/2) +",left="+((screen.width - (screen.width/2))/2);

function popUpWindow(URLStr,obj)
{

  if(popUpWin)
  {
   // if(!popUpWin.closed) popUpWin.close();
	 popUpWin.focus();
  }
  else
  {
	  URLStr = URLStr + "?return=" + obj.form.id  + '.' + obj.id;
	  popUpWin = open(URLStr, '',windowFeatures);
	  document.form1.textfield.value = popUpWin;
	  popUpWin.focus();
  }
}

function return_resualt()
{
	 eval('window.opener.document.<%=Request("return")%>.value = document.form1.textfield.value');
	 window.close();
	 window.opener.popUpWin = 0;
}

function setFirstToggle() { 
	if(window.opener != null)
	{
		window.opener.toggle=true; 
	}
}// end function 

function setSecondToggle() { 
	if(window.opener != null)
	{
		window.opener.toggle=false; 
		window.opener.popUpWin = 0;
	}
}// end function 

function onTop() { 
  if (toggle==true) { 
     popUpWin.focus(); 
  } else if (toggle==false) { 
     return; 
  }// end if 
}// end function 

alert(self.onfocus)
self.onfocus=onTop;
alert(self.onfocus)
self.onclick=onTop;
self.onunload=setSecondToggle;
self.onload=setFirstToggle;

