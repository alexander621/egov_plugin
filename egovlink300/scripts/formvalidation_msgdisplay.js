// START OF MESSAGE SCRIPT //
var MSGTIMER  = 20;
var MSGSPEED  = 2;
var MSGOFFSET = 10;
var MSGHIDE   = 5;

// build out the divs, set attributes and call the fade function //
function inlineMsg(target,string,autohide,id) {
  var msg;
  var msgcontent;
  if(!document.getElementById('msg'+id)) {
	msg                  = document.createElement('div');
    msg.id               = 'msg'+id;
	msgcontent           = document.createElement('div');
    msgcontent.id        = 'msgcontent'+id;
    msg.className        = "msg";
    msgcontent.className = "msgcontent";
	document.body.appendChild(msg);
    msg.appendChild(msgcontent);
	msg.style.filter  = 'alpha(opacity=0)';
    msg.style.opacity = 0;
	msg.alpha         = 0;
  } else {
    msg        = document.getElementById('msg'+id);
    msgcontent = document.getElementById('msgcontent'+id);
  }

  msgcontent.innerHTML = string;
  msg.style.display    = "block";
  
  var msgheight = msg.offsetHeight;
  var targetdiv = document.getElementById(target);
  
  targetdiv.focus();
  
  var targetheight = targetdiv.offsetHeight;
  var targetwidth  = targetdiv.offsetWidth;
  var topposition  = topPosition(targetdiv) - ((msgheight - targetheight) / 2);
  var leftposition = leftPosition(targetdiv) + targetwidth + MSGOFFSET;

  msg.style.top  = topposition + 'px';
  msg.style.left = leftposition + 'px';

  clearInterval(msg.timer);
  msg.timer = setInterval("fadeMsg(1,'"+id+"')", MSGTIMER);

  if(!autohide) {
     autohide = MSGHIDE;  
  }

  window.setTimeout("hideMsg('"+id+"')", (autohide * 1000));
}

// hide the form alert //
function hideMsg(id) {
  var msg = document.getElementById('msg'+id);
  if(!msg.timer) {
	msg.timer = setInterval("fadeMsg(0,'"+id+"')", MSGTIMER);
  }
}

// face the message box //
function fadeMsg(flag,id) {
  if(flag == null) {
    flag = 1;
  }
  var msg = document.getElementById('msg'+id);
  var value;
  if(flag == 1) {
	value = msg.alpha + (MSGSPEED*5);
  } else {
    value = msg.alpha - MSGSPEED;
  }

  msg.alpha = value;
  msg.style.opacity = (value / 100);
  msg.style.filter = 'alpha(opacity=' + value + ')';
  if(value >= 99) {
    clearInterval(msg.timer);
    msg.timer = null;
  } else if(value <= 1) {
	msg.style.display = "none";
    clearInterval(msg.timer);
  }
}

// calculate the position of the element in relation to the left of the browser //
function leftPosition(target) {
  var left = 0;
  if(target.offsetParent) {
    while(1) {
      left += target.offsetLeft;
      if(!target.offsetParent) {
        break;
      }
      target = target.offsetParent;
    }
  } else if(target.x) {
    left += target.x;
  }
  return left;
}

// calculate the position of the element in relation to the top of the browser window //
function topPosition(target) {
  var top = 0;
  if(target.offsetParent) {
    while(1) {
      top += target.offsetTop;
      if(!target.offsetParent) {
        break;
      }
      target = target.offsetParent;
    }
  } else if(target.y) {
    top += target.y;
  }
  return top;
}

function clearMsg(id) {
  if(document.getElementById('msg'+id)) {
     document.getElementById('msg'+id).style.display = "none";
  }
}

// preload the arrow //
if(document.images) {
  arrow = new Image(7,80); 
  arrow.src = "/eclink/images/msg_arrow.gif"; 
}
