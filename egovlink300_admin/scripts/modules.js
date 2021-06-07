<!--
    var view;
    var opacity;

	function getWindowWidth() 
	{
		var myWidth = 0;
		if( typeof( window.innerWidth ) == 'number' ) {
			//Non-IE
			myWidth = window.innerWidth;
		} else if( document.documentElement && ( document.documentElement.clientWidth ) ) {
			//IE 6+ in 'standards compliant mode'
			myWidth = document.documentElement.clientWidth;
		} else if( document.body && ( document.body.clientWidth ) ) {
			//IE 4 compatible
			myWidth = document.body.clientWidth;
		}
		return myWidth;
	}


	function getWindowHeight()
	{
		var myHeight = 0;
		if( typeof( window.innerHeight ) == 'number' ) {
			//Non-IE
			myHeight = window.innerHeight;
		} else if( document.documentElement && ( document.documentElement.clientHeight ) ) {
			//IE 6+ in 'standards compliant mode'
			myHeight = document.documentElement.clientHeight;
		} else if( document.body && ( document.body.clientHeight ) ) {
			//IE 4 compatible
			myHeight = document.body.clientHeight;
		}
		return myHeight;
	}

    function tmptoggleDisplay( obj, viewName ) {
      view = eval("document.all." + viewName);

      if (view.style.display == "none") {
        opacity = 0;
        view.style.display = "";
        setTimeout("fadeUp()", 25);
        obj.src = "../images/arrow_collapse.jpg";
      }
      else {
        opacity = 100;
        fadeDown();
        obj.src = "../images/arrow_expand.jpg";
      }
    }

    function toggleDisplay( obj, viewName ) {
      view = eval("document.all." + viewName);
      
      if (view.style.display == "none") {
        opacity = 0;
        view.style.display = "";
        setTimeout("fadeUp()", 25);
        obj.src = "images/arrow_collapse.jpg";
      }
      else {
        opacity = 100;
        fadeDown();
        obj.src = "images/arrow_expand.jpg";
      }
    }

    function fadeUp() {
      if (opacity <= 100) {
        opacity += 20;
        view.style.filter = "alpha(opacity:" + opacity + ")";
        setTimeout("fadeUp()", 25);
      }
    }

    function fadeDown() {
      if (opacity > 0) {
        opacity -= 20;
        view.style.filter = "alpha(opacity:" + opacity + ")";
        setTimeout("fadeDown()", 25);
      }
      else {
        view.style.display = "none";
      }
    }

    function doCalendar() {
      w = (screen.width - 600)/2;
      h = (screen.height - 550)/2;
      eval('window.open("events/calendar.asp?p=1", "_calendar", "width=600,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
    }

  function emailCheck (emailStr) {

/* The following variable tells the rest of the function whether or not
to verify that the address ends in a two-letter country or well-known
TLD.  1 means check it, 0 means don't. */

var checkTLD=1;

/* The following is the list of known TLDs that an e-mail address must end with. */

var knownDomsPat=/^(us|com|net|org|edu|int|mil|gov|arpa|biz|aero|name|coop|info|pro|museum)$/;

/* The following pattern is used to check if the entered e-mail address
fits the user@domain format.  It also is used to separate the username
from the domain. */

var emailPat=/^(.+)@(.+)$/;

/* The following string represents the pattern for matching all special
characters.  We don't want to allow special characters in the address. 
These characters include ( ) < > @ , ; : \ " . [ ] */

var specialChars="\\(\\)><@,;:\\\\\\\"\\.\\[\\]";

/* The following string represents the range of characters allowed in a 
username or domainname.  It really states which chars aren't allowed.*/

var validChars="\[^\\s" + specialChars + "\]";

/* The following pattern applies if the "user" is a quoted string (in
which case, there are no rules about which characters are allowed
and which aren't; anything goes).  E.g. "jiminy cricket"@disney.com
is a legal e-mail address. */

var quotedUser="(\"[^\"]*\")";

/* The following pattern applies for domains that are IP addresses,
rather than symbolic names.  E.g. joe@[123.124.233.4] is a legal
e-mail address. NOTE: The square brackets are required. */

var ipDomainPat=/^\[(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\]$/;

/* The following string represents an atom (basically a series of non-special characters.) */

var atom=validChars + '+';

/* The following string represents one word in the typical username.
For example, in john.doe@somewhere.com, john and doe are words.
Basically, a word is either an atom or quoted string. */

var word="(" + atom + "|" + quotedUser + ")";

// The following pattern describes the structure of the user

var userPat=new RegExp("^" + word + "(\\." + word + ")*$");

/* The following pattern describes the structure of a normal symbolic
domain, as opposed to ipDomainPat, shown above. */

var domainPat=new RegExp("^" + atom + "(\\." + atom +")*$");

/* Finally, let's start trying to figure out if the supplied address is valid. */

/* Begin with the coarse pattern to simply break up user@domain into
different pieces that are easy to analyze. */

var matchArray=emailStr.match(emailPat);

if (matchArray==null) {

/* Too many/few @'s or something; basically, this address doesn't
even fit the general mould of a valid e-mail address. */

alert("Email address seems incorrect (check @ and .'s)");
return false;
}
var user=matchArray[1];
var domain=matchArray[2];

// Start by checking that only basic ASCII characters are in the strings (0-127).

for (i=0; i<user.length; i++) {
if (user.charCodeAt(i)>127) {
alert("Ths username contains invalid characters.");
return false;
   }
}
for (i=0; i<domain.length; i++) {
if (domain.charCodeAt(i)>127) {
alert("Ths domain name contains invalid characters.");
return false;
   }
}

// See if "user" is valid 

if (user.match(userPat)==null) {

// user is not valid

alert("The username doesn't seem to be valid.");
return false;
}

/* if the e-mail address is at an IP address (as opposed to a symbolic
host name) make sure the IP address is valid. */

var IPArray=domain.match(ipDomainPat);
if (IPArray!=null) {

// this is an IP address

for (var i=1;i<=4;i++) {
if (IPArray[i]>255) {
    alert("Destination IP address is invalid!");
    return false;
}
}
return true;
}

// Domain is symbolic name.  Check if it's valid.
 
var atomPat=new RegExp("^" + atom + "$");
var domArr=domain.split(".");
var len=domArr.length;
for (i=0;i<len;i++) {
if (domArr[i].search(atomPat)==-1) {
    alert("The domain name does not seem to be valid.");
    return false;
   }
}

/* domain name seems valid, but now make sure that it ends in a
known top-level domain (like com, edu, gov) or a two-letter word,
representing country (uk, nl), and that there's a hostname preceding 
the domain or country. */

if (checkTLD && domArr[domArr.length-1].length!=2 && 
    domArr[domArr.length-1].search(knownDomsPat)==-1) {
    alert("The address must end in a well-known domain or two letter " + "country.");
    return false;
}

// Make sure there's a host name preceding the domain.

if (len<2) {
    alert("This address is missing a hostname!");
    return false;
}

// If we've gotten this far, everything's valid!
return true;
}

//This function checks the field length
//The ability to display the remaining characters is available.  To do so:
//  1. create a hidden field with an ID and NAME of "control_field".  
//     Set the maxlength on this field to the maximum length of the biggest field +1 that you will be checking.
//  2. create a <span> tab with an ID of "message_char_cnt"
//  3. pass in field ID (document.getElementById(...)) to evaluate against the control field
function checkFieldLength(p_value,p_limit,p_display_cnt,p_field_id) {
  lcl_length = p_value.length;
  if(lcl_length <= p_limit) {
     if(p_display_cnt=="Y") {
		document.getElementById('message_char_cnt').innerHTML = p_limit + " character limit.  Characters remaining: " + (p_limit - lcl_length);
	 }
  } else {
     p_field_id.value = document.getElementById("control_field").value.substr(0,p_limit);
     alert("Cannot exceed " + p_limit + " characters.");
  }
}

//This function checks the field length for multiple fields.
//The ability to display the remaining characters is available.  To do so:
//  1. create a hidden field with an ID and NAME of "control_field".  
//     Set the maxlength on this field to the maximum length of the biggest field +1 that you will be checking.
//  2. create a <span> tab with an ID of "message_char_cnt"
//  3. pass in field ID (document.getElementById(...)) to evaluate against the control field
//  4. pass in the rowID or some identifier that will find the proper record in the list.
function checkFieldLength_MultipleFields(p_value,p_limit,p_display_cnt,p_field_id,p_msg_id) {
  lcl_length = p_value.length;
  if(lcl_length <= p_limit) {
     if(p_display_cnt=="Y") {
		document.getElementById('message_char_cnt_'+p_msg_id).innerHTML = p_limit + " character limit.  Characters remaining: " + (p_limit - lcl_length);
	 }
  } else {
     p_field_id.value = document.getElementById("control_field").value.substr(0,p_limit);
     alert("Cannot exceed " + p_limit + " characters.");
  }
}

//The two functions handle the row highlight and un-highlight for result lists when the mouse cursor moves over and off a record
function mouseOverRow( oRow ) {
  oRow.style.backgroundColor='#93bee1';
  oRow.style.cursor='pointer';

//  oNextRow = document.getElementById(eval(parseInt(oRow.id) + 1));
//  if (oNextRow) {
//      oNextRow.style.backgroundRepeat="repeat-x";
//      oNextRow.style.backgroundImage="url(../images/shadow.png)";
//  }
}

function mouseOutRow( oRow ) {	
  oRow.style.backgroundColor='';
  oRow.style.cursor='';
//  oNextRow = document.getElementById(eval(parseInt(oRow.id) + 1));
//  if (oNextRow) {
//      oNextRow.style.backgroundImage="";
//  }
}
  //-->