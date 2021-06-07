<!--
//vbscript-like functions for javascript
//------------------------------------------------------------
//contains, lcase, ucase, pcase
//left, leftOf, right, rightOf, mid, replace
//inStr, inStrRev
//ltrim, rtrim, trim

function contains(sourcestr,searchstr) {
	//returns true if lrgstring contains smstring.
	strlen1 = searchstr.length;
	strlen2 = sourcestr.length;
	istrue = false;
	for (i=0;i<=strlen2;i++) {
		comp=sourcestr.substring(i-1,strlen2);
		comp = comp.substring(0,strlen1);
		if (comp == searchstr) {
			istrue = true;
			break;
		}
	}
	return istrue;
}

function lcase(str) {
	//returns str in all lowercase letters.
	return str.toLowerCase();
}

function left(str,n) {
	//returns the left n characters from str.
	return str.substring(0,n);
}

function leftOf(sourcestr,searchstr) {
	//returns leftmost characters of lrgstring up to smstring.
	//If user passes an empty string, change that to a space.
	if (searchstr == ""){searchstr = " ";}
	strlen1 = searchstr.length;
	strlen2 = sourcestr.length;
	foundat = 0;
	for (i=0;i<=strlen2;i++) {
		comp=sourcestr.substring(i-1,strlen2);
		comp = comp.substring(0,strlen1);	
		if (comp == searchstr) {
			foundat = i;
			break;
		}
	}
	return sourcestr.substring(0,(foundat-1));
}

function mid(str,start,n) {
	//returns a substring of str starting at 'start' that's n characters long.
	strlen = str.length;
	var jj = str.substring(start-1,strlen);
	jj = jj.substring(0,n);
	return jj;
}

function pcase(str) {
	//returns str in proper-noun case (first letter uppercase)
	strlen = str.length;
	jj = str.substring(0,1).toUpperCase();
	jj = jj + str.substring(1,strlen).toLowerCase();
	for (i = 2; i <= strlen; i++) {
		if (jj.charAt(i)==" ") {
			lefthalf = jj.substring(0,i+1);
			righthalf = jj.substring(i+1,strlen);
			righthalf = righthalf.substring(0,1).toUpperCase()+righthalf.substring(1,strlen);
			jj=lefthalf+righthalf;
		}
	}
	return jj;
}

function right(str,n) {
	//returns the right n characters of str
	strlen = str.length;
	return str.substring(strlen-n,strlen);
}

function rightOf(smstring,lrgstring) {
	//returns the rightmost characters of lrgstring back to smstring.
	//If user passes an empty string, change that to a space.
	if (smstring == ""){smstring = " ";}
	strlen1 = smstring.length;
	strlen2 = lrgstring.length;
	foundat = 0;
	for (i=strlen2;i>=0;i--) {
		comp=lrgstring.substring(i-1,strlen2);
		comp = comp.substring(0,strlen1);
		if (comp == smstring) {
			foundat = i;
			break;
		}
	}
	return lrgstring.substring(foundat,255);
}

function inStr(lrgstring,smstring) {
	//returns a number indicating the spot where smstring appears in lrgstring (left to right search)
	strlen1 = smstring.length;
	strlen2 = lrgstring.length;
	foundAt = -1;
	for (i=0;i<=strlen2;i++) {
		comp = lrgstring.substring(i-1,strlen2);
		comp = comp.substring(0,strlen1);
		if (comp == smstring) {
			foundAt = i;
			break;
		}
	}
	return foundAt;
}

function inStrRev(lrgstring,smstring) {
  //returns a number indicating the spot where smstring appears in lrgstring (right to left search)
	strlen1 = smstring.length;
	strlen2 = lrgstring.length;
	foundAt = -1;
	for (i=strlen2;i>=0;i--) {
		comp=lrgstring.substring(i-1,strlen2);
		comp = comp.substring(0,strlen1);	
		if (comp == smstring) {
			foundAt = i;
			break;
		}
	}
	return foundAt;
}

function ucase(str) {
	//returns str in all uppercase letters.
	return str.toUpperCase();
}

function replace(str, oldseq, newseq) {
  //replaces the oldseq string with the newseq string in the source string
  var r = "";
  len = oldseq.length;
  
  pos = inStr(str, oldseq);
  while (pos >= 0) {
    if (pos != 0) {
      r = r + left(str,pos-1) + newseq;
      str = mid(str,pos+len,str.length-(pos+len-1));
    }
    else {
      r = r + newseq;
      str = mid(str,pos+len+1,str.length-(pos+len-1));
    }
    pos = inStr(str, oldseq);
  }
  r = r + str;
  return r;
}

function ltrim(str) {
  //returns a string stripped of leading spaces 
  
  var i = 1;
  while (str.charAt(i) == " ")
    i++;
    
  return right(str,str.length-i);
}

function rtrim(str) {
  //returns a string stripped of trailing spaces 
  
  var i = str.length;
  while (str.charAt(i-1) == " ")
    i--;
    
  return left(str,i);
}

function trim(str) {
  //returns a string stripped of leading and trailing spaces
  str = ltrim(str);
  return rtrim(str);
}
//-->