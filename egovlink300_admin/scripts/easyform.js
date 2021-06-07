<!--
//*****************************************************************************************
// easyform.js
//
// a javascript file to validate form elements
// created 11/99 by Chip Kellam of ec link <ckellam@eclink.com>
//
// use:
//   <script language="Javascript" src="easyform.js"></script>
//   <input type=hidden name="ef:<FormElementName><-ElementType>/<arguments - req/datatype>
//   <input type=button value="Submit"
//          onclick="if (validateForm('<FormName>')) { document.<FormName>.submit(); }">
//
// example:
//   <input type=hidden name="ef:txtAge-text/req/number">
//   <input type=button value="Submit"
//          onclick="if (validateForm('MyForm')) { document.MyForm.submit(); }">
//*****************************************************************************************

//----------------------------------------------------------------------------
// User Defined Constants
//----------------------------------------------------------------------------
ERROR_HEADER = "The following problems were found with this form:\n\n";
ERROR_FOOTER = "\nPlease fix these errors, then try again."

//----------------------------------------------------------------------------
// Data Type Constants
//----------------------------------------------------------------------------
REQUIRED = "req";
REQUIRED_LONG = "required";

NUMBER = "number";
INTEGER = "integer";
DATE = "date";
EMAIL = "email";
PHONE = "phone";
ZIPCODE = "zip";
CC = "cc";
SSN = "ssn";
CURRENCY = "currency";
ACCOUNTCODE = "accountcode";

//----------------------------------------------------------------------------
// Main Functions
//----------------------------------------------------------------------------

//main handler for form validation, checks prefixes, etc..
function validateForm(formName) {
  var errMsg = "";
  var count;
  var err;
  var numElements = eval("document." + formName + ".length");

  for (count=0; count<numElements; count++) {
    if (eval("document." + formName + ".elements[" + count + "].type") == "hidden") {
      arg = eval("document." + formName + ".elements[" + count + "].name");

      if (lcase(left(arg,3)) == "ef:") {
        arg = mid(arg,4,arg.length);
        pos = inStr(arg, "/");
        if (pos > 0) {

          //determine name of element
          elementName = mid(arg,1,pos-1);

          //parse out arguments
          arg = lcase(mid(arg,elementName.length+2,arg.length-elementName.length-1));
          pos = inStr(arg, "/")
          if (pos > 0) {
            arg2 = mid(arg,pos+1,arg.length-pos);
            arg = left(arg,pos-1);
          } else {
            arg2 = "";
          }

          //determine type of element
          pos = inStr(elementName, "-")
          if (pos > 0) {
            elementType = mid(elementName, pos+1, elementName.length-pos);
            elementName = left(elementName, pos-1);
          } else {
            elementType = "text";
          }

          //determine if we have a custom element display name
          errName = eval("document." + formName + ".elements[" + count + "].value");
          if (errName == "")
            errName = elementName;

          //validate form element against arguements
          err = validateElement(formName, elementName, elementType, arg);
          if (err == "" && arg2 != "")
            err = validateElement(formName, elementName, elementType, arg2);

          if (err != "")
            errMsg = errMsg + "   - " + errName + err + "\n";
        }
        else {
          //if no arguments but a error header or footer is found..
          if (arg == ":ErrorHeader") {
            ERROR_HEADER = eval("document." + formName + ".elements[" + count + "].value");
            ERROR_HEADER = replace(ERROR_HEADER, "\\n", "\n");
          }
          if (arg == ":ErrorFooter") {
            ERROR_FOOTER = eval("document." + formName + ".elements[" + count + "].value");
            ERROR_FOOTER = replace(ERROR_FOOTER, "\\n", "\n");
          }
        }
      }
    }
  }

  if (errMsg != "") {
    alert(ERROR_HEADER + errMsg + ERROR_FOOTER);
    return false;
  }
  return true;
}

//checks the type expected against any function that may be set in place
function validateElement(formName, elementName, elementType, arg) {
var j;

  elementValue = "";
  switch( lcase(elementType) ) {
    case "text":
    case "textarea":
    case "select":
      elementValue = eval("document." + formName + "." + elementName + ".value");
      break;

    case "radio":
    case "checkbox":
    // MULTIPLE CHECK ELEMENT 
	for (j=0; j<eval("document." + formName + "." + elementName + ".length"); j++){
		if (eval("document." + formName + "." + elementName + "[j].checked"))
          elementValue = "cheese is good.";}
	 
	 // SINGLE CHECKBOX ELEMENT
		if (eval("document." + formName + "." + elementName + ".checked"))
          elementValue = "cheese is good.";

      break;
  }



  switch(arg) {
    case "req":
    case "required":
      if (elementValue == "")
        return " is a required field.";
      break;

    case "number":
      if (!isNumeric(elementValue))
        return " must be a number.";
      break;
      
    case "userpass":
      if (eval("document." + formName + ".userpass1.value") != eval("document." + formName + ".userpass2.value"))
        return " must be equal to password.";
      break;
    
    case "password":
      if (eval("document." + formName + ".password.value") != eval("document." + formName + ".password2.value"))
        return " must be equal to password.";
      break;
    
    case "integer":
      if (!isInteger(elementValue))
        return " must be an integer.";
      break;

    case "date":
      if (!isDate(elementValue))
        return " must be a valid date.";
      break;

    case "email":
      if (!isEmail(elementValue))
        return " must be a valid email address.";
      break;

    case "phone":
      if (!isPhone(elementValue))
        return " must be a valid phone number.";
      break;

    case "cc":
      if (!isCreditCard(elementValue))
        return " must be a valid Credit Card Number.";
      break;

    case "ssn":
      if (!isSSN(elementValue))
        return " must be a valid Social Security Number.";
      break;

    case "zip":
      if (!isZipCode(elementValue))
        return " must be a valid zip code.";
      break;
      
    case "passcheck":
      if (!isPassCheck(elementValue))
        return " must be 5-10 characters in length.";
      break;
   
   case "stmtdatechk":
      if (!isStmtDateChk(elementValue))
        return " \n     Statements are not available online before 199809.";
      break;

    case "currency":
      if (!isCurrency(elementValue))
        return " must be in valid currency format.";
      break;
    
    case "customerid":
      if (!isCustomerID(elementValue))
        return " must be a combination of 10 letters/numbers in length.";
      break;
	  
	case "locationid":
      if (!isLocationID(elementValue))
        return " must be a combination of 9 letters/numbers in length.";
      break;  
	
	case "ninedigits":
      if (!ninedigits(elementValue))
        return " must be a combination of 9 numbers only or less in length.";
      break;  
	
	case "twodigits":
      if (!twodigits(elementValue))
        return " must be two digits.";
      break;  
      
  	case "threedigits":
      if (!threedigits(elementValue))
        return " must be three digits.";
      break;  
      
  	case "fourdigits":
      if (!fourdigits(elementValue))
        return " must be three or four digits.";
      break; 
	  
  	case "sevendigits":
      if (!sevendigits(elementValue))
        return " must be seven digits.";
      break;  

  	case "asn":
      if (!asn(elementValue))
        return " must be SA and two digits. Example SA12.";
      break;  
      
  	case "ppn":
      if (!ppn(elementValue))
        return " must be three digits or three digits with a letter. Example 123 or 123B.";
      break;
	  
  	case "thirteenchar":
      if (!thirteenchar(elementValue))
        return " must be less than 13 characters.";
      break;  
      
      
      
  }

  return "";
}


//--------------------------------------------------------------------------
// DataType Validation Routines
//--------------------------------------------------------------------------

//checks a string to see if it is a valid zipcode
function isZipCode (value){
	var x,i;

	//Convert 5 digit zip to 9 digit by adding "0000"
	if (value.length == 5) {
		if (isNumeric(value)) {
			value=value+"0000";
			return true;
		} else {
			return false;
		}
	}

	// If length ten (XXXXX-YYYY), remove non-numeric then check length 9
	if(value.length == 10) {
		value=numericize(value);
	}

	//Check length 9
	if (value.length == 9) {
		if (isNumeric(value)) {
			return true;
		} else {
			return false;
		}
	}
	return false;
}


//checks a string to see if it is a valid customerid
function isCustomerID(value){

	//Check length 10
	if (value.length == 10) {
			return true;
	}

	return false;
}


// CHECKS TO A STRING TO MAKE SURE ITS NINE DIGITS
function ninedigits(value){

	var regmaxninedigit = new RegExp(/^\d{1,9}$/); // FIND ANY 1-4 DIGIT NUMBER
	//Check length 9
	if (value.match(regmaxninedigit)) {
			return true;
	}
	return false;
}

// CHECKS TO A STRING TO MAKE SURE ITS TWO DIGITS
function twodigits(value){

	var regmaxtwodigit = new RegExp(/^\d{2}$/); // FIND ANY 2 DIGIT NUMBER
	//Check length 2
	if (value.match(regmaxtwodigit)) {
			return true;
	}
	return false;
}

// CHECKS TO A STRING TO MAKE SURE ITS THREE DIGITS
function threedigits(value){

	var regmaxthreedigit = new RegExp(/^\d{3}$/); // FIND 3 DIGIT NUMBER
	//Check length 3
	if (value.match(regmaxthreedigit)) {
			return true;
	}
	return false;
}


// CHECKS TO A STRING TO MAKE SURE LESS THAN 13 CHARS
function thirteenchar(value){

	var regmax = new RegExp(/^\S{0,13}$/); // 
	//Check length 3
	if (value.match(regmax)) {
			return true;
	}
	return false;
}


// CHECKS TO A STRING TO MAKE SURE ITS FOUR DIGITS
function fourdigits(value){

	var regmaxfourdigit = new RegExp(/^\d{3,4}$/); // FIND ANY 3-4 DIGIT NUMBER
	
	if (value.length < 1)
	{
			return true;
	}
	
	//Check length 4
	if (value.match(regmaxfourdigit)) {
			return true;
	}
	return false;
}

// CHECKS TO A STRING TO MAKE SURE ITS SEVEN DIGITS
function sevendigits(value){

	var regmaxsevendigit = new RegExp(/^\d{1,7}$/); // FIND ANY 1-7 DIGIT NUMBER
	
	if (value.length < 1)
	{
			return true;
	}
	
	//Check length 7
	if (value.match(regmaxsevendigit)) {
			return true;
	}
	return false;
}


// CHECKS TO A STRING TO MAKE SURE ITS DIGITS ARE IN FORM SA WITH TWO DIGITS EX. SA12
function asn(value){

	var regmaxsevendigit = new RegExp(/^(SA)|(sa)|(Sa)|(sA)\d{2}/); // FIND SA with 2 trailing digits
	
	if (value.length < 1)
	{
			return true;
	}
	
	//Check length 7
	if (value.match(regmaxsevendigit)) {
			return true;
	}
	return false;
}

// CHECKS TO A STRING TO MAKE SURE FORMAT IS 123 or 123B
function ppn(value){

	var regmaxsevendigit = new RegExp(/^\d{3}[a-zA-Z]$|^\d{3}$/); // FIND 123 or 123B
	
	if (value.length < 1)
	{
			return true;
	}
	
	//Check length 7
	if (value.match(regmaxsevendigit)) {
			return true;
	}
	return false;
}


//checks a string to see if it is a valid locationid
function isLocationID(value){

	//Check length 9
	if (value.length == 9) {
			return true;
	}
	return false;
}


//returns true if string is integer
function isInteger(s) {
  if (isNumeric(s)) {
    if (parseInt(s) == s)
      return true;
  }
  return false;
}

//returns true if string is numeric
function isNumeric(s) {
  /*var i;
	for (i=0; i<s.length; i++){
	  var c = s.charAt(i);
	  if (!isDigit(c))
			return false;
	}
	return true;*/
	return (!isNaN(s));
}

//returns true if string is a currency
function isCurrency(s) {
  s = replace(s, "$", "");
  s = replace(s, ".", "");
  s = replace(s, ",", "");
  return isNumeric(s);
}

//returns true if string is a date
function isDate(s) {
  s = replace(s, "/", "");
  s = replace(s, "-", "");
  s = replace(s, ":", "");
  s = replace(s, "AM", "");
  s = replace(s, "PM", "");
  return isNumeric(s);
}

//return true is string is a valid credit card
function isCreditCard(s) {
  s = replace(s, "-", "");
  s = replace(s, " ", "");
  return isNumeric(s);
}

//returns true if c is a number
function isDigit (c) {
	return ((c >= "0") && (c <= "9"))
	
}

//returns true if date is after 199809
function isStmtDateChk (c) {
	return (c > 199808)
	
}
function numericize(s){
	// s is a string
	var i,j;
	j="";
	for (i=0;i<s.length;i++){
		if (isDigit(s.charAt(i))){
			 j = j + s.charAt(i);
		}
	}
	return j;
}

//returns true if is a valid Social Security Number
function isSSN (s) {
	var x,i;

	// Check 11 digit with dashes (strip out dashes, check if valid)
	if (s.length == 11){
		s=numericize(s);
	}

	// Check 9 digit integer
	if (s.length == 9){
		if (isNumeric(s)){
			return true;
		} else {
			return false;
		}
	}
	return false;
}

//returns true if length is correct
function isPassCheck (s) {
	return ((s.length >= 5) && (s.length <= 10))
}
	
function isPhone (value) {
	if (value.length > 10){
		value = numericize(value);
	}

	//Check if first digit is 0 or 1. Invalid phone number.
	if ((value.charAt(0) == '0') || (value.charAt(0) == '1')) {
		return false;
	}

	//Should be 3 digit area code + 7 digit phone number, as a 10 digit string
	if (value.length == 10){
		if (isNumeric(value)){
			return true;
		} else {
			return false;
		}
	}
	else{
		return false;
	}
	return true;
}

function isEmail (value){
	var i,ii;
	var j;
	var k,kk;
    var jj;
    var len;

    // Check valid email
    // Must have a "@" and a "." to be valid.
    // Must have at least 1 character before "@"
    // Must have at least 1 character after "@" and before "."
    // Must have at least 2 characters after "."
    if (value.length >0){
		i=value.indexOf("@");
		ii=value.indexOf("@",i+1);
		j=value.indexOf(".",i);
		k=value.indexOf(",");
		kk=value.indexOf(" ");
		jj=value.lastIndexOf(".")+1;
		len=value.length;
		if ((i>0) && (j>(1+1)) && (k==-1) && (ii==-1) && (kk==-1) &&
			(len-jj >=2) && (len-jj<=3)) {}
		else {
				return false;
		}
	}
    return true;
}

//--------------------------------------------------------------------------
// String Functions
//--------------------------------------------------------------------------

//returns the left n characters from str.
function left(str,n) {
	return str.substring(0,n);
}

//returns a substring of str starting at 'start' that's n characters long.
function mid(str,start,n) {
	strlen = str.length;
	var jj = str.substring(start-1,strlen);
	jj = jj.substring(0,n);
	return jj;
}

//returns a number indicating the spot where smstring appears in lrgstring (right to left search)
function inStrRev(lrgstring,smstring) {
	strlen1 = smstring.length;
	strlen2 = lrgstring.length;
	foundAt = 0;
	for (i=strlen2;i>=0;i--) {
		comp = lrgstring.substring(i-1,strlen2);
		comp = comp.substring(0,strlen1)	;
		if (comp == smstring) {
			foundAt = i;
			break;
		}
	}
	return foundAt;
}

function lcase(str) {
	//returns str in all lowercase letters.
	return str.toLowerCase()
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
//-->
