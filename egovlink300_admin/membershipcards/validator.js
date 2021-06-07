function Validator(frmname)
{
  this.formobj=document.forms[frmname];
	if(!this.formobj)
	{
	  alert("BUG: couldnot get Form object "+frmname);
		return;
	}
	if(this.formobj.onsubmit)
	{
	 this.formobj.old_onsubmit = this.formobj.onsubmit;
	 this.formobj.onsubmit=null;
	}
	else
	{
	 this.formobj.old_onsubmit = null;
	}
	this.formobj.onsubmit=form_submit_handler;
	this.addValidation = add_validation;
	this.setAddnlValidationFunction=set_addnl_vfunction;
	this.clearAllValidations = clear_all_validations;
}


function set_addnl_vfunction(functionname)
{
  this.formobj.addnlvalidation = functionname;
}
function clear_all_validations()
{
	for(var itr=0;itr < this.formobj.elements.length;itr++)
	{
		this.formobj.elements[itr].validationset = null;
	}
}
function form_submit_handler()
{
	for(var itr=0;itr < this.elements.length;itr++)
	{
		if(this.elements[itr].validationset &&
	   !this.elements[itr].validationset.validate())
		{
		  return false;
		}
	}
	if(this.addnlvalidation)
	{
	  str =" var ret = "+this.addnlvalidation+"()";
	  eval(str);
    if(!ret) return ret;
	}
	return true;
}
function add_validation(itemname,descriptor,errstr)
{
  if(!this.formobj)
	{
	  alert("BUG: the form object is not set properly");
		return;
	}//if
	var itemobj = this.formobj[itemname];
  if(!itemobj)
	{
	  alert("BUG: Couldnot get the input object named: "+itemname);
		return;
	}
	if(!itemobj.validationset)
	{
	  itemobj.validationset = new ValidationSet(itemobj);
	}
  itemobj.validationset.add(descriptor,errstr);
}
function ValidationDesc(inputitem,desc,error)
{
  this.desc=desc;
	this.error=error;
	this.itemobj = inputitem;
	this.validate=vdesc_validate;
}
function vdesc_validate()
{
 if(!V2validateData(this.desc,this.itemobj,this.error))
 {
    this.itemobj.focus();
		return false;
 }
 return true;
}
function ValidationSet(inputitem)
{
    this.vSet=new Array();
	this.add= add_validationdesc;
	this.validate= vset_validate;
	this.itemobj = inputitem;
}
function add_validationdesc(desc,error)
{
  this.vSet[this.vSet.length]= 
	  new ValidationDesc(this.itemobj,desc,error);
}
function vset_validate()
{
   for(var itr=0;itr<this.vSet.length;itr++)
	 {
	   if(!this.vSet[itr].validate())
		 {
		   return false;
		 }
	 }
	 return true;
}
function validateEmailv2(email)
{
// a very simple email validation checking. 
// you can add more complex email checking if it helps 
    if(email.length <= 0)
	{
	  return true; // I changed myself, in order to prevent the empy email from being passed.
	}
    var splitted = email.match("^(.+)@(.+)$");
    if(splitted == null) return false;
    if(splitted[1] != null )
    {
      var regexp_user=/^\"?[\w-_\.]*\"?$/;
      if(splitted[1].match(regexp_user) == null) return false;
    }
    if(splitted[2] != null)
    {
      var regexp_domain=/^[\w-\.]*\.[A-Za-z]{2,4}$/;
      if(splitted[2].match(regexp_domain) == null) 
      {
	    var regexp_ip =/^\[\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\]$/;
	    if(splitted[2].match(regexp_ip) == null) return false;
      }// if
      return true;
    }
return false;
}
function V2validateData(strValidateStr,objValue,strError) 
{ 
    var epos = strValidateStr.search("="); 
    var  command  = ""; 
    var  cmdvalue = ""; 
    if(epos >= 0) 
    { 
     command  = strValidateStr.substring(0,epos); 
     cmdvalue = strValidateStr.substr(epos+1); 
    } 
    else 
    { 
     command = strValidateStr; 
    } 
    switch(command) 
    { 
        case "req": 
        case "required": 
         { 
           if(eval(objValue.value.length) == 0) 
           { 
              if(!strError || strError.length ==0) 
              { 
                strError = objValue.name + " : Required Field"; 
              }//if 
              alert(strError); 
              return false; 
           }//if 
           break;             
         }//case required 
        case "maxlength": 
        case "maxlen": 
          { 
             if(eval(objValue.value.length) >  eval(cmdvalue)) 
             { 
               if(!strError || strError.length ==0) 
               { 
                 strError = objValue.name + " : "+cmdvalue+" characters maximum "; 
               }//if 
               alert(strError + "\n[Current length = " + objValue.value.length + " ]"); 
               return false; 
             }//if 
             break; 
          }//case maxlen 
        case "minlength": 
        case "minlen": 
           { 
             if(eval(objValue.value.length) <  eval(cmdvalue)) 
             { 
               if(!strError || strError.length ==0) 
               { 
                 strError = objValue.name + " : " + cmdvalue + " characters minimum  "; 
               }//if               
               alert(strError + "\n[Current length = " + objValue.value.length + " ]"); 
               return false;                 
             }//if 
             break; 
            }//case minlen 
        case "alnum": 
        case "alphanumeric": 
           { 
              var charpos = objValue.value.search("[^A-Za-z0-9\.]"); 
              if(objValue.value.length > 0 &&  charpos >= 0) 
              { 
               if(!strError || strError.length ==0) 
                { 
                  strError = objValue.name+": Only alpha-numeric characters allowed "; 
                }//if 
                alert(strError + "\n [Error character position " + eval(charpos+1)+"]"); 
                return false; 
              }//if 
              break; 
           }//case alphanumeric 
        case "num": 
        case "numeric": 
           { 
		      var charpos = objValue.value.search("[^0-9\.]"); 
              if(objValue.value.length > 0 &&  charpos >= 0) 
              { 
                if(!strError || strError.length ==0) 
                { 
                  strError = objValue.name+": Only digits and . are allowed "; 
                }//if               
                alert(strError + "\n [Error character position " + eval(charpos+1)+"]"); 
                return false; 
              }//if 
              break;               
           }//numeric 
        case "alphabetic": 
        case "alpha": 
           { 
              var charpos = objValue.value.search("[^A-Za-z]"); 
              if(objValue.value.length > 0 &&  charpos >= 0) 
              { 
                  if(!strError || strError.length ==0) 
                { 
                  strError = objValue.name+": Only alphabetic characters allowed "; 
                }//if                             
                alert(strError + "\n [Error character position " + eval(charpos+1)+"]"); 
                return false; 
              }//if 
              break; 
           }//alpha 
		case "alnumhyphen":
			{	
              var charpos = objValue.value.search("[^A-Za-z0-9\-_\.]"); 
              if(objValue.value.length > 0 &&  charpos >= 0) 
              { 
                  if(!strError || strError.length ==0) 
                { 
                  strError = objValue.name+": characters allowed are A-Z,a-z,0-9,- and _"; 
                }//if                             
                alert(strError + "\n [Error character position " + eval(charpos+1)+"]"); 
                return false; 
              }//if 			
			break;
			}
        case "email": 
          { 
               if(!validateEmailv2(Trim(objValue.value))) 
               { 
                 if(!strError || strError.length ==0) 
                 { 
                    strError = objValue.name+": Enter a valid Email address "; 
                 }//if                                               
                 alert(strError); 
                 return false; 
               }//if 
           break; 
          }//case email
        case "ssn": 
          { 
               if(!isSSN(objValue.value)) 
               { 
                 strError = objValue.name+": Enter a valid SSN ";                                                        
                 alert(strError); 
                 return false; 
               }//if 
           break; 
          }//case email	
        case "phone": 
          { 
               if(objValue.value.length > 0 && !isPhone(objValue.value)) 
               { 
                 strError = objValue.name+": Enter a valid Phone, like (513)4448888";                                                        
                 alert(strError); 
                 return false; 
               }//if 
           break; 
          }//case phone	
		  
        case "date": 
          { 
               if(!isDate(objValue.value)) 
               { 
                 strError = objValue.name+": Enter a valid date. For example 1/1/2004 ";                                                        
                 alert(strError); 
                 return false; 
               }//if 
           break; 
          }//case email			  
        case "lt": 
        case "lessthan": 
         { 
            if(isNaN(objValue.value)) 
            { 
              alert(objValue.name+": Should be a number "); 
              return false; 
            }//if 
            if(eval(objValue.value) >=  eval(cmdvalue)) 
            { 
              if(!strError || strError.length ==0) 
              { 
                strError = objValue.name + " : value should be less than "+ cmdvalue; 
              }//if               
              alert(strError); 
              return false;                 
             }//if             
            break; 
         }//case lessthan 
        case "gt": 
        case "greaterthan": 
         { 
            if(isNaN(objValue.value)) 
            { 
              alert(objValue.name+": Should be a number "); 
              return false; 
            }//if 
             if(eval(objValue.value) <=  eval(cmdvalue)) 
             { 
               if(!strError || strError.length ==0) 
               { 
                 strError = objValue.name + " : value should be greater than "+ cmdvalue; 
               }//if               
               alert(strError); 
               return false;                 
             }//if             
            break; 
         }//case greaterthan 
        case "regexp": 
         { 
		 	if(objValue.value.length > 0)
			{
	            if(!objValue.value.match(cmdvalue)) 
	            { 
	              if(!strError || strError.length ==0) 
	              { 
	                strError = objValue.name+": Invalid characters found "; 
	              }//if                                                               
	              alert(strError); 
	              return false;                   
	            }//if 
			}
           break; 
         }//case regexp 
        case "dontselect": 
         { 
            if(objValue.selectedIndex == null) 
            { 
              alert("BUG: dontselect command for non-select Item"); 
              return false; 
            } 
            if(objValue.selectedIndex == eval(cmdvalue)) 
            { 
             if(!strError || strError.length ==0) 
              { 
              strError = objValue.name+": Please Select one option "; 
              }//if                                                               
              alert(strError); 
              return false;                                   
             } 
             break; 
         }//case dontselect 
    }//switch 
    return true; 
}

function isSSN(value) { 
    var re = /^([0-6]\d{2}|7[0-6]\d|77[0-2])([ \-]?)(\d{2})\2(\d{4})$/; 
    if (!re.test(value)) { return false; } 
    var temp = value; 
    if (value.indexOf("-") != -1) { temp = (value.split("-")).join(""); } 
    if (value.indexOf(" ") != -1) { temp = (value.split(" ")).join(""); } 
    if (temp.substring(0, 3) == "000") { return false; } 
    if (temp.substring(3, 5) == "00") { return false; } 
    if (temp.substring(5, 9) == "0000") { return false; } 
    return true; 
}
function isDate(value){
var dateformat = /^\d{1,2}(\-|\/|\.)\d{1,2}\1(\d{2}|\d{4})$/
return dateformat.test(value) //returns true or false depending on userinput
}

function isDateTime(value){
value=Trim(value);
var dateformat = /^\d{1,2}(\-|\/|\.)\d{1,2}\1(\d{2}|\d{4})(\s{1,}[0-1]{0,1}\d:[0-5]{0,1}\d\w*((AM)|(PM)){0,1}){0,1}$/
//var dateformat = /^(?=\d)(?:(?:31(?!.(?:0?[2469]|11))|(?:30|29)(?!.0?2)|29(?=.0?2.(?:(?:(?:1[6-9]|[2-9]\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00)))(?:\x20|$))|(?:2[0-8]|1\d|0?[1-9]))([-./])(?:1[012]|0?[1-9])\1(?:1[6-9]|[2-9]\d)?\d\d(?:(?=\x20\d)\x20|$))?(((0?[1-9]|1[012])(:[0-5]\d){0,2}(\x20[AP]M))|([01]\d|2[0-3])(:[0-5]\d){1,2})?$/
return dateformat.test(value) //returns true or false depending on userinput
}

function isPhone( strValue) {
	var filter  =/(\(?\d{3}\)?)?(\-|\s|\.)?\d{3}(\-|\s|\.)?\d{4}/g;
	// \s is empty \. is dot ( | | ) is optional 
	return filter.test(strValue);
}

function Trim(str) {
	return LTrim(RTrim(str));
	}
function LTrim(str) {
	for (var i=0; str.charAt(i)==" "; i++);
	return str.substring(i,str.length);
	}
function RTrim(str) {
	for (var i=str.length-1; str.charAt(i)==" "; i--);
	return str.substring(0,i+1);
	}