/*
This file retrieved from the JS-Examples archives
http://www.js-examples.com
1000s of free ready to use scripts, tutorials, forums.
Author: Buddhike de Silva - 0
*/

/*function : format_number()  
version: 1.0.0  
This function formats a numeric value passed in to it with specified number of  
decimal values. numeric value will not be rounded.  
pnumber : numeric value to be formatted.  
decimals : number of decimal points desired.  

Author: Buddhike de Silva  
Date: 21-Nov-2002 11:16 AM*/  

/*  
revision: 1.1.0  
Author: M. Cassim Farook  
Date: 21-Nov-2002 10:16 PM  
Notes: No offense buddhike...but i had to rewrite the code  
works for ADT (any dam thing)  
usage: x = format_number(123.999, 2)  
*/  

/*  
revision: 1.2.0  
Authors: Buddhike de Silva  
Date: 22-Nov-2002 12:07 PM  
Notes: Optimized for best performence.  
usage: x = format_number(123.999, 2)  
*/  

/* 
 * Revision: 1.3 
 * Author: Mike Robb (JS-X.com) 
 * Date: May 26, 2003 
 * Notes:  Changed to deal with negative numbers. 
 *         Fixed length of final answer. 
 *         Work-around for javascript internal math problem with rounding negative numbers. 
 */ 

/*
 * Revision 1.4
 * Author: LeAnn Roberts
 * Date: September, 2003
 * Note: Modified the if logic: Math.pow()
 */

/*
 * Revision 1.5
 * Author: Robert Heggdal
 * Date: February, 2004
 * Note: Modified check for negative number by replacing parseInt with parseFloat so that negative numbers between zero and minus one are recognized as such.
 */ 

/*
 * Revision 1.6
 * Author: Naveen
 * Date: February, 2004
 * Note: Rewrote format_number to correct a logic problem.
 */

/*
 * Revision 1.7
 * Author: JS-X.com
 * Date: February, 2004
 * Note: Added wrapper around format_number as negative values were dropped from
 *       the logic.
 */

function format_number(p,d) 
{
  var r;
  //alert(p);
  if(p<0){p=-p;r=format_number2(p,d);r="-"+r;}
  else   {r=format_number2(p,d);}
  return r;
}
function format_number2(pnumber,decimals) 
{
  var strNumber = new String(pnumber);
  // If the right most char is . then add a 0 to that
  //alert(strNumber.length);
  if (strNumber.substring(strNumber.length,1) == '.')
  {
		strNumber += '0';
  }
  //alert(strNumber);
  var arrParts = strNumber.split('.');
  var intWholePart = parseInt(arrParts[0],10);
  var strResult = '';
  if (isNaN(intWholePart))
    intWholePart = '0';
  if(arrParts.length > 1)
  {
    var decDecimalPart = new String(arrParts[1]);
    var i = 0;
    var intZeroCount = 0;
     while ( i < String(arrParts[1]).length )
     {
       if( parseInt(String(arrParts[1]).charAt(i),10) == 0 )
       {
         intZeroCount += 1;
         i += 1;
       }
       else
         break;
    }
    decDecimalPart = parseInt(decDecimalPart,10)/Math.pow(10,parseInt(decDecimalPart.length-decimals-1)); 
    Math.round(decDecimalPart); 
    decDecimalPart = parseInt(decDecimalPart)/10; 
    decDecimalPart = Math.round(decDecimalPart); 

    //If the number was rounded up from 9 to 10, and it was for 1 'decimal' 
    //then we need to add 1 to the 'intWholePart' and set the decDecimalPart to 0. 

    if(decDecimalPart==Math.pow(10, parseInt(decimals)))
    { 
      intWholePart+=1; 
      decDecimalPart="0"; 
    } 
    var stringOfZeros = new String('');
    i=0;
    if( decDecimalPart > 0 )
    {
      while( i < intZeroCount)
      {
        stringOfZeros += '0';
        i += 1;
      }
    }
    decDecimalPart = String(intWholePart) + "." + stringOfZeros + String(decDecimalPart); 
    var dot = decDecimalPart.indexOf('.');
    if(dot == -1)
    {
      decDecimalPart += '.'; 
      dot = decDecimalPart.indexOf('.'); 
    } 
    var l=parseInt(dot)+parseInt(decimals); 
    while(decDecimalPart.length <= l) 
    {
      decDecimalPart += '0'; 
    }
    strResult = decDecimalPart;
  }
  else
  {
    var dot; 
    var decDecimalPart = new String(intWholePart); 

    decDecimalPart += '.'; 
    dot = decDecimalPart.indexOf('.'); 
    var l=parseInt(dot)+parseInt(decimals); 
    while(decDecimalPart.length <= l) 
    {
      decDecimalPart += '0'; 
    }
    strResult = decDecimalPart;
  }
  return strResult;
}

