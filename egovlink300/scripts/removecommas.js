/*function : removeCommas( string ) 

version: 1.0.0  
This function removes commas from a string. Use before doing curency functions.

*/  

function removeCommas( string ) 
	{
		var tstring = "";
		string = '' + string;
		splitstring = string.split(",");
		for(i = 0; i < splitstring.length; i++)
		tstring += splitstring[i];
		return tstring;
	}