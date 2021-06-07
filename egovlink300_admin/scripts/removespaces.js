/*function : removeSpaces( string ) 

version: 1.0.0  
This function removes spaces from a string.  

*/  

function removeSpaces( string ) 
	{
		var tstring = "";
		string = '' + string;
		splitstring = string.split(" ");
		for(i = 0; i < splitstring.length; i++)
		tstring += splitstring[i];
		return tstring;
	}
