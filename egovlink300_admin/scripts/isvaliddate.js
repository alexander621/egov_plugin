/*function : isValidDate( string ) 

version: 1.0.0  
This function validates a date as a real date in the format mm/dd/yyyy.  

*/  

	function isValidDate( strValue ) 
	{
		var daterege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
		var dateOk = daterege.test(strValue);

		if (! dateOk )
		{
			return false;
		}
		else
		{
			//var strSeparator = strValue.substring(2,3) 
			var arrayDate = strValue.split('/'); 
			if (arrayDate[0].length == 1)
			{
				arrayDate[0] = '0' + arrayDate[0];
			}
			//create a lookup for months not equal to Feb.
			var arrayLookup = { '01' : 31,'03' : 31, 
					'04' : 30,'05' : 31,
					'06' : 30,'07' : 31,
					'08' : 31,'09' : 30,
					'10' : 31,'11' : 30,'12' : 31}
			var intDay = parseInt(arrayDate[1],10); 

			//check if month value and day value agree
			if(arrayLookup[arrayDate[0]] != null)
			{
				if(intDay <= arrayLookup[arrayDate[0]] && intDay != 0)
					return true; //found in lookup table, good date
			}

			//check for February (bugfix 20050322)
			//bugfix  for parseInt kevin
			//bugfix  biss year  O.Jp Voutat
			var intMonth = parseInt(arrayDate[0],10);
			if (intMonth == 2) 
			{ 
				var intYear = parseInt(arrayDate[2]);
				if (intDay > 0 && intDay < 29) 
				{
					return true;
				}
				else if (intDay == 29) 
				{
					if ((intYear % 4 == 0) && (intYear % 100 != 0) || (intYear % 400 == 0)) 
					{
						// year div by 4 and ((not div by 100) or div by 400) ->ok
						return true;
					}   
				}
			}
		}  
		return false; //any other values, bad date
	}
