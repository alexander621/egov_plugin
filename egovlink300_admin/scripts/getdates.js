function getDates( iSelection, iFieldID ) 
{

	//Determine which set of fields to return the dates back into
	if(iFieldID == "Dates_Purchasing") {
		lcl_from_date = "fromDate_purchasing"
		lcl_to_date   = "toDate_purchasing"
	} else if(iFieldID == "RentalSummaryReport") {
		lcl_from_date = "sc_fromDate"
		lcl_to_date   = "sc_toDate"
	}else{
		lcl_from_date = "from" + iFieldID;
		lcl_to_date   = "to"   + iFieldID;
	}

	//ARRAY OF LAST DAYS OF MONTH
	var MonthDays = new Array(31,28,31,30,31,30,31,31,30,31,30,31);
	var Q1fromDate = '1/31'
	var d = new Date();

	switch(iSelection) {
		case '16':   
			// Today
			// Next 6 lines changed to fix for FireFox. Steve Loar 11/4/2009
			//var tempfromdt =new Date( d.getYear(), d.getMonth(), d.getDate() );
			//var tempthrudt =new Date( d.getYear(), d.getMonth(), d.getDate() );
			var tempfromdt = new Date();
			var tempthrudt = new Date();
			
			fDate = tempfromdt.getMonth()+1 + '/' + tempfromdt.getDate() + '/' + tempfromdt.getFullYear();
			tDate = tempthrudt.getMonth()+1 + '/' + tempthrudt.getDate() + '/' + tempthrudt.getFullYear();
			
			document.getElementById(lcl_from_date).value = fDate;
			document.getElementById(lcl_to_date).value = tDate;

			break;

		case '21':   
			// Today through Next Year
			var tempfromdt = new Date();

			fDate = tempfromdt.getMonth()+1 + '/' + tempfromdt.getDate() + '/' + tempfromdt.getFullYear();
			tDate =  '12' + '/31/' + (d.getFullYear()+1);

			document.getElementById(lcl_from_date).value = fDate;
			document.getElementById(lcl_to_date).value = tDate;
			break;

		case '17':   
			// Yesterday
			// Next 6 lines changed to fix for FireFox. Steve Loar 11/4/2009
			//var tempfromdt =new Date( d.getYear(), d.getMonth(), d.getDate()-1 );
			//var tempthrudt =new Date( d.getYear(), d.getMonth(), d.getDate()-1 );
			var tempfromdt = new Date();
			var tempthrudt = new Date();
			tempfromdt.setDate(d.getDate()-1);
			tempthrudt.setDate(d.getDate()-1);
			
			fDate = tempfromdt.getMonth()+1 + '/' + tempfromdt.getDate() + '/' + tempfromdt.getFullYear();
			tDate = tempthrudt.getMonth()+1 + '/' + tempthrudt.getDate() + '/' + tempthrudt.getFullYear();
			
			document.getElementById(lcl_from_date).value = fDate;
			document.getElementById(lcl_to_date).value = tDate;

			break;

		case '18':   
			// Tomorrow
			// Next 6 lines changed to fix for FireFox. Steve Loar 11/4/2009
			//var tempfromdt =new Date( d.getYear(), d.getMonth(), d.getDate()+1 );
			//var tempthrudt =new Date( d.getYear(), d.getMonth(), d.getDate()+1 );
			var tempfromdt = new Date();
			var tempthrudt = new Date();
			tempfromdt.setDate(d.getDate()+1);
			tempthrudt.setDate(d.getDate()+1);
			
			fDate = tempfromdt.getMonth()+1 + '/' + tempfromdt.getDate() + '/' + tempfromdt.getFullYear();
			tDate = tempthrudt.getMonth()+1 + '/' + tempthrudt.getDate() + '/' + tempthrudt.getFullYear();
			
			document.getElementById(lcl_from_date).value = fDate;
			document.getElementById(lcl_to_date).value = tDate;

			break;

		case '1':  
			// This Month
			fdate = (d.getMonth()+1)  + '/1/' + d.getFullYear();
			tdate = (d.getMonth()+1) + '/' + MonthDays[d.getMonth()] + '/' + d.getFullYear();
//			document.frmPFilter.fromDate.value = fdate;
//			document.frmPFilter.toDate.value = tdate;
			document.getElementById(lcl_from_date).value = fdate;
			document.getElementById(lcl_to_date).value = tdate;
			break;
		case '2': 
			// Last Month
			d = new Date(d.getFullYear(),d.getMonth()-1,1);
			fdate = (d.getMonth()+1)  + '/1/' + d.getFullYear();
			tdate = (d.getMonth()+1) + '/' + MonthDays[d.getMonth()] + '/' + d.getFullYear();

			document.getElementById(lcl_from_date).value = fdate;
			document.getElementById(lcl_to_date).value = tdate;
			break;

		case '3':   
			//This Quarter
			var iQuarter = getQuarter(d.getMonth()+1);
		
			switch(iQuarter)
			{
				case 1:
					fDate =  '1' + '/1/' + d.getFullYear();
					tDate =  '3' + '/31/' + d.getFullYear();
					break;
				case 2:
					fDate =  '4' + '/1/' + d.getFullYear();
					tDate =  '6' + '/30/' + d.getFullYear();
					break;
				case 3:
					fDate =  '7' + '/1/' + d.getFullYear();
					tDate =  '9' + '/30/' + d.getFullYear();
					break;
				case 4:
					fDate =  '10' + '/1/' + d.getFullYear();
					tDate =  '12' + '/31/' + d.getFullYear();
					break;
			}

			document.getElementById(lcl_from_date).value = fDate;
			document.getElementById(lcl_to_date).value = tDate;
			break;

		case '4':   
			// Last Quarter
			var iQuarter = getQuarter(d.getMonth()+1);

			if(iQuarter == 1) {
				fDate =  '10' + '/1/' + (d.getFullYear()-1);
				tDate =  '12' + '/31/' + (d.getFullYear()-1);
			}
			else {
		
				switch(iQuarter-1)
				{
					case 1:
						fDate =  '1' + '/1/' + d.getFullYear();
						tDate =  '3' + '/31/' + d.getFullYear();
						break;
					case 2:
						fDate =  '4' + '/1/' + d.getFullYear();
						tDate =  '6' + '/30/' + d.getFullYear();
						break;
					case 3:
						fDate =  '7' + '/1/' + d.getFullYear();
						tDate =  '9' + '/30/' + d.getFullYear();
						break;
				}
			}
	
			document.getElementById(lcl_from_date).value = fDate;
			document.getElementById(lcl_to_date).value = tDate;
			break;

		case '5':  
			// Last Year
			fDate =  '1' + '/1/' + (d.getFullYear()-1);
			tDate =  '12' + '/31/' + (d.getFullYear()-1);

			document.getElementById(lcl_from_date).value = fDate;
			document.getElementById(lcl_to_date).value = tDate;
			break;

		case '19':   
			// This Year
			fDate =  '1' + '/1/' + (d.getFullYear());
			tDate =  '12' + '/31/' + (d.getFullYear());

			document.getElementById(lcl_from_date).value = fDate;
			document.getElementById(lcl_to_date).value = tDate;
			break;

		case '20':   
			// Next Year
			fDate =  '1' + '/1/' + (d.getFullYear()+1);
			tDate =  '12' + '/31/' + (d.getFullYear()+1);

			document.getElementById(lcl_from_date).value = fDate;
			document.getElementById(lcl_to_date).value = tDate;
			break;

		case '6':   
			fDate =  '1' + '/1/' + (d.getFullYear());
			tDate =  (d.getMonth()+1) + '/' + d.getDate() + '/' + (d.getFullYear());

			document.getElementById(lcl_from_date).value = fDate;
			document.getElementById(lcl_to_date).value = tDate;
			break;

		case '7':   
			fDate =  '1/1/1900';
			tDate =  (d.getMonth()+1) + '/' + d.getDate() + '/' + (d.getFullYear());

			document.getElementById(lcl_from_date).value = fDate;
			document.getElementById(lcl_to_date).value = tDate;
			break;

		case '8':   

			break;

		case '11':   
			// THIS WEEK
			// Next 6 lines changed to fix for FireFox. Steve Loar 11/4/2009
			//var tempfromdt=new Date(d.getYear(),d.getMonth(),d.getDate()-d.getDay());
			//var tempthrudt=new Date(d.getYear(),d.getMonth(),d.getDate()+ (6-d.getDay()));
			var tempfromdt = new Date();
			var tempthrudt = new Date();
			tempfromdt.setDate(d.getDate() - d.getDay());
			tempthrudt.setDate(d.getDate() + (6 - d.getDay()));
			
			fDate = (tempfromdt.getMonth()+1) + '/' + tempfromdt.getDate() + '/' + tempfromdt.getFullYear();
			tDate = (tempthrudt.getMonth()+1) + '/' + tempthrudt.getDate() + '/' + tempthrudt.getFullYear();
			
			document.getElementById(lcl_from_date).value = fDate;
			document.getElementById(lcl_to_date).value = tDate;

			break;
 
		case '12':   
			// LAST WEEK
			// Next 6 lines changed to fix for FireFox. Steve Loar 11/4/2009
			//var tempfromdt=new Date(d.getYear(),d.getMonth(),d.getDate()-(d.getDay()+7));
			//var tempthrudt=new Date(d.getYear(),d.getMonth(),d.getDate()+ (6-d.getDay()-7));
			var tempfromdt = new Date();
			var tempthrudt = new Date();
			tempfromdt.setDate(d.getDate() - (d.getDay() + 7));
			tempthrudt.setDate(d.getDate() + ((6 - d.getDay()) - 7));
			
			fDate = (tempfromdt.getMonth()+1) + '/' + tempfromdt.getDate() + '/' + tempfromdt.getFullYear();
			tDate = (tempthrudt.getMonth()+1) + '/' + tempthrudt.getDate() + '/' + tempthrudt.getFullYear();
			
			document.getElementById(lcl_from_date).value = fDate;
			document.getElementById(lcl_to_date).value = tDate;

			break;

		case '13':   
			// NEXT MONTH
			// Next 5 lines changed to fix for FireFox. Steve Loar 11/4/2009
			//var tempfromdt=new Date(d.getYear(),d.getMonth()+1,1);
			//var tempthrudt=new Date(d.getYear(),d.getMonth()+1,1);
			d = new Date(d.getFullYear(),(d.getMonth()+1),1);
			fDate = (d.getMonth()+1) + '/1/' + d.getFullYear();
			tDate = (d.getMonth()+1) + '/' +  MonthDays[d.getMonth()] + '/' + d.getFullYear();

			document.getElementById(lcl_from_date).value = fDate;
			document.getElementById(lcl_to_date).value = tDate;


			break;

		case '14':   
			// Next WEEK
			//var tempfromdt =new Date( d.getYear(),d.getMonth(),d.getDate()+ (7-d.getDay()) );
			//var tempthrudt =new Date(d.getYear(),d.getMonth(),d.getDate()+ (13-d.getDay()));
			var tempfromdt = new Date();
			var tempthrudt = new Date();
			tempfromdt.setDate(d.getDate() + (7 - d.getDay()));
			tempthrudt.setDate(d.getDate() + (13 - d.getDay()));
			
			fDate = (tempfromdt.getMonth()+1) + '/' + tempfromdt.getDate() + '/' + tempfromdt.getFullYear();
			tDate = (tempthrudt.getMonth()+1) + '/' + tempthrudt.getDate() + '/' + tempthrudt.getFullYear();
			
			document.getElementById(lcl_from_date).value = fDate;
			document.getElementById(lcl_to_date).value = tDate;

			break;

		case '15':   
			// Next Quarter
			var iQuarter = getQuarter(d.getMonth()+1);

			if(iQuarter == 4) {
				fDate =  '1' + '/1/' + (d.getFullYear()+1);
				tDate =  '3' + '/31/' + (d.getFullYear()+1);
			}
			else {
		
				switch(iQuarter+1)
				{
					case 1:
						fDate =  '1' + '/1/' + d.getFullYear();
						tDate =  '3' + '/31/' + d.getFullYear();
						break;
					case 2:
						fDate =  '4' + '/1/' + d.getFullYear();
						tDate =  '6' + '/30/' + d.getFullYear();
						break;
					case 3:
						fDate =  '7' + '/1/' + d.getFullYear();
						tDate =  '9' + '/30/' + d.getFullYear();
						break;
					case 4:
						fDate =  '10' + '/1/' + d.getFullYear();
						tDate =  '12' + '/31/' + d.getFullYear();
						break;
				}
			}
	
			document.getElementById(lcl_from_date).value = fDate;
			document.getElementById(lcl_to_date).value = tDate;
			break;

		default:
			document.getElementById(lcl_from_date).value = '';
			document.getElementById(lcl_to_date).value = '';
			break;
	}
}


function getQuarter(iMonth) {
	var iQuarter = 0;

	switch(iMonth)
	{
		case 1:
			iQuarter = 1;
			break;
		case 2:
			iQuarter = 1;
			break;
		case 3:
			iQuarter = 1;
			break;
		case 4:
			iQuarter = 2;
			break;
		case 5:
			iQuarter = 2;
			break;
		case 6:
			iQuarter = 2;
			break;
		case 7:
			iQuarter = 3;
			break;
		case 8:
			iQuarter = 3;
			break;
		case 9:
			iQuarter = 3;
			break;
		case 10:
			iQuarter = 4;
			break;
		case 11:
			iQuarter = 4;
			break;
		case 12:
			iQuarter = 4;
			break;
		case 13:
			iQuarter = 4;
			break;
	}
	return iQuarter;
}


function doDate(returnfield, num) 
{
	w = (screen.width - 350)/2;
	h = (screen.height - 350)/2;
	eval('DatePickerWin=window.open("../qrtcalendarpicker.asp?r=" + returnfield + "&n=" + num, "_calendar", "width=350,height=250,toolbar=0,status=yes,scrollbars=0,menubar=0,left=' + w + ',top=' + h + '")');
}

