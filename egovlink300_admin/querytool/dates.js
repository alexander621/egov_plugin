function getDates(iSelection) 
{

// ARRAY OF LAST DAYS OF MONTH
var MonthDays = new Array(31,28,31,30,31,30,31,31,30,31,30,31);
var Q1fromDate = '1/31'
var d=new Date();
        switch(iSelection)
        {
        case '1':   
			fdate = (d.getMonth()+1)  + '/1/' + d.getFullYear();
			tdate = (d.getMonth()+1) + '/' + MonthDays[d.getMonth()] + '/' + d.getFullYear();
			document.frmdataselection.fromdate.value = fdate;
			document.frmdataselection.fromdate.disabled = true;
			document.frmdataselection.thrudate.value = tdate;
			document.frmdataselection.thrudate.disabled = true;
			document.getElementById("date1").disabled=true;
			document.getElementById("date2").disabled=true;
			break
	case '2':   
			d = new Date(d.getFullYear(),d.getMonth()-1,1);
			fdate = (d.getMonth()+1)  + '/1/' + d.getFullYear();
			tdate = (d.getMonth()+1) + '/' + MonthDays[d.getMonth()] + '/' + d.getFullYear();
			document.frmdataselection.fromdate.value = fdate;
			document.frmdataselection.fromdate.disabled = true;
			document.frmdataselection.thrudate.value = tdate;
			document.frmdataselection.thrudate.disabled = true;
			document.getElementById("date1").disabled=true;
			document.getElementById("date2").disabled=true;
			break
	case '3':   
			iQuarter = getQuarter(d.getMonth()+1);
			//alert(iQuarter);
		
			switch(iQuarter)
			{
				case 1:
					fDate =  '1' + '/1/' + d.getFullYear();
					tDate =  '3' + '/31/' + d.getFullYear();
					break
				case 2:
					fDate =  '4' + '/1/' + d.getFullYear();
					tDate =  '6' + '/30/' + d.getFullYear();
					break
				case 3:
					fDate =  '7' + '/1/' + d.getFullYear();
					tDate =  '9' + '/30/' + d.getFullYear();
					break
				case 4:
					fDate =  '10' + '/1/' + d.getFullYear();
					tDate =  '12' + '/31/' + d.getFullYear();
					break
			}

			document.frmdataselection.fromdate.disabled = true;
			document.frmdataselection.thrudate.disabled = true;
			document.getElementById("date1").disabled=true;
			document.getElementById("date2").disabled=true;



			document.frmdataselection.fromdate.value = fDate;
			document.frmdataselection.thrudate.value = tDate;
			break
	case '4':   
			iQuarter = getQuarter(d.getMonth()+1);
			//alert(iQuarter);
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
						break
					case 2:
						fDate =  '4' + '/1/' + d.getFullYear();
						tDate =  '6' + '/30/' + d.getFullYear();
						break
					case 3:
						fDate =  '7' + '/1/' + d.getFullYear();
						tDate =  '9' + '/30/' + d.getFullYear();
						break
				}
			}
	
			document.frmdataselection.fromdate.disabled = true;
			document.frmdataselection.thrudate.disabled = true;
			document.getElementById("date1").disabled=true;
			document.getElementById("date2").disabled=true;



			document.frmdataselection.fromdate.value = fDate;
			document.frmdataselection.thrudate.value = tDate;
			break
	case '5':   
			fDate =  '1' + '/1/' + (d.getFullYear()-1);
			tDate =  '12' + '/31/' + (d.getFullYear()-1);
	
			document.frmdataselection.fromdate.disabled = true;
			document.frmdataselection.thrudate.disabled = true;
			document.getElementById("date1").disabled=true;
			document.getElementById("date2").disabled=true;



			document.frmdataselection.fromdate.value = fDate;
			document.frmdataselection.thrudate.value = tDate;
			break
	case '6':   
			fDate =  '1' + '/1/' + (d.getFullYear());
			tDate =  (d.getMonth()+1) + '/' + d.getDate() + '/' + (d.getFullYear());
	
			document.frmdataselection.fromdate.disabled = true;
			document.frmdataselection.thrudate.disabled = true;
			document.getElementById("date1").disabled=true;
			document.getElementById("date2").disabled=true;



			document.frmdataselection.fromdate.value = fDate;
			document.frmdataselection.thrudate.value = tDate;
			break
	case '7':   
			fDate =  '1/1/1900';
			tDate =  (d.getMonth()+1) + '/' + d.getDate() + '/' + (d.getFullYear());
	
			document.frmdataselection.fromdate.disabled = true;
			document.frmdataselection.thrudate.disabled = true;
			document.getElementById("date1").disabled=true;
			document.getElementById("date2").disabled=true;



			document.frmdataselection.fromdate.value = fDate;
			document.frmdataselection.thrudate.value = tDate;
			break
	case '8':   
			document.frmdataselection.fromdate.disabled=false;
			document.frmdataselection.thrudate.disabled=false;
			document.getElementById("date1").disabled=false;
			document.getElementById("date2").disabled=false;
			break

	case '11':   
			// THIS WEEK
			var tempfromdt=new Date(d.getYear(),d.getMonth(),d.getDate()-d.getDay());
			var tempthrudt=new Date(d.getYear(),d.getMonth(),d.getDate()+ (6-d.getDay()));
			
			fDate = (tempfromdt.getMonth()+1) + '/' + tempfromdt.getDate() + '/' + tempfromdt.getFullYear();
			tDate = (tempthrudt.getMonth()+1) + '/' + tempthrudt.getDate() + '/' + tempthrudt.getFullYear();
			
			document.frmdataselection.fromdate.value = fDate;
			document.frmdataselection.fromdate.disabled = true;
			document.frmdataselection.thrudate.value = tDate;
			document.frmdataselection.thrudate.disabled = true;
			document.getElementById("date1").disabled=true;
			document.getElementById("date2").disabled=true;
			break
 
	case '12':   
			// LAST WEEK
			var tempfromdt=new Date(d.getYear(),d.getMonth(),d.getDate()-(d.getDay()+7));
			var tempthrudt=new Date(d.getYear(),d.getMonth(),d.getDate()+ (6-d.getDay()-7));
			
			fDate = (tempfromdt.getMonth()+1) + '/' + tempfromdt.getDate() + '/' + tempfromdt.getFullYear();
			tDate = (tempthrudt.getMonth()+1) + '/' + tempthrudt.getDate() + '/' + tempthrudt.getFullYear();
			
			document.frmdataselection.fromdate.value = fDate;
			document.frmdataselection.fromdate.disabled = true;
			document.frmdataselection.thrudate.value = tDate;
			document.frmdataselection.thrudate.disabled = true;
			document.getElementById("date1").disabled=true;
			document.getElementById("date2").disabled=true;
			break

	case '13':   
			// NEXT MONTH
			var tempfromdt=new Date(d.getYear(),d.getMonth()+1,1);
			var tempthrudt=new Date(d.getYear(),d.getMonth()+1,1);
			
			fDate = (tempfromdt.getMonth()+1) + '/' + '1/' + tempfromdt.getFullYear();
			tDate = (tempthrudt.getMonth()+1) + '/' +  MonthDays[tempthrudt.getMonth()] + '/' + tempthrudt.getFullYear();

			document.frmdataselection.fromdate.value = fDate;
			document.frmdataselection.fromdate.disabled = true;
			document.frmdataselection.thrudate.value = tDate;
			document.frmdataselection.thrudate.disabled = true;
			document.getElementById("date1").disabled=true;
			document.getElementById("date2").disabled=true;
			break

	default:
		// DEFAULT CASE
			document.frmdataselection.fromdate.value = '';
			document.frmdataselection.fromdate.disabled = true;
			document.frmdataselection.thrudate.value = '';
			document.frmdataselection.thrudate.disabled = true;
			document.getElementById("date1").disabled=true;
			document.getElementById("date2").disabled=true;
			break

        }
}

function getQuarter(iMonth){
	
	var iQuarter = 0;

	switch(iMonth)
	{
		case 1,2,3:
			iQuarter = 1;
			break
		case 4,5,6:
			iQuarter = 2;
			break
		case 7,8,9:
			iQuarter = 3;
			break
		case 10,11,12:
			iQuarter = 4;
			break
	}

	return iQuarter;
}

    function doDate(returnfield, num) {
      w = (screen.width - 350)/2;
      h = (screen.height - 350)/2;
      eval('DatePickerWin=window.open("../qrtcalendarpicker.asp?r=" + returnfield + "&n=" + num, "_calendar", "width=350,height=250,toolbar=0,status=yes,scrollbars=0,menubar=0,left=' + w + ',top=' + h + '")');

	}

