<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: class_signup.asp
' AUTHOR: Steve Loar
' CREATED: 03/16/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This handles the signup process for classes and events.
'
' MODIFICATION HISTORY
' 1.0 03/16/06  Steve Loar  - Initial version
' 1.1	05/10/06  Steve Loar  - Added citizen search
' 1.2	10/17/06	 Steve Loar  - Security, Header and nav changed
' 2.0	01/22/07	 Steve Loar  - New Family structure applied
' 2.1	02/19/08	 Steve Loar	 - Changes for Early Registration
' 2.2 05/28/08  David Boyer - Added Override Discount
' 2.3 01/07/08  David Boyer - Added "DisplayRosterPublic" check for Craig, CO custom registration fields.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim bRegistrationBlocked, iAvailability

'If they are coming directly to this page without a selected class, take them to the roster list so they can pick one.
 if request("classid") = "" then
   	response.redirect "roster_list.asp"
 end if

 'Session("eGovUserId") = ""

 bRegistrationBlocked = False 
 iAvailability        = 1
 sLevel               = "../"  'Override of value from common.asp
 lcl_setupCoachFields = "N"

'Check the page availability and user access rights in one call
 PageDisplayCheck "registration", sLevel	 'In common.asp

 iItemTypeId = GetItemTypeId( "recreation activity" )  'This is what kind of thing they are buying - in class_global_functions.asp

 session("RedirectPage") = "../classes/class_signup.asp?classid=" & request("classid") & "&timeid=" & request("timeid")
 session("RedirectLang") = "Return to Class/Events Signup"

 Dim sUserType, iUserid, sResidentDesc, iMemberCount, iTimeId, bMultiWeeks, sSearchName, sResults
 sUserType    = "P"
 bMultiWeeks  = False 
 iMemberCount = 0

 'response.write "Session(""eGovUserId"") = " & Session("eGovUserId") &"<br />"
 'response.write "request(""egovuserid"") = " & request("egovuserid") & "<br />"
 if request("egovuserid") <> "" then
	   iUserId = request("egovuserid")
 else
	   if session("eGovUserId") <> "" then
     		if CLng(session("eGovUserId")) <> CLng(0) then
       			iUserid = Session("eGovUserId")
     		else
       			iUserId = GetFirstUserId()  'In class_global_functions.asp
     		end if
   	else
     		iUserId = GetFirstUserId()  'In class_global_functions.asp
   	end if
 end if

 session("eGovUserId") = iUserId

'response.write iUserId & "<br />"
'response.write "Session(""eGovUserId"") = " & Session("eGovUserId") &"<br />"

'First find out what resident type they are
 sUserType = GetUserResidentType(iUserid)

'If they are not one of these (R, N), we have to figure which they are
 if sUserType <> "R" AND sUserType <> "N" then
  	'This leaves E and B - See if they are a resident, also
   	sUserType = GetResidentTypeByAddress(iUserid, Session("OrgID"))
 end if

 sResidentDesc = GetResidentTypeDesc(sUserType)

'See if a timeid was passed
 if request("timeid") <> "" then
   	iTimeId = request("timeid")
 else
   	iTimeId = 0
 end if

'Get the availability of the selected time
 iAvailability = GetActivityAvailability( iTimeId )		' In class_global_functions.asp

'See if a search term was passed
 if request("searchname") <> "" then
	   sSearchName = request("searchname")
 else
	   sSearchName = ""
 end if

 if request("results") <> "" then
   	sResults = request("results")
 else
   	sResults = ""
 end if

 if request("searchstart") <> "" then
   	sSearchStart = request("searchstart")
 else
   	sSearchStart = -1
 end if

'Check for org features
 lcl_orghasfeature_residency_verification      = orghasfeature("residency verification")
 lcl_orghasfeature_custom_registration_craigco = orghasfeature("custom_registration_craigco")
%>
<html>
<head>
 <title>E-Gov Administration Console {Class/Event Signup}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

 <script language="javascript" src="../scripts/ajaxLib.js"></script>
 <script language="javascript" src="../scripts/formatnumber.js"></script>
 <script language="javascript" src="../scripts/removespaces.js"></script>
 <script language="javascript" src="../scripts/removecommas.js"></script>
 <script language="javascript" src="../scripts/setfocus.js"></script>
 <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

 <script language="javascript">
<!--

	function doCalendar(sField) 
	{
      var w = (screen.width - 350)/2;
      var h = (screen.height - 350)/2;
      eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=PurchaseForm", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
    }

	function SearchCitizens( iSearchStart )
	{
		var optiontext;
		var optionchanged;
		//alert(document.BuyerForm.searchname.value);
		var searchtext = document.BuyerForm.searchname.value;
		var searchchanged = searchtext.toLowerCase();

		iSearchStart = parseInt(iSearchStart) + 1;
		
		for (x=iSearchStart; x < document.BuyerForm.egovuserid.length ; x++)
		{
			optiontext = document.BuyerForm.egovuserid.options[x].text;
			optionchanged = optiontext.toLowerCase();
			if (optionchanged.indexOf(searchchanged) != -1)
			{
				document.BuyerForm.egovuserid.selectedIndex = x;
				document.BuyerForm.results.value = 'Possible Match Found.';
				document.getElementById('searchresults').innerHTML = 'Possible Match Found.';
				document.BuyerForm.searchstart.value = x;
				document.BuyerForm.submit();
				return;
			}
		}
		document.BuyerForm.results.value = 'No Match Found.';
		document.getElementById('searchresults').innerHTML = 'No Match Found.';
		document.BuyerForm.searchstart.value = -1;
	}

	function ClearSearch()
	{
		document.BuyerForm.searchstart.value = -1;
	}

	function UserPick()
	{
		document.BuyerForm.searchname.value = '';
		document.BuyerForm.results.value = '';
		document.getElementById('searchresults').innerHTML = '';
		document.BuyerForm.searchstart.value = -1;
		document.BuyerForm.submit();
	}

	function UpdateFamily(iUserId)
	{
		//location.href='../dirs/family_members.asp?userid=' + iUserId;
		location.href='../dirs/family_list.asp?userid=' + iUserId;
	}

	function EditUser(iUserId)
	{
		location.href='../dirs/update_citizen.asp?userid=' + iUserId;
	}

	function NewUser()
	{
		location.href='../dirs/register_citizen.asp';
	}

	function AutoSelect(iTimeId)
	{
		// IF they start typing in a special price then check the Other Price radio
		var radioLength = document.PurchaseForm.pricetypeid.length;
		if(radioLength == undefined) {
			return;
		}
		var i = radioLength - 1;
		document.PurchaseForm.pricetypeid[i].checked = true;
	}

	function ValidateForm()	{
		var iPriceTypeCount  = 0;
  var lcl_return_false = "N";
  var lcl_focus        = "";

		// Check that a price is picked and that the amounts are formatted correctly if they are buying.
		if (document.PurchaseForm.buyorwait[0].checked) { // Buy is 0 Wait is 1
      //alert(document.PurchaseForm.minpricetypeid.value);
      //alert(document.PurchaseForm.maxpricetypeid.value);
		   	for (var p = parseInt(document.PurchaseForm.minpricetypeid.value); p <= parseInt(document.PurchaseForm.maxpricetypeid.value); p++) {
       				// Does it exist
         		if (document.getElementById("pricetypeid" + p)) {
          					// Is is checked
          					if(document.getElementById("pricetypeid" + p).checked) {
            						iPriceTypeCount += 1;

      												// Remove any extra spaces
      												document.getElementById("amount" + p).value = removeSpaces(document.getElementById("amount" + p).value);

      												// Remove commas that would cause problems in validation
      												document.getElementById("amount" + p).value = removeCommas(document.getElementById("amount" + p).value);

      												// Is the price formated correctly and not blank
      												rege = /^\d+\.\d{2}$/;
      												Ok = rege.test(document.getElementById("amount" + p).value);
      												if (! Ok) {
      						          lcl_focus = "amount" + p;
      						       		 inlineMsg(document.getElementById("amount" + p).id,'<strong>Required Field Missing: </strong>Selected prices cannot be blank and must be in currency format.',10,'amount' + p);
     						           lcl_return_false = "Y";
            						}

            						// Check if there is a dropindate entered and is in correct format
      												if (document.getElementById("dropindate" + p)) {
      						   							rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
      						   							Ok = rege.test(document.getElementById("dropindate" + p).value);
      						   							if (! Ok) {
          						          lcl_focus = "pricetypeid" + p;
      			    			       		 inlineMsg(document.getElementById("dropindate" + p).id,'<strong>Invalid Value: </strong>The Drop In date should be in the format of MM/DD/YYYY.',10,'dropindate' + p);
     						               lcl_return_false = "Y";
               							}
            						}
					          }
				       }
			   }

   			// Make sure that at least one thing was checked.
   			if (iPriceTypeCount == 0) {
          lcl_focus = "pricetypeid" + parseInt(document.PurchaseForm.minpricetypeid.value);
      			 inlineMsg(document.getElementById("pricetypeid" + parseInt(document.PurchaseForm.minpricetypeid.value)).id,'<strong>Required Field Missing: </strong>Please select at least one price.',10,'pricetypeid' + parseInt(document.PurchaseForm.minpricetypeid.value));
     					lcl_return_false = "Y";
   			}
		}

<% if lcl_orghasfeature_custom_registration_craigco then %>
  //Validate Team Registration fields
  if(document.getElementById("displayrosterpublic").value=="True") {

     //Check to see if a "coach type" has been selected.
     //If so then Full Name and at least one of the phone numbers and/or email are required.
     if(document.getElementById("rostercoachtype").value != "") {

        //Build the daytime phone
        lcl_dayphone = document.getElementById("skip_volcoachday_areacode").value;
        lcl_dayphone = lcl_dayphone + document.getElementById("skip_volcoachday_exchange").value;
        lcl_dayphone = lcl_dayphone + document.getElementById("skip_volcoachday_line").value;

        //Build the cell phone
        lcl_cellphone = document.getElementById("skip_volcoachcell_areacode").value;
        lcl_cellphone = lcl_cellphone + document.getElementById("skip_volcoachcell_exchange").value;
        lcl_cellphone = lcl_cellphone + document.getElementById("skip_volcoachcell_line").value;

        //Atleast one method of contact is required
        if(lcl_dayphone=="" && lcl_cellphone=="" && document.getElementById("rostervolunteercoachemail").value=="") {
           lcl_focus = "skip_volcoachday_areacode";
        		 inlineMsg(document.getElementById("skip_volcoachday_line").id,'<strong>Required Field Missing: </strong>One method of contact must be entered.',10,'skip_volcoachday_line');
           lcl_return_false = "Y";
        }else{

           //Validate the Email
        			if(document.getElementById("rostervolunteercoachemail").value != "" ) {
           			//var rege = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*\.(\w{2}|(com|net|org|edu|mil|gov|biz|us))$/;
					var rege = /.+@.+\..+/i;
           			var Ok = rege.test(document.getElementById("rostervolunteercoachemail").value);
           			if (! Ok) {
                  lcl_focus = "rostervolunteercoachemail";
               		 inlineMsg(document.getElementById("rostervolunteercoachemail").id,'<strong>Invalid Value: </strong>The volunteer coach email must be in a valid format.',10,'rostervolunteercoachemail');
                  lcl_return_false = "Y";
              }
           }

           //Validate the Cell Phone
           if(lcl_cellphone!="") {
              lcl_cell_areacode = document.getElementById("skip_volcoachcell_areacode").value;
              lcl_cell_exchange = document.getElementById("skip_volcoachcell_exchange").value;
              lcl_cell_line     = document.getElementById("skip_volcoachcell_line").value;

              if(lcl_cellphone.length < 10) {
                 lcl_focus = "skip_volcoachcell_areacode";
              		 inlineMsg(document.getElementById("skip_volcoachcell_line").id,'<strong>Invalid Value: </strong>One or more numbers are missing from the Cell Phone.',10,'skip_volcoachcell_line');
                 lcl_return_false = "Y";
              }else{
                 var cellPhone = new Number(lcl_cell_areacode+lcl_cell_exchange+lcl_cell_line);
                 if(cellPhone.toString() == "NaN") {
                    lcl_focus = "skip_volcoachcell_areacode";
                 		 inlineMsg(document.getElementById("skip_volcoachcell_line").id,'<strong>Invalid Value: </strong>Cell Phone must be numeric.',10,'skip_volcoachcell_line');
                    lcl_return_false = "Y";
                 }
              }
           }

           //Validate the Day Phone
           if(lcl_dayphone!="") {
              lcl_day_areacode = document.getElementById("skip_volcoachday_areacode").value;
              lcl_day_exchange = document.getElementById("skip_volcoachday_exchange").value;
              lcl_day_line     = document.getElementById("skip_volcoachday_line").value;

              if(lcl_dayphone.length < 10) {
                 lcl_focus = "skip_volcoachday_areacode";
              		 inlineMsg(document.getElementById("skip_volcoachday_line").id,'<strong>Invalid Value: </strong>One or more numbers are missing from the Day Phone.',10,'skip_volcoachday_line');
                 lcl_return_false = "Y";
              }else{
                 var dayPhone = new Number(lcl_day_areacode+lcl_day_exchange+lcl_day_line);

                 if(dayPhone.toString() == "NaN") {
                    lcl_focus = "skip_volcoachday_areacode";
                 		 inlineMsg(document.getElementById("skip_volcoachday_line").id,'<strong>Invalid Value: </strong>Day Phone must be numeric.',10,'skip_volcoachday_line');
                    lcl_return_false = "Y";
                 }
              }
           }
        }

        //Validate the Full Name
        if(document.getElementById("rostervolunteercoachname").value=="") {
           lcl_focus = "rostervolunteercoachname";
        		 inlineMsg(document.getElementById("rostervolunteercoachname").id,'<strong>Required Field Missing: </strong>Volunteer Coach - Full Name.',10,'rostervolunteercoachname');
           lcl_return_false = "Y";
        }
     }

     //Validate the Grade
     if(document.getElementById("rostergrade").value == "") {
        lcl_focus = "rostergrade";
     		 inlineMsg(document.getElementById("rostergrade").id,'<strong>Required Field Missing: </strong>Grade.',10,'rostergrade');
        lcl_return_false = "Y";
     }else{
        var rosterGrade = new Number(document.getElementById("rostergrade").value);
        if(rosterGrade.toString() == "NaN") {
           lcl_focus = "rostergrade";
        		 inlineMsg(document.getElementById("rostergrade").id,'<strong>Invalid Value: </strong>Grade must be numeric.',10,'rostergrade');
           lcl_return_false = "Y";
        }
     }
  }

<% end if %>

		// Check that the enrollment max available is not being exceeded and if so ask them about doing so

		// If the ticket field exists check that a quantity is entered
		var bexists = eval(document.PurchaseForm["quantity"]);
		if(bexists) {

  			if (document.getElementById("quantity").value == "") {
         lcl_focus = "quantity";
      		 inlineMsg(document.getElementById("quantity").id,'<strong>Required Field Missing: </strong>Ticket quantity',10,'quantity');
         lcl_return_false = "Y";
  			}else{
      			var rege = /^\d+$/;
      			var Ok   = rege.test(document.getElementById("quantity").value);

      			if (! Ok) {
             lcl_focus = "quantity";
          		 inlineMsg(document.getElementById("quantity").id,'<strong>Invalid Value: </strong>The ticket quantity must be a number.',10,'quantity');
             lcl_return_false = "Y";
      			}else{
          			// Check that the quantity is not more than what is available if they are buying
          			if (document.PurchaseForm.buyorwait.value == 'B') {
             				var iTimeId;
         								var iAvail, iQty;

         								iTimeId = getSelectedRadioValue(document.PurchaseForm.timeid)
         								//iTimeId = document.PurchaseForm.timeid.value;
         								//alert(iTimeId);
         								//get the availability for the select time 
       	  							iAvail = Number(eval('document.PurchaseForm.avail' + iTimeId + '.value'));
         								iQty   = Number(eval('document.PurchaseForm.quantity.value'));
         								//check that the ticket qty is not greater than what is available;
         								if (iQty > iAvail) {
                     lcl_focus = "quantity";
         				     		 inlineMsg(document.getElementById("quantity").id,'<strong>Invalid Value: </strong>The ticket quantity cannot be greater than the availability.',10,'quantity');
         				        lcl_return_false = "Y";
       	  			    }
         				}
         }
     }

   		if (parseInt(document.PurchaseForm.availability.value) <= 0) {
  	   			var response = confirm("This activity is full. Do you wist to continue this registration anyway?");
   		  		if ( response == false ) {
             lcl_return_false = "Y";
     				}
   		}else{
  	   			if ((parseInt(document.PurchaseForm.availability.value) - parseInt(document.PurchaseForm.quantity.value)) < 0) {
		         			var response = confirm("The quantity input exceeds the availablity of this activity. Do you wist to continue this registration anyway?");
      			   		if ( response == false ) {
                  lcl_return_false = "Y";
         					}
  	   			}
   		}

     if(lcl_return_false=="Y") {
        if(lcl_focus != "") {
           document.getElementById(lcl_focus).focus();
        }
        return false;
    	}else	{	
      		//alert('Successful');  // For Ticketed events
     		 document.PurchaseForm.submit();
   		}

		}else{
  			if (parseInt(document.PurchaseForm.availability.value) <= 0) {
     				var response = confirm("This activity is full. Do you wist to continue this registration anyway?");
				     if ( response == false ) {
             lcl_return_false = "Y";
     				}
  			}

     if(lcl_return_false=="Y") {
        if(lcl_focus != "") {
           document.getElementById(lcl_focus).focus();
        }
        return false;
    	}else	{

        //there is an error with ShowFamilyMembers that if none exist the validation errors on this ajax call
        //because the family member dropdown list does not display with no family members.

     			// Fire off AJAX check of age restrictions for registrations
		     	doAjax('check_age_restrictions.asp', 'familymemberid=' + document.PurchaseForm.familymemberid.options[document.PurchaseForm.familymemberid.selectedIndex].value + '&classid=' + document.PurchaseForm.classid.value, 'AgeCheckReturn', 'get', '0');
   		}
		}

	}

	function AgeCheckReturn( sResult )
	{
		//alert( sResult );
		if (sResult == "PASSED")
		{
			// Put call here to check for duplicate enrollment.
			doAjax('check_duplicate_enrollment.asp', 'familymemberid=' + document.PurchaseForm.familymemberid.options[document.PurchaseForm.familymemberid.selectedIndex].value + '&timeid=' + document.PurchaseForm.timeid.value, 'DupCheckReturn', 'get', '0');
		}
		else 
		{
			if (confirm("The selected family member does not meet the age requirements of this activity. \nDo you wish to register them anyway?"))
			{
				// Put call here to check for duplicate enrollment.
				doAjax('check_duplicate_enrollment.asp', 'familymemberid=' + document.PurchaseForm.familymemberid.options[document.PurchaseForm.familymemberid.selectedIndex].value + '&timeid=' + document.PurchaseForm.timeid.value, 'DupCheckReturn', 'get', '0');
			}
		}
	}

	function DupCheckReturn( sResult )
	{
		if (sResult == "NOTFOUND")
		{
			document.PurchaseForm.submit();
		}
		else 
		{
			if (confirm("The selected family member has already been registered for this activity. \nDo you wish to register them anyway?"))
			{
				document.PurchaseForm.submit();
			}
		}
	}

	function ValidateWait() {
		// This does not care for pricing, as it is a wait list addition
		// If the ticket field exists check that something is entered
		var bexists = eval(document.PurchaseForm["quantity"]);
  var lcl_return_false = "N";
  var lcl_focus        = "";  
		if(bexists) {
			if (document.PurchaseForm.quantity.value == "") {
       lcl_focus = "quantity";
       inlineMsg(document.getElementById("quantity").id,'<strong>Required Field Missing: </strong>Ticket Quantity.',10,'quantity');
       lcl_return_false = "Y";
			}
			var rege = /^\d$/;
			var Ok = rege.test(document.PurchaseForm.quantity.value);

			if (! Ok) {
       lcl_focus = "quantity";
       inlineMsg(document.getElementById("quantity").id,'<strong>Invalid Value: </strong>The ticket quantity must be a number.',10,'quantity');
       lcl_return_false = "Y";

			}
		}

  if(lcl_return_false=="Y") {
     if(lcl_focus != "") {
        document.getElementById(lcl_focus).focus();
     }
     return false;
 	}else	{	
   		//alert('Successful');
   		document.PurchaseForm.buyorwait.value = 'W'
   		document.PurchaseForm.submit();
		}
	}

	function ViewCart()
	{
		location.href='class_cart.asp';
	}

	function ValidatePrice( oPrice )
	{
		var bValid = true;
		var total = 0.00;

		// Remove any extra spaces
		oPrice.value = removeSpaces(oPrice.value);
		//Remove commas that would cause problems in validation
		oPrice.value = removeCommas(oPrice.value);

		// Validate the format of the price
		if (oPrice.value != "")
		{
			var rege = /^\d*\.?\d{0,2}$/
			var Ok = rege.exec(oPrice.value);
			if ( Ok )
			{
				oPrice.value = format_number(Number(oPrice.value),2);
			}
			else 
			{
				oPrice.value = format_number(0,2);
				bValid = false;
			}
		}

		// Calculate a new total price
		if (document.PurchaseForm.pricetypeid.length)   // If there is more than one price checkbox
		{
			var checklength = document.PurchaseForm.pricetypeid.length;
			var i = checklength - 1;

			for (l = 0; l <= i; l++)
			{
				if (document.PurchaseForm.pricetypeid[l].checked)
				{ 
					//total += Number(document.frmStatus.pricetypeid[l].value);
					total += Number(eval('document.PurchaseForm.amount' + document.PurchaseForm.pricetypeid[l].value + '.value'));
				}
			}
		}
		else   // There is only one price checkbox
		{
			if (document.PurchaseForm.pricetypeid.checked)
			{
				total += Number(eval('document.PurchaseForm.amount' + document.PurchaseForm.pricetypeid.value + '.value'));
			}
		}

		document.PurchaseForm.totalprice.value = total;
		document.getElementById("displaytotalprice").innerHTML = format_number(total,2);

		if ( bValid == false ) {
      document.getElementById(oPrice.id).focus();
      inlineMsg(oPrice.id,'<strong>Invalid Value: </strong>Prices should numbers in currency format.',10,oPrice.id);
      return false
		}
		return true;
	}

	function UpdatePriceTotal( iPrice, bChecked )
	{
		var total = 0.00;

		if (iPrice != "")
		{
			total = Number(document.PurchaseForm.totalprice.value);
			if (bChecked)
			{
				total += Number(iPrice);
			}
			else
			{
				total -= Number(iPrice);
			}
			document.PurchaseForm.totalprice.value = total;
			document.getElementById("displaytotalprice").innerHTML = format_number(total,2);
		}
	}

function setupCoachFields() {
  //Check to see if a value has been selected in the "I would like to" volunteer coach field.
  //If one has been selected then enable the other volunteer coach fields.
  //If one has not then disable them.
  lcl_type = document.getElementById("rosterCoachType").value;

  if(lcl_type!="") {
     document.getElementById("volunteerCoachInfo").style.visibility="visible";
  }else{
     document.getElementById("volunteerCoachInfo").style.visibility="hidden";
  }
}

	var isNN = (navigator.appName.indexOf("Netscape")!=-1);

	function autoTab(input,len, e) {
		var keyCode = (isNN) ? e.which : e.keyCode; 
		var filter  = (isNN) ? [0,8,9] : [0,8,9,16,17,18,37,38,39,40,46];

		if(input.value.length >= len && !containsElement(filter,keyCode)) {
  			input.value = input.value.slice(0, len);
		var addNdx = 1;

		while(input.form[(getIndex(input)+addNdx) % input.form.length].type == "hidden") {
			addNdx++;
			//alert(input.form[(getIndex(input)+addNdx) % input.form.length].type);
		}

		input.form[(getIndex(input)+addNdx) % input.form.length].focus();
	}

	function containsElement(arr, ele) {
		var found = false, index = 0;

		while(!found && index < arr.length)
			if(arr[index] == ele)
  				found = true;
			else
		  		index++;
		return found;
	}

	function getIndex(input)	{
		var index = -1, i = 0, found = false;

		while (i < input.form.length && index == -1)
			if (input.form[i] == input)index = i;
			else i++;
				return index;
	}
		return true;
	}
//-->
</script>
</head>
<body>
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->

<div id="content">
	<div id="centercontent">

<%
'Check the cart.  If items exist then display the "View Cart" button
 if CartHasItems() then
    response.write "<div id=""topbuttons"">" & vbcrlf
    response.write "<input type=""button"" name=""viewcart"" id=""viewcart"" class=""button"" value=""View Cart"" onclick=""ViewCart();"" />" & vbcrlf
    response.write "</div>" & vbcrlf
 end if

'Display "Back" and "Return to Class/Event List" buttons
 response.write "<p>" & vbcrlf
 response.write "<input type=""button"" name=""backbtn"" id=""backbtn"" value=""Back"" onclick=""location.href='class_offerings.asp?classid=" & request("classid") & "'"" />" & vbcrlf
 response.write "<input type=""button"" name=""returnToList"" id=""returnToList"" value=""Return to Class/Event List"" onclick=""location.href='roster_list.asp'"" />" & vbcrlf
 response.write "</p>" & vbcrlf
%>
	<!--<a href="javascript:history.go(-1)"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>-->
<!--	<a href="class_offerings.asp?classid=<%'request("classid")%>"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%'langBackToStart%></a> &nbsp; 
	<a href="roster_list.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Return To Class/Event List</a><br /><br /> -->
	
	<% 
	'Get the class information to display
	sSQL = "SELECT C.classname, isnull(C.startdate,'') as startdate, C.classdescription, isnull(C.enddate,'') as enddate, "
	sSQL = sSQL & " O.optionid, O.optionname, O.optiondescription, O.canpurchase, O.optiontype, C.isparent, C.classtypeid, "
	sSQL = sSQL & " isnull(C.membershipid,0) as membershipid, L.name as locationname, L.address1, isnull(C.minage,0) as minage, "
	sSQL = sSQL & " isnull(C.maxage,99) as maxage, isnull(C.pricediscountid,0) as pricediscountid, displayrosterpublic "
	sSQL = sSQL & " FROM egov_class C, egov_registration_option O, egov_class_location L "
	sSQL = sSQL & " WHERE classid = " & request("classid")
	sSQL = sSQL & " AND C.optionid = O.optionid "
	sSQL = sSQL & " AND C.locationid = L.locationid"

	set oClass = Server.CreateObject("ADODB.Recordset")
	oClass.Open sSQL, Application("DSN"), 3, 1
	
	if not oClass.eof then
  		response.write "<h3>" & oClass("classname") & " &nbsp; ( " & GetActivityNo( iTimeId ) & " )</h3>" & vbcrlf
  		response.write "<fieldset><legend><strong> Details </strong></legend>" & vbcrlf

 		'Date
  		response.write "<div><p>" & vbcrlf

  		if oClass("startdate") <> "" then
    			response.write MonthName(Month(oClass("startdate"))) & " " & Day(oClass("startdate")) & ", " & Year(oClass("startdate")) & vbcrlf
    end if

 		'handle enddate
  		if oClass("enddate") <> "" AND oClass("enddate") <> oClass("startdate") then
    			response.write " &ndash; " & MonthName(Month(oClass("enddate"))) & " " & Day(oClass("enddate")) & ", " & Year(oClass("startdate")) & vbcrlf

    			if DateDiff("d", oClass("startdate"), oClass("enddate")) > 7 then
      				bMultiWeeks = true
    			end if
    end if

		 'Days of the week
  		response.write "</p>" & vbcrlf

 		'Tell if registration, ticket, or free
  		response.write "<p><strong>" & oClass("optionname") & " &ndash; " & oClass("optiondescription") & "</strong></p>" & vbcrlf

 		'Tell about age restrictions
  		response.write "<p><strong>Age Restrictions:</strong>" & vbcrlf
  		if CDbl(oClass("minage")) = CDbl(0.0) AND CDbl(oClass("maxage")) = CDbl(99.0) then
    			response.write "<br />&nbsp;&nbsp;&nbsp;None"
  		else
		    	if CDbl(oClass("minage")) <> CDbl(0.0) then
      				response.write "<br />&nbsp;&nbsp;&nbsp;Minimum: " & oClass("minage") & " years of age"
    			end if

    			if CDbl(oClass("maxage")) <> CDbl(99.0) then
				      response.write "<br />&nbsp;&nbsp;&nbsp;Maximum: " & oClass("maxage") & " years of age"
    			end if
  		end if

  		response.write "</p>" & vbcrlf

		' Location 
		response.write vbcrlf & "<p><strong>Location:</strong><br />"
		response.write "&nbsp;&nbsp;&nbsp;" & oClass("locationname") & "<br />&nbsp;&nbsp;&nbsp;" & oClass("address1") & "</p>"

		' Display Waiver Links
		response.write "<p><strong>Waivers:</strong>&nbsp; " 
		ShowClassWaiverLinks request("classid") 
		response.write "</p>"

		'response.write vbcrlf & "<p>You are considered a " & sUserType & " - " & sResidentDesc & "</p>"
		'response.write "</div><div id=""rightdetail"">"
		'response.write "<p><strong>Description:</strong><br />" & oClass("classdescription") & "</p>" 
		response.write vbcrlf & "</div></fieldset>"

 	Select Case oClass("optiontype")
			Case "register"		' Handle registration required
				' Show pick of registered users and their detail info.
				ShowRegisteredUsers iUserId, sUserType, sResidentDesc, sSearchName, sResults, sSearchStart, iTimeId

				response.write "<form id=""PurchaseForm"" name=""PurchaseForm"" method=""post"" action=""class_addtocart.asp"">" & vbcrlf
				response.write "  <input type=""hidden"" name=""classid"" value=""" & request("classid") & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""timeid"" value=""" & iTimeId & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""userid"" value=""" & iUserId & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""optionid"" value=""" & oClass("optionid") & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""optiontype"" value=""" & oClass("optiontype") & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""isparent"" value=""" & oClass("isparent") & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""classtypeid"" value=""" & oClass("classtypeid") & """ />" & vbcrlf
				'response.write "  <input type=""hidden"" name=""buyorwait"" value=""B"" />" & vbcrlf
				response.write "  <input type=""hidden"" name=""classname"" value=""" & oClass("classname") & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""itemtypeid"" value=""" & iItemTypeId & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""availability"" value=""" & iAvailability & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""displayrosterpublic"" id=""displayrosterpublic"" value=""" & oClass("displayrosterpublic") & """ />" & vbcrlf

				'Family Member Drop Down --------------------------------------------------
				response.write "<fieldset><legend><strong>Select a Family Member to Register&nbsp;</strong></legend>" & vbcrlf

				bAllMembers = ShowFamilyMembers( iUserid, iMemberCount, oClass("membershipid") )

				response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" class=""button"" style=""width:150px;text-align:center;"" onclick=""UpdateFamily(" & iUserId & ");"" value=""Update Family Members"" />" & vbcrlf
				response.write "</fieldset>" & vbcrlf

   'Craig, CO - Custom Team Roster Registration Fields ------------------------
    if lcl_orghasfeature_custom_registration_craigco AND oClass("displayrosterpublic") then

       lcl_volunteercoach_text = getOrgDisplay(session("orgid"),"class_teamregistration_volunteercoachdesc")

       response.write "<fieldset><legend><strong>Team Registration - Additional Info&nbsp;</strong></legend>" & vbcrlf
       response.write "<p>" & vbcrlf
       response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
       response.write "  <tr>" & vbcrlf
       response.write "      <td>" & vbcrlf
       response.write "          <span style=""color:#ff0000"">*</span>Grade:&nbsp;" & vbcrlf
       response.write "          <input type=""text"" name=""rostergrade"" id=""rostergrade"" size=""3"" maxlength=""2"" onchange=""clearMsg('rostergrade')"" />" & vbcrlf
       response.write "      </td>" & vbcrlf
       response.write "      <td>" & vbcrlf
       response.write "          T-Shirt Size:&nbsp;" & vbcrlf
       response.write "          <select name=""rostershirtsize"" id=""rostershirtsize"" onchange=""clearMsg('rostershirtsize')"">" & vbcrlf
       response.write "            <option value=""Youth - Small (6-8)"">Youth - Small (6-8)</option>" & vbcrlf
       response.write "            <option value=""Youth - Medium (10-12)"">Youth - Medium (10-12)</option>" & vbcrlf
       response.write "            <option value=""Youth - Large (14-16)"">Youth - Large (14-16)</option>" & vbcrlf
       response.write "            <option value=""Adult - Small (34-36)"">Adult - Small (34-36)</option>" & vbcrlf
       response.write "            <option value=""Adult - Medium (38-40)"">Adult - Medium (38-40)</option>" & vbcrlf
       response.write "            <option value=""Adult - Large (40-42)"">Adult - Large (40-42)</option>" & vbcrlf
       response.write "            <option value=""Adult - X-Large (44-46)"">Adult - X-Large (44-46)</option>" & vbcrlf
       response.write "          </select>" & vbcrlf
       response.write "      </td>" & vbcrlf
       response.write "  </tr>" & vbcrlf
       response.write "</table>" & vbcrlf
       response.write "</p>" & vbcrlf

       if lcl_volunteercoach_text <> "" then
          response.write "<div>" & lcl_volunteercoach_text & "</div><br />" & vbcrlf
       end if

       response.write "<div>" & vbcrlf
       response.write "  I would like to:&nbsp;" & vbcrlf
       response.write "  <select name=""rostercoachtype"" id=""rostercoachtype"" onchange=""setupCoachFields();"">" & vbcrlf
       response.write "    <option value=""""></option>" & vbcrlf
       response.write "    <option value=""Head Coach"">Head Coach</option>" & vbcrlf
       response.write "    <option value=""Assistant Coach"">Assistant Coach</option>" & vbcrlf
       response.write "  </select>" & vbcrlf
       response.write "</div>" & vbcrlf
       response.write "<br />" & vbcrlf
       response.write "<div id=""volunteerCoachInfo"">" & vbcrlf
       response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
       response.write "  <tr>" & vbcrlf
       response.write "  </tr>" & vbcrlf
       response.write "  <tr>" & vbcrlf
       response.write "      <td>" & vbcrlf
       response.write "          <span style=""color:#ff0000"">*</span>Full Name:" & vbcrlf
       response.write "      </td>" & vbcrlf
       response.write "      <td width=""85%"">" & vbcrlf
       response.write "          <input type=""text"" name=""rostervolunteercoachname"" id=""rostervolunteercoachname"" size=""40"" maxlength=""100"" onchange=""clearMsg('rostervolunteercoachname');"" />" & vbcrlf
       response.write "      </td>" & vbcrlf
       response.write "  </tr>" & vbcrlf
       response.write "  <tr>" & vbcrlf
       response.write "      <td>" & vbcrlf
       response.write "          Daytime Phone:" & vbcrlf
       response.write "      </td>" & vbcrlf
       response.write "      <td>" & vbcrlf
       'response.write "          <input type=""hidden"" name=""rostervolunteercoachdayphone"" id=""rostervolunteercoachdayphone"" size=""10"" maxlength=""10"" />" & vbcrlf
       response.write "         (<input type=""text"" name=""skip_volcoachday_areacode"" id=""skip_volcoachday_areacode"" size=""3"" maxlength=""3"" onKeyUp=""return autoTab(this, 3, event);"" onchange=""clearMsg('skip_volcoachday_line');"" />)" & vbcrlf
       response.write "          <input type=""text"" name=""skip_volcoachday_exchange"" id=""skip_volcoachday_exchange"" size=""3"" maxlength=""3"" onKeyUp=""return autoTab(this, 3, event);"" onchange=""clearMsg('skip_volcoachday_line');"" />" & vbcrlf
       response.write "          &ndash;" & vbcrlf
       response.write "          <input type=""text"" name=""skip_volcoachday_line"" id=""skip_volcoachday_line"" size=""4"" maxlength=""4"" onKeyUp=""return autoTab(this, 4, event);"" onchange=""clearMsg('skip_volcoachday_line');"" />" & vbcrlf
       response.write "      </td>" & vbcrlf
       response.write "  </tr>" & vbcrlf
       response.write "  <tr>" & vbcrlf
       response.write "      <td>" & vbcrlf
       response.write "          Cell Phone:" & vbcrlf
       response.write "      </td>" & vbcrlf
       response.write "      <td>" & vbcrlf
       'response.write "          <input type=""hidden"" name=""rostervolunteercoachcellphone"" id=""rostervolunteercoachcellphone"" size=""10"" maxlength=""10"" />" & vbcrlf
       response.write "         (<input type=""text"" name=""skip_volcoachcell_areacode"" id=""skip_volcoachcell_areacode"" size=""3"" maxlength=""3"" onKeyUp=""return autoTab(this, 3, event);"" onchange=""clearMsg('skip_volcoachcell_line');"" />)" & vbcrlf
       response.write "          <input type=""text"" name=""skip_volcoachcell_exchange"" id=""skip_volcoachcell_exchange"" size=""3"" maxlength=""3"" onKeyUp=""return autoTab(this, 3, event);"" onchange=""clearMsg('skip_volcoachcell_line');"" />" & vbcrlf
       response.write "          &ndash;" & vbcrlf
       response.write "          <input type=""text"" name=""skip_volcoachcell_line"" id=""skip_volcoachcell_line"" size=""4"" maxlength=""4"" onKeyUp=""return autoTab(this, 4, event);"" onchange=""clearMsg('skip_volcoachcell_line');"" />" & vbcrlf
       response.write "      </td>" & vbcrlf
       response.write "  </tr>" & vbcrlf
       response.write "</table>" & vbcrlf
       response.write "<div>" & vbcrlf
       response.write "  Please list an email address, so you can be contacted for more information:&nbsp;" & vbcrlf
       response.write "  <input type=""text"" name=""rostervolunteercoachemail"" id=""rostervolunteercoachemail"" size=""50"" maxlength=""100"" onchange=""clearMsg('rostervolunteercoachemail');"" />" & vbcrlf
       response.write "</div>" & vbcrlf
       response.write "</div>" & vbcrlf
   				response.write "</fieldset>" & vbcrlf

       lcl_setupCoachFields = "Y"
    else
   				response.write "  <input type=""hidden"" name=""rostergrade"" id=""rostergrade"" value="""" />" & vbcrlf
   				response.write "  <input type=""hidden"" name=""rostershirtsize"" id=""rostershirtsize"" value="""" />" & vbcrlf
   				response.write "  <input type=""hidden"" name=""rostercoachtype"" id=""rostercoachtype"" value="""" />" & vbcrlf
   				response.write "  <input type=""hidden"" name=""rostervolunteercoachname"" id=""rostervolunteercoachname"" value="""" />" & vbcrlf
   				response.write "  <input type=""hidden"" name=""rostervolunteercoachdayphone"" id=""rostervolunteercoachdayphone"" value="""" />" & vbcrlf
   				response.write "  <input type=""hidden"" name=""rostervolunteercoachcellphone"" id=""rostervolunteercoachcellphone"" value="""" />" & vbcrlf
   				response.write "  <input type=""hidden"" name=""rostervolunteercoachemail"" id=""rostervolunteercoachemail"" value="""" />" & vbcrlf
    end if

   'Availability and Pricing --------------------------------------------------
				response.write "<fieldset><legend><strong> Availability and Pricing&nbsp;</strong></legend>" & vbcrlf
			'Form for selecting either ticket qty, or selecting a family member
				response.write "<div>" & vbcrlf

			'Availability
				response.write "<p><strong>Availability:</strong><br />" & vbcrlf
				DisplayClassActivities request("classid"), iTimeId, False  'In class_global_functions.asp
				response.write "</p>" & vbcrlf

				' Time Options
'				response.write vbcrlf & "<p><strong>Time:</strong><br />"
'				ShowTimeOptions request("classid"), Session("OrgID"), oClass("isparent"), oClass("classtypeid"), iTimeId
'				response.write "</p>"

				' Price options
				response.write "<p><strong>Price:</strong><br />"
				ShowPriceOptions request("classid"), Session("OrgID"), sUserType, iMemberCount, oClass("membershipid"), oClass("pricediscountid"), iUserId
				'ShowCostOptions request("classid"), sUserType, Session("OrgID"), bAllMembers, iMemberCount, oClass("isparent"), oClass("classtypeid")
				response.write vbcrlf & "</p>"

				' Purchase or Waitlist 
				response.write "<p><strong>Select:</strong><br />" & vbcrlf
				response.write "<input type=""radio"" name=""buyorwait"" id=""buyorwait"" value=""B"" checked=""checked"" /> Purchase <br />" & vbcrlf
				response.write "<input type=""radio"" name=""buyorwait"" id=""buyorwait"" value=""W"" /> Add to Wait List" & vbcrlf
				response.write "</p>" & vbcrlf

				'ShowPaymentChoices
				response.write "<p>" & vbcrlf
				response.write "<input type=""button"" name=""complete"" class=""button"" style=""width:140px;text-align:center;"" value=""Add To Cart"" onclick=""ValidateForm();"" "

				If bRegistrationBlocked Then 
  					response.write " disabled=""disabled"" "
				End If

				response.write "/>" & vbcrlf
				'response.write vbcrlf & "&nbsp;&nbsp;<strong>OR</strong>"
				'response.write vbcrlf & "&nbsp;&nbsp;<input type=""button"" name=""waitlist"" value=""Add to Wait List"" onclick=""ValidateWait();"" />"
				response.write vbcrlf & "</p>"

				response.write vbcrlf & "</div>"

				' Show the availability
				'response.write vbcrlf & "<div id=""rightprice"">"
				'ShowAvailability request("classid"), oClass("isparent")
				'response.write vbcrlf & "</div>"
				response.write vbcrlf & "</fieldset>"
				response.write vbcrlf & "</form>"
				
			Case "tickets"		' Ticketed Event
				' Show pick of registered users and their detail info.
				ShowRegisteredUsers iUserId, sUserType, sResidentDesc, sSearchName, sResults, sSearchStart, iTimeId

				response.write vbcrlf & "<fieldset><legend><strong> Ticket Availability and Pricing </strong></legend>"
				' Form for selecting either ticket qty, or selecting a family member
				response.write "<div id=""leftprice"">" & vbcrlf
    response.write "<form name=""PurchaseForm"" method=""post"" action=""class_addtocart.asp"">" & vbcrlf
				response.write "  <input type=""hidden"" name=""classid"" value=""" & request("classid") & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""timeid"" value=""" & iTimeId & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""userid"" value=""" & iUserId & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""optionid"" value=""" & oClass("optionid") & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""optiontype"" value=""" & oClass("optiontype") & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""isparent"" value=""" & oClass("isparent") & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""classtypeid"" value=""" & oClass("classtypeid") & """ />" & vbcrlf
				'response.write "  <input type=""hidden"" name=""buyorwait"" value=""B"" />" & vbcrlf
				response.write "  <input type=""hidden"" name=""classname"" value=""" & oClass("classname") & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""itemtypeid"" value=""" & iItemTypeId & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""availability"" value=""" & iAvailability & """ />" & vbcrlf

				' Availability
				response.write vbcrlf & "<p><strong>Availability:</strong><br />"
				DisplayClassActivities request("classid"), iTimeId, False   ' In class_global_functions.asp
				'ShowAvailability request("classid"), oClass("isparent"), oClass("optionid"), iTimeId
				response.write vbcrlf & "</p>"

				' Ticket Quantity
				response.write "<p>No. of Tickets: &nbsp; <input type=""text"" name=""quantity"" id=""quantity"" value=""1"" size=""6"" maxlength=""6"" /></p>" & vbcrlf
				
				' Availability
'				response.write vbcrlf & "<p><strong>Availability:</strong><br />"
'				ShowAvailability request("classid"), oClass("isparent"), oClass("optionid"), iTimeId
'				response.write vbcrlf & "</p>"

				' Time Options
'				response.write vbcrlf & "<p><strong>Time:</strong><br />"
'				ShowTimeOptions request("classid"), Session("OrgID"), oClass("isparent"), oClass("classtypeid"), iTimeId
'				response.write "</p>"

				' Price Options
				response.write vbcrlf & "<p><strong>Price:</strong><br />"
				ShowPriceOptions request("classid"), Session("OrgID"), sUserType, iMemberCount, oClass("membershipid"), oClass("pricediscountid"), iUserId
				'ShowCostOptions request("classid"), sUserType, Session("OrgID"), bAllMembers, iMemberCount, oClass("isparent"), oClass("classtypeid")
				response.write vbcrlf & "</p>"

				' Purchase or Waitlist 
				response.write vbcrlf & "<p><strong>Select:</strong><br />"
				response.write vbcrlf & "<input type=""radio"" name=""buyorwait"" value=""B"" checked=""checked"" /> Purchase <br />"
				response.write vbcrlf & "<input type=""radio"" name=""buyorwait"" value=""W"" /> Add to Wait List"
				response.write vbcrlf & "</p>"

				'ShowPaymentChoices
				response.write vbcrlf & "<p>"
				response.write vbcrlf & "<input type=""button"" class=""button"" name=""complete"" style=""width:140px;text-align:center;"" value=""Add To Cart"" onclick=""ValidateForm();"""
				If bRegistrationBlocked Then 
					response.write " disabled=""disabled"" "
				End If 
				response.write "/>"
				'response.write vbcrlf & "&nbsp;&nbsp;<strong>OR</strong>"
				'response.write vbcrlf & "&nbsp;&nbsp;<input type=""button"" name=""waitlist"" value=""Add to Wait List"" onclick=""ValidateWait();"" />"
				response.write vbcrlf & "</p>"

				response.write vbcrlf & "</div></form>"

				' Show the availability
				'response.write vbcrlf & "<div id=""rightprice"">"
				'ShowAvailability request("classid"), oClass("isparent")
				'response.write vbcrlf & "</div>"
				response.write vbcrlf & "</fieldset>"


			Case "open"		' Open Attendance
				response.write "<p><strong>Ticketing or registration is not required.</strong></p>"
			Case Else 		' Information Only
				response.write "<p><strong>This listing is for information only.</strong>  See related classes/events to register or purchase tickets.</p>"
		End Select 
	Else
		response.write "<p>No information could be found for this class.</p>"
	End If 

	oClass.close
	set oClass = nothing

 response.write "  </div>" & vbcrlf
 response.write "</div>" & vbcrlf

'Check for javascripts
 if lcl_setupCoachFields = "Y" then
    response.write "<script language=""javascript"">" & vbcrlf
    response.write "  setupCoachFields();" & vbcrlf
    response.write "</script>" & vbcrlf
 end if
%>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<!--#Include file="class_global_functions.asp"-->  

<%
'------------------------------------------------------------------------------
Sub ShowRegisteredUsers( iUserId, sUserType, sResidentDesc, sSearchName, sResults, sSearchStart, iTimeId )

	response.write vbcrlf & "<fieldset><legend><strong> Purchaser Information </strong></legend>"
	response.write vbcrlf & "<form name=""BuyerForm"" method=""post"" action=""class_signup.asp"">"
	response.write vbcrlf & "<p>Name Search: <input type=""text"" name=""searchname"" value=""" & sSearchName & """ size=""25"" maxlength=""50"" onchange=""javascript:ClearSearch();"" />"
	response.write vbcrlf & "<input type=""button"" class=""button"" value=""Search"" onclick=""javascript:SearchCitizens(document.BuyerForm.searchstart.value);"" /> &nbsp;&nbsp; <input type=""button"" class=""button"" onclick=""javascript:NewUser();"" value=""New User"" />"
	response.write vbcrlf & "<input type=""hidden"" name=""results"" value="""" />"
	response.write vbcrlf & "<input type=""hidden"" name=""timeid"" value=""" & iTimeId & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""searchstart"" value=""" & sSearchStart & """ />"
	response.write vbcrlf & "<span id=""searchresults"">" & sResults & "</span>"
	response.write vbcrlf & "<br /><div id=""searchtip"">(last name, first name)</div>"
	response.write vbcrlf & "</p>"
	response.write vbcrlf & "<p><input type=""hidden"" name=""classid"" value=""" & request("classid") & """ />"
	response.write vbcrlf & "Select Name: <select name=""egovuserid"" onchange=""javascript:UserPick();"">"
	response.write ShowUserDropDown( iUserId )
	response.write vbcrlf & "</select>"
	response.write vbcrlf & " &nbsp;&nbsp; <input type=""button"" class=""button"" onclick=""javascript:EditUser(" & iUserId & ");"" value=""Edit User Profile"" />"
	response.write vbcrlf & "</p></form>"
	ShowUserInfo iUserId, sUserType, sResidentDesc 
	response.write vbcrlf & "</fieldset>"
End Sub 


'------------------------------------------------------------------------------
' Sub ShowFamilyMembers( iUserid )
'------------------------------------------------------------------------------
Function ShowFamilyMembers( ByVal iUserid, ByRef iMemberCount, ByVal iMembershipId )
	Dim sSql, oFamily, sMember, iCount, iMonths, iAge

	iCount = 0
	iMemberCount = 0 

	sSQL = "SELECT firstname, lastname, familymemberid, relationship, birthdate, userid"
	sSQL = sSQL & " FROM egov_familymembers "
	sSQL = sSQL & " WHERE isdeleted = 0 "
 sSQL = sSQL & " AND belongstouserid = " & iUserid
 sSQL = sSQL & " ORDER BY birthdate ASC"

	set oFamily = Server.CreateObject("ADODB.Recordset")
	oFamily.Open sSQL, Application("DSN"), 3, 1

	if not oFamily.eof then
  		response.write "<select name=""familymemberid"" id=""familymemberid"">" & vbcrlf
  		do while not oFamily.eof
    			if CLng(iMembershipId) > CLng(0) then
      				sMember = DetermineMembership(oFamily("familymemberid"), iUserid, iMembershipId)   ' In class_global_functions.asp
       end if

    			if iCount = 0 then
          lcl_member_selected = " selected=""selected"""
       else
          lcl_member_selected = ""
       end if

    			response.write vbcrlf & "  <option value=""" & oFamily("familymemberid") & """" & lcl_member_selected & ">" & oFamily("firstname") & " " & oFamily("lastname") & " &ndash; " 
			
			If CLng(oFamily("userid")) = CLng(iUserid) Then
				response.write "Head of Household"
			Else 
				response.write oFamily("relationship") 
			End If 

			If UCase(oFamily("relationship")) = "CHILD" Then 
				iAge = GetChildAge(oFamily("birthdate"))
				response.write " &ndash; Age: " & iAge & " yrs"
			Else
				If UCase(oFamily("relationship")) <> "SITTER" Then
					response.write " &ndash; Adult"
				End If 
			End If 

			If CLng(iMembershipId) > CLng(0) Then 
				If sMember = "M" Then
					response.write " &ndash; Member" 
					iMemberCount = iMemberCount + 1
				Else 
					response.write " &ndash; NonMember"
				End If 
			End If 
			response.write "</option>" & vbcrlf
			iCount = iCount + 1
			oFamily.movenext
		Loop 
		response.write "</select>" & vbcrlf
'	Else
'		response.write sSql
	End If 
	
	oFamily.close
	Set oFamily = Nothing

	If iMemberCount = iCount Then 
		ShowFamilyMembers = True 
	Else
		ShowFamilyMembers = False
	End If 

End Function  


'------------------------------------------------------------------------------
' Sub ShowTimeOptions( iClassid, iorgid, bIsParent, iClassTypeId, iTimeId)
'------------------------------------------------------------------------------
Sub ShowTimeOptions( iClassid, iorgid, bIsParent, iClassTypeId, iTimeId )
	Dim sSql, oTime, iCount

	iCount = 0

	sSql = "Select starttime, endtime, timeid from egov_class_time "
	sSql = sSql & " where classid = " & iClassid & " order by starttime"

	Set oTime = Server.CreateObject("ADODB.Recordset")
	oTime.Open sSQL, Application("DSN"), 3, 1

	'response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"">"
	Do While Not oTime.EOF
		' Display new time pick
		If iCount > 0 Then 
			response.write "<br />"
		End If 
		'response.write vbcrlf & "<tr><td><input type=""radio"" "
		response.write vbcrlf & "<input type=""radio"" "
		If (iTimeId = 0 And iCount = 0) Then 
			response.write " checked=""checked"" "
		Else 
			If CLng(iTimeId) = CLng(oTime("timeid")) Then 
				response.write " checked=""checked"" "
			End If 
		End If 
		response.write "name=""timeid"" value=""" & oTime("timeid") & """> " 
		' Handle Series
		If bIsParent And iClassTypeId = 1 Then
			response.write "Entire Series "
			If CheckIfFullSeries(iClassid) Then 
				response.write "<span class=""filledstatus""> &ndash; FILLED</span>"
			End If 
		Else 
			response.write oTime("starttime") 
			sOldTimes = oTime("starttime") 
			If oTime("endtime") <> oTime("starttime") Then
				response.write " &ndash; " & oTime("endtime")
			End If 
			If CheckIfFullSingle(oTime("timeid")) Then 
				response.write "<span class=""filledstatus""> &ndash; FILLED</span>"
			End If 
		End If 
		'response.write "</td></tr>"
		iCount = iCount + 1
		oTime.movenext 
	Loop 
	'response.write "</table>"

	oTime.close
	Set oTime = Nothing

End Sub 


'------------------------------------------------------------------------------
' Sub ShowPriceOptions( iClassid, iorgid)
'------------------------------------------------------------------------------
Sub ShowPriceOptions( iClassid, iorgid, sResidentType, iMemberCount, iMembershipId, iPriceDiscountId, iUserId )
	Dim sSql, oPrice, iCount, sDiscount, bMemberTypematch, sMemberType, iMinPricetype, iMaxPriceType
	Dim iFamilyMemberId, sMemberCode, cTotalPrice

	iCount = 0
	cTotalPrice = CDbl(0.00)

	sDiscount = GetDiscountPhrase( iPriceDiscountId )

	bResTypeMatch = CheckResTypeExists(iClassid, iorgid, sResidentType)

	' IF the class has a membership requirement
'	If CLng(iMembershipId) > CLng(0) Then 
'		iFamilyMemberId = GetCitizenFamilyId( iUserId ) ' Will use to determine membership of purchaser
'		sMemberCode = DetermineMembership( iFamilyMemberId, iUserid, iMembershipId )
'	Else
'		sMemberCode = "O"
'	End If 

	' IF at least one person in the family is a member, then set up for member pricing match
	If iMemberCount > 0 Then 
		sMemberType = "M"
	Else
		sMemberType = "O"
	End If 
	'bMemberTypematch = CheckResTypeExists(iClassid, iorgid, sMemberType)

	sSql = "Select P.pricetypeid, T.pricetypename, T.ismember, P.amount, T.pricetype, P.accountid, T.isfee, T.isbaseprice, T.checkmembership, P.membershipid, T.isdropin "
	sSql = sSql & " from egov_price_types T, egov_class_pricetype_price P "
	sSql = sSql & " where T.pricetypeid = P.pricetypeid "
	sSql = sSql & " and orgid = " & iorgid & " and P.classid = " & iClassid & " order by P.pricetypeid"
	'response.write sSql & "<br />"

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSQL, Application("DSN"), 3, 1
	'response.write "<!--egov_class_pricetype_price.pricetypeid -->"

	If Not oPrice.EOF Then 
		iMinPricetype = CLng(oPrice("pricetypeid"))
		iMaxPriceType = CLng(oPrice("pricetypeid"))

		response.write "<table id=""pricetable"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbcrlf

		Do While Not oPrice.EOF
  			If CLng(oPrice("pricetypeid")) < iMinPricetype Then
		    		iMinPricetype = CLng(oPrice("pricetypeid"))
  			End If 

  			If CLng(oPrice("pricetypeid")) > iMaxPriceType Then
		    		iMaxPriceType = CLng(oPrice("pricetypeid"))
  			End If 

			 'Display new time pick
			  response.write "<tr>" & vbcrlf
     response.write "    <td class=""pricetd"" nowrap=""nowrap"" valign=""top"">" & vbcrlf
			  response.write "        <input type=""checkbox"" "

  			If oPrice("isfee") Then 
		   		'Always check a fee
				    response.write " checked=""checked"" "
				    cTotalPrice = cTotalPrice + CDbl(oPrice("amount"))
			  Else 
				    If oPrice("isbaseprice") Then 
					     'always check a base price
					      response.write " checked=""checked"" "
       				cTotalPrice = cTotalPrice + CDbl(oPrice("amount"))
    				Else
				      	If oPrice("pricetype") = sResidentType Then 
       						'if the resident type requirement matches
        						response.write " checked=""checked"" "
        						cTotalPrice = cTotalPrice + CDbl(oPrice("amount"))
      					Else
						        If oPrice("checkmembership") Then
          							If oPrice("pricetype") = sMemberType Then 
            								response.write " checked=""checked"" "
            								cTotalPrice = cTotalPrice + CDbl(oPrice("amount"))
           						End If 
        						End If 
       				End If 
    				End If 
			  End If 

			response.write "id=""pricetypeid" & oPrice("pricetypeid") & """ name=""pricetypeid"" value=""" & oPrice("pricetypeid") & """ onClick=""clearMsg('pricetypeid" & oPrice("pricetypeid") & "');UpdatePriceTotal(document.PurchaseForm.amount" & oPrice("pricetypeid") & ".value, this.checked);"" /> " & vbcrlf
			response.write "        &nbsp; " & oPrice("pricetypename") & vbcrlf
   response.write "    </td>" & vbcrlf
			response.write "    <td class=""priceentrytd"" valign=""top"">" & vbcrlf
   response.write "        <input type=""text"" id=""amount" & oPrice("pricetypeid") & """ name=""amount" & oPrice("pricetypeid") & """ value=""" & Replace(FormatNumber(CDbl(oPrice("amount")),2),",","") & """ size=""10"" maxlength=""9"" onchange=""clearMsg('amount" & oPrice("pricetypeid") & "');ValidatePrice(this);"" />"  & vbcrlf
   response.write "    </td>" & vbcrlf
			response.write "    <td class=""priceentrytd"">&nbsp;" & FormatCurrency(oPrice("amount")) & "</td>" & vbcrlf
   response.write "    <td>" & vbcrlf

   if sDiscount <> "" then
      response.write "(<input type=""checkbox"" name=""useOverrideDiscount" & oPrice("pricetypeid") & """ value=""1"">Override Discount)" & vbcrlf
   else
      response.write "<input type=""hidden"" name=""useOverrideDiscount" & oPrice("pricetypeid") & """ value=""0"">" & vbcrlf
      response.write "&nbsp;" & vbcrlf
   end if

   response.write "    </td>" & vbcrlf
			response.write "    <td class=""pricemember"">" & vbcrlf

			if oPrice("ismember") then
				 'Show the membership for the one that requires membership
				  ShowMembership iMembershipId
			else
  				response.write " &nbsp; "
			end if 

			if oPrice("isdropin") then
 				'Input for drop in date
				  response.write "Date: <input type=""text"" class=""datefield"" id=""dropindate" & oPrice("pricetypeid") & """ name=""dropindate" & oPrice("pricetypeid") & """ value=""" & FormatDateTime(date(),2) & """ />&nbsp;<span class=""calendarimg"" style=""cursor:hand;""><img src=""../images/calendar.gif"" height=""16"" width=""16"" border=""0"" onclick=""clearMsg('dropindate" & oPrice("pricetypeid") & "');doCalendar('dropindate" & oPrice("pricetypeid") & "');"" /></span>" & vbcrlf
			end if 

			response.write "    </td>" & vbcrlf
			response.write "    <td class=""pricemember"">" & vbcrlf

			if sDiscount <> "" then 
  				response.write " (" & sDiscount & ")"
			else 
		  		response.write " &nbsp; "
			end if 

			response.write "    </td>" & vbcrlf
   response.write "</tr>" & vbcrlf
			iCount = iCount + 1
			oPrice.movenext 
		Loop 
		' final display of an other price pick
	'	response.write vbcrlf & "<tr><td class=""pricetd"" nowrap=""nowrap""><input type=""radio"" name=""pricetypeid"" value=""0"" /> &nbsp; Other Price</td><td>"
	'	response.write "<input type=""text"" name=""amount"" onKeyUp=""AutoSelect();"" value="""" size=""6"" maxlength=""6"" /></td></tr>"
'		response.write vbcrlf & "<tr></tr>"
		response.write "<tr>" & vbcrlf
  response.write "    <td>Total Price</td>" & vbcrlf
  response.write "    <td><span id=""displaytotalprice"">" & Replace(FormatNumber(cTotalPrice,2),",","") & "</span></td>" & vbcrlf
  response.write "    <td colspan=""4"">&nbsp;</td>" & vbcrlf
  response.write "</tr>" & vbcrlf
		response.write "</table>" & vbcrlf
		response.write "<input type=""hidden"" id=""totalprice"" name=""totalprice"" value=""" & cTotalPrice & """ />" & vbcrlf
		response.write "<input type=""hidden"" name=""minpricetypeid"" value=""" & iMinPricetype & """ />" & vbcrlf
		response.write "<input type=""hidden"" name=""maxpricetypeid"" value=""" & iMaxPriceType & """ />" & vbcrlf
	End If 

	oPrice.close
	Set oPrice = Nothing

End Sub 


'------------------------------------------------------------------------------
' Function CheckResTypeExists(iClassid, iorgid, sResidentType)
'------------------------------------------------------------------------------
Function CheckResTypeExists(iClassid, iorgid, sResidentType)
	Dim sSql, oCheck

	CheckResTypeExists = False 
	sSql = "Select count(T.pricetype) as hits "
	sSql = sSql & " from egov_price_types T, egov_class_pricetype_price P "
	sSql = sSql & " where T.pricetypeid = P.pricetypeid "
	sSql = sSql & " and orgid = " & iorgid & " and P.classid = " & iClassid & " and T.pricetype = '" & sResidentType & "'"

	Set oCheck = Server.CreateObject("ADODB.Recordset")
	oCheck.Open sSQL, Application("DSN"), 3, 1

	If CLng(oCheck("hits")) > 0 Then 
		CheckResTypeExists = True 
	End If 

	oCheck.close
	Set oCheck = Nothing

End Function 


'------------------------------------------------------------------------------
' Function ShowUserDropDown(iUserId)
'------------------------------------------------------------------------------
Function ShowUserDropDown( iUserId )
	Dim oCmd, oResident

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserWithAddressList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgid", 3, 1, 4, Session("OrgID"))
	    Set oResident = .Execute
	End With

	Do While Not oResident.eof 
		response.write vbcrlf & "<option value=""" & oResident("userid") & """"
		If CLng(iUserId) = CLng(oResident("userid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oResident("userlname") & ", " & oResident("userfname") & " &ndash; " & oResident("useraddress") & "</option>"
		oResident.movenext
	Loop 
		
	oResident.close
	Set oResident = Nothing
	Set oCmd = Nothing
End Function 


'------------------------------------------------------------------------------
Sub ShowUserInfo(iUserId, sUserType, sResidentDesc)
	Dim oCmd, oUser, sSql

'	Set oCmd = Server.CreateObject("ADODB.Command")
'	With oCmd
'		.ActiveConnection = Application("DSN")
'	    .CommandText = "GetEgovUserInfoList"
'	    .CommandType = 4
'		.Parameters.Append oCmd.CreateParameter("@iUserId", 3, 1, 4, iUserId)
'	    Set oUser = .Execute
'	End With

	sSQL = "SELECT userfname, userlname, useraddress, useraddress2, usercity, userstate, "
	sSQL = sSQL & " userzip, usercountry, useremail, userhomephone, "
	sSQL = sSQL & " userworkphone, userfax, userbusinessname, userpassword, "
	sSQL = sSQL & " userregistered, residenttype, residencyverified, registrationblocked, "
	sSQL = sSQL & " blockeddate, blockedadminid, blockedexternalnote, blockedinternalnote "
	sSQL = sSQL & " FROM egov_users WHERE userid = " & iUserId

	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open sSQL, Application("DSN"), 3, 1

	response.write vbcrlf & "<table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""signupuserinfo"">"
	response.write vbcrlf & "<tr><td align=""right"" valign=""top"">Name:</td><td >" & oUser("userfname") & " " & oUser("userlname") & "&nbsp;&nbsp;&nbsp;<strong>" & sResidentDesc & "</strong>"

	if not oUser("residencyverified") AND oUser("residenttype") = "R" then
  		if lcl_orghasfeature_residency_verification then
    			response.write " (not verified)"
    end if
 end if

	response.write "</td></tr>"
	response.write vbcrlf & "<tr><td align=""right"" valign=""top"">Email:</td><td>" & oUser("useremail") & "</td></tr>"
	response.write vbcrlf & "<tr><td align=""right"" valign=""top"">Phone:</td><td>" & FormatPhone(oUser("userhomephone")) & "</td></tr>"
	response.write vbcrlf & "<tr><td align=""right"" valign=""top"">Address:</td><td>" & oUser("useraddress") & "<br />" 

	If oUser("useraddress2") <> "" Then 
  		response.write oUser("useraddress2") & "<br />" 
	End If 

	If oUser("usercity") <> "" Or oUser("userstate") <> "" Or oUser("userzip") <> "" Then 
		  response.write oUser("usercity") & ", " & oUser("userstate") & " " & oUser("userzip") 
	End If 

	response.write "</td></tr>"

	' Handle blocked
	If oUser("registrationblocked") Then
		bRegistrationBlocked = True 
		response.write vbcrlf & "<tr><td colspan=""2""><span id=""warningmsg""> *** Registration Blocked *** </span></td></tr>"
		response.write vbcrlf & "<tr><td align=""right"" valign=""top"">Date:</td><td>" & oUser("blockeddate") & "</td></tr>"
		response.write vbcrlf & "<tr><td align=""right"" valign=""top"">By:</td><td>" & GetAdminName( oUser("blockedadminid") ) & "</td></tr>"
		response.write vbcrlf & "<tr><td align=""right"" valign=""top"" nowwrap=""nowrap"">Internal Note:</td><td>" & oUser.Fields("blockedinternalnote") & "</td></tr>"
		response.write vbcrlf & "<tr><td align=""right"" valign=""top"" nowwrap=""nowrap"">External Note:</td><td>" & oUser.Fields("blockedexternalnote") & "</td></tr>"
	End If 

	response.write vbcrlf & "</table>"

	oUser.close
	Set oUser = Nothing
'	Set oCmd = Nothing
	
End Sub 


'------------------------------------------------------------------------------
Sub ShowPaymentChoices()

	response.write vbcrlf & "<p><strong>Payment</strong><br />"
	
	response.write vbcrlf & "<select name=""paymenttypeid"" size=""1"">"
	ShowPaymentTypes
	response.write vbcrlf & "</select>&nbsp;&nbsp;&nbsp;&nbsp;"

	response.write vbcrlf & "<select name=""paymentlocationid"" size=""1"">"
	ShowPaymentLocations
	response.write vbcrlf & "</select>&nbsp;&nbsp;&nbsp;&nbsp;"
	response.write vbcrlf & "<input type=""button"" name=""complete"" class=""button"" style=""width:140px;text-align:center;"" value=""Complete Purchase"" onclick=""ValidateForm();"" />"
	response.write vbcrlf & "</p>"
	
	response.write vbcrlf & "<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>OR</strong>"
	response.write vbcrlf & "&nbsp;&nbsp;&nbsp;<input type=""button"" name=""waitlist"" class=""button"" value=""Add to Wait List"" onclick=""ValidateWait();"" /></p>"

End Sub 


'------------------------------------------------------------------------------
Sub ShowPaymentTypes()
	Dim sSql, oTypes

	sSql = "Select paymenttypeid, paymenttypename from egov_paymenttypes order by paymenttypeid"

	Set oTypes = Server.CreateObject("ADODB.Recordset")
	oTypes.Open sSQL, Application("DSN"), 3, 1

	Do While Not oTypes.EOF
		response.write vbcrlf & "<option value=""" & oTypes("paymenttypeid") & """>" & oTypes("paymenttypename") & "</option>"
		oTypes.movenext 
	Loop

	oTypes.close
	Set oTypes = Nothing

End Sub 


'------------------------------------------------------------------------------
Sub ShowPaymentLocations()
	Dim sSql, oLocations

	sSql = "Select paymentlocationid, paymentlocationname from egov_paymentlocations order by paymentlocationid"

	Set oLocations = Server.CreateObject("ADODB.Recordset")
	oLocations.Open sSQL, Application("DSN"), 3, 1

	Do While Not oLocations.EOF
		response.write vbcrlf & "<option value=""" & oLocations("paymentlocationid") & """>" & oLocations("paymentlocationname") & "</option>"
		oLocations.movenext 
	Loop

	oLocations.close
	Set oLocations = Nothing

End Sub 


'------------------------------------------------------------------------------
Function CheckIfFullSingle( iTimeId )
	Dim sSql, oTime

	sSql = "Select isnull(max,999999) as max, enrollmentsize from egov_class_time where timeid = " & iTimeId

	Set oTime = Server.CreateObject("ADODB.Recordset")
	oTime.Open sSQL, Application("DSN"), 0, 1

	If CLng(oTime("enrollmentsize")) >= CLng(oTime("max")) Then
		CheckIfFullSingle = True 
	Else
		CheckIfFullSingle = False 
	End If 

	oTime.close
	Set oTime = Nothing
End Function 


'------------------------------------------------------------------------------
Function CheckIfFullSeries( iClassid )
	Dim sSql, oTime

	CheckIfFullSeries = False 
	sSql = "Select T.timeid from egov_class C, egov_class_time T where C.classid = T.classid and C.parentclassid = " & iClassid

	Set oTime = Server.CreateObject("ADODB.Recordset")
	oTime.Open sSQL, Application("DSN"), 0, 1

	Do While Not oTime.EOF
		If CheckIfFullSingle( oTime("timeid") ) Then
			CheckIfFullSeries = True 
			Exit Do 
		End If 
		oTime.movenext
	Loop 

	oTime.close
	Set oTime = Nothing
End Function 


'------------------------------------------------------------------------------
Sub ShowAvailability( iClassid, bIsParent, iOptionid, iTimeId )
	Dim sSql, oAvail

	If bIsParent Then 
		' Get the availability of the children events
		sSql = "Select T.timeid, T.starttime, T.endtime, isnull(T.min,0) as min, isnull(T.max,0) as max, "
		sSql = sSql & " T.enrollmentsize, T.waitlistsize, C.startdate"
		sSql = sSql & " from egov_class_time T, egov_class C"
		sSql = sSql & " where C.parentclassid = " & iClassid
		sSql = sSql & " and T.timeid = " & iTimeId 
		sSql = sSql & " and C.classid = T.classid"
		sSql = sSql & " order by C.startdate, T.starttime"
	Else 
		' Get the single event
		sSql = "Select T.timeid, T.starttime, T.endtime, isnull(T.min,0) as min, isnull(T.max,0) as max,"
		sSql = sSql & " T.enrollmentsize, T.waitlistsize, C.startdate"
		sSql = sSql & " from egov_class_time T, egov_class C where C.classid = " & iClassid 
		sSql = sSql & " and T.timeid = " & iTimeId 
		sSql = sSql & " and C.classid = T.classid order by T.starttime"
	End If 

	Set oAvail = Server.CreateObject("ADODB.Recordset")
	oAvail.Open sSQL, Application("DSN"), 3, 1
	
	If Not oAvail.EOF Then 
		response.write vbcrlf & "<table id=""tableavail"" border=""0"" cellpadding=""2"" cellspacing""0"">"
		'response.write vbcrlf & "<caption>Availability</caption>"

		If bIsParent Then
			response.write vbcrlf & "<tr><th>Date</th><th>Time</th><th>Min</th><th>Max</th>"
		Else
			response.write vbcrlf & "<tr><th>Time</th><th>Min</th><th>Max</th>"
		End If
		If iOptionid = 1 then
			response.write "<th>Enrolled</th>"
		Else 
			response.write "<th>Attending</th>"
		End If 
		response.write "<th>Available</th><th>Waiting</th></tr>"
		 
		Do While Not oAvail.EOF
			response.write vbcrlf & "<tr><td>" 
			If bIsParent Then
				response.write DatePart("m",oAvail("startdate")) & "/" & DatePart("d",oAvail("startdate")) & "</td><td>"
			End If 
			response.write oAvail("starttime") 
			If oAvail("endtime") <> oAvail("starttime") Then
				response.write "&ndash;" & oAvail("endtime")
			End If 
			response.write "</td>"
			response.write "<td align=""center"">" 
			If CLng(oAvail("min")) = 0 Then
				response.write "none"
			Else
				response.write oAvail("min")
			End If 
			response.write "</td><td align=""center"">" 
			If CLng(oAvail("max")) = 0 Then
				response.write "none"
			Else
				response.write oAvail("max")
			End If
			' enrollment
			response.write "</td><td align=""center"">" & oAvail("enrollmentsize") & "</td>"
			' availability
			response.write "<td align=""center"">" 
			If CLng(oAvail("max")) = 0 Then
				response.write "n/a"
			Else
				iAvail = CLng(oAvail("max")) - CLng(oAvail("enrollmentsize"))
				If iAvail < 0 Then 
					iAvail = 0
				End If 
				response.write CLng(oAvail("max")) - CLng(oAvail("enrollmentsize"))
			End If 
			response.write "</td>"
			' get waitlist here
			response.write "<td align=""center"">" & oAvail("waitlistsize") & "</td>"
			response.write "</tr>"
			oAvail.movenext 
		Loop 
		response.write vbcrlf & "</table>"
	End If 

	oAvail.close
	Set oAvail = Nothing

End Sub 



%>
