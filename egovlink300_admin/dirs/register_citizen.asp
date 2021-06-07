<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<% dim blnSpecialAddress, strResidenttype
blnSpecialAddress = false %>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="inc_dbfunction.asp" //-->
<!-- #include file="citizen_global_functions.asp" //-->
<!-- #include file="../../egovlink300_global/includes/inc_passencryption.asp" //-->
<%	
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: register_citizen.asp
' AUTHOR: ????
' CREATED: ????
' COPYRIGHT: Copyright 2005 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the creation of citizen users
'
' MODIFICATION HISTORY
' 1.0   ????   ???? ???? - INITIAL VERSION
' 1.1	10/05/2011	Steve Loar - Added gender selection pick
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sError, sEmailnotavailable, bShowGenderPicks
Dim bHasResidentStreets, bFound, sResidenttype, sBusinessAddress, bHasBusinessStreets, sRedirect, bHasResidentTypes

sLevel = "../"  'Override of value from common.asp

'See if they have rights and if the page is not shutdown
PageDisplayCheck "add citizens", sLevel	' In common.asp

lcl_hidden = "hidden"  'Show/Hide all hidden fields.  TEXT=Show,HIDDEN=Hide

'Org Features
lcl_orghasfeature_residency_verification = orghasfeature("residency verification")
lcl_orghasfeature_default_area_code      = orghasfeature("default area code")
lcl_orghasfeature_large_address_list     = orghasfeature("large address list")
lcl_orghasfeature_subscriptions          = orghasfeature("subscriptions")
lcl_orghasfeature_job_postings           = orghasfeature("job_postings")
lcl_orghasfeature_bid_postings           = orghasfeature("bid_postings")
lcl_orghasfeature_donotknock             = orghasfeature("donotknock")
lcl_orghasfeature_citizenregistration_novalidate_address = orghasfeature("citizenregistration_novalidate_address")
bShowGenderPicks = orgHasFeature( "display gender pick" )

lcl_orghasneighborhoods = orghasneighborhoods(session("orgid"))

if request.servervariables("request_method") = "POST" then  'This should be the page saving itself
	if session("orgid") = "60" then
		blnSpecialAddress = True
	
		strResidenttype = request("egov_users_residenttype")
		if request.form("residentstreetnumber") = "701" and request.form("skip_address") = "LAUREL ST" and lcase(request.form("egov_users_usercity")) = "menlo park" _
			and lcase(request.form("egov_users_userstate")) = "ca" and left(request.form("egov_users_userzip"),5) = "94025" then
			strResidenttype = "E"
		elseif strResidenttype = "R" then
		elseif request.form("egov_users_userbusinessname") <> "" and (request.form("skip_Baddress") <> "0000" OR request.form("egov_users_userbusinessaddress") <> "") then
			strResidenttype = "B"
		elseif left(request.form("egov_users_userzip"),5) = "94025" and lcase(request.form("egov_users_usercity")) = "menlo park" then
			strResidenttype = "U"
		else
			strResidenttype = "N"
		end if
		'response.write strResidenttype
		'response.end
	end if

	'Add them to the egov_users table
	userid = ProcessRecords()  'in inc_dbfunction.asp
	session("eGovUserId") = userid

	' because processRecords() is so bad we have to fix the crap it inserts sometimes
	If request("egov_users_gender") = "N" Then
		setGenderToNull userid
	End If 

	'Set the "Do Not Knock" values
	if lcl_orghasfeature_donotknock then
		updateDoNotKnockValues userid, request("skip_isOnDoNotKnockList_peddlers"), request("skip_isOnDoNotKnockList_solicitors"), _
			request("skip_isDoNotKnockVendor_peddlers"), request("skip_isDoNotKnockVendor_solicitors")
	end if

	if request("skip_residencyverified") = "on" then
		bResidencyVerified = True 
	else
		bResidencyVerified = False 
	end if

	if request("skip_emailnotavailable") = "on" then
		sEmailnotavailable = "1"
	else
		sEmailnotavailable = "0"
	end if

	'Insert into the Family Members table
	AddFamilyMember userid, request("egov_users_userfname"), request("egov_users_userlname"), "Yourself", "NULL", userid

	'Update egov_users with new familyid, neighborhood and relationship
	UpdateFamilyId userid, userid, request("egov_users_relationshipid"), request("egov_users_neighborhoodid"), bResidencyVerified, sEmailnotavailable

	'Add the subscriptions
	InsertMailLists userid

	'Send them an email if that is selected
	if lcase(request("skip_notifyuser")) = "on" AND request("egov_users_useremail") <> "" then
		notifyuser session( "orgid" ), session("sOrgName"), session("egovclientwebsiteurl"), request("egov_users_useremail"), request("egov_users_password")
	end if

	'Take them back to where they came from
	'Default them back to the display citizen page
	if session("RedirectPage") <> "" then
		sRedirect = session("RedirectPage") 
		session("RedirectPage") = ""
		response.redirect sRedirect
	else
		response.redirect "display_citizen.asp"
	end if
end If

%>
<html>
<head>
	<title><%=langBSCommittees%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script language="javascript" src="../scripts/ajaxLib.js"></script>
	<script language="javascript" src="../scripts/formatnumber.js"></script>
	<script language="javascript" src="../scripts/removespaces.js"></script>
	<script language="javascript" src="../scripts/setfocus.js"></script>
	<script language="javascript" src="../scripts/isvaliddate.js"></script>

	<script language="javascript" src="../prototype/prototype-1.6.0.2.js"></script>
	
	<script language="javascript" src="../scriptaculous/src/scriptaculous.js"></script>

	<script language="javascript">
	<!--

	var winHandle;
	var w = (screen.width - 640)/2;
	var h = (screen.height - 450)/2;

	function doCheck() 
	{
  		// If they are using the large address feature
  		var exists = eval(document.register["residentstreetnumber"]);
 		if(exists) 
		{
<%
			'This feature is ENABLED then it DISABLES the large address validation and simply does the form validation.
			if not lcl_orghasfeature_citizenregistration_novalidate_address then
				response.write "// If a street number was entered" & vbcrlf
				response.write "if (document.register.residentstreetnumber.value != '') {" & vbcrlf
				response.write "				checkAddress( 'FinalCheck', 'yes' );" & vbcrlf
				response.write "} else {" & vbcrlf
				response.write "				checkDuplicateCitizens('FinalUserCheckFailed');" & vbcrlf
				response.write "}" & vbcrlf
			else
				response.write "checkDuplicateCitizens('FinalUserCheckFailed');"  & vbcrlf
			end if
%>
 		} 
		else	
		{
   			checkDuplicateCitizens( 'FinalUserCheckFailed' );
 		}
	}

	function checkAddress( sReturnFunction, sSave )
	{
		// Remove any extra spaces
		$("residentstreetnumber").value = removeSpaces($("residentstreetnumber").value);

		// check the number for non-numeric values
		var rege = /^\d+$/;
		var Ok = rege.exec($("residentstreetnumber").value);

		if ( ! Ok )
		{
			alert("The Resident Street Number cannot be blank and must be numeric.");
			setfocus($("residentstreetnumber"));
			return false;
		}

		// check that they picked a street name
		if ( $("skip_address").value == '0000')
		{
			alert("Please select a street name from the list first.");
			setfocus(document.register.skip_address);
			return false;
		}

		// This is here because window.open in the Ajax callback routine will not work
		//winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&stnumber=' + document.register.residentstreetnumber.value + '&stname=' + document.register.skip_address.value + '&sCheckType=' + sReturnFunction + '", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		//self.focus();
		// Fire off Ajax routine
		doAjax('checkaddress.asp', 'stnumber=' + document.register.residentstreetnumber.value + '&stname=' + document.register.skip_address.value, sReturnFunction, 'get', '0');

	}

	function CheckResults( sResults )
	{
		// Process the Ajax CallBack when the validate address button is clicked
		if (sResults == 'FOUND')
		{
			//if(winHandle != null && ! winHandle.closed)
			//{ 
			//	winHandle.close();
			//}
			document.register.egov_users_useraddress.value = '';
			document.register.egov_users_residenttype.value = 'R';
			alert("This is a valid address in the system.");
		}
		else
		{
			//winHandle.focus();
			document.register.egov_users_residenttype.value = 'N';
			PopAStreetPicker('CheckResults', 'no');
		}
	}

	function PopAStreetPicker( sReturnFunction, sSave )
	{
		// pop up the address picker
		winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&stnumber=' + document.register.residentstreetnumber.value + '&stname=' + document.register.skip_address.value + '&sCheckType=' + sReturnFunction + '", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
	}

	function FinalCheck( sResults )
	{
		// Process the Ajax CallBack for the save process
		if (sResults == 'FOUND')
		{
			//if(winHandle != null && ! winHandle.closed)
			//{ 
			//	winHandle.close();
			//}
			document.register.egov_users_residenttype.value = 'R';
			checkDuplicateCitizens( 'FinalUserCheckFailed' );
		}
		else
		{
			//winHandle.focus();
			document.register.egov_users_residenttype.value = 'N';
			PopAStreetPicker('FinalCheck', 'yes');
		}
		
	}

	function validate() 
	{
		var msg="";
		var bUsedAddressDropdown = false;

		if ( document.register.skip_emailnotavailable.checked == false )
		{
			// They will login so validate the email and password
			if (document.register.egov_users_useremail.value == "" )
			{
				msg+="The email cannot be blank.\n";
			}
			//var rege = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*\.(\w{2}|(com|net|org|edu|mil|gov|biz|us|COM|NET|ORG|EDU|MIL|GOV|BIZ|US))$/;
			var rege = /.+@.+\..+/i;
			var Ok = rege.test(document.register.egov_users_useremail.value);

			if (! Ok)
			{
				msg+="The email must be in a valid format.\n";
			}
			if (document.register.egov_users_password.value == "" )
			{
				msg+="The password cannot be blank.\n";
			}
			if (document.register.skip_userpassword2.value == "" )
			{
				msg+="The verify password cannot be blank.\n";
			}
			if(document.register.egov_users_password.value != document.register.skip_userpassword2.value)
			{
				msg+="The passwords you have entered do not match.\n";
			}
		}
		
		if (document.register.egov_users_userfname.value == "" )
		{
			msg+="The first name cannot be blank.\n";
		}
		if (document.register.egov_users_userlname.value == "" )
		{
			msg+="The last name cannot be blank.\n";
		}


		// set the emergency phone
		if (document.register.skip_emergencyphone_areacode.value != "" || document.register.skip_emergencyphone_exchange.value != "" || document.register.skip_emergencyphone_line.value != "" )
		{
			var sPhone = document.register.skip_emergencyphone_areacode.value + document.register.skip_emergencyphone_exchange.value + document.register.skip_emergencyphone_line.value;
			if (sPhone.length < 10)
			{
				msg += "The Emergency Phone must be a valid phone number, including area code, or blank\n";
			}
			else
			{
				document.register.egov_users_emergencyphone.value = document.register.skip_emergencyphone_areacode.value + document.register.skip_emergencyphone_exchange.value + document.register.skip_emergencyphone_line.value;
				var rege = /^\d+$/;
				var Ok = rege.exec(document.register.egov_users_emergencyphone.value);
				if ( ! Ok )
				{
					msg += "The Emergency Phone must be a valid phone number, including area code, or blank\n";
				}
			}
		}

		// check the work phone
		if (document.register.skip_work_areacode.value != "" || document.register.skip_work_exchange.value != "" || document.register.skip_work_line.value != "" || document.register.skip_work_ext.value != "")
		{
			var sPhone = document.register.skip_work_areacode.value + document.register.skip_work_exchange.value + document.register.skip_work_line.value;
			if (sPhone.length < 10)
			{
				msg += "The work phone must be a valid phone number, including area code, or blank.\n";
			}
			else
			{
				document.register.egov_users_userworkphone.value = document.register.skip_work_areacode.value + document.register.skip_work_exchange.value + document.register.skip_work_line.value + document.register.skip_work_ext.value;
				var rege = /^\d+$/;
				var Ok = rege.exec(document.register.egov_users_userworkphone.value);
				if ( ! Ok )
				{
					msg += "The work phone must be a valid phone number, including area code, or blank.\n";
				}
			}
		}

		// check the fax
		if (document.register.skip_fax_areacode.value != "" || document.register.skip_fax_exchange.value != "" || document.register.skip_fax_line.value != "" )
		{
			var fPhone = document.register.skip_fax_areacode.value + document.register.skip_fax_exchange.value + document.register.skip_fax_line.value;
			if (fPhone.length < 10)
			{
				msg += "Fax must be a valid phone number, including area code, or blank.\n";
			}
			else
			{
				document.register.egov_users_userfax.value = document.register.skip_fax_areacode.value + document.register.skip_fax_exchange.value + document.register.skip_fax_line.value;
				var rege = /^\d+$/;
				var Ok = rege.exec(document.register.egov_users_userfax.value);
				if ( ! Ok )
				{
					msg += "Fax must be a valid phone number, including area code, or blank.\n";
				}
			}
		}

		// check the cell phone
		if (document.register.skip_cell_areacode.value != "" || document.register.skip_cell_exchange.value != "" || document.register.skip_cell_line.value != "" )
		{
			var cPhone = document.register.skip_cell_areacode.value + document.register.skip_cell_exchange.value + document.register.skip_cell_line.value;
			//alert(cPhone);
			if (cPhone.length < 10)
			{
				msg += "length - The cell phone must be a valid phone number, including area code, or blank.\n";
			}
			else
			{
				document.register.egov_users_usercell.value = document.register.skip_cell_areacode.value + document.register.skip_cell_exchange.value + document.register.skip_cell_line.value;
				//alert('[' + cPhone + ']');
				var cellrege = /^\d+$/;
				var cOk = cellrege.exec(cPhone);
				if ( ! cOk )
				{
					msg += "The cell phone must be a valid phone number, including area code, or blank.\n";
				}
			}
		}

		// check the home phone number
		document.register.egov_users_userhomephone.value = document.register.skip_user_areacode.value + document.register.skip_user_exchange.value + document.register.skip_user_line.value;
		if (document.register.egov_users_userhomephone.value != "" )
		{
			var hPhone = document.register.egov_users_userhomephone.value;
			if (hPhone.length < 10)
			{
				msg += "The home phone must be a valid phone number, including area code, or blank.\n";
			}
			else
			{
				var rege = /^\d+$/;
				var hOk = rege.exec(document.register.egov_users_userhomephone.value);
				if ( ! hOk )
				{
					msg += "The home phone must be a valid phone number, including area code, or blank.\n";
				}
			}
		}
	//	else
	//	{
	//		msg+="The home phone cannot be blank.\n";
	//	}

		// Handle the birthdate 
		var birthrege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
		var birthOk = birthrege.test(document.register.egov_users_birthdate.value);

		if (document.register.egov_users_birthdate.value != "")
		{
			var sBirthdate = document.register.egov_users_birthdate.value;
			var sDateParts = sBirthdate.split("/");
			if (parseInt(sDateParts[0]) == 0 || parseInt(sDateParts[1]) == 0 || parseInt(sDateParts[2]) < 1905)
			{
				birthOk = false;
			}
			if (! birthOk )
			{
				msg += "Invalid birth date value or format. \nThe birth date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again, or leave it blank.";
			}
			else
			{
				if (isValidDate( document.register.egov_users_birthdate.value ) == false)
				{
					msg += "Invalid birth date value or format. \nBirth date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again, or leave it blank.";
				}
			}
		}
		

		// Check that they selected a resident type
		var rexists = eval(document.register["skip_egov_users_residenttype"]);
		if (rexists)
		{
			if (document.register.skip_egov_users_residenttype.selectedIndex > -1)
			{
				var relement = document.register.skip_egov_users_residenttype;
				var rselectedvalue = relement.options[relement.selectedIndex].value;

				//alert( bselectedvalue );
				//  0 is the first pick that we do not want
				if (rselectedvalue != "0")
				{
					document.register.egov_users_residenttype.value = rselectedvalue;
				}
				else
				{
					msg+="Please select a resident type.\n";
				}
			}
		}


		// Process the business address if one was chosen
		var bexists = eval(document.register["skip_Baddress"]);
		if(bexists)
		{
			//See if they picked from the business dropdown and put that in the address field 
			if (document.register.skip_Baddress.selectedIndex > -1)
			{
				var belement = document.register.skip_Baddress;
				var bselectedvalue = belement.options[belement.selectedIndex].value;

				//alert( bselectedvalue );
				//  0000 is the first pick that we do not want
				if (bselectedvalue != "0000")
				{
					document.register.egov_users_userbusinessaddress.value = bselectedvalue;
				}
			}
		}

		// Process the resident address if one was chosen - this is second to set the local resident type
		var exists = eval(document.register["skip_Raddress"]);
		if(exists)
		{
			// See if they picked from the resident dropdown and put that in the address field 
			if (document.register.skip_Raddress.selectedIndex > -1)
			{
				var element = document.register.skip_Raddress;
				var selectedvalue = element.options[element.selectedIndex].value;

				//alert( selectedvalue );
				//  0000 is the first pick that we do not want
				if (selectedvalue != "0000")
				{
					document.register.egov_users_useraddress.value = selectedvalue;
					bUsedAddressDropdown = true;
				}
			}
		}

		// handle the large quantity street addresses
		exists = eval(document.register["residentstreetnumber"]);
		if(exists)
		{
			if ( document.register.residentstreetnumber.value != '' )
			{
				// See if they picked from the resident dropdown and put that in the address field 
				if (document.register.skip_address.selectedIndex > -1)
				{
					var element = document.register.skip_address;
					var selectedvalue = element.options[element.selectedIndex].value;

					//alert( selectedvalue );
					//  0000 is the first pick that we do not want
					if (selectedvalue != "0000")
					{
						document.register.egov_users_useraddress.value = document.register.residentstreetnumber.value + ' ' + selectedvalue;
						bUsedAddressDropdown = true;
					}
				}
			}
		}

		var bSelected = false;
		// If subscriptions were picked, then make sure that an email was entered.
		exists = eval(document.register["maillist"]);
		if(exists)
		{
			// see if one or more was checked
			if (document.register.maillist.length)
			{
				for (i=0; i<document.register.maillist.length; i++)
				{
					if (document.register.maillist[i].checked==true)
					{
						bSelected = true;
						break;
					}
				}
			}
			else // Only one picked
			{
				if (document.register.maillist.checked)
				{
					bSelected = true;
				}
			}

			if (bSelected)
			{
				// see if email not available was checked
				if ( document.register.skip_emailnotavailable.checked == true )
				{
					msg+="Subscriptions cannot be selected when the email address is not available.\n";
				}
				else
				{
					// see if an email address was not entered
					if (document.register.egov_users_useremail.value == "")
					{
						msg+="Subscriptions cannot be selected when the email address is not provided.\n";
					}
				}
			}
		}

		if(msg != "")
		{
			if (bUsedAddressDropdown)
			{
				document.register.egov_users_useraddress.value = '';
			}
			msg="Your form could not be submitted for the following reasons.\n\n" + msg;
			alert(msg);
		}
		else 
		{	
			// Set some final things and then submit
			if (document.register.egov_users_birthdate.value == "")
			{
				document.register.egov_users_birthdate.value = "NULL";
			}
			if ( document.register.skip_emailnotavailable.checked == true )
			{
				// NULL them out so they save as NULL
				document.register.egov_users_useremail.value = "NULL";
				document.register.egov_users_password.value = "NULL";
			}
			document.register.submit(); 
		}
	}

	function GoBack(sUrl)
	{
		//alert( sUrl);
		location.href='' + sUrl;
	}

	function GetRndPwd()
	{
		var pwd;

		pwd = Math.round(Math.random() * 1000000 );
		//alert(pwd);
		document.register.egov_users_password.value = pwd;
		document.register.skip_userpassword2.value = pwd;

	}

	var isNN = (navigator.appName.indexOf("Netscape")!=-1);

	function autoTab(input,len, e) 
	{
		var keyCode = (isNN) ? e.which : e.keyCode; 
		var filter = (isNN) ? [0,8,9] : [0,8,9,16,17,18,37,38,39,40,46];

		if(input.value.length >= len && !containsElement(filter,keyCode)) {
			input.value = input.value.slice(0, len);
		var addNdx = 1;

		while(input.form[(getIndex(input)+addNdx) % input.form.length].type == "hidden") 
		{
			addNdx++;
			//alert(input.form[(getIndex(input)+addNdx) % input.form.length].type);
		}

		input.form[(getIndex(input)+addNdx) % input.form.length].focus();
	}

	function containsElement(arr, ele) 
	{
		var found = false, index = 0;

		while(!found && index < arr.length)
			if(arr[index] == ele)
				found = true;
			else
				index++;
		return found;
	}

	function getIndex(input) 
	{
		var index = -1, i = 0, found = false;

		while (i < input.form.length && index == -1)
			if (input.form[i] == input)index = i;
			else i++;
				return index;
	}
		return true;
	}

	function checkDuplicateCitizens( sCheckReturn )
	{
		var sAlert;
		// Remove any extra spaces
		//document.register.egov_users_userlname.value = removeSpaces(document.register.egov_users_userlname.value); // Causes any spaces within the name to be removed, that is bad
		if ( document.register.egov_users_userlname.value == "" )
		{
			if(sCheckReturn == 'citizenCheckReturn')
			{
				sAlert = "This check requires that the Last Name cannot be blank. Please provide one, then try the check again.";
			}
			else
			{
				sAlert = "The Last Name cannot be blank. Please provide one, then try saving again.";
			}

			alert( sAlert );
			setfocus(document.register.egov_users_userlname);
			return false;
		}

		if ( document.register.egov_users_userfname.value == "" )
		{
			if(sCheckReturn == 'citizenCheckReturn')
			{
				sAlert = "This check requires that the First Name cannot be blank. Please provide one, then try the check again.";
			}
			else
			{
				sAlert = "The First Name cannot be blank. Please provide one, then try saving again.";
			}

			alert( sAlert );
			setfocus(document.register.egov_users_userfname);
			return false;
		}

		// Fire off Ajax routine
		doAjax('checkduplicatecitizens.asp', 'userlname=' + document.register.egov_users_userlname.value + '&userfname=' + document.register.egov_users_userfname.value, sCheckReturn, 'get', '0');
	}

	function citizenCheckReturn( sReturn )
	{
		// Process the Ajax CallBack 
		if (sReturn != 'NEWCITIZEN')
		{
			//alert( sReturn )
			alert("This citizen may be a duplicate of a citizen already in the system.\n\nPossible matches are: " + sReturn);
		}
		else
		{
			alert("This does not appear to be a duplicate of any citizen in the system.");
		}
	}

	function FinalUserCheckFailed( sReturn )
	{
		// Handle the results from the final citizen check 
		if (sReturn != 'NEWCITIZEN')
		{
			if ( confirm("This citizen may be a duplicate of a citizen already in the system.\n\nPossible matches are: " + sReturn + '\n\nDo you still wish to add this citizen?') )
			{
				validate();
			}
		}
		else
		{
			validate();
		}
	}

	function checkDuplicateEmails()
	{
		if ($F("egov_users_useremail") == '')
		{
			alert('Please enter an email before checking for duplicates.');
			$("egov_users_useremail").focus();
			return;
		}
		// Fire off Ajax routine
		doAjax('checkduplicatecitizenemail.asp', 'email=' + $F("egov_users_useremail"), "emailCheckReturn", 'get', '0');
	}

	function emailCheckReturn( sReturn )
	{
		// Process the Ajax CallBack 
		if (sReturn == 'DUPLICATE')
		{
			alert("This email is already in use.");
		}
		else
		{
			alert("This email is not already in use.");
		}
	}

	//-->
	</script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"-->
<%
'Set up default values
lcl_default_areacode = ""
lcl_default_city     = ""
lcl_default_state    = ""
lcl_default_zip      = ""

if lcl_orghasfeature_default_area_code then
 lcl_default_areacode	= getDefaultOrgValue("defaultareacode")
end if

lcl_default_city           = getDefaultOrgValue("defaultcity")
lcl_default_state          = getDefaultOrgValue("defaultstate")
lcl_default_zip            = getDefaultOrgValue("defaultzip")
lcl_default_relationshipid = getDefaultRelationshipID(session("orgid"))

response.write "<div id=""content"">"
response.write "  <div id=""centercontent"">"

response.write "  <form method=""post"" name=""register"" action=""register_citizen.asp"">"
response.write "  		<input type=""hidden"" name=""columnnameid"" value=""userid"" />"
response.write "  		<input type=""hidden"" name=""egov_users_userregistered"" value=""1"" />"
response.write "  		<input type=""hidden"" name=""egov_users_orgid"" value=""" & session("orgid") & """ />"
response.write "  		<input type=""hidden"" name=""egov_users_relationshipid"" value=""" & lcl_default_relationshipid & """ />"
response.write "  		<input type=""hidden"" name=""egov_users_residenttype"" value=""N"" />"
response.write "  		<input type=""hidden"" name=""egov_users_headofhousehold"" value=""1"" />"
response.write "  		<input type=""hidden"" name=""ef:egov_users_userfname-text/req"" value=""First name"" />"
response.write "  		<input type=""hidden"" name=""ef:egov_users_userlname-text/req"" value=""Last name"" />"

response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
response.write "  <tr>"
response.write "      <td><font size=""+1""><strong>New Citizen Registration</strong></font></td>"
response.write "  </tr>"
response.write "  <tr>"
response.write "     <td valign=""top"">"
response.write "         <div style=""font-size:10px; padding-bottom:1em;"">"

if session("RedirectPage") <> "" then
	response.write "<input type=""button"" class=""button"" value=""" & session("RedirectLang") & """ onclick=""GoBack('" & session("RedirectPage") & "');"" />&nbsp;&nbsp;"
end if

response.write "           <input type=""button"" class=""button"" value=""Create User"" onclick=""doCheck();"" />"
response.write "         </div>"
response.write "         <div class=""shadow"" id=""registershadow"">"
response.write "         <table border=""0"" class=""tableadmin"" id=""registertable"" cellpadding=""4"" cellspacing=""0"" width=""100%"">"
response.write "           <tr>"
response.write "             		<th>&nbsp;</th>"
response.write "             		<th align=""left"">(<font color=""#ff0000"">*</font> Indicates required fields)</th>"
response.write "           </tr>"
response.write "           <tr>"
response.write "             		<td class=""label"" align=""right"" nowrap=""nowrap"">"
buildRequiredFieldLabel "First Name:"
response.write "             		</td>"
response.write "             		<td>"
response.write "                   <input type=""text"" value="""" name=""egov_users_userfname"" id=""egov_users_userfname"" size=""50"" maxlength=""50"" />"
response.write "               </td>"
response.write "           </tr>"
response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
buildRequiredFieldLabel "Last Name:"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""text"" value="""" name=""egov_users_userlname"" id=""egov_users_userlname"" size=""50"" maxlength=""50"" />"
response.write "               			 &nbsp;&nbsp;"
response.write "               			 <input type=""button"" class=""button"" value=""Check for Duplicates"" onclick=""checkDuplicateCitizens( 'citizenCheckReturn' )"" />"
response.write "               </td>"
response.write "           </tr>"

If bShowGenderPicks Then 
	response.write vbcrlf & "<tr>"
	response.write "<td class=""label"" align=""right"">"
	response.write "Gender:"
	response.write "</td>"
	response.write "<td>"
	DisplayGenderPicks "egov_users_gender", "N"		' in citizen_global_functions.asp
	response.write "</td>"
	response.write "</tr>"
End If 

response.write "           <tr>"
response.write "               <td class=""label"" align=""right"">"
response.write "                   Birthdate:"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""text"" value="""" name=""egov_users_birthdate"" id=""egov_users_birthdate"" size=""10"" maxlength=""10"" /> (MM/DD/YYYY)"
response.write "               </td>"
response.write "           </tr>"
response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"" valign=""top"">"
response.write "                   &nbsp;"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""checkbox"" name=""skip_emailnotavailable"" id=""skip_emailnotavailable"" /> Email Not Available (Citizen Will Not Login)"
response.write "               </td>"
response.write "           </tr>"
response.write "           <tr>" & vbrlf
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"" valign=""top"">"
response.write "                   Email:"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""text"" name=""egov_users_useremail"" id=""egov_users_useremail"" value="""" size=""50"" maxlength=""100"" />"
response.write "                   &nbsp;&nbsp;"
response.write "                   <input type=""button"" class=""button"" value=""Check for Duplicates"" onclick=""checkDuplicateEmails();"" />"
response.write "                   <br /><input type=""checkbox"" name=""skip_notifyuser"" id=""skip_notifyuser"" checked=""checked"" /> Send Registration Confirmation to Citizen"
response.write "               </td>"
response.write "           </tr>"
response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">&nbsp;</td>"
response.write "              	<td>"
response.write "                   <input type=""button"" value=""Generate Random Password"" name=""skip_generate_pwd"" id=""skip_generate_pwd"" class=""button"" onclick=""GetRndPwd();"" />"
response.write "               </td>"
response.write "           </tr>"
response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
response.write "                   Password:"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""password"" value="""" name=""egov_users_password"" id=""egov_users_password"" size=""50"" maxlength=""50"" />"
response.write "               </td>"
response.write "           </tr>"
response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
response.write "                   Verify Password:"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""password"" value="""" name=""skip_userpassword2"" id=""skip_userpassword2"" size=""50"" maxlength=""50"" />"
response.write "               </td>"
response.write "           </tr>"

bHasResidentTypes = HasResidentTypes()
bFound = False

if bHasResidentTypes then
	response.write "           <tr>"
	response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
	buildRequiredFieldLabel "Resident Type:"
	response.write "               </td>"
	response.write "               <td>"
	if session("orgid") <> "60" then
		response.write DisplayResidentTypes()
	else
		response.write "Will Auto-populate"
	end if

	if lcl_orghasfeature_residency_verification then
		response.write "&nbsp;<input name=""skip_residencyverified"" id=""skip_residencyverified"" type=""checkbox"" checked=""checked"" /> Residency Verified"
	end if

	response.write "               </td>"
	response.write "           </tr>"
end if

response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
response.write "                   Home Phone:"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""hidden"" value="""" name=""egov_users_userhomephone"" id=""egov_users_userhomephone"" />"
response.write "                  (<input type=""text"" value=""" & lcl_default_areacode & """ name=""skip_user_areacode"" id=""skip_user_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />)&nbsp;"
response.write "                  	<input type=""text"" value="""" name=""skip_user_exchange"" id=""skip_user_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />&ndash;"
response.write "                  	<input type=""text"" value="""" name=""skip_user_line"" id=""skip_user_line"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" />"
response.write "               </td>"
response.write "           </tr>"
response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
response.write "                   Cell Phone:"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""hidden"" value="""" name=""egov_users_usercell"" />"
response.write "                		(<input type=""text"" value=""" & lcl_default_areacode & """ name=""skip_cell_areacode"" id=""skip_cell_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />)&nbsp;"
response.write "                   <input type=""text"" value="""" name=""skip_cell_exchange"" id=""skip_cell_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />&ndash;"
response.write "                   <input type=""text"" value="""" name=""skip_cell_line"" id=""skip_cell_line"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" />"
response.write "               </td>"
response.write "           </tr>"
response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
response.write "                   Fax:"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""hidden"" value="""" name=""egov_users_userfax"" id=""egov_users_userfax"" />"
response.write "                  (<input type=""text"" value="""" name=""skip_fax_areacode"" id=""skip_fax_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />)&nbsp;"
response.write "                   <input type=""text"" value="""" name=""skip_fax_exchange"" id=""skip_fax_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />&ndash;"
response.write "                   <input type=""text"" value="""" name=""skip_fax_line"" id=""skip_fax_line"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" />"
response.write "               </td>"
response.write "           </tr>"

bHasResidentStreets = HasResidentTypeStreets( "R" )
bFound = False
lcl_address_label = "Address:"

if bHasResidentStreets then
	lcl_address_label = "Address (if not listed):"

	if not lcl_orghasfeature_large_address_list then
		response.write "           <tr>"
		response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
		response.write "                   Resident Address:"
		response.write "               </td>"
		response.write "               <td>"
		'DisplayAddresses "R"
		response.write "               </td>"
		response.write "           </tr>"
	else
		response.write "           <tr>"
		response.write "               <td class=""label"" align=""right"" valign=""top"" nowrap=""nowrap"">"
		response.write "                   Resident Address:"
		response.write "               </td>"
		response.write "               <td>"
		DisplayLargeAddressList session("orgid"), "R"

		'If this feature is ENABLED then it DISABLES the large address validation and 
		'simply does the form validation.
		if not lcl_orghasfeature_citizenregistration_novalidate_address then
			response.write "<input type=""button"" class=""button"" value=""Validate Address"" onclick=""checkAddress('CheckResults', 'no');"" />"
		end if
		response.write "               </td>"
		response.write "           </tr>"
	end if
end if

response.write "           </tr>"
response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">" & lcl_address_label & "</td>"
response.write "               <td>"
response.write "                   <input type=""text"" value="""" name=""egov_users_useraddress"" id=""egov_users_useraddress"" size=""50"" maxlength=""100"" />"
response.write "               </td>"
response.write "           </tr>"
response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
response.write "                   Resident Unit:"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""text"" value="""" name=""egov_users_userunit"" id=""egov_users_userunit"" size=""11"" maxlength=""10"" />"

if lcl_orghasneighborhoods then
	response.write "               </td>"
	response.write "           </tr>"
	response.write "           <tr>"
	response.write "               <td class=""label"" align=""right"">"
	response.write "                   Neighborhood:"
	response.write "               </td>"
	response.write "               <td>"
	DisplayNeighborhoods session("orgid"), 0
else
	response.write "                   <input type=""hidden"" name=""egov_users_neighborhoodid"" id=""egov_users_neighborhoodid"" value=""0"" />"
end if

response.write "               </td>"
response.write "           </tr>"
response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
response.write "                   City:"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""text"" value=""" & lcl_default_city & """ name=""egov_users_usercity"" id=""egov_users_usercity"" size=""50"" maxlength=""100"" />"
response.write "               </td>"
response.write "           </tr>"
response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
response.write "                   State:"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""text"" value=""" & lcl_default_state & """ name=""egov_users_userstate"" id=""egov_users_userstate"" size=""5"" maxlength=""10"" />"
response.write "               </td>"
response.write "           </tr>"
response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
response.write "                   ZIP:"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""text"" value=""" & lcl_default_zip & """ name=""egov_users_userzip"" id=""egov_users_userzip"" size=""10"" maxlength=""15"" />"
response.write "               </td>"
response.write "           </tr>"
response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
response.write "                   Business Name:"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""text"" value="""" name=""egov_users_userbusinessname"" id=""egov_users_userbusinessname"" size=""50"" maxlength=""100"" />"
response.write "               </td>"
response.write "           </tr>"

bHasBusinessStreets = HasResidentTypeStreets( "B" )
bFound = False
lcl_street_label = "Business Street:"

if bHasBusinessStreets then
	lcl_street_label = "Street (if not listed):"
	response.write "           <tr>"
	response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
	response.write "                   Business Street:"
	response.write "               </td>"
	response.write "               <td>"
	DisplayAddresses "B"
	response.write "               </td>"
	response.write "           </tr>"
end if

response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">" & lcl_street_label & "</td>"
response.write "               <td>"
response.write "                   <input type=""text"" value="""" name=""egov_users_userbusinessaddress"" id=""egov_users_userbusinessaddress"" size=""50"" maxlength=""100"" />"
response.write "               </td>"
response.write "           </tr>"
response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
response.write "                   Work Phone:"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""hidden"" value="""" name=""egov_users_userworkphone"" id=""egov_users_userworkphone"" />"
response.write "                  (<input type=""text"" value="""" name=""skip_work_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />)&nbsp;"
response.write "                   <input type=""text"" value="""" name=""skip_work_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />&ndash;"
response.write "                   <input type=""text"" value="""" name=""skip_work_line"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" />&nbsp;"
response.write "                   ext. <input type=""text"" value="""" name=""skip_work_ext"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" />"
response.write "               </td>"
response.write "           </tr>"
response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
response.write "                   Emergency Contact:"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""text"" value=""" & request.form("egov_users_emergencycontact") & """ name=""egov_users_emergencycontact"" id=""egov_users_emergencycontact"" style=""width:300px;"" maxlength=""100"" />"
response.write "               </td>"
response.write "           </tr>"
response.write "           <tr>"
response.write "               <td class=""label"" align=""right"" nowrap=""nowrap"">"
response.write "                   Emergency Phone:"
response.write "               </td>"
response.write "               <td>"
response.write "                   <input type=""hidden"" value="""" name=""egov_users_emergencyphone"" />"
response.write "                  (<input type=""text"" value="""" name=""skip_emergencyphone_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />)&nbsp;"
response.write "                   <input type=""text"" value="""" name=""skip_emergencyphone_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />&ndash;"
response.write "                   <input type=""text"" value="""" name=""skip_emergencyphone_line"" size=""4"" maxlength=""4"" />"
response.write "               </td>"
response.write "           </tr>"

'Do Not Knock List Options
if lcl_orghasfeature_donotknock then
	response.write "            <tr>"
	response.write "                <td colspan=""2"">"
	response.write "                    <p>"
	response.write "                    <fieldset>"
	response.write "                      <legend><strong>""Do Not Knock"" List(s)&nbsp;</strong></legend>"
	response.write "                      <p>"
	if session("orgid") <> "56" then
	response.write "                        <input type=""checkbox"" name=""skip_isOnDoNotKnockList_peddlers"" id=""skip_isOnDoNotKnockList_peddlers"" value=""on"""     & lcl_checked_isOnDoNotKnockList_peddlers   & " />&nbsp;Is On Do Not Knock List - Peddlers<br />"
	else
	response.write "                        <input type=""hidden"" name=""skip_isOnDoNotKnockList_peddlers"" id=""skip_isOnDoNotKnockList_peddlers"""     & lcl_checked_isOnDoNotKnockList_peddlers   & " />"
	end if
	response.write "                        <input type=""checkbox"" name=""skip_isOnDoNotKnockList_solicitors"" id=""skip_isOnDoNotKnockList_solicitors"" value=""on""" & lcl_checked_isOnDoNotKnockList_solicitors & " />&nbsp;Is On Do Not Knock List - Solicitors"
	response.write "                      </p>"
	response.write "                    </fieldset>"
	response.write "                    </p>"
	response.write "                </td>"
	response.write "            </tr>"

	if session("orgid") <> "56" then
	'Do Not Knock "Is Vendor" Options
	response.write "            <tr>"
	response.write "                <td colspan=""2"">"
	response.write "                    <p>"
	response.write "                    <fieldset>"
	response.write "                      <legend><strong>""Do Not Knock"" Vendors&nbsp;</strong></legend>"
	response.write "                      <p>"
	response.write "                        <input type=""checkbox"" name=""skip_isDoNotKnockVendor_peddlers"" id=""skip_isDoNotKnockVendor_peddlers"" value=""on"""     & lcl_checked_isDoNotKnockVendor_peddlers   & " />&nbsp;Is a Do Not Knock Vendor - Peddler<br />"
	response.write "                        <input type=""checkbox"" name=""skip_isDoNotKnockVendor_solicitors"" id=""skip_isDoNotKnockVendor_solicitors"" value=""on""" & lcl_checked_isDoNotKnockVendor_solicitors & " />&nbsp;Is a Do Not Knock Vendor - Solicitor"
	response.write "                      </p>"
	response.write "                    </fieldset>"
	response.write "                    </p>"
	response.write "                </td>"
	response.write "            </tr>"
	else
	response.write "                        <input type=""hidden"" name=""skip_isDoNotKnockVendor_peddlers"" id=""skip_isDoNotKnockVendor_peddlers"""     & lcl_checked_isDoNotKnockVendor_peddlers   & " />"
	response.write "                        <input type=""hidden"" name=""skip_isDoNotKnockVendor_solicitors"" id=""skip_isDoNotKnockVendor_solicitors"" value=""on""" & lcl_checked_isDoNotKnockVendor_solicitors & " />"
	end if
end if

'response.write "            <tr>"
'response.write "                <td>&nbsp;</td>"
'response.write "                <td><font color=""#ff0000"">*</font> denotes required fields</td>"
'response.write "            </tr>"

'Display the distribution lists if the org has the feature
if lcl_orghasfeature_subscriptions then
	DisplayMaillists ""
end if

'Display the job postings if the org has the feature
if lcl_orghasfeature_job_postings then
	DisplayMaillists "JOB"
end if

'Display the bid postings if the org has the feature
if lcl_orghasfeature_bid_postings then
	DisplayMaillists "BID"
end if

response.write "           <tr>"
response.write "               <td>&nbsp;</td>"
response.write "               <td><font color=""#ff0000"">*</font> Indicates required fields</td>"
response.write "           </tr>"
response.write "         </table>"
response.write "         </div>"
response.write "         <div style=""font-size:10px; padding-top:1em;"">"
response.write "           <input type=""button"" class=""button"" value=""Create User"" onclick=""doCheck();"" />"
response.write "         </div>"
response.write "     </td>"
response.write "  </tr>"
response.write "  </form>"
response.write "</table>"
response.write "  </div>"
response.write "</div>"
%>
 <!--#Include file="../admin_footer.asp"-->  
<%
response.write vbcrlf & "</body>"
response.write vbcrlf & "</html>"


'------------------------------------------------------------------------------
sub DisplayMaillists( ByVal p_list_type )
	Dim sSql, oRs

	sSql = "SELECT * "
	sSql = sSql & " FROM egov_class_distributionlist "
	sSql = sSql & " WHERE orgid = " & session("orgid")
	sSql = sSql & " AND distributionlistdisplay = 1 "

	if p_list_type <> "" then
		sSql = sSql & " AND distributionlisttype = '" & p_list_type & "'"
	else
		sSql = sSql & " AND (distributionlisttype = '' OR distributionlisttype is null) "
	end if

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	if not oRs.eof then
		if p_list_type = "JOB" then
			lcl_legend_text = "Job Postings"
		elseif p_list_type = "BID" then
			lcl_legend_text = "Bid Postings"
		else
			lcl_legend_text = "Subscriptions"
		end if

		response.write vbcrlf & "<tr>"
		response.write "<td colspan=""2"">"
		response.write "<fieldset id=""registrationsubscription"">"
		response.write "<legend><strong>" & lcl_legend_text & "</strong></legend>"
		response.write vbcrlf & "<p>Check the mailing lists to which they would like to subscribe.</p>"

		do while not oRs.eof
			response.write vbcrlf & "<input name=""maillist"" type=""checkbox"" value=""" & oRs("distributionlistid") & """ /> " & oRs("distributionlistname") & "<br />"
			oRs.movenext
		loop

		response.write "</fieldset>"
		response.write "</td>"
		response.write "</tr>"
	end if

	oRs.close 
	set oRs = nothing 

end sub


'------------------------------------------------------------------------------
' void InsertMailLists iuserid 
'------------------------------------------------------------------------------
Sub InsertMailLists( ByVal iuserid )
	Dim sSql, oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")

	' Insert any subscription picks
	For Each list In request("maillist")
		oCmd.CommandText = "INSERT INTO egov_class_distributionlist_to_user ( userid, distributionlistid ) VALUES ( " & iuserid & ", " & list & ")"
		oCmd.Execute
	Next	

	Set oCmd = Nothing

End Sub


'------------------------------------------------------------------------------
Sub DisplayAddresses( ByVal sResidenttype )
	Dim sSql, oRs

	sSql = "SELECT residentstreetnumber, residentstreetname "
	sSql = sSql & " FROM egov_residentaddresses "
	sSql = sSql & " WHERE orgid = " & session("orgid")
	sSql = sSql & " AND residenttype = '" & sResidenttype & "' "
	sSql = sSql & " ORDER BY sortstreetname, residentstreetprefix, CAST(residentstreetnumber AS int)"
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select name=""skip_" & sResidenttype & "address"" id=""skip_" & sResidenttype & "address"">"
	response.write vbcrlf & "<option value=""0000"">Please select an address...</option>"

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" &  oRs("residentstreetnumber") & " " & oRs("residentstreetname")  & """>" & oRs("residentstreetnumber") & " " & oRs("residentstreetname") & "</option>"
		oRs.MoveNext
	Loop 

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
Function HasResidentTypeStreets( ByVal sResidenttype )
	Dim sSql, oRs

	sSql = "SELECT COUNT(residentaddressid) AS hits "
	sSql = sSql & " FROM egov_residentaddresses "
	sSql = sSql & " WHERE orgid = " & session("orgid")
	sSql = sSql & " AND residenttype = '" & sResidenttype & "'"

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If clng(oRs("hits")) > 0 Then 
		HasResidentTypeStreets = True 
	Else 
		HasResidentTypeStreets = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
Function HasResidentTypes()
	Dim sSql, oRs

	sSql = "SELECT count(resident_type) as hits "
	sSql = sSql & " FROM egov_poolpassresidenttypes "
	sSql = sSql & " WHERE orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN") , 3, 1

	If clng(oRs("hits")) > 0 Then 
		HasResidentTypes = True 
	Else 
		HasResidentTypes = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 

'------------------------------------------------------------------------------
' FUNCTION DisplayResidentTypes( )
'------------------------------------------------------------------------------
Function DisplayResidentTypes( )
	Dim sSql, oRs

	sSql = "SELECT resident_type, description FROM egov_poolpassresidenttypes where orgid = " & session("orgid") & " order by displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	DisplayResidentTypes = "<select name=""skip_egov_users_residenttype"">"	
	DisplayResidentTypes = DisplayResidentTypes &  "<option value=""0"">Please select a resident type...</option>"
		
	Do While Not oRs.EOF 
		DisplayResidentTypes = DisplayResidentTypes & "<option value=""" &  oRs("resident_type") & """"
		DisplayResidentTypes = DisplayResidentTypes & ">" & oRs("description") & "</option>"
		oRs.MoveNext
	Loop

	DisplayResidentTypes = DisplayResidentTypes & "</select>"

	oRs.Close
	Set oRs = Nothing 
	
End Function 


'------------------------------------------------------------------------------
sub NotifyUser( ByVal iOrgID, ByVal iOrgName, ByVal iEgovClientWebSiteURL, ByVal sToAddress, ByVal sPassword )
	Dim objMail2, ErrorCode, lcl_msg, oCdoConf, oCdoMail

	lcl_msg = "An account was created for you to access the e-government features of the " & iOrgName & " web site.  "
	lcl_msg = lcl_msg & "To access your account please go to " & iEgovClientWebSiteURL & "/user_login.asp.  "
	lcl_msg = lcl_msg & "Your username is (you may need to reset your password):"
	lcl_msg = lcl_msg & "Username: " & sToAddress
	'lcl_msg = lcl_msg & "Password: " & sPassword 
	'lcl_msg = lcl_msg & "This is a temporary password so please be sure to change it to something you can remember. "

	if clng(iOrgID) = clng(26) then
		lcl_msg = lcl_msg
		lcl_msg = lcl_msg & "You can use this account to submit action line requests, purchase pool memberships, "
		lcl_msg = lcl_msg & "reserve lodges and sign up for classes.  If you have any questions, "
		lcl_msg = lcl_msg & "please contact us at (513) 891-2424."
	end if

	'Build the email
	lcl_email_from    = iOrgName & " (E-Gov Website) <noreplies@egovlink.com>"
	lcl_email_sendto  = sToAddress
	lcl_email_cc      = ""
	lcl_email_subject = iOrgName & ": Web Site Registration"

	'HTMLBody
	lcl_email_htmlbody = lcl_msg
	lcl_email_htmlbody = BuildHTMLMessage(lcl_email_htmlbody,"N")

	'Send the email
	sendEmail lcl_email_from, lcl_email_sendto, lcl_email_cc, lcl_email_subject, lcl_email_htmlbody, "", "Y"

end sub

'------------------------------------------------------------------------------
sub DisplayLargeAddressList( ByVal p_orgid, ByVal sResidenttype)
 	Dim sSql, oRs, sCompareName

	'Determine if we are to validate the address (with street number AND street name) or only the street name.
	'If this feature "CitizenRegistration_NoValidate_Address" is ENABLED then the org does NOT want to validate the address
	'   with the street number ONLY the street name.
	lcl_streetnumber = ""
	lcl_streetname   = ""
	bFound           = False

	sSql = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSql = sSql & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSql = sSql & " FROM egov_residentaddresses "
	sSql = sSql & " WHERE orgid = " & p_orgid
	sSql = sSql & " AND residenttype = '" & sResidenttype & "' "
	sSql = sSql & " AND residentstreetname IS NOT NULL "
	sSql = sSql & " ORDER BY sortstreetname "

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	if not oRs.eof then
		response.write "<input type=""text"" name=""residentstreetnumber"" id=""residentstreetnumber"" value="""" size=""8"" maxlength=""10"" /> &nbsp; "
		response.write "<select name=""skip_address"" id=""skip_address"">"
		response.write "  <option value=""0000"">Choose street from dropdown...</option>"

		do while not oRs.eof
			sCompareName = ""

			if oRs("residentstreetprefix") <> "" then
				sCompareName = oRs("residentstreetprefix") & " " 
			end if

			sCompareName = sCompareName & oRs("residentstreetname")

			if oRs("streetsuffix") <> "" then
				sCompareName = sCompareName & " "  & oRs("streetsuffix")
			end if

			If oRs("streetdirection") <> "" then
				sCompareName = sCompareName & " "  & oRs("streetdirection")
			End If

			response.write "  <option value=""" & sCompareName & """>" & sCompareName & "</option>"
			oRs.MoveNext
		loop 

		response.write "</select>"
	end if

		oRs.Close
		set oRs = nothing 

end sub

'------------------------------------------------------------------------------
sub updateDoNotKnockValues(p_userid, p_isOnDoNotKnockList_peddlers, p_isOnDoNotKnockList_solicitors, _
                                     p_isDoNotKnockVendor_peddlers, p_isDoNotKnockVendor_solicitors)
	Dim sSql

	sIsOnDoNotKnockList_peddlers   = "0"
	sIsOnDoNotKnockList_solicitors = "0"
	sIsDoNotKnockVendor_peddlers   = "0"
	sIsDoNotKnockVendor_solicitors = "0"

	if p_isOnDoNotKnockList_peddlers = "on" then
		sIsOnDoNotKnockList_peddlers = "1"
	end if

	if p_isOnDoNotKnockList_solicitors = "on" then
		sIsOnDoNotKnockList_solicitors = "1"
	end if

	if p_isDoNotKnockVendor_peddlers = "on" then
		sIsDoNotKnockVendor_peddlers = "1"
	end if

	if p_isDoNotKnockVendor_solicitors = "on" then
		sIsDoNotKnockVendor_solicitors = "1"
	end if

	sSql = "UPDATE egov_users SET "
	sSql = sSql & "  isOnDoNotKnockList_peddlers = "   & sIsOnDoNotKnockList_peddlers
	sSql = sSql & ", isOnDoNotKnockList_solicitors = " & sIsOnDoNotKnockList_solicitors
	sSql = sSql & ", isDoNotKnockVendor_peddlers = "   & sIsDoNotKnockVendor_peddlers
	sSql = sSql & ", isDoNotKnockVendor_solicitors = " & sIsDoNotKnockVendor_solicitors
	sSql = sSql & " WHERE userid = " & p_userid

	set oUpdateDNK = Server.CreateObject("ADODB.Recordset")
	oUpdateDNK.Open sSql, Application("DSN"), 3, 1

	set oUpdateDNK = nothing

end Sub


'------------------------------------------------------------------------------
' void setGenderToNull userid
'------------------------------------------------------------------------------
Sub setGenderToNull( ByVal userid )
	Dim sSql

	sSql = "UPDATE egov_users SET gender = NULL WHERE userid = " & userid

	RunSQLStatement sSql

End Sub


%>
