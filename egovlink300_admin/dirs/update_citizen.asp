<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="citizen_global_functions.asp" //-->
<!-- #include file="../../egovlink300_global/includes/inc_passencryption.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: update_citizen.asp
' AUTHOR: ????
' CREATED: ????
' COPYRIGHT: Copyright 2005 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the creation and editing of citizen users
'
' MODIFICATION HISTORY
' 1.0   ????   ???? ???? - INITIAL VERSION
' 1.1	08/30/2011	Steve Loar - Fixed the case where the large addresses name matches but street number does not, on display of citizen address.
' 1.2	10/05/2011	Steve Loar - Added gender selection pick
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sError, sFirstName, sLastName, sAddress, sCity, sState, sZip, sPhone, sEmail, sFax, sCell, sBusinessName
Dim sDayPhone,sPassword,iUserID, bHasResidentStreets, bFound, sResidentType, sBusinessAddress, bHasBusinessStreets
Dim sRedirect, bHasResidentTypes, sWorkPhone, sEmergencyPhone, sEmergencyContact, iNeighborhoodid, sResidencyVerified
Dim sRegistrationBlocked, sBlockedDate, sInternalNote, sExternalNote, sBlockedAdmin, iFamilyId, sUserUnit
Dim sEmailnotavailable, sBirthdate, sStreetNumber, sStreetName, sHasFamilyMembers
Dim sIsOnDoNotKnockList_peddlers, sIsOnDoNotKnockList_solicitors, sIsDoNotKnocKVendor_peddlers, sIsDoNotKnockVendor_solicitors
Dim sGender, bShowGenderPicks, bFacilityAbuse, sFacilityAbuseNote

sLevel        = "../"  'Override of value from common.asp
sStreetNumber = ""
sStreetName   = ""
lcl_success   = ""

PageDisplayCheck "edit citizens", sLevel	 'In common.asp

If request.serverVariables("request_method") = "POST" Then   'This should be the page saving itself
	'update the egov_users table
	UpdateRecords()
	lcl_success = "SU"

	'If session("orgid") = 26 Then 
	'Send them an email if that is selected
	If LCase(request("skip_notifyuser")) = "on" And request("egov_users_useremail") <> "" Then 
		NotifyUser request("egov_users_useremail"), request("egov_users_userpassword")
	End If 
	'End If 

	'Take them back to where they came from
	If session("RedirectPage") <> "" And session("RedirectLang") <> "Return to Citizen List" Then 
		sRedirect = session("RedirectPage") 
		session("RedirectPage") = ""
		response.redirect sRedirect
		'Else
		'Default them back to the display citizen page
		'response.redirect "display_citizen.asp"
	End If 
End If 

iUserID = CLng(request("userid"))
GetRegisteredUserValues iUserID

if iFamilyId = "" then response.redirect "display_citizen.asp"

'Check this includes/common.asp function to see if they have family members
If UserHasFamilyMembers(iUserId, iFamilyId) Then 
	sHasFamilyMembers = "true"
Else 
	sHasFamilyMembers = "false"
End If 

'Org Features
lcl_orghasfeature_hasfamily              = orghasfeature("hasfamily")
lcl_orghasfeature_residency_verification = orghasfeature("residency verification")
lcl_orghasfeature_large_address_list     = orghasfeature("large address list")
lcl_orghasfeature_registration_blocking  = orghasfeature("registration blocking")
lcl_orghasfeature_permit_setup           = orghasfeature("permit setup")
lcl_orghasfeature_donotknock             = orghasfeature("donotknock")
lcl_orghasfeature_citizenregistration_novalidate_address = orghasfeature("citizenregistration_novalidate_address")
bShowGenderPicks = orgHasFeature( "display gender pick" )

'Check for a screen message
'lcl_success = request("success")
lcl_onload  = ""

If lcl_success <> "" Then 
	lcl_msg    = setupScreenMsg(lcl_success)
	lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
End If 

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
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
<!--

	var winHandle;
	var w = (screen.width - 640)/2;
	var h = (screen.height - 450)/2;

	function FlagFamilyChange()
	{
		document.register.familyaddresschanged.value = "YES";
		//alert("yes");
	}

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
          response.write "				validate();" & vbcrlf
          response.write "}" & vbcrlf
       else
          response.write "validate();" & vbcrlf
       end if
     %>
		} else	{
  			validate();
		}
}

	function checkAddress( sReturnFunction, sSave )
	{
		// Remove any extra spaces
		document.register.residentstreetnumber.value = removeSpaces(document.register.residentstreetnumber.value);
		// check the number for non-numeric values
		var rege = /^\d+$/;
		var Ok = rege.exec(document.register.residentstreetnumber.value);
	 if ( ! Ok )
		{
			alert("The Resident Street Number cannot be blank and must be numeric.");
			setfocus(document.register.residentstreetnumber);
			return false;
		}
		// check that they picked a street name
	 if ( document.register.skip_address.value == '0000')
		{
			alert("Please select a street name from the list first.");
			setfocus(document.register.skip_address);
			return false;
		}
		// This is here because window.open in the Ajax callback routine will not work
		//winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&stnumber=' + document.register.residentstreetnumber.value + '&stname=' + document.register.skip_address.value + '&sCheckType=' + sReturnFunction + 'Validate", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		//self.focus();
		// Fire off Ajax routine
		doAjax('checkaddress.asp', 'stnumber=' + document.register.residentstreetnumber.value + '&stname=' + document.register.skip_address.value, sReturnFunction, 'get', '0');

	}

	function CheckResults( sResults )
	{
		// Process the Ajax CallBack 
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
		winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&stnumber=' + document.register.residentstreetnumber.value + '&stname=' + document.register.skip_address.value + '&sCheckType=' + sReturnFunction + 'Validate", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
	}

	function FinalCheck( sResults )
	{
		// Process the Ajax CallBack 
	 if (sResults == 'FOUND')
		{
			//if(winHandle != null && ! winHandle.closed)
			//{ 
			//	winHandle.close();
			//}
			document.register.egov_users_residenttype.value = 'R';
			validate();
		}
		else
		{
			//winHandle.focus();
			document.register.egov_users_residenttype.value = 'N';
			PopAStreetPicker('FinalCheck', 'yes');
		}
		
	}

	function finalCheckValidate()
	{
		// Fire off Ajax routine - This just breaks the tie to the address window so it will close for validation to continue
		//doAjax('checkduplicatecitizens.asp', 'userlname=' + document.register.egov_users_userlname.value, 'OkToValidate', 'get', '0');
		validate();
	}

	function OkToValidate( sReturn )
	{
		//finish the validation routine 
		validate();
	}

	function validate() {
		var msg="";
		var lcl_false_count = 0;

		//Handle blocking if it exists
		var blockexists = eval(document.getElementById("egov_users_registrationblocked"));
		if(blockexists) {
			  //If they are blocking a user, make sure the notes are filled in
		   if (document.getElementById("egov_users_registrationblocked").checked == true) {
     			 if (document.getElementById("egov_users_blockedinternalnote").value == "") {
             lcl_focus = document.getElementById("egov_users_blockedinternalnote");
             inlineMsg(document.getElementById("egov_users_blockedinternalnote").id,'<strong>Required Field Missing: </strong>Internal Note<br /><span style="color:#800000;">Please include an Internal Note when blocking a user.</span>',10,'egov_users_blockedinternalnote');
             lcl_false_count = lcl_false_count + 1;
         }else{
             clearMsg("egov_users_blockedinternalnote");
     				}

     			 if (document.getElementById("egov_users_blockedexternalnote").value == "") {
             lcl_focus = document.getElementById("egov_users_blockedexternalnote");
             inlineMsg(document.getElementById("egov_users_blockedexternalnote").id,'<strong>Required Field Missing: </strong>External Note<br /><span style="color:#800000;">Please include an External Note when blocking a user.</span>',10,'egov_users_blockedexternalnote');
             lcl_false_count = lcl_false_count + 1;
         }else{
             clearMsg("egov_users_blockedexternalnote");
				     }
  			} else	{
         //Make sure that the notes are cleared. if not blocking
         clearMsg("egov_users_blockedexternalnote");
         clearMsg("egov_users_blockedinternalnote");

     				document.getElementById("egov_users_blockedexternalnote").value      = "";
 								document.getElementById("egov_users_blockedinternalnote").value      = "";
 								document.getElementById("skip_egov_users_blockedexternalnote").value = "";
 								document.getElementById("skip_egov_users_blockedinternalnote").value = "";
  			}
		} else {
			clearMsg("egov_users_blockedexternalnote");
			clearMsg("egov_users_blockedinternalnote");
		}

		//Emergency Phone
	if (document.getElementById("skip_emergencyphone_areacode").value != "" || document.getElementById("skip_emergencyphone_exchange").value != "" || document.getElementById("skip_emergencyphone_line").value != "") {
		var ePhone = "";
		ePhone += document.getElementById("skip_emergencyphone_areacode").value;
		ePhone += document.getElementById("skip_emergencyphone_exchange").value;
		ePhone += document.getElementById("skip_emergencyphone_line").value;

   		 if (ePhone.length < 10) {
          lcl_focus = document.getElementById("skip_emergencyphone_line");
          inlineMsg(document.getElementById("skip_emergencyphone_line").id,'<strong>Invalid Value: </strong>Emergency Phone must be a valid phone number, including area code, or completely blank',10,'skip_emergencyphone_line');
          lcl_false_count = lcl_false_count + 1;
   			} else	{
      				document.getElementById("egov_users_emergencyphone").value = document.getElementById("skip_emergencyphone_areacode").value + document.getElementById("skip_emergencyphone_exchange").value + document.getElementById("skip_emergencyphone_line").value;
      				var rege = /^\d+$/;
      				var Ok = rege.exec(document.getElementById("egov_users_emergencyphone").value);
      			 if ( ! Ok ) {
              lcl_focus = document.getElementById("skip_emergencyphone_line");
              inlineMsg(document.getElementById("skip_emergencyphone_line").id,'<strong>Invalid Value: </strong>Emergency Phone must be a valid phone number, including area code, or completely blank',10,'skip_emergencyphone_line');
              lcl_false_count = lcl_false_count + 1;
          }else{
              clearMsg("skip_emergencyphone_line");
      				}
   			}
		} else {
      clearMsg("skip_emergencyphone_line");
   			document.getElementById("egov_users_emergencyphone").value = '';
		}

		//Work Phone
	 if (document.getElementById("skip_work_areacode").value != "" || document.getElementById("skip_work_exchange").value != "" || document.getElementById("skip_work_line").value != "") {
   			var wPhone = "";
      wPhone += document.getElementById("skip_work_areacode").value;
      wPhone += document.getElementById("skip_work_exchange").value;
      wPhone += document.getElementById("skip_work_line").value;

   		 if (wPhone.length < 10) {
          lcl_focus = document.getElementById("skip_work_line");
          inlineMsg(document.getElementById("skip_work_line").id,'<strong>Invalid Value: </strong>Work Phone must be a valid phone number, including area code, or completely blank',10,'skip_work_line');
          lcl_false_count = lcl_false_count + 1;
   			} else	{
      				document.getElementById("egov_users_userworkphone").value = document.getElementById("skip_work_areacode").value + document.getElementById("skip_work_exchange").value + document.getElementById("skip_work_line").value;
      				var rege = /^\d+$/;
      				var Ok = rege.exec(document.getElementById("egov_users_userworkphone").value);
      			 if ( ! Ok ) {
              lcl_focus = document.getElementById("skip_work_line");
              inlineMsg(document.getElementById("skip_work_line").id,'<strong>Invalid Value: </strong>Work Phone must be a valid phone number, including area code, or completely blank',10,'skip_work_line');
              lcl_false_count = lcl_false_count + 1;
          }else{
              clearMsg("skip_work_line");
      				}
   			}
		} else {
      clearMsg("skip_work_line");
   			document.getElementById("egov_users_userworkphone").value = '';
		}


		//Business Street
		var bexists = eval(document.getElementById("skip_Baddress"));
		if(bexists) {
  			//See if they picked from the business dropdown and put that in the address field 
		   if (document.getElementById("skip_Baddress").selectedIndex > -1) {
     				var belement       = document.getElementById("skip_Baddress");
     				var bselectedvalue = belement.options[belement.selectedIndex].value;

     				//0000 is the first pick that we do not want
     			 if (bselectedvalue != "0000") {
        					document.getElementById("egov_users_userbusinessaddress").value = bselectedvalue;
     				}
  			}
		}

		//Resident Address (small address list)
		var rexists = eval(document.getElementById("skip_Raddress"));
		if(rexists) {
  			//See if they picked from the resident dropdown and put that in the address field 
		   if (document.getElementById("skip_Raddress").selectedIndex > -1) {
     				var relement       = document.getElementById("skip_Raddress");
     				var rselectedvalue = relement.options[relement.selectedIndex].value;

     				//0000 is the first pick that we do not want
     			 if (rselectedvalue != "0000") {
        					document.getElementById("egov_users_useraddress").value = rselectedvalue;
     				}
  			}
		}

 	//Resident Address (large street list)
		var exists = eval(document.getElementById("residentstreetnumber"));
		if(exists) {
 				if (document.getElementById("residentstreetnumber").value != '' ) {
    					// See if they picked from the resident dropdown and put that in the address field 
    					if (document.getElementById("skip_address").selectedIndex > -1) {
        					var element       = document.getElementById("skip_address");
						       var selectedvalue = element.options[element.selectedIndex].value;

       						//  0000 is the first pick that we do not want
       						if (selectedvalue != "0000") {
          							document.getElementById("egov_users_useraddress").value = document.getElementById("residentstreetnumber").value + ' ' + selectedvalue;
   														document.getElementById("egov_users_residenttype").value = "R";
   														bUsedAddressDropdown = true;
   										}
         }
				 }
			}

		//Fax
	 if (document.getElementById("skip_fax_areacode").value != "" || document.getElementById("skip_fax_exchange").value != "" || document.getElementById("skip_fax_line").value != "") {
   			var sFax = "";
      sFax += document.getElementById("skip_fax_areacode").value;
      sFax += document.getElementById("skip_fax_exchange").value;
      sFax += document.getElementById("skip_fax_line").value;

   		 if (sFax.length < 10) {
          lcl_focus = document.getElementById("skip_fax_line");
          inlineMsg(document.getElementById("skip_fax_line").id,'<strong>Invalid Value: </strong>Fax must be a valid phone number, including area code, or completely blank',10,'skip_fax_line');
          lcl_false_count = lcl_false_count + 1;
   			} else	{
      				document.getElementById("egov_users_userfax").value = document.getElementById("skip_fax_areacode").value + document.getElementById("skip_fax_exchange").value + document.getElementById("skip_fax_line").value;
      				var rege = /^\d+$/;
      				var Ok = rege.exec(document.getElementById("egov_users_userfax").value);
      			 if ( ! Ok ) {
              lcl_focus = document.getElementById("skip_fax_line");
              inlineMsg(document.getElementById("skip_fax_line").id,'<strong>Invalid Value: </strong>Fax must be a valid phone number, including area code, or completely blank',10,'skip_fax_line');
              lcl_false_count = lcl_false_count + 1;
          }else{
              clearMsg("skip_fax_line");
      				}
   			}
		} else {
      clearMsg("skip_fax_line");
   			document.getElementById("egov_users_userfax").value = '';
		}

		//Cell Phone
	 if (document.getElementById("skip_cell_areacode").value != "" || document.getElementById("skip_cell_exchange").value != "" || document.getElementById("skip_cell_line").value != "") {
   			var cPhone = "";
      cPhone += document.getElementById("skip_cell_areacode").value;
      cPhone += document.getElementById("skip_cell_exchange").value;
      cPhone += document.getElementById("skip_cell_line").value;

   		 if (cPhone.length < 10) {
          lcl_focus = document.getElementById("skip_cell_line");
          inlineMsg(document.getElementById("skip_cell_line").id,'<strong>Invalid Value: </strong>Cell Phone must be a valid phone number, including area code, or completely blank',10,'skip_cell_line');
          lcl_false_count = lcl_false_count + 1;
   			} else	{
      				document.getElementById("egov_users_usercell").value = document.getElementById("skip_cell_areacode").value + document.getElementById("skip_cell_exchange").value + document.getElementById("skip_cell_line").value;
      				var rege = /^\d+$/;
      				var Ok = rege.exec(document.getElementById("egov_users_usercell").value);
      			 if ( ! Ok ) {
              lcl_focus = document.getElementById("skip_cell_line");
              inlineMsg(document.getElementById("skip_cell_line").id,'<strong>Invalid Value: </strong>Cell Phone must be a valid phone number, including area code, or completely blank',10,'skip_cell_line');
              lcl_false_count = lcl_false_count + 1;
          }else{
              clearMsg("skip_cell_line");
      				}
   			}
		} else {
      clearMsg("skip_cell_line");
   			document.getElementById("egov_users_userfax").value = '';
		}

		//Home Phone
	 if (document.getElementById("skip_user_areacode").value != "" || document.getElementById("skip_user_exchange").value != "" || document.getElementById("skip_user_line").value != "") {
   			var wPhone = "";
      wPhone += document.getElementById("skip_user_areacode").value;
      wPhone += document.getElementById("skip_user_exchange").value;
      wPhone += document.getElementById("skip_user_line").value;

   		 if (wPhone.length < 10) {
          lcl_focus = document.getElementById("skip_user_line");
          inlineMsg(document.getElementById("skip_user_line").id,'<strong>Invalid Value: </strong>Home Phone must be a valid phone number, including area code, or completely blank',10,'skip_user_line');
          lcl_false_count = lcl_false_count + 1;
   			} else	{
      				document.getElementById("egov_users_userhomephone").value = document.getElementById("skip_user_areacode").value + document.getElementById("skip_user_exchange").value + document.getElementById("skip_user_line").value;
      				var rege = /^\d+$/;
      				var Ok = rege.exec(document.getElementById("egov_users_userhomephone").value);
      			 if ( ! Ok ) {
              lcl_focus = document.getElementById("skip_user_line");
              inlineMsg(document.getElementById("skip_user_line").id,'<strong>Invalid Value: </strong>Home Phone must be a valid phone number, including area code, or completely blank',10,'skip_user_line');
              lcl_false_count = lcl_false_count + 1;
          }else{
              clearMsg("skip_user_line");
      				}
   			}
		} else {
      clearMsg("skip_user_line");
   			document.getElementById("egov_users_userhomephone").value = '';
		}

		//They will login so validate the email and password
	 if (document.getElementById("skip_emailnotavailable").checked == false) {
		 /*
   		 if (document.getElementById("egov_users_userpassword").value == "" ) {
          lcl_focus = document.getElementById("egov_users_userpassword");
          inlineMsg(document.getElementById("egov_users_userpassword").id,'<strong>Required Field Missing: </strong>Password',10,'egov_users_userpassword');
          lcl_false_count = lcl_false_count + 1;
      }else{
          clearMsg("egov_users_userpassword");
      }
      */

   		 if (document.getElementById("egov_users_userpassword").value != "" && document.getElementById("skip_userpassword2").value == "" ) {
          lcl_focus = document.getElementById("skip_userpassword2");
          inlineMsg(document.getElementById("skip_userpassword2").id,'<strong>Required Field Missing: </strong>Verify Password',10,'skip_userpassword2');
          lcl_false_count = lcl_false_count + 1;
      }else{
          clearMsg("skip_userpassword2");
      }

   		 if (document.getElementById("egov_users_userpassword").value != document.getElementById("skip_userpassword2").value) {
          lcl_focus = document.getElementById("egov_users_userpassword");
          inlineMsg(document.getElementById("egov_users_userpassword").id,'<strong>Invalid Value: </strong>The Passwords you have entered do not match',10,'egov_users_userpassword');
          lcl_false_count = lcl_false_count + 1;
      }else{
          clearMsg("egov_users_userpassword");
          clearMsg("skip_userpassword2");
      }

      if (document.getElementById("egov_users_useremail").value == "") {
          lcl_focus = document.getElementById("egov_users_useremail");
          inlineMsg(document.getElementById("egov_users_useremail").id,'<strong>Required Field Missing: </strong>Email',10,'egov_users_useremail');
          lcl_false_count = lcl_false_count + 1;
      }else{
          clearMsg("egov_users_useremail");
      }

   			//var rege = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*\.(\w{2}|(com|net|org|edu|mil|gov|biz|us|COM|NET|ORG|EDU|MIL|GOV|BIZ|US))$/;
			var rege = /.+@.+\..+/i;
   			var Ok = rege.test(document.getElementById("egov_users_useremail").value);

   		 if (! Ok) {
          lcl_focus = document.getElementById("egov_users_useremail");
          inlineMsg(document.getElementById("egov_users_useremail").id,'<strong>Invalid Value: </strong> The "Email" must be in a valid format.',10,'egov_users_useremail');
          lcl_false_count = lcl_false_count + 1;
      }else{
          clearMsg("egov_users_useremail");
   			}

  }

		//Resident Type
		var rexists = eval(document.getElementById("skip_egov_users_residenttype"));
	 if (rexists) {
   		 if (document.getElementById("skip_egov_users_residenttype").selectedIndex > -1) {
      				var relement       = document.getElementById("skip_egov_users_residenttype");
      				var rselectedvalue = relement.options[relement.selectedIndex].value;

      				// 0 is the first pick that we do not want
      			 if (rselectedvalue != "0") {
         					document.getElementById("egov_users_residenttype").value = rselectedvalue;
              clearMsg("skip_egov_users_residenttype");
      				} else {
          				msg+="Please select a resident type.\n";
              lcl_focus = document.getElementById("skip_egov_users_residenttype");
              inlineMsg(document.getElementById("skip_egov_users_residenttype").id,'<strong>Required Field Missing: </strong>Resident Type',10,'skip_egov_users_residenttype');
              lcl_false_count = lcl_false_count + 1;
      				}
      }else{
          clearMsg("skip_egov_users_residenttype");
   			}
		}

		//Birthdate
		var birthrege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
		var birthOk   = birthrege.test(document.getElementById("egov_users_birthdate").value);

	 if (document.getElementById("egov_users_birthdate").value != "") {
    	 if (! birthOk ) {
          lcl_focus = document.getElementById("egov_users_birthdate");
          inlineMsg(document.getElementById("egov_users_birthdate").id,'<strong>Invalid Value: </strong> The "Birthdate" must be in date format or blank.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'egov_users_birthdate');
          lcl_false_count = lcl_false_count + 1;
   			} else {
      			 if (isValidDate(document.getElementById("egov_users_birthdate").value ) == false) {
          				msg += "Birth date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again, or leave it blank.";

              lcl_focus = document.getElementById("egov_users_birthdate");
              inlineMsg(document.getElementById("egov_users_birthdate").id,'<strong>Invalid Value: </strong> The "Birthdate" must be in date format or blank.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'egov_users_birthdate');
              lcl_false_count = lcl_false_count + 1;
          }else{
              clearMsg("egov_users_birthdate");
      				}
   			}
	}else{
      clearMsg("egov_users_birthdate");
		}

  //First and Last Names
//  if (document.getElementById("egov_users_userlname").value == "" ) {
//      lcl_focus = document.getElementById("egov_users_userlname");
//      inlineMsg(document.getElementById("egov_users_userlname").id,'<strong>Required Field Missing: </strong>Last Name',10,'egov_users_userlname');
//      lcl_false_count = lcl_false_count + 1;
//  }else{
//      clearMsg("egov_users_userlname");
//  }

//  if (document.getElementById("egov_users_userfname").value == "" ) {
//      lcl_focus = document.getElementById("egov_users_userfname");
//      inlineMsg(document.getElementById("egov_users_userfname").id,'<strong>Required Field Missing: </strong>First Name',10,'egov_users_userfname');
//      lcl_false_count = lcl_false_count + 1;
//  }else{
//      clearMsg("egov_users_userfname");
//  }

  if(lcl_false_count > 0) {
     lcl_focus.focus();
     return false;
  }else{
  			//Set some final things and then submit
  		 if (document.getElementById("egov_users_birthdate").value == "") {
     				document.getElementById("egov_users_birthdate").value = "NULL";
  			}

  		 if (document.getElementById("skip_emailnotavailable").checked == true) {
     				//NULL them out so they save as NULL
     				document.getElementById("egov_users_useremail").value    = "NULL";
     				document.getElementById("egov_users_userpassword").value = "NULL";
  			}

  		 if ((document.getElementById("familyaddresschanged").value == "YES") && (document.getElementById("hasfamilymembers").value == 'true')) {
      				var bCopyToAll = confirm("Copy changes to all family members?")
       		 if ( ! bCopyToAll ) {
         					document.register.familyaddresschanged.value = "NO"
       			}
  			} else {
    			document.getElementById("familyaddresschanged").value = "NO"
  			}

  			document.getElementById("register").submit(); 
  }
}

	function GoBack(sUrl)
	{
		//alert( sUrl);
		location.href='' + sUrl;
	}

	var isNN = (navigator.appName.indexOf("Netscape")!=-1);

	function autoTab(input,len, e, bSetFlag) 
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

		// set the flag to propagate changes to the family menbers.
	 if (bSetFlag)
		{
			document.register.familyaddresschanged.value = "YES";
		}
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

	function FamilyList( sUserId )
	{
		location.href='family_list.asp?userid=' + sUserId;
	}

	function displayScreenMsg(iMsg) 
	{
		if(iMsg!="") 
		{
			document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
			window.setTimeout("clearScreenMsg()", (10 * 1000));
		}
	}

	function clearScreenMsg() 
	{
		document.getElementById("screenMsg").innerHTML = "";
	}


//-->
</script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">
  <% ShowHeader sLevel %>
  <!--#Include file="../menu/menu.asp"--> 
<%
'BEGIN: Body Content ---------------------------------------------------------
if session("RedirectPage") <> "" then
	'response.write "<a href=""javascript:GoBack('" & session("RedirectPage") & "');""><img src='../images/arrow_2back.gif' border=""0"" align='absmiddle'>&nbsp;&nbsp;" & Session("RedirectLang") & "</a>" & vbcrlf
	lcl_returnButton_value   = "<< Back"
	lcl_returnButton_onclick = "GoBack('" & session("RedirectPage") & "');"
else
	'response.write "<div id=""backto"">" & vbcrlf
	'response.write "  <a href=""display_citizen.asp""><img src=""../images/arrow_2back.gif"" border=""0"" align=""absmiddle"">&nbsp;&nbsp;Back to All Citizens Display</a>" & vbcrlf
	'response.write "</div>" & vbcrlf
	lcl_returnButton_value   = "Back to All Citizens Display"
	lcl_returnButton_onclick = "location.href='display_citizen.asp';"
end if

'Determine if the checkbox field(s) are "checked"
lcl_emailnotavailable_checked             = ""
lcl_checked_isOnDoNotKnockList_peddlers   = ""
lcl_checked_isOnDoNotKnockList_solicitors = ""
lcl_checked_isDoNotKnockVendor_peddlers   = ""
lcl_checked_isDoNotKnockVendor_solicitors = ""

if sEmailnotavailable then
	lcl_emailnotavailable_checked = " checked=""checked"""
end if

if sIsOnDoNotKnockList_peddlers then
	lcl_checked_isOnDoNotKnockList_peddlers = " checked=""checked"""
end if

if sIsOnDoNotKnockList_solicitors then
	lcl_checked_isOnDoNotKnockList_solicitors = " checked=""checked"""
end if

if sIsDoNotKnockVendor_peddlers then
	lcl_checked_isDoNotKnockVendor_peddlers = " checked=""checked"""
end if

if sIsDoNotKnockVendor_solicitors then
	lcl_checked_isDoNotKnockVendor_solicitors = " checked=""checked"""
end if

response.write "<div id=""content"">" & vbcrlf
response.write "	 <div id=""centercontent"">" & vbcrlf

response.write "  <form method=""post"" name=""register"" id=""register"" action=""update_citizen.asp"">" & vbcrlf
response.write "    <input type=""hidden"" name=""columnnameid"" id=""columnnameid"" value=""userid"" />" & vbcrlf
response.write "    <input type=""hidden"" name=""egov_users_userregistered"" id=""egov_users_userregistered"" value=""1"" />" & vbcrlf
response.write "    <input type=""hidden"" name=""egov_users_orgid"" id=""egov_users_orgid"" value=""" & session("orgid") & """ />" & vbcrlf
response.write "    <input type=""hidden"" name=""userid"" id=""userid"" value=""" & iUserID & """ />" & vbcrlf
response.write "    <input type=""hidden"" name=""ef:egov_users_userfname-text/req"" id=""ef:egov_users_userfname-text/req"" value=""First name"" />" & vbcrlf
response.write "    <input type=""hidden"" name=""ef:egov_users_userlname-text/req"" id=""ef:egov_users_userlname-text/req"" value=""Last name"" />" & vbcrlf
response.write "    <input type=""hidden"" name=""egov_users_residenttype"" id=""egov_users_residenttype"" value=""N"" />" & vbcrlf
response.write "    <input type=""hidden"" name=""egov_users_familyid"" id=""egov_users_familyid"" value=""" & iFamilyId & """ />" & vbcrlf
response.write "    <input type=""hidden"" name=""familyaddresschanged"" id=""familyaddresschanged"" value=""NO"" />" & vbcrlf
response.write "    <input type=""hidden"" name=""hasfamilymembers"" id=""hasfamilymembers"" value=""" & sHasFamilyMembers & """ />" & vbcrlf
If Not bShowGenderPicks Then 
	response.write vbcrlf & "<input type=""hidden"" id=""egov_users_gender"" name=""egov_users_gender"" value=""N"" />"
End If 

response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" width=""100%"">" & vbcrlf
response.write "  <tr valign=""top"">" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <font size=""+1""><strong>Citizen Registration</strong></font>" & vbcrlf
response.write "          <p><input type=""button"" name=""returnButton"" id=""returnButton"" class=""button"" value=""" & lcl_returnButton_value & """ onclick=""" & lcl_returnButton_onclick & """ /></p>" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td width=""40%"" align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
response.write "  </tr>" & vbcrlf
response.write "  <tr>" & vbcrlf
response.write "      <td colspan=""2"" valign=""top"">" & vbcrlf
displayButtons "TOP", lcl_orghasfeature_hasfamily

response.write "          <div class=""shadow"" id=""registershadow"">" & vbcrlf
response.write "          <table border=""0"" class=""tableadmin"" id=""registertable"" cellpadding=""4"" cellspacing=""0"">" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <th align=""left"">Property</th>" & vbcrlf
response.write "                <th align=""left"">Value</th>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">First Name:" 
								  'buildRequiredFieldLabel "First Name:"
response.write "                </td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""text"" value=""" & sFirstName & """ name=""egov_users_userfname"" id=""egov_users_userfname"" size=""50"" maxlength=""50"" onchange=""clearMsg('egov_users_userfname');"" />" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">Last Name:"
								  'buildRequiredFieldLabel "Last Name:"
response.write "                </td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                  		<input type=""text"" value=""" & sLastName & """ name=""egov_users_userlname"" id=""egov_users_userlname"" size=""50"" maxlength=""50"" onchange=""clearMsg('egov_users_userlname');"" />" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf

If bShowGenderPicks Then 
	response.write vbcrlf & "<tr>"
	response.write "<td class=""label"" align=""right"">"
	response.write "Gender:"
	response.write "</td>"
	response.write "<td>"
	DisplayGenderPicks "egov_users_gender", sGender		' in citizen_global_functions.asp
	response.write "</td>"
	response.write "</tr>"
End If 

response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"">Birthdate:</td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                  		<input type=""text"" value=""" & sBirthdate & """ name=""egov_users_birthdate"" id=""egov_users_birthdate"" size=""10"" maxlength=""10"" onchange=""clearMsg('egov_users_birthdate');"" /> (MM/DD/YYYY)" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td>&nbsp;</td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                 			<input type=""checkbox"" name=""skip_emailnotavailable"" id=""skip_emailnotavailable"" onchange=""clearMsg('egov_users_useremail');clearMsg('egov_users_userpassword');clearMsg('skip_userpassword2');""" & lcl_emailnotavailable_checked & "/> Email Not Available (Citizen Will Not Login)" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" valign=""top"" nowrap=""nowrap"">" & vbcrlf
response.write "                    Email:" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""text"" name=""egov_users_useremail"" id=""egov_users_useremail"" value=""" & sEmail & """ size=""50"" maxlength=""100"" onchange=""clearMsg('egov_users_useremail');"" /><br />" & vbcrlf
response.write "                    <input type=""checkbox"" name=""skip_notifyuser"" id=""skip_notifyuser"" checked=""checked"" onchange=""clearMsg('egov_users_useremail');"" /> Send Registration Change Notification to Citizen" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
response.write "                    Password:" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""password"" placeholder=""Enter a new password"" value=""" &  """ name=""egov_users_userpassword"" id=""egov_users_userpassword"" size=""50"" maxlength=""100"" onchange=""clearMsg('egov_users_userpassword')"" autocomplete=""new-password"" />" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
response.write "                    Verify Password:" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""password"" placeholder=""Verify new password"" value=""" &  """ name=""skip_userpassword2"" id=""skip_userpassword2"" size=""50"" maxlength=""100"" onchange=""clearMsg('skip_userpassword2')"" />" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf

bHasResidentTypes = HasResidentTypes()
bFound            = False

If bHasResidentTypes Then 
	response.write "<tr>" & vbcrlf
	response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">Resident Type:"
	'buildRequiredFieldLabel "Resident Type:"
	response.write "</td>" & vbcrlf
	response.write "<td>" & vbcrlf
	if session("orgid") <> "60" then
		DisplayResidentTypes session("orgid"), sResidentType
	else
		response.write "<input type=""hidden"" name=""skip_egov_users_residenttype"" id=""skip_egov_users_residenttype"" value=""" & sResidentType & """ />"
		select case sResidentType
			case "R"
				response.write "Resident"
			case "N"
				response.write "Non Resident"
			case "U"
				response.write "Unincorp Menlo Park"
			case "B"
				response.write "Business"
			case "E"
				response.write "Employee"

		end select
	end if

	If lcl_orghasfeature_residency_verification Then 
		response.write "&nbsp;<input name=""egov_users_residencyverified"" id=""egov_users_residencyverified"" type=""checkbox""" & sResidencyVerified & " /> Residency Verified" & vbcrlf
	End If 

	response.write "</td>" & vbcrlf
	response.write "</tr>" & vbcrlf
End If 

response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
response.write "                    Home Phone:" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""hidden"" value=""" & sDayPhone          & """ name=""egov_users_userhomephone"" id=""egov_users_userhomephone"" />" & vbcrlf
response.write "                			(<input type=""text"" value="""   & Left(sDayPhone,3)  & """ name=""skip_user_areacode"" id=""skip_user_areacode"" onKeyUp=""return autoTab(this, 3, event, true);"" onchange=""clearMsg('skip_user_line');FlagFamilyChange();"" size=""3"" maxlength=""3"" />)&nbsp;" & vbcrlf
response.write "                			 <input type=""text"" value="""   & Mid(sDayPhone,4,3) & """ name=""skip_user_exchange"" id=""skip_user_exchange"" onKeyUp=""return autoTab(this, 3, event, true);"" onchange=""clearMsg('skip_user_line');FlagFamilyChange();"" size=""3"" maxlength=""3"" />&ndash;" & vbcrlf
response.write "                			 <input type=""text"" value="""   & Right(sDayPhone,4) & """ name=""skip_user_line"" id=""skip_user_line"" onKeyUp=""return autoTab(this, 4, event, true);"" onchange=""clearMsg('skip_user_line');FlagFamilyChange();"" size=""4"" maxlength=""4"" />" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
response.write "                    Cell Phone:" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""hidden"" value=""" & sCell          & """ name=""egov_users_usercell"" id=""egov_users_usercell"" />" & vbcrlf
response.write "                   (<input type=""text"" value="""   & Left(sCell,3)  & """ name=""skip_cell_areacode"" id=""skip_cell_areacode"" onKeyUp=""return autoTab(this, 3, event, false);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_cell_line');"" />)&nbsp;" & vbcrlf
response.write "                    <input type=""text"" value="""   & Mid(sCell,4,3) & """ name=""skip_cell_exchange"" id=""skip_cell_exchange"" onKeyUp=""return autoTab(this, 3, event, false);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_cell_line');"" />&ndash;" & vbcrlf
response.write "                    <input type=""text"" value="""   & Right(sCell,4) & """ name=""skip_cell_line"" id=""skip_cell_line"" onKeyUp=""return autoTab(this, 4, event, false);"" size=""4"" maxlength=""4"" onchange=""clearMsg('skip_cell_line');"" />" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
response.write "                    Fax:" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""hidden"" value=""" & sFax          & """ name=""egov_users_userfax"" id=""egov_users_userfax"" />" & vbcrlf
response.write "                   (<input type=""text"" value="""   & Left(sFax,3)  & """ name=""skip_fax_areacode"" id=""skip_fax_areacode"" onKeyUp=""return autoTab(this, 3, event, false);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_fax_line');"" />)&nbsp;" & vbcrlf
response.write "                    <input type=""text"" value="""   & Mid(sFax,4,3) & """ name=""skip_fax_exchange"" id=""skip_fax_exchange"" onKeyUp=""return autoTab(this, 3, event, false);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_fax_line');"" />&ndash;" & vbcrlf
response.write "                    <input type=""text"" value="""   & Right(sFax,4) & """ name=""skip_fax_line"" id=""skip_fax_line"" onKeyUp=""return autoTab(this, 4, event, false);"" size=""4"" maxlength=""4"" onchange=""clearMsg('skip_fax_line');"" />" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf

bHasResidentStreets = HasResidentTypeStreets("R")
bFound              = False

if bHasResidentStreets then
	if not lcl_orghasfeature_large_address_list then
		'Show all addresses for the city - short address solution
		response.write "            <tr>" & vbcrlf
		response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
		response.write "                    Resident Address:" & vbcrlf
		response.write "                </td>" & vbcrlf
		response.write "                <td>" & vbcrlf
		DisplayAddresses session("orgid"), "R", sAddress, bFound
		response.write "                </td>" & vbcrlf
		response.write "            </tr>" & vbcrlf
	else
		'Show the large address list solution
		response.write "<tr>" & vbcrlf
		response.write "<td class=""label"" align=""right"" valign=""top"" nowrap=""nowrap"">" & vbcrlf
		response.write "Resident Address:" & vbcrlf
		response.write "</td>" & vbcrlf
		response.write "<td>" & vbcrlf
		BreakOutAddress sAddress, sStreetNumber, sStreetName   ' In common.asp
		DisplayLargeAddressList session("orgid"), "R", lcl_orghasfeature_citizenregistration_novalidate_address, sStreetNumber, sStreetName, bFound

		'If this feature is ENABLED then it DISABLES the large address validation and 
		'simply does the form validation.
		if not lcl_orghasfeature_citizenregistration_novalidate_address then
			response.write "<input type=""button"" class=""button"" value=""Validate Address"" onclick=""checkAddress('CheckResults', 'no');"" />" & vbcrlf
		end if
		response.write "                </td>" & vbcrlf
		response.write "            </tr>" & vbcrlf
	end if
end if

lcl_address_label   = "Address:"
lcl_display_address = ""

if bHasResidentStreets then
	lcl_address_label = "Address (if not listed):"
end if

if not bFound then
	lcl_display_address = sAddress
end if

response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & lcl_address_label & "</td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""text"" value=""" & lcl_display_address & """ name=""egov_users_useraddress"" id=""egov_users_useraddress"" onchange=""FlagFamilyChange();"" size=""50"" maxlength=""100"" />" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
response.write "                    Resident Unit:" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""text"" value=""" & sUserUnit & """ name=""egov_users_userunit"" id=""egov_users_userunit"" onchange=""FlagFamilyChange();"" size=""11"" maxlength=""10"" />" & vbcrlf
		
if OrgHasNeighborhoods(session("orgid")) then
	response.write "                </td>" & vbcrlf
	response.write "            </tr>" & vbcrlf
	response.write "            <tr>" & vbcrlf
	response.write "                <td class=""label"" align=""right"">" & vbcrlf
	response.write "                    Neighborhood:" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "                <td>" & vbcrlf
	DisplayNeighborhoods session("orgid"), iNeighborhoodid
else
	response.write "                    <input type=""hidden"" name=""egov_users_neighborhoodid"" id=""egov_users_neighborhoodid"" value=""0"" />" & vbcrlf
end if

response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
response.write "                    City:</td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""text"" value=""" & sCity & """ name=""egov_users_usercity"" id=""egov_users_usercity"" onchange=""FlagFamilyChange();"" size=""50"" maxlength=""100"" />" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
response.write "                    State:" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""text"" value=""" & sState & """ name=""egov_users_userstate"" id=""egov_users_userstate"" onchange=""FlagFamilyChange();"" size=""5"" maxlength=""10"" />" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
response.write "                    ZIP:" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""text"" value=""" & sZip & """ name=""egov_users_userzip"" id=""egov_users_userzip"" onchange=""FlagFamilyChange();"" size=""10"" maxlength=""15"" />" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
response.write "                    Business Name:" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""text"" value=""" & sBusinessName & """ name=""egov_users_userbusinessname"" id=""egov_users_userbusinessname"" size=""50"" maxlength=""100"" />" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf

bHasBusinessStreets = HasResidentTypeStreets("B")
bFound              = False
lcl_street_label    = "Business Street:"
lcl_display_street  = ""

if bHasBusinessStreets then
	lcl_street_label = "Street (if not listed):"

	response.write "            <tr>" & vbcrlf
	response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
	response.write "                    Business Street:" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "                <td>" & vbcrlf
	DisplayAddresses session("orgid"), "B", sBusinessAddress, bFound
	response.write "                </td>" & vbcrlf
	response.write "            </tr>" & vbcrlf
end if

if not bFound then
	lcl_display_street = sBusinessAddress
end if

response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & lcl_street_label & "</td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""text"" value=""" & lcl_display_street & """ name=""egov_users_userbusinessaddress"" id=""egov_users_userbusinessaddress"" size=""50"" maxlength=""100"" />" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
response.write "                    Work Phone:" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""hidden"" value="""    & sWorkPhone           & """ name=""egov_users_userworkphone"" id=""egov_users_userworkphone"" />" & vbcrlf
response.write "                   (<input type=""text"" value="""      & Left(sWorkPhone,3)   & """ name=""skip_work_areacode"" id=""skip_work_areacode"" onKeyUp=""return autoTab(this, 3, event, false);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_work_line');"" />)&nbsp;" & vbcrlf
response.write "                    <input type=""text"" value="""      & Mid(sWorkPhone,4,3)  & """ name=""skip_work_exchange"" id=""skip_work_exchange"" onKeyUp=""return autoTab(this, 3, event, false);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_work_line');"" />&ndash;" & vbcrlf
response.write "                    <input type=""text"" value="""      & Mid(sWorkPhone,7,4)  & """ name=""skip_work_line"" id=""skip_work_line"" onKeyUp=""return autoTab(this, 4, event, false);"" size=""4"" maxlength=""4"" onchange=""clearMsg('skip_work_line');"" />&nbsp;" & vbcrlf
response.write "                    ext. <input type=""text"" value=""" & Mid(sWorkPhone,11,4) & """ name=""skip_work_ext"" id=""skip_work_ext"" onKeyUp=""return autoTab(this, 4, event, false);"" size=""4"" maxlength=""4"" />" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
response.write "                    Emergency Contact:" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""text"" value=""" & sEmergencyContact & """ name=""egov_users_emergencycontact"" id=""egov_users_emergencycontact"" style=""width:300px;"" maxlength=""100"" />" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "            <tr>" & vbcrlf
response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
response.write "                    Emergency Phone:" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "                <td>" & vbcrlf
response.write "                    <input type=""hidden"" value=""" & sEmergencyPhone          & """ name=""egov_users_emergencyphone"" id=""egov_users_emergencyphone"" />" & vbcrlf
response.write "                 		(<input type=""text"" value="""   & Left(sEmergencyPhone,3)  & """ name=""skip_emergencyphone_areacode"" id=""skip_emergencyphone_areacode"" onKeyUp=""return autoTab(this, 3, event, false);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_emergencyphone_line')"" />)&nbsp;" & vbcrlf
response.write "                  		<input type=""text"" value="""   & Mid(sEmergencyPhone,4,3) & """ name=""skip_emergencyphone_exchange"" id=""skip_emergencyphone_exchange"" onKeyUp=""return autoTab(this, 3, event, false);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_emergencyphone_line')"" />&ndash;" & vbcrlf
response.write "                 			<input type=""text"" value="""   & Mid(sEmergencyPhone,7,4) & """ name=""skip_emergencyphone_line"" id=""skip_emergencyphone_line"" onKeyUp=""return autoTab(this, 4, event, false);"" size=""4"" maxlength=""4"" onchange=""clearMsg('skip_emergencyphone_line')"" />" & vbcrlf
response.write "                </td>" & vbcrlf
response.write "            </tr>" & vbcrlf
if orghasfeature("facilities") or orghasfeature("recreation") then
	response.write "            <tr><td class=""label"" colspan=""2""><hr /></td></tr>" & vbcrlf
	response.write "            <tr>" & vbcrlf
	response.write "                <td>&nbsp;</td>" & vbcrlf
	response.write "                <td>" & vbcrlf
	sAbuseFlag = ""
	if bFacilityAbuse then sAbuseFlag = " checked "
	response.write "                    <input type=""checkbox"" name=""egov_users_facilityabuse"" id=""egov_users_facilityabuse""" & sAbuseFlag & " /> Facility/Rental Abuser" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "            </tr>" & vbcrlf
	response.write "            <tr>" & vbcrlf
	response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
	response.write "                    Abuse Notes:" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "                <td>" & vbcrlf
	response.write "                    <textarea name=""egov_users_facilityabusenote"" id=""egov_users_facilityabusenote"" class=""blockednotes"" onchange=""clearMsg('egov_users_facilityabusenote');"">" & sFacilityAbuseNote & "</textarea>" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "            </tr>" & vbcrlf
end if

if lcl_orghasfeature_registration_blocking then
	response.write "            <tr><td class=""label"" colspan=""2""><hr /></td></tr>" & vbcrlf
	response.write "            <tr>" & vbcrlf
	response.write "                <td>&nbsp;</td>" & vbcrlf
	response.write "                <td>" & vbcrlf
	response.write "                    <input type=""checkbox"" name=""egov_users_registrationblocked"" id=""egov_users_registrationblocked""" & sRegistrationBlocked & " /> Block Recreation Purchases" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "            </tr>" & vbcrlf
	response.write "            <tr>" & vbcrlf
	response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
	response.write "                    Blocked Date:" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "                <td>" & sBlockedDate & "</td>" & vbcrlf
	response.write "            </tr>" & vbcrlf
	response.write "            <tr>" & vbcrlf
	response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
	response.write "                    Blocked By:" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "                <td>" & sBlockedAdmin & "</td>" & vbcrlf
	response.write "            </tr>" & vbcrlf
	response.write "            <tr>" & vbcrlf
	response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
	response.write "                    Blocked External Note:" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "                <td>" & vbcrlf
	response.write "                    <input type=""hidden"" value=""" & sExternalNote & """ name=""skip_egov_users_blockedexternalnote"" id=""skip_egov_users_blockedexternalnote"" />" & vbcrlf
	response.write "                    <textarea name=""egov_users_blockedexternalnote"" id=""egov_users_blockedexternalnote"" class=""blockednotes"" onchange=""clearMsg('egov_users_blockedexternalnote');"">" & sExternalNote & "</textarea>" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "            </tr>" & vbcrlf
	response.write "            <tr>" & vbcrlf
	response.write "                <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
	response.write "                    Blocked Internal Note:" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "                <td>" & vbcrlf
	response.write "                    <input type=""hidden"" value=""" & sInternalNote & """ name=""skip_egov_users_blockedinternalnote"" id=""skip_egov_users_blockedinternalnote"" />" & vbcrlf
	response.write "                    <textarea name=""egov_users_blockedinternalnote"" id=""egov_users_blockedinternalnote"" class=""blockednotes"" onchange=""clearMsg('egov_users_blockedinternalnote');"">" & sInternalNote & "</textarea>" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "            </tr>" & vbcrlf
end if

'Do Not Knock List Options
if lcl_orghasfeature_donotknock then
	response.write "            <tr>" & vbcrlf
	response.write "                <td colspan=""2"">" & vbcrlf
	response.write "                    <p>" & vbcrlf
	response.write "                    <fieldset>" & vbcrlf
	response.write "                      <legend><strong>""Do Not Knock"" List(s)&nbsp;</strong></legend>" & vbcrlf
	response.write "                      <p>" & vbcrlf
	if session("orgid") <> "56" then
	response.write "                        <input type=""checkbox"" name=""isOnDoNotKnockList_peddlers"" id=""isOnDoNotKnockList_peddlers"" value=""on"""     & lcl_checked_isOnDoNotKnockList_peddlers   & " />&nbsp;Is On Do Not Knock List - Peddlers<br />" & vbcrlf
	else
	response.write "                        <input type=""hidden"" name=""isOnDoNotKnockList_peddlers"" id=""isOnDoNotKnockList_peddlers"""     & lcl_checked_isOnDoNotKnockList_peddlers   & " />" & vbcrlf
	end if
	response.write "                        <input type=""checkbox"" name=""isOnDoNotKnockList_solicitors"" id=""isOnDoNotKnockList_solicitors"" value=""on""" & lcl_checked_isOnDoNotKnockList_solicitors & " />&nbsp;Is On Do Not Knock List - Solicitors" & vbcrlf
	response.write "                      </p>" & vbcrlf
	response.write "                    </fieldset>" & vbcrlf
	response.write "                    </p>" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "            </tr>" & vbcrlf

	if session("orgid") <> "56" then
	'Do Not Knock "Is Vendor" Options
	response.write "            <tr>" & vbcrlf
	response.write "                <td colspan=""2"">" & vbcrlf
	response.write "                    <p>" & vbcrlf
	response.write "                    <fieldset>" & vbcrlf
	response.write "                      <legend><strong>""Do Not Knock"" Vendors&nbsp;</strong></legend>" & vbcrlf
	response.write "                      <p>" & vbcrlf
	response.write "                        <input type=""checkbox"" name=""isDoNotKnockVendor_peddlers"" id=""isDoNotKnockVendor_peddlers"" value=""on"""     & lcl_checked_isDoNotKnockVendor_peddlers   & " />&nbsp;Is a Do Not Knock Vendor - Peddler<br />" & vbcrlf
	response.write "                        <input type=""checkbox"" name=""isDoNotKnockVendor_solicitors"" id=""isDoNotKnockVendor_solicitors"" value=""on""" & lcl_checked_isDoNotKnockVendor_solicitors & " />&nbsp;Is a Do Not Knock Vendor - Solicitor" & vbcrlf
	response.write "                      </p>" & vbcrlf
	response.write "                    </fieldset>" & vbcrlf
	response.write "                    </p>" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "            </tr>" & vbcrlf
	else
	response.write "                        <input type=""hidden"" name=""isDoNotKnockVendor_peddlers"" id=""isDoNotKnockVendor_peddlers"""     & lcl_checked_isDoNotKnockVendor_peddlers   & " />" & vbcrlf
	response.write "                        <input type=""hidden"" name=""isDoNotKnockVendor_solicitors"" id=""isDoNotKnockVendor_solicitors"" value=""on""" & lcl_checked_isDoNotKnockVendor_solicitors & " />" & vbcrlf
	end if
end if

response.write "            <tr>" & vbcrlf
response.write "                <td>&nbsp;</td>" & vbcrlf
response.write "                <td><font color=""#ff0000"">*</font> denotes required fields</td>" & vbcrlf
response.write "            </tr>" & vbcrlf
response.write "          </table>" & vbcrlf
response.write "          </div>" & vbcrlf
displayButtons "BOTTOM", lcl_orghasfeature_hasfamily
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf
response.write "  </form>" & vbcrlf
response.write "</table>" & vbcrlf
response.write "  </div>" & vbcrlf
response.write "</div>" & vbcrlf

%>
<!--#Include file="../admin_footer.asp"-->  
<%

response.write "</body>" & vbcrlf
response.write "</html>" & vbcrlf



'------------------------------------------------------------------------------
Sub DisplayAddresses( ByVal iOrgID, ByVal sResidenttype, ByVal sAddress, ByRef bFound )
	Dim sSql, oRs

	sSql = "SELECT residentstreetnumber, residentstreetname "
	sSql = sSql & " FROM egov_residentaddresses_list "
	sSql = sSql & " WHERE orgid=" & iOrgID
	sSql = sSql & " AND residenttype='" & sResidenttype & "' "
	sSql = sSql & " ORDER BY sortstreetname, residentstreetprefix, cast(residentstreetnumber as int)"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write "<select name=""skip_" & sResidenttype & "address"" id=""skip_" & sResidenttype & "address"" onchange=""clearMsg('skip_" & sResidenttype & "address');FlagFamilyChange();"">" & vbcrlf
	response.write "  <option value=""0000"">Please select an address...</option>" & vbcrlf

	Do While Not oRs.EOF
		lcl_selected_businessaddress = ""

		If UCase(sAddress) = UCase(oRs("residentstreetnumber") & " " & oRs("residentstreetname")) Then 
			lcl_selected_businessaddress = " selected=""selected"""
			bFound = True
		End If 

		response.write "  <option value=""" &  oRs("residentstreetnumber") & " " & oRs("residentstreetname")  & """" & lcl_selected_businessaddress & ">" & oRs("residentstreetnumber") & " " & oRs("residentstreetname") & "</option>" & vbcrlf

		oRs.MoveNext
	Loop 

	response.write "</select>" & vbcrlf

	oRs.Close
	Set oRs = Nothing 
	
End Sub 



'------------------------------------------------------------------------------
Function HasResidentTypeStreets( ByVal sResidenttype )
	Dim sSql, oRs
	sSql = "SELECT count(residentaddressid) AS hits FROM egov_residentaddresses WHERE orgid = " & session("orgid") & " AND residenttype = '" & sResidenttype & "'"
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN") , 3, 1

	If CLng(oRs("hits")) > CLng(0) Then
		HasResidentTypeStreets = True 
	Else
		HasResidentTypeStreets = False 
	End if
	
	oRs.Close
	Set oRs = Nothing
	
End Function 

'------------------------------------------------------------------------------
Function HasResidentTypes()
	Dim sSql, oRs

	sSql = "SELECT count(resident_type) AS hits FROM egov_poolpassresidenttypes WHERE orgid = " & session("orgid") & ""
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN") , 3, 1

	If clng(oRs("hits")) > 0 Then
		HasResidentTypes = True 
	Else
		HasResidentTypes = False 
	End if
	
	oRs.Close
	Set oRs = Nothing 

End Function 

'------------------------------------------------------------------------------
sub DisplayResidentTypes( ByVal iOrgID, ByVal sResidentType )
	Dim sSql, oRs

	sSql = "SELECT resident_type, description "
	sSql = sSql & " FROM egov_poolpassresidenttypes "
	sSql = sSql & " WHERE orgid=" & iOrgID
	sSql = sSql & " ORDER BY displayorder"

	 set oRs = Server.CreateObject("ADODB.Recordset")
	 oRs.Open sSql, Application("DSN"), 3, 1

	 response.write "<select name=""skip_egov_users_residenttype"" id=""skip_egov_users_residenttype"" onchange=""clearMsg('skip_egov_users_residenttype');"">" & vbcrlf

	 do while not oRs.eof
  			lcl_selected_residenttype = ""

   	 if sResidentType = oRs("resident_type") then
     			lcl_selected_residenttype = " selected=""selected"""
   	 end if

	   	response.write "  <option value=""" &  oRs("resident_type") & """" & lcl_selected_residenttype & ">" & oRs("description") & "</option>" & vbcrlf

   		oRs.movenext
 	loop

 	response.write "</select>" & vbcrlf

 	oRs.close
	 set oRs = nothing 
	
end sub


'------------------------------------------------------------------------------
sub GetRegisteredUserValues( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT * FROM egov_users WHERE orgid = " & session("orgid") & " and  userid = " & iUserID

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	if oRs.eof then
		iUserID = 0
	else
  		sFirstName                     = oRs("userfname")
  		sLastName                      = oRs("userlname")
  		sAddress                       = oRs("useraddress")
  		sState                         = oRs("userstate")
  		sCity                          = oRs("usercity")
  		sZip                           = oRs("userzip")
  		sEmail                         = oRs("useremail")
  		sFax                           = oRs("userfax")
  		sCell                          = oRs("usercell")
  		sBusinessName                  = oRs("userbusinessname")
  		'sPassword                      = oRs("userpassword")
  		sDayPhone                      = oRs("userhomephone")
  		sWorkPhone                     = oRs("userworkphone")
  		sEmergencyContact              = oRs("emergencycontact")
	  	sEmergencyPhone                = oRs("emergencyphone")
 	 	sBirthdate                     = oRs("birthdate")
		sIsOnDoNotKnockList_peddlers   = oRs("isOnDoNotKnockList_peddlers")
		sIsOnDoNotKnockList_solicitors = oRs("isOnDoNotKnockList_solicitors")
		sIsDoNotKnockVendor_peddlers   = oRs("isDoNotKnockVendor_peddlers")
		sIsDoNotKnockVendor_solicitors = oRs("isDoNotKnockVendor_solicitors")
		bFacilityAbuse				   = oRs("FacilityAbuse")
		sFacilityAbuseNote			   = oRs("FacilityAbuseNote")
		If IsNull(oRs("gender")) Then
			sGender = "N"
		Else
			sGender = oRs("gender")
		End If 

		if IsNull(oRs("neighborhoodid")) then
			iNeighborhoodid = 0
		else
			iNeighborhoodid = oRs("neighborhoodid")
		end if

		if IsNull(oRs("residenttype")) OR oRs("residenttype") = "" then
			sResidentType = "R"
		else
			sResidentType = oRs("residenttype")
		end if

		if oRs("residencyverified") Then 
			sResidencyVerified = " checked=""checked"" "
		else
			sResidencyVerified = ""
		end if

		sBusinessAddress = oRs("userbusinessaddress")

		if oRs("registrationblocked") then
			sRegistrationBlocked = " checked=""checked"" "
		else
			sRegistrationBlocked = ""
		end if

		sBlockedDate  = oRs("blockeddate")
		sInternalNote = oRs("blockedinternalnote")
		sExternalNote = oRs("blockedexternalnote")

		if not IsNull(oRs("blockedadminid")) then
			sBlockedAdmin = GetAdminName( oRs("blockedadminid") )
		end if

		if not IsNull(oRs("familyid")) then
			iFamilyId = oRs("familyid")
		else
			iFamilyId = iUserId
		end if

		sUserUnit          = oRs("userunit")
		sEmailnotavailable = oRs("emailnotavailable")
	end if

	oRs.Close
	set oRs = nothing 

end sub


'------------------------------------------------------------------------------
Sub UpdateRecords()
	Dim sSql, iResidencyVerified, iNeighborhoodid, sGender

	if request("egov_users_residencyverified") = "on" Then
		iResidencyVerified = "1"
	else
		iResidencyVerified = "0"
	end if

	if request("egov_users_neighborhoodid") <> "" Then
		iNeighborhoodid = request("egov_users_neighborhoodid")
	else
		iNeighborhoodid = "0"
	end if

	if request("egov_users_birthdate") <> "NULL" Then
		sBirthdate = "'" & request("egov_users_birthdate") & "'"
	else
		sBirthdate = request("egov_users_birthdate")
	end If

	If request("egov_users_gender") <> "M" And request("egov_users_gender") <> "F" Then
		sGender = "NULL"
	Else
		sGender = "'" & DBsafe(request("egov_users_gender")) & "'"
	End If 


	sIsOnDoNotKnockList_peddlers   = "0"
	sIsOnDoNotKnockList_solicitors = "0"
	sIsDoNotKnockVendor_peddlers   = "0"
	sIsDoNotKnockVendor_solicitors = "0"

	if request("isOnDoNotKnockList_peddlers") = "on" then
		sIsOnDoNotKnockList_peddlers = "1"
	end if

	if request("isOnDoNotKnockList_solicitors") = "on" then
		sIsOnDoNotKnockList_solicitors = "1"
	end if

	if request("isDoNotKnockVendor_peddlers") = "on" then
		sIsDoNotKnockVendor_peddlers = "1"
	end if

	
	strDoNotKnockRegDate = ""
	if request("isDoNotKnockVendor_solicitors") = "on" then
		sIsDoNotKnockVendor_solicitors = "1"
		strDoNotKnockRegDate = ",donotknockregdate = '" & now() & "'"
	end if

	sSql = "UPDATE egov_users SET userfname = '" & DBsafe( request("egov_users_userfname") )
	sSql = sSql & "', userlname = '" &  DBsafe( request("egov_users_userlname") )
	sSql = sSql & "', gender = " & sGender
	sSql = sSql & ", useraddress = '" &  DBsafe(request("egov_users_useraddress"))
	sSql = sSql & "', usercity = '" & DBsafe(request("egov_users_usercity"))
	sSql = sSql & "', userstate = '" & DBsafe(request("egov_users_userstate"))
	sSql = sSql & "', userzip = '" & DBsafe(request("egov_users_userzip"))

	if request("egov_users_useremail") = "NULL" Then
		sSql = sSql & "', useremail = NULL"
	else 
		sSql = sSql & "', useremail = '" & request("egov_users_useremail") & "'"
	end if

	sSql = sSql & ", userbusinessname = '" & DBsafe( request("egov_users_userbusinessname") ) & "' "

	'if request("egov_users_userpassword") = "NULL" Then
		'sSql = sSql & ", userpassword = NULL"
	'else 
		'sSql = sSql & ", userpassword = '" & request("egov_users_userpassword") & "'"
	'end if
	if replace(request("egov_users_userpassword")," ","") <> "" then
		sSql = sSql & ", userpassword = NULL, password = '" & createHashedPassword(request("egov_users_userpassword")) & "'"
	end if

	If LCase(request("egov_users_facilityabuse")) = "on" Then 
		sAbuseFlag = "1"
	Else
		sAbuseFlag = "0"
	End If 

	residenttype = request("egov_users_residenttype")
	if session("orgid") = "60" then
		if request.form("residentstreetnumber") = "701" and request.form("skip_address") = "LAUREL ST" and lcase(request.form("egov_users_usercity")) = "menlo park" _
			and lcase(request.form("egov_users_userstate")) = "ca" and left(request.form("egov_users_userzip"),5) = "94025" then
			residenttype = "E"
		elseif residenttype = "R" then
		elseif request.form("egov_users_userbusinessname") <> "" and (request.form("skip_Baddress") <> "0000" OR request.form("egov_users_userbusinessaddress") <> "") then
			residenttype = "B"
		elseif left(request.form("egov_users_userzip"),5) = "94025" and lcase(request.form("egov_users_usercity")) = "menlo park" then
			residenttype = "U"
		else
			residenttype = "N"
		end if
	end if
	

	sSql = sSql & ", residenttype = '" & residenttype
	sSql = sSql & "', userhomephone = '" & request("egov_users_userhomephone")
	sSql = sSql & "', userworkphone = '" & request("egov_users_userworkphone")
	sSql = sSql & "', userfax = '" & request("egov_users_userfax")
	sSql = sSql & "', usercell = '" & request("egov_users_usercell")
	sSql = sSql & "', userbusinessaddress = '" & dbsafe(request("egov_users_userbusinessaddress"))
	sSql = sSql & "', emergencycontact = '" & dbsafe(request("egov_users_emergencycontact"))
	sSql = sSql & "', emergencyphone = '" & request("egov_users_emergencyphone")
	sSql = sSql & "', neighborhoodid = " & iNeighborhoodid
	sSql = sSql & ", birthdate = " & sBirthdate
	sSql = sSql & ", residencyverified = " & iResidencyVerified
	sSql = sSql & ", facilityabuse = '" & sAbuseFlag & "'"
	sSql = sSql & ", facilityabusenote = '" & DBSafe(request("egov_users_facilityabusenote")) & "'"
	sSql = sSql & ", blockedinternalnote = '" & DBsafe(request("egov_users_blockedinternalnote")) & "'"
	sSql = sSql & ", blockedexternalnote = '" & DBsafe(request("egov_users_blockedexternalnote")) & "'"

	if request("egov_users_registrationblocked") = "on" then
		sSql = sSql & ", registrationblocked = 1"

		if request("egov_users_blockedinternalnote") <> request("skip_egov_users_blockedinternalnote") _
			Or request("egov_users_blockedexternalnote") <> request("skip_egov_users_blockedexternalnote") then
			'Notes have changed, or are new
			sSql = sSql & ", blockedadminid = " & Session("UserId")
			sSql = sSql & ", blockeddate = '" & Date() & "' "
		end if
	else
		sSql = sSql & ", registrationblocked = 0, blockeddate = NULL, blockedadminid = NULL "
	end if

	sSql = sSql & ", userunit = '" & dbsafe(request("egov_users_userunit")) & "'"

	if request("skip_emailnotavailable") = "on" then
		sSql = sSql & ", emailnotavailable = 1" 
	else
		sSql = sSql & ", emailnotavailable = 0" 
	end if

	sSql = sSql & ", isOnDoNotKnockList_peddlers = "   & sIsOnDoNotKnockList_peddlers
	sSql = sSql & ", isOnDoNotKnockList_solicitors = " & sIsOnDoNotKnockList_solicitors
	sSql = sSql & ", isDoNotKnockVendor_peddlers = "   & sIsDoNotKnockVendor_peddlers
	sSql = sSql & ", isDoNotKnockVendor_solicitors = " & sIsDoNotKnockVendor_solicitors
	sSql = sSql & strDoNotKnockRegDate
	sSql = sSql & " WHERE userid = " & request("userid") 

	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql

	sSql = "UPDATE egov_users "
	sSql = sSql & " SET residencyverified = " & iResidencyVerified
	sSql = sSql & " WHERE familyid = " & request("egov_users_familyid")

	RunSQLStatement sSql

	sSql = "UPDATE egov_familymembers "
	sSql = sSql & " SET firstname = '" & DBsafe( request("egov_users_userfname") ) & "', "
	sSql = sSql & " lastname = '" &      DBsafe( request("egov_users_userlname") ) & "', "
	sSql = sSql & " birthdate = " &      sBirthdate
	sSql = sSql & " WHERE userid = " & request("userid") 

	RunSQLStatement sSql


	if request("familyaddresschanged") = "YES" Then
		sSql = "Update egov_users Set userhomephone = '" & request("egov_users_userhomephone")
		sSql = sSql & "', residenttype = '" & residenttype
		sSql = sSql & "', useraddress = '" &  DBsafe(request("egov_users_useraddress"))
		sSql = sSql & "', usercity = '" & DBsafe(request("egov_users_usercity"))
		sSql = sSql & "', userstate = '" & DBsafe(request("egov_users_userstate"))
		sSql = sSql & "', userzip = '" & DBsafe(request("egov_users_userzip"))
		sSql = sSql & "',  userunit = '" & dbsafe(request("egov_users_userunit")) 
		sSql = sSql & "' WHERE familyid = " & request("egov_users_familyid")
		RunSQLStatement sSql
	end if

	If lcl_orghasfeature_permit_setup Then 

		' Update any permit applicants and Primary Contacts where the permit is still open. Pull them, then loop through the set
		sSql = "SELECT P.permitid, C.permitcontactid FROM egov_permits P, egov_permitcontacts C, egov_permitstatuses S "
		sSql = sSql & " WHERE P.permitid = C.permitid AND P.permitstatusid = S.permitstatusid AND (isapplicant = 1 OR isprimarycontact = 1) "
		sSql = sSql & " AND S.iscompletedstatus = 0 AND S.cansavechanges = 1 AND S.changespropagate = 1 AND C.userid = " & request("userid") 

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		Do While Not oRs.EOF
			sSql = "UPDATE egov_permitcontacts SET firstname = '" & dbsafe(request("egov_users_userfname")) & "' "
			sSql = sSql & ", lastname = '" & dbsafe(request("egov_users_userlname")) & "' "
			if request("egov_users_userbusinessname") = "" Then 
				sSql = sSql & ", company = NULL "
			else 
				sSql = sSql & ", company = '" & dbsafe(request("egov_users_userbusinessname")) & "' "
			end if
			if request("egov_users_useraddress") = "" Then 
				sSql = sSql & ", address = NULL "
			else
				sSql = sSql & ", address = '" & dbsafe(request("egov_users_useraddress")) & "' "
			end if
			if request("egov_users_usercity") = "" Then 
				sSql = sSql & ", city = NULL "
			else
				sSql = sSql & ", city = '" & dbsafe(request("egov_users_usercity")) & "' "
			end if
			if request("egov_users_userstate") = "" Then 
				sSql = sSql & ", state = NULL "
			else
				sSql = sSql & ", state = '" & dbsafe(request("egov_users_userstate")) & "' "
			end if
			if request("egov_users_userzip") = "" Then 
				sSql = sSql & ", zip = NULL "
			else
				sSql = sSql & ", zip = '" & dbsafe(request("egov_users_userzip")) & "' "
			end if
			if request("egov_users_useremail") = "" Then 
				sSql = sSql & ", email = NULL "
			else
				sSql = sSql & ", email = '" & dbsafe(request("egov_users_useremail")) & "' "
			End If
			if request("egov_users_userhomephone") = "" Then 
				sSql = sSql & ", phone = NULL "
			else
				sSql = sSql & ", phone = '" & dbsafe(request("egov_users_userhomephone")) & "' "
			End If
			if request("egov_users_userfax") = "" Then 
				sSql = sSql & ", fax = NULL "
			else
				sSql = sSql & ", fax = '" & dbsafe(request("egov_users_userfax")) & "' "
			End If
			if request("egov_users_usercell") = "" Then 
				sSql = sSql & ", cell = NULL "
			else
				sSql = sSql & ", cell = '" & dbsafe(request("egov_users_usercell")) & "' "
			End If
			if request("skip_emailnotavailable") = "on" then
				sSql = sSql & ", emailnotavailable = 1" 
			else
				sSql = sSql & ", emailnotavailable = 0" 
			end if
			if request("egov_users_userpassword") = "" Then 
				sSql = sSql & ", userpassword = NULL "
			else
				sSql = sSql & ", userpassword = NULL, password = '" & createHashedPassword(request("egov_users_userpassword")) & "' "
			End If
			if request("egov_users_residenttype") = "" Then 
				sSql = sSql & ", residenttype = NULL "
			else
				sSql = sSql & ", residenttype = '" & residenttype & "' "
			End If
			if request("egov_users_userworkphone") = "" Then 
				sSql = sSql & ", userworkphone = NULL "
			else
				sSql = sSql & ", userworkphone = '" & dbsafe(request("egov_users_userworkphone")) & "' "
			End If
			if request("egov_users_emergencyphone") = "" Then 
				sSql = sSql & ", emergencyphone = NULL "
			else
				sSql = sSql & ", emergencyphone = '" & dbsafe(request("egov_users_emergencyphone")) & "' "
			End If
			if request("egov_users_neighborhoodid") = "" Then 
				sSql = sSql & ", neighborhoodid = NULL "
			else
				sSql = sSql & ", neighborhoodid = " & request("egov_users_neighborhoodid") 
			End If
			if request("egov_users_userunit") = "" Then 
				sSql = sSql & ", userunit = NULL "
			else
				sSql = sSql & ", userunit = '" & dbsafe(request("egov_users_userunit")) & "' "
			End If
			if request("egov_users_emergencycontact") = "" Then 
				sSql = sSql & ", emergencycontact = NULL "
			else
				sSql = sSql & ", emergencycontact = '" & dbsafe(request("egov_users_emergencycontact")) & "' "
			End If
			if request("egov_users_userbusinessaddress") = "" Then 
				sSql = sSql & ", userbusinessaddress = NULL "
			else
				sSql = sSql & ", userbusinessaddress = '" & dbsafe(request("egov_users_userbusinessaddress")) & "' "
			End If
			sSql = sSql & " WHERE permitid = " & oRs("permitid") & " AND permitcontactid = " & oRs("permitcontactid")
			'response.write sSql & "<br /><br />"

			RunSQLStatement sSql		'In common.asp

			oRs.MoveNext
		Loop

		oRs.Close 
		Set oRs = Nothing 
	End If 

End Sub 


'------------------------------------------------------------------------------
' Sub NotifyUser( sToAddress, sUserName, sPassword )
'------------------------------------------------------------------------------
Sub NotifyUser( sToAddress, sPassword )
	Dim objMail2, ErrorCode, sMsg2, sCityName, sRoot, sDefaultPhone, sDefaultEmail

	'sCityName = GetOrgName( Session("orgid") )
	'sRoot = GetVirtualName( Session("orgid") )
	'sDefaultPhone = GetDefaultPhone( Session("orgid") )
	'sDefaultEmail = GetDefaultEmail( Session("orgid") )

	sMsg2 = "Your account for the e-government features of the " & sCityName & " web site has been changed.  "
	sMsg2 = sMsg2 & "To access your account and view these changes, please go to "

	sSQL = "SELECT wpLive,OrgName, DefaultPhone, DefaultEmail, OrgPublicWebsiteURL,OrgEgovWebsiteURL FROM organizations WHERE orgid = '" & session("orgid") & "'"
	Set oO = Server.CreateObject("ADODB.Recordset")
	oO.Open sSql, Application("DSN"), 3, 1
	sCityName = oO("OrgName")
	sDefaultPhone = oO("DefaultPhone")
	sDefaultEmail = oO("DefaultEmail")
	if not oO("wpLive") then
		sMsg2 = sMsg2 & oO("OrgEgovWebsiteURL") & "/user_login.asp.  "
	else
		sMsg2 = sMsg2 & oO("OrgPublicWebsiteURL") & "login/.  "
	end if
	oO.Close
	Set oO = Nothing

	sMsg2 = sMsg2 & "Your username is (you may need to reset your password):"
	sMsg2 = sMsg2 & vbcrlf & vbcrlf & "Username: " & sToAddress
	'sMsg2 = sMsg2 & vbcrlf & "Password: " & sPassword
	'sMsg2 = sMsg2 & vbcrlf & vbcrlf & "This is a temporary password so please be sure to change it to something you can remember.  "
 if CLng(Session("orgid")) = CLng(26) Then 
		sMsg2 = sMsg2 & vbcrlf & vbcrlf & "Remember that you can use this account to submit action line requests, purchase pool memberships, reserve"
		sMsg2 = sMsg2 & " lodges and sign up for classes.  "
 end if
	sMsg2 = sMsg2 & vbcrlf & vbcrlf & "If you have any questions, please contact us at " & FormatPhone( sDefaultPhone ) & "."


	sendEmail "", sToAddress, "", sCityName & " Web Site Registration Change", "", sMsg2, "Y"
	

End Sub 

'------------------------------------------------------------------------------
Function FormatPhone( ByVal Number )

	if Len(Number) = 10 Then
		FormatPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
	else
		FormatPhone = Number
	End If

End Function

'------------------------------------------------------------------------------
Function GetDefaultPhone( ByVal iOrgId )
	Dim sSql, oName

	sSql = "Select defaultphone from organizations where orgid = " & iOrgId

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSql, Application("DSN"), 0, 1

 if Not oName.EOF Then
		GetDefaultPhone = oName("defaultphone")
 else 
		GetDefaultPhone = ""
 end if

	oName.close
	Set oName = Nothing
End Function 

'------------------------------------------------------------------------------
Function GetDefaultEmail( ByVal iOrgId )
	Dim sSql, oName

	sSql = "Select defaultemail from organizations where orgid = " & iOrgId

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSql, Application("DSN"), 0, 1

 if Not oName.EOF Then
		GetDefaultEmail = oName("defaultemail")
 else 
		GetDefaultEmail = ""
 end if

	oName.close
	Set oName = Nothing
End Function 

'------------------------------------------------------------------------------
Function GetOrgName( ByVal iOrgId )
	Dim sSql, oName

	sSql = "Select orgname from organizations where orgid = " & iOrgId

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSql, Application("DSN"), 0, 1

 if Not oName.EOF Then
		GetOrgName = oName("orgname")
 else 
		GetOrgName = ""
 end if

	oName.close
	Set oName = Nothing

End Function 

'------------------------------------------------------------------------------
Function GetVirtualName( ByVal iorgid )
	Dim sReturnValue, sSql, oRst

	sReturnValue = "UNKNOWN"

	Set oRst = Server.CreateObject("ADODB.Recordset")
	'sSql = "SELECT OrgVirtualSiteName FROM Organizations WHere orgid='" &  iorgid & "'"
	sSql = "SELECT orgegovwebsiteurl FROM Organizations WHere orgid='" &  iorgid & "'"
	oRst.open sSql,Application("DSN"),3,1

 if NOT oRst.EOF THEN
		sReturnValue = Trim(oRst("orgegovwebsiteurl"))
		'response.write Trim(oRst("orgegovwebsiteurl")) & "&nbsp;" & Len(sReturnValue) & "&nbsp;" & InstrRev(sReturnValue,"/")
		'sReturnValue = Mid(sReturnValue,1,(InstrRev(sReturnValue,"/")-1))
	END If
	oRst.close
	Set oRst = Nothing 

	GetVirtualName = sReturnValue
End Function

'------------------------------------------------------------------------------
Function DBsafe( ByVal strDB )

  If Not VarType( strDB ) = 8 Then DBsafe = strDB : Exit Function

  DBsafe = Replace( strDB, "'", "''" )

End Function

'------------------------------------------------------------------------------
'Sub DisplayLargeAddressListOld( ByVal sResidenttype, ByVal sStreetNumber, ByVal sStreetName, ByRef bFound )
'	Dim sSql, oRs

' if Not IsValidAddress( sStreetNumber, sStreetName ) Then   ' In common.asp
'  		sStreetNumber = ""
'		  sStreetName   = ""
'  		bFound        = False 
' end if

'	sSql = "SELECT distinct sortstreetname, residentstreetprefix, residentstreetname "
'	sSql = sSql & " FROM egov_residentaddresses "
' sSql = sSql & " WHERE orgid = " & session( "orgid" )
' sSql = sSql & " AND residenttype = '" & sResidenttype & "' "
'	sSql = sSql & " AND residentstreetname is not null "
' sSql = sSql & " ORDER BY sortstreetname, residentstreetprefix, residentstreetname "
	
'	Set oRs = Server.CreateObject("ADODB.Recordset")
'	oRs.Open sSql, Application("DSN"), 3, 1

' if Not oRs.EOF Then
'  		response.write "<input type=""text"" name=""residentstreetnumber"" value=""" & sStreetNumber & """ onchange=""FlagFamilyChange();"" size=""8"" maxlength=""10"" /> &nbsp; " & vbcrlf
'		  response.write "<select name=""skip_address"" onchange=""FlagFamilyChange();"">" & vbcrlf
'  		response.write "  <option value=""0000"">Choose street from dropdown</option>" & vbcrlf

'  		Do While NOT oRs.EOF 
'	     	response.write "<option value="""  & oRs("residentstreetname") & """"

'    		 if sStreetName = oRs("residentstreetname") Then
'	       		response.write " selected=""selected"" "
'   		   		bFound = True 
'    		 end if

' 	   		response.write " >"
'    			response.write oRs("residentstreetname") & "</option>" & vbcrlf
'	     	oRs.MoveNext
'  		Loop 

'  		response.write "</select>" & vbcrlf
' end if

'	oRs.Close
'	Set oRs = Nothing 

'End Sub 

'------------------------------------------------------------------------------
Sub DisplayLargeAddressList( ByVal p_orgid, ByVal sResidenttype, ByVal iOrgHasFeature_CitizenRegistration_NoValidate_Address, _
                             ByVal sStreetNumber, ByVal sStreetName, ByRef bFound )
	Dim sSql, oRs, sCompareName, lcl_streetnumber, lcl_streetname, bIsValid

	'Determine if we are to validate the address (with street number AND street name) or only the street name.
	'If this feature "CitizenRegistration_NoValidate_Address" is ENABLED then the org does NOT want to validate the address
	'   with the street number ONLY the street name.

	lcl_streetnumber = ""
	lcl_streetname = ""
	bFound = False
	bIsValid = False 

	If iOrgHasFeature_CitizenRegistration_NoValidate_Address Then 
		If IsValidAddress_byStreetName( p_orgid, sStreetName ) Then 
			lcl_streetnumber = sStreetNumber
			sStreetName = sStreetname
			bIsValid = True
		End If 
	Else 
		If isValidAddress( sStreetNumber, sStreetName ) Then 
			lcl_streetnumber = sStreetNumber
			sStreetName = sStreetName
			bIsValid = True
		End If 
	End If 
	'response.write "<h1>" & bIsValid & "</h1>"

	'if NOT IsValidAddress( sStreetNumber, sStreetName ) then  'In common.asp
	' 		sStreetNumber = ""
	'	  sStreetName   = ""
	' 		bFound        = False
	'end if

	sSql = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, "
	sSql = sSql & " residentstreetname, ISNULL(streetsuffix,'') AS streetsuffix, "
	sSql = sSql & " ISNULL(streetdirection,'') AS streetdirection "
	sSql = sSql & " FROM egov_residentaddresses "
	sSql = sSql & " WHERE orgid = " & p_orgid
	sSql = sSql & " AND residenttype = '" & sResidenttype & "' "
	sSql = sSql & " AND residentstreetname IS NOT NULL "
	sSql = sSql & " ORDER BY sortstreetname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write "<input type=""text"" name=""residentstreetnumber"" id=""residentstreetnumber"" value=""" & lcl_streetnumber & """ onchange=""FlagFamilyChange();"" size=""8"" maxlength=""10"" /> &nbsp; " & vbcrlf
		response.write "<select name=""skip_address"" id=""skip_address"" onchange=""FlagFamilyChange();"">" & vbcrlf
		response.write "  <option value=""0000"">Choose street from dropdown...</option>" & vbcrlf

		Do While Not oRs.EOF
			sCompareName = ""

			If trim(oRs("residentstreetprefix")) <> "" Then 
				sCompareName = trim(oRs("residentstreetprefix")) & " " 
			End If 

			sCompareName = sCompareName & trim(oRs("residentstreetname"))

			If trim(oRs("streetsuffix")) <> "" Then 
				sCompareName = sCompareName & " "  & trim(oRs("streetsuffix"))
			End If 

			If trim(oRs("streetdirection")) <> "" Then 
				sCompareName = sCompareName & " "  & trim(oRs("streetdirection"))
			End If 

			lcl_address_selected = ""

			If UCase(sStreetName) = UCase(sCompareName) And bIsValid Then 
				lcl_address_selected = " selected=""selected"""
				bFound = True
			End If 

			response.write "  <option value=""" & sCompareName & """" & lcl_address_selected & ">" & sCompareName & "</option>" & vbcrlf

			oRs.MoveNext
		Loop 

		response.write "</select>" & vbcrlf
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
sub displayButtons( ByVal iLocation, ByVal iOrgHasFeature_hasFamily )

  if iLocation = "BOTTOM" then
     lcl_padding = "padding-top:5px;"
  else
     lcl_padding = "padding-bottom:10px;"
  end if

  'response.write "          <div style=""font-size:10px; padding-bottom:5px;""><img src=""../images/cancel.gif"" align=""absmiddle"" />&nbsp;<a href=""display_citizen.asp"">" & langCancel & "</a>" & vbcrlf
  'response.write "          &nbsp;&nbsp;&nbsp;&nbsp;" & vbcrlf
  'response.write "          <img src=""../images/go.gif"" align=""absmiddle"">&nbsp;<a href=""javascript:doCheck();"">Update</a>" & vbcrlf
  response.write "          <div style=""" & lcl_padding & """>" & vbcrlf
  response.write "            <input type=""button"" name=""cancelButton"" id=""cancelButton"" class=""button"" value=""Cancel"" onclick=""location.href='display_citizen.asp';"" />" & vbcrlf

  if iOrgHasFeature_hasFamily then
     'response.write "&nbsp;&nbsp;<img src=""" & RootPath & "images/newgroup.gif"" width=""16"" height=""16"" align=""absmiddle"" />&nbsp;<a href=""javascript:FamilyList('" & iFamilyId & "');"">View Their Family Members</a>" & vbcrlf
     response.write "<input type=""button"" name=""viewFamilyButton"" id=""viewFamilyButton"" class=""button"" value=""View Their Family Members"" onclick=""FamilyList('" & iFamilyID & "');"" />" & vbcrlf
  end if

  response.write "            <input type=""button"" name=""saveButton"" id=""saveButton"" class=""button"" value=""Save Changes"" onclick=""clearScreenMsg();doCheck();"" />" & vbcrlf
  response.write "          </div>" & vbcrlf

end Sub


%>
