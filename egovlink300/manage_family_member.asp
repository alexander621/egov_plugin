
<!DOCTYPE html>
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="class/classFamily.asp" //-->
<% 
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: manage_family_member.asp
' AUTHOR: Steve Loar
' CREATED: 12/27/2006 - Copied from manage_account.asp
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  family member information management.
'
' MODIFICATION HISTORY
' 1.0   12/27/2006   Steve Loar - Initial code 
' 1.3	10/30/2007	Steve Loar	- Added large address list selection and popup
' 1.4	10/05/2011	Steve Loar - Added gender selection pick
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'PageDisplayCheck "registration", "", iUserId  ' In Common.asp

' USER VALUES
Dim sFirstName,sLastName,sAddress,sCity,sState,sZip,sPhone,sEmail,sFax,sCell,sBusinessName,sDayPhone,sPassword,iUserID
Dim bHasResidentStreets, bFound, sResidenttype, sBusinessAddress, bHasBusinessStreets, sWorkPhone, iNeighborhoodId
Dim sEmergencyContact, sEmergencyPhone, sBirthdate, iFamilyId, iRelationshipId, oFamily, sResidencyVerified, sRelationship
Dim sButtonText, sGender, bShowGenderPicks, bGenderIsRequired, bUserFound

Set oFamily = New classFamily
iRelationshipId = 0

'Check to see if the user is coming from the new ASP.NET Classes/Events section.
'If "yes" then check for the "useridx" cookie (used for ASP.NET).
'If "no" then check for the "userid" cookie (user for ASP).
if request.cookies("useridx") <> "" then
   sCookieUserID              = request.cookies("useridx")
   response.cookies("userid") = sCookieUserID
else
   sCookieUserID = request.cookies("userid")
end If

If sCookieUserID = "" Then
	sCookieUserID = 0
End If 

reqU = request("u")

on error resume next
	reqU = CLng(reqU)
if err.number <> 0 then
	reqU = CLng(0)

end if
on error goto 0

If CLng(reqU) = CLng(0) Then
	' For adding a new family member
	iUserId = CLng(0)
	'iFamilyId = oFamily.GetFamilyId( request.cookies("userid") )
	iFamilyId = oFamily.GetFamilyId(sCookieUserID)
	GetUnRegisteredUserValues iFamilyId 
	sButtonText = "Add This Family Member"
Else
	' For editing an existing family member
	iUserId = CLng(reqU)
	oFamily.GetFamilyId( iUserId )
	bUserFound = GetRegisteredUserValues( iUserId )
	sButtonText = "Update Family Member Information"
	If bUserFound = false Then
		response.redirect "family_list.asp"
	End If 
End If 

sRelationship = GetRelationShip( iRelationshipId )

Set oFamily = Nothing 

bShowGenderPicks = orgHasFeature( iOrgId, "display gender pick" )
bGenderIsRequired = orgHasFeature( iOrgId, "gender required" )

%>

<html lang="en">
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	<meta charset="UTF-8">
	<title>E-Gov Services <%=sOrgName%> - Manage Family Member</title>

	<link rel="stylesheet" href="css/styles.css" />
	<link rel="stylesheet" href="global.css" />
	<link rel="stylesheet" href="css/style_<%=iorgid%>.css" />
	<style>
  fieldset {
     border-radius: 6px;
  }

.address_fieldset {
   border:                1pt solid #808080;
   border-radius:         5px;
   -moz-border-radius:    5px;
   -webkit-border-radius: 5px;
}

#validaddresslist {
   border:                1pt solid #c0c0c0;
   border-radius:         6px;
  	-moz-border-radius:    6px;
   -webkit-border-radius: 6px;
   background-color:   #efefef;
   margin-top:         4px;
}

#validaddresslist legend {
   border:           1pt solid #c0c0c0;
   border-radius:    4px;
  	-moz-border-radius:    4px;
   -webkit-border-radius: 4px;
   background-color: #ffffff;
   color:            #ff0000;
   padding-left:     4px;
   padding-right:    4px;
}

div#addresspicklist {
  border-radius: 6px;
   -moz-border-radius:    6px;
   -webkit-border-radius: 5px;
}

.maintain_url {
   border:           1pt solid #000000;
   background-color: #c0c0c0;
   padding:          4px;
   color:            #000000;
   font-size:        10pt;
   display:          none;
}

.url_displaytext {
   font-size: 10pt;
   color:     #000000;
}

#screenMsg {
   text-align:  right;
   color:       #ff0000;
   font-size:   10pt;
   font-weight: bold;
}
	</style>


  	<script type="text/javascript" src="scripts/jquery-1.9.1.min.js"></script>
	<script src="scripts/modules.js"></script>
	<script src="scripts/easyform.js"></script>
	<script src="scripts/removespaces.js"></script>
	<script src="scripts/setfocus.js"></script>
	<script src="scripts/validationFunctions.js"></script>
	<script src="scripts/ajaxLib.js"></script>

	<script>
	<!--

	var selectedvalue = '0000';
	var winHandle;
	var w = (screen.width - 640)/2;
	var h = (screen.height - 450)/2;

	function openWin2(url, name) 
	{
	  popupWin = window.open(url, name,"resizable,width=500,height=450");
	}

	function UpdateFamily(iUserId)
	{
		location.href='family_members.asp?userid=' + iUserId;
	}

	function doCheck()
	{
		// If they are using the large address feature
		var exists = eval(document.register["residentstreetnumber"]);
		if(exists)
		{
			// If a street number was entered
			if (document.register.residentstreetnumber.value != '')
			{
				checkAddress( 'FinalCheckOLD', 'yes' );
			}
			else
			{
			       	document.register.egov_users_residenttype.value = "N";
				//checkDuplicateCitizens( 'FinalUserCheckFailed' );
				validate();
			}
		}
		else
		{
			//checkDuplicateCitizens( 'FinalUserCheckFailed' );
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
		//winHandle = eval('window.open("includes/addresspicker.asp?saving=' + sSave + '&stnumber=' + document.register.residentstreetnumber.value + '&stname=' + document.register.skip_address.value + '&sCheckType=' + sReturnFunction + '", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		//self.focus();
		// Fire off Ajax routine
		//doAjax('includes/checkaddress.asp', 'stnumber=' + document.register.residentstreetnumber.value + '&stname=' + document.register.skip_address.value, sReturnFunction, 'get', '0');
     		jQuery.get('includes/checkaddress.asp', {
        		stnumber:    document.register.residentstreetnumber.value,
        		stname:      document.register.skip_address.value,
           		addresstype: 'LARGE'
      		}, function(result) {
        		//displayValidAddressList(result);
			eval(sReturnFunction + "(result);");
			
     		});

	}

	function CheckResults( sResults )
	{
		// Process the Ajax CallBack when the validate address button is clicked
		if (sResults == 'FOUND CHECK')
		{
			document.register.egov_users_useraddress.value = '';
       			jQuery('#validaddresslist').hide();
			alert("This is a valid address in the system.");
		}
		else
		{
			//PopAStreetPicker('CheckResults', 'no');
			jQuery('#validaddresslist').show('slow', function() {
       				jQuery.post('includes/checkaddress.asp', {
          				addresstype: 'LARGE',
          				stnumber:   document.register.residentstreetnumber.value,
          				stname:     document.register.skip_address.value,
          				returntype: 'DISPLAY_OPTIONS'
          				}, function(result) {
             					jQuery('#addresspicklist').html(result);
						//jQuery('#validaddresslist').focus();
						document.getElementById('addblock').scrollIntoView();
       					});
     				});
			}
		}
		<%
       'BEGIN: Do Select -----------------------------------------------------
        response.write "function doSelect() {" & vbcrlf
        response.write "  if(jQuery('#stnumber').prop('selectedIndex') < 0) {" & vbcrlf
        'response.write "     inlineMsg(document.getElementById(""stnumber"").id,'<strong>Required Field Missing: </strong> Please select a valid address first.',10,'stnumber');" & vbcrlf
	response.write "	alert(""Required Field Missing: Please select a valid address first."");"
        response.write "     return false;" & vbcrlf
        response.write "  }" & vbcrlf

        'response.write "  clearScreenMsg();" & vbcrlf
        'response.write "  clearMsg('stnumber');" & vbcrlf
        response.write "  jQuery('#residentstreetnumber').val(jQuery('#stnumber').val());" & vbcrlf
        response.write "  jQuery('#egov_users_useraddress').val('');" & vbcrlf
        response.write "  FinalCheck('FOUND SELECT',0);" & vbcrlf
        response.write "}" & vbcrlf
       'END: Do Select -------------------------------------------------------

       'BEGIN: Cancel Pick ---------------------------------------------------
        response.write "function cancelPick() {" & vbcrlf
        'response.write "  clearScreenMsg();" & vbcrlf
        'response.write "  clearMsg('stnumber');" & vbcrlf
        'response.write "  displayValidAddressList('CANCEL');" & vbcrlf
        response.write "     jQuery('#validaddresslist').hide('slow');" & vbcrlf
        response.write "}" & vbcrlf
       'END: Cancel Pick -----------------------------------------------------

       'BEGIN: Do Keep -------------------------------------------------------
        response.write "function doKeep() {" & vbcrlf
        response.write "  var lcl_streetnumber = jQuery('#residentstreetnumber').val();" & vbcrlf
        response.write "  var lcl_streetname   = jQuery('#skip_address').val();" & vbcrlf
        response.write "  var lcl_streetaddress = '';" & vbcrlf

        response.write "  if(lcl_streetnumber != '') {" & vbcrlf
        response.write "     lcl_streetaddress = lcl_streetnumber;" & vbcrlf
        response.write "  }" & vbcrlf

        response.write "  if(lcl_streetname != '') {" & vbcrlf
        response.write "     if(lcl_streetaddress != '') {" & vbcrlf
        response.write "        lcl_streetaddress += ' ';" & vbcrlf
        response.write "        lcl_streetaddress += lcl_streetname;" & vbcrlf
        response.write "     } else {" & vbcrlf
        response.write "        lcl_streetaddress = lcl_streetname;" & vbcrlf
        response.write "     }" & vbcrlf
        response.write "  }" & vbcrlf

        response.write "  jQuery('#egov_users_useraddress').val(lcl_streetaddress);" & vbcrlf
        response.write "  jQuery('#residentstreetnumber').val('');" & vbcrlf
        response.write "  jQuery('#skip_address').val('');" & vbcrlf
        response.write "  jQuery('#skip_address').prop('selectedIndex',0);" & vbcrlf
        response.write "  FinalCheck('FOUND KEEP',0);" & vbcrlf
        response.write "}" & vbcrlf
       'END: Do Keep ---------------------------------------------------------
    'BEGIN: Final Check ---------------------------------------------------
     response.write "function FinalCheck( sResults, iFalseCount ) {" & vbcrlf
     response.write "  if (sResults == 'FOUND CHECK') {" & vbcrlf
     response.write "      jQuery('#validstreet').val('Y');" & vbcrlf
     response.write "      jQuery('#validaddresslist').hide('slow');" & vbcrlf
     'response.write "      enableDisableAddressFields('');" & vbcrlf
     response.write "  } else if (sResults == 'SUBMIT') {" & vbcrlf
     response.write "      if(jQuery('#egov_users_useraddress').val() == '') {" & vbcrlf
     response.write "         var lcl_streetnumber = jQuery('#residentstreetnumber').val();" & vbcrlf
     response.write "         var lcl_streetname   = jQuery('#skip_address').val();" & vbcrlf
     response.write "      }" & vbcrlf

     response.write "      if(iFalseCount > 0) {" & vbcrlf
     response.write "         return false;" & vbcrlf
     response.write "      } else {" & vbcrlf
     'response.write "         document.getElementById(""maintain_dmt_section"").submit();" & vbcrlf
     'response.write "         return true;" & vbcrlf
     response.write "      }" & vbcrlf
     response.write "  }else{" & vbcrlf
     response.write "      if ((sResults == 'FOUND SELECT')||(sResults == 'FOUND KEEP')) {" & vbcrlf
     response.write "           if (sResults == 'FOUND SELECT') {" & vbcrlf
     response.write "               jQuery('#validstreet').val('Y');" & vbcrlf
     response.write "           }else{" & vbcrlf
     response.write "               jQuery('#validstreet').val('N');" & vbcrlf
     response.write "           }" & vbcrlf
     response.write "           jQuery('#validaddresslist').hide('slow');" & vbcrlf
     'response.write "           enableDisableAddressFields('');" & vbcrlf
     response.write "      }else{" & vbcrlf
     response.write "           if(jQuery('#egov_users_useraddress').val() != '') {" & vbcrlf
     response.write "              jQuery('#validaddresslist').hide('slow');" & vbcrlf
     'response.write "              enableDisableAddressFields('');" & vbcrlf
     response.write "           } else {" & vbcrlf
     response.write "              jQuery('#validaddresslist').show('slow');" & vbcrlf
     'response.write "              enableDisableAddressFields('disabled');" & vbcrlf
     response.write "           }" & vbcrlf
     response.write "      }" & vbcrlf
     response.write "  }" & vbcrlf
     response.write "}" & vbcrlf
    'END: Final Check -----------------------------------------------------

		%>

	function PopAStreetPicker( sReturnFunction, sSave )
	{
		// pop up the address picker
		winHandle = eval('window.open("includes/addresspicker.asp?saving=' + sSave + '&stnumber=' + document.register.residentstreetnumber.value + '&stname=' + document.register.skip_address.value + '&sCheckType=' + sReturnFunction + '", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
	}

	function FinalCheckOLD( sResults )
	{
		// Process the Ajax CallBack for the save process
		if (sResults == 'FOUND CHECK')
		{
			finalCheckValidate();
		}
		else
		{
			//PopAStreetPicker('FinalCheck', 'yes');
			CheckResults('');
		}
	}

	function finalCheckValidate()
	{
		// Fire off Ajax routine - This just breaks the tie to the address window so it will close for validation to continue
		//doAjax('includes/checkduplicatecitizens.asp', 'userlname=' + document.register.egov_users_userlname.value, 'OkToValidate', 'get', '0');
		validate();
	}

	function OkToValidate( sReturn )
	{
		//finish the validation routine 
		validate();
	}

	function validate() 
	{
		var msg="";
		
		// check for missing name fields and make sure they are not numbers
		if ( document.register.egov_users_userfname.value == '' )
		{
			msg += "A first name is required.\n";
		}
		else {
			// filter out the numeric bot inputs
			var rege = /^\d+$/;
			var Ok = rege.exec(document.register.egov_users_userfname.value);
			if ( Ok )
			{
				msg += "The first name cannot be a number.\n";
			}
			else {
				if ( document.register.egov_users_userfname.length < 2) {
					msg += "The first name must be more than 1 character.\n";
				}
			}
		}
		
		if ( document.register.egov_users_userlname.value == '' )
		{
			msg += "A last name is required.\n";
		}
		else {
			// filter out the numeric bot inputs
			var rege = /^\d+$/;
			var Ok = rege.exec(document.register.egov_users_userlname.value);
			if ( Ok )
			{
				msg += "The last name cannot be a number.\n";
			}
			else {
				if ( document.register.egov_users_userfname.length < 2) {
					msg += "The first name must be more than 1 character.\n";
				}
			}
		}
			

		// Gender validation if present and is required
<%		If bShowGenderPicks And bGenderIsRequired Then	%>
			if (document.register.egov_users_gender.value == 'N') 
			{
				msg+="Selection of a valid gender is required.\n";
			}
<%		End If		%>

		// set the work phone
		if (document.register.skip_work_areacode.value != "" || document.register.skip_work_exchange.value != "" || document.register.skip_work_line.value != "" || document.register.skip_work_ext.value != "")
		{
			var sPhone = document.register.skip_work_areacode.value + document.register.skip_work_exchange.value + document.register.skip_work_line.value;
			if (sPhone.length < 10)
			{
				msg += "The Work Phone number must be a valid phone number, including area code, or blank\n";
			}
			else
			{
				document.register.egov_users_userworkphone.value = document.register.skip_work_areacode.value + document.register.skip_work_exchange.value + document.register.skip_work_line.value + document.register.skip_work_ext.value;
				var rege = /^\d+$/;
				var Ok = rege.exec(document.register.egov_users_userworkphone.value);
				if ( ! Ok )
				{
					msg += "The Work Phone number must be a valid phone number, including area code, or blank\n";
				}
			}
		}

		// set the fax
		if (document.register.skip_fax_areacode.value != "" || document.register.skip_fax_exchange.value != "" || document.register.skip_fax_line.value != "" )
		{
			var sPhone = document.register.skip_fax_areacode.value + document.register.skip_fax_exchange.value + document.register.skip_fax_line.value;
			if (sPhone.length < 10)
			{
				msg += "The Fax number must be a valid phone number, including area code, or blank\n";
			}
			else
			{
				document.register.egov_users_userfax.value = document.register.skip_fax_areacode.value + document.register.skip_fax_exchange.value + document.register.skip_fax_line.value;
				var rege = /^\d+$/;
				var Ok = rege.exec(document.register.egov_users_userfax.value);
				if ( ! Ok )
				{
					msg += "The Fax number must be a valid phone number, including area code, or blank\n";
				}
			}
		}

		// set the cell phone
		if (document.register.skip_cell_areacode.value != "" || document.register.skip_cell_exchange.value != "" || document.register.skip_cell_line.value != "" )
		{
			var cPhone = document.register.skip_cell_areacode.value + document.register.skip_cell_exchange.value + document.register.skip_cell_line.value;
			if (cPhone.length < 10)
			{
				msg += "The cell phone number must be a valid phone number, including area code, or blank\n";
			}
			else
			{
				document.register.egov_users_usercell.value = document.register.skip_cell_areacode.value + document.register.skip_cell_exchange.value + document.register.skip_cell_line.value;
				var crege = /^\d+$/;
				var cOk = crege.exec(document.register.egov_users_usercell.value);
				if ( ! cOk )
				{
					msg += "The cell phone number must be a valid phone number, including area code, or blank\n";
				}
			}
		}

		// set the Emergency Phone
		if (document.register.skip_emergencyphone_areacode.value != "" || document.register.skip_emergencyphone_exchange.value != "" || document.register.skip_emergencyphone_line.value != "" )
		{
			var sPhone = document.register.skip_emergencyphone_areacode.value + document.register.skip_emergencyphone_exchange.value + document.register.skip_emergencyphone_line.value;
			if (sPhone.length < 10)
			{
				msg += "The Emergency Phone number must be a valid phone number, including area code, or blank\n";
			}
			else
			{
				document.register.egov_users_emergencyphone.value = document.register.skip_emergencyphone_areacode.value + document.register.skip_emergencyphone_exchange.value + document.register.skip_emergencyphone_line.value;
				var rege = /^\d+$/;
				var Ok = rege.exec(document.register.egov_users_emergencyphone.value);
				if ( ! Ok )
				{
					msg += "The Emergency Phone number must be a valid phone number, including area code, or blank\n";
				}
			}
		}

		// set the home phone number
		document.register.egov_users_userhomephone.value = document.register.skip_user_areacode.value + document.register.skip_user_exchange.value + document.register.skip_user_line.value;
		if (document.register.egov_users_userhomephone.value != "" )
		{
			var hPhone = document.register.egov_users_userhomephone.value;
			if (hPhone.length < 10)
			{
				msg += "The home phone number must be a valid phone number, including area code.\n";
			}
			else
			{
				var rege = /^\d+$/;
				var Ok = rege.exec(document.register.egov_users_userhomephone.value);
				if ( ! Ok )
				{
					msg += "The home phone number must be a valid phone number, including area code.\n";
				}
			}
		}
		else
		{
			msg+="The home phone cannot be blank.\n";
		}

		// Handle the birthdate - Required for Children
		var relationshipidexists = eval(document.register["egov_users_relationshipid"]);

		//var birthrege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
		//var birthOk = birthrege.test(document.register.egov_users_birthdate.value);
		
		if (relationshipidexists)
		{
			var sObj = document.getElementById("relationship");
			var sType = sObj.type;
			if (sType != 'hidden')
			{
				// The drop down picks
				//alert(document.register.egov_users_relationshipid.options[document.register.egov_users_relationshipid.selectedIndex].text);
				document.register.skip_egov_users_relationship.value = document.register.egov_users_relationshipid.options[document.register.egov_users_relationshipid.selectedIndex].text;
				if (document.register.egov_users_relationshipid.options[document.register.egov_users_relationshipid.selectedIndex].text == 'Child')
				{
					if (document.register.egov_users_birthdate.value == "")
					{
						msg += "Please input a birth date for this child in the format of MM/DD/YYYY.";
					}
					else
					{
						if (! isValidDate(document.register.egov_users_birthdate.value))
						{
								msg += "Birth date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.";
						}
						else
						{	
							//alert(yearDiff(document.register.egov_users_birthdate.value, '<%=formatdatetime(date(),2)%>'));
							//return;
							if (yearDiff(document.register.egov_users_birthdate.value, '<%=formatdatetime(date(),2)%>') > 125)
							{
								msg += "The birthdate you entered gives an age over 125.\nPlease correct this."
							}
						}
					}
				}
				else
				{
					if (document.register.egov_users_birthdate.value != "")
					{
						if (! isValidDate(document.register.egov_users_birthdate.value))
						{
							msg += "Birth date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again, or leave it blank.";
						}
						else
						{
							if (yearDiff(document.register.egov_users_birthdate.value, '<%=formatdatetime(date(),2)%>') > 125)
							{
								msg += "The birthdate you entered gives an age over 125.\nPlease correct this."
							}
						}
					}
				}
			}
			else
			{
				if (document.register.egov_users_birthdate.value != "")
				{
					if (! isValidDate(document.register.egov_users_birthdate.value))
					{
						msg += "Birth date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again, or leave it blank.";
					}
					else
						{
							if (yearDiff(document.register.egov_users_birthdate.value, '<%=formatdatetime(date(),2)%>') > 125)
							{
								msg += "The birthdate you entered gives an age over 125.\nPlease correct this."
							}
						}
				}
			}
		}
		else
		{
			if (document.register.egov_users_birthdate.value != "")
			{
				if (! isValidDate(document.register.egov_users_birthdate.value))
				{
					msg += "Birth date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again, or leave it blank.";
				}
				else
				{
					if (yearDiff(document.register.egov_users_birthdate.value, '<%=formatdatetime(date(),2)%>') > 125)
					{
						msg += "The birthdate you entered gives an age over 125.\nPlease correct this."
					}
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
					document.register.egov_users_residenttype.value = "B";
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
				selectedvalue = element.options[element.selectedIndex].value;

				//alert( selectedvalue );
				//  0000 is the first pick that we do not want
				if (selectedvalue != "0000")
				{
					document.register.egov_users_useraddress.value = selectedvalue;
					document.register.egov_users_residenttype.value = "R";
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
					selectedvalue = element.options[element.selectedIndex].value;

					//alert( selectedvalue );
					//  0000 is the first pick that we do not want
					if (selectedvalue != "0000")
					{
						document.register.egov_users_useraddress.value = document.register.residentstreetnumber.value + ' ' + selectedvalue;
						document.register.egov_users_residenttype.value = "R";
					}
				}
			}
		}

		if(msg != "")
		{
			msg="Your form could not be submitted for the following reasons.\n\n" + msg;
			alert(msg);
			return;
		}
		else 
		{	
			if (validateForm('register')) 
			{ 
				document.register.submit(); 
				//alert("OK");
			}
		}
	}

	function GoBack(ReturnToURL)
	{
		if (ReturnToURL != "")
		{
			location.href=ReturnToURL;
		}
		else
		{
			history.go(-1);
		}
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

//-->
</script>

</head>

<!--#Include file="include_top.asp"-->

<!--BODY CONTENT-->

<%	RegisteredUserDisplay( "" ) %>


<font class="pagetitle">Family Member Information</font><br />
<br /><br />
<!--<font class="datetagline">Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%> </font><br /><br />-->

<div id="content">
<div id="centercontent">


<a href="javascript:GoBack('family_list.asp')"><img src="images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Return to Family List</a><br /><br />

<div class="box_header4">Family Member Information</div>

<div class="groupSmall2">
	<form name="register" action="update_family_members.asp" method="post">
	<input type="hidden" name="columnnameid" value="userid">
	<input type="hidden" name="userid" value="<%=iuserid%>">
	<input type="hidden" name="egov_users_orgid" value="<%=iorgid%>">
	<!--
	<input type=hidden name="ef:egov_users_useremail-text/req" value="Email Address">
	<input type=hidden name="ef:egov_users_userpassword-text/req" value="Password 1">
	<input type=hidden name="ef:skip_userpassword2-text/req" value="Password 2">
	-->
	<input type="hidden" name="ef:egov_users_userhomephone-text/req/phone" value="Phone Number">
	<input type="hidden" name="ef:egov_users_userfname-text/req" value="First name">
	<input type="hidden" name="ef:egov_users_userlname-text/req" value="Last name">
	<input type="hidden" name="egov_users_residenttype" value="<%=sResidenttype%>">
	<input type="hidden" name="skip_egov_users_relationship" value="<%=sRelationship%>" />
	<input type="hidden" name="egov_users_residencyverified" value="<%=sResidencyVerified%>" />
<%
	If Not bShowGenderPicks Then 
		response.write vbcrlf & "<input type=""hidden"" id=""egov_users_gender"" name=""egov_users_gender"" value=""N"" />"
	End If 
%>

	<table>
<%
		If errormsg <> "" Then
			response.write "<tr><td colspan=2 align=right>" & errormsg & "</td></tr>"
		End If
%>
		<tr><td colspan="2" align="center">
			<input class="actionbtn" type="button" value="<%=sButtonText%>" onClick="javascript:doCheck();">
		</td></tr>

		<tr><td colspan="2"> &nbsp; </td></tr>
		<tr><td class=label align="right">
			<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color="red">*</font></span> 
			First Name:
			</span>
		</td><td>
			<span class="cot-text-emphasized" title="This field is required"> 
			<input type="text" value="<%=sFirstName%>" name="egov_users_userfname" style="width:300px;" maxlength="100">
			</span>
		</td></tr>
		<tr><td class=label align="right">
			<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color="red">*</font></span>
			Last Name:
			</span>
		</td><td>
			<span class="cot-text-emphasized" title="This field is required">
			<input type="text" value="<%=sLastName%>" name="egov_users_userlname" style="width:300px;" maxlength="100">
			</span>
		</td></tr>
<%
		If bShowGenderPicks Then 
			response.write vbcrlf & "<tr>"
			response.write "<td class=""label"" align=""right"">"
			If bGenderIsRequired Then 
				response.write "<span class=""cot-text-emphasized"" title=""This field is required""><span class=""cot-text-emphasized""><font color=""#ff0000"">*</font></span>"
			End If 
			response.write "Gender:"
			If bGenderIsRequired Then 
				response.write "</span>"
			End If 
			response.write "</td>"
			response.write "<td>"
			DisplayGenderPicks "egov_users_gender", sGender		' in common.asp
			response.write "</td>"
			response.write "</tr>"
		End If 
%>
		<tr><td class="label" align="right">
			Birthdate:
		</td><td>
			<input type="text" value="<%=sBirthdate%>" name="egov_users_birthdate" size="10" maxlength="10" /> (MM/DD/YYYY)
		
<%
  'If CLng(iUserId) <> CLng(request.cookies("userid")) Then
		If CLng(iUserId) <> CLng(sCookieUserID) Then
%>
			</td></tr>
			<tr><td class="label" align="right">
				Relationship:
			</td><td>
				<% DisplayRelationships iOrgid, iRelationshipId %>
<%		Else %>
			<input type="hidden" name="egov_users_relationshipid" value="<%=iRelationshipId%>" id="relationship" />
<%		End If %>
		</td></tr>
<%		bHasResidentStreets = HasResidentTypeStreets( iOrgid, "R" )
		bFound = False 
		If bHasResidentStreets  Then 
			If Not OrgHasFeature( iOrgid, "large address list" ) Then %>
			<tr><td class="label" align="right" nowrap="nowrap">
					Resident Street: 
				</td><td>
					<% DisplayAddresses iorgid, "R", sAddress, bFound %>
			</td></tr>
<%			
			Else
			' Show the large address list solution
%>
				<tr><td class="label" align="right" valign="top" nowrap="nowrap" id="addblock">
						Resident Address:
					</td><td nowrap="nowrap">
<%						BreakOutAddress sAddress, sStreetNumber, sStreetName   ' In common.asp
						DisplayLargeAddressList "R", sStreetNumber, sStreetName, bFound %>&nbsp;
						<input type="button" class="button" value="Validate Address" onclick='checkAddress( "CheckResults", "no" );' />
				</td></tr>
<%			End If 
		End If %>
		<tr><td class=label align="right" valign="top">
			<% If bHasResidentStreets Then %>
				Address (if not listed):
			<% Else %>
				Address:
			<% End If %>
		</td><td>
			<input type="hidden" value="<%=sAddress%>" name="skip_old_egov_users_useraddress" />
			<input type="text" value="<%If Not bfound Then %><%=sAddress%><% End If %>" id="egov_users_useraddress" name="egov_users_useraddress" style="width:300px;" maxlength="100">
										<%
              response.write "    <fieldset id=""validaddresslist"">" & vbcrlf
              response.write "      <legend>Invalid Address</legend>" & vbcrlf
              response.write "      <p>The address you entered does not match any in the system. " & vbcrlf
              response.write "      You can select a valid address from the list, or if you are certain the address you entered is correct " & vbcrlf
              response.write "      click the ""Use the address I entered"" button, to continue.</p>" & vbcrlf
              'response.write "      <form name=""frmAddress"" action=""addresspicker.asp"" method=""post"">" & vbcrlf
              response.write "      			<div id=""addresspicklist""></div>" & vbcrlf
              response.write "      			<input type=""button"" name=""validpick"" id=""validpick"" value=""Use the valid address selected"" class=""button"" onclick=""doSelect();"" />" & vbcrlf
              response.write "      			<input type=""button"" name=""invalidpick"" id=""invalidpick"" value=""Use the address I entered"" class=""button"" onclick=""doKeep();"" />" & vbcrlf
              response.write "      			<input type=""button"" name=""cancelpick"" id=""cancelpick"" value=""Cancel"" class=""button"" onclick=""cancelPick();"" />" & vbcrlf
              'response.write "      		</form>" & vbcrlf
              response.write "    </fieldset>" & vbcrlf
	      %>
		</td></tr>

<%		If OrgHasNeighborhoods( iOrgId ) Then %>
			<tr><td class=label align="right">
				<input type="hidden" value="<%=iNeighborhoodId%>" name="skip_old_egov_users_neighborhoodid" />
				Neighborhood:
			</td><td>
				<% DisplayNeighborhoods iOrgid, iNeighborhoodId %>
			</td></tr>
<%		End If %>

		<tr><td class=label align="right">
			City:
		</td><td>
			<input type="hidden" value="<%=sCity%>" name="skip_old_egov_users_usercity" />
			<input type="text" value="<%=sCity%>" name="egov_users_usercity" style="width:300px;" maxlength="40">
		</td></tr>
		<tr><td class=label align="right">
			State / Province:
		</td><td>
			<input type="hidden" value="<%=sState%>" name="skip_old_egov_users_userstate" />
			<input type="text" value="<%=sState%>" name="egov_users_userstate" size="5" maxlength="50">
		</td></tr>
		<tr><td class=label align="right">
			ZIP / Postal Code:
		</td><td>
			<input type="hidden" value="<%=sZip%>" name="skip_old_egov_users_userzip" />
			<input type="text" value="<%=sZip%>" name="egov_users_userzip" size="10" maxlength="10">
		</td></tr>
		<tr><td class=label align="right">
			<font color=red>*</font> Home Phone:
		</td><td>
			<input type="hidden" value="<%=sDayPhone%>" name="egov_users_userhomephone">
			(<input class="phonenum" type="text" value="<%=Left(sDayPhone,3)%>" name="skip_user_areacode" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3">)&nbsp;
			<input class="phonenum" type="text" value="<%=Mid(sDayPhone,4,3)%>" name="skip_user_exchange" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3">&ndash;
			<input class="phonenum" type="text" value="<%=Right(sDayPhone,4)%>" name="skip_user_line" onKeyUp="return autoTab(this, 4, event);" size="4" maxlength="4">
		</td></tr>
		<tr><td class=label align="right">
			Cell Phone:
		</td><td>
			<input type="hidden" value="<%=sCell%>" name="egov_users_usercell">
			(<input class="phonenum" type="text" value="<%=Left(sCell,3)%>" name="skip_cell_areacode" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3">)&nbsp;
			<input class="phonenum" type="text" value="<%=Mid(sCell,4,3)%>" name="skip_cell_exchange" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3">&ndash;
			<input class="phonenum" type="text" value="<%=Right(sCell,4)%>" name="skip_cell_line" onKeyUp="return autoTab(this, 4, event);" size="4" maxlength="4">
		</td></tr>
		<tr><td class=label align="right">
			Fax:
		</td><td>
			<input type="hidden" value="<%=sFax%>" name="egov_users_userfax">
			(<input class="phonenum" type="text" value="<%=Left(sFax,3)%>" name="skip_fax_areacode" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3">)&nbsp;
			<input class="phonenum" type="text" value="<%=Mid(sFax,4,3)%>" name="skip_fax_exchange" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3">&ndash;
			<input class="phonenum" type="text" value="<%=Right(sFax,4)%>" name="skip_fax_line" onKeyUp="return autoTab(this, 4, event);" size="4" maxlength="4">
		</td></tr>
		<!--<tr><td colspan="2" align="center"><strong>If you are not a resident, please include the following.</strong></td><tr>-->
		<tr><td class=label align="right">
			Business Name:
		</td><td>
			<input type="text" value="<%=sBusinessName%>" name="egov_users_userbusinessname" style="width:300px;" maxlength="100">
		</td></tr>
<%		bHasBusinessStreets = HasResidentTypeStreets( iOrgid, "B" )
		bFound = False 
		If bHasBusinessStreets  Then %>
			<tr><td class=label align="right">
					Business Street: 
				</td><td>
					<% DisplayAddresses iorgid, "B", sBusinessAddress, bFound %>
			</td></tr>
<%		End If %>
		<tr><td class=label align="right">
			<% If bHasBusinessStreets Then %>
				Street (if not listed):
			<% Else %>
				Business Street:
			<% End If %>
		</td><td>
			<input type="text" value="<%If Not bfound then
											response.write sBusinessAddress
										End If %>" name="egov_users_userbusinessaddress" style="width:300px;" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			Work Phone:
		</td><td>
			<!--<input type="text" value="<%=sWorkPhone%>" name="egov_users_userworkphone" style="width:300;" maxlength="100">-->
			<input type="hidden" value="<%=sWorkPhone%>" name="egov_users_userworkphone">
			(<input class="phonenum" type="text" value="<%=Left(sWorkPhone,3)%>" name="skip_work_areacode" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3">)&nbsp;
			<input class="phonenum" type="text" value="<%=Mid(sWorkPhone,4,3)%>" name="skip_work_exchange" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3">&ndash;
			<input class="phonenum" type="text" value="<%=Mid(sWorkPhone,7,4)%>" name="skip_work_line" onKeyUp="return autoTab(this, 4, event);" size="4" maxlength="4">&nbsp;
			ext. <input class="phonenum" type="text" value="<%=Mid(sWorkPhone,11,4)%>" name="skip_work_ext" onKeyUp="return autoTab(this, 4, event);" size="4" maxlength="4">
		</td></tr>
		<tr><td class=label align="right">
			Emergency Contact:
		</td><td>
			<input type="text" value="<%=sEmergencyContact%>" name="egov_users_emergencycontact" style="width:300px;" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			Emergency Phone:
		</td><td>
			<input type="hidden" value="<%=sEmergencyPhone%>" name="egov_users_emergencyphone">
			(<input class="phonenum" type="text" value="<%=Left(sEmergencyPhone,3)%>" name="skip_emergencyphone_areacode" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3">)&nbsp;
			<input class="phonenum" type="text" value="<%=Mid(sEmergencyPhone,4,3)%>" name="skip_emergencyphone_exchange" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3">&ndash;
			<input class="phonenum" type="text" value="<%=Mid(sEmergencyPhone,7,4)%>" name="skip_emergencyphone_line" onKeyUp="return autoTab(this, 4, event);" size="4" maxlength="4">
		</td></tr>

		<tr><td colspan="2"> &nbsp; </td></tr>

		<tr><td colspan=2 align="center">
			<input class="actionbtn" type="button" value="<%=sButtonText%>" onClick="javascript:doCheck();">
		</td></tr>

	</table>

	</form>
	</div>

</div></div>
<%
response.write "<script>"
     response.write "jQuery(document).ready(function(){" & vbcrlf
        response.write "  jQuery('#validaddresslist').hide();" & vbcrlf
     response.write "});" & vbcrlf
response.write "</script>"
%>

   
<!--#Include file="include_bottom.asp"--> 

<!--#Include file="includes\inc_dbfunction.asp"-->  

<%

'--------------------------------------------------------------------------------------------------
' void DisplayNeighborhoods iorgid, iNeighborhoodId 
'--------------------------------------------------------------------------------------------------
Sub DisplayNeighborhoods( ByVal iorgid, ByVal iNeighborhoodId )
	Dim sSql, oRs 

	sSql = "SELECT neighborhoodid, neighborhood FROM egov_neighborhoods WHERE orgid = " & iorgid & " ORDER BY neighborhood"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select name=""egov_users_neighborhoodid"">"	
	response.write vbcrlf &  "<option value=""0"">Not on List...</option>"
		
	Do While NOT oRs.EOF 
		response.write vbcrlf & "<option value=""" &  oRs("neighborhoodid") & """"
		If CLng(iNeighborhoodId) = CLng(oRs("neighborhoodid")) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("neighborhood") & "</option>"
		oRs.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub  


'--------------------------------------------------------------------------------------------------
' void DisplayResidentAddresses iorgid, sAddress, bFound 
'--------------------------------------------------------------------------------------------------
Sub DisplayAddresses( ByVal iorgid, ByVal sResidenttype, ByVal sAddress, ByRef bFound )
	Dim sSql, oRs

	sSql = "SELECT residentstreetnumber, residentstreetname FROM egov_residentaddresses_list WHERE orgid = " & iorgid 
	sSql = sSql & " AND residenttype='" & sResidenttype & "' ORDER BY sortstreetname, residentstreetprefix, CAST(residentstreetnumber AS INT)"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write "<select name=""skip_" & sResidenttype & "address"">"	
	response.write "<option value=""0000"">Please select an address...</option>"
		
	Do While Not oRs.EOF 
		response.write vbcrlf & "<option value=""" &  oRs("residentstreetnumber") & " " & oRs("residentstreetname")  & """"
		If UCase(sAddress) = UCase(oRs("residentstreetnumber") & " " & oRs("residentstreetname")) Then
			bFound = True
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("residentstreetnumber") & " " & oRs("residentstreetname") & "</option>"
		oRs.MoveNext
	Loop

	response.write "</select>"

	oRs.Close
	Set oRs = Nothing 
	
End Sub 


'--------------------------------------------------------------------------------------------------
' void DisplayLargeAddressList sResidenttype, sStreetNumber, sStreetName, bFound 
'--------------------------------------------------------------------------------------------------
Sub DisplayLargeAddressList( ByVal sResidenttype, ByVal sStreetNumber, ByVal sStreetName, ByRef bFound )
	Dim sSql, oRs

	If Not IsValidAddress( sStreetNumber, sStreetName ) Then   ' In common.asp
		sStreetNumber = ""
		sStreetName = ""
		bFound = False 
	End If 

	sSql = "SELECT DISTINCT sortstreetname, residentstreetprefix, residentstreetname "
	sSql = sSql & "FROM egov_residentaddresses WHERE orgid = " & iOrgid & " AND residenttype = '" & sResidenttype & "' "
	sSql = sSql & "AND residentstreetname IS NOT NULL ORDER BY sortstreetname, residentstreetprefix, residentstreetname"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<input type=""text"" id=""residentstreetnumber""  name=""residentstreetnumber"" value=""" & sStreetNumber & """ size=""8"" maxlength=""10"" /> &nbsp; "
		response.write vbcrlf & "<select id=""skip_address"" name=""skip_address"">"
		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown</option>"
		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value="""  & oRs("residentstreetname") & """"
			If sStreetName = oRs("residentstreetname") Then
				response.write " selected=""selected"" "
				bFound = True 
			End If 
			response.write " >"
			response.write oRs("residentstreetname") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean HasResidentTypeStreets( iOrgid, sResidenttype )
'--------------------------------------------------------------------------------------------------
Function HasResidentTypeStreets( ByVal iOrgid, ByVal sResidenttype )
	Dim sSql, oRs

	sSql = "SELECT COUNT(residentaddressid) AS hits FROM egov_residentaddresses "
	sSql = sSql & "WHERE orgid = " & iorgid & " AND residenttype = '" & sResidenttype & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If CLng(oRs("hits")) > 0 Then
		HasResidentTypeStreets = True 
	Else
		HasResidentTypeStreets = False 
	End if
	
	oRs.Close
	Set oRs = Nothing
	
End Function 


'--------------------------------------------------------------------------------------------------
' void GetRegisteredUserValues iUserId
'--------------------------------------------------------------------------------------------------
Function GetRegisteredUserValues( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT * FROM egov_users WHERE isdeleted = 0 AND userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetRegisteredUserValues = true 
		sFirstName = oRs("userfname")
		sLastName = oRs("userlname")
		sAddress = oRs("useraddress")
		sState = oRs("userstate")
		sCity = oRs("usercity")
		sZip = oRs("userzip")
		sEmail = oRs("useremail")
		sFax = oRs("userfax")
		sCell = oRs("usercell")
		sBusinessName = oRs("userbusinessname")
		sPassword = oRs("userpassword")
		sDayPhone = oRs("userhomephone")
		sWorkPhone = oRs("userworkphone")
		iUserID = oRs("userid")
		If IsNull(oRs("residenttype")) Or oRs("residenttype") = "" Then
			sResidenttype = "N"
		Else 
			sResidenttype = oRs("residenttype")
		End If 
		sBusinessAddress = oRs("userbusinessaddress")
		If IsNull(oRs("neighborhoodid")) Then 
			iNeighborhoodId = 0
		Else 
			iNeighborhoodId = oRs("neighborhoodid")
		End If 
		sEmergencyContact = oRs("emergencycontact")
		sEmergencyPhone = oRs("emergencyphone")
		iRelationshipId = oRs("relationshipid")
		sBirthdate = oRs("birthdate")
		sResidencyVerified = oRs("residencyverified")
		If IsNull("gender") Then 
			sGender = "N"
		Else
			sGender = oRs("gender")
		End If 
	Else
		GetRegisteredUserValues = false
		iRelationshipId = 0
		sGender = "N"
	End If

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' void GetUnRegisteredUserValues iUserId
'--------------------------------------------------------------------------------------------------
Sub GetUnRegisteredUserValues( ByVal iUserId )
	Dim sSql, oRs

	' This is pulling the head of household values to apply to the family member
	sSql = "SELECT * FROM egov_users WHERE userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sFirstName = ""
		sLastName = oRs("userlname")
		sAddress = oRs("useraddress")
		sState = oRs("userstate")
		sCity = oRs("usercity")
		sZip = oRs("userzip")
		sEmail = ""
		sFax = oRs("userfax")
		sCell = oRs("usercell")
		sBusinessName = ""
		sPassword = ""
		sDayPhone = oRs("userhomephone")
		sWorkPhone = oRs("userworkphone")
		If IsNull(oRs("residenttype")) Or oRs("residenttype") = "" Then
			sResidenttype = "N"
		Else 
			sResidenttype = oRs("residenttype")
		End If 
		sBusinessAddress = ""
		If IsNull(oRs("neighborhoodid")) Then 
			iNeighborhoodId = 0
		Else 
			iNeighborhoodId = oRs("neighborhoodid")
		End If 
		sEmergencyContact = oRs("emergencycontact")
		sEmergencyPhone = oRs("emergencyphone")
		iRelationshipId = 0
		sBirthdate = ""
		sResidencyVerified = oRs("residencyverified")
		sGender = "N"
	Else
		sGender = "N"
		iNeighborhoodId = 0
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void UpdateUserNeighborhood iUserId, iNeighborhoodid 
'--------------------------------------------------------------------------------------------------
Sub UpdateUserNeighborhood( ByVal iUserId, ByVal iNeighborhoodid )
	Dim sSql

	sSql = "Update egov_users SET neighborhoodid = "
	If iNeighborhoodid = "0" Then 
		sSql = sSql & " NULL "
	Else
		sSql = sSql & iNeighborhoodid
	End If 
	sSql = sSql & " WHERE userid = " & iUserId 

	RunSQLStatement sSql 	' In common.asp

End Sub 


'--------------------------------------------------------------------------------------------------
' void DisplayRelationships iOrgid, iRelationshipId 
'--------------------------------------------------------------------------------------------------
Sub DisplayRelationships( ByVal iOrgid, ByVal iRelationshipId )
	Dim sSql, oRs 

	sSql = "SELECT relationshipid, relationship FROM egov_familymember_relationships WHERE orgid = " & iorgid & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""egov_users_relationshipid"" id=""relationship"">"	
		
	Do While NOT oRs.EOF 
		response.write vbcrlf & "<option value=""" &  oRs("relationshipid") & """"
		If CLng(iRelationshipId) = CLng(oRs("relationshipid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("relationship") & "</option>"
		oRs.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' string GetRelationShip( iRelationshipId )
'--------------------------------------------------------------------------------------------------
Function GetRelationShip( ByVal iRelationshipId )
	Dim sSql, oRs

	sSql = "SELECT relationship FROM egov_familymember_relationships WHERE relationshipid = " & iRelationshipId
	session("GetRelationShipSql") = sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	session("GetRelationShipSql") = ""

	If Not oRs.EOF Then
		GetRelationShip = oRs("relationship") 
	Else
		GetRelationShip = "" 
	End if
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void DisplayLargeAddressListOld sResidenttype, sStreetNumber, sStreetName, bFound 
'--------------------------------------------------------------------------------------------------
Sub DisplayLargeAddressListOld( ByVal sResidenttype, ByVal sStreetNumber, ByVal sStreetName, ByRef bFound )
	Dim sSql, oRs

	If Not IsValidAddress( sStreetNumber, sStreetName ) Then   ' In common.asp
		sStreetNumber = ""
		sStreetName = ""
		bFound = False 
	End If 

	sSql = "SELECT distinct sortstreetname, residentstreetprefix, residentstreetname "
	sSql = sSql & "FROM egov_residentaddresses WHERE orgid = " & iOrgid & " AND residenttype = '" & sResidenttype & "' "
	sSql = sSql & "AND residentstreetname IS NOT NULL ORDER BY sortstreetname, residentstreetprefix, residentstreetname"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<input type=""text"" id=""residentstreetnumber"" name=""residentstreetnumber"" value=""" & sStreetNumber & """ size=""8"" maxlength=""10"" /> &nbsp; "
		response.write vbcrlf & "<select id=""skip_address"" name=""skip_address"">"
		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown</option>"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<option value="""  & oRs("residentstreetname") & """"
			If sStreetName = oRs("residentstreetname") Then
				response.write " selected=""selected"" "
				bFound = True 
			End If 
			response.write " >"
			response.write oRs("residentstreetname") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void DisplayLargeAddressList sResidenttype, sStreetNumber, sStreetName, bFound 
'--------------------------------------------------------------------------------------------------
Sub DisplayLargeAddressList( ByVal sResidenttype, ByVal sStreetNumber, ByVal sStreetName, ByRef bFound )
	Dim sSql, oRs, sCompareName

	If Not IsValidAddress( sStreetNumber, sStreetName ) Then   ' In common.asp
		sStreetNumber = ""
		sStreetName = ""
		bFound = False 
	End If 

	sSql = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSql = sSql & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSql = sSql & " FROM egov_residentaddresses WHERE orgid = " & iOrgid & " AND residenttype = '" & sResidenttype & "' "
	sSql = sSql & " AND residentstreetname IS NOT NULL ORDER BY sortstreetname"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<input type=""text"" id=""residentstreetnumber"" name=""residentstreetnumber"" value=""" & sStreetNumber & """ size=""8"" maxlength=""10"" /> &nbsp; "
		response.write vbcrlf & "<select id=""skip_address"" name=""skip_address"">"
		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown</option>"
		Do While NOT oRs.EOF 
			sCompareName = ""
			If oRs("residentstreetprefix") <> "" Then
				sCompareName = oRs("residentstreetprefix") & " " 
			End If 
			sCompareName = sCompareName & oRs("residentstreetname")
			If oRs("streetsuffix") <> "" Then
				sCompareName = sCompareName & " "  & oRs("streetsuffix")
			End If
			If oRs("streetdirection") <> "" Then
				sCompareName = sCompareName & " "  & oRs("streetdirection")
			End If

			response.write vbcrlf & "<option value=""" & sCompareName & """"
			If sStreetName = sCompareName Then
				response.write " selected=""selected"" "
				bFound = True 
			End If 
			response.write " >"
			response.write sCompareName & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 




%>
