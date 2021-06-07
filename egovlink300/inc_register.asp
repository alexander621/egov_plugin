<%
 dim sError, iUserId, oFamily, sUserDefaultCity, sUserDefaultState, sUserDefaultZip, sDefaultAreaCode
 dim bHasResidentStreets, bFound, sResidenttype, sBusinessAddress, bHasBusinessStreets, bUserFound
 dim bAddressRequired, bShowGenderPicks, bGenderIsRequired

 If request.servervariables("request_method") <> "POST" Then 
	   sUserDefaultCity  = ""
   	sUserDefaultState = ""
   	sUserDefaultZip   = ""

  	'Get the default city, state and zip for the org
   	GetOrgDefaultLocation iorgid, sUserDefaultCity, sUserDefaultState, sUserDefaultZip, sDefaultAreaCode
end if
 Set oFamily = Nothing  'destroy the class instance

'Check for Org Features
 lcl_orghasfeature_payments             = orghasfeature(iOrgID, "payments")
 lcl_orghasfeature_action_line          = orghasfeature(iOrgID, "action line")
 lcl_orghasfeature_large_address_list   = orghasfeature(iOrgID, "large address list")
 lcl_orghasfeature_issue_location     = orghasfeature(iOrgID, "issue location")
 lcl_orghasfeature_no_emergency_contact = orghasfeature(iOrgID, "no emergency contact")
 lcl_orghasfeature_subscriptions        = orghasfeature(iOrgID, "subscriptions")
 lcl_orghasfeature_job_postings         = orghasfeature(iOrgID, "job_postings")
 lcl_orghasfeature_bid_postings         = orghasfeature(iOrgID, "bid_postings")
 lcl_orghasfeature_donotknock           = orghasfeature(iOrgID, "donotknock")
 lcl_orghasfeature_bidpostings_viewplanholders_requirefields = orghasfeature(iOrgID, "bidpostings_viewplanholders_requirefields")
 lcl_orghasfeature_subscriptions_distributionlist_showdesc   = orghasfeature(iOrgID, "subscriptions_distributionlist_showdesc")
 lcl_orghasfeature_citizenregistration_novalidate_address    = orghasfeature(iOrgID, "citizenregistration_novalidate_address")
 bShowGenderPicks = orgHasFeature( iOrgId, "display gender pick" )
 bGenderIsRequired = orgHasFeature( iOrgId, "gender required" )

'Determine if the address is required or not - This was for Bullhead City
 bAddressRequired = orghasfeature( iorgid, "registration req address" )

'Check for org "edit displays"
 lcl_orghasdisplay_citizen_register_maint_addressinfo = orghasdisplay(iorgid,"citizen_register_maint_addressinfo")
 lcl_orghasdisplay_donotknock_list_description        = orghasdisplay(iorgid,"donotknock_list_description")

'Determine if user is coming from a job/bid postings page.
 isBusinessName_Required  = False
 lcl_label_firstlast_name = ""

 If request("fromPostings") = "Y" AND lcl_orghasfeature_bidpostings_viewplanholders_requirefields Then 
	   isBusinessName_Required  = True
   	lcl_label_firstlast_name = "Contact "
 End If 
 %>
	<script type="text/javascript" src="prototype/prototype-1.6.0.2.js"></script>
  	<script type="text/javascript" src="scripts/jquery-1.9.1.min.js"></script>

	<script type="text/javascript" src="scripts/ajaxLib.js"></script>
	<script type="text/javascript" src="scripts/modules.js"></script>
	<script type="text/javascript" src="scripts/easyform.js"></script>
	<script type="text/javascript" src="scripts/removespaces.js"></script>
	<script type="text/javascript" src="scripts/setfocus.js"></script>
	<script type="text/javascript" src="scripts/formvalidation_msgdisplay.js"></script>

	<script type="text/javascript">
	<!--

  jQuery.noConflict();  //Allows us to use jQuery

		var winHandle;
		var w = (screen.width - 640)/2;
		var h = (screen.height - 450)/2;

		function doCheck() {
		 	// If they are using the large address feature
		 	var exists = eval(document.register["residentstreetnumber"]);
		 	if(exists) {
       <%
        'If this feature is ENABLED then it DISABLES the large address validation and simply does the form validation.
         if not lcl_orghasfeature_citizenregistration_novalidate_address then
            response.write "// If a street number was entered" & vbcrlf
            response.write "if (document.register.residentstreetnumber.value != '') {" & vbcrlf
            response.write "				checkAddress( 'FinalCheckOLD', 'yes' );" & vbcrlf
            response.write "} else {" & vbcrlf
            response.write "				//checkDuplicateCitizens( 'FinalUserCheckFailed' );" & vbcrlf
            response.write "				validate();" & vbcrlf
            response.write "}" & vbcrlf
         else
            response.write "validate();" & vbcrlf
         end if
       %>
 			} else {
   				//checkDuplicateCitizens( 'FinalUserCheckFailed' );
			   	validate();
 			}
		}

		function openWin2(url, name) 
		{
		  popupWin = window.open(url, name,"resizable,width=500,height=450");
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
				alert("Please select a street name from the list first.  If you're trying to enter an address as a non-resident, enter both the street number and name in the 'Other' box.");
				setfocus(document.register.skip_address);
				return false;
			}

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
        response.write "     inlineMsg(document.getElementById(""stnumber"").id,'<strong>Required Field Missing: </strong> Please select a valid address first.',10,'stnumber');" & vbcrlf
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
        'response.write "  jQuery('#skip_address').attr('selectedIndex',0);" & vbcrlf
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
			validate();
		}

		function OkToValidate( sReturn )
		{
			//finish the validation routine 
			validate();
		}

		function validate() 
		{
			var msg              = "";
			var lcl_return_false = "N";

			/////////////////////////////////////////////////////////////////////////////
			// Validate Fields
			// NOTE: The fields are validated in reverse order so that error messages are 
			// stacked on top of each other and the error messages for the fields that 
			// are at the top of the screen are not buried under the others.
			/////////////////////////////////////////////////////////////////////////////
			if (document.register.egov_users_emergencyphone)
			{
				//Set the Emergency Phone
				if(document.register.skip_emergencyphone_areacode.value != "" || document.register.skip_emergencyphone_exchange.value != "" || document.register.skip_emergencyphone_line.value != "" ) 
				{
					var sPhone = document.register.skip_emergencyphone_areacode.value + document.register.skip_emergencyphone_exchange.value + document.register.skip_emergencyphone_line.value;
					if(sPhone.length < 10) 
					{
						//msg += "The Emergency Phone must be a valid phone number, including area code, or blank\n";
						setfocus(document.register.skip_emergencyphone_areacode);
						inlineMsg(document.getElementById("skip_emergencyphone_line").id,'<strong>Invalid Value: </strong>The Emergency Phone number must be a valid phone number, including the area code, or blank.',10,'skip_emergencyphone_line');
						lcl_return_false = "Y";
					}	
					else	
					{
						document.register.egov_users_emergencyphone.value = document.register.skip_emergencyphone_areacode.value + document.register.skip_emergencyphone_exchange.value + document.register.skip_emergencyphone_line.value;
						var rege = /^\d+$/;
						var Ok = rege.exec(document.register.egov_users_emergencyphone.value);
						if(! Ok) 
						{
							//msg += "The Emergency Phone must be a valid phone number, including area code, or blank\n";
							setfocus(document.register.skip_emergencyphone_areacode);
							inlineMsg(document.getElementById("skip_emergencyphone_line").id,'<strong>Invalid Value: </strong>The Emergency Phone must be a valid phone number, including the area code, or blank.',10,'skip_emergencyphone_line');
							lcl_return_false = "Y";
						}
					}
				}
			}

			//Set the Work Phone
			if(document.register.skip_work_areacode.value != "" || document.register.skip_work_exchange.value != "" || document.register.skip_work_line.value != "" || document.register.skip_work_ext.value != "") 
			{
				var sPhone = document.register.skip_work_areacode.value + document.register.skip_work_exchange.value + document.register.skip_work_line.value;
				if(sPhone.length < 10) 
				{
					//msg += "The Work Phone Number must be a valid phone number, including area code, or blank\n";
					setfocus(document.register.skip_work_areacode);
					inlineMsg(document.getElementById("skip_work_line").id,'<strong>Invalid Value: </strong>The Work Phone must be a valid phone number, including the area code, or blank.',10,'skip_work_line');
					lcl_return_false = "Y";
				} 
				else 
				{
					document.register.egov_users_userworkphone.value = document.register.skip_work_areacode.value + document.register.skip_work_exchange.value + document.register.skip_work_line.value + document.register.skip_work_ext.value;
					var rege = /^\d+$/;
					var Ok = rege.exec(document.register.egov_users_userworkphone.value);
					if(! Ok) 
					{
						//msg += "The Work Phone Number must be a valid phone number, including area code, or blank\n";
						setfocus(document.register.skip_work_areacode);
						inlineMsg(document.getElementById("skip_work_line").id,'<strong>Invalid Value: </strong>The Work Phone number must be a valid phone number, including the area code, or blank.',10,'skip_work_line');
						lcl_return_false = "Y";
					}
				}
			}

			//Set the Business Name, if coming from Postings page
	<%		if isBusinessName_Required then %>
			if(document.getElementById("egov_users_userbusinessname").value == '') 
			{
				setfocus(document.register.egov_users_userbusinessname);
				inlineMsg(document.getElementById("egov_users_userbusinessname").id,'<strong>Required Field Missing: </strong>Business Name',10,'egov_users_userbusinessname');
				lcl_return_false = "Y";
			}
	  <%	end if %>

			//Set the Fax
			if(document.register.skip_fax_areacode.value != "" || document.register.skip_fax_exchange.value != "" || document.register.skip_fax_line.value != "" ) 
			{
				var sPhone = document.register.skip_fax_areacode.value + document.register.skip_fax_exchange.value + document.register.skip_fax_line.value;
				if(sPhone.length < 10) 
				{
					//msg += "The Fax must be a valid phone number, including area code, or blank\n";
					setfocus(document.register.skip_fax_areacode);
					inlineMsg(document.getElementById("skip_fax_line").id,'<strong>Invalid Value: </strong>The Fax must be a valid phone number, including the area code, or blank.',10,'skip_fax_line');
					lcl_return_false = "Y";
				} 
				else	
				{
					document.register.egov_users_userfax.value = document.register.skip_fax_areacode.value + document.register.skip_fax_exchange.value + document.register.skip_fax_line.value;
					var rege = /^\d+$/;
					var Ok = rege.exec(document.register.egov_users_userfax.value);
					if(! Ok) 
					{
						//msg += "The Fax must be a valid phone number, including area code, or blank\n";
						setfocus(document.register.skip_fax_areacode);
						inlineMsg(document.getElementById("skip_fax_line").id,'<strong>Invalid Value: </strong>The Fax must be a valid phone number, including the area code, or blank.',10,'skip_fax_line');
						lcl_return_false = "Y";
					}
				}
			}

			//Set the Cell Phone
			if(document.register.skip_cell_areacode.value != "" || document.register.skip_cell_exchange.value != "" || document.register.skip_cell_line.value != "" )	
			{
				var cPhone = document.register.skip_cell_areacode.value + document.register.skip_cell_exchange.value + document.register.skip_cell_line.value;
				if(cPhone.length < 10) 
				{
					//msg += "The cell phone number must be a valid phone number, including area code, or blank\n";
					setfocus(document.register.skip_cell_areacode);
					inlineMsg(document.getElementById("skip_cell_line").id,'<strong>Invalid Value: </strong>The Cell Phone must be a valid phone number, including the area code, or blank.',10,'skip_cell_line');
					lcl_return_false = "Y";
				} 
				else	
				{
					document.register.egov_users_usercell.value = document.register.skip_cell_areacode.value + document.register.skip_cell_exchange.value + document.register.skip_cell_line.value;
					var crege = /^\d+$/;
					var cOk = crege.exec(document.register.egov_users_usercell.value);
					if(! cOk) 
					{
						//msg += "The cell phone number must be a valid phone number, including area code, or blank\n";
						setfocus(document.register.skip_cell_areacode);
						inlineMsg(document.getElementById("skip_cell_line").id,'<strong>Invalid Value: </strong>The Cell Phone must be a valid phone number, including the area code, or blank.',10,'skip_cell_line');
						lcl_return_false = "Y";
					}
				}
			}

			//Set the Phone Number
			document.register.egov_users_userhomephone.value = document.register.skip_user_areacode.value + document.register.skip_user_exchange.value + document.register.skip_user_line.value;
			if(document.register.egov_users_userhomephone.value != "" )	
			{
				var hPhone = document.register.egov_users_userhomephone.value;
				if(hPhone.length < 10) 
				{
					//msg += "The phone number must be a valid phone number, including area code.\n";
					setfocus(document.register.skip_user_areacode);
					inlineMsg(document.getElementById("skip_user_line").id,'<strong>Invalid Value: </strong>The Home Phone must be a valid phone number, including the area code, or blank.',10,'skip_user_line');
					lcl_return_false = "Y";
				} 
				else	
				{
					var rege = /^\d+$/;
					var Ok = rege.exec(document.register.egov_users_userhomephone.value);
					if(! Ok) 
					{
						//msg += "The phone number must be a valid phone number, including area code.\n";
						setfocus(document.register.skip_user_areacode);
						inlineMsg(document.getElementById("skip_user_line").id,'<strong>Invalid Value: </strong>The Home Phone must be a valid phone number, including the area code, or blank.',10,'skip_user_line');
						lcl_return_false = "Y";
					}
				}
			} 
			else	
			{
				//msg+="The phone number cannot be blank.\n";
				setfocus(document.register.skip_user_areacode);
				inlineMsg(document.getElementById("skip_user_line").id,'<strong>Required Field Missing: </strong>Home Phone',10,'skip_user_line');
				lcl_return_false = "Y";
			}

			// Gender validation if present and is required
<%			If bShowGenderPicks And bGenderIsRequired Then	%>
				if ($("egov_users_gender").value == 'N') 
				{
					setfocus(document.register.egov_users_gender);
					inlineMsg(document.getElementById("egov_users_gender").id,'<strong>Required Field Missing: </strong>Gender',10,'egov_users_gender');
					lcl_return_false = "Y";
				}
<%			End If		%>

			//Last Name
			if ($("egov_users_userlname").value == '') 
			{
				//msg+="Last Name is required.\n"
				setfocus(document.register.egov_users_userlname);
				inlineMsg(document.getElementById("egov_users_userlname").id,'<strong>Required Field Missing: </strong>Last Name',10,'egov_users_userlname');
				lcl_return_false = "Y";
			}

			//First Name
			if ($("egov_users_userfname").value == '') 
			{
				//msg+="First Name is required.\n"
				setfocus(document.register.egov_users_userfname);
				inlineMsg(document.getElementById("egov_users_userfname").id,'<strong>Required Field Missing: </strong>First Name',10,'egov_users_userfname');
				lcl_return_false = "Y";
			}

			//Password
			var sPasswrd  = document.register.egov_users_userpassword.value;
			var sPasswrd2 = document.register.skip_userpassword2.value;

			if(sPasswrd == "") 
			{
				setfocus(document.register.egov_users_userpassword);
				inlineMsg(document.getElementById("egov_users_userpassword").id,'<strong>Required Field Missing: </strong>Password',10,'egov_users_userpassword');
				lcl_return_false = "Y";
			} 
			else 
			{
				if(sPasswrd2 == "") 
				{
					setfocus(document.register.skip_userpassword2);
					inlineMsg(document.getElementById("skip_userpassword2").id,'<strong>Required Field Missing: </strong>Verify Password',10,'skip_userpassword2');
					lcl_return_false = "Y";
				} 
				else 
				{
					if(sPasswrd.length > 50)	
					{
						//msg+="The password you have entered is too long.\n";
						setfocus(document.register.egov_users_userpassword);
						inlineMsg(document.getElementById("egov_users_userpassword").id,'<strong>Invalid Value: </strong>The Password you have entered is too long',10,'egov_users_userpassword');
						lcl_return_false = "Y";
					} 
					else 
					{
						if(document.register.egov_users_userpassword.value != document.register.skip_userpassword2.value) 
						{
							//msg+="The passwords you have entered do not match.\n";
							setfocus(document.register.egov_users_userpassword);
							inlineMsg(document.getElementById("skip_userpassword2").id,'<strong>Invalid Value: </strong>The Passwords you have entered do not match',10,'skip_userpassword2');
							lcl_return_false = "Y";
						}
					}
				}
			}

			//Email
			//var rege = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*\.(\w{2}|(com|net|org|edu|mil|gov|biz|us))$/;
			var rege = /.+@.+\..+/i;

			//var Ok = rege.test(document.register.egov_users_useremail.value.trim());
			//var Ok = rege.test($("egov_users_useremail").value.trim());
   var sEGovUsersEmail = document.getElementById("egov_users_useremail").value;
       sEGovUsersEmail = jQuery.trim(sEGovUsersEmail);

			var Ok = rege.test(sEGovUsersEmail);

			//if($("egov_users_useremail").value.trim() == '') 
   if(sEGovUsersEmail == '') 
			{
				setfocus(document.register.egov_users_useremail);
				inlineMsg(document.getElementById("egov_users_useremail").id,'<strong>Required Field Missing: </strong>Email',10,'egov_users_useremail');
				lcl_return_false = "Y";
			} 
			else 
			{
				if(! Ok) 
				{
					//msg+="The email must be in a valid format.\n";
					setfocus(document.register.egov_users_useremail);
					inlineMsg(document.getElementById("egov_users_useremail").id,'<strong>Invalid Value: </strong>The email must be in proper format.',10,'egov_users_useremail');
					lcl_return_false = "Y";
				}
			}

			/////////////////////////////////////////////////////////////////////////////
			// Process Records
			/////////////////////////////////////////////////////////////////////////////
			//Process the Neighborhood
			var neighborexists = eval(document.register["skip_neighborhoodid"]);
			if(neighborexists) 
			{
				//See if they picked from the neighborhood dropdown and put that in the neighborhood field 
				if(document.register.skip_neighborhoodid.selectedIndex > -1) 
				{	
					var nelement = document.register.skip_neighborhoodid;
					var nselectedvalue = nelement.options[nelement.selectedIndex].value;

					//0 is the first pick that we do not want
					if(nselectedvalue != "0") 
					{
						document.register.egov_users_neighborhoodid.value = nselectedvalue;
					}
				}
			}

			//Process the Business Address if one was chosen
			var bexists = eval(document.register["skip_Baddress"]);
			if(bexists)	
			{
				//See if they picked from the business dropdown and put that in the address field 
				if(document.register.skip_Baddress.selectedIndex > -1) 
				{
					var belement = document.register.skip_Baddress;
					var bselectedvalue = belement.options[belement.selectedIndex].value;

					//alert( bselectedvalue );
					//0000 is the first pick that we do not want
					if(bselectedvalue != "0000")	
					{
						document.register.egov_users_userbusinessaddress.value = bselectedvalue;
						document.register.egov_users_residenttype.value = "B";
					}
				}
			}

<%			If bAddressRequired Then	%>

			
<%			If lcl_orghasfeature_large_address_list Then	%>
				if( ($F('residentstreetnumber') == '' ||  $('skip_address').getValue() == '0000') && $F("egov_users_useraddress") == "" ) 
				{
					setfocus(document.register.residentstreetnumber);
					inlineMsg($("residentstreetnumber").id,'<strong>Required Field Missing: </strong>Address',10,'residentstreetnumber');
<%			Else	%>
				if ($F("egov_users_useraddress") == "")
				{
					setfocus(document.register.egov_users_useraddress);
					inlineMsg($("egov_users_useraddress").id,'<strong>Required Field Missing: </strong>Address',10,'egov_users_useraddress');
<%			End If											%>
					lcl_return_false = "Y";
				}

			if ($("egov_users_usercity").value == "")
			{
				setfocus(document.register.egov_users_usercity);
				inlineMsg(document.getElementById("egov_users_usercity").id,'<strong>Required Field Missing: </strong>City',10,'egov_users_usercity');
				lcl_return_false = "Y";
			}

			if ($("egov_users_userstate").value == "")
			{
				setfocus(document.register.egov_users_userstate);
				inlineMsg(document.getElementById("egov_users_userstate").id,'<strong>Required Field Missing: </strong>State',10,'egov_users_userstate');
				lcl_return_false = "Y";
			}

			if ($("egov_users_userzip").value == "")
			{
				setfocus(document.register.egov_users_userzip);
				inlineMsg(document.getElementById("egov_users_userzip").id,'<strong>Required Field Missing: </strong>Zip',10,'egov_users_userzip');
				lcl_return_false = "Y";
			}

<%			End If						%>

			if ($F('egov_users_userzip') == '123456')
			{
				setfocus(document.register.egov_users_userzip);
				inlineMsg(document.getElementById("egov_users_userzip").id,'<strong>Invalid Value: </strong>Zip',10,'egov_users_userzip');
				lcl_return_false = "Y";
			}

			//Process the Resident Address if one was chosen - this is second to set the local resident type
			var exists = eval(document.register["skip_Raddress"]);
			if(exists) 
			{
				//See if they picked from the resident dropdown and put that in the address field 
				if(document.register.skip_Raddress.selectedIndex > -1) 
				{
					var element = document.register.skip_Raddress;
					var selectedvalue = element.options[element.selectedIndex].value;

					//alert( selectedvalue );
					//  0000 is the first pick that we do not want
					if(selectedvalue != "0000") 
					{
						document.register.egov_users_useraddress.value = selectedvalue;
						document.register.egov_users_residenttype.value = "R";
					}
				}
			}

			//handle the large quantity Street Addresses
			exists = eval(document.register["residentstreetnumber"]);
			if(exists) 
			{
				if(document.register.residentstreetnumber.value != '') 
				{
						//See if they picked from the resident dropdown and put that in the address field 
						if(document.register.skip_address.selectedIndex > -1) 
						{
							var element = document.register.skip_address;
							var selectedvalue = element.options[element.selectedIndex].value;

							//alert( selectedvalue );
							//0000 is the first pick that we do not want
							if(selectedvalue != "0000") 
							{
									document.register.egov_users_useraddress.value = document.register.residentstreetnumber.value + ' ' + selectedvalue;
									document.register.egov_users_residenttype.value = "R";
							}
					 }
				}
			}

			//if(msg != "") {
				//msg="Your form could not be submitted for the following reasons.\n\n" + msg;
				//alert(msg);
			if(lcl_return_false=="Y") 
			{
<%				If lcl_orghasfeature_large_address_list Then	%>
					if( $F('residentstreetnumber') != '' &&  $('skip_address').getValue() != '0000' ) 
					{
						$('egov_users_useraddress').value = '';
					}
<%				End If											%>
				return false;
			}	
			else 	
			{	
				if (validateForm('register')) 
				{ 
					document.register.submit(); 
				}
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
<%


response.write "<tr>" & vbcrlf
response.write "<td valign=""top"">" & vbcrlf
response.write "<div id=""content"">" & vbcrlf
response.write "<div id=""centercontent"">" & vbcrlf
assistant = "alexa"
assistantname = "Alexa"
if request.querystring("googleauth") = "true" then 
	assistant = "googleauth"
	assistantname = "Google Home"
end if

if iorgid = "26" then
response.write "<p><font class=""pagetitle"">Welcome to the " & sOrgName & " " & assistantname & " Registration</font><br /></p>" & vbcrlf
response.write "<p>Registering is FREE, quick , and easy!</p>" & vbcrlf
else
response.write "<p><font class=""pagetitle"">Welcome to the " & sOrgName & " Registration</font><br /></p>" & vbcrlf
response.write "<p>Registering to use " & sOrgName & " E-Gov Services is FREE, quick and easy to establish!</p>" & vbcrlf
response.write "<p>You can access your transaction history.</p>" & vbcrlf

if lcl_orghasfeature_payments OR lcl_orghasfeature_action_line then
	response.write "<p>For example,"
end if

if lcl_orghasfeature_payments then
	  response.write " history of online payments using " & sOrgName & " E-Gov Services"
end if

if lcl_orghasfeature_payments AND lcl_orghasfeature_action_line then
	response.write " or "
end if

if lcl_orghasfeature_action_line then
	Dim oActionOrg
	Set oActionOrg = New classOrganization
	response.write " requests submitted via the " '& sOrgActionName 
	response.write oActionOrg.GetOrgFeatureName("action line")
	Set oActionOrg = Nothing 
end if

if lcl_orghasfeature_payments OR lcl_orghasfeature_action_line then
	response.write ". </p>"
end if


response.write "<p>You can choose to have contact information (such as an address & telephone number) saved with your " & vbcrlf
response.write "membership thereby eliminating the requirement to ""re-type"" this information into online forms.</p>" & vbcrlf
response.write "<div class=""box_header4"">" & sOrgName & " Registration: </div>" & vbcrlf
end if
response.write "  <div id=""registrationinformation"" class=""groupSmall4"">" & vbcrlf
response.write "<form name=""register"" action=""register.asp"" method=""post"" autocomplete=""off"">" & vbcrlf
response.write "	<input type=""hidden"" name=""columnnameid"" value=""userid"" />" & vbcrlf
response.write "	<input type=""hidden"" name=""token"" value=""" & request("token") & """ />" & vbcrlf
response.write "	<input type=""hidden"" name=""" & assistant & """ value=""" & request(assistant) & """ />" & vbcrlf
	response.write "<input type=""hidden"" name=""state"" value=""" & request("state") & """ />"
	response.write "<input type=""hidden"" name=""redirect_uri"" value=""" & request("redirect_uri") & """ />"
response.write "	<input type=""hidden"" name=""egov_users_userregistered"" value=""1"" />" & vbcrlf
response.write "  	<input type=""hidden"" name=""egov_users_orgid"" value=""" & iorgid & """ />" & vbcrlf
response.write "  	<input type=""hidden"" name=""egov_users_relationshipid"" value=""" & GetDefaultRelationShipId( iOrgid ) & """ />" & vbcrlf
response.write "  	<input type=""hidden"" name=""ef:egov_users_useremail-text/req"" value=""Email Address"" />" & vbcrlf
response.write "  	<input type=""hidden"" name=""ef:egov_users_userpassword-text/req"" value=""Password 1"" />" & vbcrlf
response.write "  	<input type=""hidden"" name=""ef:skip_userpassword2-text/req"" value=""Password 2"" />" & vbcrlf
response.write "  	<input type=""hidden"" name=""ef:egov_users_userhomephone-text/req/phone"" value=""Phone Number"" />" & vbcrlf
response.write "  	<input type=""hidden"" name=""ef:egov_users_userfname-text/req"" value=""First name"" />" & vbcrlf
response.write "  	<input type=""hidden"" name=""ef:egov_users_userlname-text/req"" value=""Last name"" />" & vbcrlf
response.write "  	<input type=""hidden"" name=""egov_users_residenttype"" value=""N"" />" & vbcrlf
response.write "  	<input type=""hidden"" name=""egov_users_neighborhoodid"" value=""0"" />" & vbcrlf
response.write "  	<input type=""hidden"" name=""egov_users_headofhousehold"" value=""1"" />" & vbcrlf

If Not bShowGenderPicks Then 
	response.write "  	<input type=""hidden"" id=""egov_users_gender"" name=""egov_users_gender"" value=""N"" />" & vbcrlf
End If 

if Not lcl_orghasfeature_donotknock then
	response.write "    <input type=""hidden"" name=""isOnDoNotKnockList_peddlers"" id=""isOnDoNotKnockList_peddlers"" value="""" />" & vbcrlf
	response.write "    <input type=""hidden"" name=""isOnDoNotKnockList_solicitors"" id=""isOnDoNotKnockList_solicitors"" value="""" />" & vbcrlf
end if

response.write "<table border=""0"" cellpadding=""2"" cellspacing=""0"" id=""registrationdisplay"">" & vbcrlf

if errormsg <> "" then
	response.write "  <tr><td colspan=""2"" align=""center"">" & errormsg & "</td></tr>" & vbcrlf
end if

response.write "  <tr>" & vbcrlf
response.write "      <td colspan=""2"" align=""center"">" & vbcrlf
response.write "          <font color=""#ff0000"">*</font>" & vbcrlf
response.write "          Indicates a required field that must be filled in order to complete your registration." & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write "          <span class=""cot-text-emphasized"" title=""This field is required""><span class=""cot-text-emphasized""><font color=""#ff0000"">*</font></span>" & vbcrlf
response.write "      			 Email:" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""text"" value=""" & request("egov_users_useremail") & """ name=""egov_users_useremail"" id=""egov_users_useremail"" class=""threehundredwide"" maxlength=""100"" onchange=""clearMsg('egov_users_useremail')"" autocomplete=""off"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

response.write " 	<tr>" & vbcrlf
response.write "      <td class=""label"" align=""right""><nobr>" & vbcrlf
response.write "			       <span class=""cot-text-emphasized"" title=""This field is required""><span class=""cot-text-emphasized""><font color=""#ff0000"">*</font></span>" & vbcrlf
response.write "        		Password:" & vbcrlf
response.write "      </nobr></td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""password"" value=""" & request("egov_users_userpassword") & """ name=""egov_users_userpassword"" id=""egov_users_userpassword"" class=""twozeroeightwide"" maxlength=""50"" onchange=""clearMsg('egov_users_userpassword')"" autocomplete=""off"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write "          <span class=""cot-text-emphasized"" title=""This field is required""><span class=""cot-text-emphasized""><font color=""#ff0000"">*</font></span>" & vbcrlf
response.write "          Verify Password:" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""password"" value=""" & request("skip_userpassword2") & """ name=""skip_userpassword2"" id=""skip_userpassword2"" class=""twozeroeightwide"" maxlength=""50"" onchange=""clearMsg('skip_userpassword2')"" autocomplete=""off"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write "          <span class=""cot-text-emphasized"" title=""This field is required""><span class=""cot-text-emphasized""><font color=""#ff0000"">*</font></span>" & vbcrlf
response.write            lcl_label_firstlast_name & "First Name:" & vbcrlf
response.write "          </span>" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <span class=""cot-text-emphasized"" title=""This field is required""> " & vbcrlf
response.write "       			<input type=""text"" value=""" & request("egov_users_userfname") & """ id=""egov_users_userfname"" name=""egov_users_userfname"" class=""threehundredwide"" maxlength=""100"" onchange=""clearMsg('egov_users_userfname')"" autocomplete=""off"" />" & vbcrlf
response.write "        		</span>" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write "          <span class=""cot-text-emphasized"" title=""This field is required""><span class=""cot-text-emphasized""><font color=""#ff0000"">*</font></span>" & vbcrlf
response.write            lcl_label_firstlast_name & "Last Name:" & vbcrlf
response.write "         	</span>" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <span class=""cot-text-emphasized"" title=""This field is required"">" & vbcrlf
response.write "       			<input type=""text"" value=""" & request("egov_users_userlname") & """ id=""egov_users_userlname"" name=""egov_users_userlname"" class=""threehundredwide"" maxlength=""100"" onchange=""clearMsg('egov_users_userlname')"" autocomplete=""off"" />" & vbcrlf
response.write "       			</span>" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

If bShowGenderPicks Then 
	response.write "<tr>" & vbcrlf
	response.write "<td class=""label"" align=""right"">" & vbcrlf

	If bGenderIsRequired Then 
		response.write "<span class=""cot-text-emphasized"" title=""This field is required""><span class=""cot-text-emphasized""><font color=""#ff0000"">*</font></span>" & vbcrlf
	End If 

	response.write "Gender:"

	If bGenderIsRequired Then 
		response.write "</span>" & vbcrlf
	End If 

	response.write "</td>" & vbcrlf
	response.write "<td>" & vbcrlf
	DisplayGenderPicks "egov_users_gender", "N"		' in common.asp
	response.write "</td>" & vbcrlf
	response.write "</tr>" & vbcrlf
End If 

response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write "          <font color=""#ff0000"">*</font> Phone Number:" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""hidden"" value="""" name=""egov_users_userhomephone"" id=""egov_users_userhomephone"" />" & vbcrlf
response.write "      			(<input type=""text"" value="""" name=""skip_user_areacode"" id=""skip_user_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_user_line')"" />)&nbsp;" & vbcrlf
response.write "       			<input type=""text"" value="""" name=""skip_user_exchange"" id=""skip_user_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_user_line')"" />&ndash;" & vbcrlf
response.write "       			<input type=""text"" value="""" name=""skip_user_line"" id=""skip_user_line"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" onchange=""clearMsg('skip_user_line')"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write "          Cell Phone:" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""hidden"" value="""" name=""egov_users_usercell"" id=""egov_users_usercell"" />" & vbcrlf
response.write "      			(<input type=""text"" value="""" name=""skip_cell_areacode"" id=""skip_cell_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_cell_line')"" />)&nbsp;" & vbcrlf
response.write "       			<input type=""text"" value="""" name=""skip_cell_exchange"" id=""skip_cell_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_cell_line')"" />&ndash;" & vbcrlf
response.write "       			<input type=""text"" value="""" name=""skip_cell_line"" id=""skip_cell_line"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" onchange=""clearMsg('skip_cell_line')"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write "          Fax:" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""hidden"" value="""" name=""egov_users_userfax"" />" & vbcrlf
response.write "       		(<input type=""text"" value="""" name=""skip_fax_areacode"" id=""skip_fax_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_fax_line')"" />)&nbsp;" & vbcrlf
response.write "       			<input type=""text"" value="""" name=""skip_fax_exchange"" id=""skip_fax_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_fax_line')"" />&ndash;" & vbcrlf
response.write "       			<input type=""text"" value="""" name=""skip_fax_line"" id=""skip_fax_line"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" onchange=""clearMsg('skip_fax_line')"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

'Show additional address info if org has "edit display"
lcl_address_info_displayid = getDisplayID("citizen_register_maint_addressinfo")
lcl_address_info           = getOrgDisplayWithID(iorgid, lcl_address_info_displayid, False)

if lcl_address_info <> "" then
	response.write "  <tr valign=""bottom"">" & vbcrlf
	response.write "      <td>&nbsp;</td>" & vbcrlf
	response.write "      <td style=""padding-top:10pt;"">" & lcl_address_info & "</td>" & vbcrlf
	response.write "  </tr>" & vbcrlf
end if

bHasResidentStreets = HasResidentTypeStreets( iOrgid, "R" )
'bHasResidentStreets = False  ' for Bullhead city testing
bFound = False 

If bHasResidentStreets Then 
	If Not lcl_orghasfeature_large_address_list Then 
		response.write "<tr>" & vbcrlf
		response.write "<td class=""label"" align=""right"">Resident Street:</td>" & vbcrlf
		response.write "<td nowrap=""nowrap"">" & vbcrlf

		DisplayAddresses iOrgid, "R"
		response.write "<input type=""hidden"" name=""egov_users_useraddress"" id=""egov_users_useraddress"" value="""" />" & vbcrlf

		response.write "</td>" & vbcrlf
		response.write "</tr>" & vbcrlf
  	Else 
    		'Show the large address list solution
    		response.write "<tr>" & vbcrlf
    		response.write "<td class=""label fullonly"" align=""right"" valign=""top"" nowrap=""nowrap"">" & vbcrlf

    		If bAddressRequired Then 
      			response.write "<font color=""#ff0000"">*</font>" & vbcrlf
    		End If

    		response.write "Address:" & vbcrlf
      	response.write "</td>" & vbcrlf
    		response.write "<td id=""addblock"">" & vbcrlf

      	response.write "    <fieldset class=""addressFieldset"">" & vbcrlf

    		DisplayLargeAddressList "R"

   		'If this feature is ENABLED then it DISABLES the large address validation and simply does the form validation.
    		If Not lcl_orghasfeature_citizenregistration_novalidate_address Then 
      			response.write "&nbsp;<input type=""button"" class=""button"" value=""Validate Address"" onclick=""checkAddress('CheckResults', 'no');"" />" & vbcrlf
    		End If

   		'Set up the Address label/value
'   		 lcl_address_label = "Address"
      	lcl_address_label = "- Or Other Not Listed <nobr>(Non-Resident)</nobr> -"
   		 lcl_address_value = ""

'   		 if bHasResidentStreets then
'   			   lcl_address_label = lcl_address_label & " (if not listed):"
'   		 else
'   		   	lcl_address_label = lcl_address_label & ":"
'   		 end if

   		 if not bfound then
   			   lcl_address_value = sAddress
   		 end if

   		 If bAddressRequired Then 
   			   response.write "<font color=""#ff0000"">*</font>" & vbcrlf
   		 End If 

      	response.write "      <br />" & lcl_address_label & "<br />" & vbcrlf
      	response.write "      <input type=""text"" name=""egov_users_useraddress"" id=""egov_users_useraddress"" value=""" & lcl_address_value & """ class=""threehundredwide"" maxlength=""100"" autocomplete=""off"" />" & vbcrlf

          'Only build the "invalid address" section if the org has the "issue location" and "large address list" features
           'if lcl_orghasfeature_issue_location AND lcl_orghasfeature_large_address_list then
           if lcl_orghasfeature_large_address_list then
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
           end if
      	response.write "    </fieldset>" & vbcrlf




    		response.write "</td>" & vbcrlf
    		response.write "</tr>" & vbcrlf
  	End If
Else
	' they do not have addresses loaded. I think this is only Pikeville'
	lcl_address_label = "Address:"
	lcl_address_value = ""
	
	response.write "<tr>" & vbcrlf
	response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf

	If bAddressRequired Then 
		response.write "<font color=""#ff0000"">*</font>" & vbcrlf
	End If 

	response.write lcl_address_label & "</td>" & vbcrlf
	response.write "<td><input type=""text"" name=""egov_users_useraddress"" id=""egov_users_useraddress"" value=""" & lcl_address_value & """ class=""threehundredwide"" maxlength=""100"" autocomplete=""off"" /></td>" & vbcrlf
	response.write "</tr>" & vbcrlf  	
End If 

'Set up the Address label/value
' lcl_address_label = "Address"
' lcl_address_value = ""

' if bHasResidentStreets then
'	   lcl_address_label = lcl_address_label & " (if not listed):"
' else
'   	lcl_address_label = lcl_address_label & ":"
' end if

' if not bfound then
'	   lcl_address_value = sAddress
' end if

' response.write "<tr>" & vbcrlf
' response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf

' If bAddressRequired Then 
'	   response.write "<font color=""#ff0000"">*</font>" & vbcrlf
' End If 

' response.write lcl_address_label & "</td>" & vbcrlf
' response.write "<td><input type=""text"" name=""egov_users_useraddress"" id=""egov_users_useraddress"" value=""" & lcl_address_value & """ class=""threehundredwide"" maxlength=""100"" autocomplete=""off"" /></td>" & vbcrlf
' response.write "  </tr>" & vbcrlf

 response.write "<tr>" & vbcrlf
 response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
 'response.write "Resident Unit:</td>" & vbcrlf
 response.write "Unit:</td>" & vbcrlf
 response.write "<td><input type=""text"" value="""" name=""egov_users_userunit"" size=""10"" maxlength=""10"" /></td>" & vbcrlf
 response.write "</tr>" & vbcrlf

 If OrgHasNeighborhoods( iorgid ) Then 
 	  response.write "<tr>" & vbcrlf
 	  response.write "<td class=""label"" align=""right"">Neighborhood:</td>" & vbcrlf
 	  response.write "<td>" & vbcrlf
 	  DisplayNeighborhoods iorgid
 	  response.write "</td>" & vbcrlf
 	  response.write "</tr>" & vbcrlf
 Else 
 	  response.write "<input type=""hidden"" name=""skip_neighborhoodid"" value=""0"" />" & vbcrlf
 End If 

 response.write "<tr>" & vbcrlf
 response.write "<td class=""label"" align=""right"">" & vbcrlf

 If bAddressRequired Then 
   	response.write "<font color=""red"">*</font>" & vbcrlf
 End If 

 response.write "City:</td>" & vbcrlf
 response.write "<td><input type=""text"" value=""" & sUserDefaultCity & """ id=""egov_users_usercity"" name=""egov_users_usercity"" class=""threehundredwide"" maxlength=""40"" autocomplete=""off"" /></td>" & vbcrlf
 response.write "</tr>" & vbcrlf

 response.write "<tr>" & vbcrlf
 response.write "<td class=""label"" align=""right"">" & vbcrlf

 If bAddressRequired Then 
	   response.write "<font color=""red"">*</font>" & vbcrlf
 End If 

 response.write "State:</td>" & vbcrlf
 response.write "<td><input type=""text"" value=""" & sUserDefaultState & """ id=""egov_users_userstate"" name=""egov_users_userstate"" size=""2"" maxlength=""2"" autocomplete=""off"" /></td>" & vbcrlf
 response.write "</tr>" & vbcrlf

 response.write "<tr>" & vbcrlf
 response.write "<td class=""label"" align=""right"">" & vbcrlf

 If bAddressRequired Then 
	   response.write "<font color=""red"">*</font>" & vbcrlf
 End If 

 response.write "ZIP:</td>" & vbcrlf
 response.write "<td><input type=""text"" value=""" & sUserDefaultZip & """ id=""egov_users_userzip"" name=""egov_users_userzip"" size=""10"" maxlength=""10"" autocomplete=""off"" /></td>" & vbcrlf
 response.write "</tr>" & vbcrlf

 response.write "<tr>" & vbcrlf
 response.write "<td class=""label"" align=""right"">" & vbcrlf

'If user is coming from a job/bid postings page to register than this field is required.
 if isBusinessName_Required then
   	response.write "<span class=""cot-text-emphasized"" title=""This field is required"">" & vbcrlf
   	response.write "<span class=""cot-text-emphasized""><font color=""#ff0000"">*</font></span>" & vbcrlf
   	response.write "  Business Name:" & vbcrlf
   	response.write "</span>" & vbcrlf

   	lcl_onchange = " onchange=""clearMsg('egov_users_userbusinessname')"""
 else 
	   response.write "Business Name:"
   	lcl_onchange = ""
 end if

 response.write "      </td>" & vbcrlf
 response.write "      <td>" & vbcrlf
 response.write "          <input type=""text"" value=""" & request("egov_users_userbusinessname") & """ name=""egov_users_userbusinessname"" id=""egov_users_userbusinessname"" class=""threehundredwide"" maxlength=""100""" & lcl_onchange & " />" & vbcrlf
 response.write "      </td>" & vbcrlf
 response.write "  </tr>" & vbcrlf

 bHasBusinessStreets = HasResidentTypeStreets( iOrgid, "B" )
 bFound = False 

 if bHasBusinessStreets  then
   	response.write "  <tr>" & vbcrlf
   	response.write "      <td class=""label"" align=""right"">" & vbcrlf
   	response.write "          Business Street:" & vbcrlf
   	response.write "      </td>" & vbcrlf
   	response.write "      <td>" & vbcrlf
   	DisplayAddresses iorgid, "B"
   	response.write "      </td>" & vbcrlf
   	response.write "  </tr>" & vbcrlf
 end if

 response.write "  <tr>" & vbcrlf
 response.write "      <td class=""label"" align=""right"">" & vbcrlf

if bHasBusinessStreets then
	response.write "Street (if not listed):" 
else
	response.write "Business Street:" 
end if

response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""text"" value=""" & request("egov_users_userbusinessaddress") & """ name=""egov_users_userbusinessaddress"" class=""threehundredwide"" maxlength=""100"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf
response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write "          Work Phone:" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""hidden"" value="""" name=""egov_users_userworkphone"" />" & vbcrlf
response.write "     	  	(<input type=""text"" value="""" name=""skip_work_areacode"" id=""skip_work_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_work_line')"" />)&nbsp;" & vbcrlf
response.write "     			  <input type=""text"" value="""" name=""skip_work_exchange"" id=""skip_work_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_work_line')"" />&ndash;" & vbcrlf
response.write "     		  	<input type=""text"" value="""" name=""skip_work_line"" id=""skip_work_line"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" onchange=""clearMsg('skip_work_line')"" />&nbsp;" & vbcrlf
response.write "        		ext. <input type=""text"" value="""" name=""skip_work_ext"" id=""skip_work_ext"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

If Not lcl_orghasfeature_no_emergency_contact Then 
	response.write "  <tr>" & vbcrlf
	response.write "      <td class=""label"" align=""right"">" & vbcrlf
	response.write "          Emergency Contact:" & vbcrlf
	response.write "      </td>" & vbcrlf
	response.write "      <td>" & vbcrlf
	response.write "          <input type=""text"" value=""" & request("egov_users_emergencycontact") & """ name=""egov_users_emergencycontact"" id=""egov_users_emergencycontact"" class=""threehundredwide"" maxlength=""100"" />" & vbcrlf
	response.write "      </td>" & vbcrlf
	response.write "  </tr>" & vbcrlf
	response.write "  <tr>" & vbcrlf
	response.write "      <td class=""label"" align=""right"" valign=""top"">" & vbcrlf
	response.write "          Emergency Phone:" & vbcrlf
	response.write "      </td>" & vbcrlf
	response.write "      <td>" & vbcrlf
	response.write "          <input type=""hidden"" value="""" name=""egov_users_emergencyphone"" />" & vbcrlf
	response.write "    		  	(<input type=""text"" value="""" name=""skip_emergencyphone_areacode"" id=""skip_emergencyphone_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_emergencyphone_line')"" />)&nbsp;" & vbcrlf
	response.write "    				  <input type=""text"" value="""" name=""skip_emergencyphone_exchange"" id=""skip_emergencyphone_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" onchange=""clearMsg('skip_emergencyphone_line')"" />&ndash;" & vbcrlf
	response.write "      				<input type=""text"" value="""" name=""skip_emergencyphone_line"" id=""skip_emergencyphone_line"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" onchange=""clearMsg('skip_emergencyphone_line')"" />" & vbcrlf
	response.write "      </td>" & vbcrlf
	response.write "  </tr>" & vbcrlf
End If 

'Do Not Knock Options
if lcl_orghasfeature_donotknock then
	if lcl_orghasdisplay_donotknock_list_description then
		lcl_dnk_description = getOrgDisplay(iorgid,"donotknock_list_description")
	else
		lcl_dnk_description = "&nbsp"
	end if

	response.write "  <tr>" & vbcrlf
	response.write "      <td colspan=""2"">" & vbcrlf
	response.write "          <p>" & vbcrlf
	response.write "          <fieldset>" & vbcrlf
	response.write "            <legend><strong>""Do Not Knock"" List(s)&nbsp;</strong></legend>" & vbcrlf
	response.write "            <div style=""text-align:center; color:#800000"">" & lcl_dnk_vendor_title & "</div>" & vbcrlf
	response.write "            <p>" & vbcrlf
	response.write "               <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
	response.write "                 <tr>" & vbcrlf
	response.write "                     <td>" & vbcrlf
	response.write "                         <input type=""checkbox"" name=""isOnDoNotKnockList_peddlers"" id=""isOnDoNotKnockList_peddlers"" value=""on"""     & lcl_checked_isOnDoNotKnockList_peddlers   & " />&nbsp;Do Not Knock - Peddlers<br />" & vbcrlf
	response.write "                         <input type=""checkbox"" name=""isOnDoNotKnockList_solicitors"" id=""isOnDoNotKnockList_solicitors"" value=""on""" & lcl_checked_isOnDoNotKnockList_solicitors & " />&nbsp;Do Not Knock - Solicitors" & vbcrlf
	response.write "                     </td>" & vbcrlf
	response.write "                     <td>" & lcl_dnk_description & "</td>" & vbcrlf
	response.write "                 </tr>" & vbcrlf
	response.write "               </table>" & vbcrlf
	response.write "            </p>" & vbcrlf
	response.write "          </fieldset>" & vbcrlf
	response.write "          </p>" & vbcrlf
	response.write "      </td>" & vbcrlf
	response.write "  </tr>" & vbcrlf
end if

'Display the Subscription List(s)
if lcl_orghasfeature_subscriptions OR lcl_orghasfeature_job_postings OR lcl_orghasfeature_bid_postings then
	  DisplayMaillists 
end if

response.write "  <tr>" & vbcrlf
response.write "      <td align=""center"" colspan=""2"">" & vbcrlf
response.write "          <input class=""actionbtn"" type=""button"" value=""Submit Registration Form"" onClick=""doCheck();"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf
response.write "</table>" & vbcrlf

'BEGIN: Problem Field --------------------------------------------------------
response.write "<p>" & vbcrlf
response.write "<div id=""problemtextfield1"">" & vbcrlf
response.write "  Internal Use Only, Leave Blank: <input type=""text"" name=""subjecttext"" id=""problemtextinput"" value="""" size=""6"" />" & vbcrlf
response.write "  <input type=""hidden"" name=""problemorg"" value=""" & iorgid & """ /><br />" & vbcrlf
response.write "  <strong>Please leave this field blank and remove any values that have been populated for it.</strong>" & vbcrlf
response.write "</div>" & vbcrlf
response.write "</p>" & vbcrlf
'END: Problem Field ----------------------------------------------------------

response.write "</form>" & vbcrlf
response.write "	</div>" & vbcrlf

response.write "	</div>" & vbcrlf
response.write "</div>" & vbcrlf
response.write " <p>&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</p>" & vbcrlf

response.write "<script>"
     response.write "jQuery(document).ready(function(){" & vbcrlf
     if lcl_orghasfeature_large_address_list then
        response.write "  jQuery('#validaddresslist').hide();" & vbcrlf
     end if
     response.write "});" & vbcrlf
response.write "</script>"

'------------------------------------------------------------------------------
' boolean HasResidentTypeStreets( iOrgid, sResidenttype )
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' void DisplayAddresses iorgid, sResidenttype
'------------------------------------------------------------------------------
Sub DisplayAddresses( ByVal iorgid, ByVal sResidenttype )
	Dim sSql, oRs

	sSql = "SELECT residentstreetnumber, residentstreetname FROM egov_residentaddresses_list WHERE orgid=" & iorgid
	sSql = sSql & " AND residenttype='" & sResidenttype & "' ORDER BY sortstreetname, residentstreetprefix, CAST(residentstreetnumber AS int)"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""skip_" & sResidenttype & "address"">"	
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
' void GetOrgDefaultLocation iOrgId, sDefaultCity, sDefaultState, sDefaultZip, sDefaultAreaCode
'------------------------------------------------------------------------------
Sub GetOrgDefaultLocation( ByVal iOrgId, ByRef sDefaultCity, ByRef sDefaultState, ByRef sDefaultZip, ByRef sDefaultAreaCode )
	Dim sSql, oRs

	sSql = "SELECT defaultcity, defaultstate, defaultzip, ISNULL(defaultareacode,'') AS defaultareacode "
	sSql = sSql & "FROM organizations WHERE orgid = " & iorgid 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		sDefaultCity = oRs("defaultcity")
		sDefaultState = oRs("defaultstate")
		sDefaultZip = oRs("defaultzip")
		sDefaultAreaCode = oRs("defaultareacode")
	End If 

	oRs.Close
	Set oRs = Nothing

End Sub
'------------------------------------------------------------------------------
' void DisplayLargeAddressList  sResidenttype 
'------------------------------------------------------------------------------
Sub DisplayLargeAddressList( ByVal sResidenttype )
	Dim sSql, oRs, sCompareName

	sSql = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSql = sSql & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSql = sSql & " FROM egov_residentaddresses WHERE orgid = " & iorgid & " AND residenttype = '" & sResidenttype & "' "
	sSql = sSql & " AND residentstreetname IS NOT NULL ORDER BY sortstreetname"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<input type=""text"" id=""residentstreetnumber"" name=""residentstreetnumber"" value="""" size=""8"" maxlength=""10"" /> &nbsp; "
		response.write vbcrlf & "<select id=""skip_address"" name=""skip_address"">"
		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown</option>"
		Do While Not oRs.EOF 
			sCompareName = ""
			If oRs("residentstreetprefix") <> "" Then
				sCompareName = trim(oRs("residentstreetprefix")) & " " 
			End If 
			sCompareName = sCompareName & trim(oRs("residentstreetname"))
			If oRs("streetsuffix") <> "" Then
				sCompareName = sCompareName & " " & trim(oRs("streetsuffix"))
			End If
			If oRs("streetdirection") <> "" Then
				sCompareName = sCompareName & " " & trim(oRs("streetdirection"))
			End If
			response.write vbcrlf & "<option value=""" & sCompareName & """>"
			response.write sCompareName & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 
'------------------------------------------------------------------------------
' void DisplayMaillists
'------------------------------------------------------------------------------
Sub DisplayMaillists( )
	Dim sSql, oRs, rs2

	sSql = "SELECT distributionlistid, distributionlistname, distributionlistdescription, "
	sSql = sSql & " distributionlistdisplay, orgid, isnull(distributionlisttype,'') as distributionlisttype, parentid "
	sSql = sSql & " FROM egov_class_distributionlist "
	sSql = sSql & " WHERE orgid = '" & iorgid & "' "
	sSql = sSql & " AND distributionlistdisplay = 1 "
	sSql = sSql & " AND parentid is null "
	sSql = sSql & " ORDER BY distributionlisttype, distributionlistname "

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	lcl_listtype_prev = "LIST"
	lcl_line_count    = 0

	if not oRs.eof then
  		response.write "<tr>"
  		response.write "    <td colspan=""2"">"
    response.write "        <p>"
		  response.write "        <fieldset>"
  		response.write "          <legend><strong>Subscriptions&nbsp;</strong></legend>"
		  response.write "          <p>Check the mailing list to which you would like to subscribe.</p>"

  		do while not oRs.eof
    			lcl_line_count = lcl_line_count + 1

    			if oRs("distributionlisttype") = "JOB" then
     				 lcl_listtype   = oRs("distributionlisttype")
 				 				lcl_list_title = "JOB POSTINGS"
			    elseif oRs("distributionlisttype") = "BID" then
 				 				lcl_listtype   = oRs("distributionlisttype")
 				 				lcl_list_title = "BID POSTINGS"
       else
 				 				lcl_listtype   = "LIST"
 				 				lcl_list_title = "DISTRIBUTION LISTS"
       end if

    			if lcl_listtype = "LIST" AND lcl_orghasfeature_subscriptions then
      				lcl_show = "Y"
		 	   elseif lcl_listtype = "JOB" AND lcl_orghasfeature_job_postings then
			 	     lcl_show = "Y"
			    elseif lcl_listtype = "BID" AND lcl_orghasfeature_bid_postings then 
      				lcl_show = "Y"
			    else
      				lcl_show = "N"
	 		   end if

   			'If the current parent (category) is different then the previous record then reset the variables
	 	  	'if (isnull(lcl_listtype_prev) OR lcl_listtype_prev <> oRs("distributionlisttype")) AND lcl_line_count > 1 then
    			if lcl_line_count > 1 then
			      	if IsNull(lcl_listtype_prev) then
        					lcl_listtype_prev = "LIST"
      				end if

      				if lcl_listtype_prev <> oRs("distributionlisttype") then
        					lcl_line_count = 1
      				end if
    			end if

    			if lcl_line_count = 1 then
				      if lcl_list_title <> "" then
        					if lcl_listtype <> "LIST" then
          						response.write "  </div>"
                response.write "</p>"
        					end if

        					if lcl_show = "Y" then
          						response.write "<p>"
    												response.write "  <strong>" & lcl_list_title & "</strong>"
    												response.write "  <hr size=""1"" width=""100%"" />"
    												response.write "  <div id=""" & lcl_listtype & """>"
    									end if
    						end if
    			end if

    			lcl_listtype_prev = oRs("distributionlisttype")

			    if lcl_show = "Y" then
     				'If this is a listtype of JOB/BID then check for a description and display it ONLY if one exists.
      				lcl_desc = ""

          if lcl_listtype <> "LIST" OR (lcl_listtype = "LIST" AND lcl_orghasfeature_subscriptions_distributionlist_showdesc) then
        					if oRs("distributionlistdescription") <> "" then
          						'lcl_desc = " <i>(" & oRs("distributionlistdescription") & ")</i>"
          						lcl_desc = " (" & oRs("distributionlistdescription") & ")"
        					else
          						lcl_desc = ""
        					end if
      				else
        					lcl_desc = ""
       			end if

      				response.write "    <input name=""maillist"" type=""checkbox"" value=""" & oRs("distributionlistid") & """ />&nbsp;<strong>" & oRs("distributionlistname") & "</strong>" & lcl_desc & "<br />"

     				'Check for any sub-categories
				    sSql = "SELECT * FROM egov_class_distributionlist "
      				sSql = sSql & " WHERE orgid = " & iorgid
      				sSql = sSql & " AND distributionlistdisplay = 1 "
					sSql = sSql & " AND UPPER(distributionlisttype) = '" & UCASE(oRs("distributionlisttype")) & "' "
      				sSql = sSql & " AND parentid = " & oRs("distributionlistid")
      				sSql = sSql & " ORDER BY UPPER(distributionlistname) "

      				set rs2 = Server.CreateObject("ADODB.Recordset")
      				rs2.Open sSql, Application("DSN"), adOpenForwardOnly, adLockReadOnly

      				if not rs2.eof then
        					do while not rs2.eof
   												'If this is a listtype of JOB/BID then check for a description and display it ONLY if one exists.
	    											lcl_desc = ""

    						      if lcl_listtype <> "LIST" OR (lcl_listtype = "LIST" AND lcl_orghasfeature_subscriptions_distributionlist_showdesc) then
    						  							if rs2("distributionlistdescription") <> "" then
    						    								'lcl_desc = " <i>(" & rs2("distributionlistdescription") & ")</i>"
    						    								lcl_desc = " (" & rs2("distributionlistdescription") & ")"
    						  							else
    						    								lcl_desc = ""
    						  							end if
    												else
    						  							lcl_desc = ""
    												end if

          						response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    												response.write "<input name=""maillist"" type=""checkbox"" value=""" & rs2("distributionlistid") & """ />&nbsp;<strong>" & rs2("distributionlistname") & "</strong>" & lcl_desc & "<br />"

    												rs2.MoveNext
  											loop
          end if

      				rs2.Close
      				set rs2 = nothing 
       end if

    			lcl_listtype_prev = oRs("distributionlisttype")

    			oRs.MoveNext 
  		loop

				if lcl_listtype <> "LIST" then
 						response.write "  </div>"
       response.write "</p>"
				end if

  		response.write "        </fieldset>"
    response.write "        </p>"
		  response.write "    </td>"
  		response.write "</tr>"

	end if

	oRs.Close 
	Set oRs = Nothing  

End Sub 


'------------------------------------------------------------------------------
' void DisplayNeighborhoods iorgid
'------------------------------------------------------------------------------
Sub DisplayNeighborhoods( ByVal iorgid )
	Dim sSql, oRs 

	sSql = "SELECT neighborhoodid, ISNULL(neighborhood,'') AS neighborhood FROM egov_neighborhoods "
	sSql = sSql & "WHERE orgid = " & iorgid & " ORDER BY neighborhood"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""skip_neighborhoodid"">"	
	response.write vbcrlf & "<option value=""0"">Not on List...</option>"
		
	Do While Not oRs.EOF 
		response.write vbcrlf & "<option value=""" &  oRs("neighborhoodid") & """>" & oRs("neighborhood") & "</option>"
		oRs.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub  


'------------------------------------------------------------------------------
' void DisplayAddresses_new iorgid, sResidenttype
'------------------------------------------------------------------------------
Sub DisplayAddresses_new( ByVal iorgid, ByVal sResidenttype )
	Dim sSql, oRs

	sSql = "SELECT DISTINCT ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname "
	sSql = sSql & " FROM egov_residentaddresses WHERE orgid = " & iorgid 
	sSql = sSql & " AND residenttype = '" & sResidenttype & "' ORDER BY residentstreetname, residentstreetprefix"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""skip_" & sResidenttype & "address"">"	
	response.write vbcrlf & "<option value=""0000"">Please select an address...</option>"
		
	Do While NOT oRs.EOF 
		response.write vbcrlf & "<option value=""" 
		response.write oRs("residentstreetprefix") & " " 
		response.write oRs("residentstreetname")  & """>"
		If oRs("residentstreetprefix") <> "" Then 
			response.write oRs("residentstreetprefix") & " "
		Else
			response.write "&nbsp;&nbsp; "
		End If 
		response.write oRs("residentstreetname") & "</option>"
		oRs.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub  



%>
