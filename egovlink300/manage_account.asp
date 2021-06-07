<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="../egovlink300_global/includes/inc_passencryption.asp" //-->
<% 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: manage_account.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Citizen account management.
'
' MODIFICATION HISTORY
' 1.3	10/29/2007	Steve Loar	- Added large address list selection and popup
' 1.4 08/17/09 David Boyer - Added check to see if coming from Job/Bid Postings.
'                            If "yes" then require Business Name and Work Phone.
' 1.5	04/13/2010	Steve Loar - Changes to require address for Bullhead City
' 1.6	02/22/2011	Steve Loar - Making city, state and zip required optionally
' 1.8	10/04/2011	Steve Loar - Added gender selection pick
' 1.9   2014-06-11  Jerry Felix - revised the email regex to be more permissive for new TLDs
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 Dim sError, oRegOrg, bAddressChanged, errormsg, sTopDisplayMessage, sRedirect
 Dim sSaveButtonText, oRs, bAddressRequired, lcl_success, bShowGenderPicks, bGenderIsRequired

 set oRegOrg = New classOrganization

 errormsg        = ""
 sSaveButtonText = "Save Changes"
 lcl_success     = request("success")

'------------------------------------------------------------------------------
' session("RedirectPage") = "pool_pass/poolpass_select.asp" 
' session("RedirectLang") = "Return to Membership Purchase"

'If they do not have a userid set, take them to the login page automatically
' if request.cookies("userid") = "" or request.cookies("userid") = "-1" then
'	   session("LoginDisplayMsg") = "Please sign in first and then we'll send you right along."
'   	response.redirect "../user_login.asp"
' end if
'------------------------------------------------------------------------------
If request("p") = "dnk" Then 
	session("RedirectPage") = "manage_account.asp?p=dnk"
	session("redirectlang") = "Manage Account"

	'If they do not have a userid set, take them to the login page automatically
	If request.cookies("userid") = "" Or request.cookies("userid") = "-1" Then 
		session("LoginDisplayMsg") = "Please sign in first and then we'll send you right along."
		response.redirect "user_login.asp?p=dnk"
	End If 
End If 

If session("DisplayMsg") <> "" Then  
  	sTopDisplayMessage    = "<div id=""pagetitle"">" & session("DisplayMsg") & "</div>"
  	session("DisplayMsg") = ""
  	sSaveButtonText       = "Save Changes and Continue"
Else  
  	sTopDisplayMessage    = ""
End If 

'Check for Org Features
lcl_orghasfeature_permit_setup         = orghasfeature(iOrgID,"permit setup")
lcl_orghasfeature_hide_family_mgmt     = orghasfeature(iOrgID,"hide family mgmt")
lcl_orghasfeature_hasfamily            = orghasfeature(iOrgID,"hasfamily")
lcl_orghasfeature_activities           = orghasfeature(iOrgID,"activities")
lcl_orghasfeature_subscriptions        = orghasfeature(iOrgID,"subscriptions")
lcl_orghasfeature_large_address_list   = orghasfeature(iOrgID,"large address list")
lcl_orghasfeature_no_emergency_contact = orghasfeature(iOrgID,"no emergency contact")
lcl_orghasfeature_donotknock           = orghasfeature(iOrgID,"donotknock")
lcl_orghasfeature_citizenregistration_novalidate_address = orghasfeature(iOrgID,"citizenregistration_novalidate_address")
bShowGenderPicks = orgHasFeature( iOrgId, "display gender pick" )
bGenderIsRequired = orgHasFeature( iOrgId, "gender required" )

'Check for org "edit displays"
lcl_orghasdisplay_donotknock_list_description = orghasdisplay(iorgid,"donotknock_list_description")
lcl_orghasdisplay_subscription_screenmsg      = orghasdisplay(iorgid,"subscription_screenmsg")

'Determine if Org has Neighborhoods
lcl_orghasneighborhoods = orghasneighborhoods(iOrgID)

'Determine if the address is required or not
bAddressRequired   = OrgHasFeature( iOrgId, "registration req address" )
lcl_required_label = ""

If bAddressRequired Then 
	lcl_required_label = "<font color=""#ff0000"">*</font>"
End If 

'Check for org "edit displays"
lcl_orghasdisplay_citizen_register_maint_addressinfo = orghasdisplay(iorgid,"citizen_register_maint_addressinfo")

bAddressChanged = False

If request.ServerVariables("REQUEST_METHOD") = "POST" Then 
   	'Call ProcessRecords()
   	UpdateExistingUser request("userid") 

   	If errormsg = "" Then 
	 	    If lcl_orghasneighborhoods Then 
     		  	UpdateUserNeighborhood request("userid"), request("egov_users_neighborhoodid")
     	 End If 

    		'check for address changes and unflag the residency verified flag in egov_users
     		If request("skip_old_egov_users_useraddress") <> request("egov_users_useraddress") Then 
       			bAddressChanged = True
     		Else 
       			If request("skip_old_egov_users_userunit") <> request("egov_users_userunit") Then 
         				bAddressChanged = True
       			Else 
          			If request("skip_old_egov_users_neighborhoodid") <> request("egov_users_neighborhoodid") Then 
            				bAddressChanged = True
        	 			Else 
           					If request("skip_old_egov_users_usercity") <> request("egov_users_usercity") Then 
             						bAddressChanged = True
          		 			Else 
            	 					If request("skip_old_egov_users_userstate") <> request("egov_users_userstate") Then 
               							bAddressChanged = True
            			 			Else 
               	 					If request("skip_old_egov_users_userzip") <> request("egov_users_userzip") Then 
                  							bAddressChanged = True
              				 			End If 
              					End If 
           					End If 
         				End If 
        		End If 
       End If 

     		If bAddressChanged Then 
       			UpdateResidencyVerified request("userid")
     		End If 

     		If lcl_orghasfeature_permit_setup Then 
      			'Update any permit applicants and Primary Contacts where the permit is still open. Pull them, then loop through the set
       			sSql = "SELECT P.permitid, C.permitcontactid "
    						sSql = sSql & " FROM egov_permits P, egov_permitcontacts C, egov_permitstatuses S "
   	 					sSql = sSql & " WHERE P.permitid = C.permitid AND P.permitstatusid = S.permitstatusid "
   		 				sSql = sSql & " AND (isapplicant = 1 OR isprimarycontact = 1) AND S.iscompletedstatus = 0 "
   			 			sSql = sSql & " AND S.cansavechanges = 1 AND S.changespropagate = 1  AND C.userid = " & request("userid") 

    						Set oRs = Server.CreateObject("ADODB.Recordset")
    						oRs.Open sSql, Application("DSN"), 0, 1

    						do while not oRs.eof
    			  				sSql = "UPDATE egov_permitcontacts "
    			  				sSql = sSql & " SET firstname = '" & dbsafe(request("egov_users_userfname")) & "' "
		   								sSql = sSql & ", lastname = '" & dbsafe(request("egov_users_userlname")) & "' "

 		  								If request("egov_users_userbusinessname") = "" Then 
	 	  				  					sSql = sSql & ", company = NULL "
		   								Else 
		   				  					sSql = sSql & ", company = '" & dbsafe(request("egov_users_userbusinessname")) & "' "
		   								End If 

 		  								If request("egov_users_useraddress") = "" Then
	 	  				  					sSql = sSql & ", address = NULL "
		   								Else 
		   				  					sSql = sSql & ", address = '" & dbsafe(request("egov_users_useraddress")) & "' "
		   								End If 

		  	 							If request("egov_users_usercity") = "" Then
		  		 		  					sSql = sSql & ", city = NULL "
		  			 	 			Else 
		  				   					sSql = sSql & ", city = '" & dbsafe(request("egov_users_usercity")) & "' "
		  					 			End If 

 		  								If request("egov_users_userstate") = "" Then
	 	  				  					sSql = sSql & ", state = NULL "
		   				 			Else 
		   				  					sSql = sSql & ", state = '" & dbsafe(request("egov_users_userstate")) & "' "
		   								End If 

  	  				 			If request("egov_users_userzip") = "" Then
		   					  				sSql = sSql & ", zip = NULL "
		   								Else 
		   							  		sSql = sSql & ", zip = '" & dbsafe(request("egov_users_userzip")) & "' "
		  	 							End If 

		  		 						If request("egov_users_useremail") = "" Then 
					           sSql = sSql & ", email = NULL "
		  				 				Else 
           					sSql = sSql & ", email = '" & dbsafe(request("egov_users_useremail")) & "' "
		  						 		End If 

 		  								If request("egov_users_userhomephone") = "" Then 
           					sSql = sSql & ", phone = NULL "
		   								Else 
           					sSql = sSql & ", phone = '" & dbsafe(request("egov_users_userhomephone")) & "' "
		   								End If 

 		  								If request("egov_users_userfax") = "" Then 
           					sSql = sSql & ", fax = NULL "
		   								Else 
           					sSql = sSql & ", fax = '" & dbsafe(request("egov_users_userfax")) & "' "
		   								End If 

		  	 							If request("egov_users_usercell") = "" Then 
           					sSql = sSql & ", cell = NULL "
		  			 					Else 
           					sSql = sSql & ", cell = '" & dbsafe(request("egov_users_usercell")) & "' "
		  					 			End If 

 		  								If request("skip_emailnotavailable") = "on" then
           					sSql = sSql & ", emailnotavailable = 1" 
		   								Else 
           					sSql = sSql & ", emailnotavailable = 0" 
		   								End If 

 		  								If request("egov_users_userpassword") = "" Then 
           					sSql = sSql & ", password = NULL "
		   								Else 
           					sSql = sSql & ", userpassword = NULL, password = '" & createHashedPassword(request("egov_users_userpassword")) & "' "
		   								End If 

		  	 							If request("egov_users_residenttype") = "" Then 
           					sSql = sSql & ", residenttype = NULL "
		  			 					Else 
										xResidentType = request("egov_users_residenttype")
										if xResidentType = "'N'" and request("egov_users_userzip") = "94025" and lcase(request("egov_users_usercity")) = "menlo park" then xResidentType = "'U'"

           					sSql = sSql & ", residenttype = '" & dbsafe() & "' "
		  					 			End If 

 		  								If request("egov_users_userworkphone") = "" Then 
           					sSql = sSql & ", userworkphone = NULL "
		   								Else 
           					sSql = sSql & ", userworkphone = '" & dbsafe(request("egov_users_userworkphone")) & "' "
		   								End If 

 		  								If request("egov_users_emergencyphone") = "" Then 
           					sSql = sSql & ", emergencyphone = NULL "
		   								Else 
           					sSql = sSql & ", emergencyphone = '" & dbsafe(request("egov_users_emergencyphone")) & "' "
		   								End If 

		  	 							If request("egov_users_neighborhoodid") = "" Then 
           					sSql = sSql & ", neighborhoodid = NULL "
		  			 					Else 
           					sSql = sSql & ", neighborhoodid = " & request("egov_users_neighborhoodid") 
		  					 			End If 

		  						 		If request("egov_users_userunit") = "" Then 
          	 				sSql = sSql & ", userunit = NULL "
		  								 Else 
          			 		sSql = sSql & ", userunit = '" & dbsafe(request("egov_users_userunit")) & "' "
 		  								End If 

 		  								If request("egov_users_emergencycontact") = "" Then 
           					sSql = sSql & ", emergencycontact = NULL "
	 	  								Else 
           					sSql = sSql & ", emergencycontact = '" & dbsafe(request("egov_users_emergencycontact")) & "' "
		   								End If 

		   								If request("egov_users_userbusinessaddress") = "" Then
           					sSql = sSql & ", userbusinessaddress = NULL "
		  		 						Else 
           					sSql = sSql & ", userbusinessaddress = '" & dbsafe(request("egov_users_userbusinessaddress")) & "' "
		  				 				End If

		  					 			sSql = sSql & " WHERE permitid = " & oRs("permitid") & " AND permitcontactid = " & oRs("permitcontactid")
		  						 		'response.write sSql & "<br /><br />"

		  							 	RunSQLStatement sSql		' In common.asp

		  								 oRs.MoveNext
  				 			loop 

       			oRs.Close
       			Set oRs = Nothing 
     		End If 

   	 	'Take them back to where they came from if needed
    	 	If Session("RedirectPage") <> "" And errormsg = "" Then 
       			sRedirect               = Session("RedirectPage") 
      	 		Session("RedirectPage") = ""

'      		 	response.redirect sRedirect
     		End If 
   	End If 

    If errormsg = "" Then 
      	errormsg = "<span id=""information_updated"">&nbsp;Information Updated - " & Now() & "&nbsp;</span>"
   	End If 

End If 

'Set the session for the family update form to come back here
session("ManageURL")  = "manage_account.asp"
session("ManageLang") = "Return to Manage Account"

If session("RedirectLang") <> "" Then 
	sBackLang     = session("RedirectLang")
	sRedirectPage = session("RedirectPage")
Else 
	sBackLang     = "Back"
	sRedirectPage = "manage_account.asp"
End If 

session("RedirectLang") = sBackLang
session("RedirectPage") = sRedirectPage

'USER VALUES
Dim sFirstName,sLastName,sGender, sAddress,sCity,sState,sZip,sPhone,sEmail,sFax,sBusinessName,sDayPhone,sPassword,iUserID
Dim bHasResidentStreets, bFound, sResidenttype, sBusinessAddress, bHasBusinessStreets, sWorkPhone, iNeighborhoodId
Dim sEmergencyContact, sEmergencyPhone, sCell, sUserUnit, sIsOnDoNotKnockList_peddlers, sIsOnDoNotKnockList_solicitors

GetRegisteredUserValues

'Check to see if the user came from a Job/Bid Postings screen.
If request("fromPostings") <> "" Then 
	lcl_fromPostings = UCASE(request("fromPostings"))
Else 
	lcl_fromPostings = ""
End If 

If request("posting_id") <> "" Then 
	lcl_posting_id = request("posting_id")
Else 
	lcl_posting_id = ""
End If 

If request("listtype") <> "" Then 
	lcl_listtype = request("listtype")
Else 
	lcl_listtype = ""
End If 

If request("dlistid") <> "" Then 
	lcl_dlistID = request("dlistid")
Else 
	lcl_dlistID = ""
End If 

'Default the label names of the required fields.
lcl_businessNameLabel = "Business Name:"
lcl_workPhoneLabel    = "Work Phone:"
lcl_phoneNumberLabel  = "Phone Number:"

'Set up the labels for all required fields.
lcl_isRequiredField_start = "<span class=""cot-text-emphasized"" title=""This field is required""><span class=""cot-text-emphasized"" style=""color:#ff0000;"">*</span>"
lcl_isRequiredField_end   = "</span>"

lcl_isRequired_BusinessName = "N"
lcl_isRequired_WorkPhone    = "N"
lcl_isRequired_PhoneNumber  = "Y"

if lcl_fromPostings = "Y" then
	lcl_isRequired_BusinessName = "Y"
	lcl_isRequired_WorkPhone    = "Y"
end if

'If subscribing to subscriptions and redirected from the Subscriptions screen, 
'because feature "Check for User Address when subscribing (public-side)" is enabled
'then do NOT require the Phone Number
if lcl_success = "SU_NA" OR lcl_success = "SA_NA" then
	lcl_isRequired_PhoneNumber = "N"
end if

'Determine if the field is required or not.  For those fields that are required proceed the label with a red (*)
If lcl_isRequired_BusinessName = "Y" Then 
	lcl_businessNameLabel = lcl_isRequiredField_start & lcl_businessNameLabel & lcl_isRequiredField_end
End If 

If lcl_isRequired_WorkPhone = "Y" Then 
	lcl_workPhoneLabel = lcl_isRequiredField_start & lcl_workPhoneLabel & lcl_isRequiredField_end
End If 

if lcl_isRequired_PhoneNumber = "Y" then
	lcl_phoneNumberLabel = lcl_isRequiredField_start & lcl_phoneNumberLabel & lcl_isRequiredField_end
end if

lcl_checked_isOnDoNotKnockList_peddlers   = ""
lcl_checked_isOnDoNotKnockList_solicitors = ""

If sIsOnDoNotKnockList_peddlers Then 
	lcl_checked_isOnDoNotKnockList_peddlers = " checked=""checked"""
End If 

If sIsOnDoNotKnockList_solicitors Then 
	lcl_checked_isOnDoNotKnockList_solicitors = " checked=""checked"""
End If 

'Check for a screen message
lcl_onload            = ""
lcl_subscriptions_msg = ""
lcl_msg               = ""

if lcl_success <> "" then
	if lcl_success = "SU_NA" OR lcl_success = "SA_NA" then
		lcl_success = replace(lcl_success,"_NA","")

		'Check for an "edit_display" to determine if there is additional information the org wants to show with the message
		'if lcl_orghasdisplay_subscription_screenmsg AND sAddress = "" then
		if lcl_orghasdisplay_subscription_screenmsg then
			lcl_subscriptions_msg = getOrgDisplay(iorgid,"subscription_screenmsg")
		end if
	end if

	lcl_msg    = setupScreenMsg(lcl_success)
	lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
end If

%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	<title>E-Gov Services - <%=oRegOrg.GetOrgName()%>, <%=oRegOrg.GetState()%> - Manage Account</title>

	<link rel="stylesheet" type="text/css" href="css/styles.css" />
	<link rel="stylesheet" type="text/css" href="global.css" />
	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

<style type="text/css">
  div.groupSmall2 {
		   padding-left:  4px;
  			padding-right: 0;
		  	width:         594px;
		}

  #subscriptionsMsg {
     margin-bottom: 5px;
  }

  #screenMsg {
     color:       #ff0000;
     font-size:   10pt;
     font-weight: bold;
     text-align:  right;
  }

  #information_updated {
     color:            #0000ff;
     background-color: #e0e0e0;
     border:           solid 1px #000000;
     font-weight:      bold;
  }
</style>

  	<script type="text/javascript" src="scripts/jquery-1.9.1.min.js"></script>
	<script language="javascript" src="prototype/prototype-1.6.0.2.js"></script>

	<script language="javascript" src="scripts/modules.js"></script>
	<script language="javascript" src="scripts/easyform.js"></script>
	<script language="javascript" src="scripts/ajaxLib.js"></script>
	<script language="javascript" src="scripts/removespaces.js"></script>
	<script language="javascript" src="scripts/setfocus.js"></script>

	<script language="javascript">
	<!--
  jQuery.noConflict();  //Allows us to use jQuery
		var selectedvalue = '0000';
		var winHandle;
		var w = (screen.width - 640)/2;
		var h = (screen.height - 450)/2;

		function doCheck() {
  		// If they are using the large address feature
		 	var exists = eval(document.register["residentstreetnumber"]);
			 if(exists) {
       <%
        'This feature is ENABLED then it DISABLES the large address validation and simply does the form validation.
         if not lcl_orghasfeature_citizenregistration_novalidate_address then
            response.write "// If a street number was entered" & vbcrlf
            response.write "if (document.register.residentstreetnumber.value != '') {" & vbcrlf
            response.write "				checkAddress( 'FinalCheckOLD', 'yes' );" & vbcrlf
            response.write "} else {" & vbcrlf
            response.write "				//checkDuplicateCitizens( 'FinalUserCheckFailed' );" & vbcrlf
									response.write "document.register.egov_users_residenttype.value = 'N';"
            response.write "				Validate();" & vbcrlf
            response.write "}" & vbcrlf
         else
									response.write "document.register.egov_users_residenttype.value = 'N';"
            response.write "Validate();" & vbcrlf
         end if
       %>
 			} else {
									document.register.egov_users_residenttype.value = "N";
    			Validate();
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
				alert("This is a valid address in our system.");
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
				//winHandle.focus();
				//PopAStreetPicker('FinalCheck', 'yes');
				CheckResults('');
			}
		}

		function finalCheckValidate()
		{
			Validate();
		}

		function OkToValidate( sReturn )
		{
			//finish the validation routine 
			Validate();
		}

		function openWin2(url, name) 
		{
		  popupWin = window.open(url, name,"resizable,width=500,height=450");
		}

		function UpdateFamily()
		{
			location.href='family_members.asp';
		}

		function FamilyList()
		{
			location.href='family_list.asp';
			//location.href='family_list.asp?userid=' + $("userid").value;
		}

		function Validate() 
		{
			var msg="";

			// Check the email 
			//var rege = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*\.(\w{2}|(com|net|org|edu|mil|gov|biz|us))$/;
			var rege = /.+@.+\..+/i;

			var Ok = rege.test(document.register.egov_users_useremail.value);

			if (! Ok)
			{
				msg+="The email must be in a valid format.\n";
			}

			// check the passwords
			if(document.register.egov_users_userpassword.value != document.register.skip_userpassword2.value)
			{
				msg+="The passwords you have entered do not match.\n";
			}

			// Gender validation if present and is required
<%			If bShowGenderPicks And bGenderIsRequired Then	%>
				if (document.register.egov_users_gender.value == 'N') 
				{
					msg+="Selection of a valid gender is required.\n";
				}
<%			End If		%>

			// Check the first name
			if (document.register.egov_users_userfname.value == "")
			{
				msg+="Your first name cannot be left blank.\n";
			}

			// Check the last name
			if (document.register.egov_users_userlname.value == "")
			{
				msg+="Your last name cannot be left blank.\n";
			}

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
			else
			{
				document.register.egov_users_userworkphone.value = '';
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
			else
			{
				document.register.egov_users_userfax.value = '';
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
			else
			{
				document.register.egov_users_usercell.value = '';
			}

			if (document.register.egov_users_emergencyphone)
			{
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
				else
				{
					document.register.egov_users_emergencyphone.value = '';
				}
			}

			// set the phone number
			document.register.egov_users_userhomephone.value = document.register.skip_user_areacode.value + document.register.skip_user_exchange.value + document.register.skip_user_line.value;
			if (document.register.egov_users_userhomephone.value != "" ) {
 				var hPhone = document.register.egov_users_userhomephone.value;
	 			if (hPhone.length < 10) {
   					msg += "The Phone Number must be a valid phone number, including area code, or blank.\n";
			 	} else {
   					var rege = /^\d+$/;
	   				var Ok = rege.exec(document.register.egov_users_userhomephone.value);
		 	  		if ( ! Ok ) {
   			  			msg += "The Phone Number must be a valid phone number, including area code, or blank.\n";
  			 		}
				 }
			}
<% if lcl_isRequired_PhoneNumber = "Y" then %>
			else {
  				msg+="The Phone Number cannot be blank.\n";
			}
<% end if %>

			iFromPostings = '<%=lcl_fromPostings%>';

			if(iFromPostings == "Y") 
			{
				if(document.getElementById("egov_users_userbusinessname").value == "") 
				{
					msg += "Required Field Missing: Business Name\n";
				}

				if(document.getElementById("egov_users_userworkphone").value == "") 
				{
					msg += "Required Field Missing: Work Phone\n";
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

<%			If bAddressRequired Then	
				If lcl_orghasfeature_large_address_list Then	%>
					if( ($F('residentstreetnumber') == '' ||  $('skip_address').getValue() == '0000') && $F("egov_users_useraddress") == "" ) 
					{
						msg += "Required Field Missing: Address\n";
					}
<%				Else	%>
					if ($F("egov_users_useraddress") == "")
					{
						msg += "Required Field Missing: Address\n";
					}
<%				End If											%>

				if ($("egov_users_usercity").value == "")
				{
					msg += "Required Field Missing: City\n";
				}

				if ($("egov_users_userstate").value == "")
				{
					msg += "Required Field Missing: State\n";
				}

				if ($("egov_users_userzip").value == "")
				{
					msg += "Required Field Missing: Zip\n";
				}

<%			End If						%>

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
						var selectedvalue = element.options[element.selectedIndex].value;

						//alert( selectedvalue );
						//  0000 is the first pick that we do not want
						if (selectedvalue != "0000")
						{
							document.register.egov_users_useraddress.value = document.register.residentstreetnumber.value + ' ' + selectedvalue;
							document.register.egov_users_residenttype.value = "R";
							bUsedAddressDropdown = true;
						}
					}
				}
			}

			if(msg != "")
			{
<%				If lcl_orghasfeature_large_address_list Then	%>
					if( $F('residentstreetnumber') != '' &&  $('skip_address').getValue() != '0000' ) 
					{
						$('egov_users_useraddress').value = '';
					}
<%				End If											%>
				msg="Your changes could not be saved for the following reasons.\n\n" + msg;
				msg+="\nPlease correct this and try saving again.";
				alert(msg);
				return;
			}
			else 
			{	
				if (validateForm('register')) 
				{ 
					// Validate that the email is not being used by anyone else via Ajax
					doAjax('includes/checkemail.asp', 'email=' + document.register.egov_users_useremail.value + '&uid=' + document.register.userid.value + '&orgid=' + document.register.egov_users_orgid.value, 'emailCheck', 'get', '0')
					//document.register.submit();  -- This is in the function emailCheck() now
				}
				else
				{
<%				If lcl_orghasfeature_large_address_list Then	%>
					if( $F('residentstreetnumber') != '' &&  $('skip_address').getValue() != '0000' ) 
					{
						$('egov_users_useraddress').value = '';
					}
<%				End If											%>

				}
			}
		}

		function emailCheck( check )
		{
			if (check == 'OK')
			{
				document.register.submit(); 
			}
			else
			{
				alert( 'The email you have entered is in use by another.\n\nPlease enter a different email.');
				// Put the address line back to blank if needed
				if (selectedvalue != "0000")
				{
					document.register.egov_users_useraddress.value = '';
				}
				document.register.egov_users_useremail.focus();
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

		function displayScreenMsg(iMsg) {
		  if(iMsg!="") {
			 document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
			 window.setTimeout("clearScreenMsg()", (10 * 1000));
		  }
		}

		function clearScreenMsg() {
		  document.getElementById("screenMsg").innerHTML = "&nbsp;";
		}

	//-->
	</script>
</head>

<!--#Include file="include_top.asp"-->
<%
'BEGIN: Body Content ---------------------------------------------------------
response.write "<font class=""pagetitle"">Welcome to " & oRegOrg.GetOrgName() & ", " & oRegOrg.GetState() & ", Account Maintenance</font><br />" & vbcrlf

RegisteredUserDisplay( "" )

response.write "<div id=""content"">" & vbcrlf
response.write "	 <div id=""centercontent"">" & vbcrlf
response.write      sTopDisplayMessage
response.write "    <div id=""screenMsg"">&nbsp;</div>" & vbcrlf
response.write "    <div id=""subscriptionsMsg"">" & lcl_subscriptions_msg & "</div>" & vbcrlf
response.write "    <div id=""manageaccountnav"">" & vbcrlf

if iorgid = "228" and session("retfromtop") <> "" then 
	response.write "<input type=""button"" class=""button"" onclick=""location.href='" & session("retfromtop") & "'"" value=""Back"" /><br /><br />"
end if

'If coming from Job/Bid Postings
if lcl_fromPostings = "Y" then
	response.write "<input type=""button"" name=""returnToPostingsButton"" id=""returnToPostingsButton"" value=""Return to Posting"" class=""button"" onclick=""location.href='postings_info.asp?posting_id=" & lcl_posting_id & "&listtype=" & lcl_listtype & "&dlistid=" & lcl_dlistID & "';"" />" & vbcrlf
end if

if not lcl_orghasfeature_hide_family_mgmt then
	if lcl_orghasfeature_hasfamily then
		response.write "<input class=""actionbtn"" type=""button"" value=""Manage Your Family Members"" onclick=""javascript:FamilyList();"" />&nbsp;&nbsp;" & vbcrlf
	else
		if lcl_orghasfeature_activities then
			response.write "<input class=""actionbtn"" type=""button"" value=""Manage Your Family Members"" onclick=""javascript:UpdateFamily();"" />&nbsp;&nbsp;" & vbcrlf
		end if
	end if
end if

if lcl_orghasfeature_subscriptions then
	response.write "<input class=""actionbtn"" type=""button"" value=""Subscribe to Email Communications"" onclick=""javascript:location.href='manage_mail_lists.asp';"" />" & vbcrlf
end if

response.write "    </div>" & vbcrlf
response.write "    <div class=""box_header4"">Your " & sOrgName & " Account Information</div>" & vbcrlf
response.write "    <div class=""groupSmall2"">" & vbcrlf
response.write "      <form name=""register"" action=""manage_account.asp"" method=""post"" />" & vbcrlf
response.write "      	 <input type=""hidden"" name=""columnnameid"" value=""userid"" />" & vbcrlf
response.write "      	 <input type=""hidden"" id=""userid"" name=""userid"" value=""" & iuserid & """ />" & vbcrlf
response.write "      	 <input type=""hidden"" name=""egov_users_orgid"" value=""" & iorgid & """ />" & vbcrlf
response.write "      	 <input type=""hidden"" name=""ef:egov_users_useremail-text/req"" value=""Email Address"" />" & vbcrlf
response.write "      	 <input type=""hidden"" name=""ef:egov_users_userpassword-text/req"" value=""Password 1"" />" & vbcrlf
response.write "      	 <input type=""hidden"" name=""ef:skip_userpassword2-text/req"" value=""Password 2"" />" & vbcrlf
response.write "      	 <input type=""hidden"" name=""ef:egov_users_userfname-text/req"" value=""First name"" />" & vbcrlf
response.write "      	 <input type=""hidden"" name=""ef:egov_users_userlname-text/req"" value=""Last name"" />" & vbcrlf
response.write "      	 <input type=""hidden"" name=""egov_users_residenttype"" value=""" & sResidenttype & """ />" & vbcrlf
response.write "      	 <input type=""hidden"" name=""fromPostings"" id=""fromPostings"" value=""" & lcl_fromPostings & """ />" & vbcrlf
response.write "      	 <input type=""hidden"" name=""listtype"" id=""listtype"" value=""" & lcl_listtype & """ />" & vbcrlf
If Not bShowGenderPicks Then 
	response.write vbcrlf & "<input type=""hidden"" id=""egov_users_gender"" name=""egov_users_gender"" value=""N"" />"
End If 

if Not lcl_orghasfeature_donotknock then
	response.write "      	 <input type=""hidden"" name=""isOnDoNotKnockList_peddlers"" id=""isOnDoNotKnockList_peddlers"" value="""" />" & vbcrlf
	response.write "      	 <input type=""hidden"" name=""isOnDoNotKnockList_solicitors"" id=""isOnDoNotKnockList_solicitors"" value="""" />" & vbcrlf
end if

response.write "<table cellpadding=""2"" cellspacing=""0"" border=""0"">" & vbcrlf

if errormsg <> "" then
	response.write "  <tr><td colspan=""2"">" & errormsg & "</td></tr>" & vbcrlf
end if

response.write "  <tr>" & vbcrlf
response.write "      <td colspan=""2"" align=""center"">" & vbcrlf
response.write "          <input class=""actionbtn"" type=""button"" value=""" & sSaveButtonText & """ onclick=""javascript:doCheck();"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

response.write "  <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
response.write "  <tr>" & vbcrlf
response.write "      <td>&nbsp;</td>" & vbcrlf
response.write "      <td><span class=""cot-text-emphasized""><font color=""#ff0000"">*</font></span> <strong>Indicates a required field.</strong></td>" & vbcrlf
response.write "  </tr>" & vbcrlf

response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write "          <span class=""cot-text-emphasized"" title=""This field is required""><span class=""cot-text-emphasized""><font color=""#ff0000"">*</font></span>" & vbcrlf
response.write "            Email:" & vbcrlf
response.write "          </span>" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""text"" value=""" & sEmail & """ name=""egov_users_useremail"" style=""width:300px;"" maxlength=""100"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write "          <span class=""cot-text-emphasized"" title=""This field is required""><span class=""cot-text-emphasized""><font color=""#ff0000"">*</font></span>" & vbcrlf
response.write "            Password:" & vbcrlf
response.write "          </span>" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""password"" value=""" & sPassword & """ name=""egov_users_userpassword"" size=""25"" maxlength=""50"" autocomplete=""new-password"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write "          <span class=""cot-text-emphasized"" title=""This field is required""><span class=""cot-text-emphasized""><font color=""#ff0000"">*</font></span>" & vbcrlf
response.write "            Verify Password:" & vbcrlf
response.write "          </span>" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""password"" value=""" & sPassword & """ name=""skip_userpassword2"" size=""25"" maxlength=""50"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write "          <span class=""cot-text-emphasized"" title=""This field is required""><span class=""cot-text-emphasized""><font color=""#ff0000"">*</font></span>" & vbcrlf
response.write "           	First Name:" & vbcrlf
response.write "          </span>" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <span class=""cot-text-emphasized"" title=""This field is required"">" & vbcrlf
response.write "       			  <input type=""text"" value=""" & sFirstName & """ name=""egov_users_userfname"" style=""width:300px;"" maxlength=""100"" />" & vbcrlf
response.write "          </span>" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write "          <span class=""cot-text-emphasized"" title=""This field is required""><span class=""cot-text-emphasized""><font color=""#ff0000"">*</font></span>" & vbcrlf
response.write "            Last Name:" & vbcrlf
response.write "          </span>" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <span class=""cot-text-emphasized"" title=""This field is required"">" & vbcrlf
response.write "         			<input type=""text"" value=""" & sLastName & """ name=""egov_users_userlname"" style=""width:300px;"" maxlength=""100"" />" & vbcrlf
response.write "       			</span>" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

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
'bHasResidentStreets = False  ' For Bullhead City testing
bFound = False

If bHasResidentStreets Then 
	If Not lcl_orghasfeature_large_address_list Then 
        response.write "  <tr>" & vbcrlf
        response.write "      <td class=""label"" align=""right"">" & vbcrlf
        response.write "          Resident Street:" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td nowrap=""nowrap"">" & vbcrlf
                                  DisplayAddresses iorgid, "R", sAddress, bFound
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
     Else 
    	'Show the large address list solution
        response.write "  <tr>" & vbcrlf
        response.write "      <td class=""label"" align=""right"" valign=""top"" nowrap=""nowrap"" id=""addblock"">" & lcl_required_label & vbcrlf
        response.write "          Resident Address:" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td nowrap=""nowrap"">" & vbcrlf

                                		BreakOutAddress sAddress, sStreetNumber, sStreetName   'In common.asp
                                		DisplayLargeAddressList iOrgID, "R", lcl_orghasfeature_citizenregistration_novalidate_address, _
                                                          sStreetNumber, sStreetName, bFound

                               		'If this feature is ENABLED then it DISABLES the large address validation and 
                               		'simply does the form validation.
                                 	if not lcl_orghasfeature_citizenregistration_novalidate_address then
                                  			response.write "<input type=""button"" class=""button"" value=""Validate Address"" onclick=""checkAddress( 'CheckResults', 'no' );"" />" & vbcrlf
                                  end if

        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
     End If 
End If 

response.write vbcrlf & "<tr>"
response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">" & lcl_required_label

if bHasResidentStreets then
	response.write "Address (if not listed):"
else
	response.write lcl_required_label
	response.write "Address:"
end If

response.write "</td>"
response.write "<td>"

lcl_address = ""

if not bfound then
	lcl_address = sAddress
end if 

response.write "          <input type=""hidden"" value=""" & sAddress & """ name=""skip_old_egov_users_useraddress"" />" & vbcrlf
response.write "       			<input type=""text"" value=""" & lcl_address & """ name=""egov_users_useraddress"" id=""egov_users_useraddress"" style=""width:300px;"" maxlength=""100"" />" & vbcrlf
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
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf
response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"" nowrap=""nowrap"">" & vbcrlf
response.write "          Resident Unit:" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""hidden"" value=""" & sUserUnit & """ name=""skip_old_egov_users_userunit"" />" & vbcrlf
response.write "          <input type=""text"" value=""" & sUserUnit & """ name=""egov_users_userunit"" size=""10"" maxlength=""10"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

if lcl_orghasneighborhoods then
	response.write " <tr>" & vbcrlf
	response.write "     <td class=""label"" align=""right"">" & vbcrlf
	response.write "         Neighborhood:" & vbcrlf
	response.write "     </td>" & vbcrlf
	response.write "     <td>" & vbcrlf
	response.write "         <input type=""hidden"" value=""" & iNeighborhoodId & """ name=""skip_old_egov_users_neighborhoodid"" />" & vbcrlf
			   DisplayNeighborhoods iOrgid, iNeighborhoodId
	response.write "     </td>" & vbcrlf
	response.write " </tr>" & vbcrlf
end if

response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write            lcl_required_label 
response.write "          City:" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""hidden"" value=""" & sCity & """ name=""skip_old_egov_users_usercity"" />" & vbcrlf
response.write "          <input type=""text"" value=""" & sCity & """ id=""egov_users_usercity"" name=""egov_users_usercity"" style=""width:300px;"" maxlength=""40"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf
response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write            lcl_required_label
response.write "          State:" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""hidden"" value=""" & sState & """ name=""skip_old_egov_users_userstate"" />" & vbcrlf
response.write "          <input type=""text"" value=""" & sState & """ id=""egov_users_userstate"" name=""egov_users_userstate"" size=""2"" maxlength=""2"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf
response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write            lcl_required_label
response.write "          ZIP:" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""hidden"" value=""" & sZip & """ name=""skip_old_egov_users_userzip"" />" & vbcrlf
response.write "          <input type=""text"" value=""" & sZip & """ id=""egov_users_userzip"" name=""egov_users_userzip"" size=""10"" maxlength=""10"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf
response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & lcl_phoneNumberLabel & "</td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""hidden"" value=""" & sDayPhone & """ name=""egov_users_userhomephone"" />" & vbcrlf
response.write "      			(<input class=""phonenum"" type=""text"" value=""" & Left(sDayPhone,3)  & """ name=""skip_user_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />)&nbsp;" & vbcrlf
response.write "       			<input class=""phonenum"" type=""text"" value=""" & Mid(sDayPhone,4,3) & """ name=""skip_user_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />&ndash;" & vbcrlf
response.write "       			<input class=""phonenum"" type=""text"" value=""" & Right(sDayPhone,4) & """ name=""skip_user_line"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf
response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write "          Cell Phone:" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "       			<input type=""hidden"" value=""" & sCell & """ name=""egov_users_usercell"" />" & vbcrlf
response.write "      			(<input class=""phonenum"" type=""text"" value=""" & Left(sCell,3)  & """ name=""skip_cell_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />)&nbsp;" & vbcrlf
response.write "       			<input class=""phonenum"" type=""text"" value=""" & Mid(sCell,4,3) & """ name=""skip_cell_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />&ndash;" & vbcrlf
response.write "       			<input class=""phonenum"" type=""text"" value=""" & Right(sCell,4) & """ name=""skip_cell_line"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf
response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write "          Fax:" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""hidden"" value=""" & sFax & """ name=""egov_users_userfax"" />" & vbcrlf
response.write "       		(<input class=""phonenum"" type=""text"" value=""" & Left(sFax,3)  & """ name=""skip_fax_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />)&nbsp;" & vbcrlf
response.write "       			<input class=""phonenum"" type=""text"" value=""" & Mid(sFax,4,3) & """ name=""skip_fax_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />&ndash;" & vbcrlf
response.write "       			<input class=""phonenum"" type=""text"" value=""" & Right(sFax,4) & """ name=""skip_fax_line"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf
response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & lcl_businessNameLabel & "</td>" & vbcrlf
response.write "      <td><input type=""text"" value=""" & sBusinessName & """ name=""egov_users_userbusinessname"" id=""egov_users_userbusinessname"" style=""width:300px;"" maxlength=""100"" /></td>" & vbcrlf
response.write "  </tr>" & vbcrlf

bHasBusinessStreets          = HasResidentTypeStreets( iOrgid, "B" )
lcl_hasBusinessStreets_label = "Business Street:"
bFound = False 

if bHasBusinessStreets then
	lcl_hasBusinessStreets_label = "Street (if not listed):"

	response.write "  <tr>" & vbcrlf
	response.write "      <td class=""label"" align=""right"">" & vbcrlf
	response.write "          Business Street:" & vbcrlf
	response.write "      </td>" & vbcrlf
	response.write "      <td>" & vbcrlf
	DisplayAddresses iorgid, "B", sBusinessAddress, bFound
	response.write "      </td>" & vbcrlf
	response.write "  </tr>" & vbcrlf
end if

response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & vbcrlf
response.write            lcl_hasBusinessStreets_label & vbcrlf
response.write "      </td>" & vbcrlf
response.write "      <td>" & vbcrlf

lcl_businessaddress = ""

if not bfound then
	lcl_businessaddress = sBusinessAddress
end if

response.write "          <input type=""text"" value=""" & lcl_businessaddress & """ name=""egov_users_userbusinessaddress"" style=""width:300px;"" maxlength=""100"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf
response.write "  <tr>" & vbcrlf
response.write "      <td class=""label"" align=""right"">" & lcl_workPhoneLabel & "</td>" & vbcrlf
response.write "      <td>" & vbcrlf
response.write "          <input type=""hidden"" value=""" & sWorkPhone & """ name=""egov_users_userworkphone"" id=""egov_users_userworkphone"" />" & vbcrlf
response.write "      			(<input class=""phonenum"" type=""text"" value=""" & Left(sWorkPhone,3) & """ name=""skip_work_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />)&nbsp;" & vbcrlf
response.write "       			<input class=""phonenum"" type=""text"" value=""" & Mid(sWorkPhone,4,3) & """ name=""skip_work_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />&ndash;" & vbcrlf
response.write "        		<input class=""phonenum"" type=""text"" value=""" & Mid(sWorkPhone,7,4) & """ name=""skip_work_line"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" />&nbsp;" & vbcrlf
response.write "       			ext. <input class=""phonenum"" type=""text"" value=""" & Mid(sWorkPhone,11,4) & """ name=""skip_work_ext"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf

if not lcl_orghasfeature_no_emergency_contact then
	response.write "  <tr>" & vbcrlf
	response.write "      <td class=""label"" align=""right"">" & vbcrlf
	response.write "         Emergency Contact:" & vbcrlf
	response.write "      </td>" & vbcrlf
	response.write "      <td>" & vbcrlf
	response.write "          <input type=""text"" value=""" & sEmergencyContact & """ name=""egov_users_emergencycontact"" style=""width:300px;"" maxlength=""100"" />" & vbcrlf
	response.write "      </td>" & vbcrlf
	response.write "  </tr>" & vbcrlf
	response.write "  <tr>" & vbcrlf
	response.write "      <td class=""label"" align=""right"">" & vbcrlf
	response.write "          Emergency Phone:" & vbcrlf
	response.write "      </td>" & vbcrlf
	response.write "      <td>" & vbcrlf
	response.write "          <input type=""hidden"" value=""" & sEmergencyPhone & """ name=""egov_users_emergencyphone"" />" & vbcrlf
	response.write "	     			(<input class=""phonenum"" type=""text"" value=""" & Left(sEmergencyPhone,3)  & """ name=""skip_emergencyphone_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />)&nbsp;" & vbcrlf
	response.write "			      	<input class=""phonenum"" type=""text"" value=""" & Mid(sEmergencyPhone,4,3) & """ name=""skip_emergencyphone_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />&ndash;" & vbcrlf
	response.write "	      			<input class=""phonenum"" type=""text"" value=""" & Mid(sEmergencyPhone,7,4) & """ name=""skip_emergencyphone_line"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" />" & vbcrlf
	response.write "      </td>" & vbcrlf
	response.write "  </tr>" & vbcrlf
end if

'Do Not Knock List Options
If lcl_orghasfeature_donotknock Then 

	'Determine if the user is a "do not knock" vendor
	lcl_userid            = request.cookies("userid")
	lcl_dnk_vendor_title  = ""
	lcl_canViewPeddlers   = checkAccessToList(lcl_userid, iorgid, "peddlers")
	lcl_canViewSolicitors = checkAccessToList(lcl_userid, iorgid, "solicitors")

	if lcl_canViewPeddlers Or lcl_canViewSolicitors then
		lcl_dnk_vendor_title = "Authorized to view ""Do Not Knock"" List"
	end if

	if lcl_canViewPeddlers then
		if lcl_canViewSolicitors then
			lcl_dnk_vendor_title = lcl_dnk_vendor_title & "s: Peddlers/Solicitors"
		else
			lcl_dnk_vendor_title = lcl_dnk_vendor_title & ": Peddlers"
		end if
	else
		if lcl_canViewSolicitors then
			lcl_dnk_vendor_title = lcl_dnk_vendor_title & ": Solicitors"
		end if
	end if

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
End If 

response.write "  <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
response.write "  <tr>" & vbcrlf
response.write "      <td colspan=""2"" align=""center"">" & vbcrlf
response.write "          <input class=""actionbtn"" type=""button"" value=""" & sSaveButtonText & """ onClick=""javascript:doCheck();"" />" & vbcrlf
response.write "      </td>" & vbcrlf
response.write "  </tr>" & vbcrlf
response.write "</table>" & vbcrlf
response.write "      </form>" & vbcrlf
response.write "    </div>" & vbcrlf
response.write "  </div>" & vbcrlf
response.write "</div>" & vbcrlf
response.write "<script>"
     response.write "jQuery(document).ready(function(){" & vbcrlf
        response.write "  jQuery('#validaddresslist').hide();" & vbcrlf
     response.write "});" & vbcrlf
response.write "</script>"

Set oRegOrg = Nothing 

response.write "<p>&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</p>" & vbcrlf
%>
   
<!--#Include file="include_bottom.asp"--> 

<!--#Include file="includes\inc_dbfunction.asp"-->  

<%
'------------------------------------------------------------------------------
Sub DisplayNeighborhoods( ByVal iorgid, ByVal iNeighborhoodId )
	Dim sSql, oRs 

	sSql = "SELECT neighborhoodid, neighborhood FROM egov_neighborhoods "
	sSql = sSql & "WHERE orgid = " & iorgid & " order by neighborhood"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select name=""egov_users_neighborhoodid"">"	
	response.write vbcrlf &  "<option value=""0"">Not on List...</option>"
		
	Do While Not oRs.EOF 
		response.write vbcrlf & "<option value=""" &  oRs("neighborhoodid") & """"
		If clng(iNeighborhoodId) = clng(oRs("neighborhoodid")) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("neighborhood") & "</option>"
		oRs.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub  

'------------------------------------------------------------------------------
Sub DisplayAddresses( ByVal iorgid, ByVal sResidenttype, ByVal sAddress, ByRef bFound )
	sSql = "SELECT residentstreetnumber, residentstreetname FROM egov_residentaddresses_list where orgid=" & iorgid & " and residenttype='" & sResidenttype & "' order by sortstreetname, residentstreetprefix, Cast(residentstreetnumber as int)"
	Set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSql, Application("DSN") , 3, 1

	response.write "<select name=""skip_" & sResidenttype & "address"">"	
	response.write "<option value=""0000"">Please select an address...</option>"
		
	Do While NOT oAddressList.EOF 
		response.write vbcrlf & "<option value=""" &  oAddressList("residentstreetnumber") & " " & oAddressList("residentstreetname")  & """"
		If UCase(sAddress) = UCase(oAddressList("residentstreetnumber") & " " & oAddressList("residentstreetname")) Then
			bFound = True
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oAddressList("residentstreetnumber") & " " & oAddressList("residentstreetname") & "</option>"
		oAddressList.MoveNext
	Loop

	response.write "</select>"

	oAddressList.close
	Set oAddressList = Nothing 
	
End Sub


'------------------------------------------------------------------------------
Sub DisplayLargeAddressList( ByVal p_orgid, ByVal sResidenttype, ByVal iOrgHasFeature_CitizenRegistration_NoValidate_Address, ByVal sStreetNumber, ByVal sStreetName, ByRef bFound )
	Dim sSql, oRs, sCompareName

	'Determine if we are to validate the address (with street number AND street name) or only the street name.
	'If this feature "CitizenRegistration_NoValidate_Address" is ENABLED then the org does NOT want to validate the address
	'   with the street number ONLY the street name.
	lcl_streetnumber = ""
	lcl_streetname   = ""
	bFound           = False

	If iOrgHasFeature_CitizenRegistration_NoValidate_Address Then 
		If IsValidAddress_byStreetName(p_orgid, sStreetName) Then 
			lcl_streetnumber = sStreetNumber
			sStreetName = sStreetname
			bFound = True
		End If 
	Else 
		If isValidAddress(sStreetNumber, sStreetName) Then 
			lcl_streetnumber = sStreetNumber
			sStreetName = sStreetName
			bfound = True
		End If 
	End If 

	'if  not IsValidAddress( sStreetNumber, sStreetName ) then
	'    sStreetNumber = ""
	'    sStreetName   = ""
	'    bFound        = False 
	'end if

	sSql = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, "
	sSql = sSql & "residentstreetname, ISNULL(streetsuffix,'') AS streetsuffix, "
	sSql = sSql & "ISNULL(streetdirection,'') AS streetdirection "
	sSql = sSql & "FROM egov_residentaddresses "
	sSql = sSql & "WHERE orgid = " & p_orgid
	sSql = sSql & " AND residenttype = '" & sResidenttype & "' AND residentstreetname IS NOT NULL "
	sSql = sSql & "ORDER BY sortstreetname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<input type=""text"" name=""residentstreetnumber"" id=""residentstreetnumber"" value=""" & lcl_streetnumber & """ size=""8"" maxlength=""10"" /> &nbsp; "
		response.write vbcrlf & "<select name=""skip_address"" id=""skip_address"">" 
		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown</option>"

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

			If UCase(sStreetName) = UCase(sCompareName) and bfound Then 
				lcl_address_selected = " selected=""selected"""
				'bFound = True
			End If 
			response.write vbcrlf & "<option value=""" & sCompareName & """ " & lcl_address_selected & ">" & sCompareName & "</option>"
			oRs.MoveNext
		Loop 

		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 

'------------------------------------------------------------------------------
' boolean HasResidentTypeStreets( iOrgid, sResidenttype )
'------------------------------------------------------------------------------
Function HasResidentTypeStreets( ByVal iOrgid, ByVal sResidenttype )
	Dim sSql, oRs 

	sSql = "SELECT count(residentaddressid) AS hits FROM egov_residentaddresses WHERE orgid = " & iorgid
	sSql = sSql & " AND residenttype = '" & sResidenttype & "'"

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
' void GetRegisteredUserValues()
'------------------------------------------------------------------------------
Sub GetRegisteredUserValues()
	Dim sSql, oRs

	If request.cookies("userid") <> "" And request.cookies("userid") <> "-1" Then
		
		sSql = "SELECT * FROM egov_users WHERE userid = " & request.cookies("userid")

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then
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
			iUserID                        = oRs("userid")
			sIsOnDoNotKnockList_peddlers   = oRs("isOnDoNotKnockList_peddlers")
			sIsOnDoNotKnockList_solicitors = oRs("isOnDoNotKnockList_solicitors")

			If IsNull("gender") Then 
				sGender = "N"
			Else
				sGender = oRs("gender")
			End If 

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
			sEmergencyPhone   = oRs("emergencyphone")
			sUserUnit         = oRs("userunit")
		End If

		oRs.Close
		Set oRs = Nothing 
	Else
		' REDIRECT TO USER LOGIN
		response.redirect("user_login.asp")
	End If 

End Sub 


'------------------------------------------------------------------------------
Sub UpdateUserNeighborhood( ByVal iUserId, ByVal iNeighborhoodid )
	Dim sSql, oCmd

	sSql = "UPDATE egov_users SET neighborhoodid = "
	If iNeighborhoodid = "0" Then 
		sSql = sSql & " NULL "
	Else
		sSql = sSql & iNeighborhoodid
	End If 
	sSql = sSql & " WHERE userid = " & iUserId 

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub 


'------------------------------------------------------------------------------
Sub UpdateResidencyVerified( ByVal iUserId )
	Dim sSql, oCmd

	sSql = "Update egov_users set residencyverified = 0 Where userid = " & iUserId 

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub 


'------------------------------------------------------------------------------
Sub UpdateExistingUser( ByVal iUserId )
	Dim sSql, oCmd, bValid, oRE, sUserfname, sUserlname, sUserstreetnumber, sUserunit, sUseraddress
	Dim sUsercity, sUserstate, sUserzip, sUseremail, sUserhomephone, sUsercell, sUserworkphone
	Dim sUserfax, sUserbusinessname, sUserbusinessaddress, sUserpassword, sResidenttype
	Dim iNeighborhoodid, sEmergencycontact, sEmergencyphone, sIsOnDoNotKnockList_peddlers
	Dim sIsOnDoNotKnockList_solicitors, sGender

	iUserId = CLng(iUserId)
	bValid = True 
	Set oRE = New RegExp
	oRE.IgnoreCase = False    
	oRE.Global = False 

	sUserfname = "'" & dbready_string(request("egov_users_userfname"),50) & "'"

	sUserlname = "'" & dbready_string(request("egov_users_userlname"),50) & "'"

	If request("egov_users_gender") <> "M" And request("egov_users_gender") <> "F" Then
		sGender = "NULL"
	Else
		sGender = "'" & dbready_string(request("egov_users_gender"),1) & "'"
	End If 

	If request("residentstreetnumber") <> "" Then
		sUserstreetnumber = "'" & dbready_string(request("residentstreetnumber"),10) & "'"
	Else
		sUserstreetnumber = "NULL"
	End If

	If request("egov_users_userunit") <> "" Then
		If clng(InStr(request("egov_users_userunit"),"http:")) > clng(0) Then 
			bValid = False 
			errormsg = errormsg & " The Unit contains invalid data.<br />"
		Else
			sUserunit = "'" & dbready_string(request("egov_users_userunit"),10) & "'"
		End If 
	Else
		sUserunit = "NULL"
	End If

	If request("egov_users_useraddress") <> "" Then 
		sUseraddress = "'" & dbready_string(request("egov_users_useraddress"),255) & "'"
	Else
		sUseraddress = "NULL"
	End If 

	If request("egov_users_usercity") <> "" Then
		sUsercity = "'" & dbready_string(request("egov_users_usercity"),50) & "'"
	Else
		sUsercity = "NULL"
	End If 

	If request("egov_users_userstate") <> "" Then
		sUserstate = UCase(dbready_string(request("egov_users_userstate"),2))
		' Spammers put a foreign country in as the state, so this will trip them up.
'		If StateNotValid( sUserstate ) Then
'			bValid = False 
'			errormsg = errormsg & " The State is invalid.<br />"  
'		Else
			sUserstate = "'" & sUserstate & "'"
'		End If 
	Else
		sUserstate = "NULL"
	End If 

	If request("egov_users_userzip") <> "" Then 
		sUserzip = "'" & dbready_string(request("egov_users_userzip"),10) & "'"
		' Check for numbers and a dash only - Spammers may put too many digits and no dash
'		If clng(InStr(sUserzip, "-")) = clng(0) Then
'	        oRE.Pattern = "^'\d{5}'$"
 '       Else
'		    oRE.Pattern = "^'\d{5}-\d{4}'$"
 '       End If
'		If Not oRE.Test(sUserzip) Then 
'			bValid = False 
'			errormsg = errormsg & " Zipcode is invalid.<br />" 
'		End If 
	Else
		sUserzip = "NULL"
	End If 

	If request("egov_users_useremail") <> "" Then 
		sUseremail = "'" & dbready_string(request("egov_users_useremail"),512) & "'"
	Else
		bValid = False    ' They have to have an email
		errormsg = errormsg & " An email address is required.<br />"
	End If 

	If request("egov_users_userhomephone") <> "" Then
		sUserhomephone = "'" & dbready_string(request("egov_users_userhomephone"),10) & "'"
		oRE.Pattern = "^'\d{10}'$"
		If Not oRE.Test(sUserhomephone) Then 
			bValid = False    
			errormsg = errormsg & " The phone number is invalid.<br />"
		End If 
	Else
		sUserhomephone = "NULL"
	End If 

	If request("egov_users_usercell") <> "" Then
		sUsercell = "'" & dbready_string(request("egov_users_usercell"),10) & "'"
		oRE.Pattern = "^'\d{10}'$"
		If Not oRE.Test(sUsercell) Then 
			bValid = False    
			errormsg = errormsg & " The cell number is invalid.<br />"
		End If 
	Else
		sUsercell = "NULL"
	End If 

	If request("egov_users_userworkphone") <> "" Then
		sUserworkphone = "'" & dbready_string(request("egov_users_userworkphone"),14) & "'"
		oRE.Pattern = "^'\d{10}\d{0,4}'$"
		If Not oRE.Test(sUserworkphone) Then
			bValid = False    
			errormsg = errormsg & " The work number is invalid.<br />"
		End If 
	Else
		sUserworkphone = "NULL"
	End If 

	If request("egov_users_userfax") <> "" Then
		sUserfax = "'" & dbready_string(request("egov_users_userfax"),10) & "'"
		oRE.Pattern = "^'\d{10}'$"
		If Not oRE.Test(sUserfax) Then
			bValid = False    
			errormsg = errormsg & " The fax number is invalid.<br />" 
		End If 
	Else
		sUserfax = "NULL"
	End If 

	If request("egov_users_userbusinessname") <> "" Then
		sUserbusinessname = "'" & dbready_string(request("egov_users_userbusinessname"),100) & "'"
	Else
		sUserbusinessname = "NULL"
	End If 

	If request("egov_users_userbusinessaddress") <> "" Then 
		sUserbusinessaddress = "'" & dbready_string(request("egov_users_userbusinessaddress"),255) & "'"
	Else
		sUserbusinessaddress = "NULL"
	End If 

	If request("egov_users_userpassword") <> "" Then
		If request("skip_userpassword2") <> request("egov_users_userpassword") Then 
			bValid = False
			errormsg = errormsg & " A password is required.<br />"
		Else 
			sUserpassword = "'" & createHashedPassword(request("egov_users_userpassword")) & "'"
		End If 
	Else
		' passwords are required
		bValid = False 
	End If

	If request("egov_users_residenttype") <> "" Then 
		sResidenttype = "'" & dbready_string(request("egov_users_residenttype"),1) & "'"


		if iOrgID = "60" then
			if request.form("residentstreetnumber") = "701" and request.form("skip_address") = "LAUREL ST" and lcase(request.form("egov_users_usercity")) = "menlo park" _
				and lcase(request.form("egov_users_userstate")) = "ca" and left(request.form("egov_users_userzip"),5) = "94025" then
				sResidenttype = "'E'"
			elseif sResidenttype = "'R'" then
			elseif request.form("egov_users_userbusinessname") <> "" and (request.form("skip_Baddress") <> "0000" OR request.form("egov_users_userbusinessaddress") <> "") then
				sResidenttype = "'B'"
			elseif left(request.form("egov_users_userzip"),5) = "94025" and lcase(request.form("egov_users_usercity")) = "menlo park" then
				sResidenttype = "'U'"
			else
				sResidenttype = "'N'"
			end if
		end if
	Else
		' This will always have a value
		bValid = False 
		errormsg = errormsg & " There is a problem with the address provided.<br />"
	End If

	If request("egov_users_neighborhoodid") <> "" Then
		If dbready_number( request("egov_users_neighborhoodid") ) Then
			iNeighborhoodid = CLng(request("egov_users_neighborhoodid"))
		Else
			bValid = False 
			errormsg = errormsg & " There is a problem with the selected neighborhood.<br />" 
		End If 
	Else
		iNeighborhoodid = "NULL"
	End If 

	If request("egov_users_emergencycontact") <> "" Then
		If clng(InStr(request("egov_users_emergencycontact"),"http:")) > clng(0) Then 
			bValid = False 
			errormsg = errormsg & " The emergecy contact contains invalid data.<br />" 
		Else 
			sEmergencycontact = "'" & dbready_string(request("egov_users_emergencycontact"),100) & "'"
		End If 
	Else
		sEmergencycontact = "NULL"
	End If

	If request("egov_users_emergencyphone") <> "" Then
		sEmergencyphone = "'" & dbready_string(request("egov_users_emergencyphone"),10) & "'"
		oRE.Pattern = "^'\d{10}'$"
		If Not oRE.Test(sEmergencyphone) Then 
			bValid = False 
			errormsg = errormsg & " The emergecy phone is invalid.<br />"
		End If 
	Else
		sEmergencyphone = "NULL"
	End If 

  sIsOnDoNotKnockList_peddlers   = "0"
  sIsOnDoNotKnockList_solicitors = "0"
  'sIsDoNotKnockVendor_peddlers   = "0"
  'sIsDoNotKnockVendor_solicitors = "0"

  if request("isOnDoNotKnockList_peddlers") = "on" then
     sIsOnDoNotKnockList_peddlers = "1"
  end if

  if request("isOnDoNotKnockList_solicitors") = "on" then
     sIsOnDoNotKnockList_solicitors = "1"
  end if

	set oRE = Nothing 

	If bValid Then 
		sSql = "UPDATE egov_users SET "
		sSql = sSql & " userfname = "                     & sUserfname                   & ", "
		sSql = sSql & " userlname = "                     & sUserlname                   & ", "
		sSql = sSql & " userstreetnumber = "              & sUserstreetnumber            & ", "
		sSql = sSql & " userunit = "                      & sUserunit                    & ", "
		sSql = sSql & " useraddress = "                   & sUseraddress                 & ", "
		sSql = sSql & " usercity = "                      & sUsercity                    & ", "
		sSql = sSql & " userstate = "                     & sUserstate                   & ", "
		sSql = sSql & " userzip = "                       & sUserzip                     & ", "
		sSql = sSql & " useremail = "                     & sUseremail                   & ", "
		sSql = sSql & " userhomephone = "                 & sUserhomephone               & ", "
		sSql = sSql & " usercell = "                      & sUsercell                    & ", "
		sSql = sSql & " userworkphone = "                 & sUserworkphone               & ", "
		sSql = sSql & " userfax = "                       & sUserfax                     & ", "
		sSql = sSql & " userbusinessname = "              & sUserbusinessname            & ", "
		sSql = sSql & " userbusinessaddress = "           & sUserbusinessaddress         & ", "
		sSql = sSql & " userpassword = NULL,"
		sSql = sSql & " password = "	                  & sUserpassword                & ", "
		sSql = sSql & " residenttype = "                  & sResidenttype                & ", "
		sSql = sSql & " neighborhoodid = "                & iNeighborhoodid              & ", "
		sSql = sSql & " emergencycontact = "              & sEmergencycontact            & ", "
		sSql = sSql & " emergencyphone = "                & sEmergencyphone              & ", "
		sSql = sSql & " isOnDoNotKnockList_peddlers = "   & sIsOnDoNotKnockList_peddlers & ", "
		sSql = sSql & " isOnDoNotKnockList_solicitors = " & sIsOnDoNotKnockList_solicitors & ", "
		sSql = sSql & " gender = " & sGender
		sSql = sSql & " WHERE userid = " & iUserId 

if request.cookies("userid") = "1150705" then
		'response.write sSql & "<br />"
		'response.end
end if

		RunSQLStatement sSql		' In common.asp
		
	Else
		errormsg = "<strong>There was a problem processing your changes.</strong><br />" & errormsg 
	End If 

End Sub

'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Your subscription selection(s) have been saved successfully."
     elseif iSuccess = "SA" then
        lcl_return = "Your subscription selection(s) have been submitted successfully."
     elseif iSuccess = "SU_NA" then
        lcl_return = "Your subscription selection(s) have been saved successfully.<br />An address is recommended to be entered."
     elseif iSuccess = "SA_NA" then
        lcl_return = "Your subscription selection(s) have been submitted successfully.<br />An address is recommended to be entered."
     elseif iSuccess = "EXISTS_BAD_PWD" then
        lcl_return = "An account already exists with this email, but the password does not match."
     elseif iSuccess = "REQUIRED_PWD" then
        lcl_return = "A password is required."
     elseif iSuccess = "BAD_PWD" then
        lcl_return = "The password entered is incorrect."
     elseif iSuccess = "NOT_EXISTS" then
        lcl_return = "We do not have an account with this email address."
     elseif iSuccess = "UNSUBSCRIBED" then
        lcl_return = "ALL subscriptions have been removed."
     elseif iSuccess = "AJAX_ERROR" then
        lcl_return = "ERROR: An error has during the AJAX routine..."
     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
'Sub DisplayLargeAddressListOld( ByVal sResidenttype, ByVal sStreetNumber, ByVal sStreetName, ByRef bFound )
'	Dim sSql, oAddressList

'	If Not IsValidAddress( sStreetNumber, sStreetName ) Then   ' In common.asp
'		sStreetNumber = ""
'		sStreetName = ""
'		bFound = False 
'	End If 

'	sSql = "SELECT distinct sortstreetname, residentstreetprefix, residentstreetname "
'	sSql = sSql & " FROM egov_residentaddresses where orgid = " & iOrgid & " and residenttype = '" & sResidenttype & "' "
'	sSql = sSql & "and residentstreetname is not null order by sortstreetname, residentstreetprefix, residentstreetname"
	
'	Set oAddressList = Server.CreateObject("ADODB.Recordset")
'	oAddressList.Open sSql, Application("DSN"), 3, 1

'	If Not oAddressList.EOF Then
'		response.write vbcrlf & "<input type=""text"" name=""residentstreetnumber"" value=""" & sStreetNumber & """ size=""8"" maxlength=""10"" /> &nbsp; "
'		response.write vbcrlf & "<select name=""skip_address"">"
'		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown</option>"
'		Do While NOT oAddressList.EOF 
'			response.write vbcrlf & "<option value="""  & oAddressList("residentstreetname") & """"
'			If sStreetName = oAddressList("residentstreetname") Then
'				response.write " selected=""selected"" "
'				bFound = True 
'			End If 
'			response.write " >"
'			response.write oAddressList("residentstreetname") & "</option>"
'			oAddressList.MoveNext
'		Loop 
'		response.write vbcrlf & "</select>"
'	End If 

'	oAddressList.Close
'	Set oAddressList = Nothing 

'End Sub 
%>
