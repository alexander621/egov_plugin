<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitapplicantedit.asp
' AUTHOR: Steve Loar
' CREATED: 02/20/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Creates and edits permit applicants
'
' MODIFICATION HISTORY
' 1.0   02/20/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iUserId, sTitle, sFirstName, sLastName, sAddress, sCity, sState, sZip, sPhone, sEmail, sFax, sCell
Dim sBusinessName, sDayPhone, sPassword, sEmailnotavailable, bHasResidentTypes, sResidencyVerified
Dim bHasResidentStreets, sStreetNumber, sStreetName, bFound, iNeighborhoodid, bHasBusinessStreets
Dim sBusinessAddress, sResidentType, sEmergencyContact, sEmergencyPhone, sUserUnit, iPermitStatusId
Dim sIsArchitect, sIsContractor, sIsOwner, sSaveButton, iFamilyId, bUpdateParent, sDetailId, iPermitId
Dim bCanSaveChanges, sHasFamilyMembers, sWorkPhone, sNotifyType, iPermitContactId

sLevel = "../" ' Override of value from common.asp

If request("userid") <> "" Then 
iUserId = CLng(request("userid") )
Else 
	iPermitContactId = CLng(request("permitcontactid")) 
	iUserId = GetUserIdByPermitContactId( iPermitContactId )
End If 


If request("permitid") <> "" Then 
	iPermitId = CLng(request("permitid"))
	iPermitStatusId = CLng(request("permitstatusid"))
	bCanSaveChanges = StatusAllowsSaveChanges( iPermitStatusId ) 	' in permitcommonfunctions.asp
Else
	iPermitId = CLng(0)
	iPermitStatusId = CLng(0)
	bCanSaveChanges = True 
End If 

If request("updatetitle") <> "" Then
	bUpdateParent = True 
Else 
	bUpdateParent = False 
End If 

sDetailId = request("detailid")

If sDetailId = "applicantdetails" Then
	sTitle = "Permit Applicant"
Else
	sTitle = "Primary Contact"
End If 

If CLng(iUserId) > CLng(0) Then
	sTitle = "Edit " & sTitle
	sSaveButton = "Update"
	If CLng(iPermitId) > CLng(0) Then
		' They are editing either the Applicant or the Primary Contact, pull from egov_permitcontacts
		GetPermitContactValues iUserId, iPermitId
		iFamilyId = GetFamilyId( iUserId )
	Else 
		' They are on the new permit screen, so pull from the egov_users table
		GetRegisteredUserValues iUserId
	End If 
	If UserHasFamilyMembers( iUserId, iFamilyId ) Then
		sHasFamilyMembers = "true"
	Else
		sHasFamilyMembers = "false"
	End If 
	sNotifyType = "Change"
Else
	sTitle = "New " & sTitle
	sSaveButton = "Create"
	sNotifyType = "Confirmation"
	sFirstName = ""
	sLastName = ""
	sAddress = ""
	sCity = ""
	sState = ""
	sZip = ""
	sPhone = ""
	sEmail = ""
	sDayPhone = ""
	sFax = ""
	sCell = ""
	sWorkPhone = ""
	sPassword = ""
	sEmailnotavailable = 1 
	sStreetNumber = ""
	sStreetName = ""
	bFound = False 
	bHasResidentTypes = False 
	bHasBusinessStreets = False 
	iNeighborhoodid = 0
	sBusinessAddress = ""
	sResidentType = "R"
	sResidencyVerified = 1 
	sEmergencyContact = ""
	sEmergencyPhone = ""
	sUserUnit = ""
	sIsArchitect = ""
	sIsContractor = "" 
	sIsOwner = ""
	iFamilyId = iUserId
	sHasFamilyMembers = "false"
End If 
%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
	<script language="JavaScript" src="../scripts/removespaces.js"></script>
	<script language="JavaScript" src="../scripts/setfocus.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

	<script language="Javascript">
	<!--
		parent.document.getElementById("modal"+window.frameElement.getAttribute("data-close")).style.width = "55%";
		parent.document.getElementById("modal"+window.frameElement.getAttribute("data-close")).style.height = "90%";
		parent.document.getElementById("modal"+window.frameElement.getAttribute("data-close")).style.left = "25%";
		parent.document.getElementById("modal"+window.frameElement.getAttribute("data-close")).style.top = "5%";

		var isNN = (navigator.appName.indexOf("Netscape")!=-1);
		var winHandle;
		var w = (screen.width - 640)/2;
		var h = (screen.height - 450)/2;

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
				document.frmPermitApplicant.familyaddresschanged.value = "YES";
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

		function FlagFamilyChange()
		{
			document.frmPermitApplicant.familyaddresschanged.value = "YES";
			//alert("yes");
		}

		function checkAddress( sReturnFunction, sSave )
		{
			// Remove any extra spaces
			document.frmPermitApplicant.residentstreetnumber.value = removeSpaces(document.frmPermitApplicant.residentstreetnumber.value);
			// check the number for non-numeric values
			var rege = /^\d+$/;
			var Ok = rege.exec(document.frmPermitApplicant.residentstreetnumber.value);
			if ( ! Ok )
			{
				alert("The Resident Street Number cannot be blank and must be numeric.");
				setfocus(document.frmPermitApplicant.residentstreetnumber);
				return false;
			}
			// check that they picked a street name
			if ( document.frmPermitApplicant.address.value == '0000')
			{
				alert("Please select a street name from the list first.");
				setfocus(document.frmPermitApplicant.address);
				return false;
			}
			// This is here because window.open in the Ajax callback routine will not work
			//winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&parentform=frmPermitApplicant' + '&stnumber=' + document.frmPermitApplicant.residentstreetnumber.value + '&stname=' + document.frmPermitApplicant.address.value + '&sCheckType=' + sReturnFunction + 'Validate", "_address", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			//self.focus();
			// Fire off Ajax routine
			doAjax('checkaddress.asp', 'stnumber=' + document.frmPermitApplicant.residentstreetnumber.value + '&stname=' + document.frmPermitApplicant.address.value, sReturnFunction, 'get', '0');
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
				document.frmPermitApplicant.useraddress.value = '';
				alert("This is a valid address in the system.");
			}
			else
			{
				//winHandle.focus();
				PopAStreetPicker('CheckResults', 'no');
			}
		}

		function PopAStreetPicker( sReturnFunction, sSave )
		{
			// pop up the address picker
			//winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&parentform=frmPermitApplicant' + '&stnumber=' + document.frmPermitApplicant.residentstreetnumber.value + '&stname=' + document.frmPermitApplicant.address.value + '&sCheckType=' + sReturnFunction + 'Validate", "_address", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('addresspicker.asp?saving=' + sSave + '&parentform=frmPermitApplicant' + '&stnumber=' + document.frmPermitApplicant.residentstreetnumber.value + '&stname=' + document.frmPermitApplicant.address.value + '&sCheckType=' + sReturnFunction + 'Validate', 'Invalid Address Selection', 90, 90);
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
				validate();
			}
			else
			{
				//winHandle.focus();
				PopAStreetPicker('FinalCheck', 'yes');
			}
			
		}

		function CloseThis()
		{
			//window.close();
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		function doCheck()
		{
			// If they are using the large address feature
			var exists = eval(document.frmPermitApplicant["residentstreetnumber"]);
			if(exists)
			{
				// If a street number was entered
				if (document.frmPermitApplicant.residentstreetnumber.value != '')
				{
					checkAddress( 'FinalCheck', 'yes' );
				}
				else
				{
					validate();
				}
			}
			else
			{
				validate();
			}
		}

		function finalCheckValidate()
		{
			// Fire off Ajax routine - This just breaks the tie to the address window so it will close for validation to continue
			doAjax('checkduplicatecitizens.asp', 'userlname=' + document.frmPermitApplicant.userlname.value, 'OkToValidate', 'get', '0');
		}

		function OkToValidate( sReturn )
		{
			//finish the validation routine 
			validate();
		}

		function validate()
		{
			var msg = "";

			if ( document.frmPermitApplicant.emailnotavailable.checked == false )
			{
				// They will login so validate the email and password
				if (document.frmPermitApplicant.useremail.value == "" )
				{
					msg+="The email cannot be blank.\n";
				}
				//var rege = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*\.(\w{2}|(com|net|org|edu|mil|gov|biz|us))$/;
				var rege = /.+@.+\..+/i;
				var Ok = rege.test(document.frmPermitApplicant.useremail.value);

				if (! Ok)
				{
					msg+="The email must be in a valid format.\n";
				}
				if (document.frmPermitApplicant.userpassword.value == "" )
				{
					msg+="The password cannot be blank.\n";
				}
				if (document.frmPermitApplicant.userpassword2.value == "" )
				{
					msg+="The verify password cannot be blank.\n";
				}
				if(document.frmPermitApplicant.userpassword.value != document.frmPermitApplicant.userpassword2.value)
				{
					msg+="The passwords you have entered do not match.\n";
				}
			}
			if (document.frmPermitApplicant.userfname.value == "" )
			{
				msg+="The first name cannot be blank.\n";
			}
			if (document.frmPermitApplicant.userlname.value == "" )
			{
				msg+="The last name cannot be blank.\n";
			}

			// set the emergency phone
			if (document.frmPermitApplicant.emergencyphone_areacode.value != "" || document.frmPermitApplicant.emergencyphone_exchange.value != "" || document.frmPermitApplicant.emergencyphone_line.value != "" )
			{
				var ePhone = document.frmPermitApplicant.emergencyphone_areacode.value + document.frmPermitApplicant.emergencyphone_exchange.value + document.frmPermitApplicant.emergencyphone_line.value;
				if (ePhone.length < 10)
				{
					msg += "The Emergency Phone must be a valid phone number, including area code, or blank\n";
				}
				else
				{
					document.frmPermitApplicant.emergencyphone.value = document.frmPermitApplicant.emergencyphone_areacode.value + document.frmPermitApplicant.emergencyphone_exchange.value + document.frmPermitApplicant.emergencyphone_line.value;
					var rege = /^\d+$/;
					var Ok = rege.exec(document.frmPermitApplicant.emergencyphone.value);
					if ( ! Ok )
					{
						msg += "The Emergency Phone must be a valid phone number, including area code, or blank\n";
					}
				}
			}

			// check the work phone
			if (document.frmPermitApplicant.work_areacode.value != "" || document.frmPermitApplicant.work_exchange.value != "" || document.frmPermitApplicant.work_line.value != "" || document.frmPermitApplicant.work_ext.value != "")
			{
				var sPhone = document.frmPermitApplicant.work_areacode.value + document.frmPermitApplicant.work_exchange.value + document.frmPermitApplicant.work_line.value;
				if (sPhone.length < 10)
				{
					msg += "Work Phone Number must be a valid phone number, including area code, or blank\n";
				}
				else
				{
					document.frmPermitApplicant.userworkphone.value = document.frmPermitApplicant.work_areacode.value + document.frmPermitApplicant.work_exchange.value + document.frmPermitApplicant.work_line.value + document.frmPermitApplicant.work_ext.value;
					var rege = /^\d+$/;
					var Ok = rege.exec(document.frmPermitApplicant.userworkphone.value);
					if ( ! Ok )
					{
						msg += "Work Phone Number must be a valid phone number, including area code, or blank\n";
					}
				}
			}

			// check the fax
			if (document.frmPermitApplicant.fax_areacode.value != "" || document.frmPermitApplicant.fax_exchange.value != "" || document.frmPermitApplicant.fax_line.value != "" )
			{
				var fPhone = document.frmPermitApplicant.fax_areacode.value + document.frmPermitApplicant.fax_exchange.value + document.frmPermitApplicant.fax_line.value;
				if (fPhone.length < 10)
				{
					msg += "Fax must be a valid phone number, including area code, or blank\n";
				}
				else
				{
					document.frmPermitApplicant.userfax.value = document.frmPermitApplicant.fax_areacode.value + document.frmPermitApplicant.fax_exchange.value + document.frmPermitApplicant.fax_line.value;
					var rege = /^\d+$/;
					var Ok = rege.exec(document.frmPermitApplicant.userfax.value);
					if ( ! Ok )
					{
						msg += "Fax must be a valid phone number, including area code, or blank\n";
					}
				}
			}

			// check the cell
			if (document.frmPermitApplicant.cell_areacode.value != "" || document.frmPermitApplicant.cell_exchange.value != "" || document.frmPermitApplicant.cell_line.value != "" )
			{
				var cPhone = document.frmPermitApplicant.cell_areacode.value + document.frmPermitApplicant.cell_exchange.value + document.frmPermitApplicant.cell_line.value;
				if (cPhone.length < 10)
				{
					msg += "The cell phone must be a valid phone number, including area code, or blank\n";
				}
				else
				{
					document.frmPermitApplicant.usercell.value = document.frmPermitApplicant.cell_areacode.value + document.frmPermitApplicant.cell_exchange.value + document.frmPermitApplicant.cell_line.value;
					var crege = /^\d+$/;
					var cOk = crege.exec(document.frmPermitApplicant.usercell.value);
					if ( ! cOk )
					{
						msg += "The cell phone must be a valid phone number, including area code, or blank\n";
					}
				}
			}

			// check the home phone number
			document.frmPermitApplicant.userhomephone.value = document.frmPermitApplicant.user_areacode.value + document.frmPermitApplicant.user_exchange.value + document.frmPermitApplicant.user_line.value;
			if (document.frmPermitApplicant.userhomephone.value != "" )
			{
				var hPhone = document.frmPermitApplicant.userhomephone.value;
				if (hPhone.length < 10)
				{
					msg += "The home phone must be a valid phone number, including area code.\n";
				}
				else
				{
					var rege = /^\d+$/;
					var Ok = rege.exec(document.frmPermitApplicant.userhomephone.value);
					if ( ! Ok )
					{
						msg += "The home phone must be a valid phone number, including area code.\n";
					}
				}
			}

			// Process the business address if one was chosen
			var bexists = eval(document.frmPermitApplicant["Baddress"]);
			if(bexists)
			{
				//See if they picked from the business dropdown and put that in the address field 
				if (document.frmPermitApplicant.Baddress.selectedIndex > -1)
				{
					var belement = document.frmPermitApplicant.Baddress;
					var bselectedvalue = belement.options[belement.selectedIndex].value;

					//alert( bselectedvalue );
					//  0000 is the first pick that we do not want
					if (bselectedvalue != "0000")
					{
						document.frmPermitApplicant.userbusinessaddress.value = bselectedvalue;
					}
				}
			}

			// Process the resident address if one was chosen - this is second to set the local resident type
			var exists = eval(document.frmPermitApplicant["Raddress"]);
			if(exists)
			{
				// See if they picked from the resident dropdown and put that in the address field 
				if (document.frmPermitApplicant.Raddress.selectedIndex > -1)
				{
					var element = document.frmPermitApplicant.Raddress;
					var selectedvalue = element.options[element.selectedIndex].value;

					//alert( selectedvalue );
					//  0000 is the first pick that we do not want
					if (selectedvalue != "0000")
					{
						document.frmPermitApplicant.useraddress.value = selectedvalue;
					}
				}
			}

			// handle the large quantity street addresses
			exists = eval(document.frmPermitApplicant["residentstreetnumber"]);
			if(exists)
			{
				if ( document.frmPermitApplicant.residentstreetnumber.value != '' )
				{
					// See if they picked from the resident dropdown and put that in the address field 
					if (document.frmPermitApplicant.address.selectedIndex > -1)
					{
						var element = document.frmPermitApplicant.address;
						var selectedvalue = element.options[element.selectedIndex].value;

						//alert( selectedvalue );
						//  0000 is the first pick that we do not want
						if (selectedvalue != "0000")
						{
							document.frmPermitApplicant.useraddress.value = document.frmPermitApplicant.residentstreetnumber.value + ' ' + selectedvalue;
							bUsedAddressDropdown = true;
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
				// Set some final things and then submit
				if (document.frmPermitApplicant.familyaddresschanged.value == "YES")
				{
					if ((document.frmPermitApplicant.userid.value != 0) && (document.frmPermitApplicant.hasfamilymembers.value == 'true'))
					{
						var bCopyToAll = confirm("Copy changes to all family members?");
						if ( ! bCopyToAll )
						{
							document.frmPermitApplicant.familyaddresschanged.value = "NO";
						}
					}
					else
					{
						document.frmPermitApplicant.familyaddresschanged.value = "NO";
					}
				}
				//alert("OK");

				// Build the post parameter 
				var sParameter = 'userid=' + encodeURIComponent(document.frmPermitApplicant.userid.value);
				sParameter += '&userfname=' + encodeURIComponent(document.frmPermitApplicant.userfname.value);
				sParameter += '&userlname=' + encodeURIComponent(document.frmPermitApplicant.userlname.value);
				sParameter += '&emailnotavailable=' + encodeURIComponent(document.frmPermitApplicant.emailnotavailable.checked);
				sParameter += '&useremail=' + encodeURIComponent(document.frmPermitApplicant.useremail.value);
				sParameter += '&userpassword=' + encodeURIComponent(document.frmPermitApplicant.userpassword.value);
				sParameter += '&residenttype=' + encodeURIComponent(document.frmPermitApplicant.residenttype.value);
				sParameter += '&residencyverified=' + encodeURIComponent(document.frmPermitApplicant.residencyverified.value);
				sParameter += '&usercity=' + encodeURIComponent(document.frmPermitApplicant.usercity.value);
				sParameter += '&userstate=' + encodeURIComponent(document.frmPermitApplicant.userstate.value);
				sParameter += '&userzip=' + encodeURIComponent(document.frmPermitApplicant.userzip.value);
				sParameter += '&userhomephone=' + encodeURIComponent(document.frmPermitApplicant.userhomephone.value);
				sParameter += '&usercell=' + encodeURIComponent(document.frmPermitApplicant.usercell.value);
				sParameter += '&userworkphone=' + encodeURIComponent(document.frmPermitApplicant.userworkphone.value);
				sParameter += '&userfax=' + encodeURIComponent(document.frmPermitApplicant.userfax.value);
				sParameter += '&emergencyphone=' + encodeURIComponent(document.frmPermitApplicant.emergencyphone.value);
				sParameter += '&neighborhoodid=' + encodeURIComponent(document.frmPermitApplicant.egov_users_neighborhoodid.value);
				sParameter += '&userunit=' + encodeURIComponent(document.frmPermitApplicant.userunit.value);
				sParameter += '&userbusinessname=' + encodeURIComponent(document.frmPermitApplicant.userbusinessname.value);
				sParameter += '&emergencycontact=' + encodeURIComponent(document.frmPermitApplicant.emergencycontact.value);
				sParameter += '&useraddress=' + encodeURIComponent(document.frmPermitApplicant.useraddress.value);
				sParameter += '&userbusinessaddress=' + encodeURIComponent(document.frmPermitApplicant.userbusinessaddress.value);
				sParameter += '&familyaddresschanged=' + encodeURIComponent(document.frmPermitApplicant.familyaddresschanged.value);
				sParameter += '&familyid=' + encodeURIComponent(document.frmPermitApplicant.familyid.value);
				sParameter += '&permitid=' + encodeURIComponent(document.frmPermitApplicant.permitid.value);
				sParameter += '&permitstatusid=' + encodeURIComponent(document.frmPermitApplicant.permitstatusid.value);
				sParameter += '&notifyuser=' + encodeURIComponent(document.frmPermitApplicant.notifyuser.checked);
				
				//if (document.frmPermitApplicant.userid.value == 0)
				//{
				//	sParameter += '&isarchitect=' + encodeURIComponent(document.frmPermitApplicant.isarchitect.checked);
				//	sParameter += '&iscontractor=' + encodeURIComponent(document.frmPermitApplicant.iscontractor.checked);
				//	sParameter += '&isowner=' + encodeURIComponent(document.frmPermitApplicant.isowner.checked);
				//}
				//alert( sParameter );
				// Save Call Here
				<% If bUpdateParent Then %>
					
					// Fire off and wait for a response
					doAjax('permitapplicantsave.asp', sParameter, 'CloseThisSaved', 'post', '0');
				<% Else %>
					// Fire off and close
					doAjax('permitapplicantsave.asp', sParameter, 'CloseMe', 'post', '0');
					
				<% End If %>
				//document.frmPermitApplicant.submit();
			}
		}

		function CloseMe( sResult )
		{
			//window.close();
			//SHOULD UPDATE PARENT DATA

			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		function CloseThisSaved( sResult )
		{
			var sDetailText;
			var sPrefix; 
			var sToolTip = '';
			//alert( sResult );
			// optionally put code here to update the parent window
			<% If bUpdateParent Then %>
				if (parseInt(document.frmPermitApplicant.userid.value) > 0)
				{
					sDetailText = '<a href="javascript:';
					<% If sDetailId = "applicantdetails" Then %>
						sDetailText += 'EditApplicant';
						sPrefix = '\'U\', ';
					<% Else %>
						sDetailText += 'EditPrimaryContact';
						sPrefix = '';
					<% End If %>
					sDetailText += '(' + sPrefix + '\'' + document.frmPermitApplicant.userid.value + '\' );" ';
					if (document.frmPermitApplicant.userfname.value != "")
					{
						sToolTip += '<strong>' + document.frmPermitApplicant.userfname.value + ' ' + document.frmPermitApplicant.userlname.value + '</strong><br />';
					}
					if (document.frmPermitApplicant.userbusinessname.value != "")
					{
						sToolTip += document.frmPermitApplicant.userbusinessname.value + '<br />';
					}
					
					if (document.frmPermitApplicant.useraddress.value != "")
					{
						sToolTip += document.frmPermitApplicant.useraddress.value + '<br />';
					}
					if (document.frmPermitApplicant.usercity.value != "")
					{
						sToolTip += document.frmPermitApplicant.usercity.value + ', ' + document.frmPermitApplicant.userstate.value + ' ' + document.frmPermitApplicant.userzip.value + '<br />';
					}
					if (document.frmPermitApplicant.user_areacode.value != "")
					{
						sToolTip += '(' + document.frmPermitApplicant.user_areacode.value + ') ' + document.frmPermitApplicant.user_exchange.value + '-' + document.frmPermitApplicant.user_line.value;
					}
					var myRegExp = /\'/g;
					sToolTip = sToolTip.replace(myRegExp, '\\&#39;');
					sDetailText += ' onMouseover="ddrivetip(\'' + sToolTip + '\', 300)"; onMouseout="hideddrivetip()"; '
					sDetailText += '>';
					sDetailText += document.frmPermitApplicant.userfname.value + ' ' + document.frmPermitApplicant.userlname.value;
					if (document.frmPermitApplicant.userbusinessname.value != "")
					{
						sDetailText += ' (' + document.frmPermitApplicant.userbusinessname.value + ')';
					}

					sDetailText += '</a>';
					//alert( sDetailText );
					//alert('<%=sDetailId%>');
					parent.document.getElementById('<%=sDetailId%>').innerHTML = sDetailText;
					//window.opener.NiceTitles.autoCreated.anchors.addElements(window.opener.document.getElementsByTagName("a"), "title");
				}
				else
				{
					// Put stuff to select the contact here
					//parent.document.getElementById("primarycontactdetails").innerHTML = document.frmPermitApplicant.userfname.value + ' ' + document.frmPermitApplicant.userlname.value;
					//parent.document.getElementById("isprimarycontactuserid").value = sResult;
					parent.document.getElementById("applicant").innerHTML = '<select name="userid" id="userid" onchange="toggleAddressSearch();"><option value="U' + sResult + '">' + document.frmPermitApplicant.userfname.value + ' ' + document.frmPermitApplicant.userlname.value + '</option></select>';
				}
			<% End If %>

			if (typeof parent.hideddrivetip !== "undefined") { 
				parent.hideddrivetip();
			}
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		function GetRndPwd()
		{
			var pwd;

			pwd = Math.round(Math.random() * 1000000 );
			//alert(pwd);
			$("userpassword").value = pwd;
			$("userpassword2").value = pwd;
		}

	//-->
	</script>

</head>

<body>
			
	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<script>parent.document.getElementById('modaltitle'+window.frameElement.getAttribute("data-close")).innerHTML = '<%=sTitle%>';</script>
			<form name="frmPermitApplicant" method="post" action="permitapplicantsave.asp">
				<input type="hidden" name="userid" value="<%=iUserId%>" />
				<input type="hidden" name="familyid" value="<%=iFamilyId%>" />
				<input type="hidden" name="familyaddresschanged" value="NO" />
				<input type="hidden" name="residencyverified" value="<%=sResidencyVerified%>" />
				<input type="hidden" name="permitid" value="<%=iPermitId%>" />
				<input type="hidden" name="permitstatusid" value="<%=iPermitStatusId%>" />
				<input type="hidden" name="hasfamilymembers" value="<%=sHasFamilyMembers%>" />

				<table border="0" id="permitapplicant" cellpadding="4" cellspacing="0">
					<tr>
						<td class="label" align="right" nowrap="nowrap">
							<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color="red">*</font></span> 
							First Name:</span>
						</td><td>
							<span class="cot-text-emphasized" title="This field is required"> 
							<input type="text" value="<%=sFirstName%>" name="userfname" size="50" maxlength="50" />
							</span>
						</td>
					</tr>
					<tr>
						<td class="label" align="right" nowrap="nowrap">
							<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color="red">*</font></span> 
							Last Name:</span>
						</td><td>
							<span class="cot-text-emphasized" title="This field is required"> 
							<input type="text" value="<%=sLastName%>" name="userlname" size="50" maxlength="50" />
							</span>
						</td>
					</tr>
					<tr><td class="label" align="right" nowrap="nowrap" valign="top">&nbsp;</td>
						<td><input type="checkbox" name="emailnotavailable"
							<% If CLng(sEmailnotavailable) = CLng(1) Then
									response.write " checked=""checked"" "
								End If %>
							/> Email Not Available (They Will Not Be Able To Login)
						</td>
					</tr>
					<tr><td class="label" align="right" valign="top" nowrap="nowrap">Email:</td>
						<td>
							<input type="text" name="useremail" value="<%=sEmail%>" size="50" maxlength="100" />
							<br /><input type="checkbox" checked="checked" name="notifyuser" /> Send Registration <%=sNotifyType%> to Citizen
						</td>
					</tr>
					<tr><td class="label" align="right" nowrap="nowrap">
							<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color="red"></font></span> 
							&nbsp;</span>
						</td>
						<td>
							<input type="button" class="button ui-button ui-widget ui-corner-all" value="Generate Random Password" name="generate_pwd" onclick="GetRndPwd();" />
						</td>
					</tr>
					<tr>
						<td class="label" align="right" nowrap="nowrap">Password:</td>
						<td><input type="password" value="<%=sPassword%>" id="userpassword" name="userpassword" size="50" maxlength="100" /></td>
					</tr>
					<tr>
						<td class="label" align="right" nowrap="nowrap">Verify Password:</td>
						<td><input type="password" value="<%=sPassword%>" id="userpassword2" name="userpassword2" size="50" maxlength="100" /></td>
					</tr>
<%					bHasResidentTypes = HasResidentTypes()
					bFound = False 
					If bHasResidentTypes Then %>
					<tr>
						<td class="label" align="right" nowrap="nowrap">
							<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color="red">*</font></span>
							Resident Type:</span></td>
						<td><%=DisplayResidentTypes( sResidentType ) %>
						</td>
					</tr>
<%					End If %>
					<tr>
						<td class="label" align="right" nowrap="nowrap">Home Phone:
<%					If Not bHasResidentTypes Then	%>
						<input type="hidden" name="residenttype" value="R" />
<%					End If	%>
						</td>
						<td>
							<input type="hidden" value="<%=sDayPhone%>" name="userhomephone" />
							(<input type="text" value="<%=Left(sDayPhone,3)%>" name="user_areacode" onKeyUp="return autoTab(this, 3, event, true);" onchange="FlagFamilyChange();" size="3" maxlength="3" />)&nbsp;
							<input type="text" value="<%=Mid(sDayPhone,4,3)%>" name="user_exchange" onKeyUp="return autoTab(this, 3, event, true);" onchange="FlagFamilyChange();" size="3" maxlength="3" />&ndash;
							<input type="text" value="<%=Right(sDayPhone,4)%>" name="user_line" onKeyUp="return autoTab(this, 4, event, true);" onchange="FlagFamilyChange();" size="4" maxlength="4" />
						</td>
					</tr>
					<tr>
						<td class="label" align="right" nowrap="nowrap">Cell Phone:</td>
						<td>
							<input type="hidden" value="<%=sCell%>" name="usercell" />
							(<input type="text" value="<%=Left(sCell,3)%>" name="cell_areacode" onKeyUp="return autoTab(this, 3, event, false);" size="3" maxlength="3" />)&nbsp;
							<input type="text" value="<%=Mid(sCell,4,3)%>" name="cell_exchange" onKeyUp="return autoTab(this, 3, event, false);" size="3" maxlength="3" />&ndash;
							<input type="text" value="<%=Right(sCell,4)%>" name="cell_line" onKeyUp="return autoTab(this, 4, event, false);" size="4" maxlength="4" />
						</td>
					</tr>
					<tr>
						<td class="label" align="right" nowrap="nowrap">Fax:</td>
						<td>
							<input type="hidden" value="<%=sFax%>" name="userfax" />
							(<input type="text" value="<%=Left(sFax,3)%>" name="fax_areacode" onKeyUp="return autoTab(this, 3, event, false);" size="3" maxlength="3" />)&nbsp;
							<input type="text" value="<%=Mid(sFax,4,3)%>" name="fax_exchange" onKeyUp="return autoTab(this, 3, event, false);" size="3" maxlength="3" />&ndash;
							<input type="text" value="<%=Right(sFax,4)%>" name="fax_line" onKeyUp="return autoTab(this, 4, event, false);" size="4" maxlength="4" />
						</td>
					</tr>
<%					bHasResidentStreets = HasResidentTypeStreets( "R" )
					bFound = False 
					If bHasResidentStreets  Then 
%>
						<tr>
							<td class="label" align="right" valign="top" nowrap="nowrap">Resident Address:</td>
							<td>
<%								BreakOutAddress sAddress, sStreetNumber, sStreetName   ' In common.asp
								DisplayLargeAddressList "R", sStreetNumber, sStreetName, bFound %>&nbsp;
								<input type="button" class="button ui-button ui-widget ui-corner-all" value="Validate Address" onclick='checkAddress( "CheckResults", "no" );' />
							</td>
						</tr>
<%
					End If 
%>
					<tr>
						<td class="label" align="right" nowrap="nowrap">
							<% If bHasResidentStreets Then %>
								Address (if not listed):
							<% Else %>
								Address:
							<% End If %>
						</td>
						<td>
							<input type="text" value="<% If Not bFound Then 
														response.write sAddress
													End If %>" name="useraddress" onchange="FlagFamilyChange();" size="50" maxlength="100" />
						</td>
					</tr>
					<tr>
						<td class="label" align="right" nowrap="nowrap">Resident Unit:</td>
						<td><input type="text" value="<%=sUserUnit%>" name="userunit" onchange="FlagFamilyChange();" size="11" maxlength="10" />

<%					If OrgHasNeighborhoods( Session("orgid") ) Then %>
						</td>
					</tr>
					<tr>
						<td class="label" align="right">Neighborhood:</td>
						<td><% DisplayNeighborhoods Session("orgid"), iNeighborhoodid   ' In common.asp %> 
<%					Else %> 
						<input type="hidden" name="egov_users_neighborhoodid" value="0" />
<%					End If %>
						</td>
					</tr>
					<tr>
						<td class="label" align="right" nowrap="nowrap">City:</td>
						<td>
							<input type="text" value="<%=sCity%>" name="usercity" onchange="FlagFamilyChange();" size="50" maxlength="100" />
						</td>
					</tr>
					<tr>
						<td class="label" align="right" nowrap="nowrap">State:</td>
						<td>
							<input type="text" value="<%=sState%>" name="userstate" onchange="FlagFamilyChange();" size="2" maxlength="2" />
						</td>
					</tr>
					<tr>
						<td class="label" align="right" nowrap="nowrap">ZIP:</td>
						<td>
							<input type="text" value="<%=sZip%>" name="userzip" onchange="FlagFamilyChange();" size="10" maxlength="15" />
						</td>
					</tr>
					<tr>
						<td class="label" align="right" nowrap="nowrap">Business Name:</td>
						<td>
							<input type="text" value="<%=sBusinessName%>" name="userbusinessname" size="50" maxlength="100" />
						</td>
					</tr>
<%					bHasBusinessStreets = HasResidentTypeStreets( "B" )
					bFound = False 
					If bHasBusinessStreets  Then %>
						<tr>
							<td class="label" align="right" nowrap="nowrap">Business Street:</td>
							<td><% DisplayAddresses  "B", sBusinessAddress, bFound %></td>
						</tr>
<%					End If %>
					<tr>
						<td class="label" align="right" nowrap="nowrap">
			<%		If bHasBusinessStreets Then %>
							Street (if not listed):
			<%		Else %>
							Business Street:
			<%		End If %>
						</td>
						<td>
							<input type="text" value="<% If Not bFound Then 
											response.write sBusinessAddress
										 End If %>" name="userbusinessaddress" size="50" maxlength="100" />
						</td>
					</tr>
					<tr>
						<td class="label" align="right" nowrap="nowrap">Work Phone:</td>
						<td>
							<input type="hidden" value="<%=sWorkPhone%>" name="userworkphone" />
							(<input type="text" value="<%=Left(sWorkPhone,3)%>" name="work_areacode" onKeyUp="return autoTab(this, 3, event, false);" size="3" maxlength="3" />)&nbsp;
							<input type="text" value="<%=Mid(sWorkPhone,4,3)%>" name="work_exchange" onKeyUp="return autoTab(this, 3, event, false);" size="3" maxlength="3" />&ndash;
							<input type="text" value="<%=Mid(sWorkPhone,7,4)%>" name="work_line" onKeyUp="return autoTab(this, 4, event, false);" size="4" maxlength="4" />&nbsp;
							ext. <input type="text" value="<%=Mid(sWorkPhone,11,4)%>" name="work_ext" onKeyUp="return autoTab(this, 4, event, false);" size="4" maxlength="4" />
						</td>
					</tr>
					<tr>
						<td class="label" align="right" nowrap="nowrap">Emergency Contact:</td>
						<td>
							<input type="text" value="<%=sEmergencyContact%>" name="emergencycontact" style="width:300px;" maxlength="100" />
						</td>
					</tr>
					<tr>
						<td class="label" align="right" nowrap="nowrap">Emergency Phone:</td>
						<td>
							<input type="hidden" value="<%=sEmergencyPhone%>" name="emergencyphone" />
							(<input type="text" value="<%=Left(sEmergencyPhone,3)%>" name="emergencyphone_areacode" onKeyUp="return autoTab(this, 3, event, false);" size="3" maxlength="3" />)&nbsp;
							<input type="text" value="<%=Mid(sEmergencyPhone,4,3)%>" name="emergencyphone_exchange" onKeyUp="return autoTab(this, 3, event, false);" size="3" maxlength="3" />&ndash;
							<input type="text" value="<%=Mid(sEmergencyPhone,7,4)%>" name="emergencyphone_line" onKeyUp="return autoTab(this, 4, event, false);" size="4" maxlength="4" />
						</td>
					</tr>	
<%				If CLng(iUserId) = CLng(-1) Then		%>
					<tr><td>&nbsp;</td><td class="datacolumn"><input type="checkbox" name="iscontact" /> &nbsp; <strong>Check to add them to the Permit Contractors</strong></td></tr>
<%				End If %>
					<tr><td align="right"><font color="red">*</font> denotes required fields</td><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td>
						<td>
							<%					
							tooltipclass=""
							tooltip = ""
							disabled = ""
							If not bCanSaveChanges Then
								tooltipclass="tooltip"
								disabled = " disabled "
								tooltip = "<span class=""tooltiptext"">You don't have permission to save changes.</span>"
							end if
							%>
							<button <%=disabled%> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="doCheck();"><%=sSaveButton%><%=tooltip%></button>&nbsp;&nbsp;&nbsp;&nbsp;
							<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:CloseThis();" value="Close" />
						</td>
					</tr>
				</table>

			</form>
		</div>
	</div>

	<!--#Include file="../admin_footer.asp"-->  
	<!--#Include file="modal.asp"-->  

</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub GetPermitContactValues( iUserId, iPermitId )
'--------------------------------------------------------------------------------------------------
Sub GetPermitContactValues( iUserId, iPermitId )
	Dim sSql, oRs

	sSql = "SELECT * FROM egov_permitcontacts WHERE userid = " & iUserId & " AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sFirstName = oRs("firstname")
		sLastName = oRs("lastname")
		sAddress = oRs("address")
		sState = oRs("state")
		sCity = oRs("city")
		sZip = oRs("zip")
		sEmail = oRs("email")
		sFax = oRs("fax")
		sCell = oRs("cell")
		sBusinessName = oRs("company")
		sPassword = oRs("userpassword")
		sDayPhone = oRs("phone")
		sWorkPhone = oRs("userworkphone")
		sEmergencyContact = oRs("emergencycontact")
		sEmergencyPhone = oRs("emergencyphone")
		If IsNull(oRs("neighborhoodid")) Then
			iNeighborhoodid = 0
		else
			iNeighborhoodid = oRs("neighborhoodid")
		End If 
		If IsNull(oRs("residenttype")) Or oRs("residenttype") = "" Then
			sResidentType = "R"
		Else 
			sResidentType = oRs("residenttype")
		End If 
		sBusinessAddress = oRs("userbusinessaddress")
		sUserUnit = oRs("userunit")
		If oRs("emailnotavailable") Then 
			sEmailnotavailable = 1
		Else
			sEmailnotavailable = 0
		End If 
		If oRs("residencyverified") Then 
			sResidencyVerified = 1
		Else
			sResidencyVerified = 0
		End If 
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub GetRegisteredUserValues( iUserId )
'--------------------------------------------------------------------------------------------------
Sub GetRegisteredUserValues( iUserId )
	Dim sSql, oRs

	sSql = "SELECT * FROM egov_users WHERE userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
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
		sEmergencyContact = oRs("emergencycontact")
		sEmergencyPhone = oRs("emergencyphone")
		sBirthdate = oRs("birthdate")
		If IsNull(oRs("neighborhoodid")) Then
			iNeighborhoodid = 0
		else
			iNeighborhoodid = oRs("neighborhoodid")
		End If 
		If IsNull(oRs("residenttype")) Or oRs("residenttype") = "" Then
			sResidentType = "R"
		Else 
			sResidentType = oRs("residenttype")
		End If 
		sBusinessAddress = oRs("userbusinessaddress")
		If Not IsNull(oRs("familyid")) Then 
			iFamilyId = oRs("familyid")
		Else
			iFamilyId = iUserId
		End If 
		sUserUnit = oRs("userunit")
		If oRs("emailnotavailable") Then 
			sEmailnotavailable = 1
		Else
			sEmailnotavailable = 0
		End If 
		If oRs("residencyverified") Then 
			sResidencyVerified = 1
		Else
			sResidencyVerified = 0
		End If 
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION HasResidentTypes( )
'--------------------------------------------------------------------------------------------------
Function HasResidentTypes
	Dim sSql, oRs

	sSql = "SELECT COUNT(resident_type) AS hits FROM egov_poolpassresidenttypes WHERE orgid = " & session("orgid") & ""

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If CLng(oRs("hits")) > CLng(0) Then
		HasResidentTypes = True 
	Else
		HasResidentTypes = False 
	End if
	
	oRs.Close
	Set oRs = nothing

End Function 


'--------------------------------------------------------------------------------------------------
' FUNCTION DisplayResidentTypes( sResidentType )
'--------------------------------------------------------------------------------------------------
Function DisplayResidentTypes( sResidentType )
	Dim sSql, oResidentType

	sSql = "SELECT resident_type, description FROM egov_poolpassresidenttypes WHERE orgid = " & session("orgid") & " ORDER BY displayorder"

	Set oResidentType = Server.CreateObject("ADODB.Recordset")
	oResidentType.Open sSql, Application("DSN"), 3, 1

	DisplayResidentTypes = "<select name=""residenttype"">"	
		
	Do While NOT oResidentType.EOF 
		DisplayResidentTypes = DisplayResidentTypes & vbcrlf & "<option value=""" &  oResidentType("resident_type") & """"
		If sResidentType = oResidentType("resident_type") Then
			DisplayResidentTypes = DisplayResidentTypes & " selected=""selected"" "
		End If 
		DisplayResidentTypes = DisplayResidentTypes & ">" & oResidentType("description") & "</option>"
		oResidentType.MoveNext
	Loop

	DisplayResidentTypes = DisplayResidentTypes & "</select>"

	oResidentType.close
	Set oResidentType = Nothing 
	
End Function 


'--------------------------------------------------------------------------------------------------
' FUNCTION HasResidentTypeStreets( sResidenttype )
'--------------------------------------------------------------------------------------------------
Function HasResidentTypeStreets( sResidenttype )
	Dim sSql, oValues

	sSql = "SELECT COUNT(residentaddressid) AS hits FROM egov_residentaddresses WHERE orgid = " & session("orgid") & " AND residenttype = '" & sResidenttype & "'"

	Set oValues = Server.CreateObject("ADODB.Recordset")
	oValues.Open sSql, Application("DSN"), 3, 1

	If CLng(oValues("hits")) > CLng(0) Then
		HasResidentTypeStreets = True 
	Else
		HasResidentTypeStreets = False 
	End if
	
	oValues.Close
	Set oValues = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' Sub  DisplayAddresses( sResidenttype, sAddress, ByRef bFound )
'--------------------------------------------------------------------------------------------------
Sub  DisplayAddresses( sResidenttype, sAddress, ByRef bFound )
	Dim sSql, oAddressList

	sSql = "SELECT residentstreetnumber, residentstreetname FROM egov_residentaddresses_list WHERE orgid = " & session("orgid") & " and residenttype='" & sResidenttype & "' ORDER BY sortstreetname, residentstreetprefix, CAST(residentstreetnumber AS int)"

	Set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select name=""" & sResidenttype & "address"" onchange=""FlagFamilyChange();"">"	
	response.write vbcrlf & "<option value=""0000"">Please select an address...</option>"
		
	Do While NOT oAddressList.EOF 
		response.write vbcrlf & "<option value=""" &  oAddressList("residentstreetnumber") & " " & oAddressList("residentstreetname")  & """"
		If UCase(sAddress) = UCase(oAddressList("residentstreetnumber") & " " & oAddressList("residentstreetname")) Then 
			response.write " selected=""selected"" "
			bFound = True 
		End If 
		response.write ">" & oAddressList("residentstreetnumber") & " " & oAddressList("residentstreetname") & "</option>"
		oAddressList.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oAddressList.close
	Set oAddressList = Nothing 
	
End Sub  


'--------------------------------------------------------------------------------------------------
' Sub DisplayLargeAddressList( sResidenttype, sStreetNumber, sStreetName, bFound )
'--------------------------------------------------------------------------------------------------
Sub DisplayLargeAddressList( ByVal sResidenttype, ByVal sStreetNumber, ByVal sStreetName, ByRef bFound )
	Dim sSql, oAddressList, sCompareName

	If Not IsValidAddress( sStreetNumber, sStreetName ) Then   ' In common.asp
		sStreetNumber = ""
		sStreetName = ""
		bFound = False 
	End If 

	sSql = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSql = sSql & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSql = sSql & " FROM egov_residentaddresses WHERE orgid = " & session( "orgid" ) & " AND residenttype = '" & sResidenttype & "' "
	sSql = sSql & "AND residentstreetname IS NOT NULL ORDER BY sortstreetname"
	
	Set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSql, Application("DSN"), 3, 1

	If Not oAddressList.EOF Then
		response.write vbcrlf & "<input type=""text"" name=""residentstreetnumber"" value=""" & sStreetNumber & """ onchange=""FlagFamilyChange();"" size=""8"" maxlength=""10"" /> &nbsp; "
		response.write vbcrlf & "<select name=""address"" onchange=""FlagFamilyChange();"">"
		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown...</option>"
		Do While NOT oAddressList.EOF 
			sCompareName = ""
			If oAddressList("residentstreetprefix") <> "" Then
				sCompareName = oAddressList("residentstreetprefix") & " " 
			End If 
			sCompareName = sCompareName & oAddressList("residentstreetname")
			If oAddressList("streetsuffix") <> "" Then
				sCompareName = sCompareName & " "  & oAddressList("streetsuffix")
			End If
			If oAddressList("streetdirection") <> "" Then
				sCompareName = sCompareName & " "  & oAddressList("streetdirection")
			End If

			response.write vbcrlf & "<option value=""" & sCompareName & """"
			If sStreetName = sCompareName Then
				response.write " selected=""selected"" "
				bFound = True 
			End If 
			response.write " >"
			response.write sCompareName & "</option>"
			oAddressList.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oAddressList.Close
	Set oAddressList = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetUserIdByPermitContactId( iPermitContactId )
'--------------------------------------------------------------------------------------------------
Function GetUserIdByPermitContactId( iPermitContactId )
	Dim sSql, oRs

	sSql = "SELECT userid FROM egov_permitcontacts WHERE permitcontactid = " & iPermitContactId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetUserIdByPermitContactId = CLng(oRs("userid"))
	Else
		GetUserIdByPermitContactId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 



%>
