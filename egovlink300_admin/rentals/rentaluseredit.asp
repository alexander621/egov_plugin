<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentaluseredit.asp
' AUTHOR: Steve Loar
' CREATED: 10/12/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Edits Registered Users (citizens)
'
' MODIFICATION HISTORY
' 1.0   10/12/2009	Steve Loar - INITIAL VERSION
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
	iUserId = CLng(0)
End If 

If CLng(iUserId) > CLng(0) Then
	sTitle = "Edit Citizen User"
	sSaveButton = "Update"

	' Pull from the egov_users table
	GetRegisteredUserValues iUserId

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
	<link rel="stylesheet" type="text/css" href="../permits/permits.css" />

	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
	<script language="JavaScript" src="../scripts/removespaces.js"></script>
	<script language="JavaScript" src="../scripts/setfocus.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="Javascript">
	<!--
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
				document.frmCitizenUser.familyaddresschanged.value = "YES";
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
			document.frmCitizenUser.familyaddresschanged.value = "YES";
			//alert("yes");
		}

		function checkAddress( sReturnFunction, sSave )
		{
			// Remove any extra spaces
			document.frmCitizenUser.residentstreetnumber.value = removeSpaces(document.frmCitizenUser.residentstreetnumber.value);
			// check the number for non-numeric values
			var rege = /^\d+$/;
			var Ok = rege.exec(document.frmCitizenUser.residentstreetnumber.value);
			if ( ! Ok )
			{
				alert("The Resident Street Number cannot be blank and must be numeric.");
				setfocus(document.frmCitizenUser.residentstreetnumber);
				return false;
			}
			// check that they picked a street name
			if ( document.frmCitizenUser.address.value == '0000')
			{
				alert("Please select a street name from the list first.");
				setfocus(document.frmCitizenUser.address);
				return false;
			}
			// Fire off Ajax routine
			doAjax('checkaddress.asp', 'stnumber=' + document.frmCitizenUser.residentstreetnumber.value + '&stname=' + document.frmCitizenUser.address.value, sReturnFunction, 'get', '0');
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
				document.frmCitizenUser.useraddress.value = '';
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
			winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&parentform=frmCitizenUser' + '&stnumber=' + document.frmCitizenUser.residentstreetnumber.value + '&stname=' + document.frmCitizenUser.address.value + '&sCheckType=' + sReturnFunction + 'Validate", "_address", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function FinalCheck( sResults )
		{
			// Process the Ajax CallBack 
			if (sResults == 'FOUND')
			{
				validate();
			}
			else
			{
				PopAStreetPicker('FinalCheck', 'yes');
			}
			
		}

		function CloseThis()
		{
			window.close();
			window.opener.focus();
		}

		function doCheck()
		{
			// If they are using the large address feature
			var exists = eval(document.frmCitizenUser["residentstreetnumber"]);
			if(exists)
			{
				// If a street number was entered
				if (document.frmCitizenUser.residentstreetnumber.value != '')
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
			doAjax('checkduplicatecitizens.asp', 'userlname=' + document.frmCitizenUser.userlname.value, 'OkToValidate', 'get', '0');
		}

		function OkToValidate( sReturn )
		{
			//finish the validation routine 
			validate();
		}

		function validate()
		{
			var msg = "";

			if ( document.frmCitizenUser.emailnotavailable.checked == false )
			{
				// They will login so validate the email and password
				if (document.frmCitizenUser.useremail.value == "" )
				{
					msg+="The email cannot be blank.\n";
				}
				//var rege = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*\.(\w{2}|(com|net|org|edu|mil|gov|biz|us))$/;
				var rege = /.+@.+\..+/i;
				var Ok = rege.test(document.frmCitizenUser.useremail.value);

				if (! Ok)
				{
					msg+="The email must be in a valid format.\n";
				}
				if (document.frmCitizenUser.userpassword.value == "" )
				{
					msg+="The password cannot be blank.\n";
				}
				if (document.frmCitizenUser.userpassword2.value == "" )
				{
					msg+="The verify password cannot be blank.\n";
				}
				if(document.frmCitizenUser.userpassword.value != document.frmCitizenUser.userpassword2.value)
				{
					msg+="The passwords you have entered do not match.\n";
				}
			}
			if (document.frmCitizenUser.userfname.value == "" )
			{
				msg+="The first name cannot be blank.\n";
			}
			if (document.frmCitizenUser.userlname.value == "" )
			{
				msg+="The last name cannot be blank.\n";
			}

			// set the emergency phone
			if (document.frmCitizenUser.emergencyphone_areacode.value != "" || document.frmCitizenUser.emergencyphone_exchange.value != "" || document.frmCitizenUser.emergencyphone_line.value != "" )
			{
				var ePhone = document.frmCitizenUser.emergencyphone_areacode.value + document.frmCitizenUser.emergencyphone_exchange.value + document.frmCitizenUser.emergencyphone_line.value;
				if (ePhone.length < 10)
				{
					msg += "The Emergency Phone must be a valid phone number, including area code, or blank\n";
				}
				else
				{
					document.frmCitizenUser.emergencyphone.value = document.frmCitizenUser.emergencyphone_areacode.value + document.frmCitizenUser.emergencyphone_exchange.value + document.frmCitizenUser.emergencyphone_line.value;
					var rege = /^\d+$/;
					var Ok = rege.exec(document.frmCitizenUser.emergencyphone.value);
					if ( ! Ok )
					{
						msg += "The Emergency Phone must be a valid phone number, including area code, or blank\n";
					}
				}
			}

			// check the work phone
			if (document.frmCitizenUser.work_areacode.value != "" || document.frmCitizenUser.work_exchange.value != "" || document.frmCitizenUser.work_line.value != "" || document.frmCitizenUser.work_ext.value != "")
			{
				var sPhone = document.frmCitizenUser.work_areacode.value + document.frmCitizenUser.work_exchange.value + document.frmCitizenUser.work_line.value;
				if (sPhone.length < 10)
				{
					msg += "Work Phone Number must be a valid phone number, including area code, or blank\n";
				}
				else
				{
					document.frmCitizenUser.userworkphone.value = document.frmCitizenUser.work_areacode.value + document.frmCitizenUser.work_exchange.value + document.frmCitizenUser.work_line.value + document.frmCitizenUser.work_ext.value;
					var rege = /^\d+$/;
					var Ok = rege.exec(document.frmCitizenUser.userworkphone.value);
					if ( ! Ok )
					{
						msg += "Work Phone Number must be a valid phone number, including area code, or blank\n";
					}
				}
			}

			// check the fax
			if (document.frmCitizenUser.fax_areacode.value != "" || document.frmCitizenUser.fax_exchange.value != "" || document.frmCitizenUser.fax_line.value != "" )
			{
				var fPhone = document.frmCitizenUser.fax_areacode.value + document.frmCitizenUser.fax_exchange.value + document.frmCitizenUser.fax_line.value;
				if (fPhone.length < 10)
				{
					msg += "Fax must be a valid phone number, including area code, or blank\n";
				}
				else
				{
					document.frmCitizenUser.userfax.value = document.frmCitizenUser.fax_areacode.value + document.frmCitizenUser.fax_exchange.value + document.frmCitizenUser.fax_line.value;
					var rege = /^\d+$/;
					var Ok = rege.exec(document.frmCitizenUser.userfax.value);
					if ( ! Ok )
					{
						msg += "Fax must be a valid phone number, including area code, or blank\n";
					}
				}
			}

			// check the cell
			if (document.frmCitizenUser.cell_areacode.value != "" || document.frmCitizenUser.cell_exchange.value != "" || document.frmCitizenUser.cell_line.value != "" )
			{
				var cPhone = document.frmCitizenUser.cell_areacode.value + document.frmCitizenUser.cell_exchange.value + document.frmCitizenUser.cell_line.value;
				if (cPhone.length < 10)
				{
					msg += "The cell phone must be a valid phone number, including area code, or blank\n";
				}
				else
				{
					document.frmCitizenUser.usercell.value = document.frmCitizenUser.cell_areacode.value + document.frmCitizenUser.cell_exchange.value + document.frmCitizenUser.cell_line.value;
					var crege = /^\d+$/;
					var cOk = crege.exec(document.frmCitizenUser.usercell.value);
					if ( ! cOk )
					{
						msg += "The cell phone must be a valid phone number, including area code, or blank\n";
					}
				}
			}

			// check the home phone number
			document.frmCitizenUser.userhomephone.value = document.frmCitizenUser.user_areacode.value + document.frmCitizenUser.user_exchange.value + document.frmCitizenUser.user_line.value;
			if (document.frmCitizenUser.userhomephone.value != "" )
			{
				var hPhone = document.frmCitizenUser.userhomephone.value;
				if (hPhone.length < 10)
				{
					msg += "The home phone must be a valid phone number, including area code.\n";
				}
				else
				{
					var rege = /^\d+$/;
					var Ok = rege.exec(document.frmCitizenUser.userhomephone.value);
					if ( ! Ok )
					{
						msg += "The home phone must be a valid phone number, including area code.\n";
					}
				}
			}

			// Process the business address if one was chosen
			var bexists = eval(document.frmCitizenUser["Baddress"]);
			if(bexists)
			{
				//See if they picked from the business dropdown and put that in the address field 
				if (document.frmCitizenUser.Baddress.selectedIndex > -1)
				{
					var belement = document.frmCitizenUser.Baddress;
					var bselectedvalue = belement.options[belement.selectedIndex].value;

					//alert( bselectedvalue );
					//  0000 is the first pick that we do not want
					if (bselectedvalue != "0000")
					{
						document.frmCitizenUser.userbusinessaddress.value = bselectedvalue;
					}
				}
			}

			// Process the resident address if one was chosen - this is second to set the local resident type
			var exists = eval(document.frmCitizenUser["Raddress"]);
			if(exists)
			{
				// See if they picked from the resident dropdown and put that in the address field 
				if (document.frmCitizenUser.Raddress.selectedIndex > -1)
				{
					var element = document.frmCitizenUser.Raddress;
					var selectedvalue = element.options[element.selectedIndex].value;

					//alert( selectedvalue );
					//  0000 is the first pick that we do not want
					if (selectedvalue != "0000")
					{
						document.frmCitizenUser.useraddress.value = selectedvalue;
					}
				}
			}

			// handle the large quantity street addresses
			exists = eval(document.frmCitizenUser["residentstreetnumber"]);
			if(exists)
			{
				if ( document.frmCitizenUser.residentstreetnumber.value != '' )
				{
					// See if they picked from the resident dropdown and put that in the address field 
					if (document.frmCitizenUser.address.selectedIndex > -1)
					{
						var element = document.frmCitizenUser.address;
						var selectedvalue = element.options[element.selectedIndex].value;

						//alert( selectedvalue );
						//  0000 is the first pick that we do not want
						if (selectedvalue != "0000")
						{
							document.frmCitizenUser.useraddress.value = document.frmCitizenUser.residentstreetnumber.value + ' ' + selectedvalue;
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
				if (document.frmCitizenUser.familyaddresschanged.value == "YES")
				{
					if ((document.frmCitizenUser.userid.value != 0) && (document.frmCitizenUser.hasfamilymembers.value == 'true'))
					{
						var bCopyToAll = confirm("Copy changes to all family members?");
						if ( ! bCopyToAll )
						{
							document.frmCitizenUser.familyaddresschanged.value = "NO";
						}
					}
					else
					{
						document.frmCitizenUser.familyaddresschanged.value = "NO";
					}
				}
				//alert("OK");

				// Build the post parameter 
				var sParameter = 'userid=' + encodeURIComponent(document.frmCitizenUser.userid.value);
				sParameter += '&userfname=' + encodeURIComponent(document.frmCitizenUser.userfname.value);
				sParameter += '&userlname=' + encodeURIComponent(document.frmCitizenUser.userlname.value);
				sParameter += '&emailnotavailable=' + encodeURIComponent(document.frmCitizenUser.emailnotavailable.checked);
				sParameter += '&useremail=' + encodeURIComponent(document.frmCitizenUser.useremail.value);
				sParameter += '&userpassword=' + encodeURIComponent(document.frmCitizenUser.userpassword.value);
				sParameter += '&residenttype=' + encodeURIComponent(document.frmCitizenUser.residenttype.value);
				sParameter += '&residencyverified=' + encodeURIComponent(document.frmCitizenUser.residencyverified.value);
				sParameter += '&usercity=' + encodeURIComponent(document.frmCitizenUser.usercity.value);
				sParameter += '&userstate=' + encodeURIComponent(document.frmCitizenUser.userstate.value);
				sParameter += '&userzip=' + encodeURIComponent(document.frmCitizenUser.userzip.value);
				sParameter += '&userhomephone=' + encodeURIComponent(document.frmCitizenUser.userhomephone.value);
				sParameter += '&usercell=' + encodeURIComponent(document.frmCitizenUser.usercell.value);
				sParameter += '&userworkphone=' + encodeURIComponent(document.frmCitizenUser.userworkphone.value);
				sParameter += '&userfax=' + encodeURIComponent(document.frmCitizenUser.userfax.value);
				sParameter += '&emergencyphone=' + encodeURIComponent(document.frmCitizenUser.emergencyphone.value);
				sParameter += '&neighborhoodid=' + encodeURIComponent(document.frmCitizenUser.egov_users_neighborhoodid.value);
				sParameter += '&userunit=' + encodeURIComponent(document.frmCitizenUser.userunit.value);
				sParameter += '&userbusinessname=' + encodeURIComponent(document.frmCitizenUser.userbusinessname.value);
				sParameter += '&emergencycontact=' + encodeURIComponent(document.frmCitizenUser.emergencycontact.value);
				sParameter += '&useraddress=' + encodeURIComponent(document.frmCitizenUser.useraddress.value);
				sParameter += '&userbusinessaddress=' + encodeURIComponent(document.frmCitizenUser.userbusinessaddress.value);
				sParameter += '&familyaddresschanged=' + encodeURIComponent(document.frmCitizenUser.familyaddresschanged.value);
				sParameter += '&familyid=' + encodeURIComponent(document.frmCitizenUser.familyid.value);
				sParameter += '&notifyuser=' + encodeURIComponent(document.frmCitizenUser.notifyuser.checked);
				
				//alert( sParameter );

				// Fire off and close
				doAjax('rentalusersave.asp', sParameter, '', 'post', '0');
				window.opener.displayScreenMsg('Updates to the Registered User have been saved.');
				window.close();

				//document.frmCitizenUser.submit();
			}
		}

		function CloseThisSaved( sResult )
		{
			var sDetailText;
			var sToolTip = '';
			//alert( sResult );
			// optionally put code here to update the parent window
			<% If bUpdateParent Then %>
				if (parseInt(document.frmCitizenUser.userid.value) > 0)
				{
					sDetailText = '<a href="javascript:';
					<% If sDetailId = "applicantdetails" Then %>
						sDetailText += 'EditApplicant';
					<% Else %>
						sDetailText += 'EditPrimaryContact';
					<% End If %>
					sDetailText += '(\'' + document.frmCitizenUser.userid.value + '\' );" ';
					if (document.frmCitizenUser.userfname.value != "")
					{
						sToolTip += '<strong>' + document.frmCitizenUser.userfname.value + ' ' + document.frmCitizenUser.userlname.value + '</strong><br />';
					}
					if (document.frmCitizenUser.userbusinessname.value != "")
					{
						sToolTip += document.frmCitizenUser.userbusinessname.value + '<br />';
					}
					
					if (document.frmCitizenUser.useraddress.value != "")
					{
						sToolTip += document.frmCitizenUser.useraddress.value + '<br />';
					}
					if (document.frmCitizenUser.usercity.value != "")
					{
						sToolTip += document.frmCitizenUser.usercity.value + ', ' + document.frmCitizenUser.userstate.value + ' ' + document.frmCitizenUser.userzip.value + '<br />';
					}
					if (document.frmCitizenUser.user_areacode.value != "")
					{
						sToolTip += '(' + document.frmCitizenUser.user_areacode.value + ') ' + document.frmCitizenUser.user_exchange.value + '-' + document.frmCitizenUser.user_line.value;
					}
					var myRegExp = /\'/g;
					sToolTip = sToolTip.replace(myRegExp, '\\&#39;');
					sDetailText += ' onMouseover="ddrivetip(\'' + sToolTip + '\', 300)"; onMouseout="hideddrivetip()"; '
					sDetailText += '>';
					sDetailText += document.frmCitizenUser.userfname.value + ' ' + document.frmCitizenUser.userlname.value;
					if (document.frmCitizenUser.userbusinessname.value != "")
					{
						sDetailText += ' (' + document.frmCitizenUser.userbusinessname.value + ')';
					}

					sDetailText += '</a>';
					//alert( sDetailText );
					//alert('<%=sDetailId%>');
					window.opener.document.getElementById('<%=sDetailId%>').innerHTML = sDetailText;
					//window.opener.NiceTitles.autoCreated.anchors.addElements(window.opener.document.getElementsByTagName("a"), "title");
				}
				else
				{
					// Put stuff to select the contact here
					window.opener.document.getElementById("primarycontactdetails").innerHTML = document.frmCitizenUser.userfname.value + ' ' + document.frmCitizenUser.userlname.value;
					window.opener.document.getElementById("isprimarycontactuserid").value = sResult;
				}
			<% End If %>

			window.close();
			window.opener.hideddrivetip();
			window.opener.focus();
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

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong><%=sTitle%></strong></font><br /><br />
			</p>
			<!--END: PAGE TITLE-->
			<form name="frmCitizenUser" method="post" action="rentalusersave.asp">
				<input type="hidden" name="userid" value="<%=iUserId%>" />
				<input type="hidden" name="familyid" value="<%=iFamilyId%>" />
				<input type="hidden" name="familyaddresschanged" value="NO" />
				<input type="hidden" name="residencyverified" value="<%=sResidencyVerified%>" />
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
					<tr><td class="label" align="right" nowrap="nowrap">Email:</td>
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
							<input type="button" value="Generate Random Password" name="generate_pwd" onclick="GetRndPwd();" />
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
								<input type="button" class="button" value="Validate Address" onclick='checkAddress( "CheckResults", "no" );' />
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
					<tr><td align="right"><font color="red">*</font> denotes required fields</td><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td>
						<td>
							<input type="button" class="button" onclick="javascript:doCheck();" value="<%=sSaveButton%>" />&nbsp;&nbsp;&nbsp;&nbsp;
							<input type="button" class="button" onclick="javascript:CloseThis();" value="Close" />
						</td>
					</tr>
				</table>

			</form>
		</div>
	</div>

	<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub GetRegisteredUserValues( iUserId )
'--------------------------------------------------------------------------------------------------
Sub GetRegisteredUserValues( iUserId )
	Dim sSql, oRs

	sSql = "SELECT * FROM egov_users WHERE orgid = " & session("orgid") & " AND userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

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
	oRs.Open sSql, Application("DSN"), 0, 1

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
	oResidentType.Open sSql, Application("DSN"), 0, 1

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
	oValues.Open sSql, Application("DSN"), 0, 1

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
	oAddressList.Open sSql, Application("DSN"), 0, 1

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
	oAddressList.Open sSql, Application("DSN"), 0, 1

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




%>
