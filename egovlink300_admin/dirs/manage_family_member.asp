<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="citizen_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: manage_family_member.asp
' AUTHOR: Steve Loar
' CREATED: 1/10/2007 - Copied from update_citizen.asp
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  family member information management.
'
' MODIFICATION HISTORY
' 1.0   1/10/2007	Steve Loar - Initial code 
' 1.1	10/05/2011	Steve Loar - Added gender selection pick
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sError, sFirstName,sLastName,sAddress,sCity,sState,sZip,sPhone,sEmail,sFax,sCell,sBusinessName,sDayPhone,sPassword,iUserID
Dim bHasResidentStreets, bFound, sResidentType, sBusinessAddress, bHasBusinessStreets, sRedirect, bHasResidentTypes, sWorkPhone
Dim sEmergencyPhone, sEmergencyContact, iNeighborhoodid, sResidencyVerified, sBirthdate, iRelationshipId, sRelationship, sUserUnit
Dim sStreetNumber, sStreetName, sGender, bShowGenderPicks, iFamilyId, sHeadOfHouseholdName, sSaveLabel

sLevel = "../" ' Override of value from common.asp

sStreetNumber = ""
sStreetName = ""

PageDisplayCheck "edit citizens", sLevel	' In common.asp

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then   ' This should be the page saving itself
	If CLng(request("userid")) <> CLng(0) Then 
		' update the egov_users table
		UpdateRecords()
		iUserId = request("userid")
		FamilyMemberUpdate iUserId, request("egov_users_userfname"), request("egov_users_userlname"), request("skip_egov_users_relationship"), request("egov_users_birthdate")
	Else
		' Do an Insert Here and retreive the new userid
		iUserId = InsertRecord()
		AddFamilyMember request("egov_users_familyid"), request("egov_users_userfname"), request("egov_users_userlname"), request("skip_egov_users_relationship"), request("egov_users_birthdate"), iUserId
	End If 
	UpdateOverflowFields iUserId, request("egov_users_userunit")

	If CLng(request("returnto")) <> CLng(0) Then 
		If CLng(request("returnto")) = CLng(-1) Then 
			sRedirect = Session("RedirectPage") 
			'Session("RedirectPage") = ""
			response.redirect sRedirect
		Else
			' Default them back to the Family List page
			response.redirect "family_list.asp?userid=" & request("returnto")
		End If 
	Else
		response.redirect Session("RedirectPage")
	End If 

End If

' This is where displaying the page starts

iUserID = request("u")
iReturn = request("iReturn")

If CLng(request("iReturn")) <> CLng(0) Then 
	If CLng(request("iReturn")) = CLng(-1) Then 
		iFamilyId = GetFamilyId( iUserID )
	Else 
		iFamilyId = GetFamilyId( iReturn )
	End If 
	
Else
	' new Family member
	iFamilyId = GetFamilyId( iUserID )
End If 

If CLng(iUserID) = CLng(0) Then
	GetNewFamilyValues iReturn 
	sSaveLabel = "Create Family Member"
Else 
	GetRegisteredUserValues iUserID 
	sSaveLabel = "Save Changes"
End If 

sRelationship = GetRelationShip( iRelationshipId )

bShowGenderPicks = orgHasFeature( "display gender pick" )

%>

<html>
<head>
	<title><%=langBSCommittees%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script language="JavaScript" src="../scripts/jquery-1.4.2.min.js"></script>

	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
	<script language="JavaScript" src="../scripts/formatnumber.js"></script>
	<script language="JavaScript" src="../scripts/removespaces.js"></script>
	<script language="JavaScript" src="../scripts/setfocus.js"></script>
	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>

<script language="javascript">
<!--
	var winHandle;
	var w = (screen.width - 640)/2;
	var h = (screen.height - 450)/2;

	function FlagFamilyChange()
	{

		<% If CLng(iUserID) > CLng(0) Then %>
			document.register.familyaddresschanged.value = "YES";
			//alert("setting to yes");
		<% else %>
			document.register.familyaddresschanged.value = "NO";
			//alert("setting to no");
		<% end if %>
		//alert(document.register.familyaddresschanged.value);
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
		//winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&stnumber=' + document.register.residentstreetnumber.value + '&stname=' + document.register.skip_address.value + '", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
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
		winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&stnumber=' + document.register.residentstreetnumber.value + '&stname=' + document.register.skip_address.value + '&sCheckType=' + sReturnFunction + 'Validate", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		//winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&stnumber=' + document.register.residentstreetnumber.value + '&stname=' + document.register.skip_address.value + '&sCheckType=' + sReturnFunction + 'Validate", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
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

	function finalCheckValidate()
	{
		validate();
	}

function validate() {
	var msg="";

	//var rege = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*\.(\w{2}|(com|net|org|edu|mil|gov|biz))$/;
	var rege = /.+@.+\..+/i;
	
	if (document.register.egov_users_userfname.value == "" )
	{
		msg+="The first name cannot be blank.\n";
	}
	if (document.register.egov_users_userlname.value == "" )
	{
		msg+="The last name cannot be blank.\n";
	}

	// set the emergency phone
	if ($("#skip_emergencyphone_areacode").val() != "" || $("#skip_emergencyphone_exchange").val() != "" || $("#skip_emergencyphone_line").val() != "" )
	{
		var ePhone = $("#skip_emergencyphone_areacode").val() + $("#skip_emergencyphone_exchange").val() + $("#skip_emergencyphone_line").val();
		if (ePhone.length < 10)
		{
			msg += "The Emergency Phone must be a valid phone number, including area code, or blank\n";
		}
		else
		{
			$("#egov_users_emergencyphone").val( $("#skip_emergencyphone_areacode").val() + $("#skip_emergencyphone_exchange").val() + $("#skip_emergencyphone_line").val() );
			var rege = /^\d+$/;
			var Ok = rege.exec($("#egov_users_emergencyphone").val());
			if ( ! Ok )
			{
				msg += "The Emergency Phone must be a valid phone number, including area code, or blank\n";
			}
		}
	}
	else
	{
		$("#egov_users_emergencyphone").val( $("#skip_emergencyphone_areacode").val() + $("#skip_emergencyphone_exchange").val() + $("#skip_emergencyphone_line").val() );
	}

	// check the work phone
	if (document.register.skip_work_areacode.value != "" || document.register.skip_work_exchange.value != "" || document.register.skip_work_line.value != "" || document.register.skip_work_ext.value != "")
	{
		var sPhone = document.register.skip_work_areacode.value + document.register.skip_work_exchange.value + document.register.skip_work_line.value;
		if (sPhone.length < 10)
		{
			msg += "Work Phone Number must be a valid phone number, including area code, or blank\n";
		}
		else
		{
			document.register.egov_users_userworkphone.value = document.register.skip_work_areacode.value + document.register.skip_work_exchange.value + document.register.skip_work_line.value + document.register.skip_work_ext.value;
			var rege = /^\d+$/;
			var Ok = rege.exec(document.register.egov_users_userworkphone.value);
			if ( ! Ok )
			{
				msg += "Work Phone Number must be a valid phone number, including area code, or blank\n";
			}
		}
	}
	else
	{
		$("#egov_users_userworkphone").val( '' );
	}

	// check the fax
	if (document.register.skip_fax_areacode.value != "" || document.register.skip_fax_exchange.value != "" || document.register.skip_fax_line.value != "" )
	{
		var fPhone = document.register.skip_fax_areacode.value + document.register.skip_fax_exchange.value + document.register.skip_fax_line.value;
		if (fPhone.length < 10)
		{
			msg += "Fax must be a valid phone number, including area code, or blank\n";
		}
		else
		{
			document.register.egov_users_userfax.value = document.register.skip_fax_areacode.value + document.register.skip_fax_exchange.value + document.register.skip_fax_line.value;
			var rege = /^\d+$/;
			var Ok = rege.exec(document.register.egov_users_userfax.value);
			if ( ! Ok )
			{
				msg += "Fax must be a valid phone number, including area code, or blank\n";
			}
		}
	}

	// check the cell phone
	if (document.register.skip_cell_areacode.value != "" || document.register.skip_cell_exchange.value != "" || document.register.skip_cell_line.value != "" )
	{
		var cPhone = document.register.skip_cell_areacode.value + document.register.skip_cell_exchange.value + document.register.skip_cell_line.value;
		if (cPhone.length < 10)
		{
			msg += "The cell phone must be a valid phone number, including area code, or blank\n";
		}
		else
		{
			document.register.egov_users_usercell.value = document.register.skip_cell_areacode.value + document.register.skip_cell_exchange.value + document.register.skip_cell_line.value;
			var crege = /^\d+$/;
			var cOk = crege.exec(document.register.egov_users_usercell.value);
			if ( ! cOk )
			{
				msg += "The cell phone must be a valid phone number, including area code, or blank\n";
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
			msg += "The home phone must be a valid phone number, including area code.\n";
		}
		else
		{
			var rege = /^\d+$/;
			var Ok = rege.exec(document.register.egov_users_userhomephone.value);
			if ( ! Ok )
			{
				msg += "The home phone must be a valid phone number, including area code.\n";
			}
		}
	}
//	else
//	{
//		msg+="The home phone cannot be blank.\n";
//	}

	// Handle the birthdate - Required for Children
	var relationshipidexists = eval(document.register["egov_users_relationshipid"]);

	var birthrege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
	var birthOk = birthrege.test( $("#egov_users_birthdate").val() );
	//alert( $("#egov_users_birthdate").val() );

	//alert( new Date($("#egov_users_birthdate").val()).toDateString() );
	
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
				if ($("#egov_users_birthdate").val() == "")
				{
					msg += "Please input a birth date for this child in the format of MM/DD/YYYY.";
				}
				else
				{
					if (! birthOk )
					{
						msg += "Birth date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.";
					}
					else
					{
						if (isValidDate( $("#egov_users_birthdate").val() ) == false)
						{
							msg += "Birth date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.";
						}
						else
						{
							if (new Date('1/1/1900') > new Date($("#egov_users_birthdate").val()))
							{
								msg += 'The Birthdate must be greater than 1/1/1900.';
							}
						}
					}
				}
			}
			else
			{
				if ($("#egov_users_birthdate").val() != "")
				{
					if (! birthOk )
					{
						msg += "Birth date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again, or leave it blank.";
					}
					else
					{
						if (isValidDate( $("#egov_users_birthdate").val() ) == false)
						{
							msg += "Birth date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.";
						}
						else
						{
							if (new Date('1/1/1900') > new Date($("#egov_users_birthdate").val()))
							{
								msg += 'The Birthdate must be greater than 1/1/1900.';
							}
						}
					}
				}
			}
		}
		else
		{
			if ($("#egov_users_birthdate").val() != "")
			{
				if (! birthOk )
				{
					msg += "Birth date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again, or leave it blank.";
				}
				else
				{
					if (isValidDate( $("#egov_users_birthdate").val() ) == false)
					{
						msg += "Birth date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.";
					}
					else
					{
						if (new Date('1/1/1900') > new Date($("#egov_users_birthdate").val()))
						{
							msg += 'The Birthdate must be greater than 1/1/1900.';
						}
					}
				}
			}
		}
	}
	else
	{
		if ($("#egov_users_birthdate").val() != "")
		{
			if (! birthOk )
			{
				msg += "Birth date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again, or leave it blank.";
			}
			else
			{
				if (isValidDate( $("#egov_users_birthdate").val() ) == false)
				{
					msg += "Birth date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.";
				}
				else
				{
					if (new Date('1/1/1900') > new Date($("#egov_users_birthdate").val()))
					{
						msg += 'The Birthdate must be greater than "1/1/1900".';
					}
				}
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

	if(msg != "")
	{
		msg="Your form could not be submitted for the following reasons.\n\n" + msg;
		alert(msg);
	}
	else 
	{	
		if (document.register.familyaddresschanged.value == "YES")
		{
			if (!confirm("Copy changes to all family members?"))
			{
				//alert("canceled");
				document.register.familyaddresschanged.value = "NO"
			}
		}
		//alert('Flag = ' + document.register.familyaddresschanged.value);
		document.register.submit(); 
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

//-->
</script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

  <% ShowHeader sLevel %>
  <!--#Include file="../menu/menu.asp"--> 

<div id="content">
	<div id="centercontent">

  <table border="0" cellpadding="10" cellspacing="0" width="100%">
	<tr>
      <td>
		<font size="+1"><b>
<%			If CLng(iUserID) = CLng(0) Then %>
				New Family Member of 
<%			Else %>
				Family Member of 
<%			End If 
			sHeadOfHouseholdName = GetFamilyOwnerName( iFamilyId )
			If sHeadOfHouseholdName = "" Then
				' if the email is blank, then GetFamilyOwner() will return nothing so try head of household
				sHeadOfHouseholdName = GetHeadOfHouseholdName( iFamilyId )
			End If 
			response.write sHeadOfHouseholdName
%>
			</b></font><br /><br />

<%	If CLng(request("iReturn")) <> CLng(0) Then 
		If CLng(request("iReturn")) = CLng(-1) Then %>
			<a href="javascript:GoBack('<%=Session("RedirectPage") %>');"><img src='../images/arrow_2back.gif' border="0" align='absmiddle'>&nbsp;&nbsp;Back</a>  
	<%	 Else 
			'Session("RedirectPage") = "family_list.asp?userid=" & iReturn
			'Session("RedirectLang") = "Return to Family List"
            if Len(Session("RedirectLang")) > 0 then
			   lcl_return_label = Session("RedirectLang")
			   lcl_return_url   = Session("RedirectPage")
			else
               lcl_return_label = "Return to the Family List"
			   lcl_return_url = "javascript:GoBack('family_list.asp?userid=" & iReturn & "');"
			end if
    %>
			<a href="<%=lcl_return_url%>"><img src='../images/arrow_2back.gif' border="0" align='absmiddle'>&nbsp;&nbsp;<%=lcl_return_label%></a>  
	<% End If %>
<%	Else 
		'Session("RedirectPage") = "display_citizen.asp?v=3"
		'Session("RedirectLang") = "Return to Citizen List"%>
		<a href="javascript:GoBack('<%=Session("RedirectPage")%>');"><img src='../images/arrow_2back.gif' border="0" align='absmiddle'>&nbsp;&nbsp;Return to Citizen List</a>  
<%	End If %>
		
	  </td>
      <td width="200">&nbsp;</td>
    </tr>
	<tr>
      <td valign="top">

	  <form method="post" name="register" action="manage_family_member.asp">
		<input type="hidden" name="columnnameid" value="userid">
		<input type="hidden" name="egov_users_userregistered" value="1">
		<input type="hidden" name="egov_users_orgid" value="<%=session("orgid")%>">
		<input type="hidden" name="userid" value="<%=iUserID%>">
		<!--<input type="hidden" name="ef:egov_users_userhomephone-text/req/phone" value="Home Phone Number">-->
		<input type="hidden" name="ef:egov_users_userfname-text/req" value="First name">
		<input type="hidden" name="ef:egov_users_userlname-text/req" value="Last name">
		<input type="hidden" name="egov_users_residenttype" value="R">
		<input type="hidden" name="egov_users_familyid" value="<%=iFamilyId%>" />
		<input type="hidden" name="returnto" value="<%=iReturn%>" />
		<input type="hidden" name="skip_egov_users_relationship" value="<%=sRelationship%>" />
		<input type="hidden" name="familyaddresschanged" value="NO" />
<%
		If Not bShowGenderPicks Then 
			response.write vbcrlf & "<input type=""hidden"" id=""egov_users_gender"" name=""egov_users_gender"" value=""N"" />"
		End If 
%>

		 <div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:FamilyList('<%=iFamilyId%>');"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:doCheck();"><%=sSaveLabel%></a>
			&nbsp;&nbsp;<img src="<%=RootPath%>images/newgroup.gif" width="16" height="16" align="absmiddle">&nbsp;<a href="javascript:FamilyList('<%=iFamilyId%>');">View Their Family Members</a>
		 </div>

		<table border="0" class="tableadmin" id="registertable" cellpadding="4" cellspacing="0">
		<tr>
		<th align="left">Property</th>
		<th align="left">Value</th></tr>
		<tr><td class="label" align="right" nowrap="nowrap">
			<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color="red">*</font></span> 
			First Name:
			</span>
		</td><td>
			<span class="cot-text-emphasized" title="This field is required"> 
			<input type="text" value="<%=sFirstName%>" name="egov_users_userfname" size="50" maxlength="100">
			</span>
		</td></tr>
		<tr><td class="label" align="right" nowrap="nowrap">
			<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color="red">*</font></span>
			Last Name:
			</span>
		</td><td>
			<span class="cot-text-emphasized" title="This field is required">
			<input type="text" value="<%=sLastName%>" name="egov_users_userlname" size="50" maxlength="100">
			</span>
		</td></tr>
<%
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
%>
		<tr><td class="label" align="right">
			Birthdate:
		</td><td>
			<input type="text" value="<%=sBirthdate%>" id="egov_users_birthdate" name="egov_users_birthdate" size="10" maxlength="10" /> (MM/DD/YYYY)

<%		If CLng(iUserId) <> CLng(iReturn) Then %>
			</td></tr>
			<tr><td class="label" align="right">
				Relationship:
			</td><td>
				<% DisplayRelationships session("orgid"), iRelationshipId %>
			</td></tr>
<%		Else %>
			<input type="hidden" name="egov_users_relationshipid" value="<%=iRelationshipId%>" id="relationship" />
<%		End If %>

		</td></tr>

<%		bHasResidentTypes = HasResidentTypes()
		bFound = False 
		If bHasResidentTypes Then %>
			<tr><td class="label" align="right" nowrap="nowrap">
				<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color=red>*</font></span>
					Resident Type:</span>
				</td><td>
					<%=DisplayResidentTypes( sResidentType ) %>
					<% 
						If OrgHasFeature( "residency verification" ) Then %>
							&nbsp; <input name="egov_users_residencyverified" type="checkbox" <%=sResidencyVerified%> /> Residency Verified
					<%	End If %>
			</td></tr>
<%		End If %>
		<tr><td class="label" align="right" nowrap="nowrap">
			<!--<font color=red>*</font>-->Home Phone:
		</td><td>
			<!--<input type="text" value="<%=sDayPhone%>" name="egov_users_userhomephone" size="50" maxlength="100">-->
			<input type="hidden" value="<%=sDayPhone%>" name="egov_users_userhomephone">
			(<input type="text" value="<%=Left(sDayPhone,3)%>" name="skip_user_areacode" onKeyUp="return autoTab(this, 3, event, true);" onchange="FlagFamilyChange();" size="3" maxlength="3">)&nbsp;
			<input type="text" value="<%=Mid(sDayPhone,4,3)%>" name="skip_user_exchange" onKeyUp="return autoTab(this, 3, event, true);" onchange="FlagFamilyChange();" size="3" maxlength="3">&ndash;
			<input type="text" value="<%=Right(sDayPhone,4)%>" name="skip_user_line" onKeyUp="return autoTab(this, 4, event, true);" onchange="FlagFamilyChange();" size="4" maxlength="4">
		</td></tr>
		<tr><td class="label" align="right" nowrap="nowrap">
			Cell Phone:
		</td><td>
			<input type="hidden" value="<%=sCell%>" name="egov_users_usercell">
			(<input type="text" value="<%=Left(sCell,3)%>" name="skip_cell_areacode" onKeyUp="return autoTab(this, 3, event, false);" size="3" maxlength="3">)&nbsp;
			<input type="text" value="<%=Mid(sCell,4,3)%>" name="skip_cell_exchange" onKeyUp="return autoTab(this, 3, event, false);" size="3" maxlength="3">&ndash;
			<input type="text" value="<%=Right(sCell,4)%>" name="skip_cell_line" onKeyUp="return autoTab(this, 4, event, false);" size="4" maxlength="4">
		</td></tr>
		<tr><td class="label" align="right" nowrap="nowrap">
			Fax:
		</td><td>
			<input type="hidden" value="<%=sFax%>" name="egov_users_userfax">
			(<input type="text" value="<%=Left(sFax,3)%>" name="skip_fax_areacode" onKeyUp="return autoTab(this, 3, event, false);" size="3" maxlength="3">)&nbsp;
			<input type="text" value="<%=Mid(sFax,4,3)%>" name="skip_fax_exchange" onKeyUp="return autoTab(this, 3, event, false);" size="3" maxlength="3">&ndash;
			<input type="text" value="<%=Right(sFax,4)%>" name="skip_fax_line" onKeyUp="return autoTab(this, 4, event, false);" size="4" maxlength="4">
		</td></tr>
<%		bHasResidentStreets = HasResidentTypeStreets( "R" )
		bFound = False 
		If bHasResidentStreets  Then 
			If Not OrgHasFeature( "large address list" ) Then 
				' Show all addresses for the city - short address solution
%>
				<tr><td class="label" align="right" nowrap="nowrap">
						Resident Address: 
					</td><td>
						<% DisplayAddresses  "R", sAddress, bFound %>
				</td></tr>
<%		
			Else
			' Show the large address list solution
%>
				<tr><td class="label" align="right" valign="top" nowrap="nowrap">
						Resident Address:
					</td><td>
<%						BreakOutAddress sAddress, sStreetNumber, sStreetName   ' In common.asp
						DisplayLargeAddressList "R", sStreetNumber, sStreetName, bFound %>&nbsp;
						<input type="button" class="button" value="Validate Address" onclick='checkAddress( "CheckResults", "no" );' />
						<!-- <br />- Or Other Not Listed - -->
				</td></tr>
<%
			End If 
		End If 
%>
		<tr><td class="label" align="right" nowrap="nowrap">
			<% If bHasResidentStreets Then %>
				Address(if not listed):
			<% Else %>
				Address:
			<% End If %>
		</td><td>
			<input type="text" value="<% If Not bFound Then 
											response.write sAddress
										 End If %>" name="egov_users_useraddress" onchange="FlagFamilyChange();" size="50" maxlength="100" />
		</td></tr>
		<tr><td class="label" align="right" nowrap="nowrap">
			Resident Unit:
		</td><td>
			<input type="text" value="<%=sUserUnit%>" name="egov_users_userunit" onchange="FlagFamilyChange();" size="11" maxlength="10" />
		<//td></tr>

<%		If OrgHasNeighborhoods( Session("orgid") ) Then %>
			<tr><td class=label align="right">
				Neighborhood:
			</td><td>
				<% DisplayNeighborhoods Session("orgid"), iNeighborhoodid %>
			</td></tr>
<%		End If %>

		<tr><td class="label" align="right" nowrap="nowrap">
			City:
		</td><td>
			<input type="text" value="<%=sCity%>" name="egov_users_usercity" onchange="FlagFamilyChange();" size="50" maxlength="100">
		</td></tr>
		<tr><td class="label" align="right" nowrap="nowrap">
			State / Province:
		</td><td>
			<input type="text" value="<%=sState%>" name="egov_users_userstate" onchange="FlagFamilyChange();" size="5" maxlength="10">
		</td></tr>
		<tr><td class="label" align="right" nowrap="nowrap">
			ZIP / Postal Code:
		</td><td>
			<input type="text" value="<%=sZip%>" name="egov_users_userzip" onchange="FlagFamilyChange();" size="10" maxlength="15">
		</td></tr>
		<tr><td class="label" align="right" nowrap="nowrap">
			Business Name:
		</td><td>
			<input type="text" value="<%=sBusinessName%>" name="egov_users_userbusinessname" size="50" maxlength="100">
		</td></tr>
<%		bHasBusinessStreets = HasResidentTypeStreets( "B" )
		bFound = False 
		If bHasBusinessStreets  Then %>
			<tr><td class="label" align="right" nowrap="nowrap">
					Business Street: 
				</td><td>
					<% DisplayAddresses  "B", sBusinessAddress, bFound %>
			</td></tr>
<%		End If %>
		<tr><td class="label" align="right" nowrap="nowrap">
			<% If bHasBusinessStreets Then %>
				Street (if not listed):
			<% Else %>
				Business Street:
			<% End If %>
		</td><td>
			<input type="text" value="<% If Not bFound Then 
											response.write sBusinessAddress
										 End If %>" name="egov_users_userbusinessaddress" size="50" maxlength="100">
		</td></tr>
		<tr><td class="label" align="right" nowrap="nowrap">
			Work Phone:
		</td><td>
			<!--<input type="text" value="<%=sWorkPhone%>" name="egov_users_userworkphone" size="50" maxlength="100">-->
			<input type="hidden" value="<%=sWorkPhone%>" id="egov_users_userworkphone" name="egov_users_userworkphone">
			(<input type="text" value="<%=Left(sWorkPhone,3)%>" name="skip_work_areacode" onKeyUp="return autoTab(this, 3, event, false);" size="3" maxlength="3" />)&nbsp;
			<input type="text" value="<%=Mid(sWorkPhone,4,3)%>" name="skip_work_exchange" onKeyUp="return autoTab(this, 3, event, false);" size="3" maxlength="3" />&ndash;
			<input type="text" value="<%=Mid(sWorkPhone,7,4)%>" name="skip_work_line" onKeyUp="return autoTab(this, 4, event, false);" size="4" maxlength="4" />&nbsp;
			ext. <input type="text" value="<%=Mid(sWorkPhone,11,4)%>" name="skip_work_ext" onKeyUp="return autoTab(this, 4, event, false);" size="4" maxlength="4" />
		</td></tr>

		<tr><td class="label" align="right" nowrap="nowrap">
			Emergency Contact:
		</td><td>
			<input type="text" value="<%=sEmergencyContact%>" name="egov_users_emergencycontact" style="width:300px;" maxlength="100">
		</td></tr>
		<tr><td class="label" align="right" nowrap="nowrap">
			Emergency Phone:
		</td><td>
			<input type="hidden" value="<%=sEmergencyPhone%>" id="egov_users_emergencyphone" name="egov_users_emergencyphone" />
			(<input type="text" value="<%=Left(sEmergencyPhone,3)%>" id="skip_emergencyphone_areacode" name="skip_emergencyphone_areacode" onKeyUp="return autoTab(this, 3, event, false);" size="3" maxlength="3" />)&nbsp;
			<input type="text" value="<%=Mid(sEmergencyPhone,4,3)%>" id="skip_emergencyphone_exchange" name="skip_emergencyphone_exchange" onKeyUp="return autoTab(this, 3, event, false);" size="3" maxlength="3" />&ndash;
			<input type="text" value="<%=Mid(sEmergencyPhone,7,4)%>" id="skip_emergencyphone_line" name="skip_emergencyphone_line" size="4" maxlength="4" />
		</td></tr>	

		</table>
		<div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:FamilyList('<%=iFamilyId%>');"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:doCheck();"><%=sSaveLabel%></a></div>

		</form>
  </td>

    </tr>
 </table>


 </div>
 </div>

<!--#Include file="../admin_footer.asp"-->  
  

</body>
</html>

<!--#Include file="inc_dbfunction.asp"-->

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' FUNCTION DisplayResidentAddresses( sAddress )
'--------------------------------------------------------------------------------------------------
Sub  DisplayAddresses( ByVal sResidenttype, ByVal sAddress, ByRef bFound )
	Dim sSql, oRs 

	sSql = "SELECT residentstreetnumber, residentstreetname FROM egov_residentaddresses_list where orgid=" & session("orgid") & " and residenttype='" & sResidenttype & "' order by sortstreetname, residentstreetprefix, Cast(residentstreetnumber as int)"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN") , 3, 1

	response.write "<select name=""skip_" & sResidenttype & "address"" onchange=""FlagFamilyChange();"">"	
	response.write "<option value=""0000"">Please select an address...</option>"
		
	Do While Not oRs.EOF 
		response.write vbcrlf & "<option value=""" &  oRs("residentstreetnumber") & " " & oRs("residentstreetname")  & """"
		If UCase(sAddress) = UCase(oRs("residentstreetnumber") & " " & oRs("residentstreetname")) Then 
			response.write " selected=""selected"" "
			bFound = True 
		End If 
		response.write ">" & oRs("residentstreetnumber") & " " & oRs("residentstreetname") & "</option>"
		oRs.MoveNext
	Loop

	response.write "</select>"

	oRs.close
	Set oRs = Nothing 
	
End Sub  

'--------------------------------------------------------------------------------------------------
' boolean HasResidentTypeStreets( sResidenttype )
'--------------------------------------------------------------------------------------------------
Function HasResidentTypeStreets( ByVal sResidenttype )
	Dim sSql, oRs

	sSql = "SELECT COUNT(residentaddressid) AS hits FROM egov_residentaddresses WHERE orgid = " & session("orgid") & " AND residenttype = '" & sResidenttype & "'"
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN") , 3, 1

	If CLng(oRs("hits")) > 0 Then
		HasResidentTypeStreets = True 
	Else
		HasResidentTypeStreets = False 
	End if
	
	oRs.Close
	Set oRs = Nothing
	
End Function 

'--------------------------------------------------------------------------------------------------
' FUNCTION HasResidentTypes( )
'--------------------------------------------------------------------------------------------------
Function HasResidentTypes()
	Dim sSql, oRs

	sSql = "SELECT count(resident_type) as hits FROM egov_poolpassresidenttypes where orgid = " & session("orgid") & ""
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN") , 3, 1

	If CLng(oRs("hits")) > 0 Then
		HasResidentTypes = True 
	Else
		HasResidentTypes = False 
	End if
	
	oRs.close
	Set oRs = Nothing 

End Function 

'--------------------------------------------------------------------------------------------------
' string DisplayResidentTypes( sResidentType )
'--------------------------------------------------------------------------------------------------
Function DisplayResidentTypes( ByVal sResidentType )
	Dim sSql, oRs

	sSql = "SELECT resident_type, description FROM egov_poolpassresidenttypes where orgid=" & session("orgid") & " order by displayorder"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN") , 3, 1

	DisplayResidentTypes = "<select name=""skip_egov_users_residenttype"">"	
	'DisplayResidentTypes = DisplayResidentTypes &  "<option value=""0"">Please select a resident type...</option>"
		
	Do While Not oRs.EOF 
		DisplayResidentTypes = DisplayResidentTypes & vbcrlf & "<option value=""" &  oRs("resident_type") & """"
		If sResidentType = oRs("resident_type") Then
			DisplayResidentTypes = DisplayResidentTypes & " selected=""selected"" "
		End If 
		DisplayResidentTypes = DisplayResidentTypes & ">" & oRs("description") & "</option>"
		oRs.MoveNext
	Loop

	DisplayResidentTypes = DisplayResidentTypes & "</select>"

	oRs.Close
	Set oRs = Nothing 
	
End Function 


'--------------------------------------------------------------------------------------------------
' void GetNewFamilyValues iRootUserId 
'--------------------------------------------------------------------------------------------------
Sub GetNewFamilyValues( ByVal iRootUserId )
	Dim sSql, oRs

	sSql = "SELECT * FROM egov_users WHERE userid = " & iRootUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sFirstName = ""
		sLastName = oRs("userlname")
		sAddress = oRs("useraddress")
		sState = oRs("userstate")
		sCity = oRs("usercity")
		sZip = oRs("userzip")
		sFax = oRs("userfax")
		sCell = oRs("usercell")
		sBusinessName = ""
		sPassword = ""
		sDayPhone = oRs("userhomephone")
		sWorkPhone = oRs("userworkphone")
		sEmergencyContact = oRs("emergencycontact")
		sEmergencyPhone = oRs("emergencyphone")
		sBirthdate = ""
		iRelationshipId = 0
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
		If oRs("residencyverified") Then 
			sResidencyVerified = " checked=""checked"" "
		Else
			sResidencyVerified = ""
		End If 
		sBusinessAddress = ""
		sUserUnit = oRs("userunit")
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void GetRegisteredUserValues iUserId 
'--------------------------------------------------------------------------------------------------
Sub GetRegisteredUserValues( ByVal iUserId )
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
		sBirthdate= oRs("birthdate")
		iRelationshipId = oRs("relationshipid")
		If IsNull(oRs("gender")) Then
			sGender = "N"
		Else
			sGender = oRs("gender")
		End If 
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
		If oRs("residencyverified") Then 
			sResidencyVerified = " checked=""checked"" "
		Else
			sResidencyVerified = ""
		End If 
		sBusinessAddress = oRs("userbusinessaddress")
		sUserUnit = oRs("userunit")
	End If

	oRs.close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void UpdateRecords
'--------------------------------------------------------------------------------------------------
Sub UpdateRecords()
	Dim sSql, iResidencyVerified, iNeighborhoodid, sSql2, sBirthdate, iRelationshipId, sGender

	If request("egov_users_residencyverified") = "on" Then
		iResidencyVerified = "1"
	Else
		iResidencyVerified = "0"
	End If 
	If request("egov_users_neighborhoodid") <> "" Then
		iNeighborhoodid = request("egov_users_neighborhoodid")
	Else
		iNeighborhoodid = "0"
	End If 
	If request("egov_users_birthdate") <> "" Then
		sBirthdate = "'" & request("egov_users_birthdate") & "'"
	Else
		sBirthdate = "NULL"
	End If 

	If request("egov_users_gender") <> "M" And request("egov_users_gender") <> "F" Then
		sGender = "NULL"
	Else
		sGender = "'" & DBsafe(request("egov_users_gender")) & "'"
	End If 

	sSql = "UPDATE egov_users SET userfname = '" & DBsafe(request("egov_users_userfname"))
	sSql = sSql & "', userlname = '" &  DBsafe(request("egov_users_userlname"))
	sSql = sSql & "', gender = " & sGender
	sSql = sSql & ", useraddress = '" &  DBSafe(request("egov_users_useraddress"))
	sSql = sSql & "', usercity = '" & DBSafe(request("egov_users_usercity"))
	sSql = sSql & "', userstate = '" & DBSafe(request("egov_users_userstate"))
	sSql = sSql & "', userzip = '" & request("egov_users_userzip")
	'sSql = sSql & "', useremail = '" & request("egov_users_useremail")
	sSql = sSql & "', userbusinessname = '" & DBsafe( request("egov_users_userbusinessname") )
	'sSql = sSql & "', userpassword = '" & request("egov_users_userpassword")
	sSql = sSql & "', userhomephone = '" & request("egov_users_userhomephone")
	sSql = sSql & "', userworkphone = '" & request("egov_users_userworkphone")
	sSql = sSql & "', userfax = '" & request("egov_users_userfax")
	sSql = sSql & "', usercell = '" & request("egov_users_usercell")
	sSql = sSql & "', residenttype = '" & request("egov_users_residenttype")
	sSql = sSql & "', userbusinessaddress = '" & DBSafe(request("egov_users_userbusinessaddress"))
	sSql = sSql & "', emergencycontact = '" & DBSafe(request("egov_users_emergencycontact"))
	sSql = sSql & "', emergencyphone = '" & request("egov_users_emergencyphone")
	sSql = sSql & "', neighborhoodid = " & iNeighborhoodid
	sSql = sSql & ",  birthdate = " & sBirthdate
	sSql = sSql & ",  relationshipid = " & request("egov_users_relationshipid")
	sSql = sSql & ",  residencyverified = " & iResidencyVerified
	sSql = sSql & " WHERE userid = " & request("userid") & ""

	'response.write sSql
	RunSQLStatement sSql

	If request("familyaddresschanged") = "YES" Then
		sSql = "Update egov_users Set userhomephone = '" & request("egov_users_userhomephone")
		sSql = sSql & "', useraddress = '" &  DBsafe(request("egov_users_useraddress"))
		sSql = sSql & "', usercity = '" & DBsafe(request("egov_users_usercity"))
		sSql = sSql & "', userstate = '" & DBsafe(request("egov_users_userstate"))
		sSql = sSql & "', userzip = '" & DBsafe(request("egov_users_userzip"))
		sSql = sSql & "',  userunit = '" & dbsafe(request("egov_users_userunit")) 
		sSql = sSql & "' WHERE familyid = " & request("egov_users_familyid")
		RunSQLStatement sSql
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer InsertRecord( )
'--------------------------------------------------------------------------------------------------
Function InsertRecord( )
	Dim iResidencyVerified, iFamilyId, oFamily, sAction, bAddressChanged, sSql, sUserfname, sUserlname
	Dim sUseraddress, sUsercity, sUserstate, sUserzip, sUserhomephone, sUsercell, sUserworkphone
	Dim sUserfax, sUserbusinessname, sUserbusinessaddress, iNeighborhoodid,sEmergencycontact
	Dim sEmergencyphone, sBirthDate, iRelationshipId, sResidentType, sResidencyVerified, sGender

	bAddressChanged = False 

	iFamilyId = request("egov_users_familyid")

	sUserfname = "'" & dbsafe(request("egov_users_userfname")) & "'"

	sUserlname = "'" & dbsafe(request("egov_users_userlname")) & "'"

	If request("egov_users_gender") <> "M" And request("egov_users_gender") <> "F" Then
		sGender = "NULL"
	Else
		sGender = "'" & dbsafe(request("egov_users_gender")) & "'"
	End If 

	If request("egov_users_useraddress") <> "" Then 
		sUseraddress = "'" & dbsafe(request("egov_users_useraddress")) & "'"
	Else
		sUseraddress = "NULL"
	End If 

	If request("egov_users_usercity") <> "" Then
		sUsercity = "'" & dbsafe(request("egov_users_usercity")) & "'"
	Else
		sUsercity = "NULL"
	End If 

	If request("egov_users_userstate") <> "" Then
		sUserstate = UCase(dbsafe(request("egov_users_userstate")))
		sUserstate = "'" & sUserstate & "'"
	Else
		sUserstate = "NULL"
	End If 

	If request("egov_users_userzip") <> "" Then 
		sUserzip = "'" & dbsafe(request("egov_users_userzip")) & "'"
	Else
		sUserzip = "NULL"
	End If 

	If request("egov_users_userhomephone") <> "" Then
		sUserhomephone = "'" & dbsafe(request("egov_users_userhomephone")) & "'"
	Else
		sUserhomephone = "NULL"
	End If 

	If request("egov_users_usercell") <> "" Then
		sUsercell = "'" & dbsafe(request("egov_users_usercell")) & "'"
	Else
		sUsercell = "NULL"
	End If 

	If request("egov_users_userworkphone") <> "" Then
		sUserworkphone = "'" & dbsafe(request("egov_users_userworkphone")) & "'"
	Else
		sUserworkphone = "NULL"
	End If 

	If request("egov_users_userfax") <> "" Then
		sUserfax = "'" & dbsafe(request("egov_users_userfax")) & "'"
	Else
		sUserfax = "NULL"
	End If 

	If request("egov_users_userbusinessname") <> "" Then
		sUserbusinessname = "'" & dbsafe(request("egov_users_userbusinessname")) & "'"
	Else
		sUserbusinessname = "NULL"
	End If 

	If request("egov_users_userbusinessaddress") <> "" Then 
		sUserbusinessaddress = "'" & dbsafe(request("egov_users_userbusinessaddress")) & "'"
	Else
		sUserbusinessaddress = "NULL"
	End If 

	If request("egov_users_neighborhoodid") <> "" Then
		iNeighborhoodid = request("egov_users_neighborhoodid")
	Else
		iNeighborhoodid = "0"
	End If 

	If request("egov_users_emergencycontact") <> "" Then
		sEmergencycontact = "'" & dbsafe(request("egov_users_emergencycontact")) & "'"
	Else
		sEmergencycontact = "NULL"
	End If

	If request("egov_users_emergencyphone") <> "" Then
		sEmergencyphone = "'" & dbsafe(request("egov_users_emergencyphone")) & "'"
	Else
		sEmergencyphone = "NULL"
	End If 

	If request("egov_users_birthdate") <> "" Then
		sBirthDate = "'" & dbsafe(request("egov_users_birthdate")) & "'"
	Else
		sBirthDate = "NULL"
	End If 

	iRelationshipId = CLng(request("egov_users_relationshipid"))

	If request("egov_users_residenttype") <> "" Then 
		sResidentType = "'" & dbsafe(request("egov_users_residenttype")) & "'"
	Else
		sResidentType = "NULL"
	End If 

	If request("egov_users_residencyverified") = "on" Then 
		sResidencyVerified = 1	' should be 0 or 1 only
	Else
		sResidencyVerified = 0
	End If 


	sSql = "INSERT INTO egov_users ( userfname, userlname, useraddress ,usercity, userstate, userzip, userhomephone, "
	sSql = sSql & "userworkphone, userbusinessname, orgid, userregistered, userbusinessaddress, emergencycontact, "
	sSql = sSql & "emergencyphone, neighborhoodid, birthdate, relationshipid, residencyverified, familyid, "
	sSql = sSql & "residenttype, usercell, gender ) VALUES ( "
	sSql = sSql & sUserfname & ", "
	sSql = sSql & sUserlname & ", "
	sSql = sSql & sUseraddress & ", "
	sSql = sSql & sUsercity & ", "
	sSql = sSql & sUserstate & ", "
	sSql = sSql & sUserzip & ", "
	sSql = sSql & sUserhomephone & ", "
	sSql = sSql & sUserworkphone & ", "
	sSql = sSql & sUserbusinessname & ", "
	sSql = sSql & session("orgid") & ", "
	sSql = sSql & "1, "
	sSql = sSql & sUserbusinessaddress & ", "
	sSql = sSql & sEmergencycontact & ", "
	sSql = sSql & sEmergencyphone & ", "
	sSql = sSql & iNeighborhoodid & ", "
	sSql = sSql & sBirthDate & ", "
	sSql = sSql & iRelationshipId & ", "
	sSql = sSql & sResidencyVerified & ", "
	sSql = sSql & iFamilyId & ", "
	sSql = sSql & sResidentType & ", "
	sSql = sSql & sUsercell & ", "
	sSql = sSql & sGender
	sSql = sSql & " )"

	InsertRecord = RunInsertStatement( sSql )	' in common.asp


	' Parameters for the stored Proc
	'@orgid int,
	'@firstname varchar(25),
	'@lastname varchar(25),
	'@businessname  varchar(50) = NULL,
	'@address1  varchar(250) = NULL,
	'@homenumber varchar(20),
	'@cellnumber varchar(20),
	'@worknumber varchar(20) = NULL,
	'@city varchar(20) = NULL,
	'@state varchar(20) = NULL,
	'@zip varchar(20) = NULL,
	'@faxnumber varchar(20) = NULL ,
	'@businessaddress varchar(255) = NULL,
	'@emergencycontact varchar(100) = NULL,
	'@emergencyphone varchar(50) = NULL,
	'@neighborhoodid int = NULL,
	'@birthdate datetime = NULL,
	'@relationshipid int = NULL, 
	'@residencyverified bit,
	'@residenttype char(1) = NULL,
	'@familyid int,
	'@userid int OUTPUT

'	If request("egov_users_residencyverified") = "on" Then
'		iResidencyVerified = "1"
'	Else
'		iResidencyVerified = "0"
'	End If 
'	If request("egov_users_neighborhoodid") <> "" Then
'		iNeighborhoodid = request("egov_users_neighborhoodid")
'	Else
'		iNeighborhoodid = "0"
'	End If 

'	Set oCmd = Server.CreateObject("ADODB.Command")
'	With oCmd
'		.ActiveConnection = Application("DSN")
'		.CommandText = "NewCitizenFamilyMember"
'		.CommandType = 4
'		.Parameters.Append oCmd.CreateParameter("@orgid", 3, 1, 4, Session("orgid"))
'		.Parameters.Append oCmd.CreateParameter("@firstname", 200, 1, 25,  DBsafe(request("egov_users_userfname")))
'		.Parameters.Append oCmd.CreateParameter("@lastname", 200, 1, 25, DBsafe(request("egov_users_userlname")))
'		If request("egov_users_userbusinessname") <> "" Then 
'			.Parameters.Append oCmd.CreateParameter("@businessname", 200, 1, 25, DBsafe( request("egov_users_userbusinessname") ))
'		Else
'			.Parameters.Append oCmd.CreateParameter("@businessname", 200, 1, 25, NULL)
'		End If 
'		If request("egov_users_useraddress") <> "" Then 
'			.Parameters.Append oCmd.CreateParameter("@address1", 200, 1, 250,  dbsafe(request("egov_users_useraddress")))
'		Else
'			.Parameters.Append oCmd.CreateParameter("@address1", 200, 1, 250, NULL)
'		End If 
'		.Parameters.Append oCmd.CreateParameter("@homenumber", 200, 1, 20, request("egov_users_userhomephone"))
'		If request("egov_users_usercell") <> "" Then 
'			.Parameters.Append oCmd.CreateParameter("@cellnumber", 200, 1, 20, request("egov_users_usercell"))
'		Else
'			.Parameters.Append oCmd.CreateParameter("@cellnumber", 200, 1, 20, NULL)
'		End If 
'		If request("egov_users_userworkphone") <> "" Then 
'			.Parameters.Append oCmd.CreateParameter("@worknumber", 200, 1, 20, request("egov_users_userworkphone"))
'		Else
'			.Parameters.Append oCmd.CreateParameter("@worknumber", 200, 1, 20, NULL)
'		End If 
'		If request("egov_users_usercity") <> "" Then 
'			.Parameters.Append oCmd.CreateParameter("@city", 200, 1, 20, dbsafe(request("egov_users_usercity")))
'		Else
'			.Parameters.Append oCmd.CreateParameter("@city", 200, 1, 20, NULL)
'		End If 
'		If request("egov_users_userstate") <> "" Then 
'			.Parameters.Append oCmd.CreateParameter("@state", 200, 1, 20, request("egov_users_userstate"))
'		Else
'			.Parameters.Append oCmd.CreateParameter("@state", 200, 1, 20, NULL)
'		End If
'		If request("egov_users_userzip") <> "" Then 
'			.Parameters.Append oCmd.CreateParameter("@zip", 200, 1, 20, request("egov_users_userzip"))
'		Else
'			.Parameters.Append oCmd.CreateParameter("@zip", 200, 1, 20, NULL)
'		End If
'		If request("egov_users_userfax") <> "" Then 
'			.Parameters.Append oCmd.CreateParameter("@faxnumber", 200, 1, 20, request("egov_users_userfax"))
'		Else
'			.Parameters.Append oCmd.CreateParameter("@faxnumber", 200, 1, 20, NULL)
'		End If
'		If request("egov_users_userbusinessaddress") <> "" Then 
'			.Parameters.Append oCmd.CreateParameter("@businessaddress", 200, 1, 255, DBSafe(request("egov_users_userbusinessaddress")))
'		Else
'			.Parameters.Append oCmd.CreateParameter("@businessaddress", 200, 1, 255, NULL)
'		End If
'		If request("egov_users_emergencycontact") <> "" Then 
'			.Parameters.Append oCmd.CreateParameter("@emergencycontact", 200, 1, 100, DBSafe(request("egov_users_emergencycontact")))
'		Else
'			.Parameters.Append oCmd.CreateParameter("@emergencycontact", 200, 1, 100, NULL)
'		End If
'		If request("egov_users_emergencyphone") <> "" Then 
'			.Parameters.Append oCmd.CreateParameter("@emergencyphone", 200, 1, 50, request("egov_users_emergencyphone"))
'		Else
'			.Parameters.Append oCmd.CreateParameter("@emergencyphone", 200, 1, 50, NULL)
'		End If
'		If CLng(iNeighborhoodid) <> CLng(0) Then 
'			.Parameters.Append oCmd.CreateParameter("@neighborhoodid", 3, 1, 4, iNeighborhoodid)
'		Else
'			.Parameters.Append oCmd.CreateParameter("@neighborhoodid", 3, 1, 4, NULL)
'		End If
'		If Trim(request("egov_users_birthdate")) <> "" Then 
'			.Parameters.Append oCmd.CreateParameter("@birthdate", 135, 1, 16, request("egov_users_birthdate"))
'		Else
'			.Parameters.Append oCmd.CreateParameter("@birthdate", 135, 1, 16, NULL)
'		End If
'		.Parameters.Append oCmd.CreateParameter("@relationshipid", 3, 1, 4, request("egov_users_relationshipid"))
'		.Parameters.Append oCmd.CreateParameter("@residencyverified", 11, 1, 1, iResidencyVerified)
'		If Trim(request("egov_users_residenttype")) <> "" Then 
'			.Parameters.Append oCmd.CreateParameter("@residenttype", 129, 1, 1, request("egov_users_residenttype"))
'		Else
'			.Parameters.Append oCmd.CreateParameter("@residenttype", 129, 1, 1, NULL)
'		End If
'		.Parameters.Append oCmd.CreateParameter("@familyid", 3, 1, 4, request("egov_users_familyid"))
'		.Parameters.Append oCmd.CreateParameter("@userid", 3, 2, 4)
'		.Execute
'	End With

'	iUserId = oCmd.Parameters("@userid").Value

'	Set oCmd = Nothing

	' Send back the new userid
'	InsertRecord = iUserId

End Function 


'--------------------------------------------------------------------------------------------------
' string FormatPhone( Number )
'--------------------------------------------------------------------------------------------------
Function FormatPhone( ByVal Number )

	If Len(Number) = 10 Then
		FormatPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
	Else
		FormatPhone = Number
	End If

End Function


'--------------------------------------------------------------------------------------------------
' string GetDefaultPhone( iOrgId )
'--------------------------------------------------------------------------------------------------
Function GetDefaultPhone( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT defaultphone FROM organizations WHERE orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetDefaultPhone = oRs("defaultphone")
	Else 
		GetDefaultPhone = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string GetDefaultEmail( iOrgId )
'--------------------------------------------------------------------------------------------------
Function GetDefaultEmail( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT defaultemail FROM organizations WHERE orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetDefaultEmail = oRs("defaultemail")
	Else 
		GetDefaultEmail = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string GetOrgName( iOrgId )
'--------------------------------------------------------------------------------------------------
Function GetOrgName( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT orgname FROM organizations WHERE orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetOrgName = oRs("orgname")
	Else 
		GetOrgName = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetVirtualName( iorgid )
'--------------------------------------------------------------------------------------------------
'Function GetVirtualName( ByVal iorgid )
'	Dim sReturnValue, sSql, oRst
'
'	sReturnValue = "UNKNOWN"
'
'	Set oRst = Server.CreateObject("ADODB.Recordset")
'	'sSql = "SELECT OrgVirtualSiteName FROM Organizations WHere orgid='" &  iorgid & "'"
'	sSql = "SELECT orgegovwebsiteurl FROM Organizations WHere orgid='" &  iorgid & "'"
'	oRst.open sSql,Application("DSN"),0,1
'
'	If NOT oRst.EOF THEN
'		sReturnValue = Trim(oRst("orgegovwebsiteurl"))
'		'response.write Trim(oRst("orgegovwebsiteurl")) & "&nbsp;" & Len(sReturnValue) & "&nbsp;" & InstrRev(sReturnValue,"/")
'		'sReturnValue = Mid(sReturnValue,1,(InstrRev(sReturnValue,"/")-1))
'	END If
'	oRst.close
'	Set oRst = Nothing 
'
'	GetVirtualName = sReturnValue
'End Function


'--------------------------------------------------------------------------------------------------
' string DBsafe( strDB )
'--------------------------------------------------------------------------------------------------
Function DBsafe( ByVal strDB )

  If Not VarType( strDB ) = 8 Then DBsafe = strDB : Exit Function

  DBsafe = Replace( strDB, "'", "''" )

End Function


'--------------------------------------------------------------------------------------------------
' void DisplayRelationships iOrgid, iRelationshipId 
'--------------------------------------------------------------------------------------------------
Sub DisplayRelationships( ByVal iOrgid, ByVal iRelationshipId )
	Dim sSql, oRs 

	sSql = "SELECT relationshipid, relationship FROM egov_familymember_relationships WHERE orgid = " & iorgid & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""egov_users_relationshipid"" id=""relationship"">"	
		
	Do While Not oRs.EOF 
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

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetRelationShip = oRs("relationship") 
	Else
		GetRelationShip = "" 
	End if
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetFamilyId( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetFamilyId( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT familyid FROM egov_users WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetFamilyId = oRs("familyid")
	Else
		GetFamilyId = iUserId
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void UpdateOverflowFields iUserId, sUserUnit 
'--------------------------------------------------------------------------------------------------
Sub UpdateOverflowFields( ByVal iUserId, ByVal sUserUnit )
	' This handles overflow fields for the egov_users table
	Dim sSql

	sSql = "UPDATE egov_users SET userunit = '" & DBsafe( sUserUnit ) & "' WHERE userid = " & iUserId

	RunSQLStatement sSql

End Sub


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
	sSql = sSql & " FROM egov_residentaddresses where orgid = " & session( "orgid" ) & " and residenttype = '" & sResidenttype & "' "
	sSql = sSql & "and residentstreetname is not null order by sortstreetname, residentstreetprefix, residentstreetname"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<input type=""text"" name=""residentstreetnumber"" value=""" & sStreetNumber & """ onchange=""FlagFamilyChange();"" size=""8"" maxlength=""10"" /> &nbsp; "
		response.write vbcrlf & "<select name=""skip_address"" onchange=""FlagFamilyChange();"">"
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
' Sub DisplayLargeAddressList( sResidenttype, sStreetNumber, sStreetName, bFound )
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
	sSql = sSql & " FROM egov_residentaddresses WHERE orgid = " & session( "orgid" ) & " AND residenttype = '" & sResidenttype & "' "
	sSql = sSql & "AND residentstreetname IS NOT NULL ORDER BY sortstreetname"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<input type=""text"" name=""residentstreetnumber"" value=""" & sStreetNumber & """ onchange=""FlagFamilyChange();"" size=""8"" maxlength=""10"" /> &nbsp; "
		response.write vbcrlf & "<select name=""skip_address"" onchange=""FlagFamilyChange();"">"
		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown...</option>"
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
