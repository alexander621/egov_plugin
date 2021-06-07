<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: action_.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Action Line Search Results.
'
' MODIFICATION HISTORY
' ?.?	05/08/07	 Steve Loar - Changes to problem location to handle larger cities
' 6.2	07/16/07 	Steve Loar - Changes to handle email/request type problems.
' 6.3 11/16/07  David Boyer - Now validates contact email address to ensure it is entered in proper format.
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sError 
' CAPTURE CURRENT PATH
Session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()
Session("RedirectLang") = "Return to Action Line"

Dim oActionOrg

Set oActionOrg = New classOrganization

'Show/Hide hidden fields.  To Hide = "HIDDEN", To Show = "TEXT"
lcl_hidden = "hidden"
%>
<html>
<head>

	<%If iorgid = 7 Then %>
		<title><%=sOrgName%></title>
	<%Else%>
		<title>E-Gov Services <%=sOrgName%></title>
	<%End If%>

 	<link rel="stylesheet" type="text/css" href="css/styles.css" />
 	<link rel="stylesheet" type="text/css" href="global.css" />
 	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

 	<script language="Javascript" src="scripts/modules.js"></script>
 	<script language="Javascript" src="scripts/easyform.js"></script>
  <script language="JavaScript" src="scripts/ajaxLib.js"></script>
  <script language="JavaScript" src="scripts/removespaces.js"></script>
  <script language="JavaScript" src="scripts/setfocus.js"></script>
<script language="Javascript">
function autoTab(input,len, e) {
  var keyCode = (isNN) ? e.which : e.keyCode; 
		var filter = (isNN) ? [0,8,9] : [0,8,9,16,17,18,37,38,39,40,46];

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

function getIndex(input) {
		var index = -1, i = 0, found = false;

		while (i < input.form.length && index == -1)
			if (input.form[i] == input)index = i;
			else i++;
				return index;
		}
		return true;
	}


		function ShowMap()
		{
			//alert(document.frmRequestAction.skip_address.value);
			if (document.frmRequestAction.skip_address.value == 0)
			{
				alert('Please Select an address from the list.');
				document.frmRequestAction.skip_address.focus();
				return;
			}
					eval('window.open("<%=Application("MAP_URL")%>action_line_map.asp?residentaddressid=' + document.frmRequestAction.skip_address.value + '", "_map", "width=700,height=500,toolbar=0,statusbar=0,resizable,scrollbars=1,menubar=0,left=0,top=0")');
		}

		function openWin2(url, name) 
		{
		  popupWin = window.open(url, name,"resizable,width=500,height=450");
		}

		// Validate Tracking added by Steve Loar - 12/30/2005
		function ValidateTracking( form )
		{
			//return true;
			var rege = /^\d+$/;
			var Ok = rege.exec(form.REQUEST_ID.value);

			if (! Ok)
			{
				alert ("Tracking Numbers must be numeric. Please try your search again.");
				form.REQUEST_ID.focus();
				form.REQUEST_ID.select();
				return false;
			}
			return true;
		}

		function ValidateInput()
		{
			if (document.frmRequestAction.subjecttext.value != '')
			{
				alert("Please remove any input from the Internal Only field at the bottom of the form.");
				document.frmRequestAction.subjecttext.focus();
				return false;
			}
			//alert(document.frmRequestAction.cot_txtDaytime_Phone.value);
			// Set the Phone number
			var Phone_exists = eval(document.frmRequestAction["cot_txtDaytime_Phone"]);
			if(Phone_exists)
			{
				document.frmRequestAction.cot_txtDaytime_Phone.value = document.frmRequestAction.skip_user_areacode.value + document.frmRequestAction.skip_user_exchange.value + document.frmRequestAction.skip_user_line.value;
			}
			// Set the Fax
			var Fexists = eval(document.frmRequestAction["cot_txtFax"]);
			if(Fexists)
			{
				document.frmRequestAction.cot_txtFax.value = document.frmRequestAction.skip_fax_areacode.value + document.frmRequestAction.skip_fax_exchange.value + document.frmRequestAction.skip_fax_line.value;
			}
			// alert(document.frmRequestAction.cot_txtDaytime_Phone.value);
			return  validateForm('frmRequestAction');
			//return true;
		}

		var isNN = (navigator.appName.indexOf("Netscape")!=-1);

		function getinfo()
		{
			if (document.frmRequestAction.chkSameAs.checked) 
			{
				// CHECK USE ABOVE
				document.frmRequestAction.ques_issue3.value = document.frmRequestAction.cot_txtCity.value;
				document.frmRequestAction.ques_issue4.value = document.frmRequestAction.cot_txtState_vSlash_Province.value;
				document.frmRequestAction.ques_issue5.value = document.frmRequestAction.cot_txtZIP_vSlash_Postal_Code.value;
			}


			else 
			{
				// UNCHECKED CLEAR VALUES
				document.frmRequestAction.ques_issue3.value = '';
				document.frmRequestAction.ques_issue4.value = '';
				document.frmRequestAction.ques_issue5.value = '';
			}
		}
 	var winHandle;
 	var w = (screen.width - 640)/2;
 	var h = (screen.height - 450)/2;

<% if OrgHasFeature(iorgid, "large address list") then %>
function checkAddress( sReturnFunction, sSave ) {
		// Remove any extra spaces
		document.frmRequestAction.residentstreetnumber.value = removeSpaces(document.frmRequestAction.residentstreetnumber.value);

		// check the number for non-numeric values
		var rege = /^\d+$/;
		var Ok = rege.exec(document.frmRequestAction.residentstreetnumber.value);

  if(document.frmRequestAction.ques_issue2.value=="") {
		   if ( ! Ok ) {
    			alert("The Resident Street Number cannot be blank and must be numeric.");
	 	   	setfocus(document.frmRequestAction.residentstreetnumber);
   		 	return false;
   		}

   		// check that they picked a street name
   		if ( document.frmRequestAction.skip_address.value == '0000') {
 		   	alert("Please select a street name from the list first.");
   	 		setfocus(document.frmRequestAction.skip_address);
   		 	return false;
   		}

   		// This is here because window.open in the Ajax callback routine will not work
   		winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&stnumber=' + document.frmRequestAction.residentstreetnumber.value + '&stname=' + document.frmRequestAction.skip_address.value + '&sCheckType=' + sReturnFunction + '&formname=frmRequestAction", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
   		//self.focus();
   		// Fire off Ajax routine
   		doAjax('checkaddress.asp', 'stnumber=' + document.frmRequestAction.residentstreetnumber.value + '&stname=' + document.frmRequestAction.skip_address.value + '&orgid=<%=iorgid%>', sReturnFunction, 'get', '0');
  }else{
     if(document.frmRequestAction.residentstreetnumber.value!="" || document.frmRequestAction.skip_address.value!="0000") {
        document.frmRequestAction.ques_issue2.value="";
      		winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&stnumber=' + document.frmRequestAction.residentstreetnumber.value + '&stname=' + document.frmRequestAction.skip_address.value + '&sCheckType=' + sReturnFunction + '&formname=frmRequestAction", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
      		doAjax('checkaddress.asp', 'stnumber=' + document.frmRequestAction.residentstreetnumber.value + '&stname=' + document.frmRequestAction.skip_address.value + '&orgid=<%=iorgid%>', sReturnFunction, 'get', '0');
     }else{
        if(sReturnFunction!="FinalCheck") {
//      		alert("A non-listed address has been entered.");
           document.frmRequestAction.validstreet.value = 'N';
        }else{
           document.frmRequestAction.validstreet.value = 'N';
           FinalCheck('FOUND');
        }
     }
  }
}

function CheckResults( sResults ) {
  // Process the Ajax CallBack when the validate address button is clicked
  if (sResults == 'FOUND') {
    		if(winHandle != null && ! winHandle.closed) { 
   	  			winHandle.close();
   			}
	 	  	document.frmRequestAction.ques_issue2.value = '';
      document.frmRequestAction.validstreet.value = 'Y';
  		 	alert("This is a valid address in the system.");
  }else{
      document.frmRequestAction.validstreet.value = 'N';
 	  		winHandle.focus();
  }
}

function FinalCheck( sResults ) {
  if (sResults == 'FOUND') {
    		if(winHandle != null && ! winHandle.closed) { 
   	  			winHandle.close();
   			}
      if(document.frmRequestAction.validstreet.value == '') {
         document.frmRequestAction.validstreet.value = 'Y';
      }

      if(ValidateInput()) {
         isemailentered();
      }
  }else{
      document.frmRequestAction.validstreet.value = 'N';
 	  		winHandle.focus();
  }
}
<% end if %>

function validateCheckbox(p_field) {
  //get the total options
  var lcl_total_options = document.getElementById("total_options_"+p_field).innerHTML;
  var lcl_total_checked = 0;

  for (i = 1; i <= lcl_total_options-1; i++) {
       if(document.getElementById(p_field+"_"+i).checked==true) {
          lcl_total_checked = lcl_total_checked + 1;
       }
  }
  if(lcl_total_checked == 0) {
     document.getElementById(p_field+"_"+lcl_total_options).checked=true;
  }else{
     document.getElementById(p_field+"_"+lcl_total_options).checked=false;
  }
}
</script>
</head>
<!--#Include file="include_top.asp"-->

<!--BODY CONTENT-->
<div id="content">
 	<div id="centercontent">
<%
 'Set up the values
  lcl_feature        = "Create Action Line Requests"
  lcl_start_datetime = "12-14-07 04:00 PM EST"
  lcl_end_datetime   = "12-14-07 05:00 PM EST"
%>
<table border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td align="center">
          The feature "<font color="#800000" style="font-size: 11px"><strong><%=lcl_feature%></strong></font>" will be unavailable between<br>
          <font color="#800000" style="font-size: 11px"><strong><%=lcl_start_datetime%></strong></font> to 
          <font color="#800000" style="font-size: 11px"><strong><%=lcl_end_datetime%></strong></font><br>
          due to scheduled maintenance outage.
 	    </td>
  </tr>
</table>
<p>
  </div>
</div>

<!--#Include file="include_bottom.asp"-->