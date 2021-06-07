<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->
<!-- #include file="correction_global_functions.asp" //-->
<% 
'------------------------------------------------------------------------------
' FILENAME: CORRECTION_ISSUE_LOCATION.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/5/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  
'
' MODIFICATION HISTORY
' 1.0	 02/05/07	 John Stullenberger - INITIAL VERSION
' 2.0  10/30/07  David Boyer - Added "Valid Address" and "Large Address List" features
' 2.1  10/03/08  David Boyer - Added "Import Address Fields" button
' 2.2  05/29/09  David Boyer - Added check to see if "Additional Information" textarea is displayed or not.
'
'------------------------------------------------------------------------------
'INITIALIZE AND DECLARE VARIABLES
 Dim sError
 sLevel = "../../"  'OVERRIDE OF VALUE FROM COMMON.ASP

'Set timezone information into session
 session("iUserOffset") = request.cookies("tz")

'Get form information
 Dim sIssueName,sIssueDesc, sHideIssueLocAddInfo

 GetFormInformation(request("requestid"))

 lcl_hidden = "hidden"  'Show/Hide hidden fields.  HIDDEN = Hide, TEXT = Show

'Determine if the user has clicked on the "Import Address Fields" button
 if request("importAddressFields") <> "" then
    lcl_importAddressFields = request("importAddressFields")
 else
    lcl_importAddressFields = "N"
 end if

'Check to see if the org has the large address list feature turned on
 lcl_large_address_list = OrgHasFeature("large address list")

 if lcl_importAddressFields = "Y" then
    if lcl_large_address_list then
       lcl_street_number  = request("residentstreetnumber")
       lcl_street_address = request("streetaddress")
    else
       lcl_street_number  = ""
       lcl_street_address = request("streetaddress")
    end if
 else
    lcl_street_number  = ""
    lcl_street_address = ""
 end if
%>
<html>
<head>
  <title>E-GovLink Administration Consule {Edit <%=sIssueName%>}</title>

  <link rel="stylesheet" type="text/css" href="../../global.css" />
  <link rel="stylesheet" type="text/css" href="../../menu/menu_scripts/menu.css" />

  <script language="javascript" src="../../scripts/modules.js"></script>
  <script language="javascript" src="../../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../../scripts/removespaces.js"></script>
  <script language="javascript" src="../../../scripts/easyform.js"></script>
  <script language="javascript" src="../../scripts/setfocus.js"></script>
  <script language="javascript" src="../scripts/selectAll.js"></script>
 	<script language="javascript" src="../../scripts/textareamaxlength.js"></script>
<script language="javascript" > 
<!--
function validateAddress() {
  // Remove any extra spaces
  document.frmlocation.residentstreetnumber.value = removeSpaces(document.frmlocation.residentstreetnumber.value);

  // check the number for non-numeric values
  var rege = /^\d+$/;
  var Ok = rege.exec(document.frmlocation.residentstreetnumber.value);

		if ( ! Ok ) {
    		alert("The Resident Street Number cannot be blank and must be numeric.");
	 	   setfocus(document.frmlocation.residentstreetnumber);
   		 return false;
  }

  // check that they picked a street name
  if ( document.frmlocation.streetaddress.value == '0000') {
 	   	alert("Please select a street name from the list first.");
    		setfocus(document.frmlocation.streetaddress);
   	 	return false;
 	}

  return true;

}

//Set timezone in cookie to retrieve later
var d=new Date();
if (d.getTimezoneOffset) {
   	var iMinutes = d.getTimezoneOffset();
  		document.cookie = "tz=" + iMinutes;
}

function save_address() {
  // CHECK TO SEE IF WE HAVE ADDRESS ON FILE OR IF IT IS A CUSTOM ADDRESS
  if (document.frmlocation.residentstreetnumber.value=="" && document.frmlocation.streetaddress.options[document.frmlocation.streetaddress.selectedIndex].value=="0000") {
 	  		// SUBMIT FORM AS IS
  	 		document.frmlocation.validstreet.value = 'N';
  }else{
 	  		// IF SELECTED ADDRESS CLEAR OUT CUSTOM ADDRESS FIELD
  	 		document.frmlocation.ques_issue2.value = '';
  	 		document.frmlocation.validstreet.value = 'Y';
  }
}

  //Set up global variables
 	var winHandle;
 	var w = (screen.width - 640)/2;
 	var h = (screen.height - 450)/2;

<% if lcl_large_address_list then %>
function checkAddress( sReturnFunction, sSave ) {
  if(document.frmlocation.ques_issue2.value=="") {
     lcl_success = validateAddress();
     if(lcl_success) {
      		// This is here because window.open in the Ajax callback routine will not work
      		//winHandle = eval('window.open("../addresspicker.asp?saving=' + sSave + '&stnumber=' + document.frmlocation.residentstreetnumber.value + '&stname=' + document.frmlocation.streetaddress.value + '&sCheckType=' + sReturnFunction + '&formname=frmlocation", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
      		//self.focus();
      		// Fire off Ajax routine
      		doAjax('../checkaddress.asp', 'stnumber=' + document.frmlocation.residentstreetnumber.value + '&stname=' + document.frmlocation.streetaddress.value, sReturnFunction, 'get', '0');
     }
  }else{
     if(document.frmlocation.residentstreetnumber.value!="" || document.frmlocation.streetaddress.value!="0000") {
        document.frmlocation.ques_issue2.value="";
      		//winHandle = eval('window.open("../addresspicker.asp?saving=' + sSave + '&stnumber=' + document.frmlocation.residentstreetnumber.value + '&stname=' + document.frmlocation.streetaddress.value + '&sCheckType=' + sReturnFunction + '&formname=frmlocation", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
      		doAjax('../checkaddress.asp', 'stnumber=' + document.frmlocation.residentstreetnumber.value + '&stname=' + document.frmlocation.streetaddress.value, sReturnFunction, 'get', '0');
     }else{
      //if(sReturnFunction!="FinalCheck") {
      //   alert("A non-listed address has been entered.");
      //   document.frmlocation.validstreet.value = 'N';
      //}else{
      //   document.frmlocation.validstreet.value = 'Y';
           FinalCheck('NOT FOUND');
      //}
     }
  }
}

function CheckResults( sResults ) {
  // Process the Ajax CallBack when the validate address button is clicked
  if (sResults == 'FOUND CHECK') {
    		//if(winHandle != null && ! winHandle.closed) { 
   	  //			winHandle.close();
   			//}
	 	  	document.frmlocation.ques_issue2.value = '';
      document.frmlocation.validstreet.value = 'Y';
  		 	alert("This is a valid address in the system.");
  }else{
      document.frmlocation.validstreet.value = 'N';
 	  		//winHandle.focus();
      PopAStreetPicker('CheckResults','no');
  }
}

function FinalCheck( sResults ) {
  if (sResults == 'FOUND CHECK') {
    		//if(winHandle != null && ! winHandle.closed) { 
   	  //			winHandle.close();
   			//}
      document.frmlocation.validstreet.value = 'Y';
      document.frmlocation.submit();
  }else{
      if ((sResults == 'FOUND SELECT')||(sResults == 'FOUND KEEP')) {
     		    //if(winHandle != null && ! winHandle.closed) { 
        	  //			winHandle.close();
   			     //}

           if (sResults == 'FOUND SELECT') {
               document.frmlocation.validstreet.value = 'Y';
           }else{
               document.frmlocation.validstreet.value = 'N';
           }

           document.frmlocation.submit();
      }else{
           //document.frmlocation.validstreet.value = 'N';
         		//if(winHandle != null && ! winHandle.closed) { 
           //   winHandle.focus();
           //}else{
           //   document.frmlocation.submit();
           //}
           if(document.frmlocation.ques_issue2.value!="") {
              document.frmlocation.submit();
           } else {
              PopAStreetPicker('FinalCheck','no');
           }
      }
  }
}

function PopAStreetPicker( sReturnFunction, sSave )	{
		// pop up the address picker
		//winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&stnumber=' + document.register.residentstreetnumber.value + '&stname=' + document.register.skip_address.value + '&sCheckType=' + sReturnFunction + '", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
  winHandle = eval('window.open("../addresspicker.asp?saving=' + sSave + '&stnumber=' + document.frmlocation.residentstreetnumber.value + '&stname=' + document.frmlocation.streetaddress.value + '&sCheckType=' + sReturnFunction + '&formname=frmlocation", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}
<% end if %>

function checkImportAddressBtn() {
  document.getElementById("importAddress").disabled=true;

  //If a non-valid address is entered then disable the "Import Address" button.
  <%
   'Determine which fields are valid depending on the which address list the org has
    lcl_address_js_code = ""

    if lcl_large_address_list then
       lcl_address_js_code = "document.getElementById(""residentstreetnumber"").value!="""" || "
       lcl_address_js_code = lcl_address_js_code & "document.getElementById(""streetaddress"").value!=""0000"""
    else
       lcl_address_js_code = "document.getElementById(""streetaddress"").value!=""0000"""
    end if
  %>


  if(<%=lcl_address_js_code%>) {
     document.getElementById("importAddress").disabled=false;
  }
}

function getAddressFields() {
  document.getElementById("importAddressFields").value = "Y";
  document.getElementById("frmlocation").action = "correction_issue_location.asp";
  document.getElementById("frmlocation").submit();
}
//-->
</script>

<STYLE>
		div.correctionsbox           { border: solid 1px #336699;padding: 4px 0px 0px 4px ; }
		div.correctionsboxnotfound   { background-color:#e0e0e0;border: solid 1px #000000;padding: 10px;color:red;font-weight:bold; }
		td.correctionslabel          { font-weight:bold; }
		th.corrections               { background-color:#93bee1;font-size:12px;padding:5px;color:#000000; }
		input.correctionstextbox     { border: solid 1px #336699;width:400px; }
		textarea.correctionstextarea { border: solid 1px #336699;width:600px;height:150px; }
		p.instructions               { padding: 10px; }
		span.maptrue                 { color:green;font-weight:bold; }
		span.mapfalse                { color:red;font-weight:bold; }
		.savemsg                     { font-size:12px;padding:5px;color:#0000ff;font-weight:bold; }
</STYLE>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="setMaxLength();checkImportAddressBtn();">
<% ShowHeader sLevel %>
<!--#Include file="../../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
 	<div id="centercontent">

<h3>Edit <%=sIssueName%></h3>
<input type="button" name="backButton" id="backButton" value="Return to Request" class="button" onclick="location.href='../action_respond.asp?control=<%=request("requestid")%>';" />
<%
	'DISPLAY TO USER THAT VALUES WERE SAVED
		If request("r") = "save" Then 
			  response.write "<p><span class=""savemsg"">Saved " & Now() & ".</span></p>"
		End If

	'GET ISSUE LOCATION INFORMATION
		SubDrawEditIssueLocationInformation request("requestid"),lcl_street_number,lcl_street_address
%>
 	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../../admin_footer.asp"-->  
</body>
</html>
<%
'------------------------------------------------------------------------------
Function SubDrawEditIssueLocationInformation(irequestid,p_street_number,p_street_address)

 'Determine if the form is diplaying the issue/problem location section
  blnIssueDisplay      = checkForIssueLocationOnForm(irequestid)
  sHideIssueLocAddInfo = checkForHideIssueLocAddInfo(irequestid)

	'Check for empty or missing userid
	if isnull(irequestid) or irequestid = "" then
  		response.write "<div class=""correctionsboxnotfound"">No information available for the issue location for this request.</div>" & vbcrlf
	else
 		'GET INFORMATION FOR SPECIFIED USER
 			sSQL = "SELECT * "
    sSQL = sSQL & " FROM egov_action_response_issue_location "
    sSQL = sSQL & " WHERE actionrequestresponseid = '" & iRequestID & "'"

 			set oIssueLocation = Server.CreateObject("ADODB.Recordset")
	  	oIssueLocation.Open sSQL, Application("DSN"), 3, 1

 			if not oIssueLocation.eof then

      'Check to see if the user wants to import the address fields because the street number/name has been changed.
       if lcl_importaddressfields = "Y" then
          if lcl_large_address_list then
             GetAddressInfoLarge p_street_number, p_street_address, sNumber, sPrefix, sAddress, sSuffix, sDirection, _
                                 sLatitude, sLongitude, sCity, sState, sZip, sCounty, sParcelID, sListedOwner, sLegalDescription, _
                                 sResidentType, sRegisteredUserID, sValidStreet
          else
      		    	GetAddressInfo p_street_address, sNumber, sPrefix, sAddress, sSuffix, sDirection, sLatitude, sLongitude, sCity, sState, _
                            sZip, sCounty, sParcelID, sListedOwner, sLegalDescription, sResidentType, sRegisteredUserID, sValidStreet
          end if

         'These fields are NOT to be overridden during the import
          sUnit             = request("streetunit")
          sComments         = request("comments")
          sSortStreetName   = request("sortstreetname")
       else
          sNumber           = oIssueLocation("streetnumber")
          sPrefix           = oIssueLocation("streetprefix")
          sAddress          = oIssueLocation("streetaddress")
          sSuffix           = oIssueLocation("streetsuffix")
          sDirection        = oIssueLocation("streetdirection")
          sCity             = oIssueLocation("city")
          sState            = oIssueLocation("state")
          sZip              = oIssueLocation("zip")
          sLatitude         = oIssueLocation("latitude")
          sLongitude        = oIssueLocation("longitude")
          sUnit             = oIssueLocation("streetunit")
          sCounty           = oIssueLocation("county")
          sParcelID         = oIssueLocation("parcelidnumber")
          sListedOwner      = oIssueLocation("listedowner")
          sLegalDescription = oIssueLocation("legaldescription")
          sComments         = oIssueLocation("comments")
          sValidStreet      = oIssueLocation("validstreet")
          sSortStreetName   = oIssueLocation("sortstreetname")
          sResidentType     = oIssueLocation("residenttype")
          sRegisterUserID   = oIssueLocation("registereduserid")
       end if

   				response.write "<form name=""frmlocation"" id=""frmlocation"" action=""correction_issue_location_cgi.asp"" method=""post"">" & vbcrlf
  		 		response.write "  <input type=""" & lcl_hidden & """ name=""status"" id=""status"" value=""" &              request("status")    & """ />" & vbcrlf
		  	 	response.write "  <input type=""" & lcl_hidden & """ name=""substatus"" id=""substatus"" value=""" &        request("substatus") & """ />" & vbcrlf
				   response.write "  <input type=""" & lcl_hidden & """ name=""requestid"" id=""requestid"" value=""" &        iRequestID           & """ />" & vbcrlf
       response.write "  <input type=""" & lcl_hidden & """ name=""validstreet"" value=""" &                       sValidStreet         & """ />" & vbcrlf
       response.write "  <input type=""" & lcl_hidden & """ name=""sortstreetname"" value=""" &                    sSortStreetName      & """ />" & vbcrlf
       response.write "  <input type=""" & lcl_hidden & """ name=""residenttype"" maxlength=""10"" value=""" &     sResidentType        & """ />" & vbcrlf
       response.write "  <input type=""" & lcl_hidden & """ name=""registereduserid"" maxlength=""20"" value=""" & sRegisterUserID      & """ />" & vbcrlf
       response.write "  <input type=""" & lcl_hidden & """ name=""importAddressFields"" id=""importAddressFields"" size=""1"" maxlength=""1"" value=""" & lcl_importAddressFields & """ />" & vbcrlf
			   	response.write "<div class=""shadow"">" & vbcrlf
   				response.write "<table class=""tablelist"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
			   	response.write "  <tr>" & vbcrlf
       response.write "      <th class=""corrections"" colspan=""2"" align=""left"">&nbsp;" & sIssueName & "</th>" & vbcrlf
       response.write "  </tr>" & vbcrlf

   			'DISPLAY INSTRUCTIONS
    			response.write "  <tr>" & vbcrlf
       response.write "      <td colspan=""2"">" & vbcrlf
       response.write "          <p class=""instructions"">" & vbcrlf
				   response.write            sIssueDesc
   				response.write "          Press <b>Save</b> when finished making changes.</p>" & vbcrlf
       response.write "      </td>" & vbcrlf
       response.write "  </tr>" & vbcrlf

				  'DISPLAY SAVE AND CANCEL BUTTONS
   				response.write "  <tr>" & vbcrlf
       response.write "      <td class=""correctionslabel"" align=""left"" colspan=""2"">" & vbcrlf

                                 displayButtons blnIssueDisplay, lcl_large_address_list, request("requestid")

       response.write "      </td></tr>" & vbcrlf

				  'DISPLAY STREET
			   	response.write "  <tr>" & vbcrlf
       response.write "      <td class=""correctionslabel"" align=""right"" valign=""top"">Street:</td>" & vbcrlf
       response.write "      <td>" & vbcrlf

      'Determine if the "other address" is populated or not
       if sValidStreet = "Y" then
          lcl_street_name = buildStreetAddress("", sPrefix, sAddress, sSuffix, sDirection)

          if lcl_large_address_list then
             DisplayLargeAddressList session("orgid"), "R", sNumber, lcl_street_name
          else
             DisplayAddress session("orgid"), "R", 1, sNumber, lcl_street_name
          end if

          lcl_display_other_address = ""
       else
          if lcl_large_address_list then
             DisplayLargeAddressList session("orgid"), "R", "", ""
          else
             DisplayAddress session("orgid"), "R", 1, "", ""
          end if

          lcl_display_other_address = sNumber
          if lcl_display_other_address <> "" then
             lcl_display_other_address = lcl_display_other_address & " " & sAddress
          else
             lcl_display_other_address = sAddress
          end if
       end if

       response.write "          <input type=""button"" id=""importAddress"" class=""button"" value=""Import Address Fields"" onclick=""getAddressFields()"" />" & vbcrlf
       response.write "          <br /> - Or Other Not Listed - <br /> " & vbcrlf
       response.write "          <input type=""text"" name=""ques_issue2"" id=""ques_issue2"" class=""correctionstextbox"" maxlength=""75"" value=""" & lcl_display_other_address & """ onchange=""save_address();checkImportAddressBtn()"" />" & vbcrlf
			   	response.write "      </td>" & vbcrlf
       response.write "  </tr>" & vbcrlf

  				'DISPLAY CITY
   				response.write "  <tr>" & vbcrlf
       response.write "      <td class=""correctionslabel"" align=""right"">City:</td>" & vbcrlf
       response.write "      <td><input name=""city"" type=""text"" class=""correctionstextbox"" maxlength=""50"" value=""" & sCity & """></td>" & vbcrlf
       response.write "  </tr>" & vbcrlf

				  'DISPLAY STATE
   				response.write "  <tr>" & vbcrlf
       response.write "      <td class=""correctionslabel"" align=""right"">State:</td>" & vbcrlf
       response.write "      <td><input name=""state"" type=""text"" class=""correctionstextbox"" maxlength=""50"" value=""" & sState & """></td>" & vbcrlf
       response.write "  </tr>" & vbcrlf

				  'DISPLAY ZIP
   				response.write "  <tr>" & vbcrlf
       response.write "      <td class=""correctionslabel"" align=""right"">Zip:</td>" & vbcrlf
       response.write "      <td><input name=""zip"" type=""text"" class=""correctionstextbox"" maxlength=""15"" value=""" & sZip & """></td>" & vbcrlf
       response.write "  </tr>" & vbcrlf

				  'DISPLAY UNIT
   				response.write "  <tr>" & vbcrlf
       response.write "      <td class=""correctionslabel"" align=""right"">Unit:</td>" & vbcrlf
       response.write "      <td><input name=""streetunit"" type=""text"" class=""correctionstextbox"" maxlength=""10"" value=""" & sUnit & """></td>" & vbcrlf
       response.write "  </tr>" & vbcrlf

       lcl_display_id = 0

      'Build the custom label
       lcl_display_id   = GetDisplayId("address grouping field")
       lcl_county_label = GetOrgDisplayWithId(session("orgid"),lcl_display_id,true)

       if lcl_county_label = "" then
          lcl_county_label = GetDisplayName(lcl_display_id)
       end if

   				response.write "  <tr>" & vbcrlf
       response.write "      <td class=""correctionslabel"" align=""right"">" & lcl_county_label & ":</td>" & vbcrlf
       response.write "      <td><input name=""county"" type=""text"" class=""correctionstextbox"" maxlength=""50"" value=""" & sCounty & """ /></td>" & vbcrlf
       response.write "  </tr>" & vbcrlf

       if OrgHasFeature("parcel_id") then
				     'Display Parcel ID
      				response.write "  <tr>" & vbcrlf
          response.write "      <td class=""correctionslabel"" align=""right"">Parcel Id No:</td>" & vbcrlf
          response.write "      <td><input name=""parcelidnumber"" type=""text"" class=""correctionstextbox"" maxlength=""50"" value=""" & sParcelID & """ /></td>" & vbcrlf
          response.write "  </tr>" & vbcrlf

   				  'Display Listed Owner
      				response.write "  <tr>" & vbcrlf
          response.write "      <td valign=""top"" class=""correctionslabel"" align=""right"">Listed Owner:</td>" & vbcrlf
          response.write "      <td><textarea id=""owner"" name=""listedowner"" rows=""3"" cols=""80"" maxlength=""250"" wrap=""soft"" class=""correctionstextarea"" style=""height: 55px"">" & sListedOwner & "</textarea></td>" & vbcrlf
          response.write "  </tr>" & vbcrlf

   				  'Display Legal Description
      				response.write "  <tr>" & vbcrlf
          response.write "      <td valign=""top"" class=""correctionslabel"" align=""right"">Legal Description:</td>" & vbcrlf
          response.write "      <td><textarea id=""owner"" name=""legaldescription"" rows=""3"" cols=""80"" maxlength=""400"" wrap=""soft"" class=""correctionstextarea"" style=""height: 55px"">" & sLegalDescription & "</textarea></td>" & vbcrlf
          response.write "  </tr>" & vbcrlf
       else
          response.write "<input type=""hidden"" name=""parcelidnumber"" maxlength=""50"" value="""    & sParcelID         & """ />" & vbcrlf
          response.write "<input type=""hidden"" name=""listedowner"" maxlength=""250"" value="""      & sListedOwner      & """ />" & vbcrlf
          response.write "<input type=""hidden"" name=""legaldescription"" maxlength=""400"" value=""" & sLegalDescription & """ />" & vbcrlf
       end if

				  'Display Additional Information
       if not sHideIssueLocAddInfo then
      				response.write "  <tr>" & vbcrlf
          response.write "      <td valign=""top"" class=""correctionslabel"" align=""right"">Additional Information:</td>" & vbcrlf
          response.write "      <td><textarea class=""correctionstextarea"" name=""comments"" maxlength=""512"">" & sComments & "</textarea></td>" & vbcrlf
          response.write "  </tr>" & vbcrlf
       end if

       'response.write "      </td>" & vbcrlf
       'response.write "  </tr>" & vbcrlf

   				response.write "  <tr>" & vbcrlf
       response.write "      <td class=""correctionslabel"" align=""left"" colspan=""2"">" & vbcrlf
                                 displayButtons blnIssueDisplay, lcl_large_address_list, request("requestid")
       response.write "      </td>" & vbcrlf
       response.write "  </tr>" & vbcrlf
   				response.write "</table>" & vbcrlf
		   		response.write "</div>" & vbcrlf
   				response.write "</form>" & vbcrlf
    Else
		 		 'NO MATCHING USER FOUND
  	 			response.write "<div class=""correctionsboxnotfound"">No information available for the issue location for this request.</div>" & vbcrlf
		 	End If
End If

End Function

'------------------------------------------------------------------------------
Function DisplayAddress( iorgid, sResidenttype, sAddress, p_street_number, p_street_name )
	
	Dim sNumber, sSQL, oAddressList, blnFound

 lcl_new_street_name = buildStreetAddress(p_street_number, "", p_street_name, "", "")

'GET LIST OF ADDRESSES FOR ORGANIZATION
	sSQL = "SELECT residentaddressid, isnull(residentstreetnumber,'') as residentstreetnumber, residentstreetprefix, "
 sSQL = sSQL & " residentstreetname, streetsuffix, streetdirection, "
 sSQL = sSQL & " isnull(latitude,0.00) as latitude, isnull(longitude,0.00) as longitude "
 sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid=" & iorgid
' sSQL = sSQL & " AND residenttype='" & sResidenttype & "' "
 sSQL = sSQL & " AND excludefromactionline = 0 "
 sSQL = sSQL & " AND residentstreetname is not null "
 sSQL = sSQL & " ORDER BY sortstreetname, residentstreetprefix, Cast(residentstreetnumber as int)"

	Set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSQL, Application("DSN"), 3, 1

'DISPLAY ADDRESS SELECT BOX
	response.write "<select name=""streetaddress"" id=""streetaddress"" onchange=""save_address();checkImportAddressBtn()"">" & vbcrlf
	response.write "  <option value=""0000"">Choose street from dropdown</option>" & vbcrlf
	
'LOOP THRU RESIDENT ADDRESS FOR CITY
	do while NOT oAddressList.eof
    lcl_original_street_name = buildStreetAddress(oAddressList("residentstreetnumber"), oAddressList("residentstreetprefix"), oAddressList("residentstreetname"), oAddressList("streetsuffix"), oAddressList("streetdirection"))

   'CHECK TO SEE IF WE HAVE MATCHING ADDRESS
    if lcl_large_address_list then
       if trim(lcl_new_street_name) = trim(lcl_original_street_name) then
    			    sSelected = "SELECTED "
     		else
    		    	sSelected = ""
     		end If
    else
       if p_street_number <> "" or not IsNull(p_street_number) then
          if trim(lcl_original_street_name) = trim(lcl_new_street_name) then
              sSelected = "SELECTED"
          else
              sSelected = ""
          end if
       else
          if p_street_name = oAddressList("residentstreetname") then
              sSelected = "SELECTED"
          else
              sSelected = ""
          end if
       end if
    end if

   	response.write "  <option " & sSelected & " value=""" & oAddressList("residentaddressid") & """>" & lcl_original_street_name & "</option>" & vbcrlf

  		oAddressList.MoveNext
	Loop

	response.write "</select>" & vbcrlf

'CLEAN UP OBJECTS
	oAddressList.close
	Set oAddressList = Nothing 

End Function

'------------------------------------------------------------------------------
Sub DisplayLargeAddressList( iAddressOrgId, sResidenttype, p_street_number, p_street_name )
Dim sSql, oAddressList

	sSQL = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSQL = sSQL & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid = " & iAddressOrgId
	sSQL = sSQL & " AND residentstreetname IS NOT NULL "
 sSQL = sSQL & " AND excludefromactionline = 0 "
 sSQL = sSQL & " ORDER BY sortstreetname "
	
Set oAddressList = Server.CreateObject("ADODB.Recordset")
oAddressList.Open sSQL, Application("DSN"), 3, 1

If Not oAddressList.EOF Then
 		response.write "<input type=""text"" name=""residentstreetnumber"" id=""residentstreetnumber"" value=""" & p_street_number & """ size=""8"" maxlength=""10"" onchange=""save_address();"" /> &nbsp; " & vbcrlf
	 	response.write "<select name=""streetaddress"" id=""streetaddress"" onchange=""save_address();checkImportAddressBtn()"">" & vbcrlf
 		response.write "  <option value=""0000"">Choose street from dropdown</option>" & vbcrlf

   do while NOT oAddressList.eof
      sCompareName = buildStreetAddress("", oAddressList("residentstreetprefix"), oAddressList("residentstreetname"), oAddressList("streetsuffix"), oAddressList("streetdirection"))
      response.write "<option value=""" & sCompareName & """"

    		if p_street_name = sCompareName then
      			response.write " selected=""selected"""
    		end if

    		response.write ">"
     	response.write sCompareName & "</option>" & vbcrlf
   			oAddressList.MoveNext
   loop 

   response.write "</select>&nbsp;" & vbcrlf
   response.write "<input type=""button"" id=""validateAddress""  class=""button"" value=""Validate Address"" onclick=""checkAddress( 'CheckResults', 'no');"" />" & vbcrlf
End If 

oAddressList.Close
Set oAddressList = Nothing 

End Sub 

'------------------------------------------------------------------------------
Sub GetFormInformation(iRequestID)

	sSQL = "SELECT *, (FirstName + ' ' + LastName) as EmployeeSubmitName, F.DeptID, F.hideIssueLocAddInfo "
 sSQL = sSQL & " FROM egov_actionline_requests "
	sSQL = sSQL & " LEFT OUTER JOIN users ON egov_actionline_requests.employeesubmitid = users.userid "
	sSQL = sSQL & " LEFT OUTER JOIN egov_users ON egov_actionline_requests.userid = egov_users.userid "
	sSQL = sSQL & " LEFT OUTER JOIN egov_action_request_forms AS F ON egov_actionline_requests.category_id = F.action_form_id "
	sSQL = sSQL & " WHERE action_autoid=" & iRequestID
	
	Set oRequest = Server.CreateObject("ADODB.Recordset")
	oRequest.Open sSQL, Application("DSN"), 3, 1

'CHECK FOR INFORMATION
	If NOT oRequest.EOF Then
 		'REQUEST FOUND GET INFORMATION	
  		sIssueName = oRequest("issuelocationname")
		  If Trim(sIssueName) = "" OR IsNull(sIssueName) Then
   			 sIssueName = "Issue/Problem Location:"
  		End If
  		sIssueDesc = oRequest("issuelocationdesc")
		  If IsNull(sIssueDesc) Then
    			sIssueDesc = "Please select the closest street number/streetname of problem location from list or select ""*not on list"". "
       sIssueDesc = sIssueDesc & " Provide any additional information on problem location in the box below."
  		End If

	End If

	Set oRequest = Nothing 

End Sub

'------------------------------------------------------------------------------
sub displayButtons(iIssueDisplay, iLargeAddressList, iRequestID)

 'Build the "onclick"
  if iIssueDisplay then
     if iLargeAddressList then
        lcl_onclick = "checkAddress( 'FinalCheck', 'no' );"
     else
        lcl_onclick = "if(document.frmlocation.streetaddress.value=='0000'&&document.frmlocation.ques_issue2.value==''){"
        lcl_onclick = lcl_onclick & "alert('Required field missing: Street');"
        lcl_onclick = lcl_onclick & "document.frmlocation.streetaddress.focus();"
        lcl_onclick = lcl_onclick & "}else{"
        lcl_onclick = lcl_onclick & "document.frmlocation.submit();"
        lcl_onclick = lcl_onclick & "}"
     end if
  else
     lcl_onclick = lcl_onclick & "document.frmlocation.submit();"
  end if

  response.write "<input type=""button"" value=""Save"" onclick=""" & lcl_onclick & """ />" & vbcrlf
  response.write "<input type=""button"" value=""Cancel"" onclick=""location.href='../action_respond.asp?control=" & iRequestID & "';"" />" & vbcrlf
end sub
%>