<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../include_top_functions.asp" //-->
<!-- #include file="../action_line_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME:  searchzoning.asp
' AUTHOR:    David Boyer
' CREATED:   06/03/09
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Display a search input field(s) and results list for the custom "County" field for Valid Addresses.
'
' MODIFICATION HISTORY
' 1.0  06/03/09  David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

'Determine which "search type" they wish to display or if we show both options.
'Search Types:
' "A" = Street Address
' "P" = Parcel Number
' ""  = Show both options
 if request("stype") <> "" then
    lcl_searchtype = UCASE(request("stype"))
 else
    lcl_searchtype = ""
 end if

'Check for a background color
 if request("bgcolor") <> "" then
    lcl_bgcolor = request("bgcolor")
 else
    lcl_bgcolor = "ffffff"
 end if

'Check for org features
 lcl_orghasfeature_large_address_list = OrgHasFeature(iorgid,"large address list")

'Setup page variables
 lcl_displaySearchField = 0
 lcl_countyLabel        = ""
 lcl_sc_parcelidnumber  = request("sc_pid")
 lcl_displayid          = GetDisplayId("address grouping field")
 lcl_sc_streetnumber    = ""
 lcl_sc_streetaddress   = request("skip_address")

 if lcl_orghasfeature_large_address_list then
    lcl_sc_streetnumber  = request("residentstreetnumber")
 end if

 if orghasdisplay(iorgid,"address grouping field") then
    lcl_countyLabel = UCASE(GetOrgDisplayWithId(iorgid, lcl_displayid, True)) & " "
 end if

 'Determine which search field(s) to display
  if UCASE(lcl_searchtype) = "P" then
     lcl_displaySearchField = lcl_displaySearchField + 1
  end if

  if UCASE(lcl_searchtype) = "A" then
     lcl_displaySearchField = lcl_displaySearchField + 1
  end if
%>
<html>
<head>

	<title></title>
	
	<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
	
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

 <script language="javascript" src="../scripts/ajaxLib.js"></script>
 <script language="javascript" src="../scripts/removespaces.js"></script>
 <script language="javascript" src="../scripts/setfocus.js"></script>

<script language="javascript">
 	var winHandle;
 	var w = (screen.width - 640)/2;
 	var h = (screen.height - 450)/2;

<%
  lcl_searchFieldCheck = ""

  if lcl_displaySearchField = 0 then
     lcl_searchFieldCheck = lcl_searchFieldCheck & "document.getElementById('btnSubmit').disabled=true;" & vbcrlf
     lcl_searchFieldCheck = lcl_searchFieldCheck & "if("
     lcl_searchFieldCheck = lcl_searchFieldCheck & "(document.getElementById(""sc_pid"").value!='')&&"

     if lcl_orghasfeature_large_address_list then
        lcl_searchFieldCheck = lcl_searchFieldCheck & " ("
        lcl_searchFieldCheck = lcl_searchFieldCheck &    "(document.getElementById(""residentstreetnumber"").value!='')"
        lcl_searchFieldCheck = lcl_searchFieldCheck & " &&(document.getElementById(""skip_address"").value!='0000')"
        lcl_searchFieldCheck = lcl_searchFieldCheck & " )"
     else
        lcl_searchFieldCheck = lcl_searchFieldCheck & "(document.getElementById(""skip_address"").value!='0000')"
     end if

     lcl_searchFieldCheck = lcl_searchFieldCheck & ") {" & vbcrlf
     lcl_searchFieldCheck = lcl_searchFieldCheck & "  alert('Invalid Value: Either a Parcel Number OR a Street Address may be entered.');" & vbcrlf
     lcl_searchFieldCheck = lcl_searchFieldCheck & "  document.getElementById('btnSubmit').disabled=false;" & vbcrlf
     lcl_searchFieldCheck = lcl_searchFieldCheck & "  setfocus(document.searchZoning.sc_pid);" & vbcrlf
     lcl_searchFieldCheck = lcl_searchFieldCheck & "}else{" & vbcrlf
     lcl_searchFieldCheck = lcl_searchFieldCheck & "  if(document.getElementById(""sc_pid"").value!='') {" & vbcrlf
     lcl_searchFieldCheck = lcl_searchFieldCheck & "     document.getElementById(""searchZoning"").submit();" & vbcrlf
     lcl_searchFieldCheck = lcl_searchFieldCheck & "  }else{" & vbcrlf

     if lcl_orghasfeature_large_address_list then
        lcl_searchFieldCheck = lcl_searchFieldCheck & "     checkAddress( 'FinalCheck', 'yes' );" & vbcrlf
     else
        lcl_searchFieldCheck = lcl_searchFieldCheck & "     document.getElementById(""searchZoning"").submit();" & vbcrlf
     end if

     lcl_searchFieldCheck = lcl_searchFieldCheck & "  }" & vbcrlf
     lcl_searchFieldCheck = lcl_searchFieldCheck & "}" & vbcrlf
  else
     lcl_searchFieldCheck = lcl_searchFieldCheck & "document.getElementById(""searchZoning"").submit();" & vbcrlf
  end if
%>
function validateFields() {
  <%=lcl_searchFieldCheck%>
}
<%
 'ONLY show function if the org has the "large address list" feature
  if lcl_orghasfeature_large_address_list then
%>
function checkAddress( sReturnFunction, sSave ) {
  //Disable the submit button
  document.getElementById("btnSubmit").disabled=true;

		// Remove any extra spaces
		document.searchZoning.residentstreetnumber.value = removeSpaces(document.searchZoning.residentstreetnumber.value);

		// check the number for non-numeric values
		var rege = /^\d+$/;
		var Ok = rege.exec(document.searchZoning.residentstreetnumber.value);

  if ( ! Ok ) {
  		alert("The Street Number cannot be blank and must be numeric.");
	  	setfocus(document.searchZoning.residentstreetnumber);
    document.getElementById("btnSubmit").disabled=false;
   	return false;
  }

  // check that they picked a street name
  if ( document.searchZoning.skip_address.value == '0000') {
 	 	alert("Please select a street name from the list first.");
  		setfocus(document.searchZoning.skip_address);
    document.getElementById("btnSubmit").disabled=false;
   	return false;
  }
  // This is here because window.open in the Ajax callback routine will not work
  // Fire off Ajax routine
  doAjax('../checkaddress.asp', 'stnumber=' + document.searchZoning.residentstreetnumber.value + '&stname=' + document.searchZoning.skip_address.value + '&orgid=<%=iorgid%>', sReturnFunction, 'get', '0');
}

function CheckResults( sResults ) {
  // Process the Ajax CallBack when the validate address button is clicked
  if (sResults == 'FOUND CHECK') {
  		 	alert("This is a valid address in the system.");
  }else{
   			PopAStreetPicker('CheckResults', 'no');
  }

  document.getElementById("btnSubmit").disabled=false
}

function FinalCheck( sResults ) {
  if (sResults == 'FOUND CHECK') {
      document.getElementById("searchZoning").submit();
  }else{
      if ((sResults == 'FOUND SELECT')||(sResults == 'FOUND KEEP')) {
         document.getElementById("searchZoning").submit();
      }else{
      			PopAStreetPicker('FinalCheck', 'yes');
      }
  }
}
<% end if %>

function PopAStreetPicker( sReturnFunction, sSave ) {
		// pop up the address picker
  winHandle = eval('window.open("../addresspicker.asp?saving=' + sSave + '&validaddressonly=Y&stnumber=' + document.searchZoning.residentstreetnumber.value + '&stname=' + document.searchZoning.skip_address.value + '&sCheckType=' + sReturnFunction + '&formname=searchZoning", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}
</script>
</head>
<body bgcolor="<%=lcl_bgcolor%>">
<%
'Display the search field(s)
 displaySearchField lcl_searchtype, lcl_bgcolor

 if request.ServerVariables("request_method") = "POST" then
   'Build the search query
   	sSQL = "SELECT residentaddressid, county "
  		sSQL = sSQL & " FROM egov_residentaddresses "
  	 sSQL = sSQL & " WHERE orgid = " & iorgid
    sSQL = sSQL & " AND county IS NOT NULL "
    sSQL = sSQL & " AND county <> '' "
    'response.write iorgid

  	'Determine if which option we are searching on.
  	 response.write "<p>" & vbcrlf
  	 response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf

  	 if lcl_sc_parcelidnumber <> "" OR lcl_sc_streetaddress <> "" then
     	 if lcl_sc_parcelidnumber <> "" then

        	 lcl_displayLabel = "Parcel Number: <span style=""color:#800000"">" & lcl_sc_parcelidnumber & "</span>"

  	         sSQL = sSQL & " AND UPPER(parcelidnumber) = '" & UCASE(lcl_sc_parcelidnumber) & "' "

     	 elseif lcl_sc_streetaddress <> "" then

  	       if lcl_orghasfeature_large_address_list then
     	       lcl_displayLabel     = "Street Address: <span style=""color:#800000"">" & lcl_sc_streetnumber & " " & lcl_sc_streetaddress & "</span>"
             lcl_sc_streetnumber  = dbsafe(lcl_sc_streetnumber)
             lcl_sc_streetaddress = dbsafe(lcl_sc_streetaddress)

  	          sSQL = sSQL & " AND UPPER(residentstreetnumber) = UPPER('" & lcl_sc_streetnumber & "') "
     	       sSQL = sSQL & " AND (residentstreetname = '" & lcl_sc_streetaddress & "' "
        	    sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix = '" & lcl_sc_streetaddress & "' "
  	          sSQL = sSQL & " OR residentstreetname + ' ' + streetdirection = '" & lcl_sc_streetaddress & "' "
     	       sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & lcl_sc_streetaddress & "' "
        	    sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname = '" & lcl_sc_streetaddress & "' "
  	          sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = '" & lcl_sc_streetaddress & "' "
     	       sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetdirection = '" & lcl_sc_streetaddress & "' "
        	    sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & lcl_sc_streetaddress & "') "
  	       else
     	      	sSQL = sSQL & " AND residentaddressid = " & lcl_sc_streetaddress
  	       end if

  	    end if

  	   	set oZoningInfo = Server.CreateObject("ADODB.Recordset")
     		oZoningInfo.Open sSQL, Application("DSN"), 3, 1

  	    if not oZoningInfo.eof then
    	    	do while not oZoningInfo.eof

  	          response.write "  <tr><td align=""center""><strong>" & lcl_countyLabel & "INFORMATION FOUND FOR " & lcl_displayLabel & "</strong><br /><br /></td></tr>" & vbcrlf
     	       response.write "  <tr><td align=""center"" style=""font-size:20px; color:#800000; font-weight:bold;"">" & oZoningInfo("county") & "</td></tr>" & vbcrlf

        	    oZoningInfo.movenext
  	       loop
     	 else
        	 response.write "  <tr><td align=""center"" style=""color:#ff0000; font-weight:bold;"">*** No Zoning Information Available... ***</td></tr>" & vbcrlf
  	    end if

  	   	oZoningInfo.close
     		set oZoningInfo = nothing

  	 else
     	 response.write "  <tr><td align=""center"" style=""color:#ff0000; font-weight:bold;"">*** No criteria entered to search on... ***</td></tr>" & vbcrlf
  	 end if

  	 response.write "</table>" & vbcrlf
  	 response.write "</p>" & vbcrlf
 end if
%>

</body>
</html>
<%
'------------------------------------------------------------------------------
sub displaySearchField(iSearchType, iBGColor)

 'Determine which search field(s) to display
  lcl_displaySearchField = 0

  if UCASE(iSearchType) = "P" then
     lcl_displaySearchField = lcl_displaySearchField + 1
  end if

  if UCASE(iSearchType) = "A" then
     lcl_displaySearchField = lcl_displaySearchField + 1
  end if

 'Display Search Options
  response.write "<fieldset>" & vbcrlf
  response.write "  <legend>Search Options&nbsp;</legend>" & vbcrlf
  response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "    <form name=""searchZoning"" id=""searchZoning"" action=""searchzoning.asp"" method=""post"">" & vbcrlf
  response.write "      <input type=""hidden"" name=""stype"" id=""stype"" size=""5"" maxlength=""5"" value=""" & lcl_searchtype & """ />" & vbcrlf
  response.write "      <input type=""hidden"" name=""bgcolor"" id=""bgcolor"" size=""5"" maxlength=""6"" value=""" & iBGColor & """ />" & vbcrlf

  if UCASE(iSearchType) = "P" OR lcl_displaySearchField = 0 then
     response.write "  <tr>" & vbcrlf
     response.write "      <td><strong>Parcel Number: </strong></td>" & vbcrlf
     response.write "      <td><input type=""text"" name=""sc_pid"" id=""sc_pid"" size=""50"" maxlength=""50"" />" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

  if UCASE(iSearchType) = "A" OR lcl_displaySearchField = 0 then

     if lcl_displaySearchField = 0 then
        response.write "  <tr>" & vbcrlf
        response.write "      <td colspan=""2"">----- OR ------</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
     end if

     response.write "  <tr>" & vbcrlf
     response.write "      <td><strong>Street Address: </strong></td>" & vbcrlf
     response.write "      <td>" & vbcrlf

 				if lcl_orghasfeature_large_address_list then
   					DisplayLargeAddressList iorgid, "R" 
 				else
   					DisplayAddress iorgid, "R"
 				end if

     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf

  if UCASE(iSearchType) = "A" OR lcl_displaySearchField = 0 then
     if lcl_orghasfeature_large_address_list AND lcl_displaySearchField > 0 then
        lcl_onclick = "document.getElementById('btnSubmit').disabled=false;"
        lcl_onclick = "checkAddress( 'FinalCheck', 'yes' );" & vbcrlf
     else
        lcl_onclick = lcl_onclick & "document.getElementById('btnSubmit').disabled=false;"
        lcl_onclick = lcl_onclick & "validateFields();" & vbcrlf
     end if
  else
     lcl_onclick = "document.getElementById('btnSubmit').disabled=false;validateFields();"
  end if

  response.write "          <input type=""button"" name=""btnSubmit"" id=""btnSubmit"" value=""Search"" style=""cursor:pointer;"" onclick=""" & lcl_onclick & """ />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

  response.write "    </form>" & vbcrlf
  response.write "  </table>" & vbcrlf

  response.write "</fieldset>" & vbcrlf

end sub

'------------------------------------------------------------------------------
function DBsafe( strDB )
 	Dim sNewString
 	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
 	sNewString = Replace( strDB, "'", "''" )
 	sNewString = Replace( sNewString, "<", "&lt;" )
 	DBsafe = sNewString
end function
%>
