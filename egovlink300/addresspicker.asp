<!DOCTYPE HTML PUBLIC "-//W3C//DTD XHTML 1.1 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="include_top_functions.asp" //-->
<%
  if request.querystring("formname") <> "" then
     lcl_form_name = request.querystring("formname")
  else
     lcl_form_name = "register"
  end if

 'Set up the form variables
  select case lcl_form_name
    case "frmRequestAction"  'Action Line - Create
         lcl_other_address = "ques_issue2"
         lcl_message_text  = "an Action Line Request"
    case "searchZoning"      'Check Your Zoning
         lcl_other_address = ""
         lcl_message_text  = "a Street Address"
    case else
         lcl_other_address = "egov_users_useraddress"
         lcl_message_text  = "a registration"
  end select

 'Determine we are to allow valid addresses only or if the user is allowed to enter an invalid address.
  if request("validaddressonly") <> "" then
     lcl_validaddressonly = UCASE(request("validaddressonly"))
  else
     lcl_validaddressonly = "N"
  end if
%>
<html>
	<head>
		<link rel="stylesheet" type="text/css" href="global.css" />
 	<link rel="stylesheet" type="text/css" href="css/styles.css" />
	 <link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

		<script language="javascript">
		<!--

			//window.onunload = function (e) {saybye('<%=request("saving")%>'); register(e)};

			function saybye( sString ) {
  			if (sString == 'yes') {
  					//alert('You need to click \'Create User\' again to complete this registration.');
		 		}
			}

			function doSelect() {
			 	// check that an address was picked
				 if(document.frmAddress.stnumber.selectedIndex < 0) {
  			 	 alert('Please select an address from the list first.');
  				 	return false;
 				}
	 			window.opener.document.<%=lcl_form_name%>.residentstreetnumber.value = document.frmAddress.stnumber.value;
     <%
       if lcl_validaddressonly = "N" then
          response.write "window.opener.document." & lcl_form_name & "." & lcl_other_address & ".value = '';" & vbcrlf

          if lcl_form_name = "frmRequestAction" then
             response.write "window.opener.document." & lcl_form_name & ".validstreet.value='Y';" & vbcrlf
          end if
       end if

			 	  if request("sCheckType") = "FinalCheck" then
          if lcl_form_name = "register" then
     				 	  response.write "window.opener.checkDuplicateCitizens( 'FinalUserCheckFailed' );" & vbcrlf
          else
             response.write "window.opener.FinalCheck('FOUND SELECT');" & vbcrlf
          end if
       elseif request("sCheckType") = "FinalCheckValidate" then
        		response.write "window.opener.finalCheckValidate();" & vbcrlf
       end if
     %>
			 	window.close();
			}

			function doKeep() {
<% if lcl_validaddressonly = "N" then %>
 				window.opener.document.<%=lcl_form_name%>.<%=lcl_other_address%>.value = document.frmAddress.oldstnumber.value + ' ' + document.frmAddress.stname.value;
	 			window.opener.document.<%=lcl_form_name%>.residentstreetnumber.value = '';
		 		window.opener.document.<%=lcl_form_name%>.skip_address.selectedIndex = 0;
     <%
       if lcl_form_name = "frmRequestAction" then
          response.write "window.opener." & lcl_form_name & ".validstreet.value='N';" & vbcrlf
       end if

			 	  if request("sCheckType") = "FinalCheck" then
          if lcl_form_name = "register" then
             response.write "window.opener.checkDuplicateCitizens( 'FinalUserCheckFailed' );" & vbcrlf
          else
             response.write "window.opener.FinalCheck('FOUND KEEP');" & vbcrlf
          end if
				   elseif request("sCheckType") = "FinalCheckValidate" then
          response.write "window.opener.finalCheckValidate();" & vbcrlf
       end if
   end if
     %>
			 	window.close();
			}
		//-->
		</script>
</head>
<body>
<div id="content">
  <div id="addresspickcontent">
<font size="+1"><strong>Invalid Address Selection</strong></font><br /><br />
<p class="addresspicker">
  The address you entered does not match any in the system. You can select a valid address from the list or, 
		if you are certain the address you entered is correct, select to use the address you supplied and continue registration.
</p>
<form name="frmAddress" action="addresspicker.asp" method="post">
<strong>The address you entered</strong><br />
  <input type="text" name="oldstnumber" value="<%=request("stnumber")%>" disabled="disabled" size="8" maxlength="10" /> &nbsp; 
		<input type="text" name="stname" value="<%=request("stname")%>" disabled="disabled" size="50" maxlength="50" />
<div id="addresspicklist">
<strong>Valid Address Choices </strong><br />
<%	ShowAddressPicks( request("stname") )	%>
</div>
  <input type="button" class="button" name="validpick" value="Use the valid address selected" onclick="doSelect();" />&nbsp;OR&nbsp;
<%
 'If we are only wanting "valid addresses" to be picked then do not show the "use the address I entered" button.
  if lcl_validaddressonly = "N" then
     lcl_invalidButtonLabel = "Use the address I entered"
  else
     lcl_invalidButtonLabel = "Cancel"
  end if

  response.write "<input type=""button"" class=""button"" name=""invalidpick"" id=""invalidpick"" value=""" & lcl_invalidButtonLabel & """ onclick=""doKeep();"" />" & vbcrlf
%>
<p class="addresspicker">
<strong>If you were saving <%=lcl_message_text%> when you got this window, you will need to resubmit the form.</strong>
<p>
</form>
  </div>
</div>
</body>
</html>
<%
'------------------------------------------------------------------------------
Sub ShowAddressPicks( sStreetName )
	Dim sSql, oAddress

 sSteetName = dbsafe(sStreetName)

	sSQL = "SELECT DISTINCT residentstreetnumber, "
 sSQL = sSQL & " residentstreetname, "
 sSQL = sSQL & " CAST(residentstreetnumber AS INT) AS ordernumb, "
	sSQL = sSQL & " ISNULL(residentstreetprefix,'') AS residentstreetprefix, "
 sSQL = sSQL & " ISNULL(streetsuffix,'') AS streetsuffix, "
 sSQL = sSQL & " ISNULL(streetdirection,'') AS streetdirection "
	sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid = " & iorgid
 sSQL = sSQL & " AND (residentstreetname = '" & sStreetName & "' "
 sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix = '" & sStreetName & "' "
 sSQL = sSQL & " OR residentstreetname + ' ' + streetdirection = '" & sStreetName & "' "
 sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & sStreetName & "' "
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname = '" & sStreetName & "' "
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = '" & sStreetName & "' "
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetdirection = '" & sStreetName & "' "
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & sStreetName & "' "
 sSQL = sSQL & " ) "
 sSQL = sSQL & " AND excludefromactionline = 0 "
	sSQL = sSQL & " ORDER BY 2, 5, 6, 4, 3, 1 "

	Set oAddress = Server.CreateObject("ADODB.Recordset")
	oAddress.Open sSQL, Application("DSN"), 0, 1
	
	if not oAddress.EOF then
  		response.write "<select name=""stnumber"" id=""stnumber"" size=""10"">" & vbcrlf
  		do while NOT oAddress.eof

       sOption = buildStreetAddress(oAddress("residentstreetnumber"), oAddress("residentstreetprefix"), oAddress("residentstreetname"), _
                                    oAddress("streetsuffix"), oAddress("streetdirection"))

    			response.write "  <option value=""" & oAddress("residentstreetnumber") & """>" & sOption & "</option>" & vbcrlf

    			oAddress.movenext
  		loop
  		response.write "</select>" & vbcrlf

	end if
	oAddress.close
	Set oAddress = Nothing
end sub

'------------------------------------------------------------------------------
Function DBsafe( strDB )
Dim sNewString
If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  	sNewString = Replace( strDB, "'", "''" )
'  	sNewString = Replace( sNewString, "<", "&lt;" )
  	DBsafe = sNewString
End Function
%>