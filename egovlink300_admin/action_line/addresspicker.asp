<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: addresspicker.asp
' AUTHOR: Steve Loar
' CREATED: 08/28/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This displays valid addresses for selection
'
' MODIFICATION HISTORY
' 1.0  --/--/----	 ---------- - INITIAL VERSION
' 2.0  04/09/2008  David Boyer - Converted address list to new format
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
  if request.querystring("formname") <> "" then
     lcl_form_name = request.querystring("formname")
  else
     lcl_form_name = "register"
  end if

 'Set up the form variables
  select case lcl_form_name
    case "frmRequestAction"  'Admin Action Line - Create
         lcl_other_address = "ques_issue2"
         lcl_message_text  = "an Action Line Request"
    case "frmlocation"       'Admin Action Line - Edit Request => Issue Location Edit
         lcl_other_address = "ques_issue2"
         lcl_message_text  = "an Action Line Request"
    case "mappoints_maint"   'Admin - Map-Points Maintain
         lcl_other_address = "ques_issue2"
         lcl_message_text  = "a Map-Point"
    case else
         lcl_other_address = "egov_users_useraddress"
         lcl_message_text  = "a registration"
  end select
%>
<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />

		<script language="Javascript">
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
		 		window.opener.document.<%=lcl_form_name%>.<%=lcl_other_address%>.value = '';

     <%
        if lcl_form_name = "frmRequestAction" OR lcl_form_name = "frmlocation" OR lcl_form_name = "mappoints_maint" then
           response.write "window.opener." & lcl_form_name & ".validstreet.value='Y';" & vbcrlf
        end if

        if request("sCheckType") = "FinalCheck" then
           if lcl_form_name = "register" then
              response.write "window.opener.checkDuplicateCitizens('FinalUserCheckFailed');" & vbcrlf
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
 				window.opener.document.<%=lcl_form_name%>.<%=lcl_other_address%>.value = document.frmAddress.oldstnumber.value + ' ' + document.frmAddress.stname.value;
	 			window.opener.document.<%=lcl_form_name%>.residentstreetnumber.value = '';
  <%
     if lcl_form_name = "frmlocation" OR lcl_form_name = "mappoints_maint" then
        response.write "window.opener.document." & lcl_form_name & ".streetaddress.selectedIndex = 0;" & vbcrlf
     else
        response.write "window.opener.document." & lcl_form_name & ".skip_address.selectedIndex = 0;" & vbcrlf
     end if

     if lcl_form_name = "frmRequestAction" OR lcl_form_name = "frmlocation" OR lcl_form_name = "mappoints_maint" then
        response.write "window.opener." & lcl_form_name & ".validstreet.value = 'N';" & vbcrlf
     end if

     if request("sCheckType") = "FinalCheck" then
        if lcl_form_name = "register" then
           response.write "window.opener.checkDuplicateCitizens('FinalUserCheckFailed');" & vbcrlf
        else
           response.write "window.opener.FinalCheck('FOUND KEEP');" & vbcrlf
        end if
     elseif request("sCheckType") = "FinalCheckValidate" then
        response.write "window.opener.finalCheckValidate();" & vbcrlf
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
		<font size="+1"><b>Invalid Address Selection</b></font><br /><br />
		<p class="addresspicker">
			The address you entered does not match any in the system. You can select a valid address from the list or, 
			if you are certain the address you entered is correct, select to use the address you supplied and 
   <% if lcl_form_name <> "frmlocation" then %>
         continue registration.
   <% else %>
         continue edit.
   <% end if %>
		</p>
		<form name="frmAddress" action="addresspicker.asp" method="post">
<script language="javascript">
</script>
			<strong>The address you entered</strong><br />
			<input type="text" name="oldstnumber" value="<%=request("stnumber")%>" disabled="disabled" size="8" maxlength="10" /> &nbsp; 
			<input type="text" name="stname" value="<%=request("stname")%>" disabled="disabled" size="50" maxlength="50" />
			<div id="addresspicklist">
				<strong>Valid Address Choices </strong><br />
    <%	ShowAddressPicks( request("stname") )	%>
			</div>
			<input type="button" class="button" name="validpick" value="Use the valid address selected" onclick="doSelect();" /> &nbsp;OR&nbsp;
			<input type="button" class="button" name="invalidpick" value="Use the address I entered" onclick="doKeep();" />
			<p class="addresspicker">
				<strong>If you were saving <%=lcl_message_text%> when you got this window, you will need to resubmit the form.</strong>
			<p>
		</form>
	</div>
	</div>
	</body>
</html>
<%
'--------------------------------------------------------------------------------------------------
sub ShowAddressPicks( sStreetName )
	dim sSql, oAddress, sOption

	sSQL = "SELECT DISTINCT residentstreetnumber, residentstreetname, CAST(residentstreetnumber AS INT) AS ordernumb, "
	sSQL = sSQL & " ISNULL(residentstreetprefix,'') AS residentstreetprefix, ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid = " & SESSION("orgid")
 sSQL = sSQL & " AND (residentstreetname = '" & dbsafe(sStreetName) & "' "
 sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix = '" & dbsafe(sStreetName) & "' "
 sSQL = sSQL & " OR residentstreetname + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' "
 sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' "
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname = '" & dbsafe(sStreetName) & "' "
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = '" & dbsafe(sStreetName) & "' "
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' "
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & dbsafe(sStreetName) & "'"
 sSQL = sSQL & " ) "
 sSQL = sSQL & " AND excludefromactionline = 0 "
	sSQL = sSQL & " ORDER BY 2, 5, 6, 4, 3, 1 "

	set oAddress = Server.CreateObject("ADODB.Recordset")
	oAddress.Open sSQL, Application("DSN"), 0, 1
	
	if NOT oAddress.eof then
		  response.write "<select id=""stnumber"" name=""stnumber"" size=""10"">" & vbcrlf
    do while NOT oAddress.eof

      'Build the street name
       sOption = buildStreetAddress(oAddress("residentstreetnumber"), oAddress("residentstreetprefix"), oAddress("residentstreetname"), oAddress("streetsuffix"), oAddress("streetdirection"))

    			response.write "<option value=""" & oAddress("residentstreetnumber") & """ >" & sOption & "</option>" & vbcrlf

    			oAddress.MoveNext
    loop

  		response.write "</select>" & vbcrlf

	end if
	oAddress.close
	set oAddress = nothing

end sub

'--------------------------------------------------------------------------------------------------
' Sub ShowAddressPicks( sStreetName )
'--------------------------------------------------------------------------------------------------
'Sub ShowAddressPicks( sStreetName )
'	Dim sSql, oAddress

'	sSql = "Select residentstreetnumber, residentstreetname From egov_residentaddresses Where orgid = " & SESSION("orgid")
'	sSql = sSql & " and residentstreetname = '" & dbsafe(sStreetName) & "' ORDER BY residentstreetname, CAST(residentstreetnumber as integer(4))"

'	Set oAddress = Server.CreateObject("ADODB.Recordset")
'	oAddress.Open sSQL, Application("DSN"), 0, 1
	
'	If not oAddress.EOF Then
'		response.write vbcrlf & "<select id=""stnumber"" name=""stnumber"" size=""10"">"
'		Do While NOT oAddress.EOF 
'			response.write vbcrlf & "<option value=""" & oAddress("residentstreetnumber") & """ >"  
'			response.write oAddress("residentstreetnumber") & "  " & oAddress("residentstreetname") & "</option>"
'			oAddress.MoveNext
'		Loop
'		response.write vbcrlf & "</select>"

'	End If
'	oAddress.close
'	Set oAddress = Nothing

'End Sub
%>