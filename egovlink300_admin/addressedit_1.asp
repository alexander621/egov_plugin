<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: addressedit.asp
' AUTHOR: Steve Loar
' CREATED: 04/04/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and edits permit address type information.
'
' MODIFICATION HISTORY
' 1.0   04/04/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iResidentAddressId, oRs, sSql, sTitle, sDisabled, sStreetnumber, sStreetprefix, sStreetname, sCity
Dim sState, sZip, sStreetsuffix, sStreettype, sUnit, sPin, sLegaldescription, sCounty, sActionLineName
Dim sPropertytype, sLandvalue, sTotalvalue, sTaxdistrict, sOwner, sSitus, sSearchName, sSearchStart
Dim sResults, sPermitContactTypeId, sMap, sLatitude, sLongitude, iRegisteredUserId, sStreetDirection
Dim bExcludeFromActionLine, sPagenum, sPageValue, sKeyword, sKeywordValue, oAddressOrg

sLevel = "" ' Override of value from common.asp

PageDisplayCheck "address list", sLevel	' In common.asp

Set oAddressOrg = New classOrganization

If request("pagenum") <> "" Then 
	sPagenum = "?pagenum=" & request("pagenum")
	sPageValue = request("pagenum")
Else
	sPagenum = ""
	sPageValue = ""
End If 

If request("keyword") <> "" Then
	If sPagenum <> "" Then
		sKeyword = "&"
	Else
		sKeyword = "?"
	End If 
	sKeyword = sKeyword & "keyword=" & request("keyword")
	sKeywordValue = request("keyword")
Else
	sKeyword = ""
	sKeywordValue = "" 
End If 

' GET contact ID
If CLng(request("residentaddressid")) = CLng(0) Then
	iResidentAddressId = 0
	sTitle = "New"
Else
	' EDIT EXISTING address
	iResidentAddressId = request("residentaddressid")
	sTitle = "Edit"
	'sDisabled = GetDisabledText( iResidentAddressId )
	sDisabled = ""
End If

sSql = "SELECT residentstreetnumber, residentstreetname, residentstreetprefix, parcelidnumber, residentcity, residentstate, residentzip, latitude, longitude, "
sSql = sSql & " residenttype, streetsuffix, residentunit, ISNULL(legaldescription,'') AS legaldescription, county, streetdirection, "
sSql = sSql & " ISNULL(listedowner,'') AS listedowner, ISNULL(registereduserid,0) AS registereduserid, excludefromactionline "
sSql = sSql & " FROM egov_residentaddresses WHERE residentaddressid = " & iResidentAddressId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 0, 1

If NOT oRs.EOF Then
	sStreetnumber = oRs("residentstreetnumber")
	sStreetprefix = oRs("residentstreetprefix")
	sStreetname = oRs("residentstreetname")
	sStreetSuffix = oRs("streetsuffix")
	sStreetDirection = oRs("streetdirection")
	sUnit = oRs("residentunit")
	sPin = oRs("parcelidnumber")
	sCity = oRs("residentcity")
	sState = oRs("residentstate")
	sZip = oRs("residentzip")
	sLegaldescription = Replace(oRs("legaldescription"),Chr(34),"&quot;")
	sOwner = oRs("listedowner")
	iRegisteredUserId = oRs("registereduserid")
	sResidentType = oRs("residenttype")
	sLatitude = oRs("latitude")
	sLongitude = oRs("longitude")
	sCounty = oRs("county")
	bExcludeFromActionLine = oRs("excludefromactionline")
Else
	sStreetnumber = ""
	sStreetprefix = ""
	sStreetname = ""
	sStreettype = ""
	sStreetSuffix = ""
	sStreetDirection = ""
	sUnit = ""
	sPin = ""
	sCity = ""
	sState = ""
	sZip = ""
	sLegaldescription = ""
	sPropertytype = ""
	sOwner = ""
	sPermitContactTypeId = 0
	iRegisteredUserId = 0
	sResidentType = "R"
	sLatitude = ""
	sLongitude = ""
	sCounty = ""
	bExcludeFromActionLine = False 
End If

oRs.close
Set oRs = Nothing 

%>


<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="global.css" />
	<link rel="stylesheet" type="text/css" href="permits/permits.css" />

	<script language="JavaScript" src="scripts/formatnumber.js"></script>
	<script language="JavaScript" src="scripts/removespaces.js"></script>
	<script language="JavaScript" src="scripts/removecommas.js"></script>
	<script language="JavaScript" src="scripts/textareamaxlength.js"></script>

	<script language="Javascript">
	<!--

		function SearchCitizens( iSearchStart )
		{
			var optiontext;
			var optionchanged;
			//alert(document.BuyerForm.searchname.value);
			var searchtext = document.frmAddress.searchname.value;
			var searchchanged = searchtext.toLowerCase();

			iSearchStart = parseInt(iSearchStart) + 1;
			
			for (x=iSearchStart; x < document.frmAddress.registereduserid.length ; x++)
			{
				optiontext = document.frmAddress.registereduserid.options[x].text;
				optionchanged = optiontext.toLowerCase();
				if (optionchanged.indexOf(searchchanged) != -1)
				{
					document.frmAddress.registereduserid.selectedIndex = x;
					document.frmAddress.results.value = 'Possible Match Found.';
					document.getElementById('searchresults').innerHTML = 'Possible Match Found.';
					document.frmAddress.searchstart.value = x;
					return;
				}
			}
			document.frmAddress.results.value = 'No Match Found.';
			document.getElementById('searchresults').innerHTML = 'No Match Found.';
			document.frmAddress.searchstart.value = -1;
		}

		function ClearSearch()
		{
			document.frmAddress.searchstart.value = -1;
		}

		function UserPick()
		{
			document.frmAddress.searchname.value = '';
			document.frmAddress.results.value = '';
			document.getElementById('searchresults').innerHTML = '';
			document.frmAddress.searchstart.value = -1;
		}

		function Validate()
		{
			var rege;
			var Ok;

			// Check that a street number is provided
			if (document.frmAddress.streetnumber.value == '')
			{
				alert('A street number is required.\nPlease correct this and try saving again.');
				document.frmAddress.streetnumber.focus();
				return;
			}
			else
			{
				rege = /^\d+$/;
				Ok = rege.test(document.frmAddress.streetnumber.value);
				if ( ! Ok )
				{
					alert("The street number must be a whole number value.\nPlease correct this and try saving again.");
					document.frmAddress.streetnumber.focus();
					return;
				}
			}

			// Check that a street name is provided
			if (document.frmAddress.streetname.value == '')
			{
				alert('A street name is required.\nPlease correct this and try saving again.');
				document.frmAddress.streetname.focus();
				return;
			}

			// check that if a Parcel ID is given that it is numeric.
			if (document.frmAddress.pin.value != '')
			{
				rege = /^\d+$/;
				Ok = rege.test(document.frmAddress.pin.value);
				if ( ! Ok )
				{
					alert("The parcel id number must be a whole number value.\nPlease correct this and try saving again.");
					document.frmAddress.pin.focus();
					return;
				}
			}

			// Check the length of the legal description
			if (document.frmAddress.legaldescription.value != '')
			{
				if (document.frmAddress.legaldescription.value.length >= document.frmAddress.legaldescription.getAttribute('maxlength'))
				{
					alert("The legal description has a limit of " + document.frmAddress.legaldescription.getAttribute('maxlength') + " characters which you have exceeded.\nPlease correct this and try saving again.");
					document.frmAddress.legaldescription.focus();
					return;
				}
			}

			// check the Latitude
			if (document.frmAddress.latitude.value.length > 0)
			{
				rege = /^-?\d{1,3}\.\d+$/;
				Ok = rege.test(document.frmAddress.latitude.value);

				if (! Ok)
				{
					alert("The latitude must be a number, or blank\n and in the range 90 to -90.");
					document.frmAddress.latitude.focus();
					return;
				}
				else
				{
					if (document.frmAddress.latitude.value > 90 || document.frmAddress.latitude.value < -90)
					{
						alert("The latitude must be a number, or blank\n and in the range 90 to -90.");
						document.frmAddress.latitude.focus();
						return;
					}
				}
			}
			// check the Longitude
			if (document.frmAddress.longitude.value.length > 0)
			{
				rege = /^-?\d{1,3}\.\d+$/;
				Ok = rege.test(document.frmAddress.longitude.value);

				if (! Ok)
				{
					alert("The longitude must be a number, or blank\n and in the range 180 to -180");
					document.frmAddress.longitude.focus();
					return;
				}
				else
				{
					if (document.frmAddress.longitude.value > 180 || document.frmAddress.longitude.value < -180)
					{
						alert("The longitude must be a number, or blank\n and in the range 180 to -180.");
						document.frmAddress.longitude.focus();
						return;
					}
				}
			}

			// Check the length of the listed owner
			if (document.frmAddress.listedowner.value != '')
			{
				if (document.frmAddress.listedowner.value.length >= document.frmAddress.listedowner.getAttribute('maxlength'))
				{
					alert("The listed owner has a limit of " + document.frmAddress.listedowner.getAttribute('maxlength') + " characters which you have exceeded.\nPlease correct this and try saving again.");
					document.frmAddress.listedowner.focus();
					return;
				}
			}
			// alert('Ok');
			document.frmAddress.submit();
		}

		function Delete() 
		{
			if (confirm("Do you wish to delete this address?"))
			{
				location.href='addressdelete.asp?residentaddressid=<%=iResidentAddressId%>';
			}
		}

		function NewPermit()
		{
			location.href='newpermit.asp?permitaddresstypeid=<%=iResidentAddressId%>';
		}


	//-->
	</script>

</head>

<body onload="setMaxLength();">

<% ShowHeader sLevel %>
<!--#Include file="menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	
	<!--BEGIN: PAGE TITLE-->
	<p>
		<font size="+1"><strong><%=sTitle%> Address</strong></font><br /><br />
		<a href="manage_address_list.asp<%=sPagenum%><%=sKeyword%>"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
	</p>
	<!--END: PAGE TITLE-->


	<!--BEGIN: EDIT FORM-->
	<%		If CLng(iResidentAddressId) = CLng(0) Then %>
				<input type="button" class="button" onclick="javascript:Validate();" value="Create Address" /><br />
	<%		Else %>
				<input type="button" class="button" onclick="javascript:Validate();" value="Save Changes" /> &nbsp; &nbsp;
				<br />
	<%		End If %>

	<form name="frmAddress" action="addressupdate.asp" method="post">
		<input type="hidden" name="residentaddressid" value="<%=iResidentAddressId%>" />
		<input type="hidden" name="searchtext" value="<%=request("searchtext")%>" />
		<input type="hidden" name="searchfield" value="<%=request("searchfield")%>" />
		<input type="hidden" name="pagenum" value="<%=sPageValue%>" />
		<input type="hidden" name="keyword" value="<%=sKeywordValue%>" />

		<div class="shadow">
		<table id="permitaddressinfo" cellpadding="2" cellspacing="0" border="0">
			<tr>
				<td align="right" class="labelcolumn">Street Number:</td><td class="datacolumn"><input type="text" name="streetnumber" value="<%=sStreetnumber%>" size="10" maxlength="10" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Street Prefix:</td><td class="datacolumn"><input type="text" name="streetprefix" value="<%=sStreetprefix%>" size="15" maxlength="15" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Street Name:</td><td class="datacolumn"><input type="text" name="streetname" value="<%=sStreetname%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Street Suffix:</td><td class="datacolumn"><input type="text" name="streetsuffix" value="<%=sStreetSuffix%>" size="15" maxlength="15" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Street Direction:</td><td class="datacolumn"><input type="text" name="streetdirection" value="<%=sStreetDirection%>" size="10" maxlength="10" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Unit/Suite:</td><td class="datacolumn"><input type="text" name="unit" value="<%=sUnit%>" size="10" maxlength="10" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Parcel Id No:</td><td class="datacolumn"><input type="text" name="pin" value="<%=sPin%>" size="10" maxlength="10" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">City:</td><td class="datacolumn"><input type="text" name="city" value="<%=sCity%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">State:</td><td class="datacolumn"><input type="text" name="state" value="<%=sState%>" size="2" maxlength="2" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Zip:</td><td class="datacolumn"><input type="text" name="zip" value="<%=sZip%>" size="10" maxlength="10" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn"><%=oAddressOrg.GetOrgDisplayName("address grouping field")%>:</td><td class="datacolumn"><input type="text" name="county" value="<%=sCounty%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn" nowrap="nowrap" valign="top">Legal Description:</td>
				<td class="datacolumn">
					<textarea id="legaldescription" name="legaldescription" rows="5" cols="80" maxlength="400" wrap="soft"><%=sLegaldescription%></textarea>
				</td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Property Type:</td>
				<td class="datacolumn">
					<select name="residenttype" >
<%						ShowResidencyTypePicks sResidentType %>
					</select>
				</td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Latitude:</td>
				<td class="datacolumn"><input type="text" name="latitude" size="11" maxlength="11" value="<%=sLatitude%>" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Longitude:</td>
				<td class="datacolumn"><input type="text" name="longitude" size="11" maxlength="11" value="<%=sLongitude%>" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn" valign="top">Listed Owner:</td><td class="datacolumn">
					<textarea id="owner" name="listedowner" rows="3" cols="80" maxlength="250" wrap="soft"><%=sOwner%></textarea>
				</td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn" valign="top">Registered User (owner):</td>
				<td class="datacolumn">
					<select name="registereduserid" onchange="javascript:UserPick();">
						<option value="0">Select a registered user from the list</option>
						<% 
							ShowUserDropDown iRegisteredUserId 
						%>
					</select>
					<br />Name Search: <input type="text" name="searchname" value="<%=sSearchName%>" size="25" maxlength="50" onchange="javascript:ClearSearch();" />
					<input type="button" class="button" value="Search" onclick="javascript:SearchCitizens(document.frmAddress.searchstart.value);" />
					<input type="hidden" name="results" value="" /><input type="hidden" name="searchstart" value="<%=sSearchStart%>" />
					<span id="searchresults"><%=sResults%></span>
					<br /><br />	
			<% If OrgHasFeature( "action line" ) Then 
				sActionLineName = GetFeatureName( "action line" )
					%>					
				</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td class="datacolumn">
					<input type="checkbox" name="excludefromactionline" 
																<% If bExcludeFromActionLine Then 
																		response.write " checked=""checked"" "
																	End If %> /> Exclude From <%= sActionLineName %>
			<%	Else %>
					<input type="hidden" name="excludefromactionline" value="off" />
			<%	End If  %>
				</td>
			</tr>
		</table>
		</div>
	</form>
	<!--END: EDIT FORM-->

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="admin_footer.asp"-->  

</body>
</html>


<%
Set oAddressOrg = Nothing 
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function GetDisabledText( iPermitAddressTypeid )
'--------------------------------------------------------------------------------------------------
Function GetDisabledText( iPermitAddressTypeid )
	Dim sSql, oRs

	'If this contact is used, keep it from being deleted

	sSql = "SELECT COUNT(residentaddressid) AS hits FROM egov_permitaddress "
	sSql = sSql & " WHERE residentaddressid = " & iPermitAddressTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			GetDisabledText = " disabled=""disabled"" " 
		Else
			GetDisabledText = "" 
		End If 
	Else
		GetDisabledText = "" 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' Sub ShowOwnerDropDown( sPermitContactTypeId )
'--------------------------------------------------------------------------------------------------
Sub ShowOwnerDropDown( sPermitContactTypeId )
	Dim sSql, oRs

	sSql = "SELECT permitcontacttypeid, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname,"
	sSql = sSql & " ISNULL(company,'') AS company, LOWER(ISNULL(lastname,company)) AS sortname FROM egov_permitcontacttypes "
	sSql = sSql & " WHERE orgid = "& session("orgid" ) & " ORDER BY 5, 2"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("permitcontacttypeid") & """"
		If CLng(sPermitContactTypeId) = CLng(oRs("permitcontacttypeid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">"
		If oRs("lastname") <> "" Then
			response.write oRs("lastname") & ", " & oRs("firstname") 
			If oRs("company") <> "" Then
				response.write " (" & oRs("company") & ")"
			End If 
		Else
			response.write oRs("company")
		End If 
		response.write "</option>"
		oRs.MoveNext
	Loop 
		
	oRs.close
	Set oRs = Nothing

End Sub  


'--------------------------------------------------------------------------------------------------
' Sub ShowResidencyTypePicks( sResidentType )
'--------------------------------------------------------------------------------------------------
Sub ShowResidencyTypePicks( sResidentType )
	Dim sSql, oRs

	sSql = "SELECT resident_type, description FROM egov_poolpassresidenttypes WHERE isforaddresses = 1 AND orgid = " & session("orgid")
	sSql = sSql & " ORDER BY description"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("resident_type") & """"
		If UCase(sResidentType) = UCase(oRs("resident_type")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">"
		response.write oRs("description")
		response.write "</option>"
		oRs.MoveNext
	Loop 
		
	oRs.close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowUserDropDown( iUserId )
'--------------------------------------------------------------------------------------------------
Sub ShowUserDropDown( iUserId )
	Dim oCmd, oResident

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserWithAddressList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgid", 3, 1, 4, Session("OrgID"))
	    Set oResident = .Execute
	End With

	Do While Not oResident.eof 
		response.write vbcrlf & "<option value=""" & oResident("userid") & """"
		If CLng(iUserId) = CLng(oResident("userid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oResident("userlname") & ", " & oResident("userfname")
		If Not IsNull(oResident("useraddress")) And oResident("useraddress") <> "" Then
			response.write " &ndash; " & oResident("useraddress")
		End If 
		response.write "</option>"
		oResident.movenext
	Loop 
		
	oResident.close
	Set oResident = Nothing
	Set oCmd = Nothing
End Sub  



%>


