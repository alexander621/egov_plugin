<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitaddresstypeedit.asp
' AUTHOR: Steve Loar
' CREATED: 02/05/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and edits permit address type information.
'
' MODIFICATION HISTORY
' 1.0   02/05/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitAddressTypeid, oRs, sSql, sTitle, sDisabled, sStreetnumber, sStreetprefix, sStreetname, sCity
Dim sState, sZip, sStreetsuffix, sStreettype, sSuite, sUnit, sPin, sSortstreetname, sLegaldescription
Dim sPropertytype, sLandvalue, sTotalvalue, sTaxdistrict, sOwner, sSitus, sSearchName, sSearchStart
Dim sResults, sPermitContactTypeId, sMap

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "permit addresses" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If

If isFeatureOffline("permit addresses") = "Y" Then 
    response.redirect "../admin/outage_feature_offline.asp"
End If 

' GET contact ID
If CLng(request("permitaddresstypeid")) = 0 Then
	iPermitAddressTypeid = 0
	sTitle = "New"
Else
	' EDIT EXISTING contact
	iPermitAddressTypeid = request("permitaddresstypeid")
	sTitle = "Edit"
	sDisabled = GetDisabledText( iPermitAddressTypeid )
End If

'sSql = "SELECT streetnumber, streetprefix, streetname, streetsuffix, streettype, suite, unit, pin, sortstreetname,"
'sSql = sSql & " city, state, zip, legaldescription, propertytype,  ISNULL(landvalue,0.00) AS landvalue,"
'sSql = sSql & " ISNULL(totalvalue,0.00) AS totalvalue, taxdistrict, owner, situs, ISNULL(permitcontacttypeid,0) AS permitcontacttypeid "
'sSql = sSql & " FROM egov_permitaddresstypes WHERE permitaddresstypeid = " & iPermitAddressTypeid 

sSql = "SELECT residentstreetnumber, residentstreetname, residentstreetprefix, parcelidnumber, residentcity, residentstate, residentzip, "
sSql = sSql & " residenttype, suite, residentunit, sortstreetname, ISNULL(legaldescription,'') AS legaldescription, ISNULL(landvalue,0.00) AS landvalue, "
sSql = sSql & " ISNULL(totalvalue,0.00) AS totalvalue, taxdistrict, ISNULL(listedowner,'' AS listedowner, ISNULL(permitownerid,0) AS permitownerid "
sSql = sSql & " FROM egov_residentaddresses WHERE residentaddressid = " & iPermitAddressTypeid

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If NOT oRs.EOF Then
	sStreetnumber = oRs("residentstreetnumber")
	sStreetprefix = oRs("residentstreetprefix")
	sStreetname = oRs("residentstreetname")
	sStreettype = oRs("streettype")
	sSuite = oRs("suite")
	sUnit = oRs("residentunit")
	sPin = oRs("parcelidnumber")
	sSortstreetname = oRs("sortstreetname")
	sCity = oRs("residentcity")
	sState = oRs("residentstate")
	sZip = oRs("residentzip")
	sLegaldescription = Replace(oRs("legaldescription"),Chr(34),"&quot;")
	sPropertytype = oRs("propertytype")
	sLandvalue = FormatNumber(oRs("landvalue"),2,,,0)
	sTotalvalue = FormatNumber(oRs("totalvalue"),2,,,0)
	sTaxdistrict = oRs("taxdistrict")
	sOwner = Replace(oRs("listedowner"),Chr(34),"&quot;")
	sPermitContactTypeId= oRs("permitownerid")
	sResidentType = oRs("residenttype")
Else
	sStreetnumber = ""
	sStreetprefix = ""
	sStreetname = ""
'	sStreetsuffix = ""
	sStreettype = ""
	sSuite = ""
	sUnit = ""
	sPin = ""
	sSortstreetname = ""
	sCity = ""
	sState = ""
	sZip = ""
	sLegaldescription = ""
	sPropertytype = ""
	sLandvalue = ""
	sTotalvalue = ""
	sTaxdistrict = ""
	sOwner = ""
	sPermitContactTypeId = 0
	sResidentType = "R"
End If

oRs.close
Set oRs = Nothing 

%>


<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="JavaScript" src="../scripts/formatnumber.js"></script>
	<script language="JavaScript" src="../scripts/removespaces.js"></script>
	<script language="JavaScript" src="../scripts/removecommas.js"></script>
	<script language="JavaScript" src="../scripts/textareamaxlength.js"></script>

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
			
			for (x=iSearchStart; x < document.frmAddress.permitcontacttypeid.length ; x++)
			{
				optiontext = document.frmAddress.permitcontacttypeid.options[x].text;
				optionchanged = optiontext.toLowerCase();
				if (optionchanged.indexOf(searchchanged) != -1)
				{
					document.frmAddress.permitcontacttypeid.selectedIndex = x;
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
			// check that if a land value is given that it is in money format.
			if (document.frmAddress.landvalue.value != '')
			{
				// Remove any extra spaces
				document.frmAddress.landvalue.value = removeSpaces(document.frmAddress.landvalue.value);
				//Remove commas that would cause problems in validation
				document.frmAddress.landvalue.value = removeCommas(document.frmAddress.landvalue.value);

				rege = /^\d*\.?\d{0,2}$/;
				Ok = rege.test(document.frmAddress.landvalue.value);
				if ( ! Ok )
				{
					alert("The land value must be in currency format or blank.\nPlease correct this and try saving again.");
					document.frmAddress.landvalue.focus();
					return;
				}
				else
				{
					document.frmAddress.landvalue.value = format_number(Number(document.frmAddress.landvalue.value),2);
				}
			}
			// check that if a total value is given that it is in money format.
			if (document.frmAddress.totalvalue.value != '')
			{
				// Remove any extra spaces
				document.frmAddress.totalvalue.value = removeSpaces(document.frmAddress.totalvalue.value);
				//Remove commas that would cause problems in validation
				document.frmAddress.totalvalue.value = removeCommas(document.frmAddress.totalvalue.value);

				rege = /^\d*\.?\d{0,2}$/;
				Ok = rege.test(document.frmAddress.totalvalue.value);
				if ( ! Ok )
				{
					alert("The total value must be in currency format or blank.\nPlease correct this and try saving again.");
					document.frmAddress.totalvalue.focus();
					return;
				}
				else
				{
					document.frmAddress.totalvalue.value = format_number(Number(document.frmAddress.totalvalue.value),2);
				}
			}
			// Check the length of the legal description
			if (document.frmAddress.legaldescription.value != '')
			{
				if (document.frmAddress.legaldescription.value.length >= document.frmAddress.legaldescription.getAttribute('maxlength'))
				{
					alert("The legal description has a limit of 400 characters which you have exceeded.\nPlease correct this and try saving again.");
					document.frmAddress.legaldescription.focus();
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
				location.href="permitaddresstypedelete.asp?permitaddresstypeid=<%=iPermitAddressTypeid%>";
			}
		}


	//-->
	</script>

</head>

<body onload="setMaxLength();">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	
	<!--BEGIN: PAGE TITLE-->
	<p>
		<font size="+1"><strong><%=sTitle%> Permit Address</strong></font><br /><br />
		<a href="permitaddresstypelist.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;<%=langBackToStart%></a>
	</p>
	<!--END: PAGE TITLE-->


	<!--BEGIN: EDIT FORM-->
	<%		If CLng(iPermitAddressTypeid) = CLng(0) Then %>
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Create" /><br />
	<%		Else %>
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Validate();" value="Update" /> &nbsp; &nbsp;
				<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:Delete();" value="Delete" <%=sDisabled%> /><br />
	<%		End If %>

	<form name="frmAddress" action="permitaddresstypeupdate.asp" method="post">
		<input type="hidden" name="permitaddresstypeid" value="<%=iPermitAddressTypeid%>" />

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
				<td align="right" class="labelcolumn">Street Suffix:</td><td class="datacolumn"><input type="text" name="streetsuffix" value="<%=sStreetsuffix%>" size="15" maxlength="15" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Street Type:</td><td class="datacolumn"><input type="text" name="streettype" value="<%=sStreettype%>" size="5" maxlength="5" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Suite:</td><td class="datacolumn"><input type="text" name="suite" value="<%=sSuite%>" size="15" maxlength="15" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Unit:</td><td class="datacolumn"><input type="text" name="unit" value="<%=sUnit%>" size="10" maxlength="10" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Parcel Id No:</td><td class="datacolumn"><input type="text" name="pin" value="<%=sPin%>" size="10" maxlength="10" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Sortstreetname:</td><td class="datacolumn"><input type="text" name="sortstreetname" value="<%=sSortstreetname%>" size="50" maxlength="50" /></td>
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
				<td align="right" class="labelcolumn" nowrap="nowrap" valign="top">Legal Description:</td>
				<td class="datacolumn">
					<textarea id="legaldescription" name="legaldescription" rows="5" cols="80" maxlength="400" wrap="soft"><%=sLegaldescription%></textarea>
				</td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Property Type:</td><td class="datacolumn"><input type="text" name="propertytype" value="<%=sPropertytype%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Land Value:</td><td class="datacolumn"><input type="text" name="landvalue" value="<%=sLandvalue%>" size="15" maxlength="15" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Total Value:</td><td class="datacolumn"><input type="text" name="totalvalue" value="<%=sTotalvalue%>" size="15" maxlength="15" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Tax District:</td><td class="datacolumn"><input type="text" name="taxdistrict" value="<%=sTaxdistrict%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn" valign="top">Listed Owner:</td><td class="datacolumn">
					<textarea id="owner" name="owner" rows="3" cols="80" maxlength="250" wrap="soft"><%=sOwner%></textarea>
				</td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn" valign="top">Owner Contact:</td>
				<td class="datacolumn">
					<select name="permitownerid" onchange="javascript:UserPick();">
						<option value="0">Select an owner from the list</option>
						<% ShowOwnerDropDown( sPermitContactTypeId )%>
					</select>
					<br />Name Search: <input type="text" name="searchname" value="<%=sSearchName%>" size="25" maxlength="50" onchange="javascript:ClearSearch();" />
					<input type="button" class="button ui-button ui-widget ui-corner-all" value="Search" onclick="javascript:SearchCitizens(document.frmAddress.searchstart.value);" />
					<input type="hidden" name="results" value="" /><input type="hidden" name="searchstart" value="<%=sSearchStart%>" />
					<span id="searchresults"><%=sResults%></span>
					<br /><div id="searchtip">(last name, first name)</div>									
				</td>
			</tr>
		</table>
		</div>
	</form>
	<!--END: EDIT FORM-->

	</div>
</div>
<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function GetDisabledText( iPermitAddressTypeid )
'--------------------------------------------------------------------------------------------------
Function GetDisabledText( iPermitAddressTypeid )
	Dim sSql, oRs

	'If this contact is used, keep it from being deleted

	sSql = "SELECT COUNT(permitaddresstypeid) AS hits FROM egov_permitaddress "
	sSql = sSql & " WHERE permitaddresstypeid = " & iPermitAddressTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

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
	oRs.Open sSQL, Application("DSN"), 3, 1

	Do While Not oRs.eof 
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



%>


