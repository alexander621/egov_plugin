<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: residentaddressedit.asp
' AUTHOR: Steve Loar
' CREATED: 04/02/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and edits address information.
'
' MODIFICATION HISTORY
' 1.0   04/02/2008	Steve Loar - INITIAL VERSION
' 1.1	07/21/2009	Steve Loar - New fields for Lansing IL
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
 dim iResidentAddressid, lcl_copyid, oRs, sSql, sTitle, sDisabled, sStreetnumber, sStreetprefix
 dim sStreetname, sCity, sState, sZip, sStreetsuffix, sStreettype, sUnit, sPin, sLegaldescription
 dim sCounty, sStreetDirection, sPropertytype, sLandvalue, sTotalvalue, sTaxdistrict, sOwner
 dim sSitus, sSearchName, sSearchStart, sResults, sPermitContactTypeId, sMap, sLatitude, sLongitude
 dim iRegisteredUserId, sSaveButtonText, oAddressOrg, bNewAddress
 dim sPropertyTaxNumber, sLotNumber, sLotWidth, sLotLength, sBlockNumber, sSubdivision, sSection
 dim sTownship, sRange, sPermanentRealEstateIndexNumber, sCollectorsTaxBillVolumeNumber

 sLevel = "../"  'Override of value from common.asp

 set oAddressOrg = New classOrganization

 iResidentAddressid = clng(request("residentaddressid"))
 lcl_copyid         = clng(0)

 if iResidentAddressid > clng(0) then
	   sTitle          = "Edit"
   	sSaveButtonText = "Save Changes"
   	bNewAddress     = False
 else
	   sTitle          = "New"
   	sSaveButtonText = "Create"
   	bNewAddress     = True

    if request("copyid") <> "" then
       lcl_copyid      = clng(request("copyid"))
       sTitle          = "Copy"
       sSaveButtonText = sSAveButtonText & " Copy"
    end if
 end if

 if sTitle <> "" then
    sTitle = sTitle & " Address"
 end if

 sSQL = "SELECT residentstreetnumber, "
 sSQL = sSQL & " residentstreetname, "
 sSQL = sSQL & " residentstreetprefix, "
 sSQL = sSQL & " streetdirection, "
 sSQL = sSQL & " parcelidnumber, "
 sSQL = sSQL & " residentcity, "
 sSQL = sSQL & " residentstate, "
 sSQL = sSQL & " residentzip, "
 sSQL = sSQL & " county, "
 sSQL = sSQL & " latitude, "
 sSQL = sSQL & " longitude, "
 sSQL = sSQL & " residenttype, "
 sSQL = sSQL & " streetsuffix, "
 sSQL = sSQL & " residentunit, "
 sSQL = sSQL & " ISNULL(legaldescription,'') AS legaldescription, "
 sSQL = sSQL & " ISNULL(listedowner,'') AS listedowner, "
 sSQL = sSQL & " ISNULL(registereduserid,0) AS registereduserid, "
 sSQL = sSQL & " ISNULL(propertytaxnumber,'') AS propertytaxnumber, "
 sSQL = sSQL & " ISNULL(lotnumber,'') AS lotnumber, "
 sSQL = sSQL & " ISNULL(lotwidth,'') AS lotwidth, "
 sSQL = sSQL & " ISNULL(lotlength,'') AS lotlength, "
 sSQL = sSQL & " ISNULL(blocknumber, '') AS blocknumber, "
 sSQL = sSQL & " ISNULL(subdivision,'') AS subdivision, "
 sSQL = sSQL & " ISNULL(section,'') AS section, "
 sSQL = sSQL & " ISNULL(township,'') AS township, "
 sSQL = sSQL & " ISNULL(range,'') AS range, "
 sSQL = sSQL & " ISNULL(permanentrealestateindexnumber,'') AS permanentrealestateindexnumber, "
 sSQL = sSQL & " ISNULL(collectorstaxbillvolumenumber,'') AS collectorstaxbillvolumenumber "
 sSQL = sSQL & " FROM egov_residentaddresses "

 if lcl_copyid > 0 then
    sSQL = sSQL & " WHERE residentaddressid = " & lcl_copyid
 else
    sSQL = sSQL & " WHERE residentaddressid = " & iResidentAddressid
 end if

 set oRs = Server.CreateObject("ADODB.Recordset")
 oRs.Open sSQL, Application("DSN"), 3, 1

 if not oRs.eof then
   	sStreetnumber      = oRs("residentstreetnumber")
   	sStreetprefix      = oRs("residentstreetprefix")
   	sStreetname        = oRs("residentstreetname")
   	sStreetSuffix      = oRs("streetsuffix")
   	sUnit              = oRs("residentunit")
   	sPin               = oRs("parcelidnumber")
   	sCity              = oRs("residentcity")
   	sState             = oRs("residentstate")
   	sZip               = oRs("residentzip")
   	sCounty            = oRs("county")
   	sLegaldescription  = replace(oRs("legaldescription"),chr(34),"&quot;")
   	sOwner             = oRs("listedowner")
   	iRegisteredUserId  = oRs("registereduserid")
   	sResidentType      = oRs("residenttype")
   	sLatitude          = oRs("latitude")
   	sLongitude         = oRs("longitude")
   	sStreetDirection   = oRs("streetdirection")
   	sPropertyTaxNumber = oRs("propertytaxnumber")
   	sLotNumber         = oRs("lotnumber")
   	sLotWidth          = oRs("lotwidth")
   	sLotLength         = oRs("lotlength")
   	sBlockNumber       = oRs("blocknumber")
   	sSubdivision       = oRs("subdivision")
   	sSection           = oRs("section")
   	sTownship          = oRs("township")
   	sRange             = oRs("range")
   	sPermanentRealEstateIndexNumber = oRs("permanentrealestateindexnumber")
   	sCollectorsTaxBillVolumeNumber  = oRs("collectorstaxbillvolumenumber")
 else
  		sStreetnumber      = ""
  		sStreetprefix      = ""
  		sStreetname        = ""
  		sStreetSuffix      = ""
  		sUnit              = ""
  		sPin               = ""
  		sCity              = ""
  		sState             = ""
  		sZip               = ""
  		sCounty            = ""
  		sLegaldescription  = ""
  		sOwner             = ""
  		iRegisteredUserId  = 0
  		sResidentType      = ""
  		sLatitude          = ""
  		sLongitude         = ""
  		sStreetDirection   = ""
  		sPropertyTaxNumber = ""
  		sLotNumber         = ""
  		sLotWidth          = ""
  		sLotLength         = ""
  		sBlockNumber       = ""
  		sSubdivision       = ""
  		sSection           = ""
  		sTownship          = ""
  		sRange             = ""
  		sPermanentRealEstateIndexNumber = ""
  		sCollectorsTaxBillVolumeNumber  = ""
 end if

 oRs.close
 set oRs = nothing 
%>
<html>
<head>
	<title>E-Gov Administration Console { Maintain Address }</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script type="text/javascript" src="../scripts/formatnumber.js"></script>
	<script type="text/javascript" src="../scripts/removespaces.js"></script>
	<script type="text/javascript" src="../scripts/removecommas.js"></script>
	<script type="text/javascript" src="../scripts/textareamaxlength.js"></script>
	<script type="text/javascript" src="../scripts/ajaxLib.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

	<script type="text/javascript">
	<!--

		function SearchCitizens( iSearchStart )
		{
			var optiontext;
			var optionchanged;
			//alert(document.BuyerForm.searchname.value);
			var searchtext = document.frmAddress.searchname.value;
			var searchchanged = searchtext.toLowerCase();

			iSearchStart = parseInt(iSearchStart) + 1;
			
			for (x=iSearchStart; x < document.frmAddress.permitownerid.length ; x++)
			{
				optiontext = document.frmAddress.permitownerid.options[x].text;
				optionchanged = optiontext.toLowerCase();
				if (optionchanged.indexOf(searchchanged) != -1)
				{
					document.frmAddress.permitownerid.selectedIndex = x;
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

		function Validate() {
 			var rege;
	 		var Ok;

		 	// Check that a street number is provided
			 if(document.frmAddress.residentstreetnumber.value == '') {
    			alert('A street number is required.\nPlease correct this and try saving again.');
    			document.frmAddress.residentstreetnumber.focus();
    			return;
 			} else {
    			rege = /^\d+$/;
   				Ok = rege.test(document.frmAddress.residentstreetnumber.value);
			   	if( ! Ok ) {
     					alert("The street number must be a whole number value.\nPlease correct this and try saving again.");
			     		document.frmAddress.residentstreetnumber.focus();
     					return;
   				}
			}

			// Check that a street name is provided
			if (document.frmAddress.residentstreetname.value == '')
			{
				alert('A street name is required.\nPlease correct this and try saving again.');
				document.frmAddress.residentstreetname.focus();
				return;
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
			//alert('Ok');
			//document.frmAddress.submit();

			var sParameter = 'residentaddressid=' + encodeURIComponent(document.frmAddress.residentaddressid.value);
			sParameter += '&residentstreetnumber=' + encodeURIComponent(document.frmAddress.residentstreetnumber.value);
			sParameter += '&residentstreetname=' + encodeURIComponent(document.frmAddress.residentstreetname.value);
			sParameter += '&residentstreetprefix=' + encodeURIComponent(document.frmAddress.residentstreetprefix.value);
			sParameter += '&streetsuffix=' + encodeURIComponent(document.frmAddress.streetsuffix.value);
			sParameter += '&residentunit=' + encodeURIComponent(document.frmAddress.residentunit.value);
			sParameter += '&parcelidnumber=' + encodeURIComponent(document.frmAddress.parcelidnumber.value);
			sParameter += '&residentcity=' + encodeURIComponent(document.frmAddress.residentcity.value);
			sParameter += '&residentstate=' + encodeURIComponent(document.frmAddress.residentstate.value);
			sParameter += '&residentzip=' + encodeURIComponent(document.frmAddress.residentzip.value);
			sParameter += '&county=' + encodeURIComponent(document.frmAddress.county.value);
			sParameter += '&legaldescription=' + encodeURIComponent(document.frmAddress.legaldescription.value);
			sParameter += '&residenttype=' + encodeURIComponent(document.frmAddress.residenttype.value);
			sParameter += '&latitude=' + encodeURIComponent(document.frmAddress.latitude.value);
			sParameter += '&longitude=' + encodeURIComponent(document.frmAddress.longitude.value);
			sParameter += '&listedowner=' + encodeURIComponent(document.frmAddress.listedowner.value);
			sParameter += '&registereduserid=' + encodeURIComponent(document.frmAddress.registereduserid.value);
			sParameter += '&streetdirection=' + encodeURIComponent(document.frmAddress.streetdirection.value);
<%			If OrgHasFeature( "extended address fields" ) Then %>
				sParameter += '&propertytaxnumber=' + encodeURIComponent(document.frmAddress.propertytaxnumber.value);
				sParameter += '&lotnumber=' + encodeURIComponent(document.frmAddress.lotnumber.value);
				sParameter += '&lotwidth=' + encodeURIComponent(document.frmAddress.lotwidth.value);
				sParameter += '&lotlength=' + encodeURIComponent(document.frmAddress.lotlength.value);
				sParameter += '&blocknumber=' + encodeURIComponent(document.frmAddress.blocknumber.value);
				sParameter += '&subdivision=' + encodeURIComponent(document.frmAddress.subdivision.value);
				sParameter += '&section=' + encodeURIComponent(document.frmAddress.section.value);
				sParameter += '&township=' + encodeURIComponent(document.frmAddress.township.value);
				sParameter += '&range=' + encodeURIComponent(document.frmAddress.range.value);
				sParameter += '&permanentrealestateindexnumber=' + encodeURIComponent(document.frmAddress.permanentrealestateindexnumber.value);
				sParameter += '&collectorstaxbillvolumenumber=' + encodeURIComponent(document.frmAddress.collectorstaxbillvolumenumber.value);
<%			End If %>

			// Do the ajax call
			doAjax('permitaddressupdate.asp', sParameter, 'CloseThisSaved', 'post', '0');
		}

		function CloseThisSaved( sResult )
		{
			//alert( sResult ); 
			doClose();
		}

		function Delete() 
		{
			if (confirm("Do you wish to delete this address?"))
			{
				location.href='permitaddresstypedelete.asp?permitaddresstypeid=<%=iPermitAddressTypeid%>&searchtext=<%=request("searchtext")%>&searchfield=<%=request("searchfield")%>';
			}
		}

		function NewPermit()
		{
			location.href='newpermit.asp?permitaddresstypeid=<%=iPermitAddressTypeid%>';
		}

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		function SelectListedOwner( )
		{
			var w = (screen.width - 600)/2;
			var h = (screen.height - 300)/2;
			//winHandle = eval('window.open("../listedownerpicker.asp", "_contact", "width=600,height=300,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			showModal('../listedownerpicker.asp', 'Listed Owner Selection', 50, 50);
		}

	//-->
	</script>

</head>

<body onload="setMaxLength();document.getElementById('residentstreetnumber').focus();">

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	
	<!--BEGIN: PAGE TITLE-->
<script>parent.document.getElementById('modaltitle'+window.frameElement.getAttribute("data-close")).innerHTML='<%=sTitle%>';</script>
	<!--END: PAGE TITLE-->

	<!--BEGIN: EDIT FORM-->
	<p>
		<input type="button" id="cancelButton" class="button ui-button ui-widget ui-corner-all" value="Cancel" onclick="doClose();" />
		<input type="button" id="savebutton" class="button ui-button ui-widget ui-corner-all" value="<%=sSaveButtonText%>" onclick="Validate();" />
	</p>

	<form name="frmAddress" action="permitaddressupdate.asp" method="post">
		<input type="hidden" id="residentaddressid" name="residentaddressid" value="<%=iResidentAddressid%>" />

		<table id="permitaddressinfo" cellpadding="2" cellspacing="0" border="0">
			<tr>
				<td align="right" class="labelcolumn">Street Number:</td><td class="datacolumn"><input type="text" id="residentstreetnumber" name="residentstreetnumber" value="<%=sStreetnumber%>" size="10" maxlength="10" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Street Prefix:</td><td class="datacolumn"><input type="text" id="residentstreetprefix" name="residentstreetprefix" value="<%=sStreetprefix%>" size="15" maxlength="15" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Street Name:</td>
				<td class="datacolumn">
					<!--input type="text" id="residentstreetname" name="residentstreetname" value="<%=sStreetname%>" size="50" maxlength="50" /-->
					<select id="residentstreetname" name="residentstreetname">
						<option><%=sStreetname%></option>
						<% GetStreetNames %>
					</select>

				</td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Street Suffix:</td><td class="datacolumn"><input type="text" id="streetsuffix" name="streetsuffix" value="<%=sStreetSuffix%>" size="15" maxlength="15" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Direction:</td><td class="datacolumn"><input type="text" id="streetdirection" name="streetdirection" value="<%=sStreetDirection%>" size="10" maxlength="10" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Unit/Suite:</td><td class="datacolumn"><input type="text" id="residentunit" name="residentunit" value="<%=sUnit%>" size="10" maxlength="10" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Parcel Id No:</td><td class="datacolumn"><input type="text" id="parcelidnumber" name="parcelidnumber" value="<%=sPin%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">City:</td><td class="datacolumn"><input type="text" id="residentcity" name="residentcity" value="<%=sCity%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">State:</td><td class="datacolumn"><input type="text" id="residentstate" name="residentstate" value="<%=sState%>" size="2" maxlength="2" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Zip:</td><td class="datacolumn"><input type="text" id="residentzip" name="residentzip" value="<%=sZip%>" size="10" maxlength="10" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn"><%=oAddressOrg.GetOrgDisplayName("address grouping field")%>:</td><td class="datacolumn"><input type="text" id="county" name="county" value="<%=sCounty%>" size="50" maxlength="50" /></td>
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
					<select id="residenttype" name="residenttype">
<%						ShowResidencyTypePicks sResidentType %>
					</select>
				</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td>For Mapping, enter the Map Coordinates below.<br />If you do not know them, you can find them <a href="http://www.batchgeocode.com/lookup/" target="_blank">here.</a></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Latitude:</td>
				<td class="datacolumn"><input type="text" id="latitude" name="latitude" size="11" maxlength="11" value="<%=sLatitude%>" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Longitude:</td>
				<td class="datacolumn"><input type="text" id="longitude" name="longitude" size="11" maxlength="11" value="<%=sLongitude%>" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn" valign="top">Listed Owner:</td><td class="datacolumn">
<%					If bNewAddress Then		%>
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Select An Existing Owner" onclick="SelectListedOwner( );" /><br />
<%					End If		%>
					<textarea id="owner" name="listedowner" rows="3" cols="80" maxlength="250" wrap="soft"><%=sOwner%></textarea>
				</td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn" valign="top">Registered User (owner):</td>
				<td class="datacolumn">
					<select id="registereduserid" name="registereduserid" onchange="javascript:UserPick();">
						<option value="0">Select a registered user from the list...</option>
						<% 
							ShowUserDropDown iRegisteredUserId 
						%>
					</select>
					<br />Name Search: <input type="text" name="searchname" value="<%=sSearchName%>" size="25" maxlength="50" onchange="javascript:ClearSearch();" />
					<input type="button" class="button ui-button ui-widget ui-corner-all" value="Search" onclick="javascript:SearchCitizens(document.frmAddress.searchstart.value);" />
					<input type="hidden" id="results" name="results" value="" /><input type="hidden" id="searchstart" name="searchstart" value="<%=sSearchStart%>" />
					<span id="searchresults"><%=sResults%></span>
					<br /><br />					
				</td>
			</tr>

<%			If OrgHasFeature( "extended address fields" ) Then %>
			<tr>
				<td align="right" class="labelcolumn">Property Tax Number:</td><td class="datacolumn"><input type="text" id="propertytaxnumber" name="propertytaxnumber" value="<%=sPropertyTaxNumber%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Lot Number:</td><td class="datacolumn"><input type="text" id="lotnumber" name="lotnumber" value="<%=sLotNumber%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Lot Width:</td><td class="datacolumn"><input type="text" id="lotwidth" name="lotwidth" value="<%=sLotWidth%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Lot Length:</td><td class="datacolumn"><input type="text" id="lotlength" name="lotlength" value="<%=sLotLength%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Block Number:</td><td class="datacolumn"><input type="text" id="blocknumber" name="blocknumber" value="<%=sBlockNumber%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Subdivision:</td><td class="datacolumn"><input type="text" id="subdivision" name="subdivision" value="<%=sSubdivision%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Section:</td><td class="datacolumn"><input type="text" id="section" name="section" value="<%=sSection%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Township:</td><td class="datacolumn"><input type="text" id="township" name="township" value="<%=sTownship%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Range:</td><td class="datacolumn"><input type="text" id="range" name="range" value="<%=sRange%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Permanent Real Estate Index Number:</td><td class="datacolumn"><input type="text" id="permanentrealestateindexnumber" name="permanentrealestateindexnumber" value="<%=sPermanentRealEstateIndexNumber%>" size="50" maxlength="50" /></td>
			</tr>
			<tr>
				<td align="right" class="labelcolumn">Collectors Tax Bill Volume Number:</td><td class="datacolumn"><input type="text" id="collectorstaxbillvolumenumber" name="collectorstaxbillvolumenumber" value="<%=sCollectorsTaxBillVolumeNumber%>" size="50" maxlength="50" /></td>
			</tr>
<%			End If %>

		</table>
	<p>
		<input type="button" id="cancelButton" class="button ui-button ui-widget ui-corner-all" value="Cancel" onclick="doClose();" />
		<input type="button" id="savebutton" class="button ui-button ui-widget ui-corner-all" value="<%=sSaveButtonText%>" onclick="Validate();" />
	</p>
	</form>
	<!--END: EDIT FORM-->

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  
<!--#Include file="modal.asp"-->  

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
	oRs.Open sSQL, Application("DSN"), 3, 1

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

Sub GetStreetNames( )
	sSQL = "SELECT residentstreetname FROM egov_residentaddresses WHERE orgid = '" & session("orgid") & "' GROUP BY residentstreetname ORDER BY residentstreetname"
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1
	Do While Not oRs.EOF
		%>
		<option><%=oRs("ResidentStreetName")%></option>
		<%
		oRs.MoveNext
	loop
	oRs.Close
	Set oRs = Nothing
End Sub


%>


