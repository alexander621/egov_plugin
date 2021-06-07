<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitaddressupdate.asp
' AUTHOR: Steve Loar
' CREATED: 03/26/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This updates the permit addresses
'
' MODIFICATION HISTORY
' 1.0   03/26/2008  Steve Loar - INITIAL VERSION
' 2.0	04/02/2008	Steve Loar - Allow updates and new addresses from permit create page
' 2.1	07/21/2009	Steve Loar - New fields for Lansing IL
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iResidentAddressId, sSql, sResidentstreetnumber, sResidentstreetprefix, sResidentstreetname
Dim sResidentunit, sStreetsuffix, sSortStreetname, sParcelidnumber, sResidentcity, sResidentstate
Dim sResidentzip, sLegaldescription, sResidentType, sLandvalue, sTotalvalue, sTaxdistrict
Dim sListedOwner, iRegisteredUserId, sLatitude, sLongitude, oRs, sCounty, sStreetDirection
Dim iPermitId, iPermitStatusId, bChangesPropagate
Dim sPropertyTaxNumber, sLotNumber, sLotWidth, sLotLength, sBlockNumber, sSubdivision, sSection
Dim sTownship, sRange, sPermanentRealEstateIndexNumber, sCollectorsTaxBillVolumeNumber

' Process the passed values
if request("residentaddressid") = "" then
	response.write "Sorry, your request cannot be processed.  Please try again later."
	response.end
end if
iResidentAddressId = CLng(request("residentaddressid"))
If request("permitid") <> "" Then 
	iPermitId = CLng(request("permitid"))
	iPermitStatusId = CLng(request("permitstatusid"))
	bChangesPropagate = StatusAllowsChangesToPropagate( iPermitStatusId )
Else
	iPermitId = CLng(0)
	iPermitStatusId = CLng(0)
	bChangesPropagate = True 
End If 

sResidentstreetnumber = "'" & Trim(request("residentstreetnumber")) & "'"
sResidentstreetprefix = "'" & dbsafe(Trim(request("residentstreetprefix"))) & "'"
sResidentstreetname = "'" & dbsafe(Trim(request("residentstreetname"))) & "'"
sResidentunit = "'" & dbsafe(Trim(request("residentunit"))) & "'"
sParcelidnumber = "'" & dbsafe(request("parcelidnumber")) & "'"
sStreetsuffix = "'" & dbsafe(Trim(request("streetsuffix"))) & "'"
sStreetDirection = "'" & dbsafe(Trim(request("streetdirection"))) & "'"
sResidentcity = "'" & dbsafe(request("residentcity")) & "'"
sResidentstate = "'" & dbsafe(request("residentstate")) & "'"
sResidentzip = "'" & dbsafe(request("residentzip")) & "'"
sCounty = "'" & dbsafe(request("county")) & "'"
sLegaldescription = "'" & dbsafe(request("legaldescription")) & "'"
sResidentType =  "'" & request("residenttype") & "'"
sListedOwner = "'" & dbsafe(request("listedowner")) & "'"

sSortStreetname = "'" & dbsafe(Trim(request("residentstreetname")))
If request("streetsuffix") <> "" Then 
	sSortStreetname = Trim(sSortStreetname & " " & dbsafe(request("streetsuffix")))
End If 
If request("streetdirection") <> "" Then 
	sSortStreetname = Trim(sSortStreetname & " " & dbsafe(request("streetdirection")))
End If 
If request("streetprefix") <> "" Then 
	sSortStreetname = Trim(sSortStreetname & " " & dbsafe(request("residentstreetprefix")))
End If 
'If request("unit") <> "" Then 
'	sSortStreetname = sSortStreetname & " " & dbsafe(request("residentunit"))
'End If 
sSortStreetname = sSortStreetname & "'"

If CLng(request("registereduserid")) > CLng(0) Then 
	iRegisteredUserId = request("registereduserid")
Else
	iRegisteredUserId = "NULL"
End If 

If request("latitude") <> "" Then
	sLatitude = request("latitude")
Else 
	sLatitude = "NULL"
End If 

If request("longitude") <> "" Then
	sLongitude = request("longitude")
Else 
	sLongitude = "NULL"
End If 

If request("propertytaxnumber") <> "" Then
	sPropertyTaxNumber = "'" & dbsafe(request("propertytaxnumber")) & "'"
Else
	sPropertyTaxNumber = "NULL"
End If 

If request("lotnumber") <> "" Then
	sLotNumber = "'" & dbsafe(request("lotnumber")) & "'"
Else
	sLotNumber = "NULL"
End If 

If request("lotwidth") <> "" Then
	sLotWidth = "'" & dbsafe(request("lotwidth")) & "'"
Else
	sLotWidth = "NULL"
End If 

If request("lotlength") <> "" Then
	sLotLength = "'" & dbsafe(request("lotlength")) & "'"
Else
	sLotLength = "NULL"
End If 

If request("blocknumber") <> "" Then
	sBlockNumber = "'" & dbsafe(request("blocknumber")) & "'"
Else
	sBlockNumber = "NULL"
End If 

If request("subdivision") <> "" Then
	sSubdivision = "'" & dbsafe(request("subdivision")) & "'"
Else
	sSubdivision = "NULL"
End If 

If request("section") <> "" Then
	sSection = "'" & dbsafe(request("section")) & "'"
Else
	sSection = "NULL"
End If 

If request("township") <> "" Then
	sTownship = "'" & dbsafe(request("township")) & "'"
Else
	sTownship = "NULL"
End If 

If request("range") <> "" Then
	sRange = "'" & dbsafe(request("range")) & "'"
Else
	sRange = "NULL"
End If 

If request("permanentrealestateindexnumber") <> "" Then
	sPermanentRealEstateIndexNumber = "'" & dbsafe(request("permanentrealestateindexnumber")) & "'"
Else
	sPermanentRealEstateIndexNumber = "NULL"
End If 

If request("collectorstaxbillvolumenumber") <> "" Then
	sCollectorsTaxBillVolumeNumber = "'" & dbsafe(request("collectorstaxbillvolumenumber")) & "'"
Else
	sCollectorsTaxBillVolumeNumber = "NULL"
End If 


If iResidentAddressId > CLng(0) Then 
	If bChangesPropagate Then 
		' Update the parent address record
		sSql = "UPDATE egov_residentaddresses "
		sSql = sSql & " SET residentstreetnumber = " & sResidentstreetnumber
		sSql = sSql & ", residentstreetprefix = " & sResidentstreetprefix
		sSql = sSql & ", residentstreetname = " & sResidentstreetname
		sSql = sSql & ", streetsuffix = " & sStreetsuffix
		sSql = sSql & ", streetdirection = " & sStreetDirection
		sSql = sSql & ", residentunit = " & sResidentunit
		sSql = sSql & ", parcelidnumber = " & sParcelidnumber
		sSql = sSql & ", sortstreetname = " & sSortstreetname
		sSql = sSql & ", residentcity = " & sResidentcity
		sSql = sSql & ", residentstate = " & sResidentstate
		sSql = sSql & ", residentzip = " & sResidentzip
		sSql = sSql & ", county = " & sCounty
		sSql = sSql & ", legaldescription = " & sLegaldescription
		sSql = sSql & ", residenttype = " & sResidentType
		sSql = sSql & ", latitude = " & sLatitude
		sSql = sSql & ", longitude = " & sLongitude
		sSql = sSql & ", listedowner = " & sListedOwner
		sSql = sSql & ", registereduserid = " & iRegisteredUserId
		sSql = sSql & ", propertytaxnumber = " & sPropertyTaxNumber
		sSql = sSql & ", lotnumber = " & sLotNumber
		sSql = sSql & ", lotwidth = " & sLotWidth
		sSql = sSql & ", lotlength = " & sLotLength
		sSql = sSql & ", blocknumber = " & sBlockNumber
		sSql = sSql & ", subdivision = " & sSubdivision
		sSql = sSql & ", section = " & sSection
		sSql = sSql & ", township = " & sTownship
		sSql = sSql & ", range = " & sRange
		sSql = sSql & ", permanentrealestateindexnumber = " & sPermanentRealEstateIndexNumber
		sSql = sSql & ", collectorstaxbillvolumenumber = " & sCollectorsTaxBillVolumeNumber
		sSql = sSql & " WHERE orgid = " & session("orgid") & " AND residentaddressid = " & iResidentAddressId
		RunSQL sSql 

		' Pull the permitaddresses that are still open and have the residentaddressid and update them
		sSql = "SELECT A.permitaddressid FROM egov_permitaddress A, egov_permits P, egov_permitstatuses S "
		sSql = sSql & " WHERE A.orgid = " & session("orgid") & " AND A.permitid = P.permitid AND P.permitstatusid = S.permitstatusid AND S.iscompletedstatus = 0  AND S.cansavechanges = 1 AND S.changespropagate = 1 "
		sSql = sSql & " AND A.residentaddressid = " & iResidentAddressId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSQL, Application("DSN"), 3, 1

		Do While Not oRs.EOF
			sSql = "UPDATE egov_permitaddress "
			sSql = sSql & " SET residentstreetnumber = " & sResidentstreetnumber
			sSql = sSql & ", residentstreetprefix = " & sResidentstreetprefix
			sSql = sSql & ", residentstreetname = " & sResidentstreetname
			sSql = sSql & ", streetsuffix = " & sStreetsuffix
			sSql = sSql & ", streetdirection = " & sStreetDirection
			sSql = sSql & ", residentunit = " & sResidentunit
			sSql = sSql & ", parcelidnumber = " & sParcelidnumber
			sSql = sSql & ", sortstreetname = " & sSortstreetname
			sSql = sSql & ", residentcity = " & sResidentcity
			sSql = sSql & ", residentstate = " & sResidentstate
			sSql = sSql & ", residentzip = " & sResidentzip
			sSql = sSql & ", county = " & sCounty
			sSql = sSql & ", legaldescription = " & sLegaldescription
			sSql = sSql & ", residenttype = " & sResidentType
			sSql = sSql & ", latitude = " & sLatitude
			sSql = sSql & ", longitude = " & sLongitude
			sSql = sSql & ", listedowner = " & sListedOwner
			sSql = sSql & ", registereduserid = " & iRegisteredUserId
			sSql = sSql & ", propertytaxnumber = " & sPropertyTaxNumber
			sSql = sSql & ", lotnumber = " & sLotNumber
			sSql = sSql & ", lotwidth = " & sLotWidth
			sSql = sSql & ", lotlength = " & sLotLength
			sSql = sSql & ", blocknumber = " & sBlockNumber
			sSql = sSql & ", subdivision = " & sSubdivision
			sSql = sSql & ", section = " & sSection
			sSql = sSql & ", township = " & sTownship
			sSql = sSql & ", range = " & sRange
			sSql = sSql & ", permanentrealestateindexnumber = " & sPermanentRealEstateIndexNumber
			sSql = sSql & ", collectorstaxbillvolumenumber = " & sCollectorsTaxBillVolumeNumber
			sSql = sSql & " WHERE permitaddressid = " & oRs("permitaddressid")
			RunSQL sSql 
			oRs.MoveNext 
		Loop

		oRs.Close
		Set oRs = Nothing 
	Else
		' Pull the permitaddress for the permit
		sSql = "SELECT A.permitaddressid FROM egov_permitaddress A, egov_permits P, egov_permitstatuses S "
		sSql = sSql & " WHERE A.orgid = " & session("orgid") & " AND A.permitid = P.permitid AND P.permitstatusid = S.permitstatusid AND S.iscompletedstatus = 0 "
		sSql = sSql & " AND A.residentaddressid = " & iResidentAddressId & " AND P.permitid = " & iPermitId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSQL, Application("DSN"), 3, 1

		Do While Not oRs.EOF
			sSql = "UPDATE egov_permitaddress "
			sSql = sSql & " SET residentstreetnumber = " & sResidentstreetnumber
			sSql = sSql & ", residentstreetprefix = " & sResidentstreetprefix
			sSql = sSql & ", residentstreetname = " & sResidentstreetname
			sSql = sSql & ", streetsuffix = " & sStreetsuffix
			sSql = sSql & ", streetdirection = " & sStreetDirection
			sSql = sSql & ", residentunit = " & sResidentunit
			sSql = sSql & ", parcelidnumber = " & sParcelidnumber
			sSql = sSql & ", sortstreetname = " & sSortstreetname
			sSql = sSql & ", residentcity = " & sResidentcity
			sSql = sSql & ", residentstate = " & sResidentstate
			sSql = sSql & ", residentzip = " & sResidentzip
			sSql = sSql & ", county = " & sCounty
			sSql = sSql & ", legaldescription = " & sLegaldescription
			sSql = sSql & ", residenttype = " & sResidentType
			sSql = sSql & ", latitude = " & sLatitude
			sSql = sSql & ", longitude = " & sLongitude
			sSql = sSql & ", listedowner = " & sListedOwner
			sSql = sSql & ", registereduserid = " & iRegisteredUserId
			sSql = sSql & ", propertytaxnumber = " & sPropertyTaxNumber
			sSql = sSql & ", lotnumber = " & sLotNumber
			sSql = sSql & ", lotwidth = " & sLotWidth
			sSql = sSql & ", lotlength = " & sLotLength
			sSql = sSql & ", blocknumber = " & sBlockNumber
			sSql = sSql & ", subdivision = " & sSubdivision
			sSql = sSql & ", section = " & sSection
			sSql = sSql & ", township = " & sTownship
			sSql = sSql & ", range = " & sRange
			sSql = sSql & ", permanentrealestateindexnumber = " & sPermanentRealEstateIndexNumber
			sSql = sSql & ", collectorstaxbillvolumenumber = " & sCollectorsTaxBillVolumeNumber
			sSql = sSql & " WHERE permitaddressid = " & oRs("permitaddressid")
			RunSQL sSql 
			oRs.MoveNext 
		Loop

		oRs.Close
		Set oRs = Nothing 

	End If 
Else
	' Insert the new address. This should only be called from the New Permit Create page
	sSql = "INSERT INTO egov_residentaddresses ( residentstreetnumber, residentstreetprefix, residentstreetname, "
	sSql = sSql & " streetsuffix, streetdirection, residentunit, parcelidnumber, sortstreetname, residentcity, "
	sSql = sSql & " residentstate, residentzip, county, legaldescription, residenttype, latitude, longitude, "
	sSql = sSql & " listedowner, registereduserid, orgid, propertytaxnumber, lotnumber, lotwidth, lotlength, "
	sSql = sSql & " blocknumber, subdivision, section, township, range, permanentrealestateindexnumber, "
	sSql = sSql & " collectorstaxbillvolumenumber ) VALUES ( "
	sSql = sSql & sResidentstreetnumber & ", " & sResidentstreetprefix & ", " & sResidentstreetname & ", " 
	sSql = sSql & sStreetsuffix & ", " & sStreetDirection & ", " & sResidentunit & ", " & sParcelidnumber & ", " 
	sSql = sSql & sSortstreetname & ", " & sResidentcity & ", " & sResidentstate & ", " & sResidentzip & ", "
	sSql = sSql & sCounty & ", " & sLegaldescription & ", " & sResidentType & ", " & sLatitude & ", "
	sSql = sSql & sLongitude & ", " & sListedOwner & ", " & iRegisteredUserId & ", " & session("orgid") & ", "
	sSql = sSql & sPropertyTaxNumber & ", " & sLotNumber & ", " & sLotWidth & ", " & sLotLength & ", "
	sSql = sSql & sBlockNumber & ", " & sSubdivision & ", " & sSection & ", " & sTownship & ", " & sRange & ", "
	sSql = sSql & sPermanentRealEstateIndexNumber & ", " & sCollectorsTaxBillVolumeNumber & " )"
	RunSQL sSql 
End If 

response.write "Success"


%>
