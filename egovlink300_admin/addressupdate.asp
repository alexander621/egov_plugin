<!-- #include file="includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: addresstypeupdate.asp
' AUTHOR: Steve Loar
' CREATED: 02/11/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the permit address types
'
' MODIFICATION HISTORY
' 1.0   02/11/2008   Steve Loar - INITIAL VERSION
' 1.1	03/27/2008	Steve Loar - Changed to add county
' 1.2	07/21/2009	Steve Loar - New fields for Lansing IL
' 1.3	07/6/2011	Steve Loar - Changing key address fields to store NULL when empty
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iResidentAddressId, sSql, sStreetnumber, sStreetprefix, sStreetname, sStreettype, sResidentType
Dim sUnit, sPin, sStreetsuffix, sCity, sState, sZip, sLegaldescription, sLatitude, sLongitude
Dim sStreetDirection, sOwner, sRegisteredUserId, sSortStreetname, sCounty, sExcludeFromActionLine
Dim sPropertyTaxNumber, sLotNumber, sLotWidth, sLotLength, sBlockNumber, sSubdivision, sSection
Dim sTownship, sRange, sPermanentRealEstateIndexNumber, sCollectorsTaxBillVolumeNumber, sfloodplain, szoningdistrict
Dim MessageFlag

iResidentAddressId = CLng(request("residentaddressid"))

If request("streetnumber") <> "" Then 
	sStreetnumber = "'" & Trim(request("streetnumber")) & "'"
Else
	sStreetnumber = "NULL"
End If 

If request("streetprefix") <> "" Then 
	sStreetprefix = "'" & dbsafe(Trim(request("streetprefix"))) & "'"
Else
	sStreetprefix = "NULL"
End If 

If request("streetname") <> "" Then 
	sStreetname = "'" & dbsafe(Trim(request("streetname"))) & "'"
Else
	sStreetname = "NULL"
End If 

If request("streetsuffix") <> "" Then 
	sStreetsuffix = "'" & dbsafe(Trim(request("streetsuffix"))) & "'"
Else
	sStreetsuffix = "NULL"
End If 

If request("streetdirection") <> "" Then 
	sStreetDirection = "'" & dbsafe(Trim(request("streetdirection"))) & "'"
Else
	sStreetDirection = "NULL"
End If 

sUnit = "'" & dbsafe(Trim(request("unit"))) & "'"
sPin = "'" & dbsafe(request("pin")) & "'"
sCity = "'" & dbsafe(request("city")) & "'"
sState = "'" & dbsafe(request("state")) & "'"
sZip = "'" & dbsafe(request("zip")) & "'"
sCounty = "'" & dbsafe(request("county")) & "'"
sLegaldescription = "'" & dbsafe(request("legaldescription")) & "'"
sResidentType =  "'" & request("residenttype") & "'"
sOwner = "'" & dbsafe(request("listedowner")) & "'"

If request("excludefromactionline") = "on" Then
	sExcludeFromActionLine = "1"
Else
	sExcludeFromActionLine = "0" 
End If 

If CLng(request("registereduserid")) > CLng(0) Then 
	sRegisteredUserId = request("registereduserid")
Else
	sRegisteredUserId = "NULL"
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

sSortStreetname = "'" & Trim(dbsafe(request("streetname")))
If request("streetsuffix") <> "" Then 
	sSortStreetname = Trim(sSortStreetname & " " & dbsafe(request("streetsuffix")))
End If 
If request("streetdirection") <> "" Then 
	sSortStreetname = Trim(sSortStreetname & " " & dbsafe(request("streetdirection")))
End If 
If request("streetprefix") <> "" Then 
	sSortStreetname = Trim(sSortStreetname & " " & dbsafe(request("streetprefix")))
End If 
sSortStreetname = sSortStreetname & "'"

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

If request("floodplain") <> "" Then
	sfloodplain = "'" & dbsafe(request("floodplain")) & "'"
Else
	sfloodplain = "NULL"
End If 

If request("zoningdistrict") <> "" Then
	szoningdistrict = "'" & dbsafe(request("zoningdistrict")) & "'"
Else
	szoningdistrict = "NULL"
End If 


If iResidentAddressId = CLng(0) Then 
	sSql = "INSERT INTO egov_residentaddresses ( orgid, residentstreetnumber, residentstreetprefix, residentstreetname, streetsuffix, latitude, longitude, "
	sSql = sSql & " residentunit, parcelidnumber, sortstreetname, residentcity, residentstate, residentzip, county, legaldescription, "
	sSql = sSql & " residenttype, listedowner, registereduserid, streetdirection, excludefromactionline, propertytaxnumber, lotnumber, "
	sSql = sSql & " lotwidth, lotlength, blocknumber, subdivision, section, township, range, permanentrealestateindexnumber, collectorstaxbillvolumenumber, floodplain, zoningdistrict ) "
	sSql = sSql & " VALUES ( " & session("orgid") & ", " & sStreetnumber & ", " & sStreetprefix & ", " & sStreetname & ", " & sStreetsuffix & ", " & sLatitude & ", " & sLongitude & ", "
	sSql = sSql & sUnit & ", " & sPin & ", " & sSortstreetname & ", " & sCity & ", " & sState & ", " & sZip & ", " & sCounty & ", " & sLegaldescription & ", "
	sSql = sSql & sResidentType & ", " & sOwner & ", " & sRegisteredUserId & ", " & sStreetDirection & ", " & sExcludeFromActionLine & ", "
	sSql = sSql & sPropertyTaxNumber & ", " & sLotNumber & ", " & sLotWidth & ", " & sLotLength & ", " & sBlockNumber & ", " & sSubdivision & ", "
	sSql = sSql & sSection & ", " & sTownship & ", " & sRange & ", " & sPermanentRealEstateIndexNumber & ", " & sCollectorsTaxBillVolumeNumber & ", " & sfloodplain & ", " & szoningdistrict & " )"

	iResidentAddressId = RunIdentityInsert( sSql ) 

	MessageFlag = "n"
Else 
	sSql = "UPDATE egov_residentaddresses "
	sSql = sSql & " SET residentstreetnumber = " & sStreetnumber
	sSql = sSql & ", residentstreetprefix = " & sStreetprefix
	sSql = sSql & ", residentstreetname = " & sStreetname
	sSql = sSql & ", streetsuffix = " & sStreetsuffix
	sSql = sSql & ", residentunit = " & sUnit
	sSql = sSql & ", parcelidnumber = " & sPin
	sSql = sSql & ", sortstreetname = " & sSortstreetname
	sSql = sSql & ", residentcity = " & sCity
	sSql = sSql & ", residentstate = " & sState
	sSql = sSql & ", residentzip = " & sZip
	sSql = sSql & ", county = " & sCounty
	sSql = sSql & ", legaldescription = " & sLegaldescription
	sSql = sSql & ", residenttype = " & sResidentType
	sSql = sSql & ", latitude = " & sLatitude
	sSql = sSql & ", longitude = " & sLongitude
	sSql = sSql & ", listedowner = " & sOwner
	sSql = sSql & ", registereduserid = " & sRegisteredUserId
	sSql = sSql & ", streetdirection = " & sStreetDirection
	sSql = sSql & ", excludefromactionline = " & sExcludeFromActionLine
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
	sSql = sSql & ", floodplain = " & sfloodplain
	sSql = sSql & ", zoningdistrict = " & szoningdistrict
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND residentaddressid = " & iResidentAddressId

	RunSQL sSql 

	MessageFlag = "u"

	' Pull the permitaddresses that are still open and have the residentaddressid and update them
sSql = "SELECT A.permitaddressid FROM egov_permitaddress A, egov_permits P, egov_permitstatuses S "
sSql = sSql & "WHERE A.orgid = " & session("orgid") & " AND A.permitid = P.permitid AND "
sSql = sSql & "P.permitstatusid = S.permitstatusid AND S.iscompletedstatus = 0 AND S.cansavechanges = 1 AND S.changespropagate = 1 "
sSql = sSql & "AND A.residentaddressid = " & iResidentAddressId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 0, 1

Do While Not oRs.EOF
	sSql = "UPDATE egov_permitaddress "
	sSql = sSql & " SET residentstreetnumber = " & sStreetnumber
	sSql = sSql & ", residentstreetprefix = " & sStreetprefix
	sSql = sSql & ", residentstreetname = " & sStreetname
	sSql = sSql & ", streetsuffix = " & sStreetsuffix
	sSql = sSql & ", residentunit = " & sUnit
	sSql = sSql & ", parcelidnumber = " & sPin
	sSql = sSql & ", sortstreetname = " & sSortstreetname
	sSql = sSql & ", residentcity = " & sCity
	sSql = sSql & ", residentstate = " & sState
	sSql = sSql & ", residentzip = " & sZip
	sSql = sSql & ", county = " & sCounty
	sSql = sSql & ", legaldescription = " & sLegaldescription
	sSql = sSql & ", residenttype = " & sResidentType
	sSql = sSql & ", latitude = " & sLatitude
	sSql = sSql & ", longitude = " & sLongitude
	sSql = sSql & ", listedowner = " & sOwner
	sSql = sSql & ", registereduserid = " & sRegisteredUserId
	sSql = sSql & ", streetdirection = " & sStreetDirection
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

If request("pagenum") <> "" Then
	sPageNum = "&pagenum=" & request("pagenum")
Else
	sPageNum = ""
End If 
If request("keyword") <> "" Then
	sKeyword = "&keyword=" & request("keyword")
Else
	sKeyword = ""
End If 

response.redirect "addressedit.asp?msg=" & MessageFlag & "&residentaddressid=" & iResidentAddressId & sKeyword & sPageNum


'-------------------------------------------------------------------------------------------------
' Sub RunSQL( sSql )
'-------------------------------------------------------------------------------------------------
Sub RunSQL( ByVal sSql )
	Dim oCmd

	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub 


'-------------------------------------------------------------------------------------------------
' Function RunIdentityInsert( sInsertStatement )
'-------------------------------------------------------------------------------------------------
Function RunIdentityInsert( ByVal sInsertStatement )
	Dim sSQL, iReturnValue, oInsert

	iReturnValue = 0

'	response.write "<p>" & sInsertStatement & "</p><br /><br />"
'	response.flush

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSQL = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.Open sSQL, Application("DSN"), 3, 3
	iReturnValue = oInsert("ROWID")
	oInsert.close
	Set oInsert = Nothing

	RunIdentityInsert = iReturnValue

End Function


%>
