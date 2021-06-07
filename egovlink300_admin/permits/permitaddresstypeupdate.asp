<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitaddresstypeupdate.asp
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
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitAddressTypeid, sSql, sStreetnumber, sStreetprefix, sStreetname, sStreettype, sResidentType
Dim sUnit, sPin, sStreetsuffix, sCity, sState, sZip, sLegaldescription, sLatitude, sLongitude
Dim sStreetDirection, sOwner, sRegisteredUserId, sSortStreetname, sCounty, sExcludeFromActionLine
Dim sPagenum

iPermitAddressTypeid = CLng(request("permitaddresstypeid") )
sStreetnumber = "'" & request("streetnumber") & "'"
sStreetprefix = "'" & dbsafe(request("streetprefix")) & "'"
sStreetname = "'" & dbsafe(request("streetname")) & "'"
sUnit = "'" & dbsafe(request("unit")) & "'"
sPin = "'" & dbsafe(request("pin")) & "'"
sStreetsuffix = "'" & dbsafe(request("streetsuffix")) & "'"
sStreetDirection = "'" & dbsafe(request("streetdirection")) & "'"
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

sSortStreetname = "'" & dbsafe(request("streetname"))
If request("streetsuffix") <> "" Then 
	sSortStreetname = sSortStreetname & " " & dbsafe(request("streetsuffix"))
End If 
If request("streetdirection") <> "" Then 
	sSortStreetname = sSortStreetname & " " & dbsafe(request("streetdirection"))
End If 
If request("streetprefix") <> "" Then 
	sSortStreetname = sSortStreetname & " " & dbsafe(request("streetprefix"))
End If 
'If request("unit") <> "" Then 
'	sSortStreetname = sSortStreetname & " " & dbsafe(request("unit"))
'End If 
sSortStreetname = sSortStreetname & "'"

If request("pagenum") <> "" Then
	sPagenum = "&pagenum=" & request("pagenum")
Else
	sPagenum = ""
End If 

If iPermitAddressTypeid = CLng(0) Then 
	sSql = "INSERT INTO egov_residentaddresses ( orgid, residentstreetnumber, residentstreetprefix, residentstreetname, streetsuffix, latitude, longitude, "
	sSql = sSql & " residentunit, parcelidnumber, sortstreetname, residentcity, residentstate, residentzip, county, legaldescription, "
	sSql = sSql & " residenttype, listedowner, registereduserid, streetdirection, excludefromactionline ) "
	sSql = sSql & " VALUES ( " & session("orgid") & ", " & sStreetnumber & ", " & sStreetprefix & ", " & sStreetname & ", " & sStreetsuffix & ", " & sLatitude & ", " & sLongitude
	sSql = sSql & ", " & sUnit & ", " & sPin & ", " & sSortstreetname & ", " & sCity & ", " & sState & ", " & sZip & ", " & sCounty & ", " & sLegaldescription
	sSql = sSql & ", " & sResidentType & ", " & sOwner & ", " & sRegisteredUserId & ", " & sStreetDirection & ", " & sExcludeFromActionLine & " )"
	iPermitAddressTypeid = RunIdentityInsert( sSql ) 
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
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND residentaddressid = " & iPermitAddressTypeid
	RunSQL sSql 

	' Pull the permitaddresses that are still open and have the residentaddressid and update them
sSql = "SELECT A.permitaddressid FROM egov_permitaddress A, egov_permits P, egov_permitstatuses S "
sSql = sSql & " WHERE A.orgid = " & session("orgid") & " AND A.permitid = P.permitid AND P.permitstatusid = S.permitstatusid AND S.iscompletedstatus = 0  AND S.cansavechanges = 1 AND S.changespropagate = 1 "
sSql = sSql & " AND A.residentaddressid = " & iPermitAddressTypeid

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

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
	sSql = sSql & " WHERE permitaddressid = " & oRs("permitaddressid")
	RunSQL sSql 
	oRs.MoveNext 
Loop

oRs.Close
Set oRs = Nothing 

End If 

response.redirect "permitaddresstypeedit.asp?permitaddresstypeid=" & iPermitAddressTypeid & "&searchtext=" & request("searchtext") & "&searchfield=" & request("searchfield") & sPagenum



%>
