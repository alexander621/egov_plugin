<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: checkaddress.asp
' AUTHOR: Steve Loar
' CREATED: 08/28/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This checks that the passed address is in the loaded address list, is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   08/28/07	Steve Loar - INITIAL VERSION
' 2.0	04/04/2008	Steve Loar - Changed to use new fields of prefix, suffix, direction
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, sStreetNumber, sStreetName, oRs, sResults

sStreetNumber = CLng(request("stnumber"))
sStreetName = request("stname")

sSql = "SELECT COUNT(residentaddressid) AS hits FROM egov_residentaddresses "
sSql = sSql & " WHERE residentstreetnumber = '" & dbsafe(sStreetNumber) & "' "
sSql = sSql & " AND (residentstreetname = '" & dbsafe(sStreetName) & "' "
sSql = sSql & " OR residentstreetname + ' ' + streetsuffix = '" & dbsafe(sStreetName) & "' "
sSQL = sSQL & " OR residentstreetname + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' "
sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' "
sSql = sSql & " OR residentstreetprefix + ' ' + residentstreetname = '" & dbsafe(sStreetName) & "' "
sSql = sSql & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = '" & dbsafe(sStreetName) & "' "
sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' "
sSql = sSql & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' )"
sSql = sSql & " AND orgid = " & Session("OrgID")  

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If Not oRs.EOF Then
	If clng(oRs("hits")) > clng(0) Then
		sResults = "FOUND"
	Else
		sResults = "NOT FOUND"
	End If 
Else
	sResults = "NOT FOUND"
End If 

oRs.Close
Set oRs = Nothing 

response.write sResults


%>
