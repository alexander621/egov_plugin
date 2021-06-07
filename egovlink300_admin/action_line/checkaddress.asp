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
' 1.0 08/28/2007	 Steve Loar - INITIAL VERSION
' 1.1	02/05/2008	 Steve Loar - Changed handling of street number to handle none provided
' 1.2 04/10/2008  David Boyer - Modified adderss format
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, sStreetNumber, sStreetName, oRs, sResults

'iStNumber = CLng(request("stnumber"))
If request("stnumber") <> "" Then 
   If IsNumeric(request("stnumber")) Then 
    		sStreetNumber = Cdbl(request("stnumber"))
  	Else
		    sStreetNumber = dbsafe(request("stnumber"))
  	End If 
Else
  	sStreetNumber = ""
End If 
'sStreetName = DBsafe(request("stname"))
sStreetName = request("stname")

'sSql = "SELECT COUNT(residentaddressid) as hits "
'sSql = sSql & " FROM egov_residentaddresses "
'sSql = sSql & " WHERE orgid = " & session("orgid")
'sSql = sSql & " AND residentstreetnumber = '" & iStNumber
'sSql = sSql & "' AND residentstreetname = '" & sStName & "' "

sSQL = "SELECT COUNT(residentaddressid) AS hits "
sSQL = sSQL & " FROM egov_residentaddresses "
sSQL = sSQL & " WHERE orgid = " & Session("OrgID")
sSQL = sSQL & " AND residentstreetnumber = '" & dbsafe(sStreetNumber) & "' "
sSQL = sSQL & " AND (residentstreetname = '" & dbsafe(sStreetName) & "' "
sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix = '" & dbsafe(sStreetName) & "' "
sSQL = sSQL & " OR residentstreetname + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' "
sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' "
sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname = '" & dbsafe(sStreetName) & "' "
sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = '" & dbsafe(sStreetName) & "' "
sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' "
sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & dbsafe(sStreetName) & "'"
sSQL = sSQL & ")"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If Not oRs.EOF Then
 	 If CLng(oRs("hits")) > CLng(0) Then
	 	   sResults = "FOUND CHECK"
   Else
		    sResults = "NOT FOUND"
  	End If 
Else
	  sResults = "NOT FOUND"
End If 

oRs.Close
Set oRs = Nothing 

response.write sResults

'--------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'--------------------------------------------------------------------------------------------------
Function DBsafe( strDB )

If Not VarType( strDB ) = vbString Then 
 		DBsafe = strDB
Else 
  	DBsafe = Replace( strDB, "'", "''" )
End If 

End Function
%>
