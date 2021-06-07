<!-- #include file="includes/common.asp" //-->
<!-- #include file="include_top_functions.asp" //-->
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
' 1.0  08/28/07	 Steve Loar - INITIAL VERSION
' 1.1	 01/15/08	 Steve Loar - Changes to handle more street number types safely.
' 1.2	 01/31/08	 Steve Loar - To handle spiders that do not have an orgid.
' 2.0	 04/07/08	 Steve Loar - Using prefix, suffix and direction
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, sStreetNumber, sStreetName, oRs, sResults, sOrgID

If request("stnumber") <> "" Then 
  	If IsNumeric(request("stnumber")) Then 
		    sStreetNumber = CLng(request("stnumber"))
  	Else
		    sStreetNumber = request("stnumber")
  	End If 
Else
  	sStreetNumber = ""
End If 

sStreetName = request("stname")

If request("orgid") = "" Then 
  	sOrgID = CLng(0)
Else 
  	sOrgID = CLng(request("orgid"))
End If 

sSQL = "SELECT COUNT(residentaddressid) AS hits " & vbcrlf
sSQL = sSQL & " FROM egov_residentaddresses " & vbcrlf
sSQL = sSQL & " WHERE orgid = " & sOrgID & vbcrlf
sSQL = sSQL & " AND residentstreetnumber = '" & dbsafe(sStreetNumber) & "' " & vbcrlf
sSQL = sSQL & " AND (ltrim(rtrim(residentstreetname)) = '" & dbsafe(sStreetName) & "' " & vbcrlf
sSQL = sSQL & " OR ltrim(rtrim(residentstreetname)) + ' ' + ltrim(rtrim(streetsuffix)) = '" & dbsafe(sStreetName) & "' " & vbcrlf
sSQL = sSQL & " OR ltrim(rtrim(residentstreetname)) + ' ' + ltrim(rtrim(streetdirection)) = '" & dbsafe(sStreetName) & "' " & vbcrlf
sSQL = sSQL & " OR ltrim(rtrim(residentstreetname)) + ' ' + ltrim(rtrim(streetsuffix)) + ' ' + ltrim(rtrim(streetdirection)) = '" & dbsafe(sStreetName) & "' " & vbcrlf
sSQL = sSQL & " OR ltrim(rtrim(residentstreetprefix)) + ' ' + ltrim(rtrim(residentstreetname)) = '" & dbsafe(sStreetName) & "' " & vbcrlf
sSQL = sSQL & " OR ltrim(rtrim(residentstreetprefix)) + ' ' + ltrim(rtrim(residentstreetname)) + ' ' + ltrim(rtrim(streetsuffix)) = '" & dbsafe(sStreetName) & "' " & vbcrlf
sSQL = sSQL & " OR ltrim(rtrim(residentstreetprefix)) + ' ' + ltrim(rtrim(residentstreetname)) + ' ' + ltrim(rtrim(streetdirection)) = '" & dbsafe(sStreetName) & "' " & vbcrlf
sSQL = sSQL & " OR ltrim(rtrim(residentstreetprefix)) + ' ' + ltrim(rtrim(residentstreetname)) + ' ' + ltrim(rtrim(streetsuffix)) + ' ' + ltrim(rtrim(streetdirection)) = '" & dbsafe(sStreetName) & "'" & vbcrlf
sSQL = sSQL & ")"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

if NOT oRs.EOF then
  	if CLng(oRs("hits")) > CLng(0) then
    		sResults = "FOUND CHECK"
  	else
    		sResults = "NOT FOUND"
   end if
else
  	sResults = "NOT FOUND"
end if

oRs.Close
Set oRs = Nothing 

response.write sResults

'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'--------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
	Dim sNewString

	If Not VarType( strDB ) = vbString Then 
  		DBsafe = strDB & ""
	Else 
'		  sNewString = Replace( strDB, "'", "''" )
'		  DBsafe    = Replace( sNewString, "<", "&lt;" )
    dbsafe = replace(strDB,"'","''")
	End If 
End Function
%>
