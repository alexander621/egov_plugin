<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: waiver_save.asp
' AUTHOR: John Stullenberger
' CREATED: 2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Remove the waiver
'
' MODIFICATION HISTORY
' 1.0   2006	John Stullenberger - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

Dim iWaiverId, iFacilityId, sName, sDescription, sBody, sSql

iWaiverId = CLng(request("iWaiverId"))
iFacilityId = CLng(request("iFacilityId"))
sName = dbsafe(request("sName"))
sDescription = dbsafe(request("sDescription"))
sBody = dbsafe(request("sBody"))


If iWaiverId = CLng(0) Then
	' Insert new records
	sSql = "INSERT INTO egov_waivers ( orgid, name, description, body ) VALUES ( " 
	sSql = sSql & Session("orgid") & ", '" & sName & "', '" & sDescription & "', '" & sBody & "' )"
Else 
	' Update existing records
	sSQL = "UPDATE egov_waivers SET name = '" & sName & "', description = '"
	sSql = sSql & sDescription & "', body = '" & sBody &"' WHERE waiverid = " & iWaiverId
End If

RunSQLStatement sSql

' REDIRECT TO facility rates page
response.redirect( "facility_waivers.asp?facilityid=" & iFacilityId )


%>