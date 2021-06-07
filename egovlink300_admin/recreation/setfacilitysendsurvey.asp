<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: setfacilitysendsurvey.asp
' AUTHOR: Steve Loar	
' CREATED: 11/05/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This sets the flag to include this facility in the nightly survey job. Called
'				from facility_management.asp via Ajax.
'
' MODIFICATION HISTORY
' 1.0   11/05/2007   Steve Loar - Initial code  
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim oCmd, iFacilityId, sSql, oRs, iSendFlag

iFacilityId = CLng(request("facilityid"))

sSql = "SELECT sendsurveys FROM egov_facility WHERE facilityid = " & iFacilityId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 0, 1

' set them to the opposite of what they are now set to
If oRs("sendsurveys") Then 
	iSendFlag = 0
Else
	iSendFlag = 1
End If 

oRs.Close
Set oRs = Nothing 

Set oCmd = Server.CreateObject("ADODB.Command")
oCmd.ActiveConnection = Application("DSN")

' only update the facility send flag if the orgid also matches.
oCmd.CommandText = "UPDATE egov_facility SET sendsurveys = " & iSendFlag & " WHERE facilityid = " & iFacilityId & " AND orgid = " & Session("OrgID")
oCmd.Execute

Set oCmd = Nothing

response.write "Completed"

%>