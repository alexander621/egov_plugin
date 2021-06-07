<!--#Include file="../includes/common.asp"-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: regattateamupdate.asp
' AUTHOR: Steve Loar
' CREATED: 04/08/2010
' COPYRIGHT: Copyright 20109 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module updates regatta teams information 

'
' MODIFICATION HISTORY
' 1.0	04/08/2010	Steve Loar  - Initial Version
' 1.1	5/14/2010	Steve Loar - Split captain name into first and last
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sSql, iRegattaTeamId, sTeamName, sCaptainFirstName, sCaptainAddress, sCaptainCity
Dim sCaptainState, sCaptainZip, sCaptainPhone, iRegattaTeamGroupId, sCaptainLastName


iRegattaTeamId = CLng(request("regattateamid"))

sTeamName = "'" & dbsafe(request("regattateam")) & "'"

iRegattaTeamGroupId = CLng(request("regattateamgroupid"))

sCaptainFirstName = "'" & dbsafe(request("captainfirstname")) & "'"
sCaptainLastName = "'" & dbsafe(request("captainlastname")) & "'"
sCaptainAddress = "'" & dbsafe(request("captainaddress")) & "'"
sCaptainCity = "'" & dbsafe(request("captaincity")) & "'"
sCaptainState = "'" & dbsafe(request("captainstate")) & "'"
sCaptainZip = "'" & dbsafe(request("captainzip")) & "'"
sCaptainPhone = "'" & dbsafe(request("captainphone")) & "'"

sSql = "UPDATE egov_regattateams SET "
sSql = sSql & "regattateam = " & sTeamName
sSql = sSql & ", captainfirstname = " & sCaptainFirstName
sSql = sSql & ", captainlastname = " & sCaptainLastName
sSql = sSql & ", captainaddress = " & sCaptainAddress
sSql = sSql & ", captaincity = " & sCaptainCity
sSql = sSql & ", captainstate = " & sCaptainState
sSql = sSql & ", captainzip = " & sCaptainZip
sSql = sSql & ", captainphone = " & sCaptainPhone
sSql = sSql & ", regattateamgroupid = " & iRegattaTeamGroupId
sSql = sSql & " WHERE regattateamid = " & iRegattaTeamId

RunSQLStatement sSql

response.redirect "regattateamlist.asp?regattateamid=" & iRegattaTeamId & "&u=s"


%>