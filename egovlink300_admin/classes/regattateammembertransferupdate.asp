<!-- #include file="../includes/common.asp" //-->
<!--#Include file="class_global_functions.asp"-->  
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: regattateammembertransferupdate.asp
' AUTHOR: Steve Loar
' CREATED: 08/03/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module transfer a team member to another team.
'
' MODIFICATION HISTORY
' 1.0   08/03/2009   Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRegattaTeamMemberId, sSql, iRegattaTeamId, iOriginalRegattaTeamId, iMemberCount

iRegattaTeamMemberId = CLng(request("regattateammemberid"))
iRegattaTeamId = CLng(request("regattateamid"))
iOriginalRegattaTeamId = CLng(request("originalregattateamid"))

' Transfer the team member to the new team
sSql = "UPDATE egov_regattateammembers SET regattateamid = " & iRegattaTeamId 
sSql = sSql & " WHERE regattateammemberid = " & iRegattaTeamMemberId
sSql = sSql & " AND orgid = " & session("orgid")
RunSQLStatement sSql

' Update the new team with the new team count
iMemberCount = GetRegattaTeamMemberCount( iRegattaTeamId )	' in class_global_functions.asp
sSql = "UPDATE egov_regattateams SET membercount = " & iMemberCount 
sSql = sSql & " WHERE regattateamid = " & iRegattaTeamId 
sSql = sSql & " AND orgid = " & session("orgid")
RunSQLStatement sSql

' Update the old team with the new team count
iMemberCount = GetRegattaTeamMemberCount( iOriginalRegattaTeamId )	' in class_global_functions.asp
sSql = "UPDATE egov_regattateams SET membercount = " & iMemberCount 
sSql = sSql & " WHERE regattateamid = " & iOriginalRegattaTeamId 
sSql = sSql & " AND orgid = " & session("orgid")
RunSQLStatement sSql

response.redirect "regattateamlist.asp?regattateamid=" & iRegattaTeamId


%>