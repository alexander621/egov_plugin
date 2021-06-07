<!-- #include file="../includes/common.asp" //-->
<!--#Include file="class_global_functions.asp"-->  
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: regattamemberdelete.asp
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
Dim iRegattaTeamMemberId, sSql, iRegattaTeamId, iMemberCount, iClassListId, iPaymentid

iRegattaTeamMemberId = CLng(request("regattateammemberid"))
iRegattaTeamId = CLng(request("regattateamid"))

iClassListId = GetTeamMemberClassListId( iRegattaTeamMemberId )

' Delete them from the team
sSql = "DELETE FROM egov_regattateammembers WHERE regattateammemberid = " & iRegattaTeamMemberId
sSql = sSql & " AND orgid = " & session("orgid")
RunSQLStatement sSql

' Update the team with the new team count
iMemberCount = GetRegattaTeamMemberCount( iRegattaTeamId )	' in class_global_functions.asp
sSql = "UPDATE egov_regattateams SET membercount = " & iMemberCount 
sSql = sSql & " WHERE regattateamid = " & iRegattaTeamId 
sSql = sSql & " AND orgid = " & session("orgid")
RunSQLStatement sSql

If Not MoreMembersOnClasslistid( iClassListId ) Then
	' Get the paymentid
	iPaymentid = GetPaymentIdFromClassListId( iClassListId )

	' Remove the class list record
	sSql = "DELETE FROM egov_class_list WHERE classlistid = " & iClassListId
	RunSQLStatement sSql

	' Remove the accounts ledger record
	sSql = "DELETE FROM egov_accounts_ledger WHERE paymentid = " & iPaymentid
	sSql = sSql & " AND orgid = " & session("orgid")
	RunSQLStatement sSql

	' Remove the class_payment record
	sSql = "DELETE FROM egov_class_payment WHERE paymentid = " & iPaymentid
	sSql = sSql & " AND orgid = " & session("orgid")
	RunSQLStatement sSql
End If 

response.redirect "regattateamlist.asp?regattateamid=" & iRegattaTeamId



'--------------------------------------------------------------------------------------------------
' Function GetTeamMemberClassListId( iRegattaTeamMemberId )
'--------------------------------------------------------------------------------------------------
Function GetTeamMemberClassListId( iRegattaTeamMemberId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(classlistid,0) AS classlistid FROM egov_regattateammembers "
	sSql = sSql & " WHERE regattateammemberid = " & iRegattaTeamMemberId
	sSql = sSql & " AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetTeamMemberClassListId = oRs("classlistid")
	Else
		GetTeamMemberClassListId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' Function MoreMembersOnClasslistid( iClassListId )
'--------------------------------------------------------------------------------------------------
Function MoreMembersOnClasslistid( iClassListId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(regattateammemberid) AS hits FROM egov_regattateammembers "
	sSql = sSql & " WHERE classlistid = " & iClassListId
	sSql = sSql & " AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			MoreMembersOnClasslistid = True 
		Else
			MoreMembersOnClasslistid = False 
		End If 
	Else
		MoreMembersOnClasslistid = False 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetPaymentIdFromClassListId( iClassListId )
'--------------------------------------------------------------------------------------------------
Function GetPaymentIdFromClassListId( iClassListId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(paymentid,0) AS paymentid FROM egov_class_list "
	sSql = sSql & " WHERE classlistid = " & iClassListId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetPaymentIdFromClassListId = oRs("paymentid")
	Else
		GetPaymentIdFromClassListId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 



%>