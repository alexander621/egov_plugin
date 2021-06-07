<%
Call subDeleteFamily(request("iFamilyMemberId"), request("iUserId"))


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB subDeleteRate(sResidentType, iRateid)
' AUTHOR: Steve Loar
' CREATED: 01/31/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subDeleteFamily(iFamilyMemberId, iUserId)
	
	' Delete from the family members table
	sSQL = "DELETE FROM egov_familymembers WHERE familymemberid = " & iFamilyMemberId  & ""

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

	' REDIRECT TO family members page
	response.redirect( "family_members.asp?userid=" & iUserId )

End Sub

%>