<%

Call subFamilyAdd(request("iuserid"), request("iorgid"), request("usertype"), request("rateid"), request("firstname"), request("lastname"), request("relation"), request("birthdate"))


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB subFamilyAdd(iuserid, iorgid, usertype, rateid, firstname, lastname, relation, birthdate)
' AUTHOR: Steve Loar
' CREATED: 01/30/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subFamilyAdd(iuserid, iorgid, usertype, rateid, firstname, lastname, relation, birthdate)
	
	' Insert new records
	sSql = "INSERT INTO egov_familymembers (firstname, lastname, relationship, birthdate, belongstouserid) Values ('" & firstname & "', '" & lastname & "', '" & relation & "', '" & birthdate & "', " & iuserid & " )"
'	response.write sSQL
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

	' REDIRECT TO select members page
	response.redirect( "select_members.asp?iuserid=" & iuserid & "&iorgid=" & iorgid & "&usertype=" & usertype & "&rateid=" & rateid )

End Sub

%>

