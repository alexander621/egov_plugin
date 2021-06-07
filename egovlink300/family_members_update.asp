<%

Call subFamilyUpdate(CLng(request("iuserid")), CLng(request("familymemberid")), request("firstname"), request("lastname"), request("relation"), request("birthdate"))


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB subFamilyupdate(iuserid, iorgid, usertype, rateid, firstname, lastname, relation, birthdate)
' AUTHOR: Steve Loar
' CREATED: 01/30/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subFamilyUpdate(iuserid, familymemberid, firstname, lastname, relation, birthdate)
	Dim sSql

	firstname = DBsafe(Proper(firstname))
	lastname = DBsafe(Proper(lastname))

	If Trim(birthdate) = "" Then 
		birthdate =  " NULL " 
	Else 
		birthdate = " '" & birthdate & "' "
	End If 

	If clng(familymemberid) = 0 Then 
	' Insert new records
		sSql = "INSERT INTO egov_familymembers (firstname, lastname, relationship, birthdate, belongstouserid) Values ('" & firstname & "', '" & lastname & "', '" & relation & "', " & birthdate & ", " & iuserid & " )"
	Else
		sSql = "Update egov_familymembers Set firstname = '" & firstname
		sSql = sSql & "', lastname = '" & lastname
		sSql = sSql & "', relationship = '" & relation
		sSql = sSql & "', birthdate = " & birthdate
		sSql = sSql & " Where familymemberid = " & familymemberid & ""
	End If 

	'response.write sSQL
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

	' REDIRECT TO family members page
	response.redirect( "family_members.asp?userid=" & iuserid )

End Sub


'--------------------------------------------------------------------------------------------------
' Function Proper( sString )
'--------------------------------------------------------------------------------------------------
Function Proper( sString )
	Proper = sString
	If Len(sString) > 0 then
		Proper = UCase(Left(sString,1)) & Mid(sString,2)
	End If 
End Function 


'--------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'--------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function


%>

