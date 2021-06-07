<!-- #include file="includes/inc_dbfunction.asp" //-->
<!-- #include file="includes/common.asp" //-->
<!-- #include file="include_top_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: update_family_members.asp
' AUTHOR: Steve Loar
' CREATED: 12/29/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This updates family information
'
' MODIFICATION HISTORY
' 1.0   12/29/2006	Steve Loar - Initial code 
' 2.0	10/05/2011	Steve Loar - Changed from class methods to dynamic SQL, and added gender selection.
'
' <!-- #include file="class/classFamily.asp" //--> - This was used but is no longer
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iUserId, iFamilyId, oFamily, sAction, bAddressChanged, sSql, sUserfname, sUserlname, sGender
Dim sUseraddress, sUsercity, sUserstate, sUserzip, sUserhomephone, sUsercell, sUserworkphone
Dim sUserfax, sUserbusinessname, sUserbusinessaddress, iNeighborhoodid,sEmergencycontact
Dim sEmergencyphone, sBirthDate, iRelationshipId, sResidentType, sResidencyVerified

bAddressChanged = False 

'Set oFamily = New classFamily

'iFamilyId = oFamily.GetFamilyId( request.cookies("userid") )
iFamilyId = GetFamilyId( request.cookies("userid") )
If iFamilyId = 0 Then 
	' if they are a bot, then the userid is not set and the familyid is 0'
	response.redirect "./"
End If 


If Len(request("egov_users_userfname")) < 2 OR Len(request("egov_users_userlname")) < 2 Then
	' This is too short and probably a bot since the UI catches these for real people
	response.redirect "./"
End If 

sUserfname = "'" & dbready_string(request("egov_users_userfname"),50) & "'"

sUserlname = "'" & dbready_string(request("egov_users_userlname"),50) & "'"

If request("egov_users_gender") <> "M" And request("egov_users_gender") <> "F" Then
	sGender = "NULL"
Else
	sGender = "'" & dbready_string(request("egov_users_gender"),1) & "'"
End If 

If request("egov_users_useraddress") <> "" Then 
	sUseraddress = "'" & dbready_string(request("egov_users_useraddress"),255) & "'"
Else
	sUseraddress = "NULL"
End If 

If request("egov_users_usercity") <> "" Then
	sUsercity = "'" & dbready_string(request("egov_users_usercity"),50) & "'"
Else
	sUsercity = "NULL"
End If 

If request("egov_users_userstate") <> "" Then
	sUserstate = UCase(dbready_string(request("egov_users_userstate"),2))
	sUserstate = "'" & sUserstate & "'"
Else
	sUserstate = "NULL"
End If 

If request("egov_users_userzip") <> "" Then 
	sUserzip = "'" & dbready_string(request("egov_users_userzip"),10) & "'"
Else
	sUserzip = "NULL"
End If 

If request("egov_users_userhomephone") <> "" Then
	sUserhomephone = "'" & dbready_string(request("egov_users_userhomephone"),10) & "'"
Else
	sUserhomephone = "NULL"
End If 

If request("egov_users_usercell") <> "" Then
	sUsercell = "'" & dbready_string(request("egov_users_usercell"),10) & "'"
Else
	sUsercell = "NULL"
End If 

If request("egov_users_userworkphone") <> "" Then
	sUserworkphone = "'" & dbready_string(request("egov_users_userworkphone"),14) & "'"
Else
	sUserworkphone = "NULL"
End If 

If request("egov_users_userfax") <> "" Then
	sUserfax = "'" & dbready_string(request("egov_users_userfax"),10) & "'"
Else
	sUserfax = "NULL"
End If 

If request("egov_users_userbusinessname") <> "" Then
	sUserbusinessname = "'" & dbready_string(request("egov_users_userbusinessname"),100) & "'"
Else
	sUserbusinessname = "NULL"
End If 

If request("egov_users_userbusinessaddress") <> "" Then 
	sUserbusinessaddress = "'" & dbready_string(request("egov_users_userbusinessaddress"),255) & "'"
Else
	sUserbusinessaddress = "NULL"
End If 

If request("egov_users_neighborhoodid") <> "" Then
	If dbready_number( request("egov_users_neighborhoodid") ) Then
		iNeighborhoodid = CLng(request("egov_users_neighborhoodid"))
	Else
		iNeighborhoodid = "NULL" 
	End If 
Else
	iNeighborhoodid = "NULL"
End If 

If request("egov_users_emergencycontact") <> "" Then
	sEmergencycontact = "'" & dbready_string(request("egov_users_emergencycontact"),100) & "'"
Else
	sEmergencycontact = "NULL"
End If

If request("egov_users_emergencyphone") <> "" Then
	sEmergencyphone = "'" & dbready_string(request("egov_users_emergencyphone"),10) & "'"
Else
	sEmergencyphone = "NULL"
End If 

If request("egov_users_birthdate") <> "" Then
	sBirthDate = "'" & dbready_string(request("egov_users_birthdate"),10) & "'"
Else
	sBirthDate = "NULL"
End If 

If request("egov_users_relationshipid") <> "" Then
	If dbready_number( request("egov_users_relationshipid") ) Then
		iRelationshipId = CLng(request("egov_users_relationshipid"))
	Else
		iRelationshipId = "NULL" 
	End If 
Else
	iRelationshipId = "NULL"
End If 

sResidentType = "'" & dbready_string(request("egov_users_residenttype"),1) & "'"
		if iOrgID = "60" then
			if request.form("residentstreetnumber") = "701" and request.form("skip_address") = "LAUREL ST" and lcase(request.form("egov_users_usercity")) = "menlo park" _
				and lcase(request.form("egov_users_userstate")) = "ca" and left(request.form("egov_users_userzip"),5) = "94025" then
				sResidentType = "'E'"
			elseif sResidentType = "'R'" then
			elseif request.form("egov_users_userbusinessname") <> "" and (request.form("skip_Baddress") <> "0000" OR request.form("egov_users_userbusinessaddress") <> "") then
				sResidentType = "'B'"
			elseif left(request.form("egov_users_userzip"),5) = "94025" and lcase(request.form("egov_users_usercity")) = "menlo park" then
				sResidentType = "'U'"
			else
				sResidentType = "'N'"
			end if
		end if

If request("egov_users_residencyverified") <> "" Then 
	If LCase(request("egov_users_residencyverified")) = "true" Then 
		sResidencyVerified = 1	' should be 0 or 1 only
	Else
		sResidencyVerified = 0
	End If 
Else
	sResidencyVerified = 0
End If 


If request("userid") <> "0" Then 
	' They exsist, so update them
	iUserId = CLng(request("userid"))

	' This was commented out because the egov_user table keeps changing, making this approach difficult to maintain.
'	oFamily.UpdateFamilymember iUserId, request("egov_users_userfname"), request("egov_users_userlname"), _
'		request("egov_users_useraddress"), request("egov_users_usercity"), request("egov_users_userstate"),request("egov_users_userzip"), _
'		request("egov_users_userhomephone"), request("egov_users_usercell"), request("egov_users_userfax"), request("egov_users_userworkphone"), request("egov_users_userbusinessname"), _
'		request("egov_users_userbusinessaddress"), request("egov_users_emergencycontact"), request("egov_users_emergencyphone"), request("egov_users_neighborhoodid"), _
'		request("egov_users_birthdate"), request("egov_users_relationshipid")

	sSql = "UPDATE EGOV_USERS "
	sSql = sSql & "SET userfname = " & sUserfname & ", "
	sSql = sSql & "userlname = " & sUserlname & ", "
	sSql = sSql & "useraddress = " & sUseraddress & ", "
	sSql = sSql & "usercity = " & sUsercity & ", "
	sSql = sSql & "userstate = " & sUserstate & ", "
	sSql = sSql & "userzip = " & sUserzip & ", "
	sSql = sSql & "userhomephone = " & sUserhomephone & ", "
	sSql = sSql & "usercell = " & sUsercell & ", "
	sSql = sSql & "userworkphone = " & sUserworkphone & ", "
	sSql = sSql & "userfax = " & sUserfax & ", "
	sSql = sSql & "userbusinessname = " & sUserbusinessname & ", "
	sSql = sSql & "userbusinessaddress = " & sUserbusinessaddress & ", "
	sSql = sSql & "emergencycontact = " & sEmergencycontact & ", "
	sSql = sSql & "emergencyphone = " & sEmergencyphone & ", "
	sSql = sSql & "neighborhoodid = " & iNeighborhoodid & ", "
	sSql = sSql & "birthdate = " & sBirthDate & ", "
	sSql = sSql & "relationshipid = " & iRelationshipId & ", "
	sSql = sSql & "residenttype = " & sResidentType & ", "
	sSql = sSql & "gender = " & sGender 
	sSql = sSql & " WHERE userid = " & iUserId
'if request.cookies("userid") = "1150705" then
'response.write sSql
'response.end
'end if

	RunSQLStatement sSql		' In common.asp

	sAction = "UPDATE"

	' check for address changes and unflag the residency verified flag in egov_users
	If request("skip_old_egov_users_useraddress") <> request("egov_users_useraddress") Then 
		bAddressChanged = True 
	Else
		If request("skip_old_egov_users_neighborhoodid") <> request("egov_users_neighborhoodid") Then 
			bAddressChanged = True 
		Else
			If request("skip_old_egov_users_usercity") <> request("egov_users_usercity") Then 
				bAddressChanged = True 
			Else
				If request("skip_old_egov_users_userstate") <> request("egov_users_userstate") Then 
					bAddressChanged = True 
				Else
					If request("skip_old_egov_users_userzip") <> request("egov_users_userzip") Then 
						bAddressChanged = True 
					End If
				End If
			End If
		End If 
	End If 

	If bAddressChanged Then 
		UpdateResidencyVerified iUserId 
	End If 

Else
	' Create the Family member's row
'	iUserId = oFamily.InsertFamilymember(request("egov_users_orgid"), request("egov_users_userfname"), request("egov_users_userlname"), _
'		request("egov_users_useraddress"), request("egov_users_usercity"), request("egov_users_userstate"),request("egov_users_userzip"), _
'		request("egov_users_userhomephone"), request("egov_users_usercell"), request("egov_users_userfax"), request("egov_users_userworkphone"), request("egov_users_userbusinessname"), _
'		request("egov_users_userbusinessaddress"), request("egov_users_emergencycontact"), request("egov_users_emergencyphone"), request("egov_users_neighborhoodid"), _
'		request("egov_users_birthdate"), request("egov_users_relationshipid"), request("egov_users_residencyverified"), request("egov_users_residenttype"), iFamilyId )

	sSql = "INSERT INTO egov_users ( userfname, userlname, useraddress ,usercity, userstate, userzip, userhomephone, "
	sSql = sSql & "userworkphone, userbusinessname, orgid, userregistered, userbusinessaddress, emergencycontact, "
	sSql = sSql & "emergencyphone, neighborhoodid, birthdate, relationshipid, residencyverified, familyid, "
	sSql = sSql & "residenttype, usercell, gender ) VALUES ( "
	sSql = sSql & sUserfname & ", "
	sSql = sSql & sUserlname & ", "
	sSql = sSql & sUseraddress & ", "
	sSql = sSql & sUsercity & ", "
	sSql = sSql & sUserstate & ", "
	sSql = sSql & sUserzip & ", "
	sSql = sSql & sUserhomephone & ", "
	sSql = sSql & sUserworkphone & ", "
	sSql = sSql & sUserbusinessname & ", "
	sSql = sSql & iOrgId & ", "
	sSql = sSql & "1, "
	sSql = sSql & sUserbusinessaddress & ", "
	sSql = sSql & sEmergencycontact & ", "
	sSql = sSql & sEmergencyphone & ", "
	sSql = sSql & iNeighborhoodid & ", "
	sSql = sSql & sBirthDate & ", "
	sSql = sSql & iRelationshipId & ", "
	sSql = sSql & sResidencyVerified & ", "
	sSql = sSql & iFamilyId & ", "
	sSql = sSql & sResidentType & ", "
	sSql = sSql & sUsercell & ", "
	sSql = sSql & sGender
	sSql = sSql & " )"

	iUserId = RunIdentityInsertStatement( sSql )	' in common.asp

	sAction = "INSERT"
End If 

' Update their FamilyId in egov_Users (found in includes/inc_dbfunction.asp)
'oFamily.UpdateFamilyId userid, iFamilyId, request("egov_users_relationshipid"), request("egov_users_neighborhoodid")

' This is the old family member table
FamilyMemberUpdate iFamilyId, iUserId, request("egov_users_userfname"), request("egov_users_userlname"), request("skip_egov_users_relationship"), request("egov_users_birthdate"), sAction 

'Set oFamily = Nothing 

response.redirect "family_list.asp"


'------------------------------------------------------------------------------------------------------------
' FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void FamilyMemberUpdate iuserid, iorgid, usertype, rateid, firstname, lastname, relation, birthdate
'--------------------------------------------------------------------------------------------------
Sub FamilyMemberUpdate( ByVal ibelongstouserid, ByVal iUserId, ByVal firstname, ByVal lastname, relation, ByVal birthdate, ByVal sAction )
	Dim sSql

	firstname = DBsafe(Proper(firstname))
	lastname = DBsafe(Proper(lastname))

	If Trim(birthdate) = "" Then 
		birthdate =  " NULL " 
	Else 
		If IsDate( birthdate ) Then 
			birthdate = " '" & CDate(birthdate) & "' "
		Else
			birthdate =  " NULL "
		End If 
	End If 

	If sAction = "INSERT" Then 
	' Insert new records
		sSql = "INSERT INTO egov_familymembers (firstname, lastname, relationship, birthdate, belongstouserid, userid) Values ('" & firstname & "', '" & lastname & "', '" & relation & "', " & birthdate & ", " & ibelongstouserid & ", " & iUserId & " )"
	Else
		sSql = "UPDATE egov_familymembers SET firstname = '" & firstname
		sSql = sSql & "', lastname = '" & lastname
		sSql = sSql & "', relationship = '" & relation
		sSql = sSql & "', birthdate = " & birthdate
		sSql = sSql & " WHERE userid = " & iUserId & ""
	End If 

	'response.write sSQL
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' string Proper( sString )
'--------------------------------------------------------------------------------------------------
Function Proper( ByVal sString )

	Proper = sString

	If Len(sString) > 0 then
		Proper = UCase(Left(sString,1)) & Mid(sString,2)
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' void UpdateResidencyVerified iUserId 
'--------------------------------------------------------------------------------------------------
Sub UpdateResidencyVerified( ByVal iUserId )
	Dim sSql, oCmd

	sSql = "UPDATE egov_users SET residencyverified = 0 WHERE userid = " & iUserId 

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute

	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' integer GetFamilyId( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetFamilyId( ByVal iUserId )
	Dim sSql, oRs, iFamilyId

	iFamilyId = 0

	If Trim(iUserId) <> "" Then 
		If IsNumeric( Trim(iUserId) ) Then
			sSql = "SELECT ISNULL(familyid,0) AS familyid "
			sSql = sSql & " FROM egov_users "
			sSql = sSql & " WHERE userid = " & CLng(iUserId)
			session("sSql") = sSql

			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open sSQL, Application("DSN"), 0, 1
			session("sSql") = ""

			If Not oRs.EOF Then 
				iFamilyId = oRs("familyid")
			Else
				iFamilyId = iUserID
			End If 

			If iFamilyId = 0 Then 
				iFamilyId = iUserID
			End If 

			oRs.Close
			Set oRs = Nothing 
		End If 
	End If 
		
	GetFamilyID = iFamilyId

End Function 

%>

