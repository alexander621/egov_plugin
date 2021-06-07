<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitapplicantsave.asp
' AUTHOR: Steve Loar
' CREATED: 02/21/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the permit applicants and Primary contacts
'
' MODIFICATION HISTORY
' 1.0   02/21/2008   Steve Loar - INITIAL VERSION
' 1.1	03/27/2008	Steve Loar	- Modified to handle permit contacts also
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iUserId, sSql, iEmailNotAvailable, sUserEmail, sUserPassword, sUserCity, sUserState, sUserZip
Dim sUserHomePhone, sUserCell, sUserWorkPhone, sUserFax, sEmergencyPhone, iNeighborhoodId, sUserUnit
Dim sUserBusinessName, sEmergencyContact, sUserAddress, sUserBusinessAddress, sResidencyVerified
Dim sRelationship, iIsOwner, iIsContractor, iIsArchitect, oRs, iPermitId, iPermitStatusId, bChangesPropagate
Dim bNotifyUser, sNotifyUserEmail, sToName, sSubject, sHTMLBody, Time0

iUserId = CLng(request("userid"))
iPermitId = CLng(request("permitid"))
iPermitStatusId = CLng(request("permitstatusid"))
bNotifyUser = False 

If iPermitId > CLng(0) Then 
	bChangesPropagate = StatusAllowsChangesToPropagate( iPermitStatusId )
Else
	bChangesPropagate = True 
End If 
'response.write "bChangesPropagate = " & bChangesPropagate & "<br /><br />"

sRelationship = ""
Time0 = timer 

'response.write "EmailNotAvailable: " & request("emailnotavailable") & "<br />"
'response.write "NotifyUser: " & request("notifyuser") & "<br />"

' ** In regular Posts check boxes come across as 'on', but in AJAX Post they are 'true'
If LCase(request("emailnotavailable")) = "true" Then
	iEmailNotAvailable = 1
	sUserEmail = "NULL"
	sUserPassword = "NULL"
Else
	iEmailNotAvailable = 0
	sUserEmail = "'" & dbsafe(request("useremail")) & "'"
	sUserPassword = "'" & dbsafe(request("userpassword")) & "'"
	If LCase(request("notifyuser")) = "true" Then
		bNotifyUser = True 
		sNotifyUserEmail = request("useremail")
		sToName = request("userfname") & " " & request("userlname")
	End If 
End If 

If request("usercity") = "" Then
	sUserCity = "NULL"
Else
	sUserCity = "'" & dbsafe(request("usercity")) & "'"
End If 

If request("userstate") = "" Then
	sUserState = "NULL"
Else
	sUserState = "'" & dbsafe(request("userstate")) & "'"
End If 

If request("userzip") = "" Then
	sUserZip = "NULL"
Else
	sUserZip = "'" & dbsafe(request("userzip")) & "'"
End If 

If request("userhomephone") = "" Then
	sUserHomePhone = "NULL"
Else
	sUserHomePhone = "'" & dbsafe(request("userhomephone")) & "'"
End If 

If request("usercell") = "" Then
	sUserCell = "NULL"
Else
	sUserCell = "'" & dbsafe(request("usercell")) & "'"
End If 

If request("userworkphone") = "" Then
	sUserWorkPhone = "NULL"
Else
	sUserWorkPhone = "'" & dbsafe(request("userworkphone")) & "'"
End If 

If request("userfax") = "" Then
	sUserFax = "NULL"
Else
	sUserFax = "'" & dbsafe(request("userfax")) & "'"
End If 

If request("emergencyphone") = "" Then
	sEmergencyPhone = "NULL"
Else
	sEmergencyPhone = "'" & dbsafe(request("emergencyphone")) & "'"
End If 

If CLng(request("neighborhoodid")) = CLng(0) Then
	iNeighborhoodId = "NULL"
Else
	iNeighborhoodId = CLng(request("neighborhoodid"))
End If 

If request("userunit") = "" Then
	sUserUnit = "NULL"
Else
	sUserUnit = "'" & dbsafe(request("userunit")) & "'"
End If 

If request("userbusinessname") = "" Then
	sUserBusinessName = "NULL"
Else
	sUserBusinessName = "'" & dbsafe(request("userbusinessname")) & "'"
End If 

If request("emergencycontact") = "" Then
	sEmergencyContact = "NULL"
Else
	sEmergencyContact = "'" & dbsafe(request("emergencycontact")) & "'"
End If 

If request("useraddress") = "" Then
	sUserAddress = "NULL"
Else
	sUserAddress = "'" & dbsafe(request("useraddress")) & "'"
End If 

If request("userbusinessaddress") = "" Then
	sUserBusinessAddress = "NULL"
Else
	sUserBusinessAddress = "'" & dbsafe(request("userbusinessaddress")) & "'"
End If 

sResidencyVerified = request("residencyverified")
'response.write "One: " & FormatNumber((timer - Time0),3) & " seconds<br />"


If iUserId > CLng(0) Then 
	If bChangesPropagate Then 
		' Update an existing account
		sSql = "UPDATE egov_users SET userfname = '" & dbsafe(request("userfname")) & "' "
		sSql = sSql & ", userlname = '" & dbsafe(request("userlname")) & "' "
		sSql = sSql & ", emailnotavailable = " & iEmailNotAvailable 
		sSql = sSql & ", useremail = " & sUserEmail 
		sSql = sSql & ", userpassword = " & sUserPassword 
		sSql = sSql & ", residenttype = '" & dbsafe(request("residenttype")) & "' "
		sSql = sSql & ", usercity = " & sUserCity 
		sSql = sSql & ", userstate = " & sUserState 
		sSql = sSql & ", userzip = " & sUserZip 
		sSql = sSql & ", userhomephone = " & sUserHomePhone 
		sSql = sSql & ", usercell = " & sUserCell 
		sSql = sSql & ", userworkphone = " & sUserWorkPhone 
		sSql = sSql & ", userfax = " & sUserFax 
		sSql = sSql & ", emergencyphone = " & sEmergencyPhone
		sSql = sSql & ", neighborhoodid = " & iNeighborhoodId
		sSql = sSql & ", userunit = " & sUserUnit 
		sSql = sSql & ", userbusinessname = " & sUserBusinessName
		sSql = sSql & ", emergencycontact = " & sEmergencyContact
		sSql = sSql & ", useraddress = " & sUserAddress 
		sSql = sSql & ", userbusinessaddress = " & sUserBusinessAddress 
		sSql = sSql & " WHERE userid = " & iUserId & " AND orgid = " & session("orgid")
		'response.write sSql & "<br /><br />"
		RunSQL sSql 
		'response.write "Update egov_users " & FormatNumber((timer - Time0),3) & " seconds<br />"

		' Sync up the family member table
		sSql = "UPDATE egov_familymembers SET firstname = '" & DBsafe( request("userfname") )
		sSql = sSql & "', lastname = '" &  DBsafe( request("userlname") )
		sSql = sSql & "' WHERE userid = " & iUserId
		'response.write sSql & "<br /><br />"
		RunSQL sSql
		'response.write "Update egov_familymembers " & FormatNumber((timer - Time0),3) & " seconds<br />"

		' If shared family data was changed, then update it for the family
		If request("familyaddresschanged") = "YES" Then
			sSql = "UPDATE egov_users SET userhomephone = " & sUserHomePhone
			sSql = sSql & ", useraddress = " & sUserAddress
			sSql = sSql & ", usercity = " & sUserCity
			sSql = sSql & ", userstate = " & sUserState
			sSql = sSql & ", userzip = " & sUserZip
			sSql = sSql & ", userunit = " & sUserUnit 
			sSql = sSql & " WHERE familyid = " & request("familyid") & " AND userid <> " & iUserId
			'response.write sSql & "<br /><br />"
			RunSQL sSql
			'response.write "Update egov_users (other family) " & FormatNumber((timer - Time0),3) & " seconds<br />"
		End If 

		' Update any permit applicants and Primary Contacts where the permit is still open. Pull them, then loop through the set
		sSql = "SELECT P.permitid, C.permitcontactid FROM egov_permits P, egov_permitcontacts C, egov_permitstatuses S "
		sSql = sSql & " WHERE P.permitid = C.permitid AND P.permitstatusid = S.permitstatusid AND (isapplicant = 1 OR isprimarycontact = 1) "
		sSql = sSql & " AND S.iscompletedstatus = 0 AND S.cansavechanges = 1 AND S.changespropagate = 1 AND C.userid = " & iUserId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		Do While Not oRs.EOF
			sSql = "UPDATE egov_permitcontacts SET firstname = '" & dbsafe(request("userfname")) & "' "
			sSql = sSql & ", lastname = '" & dbsafe(request("userlname")) & "' "
			sSql = sSql & ", company = " & sUserBusinessName
			sSql = sSql & ", address = " & sUserAddress
			sSql = sSql & ", city = " & sUserCity
			sSql = sSql & ", state = " & sUserState
			sSql = sSql & ", zip = " & sUserZip
			sSql = sSql & ", email = " & sUserEmail
			sSql = sSql & ", phone = " & sUserHomePhone
			sSql = sSql & ", fax = " & sUserFax
			sSql = sSql & ", cell = " & sUserCell
			sSql = sSql & ", emailnotavailable = " & iEmailNotAvailable 
			sSql = sSql & ", userpassword = " & sUserPassword 
			sSql = sSql & ", residenttype = '" & dbsafe(request("residenttype")) & "' "
			sSql = sSql & ", userworkphone = " & sUserWorkPhone
			sSql = sSql & ", emergencyphone = " & sEmergencyPhone 
			sSql = sSql & ", neighborhoodid = " & iNeighborhoodId
			sSql = sSql & ", userunit = " & sUserUnit 
			sSql = sSql & ", emergencycontact = " & sEmergencyContact
			sSql = sSql & ", userbusinessaddress = " & sUserBusinessAddress
			sSql = sSql & " WHERE permitid = " & oRs("permitid") & " AND permitcontactid = " & oRs("permitcontactid")
			'response.write sSql & "<br /><br />"
			RunSQL sSql
			'response.write "Update egov_permitcontacts " & FormatNumber((timer - Time0),3) & " seconds<br />"

			oRs.MoveNext
		Loop
		oRs.Close 
		Set oRs = Nothing 


	Else	' Do not propogate but update this contact for this permit
		sSql = "UPDATE egov_permitcontacts SET firstname = '" & dbsafe(request("userfname")) & "' "
		sSql = sSql & ", lastname = '" & dbsafe(request("userlname")) & "' "
		sSql = sSql & ", company = " & sUserBusinessName
		sSql = sSql & ", address = " & sUserAddress
		sSql = sSql & ", city = " & sUserCity
		sSql = sSql & ", state = " & sUserState
		sSql = sSql & ", zip = " & sUserZip
		sSql = sSql & ", email = " & sUserEmail
		sSql = sSql & ", phone = " & sUserHomePhone
		sSql = sSql & ", fax = " & sUserFax
		sSql = sSql & ", cell = " & sUserCell
		sSql = sSql & ", emailnotavailable = " & iEmailNotAvailable 
		sSql = sSql & ", userpassword = " & sUserPassword 
		sSql = sSql & ", residenttype = '" & dbsafe(request("residenttype")) & "' "
		sSql = sSql & ", userworkphone = " & sUserWorkPhone
		sSql = sSql & ", emergencyphone = " & sEmergencyPhone 
		sSql = sSql & ", neighborhoodid = " & iNeighborhoodId
		sSql = sSql & ", userunit = " & sUserUnit 
		sSql = sSql & ", emergencycontact = " & sEmergencyContact
		sSql = sSql & ", userbusinessaddress = " & sUserBusinessAddress
		sSql = sSql & " WHERE permitid = " & iPermitId & "  AND (isapplicant = 1 OR isprimarycontact = 1) AND userid = " & iUserId
		'response.write sSql & "<br /><br />"
		RunSQL sSql
		'response.write "Update egov_permitcontacts " & FormatNumber((timer - Time0),3) & " seconds<br />"
	End If 

	sSubject = Session("sOrgName") & " Web Site Registration Change"
	sHTMLBody = "The account you use to access the e-government features of the " & Session("sOrgName") & " web site has been updated.  "
	sHTMLBody = sHTMLBody & "To view the changes to your account please go to <a href=""" & session("egovclientwebsiteurl") & "/user_login.asp"">" & session("egovclientwebsiteurl") & "/user_login.asp.</a>"

Else
	' Insert a new applicant into egov_users
	sSql = "INSERT INTO egov_users ( userfname, userlname, emailnotavailable, useremail, userpassword, residenttype, "
	sSql = sSql & " usercity, userstate, userzip, userhomephone, usercell, userworkphone, userfax, emergencyphone, "
	sSql = sSql & " neighborhoodid, userunit, userbusinessname, emergencycontact, useraddress, userbusinessaddress, "
	sSql = sSql & " userregistered, residencyverified, relationshipid, headofhousehold, orgid ) VALUES ( "
	sSql = sSql & "'" & dbsafe(request("userfname")) & "', '" & dbsafe(request("userlname")) & "', " & iEmailNotAvailable & ", "
	sSql = sSql & sUserEmail & ", " & sUserPassword & ", '" & dbsafe(request("residenttype")) & "', " & sUserCity
	sSql = sSql & ", " & sUserState & ", " & sUserZip & ", " & sUserHomePhone & ", " & sUserCell & ", " & sUserWorkPhone
	sSql = sSql & ", " & sUserFax & ", " & sEmergencyPhone & ", " & iNeighborhoodId & ", " & sUserUnit & ", " 
	sSql = sSql & sUserBusinessName & ", " & sEmergencyContact & ", " & sUserAddress & ", " & sUserBusinessAddress
	sSql = sSql & ", 1, 1, " & GetRelationshipId( sRelationship ) & ", 1, " & session("orgid") & " )" 

'	response.write sSql
'	response.End
	
	iUserId = RunIdentityInsert( sSql )

	' Update their familyid
	sSql = "UPDATE egov_users SET familyid = " & iUserId & " WHERE userid = " & iUserId
	RunSQL sSql

	' Create their family member record
	AddFamilyMember iUserId, request("userfname"), request("userlname"), sRelationship, "NULL", iUserId

	' If the Permit contact field is checked, then add them to the contacts table
'	If LCase(request("iscontact")) = "true" Then
'		iIsOwner = 1
'		iIsContractor = 1
'		iIsArchitect = 1
'		sSql = "INSERT INTO egov_permitcontacttypes ( company, firstname, lastname, address, city, state, zip, email, "
'		sSql = sSql & " phone, cell, fax, isowner, iscontractor, isarchitect, userid, orgid ) VALUES ( "
'		sSql = sSql & sUserBusinessName & ", '" & dbsafe(request("userfname")) & "', '" & dbsafe(request("userlname")) & "', " 
'		sSql = sSql & sUserAddress & ", " & sUserCity & ", " & sUserState & ", " & sUserZip & ", " & sUserEmail & ", " 
'		sSql = sSql & sUserHomePhone & ", " & sUserCell & ", " & sUserFax & ", " & iIsOwner & ", " & iIsContractor & ", "
'		sSql = sSql & iIsArchitect & ", " & iUserId & ", " & session("orgid") & " )"
'		RunSQL sSql
'	End If 

	sSubject = Session("sOrgName") & " Web Site Registration"
	sHTMLBody = "An account was created for you to access the e-government features of the " & Session("sOrgName") & " web site.  "
	sHTMLBody = sHTMLBody & "To access your account please go to <a href=""" & session("egovclientwebsiteurl") & "/user_login.asp"">" & session("egovclientwebsiteurl") & "/user_login.asp</a>.  Your username and password are:"
	sHTMLBody = sHTMLBody & vbcrlf & vbcrlf & "<br /><br />Username: " & sToAddress
	sHTMLBody = sHTMLBody & vbcrlf & "<br />Password: " & sPassword
	sHTMLBody = sHTMLBody & vbcrlf & vbcrlf & "<br /><br />This is a temporary password so please be sure to change it to something you can remember. "

End If

If bNotifyUser Then
'	response.write "sToName = " & sToName & "<br />"
'	response.write "sNotifyUserEmail = " & sNotifyUserEmail & "<br />"
'	response.write "sSubject = " & sSubject & "<br />"
'	response.write "sHTMLBody = " & sHTMLBody & "<br />"
	' SendEmail( sToName, sToEmail, sFromName, sFromEmail, sSubject, sHTMLBody )
	'SendEmailPermits sToName, sNotifyUserEmail, Session("sOrgName") & " E-GOV WEBSITE", "webmaster@eclink.com", sSubject, sHTMLBody 
	' sendEmail is in common.asp
	'sendEmail "", sNotifyUserEmail & "(" & sToName & ")", "", sSubject, sHTMLBody, "", ""	
	sendEmail "", sToName & " <" & sNotifyUserEmail & ">", "", sSubject, sHTMLBody, "", "" 
	'response.write "Email Sent " & FormatNumber((timer - Time0),3) & " seconds<br />"
'Else
'	response.write "No Email Sent" & "<br />"
End If 

response.write iUserId
'response.write "<br />Done " & FormatNumber((timer - Time0),3) & " seconds<br />"


'-------------------------------------------------------------------------------------------------
' Function GetRelationshipId()
'-------------------------------------------------------------------------------------------------
Function GetRelationshipId( ByRef sRelationship )
	Dim sSql, oRs

	sSql = "SELECT relationshipid, relationship FROM egov_familymember_relationships WHERE isdefault = 1 AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetRelationshipId = oRs("relationshipid")
		sRelationship = oRs("relationship")
	Else
		GetRelationshipId = "NULL"
		sRelationship = "Spouce"
	End If
	
	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' Function AddFamilyMember( iBelongsToUserId, sFirstName, sLastName, sRelationship, sBirthDate, iUserId )
'--------------------------------------------------------------------------------------------------
Sub AddFamilyMember( iBelongsToUserId, sFirstName, sLastName, sRelationship, sBirthDate, iUserId )
	' This function adds family members to the egov_familymembers table
	Dim sSql

	sSql = "Insert Into egov_familymembers (firstname, lastname, birthdate, belongstouserid, relationship, userid) VALUES ('"
	sSql = sSql & DBsafe( sFirstName ) & "', '" & DBsafe( sLastName ) & "', " & sBirthDate & ", " & iBelongsToUserId & ", '" & sRelationship & "', " & iUserId & " )"
	RunSQL sSql

End Sub 

%>
