<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: mergecitizensmerge.asp
' AUTHOR: Steve Loar
' CREATED: 12/30/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Merges the households selected and mapped from mergecitizensmatch.asp
'
' MODIFICATION HISTORY
' 1.0   12/30/2008	Steve Loar - INITIAL VERSION
' 1.1	08/27/2012	Steve Loar - Added missing merge of rental reservations
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, iKeepFamilyId, iMergeFamilyId, iMaxMerge, iKeepHeadOfHouseholdId, iMergeHeadOfHouseholdId
Dim sUserfname, sUserlname, sUserBusinessName, sUserAddress, sUserCity, sUserState, sUserZip, sUserEmail
Dim sUserHomePhone, sUserFax, sUserCell, iEmailNotAvailable, sResidencyVerified, sUserPassword, sResidentType
Dim sUserWorkPhone, sEmergencyPhone, iNeighborhoodId, sUserUnit, sEmergencyContact, sUserBusinessAddress
Dim x, iMergeFamilyMemberId, iKeepFamilyMemberId

iKeepFamilyId = CLng(request("keepfamilyid"))
iMergeFamilyId = CLng(request("mergefamilyid"))
iMaxMerge = CLng(request("maxmerge"))
x = CLng(0)

iKeepHeadOfHouseholdId = GetHeadOfHouseholdId( iKeepFamilyId )
iMergeHeadOfHouseholdId = GetHeadOfHouseholdId( iMergeFamilyId )

' First on the list is the head of household. The rest are family members

' if the head of households are different then move household level things
If iMergeHeadOfHouseholdId <> iKeepHeadOfHouseholdId Then
	' First move the data related to being the head of household
	iMergeUserId = CLng(request("mergeuserid1"))
	iKeepUserId = CLng(request("keepuserid1"))

	' Handle Subscriptions and Job and Bid Postings
	' Remove the duplicate distribution lists
	sSql = "SELECT rowid FROM egov_class_distributionlist_to_user WHERE userid = " & iMergeUserId
	sSql = sSql & " AND distributionlistid IN (SELECT distributionlistid "
	sSql = sSql & " FROM egov_class_distributionlist_to_user WHERE userid = " & iKeepHeadOfHouseholdId & " )"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		sSql = "DELETE FROM egov_class_distributionlist_to_user WHERE rowid = " & oRs("rowid")
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql		' In common.asp
		oRs.MoveNext 
	Loop
	oRs.Close
	Set oRs = Nothing 

	' Merge the remaining distribution lists by changing the userid
	sSql = "UPDATE egov_class_distributionlist_to_user SET userid = " & iKeepHeadOfHouseholdId & " WHERE userid = " & iMergeUserId
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql		' In common.asp

	' Merge the action line requests
	sSql = "UPDATE egov_actionline_requests SET userid = " & iKeepHeadOfHouseholdId & " WHERE userid = " & iMergeUserId
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql		' In common.asp

	' Merge the Facility reservations
	sSql = "UPDATE egov_facilityschedule SET lesseeid = " & iKeepHeadOfHouseholdId & " WHERE lesseeid = " & iMergeUserId
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql		' In common.asp
	
	' Merge the permit contacts
	' These have all the same info as the egov_user table does, so pull that first and then do the swap.
	sSql = "SELECT userbusinessname, userfname, userlname, useraddress, usercity, userstate, userzip, "
	sSql = sSql & " useremail, userhomephone, usercell, userfax, userpassword, userworkphone, emergencycontact, "
	sSql = sSql & " emergencyphone, neighborhoodid, residenttype, userbusinessaddress, userunit, emailnotavailable, "
	sSql = sSql & " residencyverified FROM egov_users WHERE userid = " & iKeepHeadOfHouseholdId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If IsNull(oRs("userbusinessname")) Then
			sUserbusinessname = "NULL"
		Else
			sUserbusinessname = "'" & dbsafe(oRs("userbusinessname")) & "'"
		End If 
		If IsNull(oRs("userfname")) Then
			sUserfname = "NULL"
		Else
			sUserfname = "'" & dbsafe(oRs("userfname")) & "'"
		End If 
		If IsNull(oRs("userlname")) Then
			sUserlname = "NULL"
		Else
			sUserlname = "'" & dbsafe(oRs("userlname")) & "'"
		End If 
		If IsNull(oRs("useraddress")) Then
			sUseraddress = "NULL"
		Else
			sUseraddress = "'" & dbsafe(oRs("useraddress")) & "'"
		End If 
		If IsNull(oRs("usercity")) Then
			sUserCity = "NULL"
		Else
			sUserCity = "'" & dbsafe(oRs("usercity")) & "'"
		End If
		If IsNull(oRs("userstate")) Then
			sUserstate = "NULL"
		Else
			sUserstate = "'" & dbsafe(oRs("userstate")) & "'"
		End If
		If IsNull(oRs("userzip")) Then
			sUserzip = "NULL"
		Else
			sUserzip = "'" & dbsafe(oRs("userzip")) & "'"
		End If
		If IsNull(oRs("useremail")) Then
			sUseremail = "NULL"
		Else
			sUseremail = "'" & dbsafe(oRs("useremail")) & "'"
		End If
		If IsNull(oRs("userhomephone")) Then
			sUserhomephone = "NULL"
		Else
			sUserhomephone = "'" & oRs("userhomephone") & "'"
		End If
		If IsNull(oRs("usercell")) Then
			sUsercell = "NULL"
		Else
			sUsercell = "'" & oRs("usercell") & "'"
		End If
		If IsNull(oRs("userfax")) Then
			sUserfax = "NULL"
		Else
			sUserfax = "'" & oRs("userfax") & "'"
		End If
		sUserPassword = "'" & oRs("userpassword") & "'"
		sUserWorkPhone = "'" & oRs("userworkphone") & "'"
		sEmergencyContact = "'" & oRs("emergencycontact") & "'"
		sEmergencyPhone = "'" & oRs("emergencyphone") & "'"
		If IsNull(oRs("neighborhoodid")) Then
			iNeighborhoodid = "NULL"
		else
			iNeighborhoodid = oRs("neighborhoodid")
		End If 
		If IsNull(oRs("residenttype")) Or oRs("residenttype") = "" Then
			sResidentType = "'R'"
		Else 
			sResidentType = "'" & oRs("residenttype") & "'"
		End If 
		sUserBusinessAddress = "'" & oRs("userbusinessaddress") & "'"
		sUserUnit = "'" & oRs("userunit") & "'"
		If oRs("emailnotavailable") Then 
			sEmailnotavailable = 1
		Else
			sEmailnotavailable = 0
		End If 
		If oRs("residencyverified") Then 
			sResidencyVerified = 1
		Else
			sResidencyVerified = 0
		End If 

		oRs.Close
		Set oRs = Nothing 
		
		' Pull the set of contacts that need to be updated
		' Update any permit applicants and Primary Contacts where the permit is still open. Pull them, then loop through the set
'		sSql = "SELECT P.permitid, C.permitcontactid FROM egov_permits P, egov_permitcontacts C, egov_permitstatuses S "
'		sSql = sSql & " WHERE P.permitid = C.permitid AND P.permitstatusid = S.permitstatusid AND (isapplicant = 1 OR isprimarycontact = 1) "
'		sSql = sSql & " AND S.iscompletedstatus = 0 AND S.cansavechanges = 1 AND S.changespropagate = 1 AND C.userid = " & iMergeUserId
		' Pull all of them
		sSql = "SELECT P.permitid, C.permitcontactid FROM egov_permits P, egov_permitcontacts C, egov_permitstatuses S "
		sSql = sSql & " WHERE P.permitid = C.permitid AND P.permitstatusid = S.permitstatusid AND (isapplicant = 1 OR isprimarycontact = 1) "
		sSql = sSql & " AND C.userid = " & iMergeUserId
		'response.write sSql & "<br /><br />"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		Do While Not oRs.EOF
			sSql = "UPDATE egov_permitcontacts SET firstname = " & sUserfname
			sSql = sSql & ", lastname = " & sUserlname
			sSql = sSql & ", company = " & sUserBusinessName
			sSql = sSql & ", address = " & sUserAddress
			sSql = sSql & ", city = " & sUserCity
			sSql = sSql & ", state = " & sUserState
			sSql = sSql & ", zip = " & sUserZip
			sSql = sSql & ", email = " & sUserEmail
			sSql = sSql & ", phone = " & sUserHomePhone
			sSql = sSql & ", fax = " & sUserFax
			sSql = sSql & ", cell = " & sUserCell
			sSql = sSql & ", emailnotavailable = " & sEmailnotavailable 
			sSql = sSql & ", residencyverified = " & sResidencyVerified
			sSql = sSql & ", userpassword = " & sUserPassword 
			sSql = sSql & ", residenttype = " & sResidentType
			sSql = sSql & ", userworkphone = " & sUserWorkPhone
			sSql = sSql & ", emergencyphone = " & sEmergencyPhone 
			sSql = sSql & ", neighborhoodid = " & iNeighborhoodId
			sSql = sSql & ", userunit = " & sUserUnit 
			sSql = sSql & ", emergencycontact = " & sEmergencyContact
			sSql = sSql & ", userbusinessaddress = " & sUserBusinessAddress
			sSql = sSql & ", userid = " & iKeepHeadOfHouseholdId
			sSql = sSql & " WHERE permitid = " & oRs("permitid") & " AND permitcontactid = " & oRs("permitcontactid")
			'response.write sSql & "<br /><br />"
			RunSQLStatement sSql		' In common.asp
			oRs.MoveNext
		Loop
		oRs.Close 
		Set oRs = Nothing 

	End If 

	' Merge the Journal Entries (egov_class_payment table)
	sSql = "UPDATE egov_class_payment SET userid = " & iKeepHeadOfHouseholdId & " WHERE userid = " & iMergeUserId
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql		' In common.asp

	' Merge the payments data
	sSql = "UPDATE egov_payments SET userid = " & iKeepHeadOfHouseholdId & " WHERE userid = " & iMergeUserId
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql		' In common.asp

	' Membership purchase data
	sSql = "UPDATE egov_poolpasspurchases SET userid = " & iKeepHeadOfHouseholdId & " WHERE userid = " & iMergeUserId
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql		' In common.asp

	' Merge the rental reservations for only the public ones that have the old userid. If you do a blanket replace you will overwrite the internal reservations for admin folks
	sSql = "SELECT reservationid FROM egov_rentalreservations R, egov_rentalreservationtypes T "
	sSql = sSql & "WHERE R.reservationtypeid = T.reservationtypeid AND T.reservationtypeselector = 'public' AND rentaluserid = " & iMergeUserId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		sSql = "UPDATE egov_rentalreservations SET rentaluserid = " & iKeepHeadOfHouseholdId & " WHERE reservationid = " & oRs("reservationid")
		RunSQLStatement sSql		' In common.asp
		oRs.MoveNext
	Loop

	oRs.Close 
	Set oRs = Nothing 

End If 

' loop through all of the merging household and handle the merge as they indicated on the match page
Do While x < iMaxMerge
	x = x + CLng(1)
	
	iMergeUserId = CLng(request("mergeuserid" & x))
	iKeepUserId = CLng(request("keepuserid" & x))

	If iMergeUserId <> iKeepUserId  Then
		If iKeepUserId = CLng(-1) Then
			' -1 is "add them as a new family member"

			' Move their citizen record into the keep family
			sSql = "UPDATE egov_users SET headofhousehold = 0, userpassword = NULL, familyid = " & iKeepFamilyId & " WHERE userid = " & iMergeUserId
			'response.write sSql & "<br /><br />"
			RunSQLStatement sSql		' In common.asp

			' Move their family record to the keep family
			sSql = "UPDATE egov_familymembers SET belongstouserid = " & iKeepFamilyId & " WHERE userid = " & iMergeUserId
			'response.write sSql & "<br /><br />"
			RunSQLStatement sSql		' In common.asp

			iKeepUserId = iMergeUserId
		Else
			' This is "Move their stuff to another keep family member and delete them"


			' Add the merging account balance to the keep account balance
			dMergeAccountBalance = GetAcountBalance( iMergeUserId )
			If CDbl(dMergeAccountBalance) <> CDbl(0.00) Then 
				dKeepAccountBalance = GetAcountBalance( iKeepUserId )
				dKeepAccountBalance = CDbl(dKeepAccountBalance) + CDbl(dMergeAccountBalance)
				UpdateAccountBalance iKeepUserId, dKeepAccountBalance
			End If 

			' Merge the accounts ledger transfers to and from the citizen account
			sSql = "UPDATE egov_accounts_ledger SET accountid = " & iKeepUserId
			sSql = sSql & " WHERE paymenttypeid = 4 AND orgid = " & session("orgid") & " AND accountid = " & iMergeUserId
			'response.write sSql & "<br /><br />"
			RunSQLStatement sSql		' In common.asp


			' Mark them in the family table as deleted
			sSql = "UPDATE egov_familymembers SET isdeleted = 1 WHERE userid = " & iMergeUserId
			'response.write sSql & "<br /><br />"
			RunSQLStatement sSql		' In common.asp

			' Mark them in the user table as deleted
			sSql = "UPDATE egov_users SET headofhousehold = 0, userpassword = NULL, isdeleted = 1, transferdate = GETDATE(), transferadminid = " & session("UserID")
			sSql = sSql & ", transfertouserid = " & iKeepUserId & " WHERE userid = " & iMergeUserId
			'response.write sSql & "<br /><br />"
			RunSQLStatement sSql		' In common.asp
		End If 

		iMergeFamilyMemberId = GetFamilyMemberId( iMergeUserId )
		iKeepFamilyMemberId = GetFamilyMemberId( iKeepUserId )

		' Merge their class list records
		sSql = "UPDATE egov_class_list SET userid = " & iKeepHeadOfHouseholdId & ", familymemberid = " & iKeepFamilyMemberId
		sSql = sSql & ", attendeeuserid = " & iKeepUserId & " WHERE attendeeuserid = " & iMergeUserId
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql		' In common.asp

		' Memberships
		' This should be OK, as any existing memberships will have different poolpassid's from the one we are changing
		sSql = "UPDATE egov_poolpassmembers SET familymemberid = " & iKeepFamilyMemberId & " WHERE familymemberid = " & iMergeFamilyMemberId
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql		' In common.asp
	End If 
Loop 

' Take them somewhere to display the results.
'response.write "mergecitizensresult.asp?familyid=" & iKeepFamilyId & "&mergeuserid=" & iMergeHeadOfHouseholdId & "&keepuserid=" & iKeepHeadOfHouseholdId
response.redirect "mergecitizensresult.asp?familyid=" & iKeepFamilyId & "&mergeuserid=" & iMergeHeadOfHouseholdId & "&keepuserid=" & iKeepHeadOfHouseholdId


'--------------------------------------------------------------------------------------------------
' int Function GetHeadOfHouseholdId( iFamilyId )
'--------------------------------------------------------------------------------------------------
Function GetHeadOfHouseholdId( iFamilyId )
	Dim sSql, oRs

	sSql = "SELECT userid FROM egov_users "
	sSql = sSQl & " WHERE headofhousehold = 1 AND familyid = " & iFamilyId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetHeadOfHouseholdId = oRs("userid")
	Else
		GetHeadOfHouseholdId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' int Function GetFamilyMemberId( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetFamilyMemberId( iUserId )
	Dim sSql, oRs

	sSql = "SELECT familymemberid FROM egov_familymembers WHERE userid = " & iUserId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFamilyMemberId = oRs("familymemberid")
	Else
		GetFamilyMemberId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double Function GetAcountBalance( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetAcountBalance( iUserId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(accountbalance,0.00) AS accountbalance FROM egov_users WHERE userid = " & iUserId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetAcountBalance = oRs("accountbalance")
	Else
		GetAcountBalance = 0.00
	End If 

	oRs.Close
	Set oRs = Nothing
	
End Function


'--------------------------------------------------------------------------------------------------
' void UpdateAccountBalance( iUserId, dAccountBalance )
'--------------------------------------------------------------------------------------------------
Sub UpdateAccountBalance( ByVal iUserId, ByVal dAccountBalance )
	Dim sSql

	sSql = "UPDATE egov_users SET accountbalance = " & dAccountBalance & " WHERE userid = " & iUserId
	'response.write sSql & "<br /><br />"

	RunSQLStatement sSql		' In common.asp

End Sub



%>

