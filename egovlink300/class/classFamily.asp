<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: classFamily.asp
' AUTHOR: Steve Loar
' CREATED: 12/29/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the Family class
'
' MODIFICATION HISTORY
' 1.0   12/29/2006	Steve Loar - Initial code 
' 1.2	07/25/2008	Steve Loar - Changed to use deleted flag on family member delete
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Class classFamily
	
	Private Sub Class_Initialize()
	End Sub 

	'--------------------------------------------------------------------------------------------------
	' Public Function InsertFamilymember()
	'--------------------------------------------------------------------------------------------------
	Public Function InsertFamilymember( iOrgId, sFirstname, sLastname, sHomeAddress, sCity, sState, sZip, sHomenumber, _
		 sCellnumber,sFaxnumber, sWorknumber, sBusinessname, sBusinessaddress, sEmergencycontact, sEmergencyphone, sNeighborhoodid, _
		 sBirthdate, sRelationshipid, iResidencyVerified, sResidenttype, iFamilyid )
		Dim oCmd, iUserId

		' Parameters for the stored Proc
		'@orgid int,
		'@firstname varchar(25),
		'@lastname varchar(25),
		'@businessname  varchar(50) = NULL,
		'@address1  varchar(250) = NULL,
		'@homenumber varchar(20),
		'@cellnumber varchar(20),
		'@worknumber varchar(20) = NULL,
		'@city varchar(20) = NULL,
		'@state varchar(20) = NULL,
		'@zip varchar(20) = NULL,
		'@faxnumber varchar(20) = NULL ,
		'@businessaddress varchar(255) = NULL,
		'@emergencycontact varchar(100) = NULL,
		'@emergencyphone varchar(50) = NULL,
		'@neighborhoodid int = NULL,
		'@birthdate datetime = NULL,
		'@relationshipid int = NULL, 
		'@residencyverified bit,
		'@residenttype char(1) = NULL,
		'@familyid int,
		'@userid int OUTPUT

		Set oCmd = Server.CreateObject("ADODB.Command")
		With oCmd
			.ActiveConnection = Application("DSN")
			.CommandText = "NewCitizenFamilyMember"
			.CommandType = 4
			.Parameters.Append oCmd.CreateParameter("@orgid", 3, 1, 4, iOrgId)
			.Parameters.Append oCmd.CreateParameter("@firstname", 200, 1, 25, sFirstname)
			.Parameters.Append oCmd.CreateParameter("@lastname", 200, 1, 25, sLastname)
			If sBusinessname <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@businessname", 200, 1, 25, Left(sBusinessname,25))
			Else
				.Parameters.Append oCmd.CreateParameter("@businessname", 200, 1, 25, NULL)
			End If 
			If sHomeAddress <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@address1", 200, 1, 250, sHomeAddress)
			Else
				.Parameters.Append oCmd.CreateParameter("@address1", 200, 1, 250, NULL)
			End If 
			.Parameters.Append oCmd.CreateParameter("@homenumber", 200, 1, 20, sHomenumber)
			If sCellnumber <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@cellnumber", 200, 1, 20, sCellnumber)
			Else
				.Parameters.Append oCmd.CreateParameter("@cellnumber", 200, 1, 20, NULL)
			End If 
			If sWorknumber <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@worknumber", 200, 1, 20, sWorknumber)
			Else
				.Parameters.Append oCmd.CreateParameter("@worknumber", 200, 1, 20, NULL)
			End If 
			If sCity <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@city", 200, 1, 40, sCity)
			Else
				.Parameters.Append oCmd.CreateParameter("@city", 200, 1, 40, NULL)
			End If 
			If sState <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@state", 200, 1, 20, sState)
			Else
				.Parameters.Append oCmd.CreateParameter("@state", 200, 1, 20, NULL)
			End If
			If sZip <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@zip", 200, 1, 20, sZip)
			Else
				.Parameters.Append oCmd.CreateParameter("@zip", 200, 1, 20, NULL)
			End If
			If sFaxnumber <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@faxnumber", 200, 1, 20, sFaxnumber)
			Else
				.Parameters.Append oCmd.CreateParameter("@faxnumber", 200, 1, 20, NULL)
			End If
			If sBusinessaddress <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@businessaddress", 200, 1, 255, sBusinessaddress)
			Else
				.Parameters.Append oCmd.CreateParameter("@businessaddress", 200, 1, 255, NULL)
			End If
			If sEmergencycontact <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@emergencycontact", 200, 1, 100, sEmergencycontact)
			Else
				.Parameters.Append oCmd.CreateParameter("@emergencycontact", 200, 1, 100, NULL)
			End If
			If sEmergencyphone <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@emergencyphone", 200, 1, 50, sEmergencyphone)
			Else
				.Parameters.Append oCmd.CreateParameter("@emergencyphone", 200, 1, 50, NULL)
			End If
			If clng(sNeighborhoodid) <> clng(0) Then 
				.Parameters.Append oCmd.CreateParameter("@neighborhoodid", 3, 1, 4, sNeighborhoodid)
			Else
				.Parameters.Append oCmd.CreateParameter("@neighborhoodid", 3, 1, 4, NULL)
			End If
			If Trim(sBirthdate) <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@birthdate", 135, 1, 16, sBirthdate)
			Else
				.Parameters.Append oCmd.CreateParameter("@birthdate", 135, 1, 16, NULL)
			End If
			.Parameters.Append oCmd.CreateParameter("@relationshipid", 3, 1, 4, sRelationshipid)
			.Parameters.Append oCmd.CreateParameter("@residencyverified", 11, 1, 1, iResidencyVerified)
			If Trim(sResidenttype) <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@residenttype", 129, 1, 1, sResidenttype)
			Else
				.Parameters.Append oCmd.CreateParameter("@residenttype", 129, 1, 1, NULL)
			End If
			.Parameters.Append oCmd.CreateParameter("@familyid", 3, 1, 4, iFamilyid)
			.Parameters.Append oCmd.CreateParameter("@userid", 3, 2, 4)
			.Execute
		End With

		iUserId = oCmd.Parameters("@userid").Value

		Set oCmd = Nothing

		' Send back the new userid
		InsertFamilymember = iUserId

	End Function 


	'--------------------------------------------------------------------------------------------------
	' Public Sub UpdateFamilymember()
	'--------------------------------------------------------------------------------------------------
	Public Sub UpdateFamilymember( iUserId, sFirstname, sLastname, sHomeAddress, sCity, sState, sZip, sHomenumber, _
		 sCellnumber,sFaxnumber, sWorknumber, sBusinessname, sBusinessaddress, sEmergencycontact, sEmergencyphone, sNeighborhoodid, _
		 sBirthdate, sRelationshipid )
		Dim oCmd

		' Parameters for the stored Proc
		'@userid int,
		'@firstname varchar(25),
		'@lastname varchar(25),
		'@businessname  varchar(50) = NULL,
		'@address1  varchar(250) = NULL,
		'@homenumber varchar(20),
		'@cellnumber varchar(20) = NULL, 
		'@worknumber varchar(20) = NULL,
		'@city varchar(20) = NULL,
		'@state varchar(20) = NULL,
		'@zip varchar(20) = NULL,
		'@faxnumber varchar(20) = NULL ,
		'@businessaddress varchar(255) = NULL,
		'@emergencycontact varchar(100) = NULL,
		'@emergencyphone varchar(50) = NULL,
		'@neighborhoodid int = NULL,
		'@birthdate datetime = NULL,
		'@relationshipid int = NULL,

		Set oCmd = Server.CreateObject("ADODB.Command")
		With oCmd
			.ActiveConnection = Application("DSN")
			.CommandText = "UpdateCitizenFamilyMember"
			.CommandType = 4
			.Parameters.Append oCmd.CreateParameter("@userid", 3, 1, 4, iUserId)
			.Parameters.Append oCmd.CreateParameter("@firstname", 200, 1, 25, sFirstname)
			.Parameters.Append oCmd.CreateParameter("@lastname", 200, 1, 25, sLastname)
			If sBusinessname <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@businessname", 200, 1, 50, sBusinessname)
			Else
				.Parameters.Append oCmd.CreateParameter("@businessname", 200, 1, 50, NULL)
			End If 
			If sHomeAddress <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@address1", 200, 1, 250, sHomeAddress)
			Else
				.Parameters.Append oCmd.CreateParameter("@address1", 200, 1, 250, NULL)
			End If 
			.Parameters.Append oCmd.CreateParameter("@homenumber", 200, 1, 20, sHomenumber)
			If sCellnumber <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@cellnumber", 200, 1, 20, sCellnumber)
			Else
				.Parameters.Append oCmd.CreateParameter("@cellnumber", 200, 1, 20, NULL)
			End If 
			If sWorknumber <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@worknumber", 200, 1, 20, sWorknumber)
			Else
				.Parameters.Append oCmd.CreateParameter("@worknumber", 200, 1, 20, NULL)
			End If 
			If sCity <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@city", 200, 1, 40, sCity)
			Else
				.Parameters.Append oCmd.CreateParameter("@city", 200, 1, 40, NULL)
			End If 
			If sState <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@state", 200, 1, 20, sState)
			Else
				.Parameters.Append oCmd.CreateParameter("@state", 200, 1, 20, NULL)
			End If
			If sZip <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@zip", 200, 1, 20, sZip)
			Else
				.Parameters.Append oCmd.CreateParameter("@zip", 200, 1, 20, NULL)
			End If
			If sFaxnumber <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@faxnumber", 200, 1, 20, sFaxnumber)
			Else
				.Parameters.Append oCmd.CreateParameter("@faxnumber", 200, 1, 20, NULL)
			End If
			If sBusinessaddress <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@businessaddress", 200, 1, 255, sBusinessaddress)
			Else
				.Parameters.Append oCmd.CreateParameter("@businessaddress", 200, 1, 255, NULL)
			End If
			If sEmergencycontact <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@emergencycontact", 200, 1, 100, sEmergencycontact)
			Else
				.Parameters.Append oCmd.CreateParameter("@emergencycontact", 200, 1, 100, NULL)
			End If
			If sEmergencyphone <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@emergencyphone", 200, 1, 50, sEmergencyphone)
			Else
				.Parameters.Append oCmd.CreateParameter("@emergencyphone", 200, 1, 50, NULL)
			End If
			If clng(sNeighborhoodid) <> clng(0) Then 
				.Parameters.Append oCmd.CreateParameter("@neighborhoodid", 3, 1, 4, sNeighborhoodid)
			Else
				.Parameters.Append oCmd.CreateParameter("@neighborhoodid", 3, 1, 4, NULL)
			End If
			If Trim(sBirthdate) <> "" Then 
				.Parameters.Append oCmd.CreateParameter("@birthdate", 135, 1, 16, sBirthdate)
			Else
				.Parameters.Append oCmd.CreateParameter("@birthdate", 135, 1, 16, NULL)
			End If
			.Parameters.Append oCmd.CreateParameter("@relationshipid", 3, 1, 4, sRelationshipid)

			.Execute
		End With

		Set oCmd = Nothing

	End Sub 


	'--------------------------------------------------------------------------------------------------
	' DeleteFamilyMember iUserId, iDeletedById 
	'--------------------------------------------------------------------------------------------------
	Public Sub DeleteFamilyMember( ByVal iUserId, ByVal iDeletedById )
		Dim sSql

		sSql = "UPDATE egov_users SET isdeleted = 1, deleteddate = GETDATE(), deletedbycitizenid = "
		sSql = sSql & iDeletedById & " WHERE userid = " & iUserId
		RunSQLStatement sSql

		sSql = "UPDATE egov_familymembers SET isdeleted = 1 WHERE userid = " & iUserId
		RunSQLStatement sSql

	End Sub 


	'--------------------------------------------------------------------------------------------------
	' integer GetFamilyId( iUserId )
	'--------------------------------------------------------------------------------------------------
	Public Function GetFamilyId( ByVal iUserId )
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


	'--------------------------------------------------------------------------------------------------
	' Public Sub AddFamilyMember( iBelongsToUserId, sFirstName, sLastName, sRelationship, sBirthDate )
	'--------------------------------------------------------------------------------------------------
	Public Sub AddFamilyMember( ByVal iBelongsToUserId, ByVal sFirstName, ByVal sLastName, ByVal sRelationship, ByVal sBirthDate )
		' This function adds family members to the egov_familymembers table
		Dim sSql, oCmd
		
		sSql = "Insert Into egov_familymembers (firstname, lastname, birthdate, belongstouserid, relationship, userid) values ('"
		If sBirthDate <> "NULL" Then
			sSql = sSql & sFirstName & "', '" & sLastName & "', '" & sBirthDate & "', " & iBelongsToUserId & ", '" & sRelationship & "', " & iBelongsToUserId & " )"
		Else
			sSql = sSql & sFirstName & "', '" & sLastName & "', " & sBirthDate & ", " & iBelongsToUserId & ", '" & sRelationship & "', " & iBelongsToUserId & " )"
		End If 

		RunSQLStatement sSql

	End Sub 


	'--------------------------------------------------------------------------------------------------
	' Public Sub UpdateFamilyId( iUserId, iFamilyId, iRelationshipId )
	'--------------------------------------------------------------------------------------------------
	Public Sub UpdateFamilyId( ByVal iUserId, ByVal iFamilyId, ByVal iRelationshipId, ByVal iNeighborhoodid )
		Dim sSql

		sSql = "UPDATE egov_users SET familyid = " & iFamilyId
		
		If iRelationshipId <> "" Then 
			sSql = sSql & ", relationshipid = " & iRelationshipId
		End If 

		If iNeighborhoodid <> "" Then 
			If CLng(iNeighborhoodid) <> CLng(0) Then 
				sSql = sSql & ", neighborhoodid = " & iNeighborhoodid
			End If 
		End If 

		sSql = sSql & " WHERE userid = " & iUserId 

		session("sSql") = sSql

		'response.write sSql
		'response.End 

		RunSQLStatement sSql
		session("sSql") = ""

	End Sub 


	'--------------------------------------------------------------------------------------------------
	' Private Function DBsafe( strDB )
	'--------------------------------------------------------------------------------------------------
	Private Function DBsafe( ByVal strDB )
		Dim sNewString

		If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function

		sNewString = Replace( strDB, "'", "''" )
		sNewString = Replace( sNewString, "<", "&lt;" )
		DBsafe = sNewString

	End Function

	'-------------------------------------------------------------------------------------------------
	' void RunSQLStatement sSql 
	'-------------------------------------------------------------------------------------------------
	Private Sub RunSQLStatement( ByVal sSql )
		Dim oCmd

	'	response.write "<p>" & sSql & "</p><br /><br />"
	'	response.flush

		Set oCmd = Server.CreateObject("ADODB.Command")
		oCmd.ActiveConnection = Application("DSN")
		oCmd.CommandText = sSql
		oCmd.Execute
		Set oCmd = Nothing

	End Sub 

End Class 

%>
