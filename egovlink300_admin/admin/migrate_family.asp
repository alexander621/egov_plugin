<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->

<%
 response.write "Family migration did not run.  PLEASE DISABLE RESPONSE.END TO RUN SCRIPT."
 response.end

'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: migrate_family.asp
' AUTHOR: Steve Loar
' CREATED: 01/19/07
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module migrates old users to the new family structure
'
' MODIFICATION HISTORY
' 1.0   01/19/07	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' USER VALUES
Dim sFirstName,sLastName,sHomeAddress,sCity,sState,sZip,sPhone,sEmail,sFax,sCell,sBusinessName,sHomenumber,sPassword,iUserID
Dim bHasResidentStreets, bFound, sResidenttype, sBusinessAddress, bHasBusinessStreets, sWorkPhone, iNeighborhoodId, oUsers
Dim sEmergencyContact, sEmergencyPhone, sBirthdate, iFamilyId, iRelationshipId, oFamily, sResidencyVerified, sRelationship
Dim sSql, x, y, iResidencyVerified, oSelves, z

sLevel = "../" ' Override of value from common.asp


%>

<html>
<head>
	<title>E-Gov Users Setup Script</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

<script language="Javascript">
<!--
	// Put any JavaScript here

//-->
</script>

</head>

<body>
 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	<p>Started: <%=Now()%></p>

	<p><h4>Registered Citizen Update</h4></p>
<%

	x = 0
	y = 0
	z = 0

	sSql = "select U.orgid, U.userid, R.relationshipid, U.userlname, U.userfname from egov_users U, egov_familymember_relationships R "
	sSql = sSql & " where U.userregistered = 1 and U.familyid is Null "
	sSql = sSql & " and U.orgid = R.orgid and R.isdefault = 1 and U.useremail is not Null "
	sSql = sSql & " order by U.orgid, U.userlname, U.userfname, U.userid "

	Set oUsers = Server.CreateObject("ADODB.Recordset")
	oUsers.Open sSQL, Application("DSN"), 0, 1

	Do While Not oUsers.EOF 
		x = x + 1
		response.write vbcrlf & "<br />" & x & ". " & oUsers("orgid") & ": " & oUsers("userfname") & " " & oUsers("userlname")
		UpdateUserRow oUsers("userid"), oUsers("relationshipid")
		If x Mod 50 = 0 Then
			response.flush 
		End If 
		oUsers.MoveNext
	Loop 
	response.flush 
	oUsers.close
	Set oUsers = Nothing 

	response.write vbcrlf & "<br /><br />" & x & " Registered citizens processed.<br /><br />"

	' Update the family members for the 'Yourself' entries
	sSql = "Select familymemberid, belongstouserid from egov_familymembers where userid is NULL and relationship = 'Yourself'"
	
	Set oSelves = Server.CreateObject("ADODB.Recordset")
	oSelves.Open sSQL, Application("DSN"), 0, 1

	Do While Not oSelves.EOF 
		FamilyMemberUpdate oSelves("familymemberid"), oSelves("belongstouserid")
		z = z + 1
		oSelves.MoveNext
	Loop 
	response.flush 
	oSelves.close
	Set oSelves = Nothing 

	response.write vbcrlf & "<br /><br />" & z & " 'Yourself' Family members processed.<br /><br />"


	sSql = "select F.familymemberid, F.firstname, F.lastname, F.birthdate, R.relationshipid, "
	sSql = sSql & " U.orgid, U.userid as familyid, U.useraddress, U.usercity, U.userstate, U.userzip, U.userhomephone, "
	sSql = sSql & " U.userregistered, U.residenttype, U.residencyverified "
	sSql = sSql & " from egov_users U, egov_familymembers F, egov_familymember_relationships R "
	sSql = sSql & " where U.userid = F.belongstouserid "
	sSql = sSql & " and F.relationship = R.relationship and R.orgid = U.orgid and F.userid is null "
	sSql = sSql & " order by U.orgid, U.userlname, U.userfname, U.userid, F.familymemberid "

	Set oFamily = Server.CreateObject("ADODB.Recordset")
	oFamily.Open sSQL, Application("DSN"), 0, 1

	Do While Not oFamily.EOF 
		y = y + 1
		response.write vbcrlf & y & ". " & oFamily("orgid") & ": " & oFamily("firstname") & " " & oFamily("lastname")
		'GetUnRegisteredUserValues( oFamily("familyid") )
		iOrgId = oFamily("orgid")
		iFamilyid = oFamily("familyid")
		sFirstName = oFamily("firstname")
		sLastName = oFamily("lastname")
		sBirthdate = oFamily("birthdate")
		iRelationshipId = oFamily("relationshipid")
		sHomeAddress = oFamily("useraddress")
		sState = oFamily("userstate")
		sCity = oFamily("usercity")
		sZip = oFamily("userzip")
		sBusinessaddress = ""
		sEmail = ""
		sFax = ""
		sCell = ""
		sBusinessName = ""
		sPassword = ""
		sHomenumber = oFamily("userhomephone")
		sWorkPhone = ""
		sResidenttype = oFamily("residenttype")
		sUserRegistered = 1
		sNeighborhoodid = 0
		sEmergencycontact = ""
		sEmergencyphone = ""
		iResidencyVerified = oFamily("residencyverified")

		iUserID = InsertFamilymember( iOrgId, sFirstname, sLastname, sHomeAddress, sCity, sState, sZip, sHomenumber, _
			sCell,sFax, sWorkPhone, sBusinessname, sBusinessaddress, sEmergencycontact, sEmergencyphone, sNeighborhoodid, _
			sBirthdate, iRelationshipid, iResidencyVerified, sResidenttype, iFamilyid )

		response.write " --> " & iUserId  & "<br />"

		FamilyMemberUpdate oFamily("familymemberid"), iUserID

		If y Mod 50 = 0 Then
			response.flush 
		End If 
		oFamily.MoveNext
	Loop
	response.flush
	oFamily.close
	Set oFamily = Nothing 

	response.write vbcrlf & "<br /><br />" & y & " Family members processed.<br /><br />"
%>

	<p><hr /></p>
	<p>Finished: <%=Now()%></p>
	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub UpdateUserRow( iUserId, iRelationshipId )
'--------------------------------------------------------------------------------------------------
Sub UpdateUserRow( iUserId, iRelationshipId )
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "Update egov_users Set familyid = " & iUserId & ", relationshipid = " & iRelationshipId & " Where userid = " & iUserId
		.Execute
	End With
	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function InsertFamilymember()
'--------------------------------------------------------------------------------------------------
Function InsertFamilymember( iOrgId, sFirstname, sLastname, sHomeAddress, sCity, sState, sZip, sHomenumber, _
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
			.Parameters.Append oCmd.CreateParameter("@businessname", 200, 1, 25, sBusinessname)
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
			.Parameters.Append oCmd.CreateParameter("@city", 200, 1, 20, sCity)
		Else
			.Parameters.Append oCmd.CreateParameter("@city", 200, 1, 20, NULL)
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
' Sub GetUnRegisteredUserValues(iUserId)
'--------------------------------------------------------------------------------------------------
Sub GetUnRegisteredUserValues(iUserId)
	Dim sSql, oValues

	sSQL = "SELECT * FROM egov_users WHERE userid = " & iUserId

	Set oValues = Server.CreateObject("ADODB.Recordset")
	oValues.Open sSQL, Application("DSN"), 3, 1

	If NOT oValues.EOF Then
		'sFirstName = ""
		'sLastName = oValues("userlname")
		sAddress = oValues("useraddress")
		sState = oValues("userstate")
		sCity = oValues("usercity")
		sZip = oValues("userzip")
		sEmail = ""
		sFax = oValues("userfax")
		sCell = oValues("usercell")
		sBusinessName = ""
		sPassword = ""
		sDayPhone = oValues("userhomephone")
		sWorkPhone = oValues("userworkphone")
		If IsNull(oValues("residenttype")) Or oValues("residenttype") = "" Then
			sResidenttype = "N"
		Else 
			sResidenttype = oValues("residenttype")
		End If 
		sBusinessAddress = ""
		If IsNull(oValues("neighborhoodid")) Then 
			iNeighborhoodId = 0
		Else 
			iNeighborhoodId = oValues("neighborhoodid")
		End If 
		sEmergencyContact = oValues("emergencycontact")
		sEmergencyPhone = oValues("emergencyphone")
		'iRelationshipId = 0
		'sBirthdate = ""
		sResidencyVerified = oValues("residencyverified")
	End If

	oValues.close
	Set oValues = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' SUB FamilyMemberUpdate( ByVal iFamilyMemberId, ByVal iUserId )
'--------------------------------------------------------------------------------------------------
Sub FamilyMemberUpdate( ByVal iFamilyMemberId, ByVal iUserId )
	Dim sSql, oCmd

	sSql = "Update egov_familymembers Set userid = " & iUserId
	sSql = sSql & " Where familymemberid = " & iFamilyMemberId & ""
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

End Sub
%>
