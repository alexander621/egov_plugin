<%
Call subSaveInstructor(request("InstructorId"), request("sFirstName"), request("sMiddle"), request("sLastName"), request("sEmail"), request("sInstrPhone"), request("sMobilePhone"), request("sWebsiteURL"), request("sImageURL"), request("sBio"), request("isemailpublic"), request("isphonepublic"), request("iscellpublic"), request("sImageAlt"), request("userid"))


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB subSaveInstructor(InstructorId, sFirstName, sMiddle, sLastName, sEmail, sPhone, sMobilePhone, sWebsiteURL, sImageURL, sBio)
' AUTHOR: Terry Foster
' CREATED: 04/26/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subSaveInstructor( InstructorId, sFirstName, sMiddle, sLastName, sEmail, sPhone, sMobilePhone, sWebsiteURL, sImageURL, sBio, isemailpublic, isphonepublic, iscellpublic, sImageAlt, iUserId )
	Dim sSql, oCmd

	sFirstName = DBsafe( sFirstName )
	sMiddle = DBsafe( sMiddle )
	sLastName = DBsafe( sLastName )
	sEmail = DBsafe( sEmail )
	sPhone = DBsafe( Trim(Replace(sPhone,", ","")) )
	sMobilePhone = DBsafe( Trim(Replace(sMobilePhone,", ","")) )
	sWebsiteURL = DBsafe( sWebsiteURL )
	sImageURL = DBsafe( sImageURL )
	sImageAlt = dbsafe( sImageAlt )
	sBio = DBsafe( sBio )

	' Did they check to display the email
	If LCase(isemailpublic) = "on" Then
		isemailpublic = 1 
	Else
		isemailpublic = 0
	End If

	' Did they check to display the phone
	If LCase(isphonepublic) = "on" Then
		isphonepublic = 1 
	Else
		isphonepublic = 0
	End If

	' Did they check to display the cell
	If LCase(iscellpublic) = "on" Then
		iscellpublic = 1 
	Else
		iscellpublic = 0
	End If 

	If clng(iUserId) = clng(0) Then 
		iUserId = "NULL"
	End If 

	If InstructorId = "0" Then
		' Insert new records
		sSql = "INSERT INTO egov_class_instructor ( orgid, firstname, middle, lastname, email, phone, cellphone, websiteurl, imgurl, bio, isemailpublic, isphonepublic, iscellpublic, imgalt, userid ) Values ( " 
		sSql = sSql & Session("OrgID") & ",'" & sFirstName & "','" & sMiddle & "','" & sLastName & "','" & sEmail & "','" & sPhone & "','" 
		sSql = sSql & sMobilePhone & "','" & sWebsiteURL & "','" & sImageURL & "','" & sBio & "', " & isemailpublic & ", " & isphonepublic & ", " & iscellpublic & ", '" & sImageAlt & "', " & iUserId & " )"
	Else 
		' Update existing records
		sSQL = "UPDATE egov_class_instructor SET firstname = '" & sFirstName & "', middle= '" & sMiddle & "', lastname= '" & sLastName
		sSql = sSql & "', email= '" & sEmail & "', phone= '" & sPhone & "', cellphone= '" & sMobilePhone & "', websiteurl= '" 
		sSql = sSql & sWebsiteURL & "', imgurl= '" & sImageURL & "', bio= '" & sBio & "', isemailpublic = " & isemailpublic
		sSql = sSql & ", isphonepublic = " & isphonepublic & ", iscellpublic = " & iscellpublic & ", imgalt = '" & sImageAlt & "', userid = " & iUserId
		sSql = sSql & " WHERE Instructorid = " & InstructorId 
	End If

	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing


	' REDIRECT TO instructor management page
	 response.redirect "instructor_mgmt.asp"

End Sub


%>

<!-- #include file="../includes/common.asp" //-->
