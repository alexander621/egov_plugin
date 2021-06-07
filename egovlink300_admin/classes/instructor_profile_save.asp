<%
	Dim InstructorId, sFirstName, sMiddle, sLastName, sEmail, sPhone, sMobilePhone, sWebsiteURL, sImageURL, sBio 
	Dim isemailpublic, isphonepublic, iscellpublic, sImageAlt, iUserId 
	Dim sSql, oCmd

	If Not UserHasPermission( Session("UserId"), "instructor profile" ) Then
		response.redirect sLevel & "permissiondenied.asp"
	End If

	iUserId = request("userid")
	InstructorId = request("InstructorId")
	sFirstName = DBsafe( request("sFirstName") )
	sMiddle = DBsafe( request("sMiddle") )
	sLastName = DBsafe( request("sLastName") )
	sEmail = DBsafe( request("sEmail") )
	sPhone = DBsafe( Trim(Replace(request("sInstrPhone"),", ","")) )
	sMobilePhone = DBsafe( Trim(Replace(request("sMobilePhone"),", ","")) )
	sWebsiteURL = DBsafe( request("sWebsiteURL") )
	sBio = DBsafe( request("sBio") )

	' Did they check to display the email
	If LCase(request("isemailpublic")) = "on" Then
		isemailpublic = 1 
	Else
		isemailpublic = 0
	End If

	' Did they check to display the phone
	If LCase(request("isphonepublic")) = "on" Then
		isphonepublic = 1 
	Else
		isphonepublic = 0
	End If

	' Did they check to display the cell
	If LCase(request("iscellpublic")) = "on" Then
		iscellpublic = 1 
	Else
		iscellpublic = 0
	End If 

	If clng(iUserId) = clng(0) Then 
		iUserId = "NULL"
	End If 

	If InstructorId = "0" Then
		' Insert new records
		sSql = "INSERT INTO egov_class_instructor ( orgid, firstname, middle, lastname, email, phone, cellphone, websiteurl, bio, isemailpublic, isphonepublic, iscellpublic, userid ) Values ( " 
		sSql = sSql & Session("OrgID") & ",'" & sFirstName & "','" & sMiddle & "','" & sLastName & "','" & sEmail & "','" & sPhone & "','" 
		sSql = sSql & sMobilePhone & "','" & sWebsiteURL & "','" & sBio & "', " & isemailpublic & ", " & isphonepublic & ", " & iscellpublic & ", " & iUserId & " )"
	Else 
		' Update existing records
		sSQL = "UPDATE egov_class_instructor SET firstname = '" & sFirstName & "', middle= '" & sMiddle & "', lastname= '" & sLastName
		sSql = sSql & "', email= '" & sEmail & "', phone= '" & sPhone & "', cellphone= '" & sMobilePhone & "', websiteurl= '" 
		sSql = sSql & sWebsiteURL & "', bio= '" & sBio & "', isemailpublic = " & isemailpublic
		sSql = sSql & ", isphonepublic = " & isphonepublic & ", iscellpublic = " & iscellpublic 
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
	 response.redirect "instructor_profile.asp"

%>

<!-- #include file="../includes/common.asp" //-->
