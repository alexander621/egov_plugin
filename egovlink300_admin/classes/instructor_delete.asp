
<%
'--------------------------------------------------------------------------------------------------
' SUB subDeleteInstructor(InstructorId)
' AUTHOR: TERRY FOSTER
' CREATED: 04/26/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
	Dim sSql, oCmd
	
	sSQL = "DELETE FROM egov_class_instructor WHERE Instructorid = " &  request("InstructorId") 
'	response.write sSQL
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSQL
		.Execute
		.CommandText = "Delete from egov_class_to_instructor where instructorid = " & request("InstructorId")
		.Execute
	End With
	Set oCmd = Nothing

	' REDIRECT TO instructor management page
	response.redirect "instructor_mgmt.asp"


%>
