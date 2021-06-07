<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_removeinstructor.asp
' AUTHOR: Steve Loar
' CREATED: 04/24/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module removes instructors from a class.  It is called from list_picker.asp
'
' MODIFICATION HISTORY
' 1.0   4/24/2006   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim oCmd, iInstructorId

	Set oCmd = Server.CreateObject("ADODB.Command")

	For Each iInstructorId In request("instructorlist")
		With oCmd
			.ActiveConnection = Application("DSN")
			.CommandText = "delete from egov_class_to_instructor where classid = " & request("classid") & " and instructorid = " & iInstructorId 
			.Execute
		End With
	Next 

	Set oCmd = Nothing

	response.redirect "list_picker.asp?classid=" & request("classid") & "&listtype=" & request("listtype") & "&postcount=" & request("postcount")

%>
