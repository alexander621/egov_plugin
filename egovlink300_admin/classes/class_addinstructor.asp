<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_addinstructor.asp
' AUTHOR: Steve Loar
' CREATED: 04/24/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module add instructors to a class.  It is called from list_picker.asp
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
			.CommandText = "Insert Into egov_class_to_instructor ( classid, instructorid ) values ( " & request("classid") & ", " & iInstructorId & " )"
			.Execute
		End With
	Next 

	Set oCmd = Nothing

	response.redirect "list_picker.asp?classid=" & request("classid") & "&listtype=" & request("listtype") & "&postcount=" & request("postcount")
%>
