<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_removewaiver.asp
' AUTHOR: Steve Loar
' CREATED: 04/24/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module removes waivers from a class.  It is called from list_picker.asp
'
' MODIFICATION HISTORY
' 1.0   4/24/2006   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim oCmd, iWaiverId

	Set oCmd = Server.CreateObject("ADODB.Command")

	For Each iWaiverId In request("waiverlist")
		With oCmd
			.ActiveConnection = Application("DSN")
			.CommandText = "delete from egov_class_to_waivers where classid = " & request("classid") & " and waiverid = " & iWaiverId 
			.Execute
		End With
	Next 

	Set oCmd = Nothing

	response.redirect "list_picker.asp?classid=" & request("classid") & "&listtype=" & request("listtype") & "&postcount=" & request("postcount")

%>
