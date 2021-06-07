<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_addwaiver.asp
' AUTHOR: Steve Loar
' CREATED: 04/24/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module adds waivers to a class.  It is called from list_picker.asp
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
			.CommandText = "Insert Into egov_class_to_waivers ( classid, waiverid ) values ( " & request("classid") & ", " & iWaiverId & " )"
			.Execute
		End With
	Next 

	Set oCmd = Nothing

	response.redirect "list_picker.asp?classid=" & request("classid") & "&listtype=" & request("listtype") & "&postcount=" & request("postcount")
%>
