<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_addcategory.asp
' AUTHOR: Steve Loar
' CREATED: 04/24/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module adds categories to a class.  It is called from list_picker.asp
'
' MODIFICATION HISTORY
' 1.0   4/24/2006   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim oCmd, iCategoryId

	Set oCmd = Server.CreateObject("ADODB.Command")

	For Each iCategoryId In request("categorylist")
		With oCmd
			.ActiveConnection = Application("DSN")
			.CommandText = "Insert Into egov_class_category_to_class ( classid, categoryid ) values ( " & request("classid") & ", " & iCategoryId & " )"
			.Execute
		End With
	Next 

	Set oCmd = Nothing

	response.redirect "list_picker.asp?classid=" & request("classid") & "&listtype=" & request("listtype") & "&postcount=" & request("postcount")
%>
