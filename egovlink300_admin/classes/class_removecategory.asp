<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_removecategory.asp
' AUTHOR: Steve Loar
' CREATED: 04/24/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module removes categories from a class.  It is called from list_picker.asp
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
			.CommandText = "delete from egov_class_category_to_class where classid = " & request("classid") & " and categoryid = " & iCategoryId 
			.Execute
		End With
	Next 

	Set oCmd = Nothing

	response.redirect "list_picker.asp?classid=" & request("classid") & "&listtype=" & request("listtype") & "&postcount=" & request("postcount")

%>
