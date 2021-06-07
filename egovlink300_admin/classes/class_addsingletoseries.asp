<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_addsingletoseries.asp
' AUTHOR: Steve Loar
' CREATED: 05/02/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This script moves single classes into a series
'
' MODIFICATION HISTORY
' 1.0   5/02/2006   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		' Update the class table
		.CommandText = "MoveSingleToSeries"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@ClassId", 3, 1, 4, request("classid"))
		.Parameters.Append oCmd.CreateParameter("@ParentClassId", 3, 1, 4, request("parentclassid"))
'		.CommandText = "Update egov_class set classtypeid = 1, parentclassid = " & request("parentclassid") & " Where classid = " & request("classid") & " "
		.Execute
	End With
	Set oCmd = Nothing
	
	' Return to the edit page
	response.redirect "edit_class.asp?classid=" & request("parentclassid") ' & "#children"

%>