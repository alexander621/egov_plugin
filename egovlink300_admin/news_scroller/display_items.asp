<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: display_items.asp
' AUTHOR: Steve Loar
' CREATED: 10/31/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module toggles the display of the News Items.
'
' MODIFICATION HISTORY
' 1.0   10/31/06	Steve Loar - Initial version.
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim iNewDisplay, oCmd

	If clng(request("itemdisplay")) = clng(1) Then
		iNewDisplay = 0
	Else
		iNewDisplay = 1
	End If 
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "UPDATE egov_news_items SET itemdisplay = " & iNewDisplay & " WHERE newsitemid = " & request("newsitemid") 
		.Execute
	End With
	Set oCmd = Nothing

	' REDIRECT TO News Items page
	response.redirect("list_items.asp")

%>
