<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: setrosterdefault.asp
' AUTHOR: Steve Loar	
' CREATED: 09/13/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This sets the default season for the roster list page
'
' MODIFICATION HISTORY
' 1.0   09/13/2007   Steve Loar - Initial code  
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim oCmd, iClassSeasonId

iClassSeasonId = CLng(request("classseasonid"))

Set oCmd = Server.CreateObject("ADODB.Command")
oCmd.ActiveConnection = Application("DSN")

' reset them all to not being the default
oCmd.CommandText = "Update egov_class_seasons set isrosterdefault = 0 where orgid = " & Session("orgid")
oCmd.Execute

' set the selected one to be the default
oCmd.CommandText = "Update egov_class_seasons set isrosterdefault = 1 where classseasonid = " & iClassSeasonId
oCmd.Execute

Set oCmd = Nothing

response.write "Completed"

%>