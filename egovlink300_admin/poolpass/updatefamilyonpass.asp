<!-- #include file="poolpass_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: updatefamilyonpass.asp
' AUTHOR: Steve Loar
' CREATED: 10/31/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   10/31/07	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPoolpassid, sMemberList

iPoolpassid = CLng(request("poolpassid"))
sMemberList = ""

For each item in request("familymemberid")
	If Not IsOnPoolPass( iPoolpassid, item ) Then
		AddToPass iPoolpassid, item
	End If 
	If sMemberList = "" Then
		sMemberList = "("
	Else 
		sMemberList = sMemberList & ","
	End If 
	sMemberList = sMemberList & item 
Next
sMemberList = sMemberList & ")"

RemoveMembers iPoolpassid, sMemberList

response.redirect "poolpass_details.asp?iPoolPassId=" & iPoolpassid


'------------------------------------------------------------------------------------------------------------
' Function IsOnPoolPass( iPoolPassId, iFamilymemberid )
'------------------------------------------------------------------------------------------------------------
Function IsOnPoolPass( iPoolPassId, iFamilymemberid )
	Dim sSQL, oRs

	sSQL = "Select count(familymemberid) AS hits FROM egov_poolpassmembers WHERE poolpassid = " & iPoolPassId & " AND familymemberid = " & iFamilymemberid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If clng(oRs("hits")) > clng(0) Then 
			IsOnPoolPass = True 
		Else
			IsOnPoolPass = False 
		End If 
	Else
		IsOnPoolPass = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------------------------------------
' Sub AddToPass( iPoolpassid, iFamilymemberid )
'------------------------------------------------------------------------------------------------------------
Sub AddToPass( iPoolpassid, iFamilymemberid )
	Dim oCmd

 lcl_member_id = getNextMemberID()

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "INSERT INTO egov_poolpassmembers ( poolpassid, familymemberid, memberid ) VALUES (" & iPoolpassid & ", " & iFamilymemberid & ", " & lcl_member_id & ")"
		.Execute
	End with

	Set oCmd = Nothing 
End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub RemoveMembers( iPoolpassid, sMemberList )
'------------------------------------------------------------------------------------------------------------
Sub RemoveMembers( iPoolpassid, sMemberList )
	Dim oCmd, sSql

	sSql = "DELETE FROM egov_poolpassmembers WHERE poolpassid = " & iPoolpassid & " AND familymemberid NOT IN " & sMemberList
	'response.write sSql & "<br />"

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End with

	Set oCmd = Nothing 
End Sub 


%>
