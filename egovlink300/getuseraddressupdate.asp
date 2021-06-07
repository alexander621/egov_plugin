<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getuseraddressupdate.asp
' AUTHOR: Steve Loar
' CREATED: 10/14/2013
' COPYRIGHT: Copyright 2013 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This updates the address info for a reservation sign up.
'
' MODIFICATION HISTORY
' 1.0   10/14/2013   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iUserId, sUserAddress, sUserCity, sUserState, sUserZip, sSql, iOrgid, iResidentAddressId, sResidentType

iUserId = CLng(request("userid"))
iOrgid =  CLng(request("orgid"))

If request("residentaddressid") <> "" Then 
	iResidentAddressId =  CLng(request("residentaddressid"))
	If iResidentAddressId > CLng(0) Then
		sResidentType = "'R'"
	Else
		sResidentType = "'N'"
	End If
Else
	sResidentType = "'N'"
End If 

sUserAddress = "'" & Track_DBsafe(request("useraddress")) & "'"
sUserCity = "'" & Track_DBsafe(request("usercity")) & "'"
sUserState = "'" & Track_DBsafe(UCase(request("userstate"))) & "'"
sUserZip = "'" & Track_DBsafe(request("userzip")) & "'"

sSql = "UPDATE egov_users SET useraddress = " & sUserAddress
sSql = sSql & ", usercity = " & sUserCity
sSql = sSql & ", userstate = " & sUserState
sSql = sSql & ", userzip = " & sUserZip
sSql = sSql & ", residenttype = " & sResidentType
sSql = sSql & " WHERE userid = " & iUserId
sSql = sSql & " AND orgid = " & iOrgid
'response.write "<p>" & sSql & "</p><br /><br />"

RunSQLStatement sSql 

Response.Redirect session("RedirectPage")

'--------------------------------------------------------------------------------------------------
' string Track_DBsafe( strDB )
'--------------------------------------------------------------------------------------------------
Function Track_DBsafe( ByVal strDB )
	Dim sNewString

	If Not VarType( strDB ) = vbString Then Track_DBsafe = strDB : Exit Function

	sNewString = Replace( strDB, "'", "''" )
	sNewString = Replace( sNewString, "<", "&lt;" )
	Track_DBsafe = sNewString

End Function


'-------------------------------------------------------------------------------------------------
' void RunSQLStatement sSql 
'-------------------------------------------------------------------------------------------------
Sub RunSQLStatement( ByVal sSql )
	Dim oCmd

'	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub 


%>

