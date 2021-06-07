<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getapplicantaddressname.asp
' AUTHOR: Steve Loar
' CREATED: 9/25/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This gets the applicants address name. It is called via AJAX
'
' MODIFICATION HISTORY
' 1.0	09/25/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iUserId, sSql, oRs, sResults, sAddress, sStreetNumber, sStreetName

iUserId = CLng(request("userid"))

sSql = "SELECT ISNULL(useraddress,'') AS useraddress FROM egov_users WHERE userid = " & iUserId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If Not oRS.EOF Then
	If oRs("useraddress") <> "" Then
		sAddress = oRs("useraddress")
		BreakOutAddress sAddress, sStreetNumber, sStreetName	' In common.asp
		If sStreetName <> "" Then
			sResults = sStreetName
		Else
			sResults = "NONAME"
		End If 
	Else
		sResults = "NONAME"
	End If 
Else
	sResults = "NONAME"
End If 

oRs.Close
Set oRs = Nothing 

response.write sResults

%>
