<!-- #include file="../includes/JSON_2.0.2.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getshiptoinfo.asp
' AUTHOR: Steve Loar
' CREATED: 05/11/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  gets the shipping information for a citizen
'
' MODIFICATION HISTORY
' 1.0   05/11/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iUserId, sResponse, sSql, oRs

' Create the JSON object to pass data back to the calling page
Set sResponse = jsObject()

iUserId = CLng(request("userid"))

If iUserId > CLng(0) Then 
	sSql = "SELECT ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, "
	sSql = sSql & " ISNULL(useraddress,'') AS useraddress, ISNULL(usercity,'') AS usercity, "
	sSql = sSql & " ISNULL(userstate,'') AS userstate, ISNULL(userzip,'') AS userzip, ISNULL(userhomephone,'') AS phone "
	sSql = sSql & " FROM egov_users WHERE userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sResponse("flag") = "success"
		sResponse("captainfirstname") = Trim(oRs("userfname"))
		sResponse("captainlastname") = Trim(oRs("userlname"))
		sResponse("captainaddress") = oRs("useraddress")
		sResponse("captaincity") = oRs("usercity")
		sResponse("captainstate") = UCase(oRs("userstate"))
		sResponse("captainzip") = oRs("userzip")
		sResponse("areacode") = Left(oRs("phone"),3)
		sResponse("exchange") = Mid(oRs("phone"),4,3)
		sResponse("line") = Right(oRs("phone"),4)
	Else
		sResponse("flag") = "failed"
	End If

	oRs.Close
	Set oRs = Nothing 
Else
	sResponse("flag") = "failed"
End If 


sResponse.Flush

%>