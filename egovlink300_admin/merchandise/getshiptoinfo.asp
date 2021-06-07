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
	sSql = sSql & " ISNULL(userstate,'') AS userstate, ISNULL(userzip,'') AS userzip "
	sSql = sSql & " FROM egov_users WHERE userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		sResponse("flag") = "success"
		sResponse("shiptoname") = Trim(oRs("userfname") & " " & oRs("userlname"))
		sResponse("shiptoaddress") = oRs("useraddress")
		sResponse("shiptocity") = oRs("usercity")
		sResponse("shiptostate") = UCase(oRs("userstate"))
		sResponse("shiptozip") = oRs("userzip")
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