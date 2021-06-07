<%
' LOOP THRU ALL SELECTED USERS AND ADD TO LIST
For each iuserid in request.form("availablelist")
	Call  subAddUsertoList(iuserid,request("maillistid"))
Next

' RETURN USER MANAGEMENT SCREEN
response.redirect("dl_manage_subscribers.asp?idlid=" & request("maillistid") & "&iname=" & request("maillistname"))



'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' SUB SUBADDUSERTOLIST(IUSERID,IMAILLISTID)
'--------------------------------------------------------------------------------------------------
Sub subAddUsertoList(iuserid,imaillistid)
	
	sSQL = "INSERT INTO egov_class_distributionlist_to_user (userid,distributionlistid) VALUES ('" & iuserid & "','" & imaillistid& "')"
	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1
	Set oList = Nothing

End Sub
%>
