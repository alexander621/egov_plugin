<%
' LOOP THRU ALL SELECTED USERS AND ADD TO LIST
For each iuserid in request.form("subscribedlist")
	Call  subDeleteUserFromList(iuserid,request("maillistid"))
Next

' RETURN USER MANAGEMENT SCREEN
response.redirect("dl_manage_subscribers.asp?idlid=" & request("maillistid") & "&iname=" & request("maillistname"))



'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' SUB SUBDELETEUSERFROMLIST(IUSERID,IMAILLISTID)
'--------------------------------------------------------------------------------------------------
Sub subDeleteUserFromList(iuserid,imaillistid)
	
	sSQL = "DELETE FROM egov_class_distributionlist_to_user WHERE userid = '" & iuserid & "' AND distributionlistid = '" & imaillistid & "'"
	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1
	Set oList = Nothing

End Sub
%>
