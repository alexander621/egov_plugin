<%
' DELETE ORGANIZATION
subDeleteOrg(request("iorgid"))

' REDIRECT TO ORGANIZATION LIST PAGE
response.redirect("list_organization.asp")


'------------------------------------------------------------------------------------------------------------
' SUB SUBDELETEORG(IORGID)
'------------------------------------------------------------------------------------------------------------
Sub subDeleteOrg(iorgid)
	' GET ORGANIZATIONS VALUES
	sSQL = "DELETE From Organizations WHERE ORGID=" & IORGID
	Set oDelete = Server.CreateObject("ADODB.Recordset")
	oDelete.Open sSQL, Application("DSN") , 3, 1
	Set oDelete = Nothing
End Sub
%>