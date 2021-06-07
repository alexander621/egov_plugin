<!-- #include file="../includes/common.asp" //-->
<%
sSQL = "SELECT DocumentLinkID FROM DocumentLinks WHERE DocumentURL = '" & DBSafe(request.form("URL")) & "'"
Set oRs = Server.CreateObject("ADODB.RecordSet")
oRs.Open sSQL, Application("DSN"), 3, 1


if oRs.EOF then
    'Create new record
    sSQL = "INSERT INTO DocumentLinks (DocumentURL) VALUES('" & request.form("URL") & "')"
    Set oCmd = Server.CreateObject("ADODB.Connection")
    oCmd.Open Application("DSN")
    oCmd.Execute(sSQL)
    
    sSQL2 = "SELECT @@IDENTITY AS NewID"
    set rs = oCmd.Execute(sSQL2)
    ID = rs.Fields("NewID").value
    set rs = Nothing
    
    oCmd.Close
    Set oCmd = Nothing
else
    'return ID
    ID = oRs("DocumentLinkID")
end if

response.write "http://" & request.servervariables("SERVER_NAME") & "/d.asp?ID=" & ID

oRs.Close
Set oRs = Nothing

%>
