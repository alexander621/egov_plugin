<%
Dim sSql, oCnn, iSize

iSize = Request.QueryString("size")

sSql = "EXEC SetPageSize " & Session("UserID") & "," & iSize

Set oCnn = Server.CreateObject("ADODB.Connection")
oCnn.Open Application("DSN")
oCnn.Execute sSql,,adExecuteNoRecords
oCnn.Close
Set oCnn = Nothing

Session("PageSize") = iSize

Response.Redirect "ChangePersonalSettings.asp"
%>