<!-- #include file="../includes/common.asp" //-->
<%
Dim sDelete, oCmd

sDelete = Request.QueryString("delid")

If sDelete & "" <> "" Then
  Set oCmd = Server.CreateObject("ADODB.Command")
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "DelDiscussionMessages"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("MessageIDs", adVarChar, adParamInput, 1000, sDelete)
    .Execute
  End With
  Set oCmd = Nothing
End If

Response.Redirect "topics.asp?" & Request.QueryString()
%>