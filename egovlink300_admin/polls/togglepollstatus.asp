<!-- #include file="../includes/common.asp" //-->
<%
Dim oCmd

Set oCmd = Server.CreateObject("ADODB.Command")
With oCmd
  .ActiveConnection = Application("DSN")
  .CommandText = "TogglePollStatus"
  .CommandType = adCmdStoredProc
  .Parameters.Append oCmd.CreateParameter("VoteID", adInteger, adParamInput, 4, Request("id"))
  .Execute
End With
Set oCmd = Nothing

Response.Redirect "viewpoll.asp?" & Request.QueryString()
%>