<!-- #include file="../includes/common.asp" //-->
<%
If Not (HasPermission("CanEditAnnouncements")) Then Response.Redirect "../"

Dim oCmd, status, aID, previousURL

status = clng(request.querystring("status"))
aID = clng(request.querystring("aID"))

Set oCmd = Server.CreateObject("ADODB.Command")
With oCmd
  .ActiveConnection = Application("DSN")
  .CommandText = "UpdateAnnouncementsPublish"
  .CommandType = adCmdStoredProc
  .Parameters.Append .CreateParameter("@status", adInteger, adParamInput, 4, status)
  .Parameters.Append .CreateParameter("@aID", adInteger, adParamInput, 4, aID)
  .Execute
End With
Set oCmd = Nothing

previousURL = "default.asp"
Response.Redirect( previousURL )
%>
