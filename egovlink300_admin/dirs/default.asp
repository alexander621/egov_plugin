<!-- #include file="../includes/common.asp" //-->
<%
dim oCmd,oRst
If Session("UserID") = 0 Or Session("UserID") = "" Then
set oCmd = Server.CreateObject("ADODB.Command")
  oCmd.ActiveConnection = Application("DSN")
  oCmd.CommandText = "Login"
  oCmd.CommandType = 4
  oCmd.Parameters.Append oCmd.CreateParameter("username", adVarChar, adParamInput, 15, "guest")
  oCmd.Parameters.Append oCmd.CreateParameter("password", adVarChar, adParamInput, 15, "guest")
  oCmd.Parameters.Append oCmd.CreateParameter("SessionID", adVarChar, adParamInput, 30, Session.SessionID)
  oCmd.Parameters.Append oCmd.CreateParameter("IP", adVarChar, adParamInput, 16, Request.ServerVariables("REMOTE_ADDR"))
  
  Set oRst = Server.CreateObject("ADODB.Recordset")
  Set oRst = oCmd.Execute
  Set oCmd = Nothing

  If Not oRst.EOF Then
    Session("Permissions") = oRst("Permissions") & ""
    oRst.Close
    Set oRst = Nothing
    Set oCmd = Nothing
  end if
end if
response.redirect("display_committee.asp")
%>

