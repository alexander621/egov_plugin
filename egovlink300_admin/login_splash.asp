<!-- #include file="includes/common.asp" //-->
<%
Dim oCmd, oRst, sUsername, sPassword, sError

If Request.Form("_task") = "login" Then
  sUsername = Request.Form("Username")
  sPassword = Request.Form("Password")

  Set oCmd = Server.CreateObject("ADODB.Command")
  oCmd.ActiveConnection = Application("DSN")
  oCmd.CommandText = "Login"
  oCmd.CommandType = adCmdStoredProc
  oCmd.Parameters.Append oCmd.CreateParameter("OrgId", adInteger, adParamInput, 4, Session("OrgID"))
  oCmd.Parameters.Append oCmd.CreateParameter("Username", adVarChar, adParamInput, 15, sUserName)
  oCmd.Parameters.Append oCmd.CreateParameter("Password", adVarChar, adParamInput, 15, sPassword)
  oCmd.Parameters.Append oCmd.CreateParameter("SessionID", adVarChar, adParamInput, 30, Session.SessionID)
  oCmd.Parameters.Append oCmd.CreateParameter("IP", adVarChar, adParamInput, 16, Request.ServerVariables("REMOTE_ADDR"))
  
  Set oRst = Server.CreateObject("ADODB.Recordset")
  Set oRst = oCmd.Execute
  Set oCmd = Nothing

  If Not oRst.EOF Then
    Session("UserID") = oRst("UserID")
    Session("FullName") = oRst("FullName")
    Session("PageSize") = oRst("PageSize")
    Session("ShowStockTicker") = oRst("ShowStockTicker")
    Session("Permissions") = oRst("Permissions") & ""

    oRst.Close
    Set oRst = Nothing
    Set oCmd = Nothing

    If Request.Form("SaveLogin") = "on" Then
      Response.Cookies("User")("UserID") = Session("UserID")
      Response.Cookies("User")("FullName") = Session("FullName")
      Response.Cookies("User")("PageSize") = Session("PageSize")
      Response.Cookies("User")("ShowStockTicker") = Session("ShowStockTicker")
      Response.Cookies("User")("Permissions") = Session("Permissions")
      Response.Cookies("User").Expires = Now() + 365
    End If

    Response.Redirect("default.asp")
  Else
    sError = "<font color=#ff0000><b>" & langInvalid & "</b></font>"
  End If

  If oRst.State = adStateOpen Then oRst.Close
  Set oRst = Nothing
Else
  sError = ""
End If

If Request.QueryString() = "alo" Then
  sError = "<font color=#ff0000><b>Your session has expired and you have been logged out.</b></font><br><br>"
End If 
%>


<html>
<head>
  <title><%=langBSHome%></title>
  <link href="global.css" rel="stylesheet" type="text/css">
  <script language="Javascript" src="scripts/modules.js"></script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs 0,0%>

  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="images/icon_directory.jpg"></td>
      <td><font size="+1"><b><%=langLogIn%></b></font><br>Logging in will provide you with enhanced capabilites and preferences.</td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top"><br></td>
      <td colspan="2" valign="top">
        <form action="login.asp" method="post">
          <input type="hidden" name="_task" value="login">
          <table border="0" cellpadding="3" cellspacing="0">
            <tr>
              <td colspan="2">
                <%
                If sError <> "" Then
                  Response.Write sError
                Else
                  Response.Write "<br>"
                End If
                %>
              </td>
            </tr>
            <tr>
              <td><%=langUsername%>:</td>
              <td width="100%"><input type="text" name="Username" size="20"></td>
            </tr>
            <tr>
              <td><%=langPassword%>:</td>
              <td><input type="password" name="Password" size="20"></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td>
              <input type="checkbox" name="SaveLogin" style="margin-left:-3px;"><%=langLogMeAuto%></td>
            </tr>
            <tr>
              <td><br><br></td>
              <td valign="top"><br><input type="submit" value="<%=langLogIn%>" style="font-family:Tahoma,Arial; font-size:11px; width:70px;"></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td><font size="1">&nbsp;<a href="dirs/lookuppassword.asp"><%=langForgotPass%></a></font><br><br><br></td>
            </tr>
          </table>
        </form>

      </td>
    </tr>
  </table>
</body>
</html>