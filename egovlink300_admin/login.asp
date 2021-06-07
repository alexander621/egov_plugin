<% PageIsRequiredByLogin = True 

If InStr(LCase(request.servervariables("server_name")), "www.") < 1 And InStr(LCase(request.servervariables("server_name")), "dev4.") < 1 And InStr(LCase(request.servervariables("server_name")), "test.") < 1 And InStr(LCase(request.servervariables("server_name")), "egovernment.") < 1 And InStr(LCase(request.servervariables("server_name")), "egov.") < 1 Then
	' they did not specify prod, test or dev, so take them to prod.
	response.redirect "http://www." & LCase(request.servervariables("server_name")) & LCase(request.servervariables("script_name"))
End If 

%>

<!-- #include file="includes/common.asp" //-->
<!-- #include file="../egovlink300_global/includes/inc_passencryption.asp" //-->

<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: login.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2005 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This in the admin user login page for the system.
'
' MODIFICATION HISTORY
' 1.0	??/??/????	???? - INITIAL VERSION
' 2.0	10/14/2011	Steve Loar - Changed to not let those flagged as deleted in
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim oCmd, oRst, sUsername, sPassword, sError, sTmpUrl

sLevel = "" ' Override of value from common.asp

SetOrganizationParameters()

' ORG PARMS
session("payment_gateway") = iPaymentGatewayID
session("orgregistration") = blnOrgRegistration
session("orgquerytool") = blnQueryTool
session("orgfaq") = blnFaq
session("sitename") = sorgVirtualSiteName
session("OrgFormLetterOn") = sOrgFormLetterOn
session("OrgInternalEntry") = sOrgInternalEntry
session("iTimeOffset") = iTimeOffset
session("egovclientwebsiteurl") = sEgovWebsiteURL
session("virtualdirectory") = sorgVirtualSiteName
session("blnSeparateIndex") = blnSeparateIndex

If Request("_task") = "login" Then
	sUsername = Left(Request("Username"), 32)
	sPassword = Left(Request("Password"), 16)

	'Set oCmd = Server.CreateObject("ADODB.Command")
	'With oCmd
		'.ActiveConnection = Application("DSN")
		'.CommandText = "Login"
		'.CommandType = adCmdStoredProc
		'.Parameters.Append .CreateParameter("Username", adVarChar, adParamInput, 32, sUserName)
		'.Parameters.Append .CreateParameter("Password", adVarChar, adParamInput, 16, sPassword)
		'.Parameters.Append .CreateParameter("SessionID", adVarChar, adParamInput, 30, Session.SessionID)
		'.Parameters.Append .CreateParameter("IP", adVarChar, adParamInput, 16, Request.ServerVariables("REMOTE_ADDR"))
		'.Parameters.Append .CreateParameter("iorgid", adInteger, adParamInput, 4, iorgid)
	'End With
'
	'Set oRst = oCmd.Execute

	'oCmd.ActiveConnection = Nothing
	'Set oCmd = Nothing

	sSQL = "SELECT u.UserID, epassword, password, u.OrgID, u.FirstName + ' ' + u.LastName [FullName], u.PageSize, u.ShowStockTicker, dbo.GetPerms(u.userid) as [Permissions], g.GroupImage " _
		& " FROM Users [u] " _
  		& " LEFT JOIN UsersGroups [ug] ON ug.UserID = u.UserID AND ug.IsPrimaryGroup = 1 " _
  		& " LEFT JOIN Groups [g] ON g.GroupID = ug.GroupID " _
		& " WHERE u.Enabled = 1 AND username = '" & dbsafe(sUsername) & "' " _
		& " AND ((password = '" & dbsafe(sPassword) & "' and password <> '') OR password IS NULL) " _
		& " AND u.orgid = '" & iorgid & "' and isdeleted = 0"
	set oRst = Server.CreateObject("ADODB.RecordSet")
	oRst.Open sSQL, Application("DSN"), 3, 1

	if not oRst.EOF then
		if not isnull(oRst("password")) then
			'Successful login with plain text password
		else
			'Validate the epassword
			if not ValidateUser(sPassword, oRst("epassword")) then
				oRst.MoveLast
				oRst.MoveNext
			end if
		end if
	end if


	If Not oRst.EOF Then
		'Mark User As Logged In
		sSQL = "UPDATE Users SET IsLoggedIn = 1 WHERE UserID = " & oRst("UserID")
		RunSQLStatement sSQL

		'Log the login UNFINISHED
		sSQL = "INSERT INTO AuditLog( UserID, AuditType, AuditObject,  AuditNotes) VALUES ( " & oRst("UserID") & ", 'Login', '" & Session.SessionID & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
		RunSQLStatement sSQL

		Session("UserID") = oRst("UserID")
		Session("OrgID") = oRst("OrgID")
		Session("FullName") = oRst("FullName")
		Session("PageSize") = oRst("PageSize")
		Session("ShowStockTicker") = oRst("ShowStockTicker")
		Session("Permissions") = oRst("Permissions") & ""

		oRst.Close
		Set oRst = Nothing
		Set oCmd = Nothing

		Session("LocationId") = GetUserLocationId( Session("UserID") )

		If Request.Form("SaveLogin") = "on" Then
			Response.Cookies("User")("UserID") = Session("UserID")
			Response.Cookies("User")("OrgID") = Session("OrgID")
			Response.Cookies("User")("FullName") = Session("FullName")
			Response.Cookies("User")("LocationId") = Session("LocationId")
			Response.Cookies("User").Expires = Now() + 365
		Else
			' FORCE AUTOLOGIN TO WORK
			Response.Cookies("User")("UserID") = Session("UserID")
			Response.Cookies("User")("OrgID") = Session("OrgID")
			Response.Cookies("User")("FullName") = Session("FullName")
			Response.Cookies("User")("LocationId") = Session("LocationId")
		End If

		'redirect to initial page request if exists, otherwise to go to home page
		sTmpUrl = Session("RedirectPage")
		If sTmpUrl <> "" Then
			Session("RedirectPage") = ""
			Response.Redirect sTmpUrl
		Else
			Response.Redirect("default.asp")
		End If

	Else
		sError = "<font color=#ff0000><b>" & langInvalid & "</b></font>"
	End If

	If oRst.State = adStateOpen Then 
		oRst.Close
	End If 
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

	<link rel="stylesheet" type="text/css" href="global.css" />

	<script language="Javascript" src="scripts/modules.js"></script>

	<script language="javascript">
	<!--

		function init() 
		{
			if(document.getElementById('username') != null)
			{
				// put the cursor in the login field
				document.getElementById('username').focus();
			}
		}
		
		window.onload = init; 

	//-->
	</script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

  <% ShowHeader sLevel %>

  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><!--<img src="images/icon_directory.jpg">--> &nbsp; </td>
      <td><font size="+1"><b><%=langLogIn%></b></font><br>
          <%
          If Application("LoginIsRequired") = True Then
            Response.Write (langLoginBenefits)
          Else
            Response.Write "Logging in is required to gain access to " & Application("ProgramName") & "."
          End If
          %>
      </td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">&nbsp;<br></td>
      <td colspan="2" valign="top">
        <form name="loginForm" action="login.asp" method="post">
          <input type="hidden" name="_task" value="login" />
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
              <td width="100%"><input type="text" id="username" name="Username" size="20" maxlength="32" /></td>
            </tr>
            <tr>
              <td><%=langPassword%>:</td>
              <td><input type="password" name="Password" size="20" maxlength="16" /></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td>
              <input type="checkbox" name="SaveLogin" style="margin-left:-3px;" /><%=langLogMeAuto%></td>
            </tr>
            <tr>
              <td><br><br></td>
              <td valign="top"><br /><input class="button" type="submit" value="<%=langLogIn%>" style="font-family:Tahoma,Arial; font-size:11px; width:70px;" /></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td>
				<font size="1">&nbsp;<a href="dirs/lookuppassword.asp"><%=langForgotPass%></a></font><br /><br /><br />
			  </td>
            </tr>
          </table>
        </form>

      </td>
    </tr>
  </table>
</body>
</html>


