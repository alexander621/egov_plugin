<meta name="viewport" content="width=device-width, initial-scale=1" />
<% PageIsRequiredByLogin = True 

If InStr(LCase(request.servervariables("server_name")), "www.") < 1 And InStr(LCase(request.servervariables("server_name")), "dev4.") < 1 And InStr(LCase(request.servervariables("server_name")), "test.") < 1 And InStr(LCase(request.servervariables("server_name")), "egovernment.") < 1 And InStr(LCase(request.servervariables("server_name")), "egov.") < 1 Then
	' they did not specify prod, test or dev, so take them to prod.
	response.redirect "http://www." & LCase(request.servervariables("server_name")) & LCase(request.servervariables("script_name"))
End If 

if request.cookies("user")("userid") <> "" and request.cookies("user")("userid") <> "-1" then
	'Clear Cookie if logged in but token has expired
	sSQL = "SELECT adminauthtokensid FROM adminauthtokens WHERE userid = '" & track_dbsafe(request.cookies("user")("userid")) & "' AND token = '" & track_dbsafe(request.querystring("token")) & "' and orgid = '" & iorgid & "' and daterecorded >= '" & DateAdd("n",-5,Now()) & "'"
	Set oAuth = Server.CreateObject("ADODB.RecordSet")
	oAuth.Open sSQL, Application("DSN"), 3, 1
	if oAuth.EOF then
		response.cookies("user")("userid") = ""
	end if
	oAuth.Close
	Set oAuth = Nothing
end if

%>

<!-- #include file="includes/common.asp" //-->

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

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "Login"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter("Username", adVarChar, adParamInput, 32, sUserName)
		.Parameters.Append .CreateParameter("Password", adVarChar, adParamInput, 16, sPassword)
		.Parameters.Append .CreateParameter("SessionID", adVarChar, adParamInput, 30, Session.SessionID)
		.Parameters.Append .CreateParameter("IP", adVarChar, adParamInput, 16, Request.ServerVariables("REMOTE_ADDR"))
		.Parameters.Append .CreateParameter("iorgid", adInteger, adParamInput, 4, iorgid)
	End With

	Set oRst = oCmd.Execute

	oCmd.ActiveConnection = Nothing
	Set oCmd = Nothing

	If Not oRst.EOF Then
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
	       	if request.form("token") <> "" then
	      		'RECORD THE TOKEN AS AUTHENTICATED
	       		RecordToken request.form("token"), session("orgid")
	       	end if

		response.write "You're logged in!"
		response.end

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
<style>
form {
    max-width: 330px;
    padding: 15px;
    margin: 0 auto;
}
#problemtextfield1
{
	display:none;
}
input[type="text"], select, input[type="password"], input[type="button"], input[type="submit"]
{ 
	font-size:16px;
	width: 100% !important;
}
body
{
	    FONT-SIZE: 11px;
    FONT-FAMILY: Verdana,Tahoma;
}
</style>

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

<%
if request.cookies("user")("userid") = "" or request.cookies("user")("userid") = "-1" then%>
	<div style="display: table;margin-left: auto;margin-right: auto;">
	<font size="+1"><b><%=langLogIn%></b></font><br>
          <%
          If Application("LoginIsRequired") = True Then
            Response.Write (langLoginBenefits)
          Else
            Response.Write "Logging in is required to gain access to " & Application("ProgramName") & "."
          End If
          %>
	  </div>
        <form name="loginForm" action="basic_login.asp" method="post">
          <input type="hidden" name="_task" value="login" />
	<input type="hidden" name="token" value="<%=request.querystring("token")%>" />
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
              <td><br><br></td>
              <td valign="top"><br /><input class="button" type="submit" value="<%=langLogIn%>" /></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td>
	      <br />
	      <br />
	      <br />
				<font size="1">&nbsp;<a href="dirs/basic_lookuppassword.asp?token=<%=request.querystring("token")%>"><%=langForgotPass%></a></font><br /><br /><br />
			  </td>
            </tr>
          </table>
        </form>
<%else%>
You're Logged In!
<%end if%>

</body>
</html>

<%
sub RecordToken(sToken, iOrgID)

	'sSQL = "DELETE FROM adminauthTokens WHERE userid = '" & request.cookies("user")("userid") & "' and token = '" & track_dbsafe(sToken) & "' and OrgID = '" & iOrgID & "';" _
		'& " INSERT INTO adminauthTokens (userid, token, orgid) VALUES('" & request.cookies("user")("userid") & "','" & track_dbsafe(sToken) & "', '" & iOrgID & "')"
		sSQL =  " INSERT INTO adminauthTokens (userid, token, orgid) VALUES('" & request.cookies("user")("userid") & "','" & track_dbsafe(sToken) & "', '" & iOrgID & "')"

	Set oCmdToken = Server.CreateObject("ADODB.Command")
	oCmdToken.ActiveConnection = Application("DSN")
	oCmdToken.CommandText = sSql
	oCmdToken.Execute
	Set oCmdToken = Nothing

end sub
%>
