<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="includes/inc_recordtoken.asp" //-->
<!-- #include file="../egovlink300_global/includes/inc_email.asp" //-->
<!-- #include file="../egovlink300_global/includes/inc_passencryption.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: login.asp
' AUTHOR: ???
' CREATED: 2004
' COPYRIGHT: Copyright 2004 eclink, inc.
'			 All Rights Reserved.
'
' Description:  handles the login from user_login.asp
'
' MODIFICATION HISTORY
' 1.0   2004	INITIAL VERSION
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sRedirect, sLogInUserPassword, sLogInUserId

'if request("email") = "tfoster@eclink.com" then response.end

assistant = "alexa"
if request.form("googleauth") = "true" then assistant = "googleauth"

 If request("frmsubjecttext") <> "" Then 
   	' turned off per Jerry 6/24/2013
	'SendSpamFlag request("email"), request("password"), request("frmsubjecttext"), request("orgid"), request.servervariables("remote_addr")

   	Response.Cookies("userid") = -1
   	sQuerystring = Encode("STATUS=FAILED&USERLOGIN=" & request("email")) 
   	response.redirect "user_login.asp?" & sQuerystring
 End If 


If InStr(request("email"), "'") = 0 Then 

	If Not IsNumeric(request("orgid")) Or request("orgid") = "" Then
		response.redirect "user_login.asp"
	End If 

	' LOOK UP EMAIL ADDRESS IN DATABASE
	sSql = "SELECT userid, userpassword,password FROM egov_users "
	sSql = sSql & "WHERE headofhousehold = 1 AND isdeleted = 0 AND useremail = '" & Track_DBsafe(request("email"))
	sSql = sSql & "' AND orgid = " & CLng(request("orgid"))

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	iTotal = oRs.RecordCount

	If Not oRs.EOF Then 
		sLogInUserPassword = oRs("userpassword")
		sEncPassword = oRs("password")
		sLogInUserId = oRs("userid")
	Else
		sLogInUserPassword = "x.x"
		sEncPassword = "x.x"
		sLogInUserId = "x.x"
		iTotal = 0
	End If 

	oRs.Close 
	Set oRs = Nothing 

	' PROCESS RESULT SET
	If iTotal = 0 Then 
		' NO SUCH USER FOUND REDIRECT TO REGISTRATION PAGE
		if request.form("token") = "" and instr(request.servervariables("HTTP_REFERER"),"basic_login.asp") <= 0 then
			response.redirect "register.asp?egov_users_useremail=" & request("email")
		elseif request(assistant) = "true" then
			response.redirect "basic_register.asp?egov_users_useremail=" & request("email") & "&" & assistant & "=true&state=" & request("state") & "&redirect_uri=" & request("redirect_uri")
		else
			response.redirect "basic_register.asp?egov_users_useremail=" & request("email") & "&token=" & request.form("token")
		end if
	Else
		' CHECK USER'S PASSWORD
		blnPlainPassMatch = sLogInUserPassword = Track_DBSafe(request("password")) and sLogInUserPassword <> ""
		blnEncPassMatch = ValidateUser(request("password"), sEncPassword)


		If blnPlainPassMatch or blnEncPassMatch Then 
			' SUCCESSFUL oRs REDIRECT TO CALLING PAGE
			response.cookies("userid") = sLogInUserId

	            	if request.form("token") <> "" then
	       			'RECORD THE TOKEN AS AUTHENTICATED
	       			RecordToken request.form("token"), request("orgid")
	       		end if

				'response.write "HERE" & assistant & request.form(assistant)
				'response.end
			if request.form(assistant) = "true" then
				'GENERATE CODE
				GUID = RecordGUID(request.form("state"), request("orgid"))

				'REDIRECT USER TO URI
				if assistant = "alexa" then
					response.redirect request.form("redirect_uri") & "#state=" & request.form("state") & "&token_type=Bearer&access_token=" & GUID
				else
					'response.redirect request.form("redirect_uri") & "#access_token=" & GUID & "&token_type=bearer&state=" & request.form("state")
					'response.write "HERE<br>"
					'response.write request.form("redirect_uri") & "#access_token=" & GUID & "&token_type=bearer&state=" & request.form("state")
					'response.end
					response.status = "302 Found"
					response.addheader "Location", request.form("redirect_uri") & "#access_token=" & GUID & "&token_type=bearer&state=" & request.form("state")
					response.end
				end if
			end if


			' IF REFERER IS LOGIN PAGE REDIRECT TO EGOV HOME
			If InStr(request.ServerVariables("http_referer"),"user_login.asp") <> 0 or InStr(request.ServerVariables("http_referer"),"test_gauth.asp") <> 0 Then
				
				If Session("RedirectPage") <> "" Then

					
			if request.cookies("userid") = "1678692" then
				'response.write session("redirectpage")
				'response.end
			end if

					sRedirect = Session("RedirectPage") 
					Session("RedirectPage") = ""
					response.redirect sRedirect
				Else
					response.redirect( GetEGovDefaultPage( CLng(request("orgid")) ) )
				End If
				
			End If

			' REDIRECT TO REFERRING PAGE
			If Request.ServerVariables("http_referer") = "" Then
				' EMPTY GOTO TO HOMEPAGE
				response.redirect(GetEGovDefaultPage(request("orgid")))
			Else
				' GO TO LAST VISITED PAGE
				response.redirect Request.ServerVariables("http_referer")
			End If

		Else
			' FAILED LOGIN, RETURN TO THE LOGIN PAGE
			Response.Cookies("userid") = -1
			sQuerystring = Encode("STATUS=FAILED&USERLOGIN=" & request("email"))

			if request(assistant) = "true" then
				response.redirect "basic_login.asp?" & assistant & "=true&state=" & request("state") & "&redirect_uri=" & request("redirect_uri") & "&" & sQuerystring
			elseif request.form("token") = "" then
				response.redirect "user_login.asp?" & sQuerystring 
			else
				response.redirect "basic_login.asp?" & sQuerystring  & "&token=" & request.form("token")
			end if
		End if

	End if
Else
	' invalid character In email
	' FAILED LOGIN, RETURN TO LOGIN PAGE
	Response.Cookies("userid") = -1
	sQuerystring = Encode("STATUS=FAILED&USERLOGIN=" & request("email"))
      	if request.form("token") = "" then
		response.redirect "user_login.asp?" & sQuerystring
       	else
       		response.redirect "basic_login.asp?" & sQuerystring  & "&token=" & request.form("token")
       	end if
End If 



'--------------------------------------------------------------------------------------------------
' FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' string Encode( sInString )
'--------------------------------------------------------------------------------------------------
Function Encode( ByVal sInString )
	Dim x, y, abfrom, abto

	Encode="": ABFrom = ""

	For x = 0 To 25: ABFrom = ABFrom & Chr(65 + x): Next 
	For x = 0 To 25: ABFrom = ABFrom & Chr(97 + x): Next 
	For x = 0 To 9: ABFrom = ABFrom & CStr(x): Next 

	abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
	For x=1 to Len(sInString): y = InStr(abfrom, Mid(sInString, x, 1))
		If y = 0 Then
			Encode = Encode & Mid(sInString, x, 1)
		Else
			Encode = Encode & Mid(abto, y, 1)
		End If
	Next

End Function


'--------------------------------------------------------------------------------------------------
' string Track_DBsafe( strDB )
' -------------------------------------------------------------------------------------------------
Function Track_DBsafe( ByVal strDB )
	Dim sNewString

	If Not VarType( strDB ) = vbString Then Track_DBsafe = strDB : Exit Function

	sNewString = Replace( strDB, "'", "''" )
	sNewString = Replace( sNewString, "<", "&lt;" )

	Track_DBsafe = sNewString

End Function


'--------------------------------------------------------------------------------------------------
' void SendSpamFlag sEmail, sPassword, sHidden, iOrgId 
'--------------------------------------------------------------------------------------------------
Sub SendSpamFlag( ByVal sEmail, ByVal sPassword, ByVal sHidden, ByVal iOrgId, ByVal sIPAddress)
	Dim oCdoMail, oCdoConfm, sMsgBody, sOrgName

	sOrgName = GetOrgname( iOrgId )

	sMsgBody = "A attempt was made to log into the public side of " & sOrgName & ". <br />They populated the hidden field and bypassed the JavaScript catch.<br /><br />" & vbcrlf
	sMsgBody = sMsgBody & "Email Address: " & sEmail     & "<br />" & vbcrlf
	sMsgBody = sMsgBody & "Password: "      & sPassword  & "<br />" & vbcrlf
	sMsgBody = sMsgBody & "Hidden Field: "  & sHidden    & "<br />" & vbcrlf
 sMsgBody = sMsgBody & "IP Address: "    & sIPAddress & "<br />"

	sendEmail "", "egovsupport@eclink.com", "", UCase(sOrgName) & " E-GOV Invalid Public Login Attempt", sMsgBody, "", "N"


End Sub 


'--------------------------------------------------------------------------------------------------
' string GetOrgname( iOrgId )
'--------------------------------------------------------------------------------------------------
Function GetOrgname( ByVal iOrgId )
	Dim sSql, oRs

	sSQL = "SELECT OrgName FROM Organizations WHERE orgid = " & CLng(iOrgId)

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		GetOrgname = oRs("OrgName")
	Else
		GetOrgname = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 




%>
