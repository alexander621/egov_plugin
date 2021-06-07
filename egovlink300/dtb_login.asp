<!-- #include file="includes/start_modules.asp" //-->
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

 If request("frmsubjecttext") <> "" Then 
   	SendSpamFlag request("email"), _
                 request("password"), _
                 request("frmsubjecttext"), _
                 request("orgid"), _
                 request.servervariables("remote_addr")

   	Response.Cookies("userid") = -1
	sQuerystring = Encode("STATUS=FAILED&USERLOGIN=" & request("email")) 
	response.redirect "dtb_user_login.asp?" & sQuerystring
 End If 


If InStr(request("email"), "'") = 0 Then 

	If Not IsNumeric(request("orgid")) Or request("orgid") = "" Then
		response.redirect "dtb_user_login.asp"
	End If 

	' LOOK UP EMAIL ADDRESS IN DATABASE
	sSql = "SELECT userid, userpassword FROM egov_users "
	sSql = sSql & "WHERE isdeleted = 0 AND useremail = '" & Track_DBsafe(request("email"))
	sSql = sSql & "' AND orgid = " & CLng(request("orgid"))

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	iTotal = oRs.RecordCount

	If Not oRs.EOF Then 
		sLogInUserPassword = oRs("userpassword")
		sLogInUserId = oRs("userid")
	Else
		sLogInUserPassword = "x.x"
		sLogInUserId = "x.x"
		iTotal = 0
	End If 

	oRs.Close 
	Set oRs = Nothing 

	' PROCESS RESULT SET
	If iTotal = 0 Then 
		' NO SUCH USER FOUND REDIRECT TO REGISTRATION PAGE
		response.redirect "register.asp?email=" & request("email")
	Else
	response.Write "1 [" & session("RedirectPage") & "]<br />"
		' CHECK USER'S PASSWORD
		If sLogInUserPassword = request("password") Then 
			' SUCCESSFUL oRs REDIRECT TO CALLING PAGE
			response.cookies("userid") = sLogInUserId
	response.Write "2<br />"
			' IF REFERER IS LOGIN PAGE REDIRECT TO EGOV HOME
			If InStr(request.ServerVariables("http_referer"),"dtb_user_login.asp") <> 0 Then
	response.Write "3<br />"
				If Session("RedirectPage") <> "" Then
	response.Write "4<br />"
					sRedirect = Session("RedirectPage") 
					Session("RedirectPage") = ""
					'response.redirect sRedirect
				Else
					response.Write "5<br />"
					'response.redirect( GetEGovDefaultPage( CLng(request("orgid")) ) )
				End If
				
			End If
	response.Write "5<br />"
			' REDIRECT TO REFERRING PAGE
			If Request.ServerVariables("http_referer") = "" Then
				' EMPTY GOTO TO HOMEPAGE
				'response.redirect(GetEGovDefaultPage(request("orgid")))
			Else
				' GO TO LAST VISITED PAGE
				'response.redirect Request.ServerVariables("http_referer")
			End If
	response.Write "6<br />"

		Else
			' FAILED LOGIN, RETURN TO THE LOGIN PAGE
			Response.Cookies("userid") = -1
			sQuerystring = Encode("STATUS=FAILED&USERLOGIN=" & request("email"))
			response.redirect "dtb_user_login.asp?" & sQuerystring 
		End if

	End if
Else
	' invalid character In email
	' FAILED LOGIN, RETURN TO LOGIN PAGE
	Response.Cookies("userid") = -1
	sQuerystring = Encode("STATUS=FAILED&USERLOGIN=" & request("email"))
	response.redirect "dtb_user_login.asp?" & sQuerystring
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

	Set oCdoMail = Server.CreateObject("CDO.Message")
	Set oCdoConf = Server.CreateObject("CDO.Configuration")
	With oCdoConf
		.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("SMTP_Server")
		.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		.Fields.Update
	 End With

'	 With oCdoMail.Fields 
		' for Outlook: 
		'.Item(cdoImportance) = cdoHigh  
		'.Item(cdoPriority) = cdoPriorityUrgent  
		' for Outlook Express: 
		'.Item("urn:schemas:mailheader:X-Priority") = 1 
'		.Update 
'	End With 	
	
	With oCdoMail
		Set .Configuration = oCdoConf
		'.From = UCase(sOrgName) & " E-GOV WEBSITE <webmaster@eclink.com>"
		.From = UCase(sOrgName) & " E-GOV WEBSITE <noreply@eclink.com>"
		.To = "egovsupport@eclink.com"
		'.To = "sloar@eclink.com"
		.Subject = UCase(sOrgName) & " E-GOV Invalid Public Login Attempt"
		.HTMLBody = sMsgBody 
		.Send
	End With

	Set oCdoMail = Nothing 
	Set oCdoConf = Nothing 

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
