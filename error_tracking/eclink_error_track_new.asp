<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<% 
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: eclink_error_track.asp
' AUTHOR: ????
' CREATED: ????
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This handles errors for egov and estatements. It is loacted at 10.4.1.11 D$:\wwwroot\error_tracking
'
' MODIFICATION HISTORY
' 1.0   11/30/07   Steve Loar - Pointed the email link to new code location on production server
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

  Const lngMaxFormBytes = 200

  Dim objASPError, blnErrorWritten, strServername, strServerIP, strRemoteIP
  Dim strMethod, lngPos, datNow, strQueryString, strURL,sErrorInfo
  Dim sBrowserInfo, sMicrosoft, sTimeInfo, sErrorMessage
  Dim sAspCategory, sAspColumn, sAspDescription, sAspFile, sAspLine, sAspNumber, sAspSource

  If Response.Buffer Then
    Response.Clear
    Response.Status = "500 Internal Server Error"
    Response.ContentType = "text/html"
    Response.Expires = 0
  End If

  ' SET ERROR VALUES
  Set objASPError = Server.GetLastError
  sAspCategory = objASPError.Category
  sAspColumn = objASPError.Column
  sAspDescription = objASPError.Description
  sAspFile = objASPError.File
  sAspLine = objASPError.Line
  sAspNumber = objASPError.Number
  sAspSource = objASPError.Source
  Set objASPError = Nothing 

  sErrorMessage = sErrorMessage & "Category: " & sAspCategory & vbcrlf
  sErrorMessage = sErrorMessage & "Column: " & sAspColumn & vbcrlf
  sErrorMessage = sErrorMessage & "Description: " & sAspDescription & vbcrlf
  sErrorMessage = sErrorMessage & "File: " & sAspFile & vbcrlf
  sErrorMessage = sErrorMessage & "Line: " & sAspLine & vbcrlf
  sErrorMessage = sErrorMessage & "Error Number: " & sAspNumber & vbcrlf

%>

<html dir=ltr>
<head>
	<style>
		a:link		{font:8pt/11pt verdana; color:FF0000}
		a:visited	{font:8pt/11pt verdana; color:#4e4e4e}
		div.errorbox {text-align:left; width:90%; border-style:solid;border-color:#000000;border-width:1px;font-family: arial,tahoma; font-size: 12px; color:#000000;padding:5px;background-color:#e0e0e0;}
	</style>

	<META NAME="ROBOTS" CONTENT="NOINDEX">

	<title>APPLICATION ERROR</title>
	<META HTTP-EQUIV="Content-Type" Content="text-html; charset=Windows-1252">
</head>
<body bgcolor="FFFFFF">


<%
' RECORD ERRORS IN DATABASE
iDBID = fnRecordError()


' GET ERROR INFORMATION
sErrorMsg = "GET ERROR INFORMATION"


' BUILD LINK TO SPECIFIC ERROR FOR EMAIL
'sHyperlink = vbcrlf & vbcrlf & "For more information regarding this error message click the link below or copy/paste link into your web browser:" & vbcrlf & "http://dev.eclink.com/application_errors/view_error_log.asp?errorid=" & iDBID & vbcrlf 
sHyperlink = vbcrlf & vbcrlf & "For more information regarding this error message click the link below or copy/paste link into your web browser:" & vbcrlf & "http://www.egovlink.com/eclink/admin/errors/view_error_log.asp?errorid=" & iDBID & vbcrlf 
sMSG = BuildMessageBody( sErrorMessage & sHyperlink )

If request.servervariables("REMOTE_ADDR") <> "24.106.89.6" Then
	' SEND EMAIL TO TECH SUPPORT
	'iErrorCode = SendEmail("smtp1.eclink.com","devsupport@eclink.com","EC LINK WEB SERVER","ec link ASP Script Error - " & sAspFile, "devsupport@eclink.com", sMSG)
	SendCDOMail sAspFile, sMSG
	iErrorCode = "0"
Else
	' DONT SEND EMAIL JUST SHOW ERROR TO ECLINK EMPLOYEE
	iErrorCode = "9999"
End If


' DISPLAY THAT AN ERROR OCCURRED TO THE USER
If iErrorCode <> 0 Then
	' ASK USER TO SEND ERROR DETAILS TO USER
	response.write "<br /><font style=""font-family: arial,tahoma;font-size: 12px;"" ><div style=""text-align:left; width:90%; border-style:solid;border-color:#c0c0c0;border-width:1px;font-family: arial,tahoma; font-size: 14px; color:yellow;padding:5px;background-color:#FF0000;"" >We're sorry, but the server encountered an error processing your request.  If you continue to receive this message, please follow up with your account representative. You will need to provide them with the error details listed in the box below.  Copy and paste the text into an email or a word document.</div><br><br><b>Error Details:</b><div style=""text-align:left; width:90%;"" >" 
	If iErrorCode = "9999" Then
		' INTERNAL ERROR DETAILS
		Call subDisplayErrorInformation()
	Else
		' PUBLIC DETAILS
		response.write "<PRE>" & sErrorMsg & "</PRE>"
	End If
	response.write  "</div><hr size=1 color=""#000000"" ><center>Developed by <i>electronic commerce</i> link, inc. dba. <i>ec</i> link.</font></center>"
Else
	' TELL USER EMAIL WAS SENT TO SUPPORT
	response.write "<br><font style=""font-family: arial,tahoma;font-size: 12px;"" ><div style=""text-align:left; width:90%; border-style:solid;border-color:#c0c0c0;border-width:1px;font-family: arial,tahoma; font-size: 14px; color:yellow;padding:5px;background-color:#FF0000;"" >We're sorry, but the server encountered an error processing your request. An email notification has been sent to our technical department with the details of this error.</div><hr size=1 color=""#000000"" ><center>Developed by <i>electronic commerce</i> link, inc. dba. <i>ec</i> link.</font></center>"
End If

%>

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' FUNCTION BUILDMESSAGEBODY(STEXT)
'--------------------------------------------------------------------------------------------------
Function BuildMessageBody( sText )
	sValue="----------------------------------------------------------------------------------------------------" & vbcrlf
	sValue=sValue & "  ec link ASP SCRIPT ERROR SUMMARY - " & sAspFile & vbcrlf
	sValue=sValue & "----------------------------------------------------------------------------------------------------" 
	sValue=sValue & vbcrlf  & vbcrlf & sText & vbcrlf  & vbcrlf
	sValue=sValue &"---------------------------------------------------------------------------------------------------"  & vbcrlf
	sValue=sValue &"  ec link ASP SCRIPT ERROR SUMMARY - " & sAspFile & vbcrlf
	sValue=sValue &"---------------------------------------------------------------------------------------------------"
	sValue=sValue &   vbcrlf  & vbcrlf & "This automated message was sent by the web server because an ASP script error was encountered. Do not reply to this message."
	sValue=sValue & " Contact mailto://development@eclink.com for inquiries regarding this email."

	BuildMessageBody = sValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION SENDEMAIL(SSMTP,SFROMEMAIL,SFROMNAME,SSUBJECT,STOEMAIL,SBODY)
'--------------------------------------------------------------------------------------------------
Function SendEmail(sSMTP,sFromEmail,sFromName,sSubject,sToEmail,sBody)

	' CREATE EMAIL OBJECT AND SEND EMAIL       
	Dim objMailer, Body
	iReturnCode = 0
	Set objMailer = Server.CreateObject("ecMail.SMTP")

	' Assign some required properties
	objMailer.Host = sSMTP
	objMailer.BodyFormat= 1
	objMailer.FromName = sFromName
	objMailer.From = sFromEmail
	objMailer.Subject = sSubject
	objMailer.SendTo = sToEmail
	objMailer.Priority   = 1 'PRIORITY_HIGH
	objMailer.Body = sBody
	iReturnCode = objMailer.Send
	Set objMailer = Nothing

	' RETURN ERROR CODE
	SendEmail = iReturnCode

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION SENDEMAIL(SSMTP,SFROMEMAIL,SFROMNAME,SSUBJECT,STOEMAIL,SBODY)
'--------------------------------------------------------------------------------------------------
Sub SendCDOMail( sAspFile, sBody )
	Dim oCdoMail, oCdoConf

	Set oCdoMail = Server.CreateObject("CDO.Message")
	Set oCdoConf = Server.CreateObject("CDO.Configuration")

	With oCdoConf
		.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp1.eclink.com"
		.Fields.Update
	 End With

	With oCdoMail
		Set .Configuration = oCdoConf
		.From = " EC LINK WEB SERVER <devsupport@eclink.com>"
		.To = "devsupport@eclink.com"
		.Subject = "ec link ASP Script Error - " & sAspFile
		'.HTMLBody = sBody 
		.TextBody = clearHTMLTags( sBody )
		.Send
	End With

	Set oCdoMail = Nothing
	Set oCdoConf = Nothing

End Sub 


'------------------------------------------------------------------------------------------------------------
' Function clearHTMLTags( sInString )
'------------------------------------------------------------------------------------------------------------
Function clearHTMLTags( ByVal sInString )
	Dim re, sNewString

	Set re = New RegExp

	' Leading tag removal
	re.Pattern = "(<[a-zA-Z][^>]*>)"
	re.Global = True
	sNewString = re.Replace(sInString, "")

	' Closing tag removal
	re.Pattern = "(</[a-zA-Z][^>]*>)"
	clearHTMLTags = re.Replace(sNewString, "")

	Set re = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION FNRECORDERROR()
'--------------------------------------------------------------------------------------------------
Function fnRecordError()

	iReturnValue = 0

	sSQL="SET NOCOUNT ON;INSERT INTO errorlog (category,[column],description,[file],line,number,source,sessioncollection,applicationcollection,cookiescollection,browserinformation,servervariablescollection,webappid,requestformcollection,requestquerystringcollection) VALUES ('" & DBsafe(sAspCategory) & "','" & DBsafe(sAspColumn) & "','" & DBsafe(sAspdescription) & "','" & DBsafe(sAspfile) & "','" & DBsafe(sAspline) & "','" & DBsafe(sAspNumber) & "','" & DBsafe(sAspSource) & "','" & DBsafe(GetSessionInformation()) & "','" & DBsafe(GetApplicationInformation()) & "','" & DBsafe(GetCookieInformation()) & "','" & DBsafe(request.servervariables("HTTP_USER_AGENT")) & "','" & DBsafe(GetServerVariablesInformation()) &"','" & Application("AppErrorID") & "','" & DBsafe(GetRequestformInformation()) &"','" & DBsafe(GetRequestQueryStringInformation()) &"')SELECT @@IDENTITY AS ROWID;"

	Set oRecordError = Server.CreateObject("ADODB.Recordset")
	oRecordError.Open sSQL, Application("ErrorDSN"), 3, 3
	iReturnValue = oRecordError("ROWID")
	Set oRecordError = Nothing

	fnRecordError = iReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETSESSIONINFORMATION()
'--------------------------------------------------------------------------------------------------
Function GetSessionInformation()
	
	sReturnValue = ""

	For each key in session.contents
		sReturnValue = sReturnValue & key & ":" & session(key) & "<br />" & vbcrlf
	Next 
	
	GetSessionInformation = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETAPPLICATIONINFORMATION()
'--------------------------------------------------------------------------------------------------
Function GetApplicationInformation()
	
	sReturnValue = ""

	For each key in application.contents
		sReturnValue = sReturnValue & key & ":" & application(key) & "<BR>" & vbcrlf
	Next 
	
	GetApplicationInformation = sReturnValue

End Function

'--------------------------------------------------------------------------------------------------
' FUNCTION GETCOOKIEINFORMATION()
'--------------------------------------------------------------------------------------------------
Function GetCookieInformation()
	
	sReturnValue = ""

	For each key in request.cookies
		sReturnValue = sReturnValue & key & ":" & request.cookies(key) & "<BR>" & vbcrlf
	Next 
	
	 GetCookieInformation = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETSERVERVARIABLESINFORMATION()
'--------------------------------------------------------------------------------------------------
Function GetServerVariablesInformation()
	
	sReturnValue = ""

	For each key in request.servervariables
		sReturnValue = sReturnValue & key & ":" & request.servervariables(key) & "<BR>" & vbcrlf
	Next 
	
	GetServerVariablesInformation = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETREQUESTFORMINFORMATION()
'--------------------------------------------------------------------------------------------------
Function GetRequestFormInformation()
	
	sReturnValue = ""

	For each key in request.form
		sReturnValue = sReturnValue & key & ":" & request.form(key) & "<BR>" & vbcrlf
	Next 
	
	GetRequestFormInformation = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETREQUESTQUERYSTRINGINFORMATION
'--------------------------------------------------------------------------------------------------
Function GetRequestQueryStringInformation()
	
	sReturnValue = ""

	For each key in request.querystring
		sReturnValue = sReturnValue & key & ":" & request.querystring(key) & "<BR>" & vbcrlf
	Next 
	
	GetRequestQueryStringInformation = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION DBSAFE( STRDB )
'--------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function


'--------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYERRORINFORMATION
'--------------------------------------------------------------------------------------------------
Sub subDisplayErrorInformation
	
		' IIS ASP ERROR OBJECT INFORMATION
		response.write "<br><b>IIS ASP Error Object Information:</b><div class=errorbox>" 
		response.write "<b>File: </b> " & sASPFile & "<BR>"
		response.write "<b>Line: </b> " & sASPLine & "<BR>"
		response.write "<b>Number: </b> " & sASPNumber & "<BR>"
		response.write "<b>Description: </b> " & sASPDescription & "<BR>"
		response.write "<b>Category: </b> " & sASPCategory & "<BR>"
		response.write "<b>Column: </b> " & sASPColumn & "<BR>"
		response.write "</div>"
		
		' Browser Information
		response.write "<br><b>Client User Browser Information</b><div class=errorbox>" & request.servervariables("HTTP_USER_AGENT") & "</div>"

		' APPLICATION COLLECTION
		response.write "<br><b>ASP Application Collection</b><div class=errorbox>" & GetApplicationInformation() & "</div>"
		
		' REQUEST FORM COLLECTION
		response.write "<br><b>Request Form Collection</b><div class=errorbox>" & GetRequestFormInformation() & "</div>"
		
		' QUERYSTRING COLLECTION
		response.write "<br><b>Querystring Collection</b><div class=errorbox>" & GetRequestQueryStringInformation() & "</div>"

		' SESSION COLLECTION
		response.write "<br><b>Session Collection</b><div class=errorbox>" & GetSessionInformation() & "</div>"

		' COOKIES COLLECTION
		response.write "<br><b>Cookies Collection</b><div class=errorbox>" & GetCookieInformation() & "</div>"

		' SERVER VARIABLES COLLECTION
		response.write "<br><b>Server Variable Collection</b><div class=errorbox>" & GetServerVariablesInformation() & "</div>"


End Sub 


%>
