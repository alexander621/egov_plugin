<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<%@ language="VBScript" %>
<!--#include file="../egovlink300_global/includes/inc_email.asp"-->
<%

  Const lngMaxFormBytes = 200

  Dim objASPError, blnErrorWritten, strServername, strServerIP, strRemoteIP
  Dim strMethod, lngPos, datNow, strQueryString, strURL,sErrorInfo
  Dim sBrowserInfo,sMicrosoft,sTimeInfo
  Dim sAspCategory


  If Response.Buffer Then
    Response.Clear
    Response.Status = "500 Internal Server Error"
    Response.ContentType = "text/html"
    Response.Expires = 0
  End If

  Set objASPError = Server.GetLastError


  ' SET ERROR VALUES
  sAspCategory = objASPError.Category
  sAspColumn = objASPError.Column
  sAspDescription = objASPError.Description
  sAspFile = objASPError.File
  sAspLine = objASPError.Line
  sAspNumber = objASPError.Number
  sAspSource = objASPError.Source

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
		a:link			{font:8pt/11pt verdana; color:FF0000}
		a:visited		{font:8pt/11pt verdana; color:#4e4e4e}
		div.errorbox {text-align:left; width:90%; border-style:solid;border-color:#000000;border-width:1px;font-family: arial,tahoma; font-size: 12px; color:#000000;padding:5px;background-color:#e0e0e0;}
	</style>

	<META NAME="ROBOTS" CONTENT="NOINDEX">

	<title>APPLICATION ERROR</title>

	<META HTTP-EQUIV="Content-Type" Content="text-html; charset=Windows-1252">
</head>


<body bgcolor="ffffff">


<%
' RECORD ERRORS IN DATABASE
iDBID = fnRecordError()


' GET ERROR INFORMATION
sErrorMsg = "GET ERROR INFORMATION"


' BUILD LINK TO SPECIFIC ERROR FOR EMAIL
'sHyperlink = vbcrlf & vbcrlf & "For more information regarding this error message click the link below or copy/paste link into your web browser:" & vbcrlf & "http://dev.eclink.com/application_errors/view_error_log.asp?errorid=" & iDBID & vbcrlf 
sHyperlink = vbcrlf & vbcrlf & "For more information regarding this error message click the link below or copy/paste link into your web browser:" & vbcrlf & "http://www.egovlink.com/eclink/admin/errors/view_error_log.asp?errorid=" & iDBID & vbcrlf 
sMSG = BuildMessageBody(sErrorMessage & sHyperlink)

'response.write request.servervariables("REMOTE_ADDR")
If request.servervariables("REMOTE_ADDR") <> "10.4.24.100" and request.servervariables("REMOTE_ADDR") <> "24.106.89.6" And  request.servervariables("REMOTE_ADDR") <> "184.180.44.105" And Left(request.servervariables("REMOTE_ADDR"),7) <> "10.0.8."  And Left(request.servervariables("REMOTE_ADDR"),8) <> "10.0.48." and request.servervariables("REMOTE_ADDR") <> "74.87.250.138" Then
	' SEND EMAIL TO TECH SUPPORT
	'iErrorCode = SendEmail("smtp1.eclink.com","devsupport@eclink.com","EC LINK WEB SERVER","ec link ASP Script Error - " & sAspFile,"devsupport@eclink.com", sMSG)
	SendEmail "EC LINK WEB SERVER <devsupport@eclink.com>", "devsupport@eclink.com","support@eclink.com","ec link ASP Script Error - " & sAspFile, "", sMSG, "N"
	iErrorCode = "1"
Else
	' DONT SEND EMAIL JUST SHOW ERROR TO ECLINK EMPLOYEE
	'SendEmail "EC LINK WEB SERVER <devsupport@eclink.com>", "devsupport@eclink.com","support@eclink.com","ec link ASP Script Error - " & sAspFile, "", sMSG, "N"
	iErrorCode = "9999"
End If


' DISPLAY THAT AN ERROR OCCURRED TO THE USER
If iErrorCode <> 0 Then
	' ASK USER TO SEND ERROR DETAILS TO USER
	response.write "<br><font style=""font-family: arial,tahoma;font-size: 12px;"" ><div style=""text-align:left; width:90%; border-style:solid;border-color:#c0c0c0;border-width:1px;font-family: arial,tahoma; font-size: 14px; color:yellow;padding:5px;background-color:#FF0000;"" >" & vbcrlf
	response.write "We're sorry, but the server encountered an error processing your request.  If you continue to receive this message, please follow up with your account representative. You will need to provide them with the error details listed in the box below.  Copy and paste the text into an email or a word document.</div><br>" & vbcrlf
	response.write "<br><b>Error Details:</b><div style=""text-align:left; width:90%;"" >" 
	If iErrorCode = "9999" Then
		' INTERNAL ERROR DETAILS
		Call subDisplayErrorInformation()
	Else
		' PUBLIC DETAILS
		response.write "<PRE>" & sErrorMsg & "</PRE>" & vbcrlf
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
Function BuildMessageBody(sText)
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


''--------------------------------------------------------------------------------------------------
'' FUNCTION SENDEMAIL(SSMTP,SFROMEMAIL,SFROMNAME,SSUBJECT,STOEMAIL,SBODY)
''--------------------------------------------------------------------------------------------------
'Function SendEmail( sSMTP, sFromEmail, sFromName, sSubject, sToEmail, sBody )
'
  '' CREATE EMAIL OBJECT AND SEND EMAIL       
  'Dim objMailer, Body
  'iReturnCode = 0
  'Set objMailer = Server.CreateObject("ecMail.SMTP")
  ''Set objMailer = Server.CreateObject("smtp1.eclink.com")
  '
   '' Assign some required properties
  'objMailer.Host       = sSMTP
  'objMailer.BodyFormat = 1
  'objMailer.FromName   = sFromName
  'objMailer.From       = sFromEmail
  'objMailer.Subject    = sSubject
  'objMailer.SendTo     = sToEmail
  'objMailer.CC         = "egovsupport@eclink.com"
  'objMailer.Priority   = 1 'PRIORITY_HIGH
  'objMailer.Body       = sBody
'
  'iReturnCode = objMailer.Send
'
  'Set objMailer = Nothing
  '
  '' RETURN ERROR CODE
  'SendEmail = iReturnCode
'
'
'End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION FNRECORDERROR()
'--------------------------------------------------------------------------------------------------
Function fnRecordError()
	Dim iReturnValue, sSql

'	iReturnValue = 0

	sSql = "INSERT INTO errorlog ( category, [column], description, [file], line, number, source, "
	sSql = sSql & " sessioncollection, applicationcollection, cookiescollection, browserinformation, "
	sSql = sSql & " servervariablescollection, webappid, requestformcollection, requestquerystringcollection ) VALUES ( '"
	sSql = sSql & DBsafe(sAspCategory) & "','" & DBsafe(sAspColumn) & "','" & DBsafe(sAspdescription) & "','"
	sSql = sSql & DBsafe(sAspfile) & "','" & DBsafe(sAspline) & "','" & DBsafe(sAspNumber) & "','" & DBsafe(sAspSource) & "','"
	sSql = sSql & DBsafe(GetSessionInformation()) & "','" & DBsafe(GetApplicationInformation()) & "','"
	sSql = sSql & DBsafe(GetCookieInformation()) & "','" & DBsafe(request.servervariables("HTTP_USER_AGENT")) & "','"
	sSql = sSql & DBsafe(GetServerVariablesInformation()) &"','" & Application("AppErrorID") & "','" 
	sSql = sSql & DBsafe(GetRequestformInformation()) &"','" & DBsafe(GetRequestQueryStringInformation()) &"')"
'	response.write sSql & "<br />"

	fnRecordError = RunIdentityInsert( sSql )

'	Set oRecordError = Server.CreateObject("ADODB.Recordset")
'	oRecordError.Open sSql, Application("ErrorDSN"), 3, 3
'	iReturnValue = oRecordError("ROWID")
'	Set oRecordError = Nothing

'	fnRecordError = iReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION GETSESSIONINFORMATION()
'--------------------------------------------------------------------------------------------------
Function GetSessionInformation()
	on error resume next
	sReturnValue = ""
	For each key in session.contents
		sReturnValue = sReturnValue & key & ":" & session(key) & "<br />" & vbcrlf
	Next
	
	on error goto 0
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
	Dim sReturnValue, key
	
	sReturnValue = ""

	on error resume next
	For each key in request.Form
		If key <> "accountnumber" And key <> "cvv2" Then 
			sReturnValue = sReturnValue & key & ":" & request.form(key) & "<br />" & vbcrlf
		End If 
	Next 
	on error goto 0
	
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
		response.write "<br><b>IIS ASP Error Object Information:</b><div class=errorbox>"  & vbcrlf
		response.write "<b>File: </b> " & sASPFile & "<BR>" & vbcrlf
		response.write "<b>Line: </b> " & sASPLine & "<BR>" & vbcrlf
		response.write "<b>Number: </b> " & sASPNumber & "<BR>" & vbcrlf
		response.write "<b>Description: </b> " & sASPDescription & "<BR>" & vbcrlf
		response.write "<b>Category: </b> " & sASPCategory & "<BR>" & vbcrlf
		response.write "<b>Column: </b> " & sASPColumn & "<BR>" & vbcrlf
		response.write "</div>" & vbcrlf
		
		' Browser Information
		response.write "<br><b>Client User Browser Information</b><div class=errorbox>" & request.servervariables("HTTP_USER_AGENT") & "</div>" & vbcrlf

		' APPLICATION COLLECTION
		response.write "<br><b>ASP Application Collection</b><div class=errorbox>" & GetApplicationInformation() & "</div>" & vbcrlf
		
		' REQUEST FORM COLLECTION
		response.write "<br><b>Request Form Collection</b><div class=errorbox>" & GetRequestFormInformation() & "</div>" & vbcrlf
		
		' QUERYSTRING COLLECTION
		response.write "<br><b>Querystring Collection</b><div class=errorbox>" & GetRequestQueryStringInformation() & "</div>" & vbcrlf

		' SESSION COLLECTION
		response.write "<br><b>Session Collection</b><div class=errorbox>" & GetSessionInformation() & "</div>" & vbcrlf

		' COOKIES COLLECTION
		response.write "<br><b>Cookies Collection</b><div class=errorbox>" & GetCookieInformation() & "</div>" & vbcrlf

		' SERVER VARIABLES COLLECTION
		response.write "<br><b>Server Variable Collection</b><div class=errorbox>" & GetServerVariablesInformation() & "</div>" & vbcrlf


End Sub 


'-------------------------------------------------------------------------------------------------
' Function RunIdentityInsert( sInsertStatement )
'-------------------------------------------------------------------------------------------------
Function RunIdentityInsert( ByVal sInsertStatement )
	Dim sSql, iReturnValue, oRs

	iReturnValue = 0

'	response.write "<p>" & sInsertStatement & "</p><br /><br />"
'	response.flush

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSql = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("ErrorDSN"), 3, 3
	iReturnValue = oRs("ROWID")

	oRs.Close
	Set oRs = Nothing

	RunIdentityInsert = iReturnValue

End Function



%>

