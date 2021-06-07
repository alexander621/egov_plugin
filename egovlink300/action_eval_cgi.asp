<!-- #include file="includes/common.asp" //-->


<html>
<head>


<%If iorgid = 7 Then %>
	<title><%=sOrgName%></title>
<%Else%>
	<title>E-Gov Services <%=sOrgName%></title>
<%End If%>



<link rel="stylesheet" href="css/styles.css" type="text/css">
<link href="global.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" href="css/style_<%=iorgid%>.css" type="text/css">
<script language="Javascript" src="scripts/modules.js"></script>
<script language=javascript>
function openWin2(url, name) {
  popupWin = window.open(url, name,"resizable,width=500,height=450");
}
</script>
</head>


<!--#Include file="include_top.asp"-->





<!--BODY CONTENT-->
<p class=title>Thank you for using the <%=sOrgName%> Action Line.</P>


<%
' PARSE TO GET TITLE OF ACTION FORM
If instr(request("actionid"),"|") > 0 then
			arrForm = split(request("actionid"),"|")
			actionid = CLng(arrForm(0))
			actiontitle = arrForm(1)
			actiontitle = replace(actiontitle,"\"," > ")
			actiontitle = replace(actiontitle,"/"," > ")
Else
			actionid = CLng(request("actionid"))
			actiontitle = request("actiontitle")
End If


' MAKE SURE WE HAVE VALID REQUEST ID BEFORE PROCEEDING
If actionid = "" Then
	response.write "<font class=error>!There was an error processing this request. No action form found for this submission!</font>"
	response.End
End If

'DEFAULT EMAIL 
adminFromAddr = "webmaster@eclink.com"

' CONNECT TO DATABASE AND GET ADMIN EMAIL ADDRESSES FOR THIS ACTION FORM 
sSQLadmin = "SELECT assigned_userid,assigned_userid2,assigned_userid3 FROM egov_action_request_forms where action_form_id=" & actionid
Set oAdmin = Server.CreateObject("ADODB.Recordset")
oAdmin.Open sSQLadmin, Application("DSN"), 3, 1

If NOT oAdmin.EOF Then
		'** 1st ASSIGNED-TOP
		if oAdmin("assigned_userid") = "" or isNull(oAdmin("assigned_userid")) then
				'adminEmailAddr = "jstullenberger@eclink.com" ' NEED TO HAVE A DEFAULT INSTITUTION EMAIL ADDRESS
				'adminid = 162 ' DEFAULT INSTITUTION USERID
		else
				sSQLaddress = "SELECT email,UserId FROM users where UserId=" & oAdmin("assigned_userid") 
				Set oAddress = Server.CreateObject("ADODB.Recordset")
				oAddress.Open sSQLaddress, Application("DSN"), 3, 1
				
				If Not oAddress.EOF Then
					If iorgid = 18 Then
						' This handles Vandalia's inability to receive email from themselves
						adminFromAddr = "webmaster@eclink.com"
					Else 
						If adminFromAddr = "" Then 
							adminFromAddr = "webmaster@eclink.com"
						Else
							adminFromAddr = oAddress("email")   ' ASSIGNED ADMIN USER EMAIL
						End IF
					End If 
					adminEmailAddr = oAddress("email")   ' ASSIGNED ADMIN USER EMAIL
					adminid = oAddress("UserId")   ' ASSIGNED ADMIN USER ID
					oAddress.Close
				End If

				Set oAddress = Nothing
		end if

		'** 2nd ASSIGNED-TOP
		if oAdmin("assigned_userid2") = "" or isNull(oAdmin("assigned_userid2")) then
			''nothing
		else
				sSQLaddress = "SELECT email FROM users where UserId=" & oAdmin("assigned_userid2") 
				Set oAddress = Server.CreateObject("ADODB.Recordset")
				oAddress.Open sSQLaddress, Application("DSN"), 3, 1
			
				if oAddress.EOF=false then
						adminEmailAddr = adminEmailAddr & "," & oAddress("email")   ' ASSIGNED ADMIN USER EMAIL
						'adminid = oAddress("UserId")   ' ASSIGNED ADMIN USER ID
						oAddress.Close
				end if

				Set oAddress = Nothing
		end if

		'** 3rd ASSIGNED-TOP
		if oAdmin("assigned_userid3") = "" or isNull(oAdmin("assigned_userid3")) then
			''nothing
		else
				sSQLaddress = "SELECT email FROM users where UserId=" & oAdmin("assigned_userid3") 
				Set oAddress = Server.CreateObject("ADODB.Recordset")
				oAddress.Open sSQLaddress, Application("DSN"), 3, 1
			
				if oAddress.EOF=false then
						adminEmailAddr = adminEmailAddr & "," & oAddress("email")   ' ASSIGNED ADMIN USER EMAIL
						oAddress.Close
						'adminid = oAddress("UserId")   ' ASSIGNED ADMIN USER ID
				end if

				Set oAddress = Nothing
		end if

		oAdmin.Close
End If

Set oAdmin = Nothing


' GET QUESTIONS AND ENTERED VALUES INFORMATION
sQuestions = ""
For Each oField in Request.Form
	If Left(oField,10) = "fmquestion" Then
		sQuestionPrompt = "fmname" & replace(oField,"fmquestion","")
		sQuestions = sQuestions & "<P><b>" & request.form(sQuestionPrompt) & "</b><br>" & request.form(oField) & "</P>" & vbcrlf
	End If
Next
sQuestions2 = sQuestions


' INSERT FORM INFORMATION INTO DATABASE	
Set oNewActionRequest = Server.CreateObject("ADODB.Recordset")
oNewActionRequest.CursorLocation = 3
oNewActionRequest.Open "SELECT * FROM egov_actionline_requests", Application("DSN") , 3, 2
oNewActionRequest.AddNew


' GET USER INFORMATION
If sOrgRegistration Then
	If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" Then
		oNewActionRequest("userid") = request.cookies("userid")
	Else 
		oNewActionRequest("userid") = AddUserInformation()
	End If

Else

	oNewActionRequest("userid") = AddUserInformation()
	
End If


oNewActionRequest("assignedemployeeid") = adminid
oNewActionRequest("comment") = sQuestions
oNewActionRequest("category_id") = actionid
oNewActionRequest("category_title") = actiontitle
oNewActionRequest("orgid") = iorgid
oNewActionRequest("status") = "EVALFORM"



' TIMEZONE DIFFERENCE
datGMTDateTime = DateAdd("h",5,Now())
datOrgDateTime = DateAdd("h",iTimeOffset,datGMTDateTime)
datCurrentDate = Now()


oNewActionRequest("submit_date") = datCurrentDate
oNewActionRequest.Update
iTrackingNumber = oNewActionRequest("action_autoid")
oNewActionRequest.Close
Set oNewActionRequest = Nothing 



' GENERATE TRACKING NUMBER - (FORMULA IS SQL ROWID + HHMM)
lngTrackingNumber = iTrackingNumber & replace(FormatDateTime(datCurrentDate,4),":","")


' SEND EMAIL TO CITIZEN
	' BUILD MESSAGE 
			sMsg = sMsg & "This automated message was sent by the " & sOrgName & " E-Gov web site. Do not reply to this message.  Follow the instructions below or contact " & adminFromAddr & " for inquiries regarding this email." & vbcrlf 
			sMsg = sMsg & " " & vbcrlf 
			sMsg = sMsg & "Thank you for submitting information about " & sOrgName & " on " & datOrgDateTime & "." & vbcrlf 
			sMsg = sMsg & " " & vbcrlf 
		
			sMsg = sMsg & "To check the status, please link to the " & sOrgName & " web site at... " & sEgovWebsiteURL & "/action.asp" & vbcrlf 
			sMsg = sMsg & "Make sure that the entire URL appears in your browser's address field." & vbcrlf 
			sMsg = sMsg & " " & vbcrlf & vbcrlf 
			sMsg = sMsg & "Then enter your TRACKING NUMBER which is: " 
			sMsg = sMsg & lngTrackingNumber & " " & vbcrlf 
	 
			sMsg = sMsg & " " & vbcrlf 
			
			'sMsg = sMsg & "CATEGORY: " & UCASE(actiontitle)   & vbcrlf & vbcrlf
			sMsg = sMsg & "CATEGORY: " & actiontitle & vbcrlf & vbcrlf
			sMsg = sMsg & "SUGGESTION/ISSUE:  " 
			sMsg = sMsg & vbcrlf & vbcrlf  & fnPlainText(sQuestions) & vbcrlf & vbcrlf
		
			
			sMsg = sMsg & "We will evaluate this request and take action as appropriate." & vbcrlf 
			sMsg = sMsg & " " & vbcrlf 
			sMsg = sMsg & "Thank you for using the " & sOrgName & " messaging service.  We want to understand what you want and expect, make it easier for you to do business with us, and to respond as quickly as practical to your requests." & vbcrlf 
			sMsg = sMsg & " " & vbcrlf 


	' SEND MESSAGE
			'sendEmail "", request("cot_txtEmail"), "", SUBJECT, sMsgBody, "", "Y"

			If iorgid <> "7" Then
			' PLAIN ASCII TEXT MESSAGE
				lcl_subject = sOrgName & " E-GOV MSG - ACTION LINE REQUEST"
				lcl_HTMLBody = ""
				lcl_TXTBody = sMsg
			Else
			' HTML MESSAGE
				lcl_subject = "ECLINK HELPDESK - NEW HELPDESK TICKET"
				lcl_HTMLBody = BuildHTMLMessage(BuildHTMLBody())
				lcl_TXTBody = ""
			End If
			
		    
			' SEND EMAIL IF BOX WAS CHECKED
			If request("chkSendEmail") = "YES" Then
				'ErrorCode = objMail.Send
				lcl_validate_email = formatSendToEmail(request("cot_txtEmail"))

				if isValidEmail(lcl_validate_email) then
					sendEmail "", request("cot_txtEmail"), "", lcl_subject, lcl_HTMLBody, lcl_TXTBody, "Y"
				else
					ErrorCode = 1
				end if
				
				' ADD TO EMAIL QUEUE IF UNSUCCESSFUL 
				If ErrorCode <> 0 Then
					sMsg = Left(sMsg,5000)
					SendToAdd = request("cot_txtEmail")
					'fnPlaceEmailinQueue Application("SMTP_Server"),sOrgName & " E-GOV WEBSITE",adminFromAddr,SendToAdd,sOrgName & " E-GOV MSG - ACTION LINE REQUEST",1,sMsg,1,-1
				End If

			End If 
			
			Set objMail = Nothing
			
  		    
			If ErrorCode <> 0 Then
				' ADD LOGGING CODE HERE
				response.write "The request has been logged but there was an error sending an email notice to you.  You will not receive an email notice.<br /><br /><br />"
				'response.write ErrorCode & "<br>"
				'response.write Err.Number  & "<br>"
				'response.write Err.Description  & "<br>"
				bMailSent1 = False
			End If


' SEND EMAIL TO SITE ADMINISTRATOR
	' BUILD MESSAGE 
			sMsg2 = sMsg2 & "This automated message was sent by the " & sOrgName & " E-Gov web site. Do not reply to this message.  Contact " & adminFromAddr & " for inquiries regarding this email." & vbcrlf 
			sMsg2 = sMsg2 & " " & vbcrlf 
			sMsg2 = sMsg2 & "A " & sOrgName & " Action Line issue was submitted on " & datOrgDateTime & "." & vbcrlf 
			sMsg2 = sMsg2 & " " & vbcrlf 
			sMsg2 = sMsg2 & "ACTION LINE REQUEST DETAILS" & vbcrlf
			sMsg2 = sMsg2 & "DATE SUBMITTED: " & datOrgDateTime & vbcrlf
			sMsg2 = sMsg2 & "TRACKING NUMBER: " & lngTrackingNumber & vbcrlf
			sMsg2 = sMsg2 & "CATEGORY ID: " & actionid & vbcrlf
			sMsg2 = sMsg2 & "CATEGORY Title: " & actiontitle & vbcrlf
			sMsg2 = sMsg2 & "SUGGESTION/ISSUE: ..." 
			sMsg2 = sMsg2 & vbcrlf & vbcrlf  & fnPlainText(sQuestions) & vbcrlf & vbcrlf
			sMsg2 = sMsg2 & "ACTION LINE REQUESTER CONTACT INFORMATION" & vbcrlf
			sMsg2 = sMsg2 & "NAME: " & Request("cot_txtFirst_Name") & " " & Request("cot_txtLast_Name") & vbcrlf 
			sMsg2 = sMsg2 & "BUSINESS: " & Request("cot_txtBusiness_Name") & vbcrlf 
			sMsg2 = sMsg2 & "EMAIL: " & Request("cot_txtEmail") & vbcrlf 
			sMsg2 = sMsg2 & "PHONE: " & Request("cot_txtDaytime_Phone") & vbcrlf 
			sMsg2 = sMsg2 & "FAX: " & Request("cot_txtFax") & vbcrlf 
			sMsg2 = sMsg2 & "ADDRESS: " & Request("cot_txtStreet") & vbcrlf
			sMsg2 = sMsg2 & "" & Request("cot_txtCity") & " " & Request("cot_txtState_vSlash_Province") & " " 
			sMsg2 = sMsg2 & "" & Request("cot_txtZIP_vSlash_Postal_Code") & " " & Request("cot_txtCountry") & vbcrlf & vbcrlf



	' SEND MESSAGE
			
			If iorgid <> "7" Then
			' PLAIN ASCII TEXT MESSAGE
				lcl_subject = sOrgName & " E-GOV ACTION ITEM SUBMISSION"
				lcl_HTMLBody = ""
				lcl_TXTBody = sMsg2
			Else
			' HTML MESSAGE
				lcl_subject = "ECLINK HELPDESK - NEW HELPDESK TICKET"
				lcl_HTMLBody = BuildHTMLMessage(BuildAdminHTMLBody())
				lcl_TXTBody = ""
			End if


			lcl_validate_email = formatSendToEmail(adminEmailAddr)

			if isValidEmail(lcl_validate_email) then
				sendEmail "", adminEmailAddr, "", lcl_subject, lcl_HTMLBody, lcl_TXTBody, "Y"
			else
				ErrorCode = 1
			end if

			' ADD TO EMAIL QUEUE IF UNSUCCESSFUL 
			If ErrorCode <> 0 Then
				sMsg2 = Left(sMsg2,5000)
				'fnPlaceEmailinQueue Application("SMTP_Server"),sOrgName & " E-GOV WEBSITE",adminFromAddr,adminEmailAddr,sOrgName & " E-GOV ACTION ITEM SUBMISSION",1,sMsg2,1,-1
			End If

			Set objMail2 = Nothing

			If ErrorCode <> 0 Then
				response.write "The request has been logged but there was an error sending the email notice, so there may be a delay in processing your request.<br /><br />"
				'response.write ErrorCode & "<br>"
				'response.write Err.Number  & "<br>"
				'response.write Err.Description  & "<br>"
				bMailSent2 = False
			End If




' DISPLAY INFORMATION TO THE USER
response.write "<div style=""margin-left:20px; "" class=box_header2>Action Line Request Submitted - " & datOrgDateTime &  "</div>"
response.write "<div style=""margin-left:20px; "" class=groupsmall>"
Response.write "<p>A request has been submitted under the subject of <b><i>" & actiontitle & "</i></b>.  We'll evaluate your request and take action as appropriate.</p>" 
Response.write "<p>Your tracking number is <b>" & lngTrackingNumber & "</b>. Please record this number for your records."
response.write "<p>You can check the status of the request by visiting the <b>Action Line Request</b> main page. Simply enter the above tracking number to review the status of the request. Please allow at least 24 hours for a response to your request."
response.write "</div></div>"
%>


<!--SPACING CODE-->
<p><bR>&nbsp;<bR>&nbsp;</p>
<!--SPACING CODE-->





<!--#Include file="include_bottom.asp"-->  


<%

'------------------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'------------------------------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------------------------------
' FUNCTION ADDUSERINFORMATION()
'------------------------------------------------------------------------------------------------------------
Function AddUserInformation()

	' INSERT FORM INFORMATION INTO DATABASE	
	iReturnValue = 0
	
	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.CursorLocation = 3
	oUser.Open "SELECT * FROM egov_users WHERE 1=2", Application("DSN") , 3, 2
	oUser.AddNew
	oUser("userfname") = dbsafe(request.form("cot_txtFirst_Name"))
	oUser("userlname") = dbsafe(request.form("cot_txtLast_Name"))
	oUser("useremail") = dbsafe(request.form("cot_txtEmail"))
	oUser("userbusinessname") = request.form("cot_txtBusiness_Name")
	oUser("userhomephone") = request.form("cot_txtDaytime_Phone")
	oUser("userfax") = request.form("cot_txtFax")
	oUser("useraddress") = request.form("cot_txtStreet")
	oUser("usercity") = request.form("cot_txtCity")
	oUser("userstate") = request.form("cot_txtState_vSlash_Province")
	oUser("userzip") = request.form("cot_txtZIP_vSlash_Postal_Code")
	oUser("usercountry") =request.form("cot_txtCountry")
	oUser.Update
	iReturnValue = oUser("userid")
	oUser.Close

	Set oUser = Nothing

	AddUserInformation = iReturnValue

End Function


Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function

'------------------------------------------------------------------------------------------------------------------------------
' FUNCTION FNPLACEEMAILINQUEUE(SHOST,SFROMNAME,SFROMEMAIL,SSENDEMAIL,SSUBJECT,IBODYFORMAT,SBODYMESSAGE,IPRIORITY,IERRORCODE)
'------------------------------------------------------------------------------------------------------------------------------

Function fnPlaceEmailinQueue(sHost,sFromName,sFromEmail,sSendEmail,sSubject,iBodyFormat,sBodyMessage,iPriority,iErrorCode)
  
  Set oCmd = Server.CreateObject("ADODB.Command")
  
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "AddEmailtoFailoverQueue"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("Host", adVarChar , adParamInput, 50, sHost)
    .Parameters.Append oCmd.CreateParameter("FromName", adVarChar, adParamInput, 50, sFromName)
	.Parameters.Append oCmd.CreateParameter("FromEmail", adVarChar, adParamInput, 255, sFromEmail)
	.Parameters.Append oCmd.CreateParameter("SendEmail", adVarChar, adParamInput, 255, sSendEmail)
	.Parameters.Append oCmd.CreateParameter("Subject", adVarChar, adParamInput, 1024, sSubject)
	.Parameters.Append oCmd.CreateParameter("BodyFormat", adInteger, adParamInput, 4, iBodyFormat)
	.Parameters.Append oCmd.CreateParameter("BodyMessage", adVarChar, adParamInput, 5000, sBodyMessage)
	.Parameters.Append oCmd.CreateParameter("Priority", adInteger, adParamInput, 4, iPriority)
	.Parameters.Append oCmd.CreateParameter("ErrorCode", adVarChar, adParamInput, 10, iErrorCode)
    .Execute
  End With
  
  Set oCmd = Nothing

End Function


Function fnPlainText(sValue)
	'sValue = UCASE(sValue)  Removed per Peter on 3/14/2006 - Steve Loar
	sValue = replace(sValue,"<B>","")
	sValue = replace(sValue,"</B>","")
	sValue = replace(sValue,"<P>","")
	sValue = replace(sValue,"</P>",vbcrlf & vbcrlf)
	sValue = replace(sValue,"<BR>",vbcrlf)
	sValue = replace(sValue,"</BR>",vbcrlf)
	sValue = replace(sValue,"<b>","")
	sValue = replace(sValue,"</b>","")
	sValue = replace(sValue,"<p>","")
	sValue = replace(sValue,"</p>",vbcrlf & vbcrlf)
	sValue = replace(sValue,"<br>",vbcrlf)
	sValue = replace(sValue,"</br>",vbcrlf)
	sValue = replace(sValue,"<br />",vbcrlf)

	fnPlainText = sValue

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION BUILDHTMLMESSAGE(SBODY)
'--------------------------------------------------------------------------------------------------
Function BuildHTMLMessage(sBody)
	' BUILD MESSAGE
	Dim sLayout
	sLayout = sLayout & "<HTML>"
	sLayout = sLayout & "<HEAD>"
	sLayout = sLayout & "<STYLE> td {font-family: arial,tahoma; font-size: 12px; color: #000000;} </STYLE>"
	sLayout = sLayout & "</HEAD>"
	sLayout = sLayout & "<body bgcolor=#E0E0E0>"
	sLayout = sLayout & "<FONT face=""helvetica, arial"">"
	sLayout = sLayout & "<P style=""MARGIN: 0px""></P>"
	sLayout = sLayout & "<TABLE borderColor=#4A9E9F bgcolor=#ffffff cellSpacing=0 cellPadding=5 width=""95%"" align=center border=2 valign=top><tr><td>" & sBody
	sLayout = sLayout & "<CENTER>"
	sLayout = sLayout & "<br>"
	sLayout = sLayout & "<HR color=black SIZE=1 width=""95%"">"
	sLayout = sLayout & "<FONT size=-2>Copyright 2005. <i>electronic commerce</i> link, inc. dba <i>ec</I> link.</font></CENTER></TD></TR></TABLE></FONT>"
	sLayout = sLayout & "</BODY>"
	sLayout = sLayout & "</HTML>"

	BuildHTMLMessage = sLayout

End Function


Function BuildHTMLBody()

			sMsgNew = sMsgNew & "This automated message was sent by the ECLINK HELPDESK web site. Do not reply to this message.  Please follow the instructions below or contact <b>" & adminFromAddr & "</b> for inquiries regarding this email.<br>" 
			sMsgNew = sMsgNew & "<br> " & vbcrlf 
			sMsgNew = sMsgNew & "You created a helpdesk ticket on " & datOrgDateTime & ".<br>" & vbcrlf 
			sMsgNew = sMsgNew & "<br>" & vbcrlf 
		
			'sMsgNew = sMsgNew & "To check the status, please follow the link below:<br><br>http://www.egovlink.com/" & sorgVirtualSiteName & "/action.asp<br><br>" & vbcrlf 
			'sMsgNew = sMsgNew & "Make sure that the entire URL appears in your browser's address field.<br>" & vbcrlf 
			'sMsgNew = sMsgNew & "<br>" & vbcrlf & vbcrlf 
			sMsgNew = sMsgNew & "Tracking Number: <b> " 
			sMsgNew = sMsgNew & lngTrackingNumber & "</b><br> " & vbcrlf 
	 
			sMsgNew = sMsgNew & "<br><br> " & vbcrlf 
			
			sMsgNew = sMsgNew & "<b>HELP DESK FORM:</b> " & UCASE(actiontitle)  & "<BR><BR>" & vbcrlf & vbcrlf
			sMsgNew = sMsgNew & "<b>HELP DESK TICKET DETAILS:</b><br><br>"
			sMsgNew = sMsgNew & vbcrlf & vbcrlf  & sQuestions2 & vbcrlf & vbcrlf
		
			
			sMsgNew = sMsgNew & "<br><br><br>We will evaluate this issue and take action as appropriate.<br>" & vbcrlf 
			sMsgNew = sMsgNew & "<br> " & vbcrlf 
			sMsgNew = sMsgNew & "Thank you for using our help desk web site to better serve your business needs.<br>" & vbcrlf 
			sMsgNew = sMsgNew & "<br> " & vbcrlf 

			BuildHTMLBody = sMsgNew

End Function


Function BuildAdminHTMLBody()

			sMsgAdmin = sMsgAdmin & "This automated message was sent by the ECLINK HELPDESK web site. Do not reply to this message.  Follow the instructions below or contact <b>" & adminFromAddr & "</b> for inquiries regarding this email.<br>" 
			sMsgAdmin = sMsgAdmin & "<br> " & vbcrlf 
			sMsgAdmin = sMsgAdmin & "To follow-up with this help desk ticket please follow the link below:<br><br>http://www.egovlink.com/" & sorgVirtualSiteName & "/admin<br><br>" & vbcrlf 
			sMsgAdmin = sMsgAdmin & "<br><b>DATE SUBMITTED:</b> " & datOrgDateTime & vbcrlf
			sMsgAdmin = sMsgAdmin & "<br><b>TRACKING NUMBER:</b> " & lngTrackingNumber & vbcrlf
			sMsgAdmin = sMsgAdmin & "<br><b>HELP DESK FORM:</b> " & actiontitle & vbcrlf
			sMsgAdmin = sMsgAdmin & "<br><b>HELP DESK TICKET DETAILS:</b><br><br>" 
			sMsgAdmin = sMsgAdmin & vbcrlf & vbcrlf  & sQuestions2 & vbcrlf & vbcrlf
			sMsgAdmin = sMsgAdmin & "<br><br><b>TICKET SUBMITTER CONTACT INFORMATION</b>" & vbcrlf
			sMsgAdmin = sMsgAdmin & "<br>NAME: " & Request("cot_txtFirst_Name") & " " & Request("cot_txtLast_Name") & vbcrlf 
			sMsgAdmin = sMsgAdmin & "<br>BUSINESS: " & Request("cot_txtBusiness_Name") & vbcrlf 
			sMsgAdmin = sMsgAdmin & "<br>EMAIL: " & Request("cot_txtEmail") & vbcrlf 
			sMsgAdmin = sMsgAdmin & "<br>PHONE: " & Request("cot_txtDaytime_Phone") & vbcrlf 
			sMsgAdmin = sMsgAdmin & "<br>FAX: " & Request("cot_txtFax") & vbcrlf 
			sMsgAdmin = sMsgAdmin & "<br>ADDRESS: " & Request("cot_txtStreet") & vbcrlf
			sMsgAdmin = sMsgAdmin & "<br>" & Request("cot_txtCity") & " " & Request("cot_txtState_vSlash_Province") & " " 
			sMsgAdmin = sMsgAdmin & "<br>" & Request("cot_txtZIP_vSlash_Postal_Code") & " " & Request("cot_txtCountry") & vbcrlf & vbcrlf

			BuildAdminHTMLBody = sMsgAdmin

End Function
%>
