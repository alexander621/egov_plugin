<%

Sub sendEmail( ByVal iFromEmail, ByVal iSendToEmail, ByVal iCC, ByVal iSubject, ByVal iEmailHTMLBody, ByVal iEmailTEXTBody, ByVal iHighImportance )

	set oCdoMail = Server.CreateObject("CDO.Message")
	set oCdoConf = Server.CreateObject("CDO.Configuration")

	strSMTPServer = ""

'Set up the FROM EMAIL address
 if iFromEmail = "" then
    'iFromEmail = sOrgName & " (E-Gov Website) <webmaster@eclink.com>"
    'iFromEmail = sOrgName & " (E-Gov Website) <noreplies@eclink.com>"
    iFromEmail = sOrgName & " (E-Gov Website) <noreplies@egovlink.com>"
 end if

	With oCdoConf
		.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")  = 2
		'.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = 25
		'Community City, Sharonville, Park City, Piqua, Training City, E-Gov Support, Sycamore Township, Warrington, Bethel Park, Wilmington
		'Rye, Antioch, Rockaway, Liberty Township, Wyoming, Mariemont, Menlo Park, New Prov, amityville, Kirtland, Somers, Antioch, Sturgis
		'sugar grove (216), Harlingen (209) winterhaven (152), plymouth (174), maricopa (159), clarksburg (185), lincolnwood (40), payson (85)
		if instr(iFromEmail,"@egovlink.com") > 0 or instr(iFromEmail,"@eclink.com") > 0 _
			or instr(iFromEmail, "klewis@cityofclarksburgwv.com") > 0 _
			or session("orgid") = "5" or session("orgid") = "196" or session("orgid") = "153" or session("orgid") = "175" _
			or session("orgid") = "76" or session("orgid") = "147" or session("orgid") = "200" or session("orgid") = "113" _
			or session("orgid") = "125" or session("orgid") = "27" or session("orgid") = "103" or session("orgid") = "141" _
			or session("orgid") = "228" or session("orgid") = "169" or session("orgid") = "148" or session("orgid") = "37" _
			or session("orgid") = "60" or session("orgid") = "167" or session("orgid") = "41" or session("orgid") = "226" _
			or session("orgid") = "129" or session("orgid") = "169" or session("orgid") = "1" or session("orgid") = "216" _
			or session("orgid") = "209" or session("orgid") = "152" or session("orgid") = "174" or session("orgid") = "159" _
			or session("orgid") = "40" or session("orgid") = "85" then
			.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("SES_Server")
			.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'Can be 0 for No Authentication, 1 for basic authentication or 2 for NTLM (check with your host)
			.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1
			.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = Application("SES_UserName") 'The username log in credentials for the email account sending this email, not needed if authentication is set to 0
			.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Application("SES_Password") 'The password log in credentials for the email account sending this email, not needed if authentication is set to 0

			strSMTPServer = Application("SES_Server")

		else
			.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Application("SMTP_Server")

			strSMTPServer = Application("SMTP_Server")
		end if

  if iHighImportance = "Y" then
     .Fields.Item("urn:schemas:mailheader:X-MSMail-Priority") = "High"  'For Outlook 2003
     .Fields.Item("urn:schemas:mailheader:X-Priority")        = 2       'For Outlook 2003 also
     .Fields.Item("urn:schemas:httpmail:importance")          = 2       'For Outlook Express
  end if

		.Fields.Update
	End With

'Set up the FROM EMAIL address
 if iFromEmail = "" then
    'iFromEmail = sOrgName & " (E-Gov Website) <webmaster@eclink.com>"
    'iFromEmail = sOrgName & " (E-Gov Website) <noreplies@eclink.com>"
    iFromEmail = sOrgName & " (E-Gov Website) <noreplies@egovlink.com>"
 end if


 'OVERRIDE FOR AMAZON EMAIL (Amazon SES)
 'iFromEmail = "DO NOT REPLY " & sOrgName & " (E-Gov Website) <noreply_" & GetVDName & "@egovlink.com>"

'Format line length for email.
 if iEmailHTMLBody <> "" then

    lcl_email_body = formatToFitEmailLineLength(iEmailHTMLBody)
    'lcl_email_body = iEmailHTMLBody
 end if

'Setup the SENDTO and CC
 lcl_emailaddress_sendto = Replace(formatSendToEmail(iSendToEmail), " ","")

 if isValidEmail(lcl_emailaddress_sendto) then
    lcl_email_sendto    = lcl_emailaddress_sendto
    lcl_sendEmailToUser = "Y"
 end if

 if iCC <> "" then
    lcl_emailaddress_cc = formatSendToEmail(iCC)

    if isValidEmail(lcl_emailaddress_cc) then
       lcl_email_cc      = lcl_emailaddress_cc
       lcl_sendEmailToCC = "Y"
    end if
 end if

'Send Admin Email
 if lcl_sendEmailToUser = "Y" OR lcl_sendEmailToCC = "Y" then
   	With oCdoMail
    	 Set .Configuration = oCdoConf
   		 .From = iFromEmail

      if lcl_sendEmailToUser = "Y" then
       		.To = lcl_email_sendto
      end if

      if lcl_sendEmailToCC = "Y" then
         .Cc = lcl_email_cc
      end if

      'if session("orgid") = "167" then
         '.Bcc = "tfoster@eclink.com"
      'end if

      .Subject = iSubject

      if iEmailHTMLBody <> "" then
        .HTMLBody = lcl_email_body
      end if

     'Check for a text email body
      if iEmailTEXTBody <> "" then
         lcl_email_textbody = formatToFitEmailLineLength(iEmailTEXTBody)
         lcl_email_textbody = clearHTMLTags(lcl_email_textbody)
      else
         lcl_email_textbody = clearHTMLTags(lcl_email_body)
      end if

	  .TextBody = lcl_email_textbody

     'Remove the name on the email if it exists so that we can validate the email itself.
      lcl_emailaddress = formatSendToEmail(iSendToEmail)
	  session("sendtoaddress") = lcl_emailaddress

      if isValidEmail(lcl_emailaddress) then
		on error resume next
     	.Send
		on error goto 0
      end if

	'if lcl_email_sendto = "tfoster@eclink.com" then
      		LogSMTPEmail iFromEmail, lcl_email_sendto, lcl_email_cc, iSubject, strSMTPServer
	'end if

	   End With
 end if

	set oCdoMail = nothing
	set oCdoConf = nothing
end sub

sub LogSMTPEmail(strFrom, strTo, strCC, strSubject, strServer)

	strSQL = "INSERT INTO emaillog (fromaddress,toaddress,ccaddress,subject,smtpserver) VALUES('" & email_DBsafe(strFrom) & "','" & email_DBsafe(strTo) & "','" & email_DBsafe(strCC) & "','" & email_DBsafe(strSubject) & "','" & email_DBsafe(strServer) & "')"

	Set oCmdLog = Server.CreateObject("ADODB.Command")
	oCmdLog.ActiveConnection = Application("DSN")
	oCmdLog.CommandText = strSQL
	oCmdLog.Execute
	Set oCmdLog = Nothing

end sub

function formatToFitEmailLineLength(p_msg)
 'lcl_text_length = 800 characters.  The email standard length per line is 1000 characters.
 'Since we are breaking up the lines anyway, we will just inspect the lines at 800 characters for a "vbcrlf" to be 
 '  safe instead of pushing the limit.
  lcl_text_length = 800
  lcl_return      = p_msg
  lcl_msg         = p_msg
  lcl_newMsg      = ""

  if lcl_msg <> "" then
     if len(lcl_msg) > lcl_text_length then
        lcl_msg_len       = len(lcl_msg)
        lcl_cycles        = fix(lcl_msg_len/lcl_text_length)
        lcl_remaining     = lcl_msg_len - (lcl_text_length*lcl_cycles)
        lcl_msg_remaining = RIGHT(lcl_msg,lcl_remaining)

        i = 0

        do until i = lcl_cycles
           i = i + 1

           lcl_test_str = LEFT(lcl_msg,lcl_text_length)

           if instr(lcl_test_str,vbcrlf) < 1 then
              if instr(lcl_test_str," ") > 0 then
                 lcl_test_str = replace(lcl_test_str," ",vbcrlf,1)
              else
                 lcl_test_str = LEFT(lcl_test_str,(lcl_text_length/2)) & vbcrlf & RIGHT(lcl_test_str,(lcl_text_length/2))
              end if

           end if

           lcl_msg    = RIGHT(lcl_msg,len(lcl_msg)-lcl_text_length)
           lcl_newMsg = lcl_newMsg & lcl_test_str

        loop

        if lcl_remaining > 0 then
           lcl_newMsg = lcl_newMsg & lcl_msg_remaining
        end if

     else
        lcl_newMsg = lcl_msg
     end if

     lcl_return = lcl_newMsg

  end if

  formatToFitEmailLineLength = lcl_return

end function

function formatSendToEmail(iEmail)

  lcl_return = ""

 'Remove the name from the email address
  if iEmail <> "" then
     lcl_return = iEmail

     if instr(lcl_return,"<") > 0 then
        lcl_return = mid(lcl_return,instr(lcl_return,"<"))
        lcl_return = replace(lcl_return,"<","")
        lcl_return = replace(lcl_return,">","")
     end if

     lcl_return = trim(lcl_return)
  end if

  formatSendToEmail = lcl_return

end function

'------------------------------------------------------------------------------
' boolean isValidEmail( sEmail )
'------------------------------------------------------------------------------
function isValidEmail( ByVal sEmail )
	Dim bIsValid, rRegEx

 lcl_return = False

 if sEmail <> "" then
   	bIsValid   = False
   	Set rRegEx = New RegExp

   	rRegEx.IgnoreCase = False

   	'regEx.Pattern = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
    'rRegEx.Pattern = "^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*\.(\w{2}|(com|net|org|edu|mil|gov|biz|us|info))$"
	rRegEx.Pattern = "^.+@.+\..+$"
	rRegEx.IgnoreCase = True

    sEmail = lcase(sEmail)
    sEmail = ltrim(sEmail)
    sEmail = rtrim(sEmail)

    bIsValid       = rRegEx.Test(sEmail)

   	Set rRegEx = Nothing 
    lcl_return = bIsValid
 end if

 	'CHECK TO SEE IF THIS IS ON THE SUPPRESSION LIST
 	lcl_return = isNotSuppressed(sEmail, lcl_return)

	isValidEmail = lcl_return

end function


Function isNotSuppressed( strEmail, blnReturn )
	'CHECK TO SEE IF THIS IS ON THE SUPPRESSION LIST
	sSQL = "SELECT TOP 1 emailsuppressionid FROM emailsuppressionlist WHERE emailaddress = '" & email_DBsafe(sEmail) & "'"
	Set oRsSuppress = Server.CreateObject("ADODB.RecordSet")
	oRsSuppress.Open sSQL, Application("DSN"), 3, 1
	if not oRsSuppress.EOF then blnReturn = false
	oRsSuppress.Close
	Set oRsSuppress = Nothing

	isNotSuppressed = blnReturn
End Function

'------------------------------------------------------------------------------------------------------------
' string clearHTMLTags( sInString )
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

'------------------------------------------------------------------------------------------------------------
' string GetVDName()
'------------------------------------------------------------------------------------------------------------
Function GetVDName()
	sReturnValue = ""
	
	strURL = Request.ServerVariables("SCRIPT_NAME")
	strURL = Split(strURL, "/", -1, 1) 
	sReturnValue = "/" & strURL(1) 

	GetVDName = replace(sReturnValue,"/","")

End Function

Function email_DBsafe( ByVal strDB )
	Dim sNewString

	If Not VarType( strDB ) = vbString Then 
		sNewString = strDB
	Else 
		sNewString = Replace( strDB, "'", "''" )
		sNewString = Replace( sNewString, "<", "&lt;" )
	End If 

	email_DBsafe = sNewString
End Function

sub ShowEmailWarning
if  session("orgid") = "5" or session("orgid") = "196" or session("orgid") = "153" or session("orgid") = "175" _
	or session("orgid") = "76" or session("orgid") = "147" or session("orgid") = "200" or session("orgid") = "113" _
	or session("orgid") = "125" or session("orgid") = "27" or session("orgid") = "103" or session("orgid") = "141" _
	or session("orgid") = "228" or session("orgid") = "169" or session("orgid") = "148" or session("orgid") = "37" _
	or session("orgid") = "60" or session("orgid") = "167" or session("orgid") = "41" or session("orgid") = "226" _
	or session("orgid") = "129" or session("orgid") = "169" or session("orgid") = "1" or session("orgid") = "216" _
	or session("orgid") = "209" or session("orgid") = "152" or session("orgid") = "174" or session("orgid") = "159" _
	or session("orgid") = "185" or session("orgid") = "40" or session("orgid") = "85" then
  else%>
	  <fieldset>
<h3>Warning:  You are using an Email method that has been replaced.</h3><br />
Please contact E-Gov Link Support, to help us upgrade you to our new and improved Subscriptions email system.  The new system has better delivery results because your emails have less likelihood of getting flagged as spam.  <br />
<br />
We need your help in upgrading the software, because a small change is necessary at your end (to your DNS entry).<br />
<br />
Please contact us at <a href="mailto:support@egovlink.com">support@egovlink.com</a> to schedule this simple but important upgrade.<br />
<br />
<b>The Email method that you are using will be taken out of service, and will not work after 12/31/2020.</b><br />
	  </fieldset>

	  <script>
	  	function smtpnag()
		{
			var code = makeid(5);
			if (prompt("We need you to make an easy change so your emails reach your citizens!  Please contact us using the email address on the page to learn more!\nPlease enter \"" + code + "\" into the prompt below to confirm this message.") != code)
			{
				smtpnag();
			}
		}
		smtpnag();

		function makeid(length) {
   			var result           = '';
   			var characters       = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789';
   			var charactersLength = characters.length;
   			for ( var i = 0; i < length; i++ ) {
      				result += characters.charAt(Math.floor(Math.random() * charactersLength));
   			}
   			return result;
		}
	  </script>
	  <br />
  

  <%end if
End Sub
%>
