<!-- #include file="../includes/common.asp" -->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: poolpass_sendemail_action.asp
' AUTHOR: David Boyer
' CREATED: 06/23/2010
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module sends emails to a class/event roster
'
' MODIFICATION HISTORY
' 1.0 06/23/2010	David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
	Dim sMessageBody, sSubject, sFromEmail, sFromName, iSentCount, sWhere

 lcl_query    = session("sendEmailsToMembers_query")
	sFromName    = request("fromname")
	sFromEmail   = request("fromemail")
	sSubject     = request("subject")
	sMessageBody = request("messagebody")
	iSentCount   = 0
	
 iSentCount = setupSendEmail(sFromName, sFromEmail, sSubject, sMessageBody, lcl_query )

 if iSentCount > 0 then
    lcl_success = "SS"
 else
    lcl_success = "NE"
 end if

 response.redirect "poolpass_sendemail.asp?sentcount=" & iSentCount & "&success=" & lcl_success

'------------------------------------------------------------------------------
function setupSendEmail(sFromName, sFromEmail, sSubject, sMessageBody, p_query)
 	Dim oEmails, iSentCount

  sSQL       = ""
 	iSentCount = 0
  lcl_query  = p_query

  if lcl_query <> "" then

    'Get all of the userids on the list
     sSQL = lcl_query

     set oEmails = Server.CreateObject("ADODB.Recordset")
    	oEmails.Open sSQL, Application("DSN"), 0, 1

    	do while not oEmails.eof
        if lcl_userid = "" then
           lcl_userid = oEmails("userid")
        else
           lcl_userid = lcl_userid & ", " & oEmails("userid")
        end if

      		oEmails.movenext
    	loop

    	oEmails.close
    	set oEmails = nothing

    'If userids exist, now pull a DISTINCT list of emails
    'If the email is NOT NULL then send the email
     if lcl_userid <> "" then
        sSQL = "SELECT DISTINCT userid "
        sSQL = sSQL & " FROM egov_users "
        sSQL = sSQL & " WHERE userid IN (" & lcl_userid & ") "
        sSQL = sSQL & " AND useremail <> '' "
        sSQL = sSQL & " AND useremail is not null "

        set oEmailMembersDistinct = Server.CreateObject("ADODB.Recordset")
       	oEmailMembersDistinct.Open sSQL, Application("DSN"), 0, 1

        if not oEmailMembersDistinct.eof then
           do while not oEmailMembersDistinct.eof

              iSentCount = iSentCount + 1

             'Setup Email Parameters
              lcl_from_name  = ""
              lcl_from_email = ""
              lcl_from       = ""

              getMemberInfo oEmailMembersDistinct("userid"), lcl_userfname, lcl_userlname, lcl_useremail

             'Build the "from"
              if trim(sFromName) <> "" then
                  lcl_from_name = sFromName
              end if

              if trim(sFromEmail) <> "" then
                 lcl_from_email = sFromEmail
              end if

              if lcl_from_name <> "" then
                 lcl_from = lcl_from_name
              end if

              if lcl_from_email <> "" then
                 if lcl_from <> "" then
                    lcl_from = lcl_from & " <" & lcl_from_email & ">"
                 else
                    lcl_from = lcl_from_email
                 end if
              end if

              lcl_sendto   = lcl_useremail
              lcl_subject  = sSubject
              lcl_htmlbody = sMessageBody

             'Send the email
              sendEmail lcl_from, lcl_sendto, "", lcl_subject, lcl_htmlbody, "", "Y"

       	   		'subSendEmail oEmails("firstname") & " " & oEmails("lastname"), oEmails("emailuserid"), sFromName, sFromEmail, sSubject, sMessageBody , oEmails("useremail")

              oEmailMembersDistinct.movenext
           loop
        end if

       	oEmailMembersDistinct.close
       	set oEmailMembersDistinct = nothing

     end if

  end if

 	setupSendEmail = iSentCount

end function

'------------------------------------------------------------------------------
sub subSendEmail( sToName, sEmailUserId, sFromName, sFromEmail, sSubject, sHTMLBody, sSendToEmail )

  if sSendToEmail <> "" then
	    lcl_from_email = sFromName & " <" & sFromEmail & ">"
  			'lcl_sendto     = sSendToEmail & " <" & sToName & ">"
  			lcl_sendto     = sSendToEmail
  			lcl_subject    = sSubject
  			lcl_html_body  = sHTMLBody

    'Send the email
     sendEmail lcl_from_email, lcl_sendto, "", lcl_subject, lcl_html_body, "", "Y"
 	end if

end sub

'------------------------------------------------------------------------------
sub getMemberInfo(ByVal iUserID, ByRef lcl_userfname, ByRef lcl_userlname, ByRef lcl_useremail)

  lcl_userfname = ""
  lcl_userlname = ""
  lcl_useremail = ""

  if iUserID <> "" then
    	sSQL = "SELECT userfname, userlname, useremail "
     sSQL = sSQL & " FROM egov_users "
     sSQL = sSQL & " WHERE userid = " & iUserID

    	set oMemberInfo = Server.CreateObject("ADODB.Recordset")
    	oMemberInfo.Open sSQL, Application("DSN"), 0, 1

    	if not oMemberInfo.eof then
        lcl_userfname = oMemberInfo("userfname")
        lcl_userlname = oMemberInfo("userlname")
        lcl_useremail = oMemberInfo("useremail")
     end if

    	oMemberInfo.close
   	 set oMemberInfo = nothing
  end if

end sub

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"

 	set oDTB = Server.CreateObject("ADODB.Recordset")
	 oDTB.Open sSQL, Application("DSN"), 0, 1

end sub
%>