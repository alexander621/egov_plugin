<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<!-- #include file="merge_field_functions.asp" //-->
<%
 Dim iLetterID, itrackid, iUserid, sStatus, sSubStatus

 iAction      = request("action")
 iLetterID    = request("iletterid")
 iTrackID     = request("itrackid")
 sStatus      = request("status")
 sSubStatus   = request("substatus")
 iUserID      = getContactUserID(iTrackID)
 FLtitle      = getFormLetterInfo(iLetterID)
 iAddText     = request("add_text")
 iUserEmail   = getUserEmail(iUserID)
 iSubmitDate  = ConvertDateTimetoTimeZone()

 if iAddText <> "" then
    iAddText = replace(iAddText,"<<AMP>>","&")	
 end if

 if iAction = "WORDEXPORT" then
   	response.ContentType = "application/msword"
    response.AddHeader "Content-Disposition", "attachment;filename=formletter_" & iTrackID & hour(now) & minute(now) & second(now) & ".doc"

    sBody = fnWriteLetter(iLetterID, iAddText)
  		sBody = fnFillMergeFields(sBody,iTrackID,iUserID,iAddText,iTrackID)
    sBody = fnstriphtml(sBody)

  	 response.write sBody

 else
   'Determine if we need to build a message for the Activity Log
    if iAction = "PRINT" OR iAction = "EMAIL" then
       FLtitle = buildActivityLogMsg(FLtitle, iAddText, iUserEmail, iAction)
    end if
%>
<html>
<head>
  <title>E-GovLink Administration Consule {Action Line - Form Letter View}</title>
  <link rel="stylesheet" type="text/css" href="<%=session("egovclientwebsiteurl")%>/admin/global.css">
</head>
<body>
<table border="0" cellspacing="0" cellpadding="10" width="550">
  <tr>
      <td valign="top">
         <%
           if iAction = "PREVIEW" then
              response.write "<p align=""center""><strong>PREVIEW ONLY - No Activity Logged</strong>" & vbcrlf
           elseif iAction = "PRINT" then
              lcl_updatelog = "Y"
           elseif iAction = "EMAIL" then
              setupSendEmail iLetterID, iAddText, iUserEmail, iSubmitDate, lcl_updatelog
           end if

          'Determine if we need to build a message for the Activity Log
           if iAction = "PRINT" OR iAction = "EMAIL" then
              if lcl_updatelog = "Y" then
                 AddCommentTaskComment FLtitle, "", sStatus, iTrackID, session("userid"), session("orgid"), sSubStatus, iUserID, iUserEmail
              end if
           end if
        %>
        <p><%=fnWriteLetter(iLetterID, iAddText)%></p>
      </td>
  </tr>
</table>
<!--include file="bottom_include.asp"-->
</body>
</html>
<%
 end if

'------------------------------------------------------------------------------
function fnWriteLetter(p_letterid, p_addtext)

  lcl_return = ""

  if p_letterid <> "" then
     sSQL = "SELECT * "
     sSQL = sSQL & " FROM FormLetters "
     sSQL = sSQL & " WHERE FLid = " & p_letterid

     set oLetter = Server.CreateObject("ADODB.Recordset")
     oLetter.Open sSQL, Application("DSN"), 3, 1

   	'If there is a template adjust fields
    	if not oLetter.eof then
        lcl_return  = oLetter("FLbody")
        'lcl_return  = replace(lcl_return, "'", "''" )
        'lcl_return  = replace(lcl_return, chr(13), "<br />" ) 
        lcl_return  = fnFillMergeFields(lcl_return, iTrackID, iUserID, p_addtext, iTrackID)

       'Check to see if message contains HTML.
        if oLetter("containsHTML") then
           lcl_containsHTML = "Y"
        else
           lcl_containsHTML = "N"
        end if

        if lcl_containHTML = "N" then
           lcl_return  = replace(lcl_return, chr(13), "<br />" ) 
        end if

       'Check for custom html.  If any HTML tags exist in the email body the admin is sending then leave our HTML off of the email.
        lcl_customhtml = 0

        if instr(UCASE(lcl_return),"<HTML") > 0 then
           lcl_customhtml = lcl_customhtml + 1
        end if

        if instr(UCASE(lcl_return),"<HEAD") > 0 then
           lcl_customhtml = lcl_customhtml + 1
        end if

        if instr(UCASE(lcl_return),"<BODY") > 0 then
           lcl_customhtml = lcl_customhtml + 1
        end if

        if lcl_return <> "" then
          'If the message IS in HTML format already there is no need to set the carriage returns "chr(10)" to "<br />" tags.
          'If the message is NOT in HTML format then we want to make sure that the carriage returns are picked up 
           if lcl_containsHTML = "Y" then
              lcl_replace_linebreaks = ""
           else
              lcl_replace_linebreaks = "<br />"
           end if

           lcl_return = replace(lcl_return,vbcrlf,lcl_replace_linebreaks & vbcrlf)
        end if
     end if

     oLetter.close
     set oLetter = nothing
  end if

  fnWriteLetter = lcl_return

end function

'------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function

'------------------------------------------------------------------------------
Function JSsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  strDB = Replace( strDB, "'", "\'" )
  strDB = Replace( strDB, chr(34), "\'" )
  strDB = Replace( strDB, ";", "\;" )
  strDB = Replace( strDB, "-", "\-" )
  JSsafe = strDB
End Function

'------------------------------------------------------------------------------
'function AddCommentTaskComment(sInternalMsg,sExternalMsg,sStatus,iFormID,iUserID,iOrgID,sSubStatus,p_submitdate)

'  if sSubStatus = "" then
'     lcl_sub_status = 0
'  else
'     lcl_sub_status = sSubStatus
'  end if

'  sSQL = "INSERT egov_action_responses ("
'  sSQL = sSQL & "action_status,"
'  sSQL = sSQL & "action_internalcomment,"
'  sSQL = sSQL & "action_externalcomment,"
'  sSQL = sSQL & "action_userid,"
'  sSQL = sSQL & "action_orgid,"
'  sSQL = sSQL & "action_autoid,"
'  sSQL = sSQL & "action_sub_status_id,"
'  sSQL = sSQL & "action_editdate"
'  sSQL = sSQL & ") VALUES ("
'  sSQL = sSQL & "'" & sStatus              & "', "
'  sSQL = sSQL & "'" & DBsafe(sInternalMsg) & "', "
'  sSQL = sSQL & "'" & DBsafe(sExternalMsg) & "', "
'  sSQL = sSQL & "'" & iUserID              & "', "
'  sSQL = sSQL & "'" & iOrgID               & "', "
'  sSQL = sSQL & "'" & iFormID              & "', "
'  sSQL = sSQL       & lcl_sub_status       & ", "
'  sSQL = sSQL & "'" & p_submitdate         & "' "
'  sSQL = sSQL & ")"

'  set oComment = Server.CreateObject("ADODB.Recordset")
'  oComment.Open sSQL, Application("DSN") , 3, 1
'  set oComment = Nothing

 'Update the task sub-status
'  sSQL = "UPDATE egov_actionline_requests "
'  sSQL = sSQL & " SET sub_status_id = "   & lcl_sub_status
'  sSQL = sSQL & " WHERE action_autoid = " & iFormID

'  set oUpdate2 = Server.CreateObject("ADODB.Recordset")
'  oUpdate2.Open sSQL, Application("DSN") , 3, 1
'  set oUpdate2 = nothing

'end function

'------------------------------------------------------------------------------
function getFormLetterInfo(p_letterid)

  lcl_return = ""

  if p_letterid <> "" then
    'Get the title of the form letter
     sSQL = "SELECT FLtitle "
     sSQL = sSQL & " FROM FormLetters "
     sSQL = sSQL & " WHERE FLid = " & iLetterID

     set oFormLetter = Server.CreateObject("ADODB.Recordset")
     oFormLetter.Open sSQL, Application("DSN"), 3, 1

     if not oFormLetter.eof then
        lcl_return = oFormLetter("FLtitle")
     end if

     oFormLetter.close
     set oFormLetter = nothing
  end if

  getFormLetterInfo = lcl_return

end function

'------------------------------------------------------------------------------
function getContactUserID(p_trackid)

  lcl_return = 0

  if p_trackid <> "" then
    'Get Contact userid for this request
     sSQL = "SELECT userid "
     sSQL = sSQL & " FROM egov_actionline_requests "
     sSQL = sSQL & " WHERE action_autoid = " & p_trackid

     set oCID = Server.CreateObject("ADODB.Recordset")
     oCID.Open sSQL, Application("DSN"), 3, 2

     if not oCID.eof then
       	lcl_return = oCID("userid")
     end if

     oCID.close
     set oCID = nothing
  end if

  getContactUserID = lcl_return

end function

'------------------------------------------------------------------------------
function getUserEmail(p_userid)

  lcl_return = "UNKNOWN"

  if p_userid <> "" then
     sSQL = "SELECT userEmail "
     sSQL = sSQL & " FROM egov_users "
     sSQL = sSQL & " WHERE userid = " & p_userid

     set oUserEmail = Server.CreateObject("ADODB.Recordset")
     oUserEmail.Open sSQL, Application("DSN"), 3, 1

     if not oUserEmail.eof then
        lcl_return = oUserEmail("userEmail")
     end if

     oUserEmail.close
     set oUserEmail = nothing
  end if

  getUserEmail = lcl_return

end function

'------------------------------------------------------------------------------
function buildActivityLogMsg(p_fltitle, p_addtext, p_useremail, p_action)

 'Get the "label" of the current action so we can display it in the title.
  iActionLabel = getActionLabel(iAction)

 'Begin building the title of the form letter
  lcl_return = p_fltitle & " Form Letter " & iActionLabel

 'If sending an email then concatenate the user email to the title
  if p_action = "EMAIL" then
     lcl_return = lcl_return & " " & p_useremail
  end if

 'If additional text has been entered then format it properly.
  if isnull(p_addtext) OR p_addtext = "" then
   		lcl_addtext = ""
  else
     lcl_addtext = p_addtext
   		lcl_addtext = " - " & replace(lcl_addtext,"<br />"," ")
   		lcl_addtext = replace(lcl_addtext,"%20"," ")
  end if 

 'Finishing building the form letter title.
 	lcl_return = lcl_return & lcl_add_text

  buildActivityLogMsg = lcl_return

end function

'------------------------------------------------------------------------------
function getActionLabel(p_action)

  lcl_return = ""

  if p_action <> "" then
     lcl_action = UCASE(p_action)

     if lcl_action = "PREVIEW" then
        lcl_return = "Preview"
     elseif lcl_action = "PRINT" then
        lcl_return = "Printed"
     elseif lcl_action = "EMAIL" then
        lcl_return = "Emailed to"
     elseif lcl_action = "WORDEXPORT" then
        lcl_return = "Exported to MS Word"
     end if

  end if

  getActionLabel = lcl_return

end function

'------------------------------------------------------------------------------
sub setupSendEmail(ByVal p_letterid, ByVal p_addtext, ByVal p_useremail, ByVal p_submitdate, ByRef lcl_updatelog)

  lcl_updatelog   = "N"

  if p_letterid <> "" then
     sMsg = fnWriteLetter(p_letterid, p_addtext)
     sMsg = replace(sMsg,"<br><br>","<br>")
     sMsg = replace(sMsg,"<br>&nbsp;&nbsp;&nbsp;",vbcrlf & "   ")
     sMsg = replace(sMsg,"<br>"," ") & vbcrlf
     sMsg = replace(sMsg,"<p>",vbcrlf & vbcrlf)

     sOrgName           = getDefaultOrgValue("orgname")
     lcl_email_from     = sOrgName & " (E-Gov Website) <noreplies@egovlink.com>"
     sEmailToAddress    = p_useremail
     lcl_email_subject  = "Form Letter Email Message"
     lcl_email_htmlbody = BuildHTMLMessage(sMsg,"Y")

    'Remove the name from the email address
     lcl_validate_email = formatSendToEmail(sEmailToAddress)

    'The function isValidEmail (found in common.asp) allows an email to simply have an "@" sign at the end of the email.
    'However, this will crash the application.  Check to see if the last character in the email entered is an "@".
     if lcl_validate_email <> "" AND RIGHT(lcl_validate_email,1) <> "@" then
        if isValidEmail(lcl_validate_email) then

          'Send the email if it is valid.
           sendEmail lcl_email_from, sEmailToAddress, "", lcl_email_subject, lcl_email_htmlbody, "", "Y"

          'Display "email sent" message on the screen
           response.write "<p align=""right""><strong>Email Sent to: " & sEmailToAddress & "</strong> on " & p_submitdate & "</p>" & vbcrlf

           lcl_updatelog = "Y"
        end if
     end if
  end if

end sub

'------------------------------------------------------------------------------
function fnstriphtml(sValue)

  sValue = replace(sValue,"<br>&nbsp;&nbsp;&nbsp;",vbcrlf & "   ")
  sValue = replace(sValue,"<p>",vbcrlf & vbcrlf)

	 lcl_return = sValue

 	Dim objRegExp, strOutput
	 Set objRegExp = New Regexp

	 objRegExp.IgnoreCase = True
	 objRegExp.Global     = True
	 objRegExp.Pattern    = "<(.|\n)+?>"

	'Replace all HTML tag matches with the empty string
  lcl_return = objRegExp.replace(sValue, "")

	 set objRegExp = nothing

  fnstriphtml = lcl_return

end function
%>
