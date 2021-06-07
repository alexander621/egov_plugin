<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<%
 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"),"create requests") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Check for org features
 lcl_orghasfeature_large_address_list                       = orghasfeature("large address list")
 lcl_orghasfeature_issue_location                           = orghasfeature("issue location")
 lcl_orghasfeature_hide_actionline_details                  = orghasfeature("hide actionline details")
 lcl_orghasfeature_fileupload                               = orghasfeature("fileupload")
 lcl_orghasfeature_actionline_secure_attachments            = orghasfeature("actionline_secure_attachments")
 lcl_orghasfeature_actionline_display_attachments_to_public = orghasfeature("actionline_display_attachments_to_public")

'Check for user permissions
 lcl_userhaspermission_fileupload                               = userhaspermission(session("userid"),"fileupload")
 lcl_userhaspermission_actionline_secure_attachments            = userhaspermission(session("userid"),"actionline_secure_attachments")
 lcl_userhaspermission_actionline_display_attachments_to_public = userhaspermission(session("userid"),"actionline_display_attachments_to_public")

 dim iAdminCount, aAdminUserIDs(2), aAdminEmails(2), sAdminEmail, i
 iAdminCount = 0

'Parse to get title for action request form
 if instr(request("actionid"),"|") > 0 then
 			arrForm     = split(request("actionid"),"|")
 			actionid    = arrForm(0)
 			actiontitle = arrForm(1)
 			actiontitle = replace(actiontitle,"\"," > ")
 			actiontitle = replace(actiontitle,"/"," > ")
 else
 			actionid    = request("actionid")
 			actiontitle = request("actiontitle")
 end if

'Enable/Disable the "SAVE" button for attachments.
 if lcl_orghasfeature_fileupload AND lcl_userhaspermission_fileupload then
    if lcl_onload <> "" then
       lcl_onload = lcl_onload & "validateAttachment();"
    else
       lcl_onload = "validateAttachment();"
    end if
 end if

'Make sure we have a valid RequestID before proceeding
 if actionid = "" then
 	  response.write "<font class=""error"">!There was an error processing this request. No action form found for this submission!</font>"
	   response.end
 end if

'Get the internal default firstname, lastname, and email for this org
 lcl_internal_email = GetInternalDefaultEmail(session("orgid"))

 if lcl_internal_email = "" then
    lcl_internal_email = "Dev Support: No default email for orgid (" &session("orgid")& ") <devsupport@eclink.com>"
 else
    lcl_internal_email = "Internal Default Email <" & lcl_internal_email & ">"
 end if

'Get admin email addresses for this action request form.
 sSQL = "SELECT assigned_userid,assigned_userid2,assigned_userid3, action_form_resolved_status, deptid "
 sSQL = sSQL & " FROM egov_action_request_forms "
 sSQL = sSQL & " WHERE action_form_id=" & actionid

 set oAdmin = Server.CreateObject("ADODB.Recordset")
 oAdmin.Open sSQL, Application("DSN"), 3, 1

 if not oAdmin.eof then
    lcl_resolved_status = oAdmin("action_form_resolved_status")
    lcl_group_id        = oAdmin("deptid")

 		'** 1st ASSIGNED-TOP
 		if oAdmin("assigned_userid") = "" or isNull(oAdmin("assigned_userid")) then
 			  sAdminEmail                = GetInternalDefaultEmail(session("orgid"))
 			  adminEmailAddr             = "<" & sAdminEmail & ">"
 			  aAdminEmails(iAdminCount)  = sAdminEmail
      aAdminUserIDs(iAdminCount) = oAddress("assigned_userid")
 			  adminid                    = 0
 		else
   			sSQL = "SELECT email, userid, lastname, firstname "
      sSQL = sSQL & " FROM users "
      sSQL = sSQL & " WHERE userid = " & oAdmin("assigned_userid")

 		  	set oAddress = Server.CreateObject("ADODB.Recordset")
   			oAddress.Open sSQL, Application("DSN"), 3, 1

   			if not oAddress.eof then
         if oAddress("email") = "" then
	 			       sAdminEmail = lcl_internal_email
     				else
			 	       sAdminEmail = oAddress("firstname") & " " & oAddress("lastname") & " <" & oAddress("email") & ">"  'Assigned Admin User Email
     				end if

     				adminFromAddr              = sAdminEmail
     				adminEmailAddr             = sAdminEmail
 				    aAdminEmails(iAdminCount)  = sAdminEmail
         aAdminUserIDs(iAdminCount) = oAddress("userid")
     				adminid                    = oAddress("userid")  'Assigned Admin UserID
   		 end if

   			oAddress.close
	 	  	set oAddress = nothing
   end if

		'** 2nd ASSIGNED-TOP
 		if oAdmin("assigned_userid2") <> "" AND NOT isNull(oAdmin("assigned_userid2")) then
 			  sSQL = "SELECT email, lastname, firstname "
      sSQL = sSQL & " FROM users "
      sSQL = sSQL & " WHERE userid = " & oAdmin("assigned_userid2") 

 			  set oAddress = Server.CreateObject("ADODB.Recordset")
 			  oAddress.Open sSQL, Application("DSN"), 3, 1

   			if not oAddress.eof then
 		    		iAdminCount = iAdminCount + 1

         if oAddress("email") = "" then
 				       sAdminEmail = lcl_internal_email
     				else
 				       sAdminEmail = oAddress("firstname") & " " & oAddress("lastname") & " <" & oAddress("email") & ">"  'Assigned Admin User Email
     				end if

     				adminEmailAddr             = adminEmailAddr & ", " & sAdminEmail   
 				    aAdminEmails(iAdminCount)  = sAdminEmail
         aAdminUserIDs(iAdminCount) = oAdmin("assigned_userid2")
   			end if

   			oAddress.close
 		  	set oAddress = nothing
 		end if

		'** 3rd ASSIGNED-TOP
 		if oAdmin("assigned_userid3") <> "" AND NOT isNull(oAdmin("assigned_userid3")) then
   			sSQL = "SELECT email, lastname, firstname "
      sSQL = sSQL & " FROM users "
      sSQL = sSQL & " WHERE userid = " & oAdmin("assigned_userid3") 

 		  	set oAddress = Server.CreateObject("ADODB.Recordset")
   			oAddress.Open sSQL, Application("DSN"), 3, 1

   			if not oAddress.eof then
 		    		iAdminCount = iAdminCount + 1

         if oAddress("email") = "" then
 				       sAdminEmail = lcl_internal_email
     				else
 				       sAdminEmail = oAddress("firstname") & " " & oAddress("lastname") & " <" & oAddress("email") & ">"  'Assigned Admin User Email
     				end if

     				adminEmailAddr             = adminEmailAddr & ", " & sAdminEmail   
 				    aAdminEmails(iAdminCount)  = sAdminEmail
         aAdminUserIDs(iAdminCount) = oAdmin("assigned_userid3")
   		 end If

   			oAddress.Close
 		  	Set oAddress = Nothing
 		end if

 end if

 oAdmin.close
 set oAdmin = nothing

'This catches the case where an admin has been assigned but they do not have and email entered for them
 if trim(adminEmailAddr) = "" then
   	adminEmailAddr = "<" & GetInternalDefaultEmail(session("orgid")) & ">"
   	aAdminEmails(0)  = adminEmailAddr
   	aAdminEmails(1)  = adminEmailAddr
   	aAdminEmails(2)  = adminEmailAddr
   	adminid          = 0
    aAdminUserIDs(0) = ""
    aAdminUserIDs(1) = ""
    aAdminUserIDs(2) = ""
 end if

'Get questions and entered values
 sQuestions = ""

 for each oField in request.form
     sAnswer          = ""
     sFormattedAnswer = ""

     if left(oField,10) = "fmquestion" then

      		sQuestionPrompt = "fmname" & replace(oField,"fmquestion","")

      		sQuestions = sQuestions & "<p>" & vbcrlf
        sQuestions = sQuestions & "<b>" & request.form(sQuestionPrompt) & "</b><br>" & vbcrlf

        sAnswer = replace(request.form(oField),"default_novalue","")

        if trim(sAnswer) <> "" then
        '   sFormattedAnswer = formatToFitEmailLineLength(trim(sAnswer))
        'else
           sFormattedAnswer = sAnswer
        end if

        'sQuestions = sQuestions & replace(request.form(oField),"default_novalue","") & vbcrlf
        sQuestions = sQuestions & sFormattedAnswer & vbcrlf
        sQuestions = sQuestions & "</p>" & vbcrlf

   	 end if
 next

 sQuestions2 = sQuestions

'Timezone difference
 datOrgDateTime = ConvertDateTimetoTimeZone()

'Insert form information into database
 sSQL = "SELECT * FROM egov_actionline_requests WHERE action_autoid = 0"

 set oNewActionRequest = Server.CreateObject("ADODB.Recordset")
 oNewActionRequest.CursorLocation = 3
 oNewActionRequest.Open sSQL, Application("DSN"), 3, 2
 oNewActionRequest.AddNew
 oNewActionRequest("userid") = AddUserInformation()

 if adminid <> 0 then
  	'0 means that no admin person was assigned
   	oNewActionRequest("assignedemployeeid") = adminid
 end if

 oNewActionRequest("comment")                   = sQuestions
 oNewActionRequest("category_id")               = actionid
 oNewActionRequest("category_title")            = actiontitle
 oNewActionRequest("orgid")                     = session("orgid")
 oNewActionRequest("status")                    = "SUBMITTED"
 oNewActionRequest("submit_date")               = datOrgDateTime
 oNewActionRequest("contactmethodid")           = request("selContactMethod")
 oNewActionRequest("employeesubmitid")          = session("userid")
 oNewActionRequest("groupid")                   = lcl_group_id
 oNewActionRequest("submittedby_remoteaddress") = Request.ServerVariables("REMOTE_ADDR")

 oNewActionRequest.Update
 iTrackingNumber = oNewActionRequest("action_autoid")
 sStatus         = oNewActionRequest("status")



if session("orgid") = 153 and actionid = 17051 then
	DueDate = GetFOILDueDate(datOrgDateTime)


	sSQL = "UPDATE egov_actionline_requests SET due_date = '" & DueDate & "' WHERE action_autoid = " & iTrackingNumber
	RunSQLStatement sSQL
end if


'Record location information
 AddIssueLocation iTrackingNumber

'Update the Activity Log

 'AddCommentTaskComment "This request was submitted by " & getUserName(session("userid"),"FL") & ".","",oNewActionRequest("status"),oNewActionRequest("action_autoid"),session("userid"),Session("OrgID"), datOrgDateTime
 AddCommentTaskComment "This request was submitted by " & getUserName(session("userid"),"FL") & ".","",oNewActionRequest("status"),oNewActionRequest("action_autoid"),session("userid"),session("orgid"), "", "", ""

 oNewActionRequest.Close
 set oNewActionRequest = nothing 

'If the status is RESOLVED then populate the Activity Request Log so that the user can see that it has been resolved.
 if lcl_resolved_status = "Y" then
    sStatus = "RESOLVED"

   'First update the status on the Activity Request
    sSQL = "UPDATE egov_actionline_requests SET "
    sSQL = sSQL & "status = 'RESOLVED', "
    sSQL = sSQL & "complete_date = '" & datOrgDateTime & "'"
    sSQL = sSQL & " WHERE action_autoid = " & iTrackingNumber

    set rsu = Server.CreateObject("ADODB.Recordset")
    rsu.Open sSQL, Application("DSN"), 3, 1

    set rsu = nothing

   	'AddCommentTaskComment "Request's status was set to RESOLVED upon submission.", "", "RESOLVED", iTrackingNumber, session("userid"), session("orgid"), datOrgDateTime
   	AddCommentTaskComment "Request's status was set to RESOLVED upon submission.", "", "RESOLVED", iTrackingNumber, session("userid"), session("orgid"), "", "", ""
 end if

'Replaces blob functionality - stores data in prompt/answer format
 InsertRequestFieldsandResponses iTrackingNumber 

'Generate Tracking Number - (Formula is: SQL ROWID + HHMM)
 lngTrackingNumber = iTrackingNumber & replace(FormatDateTime(datOrgDateTime,4),":","")

'Send email
	if iorgid <> "7" then
   'Setup the problem location address
    lcl_problem_location = ""

    if lcl_orghasfeature_issue_location then
       sSQLf = "SELECT issuelocationname, action_form_display_issue, hideIssueLocAddInfo "
       sSQLf = sSQLf & " FROM egov_action_request_forms "
       sSQLf = sSQLf & " WHERE action_form_id = " & actionid

       set oForm = Server.CreateObject("ADODB.Recordset")
       oForm.Open sSQLf, Application("DSN") , 3, 1

       if not oForm.eof then
          sIssueName           = UCASE(oForm("issuelocationname"))
          blnIssueDisplay      = oForm("action_form_display_issue")
          sHideIssueLocAddInfo = oForm("hideIssueLocAddInfo")

          If Trim(sIssueName) = "" OR IsNull(sIssueName) Then
             sIssueName = "ISSUE/PROBLEM LOCATION:"
          End If

         '1. Check to see if the "issue location" feature has been "turned on" for this form.
         '2. Check to see if the org has the large address list feature "turned on"
         '3. Check to see if a street number has been entered.
          if blnIssueDisplay = True then
           		if lcl_orghasfeature_large_address_list then
                lcl_problem_location = request("residentstreetnumber")

               '4. Check to see if a value in the dropdown list has been selected.
                  'It doesn't matter if the large address feature has been turned on/off.
                  'If an org has the "issue location" feature then the dropdown will appear.
                if request("skip_address") <> "0000" then
                   if lcl_problem_location <> "" then
                      lcl_problem_location = lcl_problem_location & " " & request("skip_address")
                   else
                      lcl_problem_location = request("skip_address")
                   end if
                else
                  '5. If no value has been selected in the dropdown list then check
                     'to see if the "other" address has been populated.
                     'If it has then override the street number, if it was populated.
                     'If not then display whatever has been entered.  The screen will enforce
                     'a value to be entered for the address if the street number has been entered.
                   if request("ques_issue2") <> "" then
                      lcl_problem_location = request("ques_issue2")
                   end if
                end if
             else
               '6. If the org does NOT have the "large address list" feature "turned on"
                  'then if a value has been selected from the dropdown list retrieve the street address
                if request("skip_address") <> "0000" then
                   sSQLa = "SELECT residentstreetnumber, residentstreetname "
                  	sSQLa = sSQLa & " FROM egov_residentaddresses "
                   sSQLa = sSQLa & " WHERE residentaddressid=" & request("skip_address")

                   Set rs = Server.CreateObject("ADODB.Recordset")
                   rs.Open sSQLa, Application("DSN") , 3, 1

                   if not rs.eof then
                      if rs("residentstreetnumber") <> "" then
                         lcl_problem_location = rs("residentstreetnumber")
                      end if

                      if rs("residentstreetname") <> "" then
                         if lcl_problem_location <> "" then
                            lcl_problem_location = lcl_problem_location & " " & rs("residentstreetname")
                         else
                            lcl_problem_location = rs("residentstreetname")
                         end if
                      end if
                   end if
                else
                   if request("ques_issue2") <> "" then
                      lcl_problem_location = request("ques_issue2")
                   end if
                end if
             end if  'END orghasfeature("large address list")
          end if  'END blnIssueDisplay
       end if  'END eof
	   end if  'END orghasfeature("issue location")

		  set oOrg = New classOrganization

    lcl_featurename_actionline = getFeatureName("action line")

  		sMsg2 = sMsg2 & "<p>This automated message was sent by the " & oOrg.GetOrgName() & " E-Gov web site. Do not reply to this message.  Contact " & adminFromAddr & " for inquiries regarding this email.</p><br />"
 		 sMsg2 = sMsg2 & "<p>A " & oOrg.GetOrgName() & " " & lcl_featurename_actionline & " issue was submitted on " & FormatDateTime(datOrgDateTime,1) & ".</p><br />"
 		 sMsg2 = sMsg2 & "<p><strong>Click the following link to view this " & lcl_featurename_actionline & " Request:</strong><br />"
	  	sMsg2 = sMsg2 & "<a href=""" & oOrg.GetEgovURL() & "/admin/action_line/action_respond.asp?control=" & iTrackingNumber & "&e=Y"">"
  		sMsg2 = sMsg2 & oOrg.GetEgovURL() & "/admin/action_line/action_respond.asp?control=" & iTrackingNumber & "&e=Y</a></p>"
				
  		if not lcl_orghasfeature_hide_actionline_details then
		     sMsg2 = sMsg2 & "<br />"
    			sMsg2 = sMsg2 & "<p><strong>" & ucase(lcl_featurename_actionline) & " REQUEST DETAILS</strong><br />"
   				sMsg2 = sMsg2 & "DATE SUBMITTED: "  & datOrgDateTime    & "<br />"
			    sMsg2 = sMsg2 & "TRACKING NUMBER: " & lngTrackingNumber & "<br />"
    			sMsg2 = sMsg2 & "CATEGORY ID: "     & actionid          & "<br />"
				   sMsg2 = sMsg2 & "CATEGORY Title: "  & actiontitle       & "</p><br />"
    			sMsg2 = sMsg2 & "<p><strong>SUGGESTION/ISSUE: ...</strong>"
				   sMsg2 = sMsg2 & replace(sQuestions,"default_novalue","") & "</p>"
       'sMsg2 = sMsg2 & vbcrlf & sQuestions & vbcrlf & "</p>" & vbcrlf

       if lcl_orghasfeature_issue_location then
          if blnIssueDisplay then
          			sMsg2 = sMsg2 & "<p><strong>" & sIssueName & "</strong><br />"
             sMsg2 = sMsg2 & "LOCATION: "  & lcl_problem_location & "<br />"

             if not sHideIssueLocAddInfo then
             			sMsg2 = sMsg2 & "ADDITIONAL INFORMATION: " & request("ques_issue6") & "<br />"
             end if

          end if
       end if

     		sMsg2 = sMsg2 & "<p><strong>" & ucase(lcl_featurename_actionline) & " REQUESTER CONTACT INFORMATION</strong><br />"
    			sMsg2 = sMsg2 & "NAME: "     & Request("cot_txtFirst_Name") & " " & Request("cot_txtLast_Name") & "<br />"
    			sMsg2 = sMsg2 & "BUSINESS: " & Request("cot_txtBusiness_Name")              & "<br />"
   				sMsg2 = sMsg2 & "EMAIL: "    & Request("cot_txtEmail")                      & "<br />"
			    sMsg2 = sMsg2 & "PHONE: "    & FormatPhone(Request("cot_txtDaytime_Phone")) & "<br />"
    			sMsg2 = sMsg2 & "FAX: "      & Request("cot_txtFax")                        & "<br />"
				   sMsg2 = sMsg2 & "ADDRESS: "  & Request("cot_txtStreet")                     & "<br />"
    			sMsg2 = sMsg2 & "" & Request("cot_txtCity") & " " & Request("cot_txtState_vSlash_Province") & "<br />"
				   sMsg2 = sMsg2 & "" & Request("cot_txtZIP_vSlash_Postal_Code") & " " & Request("cot_txtCountry") & "</p>"
  		end if

   'Loop through the admin emails
				for iEmailCount = 0 to iAdminCount
        if aAdminEmails(iEmailCount) <> "" then

          'Remove the name from the email address
           lcl_validate_email = formatSendToEmail(aAdminEmails(iEmailCount))

           if isValidEmail(lcl_validate_email) then

             'Check to see if the user wishes to send out the email
              lcl_sendemail = checkSendEmail(request("doNotSendSelfEmail"),request("doNotSendAllEmail"),lcl_validate_email)

             'Send the email
              if lcl_sendemail = "Y" then

                'Check for a delegate
                 getDelegateInfo aAdminUserIDs(iEmailCount), lcl_delegateid, lcl_delegate_username, lcl_delegate_useremail

                'Setup the SENDTO and check for a DELEGATE
                 setupSendToAndDelegateEmails aAdminEmails(iEmailCount), lcl_delegate_useremail, lcl_email_sendto, lcl_email_cc

                 lcl_subject = "Submission: " & actiontitle & " (re: " & request("cot_txtFirst_Name") & " " & request("cot_txtLast_Name") & ")"
                 lcl_message = BuildHTMLMessage(sMsg2,"Y")

                 sendEmail "",lcl_email_sendto,lcl_email_cc,lcl_subject,lcl_message,"","Y"
              end if
           end if
        end if
    next

				set oOrg = nothing 

'------------------------------------------------------------------------------
 else  'orgid = 7
'------------------------------------------------------------------------------
		  sMsgBody = BuildHTMLMessage(BuildAdminHTMLBody(),"Y")

   'Send email to Helpdesk admin
    if adminEmailAddr <> "" then
       if right(adminEmailAddr,1) <> "@" then

         'Remove the name from the email address
          lcl_validate_email = formatSendToEmail(adminEmailAddr)

          if isValidEmail(lcl_validate_email) then

            'Check to see if the user wishes to send out the email
             lcl_sendemail = checkSendEmail(request("doNotSendSelfEmail"),request("doNotSendAllEmail"),lcl_validate_email)

            'Check for a delegate
             getDelegateInfo adminEmailAddr, lcl_delegateid, lcl_delegate_username, lcl_delegate_useremail

            'Setup the SENDTO and check for a DELEGATE
             setupSendToAndDelegateEmails adminEmailAddr, lcl_delegate_useremail, lcl_email_sendto, lcl_email_cc

            'Send the email
             if lcl_sendemail = "Y" then
                lcl_subject = "Submission: EC Link HelpDesk - HelpDesk Ticket"
                'sendEmail sOrgName & " EC Link HelpDesk <webmaster@eclink.com>",lcl_email_sendto,lcl_email_cc,lcl_subject,sMsgBody,"","Y"
                sendEmail sOrgName & " EC Link HelpDesk <noreply@eclink.com>",lcl_email_sendto,lcl_email_cc,lcl_subject,sMsgBody,"","Y"
             end if
          end if
       end if
    end if
 end if

'Record view of request
 if request.servervariables("REQUEST_METHOD") <> "POST" then
    'datOrgDateTime = ConvertDateTimetoTimeZone()

   	'AddCommentTaskComment "Request viewed by " & session("FullName") & ".", null, "SUBMITTED", iTrackingNumber, session("userid"), session("orgid"), datOrgDateTime
   	AddCommentTaskComment "Request viewed by " & session("FullName") & ".", null, "SUBMITTED", iTrackingNumber, session("userid"), session("orgid"), "", "", ""
 end if
%>
<html>
<head>
<title>E-Gov Administration Console {Create a Request}</title>

<!-- This metadata is for setting the priority and importance for CDO mail messages -->
<!--  
METADATA  
TYPE="typelib"  
UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
NAME="CDO for Windows 2000 Library"  
--> 

<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
<link rel="stylesheet" type="text/css" href="css/styles.css" />
<link rel="stylesheet" type="text/css" href="../global.css" />

<script language="javascript" src="scripts/modules.js"></script>

<script language="javascript">
<!--
function openWin2(url, name) {
  popupWin = window.open(url, name,"resizable,width=500,height=450");
}

function validateAttachment() {
  var lcl_attachment = document.getElementById('filAttachment').value;

  //Always initially disable the button.  The code with determine if/when it is enabled.
  document.getElementById("saveAttachmentButton").disabled=true;

  if(lcl_attachment != "") {
     document.getElementById("saveAttachmentButton").disabled=true;

     //Make sure that an .EXE file is not being uploaded.
     lcl_length     = lcl_attachment.length;
     lcl_period_loc = lcl_attachment.indexOf(".");
     lcl_ext        = lcl_attachment.substr(lcl_period_loc+1);

     if(lcl_ext.toUpperCase()=='EXE') {
        alert("Invalid Value: File Type (" + lcl_ext + ").");
        return false;
     }else{
        document.getElementById("saveAttachmentButton").disabled=false;
     }
  }
}
//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<%
'BEGIN: Display information to the user ---------------------------------------
 response.write "<div id=""content"">" & vbcrlf
 response.write "	 <div id=""centercontent"">" & vbcrlf
 response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
 response.write "  <tr>" & vbcrlf
 response.write "      <td valign=""top"">" & vbcrlf
 response.write "          <p>" & vbcrlf
 response.write "             <input type=""button"" name=""returnListButton"" id=""returnListButton"" value=""Return to E-Gov Request Form Entry List"" class=""button"" onclick=""location.href='action.asp'"" />" & vbcrlf
 response.write "             <input type=""button"" name=""requestManagerButton"" id=""requestManagerButton"" value=""E-Gov Request Manager"" class=""button"" onclick=""location.href='action_respond.asp?control=" & iTrackingNumber & "&e=Y'"" />" & vbcrlf
 response.write "          </p>" & vbcrlf
 response.write "          <p>" & vbcrlf
 response.write "          <div style=""margin-left:20px;"" class=""box_header2"">" & lcl_featurename_actionline & " Request Submitted - " & datOrgDateTime &  "</div>" & vbcrlf
 response.write "          <div style=""margin-left:20px;"" class=""groupsmall"">" & vbcrlf
 response.write "            <p>Request form  <strong>(<i>" & actiontitle & "</i>)</strong> has been submitted.</p>" & vbcrlf
 response.write "            <p>Tracking number: <strong>" & lngTrackingNumber & "</strong>.</p>" & vbcrlf
 response.write "          </div>" & vbcrlf
 response.write "          </p>" & vbcrlf
'END: Display information to the user -----------------------------------------

'BEGIN: Attachments -----------------------------------------------------------
 if lcl_orghasfeature_fileupload AND lcl_userhaspermission_fileupload then
    response.write "          <p>" & vbcrlf
    response.write "          <div style=""margin-left:20px;"">" & vbcrlf

    'lcl_isMobile                = "N"
    lcl_form_name               = ""
    lcl_formpost_directorylevel = ""

    'subDisplayAttachments "N", _
    '                      iTrackingNumber, _
    '                      sStatus, _
    '                      lcl_orghasfeature_actionline_secure_attachments, _
    '                      lcl_userhaspermission_actionline_secure_attachments, _
    '                      lcl_orghasfeature_actionline_display_attachments_to_public, _
    '                      lcl_userhaspermission_actionline_display_attachments_to_public, _
    '                      lcl_isMobile, _
    '                      lcl_form_name, _
    '                      lcl_formpost_directorylevel

    subDisplayAttachments "N", _
                          iTrackingNumber, _
                          sStatus, _
                          lcl_orghasfeature_actionline_secure_attachments, _
                          lcl_userhaspermission_actionline_secure_attachments, _
                          lcl_orghasfeature_actionline_display_attachments_to_public, _
                          lcl_userhaspermission_actionline_display_attachments_to_public, _
                          lcl_form_name, _
                          lcl_formpost_directorylevel

    response.write "          </div>" & vbcrlf
    response.write "          </p>" & vbcrlf
 end if
'END: Attachments -------------------------------------------------------------

 response.write "      </td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "</table>" & vbcrlf
 response.write "	 </div>" & vbcrlf
 response.write "</div>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->  
<%
 response.write "</body>" & vbcrlf
 response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
function AddUserInformation()

	'Insert form information into database
 	iReturnValue = 0
	
 	set oUser = Server.CreateObject("ADODB.Recordset")
 	oUser.CursorLocation = 3
 	oUser.Open "SELECT * FROM egov_users WHERE 1=2", Application("DSN"), 3, 2
 	oUser.AddNew
 	oUser("userfname")        = dbsafe(request.form("cot_txtFirst_Name"))
 	oUser("userlname")        = dbsafe(request.form("cot_txtLast_Name"))
 	oUser("useremail")        = dbsafe(request.form("cot_txtEmail"))
 	oUser("userbusinessname") = request.form("cot_txtBusiness_Name")
 	oUser("userhomephone")    = request.form("cot_txtDaytime_Phone")
 	oUser("userfax")          = request.form("cot_txtFax")
 	oUser("useraddress")      = request.form("cot_txtStreet")
 	oUser("usercity")         = request.form("cot_txtCity")
 	oUser("userstate")        = request.form("cot_txtState_vSlash_Province")
 	oUser("userzip")          = request.form("cot_txtZIP_vSlash_Postal_Code")
 	oUser("usercountry")      = request.form("cot_txtCountry")
 	oUser.Update

 	iReturnValue = oUser("userid")

 	oUser.close
 	set oUser = nothing

 	AddUserInformation = iReturnValue

end function

'------------------------------------------------------------------------------
function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
end function

'------------------------------------------------------------------------------
function BuildAdminHTMLBody()

			sMsgAdmin = sMsgAdmin & "This automated message was sent by the ECLINK HELPDESK web site. Do not reply to this message.  Follow the instructions below or contact <strong>" & adminFromAddr & "</strong> for inquiries regarding this email.<br />"
			sMsgAdmin = sMsgAdmin & "<br />"
			sMsgAdmin = sMsgAdmin & "To follow-up with this help desk ticket please follow the link below:<br /><br />http://www.egovlink.com/" & sorgVirtualSiteName & "/admin<br /><br />"
			sMsgAdmin = sMsgAdmin & "<br /><strong>DATE SUBMITTED:</strong> "  & Now()
			sMsgAdmin = sMsgAdmin & "<br /><strong>TRACKING NUMBER:</strong> " & lngTrackingNumber
			sMsgAdmin = sMsgAdmin & "<br /><strong>HELP DESK FORM:</strong> "  & actiontitle
			sMsgAdmin = sMsgAdmin & "<br /><strong>HELP DESK TICKET DETAILS:</strong><br /><br />"
			sMsgAdmin = sMsgAdmin & sQuestions2
			sMsgAdmin = sMsgAdmin & "<br /><br /><strong>TICKET SUBMITTER CONTACT INFORMATION</strong>"
			sMsgAdmin = sMsgAdmin & "<br />NAME: "     & Request("cot_txtFirst_Name") & " " & Request("cot_txtLast_Name")
			sMsgAdmin = sMsgAdmin & "<br />BUSINESS: " & Request("cot_txtBusiness_Name")
			sMsgAdmin = sMsgAdmin & "<br />EMAIL: "    & Request("cot_txtEmail")
			sMsgAdmin = sMsgAdmin & "<br />PHONE: "    & Request("cot_txtDaytime_Phone")
			sMsgAdmin = sMsgAdmin & "<br />FAX: "      & Request("cot_txtFax")
			sMsgAdmin = sMsgAdmin & "<br />ADDRESS: "  & Request("cot_txtStreet")
			sMsgAdmin = sMsgAdmin & "<br />" & Request("cot_txtCity") & " " & Request("cot_txtState_vSlash_Province")
			sMsgAdmin = sMsgAdmin & "<br />" & Request("cot_txtZIP_vSlash_Postal_Code") & " " & Request("cot_txtCountry")

			BuildAdminHTMLBody = sMsgAdmin

end function

'------------------------------------------------------------------------------
function FormatPhone( Number )
  if Len(Number) = 10 then
     FormatPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
  else
     FormatPhone = Number
  end if
end function

'------------------------------------------------------------------------------
sub AddIssueLocation( iFormID )
 Dim oLocation, sNumber, sStreetPrefix, sStreetName, sStreetSuffix, sStreetDirection, sStreetUnit, sCity, sState, sZip
 Dim sLatitude, sLongitude, sValidStreet, sCounty, sParcelID, sExludefromAL, oActionOrg

 set oActionOrg = New classOrganization

'Set up variables
 sNumber           = ""
 sStreetPrefix     = ""
 sStreetName       = ""
 sStreetSuffix     = ""
 sStreetDirection  = ""
 sStreetUnit       = ""
 sSortStreetName   = ""
 sCity             = oActionOrg.GetDefaultCity()
 sState            = oActionOrg.GetDefaultState()
 sZip              = oActionOrg.GetDefaultZip()
 sCounty           = ""
 sResidentType     = ""
 sLatitude         = "0.00"
 sLongitude        = "0.00"
 sParcelID         = ""
 sLegalDesc        = ""
 sListedOwner      = ""
 sRegisteredUserID = 0
 sExcludeFromAL    = 0
 sValidStreet      = request("validstreet")

'Insert form information into database
 sSQL = "SELECT * FROM egov_action_response_issue_location WHERE 1=2"

 set oLocation = Server.CreateObject("ADODB.Recordset")
 oLocation.CursorLocation = 3
 oLocation.Open sSQL, Application("DSN"), 3, 2
 oLocation.AddNew
 oLocation("actionrequestresponseid") = iFormID

 if request("skip_address") <> "" then
  		if lcl_orghasfeature_large_address_list then
 	    	if request("skip_address") <> "0000" then	
	 	 		   'Try to match the input street number and selected street name to those in the database
      				MatchAddressInfo request("residentstreetnumber"), request("skip_address"), sNumber, sPrefix, sStreetName, sSuffix, sDirection, _
                           sCity, sState, sZip, sLatitude, sLongitude, sCounty, sParcelID, sExcludeFromAL, _
                           sResidentType, sLegalDesc, sListedOwner, sRegisteredUserID
      	   if sNumber <> "" then
     	   				oLocation("streetnumber") = sNumber
    	     end if
          oLocation("streetprefix")          = sPrefix
   	 	 	 	oLocation("streetaddress")         = sStreetName
          oLocation("streetsuffix")          = sSuffix
          oLocation("streetdirection")       = sDirection
    		 	 	oLocation("city")                  = sCity
    			 	 oLocation("state")                 = sState
     				 oLocation("zip")                   = sZip
          oLocation("county")                = sCounty
          oLocation("parcelidnumber")        = sParcelID
          oLocation("excludefromactionline") = sExcludeFromAL
          oLocation("residenttype")          = sResidentType
          oLocation("legaldescription")      = sLegalDesc
          oLocation("listedowner")           = sListedOwner
          oLocation("registereduserid")      = sRegisteredUserID
      	else
      			'they selected Other address not listed
          BreakOutAddress request("ques_issue2"), sStreetNumber, sStreetName

          'oLocation("streetaddress") = dbsafe(request("ques_issue2"))
          oLocation("streetnumber")          = sStreetNumber
          oLocation("streetaddress")         = sStreetName
      				oLocation("city")                  = sCity
     	 			oLocation("state")                 = sState
     		 		oLocation("zip")                   = sZip
          oLocation("county")                = sCounty
          oLocation("parcelidnumber")        = sParcelID
          oLocation("excludefromactionline") = sExcludeFromAL
          oLocation("residenttype")          = sResidentType
          oLocation("legaldescription")      = sLegalDesc
          oLocation("listedowner")           = sListedOwner
          oLocation("registereduserid")      = sRegisteredUserID
    	  end if
       oLocation("validstreet") = sValidStreet
    else
		     if CLng(request("skip_address")) <> CLng(0) then
		       'Handle the dropdown addresses - These should have the residentaddressid as the selected value
		     			GetAddressInfo request("skip_address"), sNumber, sPrefix, sStreetName, sSuffix, sDirection, sCity, sState, sZip, sLatitude, _
                         sLongitude, sCounty, sParcelID, sExcludeFromAL, sResidentType, sLegalDesc, sListedOwner, sRegisteredUserID
		        oLocation("streetnumber")          = sNumber
          oLocation("streetprefix")          = sPrefix
		        oLocation("streetaddress")         = sStreetName
          oLocation("streetsuffix")          = sSuffix
          oLocation("streetdirection")       = sDirection
		        oLocation("city")                  = sCity
					     oLocation("state")                 = sState
					     oLocation("zip")                   = sZip
          oLocation("validstreet")           = "Y"
          oLocation("county")                = sCounty
          oLocation("parcelidnumber")        = sParcelID
          oLocation("excludefromactionline") = sExcludeFromAL
          oLocation("residenttype")          = sResidentType
          oLocation("legaldescription")      = sLegalDesc
          oLocation("listedowner")           = sListedOwner
          oLocation("registereduserid")      = sRegisteredUserID
	 	    else
		       'they selected Other address not listed
          BreakOutAddress request("ques_issue2"), sStreetNumber, sStreetName
          'oLocation("streetaddress") = dbsafe(request("ques_issue2"))
          oLocation("streetnumber")          = sStreetNumber
          oLocation("streetaddress")         = sStreetName
		        oLocation("city")                  = sCity
		        oLocation("state")                 = sState
		        oLocation("zip")                   = sZip
          oLocation("validstreet")           = "N"
          oLocation("county")                = sCounty
          oLocation("parcelidnumber")        = sParcelID
          oLocation("excludefromactionline") = sExcludeFromAL
          oLocation("residenttype")          = sResidentType
          oLocation("legaldescription")      = sLegalDesc
          oLocation("listedowner")           = sListedOwner
          oLocation("registereduserid")      = sRegisteredUserID
		     end if
 		 end if

 		 if CDbl(sLatitude) <> CDbl("0") then
	 	    oLocation("latitude")   = sLatitude
		     oLocation("longitude")  = sLongitude
		  end if
 else
   'This is for no problem location 
    oLocation("streetnumber")          = sNumber
    oLocation("streetprefix")          = sPrefix
    oLocation("streetaddress")         = sStreetName
    oLocation("streetsuffix")          = sSuffix
    oLocation("streetdirection")       = sDirection
    oLocation("city")                  = sCity
    oLocation("state")                 = sState
    oLocation("zip")                   = sZip
    oLocation("validstreet")           = "N"
    oLocation("county")                = sCounty
    oLocation("parcelidnumber")        = sParcelID
    oLocation("excludefromactionline") = sExcludeFromAL
    oLocation("residenttype")          = sResidentType
    oLocation("legaldescription")      = sLegalDesc
    oLocation("listedowner")           = sListedOwner
    oLocation("registereduserid")      = sRegisteredUserID
 end if

 oLocation("streetunit") = request("streetunit")
	oLocation("comments")   = request("ques_issue6")

'BEGIN: Build the SortStreetName ----------------------------------------------
 sSortStreetName = trim(sStreetName)

 if trim(sSuffix) <> "" then
    if sSortStreetName <> "" then
       sSortStreetName = sSortStreetName & " " & sSuffix
    else
       sSortStreetName = sSuffix
    end if
 end if

 if trim(sDirection) <> "" then
    if sSortStreetName <> "" then
       sSortStreetName = sSortStreetName & " " & sDirection
    else
       sSortStreetName = sDirection
    end if
 end if

 if trim(sPrefix) <> "" then
    if sSortStreetName <> "" then
       sSortStreetName = sSortStreetName & " " & sPrefix
    else
       sSortStreetName = sPrefix
    end if
 end if

 oLocation("sortstreetname") = sSortStreetName
'END: Build the SortStreetName ------------------------------------------------

'If the form has the issue location and the org has the issue location feature and the field is blank then default the value
 sSQL = "SELECT action_form_display_issue "
 sSQL = sSQL & " FROM egov_action_request_forms rf, egov_actionline_requests r "
 sSQL = sSQL & " WHERE r.category_id = rf.action_form_id "
 sSQL = sSQL & " AND r.action_autoid = " & iFormID

 set oDisplayIssue = Server.CreateObject("ADODB.Recordset")
 oDisplayIssue.Open sSQL, Application("DSN"), 3, 1

 if not oDisplayIssue.eof then
    if oDisplayIssue("action_form_display_issue") then
       if oLocation("streetaddress") = "" OR isnull(oLocation("streetaddress")) then
          oLocation("streetaddress") = "Street Address has not been entered."
       end if
    end if
 end if

	oLocation.Update
	iReturnValue = oLocation("rowid")  'This is not needed at this time


 oDisplayIssue.close
	oLocation.close

 set oDisplayIssue = nothing
	set oLocation     = nothing
	set oActionOrg    = nothing 

end sub

'------------------------------------------------------------------------------
Sub MatchAddressInfo( ByVal sResidentStreetNumber, ByVal sResidentStreetName, ByRef sNumber, ByRef sPrefix, ByRef sStreetName, _
                      ByRef sSuffix, ByRef sDirection, ByRef sCity, ByRef sState, ByRef sZip, ByRef sLatitude, ByRef sLongitude, _
                      ByRef sCounty, ByRef sParcelID, ByRef sExcludeFromAL, ByRef sResidentType, ByRef sLegalDesc, ByRef sListedOwner, _
                      ByRef sRegisteredUserID )
	Dim sSql, oAddress

	sSQL = "SELECT residentstreetnumber, residentstreetprefix, residentstreetname, streetsuffix, streetdirection, "
 sSQL = sSQL & " residentcity, residentstate, residentzip, isnull(latitude,0.00) as latitude, isnull(longitude,0.00) as longitude, "
 sSQL = sSQL & " county, parcelidnumber, excludefromactionline, residenttype, legaldescription, listedowner, isnull(registereduserid,0) as registereduserid "
 sSQL = sSQL & " FROM egov_residentaddresses "
	sSQL = sSQL & " WHERE orgid = " & session("orgid")
 sSQL = sSQL & " AND excludefromactionline = 0 "
 sSQL = sSQL & " AND UPPER(residentstreetnumber) = UPPER('" & dbsafe(sResidentStreetNumber) & "') "
 sSQL = sSQL & " AND (residentstreetname = '" & dbsafe(sResidentStreetName) & "' "
 sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix = '" & dbsafe(sResidentStreetName) & "' "
 sSQL = sSQL & " OR residentstreetname + ' ' + streetdirection = '" & dbsafe(sResidentStreetName) & "' "
 sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & dbsafe(sResidentStreetName) & "' "
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname = '" & dbsafe(sResidentStreetName) & "' "
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = '" & dbsafe(sResidentStreetName) & "' "
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetdirection = '" & dbsafe(sResidentStreetName) & "' "
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & dbsafe(sResidentStreetName) & "'"
 sSQL = sSQL & ")"

	set oAddress = Server.CreateObject("ADODB.Recordset")
	oAddress.Open sSQL, Application("DSN"), 0, 1
	
	if NOT oAddress.EOF then
  		sNumber           = oAddress("residentstreetnumber")
    sPrefix           = oAddress("residentstreetprefix")
  		sStreetName       = oAddress("residentstreetname")
    sSuffix           = oAddress("streetsuffix")
    sDirection        = oAddress("streetdirection")
		  sCity             = oAddress("residentcity")
  		sState            = oAddress("residentstate")
		  sZip              = oAddress("residentzip")
  		sLatitude         = oAddress("latitude")
		  sLongitude        = oAddress("longitude")
    sCounty           = oAddress("county")
    sParcelID         = oAddress("parcelidnumber")
    sExcludeFromAL    = oAddress("excludefromactionline")
    sResidentType     = oAddress("residenttype")
    sLegalDesc        = oAddress("legaldescription")
    sListedOwner      = oAddress("listedowner")
    sRegisteredUserID = oAddress("registereduserid")
 else
  		if sResidentStreetNumber <> "" then
    			sStreetName = dbsafe(sResidentStreetNumber) & " " & sResidentStreetName
    else
    			sStreetName = sResidentStreetName
    end if
 end if

	oAddress.Close
	set oAddress = nothing

end sub

'------------------------------------------------------------------------------
Sub GetAddressInfo( ByVal sResidentAddressId, ByRef sNumber, ByRef sPrefix, ByRef sStreetName, ByRef sSuffix, ByRef sDirection, _
                    ByRef sCity, ByRef sState, ByRef sZip, ByRef sLatitude, ByRef sLongitude, ByRef sCounty, ByRef sParcelID, _
                    ByRef sExcludeFromAL, ByRef sResidentType, ByRef sLegalDesc, ByRef sListedOwner, ByRef sRegisteredUserID )
	Dim sSql, oAddress

	sSQL = "SELECT residentstreetnumber, residentstreetprefix, residentstreetname, streetsuffix, streetdirection, "
 sSQL = sSQL & " residentcity, residentstate, residentzip, isnull(latitude,0.00) as latitude, isnull(longitude,0.00) as longitude, "
 sSQL = sSQL & " county, parcelidnumber, excludefromactionline, residenttype, legaldescription, listedowner, isnull(registereduserid,0) as registereduserid "
 sSQL = sSQL & " FROM egov_residentaddresses "
	sSQL = sSQL & " WHERE residentaddressid = " & sResidentAddressId 
 sSQL = sSQL & " AND excludefromactionline = 0 "

	Set oAddress = Server.CreateObject("ADODB.Recordset")
	oAddress.Open sSQL, Application("DSN"), 0, 1
	
	If Not oAddress.EOF Then 
  		sNumber           = oAddress("residentstreetnumber")
    sPrefix           = oAddress("residentstreetprefix")
  		sStreetName       = oAddress("residentstreetname")
    sSuffix           = oAddress("streetsuffix")
    sDirection        = oAddress("streetdirection")
	  	sCity             = oAddress("residentcity")
	  	sState            = oAddress("residentstate")
	  	sZip              = oAddress("residentzip")
	  	sLatitude         = oAddress("latitude")
	  	sLongitude        = oAddress("longitude")
    sCounty           = oAddress("county")
    sParcelID         = oAddress("parcelidnumber")
    sExcludeFromAL    = oAddress("excludefromactionline")
    sResidentType     = oAddress("residenttype")
    sLegalDesc        = oAddress("legaldescription")
    sListedOwner      = oAddress("listedowner")
    sRegisteredUserID = oAddress("registereduserid")
	End If 

	oAddress.close
	Set oAddress = Nothing

end sub

'------------------------------------------------------------------------------
'function AddCommentTaskComment(sInternalMsg,sExternalMsg,sStatus,iFormID,iUserID,iOrgID,sCurrentDate)
'		sSQL = "INSERT egov_action_responses (action_status,action_internalcomment,action_externalcomment,action_editdate,action_userid,action_orgid,action_autoid) VALUES ("
'  sSQL = sSQL & "'" & sStatus              & "', "
'  sSQL = sSQL & "'" & DBsafe(sInternalMsg) & "', "
'  sSQL = sSQL & "'" & DBsafe(sExternalMsg) & "', "
'  sSQL = sSQL & "'" & sCurrentDate         & "', "
'  sSQL = sSQL & "'" & iUserID              & "',"
'  sSQL = sSQL & "'" & iOrgID               & "',"
'  sSQL = sSQL & "'" & iFormID              & "')"

'		set oComment = Server.CreateObject("ADODB.Recordset")
'		oComment.Open sSQL, Application("DSN") , 3, 1
'		set oComment = nothing

'end function

'------------------------------------------------------------------------------
Sub InsertRequestFieldsandResponses( ByVal iRequestID )

	iFieldCount = 0

	'Enumerate fields and entered responses
	For Each oField In Request.Form

		'Get only fields and their associated value
		If Left(oField,10) = "fmquestion" Then 
			iFieldCount = iFieldCount + 1

			'Get field prompt
			sFieldPrompt = request.form("fmname" & replace(oField,"fmquestion",""))
			iFieldID     =  InsertFieldPrompt(sFieldPrompt, _
				iRequestID, _
				request.form("fieldtype")(iFieldCount), _
				request.form("answerlist")(iFieldCount), _
				request.form("isrequired")(iFieldCount), _
				request.form("sequence")(iFieldCount), _
				request.form("pushfieldid")(iFieldCount))

			'Enumerate and get field responses
			For iResponse = 1 To request.form(oField).count
				InsertFieldResponse request.form(oField)(iResponse), _
					iFieldID, _
					request.form("pdfformname")(iFieldCount), _
					request.form("pushfieldid")(iFieldCount)
			Next 
		End If 
	Next 

End Sub 

'------------------------------------------------------------------------------
function InsertFieldPrompt( ByVal sPrompt, ByVal iRequestID, ByVal iFieldType, ByVal sAnswerList, ByVal blnIsRequired, ByVal iSequence, ByVal iPushFieldID )
	Dim oAddFieldPrompt, iReturnValue, sSQL

	iReturnValue = 0 

	If Trim(iSequence) = "" Then 
		iSequence = NULL
	End If 

	If Trim(iPushFieldID) = "" Then 
		lcl_pushfieldid = 0
	Else 
		lcl_pushfieldid = trim(iPushFieldID)
	End If 

	sSQL = "SELECT * FROM egov_submitted_request_fields WHERE 1=2"

	Set oAddFieldPrompt = Server.CreateObject("ADODB.Recordset")
	oAddFieldPrompt.CursorLocation = 3
	oAddFieldPrompt.Open sSQL, Application("DSN"), 1, 3

	'Add new row
	oAddFieldPrompt.AddNew

	oAddFieldPrompt("submitted_request_field_prompt")      = sPrompt
	oAddFieldPrompt("submitted_request_field_type_id")     = iFieldType
	oAddFieldPrompt("submitted_request_field_answerlist")  = sAnswerList
	oAddFieldPrompt("submitted_request_field_isrequired")  = blnIsRequired
	oAddFieldPrompt("submitted_request_field_sequence")    = iSequence
	oAddFieldPrompt("submitted_request_id")                = iRequestID
	oAddFieldPrompt("submitted_request_field_pushfieldid") = lcl_pushfieldid

	'Save added information
	oAddFieldPrompt.Update

	'Set new rowid
	iReturnValue = oAddFieldPrompt("submitted_request_field_id")

	'Close
	oAddFieldPrompt.Close
	Set oAddFieldPrompt = Nothing 

	InsertFieldPrompt = iReturnValue

End Function 

'------------------------------------------------------------------------------
function InsertFieldResponse(sResponse,iFieldID,sPDFName,sPushFieldID)
	  
	 iReturnValue    = 0 
  lcl_response    = sResponse

  if lcl_response <> "" then
     lcl_response = replace(lcl_response,chr(10),"")
     lcl_response = replace(lcl_response,chr(13),"")
  end if

  if trim(sPushFieldID) = "" then
     lcl_pushfieldid = 0
  else
     lcl_pushfieldid = trim(sPushFieldID)
  end if

	 sSQL = "SELECT * FROM egov_submitted_request_field_responses WHERE 1=2"

	 set oAddFieldPrompt = Server.CreateObject("ADODB.Recordset")
  oAddFieldPrompt.CursorLocation = 3
  oAddFieldPrompt.Open sSQL, Application("DSN") , 1,3
	  
 'Add new row
  oAddFieldPrompt.AddNew
 
  oAddFieldPrompt("submitted_request_field_id")        = iFieldID
  oAddFieldPrompt("submitted_request_field_response")  = lcl_response
  oAddFieldPrompt("submitted_request_form_field_name") = sPDFName
  oAddFieldPrompt("submitted_request_pushfieldid")     = lcl_pushfieldid
	 
 'Save added information
  oAddFieldPrompt.Update

 'Close
  oAddFieldPrompt.Close

  InsertFieldResponse = iReturnValue

end function

'------------------------------------------------------------------------------
function getUserName(p_user_id, p_display_order)
 'Display Order is how you want the name to appear on the screen.
 'FL = "First Name" "Last Name"
 'LF = "Last Name", "First Name"

  if p_user_id <> "" then
     sSQL = "SELECT lastname, firstname "
     sSQL = sSQL & " FROM users "
     sSQL = sSQL & " WHERE UserId = " & p_user_id

     Set rs = Server.CreateObject("ADODB.Recordset")
   	 rs.Open sSQL, Application("DSN"), 3, 1

     if not rs.eof then
        if UCASE(p_display_order) = "FL" then
           lcl_name = rs("firstname") & " " & rs("lastname")
        elseif UCASE(p_display_order) = "LF" then
           lcl_name = rs("lastname") & ", " & rs("firstname")
        else
           lcl_name = rs("lastname") & ", " & rs("firstname")
        end if
     else
        lcl_name = ""
     end if

     getUserName = lcl_name

     rs.close
     set rs = nothing
  else
     lcl_name = ""
  end if

end function
%>
<!--#include file="../../egovlink300_global/includes/inc_rye.asp"-->
