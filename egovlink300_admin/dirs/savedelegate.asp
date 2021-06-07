<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: saveDelegate.asp
' AUTHOR: David Boyer
' CREATED: 07/31/2009
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description: Saves the delegateid for the user.
'
' MODIFICATION HISTORY
' 1.0  07/31/09 	David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 lcl_success = "N"
 sOrgID      = 0
 sUserID     = 0
 sDelegateID = 0
 
 if request("orgid") <> "" then
    sOrgID = request("orgid")
 end if

 if request("userid") <> "" then
    sUserID = request("userid")
 end if

 if request("delegateid") <> "" then
    sDelegateID = request("delegateid")
 end if

 if request("isAjaxRoutine") = "Y" then
    lcl_isAjaxRoutine = True
 else
    lcl_isAjaxRoutine = False
 end if

 sSQL = "UPDATE users SET delegateid = " & sDelegateID
 sSQL = sSQL & " WHERE orgid = " & sOrgID
 sSQL = sSQL & " AND userid = "  & sUserID

 set oUpdateDelegate = Server.CreateObject("ADODB.Recordset")
 oUpdateDelegate.Open sSQL, Application("DSN"), 3, 1

 set oUpdateDelegate = nothing

'BEGIN: Send notification to delegate -----------------------------------------
 lcl_user_name              = ""
 lcl_user_email             = ""
 lcl_delegate_name          = ""
 lcl_delegate_email         = ""
 lcl_actionline_featurename = GetFeatureName("action line")

'Get the user info
 sSQL = "SELECT firstname, lastname, email "
 sSQL = sSQL & " FROM users "
 sSQL = sSQL & " WHERE userid = " & sUserID

 set oUserInfo = Server.CreateObject("ADODB.Recordset")
 oUserInfo.Open sSQL, Application("DSN"), 3, 1

 if not oUserInfo.eof then
    lcl_user_name  = trim(oUserInfo("firstname") & " " & oUserInfo("lastname"))
    lcl_user_email = trim(oUserInfo("email"))
 end if

 oUserInfo.close
 set oUserInfo = nothing

'Get the delegate info
 sSQL = "SELECT firstname, lastname, email "
 sSQL = sSQL & " FROM users "
 sSQL = sSQL & " WHERE userid = " & sDelegateID

 set oGetDelegateInfo = Server.CreateObject("ADODB.Recordset")
 oGetDelegateInfo.Open sSQL, Application("DSN"), 3, 1

 if not oGetDelegateInfo.eof then
    lcl_delegate_name  = trim(oGetDelegateInfo("firstname") & " " & oGetDelegateInfo("lastname"))
    lcl_delegate_email = trim(oGetDelegateInfo("email"))

   'Build delegate email message
    lcl_msg = lcl_delegate_name & ",<br /><br />" & vbcrlf
    lcl_msg = lcl_msg & lcl_user_name & " has assigned you as his/her delegate within the E-Gov application.  As a delegate"
    lcl_msg = lcl_msg & " you will receive " & lcl_user_name & "'s emails related to " & lcl_actionline_featurename
    lcl_msg = lcl_msg & " such as: requests, email reminders, assignment notifications from newly submitted requests, alert notifications, "
    lcl_msg = lcl_msg & " etc.<br /><br />" & vbcrlf

    lcl_email_from   = session("sOrgName") & " (E-Gov Website) <noreplies@egovlink.com>"
    lcl_subject      = "Delegate Assignment: " & lcl_user_name & " has made you a delegate."
    lcl_message      = BuildHTMLMessage(lcl_msg,"Y")

   'Setup the SENDTO and check for a DELEGATE
    setupSendToAndDelegateEmails lcl_user_email, lcl_delegate_email, lcl_email_sendto, lcl_email_cc

   'Send a notification to the delegate and assignee
    sendEmail lcl_email_from,lcl_email_sendto,lcl_email_cc,lcl_subject,lcl_message,"","Y"
 end if

 oGetDelegateInfo.close
 set oGetDelegateInfo = nothing

 if lcl_isAjaxRoutine then
    if sDelegateID = 0 then
       response.write "Delegate has been unassigned"
    else
       response.write "Delegate has been assigned"
    end if
 end if
%>
