<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="subscriptionslog_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: subscriptionslog_viewemail.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an admin to modify a RSS Feed
'
' MODIFICATION HISTORY
' 1.0 06/30/09 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("subscriptionslog_maint") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if NOT userhaspermission(session("userid"), "subscriptionslog_maint") then
	   response.redirect sLevel & "permissiondenied.asp"
 end if

 if request("listtype") <> "" then
    lcl_list_type = UCASE(request("listtype"))
 else
    lcl_list_type = ""
 end if

'Retrieve the dl_logid of the subscription log record
'If no value exists then redirect them back to the main results screen
 if request("dl_logid") <> "" then
    lcl_dl_logid = request("dl_logid")
 else
    response.redirect "subscriptionslog_list.asp?listtype=" & lcl_list_type
 end if

'Retrieve the email subject and body
 sSQL = "SELECT email_body, containsHTML "
 sSQL = sSQL & " FROM egov_class_distributionlist_log "
 sSQL = sSQL & " WHERE orgid = "  & session("orgid")
 sSQL = sSQL & " AND dl_logid = " & lcl_dl_logid

	set oLogEmail = Server.CreateObject("ADODB.Recordset")
 oLogEmail.Open sSQL, Application("DSN"), 3, 1
	
	if not oLogEmail.eof then
    lcl_email_body   = oLogEmail("email_body")

    if oLogEmail("containsHTML") then
       lcl_containsHTML = "Y"
    else
       lcl_containsHTML = "N"
    end if
 else
    response.redirect("subscrptionslog_list.asp?success=NE")
 end if

 oLogEmail.close
 set oLogEmail = nothing

 lcl_showemail = BuildHTMLMessage(lcl_email_body, lcl_containsHTML)

 response.write lcl_showemail
%>