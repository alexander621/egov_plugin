<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
  Response.AddHeader "X-XSS-Protection","0"
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: preview_email.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an admin to preview the email body of the email subscription they are preparing to send out.
'
' MODIFICATION HISTORY
' 1.0 06/22/2011 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("subscriptions,job_postings,bid_postings") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

'Retrieve variables
 lcl_list_type    = ""
 lcl_email_body   = ""
 lcl_containsHTML = ""

 if request("listtype") <> "" then
    lcl_list_type = UCASE(request("listtype"))
 end if

 if request("emailbody") <> "" then
    lcl_email_body = request("emailbody")
 end if

 if request("containsHTML") <> "" then
    lcl_containsHTML = request("containsHTML")
 end if

'Retrieve the email subject and body
' sSQL = "SELECT email_body, containsHTML "
' sSQL = sSQL & " FROM egov_class_distributionlist_log "
' sSQL = sSQL & " WHERE orgid = "  & session("orgid")
' sSQL = sSQL & " AND dl_logid = " & lcl_dl_logid

'	set oLogEmail = Server.CreateObject("ADODB.Recordset")
' oLogEmail.Open sSQL, Application("DSN"), 3, 1
	
'	if not oLogEmail.eof then
'    lcl_email_body   = oLogEmail("email_body")

'    if oLogEmail("containsHTML") then
'       lcl_containsHTML = "Y"
'    else
'       lcl_containsHTML = "N"
'    end if
' else
'    response.redirect("subscrptionslog_list.asp?success=NE")
' end if

' oLogEmail.close
' set oLogEmail = nothing


 if lcl_containsHTML = "" OR lcl_containsHTML = "N" then
    lcl_email_body = replace(lcl_email_body,chr(10),vbcrlf)
 end if

 lcl_email_body = formatToFitEmailLineLength(lcl_email_body)
 lcl_showemail  = BuildHTMLMessage(lcl_email_body, lcl_containsHTML)

 response.write lcl_showemail
%>