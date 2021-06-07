<!-- #include file="../includes/common.asp" //-->
<%
Call updateSubscriptionLog(request("user_action"),request("dl_logid"), request("listtype"))

'------------------------------------------------------------------------------
sub updateSubscriptionLog(iAction, iDL_LogID, iListType)

 if iDL_LogID <> "" then
    sdl_logid = CLng(iDL_LogID)
 else
    sdl_logid = 0
 end if

'BEGIN: Delete the Subscriptions Log File. ---------------------------------
 if iAction = "DELETE" then

    sSQL = "DELETE FROM egov_class_distributionlist_log "
    sSQL = sSQL & " WHERE dl_logid = " & sdl_logid

  	 set oDeleteLog = Server.CreateObject("ADODB.Recordset")
 	  oDeleteLog.Open sSQL, Application("DSN"), 3, 1

    set oDeleteRSSOrgTitles = nothing
    set oDeleteRSSItems     = nothing
    set oDeleteLog      = nothing

    lcl_redirect_url = "subscriptionslog_list.asp?success=SD&listtype=" & iListType

 end if
'END: Delete the RSS Feed. -------------------------------------------------

 response.redirect lcl_redirect_url

end sub

'------------------------------------------------------------------------------
function DBsafe( strDB )
 	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
 	DBsafe = Replace( strDB, "'", "''" )
end function
%>