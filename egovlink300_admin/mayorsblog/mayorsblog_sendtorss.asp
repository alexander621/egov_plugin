<!-- #include file="../includes/common.asp" //-->
<%
 lcl_id = 0

 if request("id") <> "" then
    if isnumeric(request("id")) then
       lcl_id = request("id")
    end if
 end if

 if request("isAjax") <> "" then
    lcl_isAjax = UCASE(request("isAjax"))
 else
    lcl_isAjax = "N"
 end if

 if lcl_id > 0 then
    setupSendToRss lcl_id, "MAYORSBLOG", lcl_isAjax
 else
    if lcl_isAjax = "Y" then
       response.write "Failed to send to RSS - Error in AJAX Routine"
    else
       response.write "mayorsblog_list.asp?success=AJAX_ERROR"
    end if
 end if

'------------------------------------------------------------------------------
sub setupSendToRSS(iID, iFeedName, iIsAjax)

 'Get the FeedID
  lcl_feedid = getFeedIDByFeedName(iFeedName)

  if CLng(lcl_feedid) > CLng(0) then
    'Get the info needed to create the RSS record
     sSQL = "SELECT mb.orgid, mb.userid, u.firstname + ' ' + u.lastname AS blogowner, mb.title, mb.article, mb.createdbyid, "
     sSQL = sSQL & " mb.createdbydate, mb.isInactive "
     sSQL = sSQL & " FROM egov_mayorsblog mb "
     sSQL = sSQL &      " LEFT OUTER JOIN users u ON mb.userid = u.userid AND u.orgid = " & session("orgid")
     sSQL = sSQL & " WHERE mb.blogid = " & iID

    	set oRSSData = Server.CreateObject("ADODB.Recordset")
   	 oRSSData.Open sSQL, Application("DSN"), 3, 1

     if not oRSSData.eof then
        if trim(oRSSData("blogowner")) <> "" then
           lcl_blogowner = " - (" & oRSSData("blogowner") & ")"
        else
           lcl_blogowner = ""
        end if

        lcl_rssOrgID      = session("orgid")
        lcl_rssType       = iFeedName
        lcl_rssRowID      = iID
        lcl_rssTitle      = replace(oRSSData("title") & lcl_blogowner,"&","<<AMP>>")
        lcl_rssDesc       = replace(oRSSData("article"),"&","<<AMP>>")
        lcl_rssLink       = "/mayorsblog/mayorsblog.asp"
        lcl_rssPubDate    = ConvertDateTimetoTimeZone()
        lcl_createdByID   = session("userid")
        lcl_createdByUser = GetAdminName(lcl_createdByID)

        sendToRSS lcl_feedid, lcl_rssOrgID, lcl_rssRowID, lcl_rssTitle, lcl_rssDesc, lcl_rssLink, lcl_rssPubDate, lcl_createdByID, lcl_createdByUser

        lcl_success   = "RSS_SUCCESS"
        lcl_isAjaxMsg = "Successfully Sent to RSS"
     else
        lcl_success   = "RSS_ERROR"
        lcl_isAjaxMsg = "ERROR: Failed to send to RSS..."
     end if

     oRSSData.close
     set FaqList = nothing
  else
     lcl_success   = "RSS_ERROR"
     lcl_isAjaxMsg = "ERROR: Failed to send to RSS..."
  end if

  if iIsAjax = "Y" then
     response.write lcl_isAjaxMsg
  else
     response.redirect "mayorsblog_list.asp?success=" & lcl_success
  end if

end sub

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  if p_value <> "" then
     sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
    	set oDTB = Server.CreateObject("ADODB.Recordset")
   	 oDTB.Open sSQL, Application("DSN"), 3, 1
  end if

end sub
%>