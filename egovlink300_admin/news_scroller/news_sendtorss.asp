<!-- #include file="../includes/common.asp" //-->
<%
 lcl_id = 0

 if request("id") <> "" then
    if isnumeric(request("id")) then
       lcl_id = request("id")
    end if
 end if

 lcl_feedname = "NEWS"

 if request("isAjax") <> "" then
    lcl_isAjax = UCASE(request("isAjax"))
 else
    lcl_isAjax = "N"
 end if

 if lcl_id > 0 then
    setupSendToRss lcl_id, lcl_feedname, lcl_isAjax
 else
    if lcl_isAjax = "Y" then
       response.write "Failed to send to RSS - Error in AJAX Routine"
    else
       response.write "list_items.asp?success=AJAX_ERROR"
    end if
 end if

'------------------------------------------------------------------------------
sub setupSendToRSS(iID, iFeedName, iIsAjax)

 'Get the FeedID
  lcl_feedid = getFeedIDByFeedName(iFeedName)

  if CLng(lcl_feedid) > CLng(0) then
    'Get the info needed to create the RSS record
     sSQL = "SELECT newsitemid, orgid, itemtitle, itemdate, itemtext, itemlinkurl "
     sSQL = sSQL & " FROM egov_news_items "
     sSQL = sSQL & " WHERE newsitemid = " & iID

    	set oNewsList = Server.CreateObject("ADODB.Recordset")
   	 oNewsList.Open sSQL, Application("DSN"), 3, 1

     if not oNewsList.eof then

        if oNewsList("itemlinkurl") <> "" then
           lcl_itemlinkurl = " [" & oNewsList("itemlinkurl") & "]"
        else
           lcl_itemlinkurl = ""
        end if

        lcl_rssOrgID      = session("orgid")
        lcl_rssType       = iFeedName
        lcl_rssRowID      = iID
        lcl_rssTitle      = replace(oNewsList("itemtitle"),"&","<<AMP>>")
        lcl_rssDesc       = replace(oNewsList("itemtext") & lcl_itemlinkurl,"&","<<AMP>>")
        lcl_rssLink       = "/news/news.asp"
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

     oNewsList.close
     set oNewsList = nothing
  else
     lcl_success   = "RSS_ERROR"
     lcl_isAjaxMsg = "ERROR: Failed to send to RSS..."
  end if

  if iIsAjax = "Y" then
     response.write lcl_isAjaxMsg
  else
     response.redirect "list_items.asp?success=" & lcl_success
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