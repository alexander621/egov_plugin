<!-- #include file="../includes/common.asp" //-->
<%
 lcl_id = 0

 if request("id") <> "" then
    if isnumeric(request("id")) then
       lcl_id = request("id")
    end if
 end if

 if request("faqtype") <> "" then
    lcl_faqtype = UCASE(request("faqtype"))
 else
    lcl_faqtype = "FAQ"
 end if

 if request("isAjax") <> "" then
    lcl_isAjax = UCASE(request("isAjax"))
 else
    lcl_isAjax = "N"
 end if

 if lcl_id > 0 then
    setupSendToRss lcl_id, lcl_faqtype, lcl_isAjax
 else
    if lcl_isAjax = "Y" then
       response.write "Failed to send to RSS - Error in AJAX Routine"
    else
       response.write "list_faq.asp?faqtype=" & lcl_faqtype & "&success=AJAX_ERROR"
    end if
 end if

'------------------------------------------------------------------------------
sub setupSendToRSS(iID, iFAQType, iIsAjax)

 'Get the FeedID
  lcl_feedid = getFeedIDByFeedName(iFAQType)

  if CLng(lcl_feedid) > CLng(0) then
    'Get the info needed to create the RSS record
     sSQL = "SELECT f.orgid, f.faqQ, f.faqA, c.faqcategoryname "
     sSQL = sSQL & " FROM faq f LEFT OUTER JOIN faq_categories c ON f.faqcategoryid = c.faqcategoryid "
     sSQL = sSQL & " WHERE f.faqid = " & iID

    	set oFaqList = Server.CreateObject("ADODB.Recordset")
   	 oFaqList.Open sSQL, Application("DSN"), 3, 1

     if not oFaqList.eof then

        if oFaqList("faqcategoryname") <> "" then
           lcl_faqcategoryname = " - (" & oFaqList("faqcategoryname") & ")"
        else
           lcl_faqcategoryname = ""
        end if

        lcl_rssOrgID      = session("orgid")
        lcl_rssType       = iFAQType
        lcl_rssRowID      = iID
        lcl_rssTitle      = replace(oFaqList("faqQ") & lcl_faqcategoryname,"&","<<AMP>>")
        lcl_rssDesc       = replace(oFaqList("faqA"),"&","<<AMP>>")
        lcl_rssLink       = "/faq.asp?faqtype=" & iFAQType
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

     oFaqList.close
     set oFaqList = nothing
  else
     lcl_success   = "RSS_ERROR"
     lcl_isAjaxMsg = "ERROR: Failed to send to RSS..."
  end if

  if iIsAjax = "Y" then
     response.write lcl_isAjaxMsg
  else
     response.redirect "list_faq.asp?faqtype=" & iFAQType & "&success=" & lcl_success
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