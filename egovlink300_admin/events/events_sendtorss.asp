<!-- #include file="../includes/common.asp" //-->
<%
 lcl_id = 0

 if request("id") <> "" then
    if isnumeric(request("id")) then
       lcl_id = request("id")
    end if
 end if

 if request("rssType") <> "" then
    lcl_rssType = UCASE(request("rssType"))
 else
    lcl_rssType = "COMMUNITYCALENDAR"
 end if

 if request("isAjax") <> "" then
    lcl_isAjax = UCASE(request("isAjax"))
 else
    lcl_isAjax = "N"
 end if

 if lcl_id > 0 then
    setupSendToRSS lcl_id, lcl_rssType, lcl_isAjax
 else
    if lcl_isAjax = "Y" then
       response.write "Failed to send to RSS - Error in AJAX Routine"
    else
       response.write "default.asp?cal=" & lcl_rssType & "&success=AJAX_ERROR"
    end if
 end if

'------------------------------------------------------------------------------
sub setupSendToRSS(iID, iRSSType, iIsAjax)

 'Get the FeedID
  lcl_feedid = getFeedIDByFeedName(iRSSType)

  if CLng(lcl_feedid) > CLng(0) then
    'Get the info needed to create the RSS record
     sSQL = "SELECT e.orgid, e.eventdate, e.eventtimezoneid, e.subject, e.message, e.categoryid, c.categoryname "
     sSQL = sSQL & " FROM events e LEFT OUTER JOIN eventcategories c ON e.categoryid = c.categoryid "
     sSQL = sSQL & " WHERE e.eventid = " & iID

    	set oEventList = Server.CreateObject("ADODB.Recordset")
   	 oEventList.Open sSQL, Application("DSN"), 3, 1

     if not oEventList.eof then
        if oEventList("categoryname") <> "" then
           lcl_categoryname = " - (" & oEventList("categoryname") & ")"
        else
           lcl_categoryname = ""
        end if

        lcl_rssOrgID      = session("orgid")
        lcl_rssType       = iRSSType
        lcl_rssRowID      = iID
        lcl_rssTitle      = replace(cdate(oEventList("eventdate")) & " - " & oEventList("subject") & lcl_categoryname,"&","<<AMP>>")
        lcl_rssDesc       = replace(oEventList("message"),"&","<<AMP>>")
        lcl_rssLink       = "/events/calendarevents.asp?date=" & month(oEventList("eventdate")) & "-" & day(oEventList("eventdate")) & "-" & year(oEventList("eventdate"))
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

     oEventList.close
     set oEventList = nothing
  else
     lcl_success   = "RSS_ERROR"
     lcl_isAjaxMsg = "ERROR: Failed to send to RSS..."
  end if

  if iIsAjax = "Y" then
     response.write lcl_isAjaxMsg
  else
     response.redirect "default.asp?cal=" & iRSSType & "&success=" & lcl_success
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