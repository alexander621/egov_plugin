<!-- #include file="../includes/common.asp" //-->
<%
'Call displayActionLineWidget(request("orgid"),request("userid"),request("selectDateType"),request("fromdate"),request("todate"))

'------------------------------------------------------------------------------
'sub displayActionLineWidget(p_orgid, p_userid, p_selectDateType, p_fromdate, p_todate)

    p_orgid          = request("orgid")
    p_userid         = request("userid")
    p_selectDateType = request("selectDateType")
    p_fromdate       = request("fromdate")
    p_todate         = request("todate")

    lcl_displayToDate = p_todate
    lcl_queryToDate   = dateAdd("d",1,p_toDate)

   'Get request counts
    lcl_mineSubmitted  = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "MINE", "SUBMITTED")
    lcl_mineInProgress = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "MINE", "INPROGRESS")
    lcl_mineWaiting    = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "MINE", "WAITING")
    lcl_mineResolved   = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "MINE", "RESOLVED")
    lcl_mineDismissed  = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "MINE", "DISMISSED")

    lcl_deptSubmitted  = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "DEPT", "SUBMITTED")
    lcl_deptInProgress = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "DEPT", "INPROGRESS")
    lcl_deptWaiting    = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "DEPT", "WAITING")
    lcl_deptResolved   = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "DEPT", "RESOLVED")
    lcl_deptDismissed  = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "DEPT", "DISMISSED")

    lcl_allSubmitted   = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "ALL", "SUBMITTED")
    lcl_allInProgress  = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "ALL", "INPROGRESS")
    lcl_allWaiting     = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "ALL", "WAITING")
    lcl_allResolved    = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "ALL", "RESOLVED")
    lcl_allDismissed   = getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, lcl_queryToDate, "ALL", "DISMISSED")


    'lcl_mineSubmitted  = 0
    'lcl_mineInProgress = 0
    'lcl_mineWaiting    = 0
    'lcl_mineResolved   = 0
    'lcl_mineDismissed  = 0

    'lcl_deptSubmitted  = 0
    'lcl_deptInProgress = 0
    'lcl_deptWaiting    = 0
    'lcl_deptResolved   = 0
    'lcl_deptDismissed  = 0

    'lcl_allSubmitted   = 0
    'lcl_allInProgress  = 0
    'lcl_allWaiting     = 0
    'lcl_allResolved    = 0
    'lcl_allDismissed   = 0

   'Get the sub-totals and totals
    lcl_subtotal_mine_open = formatnumber(CLng(replace(lcl_mineSubmitted,"&nbsp;",0)) + CLng(replace(lcl_mineInProgress,"&nbsp;",0)) + CLng(replace(lcl_mineWaiting,"&nbsp;",0)),0)
    lcl_subtotal_dept_open = formatnumber(CLng(replace(lcl_deptSubmitted,"&nbsp;",0)) + CLng(replace(lcl_deptInProgress,"&nbsp;",0)) + CLng(replace(lcl_deptWaiting,"&nbsp;",0)),0)
    lcl_subtotal_all_open  = formatnumber(CLng(replace(lcl_allSubmitted,"&nbsp;",0))  + CLng(replace(lcl_allInProgress,"&nbsp;",0))  + CLng(replace(lcl_allWaiting,"&nbsp;",0)),0)

    lcl_subtotal_mine_closed = formatnumber(CLng(replace(lcl_mineResolved,"&nbsp;",0)) + CLng(replace(lcl_mineDismissed,"&nbsp;",0)),0)
    lcl_subtotal_dept_closed = formatnumber(CLng(replace(lcl_deptResolved,"&nbsp;",0)) + CLng(replace(lcl_deptDismissed,"&nbsp;",0)),0)
    lcl_subtotal_all_closed  = formatnumber(CLng(replace(lcl_allResolved,"&nbsp;",0))  + CLng(replace(lcl_allDismissed,"&nbsp;",0)),0)

    lcl_total_mine = CLng(lcl_subtotal_mine_open) + CLng(lcl_subtotal_mine_closed)
    lcl_total_dept = CLng(lcl_subtotal_dept_open) + CLng(lcl_subtotal_dept_closed)
    lcl_total_all  = CLng(lcl_subtotal_all_open)  + CLng(lcl_subtotal_all_closed)

    lcl_display_widget = ""
'    lcl_display_widget = lcl_display_widget & "<fieldset id=""actionlinewidget"" class=""fieldset"">" & vbcrlf
    lcl_display_widget = lcl_display_widget & "  <div id=""widget_request_summary"" style=""text-align:center"">Request Summary&nbsp;</div>" & vbcrlf
    lcl_display_widget = lcl_display_widget & "  <table border=""0"" cellspacing=""0"" cellpadding=""2"" style=""margin-top:5px; margin-left:auto;margin-right:auto;"">" & vbcrlf
    lcl_display_widget = lcl_display_widget & "    <tr>" & vbcrlf
    lcl_display_widget = lcl_display_widget & "        <td align=""center"">" & vbcrlf
    lcl_display_widget = lcl_display_widget & "            <div style=""font-size:11px;""><strong>From: </strong><span style=""color:#800000;"">" & p_fromdate & "</span>&nbsp;<strong>To: </strong><span style=""color:#800000;"">" & p_todate & "</span></div><br />" & vbcrlf
    lcl_display_widget = lcl_display_widget & "            <table border=""0"" cellspacing=""0"" cellpadding=""3"" style=""margin-top:5px; margin-left:5px; border:1pt solid #000000;"">" & vbcrlf
    lcl_display_widget = lcl_display_widget & "              <tr align=""center"">" & vbcrlf
    lcl_display_widget = lcl_display_widget & "                  <td colspan=""4"" style=""font-weight:bold; border-bottom:1pt solid #000000;"" bgcolor=""#93BEE1"">Action Line Requests</td>" & vbcrlf
    lcl_display_widget = lcl_display_widget & "              </tr>" & vbcrlf
    lcl_display_widget = lcl_display_widget & "              <tr align=""center"" bgcolor=""#336699"">" & vbcrlf
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell("Status", "", "ffffff", "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell("Mine",   "", "ffffff", "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell("Dept",   "", "ffffff", "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell("All",    "", "ffffff", "", "", "Y", "1pt", "000000", "N", "",    "")
    lcl_display_widget = lcl_display_widget & "              </tr>" & vbcrlf
    lcl_display_widget = lcl_display_widget & "              <tr align=""center"">" & vbcrlf
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell("Submitted",       "", "800000", "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_mineSubmitted, "", "",       "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_deptSubmitted, "", "",       "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_allSubmitted,  "", "",       "", "", "Y", "1pt", "000000", "N", "",    "")
    lcl_display_widget = lcl_display_widget & "              </tr>" & vbcrlf
    lcl_display_widget = lcl_display_widget & "              <tr align=""center"">" & vbcrlf
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell("In Progress",      "", "800000", "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_mineInProgress, "", "",       "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_deptInProgress, "", "",       "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_allInProgress,  "", "",       "", "", "Y", "1pt", "000000", "N", "",    "")
    lcl_display_widget = lcl_display_widget & "              </tr>" & vbcrlf
    lcl_display_widget = lcl_display_widget & "              <tr align=""center"">" & vbcrlf
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell("Waiting",       "", "800000", "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_mineWaiting, "", "",       "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_deptWaiting, "", "",       "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_allWaiting,  "", "",       "", "", "Y", "1pt", "000000", "N", "",    "")
    lcl_display_widget = lcl_display_widget & "              </tr>" & vbcrlf
    lcl_display_widget = lcl_display_widget & "              <tr align=""center"" bgcolor=""#93BEE1"">" & vbcrlf
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell("Total Open",           "", "ffffff", "Y", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_subtotal_mine_open, "", "ffffff", "Y", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_subtotal_dept_open, "", "ffffff", "Y", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_subtotal_all_open,  "", "ffffff", "Y", "", "Y", "1pt", "000000", "N", "",    "")
    lcl_display_widget = lcl_display_widget & "              </tr>" & vbcrlf

    lcl_display_widget = lcl_display_widget & "              <tr align=""center"">" & vbcrlf
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell("Resolved",       "", "800000", "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_mineResolved, "", "",      "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_deptResolved, "", "",      "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_allResolved,  "", "",      "", "", "Y", "1pt", "000000", "N", "",    "")
    lcl_display_widget = lcl_display_widget & "              </tr>" & vbcrlf
    lcl_display_widget = lcl_display_widget & "              <tr align=""center"">" & vbcrlf
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell("Dismissed",        "", "800000", "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_mineDismissed,  "", "",     "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_deptDismissed,  "", "",     "", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_allDismissed,   "", "",     "", "", "Y", "1pt", "000000", "N", "",    "")
    lcl_display_widget = lcl_display_widget & "              </tr>" & vbcrlf
    lcl_display_widget = lcl_display_widget & "              <tr align=""center"" bgcolor=""#93BEE1"">" & vbcrlf
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell("Total Closed",           "", "ffffff", "Y", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_subtotal_mine_closed, "", "ffffff", "Y", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_subtotal_dept_closed, "", "ffffff", "Y", "", "Y", "1pt", "000000", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_subtotal_all_closed,  "", "ffffff", "Y", "", "Y", "1pt", "000000", "N", "",    "")
    lcl_display_widget = lcl_display_widget & "              </tr>" & vbcrlf
    lcl_display_widget = lcl_display_widget & "              <tr align=""center"" bgcolor=""#336699"">" & vbcrlf
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell("Grand Total",  "", "ffffff", "Y", "", "N", "", "", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_total_mine, "", "ffffff", "Y", "", "N", "", "", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_total_dept, "", "ffffff", "Y", "", "N", "", "", "Y", "1pt", "000000")
    lcl_display_widget = lcl_display_widget &                    displayActionLineWidgetCell(lcl_total_all,  "", "ffffff", "Y", "", "N", "", "", "N", "",    "")
    lcl_display_widget = lcl_display_widget & "              </tr>" & vbcrlf
    lcl_display_widget = lcl_display_widget & "            </table>" & vbcrlf
    lcl_display_widget = lcl_display_widget & "        </td>" & vbcrlf
    lcl_display_widget = lcl_display_widget & "    </tr>" & vbcrlf
    'response.write "  <tr>" & vbcrlf
    'response.write "      <td style=""font-size:11px;"">" & vbcrlf
    'response.write "          <center style=""color:#800000; font-weight:bold;"">STATUSES</center>" & vbcrlf
    'response.write "          <strong>NEW</strong> - Submitted status (NOT included in Totals)<br />" & vbcrlf
    'response.write "          <strong>OPEN</strong> - Submitted, In Progress, and Waiting statuses<br />" & vbcrlf
    'response.write "          <strong>CLOSED</strong> - Resolved and Dismissed statuses" & vbcrlf
    'response.write "      </td>" & vbcrlf
    'response.write "  </tr>" & vbcrlf
    lcl_display_widget = lcl_display_widget & "  </table>" & vbcrlf
    'lcl_display_widget = lcl_display_widget & "  <input type=""button"" name=""hideWidgetButton"" id=""hideWidgetButton"" value=""Hide Results"" class=""button"" onclick=""hideWidgetResults();"" />" & vbcrlf
'    lcl_display_widget = lcl_display_widget & "</fieldset>" & vbcrlf

    response.write lcl_display_widget


'end sub

'------------------------------------------------------------------------------
function getRequestCount(p_orgid, p_userid, p_selectDateType, p_fromdate, p_todate, p_ownership, p_status)

  lcl_return = 0

 'p_ownership represents the columns in the grid (who is the data limited to).  MINE, DEPT, ALL
 'p_status represents the rows in the grid (what status is the data limited to).  NEW, OPEN, CLOSED
  if p_ownership <> "" AND p_status <> "" then
    'Determine which dates to search on

     if UCASE(p_selectDateType) = "ACTIVE" then
        varRequestCntClause = " AND ("
        varRequestCntClause = varRequestCntClause & " (submit_date >= '" & p_fromdate & "' AND submit_date < '" & p_todate & "') OR "
        varRequestCntClause = varRequestCntClause & " ( IsNull(complete_date,'" & Now & "') >= '" & p_fromdate & "' AND IsNull(complete_date,'" & Now & "') < '" & p_todate & "' ) OR "
        varRequestCntClause = varRequestCntClause & " (submit_date < '" & p_fromdate & "' AND IsNull(complete_date,'" & Now & "') > '" & p_todate & "')) "
     else 'selectDateType = SUBMIT
        varRequestCntClause = " AND (submit_date BETWEEN '" & p_fromdate & "' AND '" & p_todate & "') "
     end if

    'Build the SQL query.
     sSQL = "SELECT count(action_autoid) AS total_requests "
     sSQL = sSQL & " FROM egov_action_request_view "
     sSQL = sSQL & " WHERE orgid = " & p_orgid
     sSQL = sSQL & varRequestCntClause

    'Build the SQL statement for the p_ownership
     if UCASE(p_ownership) = "MINE" then
        sSQL = sSQL & " AND assignedemployeeid = " & p_userid
     elseif UCASE(p_ownership) = "DEPT" then
        sSQL = sSQL & " AND deptid IN (select distinct ug.groupid "
        sSQL = sSQL &                " from usersgroups ug, groups g "
        sSQL = sSQL &                " where ug.groupid = g.groupid "
        sSQL = sSQL &                " and ug.userid = " & p_userid & ") "
     else
       'For the StatusSummary report the "p_ownership" value is the GROUP BY value + the ID.
        'if instr(UCASE(p_ownership),"SUBMIT_DATE")       > 0 then
        if instr(UCASE(p_ownership),"SUBMITDATESHORT")       > 0 then
           lcl_columnvalue = replace(UCASE(p_ownership),"SUBMITDATESHORT","")
        elseif instr(UCASE(p_ownership),"STREETNAME")    > 0 then
           lcl_columnvalue = replace(UCASE(p_ownership),"STREETNAME","")
        elseif instr(UCASE(p_ownership),"ACTION_FORMID") > 0 then
           lcl_columnvalue = replace(UCASE(p_ownership),"ACTION_FORMID","")
        elseif instr(UCASE(p_ownership),"DEPTID")        > 0 then
           lcl_columnvalue = replace(UCASE(p_ownership),"DEPTID","")
        elseif instr(UCASE(p_ownership),"ASSIGNED_NAME") > 0 then
           lcl_columnvalue = replace(UCASE(p_ownership),"ASSIGNED_NAME","")
        elseif instr(UCASE(p_ownership),"SUBMITTEDBY")   > 0 then
           lcl_columnvalue = replace(UCASE(p_ownership),"SUBMITTEDBY","")
        end if

        if UCASE(p_ownership) = "SUBMITTEDBY" then
           lcl_columnname = "assigned_userid"
        else
           lcl_columnname = replace(p_ownership,lcl_columnvalue,"")
        end if

        if UCASE(lcl_columnname) <> "ALL" AND UCASE(lcl_columnvalue) <> "ALL" then
           sSQL = sSQL & " AND UPPER(" & lcl_columnname & ") = '" & UCASE(lcl_columnvalue) & "' "
        end if

     end if

    'Build the SQL statement for the p_status
     sSQL = sSQL & " AND UPPER(status) = '" & UCASE(p_status) & "' "
     'if UCASE(p_status) = "NEW" then
     '   sSQL = sSQL & " AND UPPER(status) = 'SUBMITTED' "
     'elseif UCASE(p_status) = "OPEN" then
     '   sSQL = sSQL & " AND UPPER(status) IN ('SUBMITTED','INPROGRESS','WAITING') "
     'elseif UCASE(p_status) = "CLOSED" then
     '   sSQL = sSQL & " AND UPPER(status) IN ('RESOLVED','DISMISSED') "
     'end if
'dtb_debug(sSQL)
   		set oWidget = Server.CreateObject("ADODB.Recordset")
   		oWidget.Open sSQL, Application("DSN"), 3, 1

     if not oWidget.eof then
        lcl_return = formatnumber(oWidget("total_requests"),0)
     end if

     oWidget.close
     set oWidget = nothing
  end if

  getRequestCount = lcl_return

end function

'------------------------------------------------------------------------------
function displayActionLineWidgetCell(p_text, p_colspan, p_textColor, p_textBold, p_BGColor, _
                                     p_borderBottom, p_borderBottomSize, p_borderBottomColor, _
                                     p_borderRight, p_borderRightSize, p_borderRightColor)

  lcl_return            = ""
  lcl_text              = "-"

  lcl_rowStyle          = ""
  lcl_colspan           = ""
  lcl_textColor         = "color:#000000;"
  lcl_textBold          = ""
  lcl_BGColor           = ""

  lcl_borderBottom      = ""
  lcl_borderBottomSize  = "1pt"
  lcl_borderBottomColor = "000000"

  lcl_borderRight      = ""
  lcl_borderRightSize  = "1pt"
  lcl_borderRightColor = "000000"

 'Text
  if p_text <> "" then
     lcl_text = p_text
  end if

 'Colspan
  if p_colspan <> "" then
     lcl_colspan = " colspan=""" & p_colspan & """"
  end if

 'Text Color
  if replace(p_textColor,"#","") <> "" then
     lcl_textColor = "color:#" & replace(p_textColor,"#","") & ";"
  end if

 'Text Bold
  if p_textBold <> "" then
     lcl_textBold = "font-weight:bold;"
  end if

 'Background Color
  if replace(p_BGColor,"#","") <> "" then
     lcl_BGColor = "background-color:#" & replace(p_BGColor,"#","") & ";"
  end if

 'Border Bottom
  if p_borderBottom = "Y" then
     if p_borderBottomSize <> "" then
        lcl_borderBottomSize = p_borderBottomSize
     end if

     if replace(p_borderBottomColor,"#","") <> "" then
        lcl_borderBottomColor = replace(p_borderBottomColor,"#","")
     end if

     lcl_borderBottom = "border-bottom:" & lcl_borderBottomSize & " solid #" & lcl_borderBottomColor & ";"
  end if

 'Border Right
  if p_borderRight = "Y" then
     if p_borderRightSize <> "" then
        lcl_borderRightSize = p_borderRightSize
     end if

     if replace(p_borderRightColor,"#","") <> "" then
        lcl_borderRightColor = replace(p_borderRightColor,"#","")
     end if

     lcl_borderRight = "border-right:" & lcl_borderRightSize & " solid #" & lcl_borderRightColor & ";"
  end if

  lcl_return = "<td style=""" & lcl_textColor & lcl_textBold & lcl_BGColor & lcl_borderBottom & lcl_borderRight & """" & lcl_colspan & ">" & lcl_text & "</td>" & vbcrlf

  displayActionLineWidgetCell = lcl_return

end function
%>
