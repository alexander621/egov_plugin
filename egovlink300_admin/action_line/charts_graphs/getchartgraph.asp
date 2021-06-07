<!-- #include file="../../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: getchartgraph.asp
' AUTHOR: David Boyer
' CREATED: 09/03/2010
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Inserts a record into egov_charts and returns the identity value for the new row
'
' MODIFICATION HISTORY
' 1.0 09/03/10  David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 lcl_success             = "Y"
 lcl_return              = ""
 lcl_orgid               = request("orgid")
 lcl_userid              = request("userid")
 lcl_fromdate            = ""
 lcl_todate              = ""
 lcl_includedates        = false
 lcl_selectedchart       = ""
 lcl_chartquery          = ""
 lcl_charttitle          = ""
 lcl_charttype           = ""
 lcl_pageurl             = "piechart"
 lcl_chartwidth          = "700"
 lcl_chartheight         = "500"
 lcl_showlegend          = 0
 lcl_legendtitle         = "NULL"
 lcl_collectedthreshold  = "5"
 lcl_collectedlabel      = "NULL"
 lcl_collectedlegendtext = "NULL"
 lcl_isAjaxRoutine       = "Y"

 if request("fromDate") <> "" then
    lcl_fromdate = request("fromDate")
 end if

 if request("toDate") <> "" then
    lcl_todate = request("toDate")
 end if

'If empty then default to the current date
 if lcl_toDate = "" OR IsNull(lcl_toDate) then
    lcl_toDate = today
 end if

 if lcl_fromDate = "" OR IsNull(lcl_fromDate) then
    lcl_fromDate = cdate(Month(today)& "/1/" & Year(today))
 end if

 if request("includedates") = "Y" then
    lcl_includedates = true
 end if

 if request("charttype") <> "" then
    lcl_charttype = request("charttype")
 else
    lcl_charttype = "pie"
 end if

 if request("selectedchart") <> "" then
    lcl_selectedchart = request("selectedchart")

   'Build SQL where clause
    varWhereClause = " WHERE ([Date Submitted] >= '" & lcl_fromdate & "' AND [Date Submitted] <= '" & DateAdd("d",1,lcl_todate) & "') "
    varWhereClause = varWhereClause & " AND orgid='" & lcl_orgid & "'"
    sOpenOnly      = " AND (UPPER(status) <> 'DISMISSED' AND UPPER(status) <> 'RESOLVED')"

    if lcl_selectedchart = "1" then
       lcl_charttitle = "Monthly Status"
       'lcl_charttype  = "bar"

       'lcl_chartquery = "SELECT "
       'lcl_chartquery = lcl_chartquery & " status as seriesname, "
       'lcl_chartquery = lcl_chartquery & " LEFT(DATENAME(MONTH,right(yearmonth,2) + '/1/' + LEFT(yearmonth,4)),3) + ' ' + LEFT(yearmonth,4) as xvalue, "
       'lcl_chartquery = lcl_chartquery & " sum(INPROGRESS) as yvalue "
       'lcl_chartquery = lcl_chartquery & " FROM egov_rpt_status_chart "
       'lcl_chartquery = lcl_chartquery &   varWhereClause
       'lcl_chartquery = lcl_chartquery & " GROUP BY yearmonth, status "
    elseif lcl_selectedchart = "2" then
       lcl_charttitle = "Monthly Status by Department"
    elseif lcl_selectedchart = "3" then
       lcl_charttitle = "Most Submitted Requests"
       'lcl_charttype  = "pie"

       lcl_chartquery = "SELECT "
       lcl_chartquery = lcl_chartquery & " [Form Name] as xvalue, "
       lcl_chartquery = lcl_chartquery & " Count([Form Name]) as yvalue "
       lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
       lcl_chartquery = lcl_chartquery &   varWhereClause
       lcl_chartquery = lcl_chartquery & " GROUP BY action_formid, [Form Name] "
       lcl_chartquery = lcl_chartquery & " ORDER BY COUNT([Form Name]) DESC "

    elseif lcl_selectedchart = "4" then
       lcl_charttitle = "Open Items Activity by Form"
       'lcl_charttype  = "column"

       lcl_chartquery = ""
       lcl_chartquery = lcl_chartquery & " SELECT "
       lcl_chartquery = lcl_chartquery & " 'Avg Days Open' as seriesname, "
       lcl_chartquery = lcl_chartquery & " isnull([FORM NAME],'N/A') as xvalue, "
       lcl_chartquery = lcl_chartquery & " AVG([Open]) as yvalue, "
       lcl_chartquery = lcl_chartquery & " 0 as seriesorder "
       lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
       lcl_chartquery = lcl_chartquery &   varWhereClause
       lcl_chartquery = lcl_chartquery &   sOpenOnly
       lcl_chartquery = lcl_chartquery & " GROUP BY isnull([FORM NAME],'N/A') "
       lcl_chartquery = lcl_chartquery & " UNION ALL "
       lcl_chartquery = lcl_chartquery & " SELECT "
       lcl_chartquery = lcl_chartquery & " 'Avg Days Last Activity' as seriesname, "
       lcl_chartquery = lcl_chartquery & " isnull([FORM NAME],'N/A') as xvalue, "
       lcl_chartquery = lcl_chartquery & " AVG(lastactivityindays) as yvalue, "
       lcl_chartquery = lcl_chartquery & " 1 as seriesorder "
       lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
       lcl_chartquery = lcl_chartquery &   varWhereClause
       lcl_chartquery = lcl_chartquery &   sOpenOnly
       lcl_chartquery = lcl_chartquery & " GROUP BY isnull([FORM NAME],'N/A') "
       lcl_chartquery = lcl_chartquery & " ORDER BY 4, isnull([FORM NAME],'N/A') "

    elseif lcl_selectedchart = "5" then
       lcl_charttitle = "Open Items Activity by Department"
       'lcl_charttype  = "bar"

       lcl_chartquery = ""
       lcl_chartquery = lcl_chartquery & " SELECT "
       lcl_chartquery = lcl_chartquery & " 'Avg Days Open' as seriesname, "
       lcl_chartquery = lcl_chartquery & " isnull(department,'N/A') as xvalue, "
       lcl_chartquery = lcl_chartquery & " AVG([Open]) as yvalue, "
       lcl_chartquery = lcl_chartquery & " 0 as seriesorder "
       lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
       lcl_chartquery = lcl_chartquery &   varWhereClause
       lcl_chartquery = lcl_chartquery &   sOpenOnly
       lcl_chartquery = lcl_chartquery & " GROUP BY isnull(department,'N/A') "

       lcl_chartquery = lcl_chartquery & " UNION ALL "

       lcl_chartquery = lcl_chartquery & " SELECT "
       lcl_chartquery = lcl_chartquery & " 'Avg Days Last Activity' as seriesname, "
       lcl_chartquery = lcl_chartquery & " isnull(department,'N/A') as xvalue, "
       lcl_chartquery = lcl_chartquery & " AVG(lastactivityindays) as yvalue, "
       lcl_chartquery = lcl_chartquery & " 1 as seriesorder "
       lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
       lcl_chartquery = lcl_chartquery &   varWhereClause
       lcl_chartquery = lcl_chartquery &   sOpenOnly
       lcl_chartquery = lcl_chartquery & " GROUP BY isnull(department,'N/A') "
       lcl_chartquery = lcl_chartquery & " ORDER BY 4, isnull(department,'N/A') "

    elseif lcl_selectedchart = "6" then
       lcl_charttitle = "Open Items by Department"
       'lcl_charttype  = "bar"

       lcl_chartquery = "SELECT "
       lcl_chartquery = lcl_chartquery & " 'Total Items' as seriesname, "
       lcl_chartquery = lcl_chartquery & " isnull(Department,'empty') as xvalue, "
       lcl_chartquery = lcl_chartquery & " count(department) as yvalue "
       lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
       lcl_chartquery = lcl_chartquery &   varWhereClause
       lcl_chartquery = lcl_chartquery &   sOpenOnly
       lcl_chartquery = lcl_chartquery & " GROUP BY DEPARTMENT "
       lcl_chartquery = lcl_chartquery & " ORDER BY count(department) desc, Department "

    elseif lcl_selectedchart = "7" then
       lcl_charttitle = "Open Items by Form"
       'lcl_charttype  = "bar"

       lcl_chartquery = "SELECT "
       lcl_chartquery = lcl_chartquery & " 'Total Items' as seriesname, "
       lcl_chartquery = lcl_chartquery & " isnull([FORM NAME],'empty') as xvalue, "
       lcl_chartquery = lcl_chartquery & " count([Form Name]) as yvalue "
       lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
       lcl_chartquery = lcl_chartquery &   varWhereClause
       lcl_chartquery = lcl_chartquery &   sOpenOnly
       lcl_chartquery = lcl_chartquery & " GROUP BY [FORM NAME] "
       lcl_chartquery = lcl_chartquery & " ORDER BY count([Form Name]) DESC, [FORM NAME] "

    end if

    if lcl_includedates then
       if lcl_charttitle <> "" then
          lcl_charttitle = lcl_charttitle & "\n" & lcl_fromdate & " - " & lcl_todate
       else
          lcl_charttitle = lcl_fromdate & " - " & lcl_todate
       end if
    end if

    if lcl_charttitle <> "" then
       lcl_charttitle = "'" & dbsafe(lcl_charttitle) & "'"
    else
       lcl_charttitle = "NULL"
    end if

 end if

'Get the page url for the chart/graph selected
 if lcl_charttype <> "" then
    if lcl_charttype = "bar" OR lcl_charttype = "column" then
       lcl_pageurl = "barcolumnchart"
    else
       lcl_pageurl = "piechart"
    end if

    lcl_charttype = "'" & dbsafe(lcl_charttype) & "'"
 end if

 if request("chartwidth") <> "" then
    lcl_chartwidth = request("chartwidth")
 end if

 if request("chartheight") <> "" then
    lcl_chartheight = request("chartheight")
 end if

 if request("showlegend") = "Y" then
    lcl_showlegend = 1

    if request("legendtitle") <> "" then
       lcl_legendtitle = "'" & dbsafe(request("legendtitle")) & "'"
    end if
 end if

 if request("collectedthreshold") <> "" then
    lcl_collectedthreshold = "'" & dbsafe(request("collectedthreshold")) & "'"
 end if

 if request("collectedlabel") <> "" then
    lcl_collectedlabel = "'" & dbsafe(request("collectedlabel")) & "'"
 end if

 if request("collectedlegendtext") <> "" then
    lcl_collectedlegendtext = "'" & dbsafe(request("collectedlegendtext")) & "'"
 end if

 if request("isAjaxRoutine") = "Y" then
    lcl_isAjaxRoutine = True
 else
    lcl_isAjaxRoutine = False
 end if

' lcl_chartquery = "SELECT "
' lcl_chartquery = lcl_chartquery & " ISNULL([Form Name],'empty') AS xvalue, "
' lcl_chartquery = lcl_chartquery & " COUNT([Form Name]) AS yvalue "
' lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
' lcl_chartquery = lcl_chartquery & " WHERE [Date Submitted] >= '" & lcl_fromdate & "' "
' lcl_chartquery = lcl_chartquery & " AND [Date Submitted] <= '" & lcl_todate & "' "
' lcl_chartquery = lcl_chartquery & " AND orgid = " & lcl_orgid & " "
' lcl_chartquery = lcl_chartquery & " AND UPPER(status) <> 'DISMISSED' "
' lcl_chartquery = lcl_chartquery & " AND UPPER(status) <> 'RESOLVED' "
' lcl_chartquery = lcl_chartquery & " GROUP BY [Form Name] "
' lcl_chartquery = lcl_chartquery & " ORDER BY COUNT([Form Name]) DESC, [FORM NAME]"

 if lcl_chartquery <> "" then
    lcl_chartquery = "'" & dbsafe(lcl_chartquery) & "'"
 end if

'Insert/Update the search options to get the chart url
 sSQL = "INSERT INTO egov_charts ("
 sSQL = sSQL & "orgid, "
 sSQL = sSQL & "charttype, "
 sSQL = sSQL & "chartquery, "
 sSQL = sSQL & "charttitle, "
 sSQL = sSQL & "showlegend, "
 sSQL = sSQL & "legendtitle, "
 sSQL = sSQL & "collectedthreshold, "
 sSQL = sSQL & "collectedlabel, "
 sSQL = sSQL & "collectedlegendtext, "
 sSQL = sSQL & "columnwidth, "
 sSQL = sSQL & "chartheight, "
 sSQL = sSQL & "chartwidth, "
 sSQL = sSQL & "dateadded, "
 sSQL = sSQL & "createdby"
 sSQL = sSQL & ") VALUES ("
 sSQL = sSQL & lcl_orgid               & ", "
 sSQL = sSQL & lcl_charttype           & ", "
 sSQL = sSQL & lcl_chartquery          & ", "
 sSQL = sSQL & lcl_charttitle          & ", "
 sSQL = sSQL & lcl_showlegend          & ", "
 sSQL = sSQL & lcl_legendtitle         & ", "
 sSQL = sSQL & lcl_collectedthreshold  & ", "
 sSQL = sSQL & lcl_collectedlabel      & ", "
 sSQL = sSQL & lcl_collectedlegendtext & ", "
 sSQL = sSQL & "0.6, "
 sSQL = sSQL & lcl_chartheight         & ", "
 sSQL = sSQL & lcl_chartwidth          & ", "
 sSQL = sSQL & "'" & now()             & "', "
 sSQL = sSQL & lcl_userid
 sSQL = sSQL & ") "

 lcl_chartid = RunInsertStatement(sSQL)

 if lcl_success = "Y" AND lcl_isAjaxRoutine then
    'response.write "Changes Saved"
    'response.write "http://dev4.egovlink.com/eclink/admin/charts/piechart.aspx?cid=2"
    'response.write "http://dev4.egovlink.com/eclink/admin/charts/barcolumnchart.aspx?cid=2"
    'response.write "http://dev4.egovlink.com/eclink/admin/charts/piechart.aspx?cid=" & lcl_chartid

    response.write "http://dev4.egovlink.com/eclink/admin/charts/" & lcl_pageurl & ".aspx?cid=" & lcl_chartid
 end if

'------------------------------------------------------------------------------
function dbsafe(p_value)
  lcl_return = ""

  lcl_return = replace(p_value,"'","''")

  dbsafe = lcl_return

end function
%>