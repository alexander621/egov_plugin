<%
'------------------------------------------------------------------------------
sub getChartURL(ByVal iOrgID, ByVal iUserID, ByVal iFromDate, ByVal iToDate, ByVal iIncludesDates, ByVal iSelectedChart, _
                ByVal iChartType, ByVal iChartWidth, ByVal iChartHeight, ByVal iShowLegend, ByVal iLegendTitle, _
                ByVal iCollectedThreshold, ByVal iCollectedLabel, ByVal iCollectedLegendText, ByVal iFormType, _
                ByVal iExcludeAutoResolved, ByRef lcl_chartid, ByRef lcl_charturl)

 lcl_success              = "Y"
 lcl_chartid              = 0
 lcl_charturl             = ""
 lcl_orgid                = iOrgID
 lcl_userid               = iUserID
 lcl_today                = Date()
 lcl_fromdate             = ""
 lcl_todate               = ""
 lcl_includedates         = false
 lcl_selectedchart        = "NULL"
 lcl_chartquery           = "NULL"
 lcl_charttitle           = "NULL"
 lcl_charttype            = ""
 lcl_pageurl              = "piechart"
 lcl_chartwidth           = "700"
 lcl_chartheight          = "500"
 lcl_showlegend           = 0
 lcl_legendtitle          = "NULL"
 lcl_collectedthreshold   = "5"
 lcl_collectedlabel       = "NULL"
 lcl_collectedlegendtext  = "NULL"
 lcl_formtype             = ""
 lcl_exclude_autoresolved = ""

 if iFromDate <> "" then
    lcl_fromdate = iFromDate
 end if

 if iToDate <> "" then
    lcl_todate = iToDate
 end if

'If empty then default to the current date
 if lcl_toDate = "" OR IsNull(lcl_toDate) then
    lcl_toDate = lcl_today
 end if

 if lcl_fromDate = "" OR IsNull(lcl_fromDate) then
    lcl_fromDate = cdate(Month(lcl_today)& "/1/" & Year(lcl_today))
 end if

 if iIncludesDates = "Y" then
    lcl_includedates = true
 end if

 if iChartType <> "" then
    lcl_charttype = iChartType
 else
    lcl_charttype = "pie"
 end if

 if iSelectedChart <> "" then
    lcl_selectedchart = iSelectedChart

   'Build SQL where clause
    varWhereClause = " WHERE ([Date Submitted] >= '" & lcl_fromdate & "' AND [Date Submitted] <= '" & DateAdd("d",1,lcl_todate) & "') "
    varWhereClause = varWhereClause & " AND orgid='" & lcl_orgid & "'"

    if iFormType <> "" then
       lcl_access_internalforms = ""

       if ucase(iFormType) = "PUBLIC" then
          lcl_access_internalforms = "0"
       elseif ucase(iFormType) = "INTERNAL" then
          lcl_access_internalforms = "1"
       end if

       if lcl_access_internalforms <> "" then
          varWhereClause = varWhereClause & " AND action_form_internal = " & lcl_access_internalforms
       end if
    end if

    sOpenOnly = " AND (UPPER(status) <> 'DISMISSED' AND UPPER(status) <> 'RESOLVED')"

    if iExcludeAutoResolved = "Y" then
       varWhereClause = varWhereClause & " AND action_form_resolved_status <> 'Y' "
    end if

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
       lcl_charttitle = "Open Items Activity by Form\n(Avg Days Open)"
       'lcl_charttype  = "column"

       lcl_chartquery = ""
       lcl_chartquery = lcl_chartquery & " SELECT "
       lcl_chartquery = lcl_chartquery & " 'Avg Days Open' as seriesname, "
       lcl_chartquery = lcl_chartquery & " isnull([FORM NAME],'N/A') + ' (' + cast(AVG([Open]) as varchar) + ')' as xvalue, "
       'lcl_chartquery = lcl_chartquery & " isnull([FORM NAME],'N/A') as xvalue, "
       lcl_chartquery = lcl_chartquery & " AVG([Open]) as yvalue "
       lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
       lcl_chartquery = lcl_chartquery &   varWhereClause
       lcl_chartquery = lcl_chartquery &   sOpenOnly
       lcl_chartquery = lcl_chartquery & " GROUP BY [FORM NAME] "
       lcl_chartquery = lcl_chartquery & " ORDER BY AVG([Open]) desc, isnull([FORM NAME],'N/A') "

    elseif lcl_selectedchart = "9" then
       lcl_charttitle = "Open Items Activity by Form\n(Avg Days Last Activity)"
       'lcl_charttype  = "column"

       lcl_chartquery = ""
       lcl_chartquery = lcl_chartquery & " SELECT "
       lcl_chartquery = lcl_chartquery & " 'Avg Days Last Activity' as seriesname, "
       lcl_chartquery = lcl_chartquery & " isnull([FORM NAME],'N/A') + ' (' + cast(AVG(lastactivityindays) as varchar) + ')' as xvalue, "
       'lcl_chartquery = lcl_chartquery & " isnull([FORM NAME],'N/A') as xvalue, "
       lcl_chartquery = lcl_chartquery & " AVG(lastactivityindays) as yvalue "
       lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
       lcl_chartquery = lcl_chartquery &   varWhereClause
       lcl_chartquery = lcl_chartquery &   sOpenOnly
       lcl_chartquery = lcl_chartquery & " GROUP BY [FORM NAME] "
       lcl_chartquery = lcl_chartquery & " ORDER BY AVG(lastactivityindays) desc, isnull([FORM NAME],'N/A') "

    elseif lcl_selectedchart = "5" then
       lcl_charttitle = "Open Items Activity by Department\n(Avg Days Open)"
       'lcl_charttype  = "bar"

       lcl_chartquery = ""
       lcl_chartquery = lcl_chartquery & " SELECT "
       lcl_chartquery = lcl_chartquery & " 'Avg Days Open' as seriesname, "
       lcl_chartquery = lcl_chartquery & " isnull(department,'N/A') + ' (' + cast(AVG([Open]) as varchar) + ')' as xvalue, "
       'lcl_chartquery = lcl_chartquery & " isnull(department,'N/A') as xvalue, "
       lcl_chartquery = lcl_chartquery & " AVG([Open]) as yvalue "
       lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
       lcl_chartquery = lcl_chartquery &   varWhereClause
       lcl_chartquery = lcl_chartquery &   sOpenOnly
       lcl_chartquery = lcl_chartquery & " GROUP BY department "
       lcl_chartquery = lcl_chartquery & " ORDER BY AVG([Open]) desc, department "

    elseif lcl_selectedchart = "8" then
       lcl_charttitle = "Open Items Activity by Department\n(Avg Days Last Activity)"
       'lcl_charttype  = "bar"

       lcl_chartquery = ""
       lcl_chartquery = lcl_chartquery & " SELECT "
       lcl_chartquery = lcl_chartquery & " 'Avg Days Last Activity' as seriesname, "
       lcl_chartquery = lcl_chartquery & " isnull(department,'N/A') + ' (' + cast(AVG(lastactivityindays) as varchar) + ')' as xvalue, "
       'lcl_chartquery = lcl_chartquery & " isnull(department,'N/A') as xvalue, "
       lcl_chartquery = lcl_chartquery & " AVG(lastactivityindays) as yvalue "
       lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
       lcl_chartquery = lcl_chartquery &   varWhereClause
       lcl_chartquery = lcl_chartquery &   sOpenOnly
       lcl_chartquery = lcl_chartquery & " GROUP BY department "
       lcl_chartquery = lcl_chartquery & " ORDER BY AVG(lastactivityindays), department "

    elseif lcl_selectedchart = "6" then
       lcl_charttitle = "Open Items by Department"
       'lcl_charttype  = "bar"

       lcl_chartquery = "SELECT "
       lcl_chartquery = lcl_chartquery & " 'Total Items' as seriesname, "
       lcl_chartquery = lcl_chartquery & " isnull(Department,'empty') + ' (' + cast(count(department) as varchar) + ')' as xvalue, "
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
       lcl_chartquery = lcl_chartquery & " isnull([FORM NAME],'empty') + ' (' + cast(count([Form Name]) as varchar) + ')' as xvalue, "
       lcl_chartquery = lcl_chartquery & " count([Form Name]) as yvalue "
       lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
       lcl_chartquery = lcl_chartquery &   varWhereClause
       lcl_chartquery = lcl_chartquery &   sOpenOnly
       lcl_chartquery = lcl_chartquery & " GROUP BY [FORM NAME] "
       lcl_chartquery = lcl_chartquery & " ORDER BY [FORM NAME] desc, count([Form Name]) "

'-- Original ------------------------------------------------------------------
'    elseif lcl_selectedchart = "4" then
'       lcl_charttitle = "Open Items Activity by Form"
       'lcl_charttype  = "column"

'       lcl_chartquery = ""
'       lcl_chartquery = lcl_chartquery & " SELECT "
'       lcl_chartquery = lcl_chartquery & " 'Avg Days Open' as seriesname, "
       'lcl_chartquery = lcl_chartquery & " isnull([FORM NAME],'N/A') + ' (' + cast(AVG([Open]) as varchar) + ')' as xvalue, "
'       lcl_chartquery = lcl_chartquery & " isnull([FORM NAME],'N/A') as xvalue, "
'       lcl_chartquery = lcl_chartquery & " AVG([Open]) as yvalue, "
'       lcl_chartquery = lcl_chartquery & " 0 as seriesorder "
'       lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
'       lcl_chartquery = lcl_chartquery &   varWhereClause
'       lcl_chartquery = lcl_chartquery &   sOpenOnly
'       lcl_chartquery = lcl_chartquery & " GROUP BY isnull([FORM NAME],'N/A') "
'       lcl_chartquery = lcl_chartquery & " UNION ALL "
'       lcl_chartquery = lcl_chartquery & " SELECT "
'       lcl_chartquery = lcl_chartquery & " 'Avg Days Last Activity' as seriesname, "
       'lcl_chartquery = lcl_chartquery & " isnull([FORM NAME],'N/A') + ' (' + cast(AVG(lastactivityindays) as varchar) + ')' as xvalue, "
'       lcl_chartquery = lcl_chartquery & " isnull([FORM NAME],'N/A') as xvalue, "
'       lcl_chartquery = lcl_chartquery & " AVG(lastactivityindays) as yvalue, "
'       lcl_chartquery = lcl_chartquery & " 1 as seriesorder "
'       lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
'       lcl_chartquery = lcl_chartquery &   varWhereClause
'       lcl_chartquery = lcl_chartquery &   sOpenOnly
'       lcl_chartquery = lcl_chartquery & " GROUP BY isnull([FORM NAME],'N/A') "
'       lcl_chartquery = lcl_chartquery & " ORDER BY 4, isnull([FORM NAME],'N/A') "

'    elseif lcl_selectedchart = "5" then
'       lcl_charttitle = "Open Items Activity by Department"
       'lcl_charttype  = "bar"

'       lcl_chartquery = ""
'       lcl_chartquery = lcl_chartquery & " SELECT "
'       lcl_chartquery = lcl_chartquery & " 'Avg Days Open' as seriesname, "
       'lcl_chartquery = lcl_chartquery & " isnull(department,'N/A') + ' (' + cast(AVG([Open]) as varchar) + ')' as xvalue, "
'       lcl_chartquery = lcl_chartquery & " isnull(department,'N/A') as xvalue, "
'       lcl_chartquery = lcl_chartquery & " AVG([Open]) as yvalue, "
'       lcl_chartquery = lcl_chartquery & " 0 as seriesorder "
'       lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
'       lcl_chartquery = lcl_chartquery &   varWhereClause
'       lcl_chartquery = lcl_chartquery &   sOpenOnly
'       lcl_chartquery = lcl_chartquery & " GROUP BY isnull(department,'N/A') "

'       lcl_chartquery = lcl_chartquery & " UNION ALL "

'       lcl_chartquery = lcl_chartquery & " SELECT "
'       lcl_chartquery = lcl_chartquery & " 'Avg Days Last Activity' as seriesname, "
       'lcl_chartquery = lcl_chartquery & " isnull(department,'N/A') + ' (' + cast(AVG(lastactivityindays) as varchar) + ')' as xvalue, "
'       lcl_chartquery = lcl_chartquery & " isnull(department,'N/A') as xvalue, "
'       lcl_chartquery = lcl_chartquery & " AVG(lastactivityindays) as yvalue, "
'       lcl_chartquery = lcl_chartquery & " 1 as seriesorder "
'       lcl_chartquery = lcl_chartquery & " FROM egov_rpt_actionline "
'       lcl_chartquery = lcl_chartquery &   varWhereClause
'       lcl_chartquery = lcl_chartquery &   sOpenOnly
'       lcl_chartquery = lcl_chartquery & " GROUP BY isnull(department,'N/A') "
'       lcl_chartquery = lcl_chartquery & " ORDER BY 4, isnull(department,'N/A') "

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
 else
    lcl_charttype = "NULL"
 end if

 if iChartWidth <> "" then
    lcl_chartwidth = iChartWidth
 end if

 if iChartHeight <> "" then
    lcl_chartheight = iChartHeight
 end if

 if iShowLegend = "Y" then
    lcl_showlegend = 1

    if iLegendTitle = "" then
       lcl_legendtitle = "'Legend'"
    else
       lcl_legendtitle = "'" & dbsafe(iLegendTitle) & "'"
    end if
 end if

 if iCollectedThreshold <> "" then
    lcl_collectedthreshold = "'" & dbsafe(iCollectedThreshold) & "'"
 end if

 if iCollectedLabel <> "" then
    lcl_collectedlabel = "'" & dbsafe(iCollectedLabel) & "'"
 end if

 if iCollectedLegendText <> "" then
    lcl_collectedlegendtext = "'" & dbsafe(iCollectedLegendText) & "'"
 end if

 if iFormType <> "" then
    lcl_formtype = "'" & dbsafe(iFormType) & "'"
 else
    lcl_formtype = "NULL"
 end if

 if iExcludeAutoResolved <> "" then
    lcl_exclude_autoresolved = UCASE(iExcludeAutoResolved)
    lcl_exclude_autoresolved = "'" & dbsafe(lcl_exclude_autoresolved) & "'"
 else
    lcl_exclude_autoresolved = "NULL"
 end if

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
 sSQL = sSQL & "createdby, "
 sSQL = sSQL & "formtype, "
 sSQL = sSQL & "exclude_autoresolved "
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
 sSQL = sSQL & lcl_userid              & ", "
 sSQL = sSQL & lcl_formtype            & ", "
 sSQL = sSQL & lcl_exclude_autoresolved
 sSQL = sSQL & ") "
'dtb_debug(sSQL)
 lcl_chartid = RunInsertStatement(sSQL)

 if lcl_success = "Y" then
    'Maintain the records in the table for this user.  If the user has over 20 records then delete the oldest.
     maintainChartHistory lcl_orgid, lcl_userid

    'response.write "Changes Saved"
    'response.write "http://dev4.egovlink.com/eclink/admin/charts/piechart.aspx?cid=2"
    'response.write "http://dev4.egovlink.com/eclink/admin/charts/barcolumnchart.aspx?cid=2"
    'response.write "http://dev4.egovlink.com/eclink/admin/charts/piechart.aspx?cid=" & lcl_chartid

    'response.write "http://dev4.egovlink.com/eclink/admin/charts/" & lcl_pageurl & ".aspx?cid=" & lcl_chartid
    lcl_org_sitename = getOrgVirtualSiteName(lcl_orgid)
    lcl_charturl     = Application("charts_graphs_url") & "/" & lcl_org_sitename & "/admin/charts/" & lcl_pageurl & ".aspx?cid=" & lcl_chartid
 end if

end sub

'------------------------------------------------------------------------------
sub maintainChartHistory(iOrgID, iUserID)

  'Find the total number of records for the user
   lcl_totalcharts = 0
   lcl_old_chartid = 0

   sSQL = "SELECT count(chartid) as total_charts "
   sSQL = sSQL & " FROM egov_charts "
   sSQL = sSQL & " WHERE orgid = "   & iOrgID
   sSQL = sSQL & " AND createdby = " & iUserID

  	set oGetTotalCharts = Server.CreateObject("ADODB.Recordset")
	  oGetTotalCharts.Open sSQL, Application("DSN"), 3, 1

   if not oGetTotalCharts.eof then
      lcl_totalcharts = oGetTotalCharts("total_charts")
   end if

  'The Limit is 20.  If over this limit then delete the oldest record
   if lcl_totalcharts > 20 then
     'Get the max chartid
      sSQL = "SELECT chartid "
      sSQL = sSQL & " FROM egov_charts "
      sSQL = sSQL & " WHERE dateadded = (select min(c2.dateadded) "
      sSQL = sSQL &                    " from egov_charts as c2 "
      sSQL = sSQL &                    " where orgid = "   & iOrgID
      sSQL = sSQL &                    " and createdby = " & iUserID & ") "

     	set oGetOldChartID = Server.CreateObject("ADODB.Recordset")
   	  oGetOldChartID.Open sSQL, Application("DSN"), 3, 1

      if not oGetOldChartID.eof then
         lcl_old_chartid = oGetOldChartID("chartid")
      end if

     'If there is a max chartid then delete it.
      if lcl_old_chartid > 0 then
         sSQL = "DELETE FROM egov_charts "
         sSQL = sSQL & " WHERE orgid = "   & iOrgID
         sSQL = sSQL & " AND createdby = " & iUserID
         sSQL = sSQL & " AND chartid = "   & lcl_old_chartid

        	set oDeleteMaxChartID = Server.CreateObject("ADODB.Recordset")
      	  oDeleteMaxChartID.Open sSQL, Application("DSN"), 3, 1

         set oDeleteMaxChartID = nothing
      end if

      oGetOldChartID.close
      set oGetOldChartID = nothing

   end if

   oGetTotalCharts.close
   set oGetTotalCharts = nothing

end sub

'------------------------------------------------------------------------------
function getOrgVirtualSiteName(iOrgID)

 lcl_return = "eclink"

 sSQL = "SELECT orgVirtualSiteName "
	sSQL = sSQL & " FROM Organizations "
 sSQL = sSQL & " WHERE orgid = " & iOrgID

	set oGetOrgInfo = Server.CreateObject("ADODB.Recordset")
	oGetOrgInfo.Open sSQL, Application("DSN"), 0, 1
	
	if not oGetOrgInfo.eof then
    lcl_return = oGetOrgInfo("orgVirtualSiteName")
 end if

 oGetOrgInfo.close
 set oGetOrgInfo = nothing

 getOrgVirtualSiteName = lcl_return

end function

'------------------------------------------------------------------------------
function isChartTypeSelected(iChartNum, iChartTypeSelected)

  lcl_return = ""

  if iChartNum <> "" AND iChartTypeSelected <> "" then
     if iChartNum = iChartTypeSelected then
        lcl_return = " selected=""selected"""
     end if
  end if

  isChartTypeSelected = lcl_return

end function

'------------------------------------------------------------------------------
function dbsafe(p_value)
  lcl_return = ""

  lcl_return = replace(p_value,"'","''")

  dbsafe = lcl_return

end function
%>