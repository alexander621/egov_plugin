<% 
		' OPEN ITEMS BY DEPARTMENT REPORTS
		' LIST VIEW
		sOpenOnly = " AND ((dbo.egov_rpt_actionline .status <> 'RESOLVED') AND (dbo.egov_rpt_actionline.status <> 'DISMISSED')) "
		sSQLViewList = "Select Department, [Tracking Number],[Form Name],[Date Submitted],[Open] as [Days Open],lastactivityindays as [Days Last Activity],Status,[Submitted By],[Assigned To] from egov_rpt_actionline   " & varWhereClause & sOpenOnly & " ORDER BY Department, [Date Submitted] ASC"
				
		' SUMMARY (TOTALS \ SUBTOTALS ONLY) VIEW
		sSQLViewSummary = "Select GROUPING(Department) as ecGT,Grouping([Form Name]) as ecG1, Department,Case When Grouping(department) = 1 Then 'Grand Total:' When Grouping([Form Name]) = 1 Then  Department + ' Subtotal:' Else [Form Name] End as [Form Name],COUNT(department) as Total,AVG([Open]) as [Avg Days Open],AVG(lastactivityindays) as [Avg Days Last Activity] from egov_rpt_actionline   " & varWhereClause & sOpenOnly &  " GROUP BY DEPARTMENT,[Form Name]  WITH ROLLUP ORDER BY ecGT, Department"
		

		' DETAILS ( LIST WITH TOTALS \ SUBTOTALS ) VIEW
		sSQLViewDetails = "Select Grouping([Tracking Number]) as ecG2,GROUPING(Department) as ecGT,Grouping([Form Name]) as ecGT1, Department,Case  When Grouping([Form Name]) = 1 Then  Department + ' Subtotal:' When Grouping([Date Submitted]) = 1 Then [Form Name] + ' Subtotal:' Else [Form Name] End as [Form Name],[Tracking Number],[Date Submitted],Status,[Submitted By],Case When Grouping(department) = 1 Then 'Grand Total:' Else [Assigned To] End as [Assigned To],COUNT(department) as Total,AVG([Open]) as [Avg Days Open],AVG(lastactivityindays) as [Avg Days Last Activity] from egov_rpt_actionline   " & varWhereClause & sOpenOnly &  " GROUP BY DEPARTMENT,[Form Name],[Date Submitted],Status,[Submitted By],[Assigned To],[Tracking Number]  WITH ROLLUP HAVING Grouping([Form Name]) = 1 OR Grouping([Date Submitted]) = 1 OR Grouping([Tracking Number]) = 0 ORDER BY ecGT, Department desc,ecGT1,[Form Name]"



		' MONTHLY STATUS
		' LIST VIEW
		sSQLMonthViewList = "Select yearmonth as [Year\Month],Status,[Date Submitted],complete_date as [Date Completed],[Open] as [Days Open],[Time to Complete],Department, [Tracking Number],[Form Name],[Submitted By],[Assigned To] from egov_rpt_actionline   " & varWhereClause  & " ORDER BY yearmonth desc,status,[Open] Desc,[Time to Complete] Desc"

		' SUMMARY (TOTALS \ SUBTOTALS ONLY) VIEW
		sSQLMonthViewSummary = "Select Grouping(yearmonth) as ecGT,Grouping(status) as ecG1,yearmonth as [Year\Month],Case when Grouping(yearmonth) = 1 Then 'Grand Total:' when Grouping(status) = 1 Then 'Month Subtotal:' Else Status End as Status,Count(yearmonth) as Total, AVG([Open]) as [Avg Days Open],AVG([Time to Complete]) as [Avg Time to Complete] from egov_rpt_actionline   " & varWhereClause  & " GROUP BY yearmonth, status WITH ROLLUP ORDER BY ecGT, yearmonth , ecG1, status"

		' DETAILS ( LIST WITH TOTALS \ SUBTOTALS ) VIEW
		sSQLMonthViewDetail = "SELECT grouping([Date Submitted]) as ecG2,Grouping(yearmonth) as ecGT,Grouping(status) as ecGT1,yearmonth as [Year\Month],Case when Grouping(yearmonth) = 1 Then '' when Grouping(status) = 1 Then 'Month Subtotal:' when Grouping([Date Submitted]) = 1 Then Status + ' Subtotal:'  Else Status End as Status,[Date Submitted],complete_date as [Date Completed], Department, [Tracking Number],[Form Name],[Submitted By],Case When Grouping(yearmonth) = 1 Then 'Grand Total:' Else [Assigned To] End as [Assigned To],Count(yearmonth) as Total, AVG([Open]) as [Avg Days Open],AVG([Time to Complete]) as [Avg Time to Complete] from egov_rpt_actionline   " & varWhereClause  & " GROUP BY yearmonth, status,[Date Submitted],complete_date,Department, [Tracking Number],[Form Name],[Submitted By],[Assigned To] WITH ROLLUP HAVING GROUPING(yearmonth)  = 1 or grouping(status) = 1 or grouping([Date Submitted]) =1 or grouping([Assigned To])= 0 ORDER BY ecGT, yearmonth desc,ecGT1, status"


		' MONTHLY STATUS BY DEPARMENT
		' LIST VIEW
		sSQLMonthSourceViewList = "Select yearmonth as [Year\Month],Department,Status,[Date Submitted],complete_date as [Date Completed],[Open] as [Days Open],[Time to Complete], [Tracking Number],[Form Name],[Submitted By],[Assigned To] from egov_rpt_actionline   " & varWhereClause  & " ORDER BY yearmonth desc,department,status,[Date Submitted]"

		' SUMMARY (TOTALS \ SUBTOTALS ONLY) VIEW
		sSQLMonthSourceViewSummary =  "Select grouping(department) as ecGT1,Grouping(yearmonth) as ecGT,Grouping(status) as ecG1,yearmonth as [Year\Month],Department,Case when Grouping(yearmonth) = 1 Then 'Grand Total:' when Grouping(status) = 1 Then 'Month Subtotal:' Else Status End as Status,Count(yearmonth) as Total, AVG([Open]) as [Avg Days Open],AVG([Time to Complete]) as [Avg Time to Complete] from egov_rpt_actionline   " & varWhereClause  & " GROUP BY yearmonth, department, status WITH ROLLUP ORDER BY ecGT, yearmonth , grouping(department),department,ecG1,status"

		' DETAILS ( LIST WITH TOTALS \ SUBTOTALS ) VIEW
		sSQLMonthSourceViewDetail = "SELECT grouping([Date Submitted]) as ecG2,Grouping(yearmonth) as ecGT,Grouping(status) as ecGT1,yearmonth as [Year\Month],Department,Case  when Grouping(yearmonth) = 1 Then '' when Grouping(status) = 1 Then 'Month Subtotal:' when Grouping([Date Submitted]) = 1 Then Status + ' Subtotal:'  Else Status End as Status,[Date Submitted],complete_date as [Date Completed], [Tracking Number],[Form Name],[Submitted By],Case when Grouping(yearmonth) = 1 Then 'Grand Total:' Else [Assigned To] End as [Assigned To],Count(yearmonth) as Total, AVG([Open]) as [Avg Days Open],AVG([Time to Complete]) as [Avg Time to Complete] from egov_rpt_actionline   " & varWhereClause  & " GROUP BY yearmonth, department,status,[Date Submitted],complete_date, [Tracking Number],[Form Name],[Submitted By],[Assigned To] WITH ROLLUP HAVING GROUPING(yearmonth)  = 1 or grouping(status) = 1 or grouping([Date Submitted]) =1 or grouping([Assigned To])= 0 or grouping(department) = 1 ORDER BY ecGT, yearmonth desc, grouping(Department),department,ecGT1, status,[Date Submitted]"

		
		' BUILDING AND ZONING PROPERTY MAINTENANCE COMPLAINTS
		sSQLBZComplaints = "Select [Tracking Number],streetnumber + ' ' + streetname as [Address of Violation], 'UNDER CONSTRUCTION' as [Type of Violation], [Date Submitted] as [Date of Complaint], 'UNDER CONSTRUCTION' as [Action Taken], complete_date as [Date Resolved] from egov_rpt_actionline   " & varWhereClause  & " order by complete_date"


		sOpenOnly = " AND ((egov_rpt_PastDueDaysList.status <> 'RESOLVED') AND (egov_rpt_PastDueDaysList.status <> 'DISMISSED')) "
		sSQLPastDueList = "Select Department, [Tracking Number],[Form Name],[adjustedsubmitdate] as [Adjusted Date Submitted],[Open] as [Days Open],lastactivityindays as [Days Last Activity],allowedunresolveddays as [Allowed Days],([Open]-allowedunresolveddays) as [Days Past Due],Status,[Submitted By],[Assigned To] from egov_rpt_PastDueDaysList   " & varWhereClause & sOpenOnly 
		sSQLPastDueList ="select  * from (" & sSQLPastDueList & ") lkk where  [Days Past Due]>0 ORDER BY Department, [Adjusted Date Submitted] ASC" 

'response.write sSQLPastDueList
		Select Case request("ireport")

			Case 1
				sSQL = sSQLViewList

			Case 2
				sSQL = sSQLViewSummary

			Case 3
				sSQL = sSQLViewDetails

			Case 4
				sSQL = sSQLMonthViewList
			
			Case 5 
				sSQL = sSQLMonthViewSummary
			
			Case 6 
				sSQL = sSQLMonthViewDetail
			
			Case 7
		
				sSQL = sSQLMonthSourceViewList
			
			Case 8 
				sSQL = sSQLMonthSourceViewSummary
			
			Case 9 
				sSQL = sSQLMonthSourceViewDetail
			
			Case 10
				sSQL = sSQLBZComplaints

			Case 11
				sSQL = sSQLPastDueList 

			Case Else

				sSQL = sSQLViewList

		End Select

%>
