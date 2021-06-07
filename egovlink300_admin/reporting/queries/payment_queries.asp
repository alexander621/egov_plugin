<% 
		' DAILY RECEIPTS REPORTS
		' LIST VIEW
		sSQLViewList = "Select paymentdate as [Payment Date],account as Account,Item,Cast(Credit as Money) as Sales, [Transaction ID], paymentlocationname as [Payment Source],paymenttypename as [Payment Type]  from egov_glreport_combined   " & varWhereClause & " ORDER BY paymentdateshort asc,paymenttypename, paymentdate asc"
				
		' SUMMARY (TOTALS \ SUBTOTALS ONLY) VIEW
		sSQLViewSummary = "Select GROUPING(paymentdateshort) as ecGT,GROUPING(paymenttypename) as ecG1,GROUPING(paymentdateshort) as ecG2, Case  when GROUPING(paymenttypename)  = 1 THEN null  Else paymentdateshort END  as [Payment Date], CASE WHEN GROUPING(paymentdateshort) = 1 THEN ' Grand Total:' WHEN GROUPING(paymenttypename) = 1 THEN 'Date Subtotal:'  ELSE paymenttypename + ' subtotal:' END as [Payment Type], COUNT(paymentdate) as Total, Cast(sum(credit) as MOney) as Sales from egov_glreport_combined  " & varWhereClause & " GROUP BY paymentdateshort,paymenttypename WITH ROLLUP ORDER BY GROUPING(paymentdateshort), paymentdateshort asc"

		' DETAILS ( LIST WITH TOTALS \ SUBTOTALS ) VIEW
		sSQLViewDetails = "Select GROUPING(paymenttypename) as ecGT,CASE WHEN GROUPING (account) = 1 THEN 2 ELSE 0 END  as ecG2, paymentdateshort as [Payment Date],CASE WHEN GROUPING(paymentdateshort) = 1 THEN NULL WHEN GROUPING(paymenttypename) = 1 THEN  'Date Subtotal:' WHEN GROUPING(account) > 0 THEN paymenttypename + ' Subtotal:'  ELSE  paymenttypename  END as [Payment Type], account as Account,category as Category,item as Item, Case WHEN GROUPING(paymentdateshort) = 1 THEN ' Grand Total:' ELSE paymentlocationname END as [Payment Source], COUNT(paymentdate) as Total,Cast(sum(credit) as MOney) as Sales from egov_glreport_combined  " & varWhereClause & "  GROUP BY paymentdateshort,paymenttypename,account,category,item,paymentlocationname  WITH ROLLUP HAVING GROUPING(paymenttypename) = 1 OR grouping(account) = 1 OR grouping(paymentlocationname) = 0 ORDER BY grouping(paymentdateshort),paymentdateshort asc,GROUPING(paymenttypename)"



		' MONTHLY REVENUE BY CATEGORY REPORT
		' LIST VIEW
		sSQLMonthViewList = "Select convert(char(6), Paymentdateshort, 112) as [Year\Month],Category,Item,paymentdate as [Payment Date],account as Account,count(credit) as Total, Cast(sum(credit) as money) as Sales from egov_glreport_combined  " & varWhereClause & "  group by  convert(char(6), Paymentdateshort, 112),Category,Item,account,paymentdate ORDER BY [Year\Month] ,category,item,paymentdate"

		' SUMMARY (TOTALS \ SUBTOTALS ONLY) VIEW
		sSQLMonthViewSummary = "Select GROUPING( convert(char(6), Paymentdateshort, 112)) as ecGT,GROUPING(category) as ecG1,convert(char(6), Paymentdateshort, 112) as [Year\Month],Case WHen grouping( convert(char(6), Paymentdateshort, 112)) = 1 THEN 'Grand Total:'  WHEN grouping(category)=1 THEN ' Month Subtotal:' Else Category END as Category, count(credit) as Total, Cast(sum(credit) as money) as Sales from egov_glreport_combined  " & varWhereClause & "  group by  convert(char(6), Paymentdateshort, 112),Category WITH ROLLUP order by GROUPING( convert(char(6), Paymentdateshort, 112)),[Year\Month]  "

		' DETAILS ( LIST WITH TOTALS \ SUBTOTALS ) VIEW
		sSQLMonthViewDetail = "Select GROUPING(paymentdate) as ecG4, GROUPING(category) as ecGT,convert(char(6), Paymentdateshort, 112) as [Year\Month],Case WHen grouping(convert(char(6), Paymentdateshort, 112)) =1 Then ' ' WHen grouping(category) =1 Then ' Month Subtotal:' Else category End as [Category],Case WHen grouping(convert(char(6), Paymentdateshort, 112)) =1 Then ' Grand Total:' When Grouping(item) = 1 Then category + ' Subtotal:' Else item End as Item, paymentdate as [Payment Date],account as Account, count(credit) as Total, Cast(sum(credit) as money) as Sales,CASE WHEN convert(char(6), Paymentdateshort, 112) is null Then 1 Else 0 END as ecG3 from egov_glreport_combined  " & varWhereClause & "  group by  convert(char(6), Paymentdateshort, 112),category,item,paymentdate,account  WITH ROLLUP HAVING Grouping( convert(char(6), Paymentdateshort, 112)) = 1 or grouping(account) = 0 or grouping(item) = 1 ORDER BY ecG3,[Year\Month] ,ecGT,category"


		' MONTHLY REVENUE BY SOURCE REPORT
		' LIST VIEW
		sSQLMonthSourceViewList = "Select convert(char(6), Paymentdateshort, 112) as [Year\Month],paymentlocationname as [Payment Source],paymenttypename as [Payment Type],paymentdate as [Payment Date],account as Account,count(credit) as Total, Cast(sum(credit) as money) as Sales from egov_glreport_combined  " & varWhereClause & "  group by  convert(char(6), Paymentdateshort, 112),paymentlocationname,paymenttypename,account,paymentdate"

		' SUMMARY (TOTALS \ SUBTOTALS ONLY) VIEW
		sSQLMonthSourceViewSummary = "Select GROUPING( convert(char(6), Paymentdateshort, 112)) as ecGT,GROUPING(paymentlocationname) as ecG1,convert(char(6), Paymentdateshort, 112) as [Year\Month],Case when grouping( convert(char(6), Paymentdateshort, 112)) = 1 then 'Grand Total:' when grouping(paymentlocationname) = 1 then ' Month Subtotal:' else paymentlocationname end as [Payment Source],count(credit) as Total, Cast(sum(credit) as money) as Sales from egov_glreport_combined  " & varWhereClause & "  group by  convert(char(6), Paymentdateshort, 112),paymentlocationname  WITH ROLLUP order by GROUPING( convert(char(6), Paymentdateshort, 112)),[Year\Month]"

		' DETAILS ( LIST WITH TOTALS \ SUBTOTALS ) VIEW
		sSQLMonthSourceViewDetail = "Select GROUPING(paymentdate) as ecG4, GROUPING(paymentlocationname) as ecGT,convert(char(6), Paymentdateshort, 112) as [Year\Month],Case WHen grouping(convert(char(6), Paymentdateshort, 112)) =1 Then ' ' WHen grouping(paymentlocationname) =1 Then ' Month Subtotal:' Else paymentlocationname End as [Payment Source],Case WHen grouping(convert(char(6), Paymentdateshort, 112)) =1 Then ' Grand Total:' When Grouping(item) = 1 Then paymentlocationname + ' Subtotal:' Else item End as Item, paymentdate as [Payment Date],account as Account, count(credit) as Total, Cast(sum(credit) as money) as Sales,CASE WHEN convert(char(6), Paymentdateshort, 112) is null Then 1 Else 0 END as ecG3 from egov_glreport_combined  " & varWhereClause & "  group by  convert(char(6), Paymentdateshort, 112),paymentlocationname,item,paymentdate,account  WITH ROLLUP HAVING Grouping( convert(char(6), Paymentdateshort, 112)) = 1 or grouping(account) = 0 or grouping(item) = 1 ORDER BY ecG3,[Year\Month],ecGT,paymentlocationname"
		
		
		' DAILY Payment Methods REPORTS
		' LIST VIEW
		sSQLPaymentViewList = "Select paymentdate as [Payment Date],account as Account,Item,Cast(Credit as Money) as Sales, [Transaction ID], paymentlocationname as [Payment Source],paymenttypename as [Payment Type]  from egov_glreport_paymentcombined   " & varWhereClause & " ORDER BY paymentdateshort asc,paymenttypename, paymentdate asc"
				
		' SUMMARY (TOTALS \ SUBTOTALS ONLY) VIEW
		sSQLpaymentViewSummary = "Select GROUPING(paymentdateshort) as ecGT,GROUPING(paymenttypename) as ecG1,GROUPING(paymentdateshort) as ecG2, Case  when GROUPING(paymenttypename)  = 1 THEN null  Else paymentdateshort END  as [Payment Date], CASE WHEN GROUPING(paymentdateshort) = 1 THEN ' Grand Total:' WHEN GROUPING(paymenttypename) = 1 THEN 'Date Subtotal:'  ELSE paymenttypename + ' subtotal:' END as [Payment Type], COUNT(paymentdate) as Total, Cast(sum(credit) as MOney) as Sales from egov_glreport_paymentcombined  " & varWhereClause & " GROUP BY paymentdateshort,paymenttypename WITH ROLLUP ORDER BY GROUPING(paymentdateshort), paymentdateshort asc"

		' DETAILS ( LIST WITH TOTALS \ SUBTOTALS ) VIEW
		sSQLPaymentViewDetails = "Select GROUPING(paymenttypename) as ecGT,CASE WHEN GROUPING (account) = 1 THEN 2 ELSE 0 END  as ecG2, paymentdateshort as [Payment Date],CASE WHEN GROUPING(paymentdateshort) = 1 THEN NULL WHEN GROUPING(paymenttypename) = 1 THEN  'Date Subtotal:' WHEN GROUPING(account) > 0 THEN paymenttypename + ' Subtotal:'  ELSE  paymenttypename  END as [Payment Type], account as Account,category as Category,item as Item, Case WHEN GROUPING(paymentdateshort) = 1 THEN ' Grand Total:' ELSE paymentlocationname END as [Payment Source], COUNT(paymentdate) as Total,Cast(sum(credit) as MOney) as Sales from egov_glreport_paymentcombined  " & varWhereClause & "  GROUP BY paymentdateshort,paymenttypename,account,category,item,paymentlocationname  WITH ROLLUP HAVING GROUPING(paymenttypename) = 1 OR grouping(account) = 1 OR grouping(paymentlocationname) = 0 ORDER BY grouping(paymentdateshort),paymentdateshort asc,GROUPING(paymenttypename)"



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
				sSQL = sSQLPaymentViewList

			Case 11
				sSQL = sSQLpaymentViewSummary

			Case 12
				sSQL = sSQLPaymentViewDetails

			Case Else

				sSQL = sSQLViewList

		End Select

%>