<!--#include file="pureaspgraph/pureaspgraph.asp"-->
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: RESPORTING_MODULE_CHART.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  
'
' MODIFICATION HISTORY
' 1.0   1/10/07	JOHN STULLENBERGER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
if isFeatureOffline("action line") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

'INITIALIZE AND DECLARE VARIABLES
'SPECIFY FOLDER LEVEL
sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK
If Not UserHasPermission( Session("UserId"), "advactionlinerpt" ) Then
	  'response.redirect sLevel & "permissiondenied.asp"
End If 

' PROCESS REPORT FILTER VALUES
' PROCESS DATE VALUES
fromDate = Request("fromDate")
toDate = Request("toDate")
today = Date()

' IF EMPTY DEFAULT TO CURRENT TO DATE
If toDate = "" or IsNull(toDate) Then toDate = today End If
If fromDate = "" or IsNull(fromDate) Then fromDate = cdate(Month(today)& "/1/" & Year(today)) End If

' BUILD SQL WHERE CLAUSE
sOpenOnly = " AND (UPPER(status) <> 'DISMISSED' AND UPPER(status) <> 'RESOLVED')"
varWhereClause = " WHERE ([Date Submitted] >= '" & fromDate & "' AND [Date Submitted] <= '" & DateAdd("d",1,toDate) & "') "
varWhereClause = varWhereClause & " AND orgid='" & session("orgid") & "'"
%>



<html>
<head>
  <title>E-Gov Request Charts</title>

	<link rel="stylesheet" type="text/css" href="reporting.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />


	<script language="Javascript" src="scripts/tablesort.js"></script>

	<script language="Javascript">
	  <!--
		function doCalendar(ToFrom) {
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  eval('window.open("../recreation/gr_calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}
	  //-->
	</script>

	<script language="Javascript" src="scripts/dates.js"></script>

</head>


<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">


<% ShowHeader sLevel %>


<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->


<form action="reporting_module_chart.asp" method=post name=frmPFilter >

	<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<tr>
			<td><font size="+1"><b>E-Gov Request Charts</b></font></td>
		</tr>
		<tr>
			<td>
				<fieldset >
					
					<legend ><b>Date Filters:</b></legend>
				
				<!--BEGIN: FILTERS-->
				<!--BEGIN: DATE FILTERS-->
				<P>
				<table>
					<tr>
						<td  align=right> <b>Request Submission Date: </td>
						<td>
							<input type=text name="fromDate" value="<%=fromDate%>">
							<a href="javascript:void doCalendar('From');"><img src="../images/calendar.gif" border=0></a>		 
						</td>
						<td>&nbsp;</td>
						<td >
							<b>To:</b> 
						</td>
						<td>
							<input type=text name="toDate" value="<%=toDate%>">
							<a href="javascript:void doCalendar('To');"><img src="../images/calendar.gif" border=0></a>
						</td>
						<td>&nbsp;</td>
						<td><%DrawDateChoices("Dates")%></td>
					</tr>
				</table>
				</p>
				<!--END: DATE FILTERS-->



				</fieldset>
				<!--END: FILTERS-->



				 <!--BEGIN: PREDEFINED REPORTS-->
				  <fieldset>

					<legend><b>Chart Options:</b></legend>

					<P>
					  <Select name="ireport">
							<option value=1 <%if request("ireport") = 1 Then response.write " SELECTED " End If %>> Monthly Status - Chart
							<option Value=2 <%if request("ireport") = 2 Then response.write " SELECTED " End If %>> Monthly Status by Department - Chart
							<option value=3 <%if request("ireport") = 3 Then response.write " SELECTED " End If %>> Most Submitted Requests - Chart
							<option value=6 <%if request("ireport") = 6 Then response.write " SELECTED " End If %>> Open Items by Department - Chart
							<option value=5 <%if request("ireport") = 5 Then response.write " SELECTED " End If %>> Open Items Activity by Department - Chart
							<option value=7 <%if request("ireport") = 7 Then response.write " SELECTED " End If %>> Open Items by Form - Chart
							<option value=4 <%if request("ireport") = 4 Then response.write " SELECTED " End If %>> Open Items Activity by Form - Chart
							
					  </select>
					  <input class=excelexport type=submit value="View Chart"> - (<a href="reporting_module_actionline.asp">Click Here to View Available Reports</a>)
					</P>
				 

				  </fieldset>
				 <!--END: PREDEFINED REPORTS-->

				
    </td>
  </tr>
	<tr>
 
      <td colspan="3" valign="top">
	  
	  
		<!--BEGIN: DISPLAY RESULTS-->

		<%
		
		' DETERMINE WHICH CHART TO DISPLAY AND DISPLAY IT
		Select Case request("ireport")

			Case "1"
				subDisplayRequestsbyMonth()
			Case "2"
				subDisplayRequeststatusbyDepartment()
			Case "3"
				subDisplayMostSubmittedRequests()
			Case "4"
				subDisplayOpenRequestsbyForm()
			Case "5"
				subDisplayOpenRequestsbyDepartment()
			Case "6"
				subDisplayOpenRequestsTotalbyDepartment()
			Case "7"
				subDisplayOpenRequestsTotalbyForm()
			Case "8"
				subDisplayDailyReceiptTotals()
			Case "9"
				subDisplayCategoryTotalsbyMonth()
			Case Else
				subDisplayRequestsbyMonth()

		End Select
			
		 
		%>
		

		<!-- END: DISPLAY RESULTS -->
      

	  </td>
       
    </tr>
  </table>

<P>




</P>  </form>
  

<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  


</body>
</html>



<!--#Include file="includes/report_display_functions.asp"-->  



<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------------------------------
' PUBLIC SUB SUBDISPLAYCHART(STITLE,SXAXISTITLE,SYAXESTITLE,SSQL)
'------------------------------------------------------------------------------------------------------------
Public Sub SubDisplayChart(sTitle,sXAxisTitle,sYAxesTitle,sSQL)

	' GET DATA FORM THE DATABASE
	Set oRequests = Server.CreateObject("ADODB.Recordset")

	' OPEN RECORDSET
	oRequests.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oRequests.EOF THEN
		
		' CREATE CHART OBJECT
		Dim objGraph
		Set objGraph = New PureAspGraph

		' SET CHART TITLE
		Call objGraph.setTitle("<span class=""charttitle"">" & sTitle & "</span>")
		
		
		' SET CHART XAXISTITLE
		If Trim(sXAxisTitle) <> "" Then
			Call objGraph.setXAxesTitle("<span class=""xaxistitle"">" & sXAxisTitle  & "</span>")
		End If
		
		' SET CHART YAXISTITLE
		If TRIM(sYAxesTitle) <> "" Then
			Call objGraph.setYAxesTitle("<span class=""yaxistitle"">" & sYAxesTitle & "</span>" )
		End If
		
		
		' ADD DATA TO CHART

		' GET INITIAL DATA (X AXIS AND FIRST Y AXIS VALUES)
		Call objGraph.setDataFromRecordset(oRequests, 0, 1)	' 0 IS INDEX OF THE TEXT, 1 IS INDEX OF THE DATA
		oRequests.MoveFirst
		
		' ENUMERATE THRU REMAINING COLUMNS TO ADD Y AXIS DATA
		For iFields = 2 To oRequests.Fields.Count - 1
			Call objGraph.addDataFromRecordset(oRequests, iFields) 'iFIELD IS INDEX OF THE DATA
			'Call objGraph.setBarColor(0, "#fff8855") 'CUSTOMIZE BAR COLOR 
			oRequests.MoveFirst
		Next
		
		' ENUMERATE THRU COLUMNS TO ADD LEGEND FOR Y VALUES
		For iFields = 1 To oRequests.Fields.Count - 1
			Call objGraph.addLabel(oRequests.Fields(iFields).Name)	
		Next


		' ADDITIONAL CHART FORMATTING OPTIONS
		Call objGraph.setType(0)		' 0 = HORIZ, 1= VERTICAL BARS	
		Call objGraph.setBarWidth(15)   ' WIDTH OF BARS ON CHART
		Call objGraph.setBarBorder(1)	' WIDTH OF LINES AROUND BARS ON CHART
		Call objGraph.setBarSpacing(0)	' SPACING BETWEEN BAR GROUPINGS
		Call objGraph.setShowValue(1)	' DISPLAY VALUES AT END OF BARS ON CHART 1=TRUE,0=FALSE
		
		If request("ireport") > 7 Then
			Call objGraph.setFormat("MONEY") ' DISPLAY SELECTED FORMAT
		End If

		' DRAW CHART
		Call objGraph.print()

	End If


End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYMOSTPOPULARREQUESTS()
'------------------------------------------------------------------------------------------------------------
Sub subDisplayMostSubmittedRequests()
		
		' CHART DISPLAY OPTIONS
		syaxistitle = ""
		sxaxistitle = ""
		sTitle = "Most Submitted Requests" & " - " & fromDate & " to " & toDate
		
		' CHART DATA QUERY
		sSQL = "Select [Form Name],Count([Form Name]) as [Number of Requests] from egov_rpt_actionline   " & varWhereClause  & " GROUP BY action_formid,[Form Name] ORDER BY COUNT([Form Name]) DESC"

		' CALL CHART DISPLAY ROUTINE
		response.write "<P>"
		Call SubDisplayChart(sTitle,sxaxistitle,syaxistitle,sSQL)
		response.write "</P>"

End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYREQUESTSBYMONTH()
'------------------------------------------------------------------------------------------------------------
Sub subDisplayRequestsbyMonth()
		
		' CHART DISPLAY OPTIONS
		syaxistitle = ""
		sxaxistitle = "Number of Requests"
		sTitle = "Monthly Status" & " - " & fromDate & " to " & toDate

		' CHART DATA QUERY
		sSQL = "Select LEFT(DATENAME(MONTH,right(yearmonth,2) + '/1/' + LEFT(yearmonth,4)),3) + ' ' + LEFT(yearmonth,4) as [Year\Month], sum(DISMISSED) as DISMISSED, sum(INPROGRESS) as INPROGRESS, sum(resolved) as 'RESOLVED',sum(submitted) as SUBMITTED,  sum(waiting) as 'WAITING' from egov_rpt_status_chart   " & varWhereClause  & " GROUP BY yearmonth  ORDER BY yearmonth desc"

		' CALL CHART DISPLAY ROUTINE
		response.write "<P>"
		Call SubDisplayChart(sTitle,sxaxistitle,syaxistitle,sSQL)
		response.write "</P>"

End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYOPENREQUESTSBYDEPARTMENT()
'------------------------------------------------------------------------------------------------------------
Sub subDisplayOpenRequestsbyDepartment()
		
		' CHART DISPLAY OPTIONS
		syaxistitle = ""
		sxaxistitle = "Number of Days"
		sTitle = "Open Items by Department" & " - " & fromDate & " to " & toDate

		' CHART DATA QUERY
		sSQL = "Select Isnull(Department,'empty'),AVG([Open]) as [Avg Days Open],AVG(lastactivityindays) as [Avg Days Last Activity] from egov_rpt_actionline   " & varWhereClause & sOpenOnly &  " GROUP BY DEPARTMENT  ORDER BY  Department"
	
		' CALL CHART DISPLAY ROUTINE
		response.write "<P>"
		Call SubDisplayChart(sTitle,sxaxistitle,syaxistitle,sSQL)
		response.write "</P>"

End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYREQUESTSTATUSBYDEPARTMENT()
'------------------------------------------------------------------------------------------------------------
Sub subDisplayRequeststatusbyDepartment()
		
		' CHART DISPLAY OPTIONS
		syaxistitle = ""
		sxaxistitle = "Number of Requests"
		sTitle = "Monthly Status by Department" & " - " & fromDate & " to " & toDate

		' CHART DATA QUERY
		sSQL = "Select Department, sum(DISMISSED) as DISMISSED, sum(INPROGRESS) as INPROGRESS, sum(resolved) as 'RESOLVED',sum(submitted) as SUBMITTED,  sum(waiting) as 'WAITING' from egov_rpt_status_chart   " & varWhereClause  & " GROUP BY department  ORDER BY department"

		' CALL CHART DISPLAY ROUTINE
		response.write "<P>"
		Call SubDisplayChart(sTitle,sxaxistitle,syaxistitle,sSQL)
		response.write "</P>"

End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYOPENREQUESTSBYFORM()
'------------------------------------------------------------------------------------------------------------
Sub subDisplayOpenRequestsbyForm()
		
		' CHART DISPLAY OPTIONS
		syaxistitle = ""
		sxaxistitle = "Number of Days"
		sTitle = "Open Items by Form" & " - " & fromDate & " to " & toDate

		' CHART DATA QUERY
		sSQL = "Select Isnull( [FORM NAME],'empty'),AVG([Open]) as [Avg Days Open],AVG(lastactivityindays) as [Avg Days Last Activity] from egov_rpt_actionline   " & varWhereClause & sOpenOnly &  " GROUP BY [FORM NAME]  ORDER BY   [FORM NAME]"
	
		' CALL CHART DISPLAY ROUTINE
		response.write "<P>"
		Call SubDisplayChart(sTitle,sxaxistitle,syaxistitle,sSQL)
		response.write "</P>"

End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYOPENREQUESTSTOTALBYDEPARTMENT()
'------------------------------------------------------------------------------------------------------------
Sub subDisplayOpenRequestsTotalbyDepartment()
		
		' CHART DISPLAY OPTIONS
		syaxistitle = ""
		sxaxistitle = ""
		sTitle = "Open Items by Department" & " - " & fromDate & " to " & toDate

		' CHART DATA QUERY
		sSQL = "Select Isnull(Department,'empty'),Count(department) as [Total Items] from egov_rpt_actionline   " & varWhereClause & sOpenOnly &  " GROUP BY DEPARTMENT  ORDER BY  Count(department) desc,Department"
		
		' CALL CHART DISPLAY ROUTINE
		response.write "<P>"
		Call SubDisplayChart(sTitle,sxaxistitle,syaxistitle,sSQL)
		response.write "</P>"

End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYOPENREQUESTSTOTALBYFORM()
'------------------------------------------------------------------------------------------------------------
Sub subDisplayOpenRequestsTotalbyForm()
		
		' CHART DISPLAY OPTIONS
		syaxistitle = ""
		sxaxistitle = ""
		sTitle = "Open Items by Form" & " - " & fromDate & " to " & toDate

		' CHART DATA QUERY
		sSQL = "Select Isnull( [FORM NAME],'empty'),Count([Form Name]) as [Total Items] from egov_rpt_actionline   " & varWhereClause & sOpenOnly &  " GROUP BY [FORM NAME]  ORDER BY   Count([Form Name]) DESC,[FORM NAME]"
		
		' CALL CHART DISPLAY ROUTINE
		response.write "<P>"
		Call SubDisplayChart(sTitle,sxaxistitle,syaxistitle,sSQL)
		response.write "</P>"

End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYDAILYRECEIPTTOTALS()
'------------------------------------------------------------------------------------------------------------
Sub subDisplayDailyReceiptTotals()
		
		' CHART DISPLAY OPTIONS
		syaxistitle = ""
		sxaxistitle = ""
		sTitle = "Daily Receipt Totals" & " - " & fromDate & " to " & toDate

		' CHART DATA QUERY
		sSQL = "Select paymentdateshort [Payment Date], Cast(sum(credit) as Money) as [Sales in Dollars] from egov_glreport_combined  " & replace(varWhereClause,"[Date Submitted]","paymentdate") & " GROUP BY paymentdateshort,paymenttypename  ORDER BY  paymentdateshort asc"

		
		' CALL CHART DISPLAY ROUTINE
		response.write "<P>"
		Call SubDisplayChart(sTitle,sxaxistitle,syaxistitle,sSQL)
		response.write "</P>"

End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYCATEGORYTOTALSBYMONTH()		
'------------------------------------------------------------------------------------------------------------
Sub subDisplayCategoryTotalsbyMonth()
		
		' CHART DISPLAY OPTIONS
		syaxistitle = ""
		sxaxistitle = ""
		sTitle = "Monthly Category Totals" & " - " & fromDate & " to " & toDate

		' CHART DATA QUERY
		sSQL = "Select LEFT(DATENAME(MONTH,right(yearmonth,2) + '/1/' + LEFT(yearmonth,4)),3) + ' ' + LEFT(yearmonth,4) as [Year\Month], Cast(Sum(Reservation) as Money) as Reservation,Cast(Sum([Pool Pass]) as Money) as [Pool Pass],Cast(Sum(Gift) as Money) as Gift,Cast(Sum([Class\Event]) as Money) as [Class\Event],Cast(Sum([Online Payment Service]) as Money) as 'Online Payment Service' from egov_report_payment_category_chart  " & replace(varWhereClause,"[Date Submitted]","paymentdate")  & " GROUP BY yearmonth ORDER BY yearmonth desc"
		
		' CALL CHART DISPLAY ROUTINE
		response.write "<P>"
		Call SubDisplayChart(sTitle,sxaxistitle,syaxistitle,sSQL)
		response.write "</P>"

End Sub





%>