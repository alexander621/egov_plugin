<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: reporting_module_actionline.asp
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
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

' INITIALIZE AND DECLARE VARIABLES
' SPECIFY FOLDER LEVEL
sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK
If Not UserHasPermission( Session("UserId"), "advactionlinerpt" ) Then
  	response.redirect sLevel & "permissiondenied.asp"
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
  <title>E-Gov Advanced Request Reporting</title>

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


<form action="reporting_module_actionline.asp" method=post name=frmPFilter >

	<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<tr>
			<td><font size="+1"><b>E-Gov Advanced Request Reporting</b></font></td>
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


				<!--BEGIN: OTHER FILTERS-->
				<!--<P>
				
				<%
					' BASE VIEW DATA
					' sSQLBase = "SELECT paymenttypename,category,item,paymentlocationname,account FROM egov_glreport_combined WHERE orgid='" & session("orgid") & "'"
					' GetFieldChoices(sSQLBase)
				%>
				</p>-->
				<!--END: OTHER FILTERS-->
	

				</fieldset>
				<!--END: FILTERS-->



				 <!--BEGIN: PREDEFINED REPORTS-->
				  <fieldset>

					<legend><b>Predefined Reports:</b></legend>

					<P>
					  <Select name="ireport">
							<option value=1 <%if request("ireport") = 1 Then response.write " SELECTED " End If %> > Open Items by Department - List
							<option Value=3 <%if request("ireport") = 3 Then response.write " SELECTED " End If %>> Open Items by Department - Detail
							<option value=2 <%if request("ireport") = 2 Then response.write " SELECTED " End If %>> Open Items by Department - Summary
							<option value=4 <%if request("ireport") = 4 Then response.write " SELECTED " End If %>> Monthly Status - List
							<option value=6 <%if request("ireport") = 6 Then response.write " SELECTED " End If %>> Monthly Status - Detail
							<option value=5 <%if request("ireport") = 5 Then response.write " SELECTED " End If %>> Monthly Status - Summary
							<option value=7 <%if request("ireport") = 7 Then response.write " SELECTED " End If %>> Monthly Status by by Department - List
							<option value=9 <%if request("ireport") = 9 Then response.write " SELECTED " End If %>> Monthly Status by by Department - Detail
							<option value=8 <%if request("ireport") = 8 Then response.write " SELECTED " End If %>> Monthly Status by by Department - Summary
					 
							<%If session("orgid") = 8 Then %>
							<option value=10 <%if request("ireport") = 10 Then response.write " SELECTED " End If %>> Building and Zoning - Property Maintenance Complaints
							<% End If %>
					  
							<option value=11 <%if request("ireport") = 11 Then response.write " SELECTED " End If %>> Past Due (weekdays) List

					  </select>
					  <input class=excelexport type=submit value="View Report"> - (<a href="reporting_module_chart.asp">Click Here to View Available Charts</a>)
					</P>
				 

					</fieldset>
				 <!--END: PREDEFINED REPORTS-->

				
    </td>
  </tr>
	<tr>
 
      <td colspan="3" valign="top">
	  
	  
		<!--BEGIN: DISPLAY RESULTS-->
		<!-- #include file="queries/actionline_queries.asp" //-->

		<%
		
		' DISPLAY RESULTS
		Display_Results sSQL,sOptions
		
		%>
		<!-- END: DISPLAY RESULTS -->
      

	  </td>
       
    </tr>
  </table>


  </form>
  

<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  


</body>
</html>



<!--#Include file="includes/report_display_functions.asp"-->  


