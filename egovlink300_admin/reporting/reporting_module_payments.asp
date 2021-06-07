<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: payment_reporting_module.asp
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/10/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   1/10/2007	JOHN STULLENBERGER - INITIAL VERSION
' 1.1	7/13/2007	Steve Loar - Added the Payment Method Reports
' 2.0	05/05/2010	Steve Loar - Adding rentals into this series of reports
' 2.1	7/14/2010	Steve Loar - Added in Memberships and changed CSV to Excel export
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

' INITIALIZE AND DECLARE VARIABLES
' SPECIFY FOLDER LEVEL
sLevel = "../" ' Override of value from common.asp


' USER SECURITY CHECK
If Not UserHasPermission( Session("UserId"), "advpaymentrpt" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 


' PROCESS REPORT FILTER VALUES
' PROCESS DATE VALUES
fromDate = Request("fromDate")
toDate = Request("toDate")
today = Date()

' IF EMPTY DEFAULT TO CURRENT TO DATE
If toDate = "" or IsNull(toDate) or not isDate(toDate) Then toDate = today End If
If fromDate = "" or IsNull(fromDate) or not isDate(fromDate) Then fromDate = cdate(Month(today)& "/1/" & Year(today)) End If

' BUILD SQL WHERE CLAUSE
varWhereClause = " WHERE (paymentDate >= '" & fromDate & "' AND paymentDate <= '" & DateAdd("d",1,toDate) & "') "
varWhereClause = varWhereClause & " AND orgid = " & session("orgid") '& " "
%>

<html>
<head>
  <title>E-Gov Advanced Payment Reporting</title>

	<link rel="stylesheet" type="text/css" href="reporting.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />

	<script language="JavaScript" src="../scripts/jquery-1.7.2.min.js"></script>
	<script language="Javascript" src="scripts/tablesort.js"></script>

	<script language="Javascript">
	  <!--
		function doCalendar(ToFrom) {
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  //eval('window.open("../recreation/gr_calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		  eval('window.open("calendarpicker.asp?updatefield=' + ToFrom + '&date=' + $("#" + ToFrom ).val() + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}
	  //-->
	</script>

	<script language="Javascript" src="scripts/dates.js"></script>

</head>

<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->

	<form action="reporting_module_payments.asp" method="post" name="frmPFilter">

	<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<tr>
			<td><font size="+1"><b>E-Gov Advanced Payment Reporting</b></font></td>
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
						<td  align="right"> <b>Payment Date: </td>
						<td>
							<input type=text id="fromDate" name="fromDate" value="<%=fromDate%>">
							<a href="javascript:void doCalendar('fromDate');"><img src="../images/calendar.gif" border="0"></a>		 
						</td>
						<td>&nbsp;</td>
						<td >
							<b>To:</b> 
						</td>
						<td>
							<input type="text" id="toDate" name="toDate" value="<%=toDate%>">
							<a href="javascript:void doCalendar('toDate');"><img src="../images/calendar.gif" border="0"></a>
						</td>
						<td>&nbsp;</td>
						<td><%DrawDateChoices "Dates" %></td>
					</tr>
				</table>
				</p>
				<!--END: DATE FILTERS-->

				</fieldset>
				<!--END: FILTERS-->

				 <!--BEGIN: PREDEFINED REPORTS-->
				  <fieldset>

					<legend><b>Predefined Reports:</b></legend>

					<p>
					  <Select name="ireport">
						<option value="1" <%if request("ireport") = 1 Then response.write " selected=""selected"" " End If %> > Daily Receipt - List</option>
						<option Value="3" <%if request("ireport") = 3 Then response.write " selected=""selected"" " End If %>> Daily Receipt - Detail</option>
						<option value="2" <%if request("ireport") = 2 Then response.write " selected=""selected"" " End If %>> Daily Receipt - Summary</option>
						<option value="4" <%if request("ireport") = 4 Then response.write " selected=""selected"" " End If %>> Monthly Revenue by Category - List</option>
						<option value="6" <%if request("ireport") = 6 Then response.write " selected=""selected"" " End If %>> Monthly Revenue by Category - Detail</option>
						<option value="5" <%if request("ireport") = 5 Then response.write " selected=""selected"" " End If %>> Monthly Revenue by Category - Summary</option>
						<option value="7" <%if request("ireport") = 7 Then response.write " selected=""selected"" " End If %>> Monthly Revenue by Source - List</option>
						<option value="9" <%if request("ireport") = 9 Then response.write " selected=""selected"" " End If %>> Monthly Revenue by Source - Detail</option>
						<option value="8" <%if request("ireport") = 8 Then response.write " selected=""selected"" " End If %>> Monthly Revenue by Source - Summary</option>
						<option value="10" <%if request("ireport") = 10 Then response.write " selected=""selected"" " End If %> > Daily Payment Method - List</option>
						<option Value="12" <%if request("ireport") = 12 Then response.write " selected=""selected"" " End If %>> Daily Payment Method - Detail</option>
						<option value="11" <%if request("ireport") = 11 Then response.write " selected=""selected"" " End If %>> Daily Payment Method - Summary</option>
					  </select>
					  <input type="submit" class="button excelexport" value="View Report" />
					</p>
					</fieldset>
				 <!--END: PREDEFINED REPORTS-->
    </td>
  </tr>
	<tr>
 
      <td colspan="3" valign="top">
	  
		<!--BEGIN: DISPLAY RESULTS-->
		<!-- #include file="queries/payment_queries.asp" //-->

<%
		' DISPLAY RESULTS
		Display_Results sSql, sOptions
		
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
