<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../classes/class_global_functions.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: instructor_payments.asp
' AUTHOR: Steve Loar
' CREATED: 08/02/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   08/02/07		Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iSeasonId, iInstructorId, iReportType, sRptTitle, sRptType, sWhereClause

' INITIALIZE AND DECLARE VARIABLES
' SPECIFY FOLDER LEVEL For the header
sLevel = "../" ' Override of value from common.asp


' USER SECURITY CHECK
If Not UserHasPermission( Session("UserId"), "instructor payment rpt" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

If request("classseasonid") = "" Then
	iSeasonId = GetRosterSeasonId()
Else
	iSeasonId = CLng(request("classseasonid"))
End If 

If request("instructorid") = "" Then
	iInstructorId = 0
Else
	iInstructorId = CLng(request("instructorid"))
End If 

If request("reporttype") = "" Then 
	iReportType = CLng(1)
Else
	iReportType = CLng(request("reporttype"))
End If 

If iReportType = CLng(1) Then
	sRptTitle = "Summary"
	sRptType = "Summary"
Else
	sRptTitle = "Detail"
	sRptType = "Detail"
End If 


' BUILD SQL WHERE CLAUSE
sWhereClause = ""
If iSeasonId > CLng(0) Then
	sWhereClause = " AND classseasonid = " & iSeasonId
End If 

If iInstructorId > CLng(0) Then
	sWhereClause = sWhereClause & " AND instructorid = " & iInstructorId
End If 

%>

<html>
<head>
  <title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="reporting.css" />
	<link rel="stylesheet" type="text/css" href="pageprint.css" media="print" />

	<script language="Javascript" src="scripts/tablesort.js"></script>

	<script language="Javascript">
	  <!--
		window.onload = function()
		{
		  //factory.printing.header = "Printed on &d"
		  //factory.printing.footer = "&bPrinted on &d - Page:&p/&P";
		  //factory.printing.portrait = true;
		  //factory.printing.leftMargin = 0.5;
		  //factory.printing.topMargin = 0.5;
		  //factory.printing.rightMargin = 0.5;
		  //factory.printing.bottomMargin = 0.5;
		 
		  // enable control buttons
		  //var templateSupported = factory.printing.IsTemplateSupported();
		  //var controls = idControls.all.tags("input");
		  //for ( i = 0; i < controls.length; i++ ) 
		  //{
		//	controls[i].disabled = false;
		//	if ( templateSupported && controls[i].className == "ie55" )
		//	  controls[i].style.display = "inline";
		  //}
		}

	  //-->
	</script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN: THIRD PARTY PRINT CONTROL-->
<div id="idControls" class="noprint">
	<input type="button" class="button" onclick="javascript:window.print();" value="Print" />
<%
'	<input disabled type="button" value="Print the page" onclick="factory.printing.Print(true)" />&nbsp;&nbsp;
'	<input class="ie55" disabled type="button" value="Print Preview..." onclick="factory.printing.Preview()" />
%>
</div>

<%
'<object id="factory" viewastext  style="display:none"
'  classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
'   codebase="../includes/smsx.cab#Version=6,3,434,12">
'</object>
%>
<!--END: THIRD PARTY PRINT CONTROL-->

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

	<form action="instructor_payments.asp" method="post" name="frmPFilter">

	<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<tr>
			<td><font size="+1"><strong>Instructor Payments <%=sRptTitle%></strong></font></td>
		</tr>
		<tr>
			<td>
				<fieldset>
					<legend><strong>Select</strong></legend>
					<p>
						<strong>Season: </strong><% ShowClassSeasonFilterPicks iSeasonId %>&nbsp;&nbsp;
						<strong>Report Type: </strong><% ShowReportTypes iReportType %>
					</p>
					<p>
						<strong>Instructor: </strong><% ShowInstructorPicks iInstructorId %>
					</p>
					<p>
						<input class="button" type="submit" value="View Report" />
						&nbsp;&nbsp;<input type="button" class="button" value="Download to Excel" onClick="location.href='instructor_payments_export.asp?seasonid=<%=iSeasonId%>&instructorid=<%=iInstructorId%>&reporttype=<%=iReportType%>'" />
					</p>

				</fieldset>
				<!--END: FILTERS-->
		    </td>
		</tr>
		<tr>
 
			<td colspan="3" valign="top">
	  
				<!--BEGIN: DISPLAY RESULTS-->
				<%
				
				If sRptType = "Detail" Then
					DisplayDetails sWhereClause
				Else
					DisplaySummary sWhereClause
				End If 
				
				%>
				<!-- END: DISPLAY RESULTS -->
      
			</td>
		 </tr>
	</table>
  </form>
	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' Sub DisplaySummary ( sWhereClause ) 
'------------------------------------------------------------------------------------------------------------
Sub DisplaySummary ( ByVal sWhereClause ) 
	Dim sSql, oPayments, iOldInstructorId, dSubTotal, dGrandTotal

	iOldInstructorId = CLng(0)
	dGrandTotal = CDbl(0.00)
	dSubTotal = CDbl(0.00)

	sSql = "SELECT instructorid, firstname, lastname, classname, activityno, startdate, enddate, SUM(instructorpay) AS instructorpay "
	sSql = sSql & " FROM egov_instructor_payment_details "
	sSql = sSql & " WHERE orgid = " & session("orgid") & sWhereClause
	sSql = sSql & " GROUP BY instructorid, firstname, lastname, classname, activityno, startdate, enddate "
	sSql = sSql & " ORDER BY lastname, firstname, classname, activityno"

'	response.write sSql

	Set oPayments = Server.CreateObject("ADODB.Recordset")
	oPayments.Open sSQL, Application("DSN"), 3, 1

	If oPayments.EOF then
		' EMPTY
		response.write "<p>No instructor payments found for your selection criteria.</p>"
	Else
		response.write vbcrlf & "<div class=""receiptpaymentshadow"">"
		response.Write vbcrlf & "<table cellspacing=""0"" cellpadding=""2""  border=""0"" width=""100%"" class=""receiptpayment"">"
		response.write vbcrlf & "<tr class=""tablelist""><th>Instructor</th><th>Class</th><th>Activity No.</th><th>Start<br />Date</th>"
		response.write "<th>End<br />Date</th><th>Payment</th></tr>"

		bgcolor = "#eeeeee"

		Do While Not oPayments.EOF
		If bgcolor="#eeeeee" Then
				bgcolor="#ffffff" 
			Else
				bgcolor="#eeeeee"
			End If			
			'response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>"

			If iOldInstructorId <> CLng(oPayments("instructorid")) Then 
				If iOldInstructorId <> CLng(0) Then
					' Sub total row for instructor
					response.write vbcrlf & "<tr class=""totalrow""><td colspan=""5"" align=""right"">" & sInstructorName & " Total:</td>"
					response.write "<td align=""right"">" & FormatNumber(dSubTotal,2,,,0) & "</td>"
					response.write "</tr>"
				End If 
				iOldInstructorId = CLng(oPayments("instructorid"))
				dSubTotal = CDbl(0.00)
				bgcolor="#ffffff" 
				response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>"
				sInstructorName = oPayments("firstname") & " " & oPayments("lastname")
				response.write "<td align=""left"" nowrap=""nowrap"">" & sInstructorName & "</td>"
			Else
				response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>"
				response.write "<td align=""left"">&nbsp;</td>"
			End If 
			' Print out line
			
			response.write "<td align=""left"">" & oPayments("classname") & "</td>"
			response.write "<td align=""center"">" & oPayments("activityno") & "</td>"
			response.write "<td align=""center"">" & oPayments("startdate") & "</td>"
			response.write "<td align=""center"">" & oPayments("enddate") & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oPayments("instructorpay"),2,,,0) & "</td>"
			response.write "</tr>"

			dSubTotal = dSubTotal + CDbl(FormatNumber(oPayments("instructorpay"),2,,,0))
			dGrandTotal = dGrandTotal + CDbl(FormatNumber(oPayments("instructorpay"),2,,,0))
			
			oPayments.MoveNext
		Loop 
		' Sub total row for final instructor
		If iOldInstructorId <> CLng(0) Then
			response.write vbcrlf & "<tr class=""totalrow""><td colspan=""5"" align=""right"">" & sInstructorName & " Total:</td>"
			response.write "<td align=""right"">" & FormatNumber(dSubTotal, 2,,,0) & "</td>"
			response.write "</tr>"
		End If 

		' Total for all instructors
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""5"" align=""right"">Totals:</td>"
		response.write "<td align=""right"">" & FormatNumber(dGrandTotal,2,,,0) & "</td>"
		response.write "</tr>"

		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>"
	End If 

	oPayments.Close
	Set oPayments = Nothing 

End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub DisplayDetails( sWhereClause )
'------------------------------------------------------------------------------------------------------------
Sub DisplayDetails( ByVal sWhereClause )
	Dim sSql, oDisplay, iOldInstructorId, dSubTotal, dGrandTotal, sInstructorName

	iOldInstructorId = CLng(0) 
	dSubTotal = CDbl(0.00)
	dGrandTotal = CDbl(0.00)


	sSql = "SELECT  instructorid, firstname, lastname, classname, activityno, startdate, enddate, CASE isdropin WHEN 1 THEN 'Yes' WHEN 0 THEN '&nbsp;' END AS isdropin, "
	sSql = sSql & " paymentid, paymentdate, pricetypename, instructorpercent, amount, entrytype, instructorpay "
	sSql = sSql & " FROM egov_instructor_payment_details "
	sSql = sSql & " WHERE orgid = " & session("orgid") & sWhereClause
	sSql = sSql & " ORDER BY lastname, firstname, instructorid, classname, activityno"

	Set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open sSQL, Application("DSN"), 3, 1

	If oDisplay.EOF then
		' EMPTY
		response.write "<p>No instructor payments found for your selection criteria.</p>"
	Else
		response.write vbcrlf & "<div class=""receiptpaymentshadow"">"
		response.Write vbcrlf & "<table cellspacing=""0"" cellpadding=""2""  border=""0"" width=""100%"" class=""receiptpayment"">"
		response.write vbcrlf & "<tr class=""tablelist""><th>Instructor</th><th>Class</th><th>Activity No.</th><th>Start<br />Date</th>"
		response.write "<th>End<br />Date</th><th>Receipt No.</th><th>Purchase<br />Date</th><th>Drop In</th><th>Pricing</th>"
		response.write "<th>Amount</th><th>Instr. %</th>" '<th>Credit<br />/Debit</th>
		response.write "<th>Payment</th></tr>"

		bgcolor = "#eeeeee"
		iOldAccountId = CLng(0)

		Do While Not oDisplay.EOF
			If bgcolor="#eeeeee" Then
				bgcolor="#ffffff" 
			Else
				bgcolor="#eeeeee"
			End If			
			If iOldInstructorId <> CLng(oDisplay("instructorid")) Then 
				' Put out a sub total row
				If iOldInstructorId <> CLng(0) Then
					response.write vbcrlf & "<tr class=""totalrow""><td colspan=""11"" align=""right"">" & sInstructorName & " Total:</td>"
					response.write "<td align=""right"">" & FormatNumber(dSubTotal, 2) & "</td>"
					response.write "</tr>"
				End If 
				iOldInstructorId = CLng(oDisplay("instructorid"))
				dSubTotal = CDbl(0.00)
				bgcolor="#ffffff" 
				response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>"
				sInstructorName = oDisplay("firstname") & " " & oDisplay("lastname")
				response.write "<td align=""left"" nowrap=""nowrap"">" & sInstructorName & "</td>"
			Else
				response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>"
				' Need place holders 
				response.write "<td>&nbsp;</td>"
			End If 

			response.write "<td align=""left"">" & oDisplay("classname") & "</td>"
			response.write "<td align=""center"">" & oDisplay("activityno") & "</td>"
			response.write "<td align=""center"">" &  FormatDateTime(oDisplay("startdate"),2) & "</td>"
			response.write "<td align=""center"">" &  FormatDateTime(oDisplay("enddate"),2) & "</td>"
			response.write "<td align=""center""><a href=""../classes/view_receipt.asp?iPaymentId=" & oDisplay("paymentid") & """>" & oDisplay("paymentid") & "</a></td>"
			response.write "<td align=""center"">" & FormatDateTime(oDisplay("paymentdate"),2) & "</td>"
			response.write "<td align=""center"">" & oDisplay("isdropin") & "</td>"
			response.write "<td align=""center"">" & oDisplay("pricetypename") & "</td>"
			
			response.write "<td align=""center"">" & FormatNumber(oDisplay("amount"),2) & "</td>"
			response.write "<td align=""center"">" & oDisplay("instructorpercent") & "%</td>"
			'response.write "<td align=""center"">" & oDisplay("entrytype") & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("instructorpay"),2) & "</td>"
			response.write "</tr>"

			dSubTotal = dSubTotal + CDbl(FormatNumber(oDisplay("instructorpay"),2,,,0))
			dGrandTotal = dGrandTotal + CDbl(FormatNumber(oDisplay("instructorpay"),2,,,0))
			
			oDisplay.MoveNext
		Loop 

		' Put out a sub total row
		If iOldInstructorId <> CLng(0) Then
			response.write vbcrlf & "<tr class=""totalrow""><td colspan=""11"" align=""right"">" & sInstructorName & " Total:</td>"
			response.write "<td align=""right"">" & FormatNumber(dSubTotal, 2) & "</td>"
			response.write "</tr>"
		End If 
		' Totals Row
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""11"" align=""right"">Total:</td>"
		response.write "<td align=""right"">" & FormatNumber(dGrandTotal,2) & "</td>"
		response.write "</tr>"

		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>"

	End If 

	oDisplay.Close
	Set oDisplay = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowClassSeasonFilterPicks( iClassSeasonId )
'--------------------------------------------------------------------------------------------------
Sub ShowClassSeasonFilterPicks( ByVal iClassSeasonId )
	Dim sSql, oSeasons

	sSQL = "Select C.classseasonid, C.seasonname From egov_class_seasons C, egov_seasons S  "
	sSql = sSql & " Where C.seasonid = S.seasonid and orgid = " & SESSION("orgid") & " ORDER BY C.seasonyear desc, S.displayorder desc, C.seasonname"
	' C.isclosed = 0 and -- This should include all for looking at old classes. Called from edit_class.asp

	Set oSeasons = Server.CreateObject("ADODB.Recordset")
	oSeasons.Open sSQL, Application("DSN"), 0, 1
	
	If Not oSeasons.EOF Then
		response.write vbcrlf & "<select name=""classseasonid"">" 
		Do While NOT oSeasons.EOF
			response.write vbcrlf & "<option value=""" & oSeasons("classseasonid") & """ "  
			If clng(iClassSeasonId) = clng(oSeasons("classseasonid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oSeasons("seasonname") & "</option>"
			oSeasons.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If

	oSeasons.close
	Set oSeasons = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetMaxClassSeasonId()
'--------------------------------------------------------------------------------------------------
Function GetMaxClassSeasonId()
	Dim sSql, oSeasons

	sSQL = "Select MAX(classseasonid) as classseasonid From egov_class_seasons Where orgid = " & SESSION("orgid") 

	Set oSeasons = Server.CreateObject("ADODB.Recordset")
	oSeasons.Open sSQL, Application("DSN"), 0, 1

	If Not oSeasons.EOF Then
		GetMaxClassSeasonId = clng(oSeasons("classseasonid"))
	Else
		GetMaxClassSeasonId = clng(0)
	End If 
	
	oSeasons.close
	Set oSeasons = Nothing

End Function 


'------------------------------------------------------------------------------------------------------------
' Sub ShowReportTypes( iReportType )
'------------------------------------------------------------------------------------------------------------
Sub ShowReportTypes( ByVal iReportType )
	
	response.write vbcrlf & "<select name=""reporttype"">"

	response.write vbcrlf & "<option value=""1"""
	If CLng(1) = CLng(iReportType) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Summary</option>"

	response.write vbcrlf & "<option value=""2"""
	If CLng(2) = CLng(iReportType) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Detail</option>"

	response.write vbcrlf & "</select>"
	
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowInstructorPicks( iInstructorId )
'--------------------------------------------------------------------------------------------------
Sub ShowInstructorPicks( ByVal iInstructorId )
	Dim sSql, oInstructor

	sSQL = "SELECT * FROM EGOV_CLASS_INSTRUCTOR WHERE ORGID = " & SESSION("ORGID") & " ORDER BY lastname, firstname"

	Set oInstructor = Server.CreateObject("ADODB.Recordset")
	oInstructor.Open sSQL, Application("DSN"), 0, 1
	
	If not oInstructor.EOF Then
		response.write vbcrlf & "<select name=""instructorid"">"
		response.write vbcrlf & "<option value=""0"" >All Instructors</option>"

		Do While NOT oinstructor.EOF 
			response.write vbcrlf & "<option value=""" & oInstructor("instructorid") & """ "  
			If clng(iInstructorId) = clng(oInstructor("instructorid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " >" & oInstructor("lastname") & ", " & oInstructor("firstname")& "</option>"
			oInstructor.MoveNext
		Loop

		response.write vbcrlf & "</select>"

	End If

	oInstructor.close
	Set oInstructor = Nothing

End Sub




%>
