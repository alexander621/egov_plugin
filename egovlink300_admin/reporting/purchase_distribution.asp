<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: account_distribution.asp
' AUTHOR: Steve Loar
' CREATED: 07/19/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  7/19/07		Steve Loar - INITIAL VERSION
' 1.1  6/25/08  David Boyer - Added "season" to search criteria, results list, and broke out records sub-totals by season.
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iLocationId, iAdminUserId, iPaymentLocationId, iReportType, sRptTitle, sRptType

' INITIALIZE AND DECLARE VARIABLES
' SPECIFY FOLDER LEVEL
sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK
PageDisplayCheck "purchase distribution rpt", sLevel	' In common.asp


' PROCESS REPORT FILTER VALUES
' PROCESS DATE VALUES

fromDate = Request("fromDate")
toDate   = Request("toDate")
today    = Date()

'IF EMPTY DEFAULT TO CURRENT TO DATE
If toDate = "" Or IsNull(toDate) Then 
   toDate = today
End If 

If fromDate = "" or IsNull(fromDate) Then 
	'fromDate = cdate(Month(today)& "/1/" & Year(today)) 
	fromDate = today
End If

If request("locationid") = "" Then
	iLocationId = 0
Else
	iLocationId = CLng(request("locationid"))
End If 

If request("adminuserid") = "" Then
	iAdminUserId = 0
Else
	iAdminUserId = CLng(request("adminuserid"))
End If 

If request("paymentlocationid") = "" Then
	iPaymentLocationId = 0
Else
	iPaymentLocationId = CLng(request("paymentlocationid"))
End If 

If request("reporttype") = "" Then 
	iReportType = CLng(1)
Else 
	iReportType = CLng(request("reporttype"))
End If 

If iReportType = CLng(1) Then 
	sRptTitle = "Summary"
	sRptType  = "Summary"
Else 
	sRptTitle = "Detail"
	sRptType  = "Detail"
End If 

If request("classseasonid") <> "" Then 
	iClassSeasonId = request("classseasonid")
Else 
	iClassSeasonId = "0"
End If 

'BUILD SQL WHERE CLAUSE
varWhereClause = " AND (P.paymentDate >= '" & fromDate & "' AND P.paymentDate <= '" & DateAdd("d",1,toDate) & "') "
varWhereClause = varWhereClause & " AND A.orgid = " & session("orgid") 
If iLocationId > 0 Then 
	varWhereClause = varWhereClause & " AND P.adminlocationid = " & iLocationId
End If 

If iAdminUserId > 0 Then
	varWhereClause = varWhereClause & " AND adminuserid = " & iAdminUserId
End If 

If iPaymentLocationId > 0 Then 
	If iPaymentLocationId = CLng(2) Then 
		varWhereClause = varWhereClause & " AND P.paymentlocationid = 3 " 
	Else 
		varWhereClause = varWhereClause & " AND P.paymentlocationid < 3 " 
	End If 
End If 

'Determine which season has been selected
If CLng(iClassSeasonId) <> CLng(0) Then 
	varWhereClause = varWhereClause & " AND C.classseasonid = " & iClassSeasonID
End if 

%>
<html>
<head>
  <title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="reporting.css" />
	<link rel="stylesheet" type="text/css" href="pageprint.css" media="print" />

	<script language="JavaScript" src="../scripts/jquery-1.7.2.min.js"></script>

	<script language="javascript" src="scripts/tablesort.js"></script>
	<script language="Javascript" src="../scripts/getdates.js"></script>

	<script language="javascript">
	  <!--
		function doCalendar(ToFrom) {
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  //eval('window.open("../recreation/gr_calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		  eval('window.open("calendarpicker.asp?updatefield=' + ToFrom + '&date=' + $("#" + ToFrom ).val() + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
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
</div>
<!--END: THIRD PARTY PRINT CONTROL-->

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

<form action="purchase_distribution.asp" method="post" name="frmPFilter">

<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
  <tr>
      <td><font size="+1"><strong>Purchase Distribution <%=sRptTitle%></strong></font></td>
  </tr>
  <tr>
      <td>
          <fieldset>
            <legend><strong>Select</strong></legend>

       			<!--BEGIN: FILTERS-->
      				<!--BEGIN: DATE FILTERS-->
          <p>
          <table border="0" cellpadding="0" cellspacing="0">
            <tr>
                <td><strong>Payment Date: </strong></td>
                <td>
                    <input type="text" id="fromDate" name="fromDate" value="<%=fromDate%>" size="10" maxlength="10" />
                    <a href="javascript:void doCalendar('fromDate');"><img src="../images/calendar.gif" border="0" /></a>		 
                </td>
                <td>&nbsp;</td>
                <td><strong>To:</strong></td>
                <td>
                    <input type="text" id="toDate" name="toDate" value="<%=toDate%>" size="10" maxlength="10" />
                    <a href="javascript:void doCalendar('toDate');"><img src="../images/calendar.gif" border="0" /></a>
                </td>
                <td>&nbsp;</td>
                <td><%DrawDateChoices "Date" %></td>
            </tr>
          </table>
          </p>
          <p>
          <strong>Admin Location: </strong><% ShowAdminLocations iLocationId %>&nbsp;&nbsp;
          <strong>Admin: </strong><% ShowAdminUsers iAdminUserId %>
          </p>
          <p>
          <strong>Payment Location: </strong><% ShowPaymentLocations iPaymentLocationId %> &nbsp;&nbsp;
          <strong>Report Type: </strong><% ShowReportTypes iReportType %>&nbsp;&nbsp;
          </p>
          <p>
          <strong>Season: </strong><% ShowClassSeasons iClassSeasonID %>
          </p>
          <!--END: DATE FILTERS-->
          <p>
          <input class="button" type="submit" value="View Report" />
          &nbsp;&nbsp;<input type="button" class="button" value="Download to Excel" onClick="location.href='purchase_distribution_export.asp?fromDate=<%=fromDate%>&toDate=<%=toDate%>&locationid=<%=iLocationId%>&adminuserid=<%=iAdminUserId%>&paymentlocationid=<%=iPaymentLocationId%>&reporttype=<%=iReportType%>&classseasonid=<%=iClassSeasonID%>'" />
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
			DisplayDetails varWhereClause
        Else 
            DisplaySummary varWhereClause
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
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' DrawDateChoices sName
'------------------------------------------------------------------------------------------------------------
Sub DrawDateChoices( ByVal sName )

	response.write vbcrlf & "<select onChange=""getDates(this.value, '" & sName & "');"" class=""calendarinput"" name=""" & sName & """>"
	response.write vbcrlf & "<option value=""0"">Or Select Date Range from Dropdown...</option>" 
	response.write vbcrlf & "<option value=""11"">This Week</option>"
	response.write vbcrlf & "<option value=""12"">Last Week</option>"
	response.write vbcrlf & "<option value=""1"">This Month</option>"
	response.write vbcrlf & "<option value=""2"">Last Month</option>"
	response.write vbcrlf & "<option value=""3"">This Quarter</option>"
	response.write vbcrlf & "<option value=""4"">Last Quarter</option>"
	response.write vbcrlf & "<option value=""6"">Year to Date</option>"
	response.write vbcrlf & "<option value=""5"">Last Year</option>"
	response.write vbcrlf & "<option value=""7"">All Dates to Date</option>"
	response.write vbcrlf & "</select>" 

End Sub 


'------------------------------------------------------------------------------------------------------------
' DisplayDetails sWhereClause 
'------------------------------------------------------------------------------------------------------------
Sub DisplayDetails( ByVal sWhereClause )
	Dim sSql, oRs, oDisplay, iOldAccountId, dSubTotal, dGrandTotal

	iOldAccountId = CLng(0) 
	dSubTotal     = CDbl(0.00)
	dGrandTotal   = CDbl(0.00)

	sSql = "SELECT A.accountname, A.accountnumber, L.paymentid, C.classname, T.activityno, L.accountid, L.amount, C.classseasonid "
	sSql = sSql & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P, "
	sSql = sSql & " egov_class_list CL, egov_class_time T, egov_class C "
	sSql = sSql & " WHERE ispaymentaccount = 0  AND L.paymentid = P.paymentid "
	sSql = sSql & " AND A.accountid = L.accountid  AND CL.classlistid = L.itemid "
	sSql = sSql & " AND CL.classtimeid = T.timeid AND L.entrytype = 'credit' "
	sSql = sSql & " AND C.classid = CL.classid " & sWhereClause 
	sSql = sSql & " ORDER BY A.accountname, A.accountnumber, L.accountid, L.paymentid, T.activityno, C.classseasonid"
'	response.write sSql & "<br /><br />"

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If oRs.EOF Then 
 		'EMPTY
  		response.write "<p>No account activity found.</p>" 
	Else 
		'Got some data now make a holding recordset
		Set oDisplay = server.CreateObject("ADODB.RECORDSET") 
		oDisplay.fields.append "accountid", adInteger, , adFldUpdatable
		oDisplay.fields.append "accountname", adVarChar, 50, adFldUpdatable
		oDisplay.fields.append "accountnumber", adVarChar, 20, adFldUpdatable
		oDisplay.fields.append "classseasonid", adInteger, , adFldUpdatable
		oDisplay.fields.append "receiptno", adInteger, , adFldUpdatable
		oDisplay.fields.append "classname", adVarChar, 255, adFldUpdatable
		oDisplay.fields.append "activityno", adVarChar, 20, adFldUpdatable
		oDisplay.fields.append "amount", adCurrency, , adFldUpdatable

		oDisplay.CursorLocation = 3
		'oDisplay.CursorType = 3

		oDisplay.open 

		'Loop through and build the display recordset.
		Do While Not oRs.EOF
			oDisplay.addnew 
			oDisplay("accountid")     = oRs("accountid")
			oDisplay("accountname")   = oRs("accountname")
			oDisplay("accountnumber") = oRs("accountnumber")
			oDisplay("classseasonid") = oRs("classseasonid")
			oDisplay("receiptno")     = oRs("paymentid")
			oDisplay("classname")     = oRs("classname")
			oDisplay("activityno")    = oRs("activityno")
			oDisplay("amount")        = CDbl(oRs("amount"))

			oDisplay.Update
			oRs.MoveNext
		Loop 
 		'sort the Display recordset
	 	'oDisplay.Sort = "accountname ASC, accountnumber ASC, accountid ASC, receiptno ASC"

		'Show results
		oDisplay.MoveFirst
		response.write vbcrlf & "<div class=""receiptpaymentshadow"">" 
		response.write vbcrlf & "<table cellspacing=""0"" cellpadding=""2""  border=""0"" width=""100%"" class=""receiptpayment"">"
		response.write vbcrlf & "<tr class=""tablelist"">"
		response.write "<th>Account Name</th><th>Account No.</th><th>Season</th><th>Receipt No.</th><th>Class Name</th><th>Activity No.</th><th>Amount</th>"
		response.write "</tr>"

		bgcolor = "#eeeeee"
		iOldAccountId = CLng(0)

		Do While Not oDisplay.EOF
			If bgcolor="#eeeeee" Then 
				bgcolor="#ffffff" 
			Else 
				bgcolor="#eeeeee"
			End If 

			response.write "<tr bgcolor=""" &  bgcolor  & """>"

			If iOldAccountId <> CLng(oDisplay("accountid")) Then 
				'Put out a sub total row
				If iOldAccountId <> CLng(0) Then 
					response.write "<tr class=""totalrow"">"
					response.write "<td colspan=""6"" align=""right"">Sub-Total:</td>"
					response.write "<td align=""right"">" & FormatNumber(dSubTotal, 2) & "</td>" 
					response.write "</tr>"
				End If 

				response.write "<td align=""left"">"   & oDisplay("accountname")   & "</td>" 
				response.write "<td align=""center"">" & oDisplay("accountnumber") & "</td>"

				iOldAccountId = CLng(oDisplay("accountid"))
				dSubTotal     = CDbl(0.00)
			Else 
				'Need place holders 
				response.write "<td>&nbsp;</td>"
				response.write "<td>&nbsp;</td>"
			End If 

			response.write "<td>" & getSeasonName(oDisplay("ClassSeasonID")) & "</td>"
			response.write "<td align=""center"">"
			response.write "<a href=""../classes/view_receipt.asp?iPaymentId=" & oDisplay("receiptno") & """>" & oDisplay("receiptno") & "</a>"
			response.write "</td>"
			response.write "<td>" & oDisplay("classname") & "</td>"
			response.write "<td align=""center"">" & oDisplay("activityno") & "</td>" 
			response.write "<td align=""right"">" & FormatNumber(oDisplay("amount"), 2) & "</td>" 

			dSubTotal   = dSubTotal + CDbl(oDisplay("amount"))
			dGrandTotal = dGrandTotal + CDbl(oDisplay("amount"))

			response.write "</tr>" 

			oDisplay.MoveNext
		Loop   

		'Put out a sub total row
		If iOldAccountId <> CLng(0) Then 
			response.write "<tr class=""totalrow"">"
			response.write "<td colspan=""6"" align=""right"">Sub-Total:</td>"
			response.write "<td align=""right"">" & FormatNumber(dSubTotal, 2) & "</td>"
			response.write "</tr>"
		End If 

		'Totals Row
		response.write vbcrlf & "<tr class=""totalrow"">" 
		response.write "<td colspan=""6"" align=""right"">Totals:</td>" 
		response.write "<td align=""right"">" & FormatNumber(dGrandTotal,2) & "</td>"
		response.write "</tr>"
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>"

		oDisplay.Close
		Set oDisplay = Nothing 

	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' DisplaySummary sWhereClause
'------------------------------------------------------------------------------
Sub DisplaySummary( ByVal sWhereClause )
	Dim sSql, oRs, oDisplay, iOldAccountId, dSubTotal, dGrandTotal

	iOldAccountId = CLng(0) 
	dSubTotal     = CDbl(0.00)
	dGrandTotal   = CDbl(0.00)

	'	sSql = "SELECT A.accountname, A.accountnumber, L.paymentid, T.activityno, L.accountid, L.amount "
	sSql = "SELECT A.accountname, A.accountnumber, L.accountid, C.classseasonid "
	sSql = sSql & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P, egov_class_list CL, "
	sSql = sSql & " egov_class_time T, egov_class C "
	sSql = sSql & " WHERE ispaymentaccount = 0 AND L.paymentid = P.paymentid "
	sSql = sSql & " AND A.accountid = L.accountid AND CL.classlistid = L.itemid "
	sSql = sSql & " AND CL.classtimeid = T.timeid AND L.entrytype = 'credit' "
	sSql = sSql & " AND C.classid = CL.classid " & sWhereClause
	sSql = sSql & " GROUP BY A.accountname, A.accountnumber, L.accountid, C.classseasonid "
	sSql = sSql & " ORDER BY A.accountname, A.accountnumber, L.accountid, C.classseasonid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If oRs.EOF Then 
		'EMPTY
		response.write "<p>No account activity found.</p>"
	Else 
		'Got some data now make a holding recordset
		Set oDisplay = server.CreateObject("ADODB.RECORDSET") 
		oDisplay.fields.append "accountid", adInteger, , adFldUpdatable
		oDisplay.fields.append "accountname", adVarChar, 50, adFldUpdatable
		oDisplay.fields.append "accountnumber", adVarChar, 20, adFldUpdatable
		oDisplay.fields.append "classseasonid", adInteger, , adFldUpdatable

		oDisplay.CursorLocation = 3
		oDisplay.Open 

		'Loop through and build the display recordset.
		Do While Not oRs.EOF
			oDisplay.addnew 
			oDisplay("accountid")     = oRs("accountid")
			oDisplay("accountname")   = oRs("accountname")
			oDisplay("accountnumber") = oRs("accountnumber")
			oDisplay("classseasonid") = oRs("classseasonid")
			oDisplay.Update
			oRs.MoveNext
		Loop 

		'Show results
		oDisplay.MoveFirst
		response.write vbcrlf & "<div class=""receiptpaymentshadow"">"
		response.write vbcrlf & "<table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"" class=""receiptpayment"">"
		response.write vbcrlf & "<tr class=""tablelist""><th>Account Name</th><th>Account No.</th><th>Season</th><th>Amount</th></tr>"

		bgcolor       = "#eeeeee"
		iOldAccountId = CLng(0)

		Do While Not oDisplay.EOF
			If bgcolor="#eeeeee" Then 
				bgcolor="#ffffff" 
			Else 
				bgcolor="#eeeeee"
			End If 

			response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>"

			If iOldAccountId <> CLng(oDisplay("accountid")) Then 
				response.write "<td align=""left"">"   & oDisplay("accountname")   & "</td>"
				response.write "<td align=""center"">" & oDisplay("accountnumber") & "</td>"

				iOldAccountId = CLng(oDisplay("accountid"))
				dSubTotal     = CDbl(0.00)
			Else 
				'Need place holders 
				response.write "<td>&nbsp;</td>"
				response.write "<td>&nbsp;</td>"
			End If 

			response.write "<td>" & getSeasonName(oDisplay("classseasonid")) & "</td>"

			response.write "<td align=""right"">"
			dSubTotal = GetAccountSeasonTotal( oDisplay("accountid"), oDisplay("classseasonid"), sWhereClause )
			response.write FormatNumber(dSubTotal,2)
			dGrandTotal = dGrandTotal + dSubTotal
			response.write "</td>"
			response.write "</tr>"

			oDisplay.MoveNext
		Loop 

		'Totals Row
		response.write vbcrlf & "<tr class=""totalrow"">"
		response.write "<td colspan=""3"" align=""right"">Totals:</td>"
		response.write "<td align=""right"">" & FormatCurrency(dGrandTotal,2) & "</td>"
		response.write "</tr>"
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>"

		oDisplay.Close
		Set oDisplay = Nothing 

	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------------------------------------
' double dSubTotal = GetAccountSeasonTotal( iAccountId, iClassSeasonId, sWhereClause )
'------------------------------------------------------------------------------------------------------------
Function GetAccountSeasonTotal( ByVal iAccountId, ByVal iClassSeasonId, ByVal sWhereClause )
	Dim sSql, oRs 

	sSql = "SELECT SUM(L.amount) AS sub_total "
	sSql = sSql & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P, egov_class_list CL, "
	sSql = sSql & " egov_class_time T, egov_class C "
	sSql = sSql & " WHERE ispaymentaccount = 0 AND L.paymentid = P.paymentid "
	sSql = sSql & " AND A.accountid = L.accountid AND CL.classlistid = L.itemid "
	sSql = sSql & " AND CL.classtimeid = T.timeid AND L.entrytype = 'credit' "
	sSql = sSql & " AND C.classid = CL.classid AND L.accountid = " & iAccountId
	sSql = sSql & " AND C.classseasonid = " & iClassSeasonId & " " & sWhereClause
	sSql = sSql & " GROUP BY A.accountname, A.accountnumber, L.accountid, C.classseasonid"
	sSql = sSql & " ORDER BY A.accountname, A.accountnumber, L.accountid, C.classseasonid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetAccountSeasonTotal = CDbl(oRs("sub_total"))
	Else
		GetAccountSeasonTotal = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'------------------------------------------------------------------------------------------------------------
' ShowAdminLocations iLocationId 
'------------------------------------------------------------------------------------------------------------
Sub ShowAdminLocations( ByVal iLocationId )
	Dim sSql, oRs
	
	sSql = "SELECT locationid, name FROM egov_class_location "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " ORDER BY name"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""locationid"">"
		response.write vbcrlf & "<option value=""0"" "
		If CLng(0) = CLng(iLocationId) Then ' none selected
			 response.write " selected=""selected"" "
		End If 
		response.write ">Show All Locations</option>"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("locationid") & """"
			If CLng(oRs("locationid")) = CLng(iLocationId) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("name") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'------------------------------------------------------------------------------------------------------------
' ShowPaymentLocations iPaymentLocationId 
'------------------------------------------------------------------------------------------------------------
Sub ShowPaymentLocations( ByVal iPaymentLocationId )

	response.write vbcrlf & "<select name=""paymentlocationid"">"
	response.write vbcrlf & "<option value=""0"" "
	If CLng(0) = CLng(iPaymentLocationId) Then ' none selected
		 response.write " selected=""selected"" "
	End If 
	response.write ">Web Site and Office</option>"

	response.write vbcrlf & "<option value=""1"""
	If CLng(1) = CLng(iPaymentLocationId) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Office Only</option>"

	response.write vbcrlf & "<option value=""2"""
	If CLng(2) = CLng(iPaymentLocationId) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Web Site Only</option>"

	response.write vbcrlf & "</select>"

End Sub 


'------------------------------------------------------------------------------------------------------------
' ShowReportTypes iReportType 
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


'------------------------------------------------------------
' ShowClassSeasons iClassSeasonId 
'------------------------------------------------------------
Sub ShowClassSeasons( ByVal iClassSeasonId )
	Dim sSql, oRs

	response.write vbcrlf & "<select name=""ClassSeasonID"">"
	response.write vbcrlf & "<option value=""0"""
	If CLng(iClassSeasonID) = CLng(0) Then 
		response.write " selected=""selected"" "
	End if  
	response.write ">All</option>"

	sSql = "SELECT classseasonid, seasonname "
	sSql = sSql & " FROM egov_class_seasons "
	sSql = sSql & " WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY seasonyear, seasonname "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF

		response.write vbcrlf & "<option value=""" & oRs("classseasonid") & """"
		If CLng(iClassSeasonID) = CLng(oRs("classseasonid")) Then 
			response.write " selected=""selected"" "
		End if  
		response.write ">" & oRs("seasonname") & "</option>"

		oRs.MoveNext
	Loop 

	response.write vbcrlf & "</select>"
	
End Sub


'------------------------------------------------------------------------------------------------------------
'  ShowAdminUsers iAdminUserId 
'------------------------------------------------------------------------------------------------------------
Sub ShowAdminUsers( ByVal iAdminUserId )
	Dim sSql, oRs
	
	sSql = "SELECT userid, firstname, lastname FROM users "
	sSql = sSql & "WHERE isrootadmin = 0 AND orgid = " & session("orgid") & " ORDER BY lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""adminuserid"">"
		response.write vbcrlf & "<option value=""0"" "
		If CLng(0) = CLng(iAdminUserId) Then ' none selected
			 response.write " selected=""selected"" "
		End If 
		response.write ">Show All</option>"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("userid") & """"
			If CLng(oRs("userid")) = CLng(iAdminUserId) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("firstname") & " " & oRs("lastname") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------
' ShowReportTypes iReportType
'--------------------------------------------------------------------------------------
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


'---------------------------------------------------------
' string sName = getSeasonName( iClassSeasonId )
'---------------------------------------------------------
Function getSeasonName( ByVal iClassSeasonId )
	Dim sSql, oRs

	sSql = "SELECT seasonname FROM egov_class_seasons WHERE classseasonid = " & iClassSeasonId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		getSeasonName = oRs("seasonname")
	Else
		getSeasonName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 



%>
