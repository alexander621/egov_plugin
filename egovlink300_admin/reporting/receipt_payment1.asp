<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: receipt_payment.asp
' AUTHOR: SteveLoar
' CREATED: 01/10/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   7/17/07		Steve Loar - INITIAL VERSION
' 1.1	10/4/2007	Steve Loar - Adding payments to citizen accounts to the report
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

' INITIALIZE AND DECLARE VARIABLES
' SPECIFY FOLDER LEVEL
sLevel = "../" ' Override of value from common.asp


' USER SECURITY CHECK
If Not UserHasPermission( Session("UserId"), "receipt payment rpt" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 


' PROCESS REPORT FILTER VALUES
' PROCESS DATE VALUES
Dim iLocationId, iAdminUserId, iPaymentLocationId
fromDate = Request("fromDate")
toDate = Request("toDate")
today = Date()

' IF EMPTY DEFAULT TO CURRENT TO DATE
If toDate = "" or IsNull(toDate) Then
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

' BUILD SQL WHERE CLAUSE
varWhereClause = " WHERE (paymentDate >= '" & fromDate & "' AND paymentDate <= '" & DateAdd("d",1,toDate) & "') "
varWhereClause = varWhereClause & " AND orgid = " & session("orgid") 
If iLocationId > 0 Then
	varWhereClause = varWhereClause & " AND adminlocationid = " & iLocationId
End If 
If iAdminUserId > 0 Then
	varWhereClause = varWhereClause & " AND adminuserid = " & iAdminUserId
End If 
If iPaymentLocationId > 0 Then
	If iPaymentLocationId = CLng(2) Then
		varWhereClause = varWhereClause & " AND paymentlocationid = 3 " 
	Else
		varWhereClause = varWhereClause & " AND paymentlocationid < 3 " 
	End If 
End If 


%>

<html>
<head>
  <title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="reporting.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="pageprint.css" media="print" />

	<script language="Javascript" src="scripts/tablesort.js"></script>

	<script language="Javascript">
	  <!--
		function doCalendar(ToFrom) {
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  eval('window.open("../recreation/gr_calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

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
			//  controls[i].style.display = "inline";
		  //}
		}

		function validate()
		{
			// check from date
			if (document.frmPFilter.fromDate.value != "")
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.frmPFilter.fromDate.value);
				if (! Ok)
				{
					alert("The payment from date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.frmPFilter.fromDate.focus();
					return;
				}
			}
			// check to date
			if (document.frmPFilter.toDate.value != "")
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.frmPFilter.toDate.value);
				if (! Ok)
				{
					alert("The payment to date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.frmPFilter.toDate.focus();
					return;
				}
			}

			document.frmPFilter.submit();
		}

	  //-->
	</script>

	<script language="Javascript" src="scripts/dates.js"></script>

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

<form action="receipt_payment1.asp" method="post" name="frmPFilter">

	<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<tr>
			<td><font size="+1"><b>Receipt Payment Report</b></font></td>
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
								<input type="text" name="fromDate" value="<%=fromDate%>" size="10" maxlength="10" />
								<a href="javascript:void doCalendar('From');"><img src="../images/calendar.gif" border="0" /></a>		 
							</td>
							<td>&nbsp;</td>
							<td>
								<strong>To:</strong>
							</td>
							<td>
								<input type="text" name="toDate" value="<%=toDate%>" size="10" maxlength="10" />
								<a href="javascript:void doCalendar('To');"><img src="../images/calendar.gif" border="0" /></a>
							</td>
							<td>&nbsp;</td>
							<td><%DrawDateChoices "Dates" %></td>
						</tr>
					</table>
					</p>
					<p>
						<strong>Admin Location: </strong><% ShowAdminLocations iLocationId %>&nbsp;&nbsp;
						<strong>Admin: </strong><% ShowAdminUsers iAdminUserId %>&nbsp;&nbsp;
					</p>
					<p>
						<strong>Payment Location: </strong><% ShowPaymentLocations iPaymentLocationId %>
					</p>
					<!--END: DATE FILTERS-->
					<p>
						<input class="button" type="button" value="View Report" onclick="validate();" />
						&nbsp;&nbsp;<input type="button" class="button" value="Download to Excel" onClick="location.href='receipt_payment_export.asp?fromDate=<%=fromDate%>&toDate=<%=toDate%>&locationid=<%=iLocationId%>&adminuserid=<%=iAdminUserId%>&paymentlocationid=<%=iPaymentLocationId%>'" />
					</p>

				</fieldset>
				<!--END: FILTERS-->
		    </td>
		</tr>
		<tr>
 
			<td colspan="3" valign="top">
	  
				<!--BEGIN: DISPLAY RESULTS-->
				<%
				
				' DISPLAY RESULTS
				Display_Results varWhereClause
				
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
' FUNCTION DRAWDATECHOICES(SNAME)
'------------------------------------------------------------------------------------------------------------
Sub DrawDateChoices( sName )

	response.write vbcrlf & "<select onChange=""getDates(document.frmPFilter." & sName & ".value);"" class=""calendarinput"" name=""" & sName & """>"
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
' Sub Display_Results( sWhereClause )
'------------------------------------------------------------------------------------------------------------
Sub Display_Results( sWhereClause )
	Dim sSql, oRequests, oDisplay, iOldPaymentId, dCashTotal, dCheckTotal, dCardtotal, dOtherTotal, dMemoTotal
	Dim dGrandTotal, dCCCTotal, dCCCSubTotal, bHasData

	iOldPaymentId = CLng(0) 
	dCCCTotal = CDbl(0.0)
	bHasData = False 

	' make a holding recordset
	Set oDisplay = server.CreateObject("ADODB.RECORDSET") 
	oDisplay.fields.append "paymentid", adInteger, , adFldUpdatable
	oDisplay.fields.append "paymentdate", adVariant, 10, adFldUpdatable
	oDisplay.fields.append "item", adVarChar, 50, adFldUpdatable
	oDisplay.fields.append "userid", adInteger, , adFldUpdatable
	oDisplay.fields.append "userfname", adVarChar, 50, adFldUpdatable
	oDisplay.fields.append "userlname", adVarChar, 50, adFldUpdatable
	oDisplay.fields.append "userhomephone", adVarChar, 50, adFldUpdatable
	oDisplay.fields.append "checkamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "checkno", adVarChar, 20, adFldUpdatable
	oDisplay.fields.append "cashamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "cardamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "cccsubtotal", adCurrency, , adFldUpdatable
	oDisplay.fields.append "otheramt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "memoamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "paymenttotal", adCurrency, , adFldUpdatable

	oDisplay.CursorLocation = 3
	'oDisplay.CursorType = 3
	oDisplay.open 

	' Pull Class Purchases
	If AddToDisplay( "egov_class_to_payment_method", sWhereClause, oDisplay ) Then 
		bHasData = True
	End If 


	' Pull Citizen Account Deposits 
	If AddToDisplay(  "egov_citizen_account_to_payment_method", sWhereClause, oDisplay  ) Then 
		bHasData = True
	End If 

	If bHasData Then 
		' Sort the data by paymentid
		oDisplay.sort = "paymentid"
		' Show results
		oDisplay.MoveFirst
		response.write vbcrlf & "<div class=""receiptpaymentshadow"">"
		response.Write vbcrlf & "<table cellspacing=""0"" cellpadding=""2""  border=""0"" width=""100%"" class=""receiptpayment"">"
		response.write vbcrlf & "<tr class=""tablelist""><th>Receipt</th><th>Date</th><th>Payee</th>"
		response.write "<th>Check Amt<br />Check #</th><th>Cash Amt</th><th>Card Amt</th><th>Total Chck<br />Cash, Card</th><th>Other Amt</th>"
		response.write "<th>Memo Amt</th><th>Total<br />Paid</th></tr>"
		bgcolor = "#eeeeee"
		Do While Not oDisplay.EOF
			If bgcolor="#eeeeee" Then
				bgcolor="#ffffff" 
			Else
				bgcolor="#eeeeee"
			End If			
			response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>"
			If oDisplay("item") = "Citizen Acct" Then
				response.write "<td align=""center""><a href=""../purchases/viewjournal.asp?uid=" & oDisplay("userid") & "&pid=" & oDisplay("paymentid") & "&rt=c&it=ci&jet=d"">" & oDisplay("paymentid") & "</a></td>"
			Else 
				response.write "<td align=""center""><a href=""../classes/view_receipt.asp?iPaymentId=" & oDisplay("paymentid") & """>" & oDisplay("paymentid") & "</a></td>"
			End If 
			response.write "<td align=""center"">" & oDisplay("paymentdate") & "</td>"
			response.write "<td align=""center"" valign=""top"">" & oDisplay("userfname") & " " & oDisplay("userlname") & "<br />" & FormatPhoneNumber(oDisplay("userhomephone")) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("checkamt"), 2) & "<br />" & oDisplay("checkno") & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("cashamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("cardamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("cccsubtotal"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("otheramt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("memoamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("paymenttotal"),2) & "</td>"
			dCheckTotal = dCheckTotal + CDbl(oDisplay("checkamt"))
			dCashTotal = dCashTotal + CDbl(oDisplay("cashamt"))
			dCardTotal = dCardTotal + CDbl(oDisplay("cardamt"))
			dOtherTotal = dOtherTotal + CDbl(oDisplay("otheramt"))
			dMemoTotal = dMemoTotal + CDbl(oDisplay("memoamt"))
			dGrandTotal = dGrandTotal + CDbl(oDisplay("paymenttotal"))
			dCCCTotal = dCCCTotal + CDbl(oDisplay("cccsubtotal"))
			response.write "</tr>"
			oDisplay.MoveNext
		Loop 
		' Totals Row
		If bgcolor="#eeeeee" Then
			bgcolor="#ffffff" 
		Else
			bgcolor="#eeeeee"
		End If	
		response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """ class=""totalrow""><td colspan=""3"" align=""right"">Totals:</td>"
		response.write "<td align=""right"">" & FormatNumber(dCheckTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCashTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCardTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCCCTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dOtherTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dMemoTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dGrandTotal,2) & "</td>"
		response.write "</tr>"

		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>"

	End If 

	oDisplay.Close
	Set oDisplay = Nothing 
	
End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub ShowAdminLocations( iLocationId )
'------------------------------------------------------------------------------------------------------------
Sub ShowAdminLocations( iLocationId )
	Dim sSql, oLocation
	
	sSql = "Select locationid, name from egov_class_location where orgid = " & session("orgid") & " order by name"

	Set oLocation = Server.CreateObject("ADODB.Recordset")
	oLocation.Open  sSQL, Application("DSN"), 3, 1

	If Not oLocation.EOF Then 
		response.write vbcrlf & "<select name=""locationid"">"
		response.write vbcrlf & "<option value=""0"" "
		If CLng(0) = CLng(iLocationId) Then ' none selected
			 response.write " selected=""selected"" "
		End If 
		response.write ">Show All Locations</option>"
		Do While Not oLocation.EOF 
			response.write vbcrlf & "<option value=""" & oLocation("locationid") & """"
			If CLng(oLocation("locationid")) = CLng(iLocationId) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oLocation("name") & "</option>"
			oLocation.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oLocation.close
	Set oLocation = Nothing 
End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub ShowAdminUsers( iAdminUserId )
'------------------------------------------------------------------------------------------------------------
Sub ShowAdminUsers( iAdminUserId )
	Dim sSql, oUser
	
	sSql = "Select userid, firstname, lastname from users where isrootadmin = 0 and orgid = " & session("orgid") & " order by lastname, firstname"

	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open  sSQL, Application("DSN"), 3, 1

	If Not oUser.EOF Then 
		response.write vbcrlf & "<select name=""adminuserid"">"
		response.write vbcrlf & "<option value=""0"" "
		If CLng(0) = CLng(iAdminUserId) Then ' none selected
			 response.write " selected=""selected"" "
		End If 
		response.write ">Show All</option>"
		Do While Not oUser.EOF 
			response.write vbcrlf & "<option value=""" & oUser("userid") & """"
			If CLng(oUser("userid")) = CLng(iAdminUserId) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oUser("firstname") & " " & oUser("lastname") & "</option>"
			oUser.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oUser.close
	Set oUser = Nothing 
End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub ShowPaymentLocations( iPaymentLocationId )
'------------------------------------------------------------------------------------------------------------
Sub ShowPaymentLocations( iPaymentLocationId )

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
' Function AddToDisplay( sSql )
'------------------------------------------------------------------------------------------------------------
Function AddToDisplay( ByVal sFrom, ByVal sWhereClause, ByRef oDisplay )
	Dim oRequests, bHasData, sSql

	sSql = "SELECT paymentid, orgid, userid, ISNULL(userfname,'') as userfname, ISNULL(userlname,'') AS userlname, ISNULL(userhomephone,'') AS userhomephone, paymenttotal, paymentdate, journalentrytype, amount, paymenttypename, checkno, isothermethod, "
    sSql = sSql & " requirescash, requirescreditcard, requirescitizenaccount, requirescheckno, paymentlocationname, adminlocationid, adminuserid, item, [Transaction ID] "
	sSql = sSql & " FROM " & sFrom & " " & sWhereClause
	sSql = sSql & " ORDER BY paymentid" 
	session("AddToDisplaySql") = sSql

'dtb_debug(sSql)

	'response.write sSql & "<br /><br />"

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSQL, Application("DSN"), 3, 1
	

	If Not oRequests.EOF then
		bHasData = True 

		' Loop through and build the display recordset.
		Do While Not oRequests.EOF
			If CLng(oRequests("paymentid")) <> iOldPaymentId Then
				oDisplay.addnew 
				oDisplay("paymentid") = oRequests("paymentid")
				oDisplay("paymentdate") = DateValue(oRequests("paymentdate"))
				oDisplay("item") = oRequests("item")
				oDisplay("userid") = oRequests("userid")
				oDisplay("userfname") = oRequests("userfname")
				oDisplay("userlname") = oRequests("userlname")
				oDisplay("userhomephone") = oRequests("userhomephone")
				oDisplay("paymenttotal") = oRequests("paymenttotal")
				oDisplay("checkamt") = 0.00
				oDisplay("cashamt") = 0.00
				oDisplay("cardamt") = 0.00
				oDisplay("cccsubtotal") = 0.00
				oDisplay("otheramt") = 0.00
				oDisplay("memoamt") = 0.00
				dCCCSubTotal = 0.00
				iOldPaymentId = CLng(oRequests("paymentid"))
			End If 
			If oRequests("requirescheckno") Then
				oDisplay("checkamt") = oRequests("amount")
				oDisplay("checkno") = oRequests("checkno")
				dCCCSubTotal = dCCCSubTotal + CDbl(oRequests("amount"))
			End If 
			If oRequests("requirescash") Then
				oDisplay("cashamt") = oRequests("amount")
				dCCCSubTotal = dCCCSubTotal + CDbl(oRequests("amount"))
			End If 
			If oRequests("requirescreditcard") Then
				oDisplay("cardamt") = oRequests("amount")
				dCCCSubTotal = dCCCSubTotal + CDbl(oRequests("amount"))
			End If 
			If oRequests("isothermethod") Then
				oDisplay("otheramt") = oRequests("amount")
			End If 
			If oRequests("requirescitizenaccount") Then
				oDisplay("memoamt") = oRequests("amount")
			End If 
			oDisplay("cccsubtotal") = dCCCSubTotal

			oDisplay.Update
			oRequests.MoveNext
		Loop
	Else
		bHasData = False
	End If 
	
	session("AddToDisplaySql") = ""
	oRequests.Close
	Set oRequests = Nothing

	AddToDisplay = bHasData
End Function 

sub dtb_debug(p_value)
sSQLi = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
	Set rsi = Server.CreateObject("ADODB.Recordset")
	rsi.Open sSQLi, Application("DSN"), 3, 1
end sub
%>