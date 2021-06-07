<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: Gift_payment_list.asp
' AUTHOR: JOHN STULLENBERGER
' CREATED: 07/31/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Gift purchases list
'
' MODIFICATION HISTORY
' 1.0   07/31/2006	JOHN STULLENBERGER - INITIAL VERSION
' 1.1   09/08/2006	Steve Loar - Added the ability to click the row and jump to the details page
' 2.0	07/30/2010	Steve Loar - Changed made for Point and Pay payments
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "gift report" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

orderBy = Request("orderBy")
subTotals = Request("subTotals")
showDetail = Request("showDetail")
fromDate = Request("fromDate")
toDate = Request("toDate")
today = Date()

If orderBy = "" or IsNull(orderBy) Then 
	orderBy = "date" 
End If

If toDate = "" Or IsNull(toDate) Then 
	toDate = dateAdd("d",1,Date()) 
End If

'If fromDate = "" or IsNull(fromDate) Then fromDate = dateAdd("ww",-1,today) End If
If fromDate = "" Or IsNull(fromDate) Then 
	fromDate = DateSerial(Year(Now()),1,1) 
End If

'toDate = dateAdd("d",1,toDate)
%>

<html>
<head>
	<title><%=langBSPayments%></title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="giftstyles.css" />

	<script language="javascript" src="../scripts/selectAll.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="JavaScript">
	<!--

		function checkStat() 
		{
			if ( !(form1.statusInProgress.checked) &&  !(form1.statusPending.checked) && !(form1.statusRefund.checked) && !(form1.statusDenied.checked) &&  !(form1.statusCompleted.checked) && !(form1.statusProcessed.checked)) 
			{
				alert("You must select the status.");
				form1.statusPending.focus();
				return false;
			}
		}

		function CheckAllStatus() 
		{
			if (document.form1.CheckAllStat.checked) 
			{
				document.form1.statusPending.checked = true;
				document.form1.statusCompleted.checked = true;
				document.form1.statusDenied.checked = true;
			} 
			else 
			{
				document.form1.statusPending.checked = false;
				document.form1.statusCompleted.checked = false;
				document.form1.statusDenied.checked = false;
			}
		}

		function doCalendar(sField) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=GiftForm", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function insertAtURL (textEl, text) 
		{
			if (textEl.createTextRange && textEl.caretPos) 
			{
				var caretPos = textEl.caretPos;
				caretPos.text =
				caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
				text + ' ' : text;
			}
			else
				textEl.value  = text;
		}

	//-->
	</script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="content">
	<div id="centercontent">

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
<!--    <tr>
	    <td><font size="+1"><b>(E-Gov Payment Receipt Manager) - Manage Online Submitted Gift Payments</b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.back()"><%=langBackToStart%></a></td>
    </tr>
-->
	<tr>
    <td>
		<!--BEGIN: SEARCH OPTIONS-->
		<fieldset>
			<legend><b>Search/Sorting Option(s)</b></legend>
			<form action="gift_payment_list.asp" method="post" name="GiftForm">
			<table border=0>
				<tr>
				<td valign=top>
					<strong>From: </strong>
					<input type=text name="fromDate" value="<%=fromDate%>" />
					<a href="javascript:void doCalendar('fromDate');"><img src="../images/calendar.gif" alt="" border="0"></a>		 
				</td>
				<td>&nbsp;</td>
				<td valign=top>
					<strong>To:</strong> 
					<input type=text name="toDate" value="<%=toDate%>" />
					<a href="javascript:void doCalendar('toDate');"><img src="../images/calendar.gif" alt="" border="0"></a>
				</td>
				</tr>
			</table>
			<input type="submit" value="View" class="button">
			</form>
		</fieldset>
		<!--END: SEARCH OPTIONS-->
    </td>
  </tr>
	<tr>
      <td colspan="3" valign="top">
	  <!--BEGIN: ACTION LINE REQUEST LIST -->
      
<% 
		List_Payments request("sort") 
%>
	  
	  <!-- END: ACTION LINE REQUEST LIST -->
      </td>
    </tr>
  </table>

  </div>
</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'------------------------------------------------------------------------------------------------------------
' FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' void List_Payments( sSortBy )
'------------------------------------------------------------------------------------------------------------
Sub List_Payments( ByVal sSortBy)
	Dim sSql, oRs, iRowCount

	varWhereClause = " AND (paymentDate >= '" & fromDate & "' AND paymentDate < '" & toDate & "') "

	If sSortBy = "" Then
		sSqlSUM = "SELECT SUM(giftamount) AS total FROM egov_gift_payment INNER JOIN egov_gift ON egov_gift_payment.giftid = egov_gift.giftID WHERE egov_gift_payment.paymenttype IS NOT NULL AND egov_gift.orgid = " & Session("orgid") & varWhereClause
		sSql = "SELECT *, (" & sSqlSUM  & ") AS total FROM egov_gift_payment INNER JOIN egov_gift ON egov_gift_payment.giftid = egov_gift.giftID WHERE egov_gift_payment.paymenttype IS NOT NULL AND egov_gift.orgid = " & Session("orgid") & varWhereClause
	Else
		sSqlSUM = "SELECT SUM(giftamount) AS total FROM egov_gift_payment INNER JOIN egov_gift ON egov_gift_payment.giftid = egov_gift.giftID WHERE egov_gift_payment.paymenttype IS NOT NULL AND egov_gift.orgid = " & Session("orgid") & varWhereClause & " ORDER BY " & sSortBy
		sSql = "SELECT *,(" & sSqlSUM  & ") AS total FROM egov_gift_payment INNER JOIN egov_gift ON egov_gift_payment.giftid = egov_gift.giftID WHERE egov_gift_payment.paymenttype IS NOT NULL AND egov_gift.orgid = " & Session("orgid") & varWhereClause & " ORDER BY " & sSortBy
	End If

	'response.write sSql & "<br /><br />"

	lastTitle = "Test"
	lastDate = "1/1/02"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.CursorLocation = 3
	oRs.Open sSql, Application("DSN"), 3, 1

	If oRs.EOF Then 
		Response.write "<p><b>No records found</p>"
	Else 
		curGrandTotal = oRs("total")

		' DISPLAY RECORD STATISTICS
		Dim abspage, pagecnt
		abspage = oRs.AbsolutePage
		pagecnt = oRs.PageCount

		Response.Write " <strong><font color=""blue"">" & oRs.RecordCount & "</font> Total Online Payments</strong><br /><br />"

		response.write "<div class=""shadow"">"
		Response.Write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" width=""100%"">"
		Response.Write "<tr align=""left"" class=""tablelist""><th>&nbsp;</th>"
		response.write "<th align=""left"">Payment Reference</td>"
		response.write "<th>Product</th><th align=""center"">Transaction Date</th><th align=""center"">Status</th><th align=""center"">Amount</th></tr>"

		' LOOP AND DISPLAY THE RECORDS
		oRs.MoveFirst
		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr id=""" & iRowCount & """"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
			response.write "<td onClick=""location.href='../purchases_report/gift_details.asp?igiftpaymentid=" & oRs("giftpaymentid") & "';"">&nbsp;</td>"
			response.write "<td align=""left"" onClick=""location.href='../purchases_report/gift_details.asp?igiftpaymentid=" & oRs("giftpaymentid") & "';"">"
			If IsNull(oRs("PNREf")) Then
				response.write oRs("ordernumber")
			Else
				response.write oRs("PNREf")
			End If 
			response.write "</td>"
			response.write "<td onClick=""location.href='../purchases_report/gift_details.asp?igiftpaymentid=" & oRs("giftpaymentid") & "';"">" & oRs("giftname") & "</td>"
			response.write "<td align=""center"" onClick=""location.href='../purchases_report/gift_details.asp?igiftpaymentid=" & oRs("giftpaymentid") & "';"">" & DateValue(oRs("paymentdate")) & "</td>"
			response.write "<td align=""center"" onClick=""location.href='../purchases_report/gift_details.asp?igiftpaymentid=" & oRs("giftpaymentid") & "';"">" & oRs("result") & "</td>"
			response.write "<td class=""giftamountcol"" align=""right"" onClick=""location.href='../purchases_report/gift_details.asp?igiftpaymentid=" & oRs("giftpaymentid") & "';"">" & FormatCurrency(oRs("giftamount"),2) & "</td></tr>"
			oRs.MoveNext 
		Loop

		' DISPLAY TOTAL ROW
		Response.Write "<tr class=""tablelist"" style=""background-color:#e0e0e0;""><td>&nbsp;</td><td align=""left""></td><td></td><td></td><td></td><td class=""giftamountcol"" align=""right""><strong>Grand Total: " & FormatCurrency(curGrandTotal,2)  & "</strong></td></tr>"
		response.write "</table>"
		response.write "</div>"
	End If

	oRs.Close
	Set oRs = Nothing 
 
End Sub 


'------------------------------------------------------------------------------------------------------------
' Function DisplayGateWayInformation( sText )
'------------------------------------------------------------------------------------------------------------
'Function DisplayGateWayInformation( ByVal sText)
'	
'	' USED TO STORE DICTIONARY DATA
'	Set oDictionary=Server.CreateObject("Scripting.Dictionary")
'
'	' MAKE SURE THERE IS INFORMATION TO PARSE
'	If sText <> "" Then
'	
'		' BREAK LIST INTO SEPARATE LINES
'		arrInfo = SPLIT(sTEXT,"<br>")
'
'		' BREAK LINES INTO FIELD NAME AND VALUE
'		For w = 0 to UBOUND(arrInfo)
'			arrNamedPair = SPLIT(arrInfo(w),":")
'			
'			' MATCHED SETS ARE ADDED TO DICTIONARY
'			If UBOUND(arrNamedPair) > 0 Then
'				oDictionary.Add TRIM(arrNamedPair(0)),Trim(arrNamedPair(1))
'			End If 
'		Next
'	
'	End If
'
'	' BUILD PERSONAL INFO DISPLAY
'	response.write "<B><U><FONT style=""font-size:10px;"" COLOR=BLACK>BILL TO: </font></U></b><BR>"
'	response.write UCASE(oDictionary.Item("Cardmember Name")) & "<BR>"
'	response.write UCASE(oDictionary.Item("Street Address")) & "<BR>"
'	response.write UCASE(oDictionary.Item("City")) & ", " & UCASE(oDictionary.Item("State")) & ", " & oDictionary.Item("Zipcode")
'	Set oDictionary = Nothing
'
'End Function


'------------------------------------------------------------------------------------------------------------
' Function DisplayGateWayTransactionInformation(sText)
'------------------------------------------------------------------------------------------------------------
'Function DisplayGateWayTransactionInformation( ByVal sText)
'	
'	' USED TO STORE DICTIONARY DATA
'	Set oDictionary=Server.CreateObject("Scripting.Dictionary")
'
'	' MAKE SURE THERE IS INFORMATION TO PARSE
'	If sText <> "" Then
'	
'		' BREAK LIST INTO SEPARATE LINES
'		arrInfo = SPLIT(sTEXT,"<br>")
'
'		' BREAK LINES INTO FIELD NAME AND VALUE
'		For w = 0 to UBOUND(arrInfo)
'			arrNamedPair = SPLIT(arrInfo(w),":")
'			
'			' MATCHED SETS ARE ADDED TO DICTIONARY
'			If UBOUND(arrNamedPair) > 0 Then
'				oDictionary.Add TRIM(arrNamedPair(0)),Trim(arrNamedPair(1))
'			End If 
'		Next
'
'	End If
'
'	' BUILD PERSONAL INFO DISPLAY
'	response.write "<B><FONT style=""font-size:10px;"" COLOR=BLACK>EGOV ORDER ID: </font></b><FONT style=""font-size:10px;"">" & UCASE(oDictionary.Item("Order Number")) & "</font><BR>"
'	response.write "<B><FONT style=""font-size:10px;"" COLOR=BLACK>PAYMENT ID: </font></b><FONT style=""font-size:10px;"">" & UCASE(oDictionary.Item("Transaction File Name")) & "</font><BR>"
'
'	Set oDictionary = Nothing
'
'End Function


'------------------------------------------------------------------------------------------------------------
' Function DisplayGateWayInformationPayPal(sText)
'------------------------------------------------------------------------------------------------------------
'Function DisplayGateWayTransactionInformationPayPal(sText)
'	
'
'	' USED TO STORE DICTIONARY DATA
'	Set oDictionary=Server.CreateObject("Scripting.Dictionary")
'
'	' MAKE SURE THERE IS INFORMATION TO PARSE
'	If sText <> "" Then
'	
'		' BREAK LIST INTO SEPARATE LINES
'		arrInfo = SPLIT(sText, "</br>")
'
'		' BREAK LINES INTO FIELD NAME AND VALUE
'		For w = 0 to UBOUND(arrInfo)
'			
'			arrNamedPair = SPLIT(arrInfo(w),":")
'
'			' MATCHED SETS ARE ADDED TO DICTIONARY
'			If UBOUND(arrNamedPair) > 0 Then
'				oDictionary.Add TRIM(arrNamedPair(0)),Trim(arrNamedPair(1))
'			End If 
'		Next
'
'	End If
'
'	' BUILD PERSONAL INFO DISPLAY
'	'response.write "<B><FONT COLOR=BLACK>EGOV ORDER ID: </font></b>" & UCASE(oDictionary.Item("Order Number")) & "<BR>"
'	response.write "<B><FONT style=""font-size:10px;"" COLOR=BLACK>PAL PAY REFERENCE ID: </font></b><FONT style=""font-size:10px;"">" & UCASE(oDictionary.Item("txn_id")) & "</font><BR>"
'
'	Set oDictionary = Nothing

'End Function


'------------------------------------------------------------------------------------------------------------
' Function DisplayGateWayInformationPayPal(sText)
'------------------------------------------------------------------------------------------------------------
'Function DisplayGateWayInformationPayPal(sText)
'	
'	' USED TO STORE DICTIONARY DATA
'	Set oDictionary=Server.CreateObject("Scripting.Dictionary")
'
'	' MAKE SURE THERE IS INFORMATION TO PARSE
'	If sText <> "" Then
'	
'		' BREAK LIST INTO SEPARATE LINES
'		arrInfo = SPLIT(sText, "</br>")
'
'		' BREAK LINES INTO FIELD NAME AND VALUE
'		For w = 0 to UBOUND(arrInfo)
'			
'			arrNamedPair = SPLIT(arrInfo(w),":")
'
'			' MATCHED SETS ARE ADDED TO DICTIONARY
'			If UBOUND(arrNamedPair) > 0 Then
'				oDictionary.Add TRIM(arrNamedPair(0)),Trim(arrNamedPair(1))
'			End If 
'		Next
'
'	End If
'
'	' BUILD PERSONAL INFO DISPLAY
'	response.write "<B><U><FONT style=""font-size:10px;"" COLOR=BLACK>BILL TO: </font></U></b><BR><FONT style=""font-size:10px;"">"
'	response.write UCASE(oDictionary.Item("address_name")) & "<BR>"
'	response.write UCASE(oDictionary.Item("address_street")) & "<BR>"
'	response.write UCASE(oDictionary.Item("address_city")) & ", " & UCASE(oDictionary.Item("address_state")) & ", " & oDictionary.Item("address_zip") & "</font>"
'	Set oDictionary = Nothing

'End Function


'------------------------------------------------------------------------------------------------------------
' Function DisplayGateWayInformationVerisign(sText)
'------------------------------------------------------------------------------------------------------------
'Function DisplayGateWayTransactionInformationVerisign(iID,iRef)
'	
'
'	' BUILD PERSONAL INFO DISPLAY
'	response.write "<B><FONT COLOR=BLACK>EGOV ORDER ID: </font></b>ecC" & session("orgid") & "0" & iID & "<BR>"
'	response.write "<B><FONT style=""font-size:10px;"" COLOR=BLACK>VERISIGN REFERENCE ID: </font></b><FONT style=""font-size:10px;"">" & iRef & "</font><BR>"
'
'	Set oDictionary = Nothing
'
'End Function
'

'------------------------------------------------------------------------------------------------------------
' Function DisplayGateWayInformationVerisign(sText)
'------------------------------------------------------------------------------------------------------------
'Function DisplayGateWayInformationVerisign(sName,sAddress,sCity,sState,sZip)
'	
'	' BUILD PERSONAL INFO DISPLAY
'	response.write "<B><U><FONT style=""font-size:10px;"" COLOR=BLACK>BILL TO: </font></U></b><BR><FONT style=""font-size:10px;"">"
'	response.write UCASE(sName) & "<BR>"
'	response.write UCASE(sAddress) & "<BR>"
'	response.write UCASE(sCity) & ", " & UCASE(sState) & ", " & sZip & "</font>"
'	Set oDictionary = Nothing
'
'End Function


'------------------------------------------------------------------------------------------------------------
' Function GetGrandTotal()
'------------------------------------------------------------------------------------------------------------
'Function GetGrandTotal()
'
'	response.write "<table>"
'	response.write "<tr><td></td></tr>"
'	response.write "</table>"
'
'End Function


'------------------------------------------------------------------------------------------------------------
' Function Encode(sIn)
'------------------------------------------------------------------------------------------------------------
'Function Encode(sIn)
'    dim x, y, abfrom, abto
'    Encode="": ABFrom = ""
'
 '   For x = 0 To 25: ABFrom = ABFrom & Chr(65 + x): Next 
'    For x = 0 To 25: ABFrom = ABFrom & Chr(97 + x): Next 
'    For x = 0 To 9: ABFrom = ABFrom & CStr(x): Next 
'
'    abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
'    For x=1 to Len(sin): y = InStr(abfrom, Mid(sin, x, 1))
'        If y = 0 Then
'             Encode = Encode & Mid(sin, x, 1)
'        Else
'             Encode = Encode & Mid(abto, y, 1)
'        End If
'    Next
'End Function 


%>