<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: action_line_list.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the payments list
'
' MODIFICATION HISTORY
' 1.0   ???			???? - INITIAL VERSION
' 1.1	10/12/06	Steve Loar - Security, Header and nav changed
' 1.2	04/17/2009	Steve Loar - Adding Payment Service Picks and new excel export
' 1.3	09/27/2010	Steve Loar - Modifications for Rye NY
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPaymentServiceId, sPaymentService

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "receipts" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

orderBy = Request("orderBy")
subTotals = Request("subTotals")
showDetail = Request("showDetail")
fromDate = Request("fromDate")
toDate = Request("toDate")
today = Date()

If orderBy = "" or IsNull(orderBy) Then orderBy = "date" End If
If toDate = "" or IsNull(toDate) Then toDate = dateAdd("d",0,today) End If
If fromDate = "" or IsNull(fromDate) Then fromDate = dateAdd("ww",-1,today) End If

toDate = dateAdd("d",1,toDate)


If subTotals = "yes" Then 
	subTotals = "yes"
ElseIf subTotals = "" AND Request.ServerVariables("REQUEST_METHOD") <> "POST" Then 
	subTotals = "yes"
ELSE
	subTotals = "no"
End If

If showDetail = "yes" Then 
	showDetail = "yes"
ElseIf showDetail = "" AND Request.ServerVariables("REQUEST_METHOD") <> "POST" Then 
	showDetail = "yes"
ELSE
	showDetail = "no"
End If

statusInProgress = Request("statusInProgress")
statusPending = Request("statusPending")
statusRefund = Request("statusRefund")
statusDenied = Request("statusDenied")
statusCompleted = Request("statusCompleted")
statusProcessed = Request("statusProcessed")

noStatus = true

If statusInProgress = "yes" Then 
	noStatus = false
ELSE
	statusInProgress = "no"
End If
If statusPending = "yes"  Then 
	noStatus = false 
ELSE
	statusPending = "no"
End If
If statusRefund = "yes"  Then 
	noStatus = false 
ELSE
	statusRefund = "no"
End If
If statusDenied = "yes" Then 
	noStatus = false
ELSE
	statusDenied = "no"
End If
If statusCompleted = "yes" Then 
	noStatus = false
ELSE
	statusCompleted = "no"
End If
If statusProcessed = "yes" Then 
	noStatus = false
ELSE
	statusProcessed = "no"
End If

If noStatus = True Then 
   statusInProgress = "yes"
   statusPending = "yes"
   statusRefund = "yes"
   statusDenied = "yes"
   statusCompleted = "yes"
   statusProcessed = "yes"
End If

If request("paymentserviceid") <> "" Then
	'iPaymentServiceId = CLng(request("paymentserviceid"))
	iPaymentServiceId = request("paymentserviceid")
        sPaymentServiceId = ""
	if not isnumeric(iPaymentServiceId) then
		sPaymentServiceId = iPaymentServiceId
		iPaymentServiceId = 382
	end if
	sPaymentService = GetPaymentService( iPaymentServiceId )
Else
	iPaymentServiceId = 0
End If 


%>

<html>
<head>
	<meta http-equiv="Content-type" content="text/html;charset=UTF-8"> 
	<title><%=langBSPayments%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script type="text/javascript" src="../scripts/jquery-1.4.2.min.js"></script>

	<script type="text/javascript" src="../scripts/selectAll.js"></script>
	<script type="text/javascript" src="../scripts/isvaliddate.js"></script>
	<script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script LANGUAGE="JavaScript">
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

		function doCalendar( sField ) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			var sSelectedDate = $("#"+sField).val();
			eval('window.open("calendarpicker.asp?date=' + sSelectedDate + '&p=1&updatefield=' + sField + '&updateform=form1", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function displayResults()
		{
			// check the from date
			if ($("#fromDate").val() == "")
			{
				$("#fromDate").focus();
				inlineMsg("fromDate",'<strong>Invalid Value: </strong>Please enter a From date.',8,"fromDate");
				return;
			}
			else
			{
				if (! isValidDate($("#fromDate").val()))
				{
					$("#fromDate").focus();
					inlineMsg("fromDate",'<strong>Invalid Value: </strong>The From date should be a valid date in the format of MM/DD/YYYY.',8,"fromDate");
					return;
				}
			}

			// check the to date
			if ($("#toDate").val() == "")
			{
				$("#toDate").focus();
				inlineMsg("toDate",'<strong>Invalid Value: </strong>Please enter a To date.',8,"toDate");
				return;
			}
			else
			{
				if (! isValidDate($("#toDate").val()))
				{
					$("#toDate").focus();
					inlineMsg("toDate",'<strong>Invalid Value: </strong>The To date should be a valid date in the format of MM/DD/YYYY.',8,"toDate");
					return;
				}
			}

			document.form1.submit();
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
    <tr>
	    <td><font size="+1"><strong>(E-Gov Payment Receipt Manager) - Manage Online Submitted Payments</strong></font>
		</td>
    </tr>
	<tr>
    <td>
				 <!--BEGIN: SEARCH OPTIONS-->
		<fieldset>
			<legend><b>Search/Sorting Option(s)</b></legend>
			<form action="action_line_list.asp" method="post" name="form1"  onSubmit="return checkStat()">
				<table border="0" cellpadding="0" cellspacing="0" id="paymentsortoptions">
				  <tr>
				  <td valign="top">
					  <b>From: 
					  <input type="text" id="fromDate" name="fromDate" value="<%=fromDate%>" />
					  <a href="javascript:void doCalendar('fromDate');"><img src="../images/calendar.gif" border="0" /></a>		 
				  </td>
				  <td>&nbsp;</td>
				   <td valign="top">
					<b>To:</b> 
					  <input type="text" id="toDate" name="toDate" value="<%=dateAdd("d",-1,toDate)%>" />
					  <a href="javascript:void doCalendar('toDate');"><img src="../images/calendar.gif" border="0" /></a>
				   </td>
				  </tr>
				  <tr>
					<td valign="top" colspan="3">
					<%
					If statusInProgress = "yes" then check1 = "checked=""checked"""
					If statusPending = "yes" Then check2 = "checked=""checked"""
					If statusRefund = "yes" Then check3 = "checked=""checked"""
					If statusDenied = "yes" then check4 = "checked=""checked"""
					If statusCompleted = "yes" Then check5 = "checked=""checked""" 
					If statusProcessed = "yes" Then check6 = "checked=""checked"""
					%>
					
					<b>Display:</b> 
					 <input type="checkbox" name="statusInProgress" value="yes" <%=check1%> />In Progress
					 <input type="checkbox" name="statusPending" value="yes" <%=check2%> />Pending
					 <input type="checkbox" name="statusRefund" value="yes" <%=check3%> />Refund
					 <input type="checkbox" name="statusDenied" value="yes" <%=check4%> />Denied
					 <input type="checkbox" name="statusCompleted" value="yes" <%=check5%> />Completed
					 <input type="checkbox" name="statusProcessed" value="yes" <%=check6%> />Processed<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

					 
					 <!--<input type=checkbox name="CheckAllStat" value="checked" onClick="CheckAllStatus();">Check All-->

					  <% if subTotals = "yes" then %>
							  <input type="checkbox" name="subTotals" value="yes" checked="checked" />Subtotals
					  <% else %>
							<input type="checkbox" name="subTotals" value="yes" />Subtotals
						<% end if %>

				  <% if showDetail = "yes" then %>
							  <input type="checkbox" name="showDetail" value="yes" checked="checked" />Details
					  <% else %>
							<input type="checkbox" name="showDetail" value="yes" />Details
						<% end if %>
				   </td>
				</tr>	
				<tr>
					<td colspan="3" style="height: 30px;">
						<b>Payment Service: </b>
<%							ShowPaymentServicePicks iPaymentServiceId		%>
					</td>
				</tr>
				<tr>
					<td colspan="3" style="height: 30px;">
						<b>Order By: </b>
						<select name="orderBy">
							<option value="service">Payment Service</option>
							<option value="date" <% if orderBy="date" then response.write " selected"%>>Transaction Date</option>		 
							<option value="purchase" <% if orderBy="purchase" then response.write " selected"%>>Purchase Order</option>
						</select>
					</td>
				</tr>
				<tr>
					<td valign="bottom" colspan="3" style="height: 30px;">
						<input type="button" class="button" value="Display Results" onClick="displayResults();" /> &nbsp;
<%						If CLng(iPaymentServiceId) > CLng(0) Then	%>
							<input type="button" class="button" value="Download to Excel"  onClick="location.href='paymentsexport.asp'" />&nbsp;
<%							If sPaymentService = "rye commuter permit renewal" Or sPaymentService = "rye commuter permits by name" Or sPaymentService = "rye commuter permits by qty" or sPaymentService = "rye snow field parking by qty" or sPaymentService = "rye railroad new waitlist" Then	%>
								<input type="button" class="button" value="Simple Export" onClick="location.href='simpleexport.asp'" />
<%							End If		
							If sPaymentService = "rye commuter waitlist renewal" Then	%>
								<input type="button" class="button" value="Simple Export" onClick="location.href='waitlistexport.asp'" />
<%							End If		
						End If					
%>
					</td>
				</tr>
			</table>
		</form>
	</fieldset>

	<!--END: SEARCH OPTIONS-->

    </td>
  </tr>
	<tr>
 
      <td colspan="3" valign="top">
	  <!--BEGIN: Payments LIST -->
      
<% 
		List_Payments sSortBy, iPaymentServiceId  
%>
	  
	  <!-- END: payments LIST -->
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

'--------------------------------------------------------------------------------------------------
' string addBrackets( sValue, sValue2 )
'--------------------------------------------------------------------------------------------------
Function addBrackets( ByVal sValue, ByVal sValue2 )
	Dim sReturnValue

	sReturnValue = ""

	If Ucase(sValue) = "COMPLETED" Or UCase(sValue) = "PROCESSED" Then
		sReturnValue = "<b>" & FormatCurrency(sValue2,2) & "</b>"
	Else 
		sReturnValue = "[" & FormatCurrency(sValue2,2) & "]"
	End If

	addBrackets = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' void List_Payments sSortBy, iPaymentServiceId 
'--------------------------------------------------------------------------------------------------
Sub List_Payments( ByVal sSortBy, ByVal iPaymentServiceId )
	Dim statArray(6), sSql, oRs
	
	i = 0
	If statusInProgress = "yes" Then
		statArray(i) = " paymentstatus='INPROGRESS' OR"
		i = i + 1
	End If
	If statusPending= "yes" Then
		statArray(i) = " paymentstatus='PENDING' OR"
		i = i + 1
	End If
	If statusRefund = "yes" Then 
		statArray(i) = " paymentstatus='REFUND' OR"
		i = i + 1
	End If
	If statusDenied = "yes" Then
		statArray(i) = " paymentstatus='DENIED' OR"
		i = i + 1
	End If
	If statusCompleted= "yes" Then
		statArray(i) = " paymentstatus='COMPLETED' OR"
		i = i + 1
	End If
	If statusProcessed = "yes" Then 
		statArray(i) = " paymentstatus='PROCESSED' OR"
		i = i + 1
	End If

	For u = 0 To UBound(statArray)
		varStatClause = varStatClause & "" & statArray(u)
	Next 
	lenStatClause = len(varStatClause) - 3
	If lenStatClause > 1 Then 
		varStatClause = Left(varStatClause,lenStatClause)
	End If 


	varWhereClause = " WHERE (paymentDate >= '" & fromDate & "' AND paymentDate < '" & toDate & "') "

	varWhereClause = varWhereClause & " AND (" & varStatClause & ") AND orgid = " & session("orgid")

     	If CLng(iPaymentServiceId) <> CLng(0) Then
     		varWhereClause = varWhereClause & " AND paymentserviceid = " & iPaymentServiceId
     	End If 
	
	if sPaymentServiceId <> "" then
	        varWhereClause = varWhereClause & " AND payment_information LIKE '%" & sPaymentServiceId & "%'"
	end if


	If orderBy = "date" Then 
		sSql = "SELECT *,(SELECT SUM(paymentamount) AS Expr1 FROM dbo.egov_payment_list " &varWhereClause & ") as GRANDTOTAL FROM dbo.egov_payment_list " & varWhereClause & " ORDER BY paymentdate DESC, paymentid DESC"
	elseif orderBy = "purchase" then
		sSql = "SELECT *,(SELECT SUM(paymentamount) AS Expr1 FROM dbo.egov_payment_list " &varWhereClause & ") as GRANDTOTAL FROM dbo.egov_payment_list " & varWhereClause & " ORDER BY paymentid"
	Else 
		sSql = "SELECT *,(SELECT SUM(paymentamount) AS Expr1 FROM dbo.egov_payment_list " &varWhereClause & ") as GRANDTOTAL FROM dbo.egov_payment_list " & varWhereClause & " ORDER BY paymentservicename, paymentid DESC"
	End If 

	session("sPaymentSql") = sSql
	'response.write sSql & "<br /><br />"

	lastTitle = "Test"
	lastDate = "1/1/02"

	Set oRs = Server.CreateObject("ADODB.Recordset")

	If subTotals <> "yes" Then 
		' SET PAGE SIZE AND RECORDSET PARAMETERS
		oRs.PageSize = 5
		oRs.CacheSize = 5
		oRs.CursorLocation = 3
	End If

	' OPEN RECORDSET
	oRs.Open sSql, Application("DSN"), 3, 1
 
	If oRs.EOF Then 
		Response.write "<p><strong>No payment records found.</strong></p>"
	Else 
		If subTotals <> "yes" Then 
		' SET PAGE TO VIEW
			If Len(Request("pagenum")) = 0 OR clng(Request("pagenum")) < 1  Then
				oRs.AbsolutePage = 1
			Else
				If clng(Request("pagenum")) <=oRs.PageCount Then
					oRs.AbsolutePage = Request("pagenum")
				Else
					oRs.AbsolutePage = 1
				End If
			End If

			' DISPLAY RECORD STATISTICS
			Dim abspage, pagecnt
			abspage = oRs.AbsolutePage
			pagecnt = oRs.PageCount
			Response.write "<b>Page <font color=""blue""> " & oRs.AbsolutePage & "</font> "
			Response.Write " of <font color=""blue"">" & oRs.PageCount & "</font></b> &nbsp;|&nbsp; " & vbcrlf
			Response.Write " <b><font color=""blue"">" & oRs.RecordCount & "</font> total Online Payments</b>"

			' DISPLAY FORWARD AND BACKWARD NAVIGATION TOP					

			'Response.write "<div><table width=""100%""><tr><td valign=top><table><tr><td><a href=""action_line_list.asp?pagenum="&abspage - 1&"""><img border=0 src=""../images/arrow_back.gif""></a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage - 1&""">BACK</a></td><td valign=top>&nbsp;"  & "<a href=""action_line_list.asp?pagenum="&abspage + 1&""">NEXT</a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage + 1&"""><img border=0 src=""../images/arrow_forward.gif"" valign=bottom></a></td></tr></table></td><td width=450 align=right><a href=""action_line_list_print.asp?orderBy=" & orderBy & "&statusPending=" & statusPending & "&statusCompleted=" & statusCompleted & "&statusDenied=" & statusDenied & "&toDate=" & toDate & "&fromDate=" & fromDate & "&showDetail=" & showDetail & "&subTotals=" & subTotals & """ target=new>Open New Printer Friendly Results Window</a></td></tr></table></div>"
			Response.write "<div><table width=""100%""><tr><td valign=""top""><table><tr><td><a href=""action_line_list.asp?pagenum="&abspage - 1&"&"&sQueryString&"""><img border=0 src=""../images/arrow_back.gif""></a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage - 1&"&"&sQueryString&""">BACK</a></td><td valign=top>&nbsp;"  & "<a href=""action_line_list.asp?pagenum="&abspage + 1&"&"&sQueryString&""">NEXT</a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage + 1&"&"&sQueryString&"""><img border=0 src=""../images/arrow_forward.gif"" valign=bottom></a></td><td>&nbsp;&nbsp;</td>"
			response.write "<td><img src=""../images/icon_checkmark.png"" border=""0"" /><a href=""javascript:document.all.FormProcess.submit();"" onClick=""javascript: return confirm('Performing this action will change status of the selected items to Processed. Are you sure you want to proceed?');"">Process</a></td></tr></table>"
			response.write "</td><td width=450 align=right><!--<a href=""action_line_list_print.asp?orderBy=" & orderBy & "&statusInProgress=" & statusInProgress & "&statusPending=" & statusPending & "&statusRefund=" & statusRefund & "&statusDenied=" & statusDenied & "&statusCompleted=" & statusCompleted & "&statusProcessed=" & statusProcessed & "&toDate=" & toDate & "&fromDate=" & fromDate & "&showDetail=" & showDetail & "&subTotals=" & subTotals & """ target=new>Open New Printer Friendly Results Window</a>-->    "
		
			' EXPORT LINK CODE BULLHEADCITY
			If session("orgid") = 11 Then
				response.write "<a href=""export_data_csv2.asp?options=" & Encode(varWhereClause) & """ target=""_EXPORT"" >DOWNLOAD/VIEW AS CSV FILE</a>"
			End If 

			' EXPORT LINK CODE DEMO city CODE
			'If session("orgid") = 5 Then
				'response.write "<a href=""export_data_csv.asp?options=" & Encode(varWhereClause) & """ target=""_EXPORT"" >DOWNLOAD/VIEW AS CSV FILE</a>"
			'End If 

			response.write "</td></tr></table></div>"	
		Else
			' DISPLAY TOTAL RECORDS	
			Response.write "<div><table width=""100%""><tr><td valign=""top""><table><tr><td><b><font color=""blue"">" & oRs.RecordCount & "</font> total Online Payments</b></td><td>&nbsp;&nbsp;&nbsp;<img src=../images/icon_checkmark.png border=0><a href=""javascript:document.all.FormProcess.submit();"" onClick=""javascript: return confirm('Performing this action will change status of the selected items to Processed. Are you sure you want to proceed?');"">Process</a></td></tr></table></td><td width=450 align=right><!--<a href=""action_line_list_print.asp?orderBy=" & orderBy & "&statusRefund=" & statusRefund & "&statusInProgress=" & statusInProgress & "&statusProcessed=" & statusProcessed & "&statusPending=" & statusPending & "&statusCompleted=" & statusCompleted & "&statusDenied=" & statusDenied & "&fromDate=" & fromDate & "&toDate=" & toDate & "&showDetail=" & showDetail & "&subTotals=" & subTotals & """ target=new>Open New Printer Friendly Results Window</a>-->"

			' EXPORT LINK CODE
			If session("orgid") = 11 Then
				response.write "<a href=""export_data_csv2.asp?options=" & Encode(varWhereClause) & """ target=""_EXPORT"" >DOWNLOAD/VIEW AS CSV FILE</a>"
			End If 

			' EXPORT LINK CODE DEMO city CODE
			'If session("orgid") = 5 Then
				'response.write "<a href=""export_data_csv.asp?options=" & Encode(varWhereClause) & """ target=""_EXPORT"" >DOWNLOAD/VIEW AS CSV FILE</a>"
			'End If 

			response.write "</td></tr></table></div>"
		
		End If
  
		response.write "<div class=""shadow"">"
		Response.Write "<table cellspacing=0 cellpadding=2 class=tablelist width=""100%"">"
		Response.Write "<tr class=tablelist><th>&nbsp;</th><th align=left>Payment ID</td><th>Payment Service</th><th>Transaction Date</th><th>Payment Amount</th><th>Status</th><th>Assigned To</th></tr>"
		Response.Write "<form name=FormProcess action=""process_respond.asp?orderBy=" & orderBy & "&statusInProgress=" & statusInProgress & "&statusPending=" & statusPending & "&statusRefund=" & statusRefund & "&statusDenied=" & statusDenied & "&statusCompleted=" & statusCompleted & "&statusProcessed=" & statusProcessed & "&toDate=" & toDate & "&fromDate=" & fromDate & "&showDetail=" & showDetail & "&subTotals=" & subTotals & """ method=post>"

		' LOOP AND DISPLAY THE RECORDS
		bgcolor = "#eeeeee"
	
		If subTotals = "yes" Then 
			MagicNumber = oRs.RecordCount
		Else
			MagicNumber = oRs.PageSize
		End If 	

		For intRec = 1 To MagicNumber
			If Not oRs.EOF Then
				curGrandTotal = oRs("GRANDTOTAL")

				If bgcolor="#eeeeee" Then
					bgcolor="#ffffff" 
				Else
					bgcolor="#eeeeee"
				End If
	
			  ' GET VALUES
				If oRs("paymentservicename") <> "" Then
					sTitle = oRs("paymentservicename")
				Else
					sTitle = "<font color=red><b>???</b></font>"
				End If
		
				If oRs("paymentstatus") <> "" Then
					sStatus = oRs("paymentstatus")
				Else
					sStatus = "<font color=red><b>???</b></font>"
				End If
		
				If oRs("paymentdate") <> "" Then
					datSubmitDate = oRs("paymentdate")
					sDate = FormatDateTime(datSubmitDate, vbShortDate) 
				Else
					datSubmitDate = "<font color=""red""><b>???</b></font>"
				End If
		
				'INSERT BLANK ROW IF NEW CATEGORY OR DATE
				If subTotals="yes" Then 
					If orderBy = "date" Then 
						'if sDate = lastDate then
						If DateDiff("d",sDate,lastDate) = 0 Then 
							'NO NEW LINE
						Else 
							If lastDate <> "1/1/02" Then 
								Response.Write vbcrlf & "<tr bgcolor=""#dddddd""><td colspan=""4"">&nbsp;</td><td align=""center""><b><font color=navy>" & lastDate & " - " & formatcurrency(subTotl,2) & "</td><td colspan=2>&nbsp;</td>"
							End If 
						End If 

					Else 
						If sTitle = lastTitle Then 
							'NO NEW LINE
						Else 
							If  lastTitle <> "Test" Then 
								Response.Write vbcrlf & "<tr bgcolor=""#dddddd""><td colspan=""4"">&nbsp;</td><td align=""center""><b><font color=navy>" & lastTitle & " - " & formatcurrency(subTotl,2) & "</td><td colspan=2>&nbsp;</td>"
							End If 
						End If 
					End If 
				End If 
				
				lngTrackingNumber = oRs("paymentid") & Replace(FormatDateTime(oRs("paymentdate"),4),":","")
				
				Response.Write "<tr bgcolor=" & bgcolor & " onMouseOver=""this.style.backgroundColor='#93bee1';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';"">"

				p = p + 1

				If UCase(sStatus) = "PROCESSED" Then
					checkP = "checked" 
				Else
					checkP = ""
				End If 
				Response.Write "<td><input type=""checkbox"" value=" & oRs("paymentid") & " name=process_" & p & " " & checkP & "></td>"
				Response.Write "<td onClick=""location.href='action_respond.asp?control=" & oRs("paymentid") & "';""><b>" & lngTrackingNumber & "</b></td>"
				Response.Write "<td onClick=""location.href='action_respond.asp?control=" & oRs("paymentid") & "';""><b>" & sTitle & " </b></td>"
				Response.Write "<td onClick=""location.href='action_respond.asp?control=" & oRs("paymentid") & "';"" align=center> " & datSubmitDate & "</td>"
			
				If subTotals = "yes" Then 
					If orderBy = "date" Then 
						If DateDiff("d",sDate,lastDate) = 0 And UCase(sStatus) = "COMPLETED" Then 
							subTotl = oRs("paymentamount") + subTotl
							''**
						ElseIf DateDiff("d",sDate,lastDate) = 0 and UCase(sStatus) = "PROCESSED" Then 
							subTotl = oRs("paymentamount") + subTotl
						ElseIf UCase(sStatus) = "COMPLETED" Then 
							subTotl = oRs("paymentamount")
							''**
						ElseIf UCase(sStatus) = "PROCESSED" Then 
							subTotl = oRs("paymentamount")
						ElseIf DateDiff("d",sDate,lastDate) = 0 Then 
							' nothing
						Else 
							subTotl = 0
						End If 
						Response.Write "<td align=center onClick=""location.href='action_respond.asp?control=" & oRs("paymentid") & "';"" >" & addBrackets( sStatus,oRs("paymentamount") ) & "</td>"
						lastDate = sDate
					Else 
						If sTitle = lastTitle And UCase(sStatus) = "COMPLETED" Then 
							subTotl = oRs("paymentamount") + subTotl
							''**
						ElseIf sTitle = lastTitle And UCase(sStatus) = "PROCESSED" Then 
							subTotl = oRs("paymentamount") + subTotl
						ElseIf UCase(sStatus) = "COMPLETED" Then 
							subTotl = oRs("paymentamount")
							''**
						ElseIf UCase(sStatus) = "PROCESSED" Then 
							subTotl = oRs("paymentamount")
						ElseIf sTitle = lastTitle  Then 
							' nothing 
						Else 
							subTotl = 0
						End If 
						Response.Write "<td align=center onClick=""location.href='action_respond.asp?control=" & oRs("paymentid") & "';"" >" & addBrackets( sStatus,oRs("paymentamount") ) & "</td>"
						lastTitle = sTitle
					End If 
				Else 
					Response.Write "<td align=center onClick=""location.href='action_respond.asp?control=" & oRs("paymentid") & "';"" >" & addBrackets( sStatus,oRs("paymentamount") ) & "</td>"
				End If 

				Response.Write "<td align=center onClick=""location.href='action_respond.asp?control=" & oRs("paymentid") & "';"" > " & UCase(sStatus) & "</td>"
				Response.Write "<td align=center onClick=""location.href='action_respond.asp?control=" & oRs("paymentid") & "';"" > " & oRs("assignedName") & "</td></tr>"
				
				If showDetail = "yes" Then 
					Response.Write "<tr bgcolor=""" & bgcolor & """><td align=""left"" colspan=""7"" style=""padding-left:22px;""><table width=700px;><tr><td width=230px;><font color=navy>"
			
					Select Case session("payment_gateway")

						Case 1
							' PAY PAL PAYMENT GATEWAY
						
							' DISPLAY PAYMENT GATEWAY TRANSACTION INFORMATION
							DisplayGateWayTransactionInformationPayPal oRs("paymentsummary")
							response.write "</font></td><td width=230px;><font color=""navy"">"


							' DISPLAY PAYMENT GATEWAY USER INFORMATION
							DisplayGateWayInformationPayPal oRs("paymentsummary")
							response.write "</font></td><td width=230px;><font color=""navy"">"

						Case 2
							' SKIP JACK PAYMENT GATEWAY

							' DISPLAY PAYMENT GATEWAY TRANSACTION INFORMATION
							DisplayGateWayTransactionInformation oRs("paymentsummary")
							response.write "</font></td><td width=230px;><font color=navy>"


							' DISPLAY PAYMENT GATEWAY USER INFORMATION
							DisplayGateWayInformation oRs("paymentsummary")
							response.write "</font></td><td width=230px;><font color=navy>"

						Case 4 
							' VERSIGN GATEWAY USER INFORMATOIN
							
							' DISPLAY PAYMENT GATEWAY TRANSACTION INFORMATION
							DisplayGateWayTransactionInformationVerisign oRs("paymentid"), oRs("paymentrefid")
							response.write "</font></td><td width=230px;><font color=navy>"


							' DISPLAY PAYMENT GATEWAY USER INFORMATION
							DisplayGateWayInformationVerisign oRs("userfname"), oRs("useraddress"), oRs("userstate"), oRs("usercity"), oRs("userzip") 
							response.write "</font></td><td width=230px;><font color=navy>"
					
						Case Else
							' NO PAYMENT GATEWAY SPECIFIED
							' nothing
					End Select

					' DISPLAY PAYMENT SERVICE FIELDS
					If Trim(oRs("payment_information")) <> "" Then
						response.write "<span class=""paymentserviceinfo"">PAYMENT SERVICE INFO:</span><br /><font style=""font-size:10px;"">"
						response.Write UCase(oRs("payment_information")) & "</font>" 
					End If

					Response.Write "</font></td></tr></table></td></tr>"
				End If 

				oRs.MoveNext 
			End If
	  	Next
	 
		If subTotals = "yes" Then
			If orderBy = "date" Then 
				Response.Write "<tr bgcolor=#dddddd><td colspan=4>&nbsp;</td><td align=center><b><font color=navy>" & lastDate & " - " & formatcurrency(subTotl,2) & "</td><td colspan=2>&nbsp;</td>"
			Else 
				Response.Write "<tr bgcolor=#dddddd><td colspan=4>&nbsp;</td><td align=center><b><font color=navy>" & lastTitle & " - " & formatcurrency(subTotl,2) & "</td><td colspan=2>&nbsp;</td>"				
			End If 
		End If

		' DISPLAY GRANDTOTAL
		Response.Write vbcrlf & "<tr bgcolor=""#dddddd""><td colspan=""6"">&nbsp;</td><td align=center><b><font color=""navy""> GRANDTOTAL - " & formatcurrency(curGrandTotal,2) & "</td></tr>"	
	 
		Response.Write vbcrlf & "</table>"
		response.write vbcrlf & "</div><br /><br />"

		If subTotals <> "yes" Then
			' DISPLAY FORWARD AND BACKWARD NAVIGATION BOTTOM
			'Response.write "<div><table><tr><td valign=top><a href=""action_line_list.asp?pagenum="&abspage - 1&"""><img border=0 src=""../images/arrow_back.gif""></a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage - 1&""">BACK</a></td><td valign=top>&nbsp;"  & "<a href=""action_line_list.asp?pagenum="&abspage + 1&""">NEXT</a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage + 1&"""><img border=0 src=""../images/arrow_forward.gif"" valign=bottom></a></td></tr></table></div>"
			Response.write "<div><table border=0><tr><td valign=top><a href=""action_line_list.asp?pagenum="&abspage - 1&"&"&sQueryString&"""><img border=0 src=""../images/arrow_back.gif""></a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage - 1&"&"&sQueryString&""">BACK</a></td><td valign=top>&nbsp;"  & "<a href=""action_line_list.asp?pagenum="&abspage + 1&"&"&sQueryString&""">NEXT</a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage + 1&"&"&sQueryString&"""><img border=0 src=""../images/arrow_forward.gif"" valign=bottom></a></td><td>&nbsp;&nbsp;</td><td>"
			Response.write "<img src=../images/icon_checkmark.png border=0><a href=""javascript:document.all.FormProcess.submit();"" onClick=""javascript: return confirm('Performing this action will change status of the selected items to Processed. Are you sure you want to proceed?');"">Process</a></td>"
			Response.write vbcrlf & "</tr></table></div>"
		Else
			Response.write vbcrlf & "<div><table border=""0"" cellpadding=""0"" cellspacing=""0""><tr><td valign=""top""><img src=../images/icon_checkmark.png border=""0""><a href=""javascript:document.all.FormProcess.submit();"" onClick=""javascript: return confirm('Performing this action will change status of the selected items to Processed. Are you sure you want to proceed?');"">Process</a></td></tr></table></div>"
		End If
		response.write vbcrlf & "</div>"

	End If

	oRs.Close
	Set oRs = Nothing 

	response.write "<input type=""hidden"" name=""process_total"" value=""" & p & """ />"
	response.write vbcrlf & "</form>"

End Sub


'------------------------------------------------------------------------------------------------------------
' void DisplayGateWayInformation sText
'------------------------------------------------------------------------------------------------------------
Sub DisplayGateWayInformation( ByVal sText )
	Dim oDictionary, arrInfo, w
	
	' USED TO STORE DICTIONARY DATA
	Set oDictionary = Server.CreateObject("Scripting.Dictionary")

	' MAKE SURE THERE IS INFORMATION TO PARSE
	If sText <> "" Then

		' BREAK LIST INTO SEPARATE LINES
		arrInfo = Split(sTEXT,"<br>")
	
		' BREAK LINES INTO FIELD NAME AND VALUE
		For w = 0 to UBound(arrInfo)
			arrNamedPair = Split(arrInfo(w),":")

			' MATCHED SETS ARE ADDED TO DICTIONARY
			If UBound(arrNamedPair) > 0 Then
				oDictionary.Add Trim(arrNamedPair(0)),Trim(arrNamedPair(1))
			End If 
		Next

	End If

	' BUILD PERSONAL INFO DISPLAY
	response.write "<B><U><FONT style=""font-size:10px;"" COLOR=BLACK>BILL TO: </font></U></b><BR>"
	response.write UCase(oDictionary.Item("Cardmember Name")) & "<BR>"
	response.write UCase(oDictionary.Item("Street Address")) & "<BR>"
	response.write UCase(oDictionary.Item("City")) & ", " & UCase(oDictionary.Item("State")) & ", " & oDictionary.Item("Zipcode")

	Set oDictionary = Nothing

End Sub 


'------------------------------------------------------------------------------------------------------------
' void  DisplayGateWayTransactionInformation sText
'------------------------------------------------------------------------------------------------------------
Sub DisplayGateWayTransactionInformation( ByVal sText )
	Dim oDictionary, arrInfo, w
	
	' USED TO STORE DICTIONARY DATA
	Set oDictionary = Server.CreateObject("Scripting.Dictionary")

	' MAKE SURE THERE IS INFORMATION TO PARSE
	If sText <> "" Then

		' BREAK LIST INTO SEPARATE LINES
		arrInfo = Split(sTEXT,"<br>")

		' BREAK LINES INTO FIELD NAME AND VALUE
		For w = 0 to UBound(arrInfo)
			arrNamedPair = Split(arrInfo(w),":")

			' MATCHED SETS ARE ADDED TO DICTIONARY
			If UBound(arrNamedPair) > 0 Then
				oDictionary.Add Trim(arrNamedPair(0)),Trim(arrNamedPair(1))
			End If 
		Next

	End If

	' BUILD PERSONAL INFO DISPLAY
	response.write "<B><FONT style=""font-size:10px;"" COLOR=BLACK>EGOV ORDER ID: </font></b><FONT style=""font-size:10px;"">" & UCase(oDictionary.Item("Order Number")) & "</font><BR>"
	response.write "<B><FONT style=""font-size:10px;"" COLOR=BLACK>PAYMENT ID: </font></b><FONT style=""font-size:10px;"">" & UCase(oDictionary.Item("Transaction File Name")) & "</font><BR>"

	Set oDictionary = Nothing

End Sub 


'------------------------------------------------------------------------------------------------------------
' void DisplayGateWayTransactionInformationPayPal sText
'------------------------------------------------------------------------------------------------------------
Sub DisplayGateWayTransactionInformationPayPal( ByVal sText )
	Dim oDictionary, arrInfo, w

	' USED TO STORE DICTIONARY DATA
	Set oDictionary = Server.CreateObject("Scripting.Dictionary")

	' MAKE SURE THERE IS INFORMATION TO PARSE
	If sText <> "" Then

		' BREAK LIST INTO SEPARATE LINES
		arrInfo = Split(sText, "</br>")

		' BREAK LINES INTO FIELD NAME AND VALUE
		For w = 0 to UBound(arrInfo)

			arrNamedPair = Split(arrInfo(w),":")

			' MATCHED SETS ARE ADDED TO DICTIONARY
			If UBound(arrNamedPair) > 0 Then
				oDictionary.Add Trim(arrNamedPair(0)),Trim(arrNamedPair(1))
			End If 
		Next

	End If

	' BUILD PERSONAL INFO DISPLAY
	'response.write "<B><FONT COLOR=BLACK>EGOV ORDER ID: </font></b>" & UCASE(oDictionary.Item("Order Number")) & "<BR>"
	response.write "<B><FONT style=""font-size:10px;"" COLOR=BLACK>PAY PAL REFERENCE ID: </font></b><FONT style=""font-size:10px;"">" & UCase(oDictionary.Item("txn_id")) & "</font><BR>"

	Set oDictionary = Nothing

End Sub


'------------------------------------------------------------------------------------------------------------
' void DisplayGateWayInformationPayPal sText
'------------------------------------------------------------------------------------------------------------
Sub  DisplayGateWayInformationPayPal( ByVal sText )
	Dim oDictionary, arrInfo, w
	
	' USED TO STORE DICTIONARY DATA
	Set oDictionary = Server.CreateObject("Scripting.Dictionary")

	' MAKE SURE THERE IS INFORMATION TO PARSE
	If sText <> "" Then

		' BREAK LIST INTO SEPARATE LINES
		arrInfo = Split(sText, "</br>")

		' BREAK LINES INTO FIELD NAME AND VALUE
		For w = 0 To UBound(arrInfo)

			arrNamedPair = Split(arrInfo(w),":")

			' MATCHED SETS ARE ADDED TO DICTIONARY
			If UBound(arrNamedPair) > 0 Then
				oDictionary.Add Trim(arrNamedPair(0)),Trim(arrNamedPair(1))
			End If 
		Next

	End If

	' BUILD PERSONAL INFO DISPLAY
	response.write "<B><U><FONT style=""font-size:10px;"" COLOR=BLACK>BILL TO: </font></U></b><BR><FONT style=""font-size:10px;"">"
	response.write UCase(oDictionary.Item("address_name")) & "<BR>"
	response.write UCase(oDictionary.Item("address_street")) & "<BR>"
	response.write UCase(oDictionary.Item("address_city")) & ", " & UCase(oDictionary.Item("address_state")) & ", " & oDictionary.Item("address_zip") & "</font>"
	
	Set oDictionary = Nothing

End Sub


'------------------------------------------------------------------------------------------------------------
' void DisplayGateWayTransactionInformationVerisign iID, iRef
'------------------------------------------------------------------------------------------------------------
Sub DisplayGateWayTransactionInformationVerisign( ByVal iID, ByVal iRef )
	
	' BUILD PERSONAL INFO DISPLAY
	response.write "<B><FONT COLOR=BLACK>EGOV ORDER ID: </font></b>ecC" & session("orgid") & "0" & iID & "<BR>"
	response.write "<B><FONT style=""font-size:10px;"" COLOR=BLACK>VERISIGN REFERENCE ID: </font></b><FONT style=""font-size:10px;"">" & iRef & "</font><BR>"

End Sub 


'------------------------------------------------------------------------------------------------------------
' void  DisplayGateWayInformationVerisign sName, sAddress, sCity, sState, sZip
'------------------------------------------------------------------------------------------------------------
Sub DisplayGateWayInformationVerisign( ByVal sName, ByVal sAddress, ByVal sCity, ByVal sState, ByVal sZip )
	
	' BUILD PERSONAL INFO DISPLAY
	response.write "<B><U><FONT style=""font-size:10px;"" COLOR=BLACK>BILL TO: </font></U></b><BR><FONT style=""font-size:10px;"">"
	response.write UCase(sName) & "<BR>"
	response.write UCase(sAddress) & "<BR>"
	response.write UCase(sCity) & ", " & UCase(sState) & ", " & sZip & "</font>"

End Sub 


'------------------------------------------------------------------------------------------------------------
' string  Encode( sIn )
'------------------------------------------------------------------------------------------------------------
Function Encode( ByVal sInValue )
    Dim x, y, abfrom, abto

    Encode = "": abfrom = ""

    For x = 0 To 25: abfrom = abfrom & Chr(65 + x): Next 
    For x = 0 To 25: abfrom = abfrom & Chr(97 + x): Next 
    For x = 0 To 9: abfrom = abfrom & CStr(x): Next 

    abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)

    For x = 1 to Len(sInValue): y = InStr(abfrom, Mid(sInValue, x, 1))
        If y = 0 Then
             Encode = Encode & Mid(sInValue, x, 1)
        Else
             Encode = Encode & Mid(abto, y, 1)
        End If
    Next

End Function 


'------------------------------------------------------------------------------------------------------------
' void ShowPaymentServicePicks iPaymentServiceId 
'------------------------------------------------------------------------------------------------------------
Sub ShowPaymentServicePicks( ByVal iPaymentServiceId )
	Dim sSql, oRs

	sSql = "SELECT paymentserviceid, paymentservicename FROM egov_paymentservices WHERE orgid = " & session("orgid")
	sSql = sSql & " AND paymentservice_type = 0 ORDER BY paymentservicename"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""paymentserviceid"">"
		response.write vbcrlf & "<option value=""0"" " 
		if isnumeric(iPaymentServiceId) then
		If CLng(iPaymentServiceId) = CLng(0) Then
			response.write " selected=""selected"" "
		End If 
		End If 
		response.write ">All Payment Services</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("paymentserviceid") & """ " 
			if isnumeric(iPaymentServiceId) then
			If CLng(iPaymentServiceId) = CLng(oRs("paymentserviceid")) Then
				response.write " selected=""selected"" "
			End If 
			End If 
			response.write ">" & oRs("paymentservicename") & "</option>"
			oRs.MoveNext
		Loop 
		'if request.cookies("user")("userid") = "6398" then
		if session("orgid") = "153" then
			selected = ""
			if sPaymentServiceId = "PERMITHOLDERTYPE : CURRENT RESIDENT RAILROAD PERMIT HOLDER" then selected = " selected"
			response.write "<option value=""PERMITHOLDERTYPE : CURRENT RESIDENT RAILROAD PERMIT HOLDER"" " & selected & " >CURRENT RESIDENT RAILROAD PERMIT HOLDER</option>"
			selected = ""
			if sPaymentServiceId = "PERMITHOLDERTYPE : CURRENT HIGHLAND/CEDAR PERMIT HOLDER" then selected = " selected"
			response.write "<option value=""PERMITHOLDERTYPE : CURRENT HIGHLAND/CEDAR PERMIT HOLDER"" " & selected & " >CURRENT HIGHLAND/CEDAR PERMIT HOLDER</option>"
			selected = ""
			if sPaymentServiceId = "PERMITHOLDERTYPE : Current Non-resident Railroad Permit Holder" then selected = " selected"
			response.write "<option value=""PERMITHOLDERTYPE : Current Non-resident Railroad Permit Holder"" " & selected & " >Current Non-resident Railroad Permit Holder</option>"
		end if
		'end if
		response.write vbcrlf & "</select>"
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------------------------------------
' String GetPaymentService( iPaymentServiceId )
'------------------------------------------------------------------------------------------------------------
Function GetPaymentService( ByVal iPaymentServiceId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(paymentservice,'') AS paymentservice "
	sSql = sSql & "FROM egov_paymentservices WHERE paymentserviceid = " & iPaymentServiceId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetPaymentService = oRs("paymentservice")
	Else 
		GetPaymentService = ""
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 

%>
