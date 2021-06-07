<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
orderBy = Request("orderBy")
subTotals = Request("subTotals")
showDetail = Request("showDetail")
fromDate = Request("fromDate")
toDate = Request("toDate")
today = Date()

If orderBy = "" or IsNull(orderBy) Then orderBy = "date" End If
If toDate = "" or IsNull(toDate) Then toDate = dateAdd("d",0,today) End If
If fromDate = "" or IsNull(fromDate) Then fromDate = dateAdd("yyyy",-1,today) End If

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


if noStatus = true then
   statusInProgress = "yes"
   statusPending = "yes"
   statusRefund = "yes"
   statusDenied = "yes"
   statusCompleted = "yes"
   statusProcessed = "yes"
end if
%>

<html>
<head>
  <title><%=langBSPayments%></title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>

<SCRIPT LANGUAGE="JavaScript">
  function checkStat() {
  if ( !(form1.statusInProgress.checked) &&  !(form1.statusPending.checked) && !(form1.statusRefund.checked) && !(form1.statusDenied.checked) &&  !(form1.statusCompleted.checked) && !(form1.statusProcessed.checked)) {
		alert("You must select the status.");
		form1.statusPending.focus();
		return false;
	}
  }
  function CheckAllStatus() {
	if (document.form1.CheckAllStat.checked) {
		document.form1.statusPending.checked = true;
		document.form1.statusCompleted.checked = true;
		document.form1.statusDenied.checked = true;
	} else {
		document.form1.statusPending.checked = false;
		document.form1.statusCompleted.checked = false;
		document.form1.statusDenied.checked = false;
	}
 }
 </SCRIPT>
 
   <script language="Javascript">
  <!--
    function doCalendar(ToFrom) {
      w = (screen.width - 350)/2;
      h = (screen.height - 350)/2;
      eval('window.open("calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
    }
  //-->
  </script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabPayments,1%>

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
	    <td><font size="+1"><b>(E-Gov Payment Receipt Manager) - Manage Online Submitted Payments</b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.go(-1)"><%=langBackToStart%></a></td>
    </tr>
	<tr>
    <td>
				 <!--BEGIN: SEARCH OPTIONS-->
				  <fieldset>
				  <legend><b>Search/Sorting Option(s)</b></legend>
				  <form action="action_line_list.asp" method=post name=form1  onSubmit="return checkStat()">
				  <table border=0>
				  <tr>
				  <td valign=top>
					  <b>From: 
					  <input type=text name="fromDate" value="<%=fromDate%>">
					  <a href="javascript:void doCalendar('From');"><img src="../images/calendar.gif" border=0></a>		 
				  </td>
				  <td>&nbsp;</td>
				   <td valign=top>
					<b>To:</b> 
					  <input type=text name="toDate" value="<%=dateAdd("d",-1,toDate)%>">
					  <a href="javascript:void doCalendar('To');"><img src="../images/calendar.gif" border=0></a>
				   </td>
				  </tr>
				  <tr>
					<td valign=top colspan=3>
					<%
					If statusInProgress = "yes" then check1 = "checked"
					If statusPending = "yes" Then check2 = "checked" 
					If statusRefund = "yes" Then check3 = "checked"
					If statusDenied = "yes" then check4 = "checked"
					If statusCompleted = "yes" Then check5 = "checked" 
					If statusProcessed = "yes" Then check6 = "checked"
					%>
					
					<b>Display:</b> 
					 <input type=checkbox name="statusInProgress" value="yes" <%=check1%>>In Progress
					 <input type=checkbox name="statusPending" value="yes" <%=check2%>>Pending
					 <input type=checkbox name="statusRefund" value="yes" <%=check3%>>Refund
					 <input type=checkbox name="statusDenied" value="yes" <%=check4%>>Denied
					 <input type=checkbox name="statusCompleted" value="yes" <%=check5%>>Completed
					 <input type=checkbox name="statusProcessed" value="yes" <%=check6%>>Processed<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

					 
					 <!--<input type=checkbox name="CheckAllStat" value="checked" onClick="CheckAllStatus();">Check All-->

					  <% if subTotals = "yes" then %>
							  <input type=checkbox name="subTotals" value="yes" checked>Subtotals
					  <% else %>
							<input type=checkbox name="subTotals" value="yes">Subtotals
						<% end if %>

				  <% if showDetail = "yes" then %>
							  <input type=checkbox name="showDetail" value="yes" checked>Details
					  <% else %>
							<input type=checkbox name="showDetail" value="yes">Details
						<% end if %>
				   </td>
				</tr>		
				<tr>
				  <td valign=top colspan=3>
					  <b>Order By: 
					    <select name="orderBy">
					  <% if orderBy = "date" then %>
							<option value="date">Date</option>
							<option value="service">Service</option>
						<% else %>	
							<option value="service">Service</option>
							<option value="date">Date</option>		 
						<% end if %>
					  </select>

					  <input type=submit value=" Search Payments ">
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
	  <!--BEGIN: ACTION LINE REQUEST LIST -->
      
		<% List_Action_Requests(sSortBy) %>
	  
	  <!-- END: ACTION LINE REQUEST LIST -->
      </td>
       
    </tr>
  </table>
</body>
</html>


<%
Function addBrackets(sValue,sValue2)
	sReturnValue = ""
	If Ucase(sValue) = "COMPLETED" OR Ucase(sValue) = "PROCESSED" Then
		sReturnValue = "<b>" & formatcurrency(sValue2,2) & "</b>"
	else
		sReturnValue = "[" & formatcurrency(sValue2,2) & "]"
	End If

	addBrackets = sReturnValue
End Function


Function List_Action_Requests(sSortBy)

Dim statArray(6)
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

for u = 0 to ubound(statArray)
	varStatClause = varStatClause & "" & statArray(u)
next
lenStatClause = len(varStatClause) - 3
if lenStatClause > 1 then
	varStatClause = left(varStatClause,lenStatClause)
end if

'if noStatus = true then
'	varWhereClause = " WHERE paymentDate > '1/1/2003' "
'else
'	varWhereClause = " WHERE (" & varStatClause & ")"
'end if


varWhereClause = " WHERE (paymentDate >= '" & fromDate & "' AND paymentDate < '" & toDate & "') "

varWhereClause = varWhereClause & " AND (" & varStatClause & ") AND orgid='" & session("orgid") & "'"


if orderBy = "date" then 
	'sSQL = "SELECT * FROM dbo.egov_payment_list " &varWhereClause & " ORDER BY paymentdate DESC"
	sSQL = "SELECT *,(SELECT SUM(paymentamount) AS Expr1 FROM dbo.egov_payment_list " &varWhereClause & ") as GRANDTOTAL FROM dbo.egov_payment_list " &varWhereClause & " ORDER BY paymentdate DESC"

else
	'sSQL = "SELECT * FROM dbo.egov_payment_list " &varWhereClause & " ORDER BY paymentservicename DESC"
	sSQL = "SELECT *,(SELECT SUM(paymentamount) AS Expr1 FROM dbo.egov_payment_list " &varWhereClause & ") as GRANDTOTAL FROM dbo.egov_payment_list " &varWhereClause & " ORDER BY paymentservicename DESC"
end if

lastTitle = "Test"
lastDate = "1/1/02"

Set oRequests = Server.CreateObject("ADODB.Recordset")

 If subTotals <> "yes" Then 
	 ' SET PAGE SIZE AND RECORDSET PARAMETERS
	 oRequests.PageSize = 5
	 oRequests.CacheSize = 5
	 oRequests.CursorLocation = 3
 End If

' OPEN RECORDSET
oRequests.Open sSQL, Application("DSN"), 3, 1
 
if oRequests.EOF then
	Response.write "<p><b>No records found</p>"
else

   If subTotals <> "yes" Then 
			 ' SET PAGE TO VIEW
			 If Len(Request("pagenum")) = 0 OR clng(Request("pagenum")) < 1  Then
						oRequests.AbsolutePage = 1
			 Else
						If clng(Request("pagenum")) <=oRequests.PageCount Then
							oRequests.AbsolutePage = Request("pagenum")
						Else
							oRequests.AbsolutePage = 1
						End If
			 End If

			 ' DISPLAY RECORD STATISTICS
			  Dim abspage, pagecnt
			  abspage = oRequests.AbsolutePage
			  pagecnt = oRequests.PageCount
			  Response.write "<b>Page <font color=blue> " & oRequests.AbsolutePage & "</font> "
			  Response.Write " of <font color=blue>" & oRequests.PageCount & "</font></b> &nbsp;|&nbsp; " & vbcrlf
			  Response.Write " <b><font color=blue>" & oRequests.RecordCount & "</font> total Online Payments</b>"

			 ' DISPLAY FORWARD AND BACKWARD NAVIGATION TOP					
									
			 'Response.write "<div><table width=""100%""><tr><td valign=top><table><tr><td><a href=""action_line_list.asp?pagenum="&abspage - 1&"""><img border=0 src=""../images/arrow_back.gif""></a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage - 1&""">BACK</a></td><td valign=top>&nbsp;"  & "<a href=""action_line_list.asp?pagenum="&abspage + 1&""">NEXT</a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage + 1&"""><img border=0 src=""../images/arrow_forward.gif"" valign=bottom></a></td></tr></table></td><td width=450 align=right><a href=""action_line_list_print.asp?orderBy=" & orderBy & "&statusPending=" & statusPending & "&statusCompleted=" & statusCompleted & "&statusDenied=" & statusDenied & "&toDate=" & toDate & "&fromDate=" & fromDate & "&showDetail=" & showDetail & "&subTotals=" & subTotals & """ target=new>Open New Printer Friendly Results Window</a></td></tr></table></div>"
  	    Response.write "<div><table width=""100%""><tr><td valign=top><table><tr><td><a href=""action_line_list.asp?pagenum="&abspage - 1&"&"&sQueryString&"""><img border=0 src=""../images/arrow_back.gif""></a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage - 1&"&"&sQueryString&""">BACK</a></td><td valign=top>&nbsp;"  & "<a href=""action_line_list.asp?pagenum="&abspage + 1&"&"&sQueryString&""">NEXT</a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage + 1&"&"&sQueryString&"""><img border=0 src=""../images/arrow_forward.gif"" valign=bottom></a></td><td>&nbsp;&nbsp;</td><td><img src=../images/icon_checkmark.png border=0><a href=""javascript:document.all.FormProcess.submit();"" onClick=""javascript: return confirm('Performing this action will change status of the selected items to Processed. Are you sure you want to proceed?');"">Process</a></td><td width=450 align=right><a href=csvExport.asp title=Export to CSV File>Export to CSV File</a></td></tr></table></td><td width=450 align=right>Export to CSV File</td><td width=450 align=right><!--<a href=""action_line_list_print.asp?orderBy=" & orderBy & "&statusInProgress=" & statusInProgress & "&statusPending=" & statusPending & "&statusRefund=" & statusRefund & "&statusDenied=" & statusDenied & "&statusCompleted=" & statusCompleted & "&statusProcessed=" & statusProcessed & "&toDate=" & toDate & "&fromDate=" & fromDate & "&showDetail=" & showDetail & "&subTotals=" & subTotals & """ target=new>Open New Printer Friendly Results Window</a>--></td></tr></table></div>"	
  	    

  
  Else
			' DISPLAY TOTAL RECORDS	
				Response.write "<div><table width=""100%""><tr><td valign=top><table><tr><td><b><font color=blue>" & oRequests.RecordCount & "</font> total Online Payments</b></td><td>&nbsp;&nbsp;&nbsp;<img src=../images/icon_checkmark.png border=0><a href=""javascript:document.all.FormProcess.submit();"" onClick=""javascript: return confirm('Performing this action will change status of the selected items to Processed. Are you sure you want to proceed?');"">Process</a></td></tr></table></td><td width=450 align=right><a href=csvExport.asp title=Export to CSV File>Export to CSV File</a></td><td width=450 align=right><!--<a href=""action_line_list_print.asp?orderBy=" & orderBy & "&statusRefund=" & statusRefund & "&statusInProgress=" & statusInProgress & "&statusProcessed=" & statusProcessed & "&statusPending=" & statusPending & "&statusCompleted=" & statusCompleted & "&statusDenied=" & statusDenied & "&fromDate=" & fromDate & "&toDate=" & toDate & "&showDetail=" & showDetail & "&subTotals=" & subTotals & """ target=new>Open New Printer Friendly Results Window</a>--></td></tr></table></div>"
  End If
  
  
  Response.Write "<table cellspacing=0 cellpadding=2 class=tablelist width=""100%"">"
  Response.Write "<tr class=tablelist><th>&nbsp;</th><th align=left>Payment ID</td><th>Payment Service</th><th>Transaction Date</th><th>Payment Amount</th><th>Status</th><th>Assigned To</th></tr>"
  Response.Write "<form name=FormProcess action=""process_respond.asp?orderBy=" & orderBy & "&statusInProgress=" & statusInProgress & "&statusPending=" & statusPending & "&statusRefund=" & statusRefund & "&statusDenied=" & statusDenied & "&statusCompleted=" & statusCompleted & "&statusProcessed=" & statusProcessed & "&toDate=" & toDate & "&fromDate=" & fromDate & "&showDetail=" & showDetail & "&subTotals=" & subTotals & """ method=post>"
  	
  ' LOOP AND DISPLAY THE RECORDS
  bgcolor = "#eeeeee"
	
	
	If subTotals = "yes" Then 
			MagicNumber = oRequests.RecordCount
	Else
			MagicNumber = oRequests.PageSize
	End If 	
	For intRec=1 To MagicNumber
	 	
	 	If Not oRequests.EOF Then

		curGrandTotal = oRequests("GRANDTOTAL")

				If bgcolor="#eeeeee" Then
					bgcolor="#ffffff" 
				Else
					bgcolor="#eeeeee"
				End If
	
			  ' GET VALUES
				If oRequests("paymentservicename") <> "" Then
					sTitle = oRequests("paymentservicename")
				Else
					sTitle = "<font color=red><b>???</b></font>"
				End If
		
				If oRequests("paymentstatus") <> "" Then
					sStatus = oRequests("paymentstatus")
				Else
					sStatus = "<font color=red><b>???</b></font>"
				End If
		
				If oRequests("paymentdate") <> "" Then
					datSubmitDate = oRequests("paymentdate")
					sDate = FormatDateTime(datSubmitDate, vbShortDate) 
				Else
					datSubmitDate = "<font color=red><b>???</b></font>"
				End If
		
		
				'INSERT BLANK ROW IF NEW CATEGORY OR DATE
				if subTotals="yes" then
					if orderBy = "date" then
							'if sDate = lastDate then
							if DateDiff("d",sDate,lastDate) = 0 then
									'NO NEW LINE
							else
									if lastDate <> "1/1/02" then
										Response.Write "<tr bgcolor=#dddddd><td colspan=4>&nbsp;</td><td align=center><b><font color=navy>" & lastDate & " - " & formatcurrency(subTotl,2) & "</td><td colspan=2>&nbsp;</td>"
									end if
							end if

					else
							if sTitle = lastTitle then
									'NO NEW LINE
							else
									if  lastTitle <> "Test" then
										Response.Write "<tr bgcolor=#dddddd><td colspan=4>&nbsp;</td><td align=center><b><font color=navy>" & lastTitle & " - " & formatcurrency(subTotl,2) & "</td><td colspan=2>&nbsp;</td>"
									end if
							end if
					end if
			  	end if
				
				lngTrackingNumber = oRequests("paymentid") & replace(FormatDateTime(oRequests("paymentdate"),4),":","")
				
				Response.Write "<tr bgcolor=" & bgcolor & " onMouseOver=""this.style.backgroundColor='#93bee1';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';"">"
				p = p + 1
				if UCASE(sStatus) = "PROCESSED" then checkP = "checked" else checkP = ""
				Response.Write "<td><input type=checkbox value=" & oRequests("paymentid") & " name=process_" & p & " " & checkP & "></td>"
				Response.Write "<td onClick=""location.href='action_respond.asp?control=" & oRequests("paymentid") & "';""><b>" & lngTrackingNumber & "</b></td>"
				Response.Write "<td onClick=""location.href='action_respond.asp?control=" & oRequests("paymentid") & "';""><b>" & sTitle & " </b></td>"
				Response.Write "<td onClick=""location.href='action_respond.asp?control=" & oRequests("paymentid") & "';"" align=center> " & datSubmitDate & "</td>"
			
					if subTotals="yes" then
								if orderBy = "date" then
											if DateDiff("d",sDate,lastDate) = 0 and UCASE(sStatus) = "COMPLETED" then
													subTotl = oRequests("paymentamount") + subTotl
											''**
											elseif DateDiff("d",sDate,lastDate) = 0 and UCASE(sStatus) = "PROCESSED" then
													subTotl = oRequests("paymentamount") + subTotl
											elseif UCASE(sStatus) = "COMPLETED" then
													subTotl = oRequests("paymentamount")
											''**
											elseif UCASE(sStatus) = "PROCESSED" then
													subTotl = oRequests("paymentamount")
											elseif DateDiff("d",sDate,lastDate) = 0 then
											
											else
													subTotl = 0
											end if
											Response.Write "<td align=center onClick=""location.href='action_respond.asp?control=" & oRequests("paymentid") & "';"" >" & addBrackets(sStatus,oRequests("paymentamount")) & "</td>"
											lastDate = sDate
								else
											if sTitle = lastTitle and UCASE(sStatus) = "COMPLETED" then
													subTotl = oRequests("paymentamount") + subTotl
											''**
											elseif sTitle = lastTitle and UCASE(sStatus) = "PROCESSED" then
													subTotl = oRequests("paymentamount") + subTotl
											elseif UCASE(sStatus) = "COMPLETED" then
													subTotl = oRequests("paymentamount")
											''**
											elseif UCASE(sStatus) = "PROCESSED" then
													subTotl = oRequests("paymentamount")
											elseif sTitle = lastTitle  then
											
											else
													subTotl = 0
											end if
											Response.Write "<td align=center onClick=""location.href='action_respond.asp?control=" & oRequests("paymentid") & "';"" >" & addBrackets(sStatus,oRequests("paymentamount")) & "</td>"
											lastTitle = sTitle
								end if
					else
								Response.Write "<td align=center onClick=""location.href='action_respond.asp?control=" & oRequests("paymentid") & "';"" >" & addBrackets(sStatus,oRequests("paymentamount")) & "</td>"
					end if
				Response.Write "<td align=center onClick=""location.href='action_respond.asp?control=" & oRequests("paymentid") & "';"" > " & UCASE(sStatus) & "</td>"
				Response.Write "<td align=center onClick=""location.href='action_respond.asp?control=" & oRequests("paymentid") & "';"" > " & oRequests("assignedName") & "</td></tr>"
				
				if showDetail="yes" then
					Response.Write "<tr bgcolor=" & bgcolor & "><td align=left colspan=7 style=""padding-left:22px;""><table width=700px;><tr><td width=230px;><font color=navy>"
			
				
					Select Case session("payment_gateway")

					Case 1
						' PAY PAL PAYMENT GATEWAY
					
						' DISPLAY PAYMENT GATEWAY TRANSACTION INFORMATION
						DisplayGateWayTransactionInformationPayPal(oRequests("paymentsummary"))
						response.write "</font></TD><TD  width=230px;><font color=navy>"


						' DISPLAY PAYMENT GATEWAY USER INFORMATION
						DisplayGateWayInformationPayPal(oRequests("paymentsummary"))
						response.write "</font></TD><TD  width=230px;><font color=navy>"

					Case 2
						' SKIP JACK PAYMENT GATEWAY

						' DISPLAY PAYMENT GATEWAY TRANSACTION INFORMATION
						DisplayGateWayTransactionInformation(oRequests("paymentsummary"))
						response.write "</font></TD><TD  width=230px;><font color=navy>"


						' DISPLAY PAYMENT GATEWAY USER INFORMATION
						DisplayGateWayInformation(oRequests("paymentsummary"))
						response.write "</font></TD><TD  width=230px;><font color=navy>"

					Case 4 
						' VERSIGN GATEWAY USER INFORMATOIN
						
						' DISPLAY PAYMENT GATEWAY TRANSACTION INFORMATION
						DisplayGateWayTransactionInformationVerisign oRequests("paymentid"),oRequests("paymentrefid")
						response.write "</font></TD><TD  width=230px;><font color=navy>"


						' DISPLAY PAYMENT GATEWAY USER INFORMATION
						DisplayGateWayInformationVerisign oRequests("userfname"),oRequests("useraddress"),oRequests("userstate"),oRequests("usercity"),oRequests("userzip") 
						response.write "</font></TD><TD  width=230px;><font color=navy>"

					
					Case Else
						' NO PAYMENT GATEWAY SPECIFIED
						' 

					End Select


					' DISPLAY PAYMENT SERVICE FIELDS
					If TRIM(oRequests("payment_information")) <> "" Then
						response.write "<B><U><FONT style=""font-size:10px;"" COLOR=BLACK>PAYMENT SERVICE INFO: </font></U></b><BR><FONT style=""font-size:10px;"">"
						response.Write UCASE(oRequests("payment_information")) & "</font>" 
					End If
	

					Response.Write "</font></td></tr></table></td></tr>"
				end if
		oRequests.MoveNext 
	  End If
	 
	 'If subTotals = "yes" Then
	 '			Loop
	' Else
	  		Next
	 'End if
	 
	If subTotals = "yes" Then
				if orderBy = "date" then
						Response.Write "<tr bgcolor=#dddddd><td colspan=4>&nbsp;</td><td align=center><b><font color=navy>" & lastDate & " - " & formatcurrency(subTotl,2) & "</td><td colspan=2>&nbsp;</td>"
				else
						Response.Write "<tr bgcolor=#dddddd><td colspan=4>&nbsp;</td><td align=center><b><font color=navy>" & lastTitle & " - " & formatcurrency(subTotl,2) & "</td><td colspan=2>&nbsp;</td>"				
				end if
	End If

	' DISPLAY GRANDTOTAL
	Response.Write "<tr bgcolor=#dddddd><td colspan=6>&nbsp;</td><td align=center><b><font color=navy> GRANDTOTAL - " & formatcurrency(curGrandTotal,2) & "</td></tr>"	
	 
	Response.Write "</table>"
	If subTotals <> "yes" Then
		' DISPLAY FORWARD AND BACKWARD NAVIGATION BOTTOM
		 'Response.write "<div><table><tr><td valign=top><a href=""action_line_list.asp?pagenum="&abspage - 1&"""><img border=0 src=""../images/arrow_back.gif""></a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage - 1&""">BACK</a></td><td valign=top>&nbsp;"  & "<a href=""action_line_list.asp?pagenum="&abspage + 1&""">NEXT</a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage + 1&"""><img border=0 src=""../images/arrow_forward.gif"" valign=bottom></a></td></tr></table></div>"
		  Response.write "<div><table border=0><tr><td valign=top><a href=""action_line_list.asp?pagenum="&abspage - 1&"&"&sQueryString&"""><img border=0 src=""../images/arrow_back.gif""></a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage - 1&"&"&sQueryString&""">BACK</a></td><td valign=top>&nbsp;"  & "<a href=""action_line_list.asp?pagenum="&abspage + 1&"&"&sQueryString&""">NEXT</a></td><td valign=top><a href=""action_line_list.asp?pagenum="&abspage + 1&"&"&sQueryString&"""><img border=0 src=""../images/arrow_forward.gif"" valign=bottom></a></td><td>&nbsp;&nbsp;</td><td>"
			Response.write "<img src=../images/icon_checkmark.png border=0><a href=""javascript:document.all.FormProcess.submit();"" onClick=""javascript: return confirm('Performing this action will change status of the selected items to Processed. Are you sure you want to proceed?');"">Process</a></td>"
		  Response.write "</tr></table></div>"
	Else
			Response.write "<div><table border=0><tr><td valign=top><img src=../images/icon_checkmark.png border=0><a href=""javascript:document.all.FormProcess.submit();"" onClick=""javascript: return confirm('Performing this action will change status of the selected items to Processed. Are you sure you want to proceed?');"">Process</a></td></tr></table></div>"
	End If

End If
response.write "<input type=hidden name=process_total value=" & p & ">"
response.write "</form>"
End Function


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' Function DisplayGateWayInformation(sText)
'------------------------------------------------------------------------------------------------------------
Function DisplayGateWayInformation(sText)
	
	' USED TO STORE DICTIONARY DATA
	Set oDictionary=Server.CreateObject("Scripting.Dictionary")

	' MAKE SURE THERE IS INFORMATION TO PARSE
	If sText <> "" Then
	
		' BREAK LIST INTO SEPARATE LINES
		arrInfo = SPLIT(sTEXT,"<br>")

		' BREAK LINES INTO FIELD NAME AND VALUE
		For w = 0 to UBOUND(arrInfo)
			arrNamedPair = SPLIT(arrInfo(w),":")
			
			' MATCHED SETS ARE ADDED TO DICTIONARY
			If UBOUND(arrNamedPair) > 0 Then
				oDictionary.Add TRIM(arrNamedPair(0)),Trim(arrNamedPair(1))
			End If 
		Next
	
	End If

	' BUILD PERSONAL INFO DISPLAY
	response.write "<B><U><FONT style=""font-size:10px;"" COLOR=BLACK>BILL TO: </font></U></b><BR>"
	response.write UCASE(oDictionary.Item("Cardmember Name")) & "<BR>"
	response.write UCASE(oDictionary.Item("Street Address")) & "<BR>"
	response.write UCASE(oDictionary.Item("City")) & ", " & UCASE(oDictionary.Item("State")) & ", " & oDictionary.Item("Zipcode")
	Set oDictionary = Nothing

End Function


'------------------------------------------------------------------------------------------------------------
' Function DisplayGateWayInformation(sText)
'------------------------------------------------------------------------------------------------------------
Function DisplayGateWayTransactionInformation(sText)
	
	' USED TO STORE DICTIONARY DATA
	Set oDictionary=Server.CreateObject("Scripting.Dictionary")

	' MAKE SURE THERE IS INFORMATION TO PARSE
	If sText <> "" Then
	
		' BREAK LIST INTO SEPARATE LINES
		arrInfo = SPLIT(sTEXT,"<br>")

		' BREAK LINES INTO FIELD NAME AND VALUE
		For w = 0 to UBOUND(arrInfo)
			arrNamedPair = SPLIT(arrInfo(w),":")
			
			' MATCHED SETS ARE ADDED TO DICTIONARY
			If UBOUND(arrNamedPair) > 0 Then
				oDictionary.Add TRIM(arrNamedPair(0)),Trim(arrNamedPair(1))
			End If 
		Next

	End If

	' BUILD PERSONAL INFO DISPLAY
	response.write "<B><FONT style=""font-size:10px;"" COLOR=BLACK>EGOV ORDER ID: </font></b><FONT style=""font-size:10px;"">" & UCASE(oDictionary.Item("Order Number")) & "</font><BR>"
	response.write "<B><FONT style=""font-size:10px;"" COLOR=BLACK>PAYMENT ID: </font></b><FONT style=""font-size:10px;"">" & UCASE(oDictionary.Item("Transaction File Name")) & "</font><BR>"

	Set oDictionary = Nothing

End Function


'------------------------------------------------------------------------------------------------------------
' Function DisplayGateWayInformationPayPal(sText)
'------------------------------------------------------------------------------------------------------------
Function DisplayGateWayTransactionInformationPayPal(sText)
	

	' USED TO STORE DICTIONARY DATA
	Set oDictionary=Server.CreateObject("Scripting.Dictionary")

	' MAKE SURE THERE IS INFORMATION TO PARSE
	If sText <> "" Then
	
		' BREAK LIST INTO SEPARATE LINES
		arrInfo = SPLIT(sText, "</br>")

		' BREAK LINES INTO FIELD NAME AND VALUE
		For w = 0 to UBOUND(arrInfo)
			
			arrNamedPair = SPLIT(arrInfo(w),":")

			' MATCHED SETS ARE ADDED TO DICTIONARY
			If UBOUND(arrNamedPair) > 0 Then
				oDictionary.Add TRIM(arrNamedPair(0)),Trim(arrNamedPair(1))
			End If 
		Next

	End If

	' BUILD PERSONAL INFO DISPLAY
	'response.write "<B><FONT COLOR=BLACK>EGOV ORDER ID: </font></b>" & UCASE(oDictionary.Item("Order Number")) & "<BR>"
	response.write "<B><FONT style=""font-size:10px;"" COLOR=BLACK>PAL PAY REFERENCE ID: </font></b><FONT style=""font-size:10px;"">" & UCASE(oDictionary.Item("txn_id")) & "</font><BR>"

	Set oDictionary = Nothing

End Function


'------------------------------------------------------------------------------------------------------------
' Function DisplayGateWayInformationPayPal(sText)
'------------------------------------------------------------------------------------------------------------
Function DisplayGateWayInformationPayPal(sText)
	
	' USED TO STORE DICTIONARY DATA
	Set oDictionary=Server.CreateObject("Scripting.Dictionary")

	' MAKE SURE THERE IS INFORMATION TO PARSE
	If sText <> "" Then
	
		' BREAK LIST INTO SEPARATE LINES
		arrInfo = SPLIT(sText, "</br>")

		' BREAK LINES INTO FIELD NAME AND VALUE
		For w = 0 to UBOUND(arrInfo)
			
			arrNamedPair = SPLIT(arrInfo(w),":")

			' MATCHED SETS ARE ADDED TO DICTIONARY
			If UBOUND(arrNamedPair) > 0 Then
				oDictionary.Add TRIM(arrNamedPair(0)),Trim(arrNamedPair(1))
			End If 
		Next

	End If

	' BUILD PERSONAL INFO DISPLAY
	response.write "<B><U><FONT style=""font-size:10px;"" COLOR=BLACK>BILL TO: </font></U></b><BR><FONT style=""font-size:10px;"">"
	response.write UCASE(oDictionary.Item("address_name")) & "<BR>"
	response.write UCASE(oDictionary.Item("address_street")) & "<BR>"
	response.write UCASE(oDictionary.Item("address_city")) & ", " & UCASE(oDictionary.Item("address_state")) & ", " & oDictionary.Item("address_zip") & "</font>"
	Set oDictionary = Nothing

End Function


'------------------------------------------------------------------------------------------------------------
' Function DisplayGateWayInformationVerisign(sText)
'------------------------------------------------------------------------------------------------------------
Function DisplayGateWayTransactionInformationVerisign(iID,iRef)
	

	' BUILD PERSONAL INFO DISPLAY
	response.write "<B><FONT COLOR=BLACK>EGOV ORDER ID: </font></b>ecC" & session("orgid") & "0" & iID & "<BR>"
	response.write "<B><FONT style=""font-size:10px;"" COLOR=BLACK>VERISIGN REFERENCE ID: </font></b><FONT style=""font-size:10px;"">" & iRef & "</font><BR>"

	Set oDictionary = Nothing

End Function


'------------------------------------------------------------------------------------------------------------
' Function DisplayGateWayInformationVerisign(sText)
'------------------------------------------------------------------------------------------------------------
Function DisplayGateWayInformationVerisign(sName,sAddress,sCity,sState,sZip)
	
	' BUILD PERSONAL INFO DISPLAY
	response.write "<B><U><FONT style=""font-size:10px;"" COLOR=BLACK>BILL TO: </font></U></b><BR><FONT style=""font-size:10px;"">"
	response.write UCASE(sName) & "<BR>"
	response.write UCASE(sAddress) & "<BR>"
	response.write UCASE(sCity) & ", " & UCASE(sState) & ", " & sZip & "</font>"
	Set oDictionary = Nothing

End Function


'------------------------------------------------------------------------------------------------------------
' Function GetGrandTotal()
'------------------------------------------------------------------------------------------------------------
Function GetGrandTotal()

	response.write "<table>"
	response.write "<tr><td></td></tr>"
	response.write "</table>"

End Function
%>
