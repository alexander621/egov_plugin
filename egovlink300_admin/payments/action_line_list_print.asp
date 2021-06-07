<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
orderBy = Request("orderBy")
subTotals = Request("subTotals")
showDetail = Request("showDetail")
fromDate = Request("fromDate")
toDate = Request("toDate")

If orderBy = "" or IsNull(orderBy) Then orderBy = "date" End If

'The Dates will never be blank for the print window
'If toDate = "" or IsNull(fromDate) Then toDate = dateAdd("d",0,today) End If
'If fromDate = "" or IsNull(toDate) Then fromDate = dateAdd("m",-1,today) End If
'toDate = dateAdd("d",1,toDate)

If subTotals = "yes" Then 
ELSE
	subTotals = "no"
End If

If showDetail = "yes" Then 
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
  <script language="JavaScript">
  function printit(){
 		if (window.print) {
 			window.print() ;
 		} else {
 			var WebBrowser = '<OBJECT ID="WebBrowser1" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>';
 			document.body.insertAdjacentHTML('beforeEnd', WebBrowser);
 			WebBrowser1.ExecWB(6, 2);//Use a 1 vs. a 2 for a prompting dialog box
 			WebBrowser1.outerHTML = "";
 		}
 	}
	</script></head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">


  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    
    
    <!--<tr>
      <td><font size="+1"><b>Manage Online Submitted Payments</b></font></td>
      <td width="350"><B><a href="javascript:printit()">PRINT</a></a> | <b><a href="javascript:window.close()">CLOSE WINDOW</A></td>
    </tr>-->


	<tr>

      <td colspan="2" valign="top">
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
	If Ucase(sValue) = "COMPLETED" Then
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

varWhereClause = varWhereClause & " AND (" & varStatClause & ")  AND orgid='" & session("orgid") & "'" 


if orderBy = "date" then 
	sSQL = "SELECT * FROM dbo.egov_payment_list " &varWhereClause & " ORDER BY paymentdate DESC"
else
	sSQL = "SELECT * FROM dbo.egov_payment_list " &varWhereClause & " ORDER BY paymentservicename DESC"
end if

lastTitle = "Test"
lastDate = "1/1/02"

Set oRequests = Server.CreateObject("ADODB.Recordset")

 ' OPEN RECORDSET
 oRequests.Open sSQL, Application("DSN"), 3, 1
 
 

	' DISPLAY TOTAL RECORDS	
		Response.write "<div><table><tr><td valign=top><b><font color=blue>" & oRequests.RecordCount & "</font> total Online Payments</b></td></tr></table></div>"
 
  Response.Write "<table cellspacing=0 cellpadding=2 class=tablelist width=""100%"">"
  Response.Write "<tr class=tablelist><th align=left>Payment ID</td><th>Payment Service</th><th>Transaction Date</th><th>Payment Amount</th><th>Status</th><th>Assigned To</th></tr>"

  ' LOOP AND DISPLAY THE RECORDS
  bgcolor = "#eeeeee"
	 do while Not oRequests.EOF

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
										Response.Write "<tr bgcolor=#dddddd><td colspan=3>&nbsp;</td><td align=center><b><font color=navy>" & lastDate & " - " & formatcurrency(subTotl,2) & "</td><td colspan=2>&nbsp;</td>"
									end if
							end if

					else
							if sTitle = lastTitle then
									'NO NEW LINE
							else
									if  lastTitle <> "Test" then
										Response.Write "<tr bgcolor=#dddddd><td colspan=3>&nbsp;</td><td align=center><b><font color=navy>" & lastTitle & " - " & formatcurrency(subTotl,2) & "</td><td colspan=2>&nbsp;</td>"
									end if
							end if
					end if
			  end if	
		
		lngTrackingNumber = oRequests("paymentid") & replace(FormatDateTime(oRequests("paymentdate"),4),":","")
		
		Response.Write "<tr bgcolor=" & bgcolor & " onClick=""location.href='action_respond.asp?control=" & oRequests("paymentid") & "';"" onMouseOver=""this.style.backgroundColor='#93bee1';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';""><td><b>" & lngTrackingNumber & "</b></td><td><b>" & sTitle & " </b></td><td align=center> " & datSubmitDate & "</td>"
		
				    if subTotals="yes" then
								if orderBy = "date" then
											if DateDiff("d",sDate,lastDate) = 0 and UCASE(sStatus) = "COMPLETED" then
													subTotl = oRequests("paymentamount") + subTotl
											elseif UCASE(sStatus) = "COMPLETED" then
													subTotl = oRequests("paymentamount")
											elseif DateDiff("d",sDate,lastDate) = 0 then
											
											else
													subTotl = 0
											end if
											Response.Write "<td align=center>" & addBrackets(sStatus,oRequests("paymentamount")) & "</td>"
											lastDate = sDate
								else
											if sTitle = lastTitle and UCASE(sStatus) = "COMPLETED" then
													subTotl = oRequests("paymentamount") + subTotl
											elseif UCASE(sStatus) = "COMPLETED" then
													subTotl = oRequests("paymentamount")
											elseif sTitle = lastTitle  then
											
											else
													subTotl = 0
											end if
											Response.Write "<td align=center>" & addBrackets(sStatus,oRequests("paymentamount")) & "</td>"
											lastTitle = sTitle
								end if
					else
								Response.Write "<td align=center>" & formatcurrency(oRequests("paymentamount"),2) & "</td>"
					end if
				Response.Write "<td align=center> " & UCASE(sStatus) & "</td><td align=center> " & oRequests("assignedName") & "</td></tr>"
				
				if showDetail="yes" then
					Response.Write "<tr bgcolor=" & bgcolor & "><td align=left colspan=6 style=""padding-left:22px;""><table width=700px;><tr><td width=230px;><font color=navy>"
				
					' DISPLAY PAYMENT GATEWAY TRANSACTION INFORMATION
					DisplayGateWayTransactionInformation(oRequests("paymentsummary"))
					response.write "</font></TD><TD  width=230px;><font color=navy>"


					' DISPLAY PAYMENT GATEWAY USER INFORMATION
					DisplayGateWayInformation(oRequests("paymentsummary"))
					response.write "</font></TD><TD  width=230px;><font color=navy>"


					' DISPLAY PAYMENT SERVICE FIELDS
					If TRIM(oRequests("payment_information")) <> "" Then
						response.write "<B><U><FONT COLOR=BLACK>PAYMENT SERVICE INFO: </font></U></b><BR>"
						response.Write UCASE(oRequests("payment_information")) 
					End If
	

					Response.Write "</font></td></tr></table></td></tr>"
				end if

		oRequests.MoveNext 
	  Loop
	  
	  If subTotals = "yes" Then
				if orderBy = "date" then
						Response.Write "<tr bgcolor=#dddddd><td colspan=3>&nbsp;</td><td align=center><b><font color=navy>" & lastDate & " - " & formatcurrency(subTotl,2) & "</td><td colspan=2>&nbsp;</td>"
				else
						Response.Write "<tr bgcolor=#dddddd><td colspan=3>&nbsp;</td><td align=center><b><font color=navy>" & lastTitle & " - " & formatcurrency(subTotl,2) & "</td><td colspan=2>&nbsp;</td>"				
				end if
	 End If
	 Response.Write "</table>"

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

	' BUILD PERSONAL INFO DISPLAY
	response.write "<B><U><FONT COLOR=BLACK>BILL TO: </font></U></b><BR>"
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

	' BUILD PERSONAL INFO DISPLAY
	response.write "<B><FONT COLOR=BLACK>EGOV ORDER ID: </font></b>" & UCASE(oDictionary.Item("Order Number")) & "<BR>"
	response.write "<B><FONT COLOR=BLACK>PAYMENT ID: </font></b>" & UCASE(oDictionary.Item("Transaction File Name")) & "<BR>"

	Set oDictionary = Nothing

End Function
%>