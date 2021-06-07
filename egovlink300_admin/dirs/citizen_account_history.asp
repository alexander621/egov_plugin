<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: citizen_account_history.asp
' AUTHOR: Steve Loar
' CREATED: 01/31/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the history of activity on a citizen's account
'
' MODIFICATION HISTORY
' 1.0   01/31/2007   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iUserId, sName, oRs, iActivityCount, cCurrentBalance, iRt, sRedirectPage, sRedirectLang
Dim paymentDate
response.Expires = 60
response.Expiresabsolute = Now() - 1
response.AddHeader "pragma","no-store"
response.AddHeader "cache-control","private"
response.CacheControl = "no-store" 'HTTP prevent back button

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "edit citizens" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

iUserId = request("u")
iActivityCount = 0
sName = GetCitizenName( iUserId )
cCurrentBalance = GetCitizenCurrentBalance( iUserId )

If Session("RedirectSubPage") <> "" Then 
	sRedirectPage = Session("RedirectSubPage")
	sRedirectLang = Session("RedirectSubLang")
	'Session("RedirectSubPage") = ""
	'Session("RedirectSubLang") = ""
Else
	sRedirectPage = Session("RedirectPage")
	sRedirectLang = Session("RedirectLang")
End If 

%>

<html>
<head>
	<title><%=langBSPayments%></title>
	<meta http-equiv="Content-type" content="text/html;charset=UTF-8">

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="reservationliststyles.css" />

	<script language="javascript">
	<!--

		function GoBack(sUrl)
		{
			//alert( sUrl);
			//location.href='' + sUrl;
			location.href = '<%=sRedirectPage%>';
		}

		function goToTransaction( _EntryType )
		{
			location.href = "update_citizen_account.asp?uid=<%=iUserId%>&entrytype=" + _EntryType;
		}

	//-->
	</script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="content">
	<div id="centercontent">
		<font size="+1"><b>Account History of <%=sName%></b></font><br /><br />

		<!-- <a href="javascript:GoBack('<%=sRedirectPage%>');"><img src='../images/arrow_2back.gif' border="0" align='absmiddle'>&nbsp;&nbsp;<%=sRedirectLang%></a> -->
		<input type="button" class="button" value="<< <%=sRedirectLang%>" onclick="location.href='<%=sRedirectPage%>';" />
		<br /><br />

		<h3>
			Current Balance = <%=FormatCurrency(cCurrentBalance,2)%>
		</h3>
		<p>
			<!-- <a href="update_citizen_account.asp?uid=<%=iUserId%>&entrytype=credit">Deposit</a> &nbsp; 
			<a href="update_citizen_account.asp?uid=<%=iUserId%>&entrytype=debit">Withdraw</a> &nbsp; 
			&nbsp; &nbsp; &nbsp; <a href="transfer_citizen_account.asp?uid=<%=iUserId%>">Transfer</a> -->
			<input type="button" class="button" value="Deposit Funds" onclick="goToTransaction( 'credit' )" /> &nbsp; 
			<input type="button" class="button" value="Withdraw Funds" onclick="goToTransaction( 'debit' )" />
		</p>
<%

			' Get the activities that they were part of
			sSql = "SELECT P.paymentid, P.paymentdate, P.paymenttypeid, ET.journalentrytype, I.itemtype, L.entrytype, L.amount, L.priorbalance, "
			sSql = sSql & "U.firstname + ' ' + U.lastname as adminname, P.notes, ISNULL(P.reservationid,0) AS reservationid "
			sSql = sSql & "FROM egov_class_payment P, egov_accounts_ledger L, egov_item_types I, users U, egov_journal_entry_types ET "
			sSql = sSql & "WHERE P.paymentid = L.paymentid AND L.itemtypeid = I.itemtypeid AND P.journalentrytypeid = ET.journalentrytypeid AND "
			sSql = sSql & "P.adminuserid = U.userid AND L.accountid = " & iUserId
			sSql = sSql & " ORDER BY paymentdate DESC"
			'response.write sSql & "<br /><br />"
		
			Response.Write vbcrlf & "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""citizenreport"">"
			Response.Write vbcrlf & vbtab & "<tr class=""tablelist""><th align=""left"">Date</th><th align=""left"">Prior Balance</th><th align=""left"">Deposit</th><th align=""left"">Withdrawl</th><th align=""left"">Admin</th><th align=""left"">Notes</th><th align=""left"">Receipt</th></tr>"
			bgcolor = "#eeeeee"

			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open sSQL, Application("DSN"), 3, 1

			If Not oRS.EOF Then 
				' LOOP AND DISPLAY THE RECORDS
				Do While Not oRs.EOF 
					iActivityCount = iActivityCount + 1
					If bgcolor="#eeeeee" Then
						bgcolor="#ffffff" 
					Else
						bgcolor="#eeeeee"
					End If			

					Response.Write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """ class=""tablelist"">"

					paymentDate = CDate(oRs("paymentdate"))
					response.write "<td nowrap>" & FormatDateTime( paymentDate, 2) & "</td>"

					response.write "<td>" & FormatCurrency(oRs("priorbalance"),2) & "</td>"

					response.write "<td>"
					If oRs("entrytype") = "credit" Then
						response.write FormatCurrency(oRs("amount"),2) 
					Else
						response.write "&nbsp;"
					End If
					response.write "</td>"
					response.write "<td>" 
					If oRs("entrytype") = "debit" Then
						response.write FormatCurrency(oRs("amount"),2) 
					Else 
						response.write "&nbsp;"
					End If 
					response.write "</td>"
					response.write "<td nowrap>" & oRs("adminname") & "</td>"
					response.write "<td>" & oRs("notes") & "</td>"
					If oRs("entrytype") = "debit" Then
						iRt = "d"
					Else
						iRt = "c"
					End If 
					
					If oRs("itemtype") = "recreation activity" Then
						' classes and events
						response.write "<td><a href=""../classes/view_receipt.asp?iPaymentId=" & oRs("paymentid") & """>" & oRs("paymentid") & "</a></td>"
					Else
						If oRs("itemtype") = "rentals" Then
							' rentals
							response.write "<td><a href=""../rentals/viewpaymentreceipt.asp?paymentid=" & oRs("paymentid") & """>" & oRs("paymentid") & "</a></td>"
						Else
							' everything else ??
							'response.write "<td><a href=""../purchases/viewjournal.asp?uid=" & iUserId & "&pid=" & oRs("paymentid") & "&rt=" & iRt & "&it=" & Left(oRs("itemtype"),2) & "&jet=" & Left(oRs("journalentrytype"),1) & """>View</a></td>"
							response.write "<td><a href=""../purchases/viewjournal.asp?uid=" & iUserId & "&pid=" & oRs("paymentid") & "&rt=" & iRt & "&it=" & Left(oRs("itemtype"),2) & "&jet=" & Left(oRs("journalentrytype"),1) & """>" & oRs("paymentid") & "</a></td>"
						End If 
					End If 
					
					response.write "</tr>"
					oRs.MoveNext 
				Loop 
			Else
				response.write vbcrlf & "<tr><td colspan=""5""><br />No account activity was found for this person.<br /><br /></td></tr>"
			End If 
			
		oRs.close
		Set oRs = Nothing 

		response.write vbcrlf & "</table>"

%>				  
	</div>
</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'------------------------------------------------------------------------------------------------------------


%>