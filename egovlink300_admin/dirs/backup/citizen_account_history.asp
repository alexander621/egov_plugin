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
Dim iUserId, sName, oRs, iActivityCount, cCurrentBalance, iRt

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "edit citizens" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

iUserId = request("u")
iActivityCount = 0
sName = GetCitizenName( iUserId )
cCurrentBalance = GetCitizenCurrentBalance( iUserId )

%>

<html>
<head>
	<title><%=langBSPayments%></title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />

	<script language="javascript">
	<!--

		function GoBack(sUrl)
		{
			//alert( sUrl);
			location.href='' + sUrl;
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
		<a href="javascript:GoBack('<%=Session("RedirectPage")%>');"><img src='../images/arrow_2back.gif' border="0" align='absmiddle'>&nbsp;&nbsp;<%=Session("RedirectLang")%></a><br /><br />

		<p>
			Current Balance = <%=FormatCurrency(cCurrentBalance,2)%><br /><br />
		</p>
		<p>
			<a href="update_citizen_account.asp?uid=<%=iUserId%>&entrytype=credit">Deposit</a> &nbsp; 
			<a href="update_citizen_account.asp?uid=<%=iUserId%>&entrytype=debit">Withdraw</a>
			<!--&nbsp; &nbsp; &nbsp; <a href="transfer_citizen_account.asp?uid=<%=iUserId%>">Transfer</a>-->
		</p>
<%
		
			response.write "<div class=""purchasereportshadow"">"
			Response.Write vbcrlf & "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""purchasereport"">"
			Response.Write vbcrlf & vbtab & "<tr class=""tablelist""><th align=""left"">Date</th><th align=""left"">Prior Balance</th><th align=""left"">Deposit</th><th align=""left"">Withdrawl</th><th align=""left"">Admin</th><th align=""left"">Notes</th><th align=""left"">Receipt</th></tr>"
			bgcolor = "#eeeeee"
			
			' Get the activities that they were part of
			sSql = "select P.paymentid, P.paymentdate, P.paymenttypeid, ET.journalentrytype, I.itemtype, L.entrytype, L.amount, L.priorbalance, U.firstname + ' ' + U.lastname as adminname, P.notes "
			sSql = sSql & " from egov_class_payment P, egov_accounts_ledger L, egov_item_types I, users U, egov_journal_entry_types ET "
			sSql = sSql & " where P.paymentid = L.paymentid and L.itemtypeid = I.itemtypeid and P.journalentrytypeid = ET.journalentrytypeid and "
			sSql = sSql & " P.adminuserid = U.userid and L.accountid = " & iUserId
'			and I.iscitizenaccount = 1
			sSql = sSql & " order by paymentdate desc"

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
					response.write "<td>" & oRs("paymentdate") & "</td>"
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
					response.write "<td>" & oRs("adminname") & "</td>"
					response.write "<td>" & oRs("notes") & "</td>"
					If oRs("entrytype") = "debit" Then
						iRt = "d"
					Else
						iRt = "c"
					End If 
					response.write "<td><a href=""../purchases/viewjournal.asp?uid=" & iUserId & "&pid=" & oRs("paymentid") & "&rt=" & iRt & "&it=" & Left(oRs("itemtype"),2) & "&jet=" & Left(oRs("journalentrytype"),1) & """>View</a></td>"
					response.write "</tr>"
					oRs.MoveNext 
				Loop 
			Else
				response.write vbcrlf & "<tr><td colspan=""5""><br />No Account Activity<br /><br /></td></tr>"
			End If 
			
		oRs.close
		Set oRs = Nothing 

		response.write vbcrlf & "</table>"
		response.write "</div>"

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