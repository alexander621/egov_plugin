<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: payment_accounts_update.asp
' AUTHOR: Steve Loar
' CREATED: 04/17/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page saves changes to gl accounts of payment methods
'
' MODIFICATION HISTORY
' 1.0   4/17/2007   Steve Loar - INITIAL VERSION
' 1.1	05/05/2010	Steve Loar - Changed to use RunSQLStatement
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, iRow

iRow = 1 

Do While iRow <= clng(request("maxrows"))
	' Loop through the payment methods and save accountids and default amount
	If request("defaultamount" & iRow) <> "" Then
		sAmount = ", defaultamount = " & CDbl(request("defaultamount" & iRow))
	Else
		sAmount = ", defaultamount = 0"
	End If 

	sSql = "UPDATE egov_organizations_to_paymenttypes SET accountid = " & CDbl(request("accountid" & iRow)) & sAmount
	sSql = sSql & " WHERE paymenttypeid = " & CLng(request("paymenttypeid" & iRow)) & " AND orgid = " & Session("OrgId")
	'session("payment_accounts_updateSql") = sSql
	'response.write sSql & "<br />"

	RunSQLStatement sSql
	'session("payment_accounts_updateSql") = ""

	iRow = iRow + 1
Loop 

' Return to the payment account management
response.redirect "payment_accounts.asp?s=u"

%>
