<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: correctinvoicejobvalue.asp
' AUTHOR: Steve Loar
' CREATED: 3/11/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This syncs the invoice job value to the permit job value for reports. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   3/11/2009   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sSql, oRs, dPermitJobValue, dInvoiceJobValue, dDifference, iLastInvoiceId, iLastInvoiceJobValue

iPermitId = CLng(request("permitid"))

' Get the permit job value
dPermitJobValue = GetPermitDetailItemAsNumber( iPermitId, "jobvalue", "double" )	' in permitcommonfunctions.asp

' Get the sum of invoice job values
dInvoiceJobValue = GetPriorJobValue( iPermitId )	' in permitcommonfunctions.asp

' calc the difference
dDifference =  dPermitJobValue - dInvoiceJobValue

If CDbl(dDifference) <> CDbl(0.00) then

'	get the last invoice and it's job value
	iLastInvoiceId = GetLastInvoiceIdAndJobValue( iPermitId, iLastInvoiceJobValue )
	iLastInvoiceJobValue = iLastInvoiceJobValue + dDifference

'	alter that value then update the invoice
	sSql = "UPDATE egov_permitinvoices SET netjobvalue = " & iLastInvoiceJobValue 
	sSql = sSql & " WHERE invoiceid = " & iLastInvoiceId
	RunSQL sSql

End If 

response.write "Success"


'--------------------------------------------------------------------------------------------------
' Function GetLastInvoiceIdAndJobValue( iPermitId, iLastInvoiceJobValue )
'--------------------------------------------------------------------------------------------------
Function GetLastInvoiceIdAndJobValue( iPermitId, iLastInvoiceJobValue )
	Dim sSql, oRs, iInvoiceId

	sSql = "SELECT invoiceid, ISNULL(netjobvalue, 0.00) AS netjobvalue "
	sSql = sSql & " FROM egov_permitinvoices WHERE permitid = " & iPermitId
	sSql = sSql & " ORDER BY invoiceid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		iInvoiceId = oRs("invoiceid")
		iLastInvoiceJobValue = CDbl(oRs("netjobvalue"))
		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

	GetLastInvoiceIdAndJobValue = iInvoiceId

End Function 


%>
