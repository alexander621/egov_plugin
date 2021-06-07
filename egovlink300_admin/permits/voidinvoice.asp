<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: voidinvoice.asp
' AUTHOR: Steve Loar
' CREATED: 06/17/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Voids invoices
'
' MODIFICATION HISTORY
' 1.0   06/17/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iInvoiceId, sSql, iInvoiceStatusId, dInvoicedAmount, iPermitStatusId, iPermitId, dNetJobValue

iInvoiceId = CLng(request("invoiceid"))

dNetJobValue = GetInvoiceNetJobValue( iInvoiceId ) 

' Get the void status id
iInvoiceStatusId = GetInvoiceStatusId( "isvoid" )

' Change the status of the invoice
sSql = "UPDATE egov_permitinvoices SET invoicestatusid = " & iInvoiceStatusId & ", netjobvalue = 0.00, isvoided = 1, "
sSql = sSql & " voidadmin = " & session("userid") & ", voiddate = dbo.GetLocalDate(" & Session("OrgID") & ",getdate()) "
sSql = sSql & " WHERE invoiceid = " & iInvoiceId
'response.write sSql & "<br /><br />"
RunSQL sSql

' Get the invoiced fees and decrement the invoiced total from the actual fees
sSql = "SELECT I.permitfeeid, ISNULL(I.invoicedamount,0.00) AS voidedamount, ISNULL(F.invoicedamount,0.00) AS invoicedamount "
sSql = sSql & " FROM egov_permitinvoiceitems I, egov_permitfees F "
sSql = sSql & " WHERE I.permitfeeid = F.permitfeeid AND I.invoiceid = " & iInvoiceId
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

Do While Not oRs.EOF
	dInvoicedAmount = CDbl(oRs("invoicedamount")) - CDbl(oRs("voidedamount"))
	sSql = "UPDATE egov_permitfees SET invoicedamount = " & dInvoicedAmount & " WHERE permitfeeid = " & oRs("permitfeeid")
	'response.write sSql & "<br /><br />"
	RunSQL sSql
	oRs.MoveNext
Loop 

oRs.Close
Set oRs = Nothing 

iPermitId = GetPermitIdByInvoiceId( iInvoiceId )  ' in permitcommonfunctions.asp
If dNetJobValue > CDbl(0.00) Then 
	' Move the net job value to another invoice
	SetJobValue iPermitId, dNetJobValue 
End If 

iPermitStatusId = GetPermitStatusId( iPermitId )  ' in permitcommonfunctions.asp
'MakeAPermitLogEntry( iPermitid, sActivity, sActivityComment, sInternalComment, sExternalComment, iPermitStatusId, iIsInspectionEntry, iIsReviewEntry, iIsActivityEntry, iPermitReviewId, iPermitInspectionId, iReviewStatusId, iInspectionStatusId )
MakeAPermitLogEntry iPermitId, "'Permit Invoice Voided'", "'Permit Invoice " & iInvoiceId & " Voided'", "NULL", "NULL", iPermitStatusId, 0, 0, 1, "NULL", "NULL", "NULL", "NULL"

response.write "Success"


'-------------------------------------------------------------------------------------------------
' Function GetInvoiceNetJobValue( iInvoiceId )
'-------------------------------------------------------------------------------------------------
Function GetInvoiceNetJobValue( ByVal iInvoiceId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(netjobvalue,0.00) AS netjobvalue "
	sSql = sSQl & " FROM egov_permitinvoices "
	sSql = sSQl & " WHERE invoiceid = " & iInvoiceId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetInvoiceNetJobValue = CDbl(oRs("netjobvalue")) 
	Else 
		GetInvoiceNetJobValue = CDbl(0.00) 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'-------------------------------------------------------------------------------------------------
' Sub SetJobValue( iPermitId, dNetJobValue )
'-------------------------------------------------------------------------------------------------
Sub SetJobValue( ByVal iPermitId, ByVal dNetJobValue )
	Dim sSql, oRs, dNewJobValue

	sSql = "SELECT invoiceid, ISNULL(netjobvalue,0.00) AS netjobvalue "
	sSql = sSql & " FROM egov_permitinvoices WHERE permitid = " & iPermitId
	sSql = sSql & " AND isvoided = 0 ORDER BY invoiceid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		dNewJobValue = CDbl(oRs("netjobvalue")) + dNetJobValue
		sSql = "UPDATE egov_permitinvoices SET netjobvalue = " & FormatNumber(dNewJobValue,2,,,0) & " WHERE invoiceid = " & oRs("invoiceid")
		RunSQL sSql
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 



%>
