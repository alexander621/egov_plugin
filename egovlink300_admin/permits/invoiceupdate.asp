<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: invoiceupdate.asp
' AUTHOR: Steve Loar
' CREATED: 05/14/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Creates Invoices, called via AJAX from invoicecreate.asp
'
' MODIFICATION HISTORY
' 1.0   05/14/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, iMaxFeeCount, sInvoiceTotal, iPermitContactid, iInvoiceId, iPermitFeeId, sPermitFeeInvoicedAmount
Dim iPermitStatusId, iWaiveAllFees, iInvoiceStatusId, iCurrentJobValue, iPriorJobValue, iNetJobValue
Dim sPermitFeePrefix, sPermitFee, iPermitFeeCategoryTypeId, iFeeReportingTypeId, x, iIsPercentageTypeFee
Dim iDisplayOrder

iPermitId = CLng(request("permitid"))

iMaxFeeCount = CLng(request("maxFeeCount"))

sInvoiceTotal = CDbl(request("totalamount"))

iPermitContactid = CLng(request("permitcontactid"))

iWaiveAllFees = CLng(request("allfeeswaived"))

If iWaiveAllFees = CLng(0) Then
	iInvoiceStatusId = GetInvoiceStatusId("isinitialstatus")  ' in permitcommonfunctions.asp
Else
	iInvoiceStatusId = GetInvoiceStatusId("iswaived")  ' in permitcommonfunctions.asp
End If 

' Figure out the net change in job value for the Loveland monthly report to use
iCurrentJobValue = GetCurrentJobValue( iPermitId ) 
iPriorJobValue = GetPriorJobValue( iPermitId ) 
iNetJobValue = FormatNumber((iCurrentJobValue - iPriorJobValue),2,,,0)

' Create the invoice row
sSql = "INSERT INTO egov_permitinvoices ( orgid, permitid, invoicedate, totalamount, adminuserid, permitcontactid, "
sSql = sSql & " invoicestatusid, allfeeswaived, netjobvalue ) VALUES ( "
sSql = sSql & session("orgid") & ", " & iPermitId & ", dbo.GetLocalDate(" & Session("OrgID") & ",getdate()), " 
sSql = sSql & sInvoiceTotal & ", " & session("userid") & ", " & iPermitContactid & ", " & iInvoiceStatusId & ", "
sSql = sSql & iWaiveAllFees & ", " & iNetJobValue & " )"

'response.write sSql & "<br /><br />"
iInvoiceId = RunIdentityInsert( sSql )
'response.write "iInvoiceId = " & iInvoiceId & "<br />"

'response.write "MaxFees = " & iMaxFeeCount & "<br />"

'Create the invoice item rows and update the permit fees
For x = 1 To iMaxFeeCount
	'response.write "PermitFeeId = " & CLng(request("permitfeeid" & x))
	' If the row exists
	If LCase(request("include" & x)) = "true" Then
	'If LCase(request("include" & x)) = "on" Then
		iPermitFeeId = CLng(request("permitfeeid" & x))
		'response.write "iPermitFeeId = " & iPermitFeeId & "<br />"

		If CDbl(request("invoiceamount" & x)) >= CDbl(0.00) Then
			' Get the fee values that the invoice items need to know
			GetPermitValuesForInvoiceItems iPermitFeeId, sPermitFeePrefix, sPermitFee, iPermitFeeCategoryTypeId, iFeeReportingTypeId, iIsPercentageTypeFee, iDisplayOrder
			If CLng(iFeeReportingTypeId) = CLng(0) Then 
				iFeeReportingTypeId = "NULL"
			End If 

			' Create the invoice item row
			'response.write "invoiceamount" & x & " = " & CDbl(request("invoiceamount" & x)) & "<br />"
			sSql = "INSERT INTO egov_permitinvoiceitems ( invoiceid, orgid, permitid, permitfeeid, invoicedamount, permitfeeprefix, "
			sSql = sSql & " permitfee, permitfeecategorytypeid, feereportingtypeid, ispercentagetypefee, displayorder ) VALUES ( "
			sSql = sSql & iInvoiceId & ", " & Session("orgid") & ", " & iPermitId & ", " & CLng(request("permitfeeid" & x)) & ", "
			sSql = sSql & CDbl(request("invoiceamount" & x)) & ", '" & dbsafe(sPermitFeePrefix) & "', '" & dbsafe(sPermitFee) & "', "
			sSql = sSql & iPermitFeeCategoryTypeId & ", " & iFeeReportingTypeId & ", " & iIsPercentageTypeFee & ", " & iDisplayOrder & " )"
			'response.write sSql & "<br /><br />"
			RunSQL sSql 

			' Update the permit fee with a new invoiced total
			sPermitFeeInvoicedAmount = GetPermitFeeInvoicedAmount( CLng(request("permitfeeid" & x)) )   ' in permitcommonfunctions.asp
			sPermitFeeInvoicedAmount = sPermitFeeInvoicedAmount + CDbl(request("invoiceamount" & x))
			sSql = "UPDATE egov_permitfees SET invoicedamount = " & sPermitFeeInvoicedAmount & " WHERE permitfeeid = " & CLng(request("permitfeeid" & x))
'			response.write sSql & "<br /><br />"
			RunSQL sSql
		End If 
	End If 
Next 

iPermitStatusId = GetPermitStatusId( iPermitId )  ' in permitcommonfunctions.asp
'MakeAPermitLogEntry( iPermitid, sActivity, sActivityComment, sInternalComment, sExternalComment, iPermitStatusId, iIsInspectionEntry, iIsReviewEntry, iIsActivityEntry, iPermitReviewId, iPermitInspectionId, iReviewStatusId, iInspectionStatusId )
MakeAPermitLogEntry iPermitId, "'Permit Invoice Created'", "'Permit Invoice " & iInvoiceId & " Created'", "NULL", "NULL", iPermitStatusId, 0, 0, 1, "NULL", "NULL", "NULL", "NULL"

response.write iInvoiceId

%>