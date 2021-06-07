<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: paiddateupdate.asp
' AUTHOR: Steve Loar
' CREATED: 08/20/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This updates permit invoice paid dates. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   08/20/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPaymentId, iPermitId, iInvoiceid, sPaymentDate, sSql, sOriginalPaymentDate, iPermitStatusId, sActivityComment

iPaymentId = CLng(request("paymentid"))
iPermitId = CLng(request("permitid"))
iInvoiceid = CLng(request("invoiceid"))
sPaymentDate = dbsafe(request("paymentdate"))
sOriginalPaymentDate = request("originalpaymentdate")

sSql = "UPDATE egov_class_payment SET paymentdate = '" & sPaymentDate & "' WHERE paymentid = " & iPaymentId 
sSql = sSql & " AND orgid = " & session("orgid")
RunSQL sSql

iPermitStatusId = GetPermitStatusId( iPermitId )
sActivityComment = "'The paid date of invoice " & iInvoiceid & " was changed from " & sOriginalPaymentDate & " to " & sPaymentDate & "'"

' Put an entry in the log for the change
MakeAPermitLogEntry iPermitid, "'Date Change'", sActivityComment, "NULL", "NULL", iPermitStatusId, 0, 0, 1, "NULL", "NULL", "NULL", "NULL" 

response.write "Success"


%>