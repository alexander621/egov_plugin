<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: invoicedateupdate.asp
' AUTHOR: Steve Loar
' CREATED: 08/20/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This updates permit invoice dates. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   08/20/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, iInvoiceid, sInvoiceDate, sSql, sOriginalInvoiceDate, iPermitStatusId, sActivityComment

iPermitId = CLng(request("permitid"))
iInvoiceid = CLng(request("invoiceid"))
sInvoiceDate = dbsafe(request("invoicedate"))
sOriginalInvoiceDate = request("originalinvoicedate")

sSql = "UPDATE egov_permitinvoices SET invoicedate = '" & sInvoiceDate & "' WHERE invoiceid = " & iInvoiceid 
sSql = sSql & " AND permitid = " & iPermitId & " AND orgid = " & session("orgid")
RunSQL sSql

iPermitStatusId = GetPermitStatusId( iPermitId )
sActivityComment = "'The date of invoice " & iInvoiceid & " was changed from " & sOriginalInvoiceDate & " to " & sInvoiceDate & "'"

' Put an entry in the log for the change
MakeAPermitLogEntry iPermitid, "'Date Change'", sActivityComment, "NULL", "NULL", iPermitStatusId, 0, 0, 1, "NULL", "NULL", "NULL", "NULL" 

response.write "Success"


%>