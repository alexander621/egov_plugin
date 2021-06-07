<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getpermitfeetotal.asp
' AUTHOR: Steve Loar
' CREATED: 04/10/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This get the fee total for a permit. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   04/10/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sSql, oRs, sResponse

iPermitId = CLng(request("permitid"))

sResponse = SetPermitFeeTotal( iPermitId ) ' In permitcommonfunctions.asp

'sSql = "SELECT SUM(feeamount) AS totalfee FROM egov_permitfees WHERE includefee = 1 AND permitid = " & iPermitId 

'Set oRs = Server.CreateObject("ADODB.Recordset")
'oRs.Open sSQL, Application("DSN"), 3, 1

'If Not oRs.EOF Then
'	sResponse = FormatNumber(oRs("totalfee"),2,,,0) 
'	sSql = "UPDATE egov_permits SET feetotal = " & oRs("totalfee") & " WHERE permitid = " & iPermitId 
'	RunSQL sSql
'Else
'	sResponse = "Failed"
'End If 


'oRs.Close
'Set oRs = Nothing 

response.write sResponse



%>
