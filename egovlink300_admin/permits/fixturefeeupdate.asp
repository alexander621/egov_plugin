<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: fixturefeeupdate.asp
' AUTHOR: Steve Loar
' CREATED: 04/29/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This updates the permit fixture fees. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   04/29/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeId, sSql, oRs, iPermitId, iMaxFeeCount, x, sFeesTotal, iPermitFixtureId, iQty, sFeeAmount
Dim sIsIncluded

iPermitFeeId = CLng(request("permitfeeid"))
iPermitId = CLng(request("permitid"))
iMaxFeeCount = CLng(request("maxfeecount"))
sFeesTotal = CDbl(0.00)

For x = 1 To iMaxFeeCount
	iPermitFixtureId = request("permitfixtureid" & x)
	iQty = request("qty" & x)

	sFeeAmount = GetFixtureFeeAmount( iPermitFixtureId, iQty )  ' In permitcommonfunctions.asp

	If LCase(request("include" & x)) = "true" Then 
		sFeesTotal = CDbl(sFeesTotal) + CDbl(sFeeAmount)
		sIsIncluded = 1
	Else
		sIsIncluded = 0
	End If 

	' Update the fixture row with the qty, fee amount, and is included flag
	sSql = "UPDATE egov_permitfixtures SET qty = " & iQty & ", isincluded = " & sIsIncluded & ", feeamount = " & sFeeAmount
	sSql = sSql & " WHERE permitfixtureid = " & iPermitFixtureId
	'response.write sSql & "<br /><br />"
	RunSQL sSql  ' In permitcommonfunctions.asp
Next 

' Update the fee row with the total fee
sSql = "UPDATE egov_permitfees SET feeamount = " & FormatNumber(sFeesTotal,2,,,0) & " WHERE permitfeeid = " & iPermitFeeId
'response.write sSql & "<br /><br />"
RunSQL sSql  ' In permitcommonfunctions.asp

response.write FormatNumber(sFeesTotal,2,,,0)

%>