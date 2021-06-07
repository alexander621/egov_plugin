<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitvaluationtypeupdate.asp
' AUTHOR: Steve Loar
' CREATED: 04/14/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the permit valuation types
'
' MODIFICATION HISTORY
' 1.0   04/14/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitValuationTypeid, sSql, x, iUseStepTable, iUnitQty, iUnitFeeAmount, sSuccessMsg

iPermitValuationTypeid = CLng(request("permitvaluationtypeid") )

If iPermitValuationTypeid = CLng(0) Then 
	sSql = "INSERT INTO egov_permitvaluationtypes ( orgid, permitvaluation ) VALUES ( " & session("orgid") & ", '" & dbsafe(request("permitvaluation")) & "' )"
	iPermitValuationTypeid = RunIdentityInsert( sSql ) 
	sSuccessMsg = "Permit Valuation Type Created"
Else 
	sSql = "UPDATE egov_permitvaluationtypes SET permitvaluation = '" & dbsafe(request("permitvaluation")) & "' WHERE orgid = " & session("orgid") & " AND permitvaluationtypeid = " & iPermitValuationTypeid
	RunSQL sSql 
	sSuccessMsg = "Changes Saved"
End If 

' Delete any existing step table rows
sSql = "DELETE FROM egov_permitvaluationtypestepfees WHERE permitvaluationtypeid = " & iPermitValuationTypeid
RunSQL sSql

' Insert the step fee rows
For x = 1 To CLng(request("maxrows"))
	If request("atleastvalue" & x) <> "" Then 
		sSql = "INSERT INTO egov_permitvaluationtypestepfees (permitvaluationtypeid, atleastvalue, notmorethanvalue, baseamount, unitqty, unitamount) VALUES ( "
		sSql = sSql & iPermitValuationTypeid & ", " & request("atleastvalue" & x) & ", " & request("notmorethanvalue" & x) & ", " & request("baseamount" & x) & ", " & request("unitqty" & x) & ", " & request("unitamount" & x) & " )"
		RunSQL sSql
	End If 
Next 

response.redirect "permitvaluationtypeedit.asp?permitvaluationtypeid=" & iPermitValuationTypeid & "&success=" & sSuccessMsg


%>