<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitfixturetypeupdate.asp
' AUTHOR: Steve Loar
' CREATED: 12/19/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the permit fixture types
'
' MODIFICATION HISTORY
' 1.0   12/19/2007   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFixtureTypeid, sSql, x, iUseStepTable, iUnitQty, iUnitFeeAmount, sSuccessMsg

iPermitFixtureTypeid = CLng(request("permitfixturetypeid") )

If iPermitFixtureTypeid = CLng(0) Then 
	sSql = "INSERT INTO egov_permitfixturetypes ( orgid, permitfixture, displayorder ) VALUES ( " & session("orgid") & ", '" & dbsafe(request("permitfixture")) & "', 9999 )"
	iPermitFixtureTypeid = RunIdentityInsert( sSql ) 
	sSuccessMsg = "Permit Fixture Type Created"
Else 
	sSql = "UPDATE egov_permitfixturetypes SET permitfixture = '" & dbsafe(request("permitfixture")) & "' WHERE orgid = " & session("orgid") & " AND permitfixturetypeid = " & iPermitFixtureTypeid
	RunSQL sSql 
	sSuccessMsg = "Changes Saved"
End If 

' Delete any existing step table rows
sSql = "DELETE FROM egov_permitfixturetypestepfees WHERE permitfixturetypeid = " & iPermitFixtureTypeid
RunSQL sSql

' If there is a step table being used then add the current rows
For x = 1 To CLng(request("maxrows"))
	If request("atleastqty" & x) <> "" Then 
		sSql = "INSERT INTO egov_permitfixturetypestepfees (permitfixturetypeid, atleastqty, notmorethanqty, baseamount, unitqty, unitamount) VALUES ( "
		sSql = sSql & iPermitFixtureTypeid & ", " & request("atleastqty" & x) & ", " & request("notmorethanqty" & x) & ", " & request("baseamount" & x) & ", " & request("unitqty" & x) & ", " & request("unitamount" & x) & " )"
		RunSQL sSql
	End If 
Next 

response.redirect "permitfixturetypeedit.asp?permitfixturetypeid=" & iPermitFixtureTypeid & "&success=" & sSuccessMsg


%>