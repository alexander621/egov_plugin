<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitresidentialunittypeupdate.asp
' AUTHOR: Steve Loar
' CREATED: 10/30/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the residential unit types
'
' MODIFICATION HISTORY
' 1.0   10/30/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iResidentialUnitTypeid, sSql, x, iUnitFeeAmount, sSuccessMsg

iResidentialUnitTypeid = CLng(request("residentialunittypeid") )

If iResidentialUnitTypeid = CLng(0) Then 
	sSql = "INSERT INTO egov_permitresidentialunittypes ( orgid, residentialunittype ) VALUES ( " & session("orgid") & ", '" & dbsafe(request("residentialunittype")) & "' )"
	iResidentialUnitTypeid = RunIdentityInsert( sSql ) 
	sSuccessMsg = "Residential Unit Type Created"
Else 
	sSql = "UPDATE egov_permitresidentialunittypes SET residentialunittype = '" & dbsafe(request("residentialunittype")) & "' WHERE orgid = " & session("orgid") & " AND residentialunittypeid = " & iResidentialUnitTypeid
	RunSQL sSql 
	sSuccessMsg = "Changes Saved"
End If 

' Delete any existing step table rows
sSql = "DELETE FROM egov_permitresidentialunittypestepfees WHERE residentialunittypeid = " & iResidentialUnitTypeid
RunSQL sSql

' If there is a step table being used then add the current rows
For x = 1 To CLng(request("maxrows"))
	If request("atleastqty" & x) <> "" Then 
		sSql = "INSERT INTO egov_permitresidentialunittypestepfees (residentialunittypeid, atleastqty, notmorethanqty, baseamount, unitamount) VALUES ( "
		sSql = sSql & iResidentialUnitTypeid & ", " & request("atleastqty" & x) & ", " & request("notmorethanqty" & x) & ", " & request("baseamount" & x) & ", " & request("unitamount" & x) & " )"
		RunSQL sSql
	End If 
Next 

response.redirect "permitresidentialunittypeedit.asp?residentialunittypeid=" & iResidentialUnitTypeid & "&success=" & sSuccessMsg


%>