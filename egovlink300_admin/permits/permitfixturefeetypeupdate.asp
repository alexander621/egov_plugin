<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitfixturefeetypeupdate.asp
' AUTHOR: Steve Loar
' CREATED: 01/08/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the permit fixture types
'
' MODIFICATION HISTORY
' 1.0   01/08/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeTypeid, sSql, x, sPermitFeePrefix, iPermitFeeMethodid, iAccountid, sSuccessMsg
Dim sUpFrontAmount, iFeeReportingTypeId

iPermitFeeTypeid = CLng(request("permitfeetypeid") )
x = 0

If request("permitfeeprefix") = "" Then
	sPermitFeePrefix = "NULL"
Else
	sPermitFeePrefix = "'" & dbsafe(request("permitfeeprefix")) & "'"
End If 

' Handle accountid pick not being there
If request("accountid") = "" Then
	iAccountid = "NULL"
Else
	If CLng(request("accountid")) = CLng(0) Then 
		iAccountid = "NULL"
	Else
		iAccountid = CLng(request("accountid"))
	End If 
End If 

iPermitFeeMethodid = GetFixtureFeeMethod( session("orgid") )  ' In common.asp

If request("upfrontamount") <> "" Then 
	sUpFrontAmount = CDbl(request("upfrontamount"))
Else
	sUpFrontAmount = "0.00"
End If 

If CLng(request("feereportingtypeid")) = CLng(0) Then 
	iFeeReportingTypeId = "NULL"
Else
	iFeeReportingTypeId = CLng(request("feereportingtypeid"))
End If 

If iPermitFeeTypeid = CLng(0) Then 
	sSql = "INSERT INTO egov_permitfeetypes ( orgid, permitfee, permitfeeprefix, minimumamount, accountid, permitfeecategorytypeid, "
	sSql = sSql & " isfixturetypefee, permitfeemethodid, upfrontamount, feereportingtypeid ) VALUES ( "
	sSql = sSql & session("orgid") & ", '" & dbsafe(request("permitfee")) & "', " & sPermitFeePrefix & ", " & request("minimumamount")
	sSql = sSql & ", " & iAccountid & ", " & request("permitfeecategorytypeid") & ", 1, " & iPermitFeeMethodid 
	sSql = sSql & ", " & sUpFrontAmount & ", " & iFeeReportingTypeId & " )"
	iPermitFeeTypeid = RunIdentityInsert( sSql ) 
	sSuccessMsg = "Fixture Fee Created"
Else 
	sSql = "UPDATE egov_permitfeetypes SET permitfee = '" & dbsafe(request("permitfee")) & "', permitfeeprefix = " & sPermitFeePrefix
	sSql = sSql & ", minimumamount = " & request("minimumamount")
	sSql = sSql & ", accountid = " & iAccountid
	sSql = sSql & ", permitfeecategorytypeid = " & request("permitfeecategorytypeid")
	sSql = sSql & ", upfrontamount = " & sUpFrontAmount
	sSql = sSql & ", feereportingtypeid = " & iFeeReportingTypeId 
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND permitfeetypeid = " & iPermitFeeTypeid
	RunSQL sSql 
	sSuccessMsg = "Changes Saved"
End If 

' Delete any existing fixture rows
sSql = "DELETE FROM egov_permitfeetypes_to_permitfixturetypes WHERE permitfeetypeid = " & iPermitFeeTypeid 
RunSQL sSql

' Handle the included fixtures
For Each Item In request("permitfixturetypeid")
	x = x + 1
	iDisplayOrder = GetFixtureTypeDisplayOrder( Item )
	sSql = "INSERT INTO egov_permitfeetypes_to_permitfixturetypes ( permitfeetypeid, permitfixturetypeid, displayorder ) VALUES ( "
	sSql = sSql & iPermitFeeTypeid & ", " & Item & ", " & iDisplayOrder & " )"
	RunSQL sSql
Next 

response.redirect "permitfixturefeetypeedit.asp?permitfeetypeid=" & iPermitFeeTypeid & "&success=" & sSuccessMsg


%>