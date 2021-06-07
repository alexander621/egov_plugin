<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitformulafeetypeupdate.asp
' AUTHOR: Steve Loar
' CREATED: 01/11/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the permit formula fee types
'
' MODIFICATION HISTORY
' 1.0   01/11/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeTypeid, sSql, x, sPermitFeePrefix, iPermitFeeMethodid, iAccountid, isBuildingPermitFee
Dim isUpfrontFee, isReinspectionFee, sSuccessMsg, sUpFrontAmount, iFeeReportingTypeId

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

iPermitFeeMethodid = request("permitfeemethodid")

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
	sSql = sSql & "permitfeemethodid, atleastqty, notmorethanqty, baseamount, unitqty, unitamount, upfrontamount, feereportingtypeid "
	sSql = sSql & " ) VALUES ( "
	sSql = sSql & session("orgid") & ", '" & dbsafe(request("permitfee")) & "', " & sPermitFeePrefix & ", " & request("minimumamount")
	sSql = sSql & ", " & iAccountid & ", " & request("permitfeecategorytypeid") & ", " & iPermitFeeMethodid
	sSql = sSql & ", 0, 999999999, " & request("baseamount") & ", " & request("unitqty") & "," & request("unitamount") & ", "
	sSql = sSql & sUpFrontAmount & ", " & iFeeReportingTypeId & " )"
	iPermitFeeTypeid = RunIdentityInsert( sSql ) 
	sSuccessMsg = "Formula Fee Created"
Else 
	sSql = "UPDATE egov_permitfeetypes SET permitfee = '" & dbsafe(request("permitfee")) & "', permitfeeprefix = " & sPermitFeePrefix
	sSql = sSql & ", minimumamount = " & request("minimumamount") & ", accountid = " & iAccountid & ", permitfeecategorytypeid = " & request("permitfeecategorytypeid")
	sSql = sSql & ", permitfeemethodid = " & iPermitFeeMethodid & ", baseamount = " & request("baseamount")
	sSql = sSql & ", unitqty = " & request("unitqty")
	sSql = sSql & ", unitamount = " & request("unitamount")
	sSql = sSql & ", upfrontamount = " & sUpFrontAmount
	sSql = sSql & ", feereportingtypeid = " & iFeeReportingTypeId 
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND permitfeetypeid = " & iPermitFeeTypeid
	RunSQL sSql 
	sSuccessMsg = "Changes Saved"
End If 

' Delete any existing fixture rows
sSql = "DELETE FROM egov_permitfeetypes_to_feemultipliertypes WHERE permitfeetypeid = " & iPermitFeeTypeid 
RunSQL sSql

' Handle the included fixtures
For Each Item In request("feemultipliertypeid")
	x = x + 1
	sSql = "INSERT INTO egov_permitfeetypes_to_feemultipliertypes ( permitfeetypeid, feemultipliertypeid, displayorder ) VALUES ( "
	sSql = sSql & iPermitFeeTypeid & ", " & Item & ", " & x & " )"
	RunSQL sSql
Next 

response.redirect "permitformulafeetypeedit.asp?permitfeetypeid=" & iPermitFeeTypeid & "&success=" & sSuccessMsg



%>