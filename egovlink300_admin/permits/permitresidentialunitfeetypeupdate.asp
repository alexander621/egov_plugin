<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitresidentialunitfeetypeupdate.asp
' AUTHOR: Steve Loar
' CREATED: 11/03/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the permit Residential Unit fee types
'
' MODIFICATION HISTORY
' 1.0   11/03/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeTypeid, sSql, x, sPermitFeePrefix, iPermitFeeMethodid, iAccountid, sSuccessMsg
Dim iFeeReportingTypeId

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

iPermitFeeMethodid = GetResidentialUnitFeeMethodId( session("orgid") )  ' In common.asp

If CLng(request("feereportingtypeid")) = CLng(0) Then 
	iFeeReportingTypeId = "NULL"
Else
	iFeeReportingTypeId = CLng(request("feereportingtypeid"))
End If 

If iPermitFeeTypeid = CLng(0) Then 
	sSql = "INSERT INTO egov_permitfeetypes ( orgid, permitfee, permitfeeprefix, minimumamount, accountid, permitfeecategorytypeid, "
	sSql = sSql & " isresidentialunittypefee, permitfeemethodid, feereportingtypeid ) VALUES ( "
	sSql = sSql & session("orgid") & ", '" & dbsafe(request("permitfee")) & "', " & sPermitFeePrefix & ", " & request("minimumamount")
	sSql = sSql & ", " & iAccountid & ", " & request("permitfeecategorytypeid") & ", 1, " & iPermitFeeMethodid & ", "
	sSql = sSql & iFeeReportingTypeId & " )"
	iPermitFeeTypeid = RunIdentityInsert( sSql ) 
	sSuccessMsg = "Residential Unit Fee Type Created"
Else 
	sSql = "UPDATE egov_permitfeetypes SET permitfee = '" & dbsafe(request("permitfee")) & "'"
	sSql = sSql & ", permitfeeprefix = " & sPermitFeePrefix
	sSql = sSql & ", minimumamount = " & request("minimumamount")
	sSql = sSql & ", accountid = " & iAccountid
	sSql = sSql & ", permitfeecategorytypeid = " & request("permitfeecategorytypeid")
	sSql = sSql & ", feereportingtypeid = " & iFeeReportingTypeId 
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND permitfeetypeid = " & iPermitFeeTypeid
	RunSQL sSql 
	sSuccessMsg = "Changes Saved"
End If 

' Delete any existing step table rows
sSql = "DELETE FROM egov_permitresidentialunittypestepfees WHERE permitfeetypeid = " & iPermitFeeTypeid
RunSQL sSql

' If there is a step table being used then add the current rows
For x = 1 To CLng(request("maxrows"))
	If request("atleastqty" & x) <> "" Then 
		sSql = "INSERT INTO egov_permitresidentialunittypestepfees ( permitfeetypeid, atleastqty, notmorethanqty, baseamount, unitamount ) VALUES ( "
		sSql = sSql & iPermitFeeTypeid & ", " & request("atleastqty" & x) & ", " & request("notmorethanqty" & x) & ", " & request("baseamount" & x) & ", " & request("unitamount" & x) & " )"
		RunSQL sSql
	End If 
Next 

response.redirect "permitresidentalunitfeetypeedit.asp?permitfeetypeid=" & iPermitFeeTypeid & "&success=" & sSuccessMsg



%>