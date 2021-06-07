<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitvaluationfeetypeupdate.asp
' AUTHOR: Steve Loar
' CREATED: 04/14/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the permit valuation types
'
' MODIFICATION HISTORY
' 1.0   04/14/2008   Steve Loar - INITIAL VERSION
' 2.0	11/17/2008	Steve Loar - Changed to add the valuation step fees 
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeTypeid, sSql, x, sPermitFeePrefix, iPermitFeeMethodid, iAccountid, sSuccessMsg
Dim iOnSewerFeeReport, bOnBBSFeeReport, iFeeReportingTypeId

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

iPermitFeeMethodid = GetValuationFeeMethod( session("orgid") )  ' In common.asp

If CLng(request("feereportingtypeid")) = CLng(0) Then 
	iFeeReportingTypeId = "NULL"
Else
	iFeeReportingTypeId = CLng(request("feereportingtypeid"))
End If 

If iPermitFeeTypeid = CLng(0) Then 
	sSql = "INSERT INTO egov_permitfeetypes ( orgid, permitfee, permitfeeprefix, minimumamount, accountid, permitfeecategorytypeid, "
	sSql = sSql & " isvaluationtypefee, permitfeemethodid, feereportingtypeid ) VALUES ( "
	sSql = sSql & session("orgid") & ", '" & dbsafe(request("permitfee")) & "', " & sPermitFeePrefix & ", " & request("minimumamount")
	sSql = sSql & ", " & iAccountid & ", " & request("permitfeecategorytypeid") & ", 1, " & iPermitFeeMethodid & ", "
	sSql = sSql & iFeeReportingTypeId & " )"
	iPermitFeeTypeid = RunIdentityInsert( sSql ) 
	sSuccessMsg = "Valuation Fee Created"
Else 
	sSql = "UPDATE egov_permitfeetypes SET permitfee = '" & dbsafe(request("permitfee")) & "', permitfeeprefix = " & sPermitFeePrefix
	sSql = sSql & ", minimumamount = " & request("minimumamount")
	sSql = sSql & ", accountid = " & iAccountid
	sSql = sSql & ", permitfeecategorytypeid = " & request("permitfeecategorytypeid")
	sSql = sSql & ", feereportingtypeid = " & iFeeReportingTypeId 
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND permitfeetypeid = " & iPermitFeeTypeid
	RunSQL sSql 
	sSuccessMsg = "Changes Saved"
End If 

' Delete any existing valuation rows
'sSql = "DELETE FROM egov_permitfeetypes_to_permitvaluationtypes WHERE permitfeetypeid = " & iPermitFeeTypeid 
'RunSQL sSql

' Handle the included valuations - Now there is only one per fee type
'For Each Item In request("permitvaluationtypeid")
'	x = x + 1
'	sSql = "INSERT INTO egov_permitfeetypes_to_permitvaluationtypes ( permitfeetypeid, permitvaluationtypeid, displayorder ) VALUES ( "
'	sSql = sSql & iPermitFeeTypeid & ", " & Item & ", " & x & " )"
'	RunSQL sSql
'Next 


' Delete any existing step table rows
sSql = "DELETE FROM egov_permitvaluationtypestepfees WHERE permitfeetypeid = " & iPermitFeeTypeid
RunSQL sSql

' Insert the step fee rows
For x = 1 To CLng(request("maxrows"))
	If request("atleastvalue" & x) <> "" Then 
		sSql = "INSERT INTO egov_permitvaluationtypestepfees (permitfeetypeid, atleastvalue, notmorethanvalue, baseamount, unitqty, unitamount) VALUES ( "
		sSql = sSql & iPermitFeeTypeid & ", " & request("atleastvalue" & x) & ", " & request("notmorethanvalue" & x) & ", " & request("baseamount" & x) & ", " & request("unitqty" & x) & ", " & request("unitamount" & x) & " )"
		RunSQL sSql
	End If 
Next 

response.redirect "permitvaluationfeetypeedit.asp?permitfeetypeid=" & iPermitFeeTypeid & "&success=" & sSuccessMsg



%>