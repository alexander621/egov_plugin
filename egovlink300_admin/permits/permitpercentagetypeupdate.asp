<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitpercentagetypeupdate.asp
' AUTHOR: Steve Loar
' CREATED: 09/09/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the permit percentage fee types
'
' MODIFICATION HISTORY
' 1.0   09/09/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeTypeid, sSql, x, sPermitFeePrefix, iPermitFeeMethodid, iAccountid
Dim sSuccessMsg, iPermitFeeCategoryTypeId, iMinimumAmount, iFeeReportingTypeId

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

' A value of -1 means all fees, not just one category
iPermitFeeCategoryTypeId = CLng(request("permitfeecategorytypeid"))

If request("minimumamount") <> "" Then 
	iMinimumAmount = CDbl(request("minimumamount"))
Else
	iMinimumAmount = 0.00
End If 

If CLng(request("feereportingtypeid")) = CLng(0) Then 
	iFeeReportingTypeId = "NULL"
Else
	iFeeReportingTypeId = CLng(request("feereportingtypeid"))
End If 

If iPermitFeeTypeid = CLng(0) Then 
	sSql = "INSERT INTO egov_permitfeetypes ( orgid, permitfee, permitfeeprefix, minimumamount, accountid, permitfeecategorytypeid, "
	sSql = sSql & " ispercentagetypefee, permitfeemethodid, atleastqty, notmorethanqty, "
	sSql = sSql & " baseamount, unitqty, unitamount, percentage, feereportingtypeid ) VALUES ( "
	sSql = sSql & session("orgid") & ", '" & dbsafe(request("permitfee")) & "', " & sPermitFeePrefix & ", " & iMinimumAmount
	sSql = sSql & ", " & iAccountid & ", " & iPermitFeeCategoryTypeId & ", 1, " & iPermitFeeMethodid & ", "
	sSql = sSql & "0, 999999999, NULL, NULL, NULL, " & CDbl(request("percentage")) & ", " & iFeeReportingTypeId & " )"

	iPermitFeeTypeid = RunIdentityInsert( sSql ) 

	sSuccessMsg = "Percentage Fee Created"
Else 
	sSql = "UPDATE egov_permitfeetypes SET permitfee = '" & dbsafe(request("permitfee")) & "', permitfeeprefix = " & sPermitFeePrefix
	sSql = sSql & ", minimumamount = " & iMinimumAmount
	sSql = sSql & ", accountid = " & iAccountid
	sSql = sSql & ", permitfeecategorytypeid = " & iPermitFeeCategoryTypeId
	sSql = sSql & ", permitfeemethodid = " & iPermitFeeMethodid
	sSql = sSql & ", percentage =  " & CDbl(request("percentage"))
	sSql = sSql & ", feereportingtypeid = " & iFeeReportingTypeId 
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND permitfeetypeid = " & iPermitFeeTypeid

	RunSQL sSql 

	sSuccessMsg = "Changes Saved"
End If 


response.redirect "permitpercentagetypefeeedit.asp?permitfeetypeid=" & iPermitFeeTypeid & "&success=" & sSuccessMsg



%>