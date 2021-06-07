<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: feemultiplierupdate.asp
' AUTHOR: Steve Loar
' CREATED: 12/18/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This updates the fee multiplier rates
'
' MODIFICATION HISTORY
' 1.0   12/18/2007   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFeeMultiplierTypeid, sSql, x, sSuccessMsg

iFeeMultiplierTypeid = CLng(request("feemultipliertypeid") )

If iFeeMultiplierTypeid = CLng(0) Then 
	sSql = "INSERT INTO egov_feemultipliertypes ( orgid, feemultiplier, feemultiplierrate ) VALUES ( " & session("orgid") & ", '" & dbsafe(request("feemultiplier")) & "', " & request("feemultiplierrate") & " )"
	iFeeMultiplierTypeid = RunIdentityInsert( sSql )
	sSuccessMsg = "Fee Multiplier Rate Created"
Else 
	sSql = "UPDATE egov_feemultipliertypes SET feemultiplier = '" & dbsafe(request("feemultiplier")) & "', feemultiplierrate = " & request("feemultiplierrate") & " WHERE orgid = " & session("orgid") & " AND feemultipliertypeid = " & iFeeMultiplierTypeid
	RunSQL sSql 
	sSuccessMsg = "Changes Saved"
End If 

response.redirect "feemultiplieredit.asp?feemultipliertypeid=" & iFeeMultiplierTypeid & "&success=" & sSuccessMsg & "&id=" & request("id")


%>
