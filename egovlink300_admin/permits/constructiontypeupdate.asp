<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: constructiontypeupdate.asp
' AUTHOR: Steve Loar
' CREATED: 12/12/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This updates the construction rates
'
' MODIFICATION HISTORY
' 1.0   12/12/2007   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iOccupancyTypeId, sSql, x, iIsNotPermitted, iRate, sSuccessMsg

iOccupancyTypeId = CLng(request("occupancytypeid") )

If iOccupancyTypeId = CLng(0) Then 
	sSql = "INSERT INTO egov_occupancytypes ( orgid, occupancytype, usegroupcode ) VALUES ( " & session("orgid") & ", '" & dbsafe(request("occupancytype")) & "', '" & dbsafe(request("usegroupcode")) & "' )"
	iOccupancyTypeId = RunIdentityInsert( sSql )
	sSuccessMsg = "Construction Type Rates Created"
	' Loop through and do inserts
	For x = CLng(request("min")) To CLng(request("max"))
		If request("rate" & x) <> "" Then 
			If Trim(request("rate" & x)) = "NP" Then 
				iIsNotPermitted = 1
				iRate = 0.00
			Else
				iIsNotPermitted = 0
				iRate = request("rate" & x)
			End If 
			sSql = "INSERT INTO egov_constructionfactors ( occupancytypeid, constructiontypeid, constructiontyperate, isnotpermitted ) VALUES ( " & iOccupancyTypeId & ", " & x & ", " & iRate & ", " & iIsNotPermitted & " )"
			RunSQL sSql
		End If 
	Next 
Else 
	sSql = "UPDATE egov_occupancytypes SET occupancytype = '" & dbsafe(request("occupancytype")) & "', usegroupcode = '" & dbsafe(request("usegroupcode")) & "' WHERE orgid = " & session("orgid") & " AND occupancytypeid = " & iOccupancyTypeId
	RunSQL sSql 	
	sSuccessMsg = "Changes Saved"
	' Loop through and do updates
	For x = CLng(request("min")) To CLng(request("max"))
		If request("rate" & x) <> "" Then 
			If Trim(request("rate" & x)) = "NP" Then 
				iIsNotPermitted = 1
				iRate = 0.00
			Else
				iIsNotPermitted = 0
				iRate = request("rate" & x)
			End If 
			sSql = "UPDATE egov_constructionfactors SET constructiontyperate = " & iRate & ", isnotpermitted = " & iIsNotPermitted & " WHERE occupancytypeid = " & iOccupancyTypeId & " AND constructiontypeid = " & x 
			RunSQL sSql
		End If
	Next 
End If 

response.redirect "constructiontypeedit.asp?occupancytypeid=" & iOccupancyTypeId & "&success=" & sSuccessMsg


%>