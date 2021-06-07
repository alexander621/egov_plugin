<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: updatefacilityPTG.asp
' AUTHOR: Steve Loar
' CREATED: 8/12/2013
' COPYRIGHT: COPYRIGHT 2013 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  Update the price type group id for a facility. Called via AJAX.
'
' MODIFICATION HISTORY
' 1.0   8/12/2013	Steve Loar - INITIAL VERSION
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFacilityId, iPriceTypeGroupId, sSql

If request("facilityid") <> "" Then
	iFacilityId = CLng(request("facilityid"))
Else 
	iFacilityId = 0
End If 

If request("orgid") <> "" Then
	iOrgId = CLng(request("orgid"))
Else 
	iOrgId = 0
End If 

If request("pricetypegroupid") <> "" Then
	iPriceTypeGroupId = CLng(request("pricetypegroupid"))
Else 
	iPriceTypeGroupId = 1
End If 

sSql = "UPDATE egov_facility SET pricetypegroupid = " & iPriceTypeGroupId
sSql = sSql & " WHERE facilityid = " & iFacilityId & " AND orgid = " & iOrgId

RunSQLStatement sSql

response.write "Success " '& sSql


%>

