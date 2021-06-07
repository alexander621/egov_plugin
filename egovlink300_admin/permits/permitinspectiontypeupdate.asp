<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitinspectiontypeupdate.asp
' AUTHOR: Steve Loar
' CREATED: 01/15/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the permit inspection types
'
' MODIFICATION HISTORY
' 1.0   01/15/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitInspectionTypeid, sSql, sPermitInspectionType, isBuildingPermitFee, sSuccessMsg

iPermitInspectionTypeid = CLng(request("permitinspectiontypeid") )

If request("permitinspectiontype") = "" Then
	sPermitInspectionType = "NULL"
Else
	sPermitInspectionType = "'" & dbsafe(request("permitinspectiontype")) & "'"
End If 

'If request("isbuildingpermittype") = "on" Then
'	isBuildingPermitFee = 1
'Else
'	isBuildingPermitFee = 0
'End If 

If iPermitInspectionTypeid = CLng(0) Then 
	sSql = "INSERT INTO egov_permitinspectiontypes ( orgid, inspectiondescription, permitinspectiontype "
	sSql = sSql & " ) VALUES ( " & session("orgid") & ", '" & dbsafe(request("inspectiondescription")) 
	sSql = sSql & "', " & sPermitInspectionType & " )"

	iPermitInspectionTypeid = RunIdentityInsert( sSql ) 

	sSuccessMsg = "Permit Inspection Type Created"
Else 
	sSql = "UPDATE egov_permitinspectiontypes SET inspectiondescription = '" & dbsafe(request("inspectiondescription"))
	sSql = sSql & "', permitinspectiontype = " & sPermitInspectionType
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND permitinspectiontypeid = " & iPermitInspectionTypeid

	RunSQL sSql 

	sSuccessMsg = "Changes Saved"
End If 

response.redirect "permitinspectiontypeedit.asp?permitinspectiontypeid=" & iPermitInspectionTypeid & "&success=" & sSuccessMsg



%>