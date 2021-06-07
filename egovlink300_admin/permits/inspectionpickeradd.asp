<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: inspectionpickeradd.asp
' AUTHOR: Steve Loar
' CREATED: 07/09/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Adds fees to permits
'
' MODIFICATION HISTORY
' 1.0   07/09/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, iPermitInspectionTypeId, sSql, iInspectionStatusId, iPermitInspectionId, iInspectionOrder, iPermitTypeId
Dim sPermitInspectionType, sInspectionDescription

iPermitId = CLng(request("permitid"))
iPermitInspectionTypeId = CLng(request("permitinspectiontypeid"))

sPermitInspectionType = ""
sInspectionDescription = ""

' Get the initial review status id for this org
iInspectionStatusId = GetInspectionStatusId( "isinitialstatus" )  '  in permitcommonfunctions.asp

' Get the next Inspection Order for this permit
iInspectionOrder = GetNextInspectionOrder( iPermitId )  '  in permitcommonfunctions.asp
'response.write "iInspectionOrder = " & iInspectionOrder & "<br /><br />"

' Get the permittypeid for this permit
iPermitTypeId = GetPermitTypeId( iPermitId )  '  in permitcommonfunctions.asp

' Get the info for this inspection type 
GetInspectionTypeDetails iPermitInspectionTypeId, sPermitInspectionType, sInspectionDescription

' Do the insert of the new inspection for this permit
sSql = "INSERT INTO egov_permitinspections ( orgid, permitid, permittypeid, permitinspectiontypeid, permitinspectiontype, "
sSql = sSql & " inspectiondescription, inspectionstatusid, inspectionorder, isincluded, routeorder ) VALUES ( "
sSql = sSql & session("orgid") & ", " & iPermitId & ", " & iPermitTypeId & ", " & iPermitInspectionTypeId & ", '"
sSql = sSql & sPermitInspectionType & "', '" & DBSafe(sInspectionDescription) & "', " & iInspectionStatusId & ", " & iInspectionOrder
sSql = sSql & ", 1, 999 )"

'response.write sSql & "<br />"
'response.flush

iPermitInspectionId = RunIdentityInsert( sSql )

response.write iPermitInspectionId



'-------------------------------------------------------------------------------------------------
' Sub GetInspectionTypeDetails( iPermitInspectionTypeId, sPermitInspectionType, sInspectionDescription )
'-------------------------------------------------------------------------------------------------
Sub GetInspectionTypeDetails( ByVal iPermitInspectionTypeId, ByRef sPermitInspectionType, ByRef sInspectionDescription )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(permitinspectiontype,'') AS permitinspectiontype, ISNULL(inspectiondescription,'') AS inspectiondescription "
	sSql = sSql & " FROM egov_permitinspectiontypes WHERE permitinspectiontypeid = " & iPermitInspectionTypeId 
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sPermitInspectionType = oRs("permitinspectiontype")
		sInspectionDescription = oRs("inspectiondescription")
	Else 
		sPermitInspectionType = ""
		sInspectionDescription = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 



%>
