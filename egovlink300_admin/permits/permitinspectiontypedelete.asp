<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitinspectiontypedelete.asp.asp
' AUTHOR: Steve Loar
' CREATED: 01/15/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This deletes the permit inspection types
'
' MODIFICATION HISTORY
' 1.0   01/15/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitInspectionTypeid, sSql

iPermitInspectionTypeid = CLng(request("permitinspectiontypeid") )

If PermitInspectionTypeExists( iPermitInspectionTypeid ) Then 
	' Clear out the permit types to inspection types entry
	sSql = "DELETE FROM egov_permittypes_to_permitinspectiontypes WHERE permitinspectiontypeid = " & iPermitInspectionTypeid 
	RunSQL sSql
	' Clear out the inspection type entry
	sSql = "DELETE FROM egov_permitinspectiontypes WHERE permitinspectiontypeid = " & iPermitInspectionTypeid & " AND orgid = " & session("orgid")
	RunSQL sSql
End If 

response.redirect "permitinspectiontypelist.asp"

'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Function PermitInspectionTypeExists( iPermitInspectionTypeid )
'-------------------------------------------------------------------------------------------------
Function PermitInspectionTypeExists( iPermitInspectionTypeid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitinspectiontypeid) AS hits FROM egov_permitinspectiontypes "
	sSql = sSql & " WHERE permitinspectiontypeid = " & iPermitInspectionTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			PermitInspectionTypeExists = True 
		Else
			PermitInspectionTypeExists = False 
		End If 
	Else
		PermitInspectionTypeExists = False 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


%>
