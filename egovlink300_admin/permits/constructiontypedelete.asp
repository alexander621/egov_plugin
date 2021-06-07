<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: constructiontypedelete.asp
' AUTHOR: Steve Loar
' CREATED: 12/13/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This deletes the construction rates
'
' MODIFICATION HISTORY
' 1.0   12/13/2007   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iOccupancyTypeId, sSql

iOccupancyTypeId = CLng(request("oid") )

If OccupancyTypeExists( iOccupancyTypeId ) Then 
	sSql = "DELETE FROM egov_occupancytypes WHERE occupancytypeid = " & iOccupancyTypeId & " AND orgid = " & session("orgid")
	RunSQL sSql 
	sSql = "DELETE FROM egov_constructionfactors WHERE occupancytypeid = " & iOccupancyTypeId 
	RunSQL sSql 
End If 

response.redirect "constructiontypelist.asp"

'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Function OccupancyTypeExists( iOccupancyTypeId )
'-------------------------------------------------------------------------------------------------
Function OccupancyTypeExists( iOccupancyTypeId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(occupancytypeid) AS hits FROM egov_occupancytypes WHERE occupancytypeid = " & iOccupancyTypeId
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			OccupancyTypeExists = True 
		Else
			OccupancyTypeExists = False 
		End If 
	Else
		OccupancyTypeExists = False 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


%>
