<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: checkconstructionrate.asp
' AUTHOR: Steve Loar
' CREATED: 03/11/08
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This gets the construction rate value for display. It is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   03/11/08	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, sResults, iConstructionTypeId, iOccupancyTypeId

sResults = ""
iConstructionTypeId = CLng(request("constructiontypeid"))
iOccupancyTypeId = CLng(request("occupancytypeid"))

sSql = "SELECT constructiontyperate, isnotpermitted FROM egov_constructionfactors "
sSql = sSql & " WHERE constructiontypeid = " & iConstructionTypeId
sSql = sSql & " AND occupancytypeid = " & iOccupancyTypeId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If Not oRs.EOF Then 
	If oRs("isnotpermitted") Then
		sResults = "not permitted with this combination."
	Else 
		sResults = oRs("constructiontyperate")
	End If 
Else
	sResults = "not found."
End If 

oRs.Close
Set oRs = Nothing 

response.write sResults


%>
