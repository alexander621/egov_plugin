<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: fixturetypeadd.asp
' AUTHOR: Steve Loar
' CREATED: 05/12/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Adds fees to permits
'
' MODIFICATION HISTORY
' 1.0   05/12/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, iPermitFixtureTypeId, iPermitFeeId, sSql, iNextFixtureCount, sFixtureName

iPermitId = CLng(request("permitid"))

iPermitFeeId = Clng(request("permitfeeid"))

iPermitFixtureTypeId = CLng(request("permitfixturetypeid"))

iNextFixtureCount = GetNextFixtureCount( iPermitFeeId ) 

sFixtureName = GetFixtureName( iPermitFixtureTypeId )


sSql = "INSERT INTO egov_permitfixtures ( permitid, permitfeeid, permitfixturetypeid, orgid, permitfixture, displayorder ) VALUES ( "
sSql = sSql & iPermitId & ", " & iPermitFeeId & ", " & iPermitFixtureTypeId & ", " & session("orgid") & ", '"
sSql = sSql & dbsafe(sFixtureName) & "', " & iNextFixtureCount & " )"

'response.write sSql & "<br />"

iPermitFixtureId = RunIdentityInsert( sSql )

' Input the step table entries 
CreateFixtureStepFees iPermitFixtureId, iPermitId, iPermitFeeId, iPermitFixtureTypeId  ' In permitcommonfunctions.asp

response.write "SUCCESS"


'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Function GetNextFixtureCount( iPermitFeeId ) 
'-------------------------------------------------------------------------------------------------
Function GetNextFixtureCount( iPermitFeeId ) 
	Dim sSql, oRs

	sSql = "SELECT ISNULL(MAX(displayorder),0) AS displayorder "
	sSql = sSql & " FROM egov_permitfixtures "
	sSql = sSql & " WHERE permitfeeid = " & iPermitFeeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetNextFixtureCount = CLng(oRs("displayorder")) + CLng(1)
	Else
		GetNextFixtureCount = CLng(1)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' Function GetFixtureName( iPermitFixtureTypeId )
'-------------------------------------------------------------------------------------------------
Function GetFixtureName( iPermitFixtureTypeId )
	Dim sSql, oRs

	sSql = "SELECT permitfixture FROM egov_permitfixturetypes WHERE permitfixturetypeid = " & iPermitFixtureTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetFixtureName = oRs("permitfixture")
	Else
		GetFixtureName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 



%>
