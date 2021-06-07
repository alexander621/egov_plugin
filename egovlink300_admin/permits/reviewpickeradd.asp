<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: reviewpickeradd.asp
' AUTHOR: Steve Loar
' CREATED: 06/19/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Adds fees to permits
'
' MODIFICATION HISTORY
' 1.0   06/19/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, iPermitReviewTypeId, sSql, iReviewStatusId, iPermitReviewId, iReviewOrder, iPermitTypeId
Dim sPermitReviewType, sReviewDescription

iPermitId = CLng(request("permitid"))
iPermitReviewTypeId = CLng(request("permitreviewtypeid"))

sPermitReviewType = ""
sReviewDescription = ""
iReviewerId = 0

' Get the initial review status id for this org
iReviewStatusId = GetReviewStatusId( "isinitialstatus" )  '  in permitcommonfunctions.asp

' Get the next Review Order for this permit
iReviewOrder = GetNextReviewOrder( iPermitId )

' Get the permittypeid for this permit
iPermitTypeId = GetPermitTypeId( iPermitId )  '  in permitcommonfunctions.asp

' Get the info for this review type from egov_permitreviewtypes, egov_permittypes_to_permitreviewtypes
GetReviewTypeDetails iPermitReviewTypeId, sPermitReviewType, sReviewDescription

' Do the insert of the new review for this permit
sSql = "INSERT INTO egov_permitreviews ( orgid, permitid, permittypeid, permitreviewtypeid, "
sSql = sSql & " permitreviewtype, reviewdescription, reviewstatusid, revieworder, isincluded ) VALUES ( "
sSql = sSql & session("orgid") & ", " & iPermitId & ", " & iPermitTypeId & ", " & iPermitReviewTypeId & ", '"
sSql = sSql & sPermitReviewType & "', '" & sReviewDescription & "', " & iReviewStatusId & ", " & iReviewOrder
sSql = sSql & ", 1 )"
iPermitReviewId = RunIdentityInsert( sSql )

response.write iPermitReviewId



'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Sub GetReviewTypeDetails( iPermitReviewTypeId, sPermitReviewType, sReviewDescription )
'-------------------------------------------------------------------------------------------------
Sub GetReviewTypeDetails( iPermitReviewTypeId, ByRef sPermitReviewType, ByRef sReviewDescription )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(permitreviewtype,'') AS permitreviewtype, ISNULL(reviewdescription,'') AS reviewdescription "
	sSql = sSql & " FROM egov_permitreviewtypes WHERE permitreviewtypeid = " & iPermitReviewTypeId 
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sPermitReviewType = oRs("permitreviewtype")
		sReviewDescription = oRs("reviewdescription")
	Else 
		sPermitReviewType = ""
		sReviewDescription = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'-------------------------------------------------------------------------------------------------
' Function GetNextReviewOrder( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetNextReviewOrder( iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(MAX(revieworder),0) AS revieworder FROM egov_permitreviews WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetNextReviewOrder = CLng(oRs("revieworder")) + CLng(1)
	Else
		GetNextReviewOrder = CLng(1)
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 



%>
