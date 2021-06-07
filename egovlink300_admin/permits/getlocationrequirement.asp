<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getlocationrequirement.asp
' AUTHOR: Steve Loar
' CREATED: 10/29/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This gets the location requirement for a permit type. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   10/29/2010   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitTypeId, sSql, oRs, sResponse

iPermitTypeId = CLng(request("permittypeid"))

If iPermitTypeId > CLng(0) Then 

	sSql = "SELECT R.locationtype FROM egov_permittypes P, egov_permitlocationrequirements R "
	sSql = sSql & "WHERE P.permitlocationrequirementid = R.permitlocationrequirementid "
	sSql = sSql & "AND P.permittypeid = " & iPermitTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sResponse = oRs("locationtype")
	Else
		sResponse = "none"
	End If 

	oRs.Close
	Set oRs = Nothing 

Else
	sResponse = "introtext"
End If 


response.write sResponse



%>
