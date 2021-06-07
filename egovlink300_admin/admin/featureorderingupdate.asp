<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: featureorderingupdate.asp
' AUTHOR: Steve Loar
' CREATED: 09/12/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This sets the order of a feature. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   09/12/2008	Steve Loar - INITIAL VERSION
' 1.1	04/11/2011	Steve Loar - Added mobile feature ordering
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFeatureId, iOrder, sOrderField

iFeatureId = CLng(request("featureid"))
iOrder = CLng(request("displayorder"))

If FeatureIsTopLevel( iFeatureId ) Then
	If CLng(request("admindisplayorder")) = CLng(1) Then 
		sOrderField = "admindisplayorder"
	Else
		If CLng(request("admindisplayorder")) = CLng(0) Then 
			sOrderField = "publicdisplayorder"
		Else
			sOrderField = "mobiledefaultdisplayorder"
		End If 
	End If 
Else
	sOrderField = "securitydisplayorder"
End If 


sSql = "UPDATE egov_organization_features SET " & sOrderField & " = " & iOrder & " WHERE featureid = " & iFeatureId
RunSQL sSql

response.write "UPDATED"


'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' void RunSQL sSql 
'-------------------------------------------------------------------------------------------------
Sub RunSQL( ByVal sSql )
	Dim oCmd

'	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub 


'-------------------------------------------------------------------------------------------------
' boolean FeatureIsTopLevel( iFeatureId )
'-------------------------------------------------------------------------------------------------
Function FeatureIsTopLevel( ByVal iFeatureId )
	Dim sSql, oRs

	sSql = "SELECT parentfeatureid FROM egov_organization_features WHERE featureid = " & iFeatureId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If CLng(oRs("parentfeatureid")) = CLng(0) Then
		FeatureIsTopLevel = True 
	Else
		FeatureIsTopLevel = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 



%>