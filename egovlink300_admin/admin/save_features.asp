<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: save_features.asp
' AUTHOR: Steve Loar
' CREATED: 09/19/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the updating of features for clients
'
' MODIFICATION HISTORY
' 1.0	09/19/2006	Steve Loar  - Initial Version
' 2.0	08/01/08	David Boyer - Create Job/Bid Posting statuses
' 2.1	05/29/09	David Boyer - Added "CL_postcomments_label" and "CL_postcomments_formid"
' 3.0	04/05/2011	Steve Loar - Mobile features added
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, iFeatureId, iOrgId, sFeatureName, sFeaturedescription, sPublicurl, sPublicimageurl, iPublicDisplayOrder
Dim sMobileisactivated, sMobileName, sMobileDisplayOrder, sMobileItemCount, sMobileListCount, sCLPostCommentsLabel
Dim sCLPostCommentsFormID, iJobPostingFeatureId, iBidPostingFeatureId

iOrgId = request("orgid")

iJobPostingFeatureId = CLng(getFeatureID( "job_postings" ))
iBidPostingFeatureId = CLng(getFeatureID( "bid_postings" ))

' Delete the org to features table for this org
sSql = "DELETE FROM egov_organizations_to_features WHERE orgid = " & iOrgId
'response.write sSql & "<br />"
RunSQL sSql

' Rebuild the org to features table for this org
For Each iFeatureId In Request("featureid")
	If UCase(request("featurename"&iFeatureId)) = "NULL" Then
		sFeatureName = request("featurename"&iFeatureId)
	Else 
		If request("featurename"&iFeatureId) = "" Then
			sFeatureName = "NULL"
		Else 
			sFeatureName = "'" & DBsafe(request("featurename"&iFeatureId)) & "'"
		End If 
	End If 
	If UCase(request("featuredescription"&iFeatureId)) = "NULL" Then
		sFeaturedescription = request("featuredescription"&iFeatureId)
	Else 
		If request("featuredescription"&iFeatureId) = "" Then
			sFeaturedescription = "NULL"
		Else 
			sFeaturedescription = "'" & DBsafe(request("featuredescription"&iFeatureId)) & "'"
		End If 
	End If 
	If UCase(request("publicurl"&iFeatureId)) = "NULL" Then
		sPublicurl = request("publicurl"&iFeatureId)
	Else 
		If request("publicurl"&iFeatureId) = "" Then
			sPublicurl = "NULL"
		Else
			sPublicurl = "'" & DBsafe(request("publicurl"&iFeatureId)) & "'"
		End If 
	End If 
	If UCase(request("publicimageurl"&iFeatureId)) = "NULL" Then
		sPublicimageurl = request("publicimageurl"&iFeatureId)
	Else 
		If request("publicimageurl"&iFeatureId) = "" Then
			sPublicimageurl = "NULL"
		Else
			sPublicimageurl = "'" & DBsafe(request("publicimageurl"&iFeatureId)) & "'"
		End If 
	End If 
	'			response.write "publicdisplayorder: " & request("publicdisplayorder"&iFeatureId) & "<br />"
	If request("publicdisplayorder"&iFeatureId) = "" Then
		iPublicDisplayOrder = "NULL"
	Else
		iPublicDisplayOrder = request("publicdisplayorder"&iFeatureId)
	End If 
	If request("publiccanview"&iFeatureId) = "on" Then
		' see if the feature has a public view
		If FeatureHasPublicView( iFeatureId ) Then 
			sPublicCanView = "1"
		Else
			sPublicCanView = "0"
		End If 
	Else
		sPublicCanView = "0"
	End If

	If request("CL_postcomments_label" & iFeatureID) <> "" Then 
		sCLPostCommentsLabel = "'" & dbsafe(request("CL_postcomments_label" & iFeatureID)) & "'"
	Else 
		sCLPostCommentsLabel = "NULL"
	End If 

	If request("CL_postcomments_formid" & iFeatureID) <> "" Then 
		sCLPostCommentsFormID = request("CL_postcomments_formid" & iFeatureID)
	Else 
		sCLPostCommentsFormID = "0"
	End If 

	' Mobile features
	sMobileisactivated = request("mobileisactivated" & iFeatureID)
	If request("mobilename" & iFeatureID) = "" Then
		sMobileName = "NULL"
	Else
		sMobileName = "'" & dbsafe(request("mobilename" & iFeatureID)) & "'"
	End If 
	If CLng(request("mobiledisplayorder" & iFeatureID)) = CLng(0) Then
		sMobileDisplayOrder = "NULL"
	Else
		sMobileDisplayOrder =  CLng(request("mobiledisplayorder" & iFeatureID))
	End If 
	If CLng(request("mobileitemcount" & iFeatureID)) = CLng(0) Then
		sMobileItemCount = "NULL"
	Else
		sMobileItemCount =  CLng(request("mobileitemcount" & iFeatureID))
	End If 
	If CLng(request("mobilelistcount" & iFeatureID)) = CLng(0) Then
		sMobileListCount = "NULL"
	Else
		sMobileListCount =  CLng(request("mobilelistcount" & iFeatureID))
	End If 

	sSql = "INSERT INTO egov_organizations_to_features ( "
	sSql = sSql & "orgid, "
	sSql = sSql & "featureid, "
	sSql = sSql & "featurename, "
	sSql = sSql & "featuredescription, "
	sSql = sSql & "publicurl, "
	sSql = sSql & "publicdisplayorder, "
	sSql = sSql & "publicimageurl, "
	sSql = sSql & "publiccanview, "
	sSql = sSql & "CL_postcomments_label, "
	sSql = sSql & "CL_postcomments_formid, "
	sSql = sSql & "mobileisactivated, "
	sSql = sSql & "mobilename, "
	sSql = sSql & "mobiledisplayorder, "
	sSql = sSql & "mobileitemcount, "
	sSql = sSql & "mobilelistcount "
	sSql = sSql & " ) VALUES ( "
	sSql = sSql & iOrgId & ", "
	sSql = sSql & iFeatureId & ", "
	sSql = sSql & sFeatureName & ", "
	sSql = sSql & sFeaturedescription & ", "
	sSql = sSql & sPublicurl & ", "
	sSql = sSql & iPublicDisplayOrder & ", "
	sSql = sSql & sPublicimageurl & ", "
	sSql = sSql & sPublicCanView & ", "
	sSql = sSql & sCLPostCommentsLabel & ", "
	sSql = sSql & sCLPostCommentsFormID & ", "
	sSql = sSql & sMobileisactivated & ", "
	sSql = sSql & sMobileName & ", "
	sSql = sSql & sMobileDisplayOrder & ", "
	sSql = sSql & sMobileItemCount & ", "
	sSql = sSql & sMobileListCount
	sSql = sSql & " )"
'	If iFeatureId = 2 Then 
'		response.write sSql & "<br />"
'	End If 

	RunSQL sSql

	'Check to see if this is either the Job or Bid Postings feature(s).
	'If so then check to see if any postings statuses exist for this org.
	'If none exist then create the default statuses.
	If CLng(iFeatureID) = iJobPostingFeatureId Then 
		setupPostingsStatuses( "JOB" )
	End If 

	If CLng(iFeatureID) = iBidPostingFeatureId Then 
		setupPostingsStatuses( "BID" )
	End If 

Next

' Return to the edit page
'response.redirect "manage_features.asp?orgid=" & iOrgId
response.redirect "featureselection.asp?s=upd&orgid=" & iOrgId


'--------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'--------------------------------------------------------------------------------------------------
Function DBsafe( ByVal strDB )

	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
	DBsafe = Replace( strDB, "'", "''" )

End Function


'--------------------------------------------------------------------------------------------------
' boolean FeatureHasPublicView( iFeatureId )
'--------------------------------------------------------------------------------------------------
Function FeatureHasPublicView( ByVal iFeatureId )
	Dim sSql, oRs

	sSql = "SELECT haspublicview FROM egov_organization_features WHERE featureid = " & iFeatureId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("haspublicview") Then 
			FeatureHasPublicView = True 
		Else
			FeatureHasPublicView = False 
		End If 
	Else
		FeatureHasPublicView = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' string getFeatureID( sFeature )
'------------------------------------------------------------------------------
Function getFeatureID( ByVal sFeature )
	Dim sSql, oRs

	sSql = "SELECT featureid FROM egov_organization_features "
	sSql = sSql & " WHERE feature = '" & dbsafe(sFeature) & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		getFeatureID = oRs("featureid")
	Else
		getFeatureID = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' void setupPostingsStatuses p_posting_type
'------------------------------------------------------------------------------
Sub setupPostingsStatuses( ByVal p_posting_type )
	'First check to see if any job/bid posting statuses already exist for this org.
	'If yes then do nothing.  If no then insert the default statuses.
	'Default Statuses:
	'  Open (default status)
	'  Closed
	'  Filled
	'  Cancelled

	If CLng(checkForStatuses(p_posting_type)) = CLng(0) Then 
		'Create the default statuses
		createPostingStatus p_posting_type, "Open", 1, "'Y'"
		createPostingStatus p_posting_type, "Closed", 2, "NULL"
		createPostingStatus p_posting_type, "Filled", 3, "NULL"
		createPostingStatus p_posting_type, "Cancelled", 4, "NULL"
	End If 

End Sub 

'------------------------------------------------------------------------------
' void createPostingStatus p_posting_type, p_status_name, p_display_order, p_default 
'------------------------------------------------------------------------------
Sub createPostingStatus( ByVal p_posting_type, ByVal p_status_name, ByVal p_display_order, ByVal p_default )
	Dim sSql

	sSql = "INSERT INTO egov_statuses ( status_name, status_type, status_order, active_flag, orgid, default_status ) VALUES ( "
	sSql = sSql & "'" & p_status_name & "', '" & UCASE(p_posting_type) & "', " & p_display_order & ", 'Y', " & iorgid & ", " & p_default & " )"

	RunSQL sSql

End Sub 


'------------------------------------------------------------------------------
' integer checkForStatuses( p_posting_type )
'------------------------------------------------------------------------------
function checkForStatuses( ByVal p_posting_type )
	Dim sSql, oRs

	sSql = "SELECT COUNT(status_id) AS total_count "
	sSql = sSql & " FROM egov_statuses "
	sSql = sSql & " WHERE orgid = " & iorgid
	sSql = sSql & " AND UPPER(status_type) = '" & UCASE(p_posting_type) & "' "

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		checkForStatuses = oRs("total_count")
	Else
		checkForStatuses = 0
	End If 

	oRs.Close
	set oRs = Nothing 

End Function 


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


%>