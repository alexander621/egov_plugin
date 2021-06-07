<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: featureselectionupdate.asp
' AUTHOR: Steve Loar
' CREATED: 09/12/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Saves feature selections via Ajax
'
' MODIFICATION HISTORY
' 1.0   09/12/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFeatureOrgId, iFeatureId

iFeatureOrgId = CLng(request("orgid"))
iFeatureId = CLng(request("featureid"))

' if the org has the feature then delete that row
If OrgHasFeatureById( iFeatureOrgId, iFeatureId ) Then 
	DeleteFeatureForOrg iFeatureOrgId, iFeatureId
Else 
	' else add the row
	AddFeatureForOrg iFeatureOrgId, iFeatureId
End If 

response.write "Success"


'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' boolean OrgHasFeatureById( iFeatureOrgId, iFeatureId )
'-------------------------------------------------------------------------------------------------
Function OrgHasFeatureById( ByVal iFeatureOrgId, ByVal iFeatureId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(featureid) AS hits FROM egov_organizations_to_features "
	sSql = sSql & " WHERE featureid = " & iFeatureId & " AND orgid = " & iFeatureOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then
			OrgHasFeatureById = True  
		Else
			OrgHasFeatureById = False 
		End If 
	Else
		OrgHasFeatureById = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' void DeleteFeatureForOrg( iFeatureOrgId, iFeatureId )
'-------------------------------------------------------------------------------------------------
Sub DeleteFeatureForOrg( ByVal iFeatureOrgId, ByVal iFeatureId )
	Dim sSql

	sSql = "DELETE FROM egov_organizations_to_features WHERE featureid = " & iFeatureId & " AND orgid = " & iFeatureOrgId
	RunSQL sSql

End Sub 


'-------------------------------------------------------------------------------------------------
' void AddFeatureForOrg( iFeatureOrgId, iFeatureId )
'-------------------------------------------------------------------------------------------------
Sub AddFeatureForOrg( ByVal iFeatureOrgId, ByVal iFeatureId )
 	Dim sSql

 	sSql = "INSERT INTO egov_organizations_to_features ( featureid, orgid ) VALUES ( " & iFeatureId & ", " & iFeatureOrgId & " )"
	 RunSQL sSql


 'BEGIN: Additional setup for Job/Bid Postings ---------------------------------
  iJobPostingFeatureId = clng(getFeatureID("job_postings"))
  iBidPostingFeatureId = clng(getFeatureID("bid_postings"))

 'Check to see if this is either the Job or Bid Postings feature(s).
 '- If so then check to see if any postings statuses exist for this org.
 '- If none exist then create the default statuses.
 	if clng(iFeatureID) = iJobPostingFeatureId then
   		setupPostingsStatuses iFeatureOrgID, "JOB"
 	end if 

 	if clng(iFeatureID) = iBidPostingFeatureId then
	 	  setupPostingsStatuses iFeatureOrgID, "BID"
 	end if
 'END: Additional setup for Job/Bid Postings -----------------------------------

End Sub 

'------------------------------------------------------------------------------
Sub setupPostingsStatuses(p_orgid, p_posting_type )
  dim sOrgID, sPostingType

  sOrgID       = 0
  sPostingType = ""

  if p_orgid <> "" then
     sOrgID = clng(p_orgid)
  end if

  if p_posting_type <> "" then
     if not containsApostrophe(p_posting_type) then
        sPostingType = ucase(p_posting_type)
     end if
  end if

	'First check to see if any job/bid posting statuses already exist for this org.
	'If yes then do nothing.  If no then insert the default statuses.
	'Default Statuses:
	'  Open (default status)
	'  Closed
	'  Filled
	'  Cancelled

	If CLng(checkForStatuses(p_posting_type)) = CLng(0) Then 
 		'Create the default statuses
	  	createPostingStatus sOrgID, sPostingType, "Open",      1, "'Y'"
  		createPostingStatus sOrgID, sPostingType, "Closed",    2, ""
  		createPostingStatus sOrgID, sPostingType, "Filled",    3, ""
		  createPostingStatus sOrgID, sPostingType, "Cancelled", 4, ""
	End If 

End Sub 

'------------------------------------------------------------------------------
Sub createPostingStatus(p_orgid, p_posting_type, p_status_name, p_display_order, p_default )
	Dim sSql, lcl_orgid, lcl_posting_type, lcl_status_name, lcl_display_order, lcl_default

 lcl_orgid         = 0
 lcl_posting_type  = "NULL"
 lcl_status_name   = "NULL"
 lcl_display_order = 1
 lcl_default       = "NULL"
 lcl_active_flag   = "'Y'"

 if p_orgid <> "" then
    lcl_orgid = clng(p_orgid)
 end if

 if p_posting_type <> "" then
    if not containsApostrophe(p_posting_type) then
       lcl_posting_type = ucase(p_posting_type)
       lcl_posting_type = dbsafe(lcl_posting_type)
       lcl_posting_type = "'" & lcl_posting_type & "'"
    end if
 end if

 if p_status_name <> "" then
    if not containsApostrophe(p_status_name) then
       lcl_status_name = ucase(p_status_name)
       lcl_status_name = dbsafe(lcl_status_name)
       lcl_status_name = "'" & lcl_status_name & "'"
    end if
 end if

 if p_display_order <> "" then
    lcl_display_order = clng(p_display_order)
 end if

 if p_default <> "" then
    lcl_default = ucase(p_default)
    lcl_default = dbsafe(lcl_default)
    lcl_default = "'" & lcl_default & "'"
 end if

	sSQL = "INSERT INTO egov_statuses ("
 sSQL = sSQL & " status_name, "
 sSQL = sSQL & " status_type, "
 sSQL = sSQL & " status_order, "
 sSQL = sSQL & " active_flag, "
 sSQL = sSQL & " orgid, "
 sSQL = sSQL & " default_status "
 sSQL = sSQL & " ) VALUES ( "
	sSQL = sSQL & lcl_status_name   & ", "
 sSQL = sSQL & lcl_posting_type  & ", "
 sSQL = sSQL & lcl_display_order & ", "
 sSQL = sSQL & lcl_active_flag   & ", "
 sSQL = sSQL & lcl_orgid         & ", "
 sSQL = sSQL & lcl_default
 sSQL = sSQL & ")"

	RunSQL sSQL

End Sub 

'-------------------------------------------------------------------------------------------------
' Sub RunSQL( sSql )
'-------------------------------------------------------------------------------------------------
Sub RunSQL( ByVal sSql )
	Dim oCmd

	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub
%>