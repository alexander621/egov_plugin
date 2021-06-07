<!-- #include file="../includes/common.asp" //-->
<!-- #include file="communitylink_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: saveCommunityLinkOptions.asp
' AUTHOR: David Boyer
' CREATED: 04/27/2009
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description: Saves the formid (Action Line Request - Form ID) that is to be used for a feature with the "Post a Comment" link
'              on CommunityLink.
'
' MODIFICATION HISTORY
' 1.0  04/27/09 	David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 lcl_success  = "Y"
 lcl_savetype = "FORMID"
 sOrgID       = 0
 sFeature     = ""
 sColumnName  = ""
 sFieldValue  = ""
 
 if request("orgid") <> "" then
    sOrgID = request("orgid")
 end if

 if request("feature") <> "" then
    sFeature = request("feature")
 end if

 if request("savetype") <> "" then
    lcl_savetype = UCASE(request("savetype"))
 end if

 if request("isAjaxRoutine") = "Y" then
    lcl_isAjaxRoutine = True
 else
    lcl_isAjaxRoutine = False
 end if

'Insert/Update the search options
 'saveCommunityLinkOption sOrgID, sFeature, "CL_postcomments_formid", sFormID, lcl_isAjaxRoutine, lcl_success

'BEGIN: Form ID ---------------------------------------------------------------
 if lcl_success = "Y" then
    if lcl_savetype = "FORMID" then
       if request("formid") <> "" then
          sFieldValue = request("formid")
       end if

       sColumnName = "CL_postcomments_formid"

      'Insert/Update the search options
       saveCommunityLinkOption sOrgID, sFeature, sColumnName, sFieldValue, lcl_isAjaxRoutine, lcl_success
    end if
 end if
'END: Form ID -----------------------------------------------------------------

'BEGIN: Label -----------------------------------------------------------------
 if lcl_success = "Y" then
    if lcl_savetype = "LABEL" then
       if request("label") <> "" then
          sFieldValue = request("label")
       end if

       sColumnName = "CL_postcomments_label"

      'Insert/Update the search options
       saveCommunityLinkOption sOrgID, sFeature, sColumnName, sFieldValue, lcl_isAjaxRoutine, lcl_success
    end if
 end if
'END: Label -------------------------------------------------------------------

 if lcl_success = "Y" AND lcl_isAjaxRoutine then
    response.write "Changes Saved"
 else
    response.write "Error in Ajax"
 end if
%>