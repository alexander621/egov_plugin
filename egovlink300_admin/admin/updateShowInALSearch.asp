<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: updateShowInALSearch.asp
' AUTHOR: David Boyer
' CREATED: 10/05/2009
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Updates the "showInALSearch" column for a form.
'
' MODIFICATION HISTORY
' 1.0  10/05/09 	David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 lcl_success = "Y"
 sFormID   = 0
 sALSearch = "0"

 if request("formid") <> "" then
    sFormID = request("formid")
 end if

 if request("ALSearch") <> "" then
    if UCASE(request("ALSearch")) = "ON" then
       sALSearch = "1"
    end if
 end if

 if request("isAjaxRoutine") = "Y" then
    lcl_isAjaxRoutine = True
 else
    lcl_isAjaxRoutine = False
 end if

'Update the form
 sSQL = "UPDATE egov_action_request_forms SET "
 sSQL = sSQL & " showInALSearch = " & sALSearch
 sSQL = sSQL & " WHERE action_form_id = " & sFormID

 set oSaveOpt = Server.CreateObject("ADODB.Recordset")
 oSaveOpt.Open sSQL, Application("DSN"), 3, 1

 set oSaveOpt = nothing 

 if lcl_success = "Y" AND lcl_isAjaxRoutine then
    response.write "Form Changes Saved"
 end if
%>