<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: saveItemsPerDay.asp
' AUTHOR: David Boyer
' CREATED: 07/12/2010
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Updates the Number of Items Displayed Per Day for an org's calendar(s)
'
' MODIFICATION HISTORY
' 1.0  07/12/10 	David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 lcl_success  = "Y"
 sOrgID       = 0
 sItemsPerDay = 4

 if request("orgid") <> "" then
    sOrgID = request("orgid")
 end if

 if request("itemsPerDay") <> "" then
    sItemsPerDay = request("itemsPerDay")
 end if

 if request("isAjaxRoutine") = "Y" then
    lcl_isAjaxRoutine = True
 else
    lcl_isAjaxRoutine = False
 end if

'Update the calendar with the number of items per day
 sSQL = "UPDATE organizations "
 sSQL = sSQL & " SET calendar_numitemsPerDay = " & sItemsPerDay
 sSQL = sSQL & " WHERE orgid = " & sOrgID

 set oSaveItemsPerDay = Server.CreateObject("ADODB.Recordset")
 oSaveItemsPerDay.Open sSQL, Application("DSN"), 3, 1

 set oSaveItemsPerDay = nothing

 if lcl_success = "Y" AND lcl_isAjaxRoutine then
    response.write "Successfully Updated..."
 end if
%>