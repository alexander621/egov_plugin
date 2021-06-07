<!-- #include file="../includes/common.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<!-- #include file="../customreports/customreports_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: updateactionlinecustomreport.asp
' AUTHOR: David Boyer
' CREATED: 02/23/2009
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Updates information about specified Custom Report
'
' MODIFICATION HISTORY
' 1.0  02/23/09 	David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 lcl_success       = "Y"
 sCustomReportID   = 0
 sCustomReportName = ""

 if request("customreportid") <> "" then
    sCustomReportID = request("customreportid")
 end if

 if request("customreporttype") <> "" then
    sReportType = request("customreporttype")
 else
    sReportType = "ACTIONLINE - USER"
 end if

 sReportName = request("reportname")

 if UCASE(request("isUserDefault")) <> "ON" then
    sIsUserDefault = "OFF"
 else
    sIsUserDefault = request("isUserDefault")
 end if

 if request("isAjaxRoutine") = "Y" then
    lcl_isAjaxRoutine = True
 else
    lcl_isAjaxRoutine = False
 end if

 updateCustomReport sCustomReportID, sReportType, sReportName, sIsUserDefault

 if lcl_isAjaxRoutine then
    response.write UCASE(sIsUserDefault)
 end if
%>