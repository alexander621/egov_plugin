<!-- #include file="../includes/common.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<!-- #include file="../customreports/customreports_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: saveactionlinesearchoptions.asp
' AUTHOR: David Boyer
' CREATED: 02/20/2009
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Updates/Inserts the search options on the main action line screen.
'
' MODIFICATION HISTORY
' 1.0  02/20/09 	David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 lcl_success     = "Y"
 sCustomReportID = 0

 if request("customreportid") <> "" then
    sCustomReportID = request("customreportid")
 end if

 if request("isAjaxRoutine") = "Y" then
    lcl_isAjaxRoutine = True
 else
    lcl_isAjaxRoutine = False
 end if

'Insert/Update the search options
 saveCustomReportSearchOption sCustomReportID, "selectAssignedto",        request("selectAssignedto"),        lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "orderBy",                 request("orderBy"),                 lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "recordsPer",              request("recordsPer"),              lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "reporttype",              request("reporttype"),              lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "selectFormId",            request("selectFormId"),            lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "selectDeptId",            request("selectDeptId"),            lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "pastDays",                request("pastDays"),                lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "searchDaysType",          request("searchDaysType"),          lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "fromDate",                request("fromDate"),                lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "toDate",                  request("toDate"),                  lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "fromToDateSelection",     request("fromToDateSelection"),     lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "selectDateType",          request("selectDateType"),          lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "statusDISMISSED",         request("statusDISMISSED"),         lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "statusRESOLVED",          request("statusRESOLVED"),          lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "statusWAITING",           request("statusWAITING"),           lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "statusINPROGRESS",        request("statusINPROGRESS"),        lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "statusSUBMITTED",         request("statusSUBMITTED"),         lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "substatus_hidden",        request("substatus_hidden"),        lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "selectUserFName",         request("selectUserFName"),         lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "selectUserLName",         request("selectUserLName"),         lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "selectIssueStreetNumber", request("selectIssueStreetNumber"), lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "selectIssueStreet",       request("selectIssueStreet"),       lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "selectContactStreet",     request("selectContactStreet"),     lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "selectCounty",            request("selectCounty"),            lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "selectTicket",            request("selectTicket"),            lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "selectInitialResponse",   request("selectInitialResponse"),   lcl_isAjaxRoutine, lcl_success
 saveCustomReportSearchOption sCustomReportID, "selectRequestsResolved",  request("selectRequestsResolved"),  lcl_isAjaxRoutine, lcl_success

 if lcl_success = "Y" AND lcl_isAjaxRoutine then
    response.write "Changes Saved"
 end if
%>