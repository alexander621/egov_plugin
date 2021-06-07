<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<!-- #include file="../customreports/customreports_global_functions.asp" //-->
<!--<%=request.cookies("user")("userid")%>-->
<%
 'lcl_dsn = "Driver={SQL Server}; Server=ISPS0014; Database=egovlink300; UID=egovsa; PWD=egov_4303;"
 lcl_dsn = Application("DSN")

'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel     = "../"     'Override of value from common.asp
 lcl_hidden = "HIDDEN"  'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide

 if not userhaspermission(session("userid"),"requests") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Set the permission levels
 getPermissionLevels sLevel, blnCanViewAllActionItems, blnCanViewOwnActionItems, blnCanViewDeptActionItems

'Check for org features
 lcl_orghasfeature_issuelocation                    = orghasfeature("issue location")
 lcl_orghasfeature_actionline_listfull              = orghasfeature("actionline_listfull")
 lcl_orghasfeature_responsetimereporting            = orghasfeature("responsetimereporting")
 lcl_orghasfeature_activity_log_download            = orghasfeature("activity_log_download")
 lcl_orghasfeature_display_multiple_workorders      = orghasfeature("display_multiple_workorders")
 lcl_orghasfeature_action_line_substatus            = orghasfeature("action_line_substatus")
 lcl_orghasfeature_data_export                      = orghasfeature("data export")
 lcl_orghasfeature_csv_export_parsed                = orghasfeature("csv_export_parsed")
 lcl_orghasfeature_customreports                    = orghasfeature("customreports")
 lcl_orghasfeature_customreports_codesections       = orghasfeature("customreports_codesections")
 lcl_orghasfeature_savesearchoptions_actionline     = orghasfeature("savesearchoptions_actionline")
 lcl_orghasfeature_actionline_widget_statussummary  = orghasfeature("actionline_widget_statussummary")
 lcl_orghasfeature_actionline_hide_requestlog       = orghasfeature("actionline_hide_requestlog")
 lcl_orghasfeature_actionline_issuelocation_mapit   = orghasfeature("actionline_issuelocation_mapit")
 lcl_orghasfeature_actionline_maintain_duedate      = orghasfeature("actionline_maintain_duedate")
 lcl_orghasfeature_actionline_hide_internalcomments = orghasfeature("actionline_hide_internalcomments")
 lcl_orghasfeature_csv_foil_export		    = orghasfeature("csv_foil_export")
 if request.cookies("User")("UserID") = "6398" then lcl_orghasfeature_csv_foil_export = true

'Check for user permissions
 lcl_userhaspermission_action_line_substatus           = userhaspermission(session("userid"),"action_line_substatus")
 lcl_userhaspermission_activity_log_download           = userhaspermission(session("userid"),"activity_log_download")
 lcl_userhaspermission_data_export                     = userhaspermission(session("userid"),"data export")
 lcl_userhaspermission_csv_export_parsed               = userhaspermission(session("userid"),"csv_export_parsed")
 lcl_userhaspermission_customreports                   = userhaspermission(session("userid"),"customreports")
 lcl_userhaspermission_customreports_codesections      = userhaspermission(session("userid"),"customreports_codesections")
 lcl_userhaspermission_actionline_widget_statussummary = userhaspermission(session("userid"),"actionline_widget_statussummary")
 lcl_userhaspermission_actionline_hide_requestlog      = userhaspermission(session("userid"),"actionline_hide_requestlog")
 lcl_userhaspermission_actionline_maintain_duedate     = userhaspermission(session("userid"),"actionline_maintain_duedate")
 lcl_userhaspermission_can_close_requests                       = userhaspermission(session("userid"),"can close requests")
 lcl_orghasfeature_modify_actionline_department                 = orghasfeature("modify_actionline_department")

'Determine the screen mode (Report Type) to display
'Screen Modes are: PRINT and DISPLAY and NULL
 if request("screen_mode") <> "" then
    lcl_screen_mode = request("screen_mode")
 else
    lcl_screen_mode = "DISPLAY"
 end if

'Get all of the customreport ids
 lcl_customreportid_actionline_user        = getCustomReportID("ACTIONLINE - USER",         session("orgid"), session("userid"), False)
 lcl_customreportid_actionline_lastqueried = getCustomReportID("ACTIONLINE - LAST QUERIED", session("orgid"), session("userid"), False)
 lcl_customreportid_actionline_defaults    = getCustomReportID("ACTIONLINE - DEFAULTS",     session("orgid"), session("userid"), True)

 if request.ServerVariables("REQUEST_METHOD") = "POST" then
   'Save "Last Queried" Search Options.
    lcl_success = "Y"

		 response.write "<!--SAVED2 " & lcl_customreportid_actionline_lastqueried & "-->"
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectAssignedto",        request("selectAssignedto"),        False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "orderBy",                 request("orderBy"),                 False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "recordsPer",              request("recordsPer"),              False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "reporttype",              request("reporttype"),              False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "reporttype_hideinternal", request("reporttype_hideinternal"), False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectFormId",            request("selectFormId"),            False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectDeptId",            request("selectDeptId"),            False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "pastDays",                request("pastDays"),                False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "searchDaysType",          request("searchDaysType"),          False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "fromDate",                request("fromDate"),                False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "toDate",                  request("toDate"),                  False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "fromToDateSelection",     request("Date"),                    False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectDateType",          request("selectDateType"),          False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusDISMISSED",         request("statusDISMISSED"),         False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusRESOLVED",          request("statusRESOLVED"),          False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusWAITING",           request("statusWAITING"),           False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusINPROGRESS",        request("statusINPROGRESS"),        False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusSUBMITTED",         request("statusSUBMITTED"),         False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "substatus_hidden",        request("substatus_hidden"),        False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectUserFName",         request("selectUserFName"),         False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectUserLName",         request("selectUserLName"),         False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectIssueStreetNumber", request("selectIssueStreetNumber"), False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectIssueStreet",       request("selectIssueStreet"),       False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectContactStreet",     request("selectContactStreet"),     False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectCounty",            request("selectCounty"),            False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectBusinessName",      request("selectBusinessName"),      False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectTicket",            request("selectTicket"),            False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "rockRegStreet",            request("rockRegStreet"),            False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "ryePhraseSearch",            request("ryePhraseSearch"),            False, lcl_success
		 response.write "<!--" & request("reporttype_hideinternal") & "-->"
 end if

 sReportTypeHideInternal        = ""
 sReportTypeHideInternalChecked = ""

'If user had set search options for this session then get the session values
 lcl_useSessions = request("useSessions")

 if lcl_useSessions = "" then
    lcl_useSessions = 0
 end if

 if lcl_useSessions = 1 then
   'Determine if user has a "Last Queried" record.  If any exist then retreive the "customreportid" and "isuserdefault" values
    getCustomReportInfo "ACTIONLINE - LAST QUERIED", _
                        False, _
                        session("orgid"), _
                        session("userid"), _
                        False, _
                        lcl_customreportid, _
                        lcl_reporttypeid, _
                        lcl_reportname, _
                        lcl_isuserdefault

   'Get the Org/System Defaults
    if lcl_customreportid = "" then
       getCustomReportInfo "ACTIONLINE - DEFAULTS", _
                           True, _
                           session("orgid"), _
                           session("userid"), _
                           False, _
                           lcl_customreportid, _
                           lcl_reporttypeid, _
                           lcl_reportname, _
                           lcl_isuserdefault
    end if

   'Retrieve the search options
    if lcl_customreportid <> "" then
		response.write "<!--HERE " & lcl_customreportid & "-->"

      'Retrieve the "Last Queried" Search Options
       recordsPer              = getCustomReportSearchOption(lcl_customreportid, "recordsPer")
       reporttype              = getCustomReportSearchOption(lcl_customreportid, "reporttype")
	   response.write "<!--" & reporttype & "-->"
       sReportTypeHideInternal = getCustomReportSearchOption(lcl_customreportid, "reporttype_hideinternal")
	   response.write "<!--" & sReportTypeHideInternal & "-->"
       orderBy                 = getCustomReportSearchOption(lcl_customreportid, "orderBy")
       selectAssignedto        = getCustomReportSearchOption(lcl_customreportid, "selectAssignedto")
       selectFormId            = getCustomReportSearchOption(lcl_customreportid, "selectFormId")
       selectDeptId            = getCustomReportSearchOption(lcl_customreportid, "selectDeptId")
       pastDays                = getCustomReportSearchOption(lcl_customreportid, "pastDays")
       searchDaysType          = getCustomReportSearchOption(lcl_customreportid, "searchDaysType")

       toDate                  = getCustomReportSearchOption(lcl_customreportid, "toDate")
       fromDate                = getCustomReportSearchOption(lcl_customreportid, "fromDate")
       fromToDateSelection     = getCustomReportSearchOption(lcl_customreportid, "fromToDateSelection")
       selectDateType          = getCustomReportSearchOption(lcl_customreportid, "selectDateType")

       statusDISMISSED         = getCustomReportSearchOption(lcl_customreportid, "statusDISMISSED")
       statusRESOLVED          = getCustomReportSearchOption(lcl_customreportid, "statusRESOLVED")
       statusWAITING           = getCustomReportSearchOption(lcl_customreportid, "statusWAITING")
       statusINPROGRESS        = getCustomReportSearchOption(lcl_customreportid, "statusINPROGRESS")
       statusSUBMITTED         = getCustomReportSearchOption(lcl_customreportid, "statusSUBMITTED")

       substatus_hidden        = getCustomReportSearchOption(lcl_customreportid, "substatus_hidden")

       selectUserFName         = getCustomReportSearchOption(lcl_customreportid, "selectUserFName")
       selectUserLName         = getCustomReportSearchOption(lcl_customreportid, "selectUserLName")

       selectIssueStreetNumber = getCustomReportSearchOption(lcl_customreportid, "selectIssueStreetNumber")
       selectIssueStreet       = getCustomReportSearchOption(lcl_customreportid, "selectIssueStreet")
       selectContactStreet     = getCustomReportSearchOption(lcl_customreportid, "selectContactStreet")
       selectCounty            = getCustomReportSearchOption(lcl_customreportid, "selectCounty")
       selectBusinessName      = getCustomReportSearchOption(lcl_customreportid, "selectBusinessName")
       selectTicket            = getCustomReportSearchOption(lcl_customreportid, "selectTicket")
       rockRegStreet            = getCustomReportSearchOption(lcl_customreportid, "rockRegStreet")
       ryePhraseSearch            = getCustomReportSearchOption(lcl_customreportid, "ryePhraseSearch")
    end if

 else
   'Get the modified search options.
   'These are the values on the screen that have been entered when the "SEARCH" button was pressed.
  		recordsPer              = request("recordsPer")
		  reporttype              = request("reporttype")

  		orderBy                 = request("orderBy")
		  selectFormId            = request("selectFormId")

  		if (NOT blnCanViewAllActionItems) AND (NOT blnCanViewDeptActionItems) AND blnCanViewOwnActionItems then
       if sPermissionLevel = "View Dept - Edit Dept" OR sPermissionLevel = "View Dept - Edit Own" then
     		   selectAssignedto = request("selectAssignedto")
       else
     		   selectAssignedto = session("userid")
       end if
  		else
  		   selectAssignedto = request("selectAssignedto")
  		end if

    selectFormId            = request("selectFormId")
  		selectDeptId            = request("selectDeptId")
 			pastDays                = request("pastDays")
 			searchDaysType          = request("searchDaysType")

  		selectUserFName         = request("selectUserFName")
		  selectUserLName         = request("selectUserLName")

  		fromDate                = request("fromDate")
		  toDate                  = request("toDate")
    fromToDateSelection     = request("Date")
    selectDateType          = request("selectDateType")

  		statusSUBMITTED         = request("statusSUBMITTED")
		  statusINPROGRESS        = request("statusINPROGRESS")
  		statusWAITING           = request("statusWAITING")
		  statusRESOLVED          = request("statusRESOLVED")
  		statusDISMISSED         = request("statusDISMISSED")

  		substatus_hidden        = request("substatus_hidden")
    show_hide_substatus     = request("show_hide_substatus")

    selectContactStreet     = request("selectContactStreet")
    selectIssueStreetNumber = request("selectIssueStreetNumber")
  		selectIssueStreet       = request("selectIssueStreet")
    selectCounty            = request("selectCounty")
    selectBusinessName      = request("selectBusinessName")
  		selectTicket            = request("selectTicket")
  		rockRegStreet            = request("rockRegStreet")
  		ryePhraseSearch            = request("ryePhraseSearch")
 end if

'Set status
 if statusSUBMITTED = "yes" then
    noStatus = "false"
 else
    statusSUBMITTED = "no"
 end if

 if statusINPROGRESS = "yes" then
    noStatus = "false"
 else
    statusINPROGRESS = "no"
 end if

 if statusWAITING = "yes" then
    noStatus = "false"
 else
    statusWAITING = "no"
 end if

 if statusRESOLVED = "yes" then
    noStatus = "false"
 else
    statusRESOLVED = "no"
 end if

 if statusDISMISSED = "yes" then
    noStatus = "false"
 else
    statusDISMISSED = "no"
 end if

'Determine if this is the initial time the screen has been opened.
 if request("init") = "Y" _
 OR (request("init")  = ""   AND _
    statusSUBMITTED  = "no" AND _
    statusINPROGRESS = "no" AND _
    statusWAITING    = "no" AND _
    statusRESOLVED   = "no" AND _
    statusDISMISSED  = "no" AND _
    substatus_hidden = "") then
    lcl_init = "Y"
 else
    lcl_init = "N"
 end if

 dim lcl_setLastQueryAsUser 
 lcl_setLastQueryAsUserSaved = false

 if lcl_init = "Y" then
    session("isFromEmail") = ""

    statusSUBMITTED  = "yes"
    statusINPROGRESS = "yes"
    statusWAITING    = "yes"
    statusRESOLVED   = "yes"
    statusDISMISSED  = "yes"

    if lcl_orghasfeature_savesearchoptions_actionline then
      'Determine if user has saved search options.  If any exist then retreive the "customreportid" and "isuserdefault" values
       getCustomReportInfo "ACTIONLINE - USER", _
                           False, _
                           session("orgid"), _
                           session("userid"), _
                           True, _
                           lcl_customreportid, _
                           lcl_reporttypeid, _
                           lcl_reportname, _
                           lcl_isuserdefault

      'Retreive default search values if the user has set to use his/her defaults.
       if lcl_customreportid <> "" AND lcl_isuserdefault then
 		lcl_setLastQueryAsUserSaved = true
		response.write "<!--HERE2-->"
          recordsPer              = getCustomReportSearchOption(lcl_customreportid, "recordsPer")
          reporttype              = getCustomReportSearchOption(lcl_customreportid, "reporttype")
	   	  sReportTypeHideInternal = getCustomReportSearchOption(lcl_customreportid, "reporttype_hideinternal")
          orderBy                 = getCustomReportSearchOption(lcl_customreportid, "orderBy")
          selectAssignedto        = getCustomReportSearchOption(lcl_customreportid, "selectAssignedto")
          selectFormId            = getCustomReportSearchOption(lcl_customreportid, "selectFormId")
          selectDeptId            = getCustomReportSearchOption(lcl_customreportid, "selectDeptId")
          pastDays                = getCustomReportSearchOption(lcl_customreportid, "pastDays")
          searchDaysType          = getCustomReportSearchOption(lcl_customreportid, "searchDaysType")

          toDate                  = getCustomReportSearchOption(lcl_customreportid, "toDate")
          fromDate                = getCustomReportSearchOption(lcl_customreportid, "fromDate")
          fromToDateSelection     = getCustomReportSearchOption(lcl_customreportid, "fromToDateSelection")
          selectDateType          = getCustomReportSearchOption(lcl_customreportid, "selectDateType")

          statusDISMISSED         = getCustomReportSearchOption(lcl_customreportid, "statusDISMISSED")
          statusRESOLVED          = getCustomReportSearchOption(lcl_customreportid, "statusRESOLVED")
          statusWAITING           = getCustomReportSearchOption(lcl_customreportid, "statusWAITING")
          statusINPROGRESS        = getCustomReportSearchOption(lcl_customreportid, "statusINPROGRESS")
          statusSUBMITTED         = getCustomReportSearchOption(lcl_customreportid, "statusSUBMITTED")

          substatus_hidden        = getCustomReportSearchOption(lcl_customreportid, "substatus_hidden")

          selectUserFName         = getCustomReportSearchOption(lcl_customreportid, "selectUserFName")
          selectUserLName         = getCustomReportSearchOption(lcl_customreportid, "selectUserLName")

          selectIssueStreetNumber = getCustomReportSearchOption(lcl_customreportid, "selectIssueStreetNumber")
          selectIssueStreet       = getCustomReportSearchOption(lcl_customreportid, "selectIssueStreet")
          selectContactStreet     = getCustomReportSearchOption(lcl_customreportid, "selectContactStreet")
          selectCounty            = getCustomReportSearchOption(lcl_customreportid, "selectCounty")
          selectBusinessName      = getCustomReportSearchOption(lcl_customreportid, "selectBusinessName")
          selectTicket            = getCustomReportSearchOption(lcl_customreportid, "selectTicket")
          rockRegStreet            = getCustomReportSearchOption(lcl_customreportid, "rockRegStreet")
          ryePhraseSearch            = getCustomReportSearchOption(lcl_customreportid, "ryePhraseSearch")
       end if
    end if
 end if

'Check the ReportType and redirect if needed.
'This is needed ONLY in action_line_list.asp as it is the file called from the menubar.

 if reporttype <> "" then
    if lcl_useSessions = 1 then
       lcl_usesessions_url = "?useSessions=1"
    else
       lcl_usesessions_url = ""
    end if

    if UCASE(reporttype) = "SUMMARY" OR UCASE(reporttype) = "STATUSSUMMARY" then
       response.redirect "action_line_summary.asp" & lcl_usesessions_url
    elseif UCASE(reporttype) = "RESPONSESUMMARY" then
       response.redirect "action_line_summary_response.asp" & lcl_usesessions_url
    elseif UCASE(reporttype) = "RESPONSEDETAIL" then
       response.redirect "action_line_list_response.asp"& lcl_usesessions_url
    end if

    if lcl_orghasfeature_actionline_hide_internalcomments then
       if ucase(reporttype) = "LISTFULL" and sReportTypeHideInternal = "" then
          sReportTypeHideInternal        = request("reporttype_hideinternal")
          sReportTypeHideInternalChecked = ""

          if sReportTypeHideInternal = "YES" then
             sReportTypeHideInternalChecked = " checked=""checked"""
          end if
       end if
    end if
 end if
	   response.write "<!--" & sReportTypeHideInternal & "-->"

'BEGIN: Setup Search Options --------------------------------------------------
 today = Date()

'Default the "From/To Dates will search on" field to ACTIVE
 if selectDateType = "" then
    selectDateType = "active"
 end if

'Set report type (List, ListFull, Summary, or Detail)
 if reporttype = "" or IsNull(reporttype) then
    reporttype = "List"
 end if

'Setup the report label of the report type
 lcl_reportlabel = ""

 if reporttype <> "" then
    lcl_reportlabel = reporttype
    lcl_reportlabel = replace(lcl_reportlabel,"DrillThru","Drill Through")
    lcl_reportlabel = replace(lcl_reportlabel,"Full"," (FULL)")
    lcl_reportlabel = replace(lcl_reportlabel,"statussummary","Status Summary")
    lcl_reportlabel = UCASE(lcl_reportlabel)
 end if

'Set order by
 if orderBy = "" or IsNull(orderBy) then
    orderBy = "submit_date" 
 end if

'Check for NULL values.  If empty then set to "all"
 selectFormId        = checkForNullSetToALL(selectFormId)
 selectAssignedto    = checkForNullSetToALL(selectAssignedto)
 selectDeptId        = checkForNullSetToALL(selectDeptId)
 selectUserFName     = checkForNullSetToALL(selectUserFName)
 selectUserLName     = checkForNullSetToALL(selectUserLName)
 selectContactStreet = checkForNullSetToALL(selectContactStreet)
 selectBusinessName  = checkForNullSetToALL(selectBusinessName)

'Past Days
 if searchDaysType = "PAST" then
    if pastDays = "" or IsNull(pastDays) then
       pastDays = "all"
    end if
 else
    if pastDays = "" or IsNull(pastDays) or pastDays = "0" then
       pastDays = "all"
    end if
 end if

' if pastDays = "" or IsNull(pastDays) or pastDays = "0" then
'    pastDays = "all"
' end if

'Tracking Number
 if selectTicket = "" or IsNull(selectTicket) then
    selectTicket = ""
 end if

'If a date range has been selected then calculate the from/to dates.
'Othwerwise, use what has been entered or the user's default(s).
 if fromToDateSelection <> "" AND fromToDateSelection <> "0" then
    getDatesFromDateRangeChoices fromToDateSelection, lcl_fromDate, lcl_toDate

    fromDate = lcl_fromDate
    toDate   = lcl_toDate
 else

   'From Date (last year)
    if fromDate = "" or IsNull(fromDate) then
       fromDate          = dateAdd("yyyy",-1,today)
       'iDefault_fromDate = "Y"
    end if

   'To Date (get today's date)
    if toDate = "" or IsNull(toDate) then
       toDate          = dateAdd("d",0,today)
       'iDefault_toDate = "Y"
    end if
 end if

'Records Per Page
 if lcl_screen_mode <> "PRINT" then
    recordsPer = checkRecordsPerPageFilter(25,recordsPer)
 end if
'END: Setup Search Options ----------------------------------------------------

'Initialize sub-status show/hide
 show_hide_substatus = "HIDE"

 if show_hide_substatus <> "" then
    show_hide_substatus = show_hide_substatus
 end if
%>
<html>
<head>
  <title><%=langBSActionLine%></title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />
 	<link rel="stylesheet" type="text/css" href="pageprint.css" media="print" />

<style>
	.handhover:hover
	{
		cursor: pointer;
	}
  .redText {
     color: #ff0000;
  }

  #selectSubStatus {
     display: none;
  }

  #widget_request_summary {
    text-align: center;
    font-size:  12pt;
    color:      #800000;
  }

  #buttonMapIt
  {
     cursor: pointer;
  }
</style>

  <script type="text/javascript" src="../scripts/selectAll.js"></script>
 	<script type="text/javascript" src="../scripts/ajaxLib.js"></script>
 	<script type="text/javascript" src="../scripts/getdates.js"></script>
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
  
<script type="text/javascript">
<!--
function setCookie()
{
	var state = $("#accordion h3").hasClass("ui-state-active");
	var d = new Date();
	var days = 7;
    	d.setTime(d.getTime() + (days*24*60*60*1000));
    	var expires = "expires="+ d.toUTCString();
    	document.cookie = "alsso=" + state + ";" + expires + ";";
}

function checkStat() {
  if ( !(form1.statusSUBMITTED.checked) &&  !(form1.statusINPROGRESS.checked) && !(form1.statusWAITING.checked) && !(form1.statusRESOLVED.checked) && !(form1.statusDISMISSED.checked)) {
    		alert("You must select the status.");

    		form1.statusSUBMITTED.focus();
    		return false;
 	}
}

function CheckAllStatus(checkSt) {
		if (checkSt) {
   			document.form1.statusSUBMITTED.checked  = true;
   			document.form1.statusINPROGRESS.checked = true;
   			document.form1.statusWAITING.checked    = true;
  	 		document.form1.statusRESOLVED.checked   = true;
   			document.form1.statusDISMISSED.checked  = true;
 	} else {
   			document.form1.statusSUBMITTED.checked  = false;
   			document.form1.statusINPROGRESS.checked = false;
   			document.form1.statusWAITING.checked    = false;
   			document.form1.statusRESOLVED.checked   = false;
   			document.form1.statusDISMISSED.checked  = false;
		}
}
 
function submitForm() {
  if(validateFields()) {
   	 if (document.form1.reporttype.value == "Summary" || document.form1.reporttype.value == "statussummary") {
  	   			document.forms[0].action = "action_line_summary.asp";
   		  		document.forms[0].submit();
   		} else if (document.form1.reporttype.value == "ResponseSummary") {
  	   			document.forms[0].action = "action_line_summary_response.asp";
   		  		document.forms[0].submit();
   		} else if (document.form1.reporttype.value == "responsedetail") {
  	   			document.forms[0].action = "action_line_list_response.asp";
		  		   document.forms[0].submit();
   		//}	else if (document.form1.reporttype.value == "ListFull") {
  	  // 			document.forms[0].action = "action_line_list.asp";
		  	//	   document.forms[0].submit();
   		}	else {
  	   			document.forms[0].action = "action_line_list.asp"
		  		   document.forms[0].submit();
   		}
  }
}

		function isDate(txtDate)
		{
			var currVal = txtDate;
			if(currVal == '')
				return false;
  			
  			//Declare Regex  
			var rxDatePattern = /^(\d{1,2})(\/|-)(\d{1,2})(\/|-)(\d{4})$/; 
			var dtArray = currVal.match(rxDatePattern); // is format OK?
			
			if (dtArray == null)
				return false;
 			
			//Checks for mm/dd/yyyy format.
			dtMonth = dtArray[1];
			dtDay= dtArray[3];
			dtYear = dtArray[5];
			
			if (dtMonth < 1 || dtMonth > 12)
    			return false;
			else if (dtDay < 1 || dtDay> 31)
    			return false;
			else if ((dtMonth==4 || dtMonth==6 || dtMonth==9 || dtMonth==11) && dtDay ==31)
    			return false;
			else if (dtMonth == 2)
			{
   			var isleap = (dtYear % 4 == 0 && (dtYear % 100 != 0 || dtYear % 400 == 0));
   			if (dtDay> 29 || (dtDay ==29 && !isleap))
       			return false;
			}
			return true;
		}

function validateFields() {

		var daterege         = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
		//var dateFromOk       = daterege.test(document.getElementById("fromDate").value);
		var dateFromOk       = isDate(document.getElementById("fromDate").value);
		//var dateToOk         = daterege.test(document.getElementById("toDate").value);
		var dateToOk         = isDate(document.getElementById("toDate").value);
  var lcl_return_false = 0;
  var lcl_msg_label    = "Display - Open Over Days";

  if (document.getElementById("searchDaysType").value == "PAST") {
      lcl_msg_label = "Display - Past Due Date Days";
  }

  if (document.getElementById("pastDays").value != "") {
      if (! Number(document.getElementById("pastDays").value)) {
          if (Number(document.getElementById("pastDays").value) != "0") {
              document.getElementById("pastDays").focus();
              inlineMsg(document.getElementById("pastDays").id,'<strong>Invalid Value: </strong> "' + lcl_msg_label + '" must be numeric.',10,'pastDays');
              lcl_return_false = lcl_return_false + 1;
          }
      }else{
          if(document.getElementById("pastDays").value < 0 && document.getElementById("searchDaysType").value == "OPEN") {
             document.getElementById("pastDays").focus();
             inlineMsg(document.getElementById("pastDays").id,'<strong>Invalid Value: </strong> "' + lcl_msg_label + '" must be greater than zero (0).',10,'pastDays');
             lcl_return_false = lcl_return_false + 1;
          }else{
             clearMsg("pastDays");
          }
      }
  }else{
      clearMsg("pastDays");
  }


//  if (document.getElementById("pastDays").value != "") {
//      if (! Number(document.getElementById("pastDays").value)) {
//          if (Number(document.getElementById("pastDays").value) == "0") {
//              document.getElementById("pastDays").focus();
//              inlineMsg(document.getElementById("pastDays").id,'<strong>Invalid Value: </strong> "Requests Open Over __ Days" must be greater than zero (0).',10,'pastDays');
//              lcl_return_false = lcl_return_false + 1;
//          }else{
//              document.getElementById("pastDays").focus();
//              inlineMsg(document.getElementById("pastDays").id,'<strong>Invalid Value: </strong> "Requests Open Over __ Days" must be numeric.',10,'pastDays');
//              lcl_return_false = lcl_return_false + 1;
//          }
//      }else{
//          if(document.getElementById("pastDays").value < 0) {
//             document.getElementById("pastDays").focus();
//             inlineMsg(document.getElementById("pastDays").id,'<strong>Invalid Value: </strong> "Requests Open Over __ Days" must be greater than zero (0).',10,'pastDays');
//             lcl_return_false = lcl_return_false + 1;
//          }else{
//             clearMsg("pastDays");
//          }
//      }
//  }else{
//      clearMsg("pastDays");
//  }

		if (! dateToOk ) {
      document.getElementById("toDate").focus();
      inlineMsg(document.getElementById("toDateCalPop").id,'<strong>Invalid Value: </strong> The "To Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'toDateCalPop');
      lcl_return_false = lcl_return_false + 1;
  }else{
      clearMsg("toDateCalPop");
  }

		if (! dateFromOk ) {
      document.getElementById("fromDate").focus();
      inlineMsg(document.getElementById("fromDateCalPop").id,'<strong>Invalid Value: </strong> The "From Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'fromDateCalPop');
      lcl_return_false = lcl_return_false + 1;
  }else{
      clearMsg("fromDateCalPop");
  }

  if (lcl_return_false > 0) {
      return false;
  }else{
      return true;
  }
}

function doCalendar(ToFrom) {
  w = (screen.width - 350)/2;
  h = (screen.height - 350)/2;
  eval('window.open("calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}

function changeRowColor(pID,pStatus) {
  if(pStatus=="OVER") {
     document.getElementById(pID).style.cursor          = "hand";
     document.getElementById(pID).style.backgroundColor = "#93bee1";
  }else{
     document.getElementById(pID).style.cursor          = "";
     document.getElementById(pID).style.backgroundColor = "";
  }
}

function changeSubStatus() {
  var list
  var list2
  var i
  var a

  //mainlist          = document.getElementById('status');
  sub_list          = document.getElementById('selSubStatus');
  sub_list_row      = document.getElementById('sub_status_row');
  sub_list_row_text = document.getElementById('sub_status_row_text');
  i = 0
<%
  dim oMainStatus, oSubStatus, oSubStatus_Count, line_count, lcl_sub_line_count, lcl_total_count

 'Retrieve all of the MAIN statuses
  sSqlm = "SELECT action_status_id, status_name, orgid, parent_status, display_order, active_flag "
  sSqlm = sSqlm & " FROM egov_actionline_requests_statuses "
  sSqlm = sSqlm & " WHERE orgid = 0 "
  sSqlm = sSqlm & " AND parent_status = 'MAIN' "
  sSqlm = sSqlm & " AND active_flag = 'Y' "
  sSqlm = sSqlm & " ORDER BY display_order "

  Set oMainStatus = Server.CreateObject("ADODB.Recordset")
  'oMainStatus.Open sSqlm, Application("DSN"), 0, 1
  oMainStatus.Open sSqlm, lcl_dsn, 0, 1

  If NOT oMainStatus.EOF Then
     line_count = 0
	    while NOT oMainStatus.EOF
        line_count = line_count + 1

        if line_count = 1 then
           response.write "           if(document.getElementById('status" & oMainStatus("status_name") & "').checked==true) {" & vbcrlf
        else
           response.write "           }else if(document.getElementById('status" & oMainStatus("status_name") & "').checked==true) {" & vbcrlf
        end if

	      'Get the total count of SubStatuses
        sSqlc = "SELECT count(action_status_id) AS Total_SubStatus FROM egov_actionline_requests_statuses "
        sSqlc = sSqlc & " WHERE orgid = "         & clng(Session("OrgID"))
        sSqlc = sSqlc & " AND parent_status = '"  & oMainStatus("status_name") & "' "
        sSqlc = sSqlc & " AND active_flag = 'Y' "
        Set oSubStatus_Count = Server.CreateObject("ADODB.Recordset")
        'oSubStatus_Count.Open sSqlc, Application("DSN"), 0, 1
        oSubStatus_Count.Open sSqlc, lcl_dsn, 0, 1

        lcl_total_count = oSubStatus_Count("Total_SubStatus")

        if lcl_total_count > 0 then
		  
		        'Retrieve all of the Sub-Statuses for each MAIN status for the OrgID and the form
           sSqls = "SELECT action_status_id, status_name "
           sSqls = sSqls & " FROM egov_actionline_requests_statuses "
           sSqls = sSqls & " WHERE orgid = "         & clng(Session("OrgID"))
           sSqls = sSqls & " AND parent_status = '"  & oMainStatus("status_name") & "' "
           sSqls = sSqls & " AND active_flag = 'Y' "
           sSqls = sSqls & " ORDER BY display_order "

           Set oSubStatus = Server.CreateObject("ADODB.Recordset")
           'oSubStatus.Open sSqls, Application("DSN"), 0, 1
           oSubStatus.Open sSqls, lcl_dsn, 0, 1

           If NOT oSubStatus.EOF Then
              response.write "              sub_list_row.style.display = ""block"";" & vbcrlf
              response.write "              sub_list.style.display     = ""block"";" & vbcrlf
              response.write "            //remove the current values" & vbcrlf
              response.write "          		  for(var i=0; i < sub_list.length; i++) {" & vbcrlf
              response.write "                  sub_list.remove(i);" & vbcrlf
              response.write "              }" & vbcrlf

             'Loop through the sub statuses
              lcl_sub_line_count = 0
              while NOT oSubStatus.EOF

                 response.write "            //build the new values" & vbcrlf
                 response.write "              document.forms[""form1""].selSubStatus.options[" & lcl_sub_line_count & "] = new Option(""" & oSubStatus("status_name") & """,""" & oSubStatus("action_status_id") & """);" & vbcrlf

                 lcl_sub_line_count = lcl_sub_line_count + 1
                 oSubStatus.movenext
              wend

           	  oSubStatus.Close
         			  oSubStatus_Count.Close

         			  Set oSubStatus       = Nothing
         			  Set oSubStatus_Count = Nothing 
		   
		         else
              response.write "              sub_list_row.style.display = ""none"";" & vbcrlf
              response.write "              sub_list.style.display     = ""none"";" & vbcrlf
           end if
        else
           response.write "              sub_list_row.style.display      = ""none"";" & vbcrlf
           response.write "              sub_list.style.display          = ""none"";" & vbcrlf
        end if

		      oMainStatus.movenext
	    wend

     response.write "           }else{" & vbcrlf
     response.write "              sub_list_row.style.display = ""none"";" & vbcrlf
     response.write "              sub_list.style.display     = ""none"";" & vbcrlf
     response.write "           }" & vbcrlf
  end if

  oMainStatus.Close
  Set oMainStatus = Nothing 
%>
}

var isState = false;
function checkUncheckAll() {
  isSet = document.getElementsByName("substatus");

  if(!isState) {
     for(i=0; i<isSet.length; i++){
		 isSet[i].checked = true
     }
     isState = true;
  } else if (isState) {
     for (i=0; i<isSet.length; i++){
		  isSet[i].checked = false
     }
     isState = false;
  }
}

function change_substatus_filter() {
  var lcl_substatus_display = document.getElementById('display_substatus');
  var lcl_substatus_text    = document.getElementById('substatus_hidden');
  var lcl_substatus_value;
<%
 'Get a total count of all of the active sub-statuses for this org.
  sSqlc = "SELECT count(s1.action_status_id) AS total_count "
  sSqlc = sSqlc & " FROM egov_actionline_requests_statuses s1 "
  sSqlc = sSqlc & " WHERE s1.active_flag = 'Y' "
  sSqlc = sSqlc & " AND s1.orgid = " & session("orgid")

  Set oTotal = Server.CreateObject("ADODB.Recordset")
  'oTotal.Open sSqlc, Application("DSN"), 0, 1
  oTotal.Open sSqlc, lcl_dsn, 0, 1

  lcl_total_substatuses = oTotal("total_count")

  oTotal.close
  set oTotal = nothing

 '1. Build the javascript that will cycle through all of the sub-status search criteria checkboxes and determine which ones have been checked.
 '2. Clear the hidden field (element_id = substatus_hidden) and the display field (element_id = display_substatus) that hold the values.
 '   the hidden field is used for form and query processing and display field is to show the user which values he/she has selected if the
 '   substatus list has been collapsed.
 '3. For those that have been checked then rebuild the hidden field by cycling through all of the checkboxes.
 '4. This will also rebuild the list when a value unchecked.

 'Get all of the active sub-statuses for this org.
  sSqla = "SELECT s1.action_status_id, s1.status_name, s1.display_order, s1.parent_status, s2.display_order AS parent_display_order "
  sSqla = sSqla & " FROM egov_actionline_requests_statuses s1, egov_actionline_requests_statuses s2 "
  sSqla = sSqla & " WHERE s1.parent_status = s2.status_name "
  sSqla = sSqla & " AND s1.active_flag = 'Y' "
  sSqla = sSqla & " AND s2.active_flag = 'Y' "
  sSqla = sSqla & " AND s1.orgid = " & session("orgid")
  sSqla = sSqla & " ORDER BY 5, 4, 3, 2 "

  Set oChange = Server.CreateObject("ADODB.Recordset")
  'oChange.Open sSqla, Application("DSN"), 0, 1
  oChange.Open sSqla, lcl_dsn, 0, 1

%>
     for(var i=0; i < <%=lcl_total_substatuses%>; i++) {
         lcl_substatus_display.innerHTML  = "";
         lcl_substatus_text.value         = "";
<%
  if not oChange.eof then
     do while not oChange.eof
%>
	     lcl_substatus_value = document.getElementById('SS_<%=oChange("action_status_id")%>');
		 if(lcl_substatus_value.checked==true) {
            if(lcl_substatus_display.innerHTML == "") {
	           lcl_substatus_display.innerHTML = '<%=oChange("status_name")%>';
   	 	       lcl_substatus_text.value        = "("+ lcl_substatus_value.value +")";
	        }else{
               lcl_substatus_display.innerHTML = lcl_substatus_display.innerHTML + ", " + '<%=oChange("status_name")%>';
               lcl_substatus_text.value        = lcl_substatus_text.value        + ", " + "("+ lcl_substatus_value.value +")";
            }
         }
<%
        oChange.movenext
	 loop
  end if

  oChange.close
  set oChange = nothing
%>
     }	 
}

//function showhide_substatus_criteria() {
//  var lcl_subStatusList = document.getElementById('selectSubStatus');

//  if(lcl_subStatusList.style.display == "block") {
//     lcl_subStatusList.style.display = "none";
//     document.getElementById('show_hide_substatus').value = "HIDE";
//  }else{
//     lcl_subStatusList.style.display = "block";
//     document.getElementById('show_hide_substatus').value = "SHOW";
//  }
//}

//function show_hide_init(p_value) {
//  var lcl_subStatusList = document.getElementById('selectSubStatus');

//  if(p_value=="SHOW") {
//     lcl_subStatusList.style.display = "block";
//     document.getElementById("show_hide_substatus").value = "SHOW";
//  }else if(p_value=="HIDE") {
//     lcl_subStatusList.style.display = "none";
//     document.getElementById("show_hide_substatus").value = "HIDE";
//  }
//}

//function showhide_substatus_criteria() {
//  var lcl_subStatusList     = $('#selectSubStatus')
//  var lcl_showHideSubStatus = $('#show_hide_substatus');

//  if(lcl_showHideSubStatus.val() == 'HIDE') {
//     lcl_subStatusList.slideDown(1500, function() {
//        lcl_showHideSubStatus.val('SHOW');
//        alert('here');
//     });
//  } else {
//     lcl_showHideSubStatus.val('HIDE');
//     lcl_subStatusList.slideUp('slow');
//  }
//}

//function show_hide_init(p_value) {
//  var lcl_subStatusList     = $('#selectSubStatus');
//  var lcl_showHideSubStatus = $('#show_hide_substatus');

//  if(p_value == 'SHOW') {
//     lcl_subStatusList.slideDown('slow');
//     lcl_showHideSubStatus.val('SHOW');
//  } else if(p_value == 'HIDE') {
//     lcl_subStatusList.slideUp('slow');
//     lcl_showHideSubStatus.val('HIDE');
//  }
//}

function printit() {
  if(window.print) {
	    window.print() ;
  } else {
     var WebBrowser = '<OBJECT ID="WebBrowser1" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>';
     document.body.insertAdjacentHTML('beforeEnd', WebBrowser);
     WebBrowser1.ExecWB(6, 2);//Use a 1 vs. a 2 for a prompting dialog box
     WebBrowser1.outerHTML = "";
  }
}

checked=false;
function checkedAll () {
	 var x = document.getElementById('requestlist');
	 if (checked == false) {
      checked = true
  }else{
      checked = false
  }

 	for (var i=0; i < x.elements.length; i++) {
     	 x.elements[i].checked = checked;
 	}
}

function openRequestManager(p_link_str) {
   location.href="action_respond.asp?" + p_link_str;
}

function bulkAssign()
{
	var x = document.getElementById("requestlist");
	var lcl_requestids = "";

	//Get all of the requests selected
	//isNaN = is Not A Number.  It is a "negative" check that looks for values that are NOT number.  If so then the value returned is "true".
	//    Since we are looking FOR numbers we want the value to be "false".
	for(var i=0; i < x.elements.length; i++) 
	{
		if((x.elements[i].checked==true)&&(isNaN(x.elements[i].value)==false)) 
		{
			if(lcl_requestids=="") 
			{
				lcl_requestids = x.elements[i].value;
			}
			else
			{
				lcl_requestids = lcl_requestids + "," + x.elements[i].value;
			}
		}
	}

	if (lcl_requestids != "")
	{
		lcl_param_list = "?bulkemployeeid=" + document.getElementById("bulkemployeeid").value + "&bulkdeptid=" + document.getElementById("bulkdeptid").value + "&bulkstatus=" + document.getElementById("bulkstatus").value + "&irequestid=" + lcl_requestids;
		document.getElementById("requestlist").action = "bulk_assign.asp" + lcl_param_list;
		document.requestlist.submit();
	}
	else
	{
		alert( "Please select at least one request to include in the work orders." );
	}
}
function printWorkOrders() 
{
	var x = document.getElementById("requestlist");
	var lcl_requestids = "";

	//Get all of the requests selected
	//isNaN = is Not A Number.  It is a "negative" check that looks for values that are NOT number.  If so then the value returned is "true".
	//    Since we are looking FOR numbers we want the value to be "false".
	for(var i=0; i < x.elements.length; i++) 
	{
		if((x.elements[i].checked==true)&&(isNaN(x.elements[i].value)==false)) 
		{
			if(lcl_requestids=="") 
			{
				lcl_requestids = x.elements[i].value;
			}
			else
			{
				lcl_requestids = lcl_requestids + "," + x.elements[i].value;
			}
		}
	}

	if (lcl_requestids != "")
	{
		lcl_application = '<%=application("INSTANCE")%>';
		lcl_orgid = '<%=session("orgid")%>';

		lcl_param_list = "?sys=" + lcl_application;
		lcl_param_list = lcl_param_list + "&irequestid=" + lcl_requestids;
		lcl_param_list = lcl_param_list + "&iorgid=" + lcl_orgid;

		//document.getElementById("requestlist").action = "http://secure.eclink.com/egovlink/work_order_pdf_dtb.asp" + lcl_param_list;
		//document.getElementById("requestlist").action = "work_order_pdf_dtb.asp" + lcl_param_list;
		document.getElementById("requestlist").action = "pdfview/work_order_group.aspx" + lcl_param_list;
		<% if session("orgid") = "209" then%>
		document.getElementById("requestlist").action = "pdfview/work_order_grouph.aspx" + lcl_param_list;
		<% end if %>
		document.getElementById("requestlist").target = "_blank";
		document.requestlist.submit();
	}
	else
	{
		alert( "Please select at least one request to include in the work orders." );
	}
}

function openCustomReports(p_report) {
  w = 900;
  h = 500;
  t = (screen.availHeight/2)-(h/2);
  l = (screen.availWidth/2)-(w/2);
  eval('window.open("../customreports/customreports.asp?cr='+p_report+'", "_customreports", "width='+w+',height='+h+',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
}

function openPrinterFriendlyResults() {
  w = 1000;
  h = 500;
  t = (screen.availHeight/2)-(h/2);
  l = (screen.availWidth/2)-(w/2);
  eval('window.open("action_line_list.asp?screen_mode=PRINT&useSessions=1", "_printerfriendly", "width='+w+',height='+h+',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
}

function saveSearchOptions(iCustomReportID) {
  if(validateFields()) {
     //Validate the checkboxes
     lcl_statusDISMISSED  = "no";
     lcl_statusRESOLVED   = "no";
     lcl_statusWAITING    = "no";
     lcl_statusINPROGRESS = "no";
     lcl_statusSUBMITTED  = "no";

     if(document.getElementById("statusDISMISSED").checked) {
      		lcl_statusDISMISSED = document.getElementById("statusDISMISSED").value;
     }
     if(document.getElementById("statusRESOLVED").checked) {
      		lcl_statusRESOLVED = document.getElementById("statusRESOLVED").value;
     }
     if(document.getElementById("statusWAITING").checked) {
      		lcl_statusWAITING = document.getElementById("statusWAITING").value;
     }
     if(document.getElementById("statusINPROGRESS").checked) {
      		lcl_statusINPROGRESS = document.getElementById("statusINPROGRESS").value;
     }
     if(document.getElementById("statusSUBMITTED").checked) {
      		lcl_statusSUBMITTED = document.getElementById("statusSUBMITTED").value;
     }

     //Build the parameter string
   		var sParameter = 'customreportid='           + encodeURIComponent(iCustomReportID);
     sParameter    += '&isAjaxRoutine=Y';
     sParameter    += '&selectAssignedto='        + encodeURIComponent(document.getElementById("selectAssignedto").value);
   		sParameter    += '&orderBy='                 + encodeURIComponent(document.getElementById("orderBy").value);
   		sParameter    += '&recordsPer='              + encodeURIComponent(document.getElementById("recordsPer").value);
   		sParameter    += '&reporttype='              + encodeURIComponent(document.getElementById("reporttype").value);
		if (document.getElementById("reporttype_hideinternal"))
		{
   		sParameter    += '&reporttype_hideinternal=' + encodeURIComponent(document.getElementById("reporttype_hideinternal").value);
		}
     sParameter    += '&selectFormId='            + encodeURIComponent(document.getElementById("selectFormId").value);
   		sParameter    += '&selectDeptId='            + encodeURIComponent(document.getElementById("selectDeptId").value);
   		sParameter    += '&pastDays='                + encodeURIComponent(document.getElementById("pastDays").value);
   		sParameter    += '&searchDaysType='          + encodeURIComponent(document.getElementById("searchDaysType").value);
   		sParameter    += '&fromDate='                + encodeURIComponent(document.getElementById("fromDate").value);
   		sParameter    += '&toDate='                  + encodeURIComponent(document.getElementById("toDate").value);
   		sParameter    += '&fromToDateSelection='     + encodeURIComponent(document.getElementById("fromToDateSelection").value);
   		sParameter    += '&selectDateType='          + encodeURIComponent(document.getElementById("selectDateType").value);
   		sParameter    += '&statusDISMISSED='         + encodeURIComponent(lcl_statusDISMISSED);
   		sParameter    += '&statusRESOLVED='          + encodeURIComponent(lcl_statusRESOLVED);
   		sParameter    += '&statusWAITING='           + encodeURIComponent(lcl_statusWAITING);
   		sParameter    += '&statusINPROGRESS='        + encodeURIComponent(lcl_statusINPROGRESS);
   		sParameter    += '&statusSUBMITTED='         + encodeURIComponent(lcl_statusSUBMITTED);

  <% if lcl_orghasfeature_action_line_substatus AND lcl_userhaspermission_action_line_substatus then %>
   		sParameter    += '&substatus_hidden='        + encodeURIComponent(document.getElementById("substatus_hidden").value);
  <% end if %>

   		sParameter    += '&selectUserFName='         + encodeURIComponent(document.getElementById("selectUserFName").value);
   		sParameter    += '&selectUserLName='         + encodeURIComponent(document.getElementById("selectUserLName").value);

  <% if lcl_orghasfeature_issuelocation then %>
   		sParameter    += '&selectIssueStreetNumber=' + encodeURIComponent(document.getElementById("selectIssueStreetNumber").value);
   		sParameter    += '&selectIssueStreet='       + encodeURIComponent(document.getElementById("selectIssueStreet").value);
   		sParameter    += '&selectCounty='            + encodeURIComponent(document.getElementById("selectCounty").value);
  <% end if %>

   		sParameter    += '&selectContactStreet='     + encodeURIComponent(document.getElementById("selectContactStreet").value);
   		sParameter    += '&selectBusinessName='      + encodeURIComponent(document.getElementById("selectBusinessName").value);
   		sParameter    += '&selectTicket='            + encodeURIComponent(document.getElementById("selectTicket").value);

     doAjax('saveActionLineSearchOptions.asp', sParameter, 'displayScreenMsg', 'post', '0');
  }
}

function updateCustomReport(iCustomReportID,iDefaultsType,iSaveSearchOptions) {
  lcl_customreportname = document.getElementById("userReportName").value;

  //Default Types: USER, SYSTEM
  if(iDefaultsType=="USER") {
     lcl_setSearchOptionsAsDefault = "on";
  }else{
     lcl_setSearchOptionsAsDefault = "no";
  }

  //Build the parameter string
		var sParameter  = 'isAjaxRoutine=Y';
  sParameter     += '&customreportid='   + encodeURIComponent(iCustomReportID);
  sParameter     += '&customreporttype=' + encodeURIComponent('ACTIONLINE - USER');
  sParameter     += '&reportname='       + encodeURIComponent(lcl_customreportname);
  sParameter     += '&isUserDefault='    + encodeURIComponent(lcl_setSearchOptionsAsDefault);

  //Determine if we need to save the search options along with updating the report info
  if(iSaveSearchOptions!="Y") {
     lcl_return_function = 'updateCustomSearchDisplay';
  }else{
     lcl_return_function = '';
  }

  doAjax('updateActionLineCustomReport.asp', sParameter, lcl_return_function, 'post', '0');

  //Nothing is returned from updating the report info if we are saving the search options.
  //Therefore, we need to now save the search options AND also display set the custom report info.
  if(iSaveSearchOptions=="Y") {
     saveSearchOptions(iCustomReportID);
     updateCustomSearchDisplay(lcl_setSearchOptionsAsDefault);
  }
}

function updateCustomSearchDisplay(iIsUserDefault) {
  iIsUserDefault = iIsUserDefault.toUpperCase();

  if(iIsUserDefault=="ON") {
     lcl_msg = "My Saved Options";
  }else{
     lcl_msg = "System Options";
  }

  document.getElementById("customSearchDisplay").innerHTML = lcl_msg;
}

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "";
}

//function setupPrintButtons() {
		//factory.printing.header       = "Printed on &d"
//		factory.printing.footer       = "&bPrinted on &d - Page:&p/&P";
//		factory.printing.portrait     = false;
//		factory.printing.leftMargin   = 0.5;
//		factory.printing.topMargin    = 0.5;
//		factory.printing.rightMargin  = 0.5;
//		factory.printing.bottomMargin = 0.5;
		 
		//enable control buttons
//		var templateSupported = factory.printing.IsTemplateSupported();
//		var controls = idControls.all.tags("input");
//		for ( i = 0; i < controls.length; i++ ) {
//  			controls[i].disabled = false;
//		  	if ( templateSupported && controls[i].className == "ie55" ) {
//			     controls[i].style.display = "inline";
//     }
//  }
//}

<% if lcl_orghasfeature_actionline_widget_statussummary AND lcl_userhaspermission_actionline_widget_statussummary then %>
//function getWidgetResults() {
//  var sParameter = 'orgid='           + encodeURIComponent('<%=session("orgid")%>');
//  sParameter    += '&userid='         + encodeURIComponent('<%=session("userid")%>');
//  sParameter    += '&selectDateType=' + encodeURIComponent(document.getElementById("selectAssignedto").value);
//  sParameter    += '&fromDate='       + encodeURIComponent(document.getElementById("fromDate").value);
//  sParameter    += '&toDate='         + encodeURIComponent(document.getElementById("toDate").value);

//  doAjax('displayActionLineWidget.asp', sParameter, 'displayWidget', 'post', '0');
//}

//function displayWidget(p_code) {
//  document.getElementById("actionlinewidget").innerHTML = p_code;
//}

function getWidgetResults() {
  var lcl_selectDateType = $('#selectAssignedto').val();
  var lcl_fromDate       = $('#fromDate').val();
  var lcl_toDate         = $('#toDate').val();

  $.post('displayActionLineWidget.asp', {
     orgid:          '<%=session("orgid")%>',
     userid:         '<%=session("userid")%>',
     selectDateType: lcl_selectDateType,
     fromDate:       lcl_fromDate,
     toDate:         lcl_toDate
  }, function(result) {
     $('#actionlinewidget').html(result);
  });
}
function getWidget() {
  w = (screen.width - 350)/2;
  h = (screen.height - 350)/2;
  var lcl_selectDateType = $('#selectAssignedto').val();
  var lcl_fromDate       = $('#fromDate').val();
  var lcl_toDate         = $('#toDate').val();

  eval('window.open("displayActionLineWidget.asp?orgid=<%=session("orgid")%>&userid=<%=session("userid")%>&selectDateTime=' + lcl_selectDateType + '&fromDate=' + lcl_fromDate + '&toDate=' + lcl_toDate + '", "_alwidget", "width=350,height=350,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}

function hideWidgetResults() {
  $('#actionlinewidget').slideUp('slow',function() {
//    var lcl_widget  = '<fieldset class=""fieldset"">';
      var lcl_widget  = '  <input type="button" name="widgetButton" id="widgetButton" value="Show Action Line Summary Widget" class="button ui-button ui-widget ui-corner-all" onclick="getWidgetResults();"  /><br />';
          lcl_widget += '  <div align="center" class="redText">*** Summary ONLY uses from/to dates ***<br />and "date type" for results limitation</div>';
//        lcl_widget += '</fieldset>';

     $('#actionlinewidget').html(lcl_widget);
     $('#actionlinewidget').slideDown('slow');
  });
}
<% end if %>

  $(document).ready(function() {
     $('#showHideSubStatusButton').click(function() {
       var lcl_subStatusList     = $('#selectSubStatus')
       var lcl_showHideSubStatus = $('#show_hide_substatus');

       if(lcl_showHideSubStatus.val() == 'HIDE') {
          lcl_subStatusList.css('display','block');
          lcl_showHideSubStatus.val('SHOW');
       } else {
          lcl_showHideSubStatus.val('HIDE');
          lcl_subStatusList.slideUp('slow');
       }
     });

     $('#buttonMapIt').click(function() {
        $('#frmgooglemap').submit();
     });

<% if lcl_orghasfeature_actionline_hide_internalcomments then %>
     $('#reporttype_hideinternal_span').css('display','none');

     if($('#reporttype').val() == 'ListFull')
     {
        $('#reporttype_hideinternal_span').css('display','inline');
     }     

     $('#reporttype').change(function() {
        $('#reporttype_hideinternal_span').css('display','none');

        if($('#reporttype').val() == 'ListFull')
        {
           $('#reporttype_hideinternal_span').css('display','inline');
        }
     });

     $('#reporttype_hideinternal_checkbox').change(function() {
        $('#reporttype_hideinternal').val('');

        if(document.getElementById('reporttype_hideinternal_checkbox').checked)
        {
           $('#reporttype_hideinternal').val('YES');
        }
     });
<% end if %>
  });
//-->
</script>
</head>
<!-- <body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="javascript:document.getElementById('selectSubStatus').style.display='none';"> -->
<%
if lcl_screen_mode <> "PRINT" then
   lcl_onload = ""

   if lcl_screen_mode <> "RESULTSONLY" then
     'If any substatuses have been checked and posted after clicking the Search button then we need to show the 'display' list
      if lcl_userhaspermission_action_line_substatus then
         lcl_display_substatuses = ""

         if substatus_hidden <> "" then
            lcl_display_substatuses = "change_substatus_filter();"
         end if

         'lcl_onload = "document.getElementById('selectSubStatus').style.display='none';" & lcl_display_substatuses
         lcl_onload = lcl_display_substatuses

        'Initialize the sub-status show/hide
         'show_hide_substatus = "HIDE"

         'if show_hide_substatus <> "" then
         '   show_hide_substatus = show_hide_substatus
         'end if

         'lcl_onload = lcl_onload & "show_hide_init('" & show_hide_substatus & "');"

      else
         lcl_onload = ""
      end if

     'Check to see if the org has the "Save Search Options" feature turned on.
      if lcl_orghasfeature_savesearchoptions_actionline then

        'Set up the From/To Date values if a Date Selection has been set.
         if fromToDateSelection <> "0" AND not isnull(fromToDateSelection) AND fromToDateSelection <> "" then
            lcl_onload = lcl_onload & "getDates(document.getElementById('fromToDateSelection').value, 'Date');"
         else
            lcl_onload = lcl_onload & "document.getElementById('fromDate').value='" & fromDate & "';"
            lcl_onload = lcl_onload & "document.getElementById('toDate').value='"   & toDate   & "';"
         end if
      end if
   end if

  'Check to see if we are loading the Action Line Summary Widget
'   if lcl_orghasfeature_actionline_widget_statussummary AND lcl_userhaspermission_actionline_widget_statussummary then
'      lcl_onload = lcl_onload & "getWidgetResults();"
'   end if

  'Build the onLoad
   if lcl_onload <> "" then
      lcl_onload = " onload=""" & lcl_onload & """ "
   else
      lcl_onload = ""
   end if
'else
'   lcl_onload = " onload=""setupPrintButtons();"""
end if

response.write "<body bgcolor=""#ffffff"" leftmargin=""0"" topmargin=""0"" marginheight=""0"" marginwidth=""0""" & lcl_onload & ">" & vbcrlf

'Display the navigation bar and search criteria if the screen mode is not PRINT
 if lcl_screen_mode <> "PRINT" then

   'DrawTabs tabActionline,1
    ShowHeader sLevel
%>
<!--#Include file="../menu/menu.asp"--> 
<script>
function toggleOptions()
{
	$("#searchform").toggle();
}
$( function() {
    $( "#accordion" ).accordion({
	    <% if not request.cookies("alsso") = "true" then %>
	    active:false,
	    <% end if %>
      collapsible: true
    });
    $( "#accordion2" ).accordion({
	    active:false,
      collapsible: true
    });
    $( "#defaultsearchaccord" ).accordion({
	    active:false,
      collapsible: true
    });
  } );
</script>
<style>
#maincontent *
{
	font-family:'Open Sans', sans-serif !important;
}
#maincontent .fa {
    font: normal normal normal 14px/1 FontAwesome !important; 
}
#maincontent font, #maincontent td, #maincontent select, #maincontent th
{
	font-size:14px;
}
#maincontent th
{
	font-weight:normal;
}
#maincontent input[type="checkbox"]
{
	width:15px;
	height:15px;
}
#maincontent>tbody>tr>td
{
	padding: 0 6px 0 6px;
}
.dropbtn {
    background-color: #2C5F93;
    color: white;
    padding: 4px 16px 4px 16px;
    font-size: 16px;
    border: none;
    margin-right:10px;
    margin-bottom:5px;
}

.dropdown {
    position: relative;
    display: inline-block;
    margin-left:40px;

}
#bottom .dropdown
{
	display:none;
}

.dropdown-content {
    display: none;
    position: absolute;
    background-color: #f1f1f1;
    min-width: 280px;
    box-shadow: 0px 8px 16px 0px rgba(0,0,0,0.2);
    z-index: 1;
    margin-left:-198px;
	margin-top:0;
}

.dropdown-content a {
    color: black;
    padding: 12px 16px;
    text-decoration: none;
    display: block;
}

.dropdown-content a:hover {background-color: #ddd}

.dropdown:hover .dropdown-content {
    display: block;
}

.dropdown:hover .dropbtn, .ui-state-active {
    background-color: #2C5F93 !important;
    
}
.dd-green
{
	background-color: green !important;
}
.btn-red
{
	background-color: red !important;
}
.ui-button, .ui-button:hover {
	background-color: #2C5F93;
	color:white;
	margin-bottom:5px;
}
.ui-icon-white {
  background-image: url("../images/ui-icons_ffffff_256x240.png");
}
#accordion .ui-accordion-content,
#accordion2 .ui-accordion-content,
#defaultsearchaccord .ui-accordion-content
{
	height:auto !important;
}
</style>
<%
    response.write "<table border=""0"" cellpadding=""-"" cellspacing=""0"" class=""start"" width=""100%"" id=""maincontent"">" & vbcrlf
    response.write "  <tr valign=""top"">" & vbcrlf
    response.write "      <td width=""60%"" style=""padding:6px;"">" & vbcrlf
    'response.write "          <font size=""+1""><strong>(E-Gov Request Manager) - Manage Action Line Requests</strong></font><br />" & vbcrlf
    'response.write "          <img src=""../images/arrow_2back.gif"" align=""absmiddle"" />&nbsp;" & vbcrlf
    'response.write "          <a href=""javascript:history.back();"">" & langBackToStart & "</a>" & vbcrlf
    response.write "          <span style=""font-size:16px;""><strong>Manage Action Line Requests [E-Gov Request Manager]</strong></span>" & vbcrlf
    'response.write "          <input type=""button"" name=""backButton"" id=""backButton"" value=""<< Back"" class=""button"" onclick=""javascript:history.back();"" />" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "      <td width=""40%"" align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
    response.write "  </tr>" & vbcrlf

    if lcl_screen_mode <> "RESULTSONLY" then

       response.write "  <tr valign=""top"">" & vbcrlf
       response.write "      <td colspan=""2"">" & vbcrlf

      'BEGIN: Search/Sort Options ---------------------------------------------
       'response.write "          <fieldset class=""fieldset"">" & vbcrlf
       'response.write "            <legend><strong>Search/Sorting Option(s)&nbsp;</strong><input type=""button"" value=""Show/Hide Options"" onClick=""toggleOptions();"" /></legend>" & vbcrlf
       response.write "		<div id=""accordion"" onclick=""setCookie();"">"
       response.write "            <h3><strong>Search/Sorting Options</strong></h3>" & vbcrlf
       response.write "			<div>"
       response.write "            <form name=""form1"" id=""searchform"" method=""post"" onSubmit=""return checkStat()"">" & vbcrlf
       response.write "            <table border=""0"" bordercolor=""#ff0000"" cellspacing=""0"" cellpadding=""2"" width=""100%"">" & vbcrlf
       response.write "              <tr valign=""top"">" & vbcrlf

      'Assigned To
       response.write "                  <td nowrap>" & vbcrlf
       response.write "                      <strong>Assigned To: " & vbcrlf

       if not blnCanViewAllActionItems and not blnCanViewDeptActionItems then

         'Display "Currently Logged In Administrator"
          response.write "(User " & session("userID") & ")&nbsp;&nbsp;&nbsp;"
       else
         'Draw list of employees
          DrawAssignedEmployeeSelection selectAssignedto
       end if

      'Order By
       response.write "                      Order By:" & vbcrlf
       response.write "                      <select name=""orderBy"" id=""orderBy"">" & vbcrlf
                                               displayOrderByList orderBy, _
                                                                  lcl_orghasfeature_issuelocation, _
                                                                  lcl_orghasfeature_actionline_maintain_duedate, _
                                                                  lcl_userhaspermission_actionline_maintain_duedate
        selected = ""
	if orderBy = "upvotes desc" then selected = " selected"
	response.write "			<option value=""upvotes desc"" " & selected & ">Upvotes</option>"
       response.write "                      </select>" & vbcrlf
       response.write "                      </strong>" & vbcrlf
       response.write "                  </td>" & vbcrlf


       response.write "              </tr>" & vbcrlf
       response.write "              <tr>" & vbcrlf

      'Status
       if statusSUBMITTED  = "yes" then check1 = " checked=""checked"""
       if statusINPROGRESS = "yes" then check2 = " checked=""checked"""
       if statusWAITING    = "yes" then check3 = " checked=""checked"""
       if statusRESOLVED   = "yes" then check4 = " checked=""checked"""
       if statusDISMISSED  = "yes" then check5 = " checked=""checked"""

       response.write "                  <td valign=""top"" nowrap>" & vbcrlf
       response.write "                      <strong>Status:</strong> " & vbcrlf
                                             displayStatusCheckbox "Submitted",   check1
                                             displayStatusCheckbox "In Progress", check2
                                             displayStatusCheckbox "Waiting",     check3
                                             displayStatusCheckbox "Resolved",    check4
                                             displayStatusCheckbox "Dismissed",   check5

                                             displaySubStatusOptions_new lcl_userhaspermission_action_line_substatus, _
                                                                     substatus_hidden, _
                                                                     show_hide_substatus
       response.write "                  </td>" & vbcrlf
       response.write "              </tr>" & vbcrlf

      'Category
       response.write "              <tr>" & vbcrlf
       response.write "                  <td valign=""top"" nowrap>" & vbcrlf
       response.write "                      <strong>Categories and Forms: </strong>" & vbcrlf
       response.write "                      <select name=""selectFormId"" id=""selectFormId"" onChange=""toggleregAddy(this.value);"">" & vbcrlf
       response.write "                        <option value="""">All Categories</option>" & vbcrlf
                                               fnListForms selectFormID
       response.write "                      </select>" & vbcrlf
       response.write "                  </td>" & vbcrlf
       response.write "              </tr>" & vbcrlf

	response.write "<script> function toggleregAddy(val) " _
				& " { " _
					& " if (val == '17890') " _
					& " { " _
						& " document.getElementById('regAddy').style.display=''; " _
						& " document.getElementById('rockRegStreet').disabled=false; " _
					& " } " _
					& " else " _
					& " { " _
						& " document.getElementById('regAddy').style.display='none'; " _
						& " document.getElementById('rockRegStreet').disabled=true; " _
					& " } " _
				& " } "  _
				& " </script> " 
	hideReg = ""
	disableAddy = ""
	if selectFormID <> "17890" then
		hideReg = "display:none;"
		disableAddy = "disabled=""true"""
	end if


       response.write "              <tr id=""regAddy"" style=""" & hideReg & """>" & vbcrlf
       response.write "                  <td valign=""top"" nowrap>" & vbcrlf
       response.write "                      <strong>Registration Address: </strong>" & vbcrlf
       response.write "                      <input type=""text"" name=""rockRegStreet"" id=""rockRegStreet"" value=""" & rockRegStreet & """ " & disableAddy & " />" & vbcrlf
       response.write "                  </td>" & vbcrlf
       response.write "              </tr>" & vbcrlf


       response.write "              <tr>" & vbcrlf
       response.write "                  <td valign=""top"" nowrap>" & vbcrlf

      'Department
       if blnCanViewAllActionItems OR blnCanViewDeptActionItems then
          response.write "                      <strong>Department: </strong> " & vbcrlf
          response.write "                      <select name=""selectDeptId"" id=""selectDeptId"">" & vbcrlf
          response.write "                        <option value=""all"">All Departments</option>" & vbcrlf
                                                 'Get a list of all available departments for THIS user
                                                  fnListDepts selectDeptId
          response.write "                      </select>&nbsp;&nbsp;&nbsp;"
       end if

      'Report Type
       response.write "                      <strong>Report Type: </strong>" & vbcrlf
       response.write "                      <select name=""reporttype"" id=""reporttype"">" & vbcrlf
                                               displayReportTypesList reporttype, _
                                                                      lcl_orghasfeature_actionline_listfull, _
                                                                      lcl_orghasfeature_responsetimereporting
       response.write "                      </select>" & vbcrlf

       if lcl_orghasfeature_actionline_hide_internalcomments then
          response.write "                      <span id=""reporttype_hideinternal_span""><input type=""checkbox"" name=""reporttype_hideinternal_checkbox"" id=""reporttype_hideinternal_checkbox"" value=""YES""" & sReportTypeHideInternalChecked & " />Hide Internal Comments</span>" & vbcrlf
          response.write "                      <input type=""hidden"" name=""reporttype_hideinternal"" id=""reporttype_hideinternal"" value=""" & sReportTypeHideInternal & """ />" & vbcrlf
       end if

       response.write "                  </td>" & vbcrlf
       response.write "              </tr>" & vbcrlf
       response.write "              <tr>" & vbcrlf

      'BEGIN: Date Range ------------------------------------------------------
       response.write "                  <td valign=""top"" nowrap>" & vbcrlf
       response.write "                      <fieldset class=""fieldset"">" & vbcrlf

      'From Date
       response.write "                        <strong>From: </strong>" & vbcrlf
       response.write "                        <input type=""text"" name=""fromDate"" id=""fromDate"" value=""" & fromDate & """ size=""10"" maxlength=""10"" onchange=""clearMsg('fromDateCalPop');"" />" & vbcrlf
       response.write "                        <a href=""javascript:void doCalendar('From');""><i class=""fa fa-calendar"" id=""fromDateCalPop"" border=""0"" onclick=""clearMsg('fromDateCalPop');"" ></i></a>&nbsp;" & vbcrlf

      'To Date
       response.write "                        <strong>To:</strong>" & vbcrlf
       response.write "                        <input type=""text"" name=""toDate"" id=""toDate"" value=""" & dateAdd("d",-1,toDate) & """ size=""10"" maxlength=""10"" onchange=""clearMsg('toDateCalPop');"" />" & vbcrlf
       response.write "                        <a href=""javascript:void doCalendar('To');""><i class=""fa fa-calendar"" id=""toDateCalPop"" border=""0"" onclick=""clearMsg('toDateCalPop');""></i></a>&nbsp;" & vbcrlf

      'From/To Dates Date Range options
       DrawDateChoices "Date", fromToDateSelection

      'From/To Dates will search on options
       lcl_selected_active = " selected"
       lcl_selected_submit = ""
       lcl_selected_activity = ""
       'selectUserFName     = "all"
       'selectUserLName     = "all"

       if UCASE(selectDateType) = "SUBMIT" then
          lcl_selected_active = ""
          lcl_selected_submit = " selected"
          lcl_selected_activity = ""
       end if

       if UCASE(selectDateType) = "ACTIVITY" then
          lcl_selected_active = ""
          lcl_selected_submit = ""
          lcl_selected_activity = " selected"
       end if

       if selectUserFName <> "all" then
          selectUserFName = selectUserFName
       end if

       if selectUserLName <> "all" then
          selectUserLName = selectUserLName
       end if

       response.write "                        <br /><br />" & vbcrlf
       response.write "                        <strong>From/To Dates will search on:</strong>" & vbcrlf
       response.write "                        <select name=""selectDateType"" id=""selectDateType"">" & vbcrlf
       response.write "                          <option value=""active""" & lcl_selected_active & ">Active Requests</option>" & vbcrlf
       response.write "                          <option value=""submit""" & lcl_selected_submit & ">Submit Date</option>" & vbcrlf
       response.write "                          <option value=""activity""" & lcl_selected_activity & ">Request Had Activity</option>" & vbcrlf
       response.write "                        </select>" & vbcrlf
       response.write "                      </fieldset>" & vbcrlf
       response.write "                  </td>" & vbcrlf
      'END: Date Range --------------------------------------------------------


       response.write "              </tr>" & vbcrlf
       response.write "              <tr>" & vbcrlf
       response.write "                  <td valign=""top"" nowrap>" & vbcrlf
       response.write "                      <strong>Submitted By: &nbsp;&nbsp;" & vbcrlf
       response.write "                      First: <input type=""text"" name=""selectUserFName"" id=""selectUserFName"" value=""" & selectUserFName & """ size=""12"" />&nbsp;" & vbcrlf
       response.write "                      Last:</strong>" & vbcrlf
       response.write "                       <input type=""text"" name=""selectUserLName"" id=""selectUserLName"" value=""" & selectUserLName & """ size=""12"" />" & vbcrlf
       response.write "                  </td>" & vbcrlf
       response.write "              </tr>" & vbcrlf
       response.write "              <tr>" & vbcrlf
       response.write "                  <td valign=""top"" nowrap>" & vbcrlf
       response.write "                      <strong>Contact Street Name:</strong>&nbsp;" & vbcrlf
       response.write "                      <input type=""text"" name=""selectContactStreet"" id=""selectContactStreet"" value=""" & selectContactStreet & """ />" & vbcrlf
       response.write "                  </td>" & vbcrlf
       response.write "              </tr>" & vbcrlf

       if lcl_orghasfeature_issuelocation then
         'Build the custom label for the "county" field
          lcl_display_id   = GetDisplayId("address grouping field")
          lcl_county_label = GetOrgDisplayWithId(session("orgid"),lcl_display_id,true)

          if lcl_county_label = "" then
             lcl_county_label = GetDisplayName(lcl_display_id)
          end if

         'Issue/Problem Location
          response.write "              <tr>" & vbcrlf
          response.write "                  <td valign=""top"" nowrap>" & vbcrlf
          response.write "                      <strong>Issue/Problem Location:&nbsp;&nbsp;" & vbcrlf
          response.write "                      Street Number: <input type=""text"" name=""selectIssueStreetNumber"" id=""selectIssueStreetNumber"" value=""" & selectIssueStreetNumber & """ size=""10"" maxlength=""150"" />&nbsp;" & vbcrlf
          response.write "                      Street Name: </strong> <input type=""text"" name=""selectIssueStreet"" id=""selectIssueStreet"" value=""" & selectIssueStreet & """ size=""30"" maxlength=""300"" />" & vbcrlf
          response.write "                  </td>" & vbcrlf
          response.write "              </tr>" & vbcrlf

         'County (Custom field)
          response.write "              <tr>" & vbcrlf
          response.write "                  <td valign=""top"" nowrap>" & vbcrlf
          response.write "                      <strong>" & lcl_county_label & ":</strong>" & vbcrlf
          response.write "                      <input type=""text"" name=""selectCounty"" id=""selectCounty"" value=""" & selectCounty & """ size=""30"" maxlength=""50"" />" & vbcrlf
          response.write "                  </td>" & vbcrlf
          response.write "              </tr>" & vbcrlf
       end if

      'Business Name
       response.write "              <tr>" & vbcrlf
       response.write "                  <td valign=""top"" nowrap>" & vbcrlf
       response.write "                      <strong>Business Name:</strong>&nbsp;" & vbcrlf
       response.write "                      <input type=""text"" name=""selectBusinessName"" id=""selectBusinessName"" value=""" & selectBusinessName & """ />" & vbcrlf
       response.write "                  </td>" & vbcrlf
       response.write "              </tr>" & vbcrlf

      'Tracking Number
       response.write "              <tr>" & vbcrlf
       response.write "                  <td valign=""top"" nowrap>" & vbcrlf
       response.write "                      <strong>Tracking Number:</strong>" & vbcrlf
       response.write "                      <input type=""text"" name=""selectTicket"" id=""selectTicket"" value=""" & selectTicket & """ size=""15"" />" & vbcrlf
       response.write "                  </td>" & vbcrlf
       response.write "              </tr>" & vbcrlf

       'if session("orgid") = 153 or session("orgid") = 210 or session("orgid") = 115 or session("orgid") = "209" or session("orgid") = "147" then
      'Search Phrase
       response.write "              <tr>" & vbcrlf
       response.write "                  <td valign=""top"" nowrap>" & vbcrlf
       response.write "                      <strong>Search Phrase:</strong>" & vbcrlf
       response.write "                      <input type=""text"" name=""ryePhraseSearch"" id=""ryePhraseSearch"" value=""" & ryePhraseSearch & """ size=""45"" />" & vbcrlf
       response.write "                  </td>" & vbcrlf
       response.write "              </tr>" & vbcrlf
       'end if

       'response.write "TWF" & request.cookies("user")("userid")

      'Determine which "search day type" is "selected"
       lcl_pastDays                     = ""
       lcl_selected_searchdaystype_open = " selected=""selected"""
       lcl_selected_searchdaystype_past = ""

       if searchDaysType = "PAST" then
          lcl_selected_searchdaystype_open = ""
          lcl_selected_searchdaystype_past = " selected=""selected"""
       end if

       if pastDays <> "all" then
          if searchDaysType = "PAST" OR (searchDaysType = "OPEN" and pastDays <> "0") then
             lcl_pastDays = pastDays
          end if
       end if

       response.write "              <tr>" & vbcrlf
       response.write "                  <td valign=""top"">" & vbcrlf

       if lcl_orghasfeature_actionline_maintain_duedate AND lcl_userhaspermission_actionline_maintain_duedate then
          response.write "                      <strong>Display:</strong>&nbsp;" & vbcrlf
          response.write "                      <select name=""searchDaysType"" id=""searchDaysType"">" & vbcrlf
          response.write "                        <option value=""OPEN""" & lcl_selected_searchdaystype_open & ">Open Over Days</option>" & vbcrlf
          response.write "                        <option value=""PAST""" & lcl_selected_searchdaystype_past & ">Past Due Date Days</option>" & vbcrlf
          response.write "                      </select>" & vbcrlf
          response.write "                      <input type=""text"" name=""pastDays"" id=""pastDays"" value=""" & lcl_pastDays & """ size=""2"" onchange=""clearMsg('pastDays');"" /> (in days)" & vbcrlf
       else
          response.write "                      <strong>Display Open Over " & vbcrlf
          response.write "                      <input type=""text"" name=""pastDays"" id=""pastDays"" value=""" & lcl_pastDays & """ size=""2"" onchange=""clearMsg('pastDays');"" /> Days</strong>" & vbcrlf
          response.write "                      <input type=""hidden"" name=""searchDaysType"" id=""searchDaysType"" value=""OPEN"" />" & vbcrlf
       end if

       response.write "                  </td>" & vbcrlf
       response.write "              </tr>" & vbcrlf

      'Records Per Page
       response.write "              <tr>" & vbcrlf
       response.write "                  <td valign=""top"">" & vbcrlf
       'response.write "                      <input type=""button"" class=""button"" onclick=""submitForm();"" value="" SEARCH "" />" & vbcrlf
       'response.write "                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & vbcrlf
       response.write "                      <strong>Records per Page: </strong>" & vbcrlf
       response.write "                      <input type=""text"" name=""recordsPer"" id=""recordsPer"" value=""" & recordsPer & """ size=""5"" maxlength=""4"" />" & vbcrlf
       response.write "                  </td>" & vbcrlf
       response.write "              </tr>" & vbcrlf

      'Search Button
       response.write "              <tr>" & vbcrlf
       response.write "                  <td valign=""top"">" & vbcrlf
       response.write "                      <input type=""button"" name=""searchButton"" id=""searchButton"" value="" SEARCH "" class=""button ui-button ui-widget ui-corner-all"" onclick=""clearScreenMsg();submitForm();"" />" & vbcrlf
      'BEGIN: Save Search Options (Default Search Options) --------------------
       if lcl_orghasfeature_savesearchoptions_actionline then
          'response.write "                  <td align=""center"" width=""400"" rowspan=""4"">" & vbcrlf
                                                displayCustomSearchOptions_new lcl_customreportid_actionline_user
          'response.write "                  </td>" & vbcrlf
       end if
      'END: Save Search Options (Default Search Options) ----------------------
       response.write "                  </td>" & vbcrlf
       response.write "              </tr>" & vbcrlf
       response.write "            </table>" & vbcrlf
       response.write "            </form>" & vbcrlf
       'response.write "          </fieldset>" & vbcrlf
       'response.write "          </div>" & vbcrlf
          response.write "          </div>" & vbcrlf
          response.write "          </div>" & vbcrlf

       if lcl_orghasfeature_issuelocation then
          response.write "          <div align=""right"">" & vbcrlf
          response.write "            <font style=""color: #ff0000;"">* <small><i>= Non-Listed Street Address</i></small></font>" & vbcrlf
          response.write "          </div>" & vbcrlf
       end if
      'END: Search/Sort Options --------------------------------------------------

       response.write "      </td>" & vbcrlf
       response.write "  </tr>" & vbcrlf

    end if

'------------------------------------------------------------------------------
 else  'lcl_screen_mode = "PRINT"
'------------------------------------------------------------------------------
   'BEGIN: THIRD PARTY PRINT CONTROL ------------------------------------------
    response.write "<div id=""idControls"" class=""noprint"">" & vbcrlf
    response.write "  <input type=""button"" value=""Close Window"" onclick=""parent.close()"" class=""button ui-button ui-widget ui-corner-all"" />&nbsp;&nbsp;" & vbcrlf
    response.write "  <input type=""button"" value=""Print the page"" onclick=""window.print()"" class=""button ui-button ui-widget ui-corner-all"" />&nbsp;&nbsp;" & vbcrlf
    'response.write "  <input type=""button"" value=""Print the page"" disabled onclick=""factory.printing.Print(true)"" />&nbsp;&nbsp;" & vbcrlf
    'response.write "  <input type=""button"" value=""Print Preview..."" disabled onclick=""factory.printing.Preview()"" class=""ie55"" />" & vbcrlf
    response.write "</div>" & vbcrlf
    response.write "<object id=""factory"" viewastext  style=""display:none"" classid=""clsid:1663ed61-23eb-11d2-b92f-008048fdd814"" codebase=""../includes/smsx.cab#Version=6,3,434,12""></object>" & vbcrlf
   'END: THIRD PARTY PRINT CONTROL --------------------------------------------

    response.write "<table border=""0"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
 end if

 response.write "  <tr>" & vbcrlf
 response.write "      <td valign=""top"" colspan=""2"">" & vbcrlf

'BEGIN: Action Line Request List ----------------------------------------------
 response.write "          <form name=""requestlist"" id=""requestlist"" action=""#"" method=""post"">" & vbcrlf

 List_Action_Requests session("orgid"), _
                      session("userid"), _
                      sSortBy, _
                      request.querystring, _
                      fromDate, _
                      toDate, _
                      substatus_hidden, _
                      selectFormID, _
                      selectAssignedto, _
                      selectDeptId, _
                      selectUserFName, _
                      selectUserLName, _
                      selectContactStreet, _
                      selectIssueStreetNumber, _
                      selectIssueStreet, _
                      selectCounty, _
                      selectBusinessName, _
                      selectTicket, _
		      rockRegStreet, _
		      ryePhraseSearch, _
                      searchDaysType

 response.write "          </form>" & vbcrlf
'END: Action Line Request List ------------------------------------------------

 response.write "      </td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "</table>" & vbcrlf

'BEGIN: Google Map ------------------------------------------------------------
 'response.write "<form name=""frmgooglemap"" id=""frmgooglemap"" action=""" & Application("MAP_URL") & "admin/action_line/action_line_map.asp"" method=""post"" target=""GOOGLE_MAP_WINDOW"">" & vbcrlf
 response.write "<form name=""frmgooglemap"" id=""frmgooglemap"" action=""action_line_map.asp"" method=""post"" target=""GOOGLE_MAP_WINDOW"">" & vbcrlf
 response.write "  <input type=""hidden"" name=""orgid"" value=""" & session("orgid") & """ />" & vbcrlf
 response.write "  <input type=""hidden"" name=""map_query"" value=""" & session("MAP_QUERY") & """ />" & vbcrlf

 sCurrentURLBase = request.servervariables("SERVER_NAME") & LEFT(request.servervariables("URL"),instrrev(request.servervariables("URL"),"/"))

 response.write "  <input type=""hidden"" name=""current_url"" value=""" & sCurrentURLBase & """ />" & vbcrlf
 response.write "  <input type=""hidden"" name=""orderby"" value=""" & orderby & """ />" & vbcrlf
 response.write "</form>" & vbcrlf
'END: Google Map --------------------------------------------------------------

 if lcl_screen_mode <> "PRINT" then
%>
   <!-- #include file="../admin_footer.asp" -->
<%
 end if

%>
<script>
$("#searchform").keypress(function(e){
    var checkWebkitandIE=(e.which==13 ? 1 : 0);
    var checkMoz=(e.which==13 && e.ctrlKey ? 1 : 0);

    if (checkWebkitandIE || checkMoz)
    {
	clearScreenMsg();submitForm();
    }
    
});
</script>
<%
 response.write "</body>" & vbcrlf
 response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub List_Action_Requests(iOrgID, iUserID, sSortBy, sQueryString, iFromDate, iToDate, iSubStatusHidden, _
                         iSelectFormID, iSelectAssignedTo, iSelectDeptID, iSelectUserFName, iSelectUserLName, _
                         iSelectContactStreet, iSelectIssueStreetNumber, iSelectIssueStreet, _
                         iSelectCounty, iSelectBusinessName, iSelectTicket, srockRegStreet, sryePhraseSearch, iSearchDaysType )

Dim statArray(5)
i = 0

if statusSUBMITTED = "yes" then 
	  statArray(i) = " status='SUBMITTED' OR"
	  i = i + 1
end if

if statusINPROGRESS = "yes" then
	  statArray(i) = " status='INPROGRESS' OR"
  	i = i + 1
end if

if statusWAITING = "yes" then
  	statArray(i) = " status='WAITING' OR"
  	i = i + 1
end if

if statusRESOLVED= "yes" then
  	statArray(i) = " status='RESOLVED' OR"
  	i = i + 1
end if

if statusDISMISSED = "yes" then
  	statArray(i) = " status='DISMISSED' OR"
  	i = i + 1
end if

for u = 0 to ubound(statArray)
   varStatClause = varStatClause & "" & statArray(u)
next

lenStatClause = len(varStatClause) - 3

if lenStatClause > 1 then
  	varStatClause = left(varStatClause,lenStatClause)
end if

'If start date is before from date AND finish date is NOT before from date
'OR
'If start date is NOT before from date AND start date is NOT after to date
'Check the selectDateType to determine how to use the From/To Date fields.

'Also, we add a date to the "toDate" for the SQL ONLY.  Past code added a day to the value used in the search field.
'This became a hassle with the "date selection" dropdown list.
'This SQL looked for everything LESS THAN the "toDate"
' OR
'Used the SQL BETWEEN function which also performs a LESS THAN on the "todate"
'Adding a day to the "toDate" returns the information INCLUDING the "toDate"
 lcl_query_toDate = dateAdd("d",1,iToDate)

if UCASE(selectDateType) = "ACTIVE" then
   varWhereClause = " WHERE  (egov_action_request_view.orgid=" & iOrgID & ") AND ( "    ''IsNull(complete_date,'" & Now & "')
   'varWhereClause = varWhereClause & " (submit_date >= '" & fromDate & "' AND submit_date < '" & toDate & "') OR "
   'varWhereClause = varWhereClause & " ( IsNull(complete_date,'" & Now & "') >= '" & fromDate & "' AND IsNull(complete_date,'" & Now & "') < '" & toDate & "' ) OR "
   'varWhereClause = varWhereClause & " (submit_date < '" & fromDate & "' AND IsNull(complete_date,'" & Now & "') > '" & toDate & "')  "
   varWhereClause = varWhereClause & " (submit_date >= '" & iFromDate & "' AND submit_date < '" & lcl_query_toDate & "') OR "
   varWhereClause = varWhereClause & " ( IsNull(complete_date,'" & Now & "') >= '" & iFromDate & "' AND IsNull(complete_date,'" & Now & "') < '" & lcl_query_toDate & "' ) OR "
   varWhereClause = varWhereClause & " (submit_date < '" & iFromDate & "' AND IsNull(complete_date,'" & Now & "') > '" & lcl_query_toDate & "')  "
elseif UCASE(selectDateType) = "ACTIVITY" then
   varWhereClause = " WHERE (egov_action_request_view.orgid=" & iOrgID
   varWhereClause= varWhereClause & " AND EXISTS(SELECT action_responseid FROM egov_action_responses r WHERE r.action_autoid = egov_action_request_view.action_autoid AND (action_editdate >= '" & iFromDate & "' AND action_editdate <= '" & lcl_query_toDate & "'))"
else 'selectDateType = SUBMIT
   varWhereClause = " WHERE egov_action_request_view.orgid=" & iOrgID
   varWhereClause = varWhereClause & " AND (submit_date BETWEEN '" & iFromDate & "' AND '" & lcl_query_toDate & "'"
   'varWhereClause = varWhereClause & " AND (submit_date BETWEEN '" & fromDate & "' AND '" & toDate & "'"
end if

'Sub-Status Filter
if iSubStatusHidden = "" OR isnull(iSubStatusHidden) then
  'If any Statuses are checked
   if i > 0 then
      varWhereClause = varWhereClause & " ) AND (" & varStatClause & ") "
   else
      varWhereClause = varWhereClause & " ) "
   end if
else
  'If any Statuses are checked
   if i > 0 then
      varWhereClause = varWhereClause & " ) AND ((" & varStatClause & ") "
      varWhereClause = varWhereClause & " OR sub_status_id in (" & REPLACE(REPLACE(iSubStatusHidden,"(",""),")","") & ")) "
   else
      varWhereClause = varWhereClause & " ) AND sub_status_id in (" & REPLACE(REPLACE(iSubStatusHidden,"(",""),")","") & ") "
   end if
end if

'Determine if the category or request form is to be searched on.
'If it is a "category" then retrieve all of the action_formids associated to that category.
if iSelectFormID <> "all" then 
  	if left(iSelectFormID,1)="C" then 
		    sSQLb = "SELECT action_form_id FROM egov_forms_to_categories where form_category_id = " & right(iSelectFormID,len(iSelectFormID)-1)

    		set oCategories = Server.CreateObject("ADODB.Recordset")
    		'oCategories.Open sSQLb, Application("DSN"), 0, 1
    		oCategories.Open sSQLb, lcl_dsn, 0, 1

    		if oCategories.EOF then
      			varWhereClause = varWhereClause & " AND form_category_id=999999"		
    		else
      			do while not oCategories.EOF
       					CatArray = CatArray & oCategories("action_form_id") & ","
         			oCategories.MoveNext
      			loop

         oCategories.close
         set oCategories = nothing

      			CatArray = left(CatArray,(len(CatArray)-1))
    		end if

    		varWhereClause = varWhereClause & " AND action_Formid IN (" & CatArray & ") "
  	else
    		varWhereClause = varWhereClause & " AND action_Formid = " & iSelectFormID
  	end if
end if

if iSelectAssignedTo <> "all" then
   varWhereClause = varWhereClause & " AND assignedemployeeid = " & iSelectAssignedTo
end if

'restrict what they can view by the permission level they have
 If blnCanViewDeptActionItems AND NOT blnCanViewAllActionItems Then
	  'can view dept
   	If iSelectDeptID <> "all" then 
		     varWhereClause = varWhereClause & " AND deptID = '" & iSelectDeptID & "'"
   	Else
     		varWhereClause = varWhereClause & " AND ((deptID IN (" & GetGroups(iUserID) & ")) OR ((assignedemployeeid = '" & iUserID & "') )) "
   	End If
 Else
	   If blnCanViewOwnActionItems And Not blnCanViewAllActionItems And Not blnCanViewDeptActionItems Then
    		'can view own only
     		varWhereClause = varWhereClause & " AND assignedemployeeid = " & session("userid")
   	Else 
    		'Can view all, only add to where clause if a dept is chosen
     		If iSelectDeptID <> "all" then 
       			varWhereClause = varWhereClause & " AND deptID = '" & iSelectDeptID & "' "
     		End If
   	End If 
 End If

 if iSelectUserFName <> "all" AND iSelectUserFName <> "" then
    varWhereClause = varWhereClause & " AND upper(UserFName) LIKE '%" & dbsafe(ucase(iSelectUserFName)) & "%'"
 end if

 if iSelectUserLName <> "all" AND iSelectUserLName <> "" then
    varWhereClause = varWhereClause & " AND upper(UserLName) LIKE '%" & dbsafe(ucase(iSelectUserLName)) & "%'"
 end if

'Contact Street Name
 if iSelectContactStreet <> "all" then varWhereClause = varWhereClause & " AND useraddress LIKE '%" & dbsafe(iSelectContactStreet) & "%'"

'Issue/Problem Location
 lcl_search_address = ""
 if iSelectIssueStreetNumber <> "" AND NOT isnull(iSelectIssueStreetNumber) then
    lcl_search_address = iSelectIssueStreetNumber & "%"
 end if

 lcl_search_address = lcl_search_address & iSelectIssueStreet

 if lcl_search_address <> "" then
    varWhereClause = varWhereClause & " AND UPPER(streetname) LIKE ('%" & UCASE(dbsafe(lcl_search_address)) & "%')"
 end if

'County
 if iSelectCounty <> "" AND NOT isnull(iSelectCounty) then
    varWhereClause = varWhereClause & " AND UPPER(county) LIKE ('%" & UCASE(dbsafe(iSelectCounty)) & "%')"
 end if

'Business Name
 if iSelectBusinessName <> "all" then varWhereClause = varWhereClause & " AND userbusinessname LIKE '%" & dbsafe(iSelectBusinessName) & "%'"

'Tracking Number search
if iSelectTicket <> "" then
   if IsNumeric(iSelectTicket) AND len(iSelectTicket) > 4 then 
      'response.write "<!-- HUH " & iSelectTicket & "-->" & vbcrlf
      sTicketNo = Left(iSelectTicket, (Len(iSelectTicket) - 4))
      'response.write "<!-- HUH " & sTicketNo & "-->" & vbcrlf

      'iTrackID  = CStr(CDbl(iSelectTicket))
      iTrackID  = CStr(iSelectTicket)
      'response.write "<!-- HUH " & iTrackID & "-->" & vbcrlf
      iTime     = Right(iTrackID,4)
      'response.write "<!-- HUH " & iTime & "-->" & vbcrlf
      iHour     = Left(iTime,2)
      'response.write "<!-- HUH " & iHour & "-->" & vbcrlf
      iMinute   = Right(iTime,2)
      'response.write "<!-- HUH " & iMinute & "-->" & vbcrlf
      'response.flush
      If iHour = "" or iMinute = "" Then
         iHour   = "99"
         iMinute = "99"
      End If

      varWhereClause = varWhereClause & " AND (action_autoid = " & sTicketNo & ") "
      varWhereClause = varWhereClause & " AND (DATEPART(hh, submit_date) = '"& iHour &"') "
      varWhereClause = varWhereClause & " AND (DATEPART(mi, submit_date) = '"& iMinute &"')"

   else
     'If this isn't passed then it will appear that the summary isn't working correctly because a value has been entered
     'for the Tracking Number search criteria field.  Using this statement will return the "No Records Found" in the results.
      varWhereClause = varWhereClause & " and 1=2"
   end if
end if

if srockRegStreet <> "" then
	varWhereClause = varWhereClause & " AND EXISTS(SELECT answer FROM action_submitted_questions_and_answers ad WHERE egov_action_request_view.action_autoid = ad.action_autoid and ad.question = 'Address' and ad.answer = '" & srockRegStreet & "')"
end if

if sryePhraseSearch <> "" then
	varWhereClause = varWhereClause & " AND comment LIKE '%" & sryePhraseSearch & "%'"
end if

if lcl_screen_mode <> "PRINT" then
'   sSQL = sSQL & "DateDiff(d,submit_date,ISNULL(complete_date,'" & today & "')) AS lcl_subTotal_totalDaysOpen, "
'   sSQL = sSQL & "DateDiff(d,ISNULL(adjustedsubmitdate,submit_date),ISNULL(complete_date,'" & today & "')) AS lcl_subTotal_totalDaysOpen_adjusted, "
'   sSQL = sSQL & "[dbo].[getDateDiff_NoWeekend] (egov_action_request_view.orgid,egov_action_request_view.submit_date, ISNULL(egov_action_request_view.complete_date,'" & today & "')) AS lcl_subTotal_totalDaysOpen_noweekends, "
'   sSQL = sSQL & "[dbo].[getDateDiff_NoWeekend] (egov_action_request_view.orgid,ISNULL(egov_action_request_view.adjustedsubmitdate,egov_action_request_view.submit_date),ISNULL(egov_action_request_view.complete_date,'" & today & "')) AS lcl_subTotal_totalDaysOpen_noweekends_adjusted, "
'   sSQL = sSQL & "(select ISNULL(status_name,'') "
'   sSQL = sSQL & " from egov_actionline_requests_statuses "
'   sSQL = sSQL & " where sub_status_id = action_status_id) AS sub_status_name, "
'   sSQL = sSQL & " assignedName as assigned_Name, assignedemployeeid, action_form_display_issue, due_date, "
'   sSQL = sSQL & " '' as latitude, '' as longitude, '' as streetnumber, '' as streetaddress, '' as streetname, '' as comment, '' as sortstreetname, "
'   sSQL = sSQL & " allowedunresolveddays, usesafter5adjustment, usesweekdays, '' as issuelocationname, '' as city, '' as state, '' as zip, '' as validstreet, '' as comments,  "
'   sSQL = sSQL & "(Case WHEN status <> 'RESOLVED' AND status <> 'DISMISSED' THEN DateDiff(d,submit_date,'" & date() & "') ELSE 0 END) AS numPast, "
'   sSQL = sSQL & "(DateDiff(d,due_date,'" & date() & "')) AS daysPastDueDate "

   sSQL = "SELECT "
   sSQL = sSQL & "userlname, "
   sSQL = sSQL & "userfname, "
   sSQL = sSQL & "userhomephone, "
   sSQL = sSQL & "useraddress, "
   sSQL = sSQL & "useraddress2, "
   sSQL = sSQL & "usercity, "
   sSQL = sSQL & "userstate, "
   sSQL = sSQL & "userzip, "
   sSQL = sSQL & "action_autoid, "
   sSQL = sSQL & "action_formTitle, "
   sSQL = sSQL & "subTotal_totalDaysOpen, "
   sSQL = sSQL & "subTotal_totalDaysOpen_adjusted, "
   sSQL = sSQL & "subTotal_totalDaysOpen_noweekends, "
   sSQL = sSQL & "subTotal_totalDaysOpen_noweekends_adjusted, "

'   sSQL = sSQL & "[dbo].[checkDateDiff]('DAY',submit_date,'',complete_date,'" & today &"') as lcl_subTotal_totalDaysOpen2, "
'   sSQL = sSQL & "[dbo].[checkDateDiff]('DAY',adjustedsubmitdate,submit_date,complete_date,'" & today &"') as lcl_subTotal_totalDaysOpen_adjusted2, "

   sSQL = sSQL & "submit_date, "
   sSQL = sSQL & "adjustedsubmitdate, "
   sSQL = sSQL & "complete_date, "
   sSQL = sSQL & "deptID, "
   sSQL = sSQL & "isnull(groupname,'') as deptName, "
   sSQL = sSQL & "UPPER(status) as status, "
   sSQL = sSQL & "sub_status_desc, "
   sSQL = sSQL & "assignedName as assigned_Name, "
   sSQL = sSQL & "assignedemployeeid, "
   sSQL = sSQL & "latitude, "
   sSQL = sSQL & "longitude, "
   sSQL = sSQL & "mobileoption_latitude, "
   sSQL = sSQL & "mobileoption_longitude, "
   sSQL = sSQL & "streetnumber, "
   sSQL = sSQL & "streetaddress, "
   sSQL = sSQL & "streetname, "
   sSQL = sSQL & "comment, "
   sSQL = sSQL & "ISNULL(sortstreetname,'') AS sortstreetname, "
   sSQL = sSQL & "allowedunresolveddays, "
   sSQL = sSQL & "usesafter5adjustment, "
   sSQL = sSQL & "usesweekdays, "
   sSQL = sSQL & "issuelocationname, "
   sSQL = sSQL & "city, "
   sSQL = sSQL & "state, "
   sSQL = sSQL & "zip, "
   sSQL = sSQL & "validstreet, "
   sSQL = sSQL & "comments, "
   sSQL = sSQL & "action_form_display_issue, "
   sSQL = sSQL & "due_date, "
   sSQL = sSQL & "userbusinessname, "
   sSQL = sSQL & "numPast, "
   sSQL = sSQL & "daysPastDueDate, "
   sSQL = sSQL & "upvotes "
   sSQL = sSQL & " FROM egov_action_request_view "
   sSQL = sSQL &      " LEFT OUTER JOIN groups ON deptId = groupId "
   sSQL = sSQL & varWhereClause

  'Are we searching on "Open" or "Past" dates
   if iSearchDaysType = "PAST" AND pastDays <> "all" then
      sSQL = sSQL & " AND DateDiff(d,due_date,'" & date() & "') >= " & pastDays
   else
      if pastDays <> "all" and pastDays <> "0" then
         sSQL = sSQL & " AND (status <> 'RESOLVED' "
         sSQL = sSQL & " AND  status <> 'DISMISSED' "
         sSQL = sSQL & " AND  DateDiff(d,submit_date,'" & date() & "') > " & pastDays & ") "
      end if
   end if

   'sSQL = sSQL & " AND (egov_action_request_view.orgid=" & session("orgid") & ")"

  'Set Google Map Query
   session("MAP_QUERY") = sSQL
else
   sSQL = session("MAP_QUERY")
end if

'BEGIN: Order By --------------------------------------------------------------
 if UCASE(orderBy) = "STREETNAME" then
    'lcl_order_by = "UPPER(streetaddress), CAST(streetnumber AS int) "
    lcl_order_by = "UPPER(sortstreetname), CAST(streetnumber AS int) "
 elseif UCASE(orderBy) = "SUBMITTEDBY" then
    lcl_order_by = "UPPER(userlname), UPPER(userfname) "
 elseif UCASE(orderBy) = "STATUS" then
    lcl_order_by = "status_order "
 else
    lcl_order_by = orderBy
 end if


 sSQL = sSQL & " ORDER BY " & lcl_order_by

 'if UCASE(orderBy) = "SUBMIT_DATE" OR UCASE(orderBy) = "DUE_DATE" then
 if UCASE(orderBy) = "SUBMIT_DATE" then
    sSQL = sSQL & " desc"
 end if
'END: Order By ----------------------------------------------------------------

if lcl_screen_mode <> "PRINT" then

  'Store Query for Export to CSV ----------------------------------------------
   session("DISPLAYQUERY") = buildQueryForExportToCSV(sSQL, lcl_orghasfeature_action_line_substatus, _
                                                            lcl_orghasfeature_issuelocation, _
                                                            lcl_orghasfeature_actionline_maintain_duedate, _
                                                            lcl_userhaspermission_actionline_maintain_duedate)

  'Store query for "Download Activity Log" ------------------------------------
   if lcl_orghasfeature_activity_log_download AND lcl_userhaspermission_activity_log_download then
      session("ACTLOGQUERY") = buildQueryForDownloadActivityLog(varWhereClause,lcl_order_by,orderby)
   end if

  'Store query for "Custom Reports - Code Sections" ---------------------------
   'if  lcl_orghasfeature_customreports _
   'AND lcl_userhaspermission_customreports _
   if lcl_orghasfeature_customreports_codesections AND lcl_userhaspermission_customreports_codesections then
      session("CR_CODESECTIONS") = buildQueryForCustomReports("codesections", varWhereClause,lcl_order_by,orderby)
   end if

   set oRequests = Server.CreateObject("ADODB.Recordset")

  'Set page size and recordset parameters
   oRequests.PageSize       = recordsPer
   oRequests.CacheSize      = recordsPer
   oRequests.CursorLocation = 3
else
   set oRequests = Server.CreateObject("ADODB.Recordset")
   oRequests.PageSize = 9999
end if
'dtb_debug(sSQL)
'oRequests.Open sSQL, Application("DSN"), 0, 1
'response.write "[" & lcl_useSessions & "] " & sSQL
'response.flush
 if request.cookies("user")("userid") = "1710" then
	 'response.write sSQL
	 'response.end
	 'sSQL = replace(sSQL, "ORDER BY", " AND 1=2 ORDER BY")
	
 end if
oRequests.Open sSQL, lcl_dsn, 3, 1

'Save "Last Queried" Search Options.
 lcl_success = "Y"

 if request("init") <> "Y" OR request("init") = "" OR  lcl_setLastQueryAsUserSaved = true then
		 response.write "<!--SAVED" & lcl_customreportid_actionline_lastqueried & "-->"
		 response.write "<!--" & sReportTypeHideInternal & "-->"
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectAssignedto",        selectAssignedto,        False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "orderBy",                 orderBy,                 False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "recordsPer",              recordsPer,              False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "reporttype",              reporttype,              False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "reporttype_hideinternal", sReportTypeHideInternal, False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectFormId",            selectFormId,            False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectDeptId",            selectDeptId,            False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "pastDays",                pastDays,                False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "fromDate",                fromDate,                False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "toDate",                  toDate,                  False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "fromToDateSelection",     fromToDateSection,       False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectDateType",          selectDateType,          False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusDISMISSED",         statusDISMISSED,         False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusRESOLVED",          statusRESOLVED,          False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusWAITING",           statusWAITING,           False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusINPROGRESS",        statusINPROGRESS,        False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusSUBMITTED",         statusSUBMITTED,         False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "substatus_hidden",        substatus_hidden,        False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectUserFName",         selectUserFName,         False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectUserLName",         selectUserLName,         False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectIssueStreetNumber", selectIssueStreetNumber, False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectIssueStreet",       selectIssueStreet,       False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectContactStreet",     selectContactStreet,     False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectCounty",            selectCounty,            False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectBusinessName",      selectBusinessName,      False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectTicket",            selectTicket,            False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "rockRegStreet",            rockRegStreet,            False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "ryePhraseSearch",            ryePhraseSearch,            False, lcl_success
 end if

lastTitle        = "Test"
lastDate         = "1/1/02"
lastDept         = 11798
lastDeptName     = "Test"
lastAssigned     = "bubba"
displayLastTitle = "Test"
lastSubmitted    = "bubba"
lastStatus       = "Test"

if oRequests.eof = false then
   if lcl_screen_mode <> "PRINT" then
     'SET PAGE TO VIEW
      if lcl_useSessions = 1 then
         if Len(request("pagenum")) > 0 then

           'Captures strange error with request.pagenum having multiple values
            if instr(request("pagenum"),", ") > 0 then
               lcl_comma_loc = instr(request("pagenum"),", ")
               lcl_pagenum   = mid(request("pagenum"),1,lcl_comma_loc)

               'lcl_pagenum = left(request("pagenum"),1)
            else
               lcl_pagenum = request("pagenum")
            end if

            'oRequests.AbsolutePage = clng(request("pagenum"))
            'Session("pageNum")     = clng(request("pagenum"))
            oRequests.AbsolutePage = clng(lcl_pagenum)
            Session("pageNum")     = clng(lcl_pagenum)
         else
            oRequests.AbsolutePage = 1
            Session("pageNum")     = 1
         end if
      else
        'Captures strange error with request.pagenum having multiple values
         if instr(request("pagenum"),", ") > 0 then
            lcl_comma_loc = instr(request("pagenum"),", ")
            lcl_pagenum   = mid(request("pagenum"),1,lcl_comma_loc)
            'lcl_pagenum = left(request("pagenum"),1)
         else
            lcl_pagenum = request("pagenum")
         end if

        'Determine if the pagenum is a valid number
         if dbready_number(lcl_pagenum) then
            if len(lcl_pagenum) = 0 OR clng(lcl_pagenum) < 1 then
               oRequests.AbsolutePage = 1
               session("pageNum")     = 1
            else
               if clng(lcl_pagenum) <= oRequests.PageCount then
                  oRequests.AbsolutePage = lcl_pagenum
                  session("pageNum")     = lcl_pagenum
               else
                  oRequests.AbsolutePage = 1
                  session("pageNum")     = 1
               end if
            end if
         else
            oRequests.AbsolutePage = 1
            session("pageNum")     = 1
         end if
      end if
   end if

'-------------------------------------------------------------------------------
   if lcl_screen_mode = "PRINT" then
'-------------------------------------------------------------------------------
     'Display Record Statistics
      Dim abspage, pagecnt
          abspage = oRequests.AbsolutePage
          pagecnt = oRequests.PageCount

      If request("selectAssignedto") <> "" Then
         sQueryString = replace(sQueryString,"pagenum","HFe301") 'Replace PAGENUM field with random field for navigation purposes
      Else
        'Javascript changed to keep filtering
         sQueryString = replace(sQueryString,"pagenum","HFe301")
      End If

      response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
      response.write "  <tr>" & vbcrlf
      response.write "      <td>" & vbcrlf
      response.write "          <strong><font color=""#0000ff"">" & oRequests.RecordCount  & "</font> total Action Item Requests</strong>" & vbcrlf
      response.write "          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & vbcrlf
      response.write "          <font size=""3"" color=""#3399ff""><i><strong>" & lcl_reportlabel & " REPORT</strong></i></font>" & vbcrlf
      response.write "      </td>" & vbcrlf

      if lcl_orghasfeature_issuelocation then
         response.write "      <td align=""right"">" & vbcrlf
         response.write "          <font style=""color: #ff0000;"">* <small><i>= Non-Listed Street Address</i></small></font>" & vbcrlf
         response.write "      </td>" & vbcrlf
      end if

      response.write "  </tr>" & vbcrlf
      response.write "</table>" & vbcrlf
'------------------------------------------------------------------------------
   else   'lcl_screen_mode <> "PRINT"
'------------------------------------------------------------------------------
      response.write "<font size=""3"" color=""#3399ff""><i><strong>" & lcl_reportlabel & " REPORT</strong></i></font>" & vbcrlf


      response.write "<br />Page <font color=""#0000ff"">" & oRequests.AbsolutePage & "</font>  " & vbcrlf
      response.write "of <font color=""#0000ff""> " & oRequests.PageCount & "</font> &nbsp;|&nbsp; " & vbcrlf
      response.write "<font color=""#0000ff"">" & oRequests.RecordCount & "</font> total Action Item Requests" & vbcrlf



'------------------------------------------------------------------------------
   end if
'------------------------------------------------------------------------------

   blnUpvotes = False
   intDispPage = oRequests.AbsolutePage
   if request("screen_mode") = "PRINT" then intDispPage = 1
   for intRec=1 to oRequests.PageSize
       if not oRequests.EOF then
		if clng(oRequests("upvotes")) > 0 then 
			blnUpvotes = true
		end if
       		oRequests.MoveNext
       end if
   next
   if oRequests.RecordCount > 0 then 
   	oRequests.MoveFirst
        oRequests.AbsolutePage = intDispPage
   end if

  'BEGIN: Show column headers -------------------------------------------------
   if UCASE(reporttype) = "LISTFULL" then
      response.write "<table border=""0"" cellspacing=""0"" cellpadding=""5"" class=""tablelist"">" & vbcrlf

      if lcl_screen_mode <> "PRINT" then
         response.write "  <caption>" & vbcrlf
                             display_back_next sQueryString, _
                                               session("pagenum"), _
                                               oRequests.PageCount
         response.write "  </caption>" & vbcrlf
      end if

   	  response.write "  <tr class=""tablelist"">" & vbcrlf

      if lcl_screen_mode <> "PRINT" then
         if lcl_orghasfeature_display_multiple_workorders then
            response.write "      <th><input type=""checkbox"" name=""selectAll"" onclick=""checkedAll();""></th>" & vbcrlf
         end if
      end if

	     response.write "      <th colspan=""100"">&nbsp;</th>" & vbcrlf
      response.write "  </tr>" & vbcrlf

   else

     'Display FORWARD and BACKWARD results navigation and PRINTER FRIENDLY buttons
      if lcl_screen_mode <> "PRINT" then
         display_back_next sQueryString, _
                           session("pagenum"), _
                           oRequests.PageCount
      end if

      response.write "<table border=""0"" cellspacing=""0"" cellpadding=""5"" class=""tablelist"" width=""100%"">" & vbcrlf
      response.write "  <tr valign=""bottom"" class=""tablelist"">" & vbcrlf

      if lcl_screen_mode <> "PRINT" then
         if lcl_orghasfeature_display_multiple_workorders then
            response.write "    <th><input type=""checkbox"" name=""selectAll"" onclick=""checkedAll();""></th>" & vbcrlf
         end if
      end if

      response.write "    <th>Action Line Form</th>" & vbcrlf
      if blnUpvotes then response.write "    <th>Upvotes</th>" & vbcrlf
      response.write "    <th>Date submitted</th>" & vbcrlf

      if lcl_orghasfeature_actionline_maintain_duedate AND lcl_userhaspermission_actionline_maintain_duedate then
         response.write "    <th>Due Date</th>" & vbcrlf
      end if

      if reporttype = "Detail" or reporttype = "DrillThru" then
         response.write " <th>Date Completed</th>" & vbcrlf
         response.write " <th>Days open*/To complete</th>" & vbcrlf
      end if

     'Status/Sub-Status
      response.write "    <th>" & vbcrlf
      response.write "        Status" & vbcrlf

      if lcl_userhaspermission_action_line_substatus then
         response.write "<br />[Sub-Status]"
      end if

      response.write "    </th>" & vbcrlf

	     'if lcl_userhaspermission_action_line_substatus then
      '   response.write "    <th>Sub-Status</th>" & vbcrlf
      'end if

   	  response.write "    <th>Submitted by</th>" & vbcrlf
      response.write "    <th>Contact<br />Street Name</th>" & vbcrlf
      response.write "    <th>Assigned to</th>" & vbcrlf
   	  response.write "    <th>Department</th>" & vbcrlf

      if lcl_orghasfeature_issuelocation then
         response.write "    <th>Issue/Problem Location<br />Street Name</th>" & vbcrlf
      end if

      if iSelectBusinessName <> "all" AND iSelectBusinessName <> "" then
         response.write "    <th style=""white-space:nowrap; text-align:left;"">Business Name</th>" & vbcrlf
      end if

      lcl_showPastDays = "N"

   	  if pastDays <> "all" then
         lcl_showPastDays = "Y"
      elseif lcl_screen_mode = "PRINT" AND pastDays = "0" then
         lcl_showPastDays = "Y"
   	  end if

      if lcl_showPastDays = "Y" AND ucase(reporttype) <> "DETAIL" AND ucase(reporttype) <> "DRILLTHRU" then
         if iSearchDaysType = "PAST" then
         		 response.write "    <th>Days<br />Past<br />Due</th>" & vbcrlf
         else
         		 response.write "    <th>Days<br />Open*/To<br />complete</th>" & vbcrlf
         end if
      end if

      response.write "  </tr>" & vbcrlf
   end if
  'END: Show column headers ---------------------------------------------------

  'Display Grand Total Row
   displayGrandTotalLine reporttype, varWhereClause, oRequests.RecordCount

  'BEGIN: Loop and display records --------------------------------------------
   bgcolor   = "#eeeeee"
   totalPast = 0

   for intRec=1 to oRequests.PageSize
       if not oRequests.EOF then

         'Setup variables
          bgcolor    = changeBGColor(bgcolor,"#eeeeee","#ffffff")
          sTitle     = valueExistsOrRedQuestionMarks(oRequests("action_formTitle"))
          sDate      = oRequests("submit_date")
          sDate      = formatDateTime(sDate,vbShortDate)
          sDueDate   = oRequests("due_date")
          'sDueDate   = formatDateTime(sDueDate,vbShortDate)
          sAssigned  = valueExistsOrRedQuestionMarks(oRequests("assigned_Name"))
          sSubmitted = valueExistsOrRedQuestionMarks(oRequests("userfname") & " " & oRequests("userlname"))
          sStatus    = valueExistsOrRedQuestionMarks(oRequests("status"))
          sDept      = oRequests("deptId")
          sDeptName  = trim(oRequests("deptName"))

         'The department has been deleted
          if sDept <> "" then
             if clng(sDept) > clng(0) AND sDeptName = ""then
                sDeptName = "Department<br />has been<br />deleted"
             end if
          end if

          if lcl_userhaspermission_action_line_substatus then
             'sSubStatus = oRequests("sub_status_name")
             sSubStatus = oRequests("sub_status_desc")
          else
		           sSubStatus = ""
          end if

          datSubmitDate  = ""
          datResolveDate = ""
          datDueDate     = ""

          if oRequests("complete_date") <> "" Then
             datResolveDate = oRequests("complete_date")
             datResolveDate = formatdatetime(datResolveDate,vbShortDate)
          end if

          if oRequests("due_date") <> "" Then
             datDueDate = oRequests("due_date")
             datDueDate = formatdatetime(datDueDate,vbShortDate)
          end if

          if (iSearchDaysType = "PAST" AND pastDays <> "all") OR _
             (iSearchDaysType = "OPEN" AND pastDays <> "all" AND pastDays <> "0") then

       			'if pastDays <> "all" and pastDays <> "0" then
       						numPast         = oRequests("numPast")
             daysPastDueDate = oRequests("daysPastDueDate")

             if iSearchDaysType = "PAST" then
                lcl_daysover = daysPastDueDate
             else
                lcl_daysover = numPast
             end if

	  			   	   totalPast = totalPast + lcl_daysover
       			end if

          datSubmitDate     = valueExistsOrRedQuestionMarks(oRequests("submit_date"))
          datResolveDate    = valueExistsOrRedQuestionMarks(datResolveDate)
          datDueDate        = valueExistsOrRedQuestionMarks(oRequests("due_date"))
          lngTrackingNumber = oRequests("action_autoid") & replace(FormatDateTime(oRequests("submit_date"),4),":","")

         'BEGIN: Display data rows --------------------------------------------
          if UCASE(reporttype) = "LISTFULL" then
		           Dim lcl_allowed_days, lcl_total_days, lcl_total_days_label
             if oRequests("usesafter5adjustment") = True then
			             if oRequests("usesweekdays") = True then
                   'lcl_total_days = oRequests("lcl_subTotal_totalDaysOpen_noweekends_adjusted")
                   lcl_total_days = oRequests("subTotal_totalDaysOpen_noweekends_adjusted")
				            else
                   'lcl_total_days = oRequests("lcl_subTotal_totalDaysOpen_adjusted")
                   lcl_total_days = oRequests("subTotal_totalDaysOpen_adjusted")
		 	           	end if
             else
		         	    if oRequests("usesweekdays") = True then
                   'lcl_total_days = oRequests("lcl_subTotal_totalDaysOpen_noweekends")
                   lcl_total_days = oRequests("subTotal_totalDaysOpen_noweekends")
	            			else
                   'lcl_total_days = oRequests("lcl_subTotal_totalDaysOpen")
                   lcl_total_days = oRequests("subTotal_totalDaysOpen")
            				end if
          		 end if

      		     if oRequests("allowedunresolveddays") > 0 then
                lcl_allowed_days = oRequests("allowedunresolveddays")
             else
         			    lcl_allowed_days = 0
         			 end if

         			 if IsNULL(oRequests("complete_date")) then
            				lcl_total_days_label = "Days Open"
          		 else
            				lcl_total_days_label = "Days taken to Complete"
         			 end if

             lcl_total_days_past_due       = 0
             lcl_total_days_past_due_label = ""

             if lcl_orghasfeature_actionline_maintain_duedate AND lcl_userhaspermission_actionline_maintain_duedate then
                if oRequests("due_date") <> "" then
                   if datediff("d",oRequests("due_date"),date()) >= 0 then
                      lcl_total_days_past_due = datediff("d",oRequests("due_date"),date())
                   end if
                end if

                lcl_total_days_past_due_label = "Days Past Due"
             end if

            'Build the USER City/State/Zip display value
             lcl_user_csz = buildCityStateZip(oRequests("usercity"), oRequests("userstate"), oRequests("userzip"))

            'Build the ISSUE LOCATION City/State/Zip display value
             lcl_issue_csz = buildCityStateZip(oRequests("city"), oRequests("state"), oRequests("zip"))

            'Setup the url to the Request Manager screen for the row.
             lcl_link_str = ""

             if request.querystring <> "" then
                lcl_link_str = request.querystring & "&"
             end if

             lcl_link_str = lcl_link_str & "control=" & oRequests("action_autoid")

            'Setup the javascript events for the row.
             lcl_row_onmouseover = " onMouseOver=""changeRowColor('row_" & oRequests("action_autoid") & "','OVER')"""
             lcl_row_onmouseout  = " onMouseOut=""changeRowColor('row_" & oRequests("action_autoid") & "','OUT')"""
             lcl_row_onclick     = " onClick=""location.href='action_respond.asp?" & lcl_link_str & "';"""

             response.write "<tr valign=""top"" class=""handhover"" bgcolor=""" & bgcolor & """ id=""row_" & oRequests("action_autoid") & """" & lcl_row_onmouseover & lcl_row_onmouseout & ">" & vbcrlf
             if lcl_screen_mode <> "PRINT" then
                if lcl_orghasfeature_display_multiple_workorders then
                   response.write "    <td><input type=""checkbox"" name=""p_action_autoid_" & oRequests("action_autoid") & """ value=""" & oRequests("action_autoid") & """></td>" & vbcrlf
                end if
             end if
             response.write "    <td " & lcl_row_onclick & " colspan=""100"">" & vbcrlf
             response.write "        <table border=""0"" cellspacing=""0"" cellpadding=""5"" class=""tablelist"">" & vbcrlf
             response.write "          <tr>" & vbcrlf
             response.write "              <td><font size=""3""><strong>" & lngTrackingNumber & "</strong></font></td>" & vbcrlf
             response.write "              <td><strong>Submitted Date: </strong>" & formatdatetime(datSubmitDate,vbShortDate)              & "</td>" & vbcrlf

             datDueDateLabel = "&nbsp;"
             datDueDate      = ""
         
             if lcl_orghasfeature_actionline_maintain_duedate AND lcl_userhaspermission_actionline_maintain_duedate then
                datDueDateLabel = "<strong>Due Date: </strong>"

                if oRequests("due_date") <> "" then
                   datDueDate = formatdatetime(oRequests("due_date"),vbShortDate)
                   datDueDate = valueExistsOrRedQuestionMarks(datDueDate)
                end if

                response.write "              <td>" & datDueDateLabel & datDueDate & "</td>" & vbcrlf

             else
                response.write "              <td>&nbsp;</td>" & vbcrlf
             end if

             response.write "              <td><strong>Status: </strong>"       & UCASE(sStatus) & "</td>" & vbcrlf

			          if lcl_userhaspermission_action_line_substatus then
		              response.write "           <td><strong>Sub-Status: </strong>" & sSubStatus & "</td>" & vbcrlf
             end if

             response.write "              <td><strong>Form: </strong>"        & oRequests("action_FormTitle")                             & "</td>" & vbcrlf
             response.write "          </tr>" & vbcrlf
             response.write "          <tr>" & vbcrlf
             response.write "              <td><strong>Submitted By: </strong>"    & oRequests("userfname") & " " & oRequests("userlname") & "</td>" & vbcrlf
             response.write "              <td><strong>Phone: </strong>"           & FormatPhoneNumber(oRequests("userhomephone"))         & "</td>" & vbcrlf
             response.write "              <td><strong>Contact Address: </strong>" & oRequests("useraddress")                              & "</td>" & vbcrlf
             response.write "              <td colspan=""3""><strong>Contact City/State/Zip: </strong>" & lcl_user_csz                     & "</td>" & vbcrlf
             response.write "          </tr>" & vbcrlf

             if oRequests("userbusinessname") <> "" then
                response.write "          <tr>" & vbcrlf
                response.write "              <td><strong>Business Name: </strong>" & oRequests("userbusinessname") & "</td>" & vbcrlf
                response.write "              <td colspan=""6"">&nbsp;</td>" & vbcrlf
                response.write "          </tr>" & vbcrlf
             end if

             response.write "          <tr valign=""top"">" & vbcrlf
             response.write "              <td><strong>Assigned To: </strong>"  & oRequests("assigned_Name")                   & "</td>" & vbcrlf
             response.write "              <td><strong>Department: </strong>"   & sDeptName                                    & "</td>" & vbcrlf
             response.write "              <td><strong>Days Allowed: </strong>" & lcl_allowed_days                             & " day(s)</td>" & vbcrlf
             response.write "              <td><strong>" & lcl_total_days_label & ": </strong>" & lcl_total_days & " day(s)</td>" & vbcrlf

             if lcl_orghasfeature_actionline_maintain_duedate AND lcl_userhaspermission_actionline_maintain_duedate then
                response.write "              <td colspan=""2""><strong>Days Past Due: </strong>" & lcl_total_days_past_due & " day(s)</td>" & vbcrlf
             else
                response.write "              <td colspan=""2"">&nbsp;</td>" & vbcrlf
             end if

             response.write "          </tr>" & vbcrlf

             if lcl_orghasfeature_issuelocation then
               'Set up the Issue/Problem Location label
                sIssueName = oRequests("issuelocationname")
                If Trim(sIssueName) = "" OR IsNull(sIssueName) Then
                   sIssueName = "Issue/Problem Location:"
                End If

               'Determine if the issue/problem location address is a valid address or not
                if oRequests("validstreet") <> "Y" AND oRequests("action_form_display_issue") then
                   lcl_valid_street = "<font style=""color:#ff0000;"">&nbsp;*</font>"
                else
                   lcl_valid_street = ""
                end if

             response.write "          <tr><td colspan=""3""><strong>" & sIssueName & " </strong>" & oRequests("streetname") & lcl_valid_street & "</td>" & vbcrlf
             else
             response.write "          <tr><td colspan=""3"">&nbsp;</td>" & vbcrlf
         			 end if
             response.write "              <td colspan=""3""><strong>City/State/Zip: </strong>" & lcl_issue_csz & "</td>" & vbcrlf
             response.write "          </tr>" & vbcrlf

             response.write "          <tr>" & vbcrlf
             response.write "              <td colspan=""6""><strong>Additional Info: </strong>" & oRequests("comments") & "</td>" & vbcrlf
             response.write "          </tr>" & vbcrlf
             response.write "        </table>" & vbcrlf
             response.write "        <p>" & vbcrlf
             response.write "        <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf

         			'Format the egov_action_line_request.comment
             dim lcl_comment, arrQues, iQues, sQues, sAnswer, sValue
             sValue = oRequests("comment")

            'Format and split the questions and answers
             sValue = replace(sValue,"<B>","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>")
             sValue = replace(sValue,"<b>","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>")

   	         lcl_comment = remove_html_tags(sValue)

         			 response.write "          <tr><td><hr size=""1"" width=""100%""></td></tr>" & vbcrlf
         			 response.write "          <tr><td>" & vbcrlf
         			 response.write "                  <table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
         			 response.write "                    <tr valign=""top"">" & vbcrlf
         			 response.write "                        <td>&nbsp;&nbsp;</td>" & vbcrlf
         			 response.write "                        <td>" & lcl_comment & "</td></tr>" & vbcrlf
         			 response.write "                  </table><p>" & vbcrlf
         			 response.write "              </td></tr>" & vbcrlf
             response.write "        </table>" & vbcrlf

            'Display the Request Activity Log
             if not lcl_userhaspermission_actionline_hide_requestlog then
                List_Comments oRequests("action_autoid"), _
                              sReportTypeHideInternal
             end if

             response.write "        </p>" & vbcrlf
             response.write "    </td>" & vbcrlf
             response.write "  </tr>" & vbcrlf

'------------------------------------------------------------------------------
          else  'ReportTypes NOT EQUAL to LISTFULL
'------------------------------------------------------------------------------
            'BEGIN: Insert sub-total row if Detail Report ---------------------
             if UCASE(reporttype) = "DETAIL" OR UCASE(reporttype) = "DRILLTHRU" then
                if orderBy = "submit_date" then
                   if DateDiff("d",sDate,lastDate) = 0 then
                     'NO NEW LINE
                   else
                      if lastDate <> "1/1/02" then
                         getAverages lcl_subTotal_totalRequestsOpen, lcl_subTotal_totalDaysOpen, lcl_subTotal_totalRequestsClosed, lcl_subTotal_totalDaysClosed, avOpenTotal, avClosedTotal

                         displayLastTitle     = sDate
                         displayLastAvgOpen   = avOpenTotal
                         displayLastAvgClosed = avClosedTotal

                        'Display the Sub-Total details
                         displayTotalRow "SUBTOTAL", "", lastDate, lcl_subTotal_totalRequests, lcl_subTotal_totalRequestsOpen, avOpenTotal, avClosedTotal

                      end if
                   end if

                elseif orderBy = "due_date" then
                   if DateDiff("d",sDueDate,lastDate) = 0 then
                     'NO NEW LINE
                   else
                      if lastDate <> "1/1/02" then
                         getAverages lcl_subTotal_totalRequestsOpen, lcl_subTotal_totalDaysOpen, lcl_subTotal_totalRequestsClosed, lcl_subTotal_totalDaysClosed, avOpenTotal, avClosedTotal

                         displayLastTitle     = sDueDate
                         displayLastAvgOpen   = avOpenTotal
                         displayLastAvgClosed = avClosedTotal

                        'Display the Sub-Total details
                         displayTotalRow "SUBTOTAL", "", lastDate, lcl_subTotal_totalRequests, lcl_subTotal_totalRequestsOpen, avOpenTotal, avClosedTotal

                      end if
                   end if

                elseif orderBy = "action_Formid" then
                   if sTitle = lastTitle then
                     'NO NEW LINE
                   else
                      if lastTitle <> "Test" then
                         getAverages lcl_subTotal_totalRequestsOpen, lcl_subTotal_totalDaysOpen, lcl_subTotal_totalRequestsClosed, lcl_subTotal_totalDaysClosed, avOpenTotal, avClosedTotal

                         displayLastTitle     = sTitle
                         displayLastAvgOpen   = avOpenTotal
                         displayLastAvgClosed = avClosedTotal

                        'Display the Sub-Total details
                         displayTotalRow "SUBTOTAL", "", lastTitle, lcl_subTotal_totalRequests, lcl_subTotal_totalRequestsOpen, avOpenTotal, avClosedTotal

                      end if
                   end if
			             elseif orderBy = "deptId" then
                   if sDept = lastDept then
                     'NO NEW LINE
                   else
                      if lastDept <> 11798 then
                         getAverages lcl_subTotal_totalRequestsOpen, lcl_subTotal_totalDaysOpen, lcl_subTotal_totalRequestsClosed, lcl_subTotal_totalDaysClosed, avOpenTotal, avClosedTotal

                         displayLastTitle     = sDeptName
                         displayLastAvgOpen   = avOpenTotal
                         displayLastAvgClosed = avClosedTotal

                        'Display the Sub-Total details
                         displayTotalRow "SUBTOTAL", "", lastDeptName, lcl_subTotal_totalRequests, lcl_subTotal_totalRequestsOpen, avOpenTotal, avClosedTotal

                      end if
                   end if
                elseif orderBy = "assigned_Name" then
                   if sAssigned = lastAssigned then
                     'NO NEW LINE
                   else
                      if lastAssigned <> "bubba" then						
                         getAverages lcl_subTotal_totalRequestsOpen, lcl_subTotal_totalDaysOpen, lcl_subTotal_totalRequestsClosed, lcl_subTotal_totalDaysClosed, avOpenTotal, avClosedTotal

                         displayLastTitle     = sAssigned
                         displayLastAvgOpen   = avOpenTotal
                         displayLastAvgClosed =	avClosedTotal		

                        'Display the Sub-Total details
                         displayTotalRow "SUBTOTAL", "", lastAssigned, lcl_subTotal_totalRequests, lcl_subTotal_totalRequestsOpen, avOpenTotal, avClosedTotal

                      end if
                   end if
                elseif orderBy = "submittedby" then
                   if sSubmitted = lastSubmitted then
                     'NO NEW LINE
                   else
                      if lastSubmitted <> "bubba" then
                         getAverages lcl_subTotal_totalRequestsOpen, lcl_subTotal_totalDaysOpen, lcl_subTotal_totalRequestsClosed, lcl_subTotal_totalDaysClosed, avOpenTotal, avClosedTotal

                         displayLastTitle     = sAssigned
                         displayLastAvgOpen   = avOpenTotal
                         displayLastAvgClosed =	avClosedTotal		

                        'Display the Sub-Total details
                         displayTotalRow "SUBTOTAL", "", lastSubmitted, lcl_subTotal_totalRequests, lcl_subTotal_totalRequestsOpen, avOpenTotal, avClosedTotal

                      end if
                   end if
                elseif orderBy = "status" then
                   if sStatus = lastStatus then
                     'NO NEW LINE
                   else
                      if lastStatus <> "Test" then
                         getAverages lcl_subTotal_totalRequestsOpen, lcl_subTotal_totalDaysOpen, lcl_subTotal_totalRequestsClosed, lcl_subTotal_totalDaysClosed, avOpenTotal, avClosedTotal

                         displayLastTitle     = sStatus
                         displayLastAvgOpen   = avOpenTotal
                         displayLastAvgClosed =	avClosedTotal		

                        'Display the Sub-Total details
                         displayTotalRow "SUBTOTAL", "", lastStatus, lcl_subTotal_totalRequests, lcl_subTotal_totalRequestsOpen, avOpenTotal, avClosedTotal

                      end if
                   end if
                end if
             end if		

             lcl_link_str        = ""
             lcl_row_onmouseover = ""
             lcl_row_onmouseout  = ""
             lcl_row_onclick     = ""

             if lcl_screen_mode <> "PRINT" then 
               'Setup the url to the Request Manager screen for the row.
                if request.querystring <> "" then
                   lcl_link_str = request.querystring & "&"
                end if

                lcl_link_str = lcl_link_str & "control=" & oRequests("action_autoid")

               'Setup the javascript events for the row.
                lcl_row_onmouseover = " onMouseOver=""changeRowColor('row_" & oRequests("action_autoid") & "','OVER')"""
                lcl_row_onmouseout  = " onMouseOut=""changeRowColor('row_" & oRequests("action_autoid") & "','OUT')"""
                lcl_onclick         = " onClick=""openRequestManager('" & lcl_link_str & "');"""
                lcl_td_onclick      = ""
                lcl_row_onclick     = ""

               'If the org has the "display_multiple_workorders" feature then the "onclick" will need to be on each <TD>.
               'If it is on the row then when the user clicks on the checkbox option it will open the record.
               'If the org does NOT have the feature then put the onclick on the row.
                if lcl_orghasfeature_display_multiple_workorders then
                   lcl_td_onclick  = lcl_onclick
                else
                   lcl_row_onclick = lcl_onclick
                end if

             end if

             response.write "<tr bgcolor=""" & bgcolor & """ class=""handhover"" id=""row_" & oRequests("action_autoid") & """ align=""center""" & lcl_row_onmouseover & lcl_row_onmouseout & lcl_row_onclick & ">" & vbcrlf

             if lcl_screen_mode <> "PRINT" then
                if lcl_orghasfeature_display_multiple_workorders then
                   response.write "    <td><input type=""checkbox"" name=""p_action_autoid_" & oRequests("action_autoid") & """ value=""" & oRequests("action_autoid") & """></td>" & vbcrlf
                end if
             end if

             response.write "    <td" & lcl_td_onclick & " align=""left""><strong>(" & lngTrackingNumber & ") " & sTitle & "</strong></td>" & vbcrlf
             if blnUpvotes then response.write "    <td" & lcl_td_onclick & " align=""center""><strong>" & oRequests("upvotes") & "</strong></td>" & vbcrlf
             response.write "    <td" & lcl_td_onclick & ">" & formatdatetime(datSubmitDate,vbShortDate) & "</td>" & vbcrlf

             if lcl_orghasfeature_actionline_maintain_duedate AND lcl_userhaspermission_actionline_maintain_duedate then
                lcl_display_duedate = "&nbsp;"

                if oRequests("due_date") <> "" then
                   lcl_display_duedate = formatdatetime(datDueDate,vbShortDate)
                end if

                response.write "    <td" & lcl_td_onclick & ">" & lcl_display_duedate & "</td>" & vbcrlf

             end if

             if UCASE(reporttype) = "DETAIL" OR UCASE(reporttype) = "DRILLTHRU" then
                response.write "<td" & lcl_td_onclick & " align=""center"">" & datResolveDate & "</td>" & vbcrlf
                response.write "<td" & lcl_td_onclick & " align=""center"">" & vbcrlf

                'if oRequests("lcl_subTotal_totalDaysOpen")<> "" then
                '   response.write oRequests("lcl_subTotal_totalDaysOpen") & " days</td>" & vbcrlf
                '   countDays = clng(oRequests("lcl_subTotal_totalDaysOpen"))

                if oRequests("subTotal_totalDaysOpen")<> "" then
		   if oRequests("usesweekdays") = True then
                   	response.write oRequests("subTotal_totalDaysOpen_noweekends") & " days</td>" & vbcrlf
                   	countDays = clng(oRequests("subTotal_totalDaysOpen_noweekends"))
	           else
                   	response.write oRequests("subTotal_totalDaysOpen") & " days</td>" & vbcrlf
                   	countDays = clng(oRequests("subTotal_totalDaysOpen"))
            	   end if
                else
                   openDays = dateDiff("d",oRequests("submit_date"),Now)
                   response.write openDays & " days<font style=""color:#ff0000"">* </font></td>" & vbcrlf
                   countDays = clng(openDays)
                end if
             end if

             response.write "    <td" & lcl_td_onclick & ">" & vbcrlf
             response.write          UCASE(sStatus) & vbcrlf

             if lcl_userhaspermission_action_line_substatus then
                if sSubStatus <> "" then
   		              response.write "<br /><span style=""color:#ff0000"">[" & sSubStatus & "]</span>" & vbcrlf
                end if
             end if

             response.write "    </td>" & vbcrlf

   			       'if lcl_userhaspermission_action_line_substatus then
		           '   response.write " <td" & lcl_td_onclick & ">" & sSubStatus & "</td>" & vbcrlf
			          'end if

   			       response.write "    <td" & lcl_td_onclick & ">" & oRequests("userfname") & " " & oRequests("userlname") & "</td>" & vbcrlf
             response.write "    <td" & lcl_td_onclick & " align=""left"">" & oRequests("useraddress")               & "</td>" & vbcrlf
             response.write "    <td" & lcl_td_onclick & ">" & oRequests("assigned_Name")                            & "</td>" & vbcrlf
             response.write "    <td" & lcl_td_onclick & ">" & sDeptName                                             & "</td>" & vbcrlf

             if lcl_orghasfeature_issuelocation then
                if oRequests("validstreet") <> "Y" AND oRequests("action_form_display_issue") then
                   lcl_valid_street = "<font style=""color:#ff0000;"">&nbsp;*</font>"
                else
                   lcl_valid_street = ""
                end if

                'response.write "<td>" & oRequests("streetname") & "<br />[" & oRequests("sortstreetname") & "] " & lcl_valid_street & "</td>" & vbcrlf
                response.write "<td" & lcl_td_onclick & " align=""left"">" & oRequests("streetname") & lcl_valid_street & "</td>" & vbcrlf
             end if

             if iSelectBusinessName <> "all" AND iSelectBusinessName <> "" then
                response.write "<td" & lcl_td_onclick & " align=""left"">" & oRequests("userbusinessname") & "</td>" & vbcrlf
             end if

             if lcl_showPastDays = "Y" AND ucase(reporttype) <> "DETAIL" AND ucase(reporttype) <> "DRILLTHRU" then
                'response.write "    <td" & lcl_td_onclick & ">[" & numPast & "]</td>" & vbcrlf
                response.write "    <td" & lcl_td_onclick & ">[" & lcl_daysover & "]</td>" & vbcrlf
             end if

             response.write "</tr>" & vbcrlf

          end if
         'END: Display data row -----------------------------------------------

         'BEGIN: Track the counts for each row --------------------------------
         'Depending on the group by (orderby) selected will determine what the "current value" is and what the "previous value" was.
         'The "current value" and "previous value" variables are used within the "trackSubTotals" procedure to properly calculate the
         '  subtotal row for the group.
          if UCASE(reporttype) = "DETAIL" OR UCASE(reporttype) = "DRILLTHRU" then
             if orderBy = "submit_date" then
                lcl_currentvalue  = sDate
                lcl_previousvalue = lastDate
                lastDate          = sDate
             elseif orderBy = "due_date" then
                lcl_currentvalue  = sDueDate
                lcl_previousvalue = lastDate
                lastDate          = sDueDate
             elseif orderBy = "action_Formid" then
                lcl_currentvalue  = sTitle
                lcl_previousvalue = lastTitle
                lastTitle         = sTitle
             elseif orderBy = "deptId" then
                lcl_currentvalue  = sDept
                lcl_previousvalue = lastDept
                lastDept          = sDept
                lastDeptName      = sDeptName
             elseif orderBy = "assigned_Name" then
                lcl_currentvalue  = sAssigned
                lcl_previousvalue = lastAssigned
                lastAssigned      = sAssigned 
             elseif orderBy = "submittedby" then
                lcl_currentvalue  = sSubmitted
                lcl_previousvalue = lastSubmitted
                lastSubmitted     = sSubmitted 
             elseif orderBy = "status" then
                lcl_currentvalue  = sStatus
                lcl_previousvalue = lastStatus
                lastStatus        = sStatus
             end if

            'Track each row in a group and calculate the subtotals
             trackSubTotals orderBy, _
                            sStatus, _
                            lcl_currentvalue, _
                            lcl_previousvalue, _
                            lcl_subTotal_totalRequests, _
                            lcl_subTotal_totalDaysOpen, _
                            lcl_subTotal_totalRequestsOpen, _
                            lcl_subTotal_totalDaysClosed, _
                            lcl_subTotal_totalRequestsClosed, _
                            countDays, _
                            lcl_subTotal_totalRequests, _
                            lcl_subTotal_totalDaysOpen, _
                            lcl_subTotal_totalRequestsOpen, _
                            lcl_subTotal_totalDaysClosed, _
                            lcl_subTotal_totalRequestsClosed
          end if
         'END: Track the counts for each row ----------------------------------

		response.flush
	      oRequests.MoveNext 

      end if
   next

   if UCASE(reporttype) = "DETAIL" OR UCASE(reporttype) = "DRILLTHRU" then
      lcl_searchtype = getSearchType(orderBy)

      'retrieveOpenClosedCounts "OPEN", orderBy, lcl_searchtype, varWhereClause, lcl_os_num_open, lcl_os_total_days_open, lcl_cs_num_closed, lcl_cs_total_days_closed
      'retrieveOpenClosedCounts "CLOSED", orderBy, lcl_searchtype, varWhereClause, lcl_os_num_open, lcl_os_total_days_open, lcl_cs_num_closed, lcl_cs_total_days_closed

      getAverages lcl_subTotal_totalRequestsOpen, lcl_subTotal_totalDaysOpen, lcl_subTotal_totalRequestsClosed, lcl_subTotal_totalDaysClosed, avOpenTotal, avClosedTotal

      'Track each row in a group and calculate the subtotals
       'trackSubTotals orderBy, sStatus, lcl_currentvalue, lcl_previousvalue, lcl_subTotal_totalRequests, lcl_subTotal_totalDaysOpen, _
       '               lcl_subTotal_totalRequestsOpen, countDays, lcl_subTotal_totalRequests, lcl_subTotal_totalDaysOpen, _
       '               lcl_subTotal_totalRequestsOpen

      'getAverages lcl_os_num_open, lcl_subTotal_totalDaysOpen, lcl_cs_num_closed, lcl_cs_total_days_closed, avOpenTotal, avClosedTotal

     'Display the Sub-Total details
      displayTotalRow "SUBTOTAL", "", displayLastTitle, lcl_subTotal_totalRequests, lcl_subTotal_totalRequestsOpen, avOpenTotal, avClosedTotal

   end if

   response.write "</table>" & vbcrlf

  'DISPLAY FORWARD AND BACKWARD NAVIGATION BOTTOM
   if lcl_screen_mode <> "PRINT" then

   response.write "<div id=""bottom"">" & vbcrlf
      display_back_next sQueryString, _
                        session("pagenum"), _
                        oRequests.PageCount
   response.write "</div>" & vbcrlf
   end if
else
   response.write "<p><strong>No records found</strong></p>" & vbcrlf
end if

oRequests.close
set oRequests = nothing

end sub

'------------------------------------------------------------------------------
function List_Comments(iID, iReportTypeHideInternal)

 dim lcl_display_row, lcl_display_notetocitizen, lcl_display_action_citizen, lcl_display_internal_comment
 dim lcl_reporttypehideinternal

 lcl_display_row              = ""
 lcl_display_notetocitizen    = ""
 lcl_display_action_citizen   = ""
 lcl_display_internal_comment = ""
 lcl_reporttypehideinternal   = ""

 if iReportTypeHideInternal <> "" then
    lcl_reporttypehideinternal = ucase(iReportTypeHideInternal)
 end if

'	sSQL = "SELECT * FROM egov_action_responses  LEFT OUTER JOIN egov_users ON egov_action_responses.action_userid = egov_users.userid LEFT OUTER JOIN users on egov_action_responses.action_userid=users.userid where action_autoid=" & iID & " ORDER BY action_editdate DESC"
	sSQL = "SELECT *, "
 sSQL = sSQL & " es.status_name AS sub_status_name "
	sSQL = sSQL & " FROM egov_action_responses egr "
	sSQL = sSQL & " LEFT OUTER JOIN egov_users ON egr.action_userid = egov_users.userid "
	sSQL = sSQL & " LEFT OUTER JOIN users ON egr.action_userid = users.userid "
	sSQL = sSQL & " LEFT OUTER JOIN egov_actionline_requests_statuses AS es "
	sSQL = sSQL &               "ON egr.action_sub_status_id = es.action_status_id "
	sSQL = sSQL & " WHERE egr.action_autoid = " & iID
	sSQL = sSQL & " ORDER BY egr.action_editdate DESC"

	Set oCommentList = Server.CreateObject("ADODB.Recordset")
	'oCommentList.Open sSQL, Application("DSN"), 0, 1
	oCommentList.Open sSQL, lcl_dsn, 3, 1

 sBGColor = "#E0E0E0"

	if not oCommentList.eof then
  		do while not oCommentList.eof

       lcl_display_row              = ""
       lcl_display_notetocitizen    = ""
       lcl_display_action_citizen   = ""
       lcl_display_internal_comment = ""
       lcl_comments_firstname       = oCommentList("firstname")
       lcl_comments_lastname        = oCommentList("lastname")
		     lcl_substatus_name           = oCommentList("sub_status_name")
		   
  		   if lcl_substatus_name <> "" then
		  	     lcl_substatus_name = " <i>(" & lcl_substatus_name & ")</i>"
		     end if

    			if oCommentList("action_externalcomment") <> "" then
          lcl_action_externalcomment = replace(oCommentList("action_externalcomment"),"default_novalue","")

         'Determine if an email was sent to the citizen.
          lcl_citizen_sentby_id      = "NULL"
          lcl_citizen_sentby_name    = "NULL"
          lcl_citizen_sentto_id      = "NULL"
          lcl_citizen_sentto_email   = "NULL"
          lcl_citizen_emailsent_date = "NULL"

          if oCommentList("citizen_sentto_email") <> "" then
             lcl_citizen_sentby_id      = oCommentList("citizen_sentby_id")
             lcl_citizen_sentby_name    = oCommentList("citizen_sentby_name")
             lcl_citizen_sentto_id      = oCommentList("citizen_sentto_id")
             lcl_citizen_sentto_email   = oCommentList("citizen_sentto_email")
             lcl_citizen_emailsent_date = oCommentList("citizen_emailsent_date")

            'Set up the "Note to Citizen" email in Activity Log comment
             if lcl_action_externalcomment <> "" then
                lcl_action_externalcomment = lcl_action_externalcomment & "<br />"
             end if

             lcl_action_externalcomment = lcl_action_externalcomment & lcl_citizen_sentby_name & " sent a notification of this request to contact "
             lcl_action_externalcomment = lcl_action_externalcomment & lcl_citizen_sentto_email
             lcl_action_externalcomment = lcl_action_externalcomment & " on " & lcl_citizen_emailsent_date

          end if

          lcl_display_notetocitizen = "<br />&nbsp;&nbsp;&nbsp;<strong>Note to Citizen: </strong><i>" & lcl_action_externalcomment  & "</i>" & vbcrlf
  		  	end if

    			if oCommentList("action_citizen") <> "" then
          lcl_display_action_citizen = "<br />&nbsp;&nbsp;&nbsp;<strong>" & oCommentList("userfname")  & " " & oCommentList("userlname") & " : </strong><i>" & oCommentList("action_citizen") & "</i>" & vbcrlf
  		  	end if

       if lcl_reporttypehideinternal <> "YES" then
       			if oCommentList("action_internalcomment") <> "" then
             lcl_display_internal_comment = "<br />&nbsp;&nbsp;&nbsp;<strong>Internal Note: </strong><i>" & oCommentList("action_internalcomment")  & "</i>" & vbcrlf
     		  	end if
       end if

      'Build the Activity Log row.  Display the row ONLY if data exists.
       if lcl_display_notetocitizen <> "" OR lcl_display_action_citizen <> "" OR lcl_display_internal_comment <> "" then
       			lcl_display_row = "<div style=""border-bottom:solid 1px #000000;background-color:" & sBGColor & """>" & vbcrlf
       			lcl_display_row = lcl_display_row & "<table border=""0"" cellspacing=""0"" cellpadding=""5"">" & vbcrlf
          lcl_display_row = lcl_display_row & "  <tr><td>" & lcl_comments_firstname & " " & lcl_comments_lastname & " - " & UCASE(oCommentList("action_status")) & lcl_substatus_name & " - " &  oCommentList("action_editdate") & vbcrlf

          if lcl_display_notetocitizen <> "" then
             lcl_display_row = lcl_display_row & lcl_display_notetocitizen
          end if

          if lcl_display_action_citizen <> "" then
             lcl_display_row = lcl_display_row & lcl_display_action_citizen
          end if

          if lcl_display_internal_comment <> "" then
             lcl_display_row = lcl_display_row & lcl_display_internal_comment
          end if

          lcl_display_row = lcl_display_row & "      </td></tr>" & vbcrlf
		       	lcl_display_row = lcl_display_row & "</table></div>" & vbcrlf

          response.write lcl_display_row
       end if

    			oCommentList.MoveNext

       sBGColor = changeBGColor(sBGColor,"#e0e0e0","#ffffff")

   	Loop

  	'Display Submit Date/Time and User
  		response.write "<div style=""border-bottom:solid 1px #000000;background-color:" & sBGColor & ";"">" & vbcrlf
    response.write "<table>" & vbcrlf
  		response.write "  <tr><td>" & sSubmitName & " - " & UCASE("SUBMITTED") & " - " &  datSubmitDate & "</td></tr>" & vbcrlf
		  response.write "</table>" & vbcrlf
    response.write "</div>" & vbcrlf

    oCommentList.close
    set oCommentList = nothing

	else

  	'No activity for this request
  		response.write "<div style=""border-bottom:solid 1px #000000;background-color:#e0e0e0"">" & vbcrlf
    response.write "<table>" & vbcrlf
  		response.write "  <tr><td><font style=""color:red;font-size:12px;"">&nbsp;&nbsp;&nbsp;<i>No activity Reported.</i></td></tr>" & vbcrlf
		  response.write "</table>" & vbcrlf
    response.write "</div>" & vbcrlf

  	'Display Submit Date/Time and User
  		response.write "<div style=""border-bottom:solid 1px #000000;background-color:#ffffff;"">" & vbcrlf
    response.write "<table>" & vbcrlf
  		response.write "  <tr><td>" & sSubmitName & " - " & UCASE("SUBMITTED") & " - " &  datSubmitDate & "</td></tr>" & vbcrlf
		  response.write "</table>" & vbcrlf
    response.write "</div>" & vbcrlf
 end if

End Function

'------------------------------------------------------------------------------
function remove_html_tags(p_value)
'<p>
  p_value = replace(p_value,"<P>","")
  p_value = replace(p_value,"</P>",vbcrlf & vbcrlf)
  p_value = replace(p_value,"<p>","")
  p_value = replace(p_value,"</p>",vbcrlf & vbcrlf)

'<br>
  p_value = replace(p_value,"<BR>",vbcrlf)
  p_value = replace(p_value,"</BR>",vbcrlf)
  p_value = replace(p_value,"<br>",vbcrlf)
  p_value = replace(p_value,"</br>",vbcrlf)
  p_value = replace(p_value,"<BR />",vbcrlf)
  p_value = replace(p_value,"</BR />",vbcrlf)
  p_value = replace(p_value,"<br />",vbcrlf)
  p_value = replace(p_value,"</br />",vbcrlf)

  remove_html_tags = p_value

end function

'------------------------------------------------------------------------------
function dbsafe(p_value)
  if p_value <> "" then
     lcl_value = REPLACE(p_value,"'","''")
  else
     lcl_value = ""
  end if

  dbsafe = lcl_value

end function

'------------------------------------------------------------------------------
function display_back_next (sQueryString, p_pagenum, p_page_count)
  if sQueryString <> "" then
     if left(sQueryString,8) <> "pagenum=" then
        lcl_back_querystring = "?pagenum=1"
        lcl_next_querystring = "?pagenum=2"

        lcl_back_querystring = lcl_back_querystring & "&" & sQueryString
        lcl_next_querystring = lcl_next_querystring & "&" & sQueryString
     else
        lcl_amp_pos          = instr(sQueryString,"&")
        lcl_query_string     = mid(sQueryString,lcl_amp_pos+1)

       'Set up the BACK page number
        if clng(p_pagenum)-1 > 0 then
           lcl_back_querystring = "?pagenum=" & clng(p_pagenum) - 1
        else
           lcl_back_querystring = "?pagenum=1"
        end if

       'Set up the NEXT page number
        if clng(p_pagenum)+1 <= clng(p_page_count) then
           lcl_next_querystring = "?pagenum=" & clng(p_pagenum) + 1
        else
           lcl_next_querystring = "?pagenum=" & p_pagenum
        end if

        lcl_back_querystring = lcl_back_querystring & "&" & lcl_query_string
        lcl_next_querystring = lcl_next_querystring & "&" & lcl_query_string
     end if

  else
     lcl_back_querystring = "?pagenum=1"
     lcl_next_querystring = "?pagenum=2"
  end if

 'Strip off the "init=Y" as this will not be the initial time that the use opens the screen.
  if instr(lcl_back_querystring,"init=Y") > 0 then
     lcl_back_querystring = replace(lcl_back_querystring,"&init=Y","")
  end if

  if instr(lcl_next_querystring,"init=Y") > 0 then
     lcl_next_querystring = replace(lcl_next_querystring,"&init=Y","")
  end if

 'Add "useSessions=1" so that the screen knows to pull the last saved query.
  if request("init") <> "Y" OR request("init") = "" then
     if instr(lcl_back_querystring,"useSessions") < 1 then
        lcl_back_querystring = lcl_back_querystring & "&useSessions=1"
     end if

     if instr(lcl_next_querystring,"useSessions") < 1 then
        lcl_next_querystring = lcl_next_querystring & "&useSessions=1"
     end if
  end if

  response.write "<div>" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf

     response.write "      <td valign=""top"">" & vbcrlf
     response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
   	 response.write "            <tr>" & vbcrlf
     'response.write "                <td><a href=""action_line_list.asp" & lcl_back_querystring & """><img border=""0"" src=""../images/arrow_back.gif""></a></td>" & vbcrlf
     'response.write "                <td valign=""top""><a href=""action_line_list.asp" & lcl_back_querystring & """>BACK</a></td>" & vbcrlf
     'response.write "                <td valign=""top"">&nbsp;" & "<a href=""action_line_list.asp" & lcl_next_querystring & """>NEXT</a></td>" & vbcrlf
     'response.write "                <td valign=""top""><a href=""action_line_list.asp" & lcl_next_querystring & """><img border=""0"" src=""../images/arrow_forward.gif"" valign=""bottom""></a></td>" & vbcrlf
  if clng(p_page_count) > 1 then
     response.write "                <td><input type=""button"" name=""prevRecordsButton"" id=""prevRecordsButton"" value=""<< Back"" class=""button ui-button ui-widget ui-corner-all"" onclick=""location.href='action_line_list.asp" & lcl_back_querystring & "'""  />" & vbcrlf
     response.write "                <input type=""button"" name=""nextRecordsButton"" id=""nextRecordsButton"" value=""Next >>"" class=""button ui-button ui-widget ui-corner-all"" onclick=""location.href='action_line_list.asp" & lcl_next_querystring & "'""  /></td>" & vbcrlf
    else
	    response.write "<td></td>"
  end if
	 response.write "<td>"
     'BEGIN: Insert link to export results ------------------------------------
     %>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
     <div class="dropdown" style="float:right;">
  	<button class="ui-button ui-widget ui-corner-all dd-green"><i class="fa fa-bars" aria-hidden="true"></i> Tools</button>
  	<div class="dropdown-content">
		<a href="javascript:getWidget()">Request Summary</a>
<%
     'Add Map Link
      if lcl_orghasfeature_issuelocation AND lcl_orghasfeature_actionline_issuelocation_mapit then
         response.write "<a href=""#"" name=""buttonMapIt"" id=""buttonMapIt"">Map It!</a>" & vbcrlf
      end if
      if lcl_orghasfeature_data_export then
         if lcl_userhaspermission_data_export then
            if request.querystring("reporttype") = "" or LCASE(request.querystring("reporttype")) = "list" Then
              'DISPLAY EXPORT BUTTON ONLY FOR LIST VIEW
	       response.write "<a href=""../export/csv_export.asp"">Download as CSV</a>"

	       if lcl_orghasfeature_csv_foil_export then
	       		response.write "<a href=""../export/csv_ryefoilexport.asp"">FOIL Status Export</a>"
	       		response.write "<a href=""../export/csv_ryerockexport.asp"">Rock Remov. Reg Export</a>"
	       end if

              'Check to see if the org and the user have the "Action Line - CSV Export (parsed)" feature assigned.
               if lcl_orghasfeature_csv_export_parsed AND lcl_userhaspermission_csv_export_parsed then
	          response.write "<a href=""../export/csv_export_parsed.asp"">Download as CSV (parsed)</a>"
               end if

               if lcl_orghasfeature_activity_log_download AND lcl_userhaspermission_activity_log_download then
	          response.write "<a href=""../export/csv_export_activitylog.asp"">Download Activity Log</a>"
               end if
            End If
         End If
      End If

 'Show the Custom Reports - Code Sections button if the org and user have the feature assigned
  'if  lcl_orghasfeature_customreports _
  'AND lcl_userhaspermission_customreports _
  if lcl_orghasfeature_customreports_codesections AND lcl_userhaspermission_customreports_codesections then
     'response.write "                <input type=""button"" name=""sCustomReports_CodeSections"" id=""sCustomReports_CodeSections"" value=""Custom Report - Code Violations"" style=""cursor:pointer"" onclick=""openCustomReports('codesections')"" class=""ui-button ui-widget ui-corner-all""  />" & vbcrlf
	          response.write "<a href=""javascript:openCustomReports('codesections')"">Custom Report - Code Violations</a>"
  end if

  'response.write "                <input type=""button"" name=""printerFriendlyButton"" id=""printerFriendlyButton"" value=""Printer Friendly Results"" class=""button ui-button ui-widget ui-corner-all"" onclick=""openPrinterFriendlyResults()"" />" & vbcrlf
  response.write "<a href=""javascript:openPrinterFriendlyResults()"">Printer Friendly Results</a>"
     'END: Insert link to export results --------------------------------------

     'Print Multiple Work Orders
      if lcl_orghasfeature_display_multiple_workorders then
         'response.write "<input type=""button"" value=""Print Work Order(s)"" style=""cursor:pointer"" onClick=""printWorkOrders();"" class=""ui-button ui-widget ui-corner-all"" >" & vbcrlf
	 response.write "<a href=""javascript:printWorkOrders();"">Print Work Order(s)</a>"
      end if
  response.write "</div>"
  response.write "</div>" & vbcrlf
      if session("orgid") = 5 or session("orgid") = 153 or session("orgid") = 130 then
  response.write "<div id=""bulkassign"" style=""float:right;"">"
     	response.write "                    <select name=""bulkemployeeid"" id=""bulkemployeeid"">" & vbcrlf
                                          	DrawAdminUsersAssignedHideDeleted("")
     response.write "                    </select>" & vbcrlf
     response.write "                    <select name=""bulkstatus"" id=""bulkstatus"">" & vbcrlf
     response.write "                      <option value=""SUBMITTED"""  & CheckSelected(sStatus,"SUBMITTED")  & ">SUBMITTED</option>"  & vbcrlf
     response.write "                      <option value=""INPROGRESS""" & CheckSelected(sStatus,"INPROGRESS") & ">INPROGRESS</option>" & vbcrlf
     response.write "                      <option value=""WAITING"""    & CheckSelected(sStatus,"WAITING")    & ">WAITING</option>"    & vbcrlf

     if lcl_userhaspermission_can_close_requests then
        response.write "                   <option value=""RESOLVED"""  & CheckSelected(sStatus,"RESOLVED")  & ">RESOLVED</option>" & vbcrlf
        response.write "                   <option value=""DISMISSED""" & CheckSelected(sStatus,"DISMISSED") & ">DISMISSED</option>" & vbcrlf
     end if

     response.write "                    </select>" & vbcrlf
     if lcl_orghasfeature_modify_actionline_department then
        response.write "<select name=""bulkdeptid"" id=""bulkdeptid"">" & vbcrlf
                          DrawDepartments sDeptID,"Y"
        response.write "</select>" & vbcrlf
     end if
         response.write "<input type=""button"" value=""Bulk Assign"" style=""cursor:pointer"" onClick=""bulkAssign();"" class=""ui-button ui-widget ui-corner-all"" >" & vbcrlf
  response.write "</div>"
      end if
     response.write "		     </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
   	 response.write "          </table>" & vbcrlf
     response.write "      </td>" & vbcrlf

  response.write "      </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf
end function

'------------------------------------------------------------------------------
sub retrieveOpenClosedCounts(ByVal p_querytype, ByVal p_orderby, ByVal p_value, ByVal p_where_clause, _
                             ByRef lcl_os_num_open, ByRef lcl_os_total_days_open, ByRef lcl_cs_num_closed, ByRef lcl_cs_total_days_closed)

   'Build the query
    sSQL = "SELECT count(*) as totalrequests, "
    sSQL = sSQL & " SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totaldays "
    sSQL = sSQL & " FROM egov_action_request_view "
    sSQL = sSQL & p_where_clause

   'Determine which statuses to query on
    if UCASE(p_querytype) = "OPEN" then
       sSQL = sSQL & " AND status <> 'RESOLVED' "
       sSQL = sSQL & " AND status <> 'DISMISSED' "
    else
       sSQL = sSQL & " AND (status = 'RESOLVED' OR status = 'DISMISSED') "
    end if

   'Setup the ORDER BY
    if UCASE(p_orderby) = "SUBMIT_DATE" then
       sSQL = sSQL & " AND submitdateshort = '"  & p_value & "'"
    elseif UCASE(p_orderby) = "ACTION_FORMID" then
       sSQL = sSQL & " AND action_FormTitle = '" & p_value & "'"
    elseif UCASE(p_orderby) = "DEPTID" then
       sSQL = sSQL & " AND deptId = '"           & p_value & "'"
    elseif UCASE(p_orderby) = "ASSIGNED_NAME" then
       sSQL = sSQL & " AND assignedName = '"     & p_value & "'"
    elseif UCASE(p_orderby) = "SUBMITTEDBY" then
       sSQL = sSQL & " AND userfname + ' ' + userlname = '" & p_value & "'"
    elseif UCASE(p_orderby) = "STATUS" then
       sSQL = sSQL & " AND UPPER(status) = '"    & UCASE(p_value) & "'"
    end if

    'SQLopenText    = "" & sSQL & ""
    'SQLclosedText  = "" & sSQL & ""

    set oRetrieve = Server.CreateObject("ADODB.Recordset")
    'oRetrieve.Open sSQL, Application("DSN"), 0, 1
    oRetrieve.Open sSQL, lcl_dsn, 0, 1

   'Initialize the variables
    lcl_os_num_open          = 0
    lcl_os_total_days_open   = 0
    lcl_cs_num_closed        = 0
    lcl_cs_total_days_closed = 0

    if UCASE(p_querytype) = "OPEN" then
       lcl_os_num_open        = oRetrieve("totalrequests")
       lcl_os_total_days_open = oRetrieve("totaldays")
    else
       lcl_cs_num_closed        = oRetrieve("totalrequests")
       lcl_cs_total_days_closed = oRetrieve("totaldays")
    end if

    oRetrieve.close
    set oRetrieve = nothing

end sub

'------------------------------------------------------------------------------
 sub displayTotalRow (p_total_type, p_record_count, p_description, p_subTotal, _
                      p_numOpen, p_avgOpenTotal, p_avgClosedTotal)

    lcl_font_style = " style=""color: navy; font-size: 10pt; font-weight: bold;"""

   'Determine if this is a Grand Total or Sub-Total row.
    if UCASE(p_total_type) = "GRANDTOTAL" then
       lcl_row_info = "Grand Total [" & p_record_count & " Requests]"
    else
       lcl_row_info = "Subtotal: " & p_description
    end if

    response.write "<tr bgcolor=""#dddddd""" & lcl_font_style & ">" & vbcrlf
    response.write "    <td align=""center"" colspan=""2"" nowrap=""nowrap"">" & lcl_row_info & "</td>" & vbcrlf
    response.write "    <td colspan=""2"">Total: ["                & p_subTotal       & "]</td>" & vbcrlf
    response.write "    <td colspan=""2"">Open: ["                 & p_numOpen        & "]</td>" & vbcrlf
    response.write "    <td colspan=""3"">Avg Time Still Open: ["  & p_avgOpenTotal   & "]</td>" & vbcrlf
    response.write "    <td colspan=""4"">Avg Time To Complete: [" & p_avgClosedTotal & "]</td>" & vbcrlf
    response.write "</tr>" & vbcrlf

 end sub

'------------------------------------------------------------------------------
 sub getAverages (ByVal p_os_num_open, ByVal p_os_total_days_open, ByVal p_cs_num_closed, ByVal p_cs_total_days_closed, _
                  ByRef avOpenTotal, ByRef avClosedTotal)

   'Average OPEN	
    numberOpen = clng(p_os_num_open)

    if numberOpen <> 0 then
       avOpenTotal = p_os_total_days_open / numberOpen
       avOpenTotal = formatnumber(avOpenTotal,1)
    else
       avOpenTotal = " - "
    end if

   'Average CLOSED
    numberClosed = clng(p_cs_num_closed)

    if numberClosed <> 0 then
       avClosedTotal = p_cs_total_days_closed / numberClosed
       avClosedTotal = formatnumber(avClosedTotal,1)
    else
       avClosedTotal = " - "
    end if	

 end sub

'------------------------------------------------------------------------------
 function getSearchType(iOrderBy)
   lcl_return = ""

   if iOrderBy <> "" then
      if iOrderBy     = "submit_date" then
         lcl_return = lastDate
      elseif iOrderBy = "action_Formid" then
         lcl_return = lastTitle
      elseif iOrderBy = "deptId" then
         lcl_return = lastDept
      elseif iOrderBy = "assigned_Name" then
         lcl_return = lastAssigned
      elseif iOrderBy = "status" then
         lcl_return = lastStatus
      end if
   end if

   getSearchType = lcl_return

 end function

'------------------------------------------------------------------------------
sub trackSubTotals(ByVal iOrderBy, ByVal iStatus, ByVal iCurrentValue, ByVal iPreviousValue, _
                   ByVal iSubTotal_TotalRequests, ByVal iSubTotal_TotalDaysOpen, _
                   ByVal iSubTotal_TotalRequestsOpen, ByVal iSubTotal_TotalDaysClosed, _
                   ByVal iSubTotal_TotalRequestsClosed, ByVal iCountDays, ByRef lcl_subTotal_totalRequests, _
                   ByRef lcl_subTotal_totalDaysOpen, ByRef lcl_subTotal_totalRequestsOpen, _
                   ByRef lcl_subTotal_totalDaysClosed, ByRef lcl_subTotal_totalRequestsClosed)

  lcl_isFirstLine = "Y"

 'Check to see if this is the first record (line) in the group.
  if UCASE(iOrderBy) = "SUBMIT_DATE" OR UCASE(iOrderBy) = "DUE_DATE" then
     if iCurrentValue <> "" AND iPreviousValue <> "" then
        if DateDiff("d",iCurrentValue,iPreviousValue) = 0 then
           lcl_isFirstLine = "N"
        end if
     else
        lcl_isFirstLine = "N"
     end if
  else
     if iCurrentValue = iPreviousValue then
        lcl_isFirstLine = "N"
     end if
  end if

 'If this is the "first line" in the group and the status is NOT equal to RESOLVED or DISMISSED then initialize the "count" variables.
 'Otherwise, increment the "count" variables.
  if lcl_isFirstLine = "N" then
     lcl_subTotal_totalRequests = iSubTotal_TotalRequests + 1

     if UCASE(iStatus) <> "RESOLVED" AND UCASE(iStatus) <> "DISMISSED" then
        lcl_subTotal_totalDaysOpen       = iSubTotal_TotalDaysOpen     + iCountDays
        lcl_subTotal_totalDaysClosed     = iSubTotal_TotalDaysClosed
        lcl_subTotal_totalRequestsOpen   = iSubTotal_TotalRequestsOpen + 1
        lcl_subTotal_totalRequestsClosed = iSubTotal_TotalRequestsClosed
     else
        lcl_subTotal_totalDaysOpen       = iSubTotal_TotalDaysOpen
        lcl_subTotal_totalDaysClosed     = iSubTotal_TotalDaysClosed     + iCountDays
        lcl_subTotal_totalRequestsOpen   = iSubTotal_TotalRequestsOpen
        lcl_subTotal_totalRequestsClosed = iSubTotal_TotalRequestsClosed + 1
     end if
  else
     if UCASE(iStatus) <> "RESOLVED" AND UCASE(iStatus) <> "DISMISSED" then
        lcl_subTotal_totalRequestsOpen   = 1
        lcl_subTotal_totalRequestsClosed = 0
        lcl_subTotal_totalDaysOpen       = iCountDays
        lcl_subTotal_totalDaysClosed     = 0
     else
        lcl_subTotal_totalRequestsOpen   = 0
        lcl_subTotal_totalRequestsClosed = 1
        lcl_subTotal_totalDaysOpen       = 0
        lcl_subTotal_totalDaysClosed     = iCountDays
     end if

     lcl_subTotal_totalRequests = 1

  end if

end sub

'------------------------------------------------------------------------------
sub displayGrandTotalLine(iReportType, iWhereClause, iRecordCount)

   if UCASE(reporttype) = "DETAIL" or UCASE(reporttype) = "DRILLTHRU" then

     'Total Days
      lcl_total_days        = 0
      lcl_num_total         = 0
      lcl_num_open          = 0
      lcl_total_days_open   = 0
      lcl_num_closed        = 0
      lcl_total_days_closed = 0

      'sSQLtotl    = "SELECT action_autoid,action_formTitle,DateDiff(d,submit_date,complete_date) AS lcl_subTotal_totalDaysOpen,submit_date,complete_date,deptID,groupname as deptName,status,assigned_Name FROM egov_action_request_view left outer join groups on deptId=groupId" & iWhereClause
      'oTotals.Open sSQLtotl, Application("DSN"), 0, 1
      'lcl_total_days = oTotals("lcl_subTotal_totalDaysOpen")

      sSQLtotl    = "SELECT action_autoid,action_formTitle,subTotal_totalDaysOpen,submit_date,complete_date,deptID,groupname as deptName,status,assigned_Name FROM egov_action_request_view left outer join groups on deptId=groupId" & iWhereClause
      set oTotals = Server.CreateObject("ADODB.Recordset")
      oTotals.Open sSQLtotl, lcl_dsn, 0, 1

      if not oTotals.eof then
         lcl_total_days = oTotals("subTotal_totalDaysOpen")
      end if

      oTotals.close
      set oTotals = nothing

     'Total Number
      sSQLTotal = "SELECT count(*) as numTotal FROM egov_action_request_view " & iWhereClause
      set oNumTotal = Server.CreateObject("ADODB.Recordset")
      'oNumTotal.Open sSQLTotal, Application("DSN"), 0, 1
      oNumTotal.Open sSQLTotal, lcl_dsn, 0, 1

      if not oNumTotal.eof then
         lcl_num_total = oNumTotal("numTotal")
      end if

      oNumTotal.close
      set oNumTotal = nothing

     'Total Days Open
      'sSQLopen = "SELECT count(*) as numOpen,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS lcl_subTotal_totalDaysOpenOpen FROM egov_action_request_view " & iWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') "
      'oOpen.Open sSQLopen, Application("DSN"), 0, 1

      sSQLopen = "SELECT count(*) as numOpen,SUM(subTotal_totalDaysOpen) as lcl_subTotal_totalDaysOpenOpen FROM egov_action_request_view " & iWhereClause & " AND (status<>'RESOLVED' AND status<>'DISMISSED') "
      set oOpen = Server.CreateObject("ADODB.Recordset")
      oOpen.Open sSQLopen, lcl_dsn, 0, 1

      if not oOpen.eof then
         lcl_num_open        = oOpen("numOpen")
         lcl_total_days_open = oOpen("lcl_subTotal_totalDaysOpenOpen")
      end if

      oOpen.close
      set oOpen = nothing

     'Total Days Closed
      'sSQLclosed = "SELECT count(*) as numClosed,SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS lcl_subTotal_totalDaysOpenClosed FROM egov_action_request_view " & iWhereClause & " AND (status='RESOLVED' OR status='DISMISSED') "
      'oClosed.Open sSQLclosed, Application("DSN"), 0, 1

      sSQLclosed = "SELECT count(*) as numClosed,SUM(subTotal_totalDaysOpen) AS lcl_subTotal_totalDaysOpenClosed FROM egov_action_request_view " & iWhereClause & " AND (status='RESOLVED' OR status='DISMISSED') "
      set oClosed = Server.CreateObject("ADODB.Recordset")
      oClosed.Open sSQLclosed, lcl_dsn, 0, 1

      if not oClosed.eof then
         lcl_num_closed        = oClosed("numClosed")
         lcl_total_days_closed = oClosed("lcl_subTotal_totalDaysOpenclosed")
      end if

      oClosed.close
      set oClosed = nothing

     'Average Open
      numOpen = clng(lcl_num_open)
      if numOpen<>0 and lcl_total_days_open <> 0 then
         avgOpenTotal = lcl_total_days_open / numOpen
         avgOpenTotal = formatnumber(avgOpenTotal,1)
      else
         avgOpenTotal = ""
      end if

     'Average Closed
      numClosed = clng(lcl_num_closed)
      if numClosed<>0 and lcl_total_days_closed <> 0 then
         avgClosedTotal = lcl_total_days_closed / numClosed
         avgClosedTotal = formatnumber(avgClosedTotal,1)
      else
         avgClosedTotal = ""
      end if

     'Display Grand Total Row
      displayTotalRow "GRANDTOTAL", iRecordCount, "", lcl_num_total, lcl_num_open, avgOpenTotal, avgClosedTotal
   end if

end sub

%>
