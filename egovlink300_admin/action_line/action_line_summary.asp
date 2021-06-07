<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<!-- #include file="../customreports/customreports_global_functions.asp" //-->
<%
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
 lcl_orghasfeature_savesearchoptions_actionline     = orghasfeature("savesearchoptions_actionline")
 lcl_orghasfeature_action_line_substatus            = orghasfeature("action_line_substatus")
 lcl_orghasfeature_actionline_maintain_duedate      = orghasfeature("actionline_maintain_duedate")
 lcl_orghasfeature_actionline_hide_internalcomments = orghasfeature("actionline_hide_internalcomments")

'Check for user permissions
 lcl_userhaspermission_action_line_substatus       = userhaspermission(session("userid"),"action_line_substatus")
 lcl_userhaspermission_actionline_maintain_duedate = userhaspermission(session("userid"),"actionline_maintain_duedate")

'Determine the screen mode (Report Type) to display
'Screen Modes are: PRINT and DISPLAY and NULL
 if request("screen_mode") <> "" then
    lcl_screen_mode = request("screen_mode")
 else
    lcl_screen_mode = "DISPLAY"
 end if

 ViewAll = checkViewAll()

'Get all of the customreport ids
 lcl_customreportid_actionline_user        = getCustomReportID("ACTIONLINE - USER",         session("orgid"), session("userid"), False)
 lcl_customreportid_actionline_lastqueried = getCustomReportID("ACTIONLINE - LAST QUERIED", session("orgid"), session("userid"), False)
 lcl_customreportid_actionline_defaults    = getCustomReportID("ACTIONLINE - DEFAULTS",     session("orgid"), session("userid"), True)

 if request.ServerVariables("REQUEST_METHOD") = "POST" then
   'Save "Last Queried" Search Options.
    lcl_success = "Y"

    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectAssignedto",        request("selectAssignedto"),        False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "orderBy",                 request("orderBy"),                 False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "recordsPer",              request("recordsPer"),              False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "reporttype",              request("reporttype"),              False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectFormId",            request("selectFormId"),            False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectDeptId",            request("selectDeptId"),            False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "pastDays",                request("pastDays"),                False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "searchDaysType",          request("searchDaysType"),          False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "fromDate",                request("fromDate"),                False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "toDate",                  request("toDate"),                  False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "fromToDateSelection",     request("Date"),                    False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectDateType",          request("selectDateType"),          False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusDismissed",         request("statusDismissed"),         False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusResolved",          request("statusResolved"),          False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusWaiting",           request("statusWaiting"),           False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusInprogress",        request("statusInprogress"),        False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusSubmitted",         request("statusSubmitted"),         False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "substatus_hidden",        request("substatus_hidden"),        False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectUserFName",         request("selectUserFName"),         False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectUserLName",         request("selectUserLName"),         False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectIssueStreetNumber", request("selectIssueStreetNumber"), False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectIssueStreet",       request("selectIssueStreet"),       False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectContactStreet",     request("selectContactStreet"),     False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectCounty",            request("selectCounty"),            False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectBusinessName",      request("selectBusinessName"),      False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectTicket",            request("selectTicket"),            False, lcl_success

    if UCASE(reporttype) <> "STATUSSUMMARY" then
       saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectInitialResponse",   request("selectInitialResponse"),   False, lcl_success
       saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectRequestsResolved",  request("selectRequestsResolved"),  False, lcl_success
    end if
 end if

'If user had set search options for this session then get the session values
 if request("useSessions") = 1 then
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

      'Retrieve the "Last Queried" Search Options
       recordsPer              = getCustomReportSearchOption(lcl_customreportid, "recordsPer")
       reporttype              = getCustomReportSearchOption(lcl_customreportid, "reporttype")
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

       statusDismissed         = getCustomReportSearchOption(lcl_customreportid, "statusDismissed")
       statusResolved          = getCustomReportSearchOption(lcl_customreportid, "statusResolved")
       statusWaiting           = getCustomReportSearchOption(lcl_customreportid, "statusWaiting")
       statusInprogress        = getCustomReportSearchOption(lcl_customreportid, "statusInprogress")
       statusSubmitted         = getCustomReportSearchOption(lcl_customreportid, "statusSubmitted")

       substatus_hidden        = getCustomReportSearchOption(lcl_customreportid, "substatus_hidden")

       selectUserFName         = getCustomReportSearchOption(lcl_customreportid, "selectUserFName")
       selectUserLName         = getCustomReportSearchOption(lcl_customreportid, "selectUserLName")

       selectIssueStreetNumber = getCustomReportSearchOption(lcl_customreportid, "selectIssueStreetNumber")
       selectIssueStreet       = getCustomReportSearchOption(lcl_customreportid, "selectIssueStreet")
       selectContactStreet     = getCustomReportSearchOption(lcl_customreportid, "selectContactStreet")
       selectCounty            = getCustomReportSearchOption(lcl_customreportid, "selectCounty")
       selectBusinessName      = getCustomReportSearchOption(lcl_customreportid, "selectBusinessName")
       selectTicket            = getCustomReportSearchOption(lcl_customreportid, "selectTicket")

       if UCASE(reporttype) <> "STATUSSUMMARY" then
          selectInitialResponse   = getCustomReportSearchOption(lcl_customreportid, "selectInitialResponse")
          selectRequestsResolved  = getCustomReportSearchOption(lcl_customreportid, "selectRequestsResolved")
       end if
    end if
else
   'Get the modified search options.
   'These are the values on the screen that have been entered when the "SEARCH" button was pressed.
  		recordsPer              = request("recordsPer")
		  reporttype              = request("reporttype")

  		orderBy                 = request("orderBy")
		  selectFormId            = request("selectFormId")
    selectAssignedto        = request("selectAssignedto")
  		selectDeptId            = request("selectDeptId")
 			pastDays                = request("pastDays")
 			searchDaysType          = request("searchDaysType")

  		selectUserFName         = request("selectUserFName")
		  selectUserLName         = request("selectUserLName")

  		fromDate                = request("fromDate")
		  toDate                  = request("toDate")
    fromToDateSelection     = request("Date")
    selectDateType          = request("selectDateType")

  		statusSubmitted         = request("statusSUBMITTED")
		  statusInprogress        = request("statusINPROGRESS")
  		statusWaiting           = request("statusWAITING")
		  statusResolved          = request("statusRESOLVED")
  		statusDismissed         = request("statusDISMISSED")

  		substatus_hidden        = request("substatus_hidden")
    show_hide_substatus     = request("show_hide_substatus")

    selectContactStreet     = request("selectContactStreet")
    selectIssueStreetNumber = request("selectIssueStreetNumber")
  		selectIssueStreet       = request("selectIssueStreet")
    selectCounty            = request("selectCounty")
    selectBusinessName      = request("selectBusinessName")
  		selectTicket            = request("selectTicket")

    if UCASE(reporttype) <> "STATUSSUMMARY" then
       selectInitialResponse   = request("selectInitialResponse")
       selectRequestsResolved  = request("selectRequestsResolved")
    end if
 end if

'Set status
 if statusSubmitted = "yes" then
    noStatus = "false"
 else
    statusSubmitted = "no"
 end if

 if statusInprogress = "yes" then
    noStatus = "false"
 else
    statusInprogress = "no"
 end if

 if statusWaiting = "yes" then
    noStatus = "false"
 else
    statusWaiting = "no"
 end if

 if statusResolved = "yes" then
    noStatus = "false"
 else
    statusResolved = "no"
 end if

 if statusDismissed = "yes" then
    noStatus = "false"
 else
    statusDismissed = "no"
 end if

'Determine if this is the initial time the screen has been opened.
 if request("init") = "Y" _
 OR (request("init")  = ""   AND _
     statusSubmitted  = "no" AND _
     statusInprogress = "no" AND _
     statusWaiting    = "no" AND _
     statusResolved   = "no" AND _
     statusDismissed  = "no" AND _
     substatus_hidden = "") then
    lcl_init = "Y"
 else
    lcl_init = "N"
 end if

 if lcl_init = "Y" then
    session("isFromEmail") = ""

    statusSubmitted  = "yes"
    statusInprogress = "yes"
    statusWaiting    = "yes"
    statusResolved   = "yes"
    statusDismissed  = "yes"

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
          recordsPer              = getCustomReportSearchOption(lcl_customreportid, "recordsPer")
          reporttype              = getCustomReportSearchOption(lcl_customreportid, "reporttype")
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

          statusDismissed         = getCustomReportSearchOption(lcl_customreportid, "statusDismissed")
          statusResolved          = getCustomReportSearchOption(lcl_customreportid, "statusResolved")
          statusWaiting           = getCustomReportSearchOption(lcl_customreportid, "statusWaiting")
          statusInprogress        = getCustomReportSearchOption(lcl_customreportid, "statusInprogress")
          statusSubmitted         = getCustomReportSearchOption(lcl_customreportid, "statusSubmitted")

          substatus_hidden        = getCustomReportSearchOption(lcl_customreportid, "substatus_hidden")

          selectUserFName         = getCustomReportSearchOption(lcl_customreportid, "selectUserFName")
          selectUserLName         = getCustomReportSearchOption(lcl_customreportid, "selectUserLName")

          selectIssueStreetNumber = getCustomReportSearchOption(lcl_customreportid, "selectIssueStreetNumber")
          selectIssueStreet       = getCustomReportSearchOption(lcl_customreportid, "selectIssueStreet")
          selectContactStreet     = getCustomReportSearchOption(lcl_customreportid, "selectContactStreet")
          selectCounty            = getCustomReportSearchOption(lcl_customreportid, "selectCounty")
          selectBusinessName      = getCustomReportSearchOption(lcl_customreportid, "selectBusinessName")
          selectTicket            = getCustomReportSearchOption(lcl_customreportid, "selectTicket")

          if UCASE(reporttype) <> "STATUSSUMMARY" then
             selectInitialResponse   = getCustomReportSearchOption(lcl_customreportid, "selectInitialResponse")
             selectRequestsResolved  = getCustomReportSearchOption(lcl_customreportid, "selectRequestsResolved")
          end if
       end if
    end if
 end if

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

 'iDefault_fromDate = "N"
 'iDefault_toDate   = "N"

'If a date range has been selected then calculate the from/to dates.
'Othwerwise, use what has been entered or the user's default(s).
 if fromToDateSelection <> "" AND fromToDateSelection <> "0" then
    getDatesFromDateRangeChoices fromToDateSelection, lcl_fromDate, lcl_toDate

    fromDate = lcl_fromDate
    toDate   = lcl_toDate
 else

   'From Date (last year)
    if fromDate = "" or IsNull(fromDate) then
       fromDate = dateAdd("yyyy",-1,today)
       'iDefault_fromDate = "Y"
    end if

   'To Date (get today's date)
    if toDate = "" or IsNull(toDate) then
       toDate = dateAdd("d",0,today)
       'iDefault_toDate = "Y"
    end if
 end if

'To Date (today's date + 1 day - so everything falls BETWEEN the FROM and TO Dates in the search.)
 'if lcl_screen_mode <> "PRINT" then
 '   if iDefault_toDate = "Y" then 
 '      toDate = dateAdd("d",1,toDate)
 '   end if
 'end if

'Records Per Page
 if lcl_screen_mode <> "PRINT" then
    recordsPer = checkRecordsPerPageFilter(25,recordsPer)
 end if

'Default the column calculations
 if UCASE(reporttype) <> "STATUSSUMMARY" then
    if selectInitialResponse = "" or isnull(selectInitialResponse) then
       selectInitialResponse = 1
    end if

    if selectRequestsResolved = "" or isnull(selectRequestsResolved) then
       selectRequestsResolved = 3
    end if
 end if

'Found within Organizational Features => Properties ---------------------------
'Determine if the org has the "Uses the "After 5PM is Next Day" Logic on Action Line" option turned on
'OR the "Counts Week Days Only on Action Line Calculations" option turned on then set up the display text
 sSQL = "SELECT usesafter5adjustment, usesweekdays "
 sSQL = sSQL & " FROM organizations "
 sSQL = sSQL & " WHERE orgid = " & session("orgid")

 set oAdjustedTime = Server.CreateObject("ADODB.Recordset")
 oAdjustedTime.Open sSQL, Application("DSN"), 3, 1

 if not oAdjustedTime.eof then
'    if oAdjustedTime("usesafter5adjustment") = True OR oAdjustedTime("usesweekdays") = True then
'       lcl_adjustment_label = "Weekday(s)"
'    else
'       lcl_adjustment_label = "Day(s)"
'    end if
    lcl_adjustment_label = "Day(s)"
 end if

oAdjustedTime.close
set oAdjustedTime = nothing
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
</style>

  <script type="text/javascript" src="../scripts/selectAll.js"></script>
 	<script type="text/javascript" src="../scripts/ajaxLib.js"></script>
 	<script type="text/javascript" src="../scripts/getdates.js"></script>
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.7.1.min.js"></script>

<script type="text/javascript" >
<!--
  function checkStat() {
    if ( !(form1.statusSUBMITTED.checked) &&  !(form1.statusINPROGRESS.checked) && !(form1.statusWAITING.checked) && !(form1.statusRESOLVED.checked) && !(form1.statusDISMISSED.checked)) {
          alert("You must select the status.");
      		  form1.statusSUBMITTED.focus();
          return false;
    }
  }

  function CheckAllStatus(checkSt) {
			//if (document.form1.CheckAllStat.checked) {
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
   		//} else if (document.form1.reporttype.value == "ListFull") {
   		//		document.forms[0].action = "action_line_list.asp";
   		//		document.forms[0].submit();
     } else {
       var lcl_reporttype_hideinternal = $('#reporttype_hideinternal').val();

			   	document.forms[0].action = "action_line_list.asp?reporttype_hideinternal=" + lcl_reporttype_hideinternal;
   				document.forms[0].submit();
     }
		}
}

function validateFields() {

		var daterege         = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
		var dateFromOk       = daterege.test(document.getElementById("fromDate").value);
		var dateToOk         = daterege.test(document.getElementById("toDate").value);
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

  mainlist          = document.getElementById('selStatus');
  sub_list          = document.getElementById('selSubStatus');
  sub_list_row      = document.getElementById('sub_status_row');
  sub_list_row_text = document.getElementById('sub_status_row_text');
  i = 0
<%
  dim oMainStatus, oSubStatus, oSubStatus_Count, line_count, lcl_sub_line_count, lcl_total_count

 'Retrieve all of the MAIN statuses
  sSQL = "SELECT action_status_id, status_name, orgid, parent_status, display_order, active_flag "
  sSQL = sSQL & " FROM egov_actionline_requests_statuses "
  sSQL = sSQL & " WHERE orgid = 0 "
  sSQL = sSQL & " AND parent_status = 'MAIN' "
  sSQL = sSQL & " AND active_flag = 'Y' "
  sSQL = sSQL & " ORDER BY display_order "

  set oMainStatus = Server.CreateObject("ADODB.Recordset")
  oMainStatus.Open sSQL, Application("DSN"), 3, 1

  if not oMainStatus.eof then
     line_count = 0

   	 do while not oMainStatus.eof
        line_count = line_count + 1

      		if line_count = 1 then
           response.write "if(document.getElementById('selStatus" & oMainStatus("status_name") & "').checked==true) {" & vbcrlf
        else
           response.write "}else if(document.getElementById('selStatus" & oMainStatus("status_name") & "').checked==true) {" & vbcrlf
        end if

       'Get the total count of SubStatuses
        sSQL = "SELECT count(action_status_id) AS Total_SubStatus FROM egov_actionline_requests_statuses "
        sSQL = sSQL & " WHERE orgid = "         & clng(Session("OrgID"))
        sSQL = sSQL & " AND parent_status = '"  & oMainStatus("status_name") & "' "
        sSQL = sSQL & " AND active_flag = 'Y' "

        set oSubStatus_Count = Server.CreateObject("ADODB.Recordset")
        oSubStatus_Count.Open sSQL, Application("DSN"), 3, 1

        lcl_total_count = oSubStatus_Count("Total_SubStatus")

        if lcl_total_count > 0 then
		  
      		  'Retrieve all of the Sub-Statuses for each MAIN status for the OrgID and the form
           sSQL = "SELECT action_status_id, status_name "
           sSQL = sSQL & " FROM egov_actionline_requests_statuses "
           sSQL = sSQL & " WHERE orgid = "         & clng(Session("OrgID"))
           sSQL = sSQL & " AND parent_status = '"  & oMainStatus("status_name") & "' "
           sSQL = sSQL & " AND active_flag = 'Y' "
           sSQL = sSQL & " ORDER BY display_order "

           set oSubStatus = Server.CreateObject("ADODB.Recordset")
           oSubStatus.Open sSQL, Application("DSN"), 3, 1

           if not oSubStatus.eof then
              response.write "sub_list_row.style.display = ""block"";" & vbcrlf
              response.write "sub_list.style.display     = ""block"";" & vbcrlf

             'remove the current values
              response.write "for(var i=0; i < sub_list.length; i++) {" & vbcrlf
              response.write "    sub_list.remove(i);" & vbcrlf
              response.write "}" & vbcrlf

             'Loop through the sub statuses
              lcl_sub_line_count = 0
              do while NOT oSubStatus.EOF

                'build the new values
                 response.write "document.forms[""form1""].selSubStatus.options[" & lcl_sub_line_count & "] = new Option(""" & oSubStatus("status_name") & """,""" & oSubStatus("action_status_id") & """);" & vbcrlf

                 lcl_sub_line_count = lcl_sub_line_count + 1
                 oSubStatus.movenext
              loop

          		  oSubStatus.close
          		  oSubStatus_Count.close

         			  set oSubStatus       = nothing
         			  set oSubStatus_Count = nothing 
      		   else
              response.write "sub_list_row.style.display = ""none"";" & vbcrlf
              response.write "sub_list.style.display     = ""none"";" & vbcrlf
           end if
        else
           response.write "sub_list_row.style.display = ""none"";" & vbcrlf
           response.write "sub_list.style.display     = ""none"";" & vbcrlf
        end if

      		oMainStatus.movenext
   	 loop

     response.write "}else{" & vbcrlf
     response.write "   sub_list_row.style.display = ""none"";" & vbcrlf
     response.write "   sub_list.style.display     = ""none"";" & vbcrlf
		   response.write "}" & vbcrlf
  end if

  oMainStatus.close
  Set oMainStatus = nothing 
%>
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
  oTotal.Open sSqlc, Application("DSN"), 3, 1

  lcl_total_substatuses = oTotal("total_count")

  oTotal.close
  set oTotal = nothing

 '1. Build the javascript that will cycle through all of the sub-status search criteria checkboxes and determine which ones
 'have been checked.
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
  oChange.Open sSqla, Application("DSN"), 3, 1

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

function viewDetails(iLineNum) {
<%
 'Override with custom values with the row clicked.
  if UCASE(orderby) = "SUBMIT_DATE" OR UCASE(orderby) = "DUE_DATE" then
     response.write "document.getElementById(""toDate"").value = document.getElementById(""sc_toDate"" + iLineNum).value;" & vbcrlf
     response.write "document.getElementById(""fromDate"").value = document.getElementById(""sc_fromDate"" + iLineNum).value;" & vbcrlf
  elseif UCASE(orderby) = "ACTION_FORMID" then
     response.write "document.getElementById(""selectFormId"").value = document.getElementById(""sc_selectFormId"" + iLineNum).value;" & vbcrlf
  elseif UCASE(orderby) = "DEPTID" then
     response.write "document.getElementById(""selectDeptId"").value = document.getElementById(""sc_selectDeptId"" + iLineNum).value;" & vbcrlf
  elseif UCASE(orderby) = "ASSIGNED_NAME" then
     response.write "document.getElementById(""selectAssignedto"").value = document.getElementById(""sc_selectAssignedto"" + iLineNum).value;" & vbcrlf
  elseif UCASE(orderby) = "STREETNAME" then
     response.write "document.getElementById(""selectIssueStreetNumber"").value = document.getElementById(""sc_selectIssueStreetNumber"" + iLineNum).value;" & vbcrlf
     response.write "document.getElementById(""selectIssueStreet"").value = document.getElementById(""sc_selectIssueStreet"" + iLineNum).value;" & vbcrlf
  elseif UCASE(orderby) = "SUBMITTEDBY" then
     response.write "document.getElementById(""selectUserLName"").value = document.getElementById(""sc_selectUserLName"" + iLineNum).value;" & vbcrlf
     response.write "document.getElementById(""selectUserFName"").value = document.getElementById(""sc_selectUserFName"" + iLineNum).value;" & vbcrlf
  elseif UCASE(orderby) = "STATUS" then
     response.write "document.getElementById(""statusSUBMITTED"").checked=false;" & vbcrlf
     response.write "document.getElementById(""statusINPROGRESS"").checked=false;" & vbcrlf
     response.write "document.getElementById(""statusWAITING"").checked=false;" & vbcrlf
     response.write "document.getElementById(""statusRESOLVED"").checked=false;" & vbcrlf
     response.write "document.getElementById(""statusDISMISSED"").checked=false;" & vbcrlf

     response.write "if(document.getElementById(""sc_status"" + iLineNum).value == ""SUBMITTED"") {" & vbcrlf
     response.write "   document.getElementById(""statusSUBMITTED"").checked=true;" & vbcrlf
     response.write "}else if(document.getElementById(""sc_status"" + iLineNum).value == ""INPROGRESS"") {" & vbcrlf
     response.write "   document.getElementById(""statusINPROGRESS"").checked=true;" & vbcrlf
     response.write "}else if(document.getElementById(""sc_status"" + iLineNum).value == ""WAITING"") {" & vbcrlf
     response.write "   document.getElementById(""statusWAITING"").checked=true;" & vbcrlf
     response.write "}else if(document.getElementById(""sc_status"" + iLineNum).value == ""RESOLVED"") {" & vbcrlf
     response.write "   document.getElementById(""statusRESOLVED"").checked=true;" & vbcrlf
     response.write "}else if(document.getElementById(""sc_status"" + iLineNum).value == ""DISMISSED"") {" & vbcrlf
     response.write "   document.getElementById(""statusDISMISSED"").checked=true;" & vbcrlf
     response.write "}" & vbcrlf

  end if

  if UCASE(reporttype) = "STATUSSUMMARY" then
     response.write "document.getElementById(""orderBy"").value = 'status';" & vbcrlf     
  end if

  response.write "document.getElementById(""reporttype"").value = 'Detail';" & vbcrlf
  response.write "document.getElementById(""searchButton"").click();" & vbcrlf
%>
}

//function showhide_substatus_criteria() {
//  var lcl_button = document.getElementById('selectSubStatus');

//  if(lcl_button.style.display == "block") {
//     lcl_button.style.display = "none";
//     document.getElementById('show_hide_substatus').value = "HIDE";
//  }else{
//     lcl_button.style.display = "block";
//     document.getElementById('show_hide_substatus').value = "SHOW";
//  }
//}

//function show_hide_init(p_value) {
//  var lcl_button = document.getElementById('selectSubStatus');

//  if(p_value=="SHOW") {
//     lcl_button.style.display = "block";
//     document.getElementById("show_hide_substatus").value = "SHOW";
//  }else if(p_value=="HIDE") {
//     lcl_button.style.display = "none";
//     document.getElementById("show_hide_substatus").value = "HIDE";
//  }
//}

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
  eval('window.open("action_line_summary.asp?screen_mode=PRINT&useSessions=1", "_printerfriendly", "width='+w+',height='+h+',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
}

function saveSearchOptions(iCustomReportID) {
  if(validateFields()) {
     //Validate the checkboxes
     lcl_statusDismissed  = "no";
     lcl_statusResolved   = "no";
     lcl_statusWaiting    = "no";
     lcl_statusInprogress = "no";
     lcl_statusSubmitted  = "no";

     if(document.getElementById("statusDismissed").checked) {
      		lcl_statusDismissed = document.getElementById("statusDismissed").value;
     }
     if(document.getElementById("statusResolved").checked) {
      		lcl_statusResolved = document.getElementById("statusResolved").value;
     }
     if(document.getElementById("statusWaiting").checked) {
      		lcl_statusWaiting = document.getElementById("statusWaiting").value;
     }
     if(document.getElementById("statusInprogress").checked) {
      		lcl_statusInprogress = document.getElementById("statusInprogress").value;
     }
     if(document.getElementById("statusSubmitted").checked) {
      		lcl_statusSubmitted = document.getElementById("statusSubmitted").value;
     }

     //Build the parameter string
   		var sParameter = 'customreportid='           + encodeURIComponent(iCustomReportID);

     sParameter    += '&isAjaxRoutine=Y';
     sParameter    += '&selectAssignedto='        + encodeURIComponent(document.getElementById("selectAssignedto").value);
   		sParameter    += '&orderBy='                 + encodeURIComponent(document.getElementById("orderBy").value);
   		//sParameter    += '&recordsPer='              + encodeURIComponent(document.getElementById("recordsPer").value);
   		sParameter    += '&reporttype='              + encodeURIComponent(document.getElementById("reporttype").value);
   		sParameter    += '&selectDeptId='            + encodeURIComponent(document.getElementById("selectDeptId").value);
   		sParameter    += '&pastDays='                + encodeURIComponent(document.getElementById("pastDays").value);
   		sParameter    += '&searchDaysType='          + encodeURIComponent(document.getElementById("searchDaysType").value);
   		sParameter    += '&fromDate='                + encodeURIComponent(document.getElementById("fromDate").value);
   		sParameter    += '&toDate='                  + encodeURIComponent(document.getElementById("toDate").value);
   		sParameter    += '&fromToDateSelection='     + encodeURIComponent(document.getElementById("fromToDateSelection").value);
   		sParameter    += '&selectDateType='          + encodeURIComponent(document.getElementById("selectDateType").value);
   		sParameter    += '&statusDismissed='         + encodeURIComponent(lcl_statusDismissed);
   		sParameter    += '&statusResolved='          + encodeURIComponent(lcl_statusResolved);
   		sParameter    += '&statusWaiting='           + encodeURIComponent(lcl_statusWaiting);
   		sParameter    += '&statusInprogress='        + encodeURIComponent(lcl_statusInprogress);
   		sParameter    += '&statusSubmitted='         + encodeURIComponent(lcl_statusSubmitted);

  <% if lcl_orghasfeature_action_line_substatus AND lcl_userhaspermission_action_line_substatus then %>
   		sParameter    += '&substatus_hidden='        + encodeURIComponent(document.getElementById("substatus_hidden").value);
  <% end if %>

   		sParameter    += '&selectUserFName='         + encodeURIComponent(document.getElementById("selectUserFName").value);
   		sParameter    += '&selectUserLName='         + encodeURIComponent(document.getElementById("selectUserLName").value);
   		//sParameter    += '&selectIssueStreetNumber=' + encodeURIComponent(document.getElementById("selectIssueStreetNumber").value);
   		//sParameter    += '&selectIssueStreet='       + encodeURIComponent(document.getElementById("selectIssueStreet").value);
   		//sParameter    += '&selectContactStreet='     + encodeURIComponent(document.getElementById("selectContactStreet").value);
   		//sParameter    += '&selectCounty='            + encodeURIComponent(document.getElementById("selectCounty").value);
   		sParameter    += '&selectBusinessName='      + encodeURIComponent(document.getElementById("selectBusinessName").value);
   		sParameter    += '&selectTicket='            + encodeURIComponent(document.getElementById("selectTicket").value);

  <% if UCASE(reporttype) <> "STATUSSUMMARY" then %>
   		sParameter    += '&selectInitialResponse='   + encodeURIComponent(document.getElementById("selectInitialResponse").value);
   		sParameter    += '&selectRequestsResolved='  + encodeURIComponent(document.getElementById("selectRequestsResolved").value);
  <% end if %>

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

function setupPrintButtons() {
		//factory.printing.header       = "Printed on &d"
		factory.printing.footer       = "&bPrinted on &d - Page:&p/&P";
		factory.printing.portrait     = false;
		factory.printing.leftMargin   = 0.5;
		factory.printing.topMargin    = 0.5;
		factory.printing.rightMargin  = 0.5;
		factory.printing.bottomMargin = 0.5;
		 
		//enable control buttons
		var templateSupported = factory.printing.IsTemplateSupported();
		var controls = idControls.all.tags("input");
		for ( i = 0; i < controls.length; i++ ) {
  			controls[i].disabled = false;
		  	if ( templateSupported && controls[i].className == "ie55" ) {
			     controls[i].style.display = "inline";
     }
  }
}

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
<!-- <body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"> -->
<%
if lcl_screen_mode <> "PRINT" then

   lcl_onload = ""

  'If any substatuses have been checked and posted after clicking the Search button then we need to show the 'display' list
   if lcl_userhaspermission_action_line_substatus then
      if substatus_hidden <> "" then
         lcl_display_substatuses = "change_substatus_filter();"
      else
         lcl_display_substatuses = ""
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

  'Build the onLoad
   if lcl_onload <> "" then
      lcl_onload = " onload=""javascript:" & lcl_onload & """ "
   else
      lcl_onload = ""
   end if
else
   lcl_onload = " onload=""setupPrintButtons();"""
end if

response.write "<body bgcolor=""#ffffff"" leftmargin=""0"" topmargin=""0"" marginheight=""0"" marginwidth=""0""" & lcl_onload & ">" & vbcrlf

'Display the navigation bar and search criteria if the screen mode is not PRINT
 if lcl_screen_mode <> "PRINT" then

    'DrawTabs tabActionline,1
    ShowHeader sLevel
%>
    <!--#Include file="../menu/menu.asp"--> 
<%
    response.write "<table border=""0"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
    response.write "  <tr valign=""top"">" & vbcrlf
    response.write "      <td width=""60%"">" & vbcrlf
    'response.write "          <font size=""+1""><strong>(E-Gov Request Manager) - Manage Action Line Requests</strong></font><br />" & vbcrlf
    'response.write "          <img src=""../images/arrow_2back.gif"" align=""absmiddle"" />&nbsp;" & vbcrlf
    'response.write "          <a href=""javascript:history.back();"">" & langBackToStart & "</a>" & vbcrlf
    response.write "          <font size=""+1""><strong>Manage Action Line Requests [E-Gov Request Manager]</strong></font>" & vbcrlf
    'response.write "          <input type=""button"" name=""backButton"" id=""backButton"" value=""<< Back"" class=""button"" onclick=""javascript:history.back();"" />" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "      <td width=""40%"" align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
    response.write "  <tr valign=""top"">" & vbcrlf
    response.write "      <td colspan=""2"">" & vbcrlf

   'BEGIN: Search/Sort Options --------------------------------------------------
    response.write "          <fieldset class=""fieldset"">" & vbcrlf
    response.write "            <legend><strong>Search/Sorting Option(s)&nbsp;</strong></legend>" & vbcrlf
    response.write "            <form name=""form1"" method=""post"" onSubmit=""return checkStat()"">" & vbcrlf
    response.write "            <table border=""0"" bordercolor=""#ff0000"" cellspacing=""0"" cellpadding=""2"" width=""100%"">" & vbcrlf
    response.write "              <tr valign=""top"">" & vbcrlf

   'Assigned To
    response.write "                  <td nowrap>" & vbcrlf
    response.write "                      <strong>Assigned To: " & vbcrlf

    if ViewAll = 0 then
       response.write "(User " & session("userID") & ")&nbsp;&nbsp;&nbsp;"
    end if

   'Draw list of employees
    DrawAssignedEmployeeSelection selectAssignedto

   'Group By (Order By)
    response.write "                      Group By: " & vbcrlf
    response.write "                      <select name=""orderBy"" id=""orderBy"">" & vbcrlf
                                            displayOrderByList orderBy, _
                                                               lcl_orghasfeature_issuelocation, _
                                                               lcl_orghasfeature_actionline_maintain_duedate, _
                                                               lcl_userhaspermission_actionline_maintain_duedate
    response.write "                      </select>" & vbcrlf
    response.write "                      </strong>" & vbcrlf
    response.write "                  </td>" & vbcrlf

   'Save Search Options
    if lcl_orghasfeature_savesearchoptions_actionline then

       response.write "                  <td align=""center"" width=""400"" rowspan=""4"">" & vbcrlf

       displayCustomSearchOptions lcl_customreportid_actionline_user

       response.write "                  </td>" & vbcrlf
    else
       response.write "                  <td align=""center"" width=""400"" rowspan=""4"">&nbsp;</td>" & vbcrlf
    end if

    response.write "              </tr>" & vbcrlf
    response.write "              <tr>" & vbcrlf

   'Status
	   if statusSubmitted  = "yes" then check1 = " checked=""checked"""
    if statusInprogress = "yes" then check2 = " checked=""checked""" 
	   if statusWaiting    = "yes" then check3 = " checked=""checked"""
    if statusResolved   = "yes" then check4 = " checked=""checked"""
	   if statusDismissed  = "yes" then check5 = " checked=""checked"""

    response.write "                  <td valign=""top"" nowrap>" & vbcrlf
    response.write "                      <strong>Status:</strong> " & vbcrlf
                                          displayStatusCheckbox "Submitted",   check1
                                          displayStatusCheckbox "In Progress", check2
                                          displayStatusCheckbox "Waiting",     check3
                                          displayStatusCheckbox "Resolved",    check4
                                          displayStatusCheckbox "Dismissed",   check5

                                          displaySubStatusOptions lcl_userhaspermission_action_line_substatus, _
                                                                  substatus_hidden, _
                                                                  show_hide_substatus

    response.write "                  </td>" & vbcrlf
    response.write "              </tr>" & vbcrlf
    response.write "              <tr>" & vbcrlf
    response.write "               	  <td valign=""top"" nowrap>" & vbcrlf

   'Category
    response.write "                      <strong>Categories and Forms: </strong>" & vbcrlf
    response.write "                      <select name=""selectFormId"" id=""selectFormId"">" & vbcrlf
    response.write "                        <option value="""">All Categories</option>" & vbcrlf
                                            fnListForms selectFormID
    response.write "                      </select>" & vbcrlf
    response.write "               			</td>" & vbcrlf
    response.write "              </tr>" & vbcrlf
    response.write "              <tr>" & vbcrlf
    response.write "              		  <td valign=""top"" nowrap>" & vbcrlf

   'Department
    response.write "                      <strong>Department: </strong> " & vbcrlf
    response.write "                      <select name=""selectDeptId"" id=""selectDeptId"">" & vbcrlf
    response.write "                        <option value=""all"">All Departments</option>" & vbcrlf

   'Get a list of all available departments for THIS user
    fnListDepts selectDeptId

    response.write "                      </select>&nbsp;&nbsp;&nbsp;"

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
    response.write "              		  <td valign=""top"" nowrap>" & vbcrlf

   'BEGIN: Date Range ---------------------------------------------------------
    response.write "                      <fieldset class=""fieldset"">" & vbcrlf

   'From Date
    response.write "                        <strong>From: </strong>" & vbcrlf
    response.write "                        <input type=""text"" name=""fromDate"" id=""fromDate"" value=""" & fromDate & """ size=""10"" maxlength=""10"" onchange=""clearMsg('fromDateCalPop');"" />" & vbcrlf
    response.write "                        <a href=""javascript:void doCalendar('From');""><img src=""../images/calendar.gif"" id=""fromDateCalPop"" border=""0"" onclick=""clearMsg('fromDateCalPop');"" /></a>&nbsp;" & vbcrlf

   'To Date
    response.write "                        <strong>To:</strong>" & vbcrlf
    response.write "                        <input type=""text"" name=""toDate"" id=""toDate"" value=""" & dateAdd("d",-1,toDate) & """ size=""10"" maxlength=""10"" onchange=""clearMsg('toDateCalPop');"" />" & vbcrlf
    response.write "                        <a href=""javascript:void doCalendar('To');""><img src=""../images/calendar.gif"" id=""toDateCalPop"" border=""0"" onclick=""clearMsg('toDateCalPop');"" /></a>&nbsp;" & vbcrlf

   'From/To Dates Date Range options
    DrawDateChoices "Date", fromToDateSelection

   'From/To Dates will search on options
    response.write "                        <br /><br />" & vbcrlf
    response.write "                        <strong>From/To Dates will search on:</strong>" & vbcrlf
    response.write "                        <select name=""selectDateType"" id=""selectDateType"">" & vbcrlf

    if UCASE(selectDateType) = "SUBMIT" then
       lcl_selected_active = ""
       lcl_selected_submit = " selected"
    else
       lcl_selected_active = " selected"
       lcl_selected_submit = ""
    end if

    if selectUserFName <> "all" then
       selectUserFName = selectUserFName
    'else
    '   selectUserFName = ""
    end if

    if selectUserLName <> "all" then
       selectUserLName = selectUserLName
    'else
    '   selectUserLName = ""
    end if

    response.write "                          <option value=""active""" & lcl_selected_active & ">Active Requests</option>" & vbcrlf
    response.write "                          <option value=""submit""" & lcl_selected_submit & ">Submit Date</option>" & vbcrlf
    response.write "                        </select>" & vbcrlf
    response.write "                      </fieldset>" & vbcrlf
   'END: Date Range -----------------------------------------------------------

    response.write "                  </td>" & vbcrlf
    response.write "                  <td>&nbsp;</td>" & vbcrlf
    response.write "              </tr>" & vbcrlf
    response.write "              <tr>" & vbcrlf
    response.write "              				<td valign=""top"" nowrap>" & vbcrlf
    response.write "                      <strong>Submitted By: &nbsp;&nbsp;" & vbcrlf
    response.write "                      First: <input type=""text"" name=""selectUserFName"" id=""selectUserFName"" value=""" & selectUserFName & """ size=""12"" />&nbsp;" & vbcrlf
    response.write "                      Last:</strong>" & vbcrlf
    response.write "                      <input type=""text"" name=""selectUserLName"" id=""selectUserLName"" value=""" & selectUserLName & """ size=""12"" />" & vbcrlf
    response.write "              				</td>" & vbcrlf

    if UCASE(reporttype) = "SUMMARY" then
       response.write "                  <td align=""right"" valign=""bottom"" rowspan=""6"">" & vbcrlf
       response.write "                      <fieldset class=""fieldset"" align=""right"">" & vbcrlf
       response.write "                        <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
       response.write "                          <tr>" & vbcrlf
       response.write "                              <td colspan=""2"" align=""center"" class=""redText"">" & vbcrlf
       response.write "                                  *** To be used to calculate percentages in results ***" & vbcrlf
       response.write "                                  <hr size=""1"">" & vbcrlf
       response.write "                              </td>" & vbcrlf
       response.write "                          </tr>" & vbcrlf
       'response.write "                          <tr>" & vbcrlf
       'response.write "                              <td><strong>% of Requests with Initial Response within&nbsp;</strong></td>" & vbcrlf
       'response.write "                              <td>" & vbcrlf
       'response.write "                                  <input type=""text"" name=""selectInitialResponse"" id=""selectInitialResponse"" value=""" & selectInitialResponse & """ size=""3"" maxlength=""3"" />" & vbcrlf
       'response.write "                                  <strong>&nbsp;" & lcl_adjustment_label & "</strong>" & vbcrlf
       'response.write "                              </td>" & vbcrlf
       'response.write "                          </tr>" & vbcrlf
       'response.write "                          <tr>" & vbcrlf
       'response.write "                              <td><strong>% of Requests Resolved within&nbsp;</strong></td>" & vbcrlf
       'response.write "                              <td>" & vbcrlf
       'response.write "                                  <input type=""text"" name=""selectRequestsResolved"" id=""selectRequestsResolved"" value=""" & selectRequestsResolved & """ size=""3"" maxlength=""3"" />" & vbcrlf
       'response.write "                                  <strong>&nbsp;" & lcl_adjustment_label & "</strong>" & vbcrlf
       'response.write "                              </td>" & vbcrlf
       'response.write "                          </tr>" & vbcrlf
       response.write "                          <tr>" & vbcrlf
       response.write "                              <td align=""center"">" & vbcrlf
       response.write "                                  <strong>% of Requests with Initial Response<br />within&nbsp;</strong>" & vbcrlf
       response.write "                                  <input type=""text"" name=""selectInitialResponse"" id=""selectInitialResponse"" value=""" & selectInitialResponse & """ size=""3"" maxlength=""3"" />" & vbcrlf
       response.write "                                  <strong>&nbsp;" & lcl_adjustment_label & "</strong>" & vbcrlf
       response.write "                                  <hr size=""1"">" & vbcrlf
       response.write "                              </td>" & vbcrlf
       response.write "                          </tr>" & vbcrlf
       response.write "                          <tr>" & vbcrlf
       response.write "                              <td align=""center"">" & vbcrlf
       response.write "                                  <strong>% of Requests Resolved within&nbsp;</strong>" & vbcrlf
       response.write "                                  <input type=""text"" name=""selectRequestsResolved"" id=""selectRequestsResolved"" value=""" & selectRequestsResolved & """ size=""3"" maxlength=""3"" />" & vbcrlf
       response.write "                                  <strong>&nbsp;" & lcl_adjustment_label & "</strong>" & vbcrlf
       response.write "                              </td>" & vbcrlf
       response.write "                          </tr>" & vbcrlf
       response.write "                        </table>" & vbcrlf
       response.write "                      </fieldset>" & vbcrlf
       response.write "                  </td>" & vbcrlf
    else
       response.write "                  <td align=""right"" rowspan=""3"">&nbsp</td>" & vbcrlf
    end if

    response.write "              </tr>" & vbcrlf

   'Contact Street Name
    response.write "              <tr>" & vbcrlf
    response.write "                  <td valign=""top"" nowrap>" & vbcrlf
    response.write "                      <strong>Contact Street Name:</strong>&nbsp;" & vbcrlf
    response.write "                      <input name=""selectContactStreet"" id=""selectContactStreet"" value=""" & selectContactStreet & """ type=""text"" />" & vbcrlf
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

      'County (custom field)
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
    response.write "                      <input type=""text"" name=""selectTicket"" id=""selectTicket"" value=""" & selectTicket & """ size=""15"" maxlength=""10"" />" & vbcrlf
    response.write "                  </td>" & vbcrlf
    response.write "              </tr>" & vbcrlf

   'Determine which "search day type" is "selected"
    lcl_selected_searchdaystype_open = ""
    lcl_selected_searchdaystype_past = ""

    if searchDaysType = "PAST" then
       lcl_selected_searchdaystype_open = ""
       lcl_selected_searchdaystype_past = " selected=""selected"""
    else
       lcl_selected_searchdaystype_open = " selected=""selected"""
       lcl_selected_searchdaystype_past = ""
    end if

   'Display Open Over
    'if pastDays <> "all" then
    '   lcl_pastDays = pastDays
    'else
    '   lcl_pastDays = ""
    'end if

    lcl_pastDays = ""

    if pastDays <> "all" then
       if searchDaysType = "PAST" OR (searchDaysType = "OPEN" and pastDays <> "0") then
          lcl_pastDays = pastDays
       end if
    end if

    'response.write "              <tr>" & vbcrlf
    'response.write "                  <td valign=""top"">" & vbcrlf
    'response.write "                      <strong>Display Open Over " & vbcrlf
    'response.write "                      <input type=""text"" name=""pastDays"" id=""pastDays"" value=""" & lcl_pastDays & """ size=""2"" onchange=""clearMsg('pastDays');"" /> Days</strong>" & vbcrlf
    'response.write "                  </td>" & vbcrlf
    'response.write "              </tr>" & vbcrlf

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

   'Search Button
    response.write "              <tr>" & vbcrlf
    response.write "                  <td valign=""top"" colspan=""2"">" & vbcrlf
    response.write "                      <input type=""button"" name=""searchButton"" id=""searchButton"" value=""SEARCH"" class=""button"" onclick=""clearScreenMsg();submitForm();"" />" & vbcrlf
    response.write "                  </td>" & vbcrlf
    response.write "              </tr>" & vbcrlf
    response.write "              </form>" & vbcrlf
    response.write "            </table>" & vbcrlf
    response.write "          </fieldset>" & vbcrlf
   'END: Search/Sort Options --------------------------------------------------

    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf

'------------------------------------------------------------------------------
 else  'lcl_screen_mode = "PRINT"
'------------------------------------------------------------------------------
   'BEGIN: THIRD PARTY PRINT CONTROL ------------------------------------------
    response.write "<div id=""idControls"" class=""noprint"">" & vbcrlf
    response.write "  <input type=""button"" value=""Print the page"" disabled onclick=""factory.printing.Print(true)"" />&nbsp;&nbsp;" & vbcrlf
    response.write "  <input type=""button"" value=""Print Preview..."" disabled onclick=""factory.printing.Preview()"" class=""ie55"" />" & vbcrlf
    response.write "</div>" & vbcrlf
    response.write "<object id=""factory"" viewastext  style=""display:none"" classid=""clsid:1663ed61-23eb-11d2-b92f-008048fdd814"" codebase=""../includes/smsx.cab#Version=6,3,434,12""></object>" & vbcrlf
   'END: THIRD PARTY PRINT CONTROL --------------------------------------------

    response.write "<table border=""0"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
 end if

 response.write "  <tr>" & vbcrlf
 response.write "      <td valign=""top"" colspan=""2"">" & vbcrlf

'BEGIN: Action Line Request List ----------------------------------------------
 response.write "          <form name=""requestlist"" action=""#"" method=""post"">" & vbcrlf
 'response.write "            <input type=""hidden"" name=""selectIssueStreetNumber"" id=""selectIssueStreetNumber"" value=""" & selectIssueStreetNumber & """ size=""10"" maxlength=""150"" />&nbsp;" & vbcrlf
 'response.write "            <input type=""hidden"" name=""selectIssueStreet"" id=""selectIssueStreet"" value=""" & selectIssueStreet & """ size=""30"" maxlength=""300"" />" & vbcrlf
 response.write "            <input type=""hidden"" name=""selectContactStreet"" id=""selectContactStreet"" value=""" & selectContactStreet & """ />" & vbcrlf
 response.write "            <input type=""hidden"" name=""selectBusinessName"" id=""selectBusinessName"" value=""" & selectBusinessName & """ />" & vbcrlf

 List_Action_Requests sSortBy, pastDays, searchDaysType

 response.write "          </form>" & vbcrlf
'END: Action Line Request List ------------------------------------------------

 response.write "      </td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "</table>" & vbcrlf

 if lcl_screen_mode <> "PRINT" then
%>
<!-- #include file="../admin_footer.asp" //-->  
<%
 end if

 response.write "</body>" & vbcrlf
 response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
Function List_Action_Requests(sSortBy, iPastDays, iSearchDateType)

Dim statArray(5)
	i = 0

if statusSubmitted = "yes" then 
	  statArray(i) = " UPPER(status)='SUBMITTED' OR"
	  i = i + 1
end if

if statusInprogress = "yes" then
	  statArray(i) = " UPPER(status)='INPROGRESS' OR"
  	i = i + 1
end if

if statusWaiting = "yes" then
  	statArray(i) = " UPPER(status)='WAITING' OR"
  	i = i + 1
end if

if statusResolved= "yes" then
  	statArray(i) = " UPPER(status)='RESOLVED' OR"
  	i = i + 1
end if

if statusDismissed = "yes" then
  	statArray(i) = " UPPER(status)='DISMISSED' OR"
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
 lcl_query_toDate = dateAdd("d",1,toDate)

if UCASE(selectDateType) = "ACTIVE" then
   varWhereClause = " WHERE egov_action_request_view.orgid=('"&session("orgid")&"') AND ( "    ''IsNull(complete_date,'" & Now & "')
   varWhereClause = varWhereClause & " (submit_date >= '" & fromDate & "' AND submit_date < '" & lcl_query_toDate & "') OR "
   varWhereClause = varWhereClause & " ( IsNull(complete_date,'" & Now & "') >= '" & fromDate & "' AND IsNull(complete_date,'" & Now & "') < '" & lcl_query_toDate & "' ) OR "
   varWhereClause = varWhereClause & " (submit_date < '" & fromDate & "' AND IsNull(complete_date,'" & Now & "') > '" & lcl_query_toDate & "')  "

   'varWhereClause = varWhereClause & " (submit_date >= '" & fromDate & "' AND submit_date < '" & toDate & "') OR "
   'varWhereClause = varWhereClause & " ( IsNull(complete_date,'" & Now & "') >= '" & fromDate & "' AND IsNull(complete_date,'" & Now & "') < '" & toDate & "' ) OR "
   'varWhereClause = varWhereClause & " (submit_date < '" & fromDate & "' AND IsNull(complete_date,'" & Now & "') > '" & toDate & "')  "
else 'selectDateType = SUBMIT
   varWhereClause = " WHERE egov_action_request_view.orgid=" & session("orgid")
   varWhereClause = varWhereClause & " AND (submit_date BETWEEN '" & fromDate & "' AND '" & lcl_query_toDate & "'"

   'varWhereClause = varWhereClause & " AND (submit_date BETWEEN '" & fromDate & "' AND '" & toDate & "'"
end if

'Sub-Status Filter
if substatus_hidden = "" OR isnull(substatus_hidden) then
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
      varWhereClause = varWhereClause & " OR sub_status_id in (" & REPLACE(REPLACE(substatus_hidden,"(",""),")","") & ")) "
   else
      varWhereClause = varWhereClause & " ) AND sub_status_id in (" & REPLACE(REPLACE(substatus_hidden,"(",""),")","") & ") "
   end if
end if

'restrict what they can view by the permission level they have
If blnCanViewDeptActionItems AND NOT blnCanViewAllActionItems Then
 	'can view dept
  	If selectDeptId <> "all" then 
		    varWhereClause = varWhereClause & " AND deptID = '" & selectDeptId & "' " 
  	Else
       'varWhereClause = varWhereClause & " AND deptID IN (" & GetGroups(session("userid")) & ") OR (assignedemployeeid = '" & session("userid") & "') " 
     		varWhereClause = REPLACE(varWhereClause,"AND (" & varStatClause & ")","") & " AND (((" & varStatClause & ") AND deptID IN (" & GetGroups(session("userid")) & ")) OR ((assignedemployeeid = '" & session("userid") & "') and (" & varStatClause  & "))) "
  	End If
Else
 	If blnCanViewOwnActionItems And Not blnCanViewAllActionItems And Not blnCanViewDeptActionItems Then
	   'can view own only
   		varWhereClause = varWhereClause & " AND assignedemployeeid = " & session("userid") & " "
  Else 
		  'Can view all, only add to where clause if a dept is chosen
   		If selectDeptId <> "all" then 
			     varWhereClause = varWhereClause & " AND deptID = '" & selectDeptId & "'" 
   		End If
 	End If 
End If

If selectFormId <> "all" then 
  	If left(selectFormId,1)="C" then 
		    sSQLb = "SELECT action_form_id FROM egov_forms_to_categories where form_category_id = " & right(selectFormId,len(selectFormId)-1)
    		Set oCategories = Server.CreateObject("ADODB.Recordset")

    		oCategories.Open sSQLb, Application("DSN"), 3, 1

    		if oCategories.EOF then
      			varWhereClause = varWhereClause & " AND form_category_id=999999"		
    		else
      			do while not oCategories.EOF

         			CatArray = CatArray & oCategories("action_form_id") & ","

          		oCategories.MoveNext
       		Loop
         CatArray = left(CatArray,(len(CatArray)-1))
   		 end if

      oCategories.close
      set oCategories = nothing

     	varWhereClause = varWhereClause & " AND action_Formid IN (" & CatArray & ") "
    		'varWhereClause = varWhereClause & " AND form_category_id = " & right(selectFormId,len(selectFormId)-1)
   else
     	varWhereClause = varWhereClause & " AND action_Formid = " & selectFormId
  	end if
End If

'if selectAssignedto <> "all" then varWhereClause = varWhereClause & " AND assigned_Name = '" & selectAssignedto & "'" 
'if selectIssueStreet <> "all" then varWhereClause = varWhereClause & " AND streetname LIKE '" & selectIssueStreet & "%'"
 if selectAssignedto <> "all" AND selectAssignedto <> "" then 
    varWhereClause = varWhereClause & " AND assignedemployeeid = " & selectAssignedto
 end if

 if selectDeptId <> "all" AND selectDeptId <> "" then
    varWhereClause = varWhereClause & " AND deptID = " & selectDeptId & " "
 end if

 if iSelectUserFName <> "all" AND iSelectUserFName <> "" then
    varWhereClause = varWhereClause & " AND upper(UserFName) LIKE '%" & dbsafe(ucase(iSelectUserFName)) & "%'"
 end if

 if iSelectUserLName <> "all" AND iSelectUserLName <> "" then
    varWhereClause = varWhereClause & " AND upper(UserLName) LIKE '%" & dbsafe(ucase(iSelectUserLName)) & "%'"
 end if

'Contact Street Name
 if selectContactStreet <> "all" then varWhereClause = varWhereClause & " AND useraddress LIKE '" & dbsafe(selectContactStreet) & "%'"

'Issue/Problem Location
 lcl_search_address = ""
 if selectIssueStreetNumber <> "" AND NOT isnull(selectIssueStreetNumber) then
    lcl_search_address = selectIssueStreetNumber & "%"
 end if

 lcl_search_address = lcl_search_address & selectIssueStreet

 if lcl_search_address <> "" then
    varWhereClause = varWhereClause & " AND UPPER(streetname) LIKE ('%" & dbsafe(UCASE(lcl_search_address)) & "%')"
 end if

'County
 if selectCounty <> "" AND NOT isnull(selectCounty) then
    varWhereClause = varWhereClause & " AND UPPER(county) LIKE ('%" & UCASE(dbsafe(selectCounty)) & "%')"
 end if

'Business Name
 if selectBusinessName <> "all" then varWhereClause = varWhereClause & " AND userbusinessname LIKE '" & dbsafe(selectBusinessName) & "%'"

'Tracking Number search
 if selectTicket <> "" then
    if IsNumeric(selectTicket) then
	     'remove the right 4 which are the time
      	sTicketNo = Left( selectTicket, (Len(selectTicket) - 4))
      	iTrackID  = CStr(CLng(selectTicket))
      	iTime     = Right(iTrackID,4)
      	iHour     = Left(iTime,2)
     	 iMinute   = Right(iTime,2)

      	if iHour = "" or iMinute = "" then
        		iHour   = "99"
        		iMinute = "99"
      	end if

      	varWhereClause = varWhereClause & " and (action_autoid = " & sTicketNo & ") AND (DATEPART(hh, submit_date) = '"& iHour &"') AND (DATEPART(mi, submit_date) = '"& iMinute &"')"
    else
      'If this isn't passed then it will appear that the summary isn't working correctly because a value has been entered
      'for the Tracking Number search criteria field.  Using this statement will return the "No Records Found" in the results.
       varWhereClause = varWhereClause & " and 1=2"
    end if
 end if

'This is the main query
 sSQL = setupGroupbyQuery(orderBy, varWhereClause, reporttype, iPastDays, iSearchDateType)

 set oRequests = Server.CreateObject("ADODB.Recordset")
 oRequests.Open sSQL, Application("DSN"), 3, 1

'Save "Last Queried" Search Options.
 lcl_success = "Y"

 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectAssignedto",        selectAssignedto,        False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "orderBy",                 orderBy,                 False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "recordsPer",              recordsPer,              False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "reporttype",              reporttype,              False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectFormId",            selectFormId,            False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectDeptId",            selectDeptId,            False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "pastDays",                pastDays,                False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "fromDate",                fromDate,                False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "toDate",                  toDate,                  False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "fromToDateSelection",     fromToDateSelection,     False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectDateType",          selectDateType,          False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusDismissed",         statusDismissed,         False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusResolved",          statusResolved,          False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusWaiting",           statusWaiting,           False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusInprogress",        statusInprogress,        False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "statusSubmitted",         statusSubmitted,         False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "substatus_hidden",        substatus_hidden,        False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectUserFName",         selectUserFName,         False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectUserLName",         selectUserLName,         False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectIssueStreetNumber", selectIssueStreetNumber, False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectIssueStreet",       selectIssueStreet,       False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectContactStreet",     selectContactStreet,     False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectCounty",            selectCounty,            False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectBusinessName",      selectBusinessName,      False, lcl_success
 saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectTicket",            selectTicket,            False, lcl_success

 if UCASE(reporttype) <> "STATUSSUMMARY" then
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectInitialResponse",   selectInitialResponse,   False, lcl_success
    saveCustomReportSearchOption lcl_customreportid_actionline_lastqueried, "selectRequestsResolved",  selectRequestsResolved,  False, lcl_success
 end if
	 
 if not oRequests.EOF then

   'Set up the OrderBy Column Label
	 	 if orderBy = "submit_date" then
 	 				lcl_orderBy_columnLabel = "Date"
    elseif orderBy = "due_date" then
       lcl_orderBy_columnLabel = "Due Date"
		  elseif orderBy = "streetname" then
  					lcl_orderBy_columnLabel = "Issue/Problem Location Street Name"
	 	 elseif orderBy = "action_Formid" then
 	 				lcl_orderBy_columnLabel = "Action Line Form"
		  elseif orderBy = "deptId" then
 					lcl_orderBy_columnLabel = "Department"
 		 elseif orderBy = "assigned_Name" then
  					lcl_orderBy_columnLabel = "Assigned To"
		  elseif orderBy = "submittedby" then
  					lcl_orderBy_columnLabel = "Submitted By"
    elseif orderBy = "status" then
       lcl_orderBy_columnLabel = "Status"
	 	 end if

 	  response.write "<div>" & vbcrlf
	   response.write "<table width=""100%"">" & vbcrlf
	   response.write "  <tr>" & vbcrlf
    response.write "      <td valign=""top"">" & vbcrlf
    response.write "          <font size=""3"" color=""#3399ff""><i><strong>" & lcl_reportlabel & " REPORT</strong></i></font>" & vbcrlf
    response.write "      </td>" & vbcrlf

    if lcl_screen_mode <> "PRINT" then
	      response.write "      <td width=""450"" align=""right"">" & vbcrlf
       response.write "          <input type=""button"" name=""printerFriendlyButton"" id=""printerFriendlyButton"" value=""Printer Friendly Results"" class=""button"" onclick=""openPrinterFriendlyResults()"" />" & vbcrlf
    	  response.write "      </td>" & vbcrlf
    end if

 	  response.write "  </tr>" & vbcrlf
	   response.write "</table>" & vbcrlf
	   response.write "</div>" & vbcrlf
	   response.write "<table border=""0"" cellspacing=""0"" cellpadding=""5"" class=""tablelist"" width=""100%"">" & vbcrlf
 	  response.write "  <tr class=""tablelist"">" & vbcrlf
  		response.write "      <th>" & lcl_orderBy_columnLabel & "</th>" & vbcrlf

    if UCASE(reporttype) = "STATUSSUMMARY" then
       response.write "      <th>Submitted</th>"    & vbcrlf
       response.write "      <th>In Progress</th>"  & vbcrlf
       response.write "      <th>Waiting</th>"      & vbcrlf
       response.write "      <th>Total Open</th>"   & vbcrlf
       response.write "      <th>Resolved</th>"     & vbcrlf
       response.write "      <th>Dismissed</th>"    & vbcrlf
       response.write "      <th>Total Closed</th>" & vbcrlf
       response.write "      <th>Grand Total</th>"  & vbcrlf
    else
    	  response.write "      <th>Total</th>"  & vbcrlf
       response.write "      <th>Open</th>" & vbcrlf

    	  if pastDays <> "all" then
          if lcl_orghasfeature_actionline_maintain_duedate AND lcl_userhaspermission_actionline_maintain_duedate AND iSearchDateType = "PAST" then
             response.write "      <th>" & pastDays & " day(s) past Due Date</th>" & vbcrlf
          else
   	      		 response.write "      <th>Open Items Over " & pastDays & " days</th>" & vbcrlf
          end if
    	  end if

    	  response.write "      <th>Avg. Time still Open</th>"  & vbcrlf
	      response.write "      <th>Avg. Time to Complete</th>" & vbcrlf

    	  response.write "      <th>% of Requests with<br>Initial Response within<br>" & selectInitialResponse & " " & lcl_adjustment_label & "</th>" & vbcrlf
	      response.write "      <th>% of Requests Resolved<br>within " & selectRequestsResolved & " " & lcl_adjustment_label & "</th>" & vbcrlf
    end if

 	  response.write "  </tr>" & vbcrlf

    totalSubmitted = 0
    totalOpen      = 0
    totalDays      = 0
    totalPast      = 0

   'LOOP AND DISPLAY THE RECORDS
	   bgcolor       = "#eeeeee"
    lcl_linecnt   = 0
    lcl_SQLopen   = ""
    lcl_SQLclosed = ""
    lcl_SQLpast   = ""
    lcl_SQLbdays  = ""

    lcl_columntotals_submitted   = 0
    lcl_columntotals_inprogress  = 0
    lcl_columntotals_waiting     = 0
    lcl_columntotals_resolved    = 0
    lcl_columntotals_dismissed   = 0

    lcl_columntotals_totalopen   = 0
    lcl_columntotals_totalclosed = 0
    lcl_columntotals_grandtotal  = 0


   'Set up the search options hidden fields
 		 do while not oRequests.eof
       bgcolor     = changeBGColor(bgcolor,"#eeeeee","#ffffff")
       lcl_linecnt = lcl_linecnt + 1

      'Get value(s) to display output
      '------------------------------------------------------------------------
   	   if UCASE(orderBy) = "SUBMIT_DATE" then
      '------------------------------------------------------------------------
          sTitle = formatTitle(oRequests("TheDate"))
          'lcl_search_criteria = " AND submitdateshort = '" & sTitle & "' "

          response.write "<input type=""hidden"" name=""sc_toDate" & lcl_linecnt & """ id=""sc_toDate" & lcl_linecnt & """ value=""" & oRequests("TheDate") & """ />" & vbcrlf
          response.write "<input type=""hidden"" name=""sc_fromDate" & lcl_linecnt & """ id=""sc_fromDate" & lcl_linecnt & """ value=""" & oRequests("TheDate") & """ />" & vbcrlf

      '------------------------------------------------------------------------
   	   elseif UCASE(orderBy) = "DUE_DATE" then
      '------------------------------------------------------------------------
          sTitle = formatTitle(oRequests("due_date"))
          'lcl_search_criteria = " AND submitdateshort = '" & sTitle & "' "

          response.write "<input type=""hidden"" name=""sc_toDate" & lcl_linecnt & """ id=""sc_toDate" & lcl_linecnt & """ value=""" & oRequests("due_date") & """ />" & vbcrlf
          response.write "<input type=""hidden"" name=""sc_fromDate" & lcl_linecnt & """ id=""sc_fromDate" & lcl_linecnt & """ value=""" & oRequests("due_date") & """ />" & vbcrlf

      '------------------------------------------------------------------------
       elseif UCASE(orderBy) = "ACTION_FORMID" then
      '------------------------------------------------------------------------
          sTitle = formatTitle(oRequests("action_formTitle"))
          'lcl_search_criteria = " AND action_formTitle = '" & dbsafe(sTitle) & "' "

          response.write "<input type=""hidden"" name=""sc_selectFormId" & lcl_linecnt & """ id=""sc_selectFormId" & lcl_linecnt & """ value=""" & oRequests("action_formId") & """ />" & vbcrlf
      '------------------------------------------------------------------------
       elseif UCASE(orderBy) = "DEPTID" then
      '------------------------------------------------------------------------
    						if oRequests("deptId") <> "" AND IsNull(oRequests("deptId")) = false then
 		    						sTitle = clng(oRequests("deptId"))
					 	   else
 					 			   sTitle = 0
   				 		end if

          'lcl_search_criteria = " AND deptId = " & sTitle

          response.write "<input type=""hidden"" name=""sc_selectDeptId" & lcl_linecnt & """ id=""sc_selectDeptId" & lcl_linecnt & """ value=""" & oRequests("DeptID") & """ />" & vbcrlf

    						sSQL = "SELECT groupname "
          sSQL = sSQL & " FROM groups "
          sSQL = sSQL & " WHERE orgid = " & session("OrgID")
          sSQL = sSQL & " AND groupid = " & sTitle

    			  	set oDeptName = Server.CreateObject("ADODB.Recordset")
 			   			oDeptName.Open sSQL, Application("DSN") , 3, 1

    						if oDeptName.EOF then
  		   						sTitle = "<font color=""#ff0000""><strong>???</strong></font>"
	  				   else
		   					   sTitle = oDeptName("groupname") 
   	 					end if

          oDeptName.close
          set oDeptName = nothing
      '------------------------------------------------------------------------
       elseif UCASE(orderBy) = "ASSIGNED_NAME" then
      '------------------------------------------------------------------------
          sTitle = formatTitle(trim(oRequests("assigned_name")))

          'if trim(oRequests("assigned_name")) <> "" then
          '   lcl_search_criteria = " AND assigned_Name = '" & sTitle & "' "
          '   lcl_search_criteria = lcl_search_criteria & " AND assignedemployeeid = " & oRequests("assignedemployeeid")
          '   lcl_search_criteria = lcl_search_criteria & " AND assigned_name <> '' "
          '   lcl_search_criteria = lcl_search_criteria & " AND assigned_name IS NOT NULL "
          'else
          '   lcl_search_criteria = " AND isnull(ltrim(rtrim(assigned_Name)),'') = '' "
          'end if

          response.write "<input type=""hidden"" name=""sc_selectAssignedto" & lcl_linecnt & """ id=""sc_selectAssignedto" & lcl_linecnt & """ value=""" & oRequests("assignedemployeeid") & """ />" & vbcrlf
      '------------------------------------------------------------------------
       elseif UCASE(orderBy) = "STREETNAME" then
      '------------------------------------------------------------------------
          sTitle = formatTitle(trim(oRequests("streetname")))
          'lcl_search_criteria = " AND streetname = '" & oRequests("streetname") & "'"
          'lcl_search_criteria = lcl_search_criteria & " AND streetnumber = '" & oRequests("streetnumber") & "'"

          lcl_streetnumber = ""
          lcl_streetname   = ""

          if trim(oRequests("streetnumber")) <> "" then
             lcl_streetnumber = trim(oRequests("streetnumber"))
          end if

          if trim(oRequests("streetname")) <> "" then
             lcl_streetname = trim(oRequests("streetname"))
          end if

          if lcl_streetnumber <> "" then
             lcl_formatted_streetname = replace(lcl_streetname,lcl_streetnumber,"")
          else
             lcl_formatted_streetname = lcl_streetname
          end if

          response.write "<input type=""hidden"" name=""sc_selectIssueStreetNumber" & lcl_linecnt & """ id=""sc_selectIssueStreetNumber" & lcl_linecnt & """ value=""" & trim(oRequests("streetnumber")) & """ />" & vbcrlf
          response.write "<input type=""hidden"" name=""sc_selectIssueStreet" & lcl_linecnt & """ id=""sc_selectIssueStreet" & lcl_linecnt & """ value=""" & lcl_formatted_streetname & """ />" & vbcrlf
          'response.write "<input type=""hidden"" name=""sc_selectIssueStreet" & lcl_linecnt & """ id=""sc_selectIssueStreet" & lcl_linecnt & """ value=""" & trim(oRequests("displaystreetaddress")) & """ />" & vbcrlf
      '------------------------------------------------------------------------
       elseif UCASE(orderBy) = "SUBMITTEDBY" then
      '------------------------------------------------------------------------
          sTitle = formatTitle(oRequests("userfname") & " " & oRequests("userlname"))
          'lcl_search_criteria = " AND userfname + ' ' + userlname = '" & sTitle & "' "

          response.write "<input type=""hidden"" name=""sc_selectUserLName" & lcl_linecnt & """ id=""sc_selectUserLName" & lcl_linecnt & """ value=""" & trim(oRequests("userlname")) & """ />" & vbcrlf
          response.write "<input type=""hidden"" name=""sc_selectUserFName" & lcl_linecnt & """ id=""sc_selectUserFName" & lcl_linecnt & """ value=""" & trim(oRequests("userfname")) & """ />" & vbcrlf

      '------------------------------------------------------------------------
   	   elseif UCASE(orderBy) = "STATUS" then
      '------------------------------------------------------------------------
          sTitle = formatTitle(oRequests("status"))
          'lcl_search_criteria = " AND UPPER(status) = '" & UCASE(sTitle) & "' "

          response.write "<input type=""hidden"" name=""sc_status" & lcl_linecnt & """ id=""sc_status" & lcl_linecnt & """ value=""" & trim(sTitle) & """ />" & vbcrlf

      '------------------------------------------------------------------------
       end if
      '------------------------------------------------------------------------

       if UCASE(reporttype) <> "STATUSSUMMARY" then
       		'SubTotal Submitted
   	 		   numSubmitted   = oRequests("numSubmitted")
       			totalSubmitted = totalSubmitted + numSubmitted

       		'BEGIN: SubTotal Open ------------------------------------------------
	 	   	   numOpen = CLng(oRequests("numOpen"))

       			if numOpen > 0 and oRequests("totalDaysOpen") > 0 then
	 	   	      avgOpen = (oRequests("totalDaysOpen") / numOpen)
		 	         avgOpen = formatnumber(avgOpen,1)
       			else
  	       			if numOpen > 0 then
		       	   		'The datediff is 0 but there are some open items, so they are from today.
    		 		   	   avgOpen = "< 1.0"
      				   else 
         		 			'response.write("'" & avgOpen & "'")
			           		avgOpen = "None Open"
         				end if
	 	       end if

       			totalOpen = totalOpen + numOpen
       		'END: SubTotal Open --------------------------------------------------

         'Calculate and format "% of Requests Resolved within x Business Day(s)"
          lcl_percent_3_days = ((oRequests("total_cnt_3_days") / numSubmitted)*100)
          lcl_percent_3_days = REPLACE(formatnumber(lcl_percent_3_days,3),".000","") & "%"

         'Calculate and format "% of Requests with Initial Response within x Business Day(s)"
          lcl_init_resp_1_days = ((oRequests("total_cnt_init_resp_1_days") / numSubmitted)*100)
          lcl_init_resp_1_days = REPLACE(formatnumber(lcl_init_resp_1_days,3),".000","") & "%"

       			if not IsNull(oRequests("totalDaysOpen")) then
	 	   	   		 totalOpenDays = totalOpenDays + oRequests("totalDaysOpen")
    		   	end if

       		'BEGIN: SubTotal Closed ----------------------------------------------
	 	   	   numClosed = clng(oRequests("numClosed")) 

       			if numClosed <> 0 and (oRequests("totalDaysClosed") <> 0 or oRequests("totalDaysClosed") = "") then
  	   	   			avgClosed = oRequests("totalDaysClosed") / numClosed
	  			   	   avgClosed = formatnumber(avgClosed,1)
       			else
			         	if numClosed > 0 then
 				          'Handle datediff is 0 but some have been completed, so they were completed the same day.
	    	 		   	   avgClosed = " < 1.0 "
      				   else
			      		     avgClosed = " None Completed "
   			       end if
       			end if

       			totalClosed = totalClosed + numClosed
       		'END: SubTotal Closed ------------------------------------------------

       			if not IsNull(oRequests("totalDaysClosed")) then
  	   	   			totalClosedDays = totalClosedDays + oRequests("totalDaysClosed")
    		   	end if

       			if pastDays <> "all" then
  	   	   			numPast   = oRequests("numPast")
	  			   	   totalPast = totalPast + numPast
       			end if
       end if

       'lcl_td_onclick = " onClick=""location.href='" & detaillink & "';"""
       lcl_td_onclick = " onClick=""javascript:viewDetails('" & lcl_linecnt & "');"""

    			response.write "  <tr bgcolor=" & bgcolor & " align=""center"" onMouseOver=""this.style.backgroundColor='#93bee1';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';"">" & vbcrlf
       response.write "      <td" & lcl_td_onclick & "><strong>" & replace(sTitle,"INPROGRESS","IN PROGRESS") & " </strong></td>" & vbcrlf

       if UCASE(reporttype) = "STATUSSUMMARY" then
          lcl_displayToDate = todate
          lcl_queryToDate   = dateAdd("d",1,toDate)

         'BEGIN: Calculate the individual table data cells
          lcl_allSubmitted  = oRequests("total_submitted")
          lcl_allInProgress = oRequests("total_inprogress")
          lcl_allWaiting    = oRequests("total_waiting")
          lcl_allResolved   = oRequests("total_resolved")
          lcl_allDismissed  = oRequests("total_dismissed")

          lcl_totalopen     = lcl_allSubmitted + lcl_allInProgress + lcl_allWaiting
          lcl_totalclosed   = lcl_allResolved + lcl_allDismissed
          lcl_grandtotal    = lcl_totalopen + lcl_totalclosed
         'END: Calculate the individual table data cells

         'BEGIN: Track the running totals to display in the last row.
          lcl_columntotals_submitted   = lcl_columntotals_submitted  + lcl_allSubmitted
          lcl_columntotals_inprogress  = lcl_columntotals_inprogress + lcl_allInProgress
          lcl_columntotals_waiting     = lcl_columntotals_waiting    + lcl_allWaiting
          lcl_columntotals_resolved    = lcl_columntotals_resolved   + lcl_allResolved
          lcl_columntotals_dismissed   = lcl_columntotals_dismissed  + lcl_allDismissed

          lcl_columntotals_totalopen   = lcl_columntotals_totalopen   + lcl_totalopen
          lcl_columntotals_totalclosed = lcl_columntotals_totalclosed + lcl_totalclosed
          lcl_columntotals_grandtotal  = lcl_columntotals_grandtotal  + lcl_grandtotal
         'END: Track the running totals to display in the last row.

          response.write "      <td" & lcl_td_onclick & ">"         & lcl_allSubmitted  & "</td>"   & vbcrlf
          response.write "      <td" & lcl_td_onclick & ">"         & lcl_allInProgress & "</td>"   & vbcrlf
          response.write "      <td" & lcl_td_onclick & ">"         & lcl_allWaiting    & "</td>"   & vbcrlf
          response.write "      <td" & lcl_td_onclick & "><strong>" & lcl_totalopen     & "</strong></td>" & vbcrlf
          response.write "      <td" & lcl_td_onclick & ">"         & lcl_allResolved   & "</td>"   & vbcrlf
          response.write "      <td" & lcl_td_onclick & ">"         & lcl_allDismissed  & "</td>"   & vbcrlf
          response.write "      <td" & lcl_td_onclick & "><strong>" & lcl_totalclosed   & "</strong></td>" & vbcrlf
          'response.write "      <td" & lcl_td_onclick & " style=""color:#800000; font-weight:bold;"">" & lcl_grandtotal & "</td>" & vbcrlf
          response.write "      <td" & lcl_td_onclick & " style=""color:#800000; font-weight:bold;"">" & lcl_grandtotal & "</td>" & vbcrlf
       else
    			   response.write "      <td" & lcl_td_onclick & ">" & numSubmitted         & "</td>" & vbcrlf
    			   response.write "      <td" & lcl_td_onclick & ">" & numOpen              & "</td>" & vbcrlf

       			if pastDays <> "all" then
             response.write "      <td" & lcl_td_onclick & ">" & numPast           & "</td>" & vbcrlf
       			end if

       			response.write "      <td" & lcl_td_onclick & ">" & avgOpen              & "</td>" & vbcrlf
	 		      response.write "      <td" & lcl_td_onclick & ">" & avgClosed            & "</td>" & vbcrlf

       			response.write "      <td" & lcl_td_onclick & ">" & lcl_init_resp_1_days & "</td>" & vbcrlf
	 		      response.write "      <td" & lcl_td_onclick & ">" & lcl_percent_3_days   & "</td>" & vbcrlf
       end if

    			response.write "  </tr>" & vbcrlf

     		oRequests.movenext
    loop

    oRequests.close
    set oRequests = nothing

   'Calculate the average Open/Closed
   	if totalClosed <> 0 and totalClosedDays <> 0 then
		   		avgClosed = totalClosedDays / totalClosed
  	 			avgClosed = formatnumber(avgClosed,1)
   	else
	 	  		if totalClosed > 0 then
      			'Handle datediff is 0 but some have been completed, so they were completed the same day.
     					avgClosed = " < 1.0 "
  		   else
    	 				avgClosed = ""
  				 end if
    end if

   	if totalOpen<>0 and totalOpenDays<> 0 then
	 	  		avgOpen = totalOpenDays / totalOpen
   				avgOpen = formatnumber(avgOpen,1)

      'If the avgOpen value is less than 1 day then show the following.
       if avgOpen < 1 then
          avgOpen = "< 1.0"
       end if
   	else
     		if totalOpen > 0 then
   		   	'the datediff is 0 but there are some open items, so the are from today.
     					avgOpen = "< 1.0"
  		 		else
        		'response.write("'" & avgOpen & "'")
		  	     avgOpen = ""
       end if
   	end if

   	response.write "  <tr align=""center"" style=""font-size:12px; color:#000080; font-weight:bold; background-color:#dddddd;"">" & vbcrlf
    response.write "      <td align=""center"">TOTAL</td>" & vbcrlf

    if UCASE(reporttype) = "STATUSSUMMARY" then
       response.write "      <td>"    & lcl_columntotals_submitted   & "</td>" & vbcrlf
       response.write "      <td>"    & lcl_columntotals_inprogress  & "</td>" & vbcrlf
       response.write "      <td>"    & lcl_columntotals_waiting     & "</td>" & vbcrlf
       response.write "      <td>"    & lcl_columntotals_totalopen   & "</td>" & vbcrlf
       response.write "      <td>"    & lcl_columntotals_resolved     & "</td>" & vbcrlf
       response.write "      <td>"    & lcl_columntotals_dismissed   & "</td>" & vbcrlf
       response.write "      <td>"    & lcl_columntotals_totalclosed & "</td>" & vbcrlf
       response.write "      <td>"    & lcl_columntotals_grandtotal  & "</td>" & vbcrlf
    else
       response.write "      <td>"    & totalSubmitted & "</td>" & vbcrlf
       response.write "      <td>"    & totalOpen      & "</td>" & vbcrlf

      	if pastDays <> "all" then
	 	  	   	response.write "      <td>" & totalPast      & "</td>" & vbcrlf
    	  end if

      	response.write "      <td>"    & avgOpen        & " days</td>" & vbcrlf
       response.write "      <td>"    & avgClosed      & " days</td>" & vbcrlf
    	  response.write "      <td colspan=""2"">&nbsp;</td>" & vbcrlf
    end if

   	response.write "  </tr>" & vbcrlf
    response.write "</table>" & vbcrlf

 else
    response.write "<p><strong>No records found</strong></p>" & vbcrlf
 end if

end function

'-----------------------------------------------------------------------------
function setupGroupbyQuery(p_order_by, p_where_clause, iReportType, p_pastDays, p_searchDaysType)

  sSQL = ""

 'Setup query limitation for "Display Open Over __ Days".
  if p_searchDaysType = "PAST" AND p_pastDays <> "all" then
     lcl_displayOpenOver_SQL = " AND DateDiff(d,due_date,'" & date() & "') >= " & p_pastDays
  elseif p_searchDaysType = "OPEN" AND p_pastDays <> "all" AND p_pastDays <> "0" then
     lcl_displayOpenOver_SQL = " AND (UPPER(status) <> 'RESOLVED' "
     lcl_displayOpenOver_SQL = lcl_displayOpenOver_SQL & " AND  UPPER(status) <> 'DISMISSED' "
     lcl_displayOpenOver_SQL = lcl_displayOpenOver_SQL & " AND  DateDiff(d,submit_date,'" & date() & "') > " & p_pastDays & ") "
  else
     lcl_displayOpenOver_SQL = ""
  end if

  'Are we searching on "Open" or "Past" dates
'   if pastDays <> "all" and pastDays <> "0" then
'      if searchDaysType = "PAST" then
'         sSQL = sSQL & " AND  DateDiff(d,due_date,'" & date() & "') > " & pastDays
'      else  'OPEN
'         sSQL = sSQL & " AND (status <> 'RESOLVED' "
'         sSQL = sSQL & " AND  status <> 'DISMISSED' "
'         sSQL = sSQL & " AND  DateDiff(d,submit_date,'" & date() & "') > " & pastDays & ") "
'      end if
'   end if

  sSQL = buildColumnSummaryTotals(iReportType, p_order_by, p_where_clause, lcl_displayOpenOver_SQL, p_pastDays, selectInitialResponse, selectRequestsResolved, "N")

 'Determine which column to pull in the query based on the orderby (group by field) selected.
  if UCASE(p_order_by) = "SUBMIT_DATE" then
     sSQL = sSQL & ", submitdateshort as TheDate "
  elseif UCASE(p_order_by) = "DUE_DATE" then
     sSQL = sSQL & ", due_date "
  elseif UCASE(p_order_by) = "STREETNAME" then
     sSQL = sSQL & ", streetname, "
     sSQL = sSQL & "  streetnumber "
  elseif UCASE(p_order_by) = "ACTION_FORMID" then
     sSQL = sSQL & ", action_formTitle, action_Formid "
  elseif UCASE(p_order_by) = "DEPTID" then
     sSQL = sSQL & ", deptID "
  elseif UCASE(p_order_by) = "ASSIGNED_NAME" then
     sSQL = sSQL & ", assigned_name, assignedemployeeid "
     'sSQL = sSQL & " (CASE WHEN isnull(ltrim(rtrim(assigned_name)),'') = '' THEN 0 ELSE assignedemployeeid END) AS assignedemployeeid "
  elseif UCASE(p_order_by) = "SUBMITTEDBY" then
     sSQL = sSQL & ", userlname, userfname "
  elseif UCASE(p_order_by) = "STATUS" then
     sSQL = sSQL & ", status_order, UPPER(status) AS status "
  end if

  sSQL = sSQL & " FROM egov_action_request_view "
  sSQL = sSQL & p_where_clause

 'Check for any "Display Open Over __ Days" limitation
  sSQL = sSQL & lcl_displayOpenOver_SQL

 'Build the special limitation, GROUP BY, and ORDER BYs
  if UCASE(p_order_by) = "SUBMIT_DATE" then
     if UCASE(iReportType) <> "STATUSSUMMARY" then
        sSQL = sSQL & " GROUP BY submitdateshort "
     end if

  elseif UCASE(p_order_by) = "DUE_DATE" then
     if UCASE(iReportType) <> "STATUSSUMMARY" then
        sSQL = sSQL & " GROUP BY due_date "
     end if

  elseif UCASE(p_order_by) = "STREETNAME" then
     if UCASE(iReportType) <> "STATUSSUMMARY" then
        sSQL = sSQL & " GROUP BY streetname, streetnumber "
     end if

     sSQL = sSQL & " ORDER BY streetname, streetnumber "

  elseif UCASE(p_order_by) = "ACTION_FORMID" then
     sSQL = sSQL & " GROUP BY action_formTitle, action_Formid "
     sSQL = sSQL & " ORDER BY action_formTitle "

  elseif UCASE(p_order_by) = "DEPTID" then
     if UCASE(iReportType) <> "STATUSSUMMARY" then
        sSQL = sSQL & " GROUP BY deptID "
     end if

     sSQL = sSQL & " ORDER BY deptID "

  elseif UCASE(p_order_by) = "ASSIGNED_NAME" then
     sSQL = sSQL & " AND assigned_name <> '' AND assigned_name IS NOT NULL "

    'The first part of the query retreives ALL records that HAVE a value entered into the assigned_name.
    'Since we excluded the null values in the first part of the query we now need to get the remaining records 
    '  that do NOT have an assigned_name value.  This calls for a UNION.
     sStatusTotalsForNULLs = buildColumnSummaryTotals(iReportType, p_order_by, p_where_clause, _
                                                      lcl_displayOpenOver_SQL, pastDays, _
                                                      selectInitialResponse, selectRequestsResolved, "Y")

     if UCASE(iReportType) <> "STATUSSUMMARY" then
        sSQL = sSQL & " GROUP BY assigned_name, assignedemployeeid "
     end if

     sSQL = sSQL & " UNION ALL "
     sSQL = sSQL & sStatusTotalsForNULLs
     sSQL = sSQL & ", '' AS assigned_name, 0 as assignedemployeeid "
     sSQL = sSQL & " FROM egov_action_request_view "
     sSQL = sSQL & p_where_clause
     sSQL = sSQL & lcl_displayOpenOver_SQL
     sSQL = sSQL & " AND (assigned_name = '' OR assigned_name IS NULL) "

     sSQL = sSQL & " GROUP BY assigned_name, assignedemployeeid "
     sSQL = sSQL & " ORDER BY assigned_name, assignedemployeeid "
  elseif UCASE(p_order_by) = "SUBMITTEDBY" then
     sSQL = sSQL & " GROUP BY userlname, userfname "
     sSQL = sSQL & " ORDER BY userlname, userfname "
  elseif UCASE(p_order_by) = "STATUS" then
     sSQL = sSQL & " GROUP BY status_order, status "
     sSQL = sSQL & " ORDER BY status_order, status "
  end if

  setupGroupbyQuery = sSQL

end function

'------------------------------------------------------------------------------
function checkViewAll()
  lcl_return = 0

  sSQL = "SELECT * "
  sSQL = sSQL & " FROM dbo.UsersGroupsPlus "
  sSQL = sSQL & " WHERE UserID = " & session("userid")

  set oAdmin = Server.CreateObject("ADODB.Recordset")
  oAdmin.Open sSQL, Application("DSN"), 3, 1

  if oAdmin.eof then
	    lcl_return = 0
  elseif oAdmin("GroupName") = "Administrators" then 
    	lcl_return = 1 
  end if

  oAdmin.close
  set oAdmin = nothing

  checkViewAll = lcl_return

end function

'------------------------------------------------------------------------------
function buildColumnSummaryTotals(iReportType, iOrderBy, iWhereClause, iDisplayOpenOverSQL, iPastDays, _
                                  iSelectInitialResponse, iSelectRequestsResolved, iSearchingForNULLs)
  lcl_return         = ""
  lcl_select_orderby = ""

  if UCASE(iReportType) = "STATUSSUMMARY" then
    'For the StatusSummary report the "p_order_by" value is the GROUP BY value + the ID.
     if UCASE(iOrderBy) = "SUBMIT_DATE" then
        lcl_select_orderby = " AND v2.submitdateshort = egov_action_request_view.submitdateshort "
     elseif UCASE(iOrderBy) = "DUE_DATE" then
        lcl_select_orderby = " AND v2.due_date = egov_action_request_view.due_date "
     elseif UCASE(iOrderBy) = "STREETNAME" then
        lcl_select_orderby = " AND v2.streetname = egov_action_request_view.streetname " 
        lcl_select_orderby = lcl_select_orderby & " AND v2.streetnumber = egov_action_request_view.streetnumber "
     elseif UCASE(iOrderBy) = "ACTION_FORMID" then
        lcl_select_orderby = " AND v2.action_formid = egov_action_request_view.action_formid and v2.action_formTitle = egov_action_request_view.action_formTitle "
     elseif UCASE(iOrderBy) = "DEPTID" then
        lcl_select_orderby = " AND v2.deptid = egov_action_request_view.deptid "
     elseif UCASE(iOrderBy) = "ASSIGNED_NAME" then

       'There reason for this check for NULLs is that for the ASSIGNED_NAME we have broken the query into a UNION.
       'The first part of the UNION will search for all records WITH a value in the assigned_name field.
       'The second part of the UNION will search for all records WITHOUT a value in the assigned_name field.
        if iSearchingForNULLs = "Y" then
           lcl_select_orderby = lcl_select_orderby & " AND (v2.assigned_name = '' OR v2.assigned_name IS NULL) "
        else
           lcl_select_orderby = " AND v2.assignedemployeeid = egov_action_request_view.assignedemployeeid "
           lcl_select_orderby = lcl_select_orderby & " AND v2.assigned_name <> '' AND v2.assigned_name IS NOT NULL "
        end if

        'lcl_select_orderby = " AND v2.assigned_name = egov_action_request_view.assigned_name "
        'lcl_select_orderby = lcl_select_orderby & " AND v2.assigned_userid = egov_action_request_view.assigned_userid "
     elseif UCASE(iOrderBy) = "SUBMITTEDBY" then
        lcl_select_orderby = " AND (v2.userlname = egov_action_request_view.userlname "
        lcl_select_orderby = lcl_select_orderby & " AND v2.userfname = egov_action_request_view.userfname) "
     elseif UCASE(iOrderBy) = "STATUS" then
        lcl_select_orderby = " AND UPPER(V2.status) = UPPER(egov_action_request_view.status) "
     end if

     lcl_return = "SELECT distinct "

    'Submitted
     sStatusTotals = ""
     sStatusTotals = sStatusTotals & " (SELECT count(v2.action_autoid) "
     sStatusTotals = sStatusTotals & "  FROM egov_action_request_view v2 "
     sStatusTotals = sStatusTotals &    replace(iWhereClause,"egov_action_request_view","v2")
     sStatusTotals = sStatusTotals & "  AND UPPER(v2.status) = 'SUBMITTED' "
     sStatusTotals = sStatusTotals &    iDisplayOpenOverSQL
     sStatusTotals = sStatusTotals &    lcl_select_orderby
     sStatusTotals = sStatusTotals & " ) AS total_submitted, "

    'InProgress
     sStatusTotals = sStatusTotals & " (SELECT count(v2.action_autoid) "
     sStatusTotals = sStatusTotals & "  FROM egov_action_request_view v2 "
     sStatusTotals = sStatusTotals &    replace(iWhereClause,"egov_action_request_view","v2")
     sStatusTotals = sStatusTotals & "  AND UPPER(v2.status) = 'INPROGRESS' "
     sStatusTotals = sStatusTotals &    iDisplayOpenOverSQL
     sStatusTotals = sStatusTotals &    lcl_select_orderby
     sStatusTotals = sStatusTotals & " ) AS total_inprogress, "

    'Waiting
     sStatusTotals = sStatusTotals & " (SELECT count(v2.action_autoid) "
     sStatusTotals = sStatusTotals & "  FROM egov_action_request_view v2 "
     sStatusTotals = sStatusTotals &    replace(iWhereClause,"egov_action_request_view","v2")
     sStatusTotals = sStatusTotals & "  AND UPPER(v2.status) = 'WAITING' "
     sStatusTotals = sStatusTotals &    iDisplayOpenOverSQL
     sStatusTotals = sStatusTotals &    lcl_select_orderby
     sStatusTotals = sStatusTotals & " ) AS total_waiting, "

    'Resolved
     sStatusTotals = sStatusTotals & " (SELECT count(v2.action_autoid) "
     sStatusTotals = sStatusTotals & "  FROM egov_action_request_view v2 "
     sStatusTotals = sStatusTotals &    replace(iWhereClause,"egov_action_request_view","v2")
     sStatusTotals = sStatusTotals & "  AND UPPER(v2.status) = 'RESOLVED'"
     sStatusTotals = sStatusTotals &    iDisplayOpenOverSQL
     sStatusTotals = sStatusTotals &    lcl_select_orderby
     sStatusTotals = sStatusTotals & " ) AS total_resolved, "

    'Dismissed
     sStatusTotals = sStatusTotals & " (SELECT count(v2.action_autoid) "
     sStatusTotals = sStatusTotals & "  FROM egov_action_request_view v2 "
     sStatusTotals = sStatusTotals &    replace(iWhereClause,"egov_action_request_view","v2")
     sStatusTotals = sStatusTotals & "  AND UPPER(v2.status) = 'DISMISSED' "
     sStatusTotals = sStatusTotals &    iDisplayOpenOverSQL
     sStatusTotals = sStatusTotals &    lcl_select_orderby
     sStatusTotals = sStatusTotals & " ) AS total_dismissed"

     lcl_return = lcl_return & sStatusTotals

  else
    	if iPastDays <> "all" then
        pastDate = clng(iPastDays)
     else
        pastDate = 10000
     end if

   	 lcl_return = "SELECT sum(responsetime) as totalresponsetime, sum(viewedrequests) as ttlviewedrequests, "
     lcl_return = lcl_return & " SUM(DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "'))) AS totalDays, count(action_autoid) as numSubmitted, "
     lcl_return = lcl_return & " SUM(Case WHEN (UPPER(status) <> 'RESOLVED' AND UPPER(status) <> 'DISMISSED') THEN 1 ELSE 0 END) AS numOpen,"
     'lcl_return = lcl_return & " SUM(Case WHEN (UPPER(status) <> 'RESOLVED' AND UPPER(status) <> 'DISMISSED') THEN DateDiff(d,submit_date,IsNull(complete_date,'11/3/2008 2:59:16 PM')) ELSE 0 END) AS totalDaysOpen,"
     lcl_return = lcl_return & " SUM(Case WHEN (UPPER(status) <> 'RESOLVED' AND UPPER(status) <> 'DISMISSED') THEN DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "')) ELSE 0 END) AS totalDaysOpen,"
     lcl_return = lcl_return & " SUM(Case WHEN (UPPER(status)='RESOLVED' OR UPPER(status)='DISMISSED') THEN 1 ELSE 0 END) AS numClosed,"
     'lcl_return = lcl_return & " SUM(Case WHEN (UPPER(status)='RESOLVED' OR UPPER(status)='DISMISSED') THEN DateDiff(d,submit_date,IsNull(complete_date,'11/3/2008 2:59:16 PM')) ELSE 0 END) AS totalDaysClosed,"
     lcl_return = lcl_return & " SUM(Case WHEN (UPPER(status)='RESOLVED' OR UPPER(status)='DISMISSED') THEN DateDiff(d,submit_date,IsNull(complete_date,'" & Now & "')) ELSE 0 END) AS totalDaysClosed,"
     lcl_return = lcl_return & " SUM(Case WHEN DateDiff(d,submit_date,'" & date() & "') > " & pastDate                & " AND UPPER(status) <> 'RESOLVED' AND UPPER(status) <> 'DISMISSED' THEN 1 ELSE 0 END) AS numPast,"
     lcl_return = lcl_return & " SUM(Case WHEN DateDiff(d,submit_date,firstactiondate) <= " & iSelectInitialResponse  & " AND (UPPER(status) = 'RESOLVED' OR UPPER(status) = 'DISMISSED') THEN 1 ELSE 0 END) AS total_cnt_init_resp_1_days,"
     lcl_return = lcl_return & " SUM(Case WHEN DateDiff(d,submit_date,complete_date) <= "   & iSelectRequestsResolved & " AND (UPPER(status) = 'RESOLVED' OR UPPER(status) = 'DISMISSED') THEN 1 ELSE 0 END) AS total_cnt_3_days "
  end if

  buildColumnSummaryTotals = lcl_return

end function
%>
