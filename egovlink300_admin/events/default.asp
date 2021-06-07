<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="events_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: default.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the calendar
'
' MODIFICATION HISTORY
' 1.0 ???		   	????        - INITIAL VERSION
' 1.1	10/11/06	Steve Loar  - Security, Header and nav changed
' 1.2 08/05/08 David Boyer - Added Custom Calendar
' 1.3 05/08/09 David Boyer - Added "Send to RSS" and "View Send Log" buttons
' 1.4 06/09/09	David Boyer - Added checkbox for "send to" function.  (Send to features like RSS and eventually Twitter, etc.)
' 1.5 06/15/09 David Boyer - Added search options.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim bUserHasiCalExportPermission, iSelectedDOW, iHour, iMinute, sAmPm, sMinute

'Check to see if the feature is offline
 if isFeatureOffline("calendar") = "Y" Or isFeatureOffline("custom_calendars") = "Y" then
 	response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

'Used in custom reporting
 session("RSSType") = "COMMUNITYCALENDAR"

'Check to see if this is a Custom Calendar
 lcl_calendarfeatureid    = ""
 lcl_calendarfeature      = ""
 lcl_calendarfeature_url  = ""
 lcl_calendarfeature_name = ""
 lcl_feature_rssfeeds     = "rssfeeds_events_communitycalendar"

'Allow the user to maintain the Events if any/all of the following:
'1. The user has the "edit events" permission assigned
'2. The user has a specific Custom Calendar feature assigned: [session("calendarfeature") <> ""]
 if Trim(request("cal")) <> "" then
	if not IsNumeric(Trim(request("cal"))) then
		response.redirect sLevel & "permissiondenied.asp"
	else
		lcl_calendarfeatureid = CLng(Trim(request("cal")))
		lcl_calendarfeature   = getFeatureByID(session("orgid"), lcl_calendarfeatureid)

		'if OrgHasFeature(trim(request("cal"))) AND UserHasPermission(session("userid"), trim(request("cal"))) then
		if OrgHasFeature(lcl_calendarfeature) AND UserHasPermission(session("userid"), lcl_calendarfeature) then
			lcl_calendarfeature_url  = "&cal=" & lcl_calendarfeatureid
			lcl_calendarfeature_name = " [" & getFeatureName(lcl_calendarfeature) & "]"
		else 
			response.redirect sLevel & "permissiondenied.asp"
		end if
	end if

	'if OrgHasFeature(trim(request("cal"))) AND UserHasPermission(session("userid"), trim(request("cal"))) then
	'session("calendarfeature") = trim(request("cal"))
	'lcl_calendarfeature_url    = "&cal=" & session("calendarfeature")
	'lcl_calendar_name          = " [" & GetFeatureName(session("calendarfeature")) & "]"
	'   lcl_calendarfeature     = trim(request("cal"))
	'   lcl_calendarfeature_url = "&cal=" & lcl_calendarfeature
	'   lcl_calendar_name       = " [" & GetFeatureName(lcl_calendarfeature) & "]"
	'   lcl_feature_rssfeeds    = "rssfeeds_events_" & lcl_calendarfeature
	'else
	'  	response.redirect sLevel & "permissiondenied.asp"
	'end if
 else 
	if not UserHasPermission( session("userid"), "edit events" ) then
		response.redirect sLevel & "permissiondenied.asp"
	end if

	'session("calendarfeature") = ""
 end if

'Check for org features
 lcl_orghasfeature_rssfeeds_events = orghasfeature(lcl_feature_rssfeeds)
 lcl_orghasfeature_communitylink   = orghasfeature("communitylink")
 lcl_orghasfeature_events_scroller = orghasfeature("events_scroller")

'Check for user permissions
 lcl_userhaspermission_rssfeeds_events = userhaspermission(session("userid"),lcl_feature_rssfeeds)
 lcl_userhaspermission_communitylink   = userhaspermission(session("userid"),"communitylink")
 lcl_userhaspermission_create_events   = userhaspermission(session("userid"),"create events")
 lcl_userhaspermission_events_scroller = userhaspermission(session("userid"),"events_scroller")
 bUserHasiCalExportPermission          = userhaspermission( session("userid"), "ical export" )

 dim sEvents, index, arrColors(2), truncMessage, iCount, sDesc, sLinks, bShown

'BEGIN: Check for search criteria and session variables -----------------------
 lcl_sc_fromdate = ""
 lcl_sc_todate   = ""
 lcl_sc_subject  = ""
 lcl_sc_category = ""
 lcl_sc_orderby  = "ASC"

'Clear the session variables if this is the initial screen load.
 if request("init") = "Y" then
	session("sc_fromdate") = ""
	session("sc_todate")   = ""
	session("sc_subject")  = ""
	session("sc_category") = ""
	session("sc_orderby")  = ""

	lcl_sc_fromdate = formatdatetime(date(),vbshortdate)
	lcl_sc_todate   = formatdatetime(dateadd("d",1,dateadd("m",1,date())),vbshortdate)
 else
	if request("useSessions") <> "Y" OR trim(request("useSessions")) = "" then
		session("sc_fromdate") = request("fromDate")
		session("sc_todate")   = request("toDate")
		session("sc_subject")  = request("sc_subject")
		session("sc_category") = request("sc_category")
		session("sc_orderby")  = request("sc_orderby")
	end if

	if trim(session("sc_fromdate")) <> "" then
		lcl_sc_fromdate = trim(session("sc_fromdate"))
	end if

	if trim(session("sc_todate")) <> "" then
		lcl_sc_todate = trim(session("sc_todate"))
	end if

	if trim(session("sc_subject")) <> "" then
		lcl_sc_subject = trim(session("sc_subject"))
	end if

	if trim(session("sc_category")) <> "" then
		lcl_sc_category = trim(session("sc_category"))
	end if

	if trim(session("sc_orderby")) <> "" then
		lcl_sc_orderby = trim(session("sc_orderby"))
	end if

end if

'If the date fields are blank then default the values
if lcl_sc_fromdate = "" then
	lcl_sc_fromdate = formatdatetime(date(),vbshortdate)
end if

if lcl_sc_todate = "" then
	lcl_sc_todate = formatdatetime(dateadd("d",1,dateadd("m",1,date())),vbshortdate)
end If

If request("hour") <> "" Then
	If request("hour") = "none" Then
		iHour = -1
	Else 
		iHour = clng(request("hour"))
	End If 
Else
	iHour = -1
End If 

If request("minute") <> "" Then
	If request("minute") = "none" Then
		iMinute = "-1"
		sMinute = "-1"
	Else 
		iMinute = clng(request("minute"))
		sMinute = request("minute")
	End If 
Else
	iMinute = "-1"
	sMinute = "-1"
End If 

If request("ampm") <> "" Then
	If request("ampm") = "none" Then
		sAmPm = ""
	Else 
		sAmPm = request("ampm")
	End If 
Else
	sAmPm = ""
End If 

If request("dowpick") <> "" Then
	iSelectedDOW = CLng(request("dowpick"))
Else
	iSelectedDOW = 0
End If 

if lcl_sc_orderby = "DESC" then
	lcl_selected_sc_orderby_desc = " selected=""selected"""
	lcl_selected_sc_orderby_asc  = ""
else
	lcl_selected_sc_orderby_desc = ""
	lcl_selected_sc_orderby_asc  = " selected=""selected"""
end if
'END: Check for search criteria and session variables -------------------------

'Check for a screen message
lcl_onload  = ""
lcl_success = request("success")

if lcl_success <> "" then
	lcl_msg    = setupScreenMsg(lcl_success)
	lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
end if

'Determine if there is any additional processing needed from the past update
if lcl_success = "SU" OR lcl_success = "SA" then
	if request("sendTo_RSS") <> "" then
		lcl_onload = lcl_onload & "sendToRSS('" & request("sendTo_RSS") & "');"
	end if
end if

if lcl_orghasfeature_events_scroller then
	lcl_onload        = lcl_onload & "adjustEventScroller_toDate();"
	lcl_showDaysLimit = getEventsDaysLimit(session("orgid"))
end if

%>
<html lang="en">
<head>
	<meta charset="utf-8" />
	
 	<title><%=langBSEvents%><%=lcl_calendarfeature_name%></title>

 	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="eventstyles.css" />
	<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css">

	<script src="https://code.jquery.com/jquery-1.9.1.js"></script>
  	<script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>
  	<script src="../scripts/isvaliddate.js"></script>
	<script src="../scripts/selectAll.js"></script>
	<script src="../scripts/getdates.js"></script>
	<script src="../scripts/ajaxLib.js"></script>
	<script src="../scripts/modules.js"></script>
	<script src="../scripts/formvalidation_msgdisplay.js"></script>

	<script>
	<!--
		$(document).ready(function(){
			//$('#eventsScrollerDays').val('14');

			$('#searchButton').click(function() {
				$('#eventMaint').prop('action','default.asp?cal=<%=lcl_calendarfeatureid%>');
				validateFields();
			});

			$('#deleteButton').click(function() {
				$('#eventMaint').prop('action','deleteevents.asp');
				$('#eventMaint').submit();
			});

			$('#fromDateCalPop').click(function() {
				clearMsg('fromDateCalPop');
				//doCalendar('fromDate');
			});

			$('#toDateCalPop').click(function() {
				clearMsg('toDateCalPop'); 
				//doCalendar('toDate');
			});

			$('#chkSelectAll').click(function() {
				var lcl_selectall = true;

				if(! $('#chkSelectAll').prop('checked')) {
					lcl_selectall = false;
				}

				$('input[id^="del_"]').each(function() {
					$(this).prop('checked',lcl_selectall);
				});
			});

		<% if lcl_orghasfeature_events_scroller AND lcl_userhaspermission_events_scroller then %>
			$('#saveScrollerDaysButton').click(function() {
				var lcl_showDaysLimit = '<%=lcl_showDaysLimit%>';

				if($('#eventsScrollerDays').val() != '') {
					if(isNaN($('#eventsScrollerDays').val())) {
						$('#eventsScrollerDays').focus();
						inlineMsg(document.getElementById("eventsScrollerDays").id,'<strong>Invalid Value: </strong> Days must be numeric',10,'eventsScrollerDays');
						return false;            
					} else {
						clearMsg('eventsScrollerDays');
						lcl_showDaysLimit = parseInt($('#eventsScrollerDays').val());

						$.post('update_eventScrollerDateLimit.asp', {
							orgid: '<%=session("orgid")%>',
							eventScrollerDateLimit: lcl_showDaysLimit,
							isAjax: 'Y'
						}, function(result) {
							displayScreenMsg(result);
							adjustEventScroller_toDate();
						});
					}
				} else {
					lcl_showDaysLimit = '14';
					$('#eventsScrollerDays').val(lcl_showDaysLimit);
					inlineMsg(document.getElementById("eventsScrollerDays").id,'<strong>Required Field Missing: </strong> A numeric value must be entered for the days limit.  It has been defaulted to <span style=\"color:#800000\">(' + lcl_showDaysLimit + ')</span>',10,'eventsScrollerDays');
					return false;
				}
			});
		<% end if %>
		});

		<% if lcl_orghasfeature_events_scroller then %>
			function adjustEventScroller_toDate() {
				//Determine the new "to" date
				var lcl_showDaysLimit = '';
				var lcl_newDate       = new Date();
				var month             = new Array();
				var lcl_newToDate     = '';
				var lcl_newMonth      = '';
				var lcl_newDay        = '';
				var lcl_newYear       = '';

				month[0]  = '1';
				month[1]  = '2';
				month[2]  = '3';
				month[3]  = '4';
				month[4]  = '5';
				month[5]  = '6';
				month[6]  = '7';
				month[7]  = '8';
				month[8]  = '9';
				month[9]  = '10';
				month[10] = '11';
				month[11] = '12';

				if($('#eventsScrollerDays').val() != '') {
					lcl_showDaysLimit = parseInt($('#eventsScrollerDays').val());
				} else {
					lcl_showDaysLimit = Number(14);
				}

				lcl_newDate.setDate(lcl_newDate.getDate() + lcl_showDaysLimit);
				lcl_newMonth  = month[lcl_newDate.getMonth()];
				lcl_newDay    = lcl_newDate.getDate();
				lcl_newYear   = lcl_newDate.getFullYear();

				lcl_newToDate = lcl_newMonth + '/' + lcl_newDay + '/' + lcl_newYear;
				$('#eventsScroller_toDate').html(lcl_newToDate);
			}
		<% end if %>


		//function doCalendar( sField ) 
		//{
		//	var w = (screen.width - 350)/2;
		//	var h = (screen.height - 350)/2;
		//	eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=eventMaint", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		//}
		var datesAreValid = function() {
			var okToPost = true;
			// check from date
			if ($("#fromDate").val() != "") {
				if (! isValidDate($("#fromDate").val()) ) {
					inlineMsg("fromDate","<strong>Invalid Value: </strong>The From date should be a valid date in the format of MM/DD/YYYY.");
					okToPost = false;
				}
			}
			// check to date
			if ($("#toDate").val() != "") {
				if (! isValidDate($("#toDate").val()) ) {
					inlineMsg("toDate","<strong>Invalid Value: </strong>The To date should be a valid date in the format of MM/DD/YYYY.");
					okToPost = false;
				}
			}
			
			return okToPost;
			
		};

		var validateFields = function() { 
			
			if ( datesAreValid() ) {
				$('#eventMaint').submit();
				return true;
			} else {
				return false;
			}

		};

		function viewRSSLog(pID) 
		{
			var lcl_width  = 900;
			var lcl_height = 400;
			var lcl_left   = (screen.availWidth/2) - (lcl_width/2);
			var lcl_top    = (screen.availHeight/2) - (lcl_height/2);
			var popupWin = window.open("../customreports/customreports.asp?CR=RSSLOG&id=" + pID, "_blank","width=" + lcl_width + ",height=" + lcl_height + ",left=" + lcl_left + ",top=" + lcl_top + ",scrollbars=1");
		}

		function sendToRSS(pID) 
		{
			var sParameter = 'id=' + encodeURIComponent(pID);
			sParameter    += '&isAjax=Y';

			doAjax('events_sendToRSS.asp', sParameter, 'displayScreenMsg', 'post', '0');
		}

		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
			{
				document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			document.getElementById("screenMsg").innerHTML = "";
		}
		
		// these function set up the date pickers
		$(function() {
			$( "#toDate" ).datepicker({
				showOn: "button",
				buttonImage: "../images/calendar.gif",
				buttonImageOnly: true,
				changeMonth: true,
				changeYear: true
			});
		});

		$(function() {
			$( "#fromDate" ).datepicker({
				showOn: "button",
				buttonImage: "../images/calendar.gif",
				buttonImageOnly: true,
				changeMonth: true,
				changeYear: true
			});
		});
		
	//-->
	</script>
</head>
<body onload="<%=lcl_onload%>">
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<form name=""eventMaint"" id=""eventMaint"" method=""post"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""cal"" id=""cal"" value=""" & lcl_calendarfeatureid & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""sessiontodate"" value=""" & session("sc_todate") & """ />" & vbcrlf
  response.write "<table border=""0"" cellpadding=""2"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td><font size=""+1""><strong>Events" & lcl_calendarfeature_name & "</strong></font></td>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

 'BEGIN: Search Criteria ------------------------------------------------------
  response.write "  <tr>" & vbcrlf
  response.write "  	<td colspan=""2"">" & vbcrlf
  response.write "          <fieldset class=""fieldset"">" & vbcrlf
  response.write "            <legend>Search Options&nbsp;</legend>" & vbcrlf
  response.write "            <table id=""searchCriteriaTable"" border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>From:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <input type=""text"" name=""fromDate"" id=""fromDate"" size=""10"" maxlength=""10"" value=""" & lcl_sc_fromdate & """ onchange=""clearMsg('fromDate')"" />&nbsp;" 
  'response.write "                      &nbsp;<img src=""../images/calendar.gif"" border=""0"" name=""fromDateCalPop"" id=""fromDateCalPop"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      To:&nbsp;" & vbcrlf
  response.write "                      <input type=""text"" name=""toDate"" id=""toDate"" size=""10"" maxlength=""10"" value=""" & lcl_sc_todate & """ onchange=""clearMsg('toDate')"" />&nbsp;"
  'response.write "                      &nbsp;<img src=""../images/calendar.gif"" border=""0"" name=""toDateCalPop"" id=""toDateCalPop"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
                                        DrawDateChoices "Date", ""
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Subject:</td>" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf
  response.write "                      <input type=""text"" name=""sc_subject"" id=""sc_subject"" size=""41"" maxlength=""10"" value=""" & lcl_sc_subject & """ />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <label for=""sc_category"">Category:</label>&nbsp;" & vbcrlf
  response.write "                      <select name=""sc_category"" id=""sc_category"">" & vbcrlf

                                          displayEventCategoryOptions lcl_sc_category, _
                                                                      lcl_calendarfeature

  response.write "                      </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Start Time:</td>" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf

                                        displayTimePicks iHour, _
                                                         iMinute, _
                                                         sAmPm

  response.write "                  </td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <label for=""dowpick"">Day of Week:</label> " & vbcrlf
                                        displayDayOfWeekPicks iSelectedDOW
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Order By:</td>" & vbcrlf
  response.write "                  <td colspan=""3"">" & vbcrlf
  response.write "                      <select name=""sc_orderby"" id=""sc_orderby"">" & vbcrlf
  response.write "                        <option value=""ASC"" " & lcl_selected_sc_orderby_asc & ">Date (oldest to newest)</option>" & vbcrlf
  response.write "                        <option value=""DESC"" " & lcl_selected_sc_orderby_desc & ">Date (newest to oldest)</option>" & vbcrlf
  response.write "                      </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td colspan=""4"">" & vbcrlf
  response.write "                      <input type=""button"" name=""searchButton"" id=""searchButton"" value=""Search"" class=""button"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "            </table>" & vbcrlf
  response.write "          </fieldset>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
 'END: Search Criteria --------------------------------------------------------

  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"" valign=""top"">" & vbcrlf

 'BEGIN: Events Scroller - Days Limit Option ----------------------------------
  if lcl_orghasfeature_events_scroller then
     if lcl_userhaspermission_events_scroller then
        lcl_display_eventsScrollerDays      = "text"
        lcl_display_eventsScrollerDays_text = " days"
        lcl_display_saveScrollerDaysButton  = true
     else
        lcl_display_eventsScrollerDays      = "hidden"
        lcl_display_eventsScrollerDays_text = ""
        lcl_display_saveScrollerDaysButton  = false
     end if

     response.write "<div id=""eventsScrollerDiv"" align=""right"">" & vbcrlf
     response.write "Show all events in scroller between <span id=""eventsScroller_fromDate"">" & date() & "</span> and <span id=""eventsScroller_toDate""></span>.<br />" & vbcrlf
     response.write "<input type=""" & lcl_display_eventsScrollerDays & """ name=""eventsScrollerDays"" id=""eventsScrollerDays"" value=""" & lcl_showDaysLimit & """ size=""3"" maxlength=""5"" onchange=""clearMsg('eventsScrollerDays');"" />" & lcl_display_eventsScrollerDays_text & vbcrlf

     if lcl_display_saveScrollerDaysButton then
        response.write "<input type=""button"" name=""saveScrollerDaysButton"" id=""saveScrollerDaysButton"" value=""Save Days"" class=""button"" />" & vbcrlf
     end if

     response.write "</div>" & vbcrlf
  end if
 'END: Events Scroller - Days Limit Option ------------------------------------

 'BEGIN: Display Button Row ---------------------------------------------------
  lcl_eventCategoriesButton = "&nbsp;"

  if lcl_calendarfeatureid <> "" then
     lcl_eventCategoriesButton = "<input type=""button"" value=""Event Categories"" class=""button"" onclick=""location.href='eventcategories.asp?cal=" & lcl_calendarfeatureid & "';"" />"
  end if

  response.write "<table id=""buttondisplay"" border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf

                            displayButtons lcl_calendarfeature, _
                                           lcl_calendarfeatureid, _
                                           bUserHasiCalExportPermission, _
                                           lcl_userhaspermission_create_events

  response.write "      </td>" & vbcrlf
  response.write "      <td align=""right"">" & lcl_eventCategoriesButton & "</td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
 'END: Display Button Row -----------------------------------------------------

 'BEGIN: Results List ---------------------------------------------------------
  displayEventsList session("orgid"), _
                    lcl_calendarfeatureid, _
                    lcl_calendarfeature, _
                    lcl_orghasfeature_rssfeeds_events, _
                    lcl_userhaspermission_rssfeeds_events, _
                    lcl_sc_fromdate, _
                    lcl_sc_todate, _
                    lcl_sc_subject, _
                    lcl_sc_category, _
                    lcl_sc_orderby, _
                    iSelectedDOW, _
                    iHour, _
                    sMinute, _
                    sAmPm
 'END: Results List -----------------------------------------------------------


  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</form>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
	<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayEventsList(ByVal iOrgID, _
                      ByVal iCalendarFeatureID, _
                      ByVal iCalendarFeature, _
                      ByVal iOrgHasFeature_rssFeedsEvents, _
                      ByVal iUserHasPermission_rssFeedsEvents, _
                      ByVal iSCFromDate, _
                      ByVal iSCToDate, _
                      ByVal iSCSubject, _
                      ByVal iSCCategory, _
                      ByVal iSCOrderBy, _
                      ByVal iSelectedDOW, _
                      ByVal iHour, _
                      ByVal iMinute, _
                      ByVal sAmPm)

	dim sOrgID, sCalendarFeatureID, sCalendarFeature, sCalendarFeatureSQL
    dim sOrgHasFeature_rssFeedsEvents, sUserHasPermission_rssFeedsEvents
	dim sSCFromDate, sSCToDate, sSCSubject, sSCCategory, sSCOrderBy, iRowCount
    dim sClass, oGetEvents, sSql, sDOWFilter, sTimeFilter

	sOrgID                            = 0
	sCalendarFeatureID                = ""
	sCalendarFeature                  = ""
	sCalendarFeatureSQL               = ""
	sOrgHasFeature_rssFeedsEvents     = false
	sUserHasPermission_rssFeedsEvents = false
	sSCFromDate                       = ""
	sSCToDate                         = ""
	sSCSubject                        = ""
	sSCCategory                       = ""
	sSCOrderBy                        = ""
	arrColors(0)                      = "ffffff"
	arrColors(1)                      = "eeeeee"
	index                             = 0
	iCount                            = 1
	iRowCount                         = 0

	If iOrgID <> "" Then 
		sOrgID = CLng(iOrgID)
	End If 

	If iCalendarFeatureID <> "" Then 
		sCalendarFeatureID = CLng(iCalendarFeatureID)
	End If 

	If iCalendarFeature <> "" Then 
		If Not containsApostrophe( iCalendarFeature ) Then 
			sCalendarFeature = iCalendarFeature
		End If 
	End If 

	If iOrgHasFeature_rssFeedsEvents <> "" Then 
		sOrgHasFeature_rssFeedsEvents = iOrgHasFeature_rssFeedsEvents
	End If 

	If iUserHasPermission_rssFeedsEvents <> "" Then 
		sUserHasPermission_rssFeedsEvents = iUserHasPermission_rssFeedsEvents
	End If 

	If iSCFromDate <> "" Then 
		If Not containsApostrophe(iSCFromDate) Then 
			sSCFromDate = iSCFromDate
			sSCFromDate = dbsafe(sSCFromDate)
			sSCFromDate = "'" & sSCFromDate & "'"
		End If 
	End If 

	If iSCToDate <> "" Then 
		If Not containsApostrophe(iSCToDate) Then 
			sSCToDate = iSCToDate
			sSCToDate = dbsafe(sSCToDate)
			sSCToDate = "'" & sSCToDate & "'"
		End If 
	End If 

	If iSCSubject <> "" Then 
		sSCSubject = iSCSubject
		sSCSubject = UCase(sSCSubject)
		sSCSubject = dbsafe(sSCSubject)
		sSCSubject = "'%" & sSCSubject & "%'"
	End If 

	If iSCCategory <> "" Then 
		If Not containsApostrophe(iSCCategory) And iSCCategory <> "ALL" Then 
			sSCCategory = clng(iSCCategory)
		End If 
	End If 

	If iSCOrderBy <> "" Then 
		If Not containsApostrophe(iSCOrderBy) Then 
			sSCOrderBy = UCase(iSCOrderBy)
			sSCOrderBy = dbsafe(sSCOrderBy)
		End If 
	End If 

	If CLng(iSelectedDOW) > CLng(0) Then 
		sDOWFilter = " AND DATEPART(dw, e.eventdate) = " & iSelectedDOW
	Else
		sDOWFilter = ""
	End If 

	If clng(iHour) <> clng(-1) And clng(iMinute) <> clng(-1) And sAmPm <> "" Then
		sTimeFilter = " AND LTRIM(RIGHT(CONVERT(VARCHAR(20), e.EventDate, 100), 7)) = '" & iHour & ":" & iMinute & sAmPm & "' "
	Else
		sTimeFilter = ""
	End If 

	'Setup the query depending on if this is a custom calendar or not.
	If sCalendarFeature <> "" Then 
		sCalendarFeatureSQL = " AND UPPER(e.calendarfeature) = '" & UCase(sCalendarFeature) & "' "
	Else 
		sCalendarFeatureSQL = " AND (e.calendarfeature = '' OR e.calendarfeature IS NULL) "
	End If 

	response.write "  <table width=""100%"" cellpadding=""5"" cellspacing=""0"" border=""0"" class=""tablelist"" id=""eventslist"">" & vbcrlf
	response.write "    <tr>" & vbcrlf
	'response.write "        <th width=""20""><input class=""listCheck"" type=""checkbox"" name=""chkSelectAll"" onClick=""selectAll('DelEvent', this.checked)""></th>" & vbcrlf
	response.write "        <th width=""20""><input type=""checkbox"" id=""chkSelectAll"" name=""chkSelectAll"" class=""listCheck"" /></th>" & vbcrlf
	response.write "        <th align=""left"" width=""10%"">Category</th>" & vbcrlf
	response.write "        <th align=""left"" width=""20%"">Subject</th>" & vbcrlf
	response.write "        <th align=""left"" width=""12%"">Date</th>" & vbcrlf
	response.write "        <th align=""left"">Message</th>" & vbcrlf
	response.write "        <th>&nbsp;</th>" & vbcrlf

	If sOrgHasFeature_rssFeedsEvents And sUserHasPermission_rssFeedsEvents Then 
		response.write "        <th>Send to<br />RSS</th>" & vbcrlf
		response.write "        <th>RSS<br />Send<br />Log</th>" & vbcrlf
	End If 

	response.write "    </tr>" & vbcrlf

	'Retrieve all of the events
	sSQL = "SELECT e.EventID, "
 sSQL = sSQL & " e.EventDate, "
 sSQL = sSQL & " t.TZAbbreviation, "
 sSQL = sSQL & " e.Subject, "
 sSQL = sSQL & " e.Message, "
	sSQL = sSQL & " u.FirstName, "
 sSQL = sSQL & " u.LastName, "
 sSQL = sSQL & " u.Email, "
 sSQL = sSQL & " e.ModifiedDate, "
 sSQL = sSQL & " c.Color, "
	sSQL = sSQL & " c.CategoryName, "
 sSQL = sSQL & " e.calendarfeature "
	sSQL = sSQL & " FROM Events as e "
	sSQL = sSQL &      " LEFT JOIN TimeZones as t ON t.TimeZoneID = e.EventTimeZoneID "
	sSQL = sSQL &      " LEFT JOIN EventCategories as c ON e.CategoryID = c.CategoryID "
	sSQL = sSQL &      " LEFT JOIN Users as u ON u.UserID = e.ModifierUserID "
	sSQL = sSQL & " WHERE e.OrgID = " & sOrgID
	sSQL = sSQL & sCalendarFeatureSQL & sDOWFilter & sTimeFilter

	'Build the search option(s) into the query
	'From/To Dates
	If sSCFromDate <> "" And sSCToDate <> "" Then 
		sSQL = sSQL & " AND e.eventdate BETWEEN " & sSCFromDate & " AND " & sSCToDate
	Else 
		If sSCFromDate <> "" Then 
			sSQL = sSQL & " AND e.eventdate >= " & sSCFromDate
		ElseIf sSCToDate <> "" Then 
			sSQL = sSQL & " AND e.eventdate <= " & sSCToDate
		End If 
	End If 

	'Subject
	If sSCSubject <> "" Then 
		sSQL = sSQL & " AND UPPER(e.Subject) LIKE (" & sSCSubject & ") "
	End If 

	'Category
	If sSCCategory <> "" And sSCCategory <> "ALL" Then 
		sSQL = sSQL & " AND e.categoryid = " & sSCCategory
	End If 

	'Order By
	If sSCOrderBy <> "" Then 
		sSQL = sSQL & " ORDER BY e.eventdate " & sSCOrderBy
	End If 

	'response.write sSQL & "<br /><br />"

	Set oGetEvents = Server.CreateObject("ADODB.Recordset")
	oGetEvents.Open sSQL, Application("DSN"), 3, 1

	Do While Not oGetEvents.EOF
		iRowCount      = iRowCount + 1
		lDateDiff      = DateDiff("d", Now(), oGetEvents("EventDate"))
		sClass         = ""
		lcl_td_onclick = ""
		lcl_td_title   = ""
		lcl_expiredmsg = "&nbsp;"
		lcl_eventdate  = "&nbsp;"

		'TEMPORARY BUG FIX FOR WARRINGTON!!!
		truncMessage = oGetEvents("Message")
		truncMessage = RemoveAnchorTags(truncMessage)

		If Len(truncMessage) > 250 Then 
			truncMessage = Left(truncMessage,248) & "..."
		End If 

		If iRowCount Mod 2 = 0 Then 
			sClass = " class=""altrow"""
		End If 

		lcl_td_onclick         = " onclick=""location.href='updateevent.asp?cal=" & sCalendarFeatureID & "&id=" & oGetEvents("EventID") & "';"""
		lcl_td_title           = " title=""click to edit"""

		If lDateDiff < 0 Then 
			lcl_expiredmsg = "<font color=""#ff0000"" size=""1""><strong>Expired</strong></font>"
		End If 

		'If oGetEvents("eventdate") <> "" Then 
		'	lcl_eventdate = MyFormatDateTime(oGetEvents("eventdate"),"<br />")
		'End If 

		If oGetEvents("eventdate") <> "" Then 
			lcl_eventdate = eventFormatDateTime(oGetEvents("eventdate"))
		End If 

		response.write "    <tr id=""" & iRowCount & """" & sClass & " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" valign=""top"">" & vbcrlf
		response.write "        <td><input type=""checkbox"" name=""del_" & oGetEvents("eventid") & """ id=""del_" & oGetEvents("eventid") & """ class=""listcheck"" /></td>" & vbcrlf
		response.write "        <td" & lcl_td_title & lcl_td_onclick & "><font color=""" & oGetEvents("Color") & """>" & oGetEvents("CategoryName") & "</font></td>" & vbcrlf
		response.write "        <td" & lcl_td_title & lcl_td_onclick & ">" & oGetEvents("Subject") & "</td>" & vbcrlf
		response.write "        <td" & lcl_td_title & lcl_td_onclick & ">" & lcl_eventdate   & "</td>" & vbcrlf
		response.write "        <td" & lcl_td_title & lcl_td_onclick & ">" & truncMessage    & "</td>" & vbcrlf
		response.write "        <td" & lcl_td_title & lcl_td_onclick & " align=""right"">" & lcl_expiredmsg & "</td>" & vbcrlf

		If sOrgHasFeature_rssFeedsEvents And sUserHasPermission_rssFeedsEvents Then 
			lcl_viewRSSLogButton = "&nbsp;"

			If checkRSSLogExists(sOrgID,oGetEvents("eventid"),"COMMUNITYCALENDAR") Then 
				lcl_viewRSSLogButton = "<input type=""button"" name=""viewRSSLog" & iRowCount & """ id=""viewRSSLog" & iRowCount & """ value=""View"" class=""button"" onclick=""viewRSSLog('" & oGetEvents("eventid") & "');"" />"
			End If 

			response.write "        <td align=""center"">" & vbcrlf
			response.write "            <input type=""button"" name=""sendToRSS" & iRowCount & """ id=""sendToRSS" & iRowCount & """ value=""Send"" class=""button"" onclick=""sendToRSS('" & oGetEvents("eventid") & "');"" />" & vbcrlf
			response.write "        </td>" & vbcrlf
			response.write "        <td align=""center"">" & lcl_viewRSSLogButton & "</td>" & vbcrlf
		End If 

		response.write "    </tr>" & vbcrlf

		index  = 1 - index 'flip the index
		iCount = iCount + 1

		oGetEvents.movenext
	Loop 

	If iRowCount = 0 Then 
		response.write "    <tr><td colspan=""5"">No events have been created.</td></tr>" & vbcrlf
	End If 

	oGetEvents.Close
	Set oGetEvents = Nothing 

	response.write "  </table>" & vbcrlf

	response.write "<div align=""right"" id=""totalevents"">Total Events [" & iRowCount & "]</div>" & vbcrlf

End Sub 


'------------------------------------------------------------------------------
Sub displayButtons(ByVal iCalendarFeature, _
                   ByVal iCalendarFeatureID, _
                   ByVal bUserHasiCalExportPermission, _
                   ByVal iUserHasPermission_createEvents)

    dim sCalendarFeature, sCalendarFeatureID, lcl_show_new_button

	sCalendarFeature    = ""
	sCalendarFeatureID  = ""
	lcl_show_new_button = False

	if iCalendarFeature <> "" then
		if not containsApostrophe(iCalendarFeature) then
			sCalendarFeature = iCalendarFeature
		end if
	end if

	if iCalendarFeatureID <> "" then
		if not containsApostrophe(iCalendarFeatureID) then
			sCalendarFeatureID = CLng(iCalendarFeatureID)
		end if
	end if

	if sCalendarFeature <> "" then
		lcl_show_new_button = True
	else
		if iUserHasPermission_createEvents then
			lcl_show_new_button = True
		end if
	end if

	response.write "<div id=""buttonRow"">" & vbcrlf

	if lcl_show_new_button then
		response.write vbcrlf & "<input type=""button"" value=""New Event"" class=""button"" onClick=""location.href='newevent.asp?cal=" & sCalendarFeatureID & "'"" /> &nbsp; "
	end if

	response.write vbcrlf & "  <input type=""button"" name=""deleteButton"" id=""deleteButton"" value=""Delete"" class=""button"" />" 

	if bUserHasiCalExportPermission then
		response.write "&nbsp; <input type=""button"" value=""iCal Export"" class=""button"" onClick=""location.href='ical_3month_export.asp'"" />"
	end if

	response.write "</div>" & vbcrlf

End Sub 


'------------------------------------------------------------------------------
Sub displayEventCategoryOptions( ByVal iSelectedValue, ByVal sCalendar )
	Dim sSql, oSCCat, lcl_selected_nocategory, lcl_selected_all, lcl_selected

	if iSelectedValue = "ALL" then
		lcl_selected_all = " selected=""selected"""
	else
		lcl_selected_all = ""
	end if

	response.write "<option value=""ALL""" & lcl_selected_all & ">[All Categories]</option>" & vbcrlf

	if iSelectedValue = "0" then
		lcl_selected_nocategory = " selected=""selected"""
	else
		lcl_selected_nocategory = ""
	end if

	response.write vbcrlf & "<option value=""0""" & lcl_selected_nocategory & ">[No Category Assigned]</option>"

	sSql = "SELECT DISTINCT c.CategoryName, c.categoryid  "
	sSql = sSql & " FROM EventCategories c, Events e "
	sSql = sSql & " WHERE c.categoryid = e.categoryid "
	sSql = sSql & " AND c.orgid = e.orgid "
	sSql = sSql & " AND c.orgid = " & session("orgid")

	if sCalendar <> "" then
		sSql = sSql & " AND UPPER(e.calendarfeature) = '" & UCASE(sCalendar) & "' "
	else
		sSql = sSql & " AND (e.calendarfeature IS NULL OR e.calendarfeature = '') "
	end if

	sSql = sSql & " ORDER BY c.CategoryName"

	'response.write vbcrlf & " " & sSql & "<br />"

	set oSCCat = Server.CreateObject("ADODB.Recordset")
	oSCCat.Open sSql, Application("DSN"), 0, 1

	If Not oSCCat.EOF Then 
		Do While Not oSCCat.EOF
			If CStr(iSelectedValue) = CStr(oSCCat("categoryid")) Then 
				lcl_selected = " selected=""selected"""
			Else 
				lcl_selected = ""
			End If 

			response.write vbcrlf & "<option value=""" & oSCCat("categoryid") & """" & lcl_selected & ">" & oSCCat("categoryname") & "</option>"
			oSCCat.MoveNext
		Loop 
	End If 

	oSCCat.Close
	Set oSCCat = Nothing 

End Sub

'------------------------------------------------------------------------------
Function getEventsDaysLimit( ByVal iOrgID )
	dim lcl_return, sOrgID, sSql, oGetOrgDateLimit

	lcl_return = ""
	sOrgID     = 0

	If iOrgID <> "" Then 
		sOrgID = CLng(iOrgID)
	End If 

	If sOrgID > 0 Then 
		sSql = "SELECT eventScrollerDateLimit "
		sSql = sSql & " FROM organizations "
		sSql = sSql & " WHERE orgid = " & sOrgID

		Set oGetOrgDateLimit = Server.CreateObject("ADODB.Recordset")
		oGetOrgDateLimit.Open sSql, Application("DSN"), 3, 1

		If Not oGetOrgDateLimit.EOF Then 
			If oGetOrgDateLimit("eventScrollerDateLimit") <> "" Then 
				lcl_return = oGetOrgDateLimit("eventScrollerDateLimit")
			End If 
		End If 

		oGetOrgDateLimit.close
		Set oGetOrgDateLimit = Nothing 

	End If 

	If lcl_return = "" Then 
		lcl_return = 14
	End If 

	getEventsDaysLimit = lcl_return

End Function 

'------------------------------------------------------------------------------
Sub displayDayOfWeekPicks( ByVal iSelectedDOW )
	Dim sundaySelected, mondaySelected, tuesdaySelected, wednesdaySelected
	Dim thursdaySelected, fridaySelected, saturdaySelected, anySelected

	sundaySelected = ""
	mondaySelected = ""
	tuesdaySelected = ""
	wednesdaySelected = ""
	thursdaySelected = ""
	fridaySelected = ""
	saturdaySelected = ""

	Select Case iSelectedDOW
		Case 0
			anySelected = " selected=""selected"""
		Case 1
			sundaySelected = " selected=""selected"""
		Case 2
			mondaySelected = " selected=""selected"""
		Case 3
			tuesdaySelected = " selected=""selected"""
		Case 4
			wednesdaySelected = " selected=""selected"""
		Case 5
			thursdaySelected = " selected=""selected"""
		Case 6
			fridaySelected = " selected=""selected"""
		Case 7
			saturdaySelected = " selected=""selected"""
	End Select
	
	response.write "<select id=""dowpick"" name=""dowpick"">"
	response.write "  <option value=""0""" & anySelected       & ">Any Day</option>" & vbcrlf
	response.write "  <option value=""1""" & sundaySelected    & ">Sunday</option>" & vbcrlf
	response.write "  <option value=""2""" & mondaySelected    & ">Monday</option>" & vbcrlf
	response.write "  <option value=""3""" & tuesdaySelected   & ">Tuesday</option>" & vbcrlf
	response.write "  <option value=""4""" & wednesdaySelected & ">Wednesday</option>" & vbcrlf
	response.write "  <option value=""5""" & thursdaySelected  & ">Thursday</option>" & vbcrlf
	response.write "  <option value=""6""" & fridaySelected    & ">Friday</option>" & vbcrlf
	response.write "  <option value=""7""" & saturdaySelected  & ">Saturday</option>" & vbcrlf
	response.write "</select>"

End Sub 


'------------------------------------------------------------------------------
Sub displayTimePicks(ByVal iHour, _
                     ByVal iMinute, _
                     ByVal sAmPm)

    response.write "<select name=""hour"" id=""hour"" class=""time"">" & vbcrlf
	response.write "  <option value=""none""></option>" & vbcrlf
                      'buildOption "HOUR", 0,  iHour
                      buildOption "HOUR", 1,  iHour
                      buildOption "HOUR", 2,  iHour
                      buildOption "HOUR", 3,  iHour
                      buildOption "HOUR", 4,  iHour
                      buildOption "HOUR", 5,  iHour
                      buildOption "HOUR", 6,  iHour
                      buildOption "HOUR", 7,  iHour
                      buildOption "HOUR", 8,  iHour
                      buildOption "HOUR", 9,  iHour
                      buildOption "HOUR", 10, iHour
                      buildOption "HOUR", 11, iHour
                      buildOption "HOUR", 12, iHour
	response.write "</select>" & vbcrlf
	response.write "&nbsp;:" & vbcrlf
	response.write "<select name=""minute"" id=""minute"" class=""time"">" & vbcrlf
	response.write "  <option value=""none""></option>" & vbcrlf

                      for i = 0 to 59 step 5
                         lcl_displayMinute = i

                         if i < 10 then
                            lcl_displayMinute = "0" & i
                         end if

                         buildOption "MINUTE", lcl_displayMinute, iMinute
                      next

                      'buildOption "MINUTE", "00", iMinute
                      'buildOption "MINUTE", "15", iMinute
                      'buildOption "MINUTE", "30", iMinute
                      'buildOption "MINUTE", "45", iMinute
	response.write "</select>" & vbcrlf
	response.write "<select id=""ampm"" name=""ampm"" class=""time"">" & vbcrlf
	response.write "  <option value=""none""></option>" & vbcrlf
                      buildOption "AMPM", "AM", sAmPm
                      buildOption "AMPM", "PM", sAmPm
	response.write "</select>" & vbcrlf

End Sub 


'------------------------------------------------------------------------------
'function RemoveAnchorTags(sString )
'	Dim sNewString, iStart, iEnd

'	sNewString = sString

'	if clng(len(sNewString)) > clng(2) then
'    if instr(sNewString,"<a") > 0 then
'       do while instr(sNewString,"<a") > 0
'          lcl_string_length = len(sNewString)

         'Remove all anchor tags
'          lcl_string_length   = len(sNewString)
'          lcl_anchortag_start = clng(instr(sNewString,"<a") - 1)
'          lcl_anchortag_close = clng(instr(sNewString,"</a>"))
'          lcl_left_string     = ""
'          lcl_right_string    = ""

         'Break the string up to the "left" side, part BEFORE the start of the opening anchor tag
         'and the "right" side, rest of string STARTING at the opening anchor tag
'          lcl_left_string  = left(sNewString,lcl_anchortag_start)

'          if  clng(instr(lcl_anchortag_start+1,sNewString,">")) > lcl_anchortag_start _
'          AND clng(instr(lcl_anchortag_start+1,sNewString,">")) < lcl_anchortag_close then
'              lcl_anchortag_end = clng(instr(lcl_anchortag_start+1,sNewString,">") + 1)
'              lcl_right_string  = mid(sNewString,lcl_anchortag_end)
'              lcl_right_string  = replace(lcl_right_string,"</a>","",1)
'          else
'              lcl_left_string = replace(lcl_left_string,"</a>","",1)
'          end if

'          sNewString = lcl_left_string & lcl_right_string
'       loop
'    end if
' end if

' RemoveAnchorTags = sNewString

'end function
%>
