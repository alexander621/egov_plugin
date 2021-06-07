<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="events_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: newevent.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates new calendar events
'
' MODIFICATION HISTORY
' 1.0   ???			 ???? - INITIAL VERSION
' 1.1	10/12/2006	Steve Loar - Security, Header and nav changed
' 1.2	04/19/2007	Steve Loar - Changed so default time is midnight.
' 1.3	12/07/2007	Steve Loar - Closed some objects that were not being closed
' 1.4	01/24/2008	Steve Loar - Put check on mesage length to prevent crashes when length is greater than 1000.
' 1.5	08/06/2008  David Boyer - Added Custom Calendar
' 1.6	11/11/2008	David Boyer - Add new error messages
' 1.7	06/09/2009	David Boyer - Added checkbox for "send to" function.  (Send to features like RSS and eventually Twitter, etc.)
' 1.6	02/10/2012  David Boyer - Expanded the "message" length to 1500
' 1.8	04/09/2013	Steve Loar - Store the event date in a session variable so that it can be set to that on following new events. For Mason, OH
' 1.9	04/11/2014	Steve Loar - Putting repeating options into the page, splitting out the save functionality into a seperate file
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
if isFeatureOffline("calendar") = "Y" OR isFeatureOffline("custom_calendars") = "Y" then
	response.redirect "../admin/outage_feature_offline.asp"
end if

sLevel = "../"  'Override of value from common.asp

'Check to see if this is a Custom Calendar
lcl_calendarfeatureid      = ""
lcl_calendarfeature        = ""
lcl_calendarfeature_url    = ""
lcl_calendarfeature_name   = ""
lcl_feature_rssfeeds       = "rssfeeds_events_communitycalendar"
lcl_pushcontent            = "pushcontent_communitycalendar"
lcl_displayHistory_feature = "displayhistoryinfo"
lcl_eventdate              = Date()
lcl_subject                = ""
lcl_details                = ""
lcl_checked                = " checked=""checked"""
iPushedFromRequestID       = ""
iHour                      = 12
iMinute                    = "00"
sAmPm                      = "AM"
lDuration                  = ""
sDurationInterval          = ""
lcl_customcategory         = ""
iCategoryID                = ""

'Allow the user to maintain the Event Categories if any/all of the following:
'1. The user has the "categories" permission assigned
'2. The user has a specific Custom Calendar feature assigned: [session("calendarfeature") <> ""]
if Trim(request("cal")) <> "" then
	if not isnumeric(trim(request("cal"))) then
		response.redirect sLevel & "permissiondenied.asp"
	else
		lcl_calendarfeatureid      = CLng(Trim(request("cal")))
		lcl_calendarfeature        = getFeatureByID(session("orgid"), lcl_calendarfeatureid)
		lcl_displayHistory_feature = "displayhistoryinfo_customcalendars"

		if OrgHasFeature(lcl_calendarfeature) AND UserHasPermission(session("userid"), lcl_calendarfeature) then
			lcl_calendarfeature_url  = "?cal=" & lcl_calendarfeatureid
			lcl_calendarfeature_name = " [" & getFeatureName(lcl_calendarfeature) & "]"
			lcl_feature_rssfeeds    = "rssfeeds_events_" & lcl_calendarfeature
			lcl_pushcontent         = "pushcontent_" & lcl_calendarfeature
		else
			response.redirect sLevel & "permissiondenied.asp"
		end if
	end if
else
	if NOT userhaspermission( session("userid"), "edit events" ) then
		response.redirect sLevel & "permissiondenied.asp"
	end if
end if

'Check for org features
lcl_orghasfeature_rssfeeds_events    = orghasfeature(lcl_feature_rssfeeds)
lcl_orghasfeature_pushcontent        = orghasfeature(lcl_pushcontent)
lcl_orghasfeature_displayHistoryInfo = orghasfeature(lcl_displayHistory_feature)

'Check for user permissions
lcl_userhaspermission_rssfeeds_events = userhaspermission(session("userid"),lcl_feature_rssfeeds)
lcl_userhaspermission_pushcontent     = userhaspermission(session("userid"),lcl_pushcontent)

Dim oCmd, oRst, dDate, lDuration, sTimeZones, sLinks, bShown

'if request("_task") = "newevent" then

'    dDate                           = CDate(Request.Form("DatePicker") & " " & Request.Form("Hour") & ":" & Request.Form("Minute") & " " & Request.Form("AMPM"))
'	session("eventdate")			= Request.Form("DatePicker")
'    lDuration                       = -1
'    sDurationInterval               = request("DurationInterval")
'    lcl_customcategory              = ""
'    iCategoryID                     = 0
'    lcl_isHiddenCL                  = 1
'    lcl_pushedfrom_requestid        = "NULL"
'    lcl_displayHistoryToPublic      = "0"
'    lcl_displayHistoryOption        = ""
'    lcl_insert_eventdate            = "''"
'    lcl_insert_subject              = "''"
'    lcl_insert_message              = "''"
'    lcl_insert_calendarfeature      = "NULL"
'    lcl_insert_displayHistoryOption = "NULL"

'    if request("duration") <> "" then
'       lDuration = request("Duration")
'       lDuration = CLng(lDuration) * clng(sDurationInterval)
'    end if

   'Create a New Category for this Organization
'   if request("CustomCategory") <> "" then
'       lcl_customcategory = request("CustomCategory")

'       newCategory session("orgid"), lcl_customcategory, "#000000", lcl_calendarfeature, lcl_identity

'       iCategoryID = lcl_identity
'    else
'       iCategoryID = request("Category")
'   	end if

'    if request("isHiddenCL") = "on" then
'       lcl_isHiddenCL = 0
'    end if

'    if request("displayHistoryToPublic") = "Y" then
'       lcl_displayHistoryToPublic = "1"
'    end if

'    if request("displayHistoryOption") <> "" then
'       lcl_displayHistoryOption = request("displayHistoryOption")
'    end if

   'Check to see if this record is being created from a request.
'    if request("requestid") <> "" then
'       lcl_pushedfrom_requestid = request("requestid")
'    end if

   'Set up the fields to be inserted into the table
'    if dDate <> "" then
'       lcl_insert_eventdate = dDate
'       lcl_insert_eventdate = dbsafe(lcl_insert_eventdate)
'       lcl_insert_eventdate = "'" & lcl_insert_eventdate & "'"
'    end if

'    if request("subject") <> "" then
'       lcl_insert_subject = request("subject")
'       lcl_insert_subject = dbsafe(lcl_insert_subject)
'       lcl_insert_subject = left(lcl_insert_subject,50)
'       lcl_insert_subject = "'" & lcl_insert_subject & "'"
'    end if

'    if request("message") <> "" then
'       lcl_insert_message = request("message")
'       lcl_insert_message = dbsafe(lcl_insert_message)
'       lcl_insert_message = left(lcl_insert_message,1500)
'       lcl_insert_message = "'" & lcl_insert_message & "'"
'    end if

'    if lcl_calendarfeature <> "" then
'       lcl_insert_calendarfeature = lcl_calendarfeature
'       lcl_insert_calendarfeature = dbsafe(lcl_insert_calendarfeature)
'       lcl_insert_calendarfeature = "'" & lcl_insert_calendarfeature & "'"
'    end if

'    if lcl_displayHistoryOption <> "" then
'       lcl_insert_displayHistoryOption = lcl_displayHistoryOption
'       lcl_insert_displayHistoryOption = dbsafe(lcl_insert_displayHistoryOption)
'       lcl_insert_displayHistoryOption = "'" & lcl_insert_displayHistoryOption & "'"
'    end if

   'Create the event
'    sSql = "INSERT INTO Events ("
'    sSql = sSql & "OrgID, "
'    sSql = sSql & "CreatorUserID, "
'    sSql = sSql & "EventDate, "
'    sSql = sSql & "EventTimeZoneID, "
'    sSql = sSql & "EventDuration, "
'    sSql = sSql & "[Subject], "
'    sSql = sSql & "[Message], "
'    sSql = sSql & "ModifierUserID, "
'    sSql = sSql & "CategoryID, "
'    sSql = sSql & "calendarfeature, "
'    sSql = sSql & "isHiddenCL, "
'    sSql = sSql & "pushedfrom_requestid, "
'    sSql = sSql & "displayHistoryToPublic, "
'    sSql = sSql & "displayHistoryOption "
'    sSql = sSql & ") VALUES ("
'    sSql = sSql & session("orgid")           & ", "
'    sSql = sSql & session("userid")          & ", "
'    sSql = sSql & lcl_insert_eventdate       & ", "
'    sSql = sSql & request("timezone")        & ", "
'    sSql = sSql & lDuration                  & ", "
'    sSql = sSql & lcl_insert_subject         & ", "
'    sSql = sSql & lcl_insert_message         & ", "
'    sSql = sSql & session("userid")          & ", "
'    sSql = sSql & iCategoryID                & ", "
'    sSql = sSql & lcl_insert_calendarfeature & ", "
'    sSql = sSql & lcl_isHiddenCL             & ", "
'    sSql = sSql & lcl_pushedfrom_requestid   & ", "
'    sSql = sSql & lcl_displayHistoryToPublic & ", "
'    sSql = sSql & lcl_insert_displayHistoryOption
'    sSql = sSql & ") "

'   	session("eventsql")   = sSql
	'response.write sSql
'    lcl_newEventID        = RunIdentityInsertStatement(sSql)
'    lcl_return_parameters = ""

'    if lcl_orghasfeature_rssfeeds_events AND lcl_userhaspermission_rssfeeds_events AND request("sendTo_RSS") = "on" then
'       lcl_return_parameters = "&sendTo_RSS=" & lcl_newEventID
'    end if

'    response.redirect "default.asp?init=Y&success=SA&cal=" & lcl_calendarfeatureid & lcl_return_parameters
'end If

' this was added for Mason OH on 4/9/2013
If session("eventdate") <> "" Then
	lcl_eventdate = session("eventdate")
End If 

'Check to see if this record is being "pushed" from a request
if lcl_orghasfeature_pushcontent And lcl_userhaspermission_pushcontent Then 
    iPushedFromRequestID = request("requestid")

    if iPushedFromRequestID <> "" then
       iPushedFromTrackingNum = ""

       sSql = "SELECT a.[Tracking Number] as trackingnumber "
       sSql = sSql & " FROM egov_rpt_actionline a "
       sSql = sSql & " WHERE a.action_autoid = " & iPushedFromRequestID

       set oTrackNum = Server.CreateObject("ADODB.Recordset")
       oTrackNum.Open sSql, Application("DSN"), 3, 1

       if not oTrackNum.eof then
          iPushedFromTrackingNum = oTrackNum("trackingnumber")
       end if

       oTrackNum.close
       set oTrackNum = nothing

       lcl_eventdate = getPushFieldAnswer("calendar", "events", "eventdate", iPushedFromRequestID)
       lcl_subject   = getPushFieldAnswer("calendar", "events", "subject",   iPushedFromRequestID)
       lcl_details   = getPushFieldAnswer("calendar", "events", "message",   iPushedFromRequestID)

    end if
end if

'Setup the BODY onload
lcl_onload = ""
lcl_onload = lcl_onload & "setMaxLength();"
lcl_onload = lcl_onload & "document.getElementById('DatePicker').focus();"

%>

<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	<title><%=langBSEVents%><%=lcl_calendarfeature_name%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />	
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="eventstyles.css" />

	<script language="javascript" src="../scripts/ajaxLib.js"></script>
	<script language="javascript" src="../scripts/isvaliddate.js"></script>
	<script language="javascript" src="../scripts/textareamaxlength.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
	<script type="text/javascript" src="../scripts/jquery-1.7.1.min.js"></script>


	<script language="javascript">
	<!--

		$(document).ready(function() {
			$('#displayHistoryOption').prop('disabled','disabled');

			$('#displayHistoryToPublic').click(function() {
				if($('#displayHistoryToPublic').prop('checked')) {
					$('#displayHistoryOption').prop('disabled','');
				} else {
					$('#displayHistoryOption').prop('disabled','disabled');
				}
			});
		});

		function doCalendar( sField ) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=NewEvent", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function doPicker(sFormField, p_displayDocuments, p_displayActionLine, p_displayPayments, p_displayURL) 
		{
			w = 600;
			h = 400;
			l = (screen.availWidth/2)-(w/2);
			t = (screen.availHeight/2)-(h/2);

			lcl_showFolderStart = "";
			lcl_folderStart     = 0;
			//  lcl_displayLinkText = "&displayLinkText=Y"

			//Determine which options will be displayed
			if((p_displayDocuments=="")||(p_displayDocuments==undefined)) 
			{
				lcl_displayDocuments = "";
			}
			else
			{
				lcl_displayDocuments = "&displayDocuments=Y";
				lcl_folderStart = lcl_folderStart + 1;
			}

			if((p_displayActionLine=="")||(p_displayActionLine==undefined)) 
			{
				lcl_displayActionLine = "";
			}
			else
			{
				lcl_displayActionLine = "&displayActionLine=Y";
				lcl_folderStart = lcl_folderStart + 1;
			}

			if((p_displayPayments=="")||(p_displayPayments==undefined)) 
			{
				lcl_displayPayments = "";
			}
			else
			{
				lcl_displayPayments = "&displayPayments=Y";
				lcl_folderStart = lcl_folderStart + 1;
			}

			if((p_displayURL=="")||(p_displayURL==undefined)) 
			{
				lcl_displayURL = "";
			}
			else
			{
				lcl_displayURL = "&displayURL=Y";
			}

			if(lcl_folderStart > 0) 
			{
				//lcl_showFolderStart = "&folderStart=unpublished_documents";
				lcl_showFolderStart = "&folderStart=CITY_ROOT";
			}

			pickerURL  = "../picker_new/default.asp";
			pickerURL += "?name=" + sFormField;
			pickerURL += lcl_showFolderStart;
			pickerURL += lcl_displayDocuments;
			pickerURL += lcl_displayActionLine;
			pickerURL += lcl_displayPayments;
			pickerURL += lcl_displayURL;
			//  pickerURL += lcl_displayLinkText;

			eval('window.open("' + pickerURL + '", "_picker", "width=' + w + ',height=' + h + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
		}

		function insertAtCaret (textEl, text) 
		{
			if (textEl.createTextRange && textEl.caretPos) 
			{
				var caretPos = textEl.caretPos;
				caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ? text + ' ' : text;
			}
			else 
			{
				textEl.value = textEl.value + text;
			}
		}

		function storeCaret (textEl) 
		{
			if (textEl.createTextRange) 
			{
				textEl.caretPos = document.selection.createRange().duplicate();
			}
		}

		
		function validate()	
		{
			var lcl_return_false = "N";

			clearAllMsgs();

			if (document.getElementById("message").value.length > 1500)	
			{
				document.getElementById("message").focus();
				inlineMsg(document.getElementById("message").id,'<strong>Invalid Value: </strong>Details cannot be longer than 1500 characters [current length: ' + document.getElementById("message").value.length + ']',5,'message');
				lcl_return_false = "Y";
			}

			if (document.getElementById("subject").value == '') 
			{
				document.getElementById("subject").focus();
				inlineMsg(document.getElementById("subject").id,'<strong>Required Field: </strong>Subject',5,'Subject');
				lcl_return_false = "Y";
			}

			if (document.getElementById("subject").value.length > 50)	
			{
				document.getElementById("subject").focus();
				inlineMsg(document.getElementById("subject").id,'<strong>Invalid Value: </strong>Subject cannot be longer than 50 characters [current length: ' + document.getElementById("subject").value.length + ']',5,'message');
				lcl_return_false = "Y";
			}

			//if (document.NewEvent.Subject.value == '') {
			//  	document.getElementById("subject").focus();
			//		 inlineMsg(document.getElementById("subject").id,'<strong>Required Field: </strong>Subject',10,'Subject');
			//   lcl_return_false = "Y";
			//}

			if(document.getElementById("Duration").value!="") 
			{
				var rege = /^\d+$/;
				var Ok = rege.exec(document.getElementById("Duration").value);

				if ( ! Ok ) 
				{
					document.getElementById("Duration").focus();
					inlineMsg(document.getElementById("DurationInterval").id,'<strong>Invalid Value: </strong>Duration must be a numeric value.',5,'DurationInterval');
					lcl_return_false = "Y";
				}	
			}

			if(document.getElementById("DatePicker").value == "") 
			{
				document.getElementById("DatePicker").focus();
				inlineMsg(document.getElementById("DatePickerLOV").id,'<strong>Required Field: </strong>Date',5,'DatePickerLOV');
				lcl_return_false = "Y";
			}
			else
			{
				if(! isValidDate(document.getElementById("DatePicker").value)) 
				{
					document.getElementById("DatePicker").focus();
					inlineMsg(document.getElementById("DatePickerLOV").id,'<strong>Invalid Value: </strong>Date must be in the format of MM/DD/YYYY',5,'DatePickerLOV');
					lcl_return_false = "Y";
				}
			}

			// check the recurring fields here if the repeating checkbox is checked
			if ($('#isrepeating').is(":checked"))
			{
				var errorCount = 0;
				errorCount = CheckRecur();
				if ( errorCount > 0 )
				{
					lcl_return_false = "Y";
				}
			}

			if (lcl_return_false == "Y") 
			{
				return false;
			}
			else
			{
				document.NewEvent.submit();
				return true;
			}
		}

		
		function CheckRecur()
		{
			var errorCount = 0;
			var rege;
		 	var Ok;

			// Range of Recurrence checks, at bottom of repeating section
			var howLong = $('input:radio[name=howlong]:checked').val();
			if (howLong == 'till')
			{
				if ($("#howmany").val() == "")
				{
					inlineMsg(document.getElementById("howmany").id,'<strong>Required Field: </strong>Occurrences',5,'howmany');
					errorCount++;
				}
				else
				{
					// validate that this is numeric and a whole number
					rege = /^\d+$/;
      				Ok = rege.test($("#howmany").val());
      				if (! Ok) 
					{
						inlineMsg(document.getElementById("howmany").id,'<strong>Invalid Value: </strong> "Occurrences" must be a numeric value',5,'howmany');
						errorCount++;
					}
				}
			}
			else	// howLong is 'endby'
			{
				if ($("#endbydate").val() == "")
				{
					inlineMsg(document.getElementById("endbydate").id,'<strong>Required Field: </strong>End by Date',5,'endbydate');
					errorCount++;
				}
				else
				{
					// validate the format of the date
					rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
      				Ok = rege.test($("#endbydate").val());

      				if (! Ok) 
					{
						inlineMsg(document.getElementById("endbydate").id,'<strong>Invalid Value: </strong>The "End By Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',5,'endbydate');
						errorCount++;
					}
				}
			}

			// Recurrence checks, at top of repeating section
			var recPick = $("#recur").val();

			switch (recPick)
			{
				case "dd":
					// if recur is daily
					var dayOften = $('input:radio[name=dayoften]:checked').val();
					if (dayOften == 'days')
					{
						if ($("#days").val() == "")
						{
							inlineMsg(document.getElementById("days").id,'<strong>Required Field: </strong>Days',5,'days');
							errorCount++;
						}
						else
						{
							// validate the format of the days
							rege = /^\d+$/;
							Ok = rege.test($("#days").val());

							if (! Ok) 
							{
								inlineMsg(document.getElementById("days").id,'<strong>Invalid Value: </strong>"Days" must be a numeric value',5,'days');
								errorCount++;
							}
						}
					}
					break;

				case "ww":
					// if recur is weekly
					if ($("#weeks").val() == "")
					{
						inlineMsg(document.getElementById("weeks").id,'<strong>Required Field: </strong>Occurrences',5,'weeks');
						errorCount++;
					}
					else
					{
						// validate the format of the days
						rege = /^\d+$/;
						Ok = rege.test($("#weeks").val());

						if (! Ok) 
						{
							inlineMsg(document.getElementById("weeks").id,'<strong>Invalid Value: </strong>Occurrences must be a numeric value',5,'weeks');
							errorCount++;
						}
					}
					// make sure at least one DOW is checked
					var total_weekdays = document.NewEvent.WeekDayNum.length;
					var weekDaySelected = "no"

					//Determine if at least one value is "checked"
					for (i = 0; i <= total_weekdays -1; i++)
					{
						if(document.NewEvent.WeekDayNum[i].checked == true) 
						{
							weekDaySelected = "yes";
							break;
						}
					}

					//No day(s) have been selected
					if(weekDaySelected == "no") 
					{
						inlineMsg(document.getElementById("Wednesday").id,'<strong>Required Field Missing: </strong> At least one day must be selected',5,'Wednesday');
					}
					break;

				case "mm":
					// if recur is monthly
					var monthOften = $('input:radio[name=monthOften]:checked').val();
					if (monthOften == 'absolute')
					{

						if ($("#monthDay").val() == "")
						{
							inlineMsg(document.getElementById("monthDay").id,'<strong>Required Field: </strong>Day',5,'monthDay');
							errorCount++;
						}
						else
						{
							// validate the format of the days
							rege = /^\d+$/;
							Ok = rege.test($("#monthDay").val());

							if (! Ok) 
							{
								inlineMsg(document.getElementById("monthDay").id,'<strong>Invalid Value: </strong>"Day" must be a numeric value',5,'monthDay');
								errorCount++;
							}
						}

						if ($("#monthQty").val() == "")
						{
							inlineMsg(document.getElementById("monthQty").id,'<strong>Required Field: </strong>Months',5,'monthQty');
							errorCount++;
						}
						else
						{
							// validate the format of the month qty
							rege = /^\d+$/;
							Ok = rege.test($("#monthQty").val());

							if (! Ok) 
							{
								inlineMsg(document.getElementById("monthQty").id,'<strong>Invalid Value: </strong>"Months" must be a numeric value',5,'monthQty');
								errorCount++;
							}
						}

					}
					else
					{
						if ($("#Month").val() == "")
						{
							inlineMsg(document.getElementById("Month").id,'<strong>Required Field: </strong>Day',5,'Month');
							errorCount++;
						}
						else
						{
							// validate the format of the days
							rege = /^\d+$/;
							Ok = rege.test($("#Month").val());

							if (! Ok) 
							{
								inlineMsg(document.getElementById("Month").id,'<strong>Invalid Value: </strong>"Month" must be a numeric value',5,'Month');
								errorCount++;
							}
						}
					}
					break;

				case "yy":
					// if recur is yearly
					var yearOften = $('input:radio[name=yearOften]:checked').val();
					if (yearOften == 'every')
					{

						if ($("#yearDay").val() == "")
						{
							inlineMsg(document.getElementById("yearDay").id,'<strong>Required Field: </strong>Day',5,'yearDay');
							errorCount++;
						}
						else
						{
							// validate the format of the days
							rege = /^\d+$/;
							Ok = rege.test($("#yearDay").val());

							if (! Ok) 
							{
								inlineMsg(document.getElementById("yearDay").id,'<strong>Invalid Value: </strong>"Day" must be a numeric value',5,'yearDay');
								errorCount++;
							}
						}
					}
					else
					{
						if ($("#yearDayNum").val() == "")
						{
							inlineMsg(document.getElementById("yearDayNum").id,'<strong>Required Field: </strong>Day',5,'yearDayNum');
							errorCount++;
						}
						else
						{
							// validate the format of the days
							rege = /^\d+$/;
							Ok = rege.test($("#yearDayNum").val());

							if (! Ok) 
							{
								inlineMsg(document.getElementById("yearDayNum").id,'<strong>Invalid Value: </strong>"Day" must be a numeric value',5,'yearDayNum');
								errorCount++;
							}
						}

						if ($("#yearMonths").val() == "")
						{
							inlineMsg(document.getElementById("yearMonths").id,'<strong>Required Field: </strong>Months',5,'yearMonths');
							errorCount++;
						}
						else
						{
							// validate the format of the month qty
							rege = /^\d+$/;
							Ok = rege.test($("#yearMonths").val());

							if (! Ok) 
							{
								inlineMsg(document.getElementById("yearMonths").id,'<strong>Invalid Value: </strong>"Months" must be a numeric value',5,'yearMonths');
								errorCount++;
							}
						}
					}
					break;
			}

			return errorCount;
		}


		function toggleRepeatingBlock()
		{
			// show and hide the repeating block
			if ($('#isrepeating').is(":checked"))
			{
				$("#repeatingdisplay").slideDown( "slow" );
			}
			else
			{
				clearAllMsgs();
				$("#repeatingdisplay").slideUp(  );
			}
		}

		function showPicks()
		{
			clearAllMsgs();
			$("#daypicks").hide( );
			$("#weekpicks").hide(  );
			$("#monthpicks").hide(  );
			$("#yearpicks").hide(  );

			var recPick = $("#recur").val();

			switch (recPick)
			{
				case "dd":
					$("#daypicks").slideDown( "slow" );
					break;
				case "ww":
					$("#weekpicks").slideDown( "slow" );
					break;
				case "mm":
					$("#monthpicks").slideDown( "slow" );
					break;
				case "yy":
					$("#yearpicks").slideDown( "slow" );
					break;
			}
		}

		function clearAllMsgs()
		{
			 clearMsg('DatePicker');
			 clearMsg('Duration');
			 clearMsg('subject');
			 clearMsg('message');
			 clearMsg('days');
			 clearMsg('weeks');
			 clearMsg('Wednesday');
			 clearMsg('monthDay');
			 clearMsg('monthQty');
			 clearMsg('Month');
			 clearMsg('yearDay');
			 clearMsg('yearDayNum');
			 clearMsg('yearMonths');
			 clearMsg('howmany');
			 clearMsg('endbydate');
		}

	//-->
	</script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 
<%
	response.write "<form name=""NewEvent"" id=""NewEvent"" method=""post"" action=""neweventcreate.asp"" accept-charset=""UTF-8"">" & vbcrlf
	response.write "  <input type=""hidden"" name=""_task"" id=""_task"" value=""newevent"" />" & vbcrlf
	response.write "  <input type=""hidden"" name=""timezone"" id=""timezone"" value=""1"" />" & vbcrlf
	response.write "  <input type=""hidden"" name=""cal"" id=""cal"" value=""" & lcl_calendarfeatureid & """ size=""20"" maxlength=""50"" />" & vbcrlf
	response.write "  <input type=""hidden"" name=""control_field"" id=""control_field"" value="""" size=""20"" maxlength=""4001"" />" & vbcrlf
	response.write "  <input type=""hidden"" name=""requestid"" id=""requestid"" value=""" & iPushedFromRequestID & """ />" & vbcrlf

	response.write "<div id=""content"">" & vbcrlf
	response.write "	 <div id=""centercontent"">" & vbcrlf


	response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
	response.write "  <tr valign=""top"">" & vbcrlf
	response.write "      <td>" & vbcrlf
	response.write "          <font size=""+1""><strong>Events: New" & lcl_calendar_name & "</strong></font><br />" & vbcrlf

	if iPushedFromRequestID <> "" then
		response.write "<input type=""button"" name=""returnToRequestButton"" id=""returnToRequestButton"" class=""button"" value=""Return to Request"" onclick=""location.href='../action_line/action_respond.asp?control=" & iPushedFromRequestID & "';"" />" & vbcrlf
	end if

	response.write "      </td>" & vbcrlf
	response.write "      <td align=""right""><span id=""screenMsg""></span></td>" & vbcrlf
	response.write "  </tr>" & vbcrlf
	response.write "  <tr>" & vbcrlf
	response.write "      <td colspan=""2"" valign=""top"">" & vbcrlf
	response.write "          <div class=""displayButtonsDIV"">" & vbcrlf
	displayButtons
	response.write "        		</div>" & vbcrlf
	'response.write "          <div class=""shadow"">" & vbcrlf

	response.write "		        <table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"" class=""tableadmin"" id=""neweventinput"">" & vbcrlf
	response.write "            <tr>" & vbcrlf
	response.write "                <th align=""left"" colspan=""2"">" & langNewEvent & "</th>" & vbcrlf
	response.write "            </tr>" & vbcrlf

	'Date
	response.write "            <tr>" & vbcrlf
	response.write "                <td>" & langDate & ":</td>" & vbcrlf
	response.write "                <td>" & vbcrlf
	response.write "                    <input type=""text"" name=""DatePicker"" id=""DatePicker"" maxlength=""10"" value=""" & lcl_eventdate & """ onchange=""clearMsg('DatePicker')"" />&nbsp;" & vbcrlf
	response.write "                    <a href=""javascript:void doCalendar('DatePicker');""><img src=""../images/calendar.gif"" border=""0"" /></a>" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "            </tr>" & vbcrlf

	'Time
	response.write "            <tr>" & vbcrlf
	response.write "                <td>" & langStartTime & ":</td>" & vbcrlf
	response.write "                <td>" & vbcrlf
	response.write "                    <select name=""Hour"" id=""hour"" class=""time"">" & vbcrlf
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
	response.write "                    </select>" & vbcrlf
	response.write "                    :" & vbcrlf
	response.write "                    <select name=""Minute"" id=""minute"" class=""time"">" & vbcrlf

                                        for i = 0 to 59 step 5
                                           lcl_displayMinute = i

                                           if i < 10 then
                                              lcl_displayMinute = "0" & i
                                           end if

                                           buildOption "MINUTE", lcl_displayMinute, iMinute
                                        next

	response.write "                    </select>" & vbcrlf
	response.write "                    <select name=""AMPM"" class=""time"">" & vbcrlf
                                        buildOption "AMPM", "AM", sAmPm
                                        buildOption "AMPM", "PM", sAmPm
	response.write "                    </select>" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "            </tr>" & vbcrlf

	'Duration
	response.write "            <tr>" & vbcrlf
	response.write "                <td>" & langDuration & ":</td>" & vbcrlf
	response.write "                <td>" & vbcrlf
	response.write "                    <input type=""text"" name=""Duration"" id=""Duration"" maxlength=""5"" value=""" & lDuration & """ onchange=""clearMsg('Duration');"" />" & vbcrlf
	response.write "                    <select name=""DurationInterval"" id=""DurationInterval"" class=""time"" style=""width:80px;"" onchange=""clearMsg('Duration');"">" & vbcrlf
                                        buildOption "DURATION", "1",     sDurationInterval
                                        buildOption "DURATION", "60",    sDurationInterval
                                        buildOption "DURATION", "1440",  sDurationInterval
                                        buildOption "DURATION", "10080", sDurationInterval
	response.write "                    </select>" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "            </tr>" & vbcrlf

	'Category
	response.write "            <tr>" & vbcrlf
	response.write "                <td>Category:</td>" & vbcrlf
	response.write "                <td>" & vbcrlf
	response.write "                   	Choose:" & vbcrlf
	response.write "                    <select name=""Category"" id=""category"" class=""time"">" & vbcrlf
	response.write "                      <option value=""0"">None</option>" & vbcrlf
                                        getEventCategoryOptions session("orgid"), lcl_calendarfeature, iCategory
	response.write "                    </select>" & vbcrlf
	response.write "                    OR New Category:" & vbcrlf
	response.write "                    <input type=""text"" name=""CustomCategory"" id=""customcategory"" maxlength=""50"" />" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "            </tr>" & vbcrlf

	'Subject
	response.write "            <tr>" & vbcrlf
	response.write "                <td>" & langSubject & ":</td>" & vbcrlf
	response.write "                <td><input type=""text"" name=""subject"" id=""subject"" size=""65"" maxlength=""50"" value=""" & lcl_subject & """ /></td>" & vbcrlf
	response.write "            </tr>" & vbcrlf

	'Details
	response.write "            <tr>" & vbcrlf
	response.write "                <td valign=""top"">" & langDetails & ":&nbsp;</td>" & vbcrlf
	response.write "                <td>" & vbcrlf
	response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
	response.write "                      <tr>" & vbcrlf
	response.write "                          <td width=""400"">" & vbcrlf
	response.write "                              <textarea name=""message"" id=""message"" cols=""100"" rows=""12"" maxlength=""1500"" onchange=""clearMsg('message');"">" & lcl_details & "</textarea>" & vbcrlf
	response.write "                          </td>" & vbcrlf
	response.write "                          <td align=""left"" valign=""top"">" & vbcrlf
	response.write "                              <input type=""button"" class=""button"" value=""Add Link"" onclick=""doPicker('NewEvent.message','Y','Y','Y','Y');"">" & vbcrlf
	response.write "                          </td>" & vbcrlf
	response.write "                      </tr>" & vbcrlf
	response.write "                    </table>" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "            </tr>" & vbcrlf

	' Repeating picks
	response.write "            <tr>" & vbcrlf
	response.write "                <td>&nbsp;</td>"
	response.write "                <td>" & vbcrlf

	%>
										<input type="checkbox" name="isrepeating" id="isrepeating" value="on" onclick="toggleRepeatingBlock()" />&nbsp;This event repeats.
										<div id="repeatingdisplay">
											<label for="recur">Recurrence: </label>
											<select id="recur" name="recur" class="time" onchange="showPicks()">
												<option value="dd" selected>Daily</option>
												<option value="ww">Weekly</option>
												<option value="mm">Monthly</option>
												<option value="yy">Yearly</option>
											</select>

											<div id="daypicks" class="recurpicks">
												<div class="eventpicks">
													<input type="radio" id="everydaypick" name="dayoften" value="days" checked="checked" /> 
													Every <input type="text" name="days" id="days" size="3" maxlength="3" value="1" /> day(s).
												</div>
												<div class="eventpicks">
													<input type="radio" id="weekdaypick" name="dayoften" value="weekdays" /> Every weekday.
												</div>
											</div>

											<div id="weekpicks" class="recurpicks">
												Recur every <input type="text" name="weeks" id="weeks" class="eventinput" maxlength="2" value="1" /> week(s) on: <br />
												<div class="eventpicks">
													<input type="checkbox" name="WeekDayNum" value="1" id="Sunday" /> Sunday 
													<input type="checkbox" name="WeekDayNum" value="2" id="Monday" /> Monday 
													<input type="checkbox" name="WeekDayNum" value="3" id="Tuesday" /> Tuesday 
													<input type="checkbox" name="WeekDayNum" value="4" id="Wednesday" /> Wednesday 
												</div>
												<div class="eventpicks">
													<input type="checkbox" name="WeekDayNum" value="5" id="Thursday" /> Thursday
													<input type="checkbox" name="WeekDayNum" value="6" id="Friday" /> Friday
													<input type="checkbox" name="WeekDayNum" value="7" id="Saturday" /> Saturday
												</div>
											</div>

											<div id="monthpicks" class="recurpicks">
												<div class="eventpicks">
													<input type="radio" name="monthOften" value="absolute" checked="checked" /> Day
													<input type="text" name="monthDay" id="monthDay" class="eventinput" maxlength="2" value="1" /> of every 
													<input type="text" name="monthQty" id="monthQty" class="eventinput" maxlength="2" value="1" /> month(s).<br />
												</div>
												<div class="eventpicks">
													<input type="radio" name="monthOften" value="relative" /> The
													<select name="monthOrdinal" class="time">
<%														displayOrdinalOptions	%>
													</select>
													<select name="DayLike" class="time">
<%														displayDayLikeOptions	%>
													</select> 
													of every <input type="text" name="Month" id="Month" class="eventinput" maxlength="2" value="1" /> month(s).
												</div>
											</div>

											<div id="yearpicks" class="recurpicks">
												<div class="eventpicks">
													<input type="radio" name="yearOften" value="every" checked="checked" /> Every 
													<select name="yearMonth" class="time">
<%														displayMonthOptions		%>
													</select>
													<input type="text" name="yearDay" id="yearDay" class="eventinput" maxlength="2" value="1" />.
												</div>
												<div class="eventpicks">
													<input type="radio" name="yearOften" value="absolute" /> Day
													<input type="text" name="yearDayNum" id="yearDayNum" class="eventinput" maxlength="2" value="1" /> of every 
													<input type="text" name="yearMonths" id="yearMonths" class="eventinput" maxlength="2" value="1" /> month(s).
												</div>
												<div class="eventpicks">
													<input type="radio" name="yearOften" value="relative" /> The 
													<select name="yearOrdinal" class="time">
<%														displayOrdinalOptions		%>
													</select> 
													<select name="yearDayPick" class="time">
<%														displayDayLikeOptions		%>
													</select> of 
													<select name="yearMonthPick" class="time">
<%														displayMonthOptions		%>
													</select>
												</div>
											</div>

											<div class="eventsectiontitle">Range of Recurrence:</div>
											<div id="occurrencepicks">
												<div class="eventpicks">
													<input type="radio" id="howlong1" name="howlong" value="till" checked="checked" /> End after 
													<input type="text" name="howmany" id="howmany" class="eventinput" maxlength="4" value="1" /> additional occurrences.
												</div>
												<div class="eventpicks">
													<input type="radio" id="howlong2" name="howlong" value="endby" /> End by <input type="text" name="endbydate" id="endbydate" maxlength="50" value="<%=Date()%>" />
								                    <a href="javascript:void doCalendar('endbydate');"><img src="../images/calendar.gif" border="0" /></a>
												</div>
											</div>
										</div>
<%

	response.write "                </td>" & vbcrlf
	response.write "            </tr>" & vbcrlf
	
	' Show on Community Link
	response.write "  	       <tr>" & vbcrlf
	response.write "                <td>&nbsp;</td>" & vbcrlf
	response.write "                <td><input type=""checkbox"" name=""isHiddenCL"" id=""isHiddenCL"" value=""on""" & lcl_checked & " />&nbsp;Show on Community Link</td>" & vbcrlf
	response.write "            </tr>" & vbcrlf

	if lcl_orghasfeature_rssfeeds_events AND lcl_userhaspermission_rssfeeds_events then
		response.write "            <tr valign=""top"">" & vbcrlf
		response.write "                <td>On Create Send To:</td>" & vbcrlf
		response.write "                <td>" & vbcrlf
		displaySendToOption "RSS", "ADD", "Y", lcl_orghasfeature_rssfeeds_events, lcl_userhaspermission_rssfeeds_events
		response.write "                </td>" & vbcrlf
		response.write "            </tr>" & vbcrlf
	end if

	'Display History Info
	response.write "            <tr>" & vbcrlf
	response.write "                <td colspan=""2"">" & vbcrlf
	response.write "                    <fieldset class=""fieldset"">" & vbcrlf
	response.write "                      <legend>History Log&nbsp;</legend>" & vbcrlf

	if iCreatedByID <> "" OR (iPushedFromRequestID <> "" AND lcl_orghasfeature_pushcontent AND lcl_userhaspermission_pushcontent) then
		response.write "                      <table border=""0"" cellspacing=""0"" cellpadding=""2"" style=""margin-top:5px;"">" & vbcrlf

		'Created By
		if iCreatedByID <> "" then
			lcl_createdby = iCreatedByName & " on " & iCreatedDate

			response.write "                        <tr>" & vbcrlf
			response.write "                            <td><strong>Created By:</strong></td>" & vbcrlf
			response.write "                            <td style=""color:#800000"">" & lcl_createdby & "</td>" & vbcrlf
			response.write "                        </tr>" & vbcrlf
		end if

		'Last Updated By
		if iLastUpdatedByID <> "" then
			lcl_lastupdatedby = iLastUpdatedByName & " on " & iLastUpdatedDate

			response.write "                        <tr>" & vbcrlf
			response.write "                            <td><strong>Last Updated By:</strong></td>" & vbcrlf
			response.write "                            <td style=""color:#800000"">" & lcl_lastupdatedby & "</td>" & vbcrlf
			response.write "                        </tr>" & vbcrlf
		end if

		'Originated From
		if iPushedFromRequestID <> "" AND lcl_orghasfeature_pushcontent AND lcl_userhaspermission_pushcontent then
			response.write "                        <tr>" & vbcrlf
			response.write "                            <td><strong>Originated From:</strong></td>" & vbcrlf
			response.write "                            <td><a href=""../action_line/action_respond.asp?control=" & iPushedFromRequestID & """>" & iPushedFromTrackingNum & "</a></td>" & vbcrlf
			response.write "                        </tr>" & vbcrlf
		end if

		response.write "                      </table>" & vbcrlf
	end if

	'Display History Log Options
	if lcl_orghasfeature_displayHistoryInfo then
		response.write "                      <p>" & vbcrlf
		response.write "                        <input type=""checkbox"" name=""displayHistoryToPublic"" id=""displayHistoryToPublic"" value=""Y"" />" & vbcrlf
		response.write "                        Display history on public calendar" & vbcrlf
		response.write "                        <select name=""displayHistoryOption"" id=""displayHistoryOption"">" & vbcrlf
		response.write "                          <option value="""">All Info</option>" & vbcrlf
		response.write "                          <option value=""Names Only"">Names Only</option>" & vbcrlf
		response.write "                          <option value=""Date/Time Only"">Date/Time Only</option>" & vbcrlf
		response.write "                        </select>" & vbcrlf
		response.write "                      </p>" & vbcrlf
	end if

	response.write "                    </fieldset>" & vbcrlf
	response.write "                </td>" & vbcrlf
	response.write "            </tr>" & vbcrlf
	response.write "          </table>" & vbcrlf
	'response.write "      		  </div>" & vbcrlf
	response.write "          <div class=""displayButtonsDIV"">" & vbcrlf
	displayButtons
	response.write "          </div>" & vbcrlf
	response.write "      </td>" & vbcrlf
	response.write "    </tr>" & vbcrlf
	response.write "  </table>" & vbcrlf
	response.write "  	</div>" & vbcrlf
	response.write "</div>" & vbcrlf
	response.write "</form>" & vbcrlf
%>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%

'------------------------------------------------------------------------------
Sub displayButtons( )

	response.write "<input type=""button"" value=""" & langCancel & """ class=""button"" onclick=""history.back();"" />" & vbcrlf
	response.write "<input type=""button"" value=""" & langCreate & """ class=""button"" onclick=""return validate();"" />" & vbcrlf

end sub


'------------------------------------------------------------------------------
Function dbsafe( ByVal p_value )
	Dim lcl_return

	lcl_return = ""
	lcl_return = Replace(p_value,"'","''")
	dbsafe = lcl_return

End Function 


'------------------------------------------------------------------------------
Sub displayOrdinalOptions( )

	response.write "<option value=""1"">First</option>" & vbcrlf
	response.write "<option value=""2"">Second</option>" & vbcrlf
	response.write "<option value=""3"">Third</option>" & vbcrlf
	response.write "<option value=""4"">Fourth</option>" & vbcrlf
	response.write "<option value=""5"">Last</option>" & vbcrlf

End Sub 


'------------------------------------------------------------------------------
Sub displayDayLikeOptions( )

	response.write "<option value=""d"">Day</option>" & vbcrlf
	response.write "<option value=""wd"">Weekday</option>" & vbcrlf
	response.write "<option value=""wed"">Weekend Day</option>" & vbcrlf
	response.write "<option value=""1"">Sunday</option>" & vbcrlf
	response.write "<option value=""2"">Monday</option>" & vbcrlf
	response.write "<option value=""3"">Tuesday</option>" & vbcrlf
	response.write "<option value=""4"">Wednesday</option>" & vbcrlf
	response.write "<option value=""5"">Thursday</option>" & vbcrlf
	response.write "<option value=""6"">Friday</option>" & vbcrlf
	response.write "<option value=""7"">Saturday</option>" & vbcrlf

End Sub 



'------------------------------------------------------------------------------
Sub displayMonthOptions( )

	response.write "<option value=""1"">January</option>" & vbcrlf
	response.write "<option value=""2"">February</option>" & vbcrlf
	response.write "<option value=""3"">March</option>" & vbcrlf
	response.write "<option value=""4"">April</option>" & vbcrlf
	response.write "<option value=""5"">May</option>" & vbcrlf
	response.write "<option value=""6"">June</option>" & vbcrlf
	response.write "<option value=""7"">July</option>" & vbcrlf
	response.write "<option value=""8"">August</option>" & vbcrlf
	response.write "<option value=""9"">September</option>" & vbcrlf
	response.write "<option value=""10"">October</option>" & vbcrlf
	response.write "<option value=""11"">November</option>" & vbcrlf
	response.write "<option value=""12"">December</option>" & vbcrlf

End Sub 





%>
