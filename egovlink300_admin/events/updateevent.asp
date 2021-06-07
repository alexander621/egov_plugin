<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="events_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: updateevents.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the calendar
'
' MODIFICATION HISTORY
' 1.0 ??/??/??	???? - Initial Version
' 1.1	10/11/2006	Steve Loar - Security, Header and nav changed
' 1.2	01/24/2008	Steve Loar - Put check on mesage length to prevent crashes when length is greater than 1000.
' 1.3 08/05/2008 David Boyer - Added Custom Calendar
' 1.4 04/07/2009 David Boyer - Added checkbox "Show on CommunityLink"
' 1.5 06/09/2009	David Boyer - Added checkbox for "send to" function.  (Send to features like RSS and eventually Twitter, etc.)
' 1.6 02/10/2012 David Boyer - Expanded the "message" length to 1500
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("calendar") = "Y" OR isFeatureOffline("custom_calendars") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 dim oCmd, oRst, lcl_subject, intID, intTZ, lcl_details, sDate, iHour
 dim iMinute, sAmPm, dDate, lDuration, sDurationInterval
 dim sLinks, bShown, sTimezones

'Check to see if this is a Custom Calendar
' lcl_calendarfeature     = trim(request("cal"))
 lcl_calendarfeatureid      = ""
 lcl_calendarfeature        = ""
 lcl_calendarfeature_url    = ""
 lcl_calendar_name          = ""
 lcl_feature_rssfeeds       = "rssfeeds_events_communitycalendar"
 lcl_pushcontent            = "pushcontent_communitycalendar"
 lcl_displayHistory_feature = "displayhistoryinfo"

 if trim(request("cal")) <> "" then
    if not isnumeric(trim(request("cal"))) then
      	response.redirect sLevel & "permissiondenied.asp"
    else
       lcl_calendarfeatureid      = CLng(trim(request("cal")))
       lcl_calendarfeature        = getFeatureByID(session("orgid"), lcl_calendarfeatureid)
       lcl_displayHistory_feature = "displayhistoryinfo_customcalendars"

       'if OrgHasFeature(trim(request("cal"))) AND UserHasPermission(session("userid"), trim(request("cal"))) then
       if OrgHasFeature(lcl_calendarfeature) AND UserHasPermission(session("userid"), lcl_calendarfeature) then
          lcl_calendarfeature_url  = "?cal=" & lcl_calendarfeatureid
          lcl_calendarfeature_name = " [" & getFeatureName(lcl_calendarfeature) & "]"
          lcl_feature_rssfeeds     = "rssfeeds_events_" & lcl_calendarfeature
          lcl_pushcontent          = "pushcontent_" & lcl_calendarfeature
       else
         	response.redirect sLevel & "permissiondenied.asp"
       end if
    end if
 end if

 if lcl_calendarfeature = "" AND NOT UserHasPermission( Session("UserId"), "edit events" ) then
 	  response.redirect sLevel & "permissiondenied.asp"
 end if

'Check for org features
 lcl_orghasfeature_rssfeeds_events    = orghasfeature(lcl_feature_rssfeeds)
 lcl_orghasfeature_pushcontent        = orghasfeature(lcl_pushcontent)
 lcl_orghasfeature_displayHistoryInfo = orghasfeature(lcl_displayHistory_feature)

'Check for user permissions
 lcl_userhaspermission_rssfeeds_events = userhaspermission(session("userid"),lcl_feature_rssfeeds)
 lcl_userhaspermission_pushcontent     = userhaspermission(session("userid"),lcl_pushcontent)

'Get the eventid and verify it is in proper format
 if request("id") <> "" then
    lcl_eventid = CLng(request("id"))
 else
    response.redirect "default.asp" & lcl_calendarfeature_url
 end if

If Request.Form("_task") <> "" Then

  dDate     = CDate(Request.Form("DatePicker") & " " & Request.Form("Hour") & ":" & Request.Form("Minute") & " " & Request.Form("AMPM"))
  lDuration = Request.Form("Duration")

  If lDuration & "" <> "" Then
     lDuration = CLng(lDuration) * clng(Request.Form("DurationInterval"))
  Else
     lDuration = -1
  End If

  if Request.Form("CustomCategory") <> "" Then

    'Create a New Category for this Organization
     'newCategory session("orgid"), request.form("CustomCategory"), "#000000", session("calendarfeature"), lcl_identity
     newCategory session("orgid"), request.form("CustomCategory"), "#000000", lcl_calendarfeature, lcl_identity

     iCategoryID = lcl_identity

 	else
	    iCategoryID = request("Category")
  end if

  if request("isHiddenCL") = "on" then
     lcl_isHiddenCL = 0
  else
     lcl_isHiddenCL = 1
  end if

 'Set up the fields to be inserted into the table
  lcl_update_eventdate              = "NULL"
  lcl_update_subject                = "NULL"
  lcl_update_message                = "''"
  lcl_update_calendarfeature        = "NULL"
  lcl_update_displayHistoryToPublic = "0" 
  lcl_update_displayHistoryOption   = "NULL"

  if dDate <> "" then
     lcl_update_eventdate = dDate
     lcl_update_eventdate = dbsafeWithHTML(lcl_update_eventdate)
     lcl_update_eventdate = "'" & lcl_update_eventdate & "'"
  end if

  if request("subject") <> "" then
  	lcl_update_subject = request("subject")
	  lcl_update_subject = dbsafeWithHTML(lcl_update_subject)
  	lcl_update_subject = Left(lcl_update_subject,50)
	  lcl_update_subject = "'" & lcl_update_subject & "'"
  end if

  if request("message") <> "" then
     lcl_update_message = request("message")
     lcl_update_message = dbsafeWithHTML(lcl_update_message)
     lcl_update_message = Left(lcl_update_message,1500)
     lcl_update_message = "'" & lcl_update_message & "'"
  end if

  if request("displayHistoryToPublic") <> "" then
     if request("displayHistoryToPublic") = "Y" then
        lcl_update_displayHistoryToPublic = "1"
     end if

     if request("displayHistoryOption") <> "" then
        lcl_update_displayHistoryOption = request("displayHistoryOption")
        lcl_update_displayHistoryOption = dbsafe(lcl_update_displayHistoryOption)
        lcl_update_displayHistoryOption = "'" & lcl_update_displayHistoryOption & "'"
     end if
  end if

  if lcl_calendarfeature <> "" then
     lcl_update_calendarfeature = lcl_calendarfeature
     lcl_update_calendarfeature = dbsafeWithHTML(lcl_update_calendarfeature)
     lcl_update_calendarfeature = "'" & lcl_update_calendarfeature & "'"
  end if

 'Update the Event
  sSQL = "UPDATE Events SET "
  sSQL = sSQL & " EventDate = "              & lcl_update_eventdate              & ", "
  sSQL = sSQL & " EventTimeZoneID = "        & request("timezone")               & ", "
  sSQL = sSQL & " EventDuration = "          & lDuration                         & ", "
  sSQL = sSQL & " CategoryID = "             & iCategoryID                       & ", "
  sSQL = sSQL & " Subject = "                & lcl_update_subject                & ", "
  sSQL = sSQL & " Message = "                & lcl_update_message                & ", "
  sSQL = sSQL & " ModifierUserID = "         & session("userid")                 & ", "
  sSQL = sSQL & " ModifiedDate = '"          & now()                             & "', "
  sSQL = sSQL & " calendarfeature = "        & lcl_update_calendarfeature        & ", "
  sSQL = sSQL & " isHiddenCL = "             & lcl_isHiddenCL                    & ", "
  sSQL = sSQL & " displayHistoryToPublic = " & lcl_update_displayHistoryToPublic & ", "
  sSQl = sSQL & " displayHistoryOption = "   & lcl_update_displayHistoryOption
  sSQL = sSQL & " WHERE eventid = " & lcl_eventid

  set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open sSQL, Application("DSN"), 3, 1

  set rs = nothing

  if lcl_calendarfeature_url <> "" then
     lcl_return_parameters = lcl_calendarfeature_url & "&id=" & lcl_eventid & "&success=SU"
  else
     lcl_return_parameters = "?id=" & lcl_eventid & "&success=SU"
  end if

 'Check to see if there is any aditional processing we will need to do.
  if lcl_orghasfeature_rssfeeds_events AND lcl_userhaspermission_rssfeeds_events AND request("sendTo_RSS") = "on" then
     lcl_return_parameters = lcl_return_parameters & "&sendTo_RSS=" & lcl_eventid
  end if

  response.redirect "updateevent.asp" & lcl_return_parameters

Else

 'Retreive the event data
  sSQL = "SELECT e.EventID, "
  sSQL = sSQL & " e.EventDate, "
  sSQL = sSQL & " e.EventTimeZoneID, "
  sSQL = sSQL & " t.TZAbbreviation, "
  sSQL = sSQL & " e.EventDuration, "
  sSQL = sSQL & " e.Subject, "
  sSQL = sSQL & " e.Message, "
  sSQL = sSQL & " e.CategoryID, "
  sSQL = sSQL & " e.calendarfeature, "
  sSQL = sSQL & " e.isHiddenCL, "
  sSQL = sSQL & " e.pushedfrom_requestid, "
  sSQL = sSQL & " e.CreatorUserID, "
  sSQL = sSQL & " e.CreateDate, "
  sSQL = sSQL & " e.ModifierUserID, "
  sSQL = sSQL & " e.ModifiedDate, "
  sSQL = sSQL & " e.displayHistoryToPublic, "
  sSQL = sSQL & " e.displayHistoryOption, "
  sSQL = sSQL & " (select u.firstname + ' ' + u.lastname from users u where u.userid = e.CreatorUserID) as createdbyname, "
  sSQL = sSQL & " (select u.firstname + ' ' + u.lastname from users u where u.userid = e.ModifierUserID) as lastupdatedbyname, "
  sSQL = sSQL & " (select a.[Tracking Number] from egov_rpt_actionline a where a.action_autoid = e.pushedfrom_requestid) as trackingnumber "
  sSQL = sSQL & " FROM events e "
  sSQL = sSQL &      " LEFT JOIN TimeZones t ON t.TimeZoneID = e.EventTimeZoneID "
  sSQL = sSQL & " WHERE eventid = " & lcl_eventid
  sSQL = sSQL & " AND orgid = " & session("orgid")

  set oRst = Server.CreateObject("ADODB.Recordset")
  oRst.Open sSQL, Application("DSN"), 3, 1

  if not oRst.eof then
    intID                      = CLng(oRst("EventID"))
    sDate                      = oRst("EventDate")
    iHour                      = Hour(sDate)
    iMinute                    = Minute(sDate)
    sAmPm                      = Right(sDate,2)
    sDate                      = FormatDateTime(sDate, vbShortDate)
    lDuration                  = oRst("EventDuration")
    iCategory                  = oRst("CategoryID")
    intTZ                      = oRst("EventTimeZoneID")
    lcl_subject                = escDblQuote(oRst("Subject"))
    lcl_details                = escDblQuote(oRst("Message"))
    sIsHiddenCL                = oRst("isHiddenCL")
    iCreatedByID               = oRst("CreatorUserID")
    iCreatedByName             = oRst("createdbyname")
    iCreatedDate               = oRst("CreateDate")
    iLastUpdatedByID           = oRst("ModifierUserID")
    iLastUpdatedByName         = oRst("lastupdatedbyname")
    iLastUpdatedDate           = oRst("ModifiedDate")
    iPushedFromRequestID       = oRst("pushedfrom_requestid")
    iPushedFromTrackingNum     = oRst("trackingnumber")
    lcl_displayHistoryToPublic = oRst("displayHistoryToPublic")
    lcl_displayHistoryOption   = oRst("displayHistoryOption")

    if iHour > 12 then
       iHour = iHour - 12
    end if

    if iHour = 0 then
       iHour = 12
    end if

    if lDuration Mod 10080 = 0 then
       sDurationInterval = "w"
       lDuration         = lDuration / 10080
    elseif lDuration Mod 1440 = 0 then
       sDurationInterval = "d"
       lDuration         = lDuration / 1440
    elseif lDuration Mod 60 = 0 then
       sDurationInterval = "h"
       lDuration         = lDuration / 60
    elseif lDuration >= 0 then
       sDurationInterval = "m"
    else
       lDuration = 0
    end if

    if sIsHiddenCL then
       lcl_checked = ""
    else
       lcl_checked = " checked=""checked"""
    end if

  End If

  set oRst = nothing

end if

'Setup the BODY onload
 lcl_onload = ""
 lcl_onload = lcl_onload & "setMaxLength();"
 lcl_onload = lcl_onload & "document.getElementById('DatePicker').focus();"

'Check for a screen message
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if

'Determine if there is any additional processing needed from the past update
 if lcl_orghasfeature_rssfeeds_events AND lcl_userhaspermission_rssfeeds_events AND lcl_success = "SU" then
    if request("sendTo_RSS") <> "" then
       lcl_onload = lcl_onload & "sendToRSS('" & request("sendTo_RSS") & "');"
    end if
 end if

	sTempEventID = request("id")
%>
<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	<title><%=langBSEvents%><%=lcl_calendar_name%></title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />	
	<link rel="stylesheet" type="text/css" href="eventstyles.css" />

	<script language="javascript" src="../scripts/ajaxLib.js"></script>
	<script language="javascript" src="../scripts/isvaliddate.js"></script>
	<script language="javascript" src="../scripts/textareamaxlength.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script type="text/javascript" src="../scripts/jquery-1.6.1.min.js"></script>

	<script language="javascript">
	<!--

		$(document).ready(function() 
		{
			$('#displayHistoryOption').prop('disabled','disabled');

			//Initialize fields
			if($('#displayHistoryToPublic').prop('checked')) 
			{
				$('#displayHistoryOption').prop('disabled','');
			}

			//Display History to Public: Click
			$('#displayHistoryToPublic').click(function() 
			{
				if($('#displayHistoryToPublic').prop('checked')) 
				{
					$('#displayHistoryOption').prop('disabled','');
				}
				else
				{
					$('#displayHistoryOption').prop('disabled','disabled');
				}
			});
		});

		function doCalendar( sField ) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=UpdateEvent", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function doPicker(sFormField, p_displayDocuments, p_displayActionLine, p_displayPayments, p_displayURL) {
			w = 600;
			h = 400;
			l = (screen.AvailWidth/2)-(w/2);
			t = (screen.AvailHeight/2)-(h/2);

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

		function fnCheckSubject() 
		{
			if (document.UpdateEvent.Subject.value != '') 
			{
				return true;
			}
			else
			{
				return false;
			}
		}

function validate() {
  var lcl_false_count = 0;
		var rege;
 	var Ok;

		if (document.getElementById("message").value.length > 1500) {
      document.getElementById("message").focus();
   		 inlineMsg(document.getElementById("message").id,'<strong>Invalid Value: </strong>Details cannot be longer than 1500 characters [current length: ' + document.getElementById("message").value.length + ']',10,'message');
      lcl_false_count = lcl_false_count + 1;
		}else{
      clearMsg("message");
  }

  if (document.getElementById("subject").value == "") {
      document.getElementById("subject").focus();
      inlineMsg(document.getElementById("subject").id,'<strong>Required Field Missing: </strong> Subject',10,'subject');
      lcl_false_count = lcl_false_count + 1;
  }else{
    		if (document.getElementById("subject").value.length > 50)	{
         	document.getElementById("subject").focus();
       		 inlineMsg(document.getElementById("subject").id,'<strong>Invalid Value: </strong>Subject cannot be longer than 50 characters [current length: ' + document.getElementById("ubject").value.length + ']',10,'message');
	 	       lcl_return_false = "Y";
      }else{
          clearMsg("subject");
      }
  }

  if(document.getElementById("Duration").value != "") {
 				rege = /^\d+$/;
     Ok = rege.test(document.getElementById("Duration").value);
     if (! Ok) {
         document.getElementById("Duration").focus();
         inlineMsg(document.getElementById("Duration").id,'<strong>Invalid Value: </strong> "Duration" must be a numeric value',10,'Duration');
         lcl_false_count = lcl_false_count + 1;
     }else{
         clearMsg("Duration");
     }
  }

		if (! isValidDate(document.getElementById("DatePicker").value)) {
      document.getElementById("DatePicker").focus();
      inlineMsg(document.getElementById("DatePicker").id,'<strong>Invalid Value: </strong> The "Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'DatePicker');
      lcl_false_count = lcl_false_count + 1;
		}else{
      clearMsg("DatePicker");
  }

  if(lcl_false_count > 0) {
     return false;
  }else{
  			document.getElementById("UpdateEvent").submit();
     return true;
		}
}

<% if lcl_orghasfeature_rssfeeds_events AND lcl_userhaspermission_rssfeeds_events then %>
function sendToRSS(pID) {
  var sParameter = 'id=' + encodeURIComponent(pID);
  sParameter    += '&isAjax=Y';

  doAjax('events_sendToRSS.asp', sParameter, 'displayScreenMsg', 'post', '0');
}
<% end if %>

function displayScreenMsg(iMsg) 
{
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "";
}

//-->
</script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<form name=""UpdateEvent"" id=""UpdateEvent"" method=""post"" action=""updateevent.asp"" accept-charset=""UTF-8"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""_task"" id=""_task"" value=""update"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""id"" id=""id"" value=""" & intID & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""Timezone"" id=""Timezone"" value=""1"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""cal"" id=""cal"" value=""" & lcl_calendarfeatureid & """ size=""20"" maxlength=""50"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""control_field"" id=""control_field"" value="""" size=""20"" maxlength=""4001"" />" & vbcrlf

  response.write "<div id=""content"">" & vbcrlf
  response.write "	 <div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <font size=""+1""><strong>Events: Update" & lcl_calendar_name & "</strong></font><br />" & vbcrlf
  response.write "          <input type=""button"" name=""backButton"" id=""backButton"" value=""<< Back to Event List"" class=""button"" onclick=""location.href='default.asp?useSessions=Y&cal=" & lcl_calendarfeatureid & "'"" />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"" valign=""top"">" & vbcrlf
  response.write "          <div class=""displayButtonsDIV"">" & vbcrlf
                              displayButtons
  response.write "        		</div>" & vbcrlf

  response.write "		        <table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"" class=""tableadmin"" id=""neweventinput"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <th align=""left"" colspan=""2"">" & langUpdateEvent & "</th>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Date
  response.write "            <tr>" & vbcrlf
  response.write "                <td>" & langDate & ":</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <input type=""text"" name=""DatePicker"" id=""DatePicker"" maxlength=""50"" value=""" & sDate & """ onchange=""clearMsg('DatePicker')"" />&nbsp;" & vbcrlf
  response.write "                    <a href=""javascript:void doCalendar('DatePicker');""><img src=""../images/calendar.gif"" border=""0"" /></a>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Time
  response.write "            <tr>" & vbcrlf
  response.write "                <td>" & langTime & ":</td>" & vbcrlf
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
  response.write "                    <select name=""DurationInterval"" id=""DurationInterval"" class=""time"" style=""width:80px;"">" & vbcrlf
                                        buildOption "DURATION", "1",     sDurationInterval
                                        buildOption "DURATION", "60",    sDurationInterval
                                        buildOption "DURATION", "1440",  sDurationInterval
                                        buildOption "DURATION", "10080", sDurationInterval
  response.write "                    </select>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Recurrence
  response.write "            <tr>" & vbcrlf
  response.write "                <td valign=""top"">Recurrence:</td>" & vbcrlf
  response.write "                <td><input type=""button"" value=""Choose Recurrence..."" class=""button"" onclick=""location.href='recurevent.asp?cal=" & lcl_calendarfeatureid & "&eventid=" & sTempEventID & "'""></td>" & vbcrlf
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
  response.write "                <td><input type=""text"" name=""Subject"" id=""subject"" size=""65"" maxlength=""50"" value=""" & lcl_subject & """ onchange=""clearMsg('subject')"" /></td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Details
  response.write "            <tr>" & vbcrlf
  response.write "                <td valign=""top"">" & langDetails & ":&nbsp;</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "                      <tr>" & vbcrlf
  response.write "                          <td width=""400"">" & vbcrlf
  response.write "                              <textarea name=""Message"" id=""message"" cols=""100"" rows=""12"" maxlength=""1500"">" & lcl_details & "</textarea>" & vbcrlf
  response.write "                          </td>" & vbcrlf
  response.write "                          <td align=""left"" valign=""top"">" & vbcrlf
  response.write "                              <input type=""button"" class=""button"" value=""Add Link"" onclick=""doPicker('UpdateEvent.message','Y','Y','Y','Y');"">" & vbcrlf
  response.write "                          </td>" & vbcrlf
  response.write "                      </tr>" & vbcrlf
  response.write "                    </table>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "		  	       <tr>" & vbcrlf
  response.write "                <td>&nbsp;</td>" & vbcrlf
  response.write "                <td><input type=""checkbox"" name=""isHiddenCL"" id=""isHiddenCL"" value=""on""" & lcl_checked & " />&nbsp;Show on Community Link</td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

  if lcl_orghasfeature_rssfeeds_events AND lcl_userhaspermission_rssfeeds_events then
     response.write "            <tr valign=""top"">" & vbcrlf
     response.write "                <td>On Update Send To:</td>" & vbcrlf
     response.write "                <td>" & vbcrlf
                                         displaySendToOption "RSS", "EDIT", "Y", _
                                                             lcl_orghasfeature_rssfeeds_events, _
                                                             lcl_userhaspermission_rssfeeds_events
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  end if

 'Display History Info
  response.write "            <tr>" & vbcrlf
  response.write "                <td colspan=""2"">" & vbcrlf
  response.write "                    <fieldset class=""fieldset"">" & vbcrlf
  response.write "                      <legend>History Log&nbsp;</legend>" & vbcrlf

  if iCreatedByID <> "" OR (iPushedFromRequestID <> "" AND lcl_orghasfeature_pushcontent AND lcl_userhaspermission_pushcontent) then
     response.write "                   <table border=""0"" cellspacing=""0"" cellpadding=""2"" style=""margin-top:5px;"">" & vbcrlf

    'Created By
     if iCreatedByID <> "" then
        lcl_createdby = iCreatedByName & " on " & iCreatedDate

        response.write "                     <tr>" & vbcrlf
        response.write "                         <td><strong>Created By:</strong></td>" & vbcrlf
        response.write "                         <td style=""color:#800000"">" & lcl_createdby & "</td>" & vbcrlf
        response.write "                     </tr>" & vbcrlf
     end if

    'Last Updated By
     if iLastUpdatedByID <> "" then
        lcl_lastupdatedby = iLastUpdatedByName & " on " & iLastUpdatedDate

        response.write "                     <tr>" & vbcrlf
        response.write "                         <td><strong>Last Updated By:</strong></td>" & vbcrlf
        response.write "                         <td style=""color:#800000"">" & lcl_lastupdatedby & "</td>" & vbcrlf
        response.write "                     </tr>" & vbcrlf
     end if

    'Originated From
     if iPushedFromRequestID <> "" AND lcl_orghasfeature_pushcontent AND lcl_userhaspermission_pushcontent then
        response.write "                     <tr>" & vbcrlf
        response.write "                         <td><strong>Originated From:</strong></td>" & vbcrlf
        response.write "                         <td><a href=""../action_line/action_respond.asp?control=" & iPushedFromRequestID & """>" & iPushedFromTrackingNum & "</a></td>" & vbcrlf
        response.write "                     </tr>" & vbcrlf
     end if

     response.write "                   </table>" & vbcrlf
  end if

 'Display History Log Options
  if lcl_orghasfeature_displayHistoryInfo then
     lcl_checked_displayHistoryToPublic      = ""
     lcl_selected_displayHistoryOption_all   = ""
     lcl_selected_displayHistoryOption_names = ""
     lcl_selected_displayHistoryOption_dates = ""

     if lcl_displayHistoryToPublic then
        lcl_checked_displayHistoryToPublic = " checked=""checked"""
     end if

     if lcl_displayHistoryOption = "Names Only" then
        lcl_selected_dho_names = " selected=""selected"""
     elseif lcl_displayHistoryOption = "Date/Time Only" then
        lcl_selected_dho_dates = " selected=""selected"""
     else
        lcl_selected_dho_all = " selected"
     end if

     response.write "                      <p>" & vbcrlf
     response.write "                        <input type=""checkbox"" name=""displayHistoryToPublic"" id=""displayHistoryToPublic"" value=""Y""" & lcl_checked_displayHistoryToPublic & " />" & vbcrlf
     response.write "                        Display history on public calendar" & vbcrlf
     response.write "                        <select name=""displayHistoryOption"" id=""displayHistoryOption"">" & vbcrlf
     response.write "                          <option value="""""               & lcl_selected_dho_all   & ">All Info</option>" & vbcrlf
     response.write "                          <option value=""Names Only"""     & lcl_selected_dho_names & ">Names Only</option>" & vbcrlf
     response.write "                          <option value=""Date/Time Only""" & lcl_selected_dho_dates & ">Date/Time Only</option>" & vbcrlf
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
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
Function escDblQuote( ByVal strDB )

	If VarType( strDB ) = vbString Then 
		strDB = Replace( strDB, Chr(34), "&quot;" )
	End If 

	escDblQuote = strDB

End Function

'------------------------------------------------------------------------------
sub displayButtons()
  response.write "<input type=""button"" value=""" & langCancel & """ class=""button"" onClick=""history.back();"" />" & vbcrlf
  response.write "<input type=""button"" value=""Save Changes"" class=""button"" onClick=""return validate();"" />" & vbcrlf
end sub
%>
