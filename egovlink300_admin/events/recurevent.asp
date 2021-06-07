<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="events_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: recurevent.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the calendar
'
' MODIFICATION HISTORY
' 1.0   ???			 ???? - INITIAL VERSION
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
' 1.2	12/05/07	Steve Loar - Added CheckRecur() function
' 1.3 08/06/08 David Boyer - Added Custom Calendar
' 1.4 06/11/09 David Boyer - Added check for "Show on CommunityLink"
' 1.5 05/27/10 David Boyer - Added validation on "number" fields
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
if isFeatureOffline("calendar") = "Y" OR isFeatureOffline("custom_calendars") = "Y" then
	response.redirect "../admin/outage_feature_offline.asp"
end if

sLevel = "../"  'Override of value from common.asp

'Check to see if this is a Custom Calendar
' lcl_calendarfeature     = trim(request("cal"))
' lcl_calendarfeature_url = ""
' lcl_calendar_name       = ""

' if trim(lcl_calendarfeature) <> "" then
'    if OrgHasFeature(trim(lcl_calendarfeature)) AND UserHasPermission(session("userid"), trim(lcl_calendarfeature)) then
'       lcl_calendarfeature_url = "?cal=" & lcl_calendarfeature
'       lcl_calendar_name       = " [" & GetFeatureName(lcl_calendarfeature) & "]"
'    else
'      	response.redirect sLevel & "permissiondenied.asp"
'    end if
' else
'    if NOT UserHasPermission( Session("UserId"), "edit events" ) then
'      	response.redirect sLevel & "permissiondenied.asp"
'    end if
' end if

lcl_calendarfeatureid   = ""
lcl_calendarfeature     = "?useSessions=Y"
lcl_calendarfeature_url = ""
lcl_calendar_name       = ""

if trim(request("cal")) <> "" then
	if not isnumeric(trim(request("cal"))) then
		response.redirect sLevel & "permissiondenied.asp"
	else
		lcl_calendarfeatureid = CLng(trim(request("cal")))
		lcl_calendarfeature   = getFeatureByID(session("orgid"), lcl_calendarfeatureid)

		if OrgHasFeature(lcl_calendarfeature) AND UserHasPermission(session("userid"), lcl_calendarfeature) then
			lcl_calendarfeature_url  = "?cal=" & lcl_calendarfeatureid & "&useSessions=Y"
			lcl_calendarfeature_name = " [" & getFeatureName(lcl_calendarfeature) & "]"
		else
			response.redirect sLevel & "permissiondenied.asp"
		end if
	end if
else
	if NOT userhaspermission( session("userid"), "edit events" ) then
		response.redirect sLevel & "permissiondenied.asp"
	end if
end if

Dim oCmd, oRst, dDate, iDuration, sTimeZones, sLinks, bShown, dNext

'If Not HasPermission("CanEditEvents") Then Response.Redirect "../default.asp"

if Request.Form("_task") = "newevent" then
	' create a new event if that is the task
	dDate     = CDate(Request("DatePicker") & " " & Request("Hour") & ":" & Request("Minute") & " " & Request("AMPM"))
	iDuration = Request.Form("Duration")

	if iDuration & "" <> "" then
		iDuration = CLng(iDuration) * clng(Request("DurationInterval"))
	else
		iDuration = -1
	end if

	if Request.Form("CustomCategory") <> "" then
		'Create a New Category for this Organization
		'newCategory session("orgid"), request.form("CustomCategory"), "#000000", session("calendarfeature"), lcl_identity
		newCategory session("orgid"), request.form("CustomCategory"), "#000000", lcl_calendarfeature, lcl_identity

		iCategoryID = lcl_identity

	else
		iCategoryID = Request("Category")
	end if

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "NewEvent"
		.CommandType = adCmdStoredProc
		.Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
		.Parameters.Append oCmd.CreateParameter("CreatorID", adInteger, adParamInput, 4, Session("UserID"))
		.Parameters.Append oCmd.CreateParameter("EventDate", adDateTime, adParamInput, 4, dDate)
		.Parameters.Append oCmd.CreateParameter("TimeZone", adInteger, adParamInput, 4, Request.Form("TimeZone"))
		.Parameters.Append oCmd.CreateParameter("Duration", adInteger, adParamInput, 4, iDuration)
		.Parameters.Append oCmd.CreateParameter("CategoryID", adInteger, adParamInput, 4, iCategoryID)
		.Parameters.Append oCmd.CreateParameter("Subject", adVarChar, adParamInput, 50, Request.Form("Subject"))
		.Parameters.Append oCmd.CreateParameter("Message", adVarChar, adParamInput, 5000, Request.Form("Message"))
		.Parameters.Append oCmd.CreateParameter("calendarfeature", adVarChar, adParamInput, 50, lcl_calendarfeature)
		.Execute
	End With
	Set oCmd = Nothing

	response.redirect "../events/default.asp" & lcl_calendarfeature_url

Else 
	' does this do anthing meaningful??
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "ListTimeZones"
		.CommandType = adCmdStoredProc
		.Execute
	End With

	Set oRst = Server.CreateObject("ADODB.Recordset")
	With oRst
		.CursorLocation = adUseClient
		.CursorType     = adOpenStatic
		.LockType       = adLockReadOnly
		.Open oCmd
	End With
	Set oCmd = Nothing

	do while not oRst.eof
		if oRst("TimeZoneID") = 1 then
			lcl_selected_timezone = " selected=""selected"""
		else
			lcl_selected_timezone = ""
		end if

		sTimeZones = sTimeZones & "<option value=""" & oRst("TimeZoneID") & """" & lcl_selected_timezone & ">" & oRst("TZName") & "</option>" & vbcrlf

		oRst.movenext
	Loop 

	if oRst.State=1 then oRst.Close
	set oRst = nothing

End If 




'Get the eventid and verify it is in proper format
EventID = 0

if request("eventid") <> "" then
	EventID = CLng(request("eventid"))
end if

'Retrieve querystring parameters
lcl_recur = "dd"


' This seems to be where processing the repeating dates happens
if request("Recur") <> "" then
	lcl_recur = request("Recur")
end if

'Set up the recurring records.
If lcl_recur <> "" Then 

	'EventID = request("EventID")

	if request("isHiddenCL") = "on" then
		lcl_isHiddenCL = 0
	else
		lcl_isHiddenCL = 1
	end if

	'Here is the recurrence problem - Does not use original event date
	'dNext=Date()
	dNext = getEventDate( EventID ) ' Get the event date as a string
	dNext = DateSerial(Year(dNext),Month(dNext),Day(dNext)) ' convert it to a real date

	dStart   = dNext
	tOften   = Request("Often")
	tHowLong = Request("HowLong")
	iHowMany = Request("HowMany")

	'Weekly
	iWeeks   = Request("Weeks")
	iDays    = Request("WeekDayNum")

	'Monthly
	iDay     = Request("Day")
	iMonth   = Request("Month")
	iMonths  = Request("Months")

	' if there is an end by date, then set that here
	if IsDate(Request("DatePicker")) then
		dEnd = CDate(Request("DatePicker"))
	else
		dEnd = dStart
	end if

	dDates = ""

	Select Case lcl_recur
		Case "dd"
			sDD = " selected=""selected"""

			Select Case Request("Often")
				case "days"
					if tHowLong = "endby" then
						do while DateAdd("d",Request("Days"),dNext) < dEnd
							dNext = DateAdd("d",Request("Days"),dNext)
							if dNext > dStart then dDates = dNext & "," & dDates
						loop
					elseif tHowlong = "till" then
						for Idx = 1 to iHowMany
							dNext = DateAdd("d",Request("Days"),dNext)
							if dNext > dStart then dDates = dNext & "," & dDates
						next
					end if

					RecurEvent EventID, dDates, lcl_isHiddenCL
					response.redirect "../events/default.asp" & lcl_calendarfeature_url

				case "weekdays"
					if tHowLong = "till" then
						for Idx = 1 to iHowMany
							dNext = AddWeekDays(1,dNext)
							if dNext > dStart then dDates = dNext & "," & dDates
						next
					elseif tHowlong = "endby" then
						do while AddWeekDays(1,dNext) <= dEnd
							dNext = AddWeekDays(1,dNext)
							if dNext > dStart then dDates = dNext & "," & dDates
						loop
					end if

					RecurEvent EventID, dDates, lcl_isHiddenCL
					response.redirect "../events/default.asp" & lcl_calendarfeature_url

				case else
			End Select

		Case "ww"
			sWW = " selected=""selected"""
			If iWeeks <> "" AND Request.Form("WeekDayNum") <> "" Then
				dStart = dNext
				Select Case Request.Form("HowLong")
					Case "till"
						iDays = Split(Request.Form("WeekDayNum"),",")
						For Each iDay In iDays
							dNext = dStart
							dNext = GetNextWeekDay(clng(iDay),dNext)
							For Idx = 1 To iHowMany
								dNext = DateAdd("ww",iWeeks,dNext)
								If dNext > dStart Then dDates = dNext & "," & dDates
							Next
						Next
					Case "endby"
						iDays = Split(Request.Form("WeekDayNum"),",")
						For Each iDay In iDays
							dNext = dStart
							dNext = GetNextWeekDay(clng(iDay),dNext)
							Do While DateAdd("ww",iWeeks,dNext) < dEnd
								dNext = DateAdd("ww",iWeeks,dNext)
								If dNext > dStart Then dDates = dNext & "," & dDates
							Loop
						Next
				End Select

				RecurEvent EventID, dDates, lcl_isHiddenCL
				response.redirect "../events/default.asp" & lcl_calendarfeature_url
			End If

		Case "mm"
			sMM = " selected=""selected"""
			If iMonth <> "" OR iDay <> "" Then
				Select Case tHowLong
					Case "till"
						If tOften = "absolute" Then
							dNext = NextDayOfMonth(clng(iDay),dNext)
							For Idx = 1 To clng(iHowMany)
								dNext = NextDayOfMonth(clng(iDay),dNext)
								' Add DB record
								If dNext > dStart Then dDates = dNext & "," & dDates
								dNext = DateAdd("m",clng(iMonths),dNext)
							Next
						ElseIf tOften = "relative" Then
							For Idx = 1 To clng(iHowMany)
								dNext = OrdinalDate(Request.Form("DayLike"),Request.Form("Ordinal"),dNext)
								If dNext > dStart Then dDates = dNext & "," & dDates
								dNext = DateAdd("m",clng(iMonth),dNext)
							Next
						End If
					Case "endby"
						If tOften = "absolute" Then
							'dNext = NextDayOfMonth(clng(iDay),dNext)
							Do While NextDayOfMonth(clng(iDay),dNext) < dEnd
								dNext = NextDayOfMonth(clng(iDay),dNext)
								If dNext > dStart Then dDates = dNext & "," & dDates
								dNext = DateAdd("m",iMonths,dNext)
							Loop
						ElseIf tOften = "relative" Then
							Do While OrdinalDate(Request.Form("DayLike"),Request.Form("Ordinal"),dNext) < dEnd
								dNext = OrdinalDate(Request.Form("DayLike"),Request.Form("Ordinal"),dNext)
								If dNext > dStart Then dDates = dNext & "," & dDates
								dNext = DateAdd("m",iMonth,dNext)
							Loop
						End If
				End Select

				RecurEvent EventID, dDates, lcl_isHiddenCL
				response.redirect "../events/default.asp" & lcl_calendarfeature_url
			End If

		Case "yy"
			sYY = " selected=""selected"""
			If (iMonth <> "" AND iDay <> "") OR (iMonths <> "" AND Request.Form("DayNum") <> "") OR (Request.Form("Ordinal") <> "" AND Request.Form("DayPick") <> "" AND Request.Form("MonthPick")) Then
				Select Case tHowLong
					Case "till"
						Select Case tOften
							Case "every"
								dNext = NextDate(clng(iMonth),clng(iDay),dNext)
								For Idx = 1 To clng(iHowMany)
									If dNext > dStart Then dDates = dNext & "," & dDates
									dNext = DateAdd("yyyy",1,dNext)
								Next
							Case "absolute"
								dNext = NextDayOfMonth(clng(Request.Form("DayNum")),dNext)
								Do While DatePart("yyyy",NextDayOfMonth(clng(Request.Form("DayNum")),dNext))-DatePart("yyyy",dStart) <= clng(iHowMany)
									dNext = NextDayOfMonth(clng(Request.Form("DayNum")),dNext)
									If dNext > dStart Then dDates = dNext & "," & dDates
									dNext = DateAdd("m",clng(iMonths),dNext)
								Loop
							Case "relative"
								dNext = NextDate(Request.Form("MonthPick"),1,dNext)

								For Idx = 1 To clng(iHowMany)
									dNext = OrdinalDate(Request.Form("DayPick"),Request.Form("Ordinal"),dNext)
									If dNext > dStart Then dDates = dNext & "," & dDates
									dNext = DateAdd("yyyy",1,dNext)
								Next
						End Select
					Case "endby"
						Select Case tOften
							Case "every"
								dNext = NextDate(clng(iMonth),clng(iDay),dNext)
								Do While dNext < dEnd
									If dNext > dStart Then dDates = dNext & "," & dDates
										dNext = DateAdd("yyyy",1,dNext)
								Loop
							Case "absolute"
								Do While NextDayOfMonth(clng(Request.Form("DayNum")),dNext) < dEnd
									dNext = NextDayOfMonth(clng(Request.Form("DayNum")),dNext)
									If dNext > dStart Then dDates = dNext & "," & dDates
										dNext = DateAdd("m",clng(iMonths),dNext)
								Loop
							Case "relative"
								dNext = NextDate(Request.Form("MonthPick"),1,dNext)

								Do While OrdinalDate(Request.Form("DayPick"),Request.Form("Ordinal"),dNext) < dEnd
									dNext = OrdinalDate(Request.Form("DayPick"),Request.Form("Ordinal"),dNext)
									If dNext > dStart Then dDates = dNext & "," & dDates
									dNext = DateAdd("yyyy",1,dNext)
								Loop
						End Select
				End Select

 				RecurEvent EventID, dDates, lcl_isHiddenCL
				response.redirect "../events/default.asp" & lcl_calendarfeature_url

			End If
	End Select
End If


%>
<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	<title><%=langBSEVents%><%=lcl_calendar_name%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />	
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="eventstyles.css" />

	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

  <script language="Javascript">
  <!--
		function doCalendar( sField ) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=RecurEvent", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

     function storeCaret (textEl) 
	 {
       if (textEl.createTextRange)
         textEl.caretPos = document.selection.createRange().duplicate();
     }

     function insertAtCaret (textEl, text) 
	 {
       if (textEl.createTextRange && textEl.caretPos) {
         var caretPos = textEl.caretPos;
         caretPos.text =
           caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
             text + ' ' : text;
       }
       else
         textEl.value  = text;
     }

    function doPicker(sFormField) 
	{
      w = (screen.width - 350)/2;
      h = (screen.height - 350)/2;
      eval('window.open("../picker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
    }

	function fnCheckSubject()
	{

		if (document.NewEvent.Subject.value != '') {
			return true;
		}
		else
		{
			return false;
		}
	}

 function CheckRecur() {
   var total_options   = 0
   var lcl_false_count = 0;
   var lcl_recur       = '<%=lcl_recur%>';
   var lcl_focus       = "";
   var lcl_value       = ""
 		var rege;
	 	var Ok;

   //Get the total number of options in the "Range of recurrence" radio list
   total_options = document.RecurEvent.HowLong.length;

   //find the value "checked"
   for (i = 0; i <= total_options-1; i++){
      if(document.RecurEvent.HowLong[i].checked == true) {
         lcl_value = document.RecurEvent.HowLong[i].value;
      }
   }

 		//if (document.RecurEvent.HowLong[0].checked) {
   if (lcl_value == "till") {
       if(document.getElementById("HowMany").value == "") {
          lcl_focus = "HowMany";
          inlineMsg(document.getElementById("HowMany").id,'<strong>Required Field Missing: </strong> Occurrences',10,'HowMany');
          lcl_false_count = lcl_false_count + 1;
       }else{
      				rege = /^\d+$/;
      				Ok = rege.test(document.getElementById("HowMany").value);
      				if (! Ok) {
             lcl_focus = "HowMany";
             inlineMsg(document.getElementById("HowMany").id,'<strong>Invalid Value: </strong> "Occurrences" must be a numeric value',10,'HowMany');
             lcl_false_count = lcl_false_count + 1;
          }else{
             clearMsg("HowMany");
          }
       }
		 }else{  //lcl_value = "endby"
    			if (document.getElementById("DatePicker").value == "") {
          lcl_focus = "DatePicker";
          inlineMsg(document.getElementById("DatePickerPopup").id,'<strong>Required Field Missing: </strong> Date',10,'DatePickerPopup');
          lcl_false_count = lcl_false_count + 1;
       }else{
      				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
      				Ok   = rege.test(document.getElementById("DatePicker").value);

      				if (! Ok) {
             lcl_focus = "DatePicker";
             inlineMsg(document.getElementById("DatePickerPopup").id,'<strong>Invalid Value: </strong> The "End Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'DatePickerPopup');
             lcl_false_count = lcl_false_count + 1;
          }else{
             clearMsg("DatePickerPopup");
          }
       }
		 }

   if(lcl_recur == "dd") {
      //Get the total number of options in the drop down list
      total_options = document.RecurEvent.Often.length;
      lcl_value     = ""

      //find the value "checked"
      for (i = 0; i <= total_options-1; i++){
           if(document.RecurEvent.Often[i].checked == true) {
              lcl_value = document.RecurEvent.Often[i].value;
           }
      }

      //If the option selected requires a number (days) to be entered then:
      //1. ensure that the field is not null
      //2. only allow numbers
      if(lcl_value == "days") {
         if(document.getElementById("Days").value == "") {
            lcl_focus = "Days";
            inlineMsg(document.getElementById("Days").id,'<strong>Required Field Missing: </strong> Days',10,'Days');
            lcl_false_count = lcl_false_count + 1;
         }else{
        				rege = /^\d+$/;
        				Ok = rege.test(document.getElementById("Days").value);

        				if (! Ok) {
               lcl_focus = "Days";
               inlineMsg(document.getElementById("Days").id,'<strong>Invalid Value: </strong> "Days" must be a numeric value',10,'Days');
               lcl_false_count = lcl_false_count + 1;
            }else{
               clearMsg("Days");
            }
         }
      }

   } else if(lcl_recur == "ww") {

      //Get the total number of options in the drop down list
      total_options = document.RecurEvent.WeekDayNum.length;
      lcl_value     = ""

      //Determine if at least one value is "checked"
      for (i = 0; i <= total_options-1; i++){
           if(document.RecurEvent.WeekDayNum[i].checked == true) {
              lcl_value = "Y";
           }
      }

      //No day(s) have been selected
      if(lcl_value == "") {
         lcl_focus = "Wednesday";
         inlineMsg(document.getElementById("Wednesday").id,'<strong>Required Field Missing: </strong> At least one day must be selected',10,'Wednesday');
         lcl_false_count = lcl_false_count + 1;
      }else{
         clearMsg("Wednesday");
      }

      //Validate the Number of Weekly Occurrences
      if(document.getElementById("Weeks").value == "") {
         lcl_focus = "Weeks";
         inlineMsg(document.getElementById("Weeks").id,'<strong>Required Field Missing: </strong> Occurrences',10,'Weeks');
         lcl_false_count = lcl_false_count + 1;
      }else{
     				rege = /^\d+$/;
     				Ok = rege.test(document.getElementById("Weeks").value);

     				if (! Ok) {
             lcl_focus = "Weeks";
             inlineMsg(document.getElementById("Weeks").id,'<strong>Invalid Value: </strong> Occurrences must be a numeric value',10,'Weeks');
             lcl_false_count = lcl_false_count + 1;
         }else{
             clearMsg("Weeks");
         }
      }
   } else if(lcl_recur == "mm") {
      //Get the total number of options in the drop down list
      total_options = document.RecurEvent.Often.length;
      lcl_value     = ""

      //find the value "checked"
      for (i = 0; i <= total_options-1; i++){
           if(document.RecurEvent.Often[i].checked == true) {
              lcl_value = document.RecurEvent.Often[i].value;
           }
      }

      if(lcl_value == "absolute") {
         //validate the 2nd input box (Months)
         if(document.getElementById("Months").value == "") {
            lcl_focus = "Months";
            inlineMsg(document.getElementById("Months").id,'<strong>Required Field Missing: </strong> Months',10,'Months');
            lcl_false_count = lcl_false_count + 1;
         }else{
        				rege = /^\d+$/;
        				Ok = rege.test(document.getElementById("Months").value);

        				if (! Ok) {
               lcl_focus = "Months";
               inlineMsg(document.getElementById("Months").id,'<strong>Invalid Value: </strong> "Months" must be a numeric value',10,'Months');
               lcl_false_count = lcl_false_count + 1;
            }else{
               clearMsg("Months");
            }
         }

         //validate the 1st input box (Day)
         if(document.getElementById("Day").value == "") {
            lcl_focus = "Day";
            inlineMsg(document.getElementById("Day").id,'<strong>Required Field Missing: </strong> Days',10,'Day');
            lcl_false_count = lcl_false_count + 1;
         }else{
        				rege = /^\d+$/;
        				Ok = rege.test(document.getElementById("Day").value);

        				if (! Ok) {
               lcl_focus = "Day";
               inlineMsg(document.getElementById("Day").id,'<strong>Invalid Value: </strong> "Day" must be a numeric value',10,'Day');
               lcl_false_count = lcl_false_count + 1;
            }else{
               clearMsg("Day");
            }
         }
      }else{  //lcl_value == "relative"
         //validate the 2nd input box (Months)
         if(document.getElementById("Month").value == "") {
            lcl_focus = "Month";
            inlineMsg(document.getElementById("Month").id,'<strong>Required Field Missing: </strong> Months',10,'Month');
            lcl_false_count = lcl_false_count + 1;
         }else{
        				rege = /^\d+$/;
        				Ok = rege.test(document.getElementById("Month").value);

        				if (! Ok) {
               lcl_focus = "Month";
               inlineMsg(document.getElementById("Month").id,'<strong>Invalid Value: </strong> "Months" must be a numeric value',10,'Month');
               lcl_false_count = lcl_false_count + 1;
            }else{
               clearMsg("Month");
            }
         }
      }
   } else if(lcl_recur == "yy") {

      //Get the total number of options in the drop down list
      total_options = document.RecurEvent.Often.length;
      lcl_value     = ""

      //find the value "checked"
      for (i = 0; i <= total_options-1; i++){
           if(document.RecurEvent.Often[i].checked == true) {
              lcl_value = document.RecurEvent.Often[i].value;
           }
      }

      //If the option selected requires a number (days) to be entered then:
      //1. ensure that the field is not null
      //2. only allow numbers
      if(lcl_value == "every") {
         if(document.getElementById("Day").value == "") {
            lcl_focus = "Day";
            inlineMsg(document.getElementById("Day").id,'<strong>Required Field Missing: </strong> Day',10,'Day');
            lcl_false_count = lcl_false_count + 1;
         }else{
        				rege = /^\d+$/;
        				Ok = rege.test(document.getElementById("Day").value);

        				if (! Ok) {
               lcl_focus = "Day";
               inlineMsg(document.getElementById("Day").id,'<strong>Invalid Value: </strong> "Day" must be a numeric value',10,'Day');
               lcl_false_count = lcl_false_count + 1;
            }else{
               clearMsg("Day");
            }
         }
      }else if(lcl_value == "absolute") {

         if(document.getElementById("Months").value == "") {
            lcl_focus = "Months";
            inlineMsg(document.getElementById("Months").id,'<strong>Required Field Missing: </strong> Months',10,'Months');
            lcl_false_count = lcl_false_count + 1;
         }else{
        				rege = /^\d+$/;
        				Ok = rege.test(document.getElementById("Months").value);

        				if (! Ok) {
               lcl_focus = "Months";
               inlineMsg(document.getElementById("Months").id,'<strong>Invalid Value: </strong> "Months" must be a numeric value',10,'Months');
               lcl_false_count = lcl_false_count + 1;
            }else{
               clearMsg("Months");
            }
         }

         if(document.getElementById("DayNum").value == "") {
            lcl_focus = "DayNum";
            inlineMsg(document.getElementById("DayNum").id,'<strong>Required Field Missing: </strong> Day',10,'DayNum');
            lcl_false_count = lcl_false_count + 1;
         }else{
        				rege = /^\d+$/;
        				Ok = rege.test(document.getElementById("DayNum").value);

        				if (! Ok) {
               lcl_focus = "DayNum";
               inlineMsg(document.getElementById("DayNum").id,'<strong>Invalid Value: </strong> "Day" must be a numeric value',10,'DayNum');
               lcl_false_count = lcl_false_count + 1;
            }else{
               clearMsg("DayNum");
            }
         }
      }
   }

   if(lcl_false_count > 0) {
      document.getElementById(lcl_focus).focus();
      return false;
  	}else{
    		document.RecurEvent.submit();
   }
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

  //-->
  </script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<div id="content">
	<div id="centercontent">

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <!--<td width="151" align="center"><img src="../images/icon_home.jpg"></td>-->
      <td>
          <font size="+1"><strong><%=langEvents%>: Recur Event <%=lcl_calendar_name%></strong></font><br />

          <input type="button" name="returnButton" id="returnButton" value="<< Back to Event List" class="button" onclick="location.href='default.asp<%=lcl_calendarfeature_url%>';" />
      </td>
      <td width="200" align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
    </tr>
    <tr>
      <td colspan="2" valign="top">
        <form name="RecurEvent" method="post" action="recurevent.asp" method="post">
          <input type="hidden" name="_task" value="recurevent" />
          <input type="hidden" name="EventID" value="<%=Request("EventID")%>" />
          <input type="hidden" name="dNext" value="<%=dNext%>" />
          <input type="hidden" name="cal" value="<%=lcl_calendarfeatureid%>" />

          <div class="displayButtonsDIV">
            <% displayButtons eventid, lcl_calendarfeature_url %>
          </div>

          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tableadmin" id="neweventinput">
            <tr>
                <th align="left" colspan="2">New Recurrence</th>
            </tr>
            <tr>
                <td valign="top">Recurrence:</td>
                <td width="100%">
                    <select name="Recur" class="time" onchange="document.RecurEvent.submit();">
                      <option value="dd"<%=sDD%>>Daily</option>
                      <option value="ww"<%=sWW%>>Weekly</option>
                      <option value="mm"<%=sMM%>>Monthly</option>
                      <option value="yy"<%=sYY%>>Yearly</option>
                    </select>
                </td>
            </tr>
            <% displayRecurOptions lcl_recur %>
            <tr>
                <td valign="top" nowrap>Range of Recurrence:</td>
                <td width="100%">
                    <!--Start <input type="text" name="DatePicker" style="width:133px;" maxlength="50" value="<%= Date() %>">&nbsp;<a href="javascript:void doCalendar();">Choose</a>-->
                </td>
            </tr>
            <tr>
                <td valign="top" nowrap></td>
                <td width="100%">
                    <input type="radio" name="HowLong" value="till" checked="checked" onclick="clearMsg('DatePickerPopup');" /> End after 
                    <input type="text" name="HowMany" id="HowMany" style="width:25px;" maxlength="4" value="1" onchange="clearMsg('HowMany');" /> occurrences.
                </td>
            </tr>
            <tr>
                <td valign="top" nowrap></td>
                <td width="100%">
                    <input type="radio" name="HowLong" value="endby" onclick="clearMsg('HowMany');" /> End by <input type="text" name="DatePicker" id="DatePicker" style="width:133px;" maxlength="50" value="<%=Date()%>" onchange="clearMsg('DatePickerPopup');" >
                    <a href="javascript:void doCalendar('DatePicker');"><img src="../images/calendar.gif" border="0" name="DatePickerPopup" id="DatePickerPopup" onclick="clearMsg('DatePickerPopup');" /></a>
                </td>
            </tr>
            <tr>
                <td colspan="2">&nbsp;</td>
            </tr>
            <tr>
                <td>&nbsp;</td>
                <td width="100%">
                    <input type="checkbox" name="isHiddenCL" id="isHiddenCL" value="on" />&nbsp;Show on CommunityLink
                </td>
            </tr>
          </table>

          <div class="displayButtonsDIV">
            <% displayButtons eventid, lcl_calendarfeature_url %>
          </div>
        </form>
      </td>
        <!-- END: NEW EVENT -->
    </tr>
  </table>

  	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%

'------------------------------------------------------------------------------
Function getEventDate( ByVal iEventID )
	Dim sSql, oEvent

	sSql = "SELECT eventdate FROM events WHERE eventid = " & iEventID 

	Set oEvent = Server.CreateObject("ADODB.Recordset")
	oEvent.Open sSql, Application("DSN"), 3, 1

	getEventDate = oEvent("eventdate")

	oEvent.Close
	Set oEvent = Nothing

End Function


'------------------------------------------------------------------------------
Function RecurEvent( ByVal EventID, ByVal dDates, ByVal p_isHiddenCL )
	Dim sSql, oRst, lcl_return

	lcl_return = 0
	dDates = Split(dDates,",")

	sSql = "SELECT e.EventID, e.EventDate, e.EventTimeZoneID, t.TZAbbreviation, e.EventDuration, e.Subject, e.Message, "
	sSql = sSql & " e.CategoryID, e.CalendarFeature "
	sSql = sSql & " FROM Events AS e "
	sSql = sSql & " LEFT JOIN TimeZones AS t ON t.TimeZoneID = e.EventTimeZoneID "
	sSql = sSql & " WHERE EventID = " & EventID

	Set oRst = Server.CreateObject("ADODB.Recordset")
	oRst.Open sSql, Application("DSN"), 3, 1

	If Not oRst.EOF Then 
		For Each dDate In dDates 
			If dDate <> "" Then 
				If DateDiff("d",dDate,oRst("EventDate")) <> 0 Then 
					dDate = dDate & " " & FormatDateTime(oRst("EventDate"),vbLongTime)

					'Create the New Recurring Event
					newRecurEvent session("orgid"), eventid, dDate, session("userid"), oRst("EventTimeZoneID"), oRst("EventDuration"), oRst("Subject"), oRst("Message"), oRst("CategoryID"), oRst("CalendarFeature"), p_isHiddenCL, lcl_identity
				End If 
			End If 
		Next 

		lcl_return = 1

	End If 

	oRst.Close
	Set oRst = Nothing 

	RecurEvent = lcl_return

End Function 


'------------------------------------------------------------------------------
Function AddWeekDays( ByVal nDaystoAdd, ByVal dtStartDate )
	'Adds working days based on a five day week

	Dim dtEndDate, iLoop

	'First add whole weeks
	dtEndDate = DateAdd("ww",Int(nDaysToAdd/5),dtStartDate)

	'Add any odd days
	For iLoop = 1 To (nDaysToAdd Mod 5)
		dtEndDate = DateAdd("d",1,dtEndDate)
		'If Saturday increment to following Monday
		If WeekDay(dtEndDate) = vbSaturday Then
			'Increment date to following Monday
			dtEndDate = DateAdd("d",2,dtEndDate)
		End If
	Next

	AddWeekDays=dtEndDate

End Function


'------------------------------------------------------------------------------
Function GetNextWeekDay( ByVal iWeekDay, ByVal dTemp)
	'Increments a date until the next specified weekday
	Dim Idx

	'Add any odd days
	Do While Not DatePart("w",dTemp) = iWeekDay
		dTemp = DateAdd("d",1,dTemp)
	Loop

	GetNextWeekDay = dTemp

End Function


'------------------------------------------------------------------------------
Function GetAWeekday( ByVal iNumDays, ByVal iWeekDay, ByVal dStart )
	'Increments a date until the next specified weekday
	Dim Idx, dEnd

	dEnd = dStart

	For Idx = 1 To iNumDays
		'Add any odd days
		Do While Not DatePart("w",dEnd) = iWeekDay
			dEnd = DateAdd("d",1,dEnd)
		Loop

		If Idx <> iNumDays Then
			dEnd = DateAdd("d",1,dEnd)
		End If
	Next

	GetAWeekday = dEnd

End Function


'------------------------------------------------------------------------------
Function GetLastWeekday( ByVal dStart )
	'Increments a date until the next specified weekday
	Dim Idx, dEnd

	dEnd = dStart
	dEnd = GetLastDayInMonth( dEnd )


	'Add any odd days
	Do While Weekend(DatePart("w",dEnd))
		dEnd = DateAdd("d",-1,dEnd)
	Loop

	GetLastWeekDay = dEnd

End Function


'------------------------------------------------------------------------------
Function GetLastWeekend( ByVal dStart )
	'Increments a date until the next specified weekday
	Dim Idx, dEnd

	dEnd = dStart
	dEnd = GetLastDayInMonth( dEnd )

	'Add any odd days
	Do While Not Weekend(dEnd)
		dEnd = DateAdd("d",-1,dEnd)
	Loop

	GetLastWeekend = dEnd

End Function


'------------------------------------------------------------------------------
Function GetALastWeekday( ByVal iDay, ByVal dStart )
	'Increments a date until the next specified weekday
	Dim Idx, dEnd

	dEnd = dStart
	dEnd = GetLastDayInMonth(dEnd)

	'Add any odd days
	Do While Not DatePart("w",dEnd) = iDay
		dEnd = DateAdd("d",-1,dEnd)
	Loop

	GetALastWeekday = dEnd

End Function


'------------------------------------------------------------------------------
Function GetAnyWeekday( ByVal iNumDays, ByVal dStart )
	'Increments a date until the next weekday
	Dim Idx, dEnd

	dEnd = CDate(dStart)

	'Add any odd days
	For Idx = 1 To iNumDays
		Do While Weekend(dEnd)
			dEnd = DateAdd("d",1,dEnd)
		Loop
		
		If Idx <> iNumDays Then
			dEnd = DateAdd("d",1,dEnd)
		End If
	Next

	GetAnyWeekDay = dEnd

End Function


'------------------------------------------------------------------------------
Function NextDate( ByVal iMonth, ByVal iDay, ByVal dEnd)
	'Increments a date until the next weekday
	Dim Idx, dTemp

	If GetDaysInMonth( dEnd ) < iDay Then 
		iDay = GetDaysInMonth(dEnd)
	End If 

	dTemp = DateSerial(DatePart("yyyy",dEnd),iMonth,iDay)
	If dTemp < dEnd Then
		dTemp = DateAdd("yyyy",1,dTemp)
	End If

	NextDate = dTemp

End Function


'------------------------------------------------------------------------------
Function GetAnyWeekend( ByVal iNumDays, ByVal dStart )
	'Increments a date until the next weekend
	Dim Idx, dEnd

	dEnd = dStart

	For Idx = 1 To iNumDays
		'Add any odd days
		Do While Not Weekend(dEnd)
			dEnd = DateAdd("d",1,dEnd)
		Loop

		If Idx <> iNumDays Then
			dEnd = DateAdd("d",1,dEnd)
		End If
	Next

	GetAnyWeekend = dEnd

End Function


'------------------------------------------------------------------------------
Function GetDaysInMonth( ByVal dDate )
	Dim dTemp, iYear, iMonth

	iYear = DatePart("yyyy",dDate)
	iMonth = DatePart("m",dDate)
	dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))
	GetDaysInMonth = Day(dTemp)

End Function


'------------------------------------------------------------------------------
Function GetLastDayInMonth( ByVal dDate )
	Dim dTemp, iYear, iMonth

	iYear = DatePart("yyyy",dDate)
	iMonth = DatePart("m",dDate)
	dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))
	GetLastDayInMonth = dTemp

End Function


'------------------------------------------------------------------------------
Function FirstDayInMonth( ByVal dDate )
	Dim iYear, iMonth

	iYear = DatePart("yyyy",dDate)
	iMonth = DatePart("m",dDate)
	dTemp = DateSerial(iYear, iMonth, 1)
	FirstDayInMonth = CDate(dTemp)

End Function


'------------------------------------------------------------------------------
Function Weekend( ByVal dDate )

	If WeekDay( dDate ) = VBSaturday or WeekDay( dDate ) = VBSunday Then
		Weekend = True
	Else
		Weekend = False
	End If

End Function


'------------------------------------------------------------------------------
Function NextDayOfMonth( ByVal iDay, ByVal dDate)

	If iDay <= GetDaysInMonth(dDate) Then
		Do While (Not DatePart("d",dDate) = iDay)
			dDate = DateAdd("d", 1, dDate)
		Loop
	Else
		dDate = GetLastDayInMonth( dDate )
	End If

	NextDayOfMonth = dDate

End Function


'------------------------------------------------------------------------------
Function OrdinalDate( ByVal tType, ByVal iOrd, ByVal dDate )
	Select Case tType
    Case "d"   'Day
	  	Select Case clng(iOrd)
	  		Case 1
	  		  dDate = FirstDayInMonth(dDate)
	  		Case 2
	  		  dDate = DateAdd("d",1,FirstDayInMonth(dDate))
	  		Case 3
	  		  dDate = DateAdd("d",2,FirstDayInMonth(dDate))
	  		Case 4
	  		  dDate = DateAdd("d",3,FirstDayInMonth(dDate))
	  		Case 5
	  		  dDate = GetLastDayInMonth(dDate)
	    End Select
  	Case "wd"  'Weekday
  		Select Case iOrd
	  		Case 1
	  		  dDate = FirstDayInMonth(dDate)
	  		  dDate = GetAnyWeekday(1,dDate)
	  		Case 2
	  		  dDate = FirstDayInMonth(dDate)
	  		  dDate = GetAnyWeekday(2,dDate)
	  		Case 3
	  		  dDate = FirstDayInMonth(dDate)
	  		  dDate = GetAnyWeekday(3,dDate)
	  		Case 4
	  		  dDate = FirstDayInMonth(dDate)
	  		  dDate = GetAnyWeekday(4,dDate)
	  		Case 5
	  		  dDate = GetLastWeekDay(dDate)
	    End Select
  	Case "wed" 'Weekend day
  		Select Case iOrd
	  		Case 1
	  		  dDate = FirstDayInMonth(dDate)
	  		  dDate = GetAnyWeekend(1,dDate)
	  		Case 2
	  		  dDate = FirstDayInMonth(dDate)
	  		  dDate = GetAnyWeekend(2,dDate)
	  		Case 3
	  		  dDate = FirstDayInMonth(dDate)
	  		  dDate = GetAnyWeekend(3,dDate)
	  		Case 4
	  		  dDate = FirstDayInMonth(dDate)
	  		  dDate = GetAnyWeekend(4,dDate)
	  		Case 5
	  		  dDate = GetLastWeekend(dDate)
	    End Select
    Case "1","2","3","4","5","6","7"
		Select Case iOrd
	  		Case 1
	  		  dDate = FirstDayInMonth(dDate)
	  		  dDate = GetAWeekday(1,clng(tType),dDate)
	  		Case 2
	  		  dDate = FirstDayInMonth(dDate)
	  		  dDate = GetAWeekday(2,clng(tType),dDate)
	  		Case 3
	  		  dDate = FirstDayInMonth(dDate)
	  		  dDate = GetAWeekday(3,clng(tType),dDate)
	  		Case 4
	  		  dDate = FirstDayInMonth(dDate)
	  		  dDate = GetAWeekday(4,clng(tType),dDate)
	  		Case 5
	  		  dDate = GetALastWeekday(clng(tType),dDate)
	    End Select
    End Select

	OrdinalDate = dDate

End Function


'------------------------------------------------------------------------------
Sub displayButtons( ByVal iEventID, ByVal iCalendarFeature_URL )
	Dim lcl_return_url

	'Build return to update url
	lcl_return_url = ""

	If iCalendarFeature_URL = "" Then 
		lcl_return_url = "updateevent.asp?id=" & iEventID
	Else 
		lcl_return_url = "updateevent.asp" & iCalendarFeature_URL & "&id=" & iEventID
	End If 

	'response.write "<input type=""button"" value=""" & langCancel & """ class=""button"" onclick=""history.back();"" />" & vbcrlf
	response.write "<input type=""button"" value=""" & langCancel & """ class=""button"" onclick=""location.href='" & lcl_return_url & "';"" />" & vbcrlf
	response.write "<input type=""button"" value=""" & langCreate & """ class=""button"" onclick=""CheckRecur();"" />" & vbcrlf

End Sub 


'------------------------------------------------------------------------------
Sub displayOrdinalOptions( )

	response.write "  <option value=""1"">First</option>" & vbcrlf
	response.write "  <option value=""2"">Second</option>" & vbcrlf
	response.write "  <option value=""3"">Third</option>" & vbcrlf
	response.write "  <option value=""4"">Fourth</option>" & vbcrlf
	response.write "  <option value=""5"">Last</option>" & vbcrlf

End Sub 


'------------------------------------------------------------------------------
Sub displayDayLikeOptions( )

	response.write "  <option value=""d"">Day</option>" & vbcrlf
	response.write "  <option value=""wd"">Weekday</option>" & vbcrlf
	response.write "  <option value=""wed"">Weekend Day</option>" & vbcrlf
	response.write "  <option value=""1"">Sunday</option>" & vbcrlf
	response.write "  <option value=""2"">Monday</option>" & vbcrlf
	response.write "  <option value=""3"">Tuesday</option>" & vbcrlf
	response.write "  <option value=""4"">Wednesday</option>" & vbcrlf
	response.write "  <option value=""5"">Thursday</option>" & vbcrlf
	response.write "  <option value=""6"">Friday</option>" & vbcrlf
	response.write "  <option value=""7"">Saturday</option>" & vbcrlf

End Sub 


'------------------------------------------------------------------------------
Sub displayMonthOptions( )

	response.write "  <option value=""1"">January</option>" & vbcrlf
	response.write "  <option value=""2"">February</option>" & vbcrlf
	response.write "  <option value=""3"">March</option>" & vbcrlf
	response.write "  <option value=""4"">April</option>" & vbcrlf
	response.write "  <option value=""5"">May</option>" & vbcrlf
	response.write "  <option value=""6"">June</option>" & vbcrlf
	response.write "  <option value=""7"">July</option>" & vbcrlf
	response.write "  <option value=""8"">August</option>" & vbcrlf
	response.write "  <option value=""9"">September</option>" & vbcrlf
	response.write "  <option value=""10"">October</option>" & vbcrlf
	response.write "  <option value=""11"">November</option>" & vbcrlf
	response.write "  <option value=""12"">December</option>" & vbcrlf

End Sub 


'------------------------------------------------------------------------------
sub displayRecurOptions( ByVal iRecur)

  if iRecur = "dd" OR iRecur = "" then
     response.write "  <tr>" & vbcrlf
     response.write "      <td valign=""top"" nowrap></td>" & vbcrlf
     response.write "      <td width=""100%"">" & vbcrlf
     response.write "          <input type=""radio"" name=""Often"" value=""days"" checked=""checked"" />" & vbcrlf
     response.write "          Every <input type=""text"" name=""Days"" id=""Days"" size=""3"" maxlength=""3"" value=""1"" onchange=""clearMsg('Days');"" /> day(s)." & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td valign=""top"" nowrap></td>" & vbcrlf
     response.write "      <td>" & vbcrlf
     response.write "          <input type=""radio"" name=""Often"" value=""weekdays"" onclick=""clearMsg('Days');"" /> Every weekday." & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  elseif iRecur = "ww" then
     response.write "	           <tr>" & vbcrlf
     response.write "                <td valign=""top"" nowrap></td>" & vbcrlf
     response.write "                <td width=""100%"">" & vbcrlf
     response.write "                    Recur every <input type=""text"" name=""Weeks"" id=""Weeks"" style=""width:25px;"" maxlength=""2"" value=""1"" onchange=""clearMsg('Weeks');"" /> week(s) on:" & vbcrlf
     response.write "            </tr>" & vbcrlf
     response.write "            <tr>" & vbcrlf
     response.write "                <td valign=""top"" nowrap></td>" & vbcrlf
     response.write "                <td>" & vbcrlf
     response.write "                    <input type=""checkbox"" name=""WeekDayNum"" value=""1"" id=""Sunday"" onclick=""clearMsg('Wednesday');"" /> Sunday" & vbcrlf
     response.write "                    <input type=""checkbox"" name=""WeekDayNum"" value=""2"" id=""Monday"" onclick=""clearMsg('Wednesday');"" /> Monday" & vbcrlf
     response.write "                    <input type=""checkbox"" name=""WeekDayNum"" value=""3"" id=""Tuesday"" onclick=""clearMsg('Wednesday');"" /> Tuesday" & vbcrlf
     response.write "                    <input type=""checkbox"" name=""WeekDayNum"" value=""4"" id=""Wednesday"" onclick=""clearMsg('Wednesday');"" /> Wednesday" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
     response.write "            <tr>" & vbcrlf
     response.write "                <td valign=""top"" nowrap></td>" & vbcrlf
     response.write "                <td>" & vbcrlf
     response.write "                    <input type=""checkbox"" name=""WeekDayNum"" value=""5"" id=""Thursday"" onclick=""clearMsg('Wednesday');"" /> Thursday" & vbcrlf
     response.write "                    <input type=""checkbox"" name=""WeekDayNum"" value=""6"" id=""Friday"" onclick=""clearMsg('Wednesday');"" /> Friday" & vbcrlf
     response.write "                    <input type=""checkbox"" name=""WeekDayNum"" value=""7"" id=""Saturday"" onclick=""clearMsg('Wednesday');"" /> Saturday" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  elseif iRecur = "mm" then
     response.write "	           <tr>" & vbcrlf
     response.write "               <td valign=""top"" nowrap></td>" & vbcrlf
     response.write "               <td width=""100%"">" & vbcrlf
     response.write "                   <input type=""radio"" name=""Often"" value=""absolute"" checked=""checked"" onclick=""clearMsg('Month');"" /> Day" & vbcrlf
     response.write "                   <input type=""text"" name=""Day"" id=""Day"" style=""width:25px;"" maxlength=""2"" value=""1"" onchange=""clearMsg('Day');"" /> of every " & vbcrlf
     response.write "                   <input type=""text"" name=""Months"" id=""Months"" style=""width:25px;"" maxlength=""2"" value=""1"" onchange=""clearMsg('Months');"" /> month(s)." & vbcrlf
     response.write "               </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
     response.write "            <tr>" & vbcrlf
     response.write "                <td valign=""top"" nowrap></td>" & vbcrlf
     response.write "                <td width=""100%"">" & vbcrlf
     response.write "                    <input type=""radio"" name=""Often"" value=""relative"" onclick=""clearMsg('Day');clearMsg('Months');"" /> The" & vbcrlf
     response.write "                    <select name=""Ordinal"" class=""time"">" & vbcrlf
                                              displayOrdinalOptions
     response.write "                    </select>" & vbcrlf
     response.write "                    <select name=""DayLike"" class=""time"">" & vbcrlf
                                              displayDayLikeOptions
     response.write "                    </select>" & vbcrlf
     response.write "                    of every <input type=""text"" name=""Month"" id=""Month"" style=""width:25px;"" maxlength=""2"" value=""1"" onchange=""clearMsg('Month');"" /> month(s)." & vbcrlf
     response.write "           			  </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  elseif iRecur = "yy" then
     response.write " 	          <tr>" & vbcrlf
     response.write "                <td valign=""top"" nowrap></td>" & vbcrlf
     response.write "                <td width=""100%"">" & vbcrlf
     response.write "                    <input type=""radio"" name=""Often"" value=""every"" checked=""checked"" onclick=""clearMsg('DayNum');clearMsg('Months');"" /> Every" & vbcrlf
     response.write "                    <select name=""Month"" class=""time"">" & vbcrlf
                                              displayMonthOptions
     response.write "                    </select>" & vbcrlf
     response.write "                    <input type=""text"" name=""Day"" id=""Day"" style=""width:25px;"" maxlength=""2"" value=""1"" />." & vbcrlf
     response.write "            				</td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
     response.write " 	          <tr>" & vbcrlf
     response.write "                <td valign=""top"" nowrap></td>" & vbcrlf
     response.write "                <td width=""100%"">" & vbcrlf
     response.write "                    <input type=""radio"" name=""Often"" value=""absolute""  onclick=""clearMsg('Day');""/> Day" & vbcrlf
     response.write "                    <input type=""text"" name=""DayNum"" id=""DayNum"" style=""width:25px;"" maxlength=""2"" value=""1"" onchange=""clearMsg('DayNum');"" /> of every " & vbcrlf
     response.write "                    <input type=""text"" name=""Months"" id=""Months"" style=""width:25px;"" maxlength=""2"" value=""1"" onchange=""clearMsg('Months');"" /> month(s)." & vbcrlf
     response.write "            </tr>" & vbcrlf
     response.write "            <tr>" & vbcrlf
     response.write "                <td valign=""top"" nowrap></td>" & vbcrlf
     response.write "                <td width=""100%"">" & vbcrlf
     response.write "                    <input type=""radio"" name=""Often"" value=""relative"" onclick=""clearMsg('Day');clearMsg('DayNum');clearMsg('Months');"" /> The" & vbcrlf
     response.write "                    <select name=""Ordinal"" class=""time"">" & vbcrlf
                                              displayOrdinalOptions
     response.write "                    </select>" & vbcrlf
     response.write "                    <select name=""DayPick"" class=""time"">" & vbcrlf
                                              displayDayLikeOptions
     response.write "                    </select>" & vbcrlf
     response.write "                    of" & vbcrlf
     response.write "                    <select name=""MonthPick"" class=""time"">" & vbcrlf
                                              displayMonthOptions
     response.write "                    </select>" & vbcrlf
     response.write "            			 </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  end If
  
end Sub


%>
