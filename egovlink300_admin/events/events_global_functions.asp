<%
	'------------------------------------------------------------------------------
function getFeatureByID( ByVal p_orgid, ByVal p_featureid )
	Dim sSql, lcl_return, oGetFeatureByID
	lcl_return = ""

	'sSql = "SELECT ISNULL(FO.featurename,F.featurename) AS featurename "
	sSql = "SELECT f.feature "
	sSql = sSql & " FROM egov_organizations_to_features FO, egov_organization_features F "
	sSql = sSql & " WHERE FO.featureid = F.featureid "
	sSql = sSql & " AND FO.orgid = " & p_orgid
	sSql = sSql & " AND f.featureid = " & p_featureid

	set oGetFeatureByID = Server.CreateObject("ADODB.Recordset")
	oGetFeatureByID.Open sSql, Application("DSN"), 0, 1

	if not oGetFeatureByID.eof then
		lcl_return = oGetFeatureByID("feature")
	end if

	oGetFeatureByID.Close
	set oRs = nothing

	getFeatureByID = lcl_return

end function

'------------------------------------------------------------------------------
sub getEventCategoryOptions(p_orgid, p_calendarfeature, p_selected)

   'Get and display the options, if any exist, in the Event Category dropdown list.
    sSql = "SELECT * "
    sSql = sSql & " FROM eventcategories "
    sSql = sSql & " WHERE orgid = " & p_orgid

    if p_calendarfeature <> "" then
       sSql = sSql & " AND UPPER(calendarfeature) = '" & UCASE(dbsafe(p_calendarfeature)) & "'"
    else
       sSql = sSql & " AND (calendarfeature = '' OR calendarfeature IS NULL)"
    end if

'    sSql = sSql & " AND '" & checkForCustomCalendars(p_orgid) & "' = 'Y' "

    sSql = sSql & " ORDER BY categoryname "

    set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sSql, Application("DSN"), 0, 1

    if not rs.eof then
       while NOT rs.eof
          if rs("CategoryID") = p_selected then
             sSel = "SELECTED"
   					  else
			   		     sSel = ""
          end if

          response.write "<option value=""" & rs("CategoryID") & """ " & sSel & ">" & rs("CategoryName") & "</option>" & vbcrlf

          rs.movenext
       wend
    end if

    set rs = nothing
end sub

'---------------------------------------------------------------
 function getCategoryID( p_orgid, p_categoryname, p_calendarfeature )
    lcl_return = ""

    sSql = "SELECT categoryid "
    sSql = sSql & " FROM eventcategories "
    sSql = sSql & " WHERE categoryname = '" & p_categoryname & "' "
    sSql = sSql & " AND orgid = " & p_orgid

    if p_calendarfeature <> "" then
       sSql = sSql & " AND UPPER(calendarfeature) = '" & UCASE(dbsafe(p_calendarfeature)) & "'"
    else
       sSql = sSql & " AND (calendarfeature = '' OR calendarfeature IS NULL)"
    end if

    set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sSql, Application("DSN"), 0, 1

    if not rs.eof then
       lcl_return = rs("categoryid")
    end if

    set rs = nothing

    getCategoryID = lcl_return

 end function

'---------------------------------------------------------------
 sub newCategory( ByVal p_orgid, ByVal p_categoryname, ByVal p_color, ByVal p_calendarfeature, ByRef lcl_identity )
    lcl_calendar_feature = ""

    if p_calendarfeature <> "" then
       lcl_calendar_feature = dbsafe(p_calendarfeature)
    else
       lcl_calendar_feature = ""
    end if

    sSql = "INSERT INTO eventcategories(orgid,categoryname,color,calendarfeature) VALUES ("
    sSql = sSql &       p_orgid                & ", "
    sSql = sSql & "'" & dbsafe(p_categoryname) & "', "
    sSql = sSql & "'" & dbsafe(p_color)        & "', "
    sSql = sSql & "'" & lcl_calendar_feature   & "' "
    sSql = sSql & ")"

    set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sSql, Application("DSN"), 3, 1

   'Retrieve the posting_id that was just inserted
    sSqlid = "SELECT IDENT_CURRENT('eventcategories') as NewID"
    rs.Open sSqlid, Application("DSN") , 3, 1
    lcl_identity = rs("NewID").value

    set rs = Nothing
    
 end sub


'---------------------------------------------------------------
Sub newRecurEvent(ByVal p_orgid, ByVal p_eventid, ByVal p_eventdate, ByVal p_creatoruserid, ByVal p_eventtimezoneid, _
                  ByVal p_eventduration, ByVal p_subject, ByVal p_message, ByVal p_categoryid, ByVal p_calendarfeature, _
                  ByVal p_isHiddenCL, ByRef lcl_identity)

	Dim sSql, lcl_isHiddenCL, rsu, rs

    if p_isHiddenCL then
       lcl_isHiddenCL = 1
    else
       lcl_isHiddenCL = 0
    end if

   'Update the Events table.  Set this event as a "recurring event"
    sSql = "UPDATE events SET recurid = " & p_eventid & " WHERE eventid = " & p_eventid
    Set rsu = Server.CreateObject("ADODB.Recordset")
    rsu.Open sSql, Application("DSN"), 3, 1

	Set rsu = Nothing 

   'Create the "recurring event"
    sSql = "INSERT INTO events( orgid, creatoruserid, eventdate, eventtimezoneid, eventduration, [Subject], [Message], "
    sSql = sSql & " modifieruserid, categoryid, recurid, calendarfeature, isHiddenCL ) "
    sSql = sSql & " VALUES ( "
    sSql = sSql &       p_orgid                   & ", "
    sSql = sSql &       p_creatoruserid           & ", "
    sSql = sSql & "'" & dbsafe(p_eventdate)       & "', "
    sSql = sSql &       p_eventtimezoneid         & ", "
    sSql = sSql &       p_eventduration           & ", "
    sSql = sSql & "'" & dbsafe(p_subject)         & "', "
    sSql = sSql & "'" & dbsafe(p_message)         & "', "
    sSql = sSql &       p_creatoruserid           & ", "
    sSql = sSql &       p_categoryid              & ", "
    sSql = sSql &       p_eventid                 & ", "
    sSql = sSql & "'" & dbsafe(p_calendarfeature) & "', "
    sSql = sSql &       lcl_isHiddenCL
    sSql = sSql & " )"
	'response.write sSql & "<br /><br />"

    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sSql, Application("DSN"), 3, 1

   'Retrieve the posting_id that was just inserted
    sSql = "SELECT IDENT_CURRENT('events') as NewID"
    rs.Open sSql, Application("DSN"), 3, 1
    lcl_identity = rs("NewID").value

    Set rs = Nothing 

End Sub 


'---------------------------------------------------------------
 function getCategoryName( ByVal p_categoryid )
	Dim lcl_return, sSql, rs

	lcl_return = ""

	sSql = "SELECT categoryname "
	sSql = sSql & " FROM eventcategories "
	sSql = sSql & " WHERE categoryid = " & p_categoryid

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSql, Application("DSN"), 0, 1

	if not rs.eof then
		lcl_return = rs("categoryname")
	end if

	set rs = nothing

	getCategoryName = lcl_return

 end function

'---------------------------------------------------------------
 function updateCategory( ByVal p_categoryid, ByVal p_categoryname, ByVal p_color, ByVal p_calendarfeature, ByRef lcl_identity )
	Dim sSql, rs

	sSql = "UPDATE eventcategories SET "
	sSql = sSql & " categoryname = '"    & p_categoryname    & "', "
	sSql = sSql & " color = '"           & p_color           & "', "
	sSql = sSql & " calendarfeature = '" & p_calendarfeature & "' "
	sSql = sSql & " WHERE categoryid = " & p_categoryid

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSql, Application("DSN"), 3, 1

	'Retrieve the categoryid that was just inserted
	sSql = "SELECT IDENT_CURRENT('eventcategories') as NewID"
	rs.Open sSql, Application("DSN"), 3, 1
	lcl_identity = rs("NewID").value

	set rs  = nothing

 end function


'---------------------------------------------------------------
 sub delCategory( ByVal p_categoryid )
	Dim sSql, rsd, rs

	'Delete the event category
	sSql = "DELETE FROM eventcategories WHERE categoryid = " & p_categoryid
	set rsd = Server.CreateObject("ADODB.Recordset")
	rsd.Open sSql, Application("DSN"), 3, 1

	'Remove the event category from any events that have the event category being deleted.
	sSql = "UPDATE events SET categoryid = 0 WHERE categoryid = " & p_categoryid
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSql, Application("DSN"), 3, 1

	set rsd = nothing
	set rs  = nothing

 end sub


'---------------------------------------------------------------
'When a user accesses "View Internal Calendars" a default calendar is not queried on.  We need to get the first one in the
'calendar dropdown list and set that as the default.
'Pull all of the custom calendars that are considered "internal"
'This is based off of the "haspublicview" column.  Here we want it to be "false"
 function getFirstCalendarInList( ByVal p_orgid )
	Dim lcl_return, sSql, oExists

	lcl_return = ""

	'sSql = "SELECT TOP 1 F.feature AS feature "
	sSql = "SELECT f.featureid, f.feature "
	sSql = sSql & "FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
	sSql = sSql & "WHERE "
	'sSql = sSql & " FO.publiccanview = 1 "
	sSql = sSql & " F.haspublicview = 0 "
	sSql = sSql & " AND f.feature <> 'internal_calendars' "
	sSql = sSql & " AND O.orgid = FO.orgid "
	sSql = sSql & " AND FO.featureid = F.featureid "
	sSql = sSql & " AND O.orgid = " & p_orgid
	sSql = sSql & " AND f.parentfeatureid = (select f2.featureid from egov_organization_features as f2 where UPPER(f2.feature) = 'CUSTOM_CALENDARS')"
	sSql = sSql & " ORDER BY FO.publicdisplayorder,F.publicdisplayorder"

	Set oExists = Server.CreateObject("ADODB.Recordset")
	oExists.Open sSql, Application("DSN"), 0, 1

	if Not oExists.eof then
		lcl_linecnt = 0
		do while NOT oExists.eof
			if userhaspermission(session("userid"),oExists("feature")) then
				lcl_return = oExists("featureid")
				exit do
			end if
			oExists.movenext
		loop
	end if

	getFirstCalendarInList = lcl_return

	set oExists = nothing

 end function

'---------------------------------------------------------------
'Pull all of the custom calendars that are considered "internal"
'This is based off of the "haspublicview" column.  Here we want it to be "false"
 function checkForCustomCalendars( ByVal p_orgid )
   lcl_return = "N"
   lcl_count  = 0

		 sSql = "SELECT f.feature "
		 sSql = sSql & "FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		 sSql = sSql & "WHERE "
'   sSql = sSql & " FO.publiccanview = 1 "
   sSql = sSql & " F.haspublicview = 0 "
   sSql = sSql & " AND f.feature <> 'internal_calendars' "
   sSql = sSql & " AND O.orgid = FO.orgid "
   sSql = sSql & " AND FO.featureid = F.featureid "
   sSql = sSql & " AND O.orgid = " & p_orgid
   sSql = sSql & " AND f.parentfeatureid = (select f2.featureid from egov_organization_features as f2 where UPPER(f2.feature) = 'CUSTOM_CALENDARS')"

		 Set oExists = Server.CreateObject("ADODB.Recordset")
		 oExists.Open sSql, Application("DSN"), 0, 1

   if NOT oExists.eof then
      while NOT oExists.eof
         'if userhaspermission(session("userid"),oExists("feature")) then
            lcl_count = lcl_count + 1
         'else
         '   lcl_count = lcl_count
         'end if
         oExists.movenext
      wend
   end if

   if lcl_count > 0 then 
      lcl_return = "Y"
   end if

   checkForCustomCalendars = lcl_return

   set oExists = nothing

 end function

'---------------------------------------------------------------
'Pull all of the custom calendars that are considered "internal"
'This is based off of the "haspublicview" column.  Here we want it to be "false"
 sub displayCustomCalendarOptions(p_orgid, p_calendarfeatureid)
		 'sSql = "SELECT isnull(FO.featurename,F.featurename) as featurename, F.feature "
		 sSql = "SELECT F.featureid, isnull(FO.featurename,F.featurename) as featurename, F.feature "
		 sSql = sSql & "FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		 sSql = sSql & "WHERE "
   'sSql = sSql & " FO.publiccanview = 1 "
   sSql = sSql & " F.haspublicview = 0 "
   sSql = sSql & " AND O.orgid = FO.orgid "
   sSql = sSql & " AND FO.featureid = F.featureid "
   sSql = sSql & " AND O.orgid = " & p_orgid
   sSql = sSql & " AND f.parentfeatureid = (select f2.featureid from egov_organization_features as f2 where UPPER(f2.feature) = 'CUSTOM_CALENDARS')"
   sSql = sSql & " AND f.feature <> 'internal_calendars' "
		 sSql = sSql & " ORDER BY FO.publicdisplayorder,F.publicdisplayorder"

		 set oExists = Server.CreateObject("ADODB.Recordset")
		 oExists.Open sSql, Application("DSN"), 0, 1

   if NOT oExists.eof then
      while NOT oExists.eof
         'if userhaspermission(session("userid"),oExists("feature")) then
            'if UCASE(oExists("feature")) = UCASE(p_calendarfeature) then
            if UCASE(oExists("featureid")) = UCASE(p_calendarfeatureid) then
               lcl_selected = " selected=""selected"""
            else
               lcl_selected = ""
            end if

            response.write "  <option value=""" & oExists("featureid") & """" & lcl_selected & ">" & oExists("featurename") & "</option>" & vbcrlf
         'end if

         oExists.movenext
      wend
   end if

   set oExists = nothing

 end sub

'------------------------------------------------------------------------------
function getItemsPerDay(p_orgid)
  lcl_return = 4

  if p_orgid <> "" then
     sSql = "SELECT isnull(calendar_numItemsPerDay,4) as calendar_numItemsPerDay "
     sSql = sSql & " FROM organizations "
     sSql = sSql & " WHERE orgid = " & p_orgid

  		 set oGetItemsPerDay = Server.CreateObject("ADODB.Recordset")
		   oGetItemsPerDay.Open sSql, Application("DSN"), 0, 1

     if not oGetItemsPerDay.eof then
        lcl_return = oGetItemsPerDay("calendar_numItemsPerDay")
     end if

     oGetItemsPerDay.close
     set oGetItemsPerDay = nothing

  end if

  getItemsPerDay = lcl_return

end function

'------------------------------------------------------------------------------
function getPushFieldAnswer(iPushToFeature, iPushTable, iPushColumn, iRequestID)

  lcl_return = ""

  if  iPushToFeature <> "" _
  AND iPushTable     <> "" _
  AND iPushColumn    <> "" _
  AND iRequestID     <> "" then
      lcl_pushtofeature = UCASE(iPushToFeature)
      lcl_pushtable     = UCASE(iPushTable)
      lcl_pushcolumn    = UCASE(iPushColumn)

      sSql = "SELECT rfr.submitted_request_field_response, "
      sSql = sSql & " pf.pushfieldid, "
      sSql = sSql & " pf.push_table, "
      sSql = sSql & " pf.push_column, "
      sSql = sSql & " pf.push_column_datatype, "
      sSql = sSql & " pf.push_column_label, "
      sSql = sSql & " pf.push_to_feature, "
      sSql = sSql & " pf.push_feature_permission "
      sSql = sSql & " FROM egov_actionline_pushfields AS pf, "
      sSql = sSql &      " egov_submitted_request_field_responses rfr "
      sSql = sSql & " WHERE pf.pushfieldid = rfr.submitted_request_pushfieldid "
      sSql = sSql & " AND rfr.submitted_request_field_id IN (select rf.submitted_request_field_id "
      sSql = sSql &                                        " from egov_submitted_request_fields rf "
      sSql = sSql &                                        " where rf.submitted_request_id = " & iRequestID & ") "
      sSql = sSql & " AND rfr.submitted_request_field_response NOT LIKE 'default_novalue' "
      sSql = sSql & " AND rfr.submitted_request_field_response NOT LIKE '' "
      sSql = sSql & " AND rfr.submitted_request_field_response IS NOT NULL "
      sSql = sSql & " AND UPPER(pf.push_to_feature) = '" & lcl_pushtofeature & "' "
      sSql = sSql & " AND UPPER(pf.push_table) = '"      & lcl_pushtable     & "' "
      sSql = sSql & " AND UPPER(pf.push_column) = '"     & lcl_pushcolumn    & "' "

      set oAnswerData = Server.CreateObject("ADODB.Recordset")
      oAnswerData.Open sSql, Application("DSN"), 0, 1

      if not oAnswerData.eof then

         if lcl_pushcolumn = "EVENTDATE" then
            lcl_pushdate       = ""
            lcl_pushdate_month = ""
            lcl_pushdate_day   = ""
            lcl_pushdate_year  = ""

            if oAnswerData("submitted_request_field_response") <> "" then
               do while not oAnswerData.eof
                  if UCASE(oAnswerData("push_column_label")) = "EVENT DATE (MONTH)" then
                     lcl_pushdate_month = getMonth(oAnswerData("submitted_request_field_response"))
                     lcl_pushdate_day   = lcl_pushdate_day
                     lcl_pushdate_year  = lcl_pushdate_year
                  elseif UCASE(oAnswerData("push_column_label")) = "EVENT DATE (DAY)" then
                     lcl_pushdate_month = lcl_pushdate_month
                     lcl_pushdate_day   = oAnswerData("submitted_request_field_response")
                     lcl_pushdate_year  = lcl_pushdate_year
                  elseif UCASE(oAnswerData("push_column_label")) = "EVENT DATE (YEAR)" then
                     lcl_pushdate_month = lcl_pushdate_month
                     lcl_pushdate_day   = lcl_pushdate_day
                     lcl_pushdate_year  = oAnswerData("submitted_request_field_response")
                  end if

                  oAnswerData.movenext
               loop

              'Build the Date
               if lcl_pushdate_month <> "" AND lcl_pushdate_day <> "" then
                  if lcl_pushdate_year = "" then
                     lcl_pushdate_year = Year(Now)
                  end if

                  lcl_pushdate = lcl_pushdate_month & "/" & lcl_pushdate_day & "/" & lcl_pushdate_year
               end if
              
              'If the date is LESS THAN the current date, add a single year to the date.
               if lcl_pushdate <> "" then
                  if datediff("d",date(),lcl_pushdate) < 0 then
                     lcl_pushdate = dateadd("yyyy",1,lcl_pushdate)
                  end if
               end if

               lcl_return = lcl_pushdate

            end if
         else
            lcl_return = oAnswerData("submitted_request_field_response")
         end if

      end if

      oAnswerData.close
      set oAnswerData = nothing
  end if

  getPushFieldAnswer = lcl_return

end function

'------------------------------------------------------------------------------
function getMonth(iMonthName)
  lcl_return = ""

  if iMonthName <> "" then
     lcl_monthname = UCASE(iMonthName)

     if     lcl_monthname = "JANUARY" then
        lcl_return = "1"
     elseif lcl_monthname = "FEBRUARY" then
        lcl_return = "2"
     elseif lcl_monthname = "MARCH" then
        lcl_return = "3"
     elseif lcl_monthname = "APRIL" then
        lcl_return = "4"
     elseif lcl_monthname = "MAY" then
        lcl_return = "5"
     elseif lcl_monthname = "JUNE" then
        lcl_return = "6"
     elseif lcl_monthname = "JULY" then
        lcl_return = "7"
     elseif lcl_monthname = "AUGUST" then
        lcl_return = "8"
     elseif lcl_monthname = "SEPTEMBER" then
        lcl_return = "9"
     elseif lcl_monthname = "OCTOBER" then
        lcl_return = "10"
     elseif lcl_monthname = "NOVEMBER" then
        lcl_return = "11"
     elseif lcl_monthname = "DECEMBER" then
        lcl_return = "12"
     end if
  end if

  getMonth = lcl_return

end function

'------------------------------------------------------------------------------
sub buildOption( ByVal iOptionType, ByVal iValue, ByVal iCurrentValue )

  lcl_isSelected   = ""
  lcl_optiontype   = ""
  lcl_value        = ""
  lcl_currentvalue = ""
  lcl_displayvalue = ""

  if iOptionType <> "" then
     lcl_optiontype = ucase(iOptionType)
  end if

  if iValue <> "" then
     lcl_value = iValue
  end if

  if iCurrentValue <> "" then
     lcl_currentvalue = iCurrentValue
  end if

 'Determine which option is select for the option type being built
  if lcl_optiontype = "DURATION" then
     if (lcl_value = "1"     AND lcl_currentvalue = "m") _
     or (lcl_value = "60"    AND lcl_currentvalue = "h") _
     or (lcl_value = "1440"  AND lcl_currentvalue = "d") _
     or (lcl_value = "10080" AND lcl_currentvalue = "w") then
        lcl_isSelected = " selected=""selected"""
     end if

     if lcl_value = "1" then
        lcl_displayvalue = "Minutes"
     elseif lcl_value = "60" then
        lcl_displayvalue = "Hours"
     elseif lcl_value = "1440" then
        lcl_displayvalue = "Days"
     elseif lcl_value = "10080" then
        lcl_displayvalue = "Weeks"
     end if
  elseif lcl_optiontype = "MINUTE" then
     'if (lcl_value = "00" AND (lcl_currentvalue >= 0  AND lcl_currentvalue < 15)) _
     'or (lcl_value = "15" AND (lcl_currentvalue >= 15 AND lcl_currentvalue < 30)) _
     'or (lcl_value = "30" AND (lcl_currentvalue >= 30 AND lcl_currentvalue < 45)) _
     'or (lcl_value = "45" AND (lcl_currentvalue >= 45 AND lcl_currentvalue < 60)) then
     '   lcl_isSelected = " selected=""selected"""
     'end if

     if (lcl_value = "00" AND (lcl_currentvalue  = 0  AND lcl_currentvalue < 5)) _
     or (lcl_value = "05" AND (lcl_currentvalue >= 5  AND lcl_currentvalue < 10)) _
     or (lcl_value = "10" AND (lcl_currentvalue >= 10 AND lcl_currentvalue < 15)) _
     or (lcl_value = "15" AND (lcl_currentvalue >= 15 AND lcl_currentvalue < 20)) _
     or (lcl_value = "20" AND (lcl_currentvalue >= 20 AND lcl_currentvalue < 25)) _
     or (lcl_value = "25" AND (lcl_currentvalue >= 25 AND lcl_currentvalue < 30)) _
     or (lcl_value = "30" AND (lcl_currentvalue >= 30 AND lcl_currentvalue < 35)) _
     or (lcl_value = "35" AND (lcl_currentvalue >= 35 AND lcl_currentvalue < 40)) _
     or (lcl_value = "40" AND (lcl_currentvalue >= 40 AND lcl_currentvalue < 45)) _
     or (lcl_value = "45" AND (lcl_currentvalue >= 45 AND lcl_currentvalue < 50)) _
     or (lcl_value = "50" AND (lcl_currentvalue >= 50 AND lcl_currentvalue < 55)) _
     or (lcl_value = "55" AND (lcl_currentvalue >= 55 AND lcl_currentvalue < 60)) then
        lcl_isSelected = " selected=""selected"""
     end if

     lcl_displayvalue = lcl_value

  else  'HOUR, AMPM
     if lcl_value = lcl_currentvalue then
        lcl_isSelected   = " selected=""selected"""
     end if

     lcl_displayvalue = lcl_value

  end if

 'Build the option
  response.write "<option value=""" & lcl_value & """" & lcl_isSelected & ">" & lcl_displayvalue & "</option>" & vbcrlf

end sub

'------------------------------------------------------------------------------
function eventFormatDateTime(iEventDate)
  Dim iHour, sTime, lcl_return

  lcl_return = ""
  iHour      = Hour(iEventDate)

  If iHour > 12 Then
     iHour = iHour - 12
  end if

  sTime = iHour & ":" & Right("00" & Minute(iEventDate),2) & " " & Right(iEventDate, 2)
  
  'check if 12 am
  If iHour = 0 then
    'sTime ="12:00 AM"
    sTime = "12:" & Right("00" & Minute(iEventDate),2)
    sTime = sTime & " AM"
  end if

  lcl_return = FormatDateTime(iEventDate, vbShortDate) & "<br />" & sTime

  eventFormatDateTime = lcl_return

end function

'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Successfully Updated..."
     elseif iSuccess = "SA" then
        lcl_return = "Successfully Created..."
     elseif iSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
     elseif iSuccess = "RSS_SUCCESS" then
        lcl_return = "Successfully Sent to RSS..."
     elseif iSuccess = "RSS_ERROR" then
        lcl_return = "ERROR: Failed to send to RSS..."
     elseif iSuccess = "AJAX_ERROR" then
        lcl_return = "ERROR: An error has during the AJAX routine..."
     end if
  end if

  setupScreenMsg = lcl_return

end function

'---------------------------------------------------------------
 function dbsafe(p_value)
    lcl_return = ""

    if trim(p_value) <> "" then
       lcl_return = replace(p_value,"'","''")
    end if

    dbsafe = lcl_return
 end function
%>
