<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
  dim sOrgID, sRentalID, sPeriodTypeSelector, sReservationTempID
 	dim aWantedDates(1,0)

  sOrgID              = 0
  sRentalID           = 0
  sPeriodTypeSelector = ""
  sReservationTempID  = 0

  if request("orgid") <> "" then
     sOrgID = clng(request("orgid"))
  end if

  if request("rentalid") <> "" then
     sRentalID = clng(request("rentalid"))
  end if

  if request("periodtypeselector") <> "" then
     if not containsApostrophe(request("periodtypeselector")) then
        sPeriodTypeSelector = request("periodtypeselector")
     end if
  end if

  if request("reservationtempid") <> "" then
     sReservationTempID = clng(request("reservationtempid"))
  end if

  if sOrgID > 0 AND sReservationTempID > 0 then
    	iRowCount  = 0
     lcl_return = ""

     lcl_return = lcl_return & "<table id=""reservationtempdates"" cellpadding=""0"" cellspacing=""0"" border=""0"">" & vbcrlf
     lcl_return = lcl_return & "  <tr>" & vbcrlf
     lcl_return = lcl_return & "      <th class=""firstcell"">Include</th>" & vbcrlf
     lcl_return = lcl_return & "      <th>Date</th>" & vbcrlf
     lcl_return = lcl_return & "      <th>Start Time</th>" & vbcrlf
     lcl_return = lcl_return & "      <th>End Time</th>" & vbcrlf
     lcl_return = lcl_return & "      <th class=""lastcell"">Available</th>" & vbcrlf
     lcl_return = lcl_return & "  </tr>" & vbcrlf

    'Get the temp dates
    	sSQL = "SELECT "
     sSQL = sSQL & " reservationstarttime, "
     sSQL = sSQL & " reservationendtime, "
     sSQL = sSQL & " endday "
     sSQL = sSQL & " FROM egov_rentalreservationdatestemp "
    	sSQL = sSQL & " WHERE reservationtempid = " & sReservationTempID
     sSQL = sSQL & " AND orgid = " & sOrgID
    	sSQL = sSQL & " ORDER BY reservationstarttime"

    	set oRs = Server.CreateObject("ADODB.Recordset")
    	oRs.Open sSQL, Application("DSN"), 0, 1

    	do while not oRs.eof
       	iRowCount             = iRowCount + 1
      		sReservationStartTime = oRs("reservationstarttime")
      		sReservationEndTime   = oRs("reservationendtime")
      		iEndDay               = oRs("endday")
       	bOffSeasonFlag        = GetOffSeasonFlag( sRentalID, DateValue(sReservationStartTime) )

      		if RentalIsAllDay( sRentalID, bOffSeasonFlag, Weekday(DateValue(sReservationStartTime)) ) then
        			bIsAllDayOnly   = true
        			sDisabledOption = "disabled"
      		else
        			bIsAllDayOnly = false
        			sDisabledOption = ""
       	end if

     		'If this is for all day, or the rental is only available for all day reservations then we need the opening and closing times on that day
      		if sPeriodTypeSelector = "allday" Or bIsAllDayOnly then
        			SetHoursToOpenAndClose sRentalID, DateValue(sReservationStartTime), sReservationStartTime, sReservationEndTime, iEndDay
      		end if

	      'Get the availability flag on that date and time
      		aWantedDates(0,0) = oRs("reservationstarttime")
      		aWantedDates(1,0) = oRs("reservationendtime")

        if iRowCount > 1 then
         	'The seperator Row
         		lcl_return = lcl_return & "<tr><td colspan=""5"" class=""tempseparator"">&nbsp;</td></tr>" & vbcrlf
        end if

   	  	'Build the date time row
      		lcl_return = lcl_return & "<tr class=""dateline"">" & vbcrlf
      		lcl_return = lcl_return & "    <td class=""firstcell_nopadding"" align=""center"">" & vbcrlf
      		'lcl_return = lcl_return & "        <input type=""checkbox"" id=""reservationtime" & iRowCount & """ name=""reservationtime" & iRowCount & """ />&nbsp;"
        lcl_return = lcl_return & "        <input type=""checkbox"" name=""includereservationtime_"       & sRentalID & "_" & iRowCount & """ id=""includereservationtime_"         & sRentalID & "_" & iRowCount & """ value=""Y"" checked=""checked"" onclick=""clearMsg('endday_" & sRentalID & "_" & iRowCount & "');enableDisableFields('" & sRentalID & "_" & iRowCount & "')"" />" & vbcrlf
        lcl_return = lcl_return & "        <input type=""hidden"" name=""errorCheck_reservationDateTime_" & sRentalID & "_" & iRowCount & """ id=""errorCheck_reservationDateTime_" & sRentalID & "_" & iRowCount & """ value="""" size=""20"" maxlength=""25"" />" & vbcrlf
        lcl_return = lcl_return & "        <input type=""hidden"" name=""errorCheck_errorCode_"           & sRentalID & "_" & iRowCount & """ id=""errorCheck_errorCode_"           & sRentalID & "_" & iRowCount & """ value="""" size=""10"" maxlength=""50"" />" & vbcrlf
        lcl_return = lcl_return & "        <input type=""hidden"" name=""errorCheck_okToContinue_"        & sRentalID & "_" & iRowcount & """ id=""errorCheck_okToContinue_"        & sRentalID & "_" & iRowCount & """ value="""" size=""10"" maxlength=""50"" />" & vbcrlf
        'lcl_return = lcl_return & "				    <input type=""hidden"" name=""rentalid_" & sRentalID & "_" & iRowCount & """ id=""rentalid_" & sRentalID & "_" & iRowCount & """ value=""" & sRentalID & """ />" & vbcrlf
        lcl_return = lcl_return & "    </td>" & vbcrlf
      		lcl_return = lcl_return & "    <td>" & vbcrlf
       	lcl_return = lcl_return & "        <input type=""text"" name=""startdate_" & sRentalID & "_" & iRowCount & """ id=""startdate_" & sRentalID & "_" & iRowCount & """ value=""" & DateValue(oRs("reservationstarttime")) & """ readonly=""readonly"" size=""10"" maxlength=""10"" onclick=""clearMsg('endday_" & sRentalID & "_" & iRowCount & "');doCalendar('startdate_" & sRentalID & "_" & iRowCount & "');"" />&nbsp;" & vbcrlf
       	lcl_return = lcl_return & "        <span class=""calendarimg""><img src=""../images/calendar.gif"" id=""startdate_popup_" & sRentalID & "_" & iRowCount & """ height=""16"" width=""16"" border=""0"" onclick=""clearMsg('endday_" & sRentalID & "_" & iRowCount & "');doCalendar('startdate_" & sRentalID & "_" & iRowCount & "');"" /></span>" & vbcrlf
      		lcl_return = lcl_return & "    </td>" & vbcrlf
      		lcl_return = lcl_return & "    <td align=""center"">" & vbcrlf
    				lcl_return = lcl_return &          ShowHourPicksNew("starthour_" & sRentalID & "_" & iRowCount, GetHourFromDateTime( sReservationStartTime, sAmPm ), sDisabledOption, "endday_" & sRentalID & "_" & iRowCount)  ' Original function in rentalsguifunctions.asp
    				lcl_return = lcl_return &          ":" & vbcrlf
    				lcl_return = lcl_return &          ShowMinutePicksNew("startminute_" & sRentalID & "_" & iRowCount, Minute(sReservationStartTime), sDisabledOption, "endday_" & sRentalID & "_" & iRowCount)	  ' Original function in rentalsguifunctions.asp
    				lcl_return = lcl_return &          " " & vbcrlf
    				lcl_return = lcl_return &          ShowAmPmPicksNew("startampm_" & sRentalID & "_" & iRowCount, sAmPm, sDisabledOption, "endday_" & sRentalID & "_" & iRowCount)	  ' Original function in rentalsguifunctions.asp
    				lcl_return = lcl_return & "    </td>" & vbcrlf
      		lcl_return = lcl_return & "    <td align=""center"">" & vbcrlf
        lcl_return = lcl_return &          ShowHourPicksNew("endhour_" & sRentalID & "_" & iRowCount, GetHourFromDateTime( sReservationEndTime, sAmPm ), sDisabledOption, "endday_" & sRentalID & "_" & iRowCount)	  ' Original function in rentalsguifunctions.asp
      		lcl_return = lcl_return &          ":" & vbcrlf
    				lcl_return = lcl_return &          ShowMinutePicksNew("endminute_" & sRentalID & "_" & iRowCount, Minute(sReservationEndTime), sDisabledOption, "endday_" & sRentalID & "_" & iRowCount)	  ' Original function in rentalsguifunctions.asp
    				lcl_return = lcl_return &          " " & vbcrlf
    				lcl_return = lcl_return &          ShowAmPmPicksNew("endampm_" & sRentalID & "_" & iRowCount, sAmPm, sDisabledOption, "endday_" & sRentalID & "_" & iRowCount)	  ' Original function in rentalsguifunctions.asp
    				lcl_return = lcl_return &          " " & vbcrlf
    				lcl_return = lcl_return &          ShowSameNextDayPickNew("endday_" & sRentalID & "_" & iRowCount, iEndDay, sDisabledOption)	  ' Original function in rentalsguifunctions.asp
    				lcl_return = lcl_return & "    </td>" & vbcrlf
      		lcl_return = lcl_return & "    <td class=""lastcell"" align=""center"">" & vbcrlf
      		lcl_return = lcl_return &           ShowRentalAvailabilityFlagNew(sRentalID, aWantedDates, sPeriodTypeSelector, False)
    				lcl_return = lcl_return & "    </td>" & vbcrlf
    		  lcl_return = lcl_return & "</tr>" & vbcrlf

  		  	'Get rental details for that date
  		  		lcl_return = lcl_return & "<tr>" & vbcrlf
        lcl_return = lcl_return & "    <td class=""firstcell_nopadding"">&nbsp;</td>" & vbcrlf
  		  		lcl_return = lcl_return & "    <td colspan=""2"">" & vbcrlf
  		  		lcl_return = lcl_return &          WeekDayName(Weekday(DateValue(oRs("reservationstarttime"))))
  		  		lcl_return = lcl_return &          " &ndash; " & GetRentalSeason( sRentalID, DateValue(oRs("reservationstarttime")) )
  		  		lcl_return = lcl_return &          GetRentalHours( sRentalID, DateValue(oRs("reservationstarttime")) )
  		  		lcl_return = lcl_return & "    </td>" & vbcrlf

 		  		'Get the other reservations, etc for this date
  		  		lcl_return = lcl_return & "    <td class=""lastcell"" colspan=""2"" valign=""top"">Also happening on this date" & vbcrlf
  		  		lcl_return = lcl_return &          ShowOtherReservationsForDateNew(sRentalID, DateValue(oRs("reservationstarttime")))
  		  		lcl_return = lcl_return & "    </td>" & vbcrlf
  		  		lcl_return = lcl_return & "</tr>" & vbcrlf

      		oRs.MoveNext
     loop
	
    	oRs.Close
    	set oRs = nothing 

    	lcl_return = lcl_return & "<tr><td colspan=""5"" class=""tempseparatorlast"">&nbsp;</td></tr>" & vbcrlf
     lcl_return = lcl_return & "</table>" & vbcrlf
     lcl_return = lcl_return & "<input type=""hidden"" name=""maxrows"                     & sRentalID & """ id=""maxrows"                     & sRentalID & """ size=""3"" maxlength=""10"" value=""" & iRowCount & """ />" & vbcrlf
     lcl_return = lcl_return & "<input type=""hidden"" name=""totalErrors_fieldValidation" & sRentalID & """ id=""totalErrors_fieldValidation" & sRentalID & """ size=""3"" maxlength=""10"" value=""0"" />" & vbcrlf
     lcl_return = lcl_return & "<input type=""hidden"" name=""totalErrors_alertsWarnings"  & sRentalID & """ id=""totalErrors_alertsWarnings"  & sRentalID & """ size=""3"" maxlength=""10"" value=""0"" />" & vbcrlf

  else
     lcl_return = "error"
  end if

  response.write lcl_return

'------------------------------------------------------------------------------
function showHourPicksNew( sSelectName, iHour, sDisabledOption, sClearMsgID)
 	dim x, lcl_return, lcl_onchange

  lcl_return   = ""
  lcl_onchange = ""

	if sDisabledOption = "disabled" then
  		lcl_return = lcl_return & "<input type=""hidden"" id=""" & sSelectName & """ name=""" & sSelectName & """ value=""" & iHour & """ />" & vbcrlf
  		lcl_return = lcl_return & iHour & vbcrlf
	else
     if sClearMsgID <> "" then
        lcl_onchange = " onchange=""clearMsg('" & sClearMsgID & "');"""
     end if

  		lcl_return = lcl_return & "<select id=""" & sSelectName & """ name=""" & sSelectName & """" & lcl_onchange & ">" & vbcrlf

  		for x = 1 to 12
    			if clng(x) = clng(iHour) then
      				lcl_selected_pick = " selected=""selected"""
       else
          lcl_selected_pick = ""
    			end if

		    	lcl_return = lcl_return & "  <option value=""" & x & """" & lcl_selected_pick & ">" & x & "</option>" & vbcrlf
  		next

  		lcl_return = lcl_return & "</select>" & vbcrlf

	end if

 showHourPicksNew = lcl_return

end function

'--------------------------------------------------------------------------------------------------
function showMinutePicksNew(sSelectName, iMinute, sDisabledOption, sClearMsgID)
 	Dim x, sMinutePrefix, lcl_return, lcl_onchange

  sMinutePrefix = ""
  lcl_return    = ""
  lcl_onchange  = ""

 	if sDisabledOption = "disabled" then
	   	if clng(iMinute) < clng(10) then
     			sMinutePrefix = "0"
   		end if

   		lcl_return = lcl_return & right(sMinutePrefix & iMinute,2) & vbcrlf
   		lcl_return = lcl_return & "<input type=""hidden"" id=""" & sSelectName & """ name=""" & sSelectName & """ value=""" & sMinutePrefix & iMinute & """ />" & vbcrlf
  else
     if sClearMsgID <> "" then
        lcl_onchange = " onchange=""clearMsg('" & sClearMsgID & "');"""
     end if

		   lcl_return = lcl_return & "<select id=""" & sSelectName & """ name=""" & sSelectName & """" & lcl_onchange & ">" & vbcrlf

   		for x = 0 to 59
        sMinutePrefix = ""

     			if clng(x) < clng(10) then
       				sMinutePrefix = "0"
     			end if

     			if clng(x) = clng(iMinute) then
       				lcl_selected_minuteprefix = " selected=""selected"""
        else
           lcl_selected_minuteprefix = ""
     			end if

     			lcl_return = lcl_return & "  <option value=""" & sMinutePrefix & x & """" & lcl_selected_minuteprefix & ">" & sMinutePrefix & x & "</option>" & vbcrlf
   		next

   		lcl_return = lcl_return & "</select>" & vbcrlf
	 end if

  showMinutePicksNew = lcl_return

end function

'------------------------------------------------------------------------------
function showAmPmPicksNew(sSelectName, sAmPm, sDisabledOption, sClearMsgID)

  dim lcl_return, lcl_onchange

  lcl_return   = ""
  lcl_onchange = ""

 	if sDisabledOption = "disabled" then
	   	lcl_return = lcl_return & "<input type=""hidden"" id=""" & sSelectName & """ name=""" & sSelectName & """ value=""" & sAmPm & """ onchange=""clearMsg('" & sClearMsgID & "');"" />" & vbcrlf
   		lcl_return = lcl_return & sAmPm & vbcrlf
 	else

     if sClearMsgID <> "" then
        lcl_onchange = " onchange=""clearMsg('" & sClearMsgID & "');"""
     end if

     lcl_selected_am = ""
     lcl_selected_pm = ""

   		if sAmPm = "PM" then
			     lcl_selected_pm = " selected=""selected"""
     else
        lcl_selected_am = " selected=""selected"""
   		end if

	   	lcl_return = lcl_return & "<select id=""" & sSelectName & """ name=""" & sSelectName & """" & lcl_onchange & ">" & vbcrlf
   		lcl_return = lcl_return & "  <option value=""AM""" & lcl_selected_am & ">AM</option>" & vbcrlf
   		lcl_return = lcl_return & "  <option value=""PM""" & lcl_selected_pm & ">PM</option>" & vbcrlf
   		lcl_return = lcl_return & "</select>" & vbcrlf
 	end if

  showAmPmPicksNew = lcl_return

end function

'------------------------------------------------------------------------------
function showSameNextDayPickNew(sSelectName, sDay, sDisabledOption)
  dim lcl_return

  lcl_return = ""

 	if sDisabledOption = "disabled" then
   		lcl_return = lcl_return & "<input type=""hidden"" id=""" & sSelectName & """ name=""" & sSelectName & """ value=""" & sDay & """ />" & vbcrlf

   		if clng(sDay) = clng(0) then
     			lcl_return = lcl_return & "That Day"
   		else
     			lcl_return = lcl_return & "The Next Day"
   		end if
 	else
  			lcl_selected_thatday    = ""
     lcl_selected_thenextday = ""

   		if clng(sDay) = clng(0) then
     			lcl_selected_thatday = " selected=""selected"""
   		end if

   		if clng(sDay) = clng(1) then
     			lcl_selected_thenextday = " selected=""selected"""
   		End If

	   	lcl_return = lcl_return & "<select id=""" & sSelectName & """ name=""" & sSelectName & """ onchange=""clearMsg('" & sSelectName & "');"" >" & vbcrlf
   		lcl_return = lcl_return & "  <option value=""0""" & lcl_selected_thatday & ">That Day</option>" & vbcrlf
   		lcl_return = lcl_return & "  <option value=""1""" & lcl_selected_thenextday & ">The Next Day</option>" & vbcrlf
   		lcl_return = lcl_return & "</select>" & vbcrlf
 	end if

  showSameNextDayPickNew = lcl_return

end function

'------------------------------------------------------------------------------
function showRentalAvailabilityFlagNew( ByVal iRentalid, ByRef aWantedDates, ByVal sPeriodType, ByVal bIsForClass )
 	Dim xbOffSeasonFlag, sFlag, dWantedStartTime, dWantedEndTime
  dim lcl_return

 	sFlag      = "Yes"
  lcl_return = ""

	'Check each day
	 for x = 0 to UBound(aWantedDates,2)
 	  'Set the start and end times
		   dWantedStartTime = cdate(aWantedDates(0,x))
   		dWantedEndTime   = cdate(aWantedDates(1,x))

  		'Is this date in season or off season
   		bOffSeasonFlag = GetOffSeasonFlag( iRentalid, DateValue(dWantedStartTime) )
   		'response.write bOffSeasonFlag

  		'Check if the rental is only available all day on this DOW and if so then adjust the times
   		if RentalIsAllDay( iRentalid, bOffSeasonFlag, Weekday(DateValue(dWantedStartTime)) ) then
     			GetAllDayHours iRentalid, bOffSeasonFlag, Weekday(DateValue(dWantedStartTime)), dWantedStartTime, dWantedEndTime
   		end if

  		'Now check that the wanted date fits into the hours of the rental itself
   		sFlag = CheckRentalHours( iRentalid, dWantedStartTime, dWantedEndTime, sPeriodType, bOffSeasonFlag )

   		if ucase(sFlag) = "YES" then
     		'Finally check if anyone else has a conflicting reservation - buffer time is now optional so do not check it
      		sFlag = CheckForExistingReservations( iRentalid, dWantedStartTime, dWantedEndTime, sPeriodType, bOffSeasonFlag, bIsForClass, False )
   		end if

   		if ucase(sFlag) = "NO" then
    			'When we hit a no We are done with looking
     			exit for
     end if
  next

  lcl_return = sFlag

  showRentalAvailabilityFlagNew = lcl_return

end function

'------------------------------------------------------------------------------
function showOtherReservationsForDateNew(iRentalId, dStartDateTime)
 	dim oShowOtherReservations, sSQL, dWantedEndTime, sFirstName, sLastName, sPhone
  dim lcl_return

  lcl_return = ""

	'This gets anything that starts anytime on the passed date
 	dStartDateTime = dStartDateTime & " 0:00 AM" ' Add the time of midnight to the passed in date
	 dWantedEndTime = DateAdd("d", 1, CDate(dStartDateTime)) ' set this to midnight of the next day

 	sSQL = "SELECT "
  sSQL = sSQL & " D.reservationid, "
  sSQL = sSQL & " D.reservationstarttime, "
  sSQL = sSQL & " D.billingendtime, "
  sSQL = sSQL & " D.reservationendtime, "
  sSQL = sSQL & " T.reservationtype, "
  sSQL = sSQL & " R.timeid, "
  sSQL = sSQL & " T.reservationtypeselector, "
 	sSQL = sSQL & " ISNULL(R.rentaluserid,0) AS rentaluserid, "
  sSQL = sSQL & " ISNULL(R.adminuserid,'') AS adminuserid, "
  sSQL = sSQL & " T.isreservation, "
  sSQL = sSQL & " T.isclass "
 	sSQL = sSQL & " FROM egov_rentalreservationdates D, "
  sSQL = sSQL &      " egov_rentalreservations R, "
  sSQL = sSQL &      " egov_rentalreservationtypes T "
 	sSQL = sSQL & " WHERE D.reservationid = R.reservationid "
  sSQL = sSQL & " AND R.reservationtypeid = T.reservationtypeid "
  sSQL = sSQL & " AND D.rentalid = " & iRentalid
 	sSQL = sSQL & " AND D.statusid IN (SELECT reservationstatusid "
  sSQL = sSQL &                    " FROM egov_rentalreservationstatuses "
  sSQL = sSQL &                    " WHERE iscancelled = 0) "
 	sSQL = sSQL & " AND D.reservationstarttime BETWEEN '" & dStartDateTime & "' AND '" & dWantedEndTime & "' "
  sSQL = sSQL & " AND R.orgid = " & session("orgid")
 	sSQL = sSQL & " ORDER BY D.reservationstarttime"

 	set oShowOtherReservations = Server.CreateObject("ADODB.Recordset")
 	oShowOtherReservations.Open sSQL, Application("DSN"), 0, 1

  if not oShowOtherReservations.eof then
    	do while not oShowOtherReservations.eof
     		'GetTimeFormat is in common.asp
      		lcl_return = lcl_return & "<br /><a href=""reservationedit.asp?reservationid=" & oShowOtherReservations("reservationid") & """ target=""_blank"">"
      		lcl_return = lcl_return & GetTimeFormat(oShowOtherReservations("reservationstarttime")) & " to " & GetTimeFormat(oShowOtherReservations("billingendtime")) & " &ndash; "
      		lcl_return = lcl_return & oShowOtherReservations("reservationtype") & " &ndash; "

      		if oShowOtherReservations("isreservation") then
			        if oShowOtherReservations("reservationtypeselector") = "public" then
          			'Show the citizen name
        	   		lcl_return = lcl_return & ShowShortCitizenNameNew(oShowOtherReservations("rentaluserid"))
        			else
          				lcl_return = lcl_return & GetAdminName(oShowOtherReservations("rentaluserid"))
         		end if
      		else
        			if oShowOtherReservations("isclass") then
          			'Show the class name
       		   		lcl_return = lcl_return & ShowShortClassNameNew(oShowOtherReservations("timeid"))
        			else
          			'The admin who made the hold or block or whatever - in rentlascommonfunctions.asp
           			GetAdminNameAndPhone oShowOtherReservations("adminuserid"), sFirstName, sLastName, sPhone

          				if sLastName <> "" AND sFirstName <> "" then
            					lcl_return = lcl_return & Left(UCase(Left(sFirstName,1)) & ". " & sLastName,30)
          				end if
        			end if
      		end if

      		lcl_return = lcl_return & "</a>" & vbcrlf

      		oShowOtherReservations.movenext
   	 loop
  end if

 	oShowOtherReservations.close
	 set oShowOtherReservations = nothing 

  showOtherReservationsForDateNew = lcl_return

end function

'------------------------------------------------------------------------------
function ShowShortCitizenNameNew(iUserId)
 	Dim oRs, sSQL, lcl_return 

  lcl_return = ""

 	sSQL = "SELECT ISNULL(userlname,' ') AS userlname, "
  sSQL = sSQL & " ISNULL(userfname,' ') AS userfname "
  sSQL = sSQL & " FROM egov_users "
  sSQL = sSQL & " WHERE userid = " & iUserid

 	set oRs = Server.CreateObject("ADODB.Recordset")
 	oRs.Open sSQL, Application("DSN"), 0, 1

 	if not oRs.eof then
   		lcl_return = Left(UCase(Left(oRs("userfname"),1)) & ". " & oRs("userlname"),30)
 	end if

 	oRs.close
	 set oRs = nothing

  showShortCitizenNameNew = lcl_return

end function

'------------------------------------------------------------------------------
function showShortClassNameNew(iTimeId)
 	Dim oRs, sSQL, lcl_return

  lcl_return = ""

 	sSQL = "SELECT C.classname "
  sSQL = sSQL & " FROM egov_class_time T, "
  sSQL = sSQL &      " egov_class C "
 	sSQL = sSQL & " WHERE T.classid = C.classid "
  sSQL = sSQL & " AND T.timeid = " & iTimeId

 	set oRs = Server.CreateObject("ADODB.Recordset")
 	oRs.Open sSQL, Application("DSN"), 0, 1

 	if not oRs.eof then
   		lcl_return = left(ors("classname"),30)
 	end if

 	oRs.close
 	set oRs = nothing

  showShortClassNameNew = lcl_return

end function
%>
