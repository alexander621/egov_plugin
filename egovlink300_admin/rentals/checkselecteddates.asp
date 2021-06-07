<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: checkselecteddates.asp
' AUTHOR: Steve Loar
' CREATED: 10/20/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This checks that the passed dates and times are OK, is called via AJAX
'
' MODIFICATION HISTORY
' 1.0 10/20/2009	Steve Loar - INITIAL VERSION
' 1.1 03/10/2012 David Boyer - Added option to concatenate "rentalid" to front of "sReturn"
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
 dim sSql, oRs, iRentalId, iMaxRows, x, iReservationTempId, iEndDay, sStartDateTime, sEndDateTime
 dim iPosition, iEndHour, iEndMinute, sEndAmPm, sReturn, sOffSeasonFlag, iReservationTypeId, iRentalUserid
 dim lcl_includeRentalID, iIncludeReservationTime, lcl_lineNum
 dim lcl_errorCheck_reservationDateTime, lcl_errorCheck_errorCode, lcl_errorCheck_okToContinue

 iRentalID               = clng(request("rentalid"))
 iMaxRows                = clng(request("maxrows"))
 iReservationTempId      = clng(request("rti"))
 iReservationTypeId      = clng(request("reservationtypeid"))
 iRentalUserid           = clng(request("rentaluserid"))
 iPosition               = 0
 sReturn                 = "OK"
 lcl_includeRentalID     = false
 lcl_id_rentalid         = ""
 sIncludeReservationTime = true
 lcl_lineNum             = 0

 if request("includereservationtime") <> "" then
    sIncludeReservationTime = request("includereservationtime")
 end if

 if request("includeRentalID") <> "" then
    lcl_includeRentalID = request("includeRentalID")
 end if

 if sIncludeReservationTime then
   'Clear out the old date rows
    sSql = "DELETE FROM egov_rentalreservationdatestemp WHERE reservationtempid = " & iReservationTempId
    RunSQLStatement sSql

   'If this is being called from the rentaldateselection.asp then we are already looping through
   'the parameters (i.e. startdate, endday, etc)
    if lcl_includeRentalID then
   	   if request("startdate") <> "" then
        		iPosition      = iPosition + 1
          sStartDate     = request("startdate")
          sStartHour     = request("starthour")
          sStartMinute   = request("startminute")
          sStartAMPM     = request("startampm")
          sEndHour       = request("endhour")
          sEndMinute     = request("endminute")
          sEndAMPM       = request("endampm")
        		iEndDay        = request("endday")
          sStartDateTime = ""
          sEndDateTime   = ""

     	   	'sStartDateTime = request("startdate") & " " & request("starthour") & ":" & request("startminute") & " " & request("startampm")
        		'sEndDateTime   = request("startdate") & " " & request("endhour")   & ":" & request("endminute")   & " " & request("endampm")

          sStartDateTime = buildReservationDateTime(request("startdate"), _
                                                    request("starthour"), _
                                                    request("startminute"), _
                                                    request("startampm"))

          sEndDateTime = buildReservationDateTime(request("startdate"), _
                                                  request("endhour"), _
                                                  request("endminute"), _
                                                  request("endampm"))

        		if iEndDay = "1" then
          			sEndDateTime = CStr(DateAdd("d", 1, CDate(sEndDateTime)))
     	   	end if

       		'Round up as required by the org to the next wanted interval
        		CheckOrgRentalRoundUp sStartDateTime, _
                                sEndDateTime, _
                                iEndHour, _
                                iEndMinute, _
                                sEndAmPm

       		'Save the row to the temp table
        		SaveTempWantedDates iReservationTempId, _
                              sStartDateTime, _
                              sEndDateTime, _
                              iEndDay, _
                              iPosition

         	if sReturn = "OK" then
          			sReturn = CheckRentalAvailability( iRentalid, sStartDateTime, sEndDateTime, false )	 'In rentalscommonfunctions.asp

            'Now that we have processed the date/times entered we need to check to see if this date has:
            '  1. already been validated and required user interaction to continue
            '  2. if "yes" it IS okay to continue then if the "sReturn" matches the errorCode passed in
            '     AND the current date/time matches the "errorCheck" date/time that was passed in then
            '     we know that the same date was passed in that was just validated AND that the user "Okayed"
            '     it, so we do NOT need to send back a message asking them to verify the date again and
            '     override the "sReturn" value with sReturn + 'OK'.
            '  3. if "no" it is NOT okay to continue then return with the value in "sReturn"
             lcl_current_reservationDateTime = request("startdate")
             lcl_current_reservationDateTime = lcl_current_reservationDateTime & "_"
             lcl_current_reservationDateTime = lcl_current_reservationDateTime & request("starthour") & ":" & request("startminute") & "_" & request("startampm")
             lcl_current_reservationDateTime = lcl_current_reservationDateTime & "_"
             lcl_current_reservationDateTime = lcl_current_reservationDateTime & request("endhour") & ":" & request("endminute") & "_" & request("endampm")
             lcl_current_reservationDateTime = lcl_current_reservationDateTime & "_"
             lcl_current_reservationDateTime = lcl_current_reservationDateTime & iEndDay

             lcl_errorCheck_reservationDateTime = request("errorCheck_reservationDateTime")
             lcl_errorCheck_errorCode           = request("errorCheck_errorCode")
             lcl_errorCheck_okToContinue        = request("errorCheck_okToContinue")

             if lcl_errorCheck_errorCode <> "" then
                lcl_errorCheck_errorCode = ucase(lcl_errorCheck_errorCode)
             end if

             if lcl_errorCheck_okToContinue <> "" then
                lcl_errorCheck_okToContinue = ucase(lcl_errorCheck_okToContinue)
             end if

            'Compare the date/time/endday values ("current" with "lastchecked")
             if lcl_current_reservationDateTime = lcl_errorCheck_reservationDateTime then

               'Only check for user input IF the errorCode is one that requires user interaction to continue
                if lcl_errorCheck_errorCode = "SHORT"       OR _
                   lcl_errorCheck_errorCode = "BUFFER"      OR _
                   lcl_errorCheck_errorCode = "BUFFERSHORT" OR _
                   lcl_errorCheck_errorCode = "SHORTOK"     OR _
                   lcl_errorCheck_errorCode = "BUFFEROK"    OR _
                   lcl_errorCheck_errorCode = "BUFFERSHORTOK" then

                  'Verify that the user is "OK to Continue"
                   if lcl_errorCheck_okToContinue = "Y" then
                      sReturn = sReturn & "OK"
                   end if
                end if
             end if
        		end if
       end if

    else
       for x = 1 to iMaxRows
   	      if request("startdate" & x) <> "" then
           		iPosition      = iPosition + 1
           		sStartDateTime = request("startdate" & x) & " " & request("starthour" & x) & ":" & request("startminute" & x) & " " & request("startampm" & x)
     	      	iEndDay        = request("endday"    & x)
        	   	sEndDateTime   = request("startdate" & x) & " " & request("endhour" & x) & ":" & request("endminute" & x) & " " & request("endampm" & x)

           		If iEndDay = "1" Then 
             			sEndDateTime = CStr(DateAdd("d", 1, CDate(sEndDateTime)))
     	      	End If 

          		'Round up as required by the org to the next wanted interval
           		CheckOrgRentalRoundUp sStartDateTime, _
                                   sEndDateTime, _
                                   iEndHour, _
                                   iEndMinute, _
                                   sEndAmPm

          		'Save the row to the temp table
           		SaveTempWantedDates iReservationTempId, _
                                 sStartDateTime, _
                                 sEndDateTime, _
                                 iEndDay, _
                                 iPosition

            	if sReturn = "OK" then
             			sReturn = CheckRentalAvailability( iRentalid, sStartDateTime, sEndDateTime, False )	' In rentalscommonfunctions.asp
     	      	end if
          end if
       next
    end if

   'Do a check for if a rentaluserid is needed and not picked. This is for the "Check and Reserve" button only.
    if iReservationTypeId > CLng(0) AND IsReservation( iReservationTypeId ) AND iRentalUserid = CLng(0) then
	      sReturn = "nouser"
    end if
 end if

 if sReturn = "" then
	   sReturn = "OK"
 end if

'When validating the rentals and date selection(s) on rentaldateselection.asp there may be multiple rentals to
'validate.  Including the rentalID with the return will give us the one we are working with so that if there is
'an error we know which one to show the error with.  Before this sceen only worked with a single rental location.
'Now it can display multiple rental locations.  We are also going to include the "line number" (above) so that we
'can know specifically which row we are on to help with javascript validation on the return.
'Return value format: [RentalID]_[lineNum]_[errorCode]
 if lcl_includeRentalID then

   'These return codes require input from the user so we need to identify which date they refer to.
    'if sReturn = "short" OR sReturn = "buffer" OR sReturn = "buffershort" then
    '   sReturn = sReturn & request("startdate")
    'end if

    if request("linenum") <> "" then
       if not containsApostrophe(request("linenum")) then
          lcl_lineNum = clng(request("linenum"))
       end if
    end if

   'Add the lineNumber to the results.  This tells us specifically which row we are working with.
    sReturn = lcl_lineNum & "_" & sReturn

   'Add the rentalID to the results.  This tells us which rental row we are working with.
    sReturn = iRentalID & "_" & sReturn
 end if

 response.write sReturn
%>
