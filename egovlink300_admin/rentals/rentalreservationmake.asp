<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: rentalreservationmake.asp
' AUTHOR: Steve Loar
' CREATED: 10/21/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Creates The initial Rental reservation.
'
' MODIFICATION HISTORY
' 1.0 10/21/2009	Steve Loar - INITIAL VERSION
' 1.1 04/25/2012 David Boyer - Modified to accept multiple rentalids
' 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim iRentalId, sSQL, oRs, iReservationID, iReservationTempID, iRentalUserid, sReservationTypeID
 dim iInitialStatusId, iAdminUserId, iMaxRows, x, sStartDateTime, iEndDay, sBillingEndDateTime
 dim sEndDateTime, iReservationDateId, iWeekday, bIsReservation, sUserType, bPublicReservation
 dim dTotalAmount, iAdminLocationId, iPaymentTypeId, iLedgerId, sPlusMinus, cPriorBalance
 dim iAccountId, iItemTypeID, sReturn, bNewreservation, sSuccessFlag
 dim lcl_selected_rentalids, lcl_unavailable_rentalids
 dim lcl_iorgid

 lcl_selected_rentalids    = ""
 lcl_unavailable_rentalids = ""
 iRentalID                 = 0
 lcl_iorgid = session("orgid")


 if request("selected_rentalids") <> "" then
    if not containsApostrophe(request("selected_rentalids")) then
       lcl_selected_rentalids = request("selected_rentalids")
    end if
 end if

 if request("rentalid") <> "" then
    if not containsApostrophe(request("rentalid")) then
       iRentalID = clng(request("rentalid"))
    end if
 end if

 if lcl_selected_rentalids = "" then
    if iRentalID > 0 then
       lcl_selected_rentalids = iRentalID
    end if
 end if

 if lcl_selected_rentalids <> "" then
   'BEGIN: Check availability one last time -----------------------------------
    sSQL = "SELECT distinct rentalid "
    sSQL = sSQL & " FROM egov_rentals "
    sSQL = sSQL & " WHERE rentalid IN (" & lcl_selected_rentalids & ") "
    sSQL = sSQL & " AND orgid = " & session("orgid")

   	set oGetRentalIDsCheckUnavailable = Server.CreateObject("ADODB.Recordset")
   	oGetRentalIDsCheckUnavailable.Open sSQL, Application("DSN"), 0, 1

    if not oGetRentalIDsCheckUnavailable.eof then
       do while not oGetRentalIDsCheckUnavailable.eof

          'iRentalId = CLng(request("rentalid"))
          'iMaxRows  = CLng(request("maxrows"))
          sReturn   = "OK"
          iMaxRows  = 0
          iRentalID = oGetRentalIDsCheckUnavailable("rentalid")

          if request("maxrows" & iRentalID) <> "" then
             if not containsApostrophe(request("maxrows" & iRentalID)) then
                iMaxRows = clng(request("maxrows" & iRentalID))
             end if
          end if

         'BEGIN: Cycle through reserations checking availability --------------
          for x = 1 to iMaxRows
             sIncludeReservationTime = ""

             if request("includereservationtime_" & iRentalID & "_" & x) <> "" then
                sIncludeReservationTime = request("includereservationtime_" & iRentalID & "_" & x)
             end if

             if sIncludeReservationTime = "Y" then
                lcl_startdate = ""

                if request("startdate_" & iRentalID & "_" & x) <> "" then
                   if not containsApostrophe(request("startdate_" & iRentalID & "_" & x)) then
                      lcl_startdate = request("startdate_" & iRentalID & "_" & x)
                   end if
                end if

                if lcl_startdate <> "" then
					iPosition       = iPosition + 1
					lcl_starthour   = ""
					lcl_startminute = ""
					lcl_startampm   = ""
					lcl_enddate	    = ""
					lcl_endhour     = ""
					lcl_endminute   = ""
					lcl_endampm     = ""
					sStartDateTime  = ""
					sEndDateTime    = ""
					iEndDay         = "0"

                   if request("starthour_" & iRentalID & "_" & x) <> "" then
                      if not containsApostrophe(request("starthour_" & iRentalID & "_" & x)) then
                         lcl_starthour = request("starthour_" & iRentalID & "_" & x)
                      end if
                   end if

                   if request("startminute_" & iRentalID & "_" & x) <> "" then
                      if not containsApostrophe(request("startminute_" & iRentalID & "_" & x)) then
                         lcl_startminute = request("startminute_" & iRentalID & "_" & x)
                      end if
                   end if

                   if request("startampm_" & iRentalID & "_" & x) <> "" then
                      if not containsApostrophe(request("startampm_" & iRentalID & "_" & x)) then
                         lcl_startampm = request("startampm_" & iRentalID & "_" & x)
                      end if
                   end if

                   if request("endhour_" & iRentalID & "_" & x) <> "" then
                      if not containsApostrophe(request("endhour_" & iRentalID & "_" & x)) then
                         lcl_endhour = request("endhour_" & iRentalID & "_" & x)
                      end if
                   end if

                   if request("endminute_" & iRentalID & "_" & x) <> "" then
                      if not containsApostrophe(request("endminute_" & iRentalID & "_" & x)) then
                         lcl_endminute = request("endminute_" & iRentalID & "_" & x)
                      end if
                   end if

                   if request("endampm_" & iRentalID & "_" & x) <> "" then
                      if not containsApostrophe(request("endampm_" & iRentalID & "_" & x)) then
                         lcl_endampm = request("endampm_" & iRentalID & "_" & x)
                      end if
                   end if

                   if request("endday_" & iRentalID & "_" & x) <> "" then
                      if not containsApostrophe(request("endday_" & iRentalID & "_" & x)) then
                       		iEndDay = request("endday_" & iRentalID & "_" & x)
                      end if
                   end If
                   
					 

                 		'sStartDateTime = lcl_startdate & " " & lcl_starthour & ":" & lcl_startminute & " " & lcl_startampm
                 		'sEndDateTime   = lcl_startdate & " " & lcl_endhour   & ":" & lcl_endminute   & " " & lcl_endampm
		
					sStartDateTime = buildReservationDateTime( lcl_startdate, lcl_starthour, lcl_startminute, lcl_startampm )

					If iEndDay = "1" Then
						lcl_enddate = CStr(DateAdd("d", 1, CDate(lcl_startdate)))
						'sEndDateTime = CStr(DateAdd("d", 1, CDate(sEndDateTime)))
					Else
						lcl_enddate = lcl_startdate
					End If
					sEndDateTime = buildReservationDateTime( lcl_enddate, lcl_endhour, lcl_endminute, lcl_endampm )
					'response.write "sEndDateTime: " & sEndDateTime
					'response.end

                 	if sReturn = "OK" then
                  	  	sReturn = CheckRentalAvailability( iRentalid, sStartDateTime, sEndDateTime, False )	 'In rentalscommonfunctions.asp

                     'Now that we have processed the date/times entered we need to check to see if this date has:
                     '  1. already been validated and required user interaction to continue
                     '  2. if "yes" it IS okay to continue then if the "sReturn" matches the errorCode passed in
                     '     AND the current date/time matches the "errorCheck" date/time that was passed in then
                     '     we know that the same date was passed in that was just validated AND that the user "Okayed"
                     '     it, so we do NOT need to send back a message asking them to verify the date again and
                     '     override the "sReturn" value with sReturn + 'OK'.
                     '  3. if "no" it is NOT okay to continue then return with the value in "sReturn"
                      lcl_current_reservationDateTime    = ""
                      lcl_errorCheck_reservationDateTime = ""
                      lcl_errorCheck_errorCode           = ""
                      lcl_errorCheck_okToContinue        = ""

                      lcl_current_reservationDateTime = lcl_startdate
                      lcl_current_reservationDateTime = lcl_current_reservationDateTime & "_"
                      lcl_current_reservationDateTime = lcl_current_reservationDateTime & lcl_starthour & ":" & lcl_startminute & "_" & lcl_startampm
                      lcl_current_reservationDateTime = lcl_current_reservationDateTime & "_"
                      lcl_current_reservationDateTime = lcl_current_reservationDateTime & lcl_endhour & ":" & lcl_endminute & "_" & lcl_endampm
                      lcl_current_reservationDateTime = lcl_current_reservationDateTime & "_"
                      lcl_current_reservationDateTime = lcl_current_reservationDateTime & iEndDay

                      if request("errorCheck_reservationDateTime_" & iRentalID & "_" & x) <> "" then
                         if not containsApostrophe(request("errorCheck_reservationDateTime_" & iRentalID & "_" & x)) then
                            lcl_errorCheck_reservationDateTime = request("errorCheck_reservationDateTime_" & iRentalID & "_" & x)
                         end if
                      end if

                      if request("errorCheck_errorCode_" & iRentalID & "_" & x) <> "" then
                         if not containsApostrophe(request("errorCheck_errorCode_" & iRentalID & "_" & x)) then
                            lcl_errorCheck_errorCode = ucase(request("errorCheck_errorCode_" & iRentalID & "_" & x))
                         end if
                      end if

                      if request("errorCheck_okToContinue_" & iRentalID & "_" & x) <> "" then
                         if not containsApostrophe(request("errorCheck_okToContinue_" & iRentalID & "_" & x)) then
                            lcl_errorCheck_okToContinue = ucase(request("errorCheck_okToContinue_" & iRentalID & "_" & x))
                         end if
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
                 		'else
                   '			exit for
                 		end if

                  'BEGIN: Check for unavailable rentals -----------------------
                 	'Take them to a failed page so they can start over. As is done on the public side.
                   sReturn = ucase(sReturn)

                   if sReturn <> "OK"            AND _
                      sReturn <> "SHORT"         AND _
                      sReturn <> "BUFFERSHORT"   AND _
                      sReturn <> "BUFFER"        AND _
                      sReturn <> "SHORTOK"       AND _
                      sReturn <> "BUFFERSHORTOK" AND _
                      sReturn <> "BUFFEROK" then

                         if lcl_unavailable_rentalids = "" then
                            lcl_unavailable_rentalids = iRentalID
                         else
                            lcl_unavailable_rentalids = lcl_unavailable_rentalids & "," & iRentalID
                         end if

                     	   'response.redirect "rentalunavailable.asp?rentalid=" & iRentalId
                   end if
                  'END: Check for unavailable rentals -------------------------

                end if
             end if
          next
         'BEGIN: Cycle through reserations checking availability --------------

          oGetRentalIDsCheckUnavailable.movenext
       loop
    end if

    oGetRentalIDsCheckUnavailable.close
    set oGetRentalIDsCheckUnavailable = nothing
   'END: Check availability one last time -------------------------------------

   'BEGIN: Create Rental Reservations if ALL reservations are available -------
    iReservationID = ""
    iLineCount     = 0

    if iReservationID = "" OR iReservationID = "0" then
       if request("rid") <> "" then
          iReservationID    = clng(request("rid"))
       end if
    end if

    if lcl_unavailable_rentalids <> "" then
  	    response.redirect "rentalunavailable.asp?rentalids=" & lcl_unavailable_rentalids
    else
       sReservationTypeID = 0
       dTotalAmount       = CDbl(0.0000)

       sSQL = "SELECT distinct rentalid "
       sSQL = sSQL & " FROM egov_rentals "
       sSQL = sSQL & " WHERE rentalid IN (" & lcl_selected_rentalids & ") "
       sSQL = sSQL & " AND orgid = " & session("orgid")

      	set oGetRentalIDsCreateReservations = Server.CreateObject("ADODB.Recordset")
      	oGetRentalIDsCreateReservations.Open sSQL, Application("DSN"), 0, 1

       if not oGetRentalIDsCreateReservations.eof then
          do while not oGetRentalIDsCreateReservations.eof
             iLineCount = iLineCount + 1

           	'BEGIN: Need to know if they are admin or public ------------------
             if request("reservationtypeid") <> "" then
                sReservationTypeID = clng(request("reservationtypeid"))
             end if

             bIsReservation            = IsReservation( sReservationTypeID )	 'in rentalscommonfunctions.asp 
             sReservationTypeSelection = GetReservationTypeSelection( sReservationTypeID )

             'response.write "sReservationTypeSelection = " & sReservationTypeSelection & "<br /><br />"

             if bIsReservation then
                iRentalUserid = 0

                if request("rentaluserid") <> "" then
                   iRentalUserid = clng(request("rentaluserid"))
                end if

               	if sReservationTypeSelection = "public" then
             		    sUserType = GetUserResidentType( iRentalUserid )

                		'If they are not one of these (R, N), we have to figure which they are
                  	if sUserType <> "R" AND sUserType <> "N" then

                   		'This leaves E and B - See if they are a resident, also
                   			sUserType = GetResidentTypeByAddress( iRentalUserid, session("orgid") )
                 		end if
               	else
             		   'Admin type
             		    sUserType = "E" ' employee
               	end if
             else
             	 'Blocked type
             	  iRentalUserid = "NULL"
             	  sUserType     = "E"
             end if
           	'END: Need to know if they are admin or public --------------------

           	'BEGIN: Create the reservation row for new reservations -----------
             iInitialStatusId = GetInitialReservationStatusId()
             iAdminUserId     = session("userid")
             iMaxRows         = 0
             iRentalID        = oGetRentalIDsCreateReservations("rentalid")

             'if iReservationID = "" OR iReservationID = "0" then
             '   if request("rid") <> "" then
             '      iReservationID = clng(request("rid"))
             '   end if
             'end if


             if iReservationID = "" OR iReservationID = "0" then
                sSQL = "INSERT INTO egov_rentalreservations ("
                sSQL = sSQL & "orgid, "
                sSQL = sSQL & "reservationtypeid, "
                sSQL = sSQL & "reservationstatusid, "
                sSQL = sSQL & "rentaluserid, "
                sSQL = sSQL & "adminuserid, "
                sSQL = sSQL & "reserveddate, "
                sSQL = sSQL & "originalrentalid"
                sSQL = sSQL & ") VALUES ("
                sSQL = sSQL & session("orgid")   & ", "
                sSQL = sSQL & sReservationTypeID & ", "
                sSQL = sSQL & iInitialStatusId   & ", "
                sSQL = sSQL & iRentalUserid      & ", "
                sSQL = sSQL & iAdminUserId       & ", "
                sSQL = sSQL & "dbo.GetLocalDate(" & session("orgid") & ",getdate()), "
                sSQL = sSQL & iRentalID
                sSQL = sSQL & ")"

                response.write sSQL & "<br /><br />"

                iReservationID  = RunInsertStatement( sSQL )
                bNewreservation = true
                sSuccessFlag    = "&sf=rc"
             else
           	   'Set values for existing reservations
              	 bNewreservation = false
              	 sSuccessFlag    = "&sf=ru"

               'Because we are in a loop, we ONLY want to pull the total amount from the
               'egov_rentalreservations table on the FIRST pull and NOT on every row in the loop.
               'The "dTotalAmount" will maintain its value through the loop.
                if iLineCount < 2 then
                 	 dTotalAmount = dTotalAmount + GetReservationTotalAmount( iReservationID, "totalamount" )
                end if
             end if
           	'END: Create the reservation row for new reservations -------------

            'BEGIN: Get the dates from the passed values ----------------------
             if request("maxrows" & iRentalID) <> "" then
                if not containsApostrophe(request("maxrows" & iRentalID)) then
                   iMaxRows = clng(request("maxrows" & iRentalID))
                end if
             end if

             for x = 1 to iMaxRows
                sIncludeReservationTime = ""

                if request("includereservationtime_" & iRentalID & "_" & x) <> "" then
                   sIncludeReservationTime = request("includereservationtime_" & iRentalID & "_" & x)
                end if

                if sIncludeReservationTime = "Y" then
                   lcl_startdate = ""

                   if request("startdate_" & iRentalID & "_" & x) <> "" then
                      if not containsApostrophe(request("startdate_" & iRentalID & "_" & x)) then
                         lcl_startdate = request("startdate_" & iRentalID & "_" & x)
                      end if
                   end if

                   if lcl_startdate <> "" then
						iPosition           = iPosition + 1
						lcl_starthour       = ""
						lcl_startminute     = ""
						lcl_startampm       = ""
						lcl_endhour         = ""
						lcl_endminute       = ""
						lcl_endampm         = ""
						sStartDateTime      = ""
						sBillingEndDateTime = ""
						sEndDateTime        = ""
						iEndDay             = ""

                      if request("starthour_" & iRentalID & "_" & x) <> "" then
                         if not containsApostrophe(request("starthour_" & iRentalID & "_" & x)) then
                            lcl_starthour = request("starthour_" & iRentalID & "_" & x)
                         end if
                      end if

                      if request("startminute_" & iRentalID & "_" & x) <> "" then
                         if not containsApostrophe(request("startminute_" & iRentalID & "_" & x)) then
                            lcl_startminute = request("startminute_" & iRentalID & "_" & x)
                         end if
                      end if

                      if request("startampm_" & iRentalID & "_" & x) <> "" then
                         if not containsApostrophe(request("startampm_" & iRentalID & "_" & x)) then
                            lcl_startampm = request("startampm_" & iRentalID & "_" & x)
                         end if
                      end if

                      if request("endhour_" & iRentalID & "_" & x) <> "" then
                         if not containsApostrophe(request("endhour_" & iRentalID & "_" & x)) then
                            lcl_endhour = request("endhour_" & iRentalID & "_" & x)
                         end if
                      end if

                      if request("endminute_" & iRentalID & "_" & x) <> "" then
                         if not containsApostrophe(request("endminute_" & iRentalID & "_" & x)) then
                            lcl_endminute = request("endminute_" & iRentalID & "_" & x)
                         end if
                      end if

                      if request("endampm_" & iRentalID & "_" & x) <> "" then
                         if not containsApostrophe(request("endampm_" & iRentalID & "_" & x)) then
                            lcl_endampm = request("endampm_" & iRentalID & "_" & x)
                         end if
                      end if

                      if request("endday_" & iRentalID & "_" & x) <> "" then
                         if not containsApostrophe(request("endday_" & iRentalID & "_" & x)) then
                          		iEndDay = request("endday_" & iRentalID & "_" & x)
                         end if
                      end if

                    		'sStartDateTime      = lcl_startdate & " " & lcl_starthour & ":" & lcl_startminute & " " & lcl_startampm
                    		'sBillingEndDateTime = lcl_startdate & " " & lcl_endhour   & ":" & lcl_endminute   & " " & lcl_endampm

                      sStartDateTime = buildReservationDateTime( lcl_startdate, lcl_starthour, lcl_startminute, lcl_startampm )

                      If iEndDay = "1" Then 
                       		'sBillingEndDateTime = CStr(DateAdd("d", 1, CDate(sBillingEndDateTime)))
							lcl_enddate = CStr(DateAdd("d", 1, CDate(lcl_startdate)))
					  Else 
							lcl_enddate = lcl_startdate
                      End If 
					  sBillingEndDateTime = buildReservationDateTime( lcl_enddate, lcl_endhour, lcl_endminute, lcl_endampm )

					  bOffSeasonFlag = GetOffSeasonFlag( iRentalID, DateValue(CDate(sStartDateTime)) )
					  'iEndDay = request("endday" & x)

					  'The real end time includes the end buffer if this is not to closing time for the rental
					  if EndTimeIsNotClosingTime( iRentalID, sBillingEndDateTime, bOffSeasonFlag, sStartDateTime ) then

						'Add on the end buffer
						sEndDateTime = AddPostBufferTime( iRentalID, bOffSeasonFlag, sBillingEndDateTime, sStartDateTime )
					  else
						sEndDateTime = sBillingEndDateTime
					  end if

                      sSQL = "INSERT INTO egov_rentalreservationdates ("
                      sSQL = sSQL & "reservationid, "
                      sSQL = sSQL & "rentalid, "
                      sSQL = sSQL & "orgid, "
                      sSQL = sSQL & "statusid, "
                      sSQL = sSQL & "reservationstarttime, "
                      sSQL = sSQL & "reservationendtime, "
                      sSQL = sSQL & "billingendtime, "
                      sSQL = sSQL & "actualstarttime, "
                      sSQL = sSQL & "actualendtime, "
                      sSQL = sSQL & "adminuserid, "
                      sSQL = sSQL & "reserveddate"
                      sSQL = sSQL & ") VALUES ("
                      sSQL = sSQL &       iReservationID      & ", "
                      sSQL = sSQL &       iRentalID           & ", "
                      sSQL = sSQL &       session("orgid")    & ", "
                      sSQL = sSQL &       iInitialStatusId    & ", "
                      sSQL = sSQL & "'" & sStartDateTime      & "', "
                      sSQL = sSQL & "'" & sEndDateTime        & "', "
                      sSQL = sSQL & "'" & sBillingEndDateTime & "', "
                      sSQL = sSQL & "'" & sStartDateTime      & "', "
                      sSQL = sSQL & "'" & sBillingEndDateTime & "', "
                      sSQL = sSQL &       iAdminUserId        & ", "
                      sSQL = sSQL & "dbo.GetLocalDate(" & session("orgid") & ",getdate())"
                      sSQL = sSQL & ")"

                    		response.write sSQL & "<br /><br />"

                    		iReservationDateId = RunInsertStatement( sSQL )

						if bIsReservation then
							iWeekday = Weekday(sStartDateTime)

							'Create the rental reservation date fees rows  
							'If sReservationTypeSelection = "public" Or sReservationTypeSelection = "admin" Then
							CreateRentalReservationDateFees iReservationDateId, iReservationID, iRentalID, bOffSeasonFlag, iWeekday,  sUserType, sStartDateTime, sBillingEndDateTime, dTotalAmount, sReservationTypeSelection
 'if lcl_iorgid = "33" then response.end
							'response.End 
							'End If

							'Create the rental reservation date items rows - These are things like tables and chairs
							CreateRentalReservationDateItems iReservationDateId, iReservationID, iRentalID, sReservationTypeSelection
						end if
                   end if
                end if
             next
            'END: Get the dates from the passed values ------------------------

           	'Create the rental reservation fees on new reservations - These are things like deposits, alcohol surcharge, damages charges
             if sReservationTypeSelection = "public" AND bNewreservation then
               	CreateRentalReservationFees iReservationID, iRentalID, dTotalAmount
             end if

            'If the rental is $0 fees, or this is an internal reservation, and there are no items, then generate the initial payment rows
             If ((sReservationTypeSelection = "public" AND RentalHasNoCosts( iRentalID ) AND NOT RentalHasItems( iRentalID )) OR sReservationTypeSelection = "admin") Then 

          	    'Get the itemtype for the payment
               	iItemTypeID = GetItemTypeId( "rentals" )

             		'Process new reservations 
             		'This is where the admin person is working today
               	if bNewreservation then
                 		dPaymentTotal       = CDbl(0.00) ' Payment total
                 		iJournalEntryTypeID = GetJournalEntryTypeID( "rentalpayment" )
              	   	iAdminLocationId    = 0
                			sPurchaseNotes      = "No Charge Reservation"
                  	iPaymentLocationId  = "1"	 'we do not know this, just pick the first one - that is walk in admin side

                   if session("LocationId") <> "" then
                   			iAdminLocationId = session("LocationId")
                   end if

                 		if sReservationTypeSelection = "admin" then
                   			sPurchaseNotes = "Internal Reservation"
                   end if

                   if sPurchaseNotes <> "" then
                      sPurchaseNotes = dbsafe(sPurchaseNotes)
                      sPurchaseNotes = "'" & sPurchaseNotes & "'"
                   end if

                		'Insert the egov_class_payment row (Journal entry)
                 		sSQL = "INSERT INTO egov_class_payment ("
               		  sSQL = sSQL & "paymentdate, "
               		  sSQL = sSQL & "paymentlocationid, "
               		  sSQL = sSQL & "orgid, "
            		     sSQL = sSQL & "adminlocationid, "
               				sSQL = sSQL & "userid, "
               		  sSQL = sSQL & "adminuserid, "
               		  sSQL = sSQL & "paymenttotal, "
               		  sSQL = sSQL & "journalentrytypeid, "
               		  sSQL = sSQL & "notes, "
            		     sSQL = sSQL & "isforrentals, "
               		  sSQL = sSQL & "reservationid"
               		  sSQL = sSQL & ") VALUES ("
               		  sSQL = sSQL & "dbo.GetLocalDate(" & session("orgid") & ",GetDate()), " 
               				sSQL = sSQL & iPaymentLocationId  & ", "
               		  sSQL = sSQL & session("orgid")    & ", "
            		     sSQL = sSQL & iAdminLocationId    & ", "
               				sSQL = sSQL & iRentalUserid       & ", "
               		  sSQL = sSQL & session("userid")   & ", "
               		  sSQL = sSQL & dPaymentTotal       & ", "
               		  sSQL = sSQL & iJournalEntryTypeID & ", "
               		  sSQL = sSQL & sPurchaseNotes      & ", "
            		     sSQL = sSQL & "1, "
               		  sSQL = sSQL & iReservationID
               		  sSQL = sSQL & ")"

               				response.write sSQL & "<br /><br />"

               				iPaymentId = RunInsertStatement( sSQL )

               			'Now process the payment using the "other" payment type
               				iCitizenAccountId = "NULL"
               				sCheck            = "NULL"
               				sPlusMinus        = "+"
               				cPriorBalance     = "NULL"
            			   	iPaymentTypeId    = GetRentalPaymentTypeId( session("orgid"), "isothermethod" )	' in rentalscommonfunctions.asp
               				iAccountId        = GetPaymentAccountId( session("orgid"), iPaymentTypeId )		' In common.asp

              				'Make the ledger entry for the payment
               				'iLedgerId = MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, sPlusMinus, iItemId, iIsPaymentAccount, iPaymentTypeId, cPriorBalance, iPriceTypeid )
               				'iLedgerId = MakeLedgerEntry( session("orgid"), iAccountId, iPaymentId, CDbl(request("paymentamount" & x)), "NULL", "debit", sPlusMinus, "NULL", 1, x, cPriorBalance, "NULL" )
               				sSQL = "INSERT INTO egov_accounts_ledger ("
               		  sSQL = sSQL & "paymentid, "
            		     sSQL = sSQL & "orgid, "
               		  sSQL = sSQL & "entrytype, "
               		  sSQL = sSQL & "accountid, "
               		  sSQL = sSQL & "amount, "
               		  sSQL = sSQL & "itemtypeid, "
               		  sSQL = sSQL & "plusminus, "
            			   	sSQL = sSQL & "itemid, "
               		  sSQL = sSQL & "ispaymentaccount, "
               		  sSQL = sSQL & "paymenttypeid, "
               		  sSQL = sSQL & "priorbalance, "
               		  sSQL = sSQL & "pricetypeid, "
               		  sSQL = sSQL & "reservationid"
            		     sSQL = sSQL & ") VALUES ("
               				sSQL = sSQL & iPaymentId       & ", "
               		  sSQL = sSQL & session("orgid") & ", "
               		  sSQL = sSQL & "'debit', "
               		  sSQL = sSQL & iAccountId       & ", "
               		  sSQL = sSQL & CDbl(0.00)       & ", "
            		     sSQL = sSQL & "NULL, "
               		  sSQL = sSQL & "'" & sPlusMinus & "', "
               		  sSQL = sSQL & "NULL, "
               		  sSQL = sSQL & "1, "
               		  sSQL = sSQL & iPaymentTypeId   & ", "
               		  sSQL = sSQL & cPriorBalance    & ", "
            		     sSQL = sSQL & "NULL, "
               		  sSQL = sSQL & iReservationID
               		  sSQL = sSQL & ")"

               				response.write sSQL & "<br /><br />"

               				iLedgerId = RunInsertStatement( sSQL )

               			'Make the entry in the egov_verisign_payment_information table
               				'InsertPaymentInformation iPaymentId, iLedgerId, x, CDbl(request("amount" & x)), "APPROVED", sCheck, iCitizenAccountId
               				'InsertPaymentInformation iPaymentId, iLedgerId, iPaymentTypeId, sAmount, sStatus, sCheckNo, iAccountId
               				sSQL = "INSERT INTO egov_verisign_payment_information ("
               		  sSQL = sSQL & "paymentid, "
            		     sSQL = sSQL & "ledgerid, "
               		  sSQL = sSQL & "paymenttypeid, "
               		  sSQL = sSQL & "amount, "
               				sSQL = sSQL & "paymentstatus, "
               		  sSQL = sSQL & "checkno, "
               		  sSQL = sSQL & "citizenuserid"
            		     sSQL = sSQL & ") VALUES ("
               		  sSQL = sSQL & iPaymentId     & ", "
               		  sSQL = sSQL & iLedgerId      & ", " 
               				sSQL = sSQL & iPaymentTypeId & ", "
               		  sSQL = sSQL & CDbl(0.00)     & ", "
               		  sSQL = sSQL & "'APPROVED', "
            		     sSQL = sSQL & sCheck         & ", "
               		  sSQL = sSQL & iCitizenAccountId
               		  sSQL = sSQL & ")"

               				response.write sSQL & "<br /><br />"

               				RunSQLStatement sSQL
                else
                		'Get the existing paymentid for the existing reservation
                 		iPaymentId = GetRentalPaymentId( iReservationID )		' in rentalscommonfunctions.asp
                end if

               'Now pull the daily rates that do not already have ledger account rows and process them
               	sSQL = "SELECT reservationdatefeeid, "
              	 sSQL = sSQL & " reservationdateid "
              	 sSQL = sSQL & " FROM egov_rentalreservationdatefees "
           	    sSQL = sSQL & " WHERE reservationid = " & iReservationID
              		sSQL = sSQL & " AND reservationdatefeeid NOT IN (SELECT reservationfeetypeid "
              	 sSQL = sSQL &                                  " FROM egov_accounts_ledger "
              	 sSQL = sSQL &                                  " WHERE ispaymentaccount = 0 "
              	 sSQL = sSQL &                                  " AND reservationid = " & iReservationID
           	    sSQL = sSQL &                                  ") "

              		response.write sSQL & "<br /><br />"

              		set oRs = Server.CreateObject("ADODB.Recordset")
              		oRs.Open sSQL, Application("DSN"), 0, 1

                if not oRs.eof then
                 		do while not oRs.eof
                 		  'Get the account for the rate
                      iAccountId = GetReservationAccountId( oRs("reservationdatefeeid"), "reservationdatefeeid", "egov_rentalreservationdatefees" )	' In rentalscommonfunctions.asp

                 			 'Add to Accounts Ledger Row
                 		   sSQL = "INSERT INTO egov_accounts_ledger ("
               		     sSQL = sSQL & "paymentid, "
               		     sSQL = sSQL & "orgid, "
         		           sSQL = sSQL & "entrytype, "
         		           sSQL = sSQL & "accountid, "
            		        sSQL = sSQL & "amount, "
               		     sSQL = sSQL & "itemtypeid, "
               		     sSQL = sSQL & "plusminus, "
               		   		sSQL = sSQL & "itemid, "
            		        sSQL = sSQL & "ispaymentaccount, "
         		           sSQL = sSQL & "paymenttypeid, "
            		        sSQL = sSQL & "priorbalance, "
            		        sSQL = sSQL & "reservationid, "
               		     sSQL = sSQL & "reservationfeetypeid, "
               		     sSQL = sSQL & "reservationfeetype, "
            		        sSQL = sSQL & "reservationdateid"
            		        sSQL = sSQL & ") VALUES ("
               		     sSQL = sSQL & iPaymentId                  & ", "
            		        sSQL = sSQL & session("orgid")            & ", "
               		     sSQL = sSQL & "'credit', "
               		     sSQL = sSQL & iAccountId                  & ", "
            		        sSQL = sSQL & CDbl(0.00)                  & ", "
            		        sSQL = sSQL & iItemTypeId                 & ", "
               		     sSQL = sSQL & "'+', "
            		        sSQL = sSQL & iReservationID              & ", "
            		        sSQL = sSQL & "0, "
               		     sSQL = sSQL & "NULL, "
            		        sSQL = sSQL & "NULL, "
            		        sSQL = sSQL & iReservationID              & ", "
               		     sSQL = sSQL & oRs("reservationdatefeeid") & ", "
               		     sSQL = sSQL & "'reservationdatefeeid', "
            		        sSQL = sSQL & oRs("reservationdateid")
            		        sSQL = sSQL & ")"

               		   		response.write sSQL & "<br /><br />"

               		   		RunSQLStatement sSQL

               		   		oRs.movenext
               		  loop
                end if
	
            		  oRs.close
            		  set oRs = nothing
            	end if

             oGetRentalIDsCreateReservations.movenext
          loop
       end if

       oGetRentalIDsCreateReservations.close
       set oGetRentalIDsCreateReservations = nothing
    end if
   'END: Create Rental Reservations if ALL reservations are available ---------

 end if

'BEGIN: Update the total amount due on the reservation ------------------------
 sSQL = "UPDATE egov_rentalreservations SET "
 sSQL = sSQL & " totalamount = " & CDbl(dTotalAmount)
 sSQL = sSQL & " WHERE reservationid = " & iReservationID

 response.write sSQL & "<br /><br />"
 RunSQLStatement sSQL
'END: Update the total amount due on the reservation --------------------------

'BEGIN: delete the temp data --------------------------------------------------
 iReservationTempID = 0

 if request("rti") <> "" then
    iReservationTempID = clng(request("rti"))
 end if

 sSQL = "DELETE FROM egov_rentalreservationdatestemp WHERE reservationtempid = " & iReservationTempID
 response.write sSQL & "<br /><br />"
 RunSQLStatement sSQL

 sSQL = "DELETE FROM egov_rentalreservationstemp WHERE reservationtempid = " & iReservationTempID
 response.write sSQL & "<br /><br />"
 RunSQLStatement sSQL
'END: delete the temp data ----------------------------------------------------

'Take them to the edit page for this reservation
 response.redirect "reservationedit.asp?reservationid=" & iReservationID & sSuccessFlag
%>
