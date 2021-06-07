<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: reservationtimechangedo.asp
' AUTHOR: Steve Loar
' CREATED: 06/16/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This makes time changes to reservations. It is called from reservationtimechange.asp
'
' MODIFICATION HISTORY
' 1.0   06/16/2010	Steve Loar - INITIAL VERSION
' 1.1	10/08/2010	Steve Loar - Added actual start and end times to update for Menlo Park
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, iReservationDateId, iReservationId, iRentalId, sStartDateTime, iEndDay, bOffSeasonFlag
Dim sBillingEndDateTime, sEndDateTime, iReservationTypeId, sReservationTypeSelection, iRentalUserid
Dim iWeekday, dTotalCharges

If request("rdi") = "" Then 
	response.redirect "reservationlist.asp"
End If 

If Not IsNumeric(request("rdi")) Then
	response.redirect "reservationlist.asp"
End If 

iReservationDateId = CLng(request("rdi"))

iReservationId = CLng(request("reservationid"))

iRentalId = CLng(request("rentalid"))

sStartDateTime = request("startdate") & " " & request("starthour") & ":" & request("startminute") & " " & request("startampm")
response.write "sStartDateTime" & x & " = " & sStartDateTime & "<br /><br />"

iEndDay = request("endday")

bOffSeasonFlag = GetOffSeasonFlag( iRentalid, DateValue(CDate(sStartDateTime)) )

sBillingEndDateTime = request("startdate") & " " & request("endhour") & ":" & request("endminute") & " " & request("endampm")

If iEndDay = "1" Then 
	sBillingEndDateTime = CStr(DateAdd("d", 1, CDate(sBillingEndDateTime)))
End If 

If ReservationNeedsBufferTimeAdded( iReservationId ) Then 
	' The real end time includes the end buffer if this is not to closing time for the rental
	If EndTimeIsNotClosingTime( iRentalId, sBillingEndDateTime, bOffSeasonFlag, sStartDateTime ) Then 
		' Add on the end buffer
		sEndDateTime = AddPostBufferTime( iRentalid, bOffSeasonFlag, sBillingEndDateTime, sStartDateTime )
	Else
		sEndDateTime = sBillingEndDateTime
	End If 
Else
	sEndDateTime = sBillingEndDateTime
End If 

sSql = "UPDATE egov_rentalreservationdates SET "
sSql = sSql & "reservationstarttime = '" & sStartDateTime & "', "
sSql = sSql & "reservationendtime = '" & sEndDateTime & "', "
sSql = sSql & "actualstarttime = '" & sStartDateTime & "', "
sSql = sSql & "actualendtime = '" & sBillingEndDateTime & "', "
sSql = sSql & "billingendtime = '" & sBillingEndDateTime & "' "
sSql = sSql & "WHERE reservationdateid = " & iReservationDateId
sSql = sSql & " AND orgid = " & session("orgid")
response.write sSql & "<br /><br />"
RunSQLStatement sSql

iReservationTypeId = GetReservationTypeId( iReservationId )

bIsReservation = IsReservation( iReservationTypeId )
sReservationTypeSelection = GetReservationTypeSelection( iReservationTypeId )
response.write "sReservationTypeSelection = " & sReservationTypeSelection & "<br /><br />"

' If this is a reservation type (internal or public) then we need to update the date fees and the reservation total
If bIsReservation Then 
	' Need to know if they are admin or public
	iRentalUserid =  GetReservationRentalUserId( iReservationId )	' In rentalscommonfunctions.asp
	If sReservationTypeSelection = "public" Then 
		sUserType = GetUserResidentType( iRentalUserid )
		'If they are not one of these (R, N), we have to figure which they are
		If sUserType <> "R" And sUserType <> "N" Then 
			'This leaves E and B - See if they are a resident, also
			sUserType = GetResidentTypeByAddress( iRentalUserid, Session("OrgID") )
		End If 
	Else
		' Admin type
		sUserType = "E" ' employee
	End If 

	iWeekday = Weekday(sStartDateTime)

	ResetReservationDateFees iReservationDateId, iReservationId, iRentalid, bOffSeasonFlag, iWeekday, sUserType, sStartDateTime, sBillingEndDateTime, sReservationTypeSelection

	dTotalCharges = CalculateReservationTotal( iReservationId, "feeamount" )	' in rentalscommonfunctions.asp

	' Update the total for the reservation
	sSql = "UPDATE egov_rentalreservations SET totalamount = " & CDbl(dTotalCharges)
	sSql = sSql & " WHERE reservationid = " & iReservationId
	response.write sSql & "<br /><br />"
	RunSQLStatement sSql

End If 

		
'response.write "Done."
response.redirect "reservationedit.asp?reservationid=" & iReservationId & "&sf=tc"




'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void ResetReservationDateFees iReservationDateId, iReservationId, iRentalid, bOffSeasonFlag, iWeekday, sUserType, sStartDateTime, sEndDateTime, sReservationTypeSelection
'--------------------------------------------------------------------------------------------------
Sub ResetReservationDateFees( ByVal iReservationDateId, ByVal iReservationId, ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday, ByVal sUserType, ByVal sStartDateTime, ByVal sEndDateTime, ByVal sReservationTypeSelection )
	Dim sSql, oRs, sNeedFee, iPaymentId, iItemTypeID, sAccount

	' Get the Rental Reservation Fees for the day
	sSql = "SELECT R.pricetypeid, P.pricetypename, ISNULL(accountid,0) AS accountid, R.ratetypeid, ISNULL(amount,0.00) AS amount, "
	sSql = sSql & "ISNULL(R.starthour,0) AS starthour, dbo.AddLeadingZeros(ISNULL(R.startminute,0),2) AS startminute, "
	sSql = sSql & "ISNULL(R.startampm,'AM') AS startampm, P.pricetype, P.isbaseprice, P.isfee, P.hasstarttime, P.isweekendsurcharge, "
	sSql = sSql & "ISNULL(P.basepricetypeid,0) AS basepricetypeid, P.checkresidency, P.isresident, T.datediffstring, alwaysadd "
	sSql = sSql & "FROM egov_rentaldayrates R, egov_rentaldays D, egov_price_types P, egov_rentalratetypes T "
	sSql = sSql & "WHERE D.dayid = R.dayid AND D.rentalid = R.rentalid AND R.pricetypeid = P.pricetypeid "
	sSql = sSql & " AND T.ratetypeid = R.ratetypeid AND D.rentalid = " & iRentalid
	sSql = sSql & " AND D.isoffseason = " & bOffSeasonFlag & " AND D.dayofweek = " & iWeekday
	response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If oRs("accountid") = CLng(0) Then 
			sAccount = "NULL"
		Else
			sAccount = oRs("accountid")
		End If 
		
		If oRs("isfee") Then	
			' These should be the weekend surcharge
			If oRs("hasstarttime") Then
				sHour = oRs("starthour")
				sMinute = oRs("startminute")
				sAmPm = "'" & oRs("startampm") & "'"
				sAmPmValue = oRs("startampm")
				response.write "sSurchargeStart = " & DateValue(sStartDateTime) & " " & sHour & ":" & sMinute & " " & sAmPmValue & "<br /><br />"
				sSurchargeStart = CDate(DateValue(sStartDateTime) & " " & sHour & ":" & sMinute & " " & sAmPmValue)
			Else
				sSurchargeStart = sStartDateTime
			End If 

			'response.write DateDiff( "n", CDate(sEndDateTime), CDate(sSurchargeStart)) & "<br /><br />"
			If DateDiff( "n", CDate(sEndDateTime), CDate(sSurchargeStart)) < 0 Then
				If ReservationDateHasFee( iReservationDateId, oRs("pricetypeid") ) Then
					sNeedFee = "update" 
				Else
					sNeedFee = "insert"
				End If 
				'response.write DateDiff( "n", CDate(sStartDateTime), CDate(sSurchargeStart)) & "<br /><br />"
				If DateDiff( "n", CDate(sStartDateTime), CDate(sSurchargeStart)) < 0 Then
					iDuration = DateDiff("n", CDate(sStartDateTime), CDate(sEndDateTime))
					If oRs("datediffstring") = "h" Then
						iDuration = CDbl(iDuration / 60)
					End If 
					sRate = CDbl(oRs("amount"))
					iFeeAmount = CDbl(iDuration) * sRate
				Else
					iDuration = DateDiff("n", sSurchargeStart, CDate(sEndDateTime))
					If oRs("datediffstring") = "h" Then
						iDuration = CDbl(iDuration / 60)
					ElseIf oRs("datediffstring") = "d" Then
						iDuration = CDbl(1.00)
					End If 
					sRate = CDbl(oRs("amount"))
					iFeeAmount = CDbl(iDuration) * sRate
				End If 
			Else
				If ReservationDateFeeHasPayment( iReservationDateId, oRs("pricetypeid") ) And sReservationTypeSelection = "public"  Then
					iDuration = CDbl(0.00)
					sRate = CDbl(oRs("amount"))
					iFeeAmount = CDbl(0.00)
					sNeedFee = "update" 
				Else 
					sNeedFee = "delete" 
				End If 
			End If 
		Else 
			sHour = "NULL"
			sMinute = "NULL"
			sAmPm = "NULL"

			' These are evey other type of date fee and they should be in the system to be updated
			If ReservationDateHasFee( iReservationDateId, oRs("pricetypeid") ) Then 
				' They already have these fees so we just want to update them
				iDuration = DateDiff("n", CDate(sStartDateTime), CDate(sEndDateTime))
				If oRs("datediffstring") = "h" Then
					iDuration = CDbl(iDuration / 60)
				ElseIf oRs("datediffstring") = "d" Then
					iDuration = CDbl(1.00)
				End If 
				sRate = CDbl(oRs("amount"))
				iFeeAmount = CDbl(iDuration) * sRate
				sNeedFee = "update"
			Else
				sNeedFee = "ignore"
			End If 
		End If 

		If sNeedFee <> "ignore" Then 
			' The internal reservations are always $0.00
			If sReservationTypeSelection = "admin" Then 
				iFeeAmount = CDbl(0.00)
				sRate = CDbl(0.00) 
				
			End If 

			Select Case sNeedFee
				Case "update" 
					iReservationDateFeeId = GetReservationDateFeeId( iReservationDateId, oRs("pricetypeid") )
					' Update the existing fee
					sSql = "UPDATE egov_rentalreservationdatefees "
					sSql = sSql & "SET amount = " & sRate & ", "
					sSql = sSql & "feeamount = " & iFeeAmount & ", "
					sSql = sSql & "duration = " & iDuration & ", "
					sSql = sSql & "starthour = " & sHour & ", "
					sSql = sSql & "startminute = " & sMinute & ", "
					sSql = sSql & "startampm = " & sAmPm 
					sSql = sSql & " WHERE reservationdatefeeid = " & iReservationDateFeeId
					response.write sSql & "<br /><br />"
					RunSQLStatement sSql

				Case "insert"
					' Insert the new fee
					sSql = "INSERT INTO egov_rentalreservationdatefees (reservationdateid, reservationid, rentalid, pricetypeid, "
					sSql = sSql & "accountid, ratetypeid, amount, starthour, startminute, startampm, feeamount, paidamount, "
					sSql = sSql & "refundamount, duration, datediffstring ) VALUES ( " & iReservationDateId & ", " & iReservationId & ", " & iRentalid & ", "
					sSql = sSql & oRs("pricetypeid") & ", " & sAccount & ", " & oRs("ratetypeid") & ", " & sRate & ", "
					sSql = sSql & sHour & ", " & sMinute & ", " & sAmPm & ", " & iFeeAmount & ", 0.0000, 0.0000, "
					sSql = sSql & iDuration & ", '" & oRs("datediffstring") & "' )"
					response.write sSql & "<br /><br />"
					iReservationDateFeeId = RunInsertStatement( sSql )

					If sReservationTypeSelection = "admin" Then 
						' Get Key Values needed for the account ledger entry
						iPaymentId = GetReservationPaymentId( iReservationId )
						iItemTypeID = GetItemTypeId( "rentals" )	' in rentalscommonfunctions.asp

						' Insert into the accounts ledger table 
						sSql = "INSERT Into egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, "
						sSql = sSql & "plusminus, itemid, ispaymentaccount, paymenttypeid, priorbalance, reservationid, "
						sSql = sSql & "reservationfeetypeid, reservationfeetype, reservationdateid ) VALUES ( "
						sSql = sSql & iPaymentId & ", " & session("orgid") & ", 'credit', " & sAccount & ", " & CDbl(0.00) & ", "
						sSql = sSql & iItemTypeId & ", '+', " & iReservationId & ", 0, NULL, NULL, " & iReservationId & ", "
						sSql = sSql & iReservationDateFeeId & ", 'reservationdatefeeid', " & iReservationDateId & " )"
						response.write sSql & "<br /><br />"
						RunSQLStatement sSql
					End If 

				Case "delete"
					iReservationDateFeeId = GetReservationDateFeeId( iReservationDateId, oRs("pricetypeid") )

					If iReservationDateFeeId > CLng(0) Then 
						If sReservationTypeSelection = "admin" Then 
							' Delete from the accounts ledger table 
							sSql = "DELETE FROM egov_accounts_ledger WHERE reservationfeetype = 'reservationdatefeeid' AND "
							sSql = sSql & "reservationfeetypeid = " & iReservationDateFeeId
							response.write sSql & "<br /><br />"
							RunSQLStatement sSql
						End If 

						' Delete the existing fee - This will only be the surcharges
						sSql = "DELETE FROM egov_rentalreservationdatefees WHERE reservationdatefeeid = " & iReservationDateFeeId
						response.write sSql & "<br /><br />"
						RunSQLStatement sSql
					End If 

			End Select 
		End If 

		oRs.MoveNext 
	Loop 

	oRs.Close 
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer GetReservationDateFeeId( iReservationDateId, iPriceTypeId )
'--------------------------------------------------------------------------------------------------
Function GetReservationDateFeeId( ByVal iReservationDateId, ByVal iPriceTypeId )
	Dim sSql, oRs

	sSql = "SELECT reservationdatefeeid FROM egov_rentalreservationdatefees WHERE reservationdateid = " & iReservationDateId
	sSql = sSql & " AND pricetypeid = " & iPriceTypeId
	response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationDateFeeId = CLng(oRs("reservationdatefeeid"))
	Else
		GetReservationDateFeeId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' integer GetReservationPaymentId iReservationId
'--------------------------------------------------------------------------------------------------
Function GetReservationPaymentId( ByVal iReservationId )
	Dim sSql, oRs

	sSql = "SELECT paymentid FROM egov_class_payment WHERE reservationid = " & iReservationId
	response.write sSql & "<br /><br />"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationPaymentId = CLng(oRs("paymentid"))
	Else
		GetReservationPaymentId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' boolean ReservationDateHasFee( iReservationDateId, iPriceTypeId )
'--------------------------------------------------------------------------------------------------
Function ReservationDateHasFee( ByVal iReservationDateId, ByVal iPriceTypeId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(reservationdatefeeid) AS hits FROM egov_rentalreservationdatefees "
	sSql = sSql & "WHERE reservationdateid = " & iReservationDateId
	sSql = sSql & " AND pricetypeid = " & iPriceTypeId
	response.write sSql & "<br /><br />"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then 
			ReservationDateHasFee = True
		Else
			ReservationDateHasFee = False
		End If 
	Else
		ReservationDateHasFee = False
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean ReservationDateFeeHasPayment( iReservationDateId, iPriceTypeId )
'--------------------------------------------------------------------------------------------------
Function ReservationDateFeeHasPayment( ByVal iReservationDateId, ByVal iPriceTypeId )
	Dim sSql, oRs

	sSql = "SELECT paidamount FROM egov_rentalreservationdatefees "
	sSql = sSql & "WHERE reservationdateid = " & iReservationDateId
	sSql = sSql & " AND pricetypeid = " & iPriceTypeId
	response.write sSql & "<br /><br />"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CDbl(oRs("paidamount")) > CDbl(0.00) Then 
			ReservationDateFeeHasPayment = True
		Else
			ReservationDateFeeHasPayment = False
		End If 
	Else
		ReservationDateFeeHasPayment = False
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 




%>