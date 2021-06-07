<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: classMembership.asp
' AUTHOR: Steve Loar
' CREATED: 08/07/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the membership class
'
' MODIFICATION HISTORY
' 1.0  08/07/06 Steve Loar - Initial code
' 1.1  09/11/08 David Boyer - Added Membership Renewals
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Class classMembership
	
	Private Sub Class_Initialize()
	End Sub 

	'------------------------------------------------------------------------------
	Public Function GetMembershipId( ByVal sMembership )
		Dim sSql, oRs

		sSql = "SELECT membershipid FROM egov_memberships WHERE orgid = " & iOrgId & " AND membership = '" & sMembership & "' "

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then 
			GetMembershipId = oRs("membershipid") 
		Else
			GetMembershipId = 0
		End If
			
		oRs.Close
		Set oRs = Nothing
		
	End Function

	
	'----------------------------------------------------------------------------------------
	' Public Function Get MembershipPeriod()
	'----------------------------------------------------------------------------------------
	Public Function GetMembershipPeriod()
		Dim sSql, oRs

		sSql = "SELECT ISNULL(period_desc,'') AS period_desc FROM membership_periods WHERE orgid = " & iOrgId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then 
			MembershipPeriod = oRs("period_desc") 
		Else
			MembershipPeriod = ""
		End If

		oRs.close
		Set oRs = Nothing
		
	End Function   


	'----------------------------------------------------------------------------------------
	' Public Function GetPeriodId( sPeriod )
	'----------------------------------------------------------------------------------------
	Public Function GetPeriodId( ByVal sPeriod )
		Dim sSql, oRs

		sSql = "SELECT periodid FROM egov_membership_periods WHERE orgid = " & iOrgId & " AND period_desc = '" & sPeriod & "' "

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then 
			GetPeriodId = oRs("periodid") 
		Else
			GetPeriodId = 0
		End If

		oRs.close
		Set oRs = Nothing
		
	End Function


	'----------------------------------------------------------------------------------------
	' Public Function HasActiveMembership( iUserId )
	'----------------------------------------------------------------------------------------
	Public Function HasActiveMembership( ByVal iUserId )
		Dim sSql, oRs

		sSql = "SELECT P.paymentdate, period_desc FROM egov_poolpasspurchases p, egov_membership_periods M "
		sSql = sSql & " WHERE P.userid = " & iUserId & " AND P.orgid = " & iOrgid & " AND (P.paymentresult = 'Paid' OR P.paymentresult = 'APPROVED') "
		sSql = sSql & " AND P.periodid = M.periodid ORDER BY P.paymentdate DESC"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then 
			If LCase(oRs("period_desc")) = "season" Then 
				' Season memberships are only good for the year bought
				If Year( oRs("paymentdate")) = Year(Date()) Then 
					IsActiveMember = True 
				Else
					IsActiveMember = False 
				End If 
			End If 
		Else
			IsActiveMember = False 
		End If
			
		oRs.close
		Set oRs = Nothing	
		
	End Function 


	'------------------------------------------------------------------------------
	public sub MembershipPurchase( ByVal iUserId, ByVal iorgid, ByVal iRateId, ByVal nAmount, ByVal iMembershipID, ByVal iPeriodID, _
			ByVal sPaymentType, ByVal sPaymentLocation, ByVal sPoolPassID, ByVal lcl_startdate, ByRef iPassID )
		Dim sResult, sSql, oRs

		if LCASE(sPaymentType) = "creditcard" then
			sResult = "Pending"
		else
			sResult = "Paid"
		end if

		if lcl_startdate = "" then
			lcl_startdate = Date()
		end if

		lcl_expirationdate = getExpirationDate( iPeriodID, lcl_startdate )

		sSql = "INSERT INTO egov_poolpasspurchases ( "
		sSql = sSql & "userid, "
		sSql = sSql & "orgid, "
		sSql = sSql & "rateid, "
		sSql = sSql & "membershipid, "
		sSql = sSql & "periodid, "
		sSql = sSql & "paymentamount, "
		sSql = sSql & "paymenttype, "
		sSql = sSql & "paymentlocation, "
		sSql = sSql & "paymentresult, "
		sSql = sSql & "startdate, "
		sSql = sSql & "expirationdate "

		if sPoolPassID <> "" then
			sSql = sSql & ", previous_poolpassid "
		end if

		sSql = sSql & ")	VALUES ("
		sSql = sSql & iUserID                  & ", "
		sSql = sSql & iorgid                   & ", "
		sSql = sSql & iRateID                  & ", "
		sSql = sSql & iMembershipID            & ", "
		sSql = sSql & iPeriodID                & ", "
		sSql = sSql & nAmount                  & ", "
		sSql = sSql & "'" & sPaymentType       & "', "
		sSql = sSql & "'" & sPaymentLocation   & "', "
		sSql = sSql & "'" & sResult            & "', "
		sSql = sSql & "'" & lcl_startdate      & "', "
		sSql = sSql & "'" & lcl_expirationdate & "' "

		if sPoolPassID <> "" then
			sSql = sSql & ", " & sPoolPassID
		end if

		sSql = sSql & ")"

		'set oRs = Server.CreateObject("ADODB.Recordset")
		'oRs.Open sSql, Application("DSN"), 3, 1

		'Retrieve the poolpassid that was just inserted
		'sSql = "SELECT IDENT_CURRENT('egov_poolpasspurchases') as NewID"
		
		'oRs.Open sSql, Application("DSN"), 3, 1
		'iPassID = oRs("NewID").value

		'set oRs = nothing
		
		iPassID = RunIdentityInsertStatement( sSql )

	end sub


	'--------------------------------------------------------------------------------------------------
	public sub AddMember( ByVal iMembershipId, ByVal iFamilymemberId, ByVal iPrevPoolPassID, ByVal iIsPunchcard, ByVal iPunchcardLimit )
		Dim sSql, oRs, lcl_member_id, lcl_card_printedm, lcl_printed_count
		
		lcl_member_id     = ""
		lcl_card_printed  = "N"
		lcl_printed_count = 0

		if iPrevPoolPassID <> "" then
			'Get the current member data
			getCurrentMemberInfo iPrevPoolPassID, iFamilyMemberID, lcl_member_id, lcl_card_printed, lcl_printed_count
		else
			lcl_member_id = getNextMemberID()
		end if

		if iIsPunchcard then
			lcl_isPunchcard = 1
		else
			lcl_isPunchcard = 0
		end if

		'Insert new member records
		sSql = "INSERT INTO egov_poolpassmembers ("
		sSql = sSql & "poolpassid, "
		sSql = sSql & "familymemberid, "
		sSql = sSql & "memberid, "
		sSql = sSql & "card_printed, "
		sSql = sSql & "printed_count, "
		sSql = sSql & "isPunchcard, "
		sSql = sSql & "punchcard_limit, "
		sSql = sSql & "pcard_remaining_cnt "
		sSql = sSql & ") VALUES ("
		sSql = sSql & iMembershipID     & ", "
		sSql = sSql & iFamilyMemberID   & ", "
		sSql = sSql & lcl_member_id     & ", "
		sSql = sSql & "'" & lcl_card_printed  & "', "
		sSql = sSql & lcl_printed_count & ", "
		sSql = sSql & lcl_isPunchcard   & ", "
		sSql = sSql & iPunchcardLimit   & ", "
		sSql = sSql & iPunchcardLimit
		sSql = sSql & ")"

		RunSQLStatement sSql 

		'set oRs = Server.CreateObject("ADODB.Recordset")
		'oRs.Open sSql, Application("DSN"), 3, 1

		'set oRs = nothing

	end sub


	'-----------------------------------------------------------------------------
	Public Function PublicCanPurchase( ByVal sMembership )
		Dim sSql, oRs

		sSql = "SELECT publicpurchase FROM egov_memberships WHERE orgid = " & iOrgId & " AND membership = '" & sMembership & "' "

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then 
			PublicCanPurchase = oRs("publicpurchase") 
		Else
			PublicCanPurchase = False 
		End If
			
		oRs.close
		Set oRs = Nothing
		
	End Function 


	'-----------------------------------------------------------------------------
	public function RatesAreVisible( ByVal sMembership, ByVal sUserType )
		'This is to see if the rate types they can buy are visible
		Dim sSql, oRs

		RatesAreVisible = False

		sSql = "SELECT public_display "
		sSql = sSql & " FROM egov_membership_rate_displays R, egov_memberships M "
		sSql = sSql & " WHERE orgid = " & iOrgid 
		sSql = sSql & " AND resident_type = '" & sUserType & "' "
		sSql = sSql & " AND membership = '" & sMembership & "' "
		sSql = sSql & " AND R.membershipid = M.membershipid "

		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		if not oRs.EOF then
			RatesAreVisible = oRs("public_display")
		end if

		oRs.close
		set oRs = nothing

	end function


	'----------------------------------------------------------------------------------------
	' Public Sub ShowMembershipIntro()
	'----------------------------------------------------------------------------------------
	Public Sub ShowMembershipIntro( ByVal sMembership )
		Dim sSql, oRs

		sSql = "SELECT ISNULL(introtext,'') AS introtext FROM egov_memberships WHERE orgid = " & iOrgId & " AND membership = '" & sMembership & "' "

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then 
			response.write oRs("introtext") 
		End If

		oRs.close
		Set oRs = Nothing
		
	End Sub


	'------------------------------------------------------------------------------
	public sub ShowMembershipResidencyRates( ByVal iMembershipId, ByVal sUserType, ByVal iPeriodId )
		Dim sSql, oRs

		sSql = "SELECT R.resident_type, P.description "
		sSql = sSql & " FROM egov_poolpassresidenttypes P, "
		sSql = sSql & " egov_membership_rate_displays R, "
		sSql = sSql & " egov_memberships M "
		sSql = sSql & " WHERE R.resident_type = P.resident_type "
		sSql = sSql & " AND M.membershipid = R.membershipid "
		sSql = sSql & " AND M.orgid = P.orgid "
		sSql = sSql & " AND R.public_display = 1 "
		sSql = sSql & " AND R.resident_type IN (SELECT residenttype "
		sSql = sSql & " FROM egov_poolpassrates "
		sSql = sSql & " WHERE membershipid = " & iMembershipId
		sSql = sSql & " AND periodid = "       & iPeriodId & ") "
		sSql = sSql & " AND M.orgid = "          & iOrgId
		sSql = sSql & " AND M.membershipid= "    & iMembershipId
		sSql = sSql & " AND R.resident_type = '" & sUserType & "' "
		sSql = sSql & " ORDER BY displayorder "

		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1
		if oRs.EOF then
			response.write "  No Rates Available" & vbcrlf
		end if

		do while not oRs.eof
			response.write "<div class=""picktitle"">" & oRs("description") & "</div>" & vbcrlf
			response.write "<div class=""pickchoice"">" & vbcrlf
			response.write "  <table class=""picktable"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbcrlf

			'Get the pickable rates here 
			ShowResidentRates iMembershipId, oRs("resident_type"), sUserType, iPeriodId

			response.write "  </table>" & vbcrlf
			response.write "</div>" & vbcrlf

			oRs.movenext
		loop

		oRs.close
		set oRs = nothing

	end sub


	'------------------------------------------------------------------------------
	private sub ShowResidentRates( ByVal iMembershipId, ByVal sResidenttype, ByVal sUserType, ByVal iPeriodId)
		Dim sDisabled, sSql, oRs, iRow
		
		sDisabled  = ""
		iRow       = 0

		if sUserType <> sResidentType then
		sDisabled = " disabled"
		end if

		sSql = "SELECT rateid, description, amount "
		sSql = sSql & " FROM egov_poolpassrates R, "
		sSql = sSql & " egov_memberships M, "
		sSql = sSql & " egov_membership_periods P "
		sSql = sSql & " WHERE P.periodid = R.periodid "
		sSql = sSql & " AND R.membershipid = M.membershipid "
		sSql = sSql & " AND M.orgid = P.orgid "
		sSql = sSql & " AND R.publiccanpurchase = 1 "
		sSql = sSql & " AND R.isEnabled = 1 "
		sSql = sSql & " AND M.orgid = "         & iOrgId
		sSql = sSql & " AND residenttype = '"   & sResidenttype & "' "
		sSql = sSql & " AND M.membershipid = "  & iMembershipId 
		sSql = sSql & " AND P.periodid = "      & iPeriodId
		sSql = sSql & " AND R.residenttype = '" & sUserType & "' "
		sSql = sSql & " ORDER BY displayorder "
		'response.write "<!-- " & sSql & " -->"

		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1


		do while not oRs.eof
			iRow = iRow + 1

			response.write "  <tr>" & vbcrlf
			response.write "      <td valign=""top""><input type=""radio"" name=""rateid"" id=""rateid_" & oRs("rateid") & """ value=""" & oRs("rateid") & """" & sDisabled & " onclick=""enableDisableContinueButton('" & oRs("rateid") & "');"" />" & oRs("description") & "</td>" & vbcrlf
			response.write "      <td class=""pickprice"">" & FormatCurrency(oRs("amount"),2) & "</td>" & vbcrlf
			response.write "  </tr>" & vbcrlf

			oRs.movenext
		loop

		oRs.close
		set oRs = nothing

		if iRow = 0 then
			response.write "  <tr><td colspan=""2"">&nbsp; None Available</td></tr>" & vbcrlf
		end if

	end sub


	'------------------------------------------------------------------------------
	public sub ShowResidentRates_AltLayout( ByVal iMembershipId, ByVal sUserType, ByVal iRateDesc )
		Dim sDisabled, sSql, oRs, bPreSelect, iRow, iCurrYear
		
		iRow = 0 

  		lcl_rate_desc = getDistinctRateDescList( iMembershipID, sUserType )

		sSql = "SELECT DISTINCT R.description "
		sSql = sSql & " FROM egov_poolpassrates R, "
		sSql = sSql & " egov_membership_periods MP "
		sSql = sSql & " WHERE R.periodid = MP.periodid "
		sSql = sSql & " AND R.orgid = MP.orgid "
		sSql = sSql & " AND R.isEnabled = 1 "
		sSql = sSql & " AND R.publiccanpurchase = 1 "
		sSql = sSql & " AND R.orgid = "         & iorgid
		sSql = sSql & " AND R.membershipid = "  & iMembershipId
		sSql = sSql & " AND R.residenttype = '" & sUserType     & "' "
		sSql = sSql & " AND R.description IN (" & lcl_rate_desc & ") "
		sSql = sSql & " ORDER BY R.description "

		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		if not oRs.eof then
			response.write "<select name=""rateDesc"" id=""rateDesc"" onchange=""submitPoolPassForm();"">" & vbcrlf

			do while not oRs.eof
				iRow = iRow + 1

				if UCASE(oRs("description")) = UCASE(iRateDesc) then
					lcl_selected_rates = " selected=""selected"""
				else
					lcl_selected_rates = ""
				end if

				response.write "<option value=""" & oRs("description") & """" & lcl_selected_rates & ">" & oRs("description") & "</option>" & vbcrlf

				oRs.MoveNext
			loop

			response.write "</select>" & vbcrlf

		end if

		oRs.Close
		set oRs = nothing

	end sub


	'------------------------------------------------------------------------------
	function getFirstMembershipRateOption_AltLayout( ByVal sUserType, ByVal iMembershipId )
		Dim sSql, oRs, lcl_return
		
		lcl_return = ""

		lcl_rate_desc = getDistinctRateDescList( iMembershipID, sUserType )

		'Now grab the first rate.
		sSql = "SELECT R.description, "
		sSql = sSql & " R.amount, "
		sSql = sSql & " R.displayorder "
		sSql = sSql & " FROM egov_poolpassrates R, "
		sSql = sSql & " egov_membership_periods MP "
		sSql = sSql & " WHERE R.periodid = MP.periodid "
		sSql = sSql & " AND R.orgid = MP.orgid "
		sSql = sSql & " AND R.isEnabled = 1 "
		sSql = sSql & " AND R.publiccanpurchase = 1 "
		sSql = sSql & " AND R.orgid = "         & iorgid
		sSql = sSql & " AND R.membershipid = "  & iMembershipId
		sSql = sSql & " AND R.residenttype = '" & sUserType     & "' "
		sSql = sSql & " AND R.description IN (" & lcl_rate_desc & ") "
		sSql = sSql & " ORDER BY R.displayorder, R.description "

		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		if not oRs.eof then
			lcl_return = oRs("description")
		end if

		oRs.Close
		set oRs = nothing

		getFirstMembershipRateOption_AltLayout = lcl_return

	end function

	'------------------------------------------------------------------------------
	function getDistinctRateDescList( ByVal iMembershipID, ByVal sUserType )
		Dim sSql, oRs, lcl_return
		
		lcl_return = "0"

		'Get a distinct list of rate descriptions
		sSql = "SELECT distinct R.description "
		sSql = sSql & " FROM egov_poolpassrates R, "
		sSql = sSql & " egov_membership_periods MP "
		sSql = sSql & " WHERE R.periodid = MP.periodid "
		sSql = sSql & " AND R.orgid = MP.orgid "
		sSql = sSql & " AND R.isEnabled = 1 "
		sSql = sSql & " AND R.publiccanpurchase = 1 "
		sSql = sSql & " AND R.orgid = "         & iorgid
		sSql = sSql & " AND R.membershipid = "  & iMembershipId
		sSql = sSql & " AND R.residenttype = '" & sUserType     & "' "

		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		if not oRs.eof then
			lcl_rate_desc = ""

			do while not oRs.eof
				if lcl_rate_desc <> "" then
					lcl_rate_desc = lcl_rate_desc & ",'" & oRs("description") & "'"
				else
					lcl_rate_desc = "'" & oRs("description") & "'"
				end if

				oRs.movenext
			loop

			if lcl_rate_desc <> "" then
				lcl_return = lcl_rate_desc
			end if

		end if

		oRs.close
		set oRs = nothing

		getDistinctRateDescList = lcl_return

	end function
	

	'------------------------------------------------------------------------------
	public function GetInitialPeriod( ByVal iMembershipId )
		Dim sSql, oRs
		lcl_return = 0

		sSql = "SELECT DISTINCT P.periodid, "
		sSql = sSql & " P.period_desc "
		sSql = sSql & " FROM egov_membership_periods P, "
		sSql = sSql & " egov_poolpassrates R "
		sSql = sSql & " WHERE R.periodid = P.periodid "
		sSql = sSql & " AND R.isEnabled = 1 "
		sSql = sSql & " AND R.publiccanpurchase = 1 "
		sSql = sSql & " AND P.orgid = "        & iOrgId
		sSql = sSql & " AND R.membershipid = " & iMembershipId

		if checkDefaultPeriodExists() then
			sSql = sSql & " AND P.isDefault = 1 "
		end if

		sSql = sSql & " ORDER BY P.period_desc DESC"
		'response.write sSql

		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		if not oRs.eof then
			lcl_return = CLng(oRs("periodid"))
		end if

		oRs.close
		set oRs = nothing

		GetInitialPeriod = lcl_return

	end function


	'------------------------------------------------------------------------------
	public function checkDefaultPeriodExists()
		Dim sSql, oRs, lcl_return
		
		lcl_return = False

		sSql = "SELECT periodid "
		sSql = sSql & " FROM egov_membership_periods "
		sSql = sSql & " WHERE orgid = " & iorgid
		sSql = sSql & " AND isdefault = 1 "

		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		if not oRs.eof then
			lcl_return = True
		end if

		oRs.close
		set oRs = nothing

		checkDefaultPeriodExists = lcl_return

	end function


	'-----------------------------------------------------------------------------
	 public sub ShowPeriodPicksForMembership( ByVal iMembershipId, ByVal iPeriodId )
		Dim sSql, oRs

		'Get the Periods
		sSql = "SELECT DISTINCT P.periodid, "
		sSql = sSql & " P.period_desc "
		sSql = sSql & " FROM egov_membership_periods P, "
		sSql = sSql & " egov_poolpassrates R "
		sSql = sSql & " WHERE R.periodid = P.periodid "
		sSql = sSql & " AND R.isEnabled = 1 "
		sSql = sSql & " AND R.publiccanpurchase = 1 "
		sSql = sSql & " AND P.orgid = "        & iOrgId
		sSql = sSql & " AND R.membershipid = " & iMembershipId
		sSql = sSql & " ORDER BY P.period_desc DESC"

		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1
		
		if not oRs.EOF then
  			response.write "<select name=""periodid"" id=""periodid"" onchange=""submitPoolPassForm();"">" & vbcrlf

			do while not oRs.eof
				if CLng(iPeriodId) = CLng(oRs("periodid")) then
					lcl_selected_period = " selected=""selected"""
				else
					lcl_selected_period = ""
				end if

				response.write "  <option value=""" & oRs("periodid") & """" & lcl_selected_period & ">" & oRs("period_desc") & "</option>" & vbcrlf

				oRs.movenext
			loop

			response.write "</select>" & vbcrlf

			end if

		oRs.close
		set oRs = nothing

	end sub


	'------------------------------------------------------------------------------
	public sub ShowPeriodPicksForMembership_AltLayout ( ByVal iMembershipId, ByVal iPeriodId, ByVal sUserType, ByVal iRateDesc )
		Dim sSql, oRs

		'Get the Periods
		sSql = "SELECT DISTINCT P.periodid, "
		sSql = sSql & " P.period_desc, "
		sSql = sSql & " R.amount "
		sSql = sSql & " FROM egov_membership_periods P, "
		sSql = sSql & " egov_poolpassrates R "
		sSql = sSql & " WHERE R.periodid = P.periodid "
		sSql = sSql & " AND R.isEnabled = 1 "
		sSql = sSql & " AND R.publiccanpurchase = 1 "
		sSql = sSql & " AND P.orgid = "               & iOrgId
		sSql = sSql & " AND R.membershipid = "        & iMembershipId
		sSql = sSql & " AND R.residenttype = '"       & sUserType        & "' "
		sSql = sSql & " AND UPPER(R.description) = '" & UCASE(iRateDesc) & "' "
		sSql = sSql & " ORDER BY P.period_desc DESC"
		response.write "<!-- " & sSql & " -->"

		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		response.write "<div class=""picktitle"">&nbsp;Membership Choices</div>" & vbcrlf
		response.write "<div class=""pickchoice"">" & vbcrlf
		response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" id=""picktable_altlayout"">" & vbcrlf

		if not oRs.eof then
			'response.write "  <tr><th colspan=""2""> &nbsp; " & oTypes("description") & "</th></tr>" & vbcrlf

			i = 0
			do while not oRs.eof
				i = i + 1
				if iPeriodID <> "" then
					if CLng(iPeriodId) = CLng(oRs("periodid")) then
						lcl_checked_period = " checked=""checked"""
						lcl_updaterateid   = "Y"
					else
						lcl_checked_period = ""
						lcl_updaterateid   = "N"
					end if
				else
					lcl_checked_period = ""
					lcl_updaterateid   = "N"
				end if

				lcl_rateid = getRateID( iMembershipId, sUserType, iRateDesc, oRs("periodid") )

				response.write "  <tr>" & vbcrlf
				response.write "      <td>&nbsp;<input type=""radio"" name=""periodid"" id=""periodid_" & oRs("periodid") & """ value=""" & oRs("periodid") & """" & sDisabled & lcl_checked_period & " onclick=""updateRateID('" & lcl_rateid & "');enableDisableContinueButton('" & oRs("periodid") & "');submitPoolPassForm();"" />" & oRs("period_desc") & "</td>" & vbcrlf
				response.write "      <td align=""right"">" & FormatCurrency(oRs("amount")) & "&nbsp;</td>" & vbcrlf
				response.write "  </tr>" & vbcrlf

				'This updates the rateid field
				if lcl_updaterateid = "Y" then
				response.write "<script language=""javascript"">" & vbcrlf
				response.write "  updateRateID('" & lcl_rateid & "');" & vbcrlf
				response.write "</script>" & vbcrlf
				end if

				oRs.MoveNext
			loop

		else
			response.write "  <tr><td colspan=""2"">No Membership Rates Available</td></tr>" & vbcrlf
		end if

		response.write "</table>" & vbcrlf
		response.write "</div>" & vbcrlf

		oRs.Close
		set oRs = nothing

	end sub


	'------------------------------------------------------------------------------
	function getRateID( ByVal iMembershipId, ByVal sUserType, ByVal iRateDesc, ByVal iPeriodID )
		Dim sSql, oRs

		lcl_return = 0

		'Get the RateID
		sSql = "SELECT DISTINCT R.rateid "
		sSql = sSql & " FROM egov_membership_periods P, "
		sSql = sSql & " egov_poolpassrates R "
		sSql = sSql & " WHERE R.periodid = P.periodid "
		sSql = sSql & " AND P.orgid = "               & iorgid
		sSql = sSql & " AND R.membershipid = "        & iMembershipId
		sSql = sSql & " AND R.residenttype = '"       & sUserType        & "' "
		sSql = sSql & " AND UPPER(R.description) = '" & UCASE(iRateDesc) & "' "
		sSql = sSql & " AND P.periodid = "            & iPeriodID

		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		if not oRs.eof then
			lcl_return = oRs("rateid")
		end if

		oRs.close
		set oRs = nothing

		getRateID = lcl_return

	end function


	'------------------------------------------------------------------------------
	public function GetMembershipNameById( ByVal iMembershipId )
 		Dim sSql, oRs

 		sSql = "SELECT membershipdesc FROM egov_memberships WHERE membershipid = " & iMembershipId

 		set oRs = Server.CreateObject("ADODB.Recordset")
 		oRs.Open sSql, Application("DSN"), 3, 1

 		if not oRs.EOF then
	   		GetMembershipNameById = oRs("membershipdesc")
		 else
   			GetMembershipNameById = ""
   		end if

 		oRs.Close
	 	set oRs = nothing

 	end function


	'------------------------------------------------------------------------------
	public function getMembershipIdByMembership( ByVal iMembershipType )
		Dim sSql, oRs, lcl_return
		
		lcl_return = 0

		if iMembershipType <> "" then
			lcl_membershiptype = UCASE(iMembershipType)
		else
			lcl_membershiptype = "POOL"
		end if

		sSql = "SELECT membershipid "
		sSql = sSql & " FROM egov_memberships "
		sSql = sSql & " WHERE orgid = " & iorgid
		sSql = sSql & " AND UPPER(membership) = '" & UCASE(lcl_membershiptype) & "' "

		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		if not oRs.eof then
			lcl_return = oRs("membershipid")
		end if

		oRs.close
		set oRs = nothing

		getMembershipIdByMembership = lcl_return

	end function


	'------------------------------------------------------------------------------
	public function GetMembershipPeriodName( ByVal iPeriodId )
 		Dim sSql, oRs

	 	sSql = "SELECT period_desc FROM egov_membership_periods WHERE periodid = " & iPeriodId

 		set oRs = Server.CreateObject("ADODB.Recordset")
 		oRs.Open sSql, Application("DSN"), 3, 1

		 if not oRs.EOF then
   			GetMembershipPeriodName = oRs("period_desc")
  		else
   			GetMembershipPeriodName = ""
		 end if

 		oRs.Close
 		set oRs = nothing

 	end function


	'------------------------------------------------------------------------------
	function getExpirationDate( ByVal p_periodid, ByVal p_startdate )
		Dim sSql, oRs, lcl_expirationdate
		
		lcl_expirationdate = ""

		if p_startdate <> "" then
			'Calculate the expiration date
			sSql = "SELECT CAST(dbo.fn_getMembershipExpirationdate(MP.is_seasonal,MP.period_interval,MP.period_qty,'" & p_startdate & "') AS datetime) AS expirationdate "
			sSql = sSql & " FROM egov_membership_periods MP "
			sSql = sSql & " WHERE MP.orgid = " & iorgid
			sSql = sSql & " AND MP.periodid = " & p_periodid

			set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open sSql, Application("DSN"), 3, 1

			if not oRs.eof then
				lcl_expirationdate = oRs("expirationdate")
			end if

			oRs.Close
			set oRs = nothing

		end if

		getExpirationDate = lcl_expirationdate

	end function


	'------------------------------------------------------------------------------
	function getMembershipStartDate( ByVal p_poolpassid )
		Dim sSql, oRs, lcl_return
		
		lcl_return = Date()

		if p_poolpassid <> "" then
			sSql = "SELECT expirationdate FROM egov_poolpasspurchases WHERE poolpassid = " & p_poolpassid

			set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open sSql, Application("DSN"), 3, 1

			if not oRs.eof then
				'Add one day to the renewal start date since we do not want two memberships active on the same date.
				lcl_return = DATEADD("d",1,oRs("expirationdate"))
			end if

			oRs.Close
			set oRs = nothing
		end if

		getMembershipStartDate = lcl_return

	end function
	
	
	

	'-------------------------------------------------------------------------------------------------
	' void RunSQLStatement sSql 
	'-------------------------------------------------------------------------------------------------
	Private Sub RunSQLStatement( ByVal sSql )
		Dim oCmd

	'	response.write "<p>" & sSql & "</p><br /><br />"
	'	response.flush

		Set oCmd = Server.CreateObject("ADODB.Command")
		oCmd.ActiveConnection = Application("DSN")
		oCmd.CommandText = sSql
		oCmd.Execute
		Set oCmd = Nothing

	End Sub 


	'-------------------------------------------------------------------------------------------------
	' integer RunIdentityInsertStatement( sInsertStatement )
	'-------------------------------------------------------------------------------------------------
	Private Function RunIdentityInsertStatement( ByVal sInsertStatement )
		Dim sSql, iReturnValue, oInsert

		iReturnValue = 0

		'response.write "<p>" & sInsertStatement & "</p><br /><br />"
		'response.flush

		'INSERT NEW ROW INTO DATABASE AND GET ROWID
		sSql = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

		Set oInsert = Server.CreateObject("ADODB.Recordset")
		oInsert.CursorLocation = 3
		oInsert.Open sSql, Application("DSN"), 3, 3
		iReturnValue = oInsert("ROWID")
		oInsert.Close
		Set oInsert = Nothing

		RunIdentityInsertStatement = iReturnValue

	End Function



'------------------------------------------------------------------------------
end class
%>
