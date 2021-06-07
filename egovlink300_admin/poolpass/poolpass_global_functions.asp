<%
'------------------------------------------------------------------------------
function getUserResidentType(p_userid)
  lcl_return = "N"

  sSQL = "SELECT isnull(residenttype,'N') AS residenttype "
  sSQL = sSQL & " FROM egov_users "
  sSQL = sSQL & " WHERE userid = " & p_userid

 	set rs = Server.CreateObject("ADODB.Recordset")
	 rs.Open sSQL, Application("DSN"), 3, 1

  if not rs.eof then
     lcl_return = rs("residenttype")
  end if

  set rs = nothing

  getUserResidentType = lcl_return

end function

'------------------------------------------------------------------------------
function getResidentTypeDesc(p_residenttype)

  lcl_return       = "N"
  lcl_residentDesc = ""

  if p_residenttype <> "" then
     lcl_residentDesc = dbsafe(p_residenttype)
  end if

  lcl_residentDesc = "'" & lcl_residentDesc & "'"

  sSQL = "SELECT description "
  sSQL = sSQL & " FROM egov_poolpassresidenttypes "
  sSQL = sSQL & " WHERE resident_type = " & lcl_residentDesc

 	set rs = Server.CreateObject("ADODB.Recordset")
	 rs.Open sSQL, Application("DSN"), 3, 1

  if not rs.eof then
     lcl_return = rs("description")
  end if

  set rs = nothing

  getResidentTypeDesc = lcl_return

end function

'------------------------------------------------------------------------------
function getFirstUserID()
  lcl_return = 0

  sSQL = "SELECT TOP 1 userid "
  sSQL = sSQL & " FROM egov_users "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")
  sSQL = sSQL & " AND userregistered = 1 "
  sSQL = sSQL & " AND headofhousehold = 1 "
  sSQL = sSQL & " AND userfname IS NOT NULL "
  sSQL = sSQL & " AND userlname IS NOT NULL "
  sSQL = sSQL & " AND userfname <> '' "
  sSQL = sSQL & " AND userlname <> '' "
  sSQL = sSQL & " ORDER BY userlname, userfname "

 	set rs = Server.CreateObject("ADODB.Recordset")
	 rs.Open sSQL, Application("DSN"), 3, 1

  if not rs.eof then
     lcl_return = rs("userid")
  end if

  set rs = nothing

  getFirstUserID = lcl_return

end function

'------------------------------------------------------------------------------
function formatPhone(p_number)
  lcl_return = ""

  if len(p_number) = 10 then
     lcl_return = "(" & left(p_number,3) & ") " & mid(p_number, 4, 3) & "-" & right(p_number,4)
  else
     lcl_return = p_number
  end if

  formatPhone = lcl_return

end function

'------------------------------------------------------------------------------
sub getRateInfo(ByVal p_rateid, _
                ByRef lcl_rate_description, _
                ByRef lcl_rate_residenttype)

  lcl_rate_description  = ""
  lcl_rate_residenttype = ""
  lcl_residenttype_desc = ""

 'Retrieve the RATE information if it exists
  sSQL = "SELECT description, residenttype "
  sSQL = sSQL & " FROM egov_poolpassrates "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")
  sSQL = sSQL & " AND rateid = "  & p_rateid

 	set oRate = Server.CreateObject("ADODB.Recordset")
	 oRate.Open sSQL, Application("DSN"), 3, 1

  if not oRate.eof then
     lcl_description       = oRate("description")
     lcl_rate_residenttype = oRate("residenttype")
  end if

  set oRate = nothing

 'Retrieve the ResidentType information if it exists
  sSQL = "SELECT description "
  sSQL = sSQL & " FROM egov_poolpassresidenttypes "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")
  sSQL = sSQL & " AND UPPER(resident_type) = '" & UCASE(lcl_rate_residenttype) & "'"

 	set oType = Server.CreateObject("ADODB.Recordset")
	 oType.Open sSQL, Application("DSN"), 3, 1

  if not oType.eof then
     lcl_residenttype_desc = oType("description")
  end if

  set oType = nothing

 'Combine the values together to make the complete description
  if lcl_description <> "" then
     lcl_rate_description = lcl_description
  end if

  if lcl_residenttype_desc <> "" and lcl_rate_description <> "" then
        lcl_rate_description = lcl_rate_description & " - " & lcl_residenttype_desc
  end if

end sub

'------------------------------------------------------------------------------
sub getPoolPassInfo(ByVal lcl_poolpassid, ByRef iUserId, ByRef iRateId, ByRef iMembershipId, ByRef iPeriodId, ByRef lcl_isSeasonal)
  iUserID        = 0
  iRateID        = 0
  iMembershipId  = 0
  iPeriodID      = 0
  lcl_isSeasonal = False

  sSQL = "SELECT P.userid, P.rateid, P.membershipid, P.periodid, MP.is_seasonal "
  sSQL = sSQL & " FROM egov_poolpasspurchases P "
  sSQL = sSQL &      " LEFT OUTER JOIN egov_membership_periods MP ON MP.periodid = P.periodid "
  sSQL = sSQL &      " AND MP.orgid = " & session("orgid")
  sSQL = sSQL & " WHERE P.poolpassid = " & lcl_poolpassid

 	set oPool = Server.CreateObject("ADODB.Recordset")
	 oPool.Open sSQL, Application("DSN") , 3, 1

  if not oPool.eof then
     iUserID        = oPool("userid")
     iRateId        = oPool("rateid")
     iMembershipId  = oPool("membershipid")
     iPeriodID      = oPool("periodid")
     lcl_isSeasonal = oPool("is_seasonal")
  end if

  set oPool = nothing

end sub

'------------------------------------------------------------------------------
sub getPoolPassRateAmount(ByVal iRateID, ByRef nAmount, ByRef sDescription, ByRef sType)
  nAmount      = 0
  sDescription = ""
  sType        = ""

  sSQL = "SELECT R.amount, R.description AS rate_desc, T.description as residenttype_desc "
  sSQL = sSQL & " FROM egov_poolpassrates R, egov_poolpassresidenttypes T "
  sSQL = sSQL & " WHERE R.rateid = " & iRateID
  sSQL = sSQL & " AND R.residenttype = T.resident_type "
  sSQL = sSQL & " AND T.orgid = R.orgid "
  sSQL = sSQL & " AND R.orgid = " & session("orgid")

 	set rs = Server.CreateObject("ADODB.Recordset")
	 rs.Open sSQL, Application("DSN"), 3, 1

  if not rs.eof then
     nAmount      = rs("amount")
     sDescription = rs("rate_desc")
     sType        = rs("residenttype_desc")
  end if

  set rs = nothing

end sub

'------------------------------------------------------------------------------
function GetCityName()
	 lcl_return = ""

 	sSQL = "SELECT orgname FROM organizations WHERE orgid = " & session("orgid")

	 set oName = Server.CreateObject("ADODB.Recordset")
	 oName.Open sSQL, Application("DSN"), 3, 1

	 if not oName.eof then
   		lcl_return = oName("orgname")
 	end if

 	oName.close
	 set oName = nothing

  GetCityName = lcl_return

end function

'------------------------------------------------------------------------------
function MakeProper( sString )
 lcl_return = ""

	if sString <> "" then
		  lcl_return = UCASE(Left(sString,1)) & LCASE(Mid(sString,2))
 end if

 MakeProper = lcl_return

end function

'------------------------------------------------------------------------------
sub getMembershipPeriodInfo (ByVal lcl_periodid, ByRef lcl_isSeasonal, ByRef lcl_periodtype, ByRef lcl_periodinterval, ByRef lcl_periodqty)
  lcl_isSeasonal     = ""
  lcl_periodtype     = ""
  lcl_periodinterval = ""
  lcl_periodqty      = ""

  if lcl_periodid <> "" then
     sSQL = "SELECT is_seasonal, period_type, period_interval, period_qty "
     sSQL = sSQL & " FROM egov_membership_periods "
     sSQL = sSQL & " WHERE periodid = " & lcl_periodid
     sSQL = sSQL & " AND orgid = " & session("orgid")

    	set oPeriod = Server.CreateObject("ADODB.Recordset")
   	 oPeriod.Open sSQL, Application("DSN"), 3, 1

     if not oPeriod.eof then
        lcl_isSeasonal     = oPeriod("is_seasonal")
        lcl_periodtype     = oPeriod("period_type")
        lcl_periodinterval = oPeriod("period_interval")
        lcl_periodqty      = oPeriod("period_qty")
     end if
  end if

  set oPeriod = nothing

end sub

'------------------------------------------------------------------------------
sub getCurrentMemberInfo(ByVal iPoolPassID, ByVal iFamilyMemberID, ByRef lcl_member_id, ByRef lcl_card_printed, ByRef lcl_printed_count)
  lcl_member_id     = ""
  lcl_card_printed  = "N"
  lcl_printed_count = 0

  sSQL = "SELECT memberid, card_printed, printed_count "
  sSQL = sSQL & " FROM egov_poolpassmembers "
  sSQL = sSQL & " WHERE poolpassid = " & iPoolPassID
  sSQL = sSQL & " AND familymemberid = " & iFamilyMemberID

 	set oMember = Server.CreateObject("ADODB.Recordset")
	 oMember.Open sSQL, Application("DSN"), 3, 1

  if not oMember.eof then
     lcl_member_id     = oMember("memberid")
     lcl_card_printed  = oMember("card_printed")
     lcl_printed_count = oMember("printed_count")
  end if

 'If there is no memberid then get the next ID available.
  if lcl_member_id = "" then
     lcl_member_id = getNextMemberID()
  end if

  set oMember = nothing

end sub

'------------------------------------------------------------------------------
function getNextMemberID()
  lcl_return = 1

  sSQL = "SELECT MAX(memberid) + 1 AS newID "
  sSQL = sSQL & " FROM egov_poolpassmembers "

 	set oMax = Server.CreateObject("ADODB.Recordset")
	 oMax.Open sSQL, Application("DSN"), 3, 1

  if not oMax.eof then
     lcl_return = oMax("newID")
  end if

  set oMax = nothing

  getNextMemberID = lcl_return

end function

'------------------------------------------------------------------------------
function IsOnPoolPass( iPoolPassId, iFamilymemberid )

	sSQL = "SELECT count(familymemberid) AS hits "
 sSQL = sSQL & " FROM egov_poolpassmembers "
 sSQL = sSQL & " WHERE poolpassid = " & iPoolPassId
 sSQL = sSQL & " AND familymemberid = " & iFamilymemberid

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	if not oRs.eof then
  		if clng(oRs("hits")) > clng(0) then
    			IsOnPoolPass = True
  		else
     		IsOnPoolPass = False
    end if
 else
  		IsOnPoolPass = False
 end if

	oRs.Close
	set oRs = nothing

end function

'------------------------------------------------------------------------------
function getCurrentPoolPassID(p_member_id)
  lcl_return          = 0
		lcl_current_date    = Now()
  lcl_poolpassid      = 0
  lcl_expiration_date = ""

  'sSQL = "SELECT P.poolpassid, DATEADD(yy, 1, P.paymentdate) AS expiration_date, P.paymentdate, MP.period_interval, "

  if isnumeric(p_member_id) then
     sSQL = "SELECT P.poolpassid, P.paymentdate, P.startdate, P.expirationdate, MP.period_interval, MP.period_qty, MP.period_type "
     sSQL = sSQL & " FROM egov_poolpasspurchases AS P, egov_poolpassmembers AS ppm, egov_membership_periods AS MP  "
     sSQL = sSQL & " WHERE P.poolpassid = ppm.poolpassid "
     sSQL = sSQL & " AND P.periodid = MP.periodid "
     sSQL = sSQL & " AND ppm.memberid = " & p_member_id
     sSQL = sSQL & " AND P.orgid = "      & session("orgid")
     sSQL = sSQL & " ORDER BY P.poolpassid DESC "
     'response.write sSQL

    	set oPool = Server.CreateObject("ADODB.Recordset")
   	 oPool.Open sSQL, Application("DSN"), 3, 1

     if not oPool.eof then
        while not oPool.eof

           if  CDate(date()) >= CDate(oPool("startdate")) _
           AND CDate(date()) <= CDate(oPool("expirationdate")) then
              lcl_poolpassid      = oPool("poolpassid")
              lcl_expiration_date = datevalue(oPool("expirationdate"))
           else
              lcl_poolpassid      = lcl_poolpassid
              lcl_expiration_date = lcl_expiration_date
           end if

           oPool.movenext
        wend
     end if
  end if

	'if datevalue(lcl_current_date) <= datevalue(lcl_expiration_date) then
    lcl_return = lcl_poolpassid
 'end if

  set oPool = nothing

  getCurrentPoolPassID = lcl_return

end function

'------------------------------------------------------------------------------
function getCheckInResult(iOrgID, p_member_id )

	dim lcl_member_id, lcl_current_date, lcl_expiration_date
 dim lcl_orghasfeature_memberships_usekeycards

 lcl_orghasfeature_memberships_usekeycards = orghasfeature("memberships_usekeycards")

'We now check to see if the org is using keycards/barcodes instead of memberids.
'If "yes" then we DO know that if we have gotten to this point that a barcode DOES exist AND that 
'   it is associated to a memberid.  We we need to do know is check the status of the barcode.
'
'If "no" then have to determine if this memberid is a "renewal id" or not.
'   This means that this memberid could potentially be on more than one PoolPassPurchase.
'   If this is the case then we have to determine if the membership is valid at the current 
'   time (i.e. membership may be currently invalid/expired, but valid for the next season, 
'   BUT the user is attempting to use it right now.)

' if lcl_orghasfeature_memberships_usekeycards then
'    getCheckInResult = "barcode_status"
' else
	'response.write p_member_id
    lcl_poolpassid = getCurrentPoolPassID(p_member_id)

   	sSQL = "SELECT count(memberid) AS total_cnt "
    sSQL = sSQL & " FROM egov_poolpassmembers m, "
    sSQL = sSQL &      " egov_poolpasspurchases p "
    sSQL = sSQL & " WHERE m.poolpassid = p.poolpassid "
    sSQL = sSQL & " AND m.memberid = "   & p_member_id
    sSQL = sSQL & " AND p.orgid = "      & session("orgid")
    sSQL = sSQL & " AND p.poolpassid = " & lcl_poolpassid
    'response.write sSQL

   	set oCard = Server.CreateObject("ADODB.Recordset")
   	oCard.Open sSQL, Application("DSN"), 3, 1

   	if CLng(oCard("total_cnt")) = CLng(0) then
     		getCheckInResult = "none"
   	else
		     lcl_current_date = Date()

     		sSQL2 = "SELECT P.paymentdate, "
       sSQL2 = sSQL2 & " P.startdate, "
       sSQL2 = sSQL2 & " P.expirationdate, "
       sSQL2 = sSQL2 & " MP.period_interval, "
       sSQL2 = sSQL2 & " MP.period_qty, "
       sSQL2 = sSQL2 & " MP.period_type "
     		sSQL2 = sSQL2 & " FROM egov_poolpasspurchases P, "
       sSQL2 = sSQL2 &      " egov_poolpassmembers ppm, "
       sSQL2 = sSQL2 &      " egov_membership_periods MP "
		     sSQL2 = sSQL2 & " WHERE P.poolpassid = ppm.poolpassid "
       sSQL2 = sSQL2 & " AND P.periodid = MP.periodid "
		     sSQL2 = sSQL2 & " AND ppm.memberid = " & p_member_id
       sSQL2 = sSQL2 & " AND P.poolpassid = " & lcl_poolpassid

     		set oCardExp = Server.CreateObject("ADODB.Recordset")
		     oCardExp.Open sSQL2, Application("DSN"), 3, 1

       if not oCardExp.eof then
          'if UCASE(oCardExp("period_type")) = "SEASON" then
          '   lcl_expiration_date = FormatDateTime(CDate("12/31/" & Year(oCardExp("paymentdate"))), vbshortdate)
          'else
          '   lcl_expiration_date = FormatDateTime(DateAdd(oCardExp("period_interval"),clng(oCardExp("period_qty")),DateValue(oCardExp("paymentdate"))), vbshortdate)
          'end if
          lcl_expiration_date = oCardExp("expirationdate")

        		if CDate(lcl_current_date) > CDate(lcl_expiration_date) then
		     	     getCheckInResult = "expired"
   	      else
             if  datevalue(lcl_current_date) >= datevalue(oCardExp("startdate")) _
             AND datevalue(lcl_current_date) <= datevalue(lcl_expiration_date) then
             			'lcl_return = p_member_id
                 lcl_return = "valid"

               'Perform one last check to see if the member has already signed in earlier on the same day.
               'If so then sound ALERT otherwise the member is VALID.
                lcl_lastscanned = getLastScannedDate(iOrgID, p_member_id,"")

               'This is for potential future work.  Check to see if the ID has alerady been scanned for current date
                if lcl_lastscanned <> "" then
                   if formatdatetime(lcl_lastscanned,vbshortdate) = formatdatetime(lcl_current_date,vbshortdate) then
                      lcl_return = "valid_multiplescan"
                   end if
                end if

          	   		getCheckInResult = lcl_return
             else
                getCheckInResult = "none"
             end if
         	end if

      	  	oCardExp.Close
     	   	set oCardExp = nothing
       else
          getCheckInResult = "none"
       end if

   	end if

   	oCard.Close 
   	set oCard = nothing
 'end if

end function

'------------------------------------------------------------------------------
function isRenewedPoolPass(p_poolpassid)
  lcl_return = "N"

 'Check to see if this PoolPassID exists as a "previous_poolpassid" on any other purchase(s).
 'If it does then this PoolPassID is considered to be "Renewed".
  if p_poolpassid <> "" then
     sSQL = "SELECT count(poolpassid) as total_count "
     sSQL = sSQL & " FROM egov_poolpasspurchases "
     sSQL = sSQL & " WHERE previous_poolpassid = " & p_poolpassid

    	set rs = Server.CreateObject("ADODB.Recordset")
   	 rs.Open sSQL, Application("DSN"), 3, 1

     if rs("total_count") > 0 then
        lcl_return = "Y"
     end if

     set rs = nothing

  end if

  isRenewedPoolPass = lcl_return

end function

'------------------------------------------------------------------------------
function showHideRenewalButton(p_poolpassid)
 lcl_return   = "N"

'We first have to determine if this memberid is a "renewal id" or not.
'This means that this memberid could potentially be on more than one PoolPassPurchase.
'If this is the case then we have to determine if the membership is valid at the current time (i.e. membership may be
' currently invalid/expired, but valid for the next season, BUT the user is attempting to use it right now.)
 lcl_poolpassid = p_poolpassid
	sSQL = "SELECT count(poolpassid) AS total_cnt "
 sSQL = sSQL & " FROM egov_poolpasspurchases p "
 sSQL = sSQL & " WHERE p.poolpassid = " & lcl_poolpassid

	set oCheck = Server.CreateObject("ADODB.Recordset")
	oCheck.Open sSQL, Application("DSN"), 3, 1

	if CLng(oCheck("total_cnt")) = CLng(0) then
  		lcl_return = "N"
	else
  		lcl_current_date = Date()

				sSQL2 = "SELECT P.paymentdate, "
    sSQL2 = sSQL2 & " P.startdate, "
    sSQL2 = sSQL2 & " P.expirationdate, "
    sSQL2 = sSQL2 & " P.periodid, "
    sSQL2 = sSQL2 & " MP.period_interval, "
    sSQL2 = sSQL2 & " MP.period_qty, "
    sSQL2 = sSQL2 & " MP.period_type, "
    sSQL2 = sSQL2 & " r.renewalstartdate, "
    sSQL2 = sSQL2 & " r.isRenewable, "
    sSQL2 = sSQL2 & " r.renewalTimeAfterExpire "
				sSQL2 = sSQL2 & " FROM egov_membership_periods MP, "
    sSQL2 = sSQL2 &      " egov_poolpasspurchases P LEFT OUTER JOIN egov_poolpassrates r ON p.rateid = r.rateid "
				sSQL2 = sSQL2 & " WHERE P.periodid = MP.periodid "
    sSQL2 = sSQL2 & " AND r.isEnabled = 1 "
				sSQL2 = sSQL2 & " AND p.poolpassid = " & lcl_poolpassid

				set oCheckExp = Server.CreateObject("ADODB.Recordset")
				oCheckExp.Open sSQL2, Application("DSN"), 3, 1

		  if not oCheckExp.eof then
       if oCheckExp("expirationdate") <> "" then
          lcl_expiration_date = CDate(oCheckExp("expirationdate"))
       else
          lcl_expiration_date = oMembership.getExpirationDate(oCheckExp("periodid"),oCheckExp("startdate"))
       end if

       if oCheckExp("isRenewable") <> "" then
          lcl_isRenewable = oCheckExp("isRenewable")
       else
          lcl_isRenewable = False
       end if

      'If there isn't a Renewal Start Date then set it to the current date -1 day.
       if oCheckExp("renewalstartdate") <> "" then
          lcl_renewalstartdate = oCheckExp("renewalstartdate")
       else
          lcl_renewalstartdate = DATEADD("d",-1,lcl_current_date)
       end if

       if oCheckExp("renewalTimeAfterExpire") <> "" then
          lcl_renewalTimeAfterExpire = oCheckExp("renewalTimeAfterExpire")
       else
          lcl_renewalTimeAfterExpire = 0
       end if

      'RENEWAL BUTTON RULES FOR POOLPASSID:
      '1. NOT a "previous_poolpassid" on any other record.
      '   *** If it IS a "previous_poolpassid" on another record then that means it has ALREADY been renewed. ***
      '2. Check to see if the rate associated to is is set as a "renewal".
      '3. Check to make sure that the start date is GREATER THAN or EQUAL TO the current date before allowing a renewal

      '-- Rules 4 and 5 do not apply on ADMIN site ----------------------------
      '4. The current date is EQUAL TO or GREATER THAN the renewal start date.
      '5. The current date is LESS THAN or EQUAL TO the date generated by the "expiration date" + "Days to Renew After Expiration Date".
      '------------------------------------------------------------------------
       if  isRenewedPoolPass(lcl_poolpassid) = "N" _
       AND lcl_isRenewable = True _
       AND datevalue(lcl_current_date) >= datevalue(oCheckExp("startdate")) then
       'AND datevalue(lcl_current_date) >= datevalue(lcl_renewalstartdate) _
       'AND clng(DATEDIFF("d",lcl_current_date,DATEADD("d",lcl_renewalTimeAfterExpire,lcl_expiration_date))) >= clng(0) then
 					     lcl_return = "Y"
   			 else
    		     lcl_return = "N"
  		  	end if

   				oCheckExp.Close
		   		set oCheckExp = nothing
  		else
    		 lcl_return = "N"
		  end if

 			oCheck.Close
	 		set oCheck = nothing

 end if

 showHideRenewalButton = lcl_return

end function

'------------------------------------------------------------------------------
function checkRateEnabled(iRateID)

  dim lcl_return, sSQL

  lcl_return = false

  sSQL = "SELECT isEnabled "
  sSQL = sSQL & " FROM egov_poolpassrates "
  sSQL = sSQL & " WHERE rateid = " & iRateID

		set oIsRateEnabled = Server.CreateObject("ADODB.Recordset")
		oIsRateEnabled.Open sSQL, Application("DSN"), 3, 1

  if not oIsRateEnabled.eof then
     if oIsRateEnabled("isEnabled") then
        lcl_return = oIsRateEnabled("isEnabled")
     end if
  end if

  oIsRateEnabled.close
  set oIsRateEnabled = nothing

  checkRateEnabled = lcl_return

end function

'------------------------------------------------------------------------------
function getMemberIDByBarcode(iOrgID, _
                              iBarcode)

  dim lcl_return, sSQL, sBarcode

  lcl_return = 0
  sBarcode   = ""

  if iBarcode <> "" then
     sBarcode = ucase(iBarcode)
     sBarcode = dbsafe(sBarcode)
  end if

  sBarcode = "'" & sBarcode & "'"

  sSQL = "SELECT memberid "
  'sSQL = sSQL & " FROM egov_poolpassmembers_to_barcodes "
  'sSQL = sSQL & " WHERE upper(barcode) = " & sBarcode
  'sSQL = sSQL & " AND orgid = " & iOrgID
  sSQL = sSQL & " FROM egov_poolpassmembers_to_barcodes b "
  sSQL = sSQL & " INNER JOIN egov_poolpassmembers_barcode_statuses bs ON bs.statusid = b.barcode_statusid "
  sSQL = sSQL & " WHERE upper(barcode) = " & sBarcode & " AND b.orgid = " & iOrgID & " and bs.isActiveStatus = 1 and bs.isEnabled = 1 "
  'response.write sSQL

  set oGetMemberIDByBarcode = Server.CreateObject("ADODB.Recordset")
  oGetMemberIDByBarcode.Open sSQL, Application("DSN"), 3, 1

  if not oGetMemberIDByBarcode.eof then
     lcl_return = oGetMemberIDByBarcode("memberid")
  end if

  oGetMemberIDByBarcode.close
  set oGetMemberIDByBarcode = nothing

  getMemberIDByBarcode = lcl_return

end function

'------------------------------------------------------------------------------
function isBarcodeStatusActive(iOrgID, _
                               iMemberID, _
                               iBarcode)
  dim lcl_return, sSQL, sBarcode

  lcl_return = false
  sBarcode   = ""

  if iBarcode <> "" then
     sBarcode = ucase(iBarcode)
     sBarcode = dbsafe(sBarcode)
  end if

  sBarcode = "'" & sBarcode & "'"

  sSQL = "SELECT isActiveStatus "
  sSQL = sSQL & " FROM egov_poolpassmembers_barcode_statuses "
  sSQL = sSQL & " WHERE statusid = (select barcode_statusid "
  sSQL = sSQL &                   " from egov_poolpassmembers_to_barcodes "
  sSQL = sSQL &                   " where memberid = " & iMemberID
  sSQL = sSQL &                   " and upper(barcode) = " & sBarcode
  sSQL = sSQL &                   " and orgid = " & iOrgID & ") "
  sSQL = sSQL & " AND orgid = " & iOrgID

 	set oCheckBarcodeActiveStatus = Server.CreateObject("ADODB.Recordset")
 	oCheckBarcodeActiveStatus.Open sSQL, Application("DSN"), 3, 1

  if not oCheckBarcodeActiveStatus.eof then
     if oCheckBarcodeActiveStatus("isActiveStatus") then
        lcl_return = true
     end if
  end if

  oCheckBarcodeActiveStatus.close
  set oCheckBarcodeActiveStatus = nothing

  isBarcodeStatusActive = lcl_return

end function

'------------------------------------------------------------------------------
function getBarcodeStatus(iMemberID, _
                          iBarcode)

  dim lcl_return , sSQL, sBarcode

  lcl_return = ""
  sBarcode   = ""

  if iBarcode <> "" then
     sBarcode = ucase(iBarcode)
     sBarcode = dbsafe(sBarcode)
  end if

  sBarcode = "'" & sBarcode & "'"

  sSQL = "SELECT statusname "
  sSQL = sSQL & " FROM egov_poolpassmembers_barcode_statuses "
  sSQL = sSQL & " WHERE statusid = (select barcode_statusid "
  sSQL = sSQL &                   " from egov_poolpassmembers_to_barcodes "
  sSQL = sSQL &                   " where memberid = " & iMemberID
  sSQL = sSQL &                   " and upper(barcode) = " & sBarcode & ") "

  set oGetBarcodeStatus = Server.CreateObject("ADODB.Recordset")
  oGetBarcodeStatus.Open sSQL, Application("DSN"), 3, 1

  if not oGetBarcodeStatus.eof then
     lcl_return = oGetBarcodeStatus("statusname")
  end if

  oGetBarcodeStatus.close
  set oGetBarcodeStatus = nothing

  getBarcodeStatus = lcl_return

end function

'------------------------------------------------------------------------------
function dtb_debug(p_value)
  sSQLi = "INSERT INTO my_table_dtb(notes) VALUES('" & replace(p_value,"'","''") & "')"
 	set rsi = Server.CreateObject("ADODB.Recordset")
	 rsi.Open sSQLi, Application("DSN"), 3, 1

  set rsi = nothing

end function
%>
