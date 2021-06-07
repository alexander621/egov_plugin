<%
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
function getUserResidentType(p_userid)
  lcl_return = "N"

  sSQL = "SELECT isnull(residenttype,'N') AS residenttype,useraddress,orgid "
  sSQL = sSQL & " FROM egov_users "
  sSQL = sSQL & " WHERE userid = " & p_userid

 	set rs = Server.CreateObject("ADODB.Recordset")
	 rs.Open sSQL, Application("DSN"), 3, 1

  if not rs.eof then
     lcl_return = rs("residenttype")
     if (rs("useraddress") = "" or isnull(rs("useraddress"))) and rs("orgid") = "228" then lcl_return = "Z"
  end if

  set rs = nothing

  getUserResidentType = lcl_return

end function

'------------------------------------------------------------------------------
function getResidentTypeDesc(p_residenttype)

  lcl_return = "Unknown Resident Type"

  sSQL = "SELECT description "
  sSQL = sSQL & " FROM egov_poolpassresidenttypes "
  sSQL = sSQL & " WHERE resident_type = '" & p_residenttype & "'"

 	set rs = Server.CreateObject("ADODB.Recordset")
	 rs.Open sSQL, Application("DSN"), 3, 1

  if not rs.eof then
     lcl_return = rs("description")
  end if

  set rs = nothing

  getResidentTypeDesc = lcl_return

end function

'------------------------------------------------------------------------------
sub ShowUserInfo(iUserID, ByRef sUserType, sResidentDesc)

  dim sUserID

  sUserID = 0

  if iUserID <> "" then
     sUserID = clng(iUserID)
  end if

  sSQL = "SELECT userfname, "
  sSQL = sSQL & " userlname, "
  sSQL = sSQL & " useraddress, "
  sSQL = sSQL & " useraddress2, "
  sSQL = sSQL & " userunit, "
  sSQL = sSQL & " usercity, "
  sSQL = sSQL & " userstate, "
  sSQL = sSQL & " userzip, "
  sSQL = sSQL & " usercountry, "
  sSQL = sSQL & " useremail, "
  sSQL = sSQL & " userhomephone, "
  sSQL = sSQL & " userworkphone, "
  sSQL = sSQL & " userfax, "
  sSQL = sSQL & " userbusinessname, "
  sSQL = sSQL & " userpassword, "
  sSQL = sSQL & " userregistered, "
  sSQL = sSQL & " residenttype, "
  sSQL = sSQL & " registrationblocked, "
  sSQL = sSQL & " blockeddate, "
  sSQL = sSQL & " blockedadminid, "
  sSQL = sSQL & " blockedexternalnote, "
  sSQL = sSQL & " blockedinternalnote "
  sSQL = sSQL & " FROM egov_users "
  sSQL = sSQL & " WHERE userid = " & sUserID

 	set oUser = Server.CreateObject("ADODB.Recordset")
	 oUser.Open sSQL, Application("DSN"), 3, 1

  if not oUser.eof then
    	sUserType = oUser("residenttype")

     response.write "<table border=""0"" cellpadding=""5"" cellspacing=""0"">" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td>Name</td>" & vbcrlf
     response.write "      <td>" & oUser("userfname") & " " & oUser("userlname") & " - <strong>" & sResidentDesc & "</strong></td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td valign=""top"">Email</td>" & vbcrlf
     response.write "      <td>" & oUser("useremail") & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td valign=""top"">Phone</td>" & vbcrlf
     response.write "      <td>" & FormatPhone(oUser("userhomephone")) & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td valign=""top"">Address</td>" & vbcrlf
     response.write "      <td>" & oUser("useraddress") & "<br />"

    	if oUser("useraddress2") = "" then
      		response.write oUser("useraddress2") & "<br />"
     end if

     response.write trim(oUser("usercity")) & ", " & oUser("userstate") & " " & oUser("userzip") & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td valign=""top"">Business</td>" & vbcrlf
     response.write "      <td>" & oUser("userbusinessname") & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "</table>" & vbcrlf

    	oUser.close
     set oUser = nothing
  end if
end sub

'------------------------------------------------------------------------------
function getFamilyMembers( ByVal iUserID, ByVal iRateID, ByVal iPoolPassID )
  Dim sSql, oUser, lcl_return
  
  lcl_return = 0

	'Get the preseelcted family member types
 	sPreselected = getPreselected( iRateId )

  sSql = "SELECT familymemberid, firstname, lastname, birthdate, relationship "
  sSql = sSql & " FROM egov_familymembers "
  sSql = sSql & " WHERE isdeleted = 0 "
  sSql = sSql & " AND belongstouserid = " & iUserID
  sSql = sSql & " ORDER BY birthdate ASC "

 	set oUser = Server.CreateObject("ADODB.Recordset")
	 oUser.Open sSql, Application("DSN"), 3, 1

  if not oUser.eof then

    	iFamilyCount = 0
     lcl_bgcolor  = "#eeeeee"

     while not oUser.eof
        iFamilyCount = iFamilyCount + 1

       'If this is a renewal then show the memberids.
        if iPoolPassID <> "" then
           sSql = "SELECT memberid FROM egov_poolpassmembers "
           sSql = sSql & " WHERE poolpassid = " & iPoolPassID
           sSql = sSql & " AND familymemberid = " & oUser("familymemberid")

          	set oGetMemberID = Server.CreateObject("ADODB.Recordset")
          	oGetMemberID.Open sSql, Application("DSN"), 3, 1

           if not oGetMemberID.eof then
              lcl_member_id = oGetMemberID("memberid")
           else
              lcl_member_id = "&nbsp;"
           end if
          
            oGetMemberID.Close
           set oGetMemberID = nothing
        end if

        response.write "  <tr align=""center"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
        response.write "      <td>" & vbcrlf

      		if InStr(sPreselected, oUser("relationship")) > 0 then
        			lcl_checked = " checked=""checked"""
        else
           lcl_checked = ""
        end if

      		if isnull(oUser("birthdate")) then
           lcl_birthdate = "&nbsp;"
        else
           lcl_birthdate = oUser("birthdate")
        end if

        response.write "          <input type=""checkbox"" name=""passIncl"" value=""" & oUser("familymemberid") & """" & lcl_checked & " />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & oUser("firstname")    & "</td>" & vbcrlf
        response.write "      <td>" & oUser("lastname")     & "</td>" & vbcrlf

       'Show the memberid if this is a renewal
        if iPoolPassID <> "" then
           response.write "      <td>" & lcl_member_id & "</td>" & vbcrlf
        end if

        response.write "      <td>" & oUser("relationship") & "</td>" & vbcrlf
        response.write "      <td>" & lcl_birthdate         & "</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")

      		oUser.movenext
     wend

     lcl_return = iFamilyCount

  end if
  
  oUser.Close
  set oUser = nothing

  getFamilyMembers = lcl_return

end function

'------------------------------------------------------------------------------
function getPreselected( iRateId )
	 lcl_return = ""

	'This builds a string that can be searched to see if the family member is preselected for that rate
  sSQL = "SELECT relation FROM egov_poolpasspreselected WHERE rateid = " & iRateId

 	set oRelation = Server.CreateObject("ADODB.Recordset")
	 oRelation.Open sSQL, Application("DSN"), 3, 1

  if not oRelation.eof then
     while not oRelation.eof
   		   lcl_return = lcl_return & oRelation("relation")
       	oRelation.movenext
     wend
  end if

 	oRelation.close
 	set oRelation = nothing

  getPreselected = lcl_return

end function

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
sub getPoolPassInfo(ByVal lcl_poolpassid, ByRef iUserId, ByRef iRateId, ByRef iMembershipId, ByRef iPeriodId, ByRef lcl_isSeasonal)
  iUserID        = 0
  iRateID        = 0
  iMembershipId  = 0
  iPeriodID      = 0
  lcl_isSeasonal = False

  sSQL = "SELECT P.userid, P.rateid, P.membershipid, P.periodid, MP.is_seasonal "
  sSQL = sSQL & " FROM egov_poolpasspurchases P"
  sSQL = sSQL &      " LEFT OUTER JOIN egov_membership_periods MP ON MP.periodid = P.periodid "
  sSQL = sSQL &      " AND MP.orgid = " & iorgid
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
 public sub getRateInfo(ByVal iRateID, _
                        ByRef nAmount, _
                        ByRef sMessage, _
                        ByRef sDescription, _
                        ByRef iMaxsignups, _
                        ByRef iAttendanceTypeID, _
                        ByRef lcl_rate_residenttype, _
                        ByRef lcl_rate_residenttypedesc, _
                        ByRef lcl_isPunchcard, _
                        ByRef lcl_punchcard_limit)

   nAmount                   = 0
   sDescription              = ""
   iMaxsignups               = ""
   iAttendanceTypeID         = ""
   lcl_rate_residenttype     = ""
   lcl_rate_residenttypedesc = ""
   lcl_isPunchcard           = False
   lcl_punchcard_limit       = 0

   sSQL = "SELECT R.amount, "
   sSQL = sSQL & " R.description AS rate_desc, "
   sSQL = sSQL & " R.maxsignups, "
   sSQL = sSQL & " R.attendancetypeid, "
   sSQL = sSQL & " R.message, "
   sSQL = sSQL & " R.residenttype, "
   sSQL = sSQL & " T.description as residenttype_desc, "
   sSQL = sSQL & " R.isPunchcard, "
   sSQL = sSQL & " R.punchcard_limit "
   sSQL = sSQL & " FROM egov_poolpassrates R, "
   sSQL = sSQL &      " egov_poolpassresidenttypes T "
   sSQL = sSQL & " WHERE R.rateid = " & iRateID
   sSQL = sSQL & " AND R.residenttype = T.resident_type "
   sSQL = sSQL & " AND T.orgid = R.orgid "
   sSQL = sSQL & " AND R.orgid = " & iorgid

  	set oRateInfo = Server.CreateObject("ADODB.Recordset")
 	 oRateInfo.Open sSQL, Application("DSN"), 3, 1

   if not oRateInfo.eof then
      nAmount                   = oRateInfo("amount")
      sMessage                  = oRateInfo("message")
      sDescription              = oRateInfo("rate_desc")
      iMaxsignups               = oRateInfo("maxsignups")
      iAttendanceTypeID         = oRateInfo("attendancetypeid")
      lcl_rate_residenttype     = oRateInfo("residenttype")
      lcl_rate_residenttypedesc = oRateInfo("residenttype_desc")
      lcl_isPunchcard           = oRateInfo("isPunchcard")
      lcl_punchcard_limit       = oRateInfo("punchcard_limit")
   end if

   set oRateInfo = nothing

 end sub

'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
function getFirstUserID()
  lcl_return = 0

  sSQL = "SELECT TOP 1 userid "
  sSQL = sSQL & " FROM egov_users "
  sSQL = sSQL & " WHERE orgid = " & iorgid
  sSQL = sSQL & " AND userregistered = 1 "
  sSQL = sSQL & " AND headofhousehold = 1 "
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
function GetCityName()
	 lcl_return = ""

 	sSQL = "SELECT orgname FROM organizations WHERE orgid = " & iorgid

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
     sSQL = sSQL & " AND orgid = " & iorgid

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
  lcl_return          = ""
		lcl_current_date    = Now()
  lcl_poolpassid      = lcl_poolpassid
  lcl_expiration_date = lcl_expiration_date

  sSQL = "SELECT P.poolpassid, DATEADD(yy, 1, P.paymentdate) AS expiration_date, P.paymentdate, MP.period_interval, "
  sSQL = sSQL & " MP.period_qty, MP.period_type "
  sSQL = sSQL & " FROM egov_poolpasspurchases AS P, egov_poolpassmembers AS ppm, egov_membership_periods AS MP  "
  sSQL = sSQL & " WHERE P.poolpassid = ppm.poolpassid "
  sSQL = sSQL & " AND P.periodid = MP.periodid "
  sSQL = sSQL & " AND ppm.memberid = " & p_member_id

 	set oPool = Server.CreateObject("ADODB.Recordset")
	 oPool.Open sSQL, Application("DSN"), 3, 1

  if not oPool.eof then
     while not oPool.eof
        if UCASE(oPool("period_type")) = "SEASON" then
           if year(date()) = year(oPool("expiration_date")) then
              lcl_poolpassid      = oPool("poolpassid")
        	   		lcl_expiration_date = FormatDateTime(CDate("12/31/" & Year(oPool("expiration_date"))), vbshortdate)
           else
              lcl_poolpassid      = lcl_poolpassid
              lcl_expiration_date = lcl_expiration_date
           end if
      		else
           if year(date()) = year(oPool("expiration_date")) then
              lcl_poolpassid      = oPool("poolpassid")
        	   		lcl_expiration_date = FormatDateTime(DateAdd(oPool("period_interval"),clng(oPool("period_qty")),DateValue(oPool("paymentdate"))), vbshortdate)
           else
              lcl_poolpassid      = lcl_poolpassid
              lcl_expiration_date = lcl_expiration_date
           end if
      		end if
        oPool.movenext
     wend

   		if CDate(lcl_current_date) <= CDate(lcl_expiration_date) then
			     lcl_return = lcl_poolpassid
     end if

  end if

  set oPool = nothing

  getCurrentPoolPassID = lcl_return

end function

'------------------------------------------------------------------------------
function getCheckInResult( p_member_id )
	Dim lcl_member_id, lcl_current_date, lcl_expiration_date

'We first have to determine if this memberid is a "renewal id" or not.
'This means that this memberid could potentially be on more than one PoolPassPurchase.
'If this is the case then we have to determine if the membership is valid at the current time (i.e. membership may be
' currently invalid/expired, but valid for the next season, BUT the user is attempting to use it right now.)

 lcl_poolpassid = getCurrentPoolPassID(p_member_id)

	sSQL = "SELECT count(memberid) AS total_cnt "
 sSQL = sSQL & " FROM egov_poolpassmembers m, egov_poolpasspurchases p "
 sSQL = sSQL & " WHERE m.poolpassid = p.poolpassid "
 sSQL = sSQL & " AND m.memberid = " & p_member_id
 sSQL = sSQL & " AND p.orgid = " & iorgid

	Set oCard = Server.CreateObject("ADODB.Recordset")
	oCard.Open sSQL, Application("DSN"), 3, 1

	If CLng(oCard("total_cnt")) = CLng(0) Then 
  		getCheckInResult = "none"
	Else 
		lcl_current_date = Date()

		sSQL2 = "SELECT DateAdd(yy,1,P.paymentdate) as expiration_date, P.paymentdate, MP.period_interval, MP.period_qty, MP.period_type "
		sSQL2 = sSQL2 & " FROM egov_poolpasspurchases P, egov_poolpassmembers ppm, egov_membership_periods MP "
		sSQL2 = sSQL2 & " WHERE P.poolpassid = ppm.poolpassid "
  sSQL2 = sSQL2 & " AND P.periodid = MP.periodid "
		sSQL2 = sSQL2 & " AND ppm.memberid = " & p_member_id

  if lcl_poolpassid <> "" then
     sSQL2 = sSQL2 & " AND P.poolpassid = " & lcl_poolpassid
  end if

		Set oCardExp = Server.CreateObject("ADODB.Recordset")
		oCardExp.Open sSQL2, Application("DSN"), 3, 1

  if not oCardExp.eof then
   		if UCASE(oCardExp("period_type")) = "SEASON" then
        'lcl_expiration_date = FormatDateTime(CDate("12/31/" & Year(oCardExp("paymentdate"))), vbshortdate)
  	   		lcl_expiration_date = FormatDateTime(CDate("12/31/" & Year(oCardExp("expiration_date"))), vbshortdate)
   		else
     			lcl_expiration_date = FormatDateTime(DateAdd(oCardExp("period_interval"),clng(oCardExp("period_qty")),DateValue(oCardExp("paymentdate"))), vbshortdate)
   		end if

   		if CDate(lcl_current_date) > CDate(lcl_expiration_date) then
			     getCheckInResult = "expired"
   	 else
        if CDate(lcl_current_date) <= CDate(lcl_expiration_date) AND year(lcl_current_date) = year(lcl_expiration_date) then
        			lcl_return = p_member_id

          'Perform one last check to see if the member has already signed in earlier on the same day.
          'If so then sound ALERT otherwise the member is VALID.
           lcl_lastscanned = getLastScannedDate(p_member_id,"")

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

end function

'------------------------------------------------------------------------------
function isRenewedPoolPass(lcl_poolpassid)
  lcl_return = "N"

 'Check to see if this PoolPassID exists as a "previous_poolpassid" on any other purchase(s).
 'If it does then this PoolPassID is considered to be "Renewed".
  sSQL = "SELECT count(poolpassid) as total_count "
  sSQL = sSQL & " FROM egov_poolpasspurchases "
  sSQL = sSQL & " WHERE previous_poolpassid = " & lcl_poolpassid

 	set rs = Server.CreateObject("ADODB.Recordset")
	 rs.Open sSQL, Application("DSN"), 3, 1

  if not rs.eof then
     if rs("total_count") > 0 then
        lcl_return = "Y"
     end if
  end if

  set rs = nothing

  isRenewedPoolPass = lcl_return

end function

'------------------------------------------------------------------------------
function showHideRenewalButton(p_poolpassid)
 lcl_return   = "N"

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
      '3. The current date is EQUAL TO or GREATER THAN the renewal start date.
      '4. The current date is LESS THAN or EQUAL TO the date generated by the "expiration date" + "Days to Renew After Expiration Date".
      '5. Check to make sure that the start date is GREATER THAN or EQUAL TO the current date before allowing a renewal
       if  isRenewedPoolPass(lcl_poolpassid) = "N" _
       AND lcl_isRenewable = True _
       AND datevalue(lcl_current_date) >= datevalue(lcl_renewalstartdate) _
       AND clng(DATEDIFF("d",lcl_current_date,DATEADD("d",lcl_renewalTimeAfterExpire,lcl_expiration_date))) >= clng(0) _
       AND datevalue(lcl_current_date) >= datevalue(oCheckExp("startdate")) then
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
function dtb_debug(p_value)
  sSQLi = "INSERT INTO my_table_dtb(notes) VALUES('" & replace(p_value,"'","''") & "')"
 	set rsi = Server.CreateObject("ADODB.Recordset")
	 rsi.Open sSQLi, Application("DSN"), 3, 1

  set rsi = nothing

end function
'------------------------------------------------------------------------------
' string GetResidentTypeByAddress( iUserid, iorgid )
'------------------------------------------------------------------------------
Function GetResidentTypeByAddress( ByVal iUserid, ByVal iorgid )
	' Try to match the person's address to one of the resident addresses
	Dim sSql, oRs
	
	GetResidentTypeByAddress = "N"

	sSql = "SELECT COUNT(R.residentaddressid) AS hits FROM egov_residentaddresses R, egov_users U"
	sSql = sSql & " WHERE R.orgid = U.orgid AND "
	sSql = sSql & " R.residentstreetnumber + ' ' + R.residentstreetname = U.useraddress AND "
	sSql = sSql & " R.residenttype = 'R' AND "
	sSql = sSql & " R.orgid = " & iorgid & " AND U.userid = " & iUserid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			' Match found
			GetResidentTypeByAddress = "R"
		End If 
	End if

	oRs.Close
	Set oRs = Nothing

End Function 
%>
