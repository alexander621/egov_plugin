<!-- #include file="../includes/common.asp" //-->
<!-- #include file="poolpass_global_functions.asp" //-->
<%
  dim sPurchaseSection, sOrgID, sNameSearch, sUserID, sUserType
  dim sMembershipID, sRateDesc, sPeriodID

  sOrgID        = 0
  sUserID       = 0
  sAdminUserID  = 0
  sMembershipID = 0
  sPeriodID     = 0

  sPurchaseSection = ""
  sNameSearch      = ""
  sUserType        = ""
  sRateDesc        = ""

  if request("orgid") <> "" then
     if not containsApostrophe(request("orgid")) then
        sOrgID = clng(request("orgid"))
     end if
  end if

  if request("purchaseSection") <> "" then
     sPurchaseSection = ucase(request("purchaseSection"))
  end if

 '-----------------------------------------------------------------------------
 'There are 2 layouts for the "right-side" of this purchase screen.
 'LAYOUT #1 is the default.  LAYOUT #2 is enabled if the "purchase_membership_alt_layout" feature is enabled
 '
 'LAYOUT #1: The original layout.  It has a dropdown list of Membership Periods and 
 '           based on the Membership Period selected a list of Membership Rates will be displayed.
 '
 'LAYOUT #2: The alternate layout.  It has a dropdown list of Membership Types (membership rates) and
 '           based on the Membership Type selected a list of Membership Periods will be displayed.
 '           * NOTE: This alternative layout was specifically designed for Montgomery and they are
 '                   are currently the only one that has the feature turned on for them.
 '-----------------------------------------------------------------------------
  if sPurchaseSection = "NAMESEARCH_DROPDOWN_OPTIONS" then

     if request("namesearch") <> "" then
        sNameSearch = request("namesearch")
     end if

     showUserDropdown sOrgID, _
                      sNameSearch

  else
     if request("userid") <> "" then
        if not containsApostrophe(request("userid")) then
           sUserID = clng(request("userid"))
        end if
     end if

     if request("membershipid") <> "" then
        if not containsApostrophe(request("membershipid")) then
           sMembershipID = clng(request("membershipid"))
        end if
     end if

     if request("adminuserid") <> "" then
        if not containsApostrophe(request("adminuserid")) then
           sAdminUserID = clng(request("adminuserid"))
        end if
     end if

     if request("periodid") <> "" then
        if not containsApostrophe(request("periodid")) then
           sPeriodID = clng(request("periodid"))
        end if
     end if

     sUserType = GetUserResidentType(sUserID)

     if sPurchaseSection = "SHOW_USER_INFO" then

        showUserInfo sUserID, _
                     sUserType

     elseif sPurchaseSection = "SHOW_RENEWAL_INFO" then

        showRenewalMembership sOrgID, _
                              sAdminUserID, _
                              sUserID, _
                              sMembershipID, _
                              sPeriodID

     elseif left(sPurchaseSection,14) = "RIGHTSIDEPICKS" then

        if request("rateDesc") <> "" then
           sRateDesc = request("rateDesc")
        else
           sRateDesc = getFirstMembershipRateOption_AltLayout(sOrgID, _
                                                              sUserType, _
                                                              sMembershipId)
        end if

        if sPurchaseSection = "RIGHTSIDEPICKS_BOTTOM" then

           buildRightSidePicks_bottom sOrgID, _
                                      sMembershipID, _
                                      sPeriodID, _
                                      sUserType, _
                                      sRateDesc
        else
           buildRightSidePicks_top sOrgID, _
                                   sMembershipID, _
                                   sPeriodID, _
                                   sUserType, _
                                   sRateDesc
        end if
     end if
  end if

'------------------------------------------------------------------------------
sub showUserDropdown(iOrgID, _
                     iNameSearch)

  dim lcl_nameSearch, lcl_userInfo

  lcl_nameSearch = ""
  lcl_userInfo   = ""

  if trim(iNameSearch) <> "" then
     lcl_nameSearch = ucase(trim(iNameSearch))
     lcl_nameSearch = dbsafe(lcl_nameSearch)
  end if

  if lcl_nameSearch <> "" then
     lcl_nameSearch = "'%" & lcl_nameSearch & "%'"

     sSQL = "SELECT userid, "
     sSQL = sSQL & " userfname, "
     sSQL = sSQL & " userlname, "
     sSQL = sSQL & " useraddress "
     sSQL = sSQL & " FROM egov_users "
     sSQL = sSQL & " WHERE orgid = " & iOrgID
     sSQL = sSQL & " AND isdeleted = 0 "
     sSQL = sSQL & " AND userregistered = 1 "
     sSQL = sSQL & " AND headofhousehold = 1 "
     sSQL = sSQL & " AND userfname IS NOT NULL "
     sSQL = sSQL & " AND userlname IS NOT NULL "
     sSQL = sSQL & " AND userfname <> '' "
     sSQL = sSQL & " AND userlname <> '' "
     'sSQL = sSQL & " AND upper(userfname) like (" & lcl_nameSearch & ") "
     'sSQL = sSQL & "  OR upper(userlname) like (" & lcl_nameSearch & ") "
     'sSQL = sSQL & "  OR upper(userlname) + ', ' + upper(userfname) like (" & lcl_nameSearch & ") "
     sSQL = sSQL & " AND upper(userlname) + ', ' + upper(userfname) like (" & lcl_nameSearch & ") "
     sSQL = sSQL & " ORDER BY userlname, userfname "

    	set oResident = Server.CreateObject("ADODB.Recordset")
   	 oResident.Open sSQL, Application("DSN"), 3, 1

     if not oResident.eof then
        response.write "Select Name:" & vbcrlf
        response.write "<select name=""userid"" id=""userid"" onchange=""updateUserInfo(this.value);"">" & vbcrlf

        do while not oResident.eof
           lcl_selected = ""
           lcl_userInfo = ""

           if oResident("userlname") <> "" then
              lcl_userInfo = oResident("userlname")
           end if

           if oResident("userfname") <> "" then
              if lcl_userInfo <> "" then
                 lcl_userInfo = lcl_userInfo & ", " & oResident("userfname")
              else
                 lcl_userInfo = oResident("userfname")
              end if
           end if

           if oResident("useraddress") <> "" then
              if lcl_userInfo <> "" then
                 lcl_userInfo = lcl_userInfo & " - " & oResident("useraddress")
              else
                 lcl_userInfo = oResident("useraddress")
              end if
           end if

           response.write "  <option value=""" & oResident("userid") & """" & lcl_selected & ">" & lcl_userInfo & "</option>" & vbcrlf

           oResident.movenext
        loop

        response.write "</select>" & vbcrlf

     end if

     oResident.close
     set oResident = nothing
  end if

end sub

'------------------------------------------------------------------------------
sub showUserInfo(iUserID, _
                 iUserType)

  dim sResidentDesc

  sResidentDesc = GetResidentTypeDesc(iUserType)

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
  sSQL = sSQL & " WHERE userid = " & iUserID

 	set oUser = Server.CreateObject("ADODB.Recordset")
	 oUser.Open sSQL, Application("DSN"), 3, 1

  if not oUser.eof then
    'Build the City, State, and Zip display value
     lcl_cityStateZip = ""

     if oUser("usercity") <> "" then
        lcl_cityStateZip = oUser("usercity")
     end if

     if oUser("userstate") <> "" then
        if lcl_cityStateZip <> "" then
           lcl_cityStateZip = lcl_cityStateZip & ", " & oUser("userstate")
        else
           lcl_cityStateZip = oUser("userstate")
        end if
     end if

     if oUser("userzip") <> "" then
        if lcl_cityStateZip <> "" then
           lcl_cityStateZip = lcl_cityStateZip & " " & oUser("userzip")
        else
           lcl_cityStateZip = oUser("userzip")
        end if
     end if

     response.write "<fieldset id=""userInfo"" class=""fieldset"">" & vbcrlf
     response.write "  <legend>User Info</legend>" & vbcrlf

     response.write "<table border=""0"">" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"" valign=""top"">Name:</td>" & vbcrlf
     response.write "      <td width=""60%"">" & oUser("userfname") & " " & oUser("userlname") & "&nbsp;&nbsp;&nbsp;<strong>" & sResidentDesc & "</strong></td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"" valign=""top"">Email:</td>" & vbcrlf
     response.write "      <td>" & oUser("useremail") & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"" valign=""top"">Phone:</td>" & vbcrlf
     response.write "      <td>" & FormatPhone(oUser("userhomephone")) & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"" valign=""top"">Address:</td>" & vbcrlf
     response.write "      <td>" & vbcrlf
     response.write            oUser("useraddress") & "<br />" 

     if oUser("useraddress2") = "" then
      		response.write oUser("useraddress2") & "<br />" & vbcrlf
     end if

     response.write            lcl_cityStateZip & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"" valign=""top"">Business:</td>" & vbcrlf
     response.write "      <td>" & oUser("userbusinessname") & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "</table>" & vbcrlf
     response.write "</fieldset>" & vbcrlf
     response.write "<input type=""hidden"" name=""iuserid"" id=""iuserid"" value="""   & iUserID   & """ />" & vbcrlf
     response.write "<input type=""hidden"" name=""usertype"" id=""usertype"" value=""" & iUserType & """ />" & vbcrlf
  end if

 	oUser.close
 	set oUser = nothing

end sub

'------------------------------------------------------------------------------
sub buildRightSidePicks_top(iOrgID, _
                            iMembershipID, _
                            iPeriodID, _
                            iUserType, _
                            iRateDesc)

  dim lcl_orghasfeature_purchase_membership_alt_layout

  lcl_orghasfeature_purchase_membership_alt_layout = orghasfeature("purchase_membership_alt_layout")

  if lcl_orghasfeature_purchase_membership_alt_layout then
     'response.write "<div id=""rightpicks_altlayout"">" & vbcrlf

    'Membership Rates Dropdown
     response.write "  <p id=""membershipRatesDropdown"">" & vbcrlf
     response.write "    Membership Types:<br />" & vbcrlf
                         showMembershipRates_AltLayout iOrgID, _
                                                       iUserType, _
                                                       iMembershipID, _
                                                       iRateDesc
     response.write "  </p>" & vbcrlf
  else
    'Membership Periods Dropdown
     response.write "  <p>" & vbcrlf
     response.write "    Membership Period:<br />" & vbcrlf
                         showPeriodPicksForMembership iOrgID, _
                                                      iMembershipID, _
                                                      iPeriodID
     response.write "  </p>" & vbcrlf
  end if

end sub

'------------------------------------------------------------------------------
sub buildRightSidePicks_bottom(iOrgID, _
                               iMembershipID, _
                               iPeriodID, _
                               iUserType, _
                               iRateDesc)

  dim lcl_orghasfeature_purchase_membership_alt_layout

  lcl_orghasfeature_purchase_membership_alt_layout = orghasfeature("purchase_membership_alt_layout")

  if lcl_orghasfeature_purchase_membership_alt_layout then
    'Membership Periods Checkboxes
     response.write "  <div class=""shadow"">" & vbcrlf
                         showPeriodPicksForMembership_AltLayout iOrgID, _
                                                                iMembershipID, _
                                                                iPeriodID, _
                                                                iUserType, _
                                                                iRateDesc
     response.write "  </div>" & vbcrlf
  else
    'Membership Rates Checkboxes
     response.write "  <div class=""shadow"">" & vbcrlf
                         showMembershipResidentRates iOrgID, _
                                                     iUserType, _
                                                     iMembershipID, _
                                                     iPeriodID
     response.write "  </div>" & vbcrlf
  end if

end sub

'------------------------------------------------------------------------------
sub showMembershipRates(iOrgID, _
                        iResidentType, _
                        iUserType, _
                        iMembershipID, _
                        iPeriodID)

		dim sDisabled, sSQL, oRates, bPreSelect, iRow, iCurrYear
  dim sOrgID, lcl_residenttype, lcl_usertype

		iRow   = 0
  sOrgID = 0

		sDisabled         = ""
		bPreSelect        = false
		iCurrYear         = Year(Now())
  lcl_checked_rates = ""
  lcl_usertype      = ""
  lcl_residenttype  = ""

		if iUserType <> iResidentType then
  			sDisabled = " disabled"
		end if

		if sDisabled = "" and iRow = 1 then
 				lcl_checked_rates = " checked=""checked"""
  end if

  if iResidentType <> "" then
     lcl_residenttype = dbsafe(iResidentType)
  end if

  if iUserType <> "" then
     lcl_usertype = dbsafe(iUserType)
  end if

  lcl_residenttype = "'" & lcl_residenttype & "'"
  lcl_usertype     = "'" & lcl_usertype     & "'"

		sSQL = "SELECT R.rateid, "
  sSQL = sSQL & " R.description, "
  sSQL = sSQL & " R.amount , "
  sSQL = sSQL & " MP.period_desc "
  sSQL = sSQL & " FROM egov_poolpassrates R, "
  sSQL = sSQL &      " egov_membership_periods MP "
		sSQL = sSQL & " WHERE R.periodid = MP.periodid "
		sSQL = sSQL & " AND R.orgid = MP.orgid "
  sSQL = sSQL & " AND R.isEnabled = 1 "
		sSQL = sSQL & " AND R.orgid = "        & iOrgID
		sSQL = sSQL & " AND R.residenttype = " & lcl_residenttype
		sSQL = sSQL & " AND R.membershipid = " & iMembershipId
		sSQL = sSQL & " AND MP.periodid = "    & iPeriodid
  sSQL = sSQL & " AND R.residenttype = " & lcl_usertype
		sSQL = sSQL & " ORDER BY R.displayorder"

		set oRates = Server.CreateObject("ADODB.Recordset")
		oRates.Open sSQL, Application("DSN"), 3, 1
		
		do while not oRates.eof
  			iRow = iRow + 1

  			response.write "  <tr>" & vbcrlf
     response.write "      <td><input type=""radio""  "
	if iRow = 1 then response.write " checked "
     response.write " name=""rateid"" id=""rateid_" & oRates("rateid") & """ value=""" & oRates("rateid") & """" & sDisabled & lcl_checked_rates & " onclick=""enableDisableContinueButton('" & oRates("rateid") & "');updateRenewalInfo();"" />" & oRates("description") & "</td>" & vbrlf
     response.write "      <td class=""amount"">" & FormatCurrency(oRates("amount")) & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf

  			oRates.MoveNext
  loop

		oRates.Close
		set oRates = nothing
		
 end sub

'------------------------------------------------------------------------------
sub ShowMembershipRates_AltLayout(iOrgID, _
                                  iUserType, _
                                  iMembershipID, _
                                  iRateDesc)

		dim sDisabled, sSQL, oRates, bPreSelect, sRow, iCurrYear
  dim lcl_usertype

		sRow         = 0 
  lcl_usertype = ""

  lcl_rate_desc = getDistinctRateDescList(iOrgID, _
                                          iMembershipID, _
                                          iUserType)

  if iUserType <> "" then
     lcl_usertype = dbsafe(iUserType)
  end if

  lcl_usertype = "'" & lcl_usertype & "'"

  sSQL = "SELECT distinct R.description "
  sSQL = sSQL & " FROM egov_poolpassrates R, "
  sSQL = sSQL &      " egov_membership_periods MP "
  sSQL = sSQL & " WHERE R.periodid = MP.periodid "
  sSQL = sSQL & " AND R.orgid = MP.orgid "
  sSQL = sSQL & " AND R.isEnabled = 1 "
  sSQL = sSQL & " AND R.orgid = "         & iOrgID
  sSQL = sSQL & " AND R.membershipid = "  & iMembershipID
  sSQL = sSQL & " AND R.residenttype = "  & lcl_usertype
  sSQL = sSQL & " AND R.description IN (" & lcl_rate_desc & ") "
  sSQL = sSQL & " ORDER BY R.description "
'response.write sSQL & "<br /><br />" & vbcrlf
  set oRates = Server.CreateObject("ADODB.Recordset")
  oRates.Open sSQL, Application("DSN"), 3, 1

  if not oRates.eof then
     response.write "<select name=""rateDesc"" id=""rateDesc"" onchange=""updateMembershipOptions();"">" & vbcrlf

     do while not oRates.eof
     			iRow = iRow + 1

        lcl_selected_rates = ""

     		 if ucase(oRates("description")) = ucase(iRateDesc) then
      		 		lcl_selected_rates = " selected=""selected"""
        end if

        response.write "  <option value=""" & oRates("description") & """" & lcl_selected_rates & ">" & oRates("description") & "</option>" & vbcrlf

     			oRates.movenext
     loop

     response.write "</select>" & vbcrlf

  end if

		oRates.Close
		set oRates = nothing
		
 end sub

'------------------------------------------------------------------------------
function getDistinctRateDescList(iOrgID, _
                                 iMembershipID, _
                                 iUserType)

  dim lcl_return, lcl_usertype, sSQL

  lcl_return   = "0"
  lcl_usertype = ""

  if iUserType <> "" then
     lcl_usertype = iUserType
  end if

  lcl_usertype = "'" & lcl_usertype & "'"

 'Get a distinct list of rate descriptions
		sSQL = "SELECT distinct R.description "
  sSQL = sSQL & " FROM egov_poolpassrates R, "
  sSQL = sSQL &      " egov_membership_periods MP "
		sSQL = sSQL & " WHERE R.periodid = MP.periodid "
		sSQL = sSQL & " AND R.orgid = MP.orgid "
  sSQL = sSQL & " AND R.isEnabled = 1 "
		sSQL = sSQL & " AND R.orgid = "        & iOrgID
		sSQL = sSQL & " AND R.membershipid = " & iMembershipID
  sSQL = sSQL & " AND R.residenttype = " & lcl_usertype

		set oDistinctRates = Server.CreateObject("ADODB.Recordset")
		oDistinctRates.Open sSQL, Application("DSN"), 3, 1

  if not oDistinctRates.eof then
     lcl_rate_desc = ""

     do while not oDistinctRates.eof
        if lcl_rate_desc <> "" then
           lcl_rate_desc = lcl_rate_desc & ",'" & oDistinctRates("description") & "'"
        else
           lcl_rate_desc = "'" & oDistinctRates("description") & "'"
        end if

        oDistinctRates.movenext
     loop

     if lcl_rate_desc <> "" then
        lcl_return = lcl_rate_desc
     end if

  end if

  'oDistinctRates.close
  set oDistinctRates = nothing

  getDistinctRateDescList = lcl_return

end function

 '-----------------------------------------------------------------------------
sub ShowPeriodPicksForMembership(iOrgID, _
                                 iMembershipID, _
                                 iPeriodID)
		Dim sSQL, oPeriods

	'Get the Periods
		sSQL = "SELECT DISTINCT P.periodid, "
  sSQL = sSQL & " P.period_desc "
  sSQL = sSQL & " FROM egov_membership_periods P, "
  sSQL = sSQL &      " egov_poolpassrates R "
  sSQL = sSQL & " WHERE R.periodid = P.periodid "
  sSQL = sSQL & " AND R.isEnabled = 1 "
		sSQL = sSQL & " AND P.orgid = "        & iOrgId
  sSQL = sSQL & " AND R.membershipid = " & iMembershipID
  sSQL = sSQL & " ORDER BY P.period_desc DESC"

		set oPeriods = Server.CreateObject("ADODB.Recordset")
		oPeriods.Open sSQL, Application("DSN"), 3, 1
		
		if not oPeriods.eof then
  			response.write "<select name=""periodid"" id=""periodid"" onchange=""updateMembershipOptions();"">" & vbcrlf

   		do while not oPeriods.eof
        lcl_selected_period = ""

        if CLng(iPeriodId) = CLng(oPeriods("periodid")) then
      					lcl_selected_period = " selected=""selected"""
        end if

    				response.write "  <option value=""" & oPeriods("periodid") & """" & lcl_selected_period & ">" & oPeriods("period_desc") & "</option>" & vbcrlf

    				oPeriods.movenext
     loop

   		response.write "</select>" & vbcrlf
  end if

		oPeriods.close
		set oPeriods = nothing

 end sub


'------------------------------------------------------------------------------
sub ShowPeriodPicksForMembership_AltLayout(iOrgID, _
                                           iMembershipID, _
                                           iPeriodId, _
                                           iUserType, _
                                           iRateDesc)

		dim sSQL, oPeriods, lcl_usertype, lcl_ratedesc, sOrgID

  sOrgID    = 0

  lcl_usertype = ""
  lcl_ratedesc = ""

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iUserType <> "" then
     lcl_usertype = iUserType
  end if

  if iRateDesc <> "" then
     lcl_ratedesc = ucase(iRateDesc)
     lcl_ratedesc = dbsafe(lcl_ratedesc)
  end if

  lcl_usertype = "'" & lcl_usertype & "'"
  lcl_ratedesc = "'" & lcl_ratedesc & "'"

 'Get the Periods
  sSQL = "SELECT DISTINCT P.periodid, "
  sSQL = sSQL & " P.period_desc, "
  sSQL = sSQL & " R.amount "
  sSQL = sSQL & " FROM egov_membership_periods P, "
  sSQL = sSQL &      " egov_poolpassrates R "
  sSQL = sSQL & " WHERE R.periodid = P.periodid "
  sSQL = sSQL & " AND R.isEnabled = 1 "
  sSQL = sSQL & " AND P.orgid = "              & sOrgID
  sSQL = sSQL & " AND R.membershipid = "       & iMembershipID
  sSQL = sSQL & " AND R.residenttype = "       & lcl_usertype
  sSQL = sSQL & " AND UPPER(R.description) = " & lcl_ratedesc
  sSQL = sSQL & " ORDER BY P.period_desc DESC"
'response.write sSQL & "<br /><br />" & vbcrlf
		set oPeriods = Server.CreateObject("ADODB.Recordset")
		oPeriods.Open sSQL, Application("DSN"), 3, 1

		response.write "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" class=""tableright"">" 
 	response.write "  <tr><th colspan=""2"">&nbsp;Membership Choices</th></tr>" 
		
		if not oPeriods.eof then
		  	'response.write "  <tr><th colspan=""2""> &nbsp; " & oTypes("description") & "</th></tr>" 

     i = 0
     lcl_cound_checked = 0

     do while not oPeriods.eof
        i = i + 1
    				if CLng(iPeriodId) = CLng(oPeriods("periodid")) then
				      	lcl_checked_period = " checked=""checked"""
        else
				      	lcl_checked_period = ""
     			end if

        lcl_rateid = getRateID(iOrgID, _
                               iMembershipId, _
                               iUserType, _
                               iRateDesc, _
                               oPeriods("periodid"))

     			response.write "  <tr>" 
        'response.write "      <td><input type=""radio"" name=""periodid"" id=""periodid_" & oPeriods("periodid") & """ value=""" & oPeriods("periodid") & """" & sDisabled & lcl_checked_period & " onclick=""updateRateID('" & lcl_rateid & "');enableDisableContinueButton('" & oPeriods("periodid") & "');submitPoolPassForm();"" />" & oPeriods("period_desc") & "</td>" 
        response.write "      <td><input type=""radio"" name=""periodid"" id=""periodid_" & oPeriods("periodid") & """ value=""" & oPeriods("periodid") & """" & sDisabled & lcl_checked_period & " onclick=""updateRateID('" & lcl_rateid & "');enableDisableContinueButton('" & oPeriods("periodid") & "');updateRenewalInfo();"" />" & oPeriods("period_desc") & "</td>" 
        response.write "      <td class=""amount"">" & FormatCurrency(oPeriods("amount")) & "</td>" 
        response.write "  </tr>" 

        if lcl_checked_period <> "" then
           lcl_count_checked = lcl_count_checked + 1
           lcl_update_rateid = lcl_rateid
        else
           lcl_count_checked = lcl_count_checked
           lcl_update_rateid = lcl_update_rateid
        end if

    				oPeriods.MoveNext
  			loop

     if lcl_count_checked > 0 then
        response.write "<script language=""javascript"">" 
        response.write "  updateRateID('" & lcl_update_rateid & "');" 
        response.write "</script>" 
     end if

  else
     response.write "  <tr><td colspan=""2"">No Membership Rates Available</td></tr>" 
  end if

		response.write "</table>" & vbrlf

		oPeriods.Close
		set oPeriods = nothing

 end sub

'------------------------------------------------------------------------------
function getRateID(iOrgID, _
                   iMembershipId, _
                   iUserType, _
                   iRateDesc, _
                   iPeriodID)

  dim sSQL, lcl_return, lcl_orgid, lcl_usertype, lcl_ratedesc

  lcl_return = 0
  lcl_orgid  = 0

  lcl_usertype = ""
  lcl_ratedesc = ""

  if iOrgID <> "" then
     lcl_orgid = clng(iOrgID)
  end if

  if iUserType <> "" then
     lcl_usertype = iUserType
  end if

  if iRateDesc <> "" then
     lcl_ratedesc = ucase(iRateDesc)
  end if

  lcl_usertype = "'" & lcl_usertype & "'"
  lcl_ratedesc = "'" & lcl_ratedesc & "'"

	'Get the RateID
		sSQL = "SELECT DISTINCT R.rateid "
  sSQL = sSQL & " FROM egov_membership_periods P, "
  sSQL = sSQL &      " egov_poolpassrates R "
		sSQL = sSQL & " WHERE R.periodid = P.periodid "
  sSQL = sSQL & " AND P.orgid = "              & lcl_orgid
  sSQL = sSQL & " AND R.membershipid = "       & iMembershipID
  sSQL = sSQL & " AND R.residenttype = "       & lcl_usertype
  sSQL = sSQL & " AND UPPER(R.description) = " & lcl_ratedesc
  sSQL = sSQL & " AND P.periodid = "           & iPeriodID

		set oRateID = Server.CreateObject("ADODB.Recordset")
		oRateID.Open sSQL, Application("DSN"), 3, 1

  if not oRateID.eof then
     lcl_return = oRateID("rateid")
  end if

  oRateID.close
  set oRateID = nothing

  getRateID = lcl_return

end function

'------------------------------------------------------------------------------
sub ShowMembershipResidentRates(iOrgID, _
                                iUserType, _
                                iMembershipID, _
                                iPeriodID)

		dim sSQL, oTypes, lcl_usertype, lcl_orgid

  lcl_orgid = 0

  lcl_usertype = ""

  if iOrgID <> "" then
     lcl_orgid = clng(iOrgID)
  end if

  if iUserType <> "" then
     lcl_usertype = iUserType
  end if

  lcl_usertype = "'" & lcl_usertype & "'"

  sSQL = "SELECT DISTINCT T.resident_type, "
  sSQL = sSQL & " T.description, "
  sSQL = sSQL & " T.displayorder "
  sSQL = sSQL & " FROM egov_poolpassresidenttypes T, "
  sSQL = sSQL &      " egov_poolpassrates R "
  sSQL = sSQL & " WHERE T.resident_type = R.residenttype "
  sSQL = sSQL & " AND R.isEnabled = 1 "
  sSQL = sSQL & " AND T.orgid = "        & lcl_orgid
  sSQL = sSQL & " AND R.membershipid = " & iMembershipID
  sSQL = sSQL & " AND R.periodid = "     & iPeriodID
  sSQL = sSQL & " AND R.residenttype = " & lcl_usertype
  sSQL = sSQL & " ORDER BY T.displayorder "

  set oTypes = Server.CreateObject("ADODB.Recordset")
  oTypes.Open sSQL, Application("DSN"), 3, 1

if oTypes.EOF then response.write "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" class=""tableright""><tr><td>No Rates</td></tr></table></div>"
  do while not oTypes.eof
  			response.write "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" class=""tableright"">" & vbcrlf
		  	response.write "  <tr><th colspan=""2""> &nbsp; " & oTypes("description") & "</th></tr>" & vbcrlf

  		'Get the rates here 
		  	showMembershipRates iOrgID, _
                         oTypes("resident_type"), _
                         iUserType, _
                         iMembershipID, _
                         iPeriodid

  			response.write "</table>" & vbrlf

  			oTypes.MoveNext
		loop

		oTypes.Close
		set oTypes = nothing

 end sub

'------------------------------------------------------------------------------
sub showRenewalMembership(iOrgID, _
                          iAdminUserID, _
                          p_userid, _
                          p_membershipid, _
                          p_periodid)

  dim sSQL, iPassCount, iRowCount
  dim lcl_orghasfeature_membership_renewals, lcl_userhaspermission_membership_renewals
  dim lcl_membershiprenewals_feature

		iPassCount = 0
  iRowCount  = 0

  lcl_orghasfeature_membership_renewals = false
  lcl_membershiprenewals_feature        = false

'Check for org features
 lcl_orghasfeature_membership_renewals = orghasfeature("membership_renewals")

'Check for user permissions
 lcl_userhaspermission_membership_renewals = userhaspermission(iAdminUserID, _
                                                               "membership_renewals")

'Check to see if the org has the feature turned-on and the user has it assigned
 lcl_membershiprenewals_feature = "N"

 if lcl_orghasfeature_membership_renewals AND lcl_userhaspermission_membership_renewals then
    lcl_membershiprenewals_feature = "Y"
 end if


 'Build query to check for a renewal for user selected
		sSQL = "SELECT P.poolpassid, "
  sSQL = sSQL & " P.rateid, "
  sSQL = sSQL & " P.periodid, "
  sSQL = sSQL & " U.userfname, "
  sSQL = sSQL & " U.userlname, "
  sSQL = sSQL & " P.paymentamount, "
  sSQL = sSQL & " P.paymenttype, "
  sSQL = sSQL & " P.paymentdate, "
		sSQL = sSQL & " P.paymentlocation, "
  sSQL = sSQL & " P.paymentresult, "
  sSQL = sSQL & " M.membershipdesc, "
  sSQL = sSQL & " MP.period_desc, "
  sSQL = sSQL & " P.previous_poolpassid "
		sSQL = sSQL & " FROM egov_users U, "
  sSQL = sSQL &      " egov_memberships M, "
  sSQL = sSQL &      " egov_membership_periods MP, "
  sSQL = sSQL &      " egov_poolpasspurchases P "
		sSQL = sSQL & " WHERE P.orgid = " & iOrgID
  sSQL = sSQL & " AND P.paymentresult <> 'Pending' "
		sSQL = sSQL & " AND P.paymentresult <> 'Declined' "
  sSQL = sSQL & " AND U.userid = P.userid "
  sSQL = sSQL & " AND P.membershipid = M.membershipid "
  sSQL = sSQL & " AND P.periodid = MP.periodid "
  'sSQL = sSQL & " AND (P.paymentdate >= '" & fromDate & "' AND P.paymentdate < '" & toDate & "') "
  sSQL = sSQL & " AND U.userid = "       & p_userid
  sSQL = sSQL & " AND P.membershipid = " & p_membershipid
  sSQL = sSQL & " AND P.periodid = "     & p_periodid
		sSQL = sSQL & " ORDER BY P.poolpassid "

 	set oRequests = Server.CreateObject("ADODB.Recordset")
	 oRequests.Open sSQL, Application("DSN"), 3, 1

  if not oRequests.eof then
     bgcolor            = "#eeeeee"
     lcl_showheader_row = "Y"
     lcl_close_renewal  = "N"
     lcl_isRateEnabled  = true

     'response.write "<fieldset class=""fieldset"">" & vbcrlf
     'response.write "  <legend>Membership(s) Available for Renewal</legend>" & vbcrlf

   		do while not oRequests.eof
  	   		iPassCount                = iPassCount + 1
   		  	iRowCount                 = iRowCount + 1
        lcl_showHideRenewalButton = "Y"

       'Retrieve the RATE info
        lcl_rate_description = getRateDesc(iOrgID, _
                                           oRequests("rateid"))

       'Determine if the Renewal column/button are displayed
        if lcl_membershiprenewals_feature = "Y" then
           lcl_showHideRenewalButton = showHideRenewalRow(iOrgID, _
                                                           oRequests("poolpassid"))
        end if

       'Display only the record that is to be renewed.
        if lcl_showHideRenewalButton = "Y" then

          'Show the first row of column headers.
           if lcl_showheader_row = "Y" then
              lcl_showheader_row = "N"
              lcl_close_renewal  = "Y"

              'lcl_style_width = " style=""width:600px"""
              lcl_style_width = " style=""width:100% !important;"""

              response.write "<div id=""membershipRenewals"">" & vbcrlf
              response.write "<fieldset class=""fieldset"">" & vbcrlf
              response.write "  <legend>Membership(s) Available for Renewal&nbsp;</legend>" & vbcrlf
              response.write "<br />" & vbcrlf
              response.write "<div class=""shadow""" & lcl_style_width & ">" & vbcrlf
              response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tablelist""" & lcl_style_width & ">" & vbcrlf
              response.write "  <tr class=""tablelist"" align=""left"">" & vbcrlf
              response.write "      <th>&nbsp;</th>" & vbcrlf
              response.write "      <th style=""text-align: center"">Pass<br />ID</th>" & vbcrlf
              response.write "      <th style=""white-space: nowrap"">Membership Type</th>" & vbcrlf
              response.write "      <th>Purchase Date</th>" & vbcrlf
              response.write "      <th>Purchaser</th>" & vbcrlf
              response.write "      <th>Payment<br />Amount</th>" & vbcrlf
              response.write "      <th>Payment Method</th>" & vbcrlf
              response.write "      <th>Status</th>" & vbcrlf
              response.write "      <th id=""column_renewal"">Renewal<br />of Pass ID</th>" & vbcrlf
              'response.write "      <th colspan=""2"">&nbsp;</th>" & vbcrlf
              response.write "      <th>&nbsp;</th>" & vbcrlf
              response.write "  </tr>" & vbcrlf
           else
              lcl_showheader_row = lcl_showheader_row
              lcl_close_renewal  = lcl_close_renewal
           end if

           lcl_row_mouseover = ""
           lcl_row_mouseout  = ""
           lcl_td_mouseover  = ""
           lcl_td_mouseout   = ""
           lcl_onclick       = " onClick=""location.href='poolpass_details.asp?iPoolPassId=" & oRequests("poolpassid") & "';"""

        			response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & bgcolor & """" & lcl_row_mouseover & lcl_row_mouseout & ">" & vbcrlf
  	      		response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">&nbsp;</td>" & vbcrlf
   		     	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & " align=""left"">" & oRequests("poolpassid") & "</td>" & vbcrlf
     	   		response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & oRequests("membershipdesc") & " &ndash; " & oRequests("period_desc") & "<br />" & lcl_rate_description & "</td>" & vbcrlf
      		  	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & DateValue(oRequests("paymentdate")) & "</td>" & vbcrlf
  	      		response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & oRequests("userfname") & " " & oRequests("userlname") & "</td>" & vbcrlf
        			response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & formatcurrency(oRequests("paymentamount"),2) & "</td>" & vbcrlf

        			cTotalAmount = cTotalAmount + CDbl(oRequests("paymentamount"))

        			response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & MakeProper(oRequests("paymentlocation")) & " &mdash; " & MakeProper(oRequests("paymenttype")) & "</td>" & vbcrlf
		        	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & oRequests("paymentresult") & "</td>" & vbcrlf
   		     	response.write "      <td" & lcl_td_mouseover & lcl_td_mouseout & lcl_onclick & ">" & oRequests("previous_poolpassid") & "</td>" & vbcrlf
     	   		response.write "      <td><input type=""button"" name=""renew"" id=""button_renew_" & oRequests("poolpassid") & """ value=""Renew"" onclick=""renewPass('" & oRequests("poolpassid") & "');"" class=""button"" /></td>" & vbcrlf
   		     	response.write "</tr>" & vbcrlf

           bgcolor = changeBGColor(bgcolor,"#eeeeee","#ffffff")

        end if

        oRequests.movenext
     loop

     if lcl_close_renewal = "Y" then
      		response.write "</table>" & vbcrlf
        response.write "</div>" & vbcrlf
        response.write "</fieldset>" & vbcrlf
        response.write "</div>" & vbcrlf
        response.write "</p>" & vbcrlf
     end if

  end if

  oRequests.close
  set oRequests = nothing 

'response.write "here"

end sub

'------------------------------------------------------------------------------
function getRateDesc(iOrgID, _
                     iRateID)

  dim lcl_return, sSQL, lcl_description, lcl_rate_description, lcl_rate_residenttype, lcl_residenttype_desc

  lcl_return            = ""
  lcl_description       = ""
  lcl_rate_description  = ""
  lcl_rate_residenttype = ""
  lcl_residenttype_desc = ""

 'Retrieve the RATE information if it exists
  sSQL = "SELECT description, "
  sSQL = sSQL & " residenttype "
  sSQL = sSQL & " FROM egov_poolpassrates "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
  sSQL = sSQL & " AND rateid = "  & iRateID

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
  sSQL = sSQL & " WHERE orgid = " & iOrgID
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

  if lcl_residenttype_desc <> "" then
     if lcl_rate_description <> "" then
        lcl_rate_description = lcl_rate_description & " - " & lcl_residenttype_desc
     end if
  end if

  lcl_return = lcl_rate_description

  getRateDesc = lcl_return

end function

'------------------------------------------------------------------------------
function showHideRenewalRow(iOrgID, _
                            iPoolPassID)

 dim sSQL, sSQL2, lcl_return, lcl_poolpassid

 lcl_return     = "N"
 lcl_poolpassid = 0

 if iPoolPassID <> "" then
    lcl_poolpassid = clng(iPoolPassID)
 end if

'We first have to determine if this memberid is a "renewal id" or not.
'This means that this memberid could potentially be on more than one PoolPassPurchase.
'If this is the case then we have to determine if the membership is valid at the current time (i.e. membership may be
' currently invalid/expired, but valid for the next season, BUT the user is attempting to use it right now.)
	sSQL = "SELECT count(poolpassid) AS total_cnt "
 sSQL = sSQL & " FROM egov_poolpasspurchases p "
 sSQL = sSQL & " WHERE p.poolpassid = " & lcl_poolpassid

	set oCheck = Server.CreateObject("ADODB.Recordset")
	oCheck.Open sSQL, Application("DSN"), 3, 1

	if clng(oCheck("total_cnt")) = clng(0) then
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
       lcl_isRenewable            = false
       lcl_renewalTimeAfterExpire = 0
       lcl_renewalstartdate       = DATEADD("d",-1,lcl_current_date)

       if oCheckExp("expirationdate") <> "" then
          lcl_expiration_date = CDate(oCheckExp("expirationdate"))
       else
          'lcl_expiration_date = oMembership.getExpirationDate(oCheckExp("periodid"),oCheckExp("startdate"))
          lcl_expiration_date = getMembershipExpirationDate(iOrgID, _
                                                            oCheckExp("periodid"), _
                                                            oCheckExp("startdate"))
       end if

       if oCheckExp("isRenewable") <> "" then
          lcl_isRenewable = oCheckExp("isRenewable")
       end if

      'If there isn't a Renewal Start Date then set it to the current date -1 day.
       if oCheckExp("renewalstartdate") <> "" then
          lcl_renewalstartdate = oCheckExp("renewalstartdate")
       end if

       if oCheckExp("renewalTimeAfterExpire") <> "" then
          lcl_renewalTimeAfterExpire = oCheckExp("renewalTimeAfterExpire")
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

 showHideRenewalRow = lcl_return

end function

'------------------------------------------------------------------------------
 function getMembershipExpirationDate(iOrgID, _
                                      iPeriodID, _
                                      iStartDate)

  dim lcl_expirationdate, sSQLe, sStartDate

  lcl_expirationdate = ""

  if iStartDate <> "" then
     sStartDate = dbsafe(iStartDate)
  end if

  sStartDate = "'" & sStartDate & "'"

  if p_startdate <> "" then
    'Calculate the expiration date
     sSQLe = "SELECT CAST(dbo.fn_getMembershipExpirationdate(MP.is_seasonal,MP.period_interval,MP.period_qty," & sStartDate & ") AS datetime) AS expirationdate "
     sSQLe = sSQLe & " FROM egov_membership_periods MP "
     sSQLe = sSQLe & " WHERE MP.orgid = " & iOrgID
     sSQLe = sSQLe & " AND MP.periodid = " & iPeriodID

     set rse = Server.CreateObject("ADODB.Recordset")
     rse.Open sSQLe, Application("DSN"), 3, 1

     if not rse.eof then
        lcl_expirationdate = rse("expirationdate")
     'else
        'lcl_expirationdate = DATEADD(yy,1,p_startdate)
     end if

     set rse = nothing

   end if

   getMembershipExpirationDate = lcl_expirationdate

 end function

'------------------------------------------------------------------------------
function getFirstMembershipRateOption_AltLayout(iOrgID, _
                                                iUserType, _
                                                iMembershipId)

  dim lcl_return, sSQL, lcl_usertype, sMembershipID, lcl_rate_desc

  lcl_return    = ""
  lcl_rate_desc = ""
  lcl_user_type = ""
  sMembershipID = 0

  lcl_rate_desc = getDistinctRateDescList(iOrgID, _
                                          iMembershipID, _
                                          iUserType)

  if iUserType <> "" then
     lcl_usertype = dbsafe(iUserType)
  end if

  lcl_usertype = "'" & lcl_usertype & "'"

 'Now grab the first rate.
  sSQL = "SELECT R.description, "
  sSQL = sSQL & " R.amount, "
  sSQL = sSQL & " R.displayorder "
  sSQL = sSQL & " FROM egov_poolpassrates R, "
  sSQL = sSQL &      " egov_membership_periods MP "
  sSQL = sSQL & " WHERE R.periodid = MP.periodid "
  sSQL = sSQL & " AND R.orgid = MP.orgid "
  sSQL = sSQL & " AND R.isEnabled = 1 "
  sSQL = sSQL & " AND R.orgid = "         & iOrgID
  sSQL = sSQL & " AND R.membershipid = "  & iMembershipID
  sSQL = sSQL & " AND R.residenttype = "  & lcl_usertype
  sSQL = sSQL & " AND R.description IN (" & lcl_rate_desc & ") "
  sSQL = sSQL & " ORDER BY R.displayorder, R.description "

  set oFirstRate = Server.CreateObject("ADODB.Recordset")
  oFirstRate.Open sSQL, Application("DSN"), 3, 1

  if not oFirstRate.eof then
     lcl_return = oFirstRate("description")
  end if

  oFirstRate.Close
  set oFirstRate = nothing

  getFirstMembershipRateOption_AltLayout = lcl_return
		
end function

%>
