<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: classMembership.asp
' AUTHOR: Steve Loar
' CREATED: 07/11/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the membership class
'
' MODIFICATION HISTORY
' 1.0   07/11/2006  Steve Loar - Initial code 
' 2.0	07/28/2010	Steve Loar - Changes for Point and Pay payments
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Class classMembership
	Private iMembershipId, iOrgId

	Private Sub Class_Initialize( )
		iOrgId = Session("OrgID")
	End Sub 

	Public Property Let MembershipId( sMembershipId )
		iMembershipId = sMembershipId
	End Property 

	Public Property Get MembershipId()
		MembershipId = iMembershipId
	End Property 


 '-----------------------------------------------------------------------------
	' Function DBsafe( strDB )
 '-----------------------------------------------------------------------------
	Private Function DBsafe( strDB )
	  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
	  DBsafe = Replace( strDB, "'", "''" )
	End Function


 '-----------------------------------------------------------------------------
	' Function GetInitialMembershipId( )
 '-----------------------------------------------------------------------------
	Public Function GetInitialMembershipId( )
		Dim sSql, oMember

		sSql = "Select MIN(membershipid) as membershipid FROM egov_memberships WHERE orgid = " & iOrgId 
		
		Set oMember = Server.CreateObject("ADODB.Recordset")
		oMember.Open sSql, Application("DSN"), 3, 1
		
		If IsNull(oMember("membershipid")) Then
			GetInitialMembershipId = 0
		Else
			GetInitialMembershipId = oMember("membershipid")
		End If 
		
		oMember.close
		Set oMember = Nothing
	End Function 


 '-----------------------------------------------------------------------------
	' Public Sub SetMembershipId( sMembership )
 '-----------------------------------------------------------------------------
	Public Sub SetMembershipId( sMembership )
		iMembershipId = GetMembershipId( sMembership )
	End Sub 

	Public Function GetFirstMembershipType()
		sSql = "Select top 1 membership FROM egov_memberships WHERE orgid = " & iOrgId & " "
		
		Set oMember = Server.CreateObject("ADODB.Recordset")
		oMember.Open sSql, Application("DSN"), 3, 1
		
		GetFirstMembershipType = oMember("membership")
		
		oMember.close
		Set oMember = Nothing

	End Function


 '-----------------------------------------------------------------------------
	' Function SetMembershipById( sMembershipId )
 '-----------------------------------------------------------------------------
	Public Sub SetMembershipById( sMembershipId )
		iMembershipId = sMembershipId
	End Sub 


 '-----------------------------------------------------------------------------
	' Function GetMembershipId( )
 '-----------------------------------------------------------------------------
	Public Function GetMembershipId( sMembership )
		Dim sSql, oMember

		sSql = "Select membershipid FROM egov_memberships WHERE orgid = " & iOrgId & " and membership = '" & sMembership & "' "
		
		Set oMember = Server.CreateObject("ADODB.Recordset")
		oMember.Open sSql, Application("DSN"), 3, 1
		
		GetMembershipId = oMember("membershipid")
		
		oMember.close
		Set oMember = Nothing
	End Function 


 '-----------------------------------------------------------------------------
	' Public Function GetMembershipName( )
 '-----------------------------------------------------------------------------
	Public Function GetMembershipName( )
		Dim sSql, oMember

		sSql = "SELECT membershipdesc FROM egov_memberships WHERE membershipid = " & iMembershipId
		
		Set oMember = Server.CreateObject("ADODB.Recordset")
		oMember.Open sSql, Application("DSN"), 3, 1
		
		If Not oMember.EOF Then 
			GetMembershipName = oMember("membershipdesc")
		Else 
			GetMembershipName = ""
		End If 
		
		oMember.Close
		Set oMember = Nothing
	End Function 

'------------------------------------------------------------------------------
 public function getMembershipIdByMembership(iMembershipType)

   lcl_return = 0

   if iMembershipType = "" then
      iMembershipType = "pool"
   end if

   sSql = "SELECT membershipid "
   sSql = sSql & " FROM egov_memberships "
   sSql = sSql & " WHERE orgid = " & iorgid
   sSql = sSql & " AND UPPER(membership) = '" & UCASE(iMembershipType) & "' "

 		set oMemberID = Server.CreateObject("ADODB.Recordset")
 		oMemberID.Open sSql, Application("DSN"), 3, 1

   if not oMemberID.eof then
      lcl_return = oMemberID("membershipid")
   end if

   oMemberID.close
   set oMemberID = nothing

   getMembershipIdByMembership = lcl_return

 end function

 '-----------------------------------------------------------------------------
	' Public Function GetRateName( iRateId )
 '-----------------------------------------------------------------------------
	Public Function GetRateName( iRateId )
		Dim sSql, oRate

		sSql = "SELECT description FROM egov_poolpassrates WHERE rateid = " & iRateId

		Set oRate = Server.CreateObject("ADODB.Recordset")
		oRate.Open sSql, Application("DSN"), 3, 1

		If Not oRate.EOF Then 
			GetRateName = oRate("description")
		Else
			GetRateName = ""
		End If
			
		oRate.Close
		Set oRate = Nothing
	End Function 


 '-----------------------------------------------------------------------------
	' Public Function RateHasPublicDisplay( sResidentType )
 '-----------------------------------------------------------------------------
	Public Function RateHasPublicDisplay( sResidentType )
		Dim sSql

		sSql = "SELECT public_display "
  sSql = sSql & " FROM egov_membership_rate_displays "
  sSql = sSql & " WHERE membershipid = " & iMembershipId
  sSql = sSql & " AND resident_type = '" & sResidentType & "'"

		Set oDisplay = Server.CreateObject("ADODB.Recordset")
		oDisplay.Open sSql, Application("DSN"), 3, 1

		If Not oDisplay.EOF Then 
			If oDisplay("public_display") Then
				RateHasPublicDisplay = True
			Else
				RateHasPublicDisplay = False 
			End If 
		Else
			RateHasPublicDisplay = False 
		End If 

		oDisplay.close
		Set oDisplay = Nothing
	End Function 


 '-----------------------------------------------------------------------------
	' Sub SaveMembershipIntro(sIntroText)
 '-----------------------------------------------------------------------------
	Public Sub SaveMembershipIntro( sIntroText )
		Dim sSql, oCmd
		
		sIntroText = DBsafe(sIntroText)

		sSql = "Update egov_memberships set introtext = '" & sIntroText & "' where membershipid = " & iMembershipId
		Set oCmd = Server.CreateObject("ADODB.Command")
		With oCmd
			.ActiveConnection = Application("DSN")
			.CommandText = sSql
			.Execute
		End With
		Set oCmd = Nothing

	End Sub 


 '-----------------------------------------------------------------------------
	' Sub SetPublicDisplay( iPublicPurchase )
 '-----------------------------------------------------------------------------
	Public Sub SetPublicDisplay( iPublicPurchase )
		Dim sSql, oCmd

		sSql = "Update egov_memberships set publicpurchase = " & iPublicPurchase & " Where membershipid = " & iMembershipId 

		Set oCmd = Server.CreateObject("ADODB.Command")
		With oCmd
			.ActiveConnection = Application("DSN")
			.CommandText = sSql
			.Execute
		End With
		Set oCmd = Nothing
	End sub


 '-----------------------------------------------------------------------------
	' Public Function ShowMembershipIntro( )
 '-----------------------------------------------------------------------------
	Public Function ShowMembershipIntro( )
		Dim sSql, oIntro

		sSql = "Select introtext FROM egov_memberships WHERE membershipid = " & iMembershipId

		Set oIntro = Server.CreateObject("ADODB.Recordset")
		oIntro.Open sSql, Application("DSN"), 3, 1

		If Not oIntro.eof Then 
			ShowMembershipIntro = Trim(oIntro("introtext"))
		End If
			
		oIntro.close
		Set oIntro = Nothing

	End Function  


 '-----------------------------------------------------------------------------
	' Function ShowMembershipPicks(iMembershipId, iOrgId)
 '-----------------------------------------------------------------------------
	Public Function ShowMembershipPicks()
		Dim sSql, oMembers

		' Get the memberships
		sSql = "Select membershipid, membershipdesc FROM egov_memberships WHERE orgid = " & iOrgId & " order by membershipdesc"
		ShowMembershipPicks = ""

		Set oMembers = Server.CreateObject("ADODB.Recordset")
		oMembers.Open sSql, Application("DSN"), 3, 1
		
		Do While not oMembers.eof 
			ShowMembershipPicks = ShowMembershipPicks  & "<option value=""" & oMembers("membershipid") & """ "
			If clng(iMembershipId) = clng(oMembers("membershipid"))  Then
				ShowMembershipPicks = ShowMembershipPicks & " selected=""selected"" "
			End If 
			ShowMembershipPicks = ShowMembershipPicks & ">" & oMembers("membershipdesc") & "</option>"
			oMembers.movenext
		Loop 

		oMembers.close
		Set oMembers = Nothing

	End Function 


 '-----------------------------------------------------------------------------
	' Function ShowPublicDisplayCheck( )
 '-----------------------------------------------------------------------------
	Public Function ShowPublicDisplayCheck( )
		Dim sSql, oDisplay

		ShowPublicDisplayCheck = ""

		sSql = "Select publicpurchase FROM egov_memberships Where membershipid = " & iMembershipId 
		Set oDisplay = Server.CreateObject("ADODB.Recordset")
		oDisplay.Open sSql, Application("DSN"), 3, 1

		If Not oDisplay.EOF Then 
			If oDisplay("publicpurchase") Then
				ShowPublicDisplayCheck = "checked=""checked"" "
			End If 
		End If 

		oDisplay.close
		Set oDisplay = Nothing

	End Function 


 '-----------------------------------------------------------------------------
	' Public Sub ShowMembershipPeriodPicks( iPeriodId )
 '-----------------------------------------------------------------------------
	Public Sub ShowMembershipPeriodPicks( iPeriodId )
		Dim sSql, oPeriods

		' Get the Periods
		sSql = "Select periodid, period_desc FROM egov_membership_periods WHERE orgid = " & iOrgId & " order by period_desc DESC"

		Set oPeriods = Server.CreateObject("ADODB.Recordset")
		oPeriods.Open sSql, Application("DSN"), 3, 1
		
		Do While not oPeriods.eof 
			response.write vbcrlf & "<option value=""" & oPeriods("periodid") & """ "
			If clng(iPeriodId) = clng(oPeriods("periodid"))  Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oPeriods("period_desc") & "</option>"
			oPeriods.movenext
		Loop 

		oPeriods.close
		Set oPeriods = Nothing

	End Sub  

 '-----------------------------------------------------------------------------
	Public Sub ShowPeriodPicksForMembership( iMembershipId, iPeriodId )
		Dim sSql, oPeriods

	'Get the Periods
		sSql = "SELECT DISTINCT P.periodid, "
  sSql = sSql & " P.period_desc "
  sSql = sSql & " FROM egov_membership_periods P, "
  sSql = sSql &      " egov_poolpassrates R "
  sSql = sSql & " WHERE R.periodid = P.periodid "
  sSql = sSql & " AND R.isEnabled = 1 "
		sSql = sSql & " AND P.orgid = "        & iOrgId
  sSql = sSql & " AND R.membershipid = " & iMembershipId
  sSql = sSql & " ORDER BY P.period_desc DESC"

		set oPeriods = Server.CreateObject("ADODB.Recordset")
		oPeriods.Open sSql, Application("DSN"), 3, 1
		
		if not oPeriods.eof then
  			response.write "<select name=""periodid"" id=""periodid"" onchange=""submitPoolPassForm();"">" 

   		do while not oPeriods.eof
        if CLng(iPeriodId) = CLng(oPeriods("periodid")) then
      					lcl_selected_period = " selected=""selected"""
        else
           lcl_selected_period = ""
        end if

    				response.write "  <option value=""" & oPeriods("periodid") & """" & lcl_selected_period & ">" & oPeriods("period_desc") & "</option>" 

    				oPeriods.movenext
     loop

   		response.write "</select>" 
  end if

		oPeriods.close
		set oPeriods = nothing

 end sub

'------------------------------------------------------------------------------
	public sub ShowPeriodPicksForMembership_AltLayout ( iMembershipId, iPeriodId, sUserType, iRateDesc )
		Dim sSql, oPeriods

	'Get the Periods
		sSql = "SELECT DISTINCT P.periodid, P.period_desc, R.amount "
  sSql = sSql & " FROM egov_membership_periods P, "
  sSql = sSql &      " egov_poolpassrates R "
		sSql = sSql & " WHERE R.periodid = P.periodid "
  sSql = sSql & " AND R.isEnabled = 1 "
  sSql = sSql & " AND P.orgid = "               & iOrgId
  sSql = sSql & " AND R.membershipid = "        & iMembershipId
  sSql = sSql & " AND R.residenttype = '"       & sUserType        & "' "
  sSql = sSql & " AND UPPER(R.description) = '" & UCASE(iRateDesc) & "' "
  sSql = sSql & " ORDER BY P.period_desc DESC"

		set oPeriods = Server.CreateObject("ADODB.Recordset")
		oPeriods.Open sSql, Application("DSN"), 3, 1

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

        lcl_rateid = getRateID(iMembershipId, sUserType, iRateDesc, oPeriods("periodid"))

     			response.write "  <tr>" 
        response.write "      <td><input type=""radio"" name=""periodid"" id=""periodid_" & oPeriods("periodid") & """ value=""" & oPeriods("periodid") & """" & sDisabled & lcl_checked_period & " onclick=""updateRateID('" & lcl_rateid & "');enableDisableContinueButton('" & oPeriods("periodid") & "');submitPoolPassForm();"" />" & oPeriods("period_desc") & "</td>" 
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
	function countPeriodPicksForMembership ( iMembershipId, iPeriodId, iRateDesc )
		Dim sSql, oPeriods
  lcl_return = 0

	'Count the Periods
		sSql = "SELECT count(DISTINCT P.periodid) as total_periods "
  sSql = sSql & " FROM egov_membership_periods P, "
  sSql = sSql &      " egov_poolpassrates R "
		sSql = sSql & " WHERE R.periodid = P.periodid "
  sSql = sSql & " AND R.isEnabled = 1 "
  sSql = sSql & " AND P.orgid = "               & iOrgId
  sSql = sSql & " AND R.membershipid = "        & iMembershipId
  sSql = sSql & " AND UPPER(R.description) = '" & UCASE(iRateDesc) & "' "
  sSql = sSql & " ORDER BY P.period_desc DESC"

		set oPeriodsCnt = Server.CreateObject("ADODB.Recordset")
		oPeriodsCnt.Open sSql, Application("DSN"), 3, 1
		
		if not oPeriodsCnt.eof then
     lcl_return = oPeriodsCnt("total_periods")
  end if

		oPeriodsCnt.Close
		set oPeriodsCnt = nothing

  countPeriodPicksForMembership = lcl_return

 end function

'------------------------------------------------------------------------------
function getRateID(iMembershipId, sUserType, iRateDesc, iPeriodID)
  lcl_return = 0

	'Get the RateID
		sSql = "SELECT DISTINCT R.rateid "
  sSql = sSql & " FROM egov_membership_periods P, "
  sSql = sSql &      " egov_poolpassrates R "
		sSql = sSql & " WHERE R.periodid = P.periodid "
  sSql = sSql & " AND P.orgid = "               & iorgid
  sSql = sSql & " AND R.membershipid = "        & iMembershipId
  sSql = sSql & " AND R.residenttype = '"       & sUserType        & "' "
  sSql = sSql & " AND UPPER(R.description) = '" & UCASE(iRateDesc) & "' "
  sSql = sSql & " AND P.periodid = "            & iPeriodID

		set oRateID = Server.CreateObject("ADODB.Recordset")
		oRateID.Open sSql, Application("DSN"), 3, 1

  if not oRateID.eof then
     lcl_return = oRateID("rateid")
  end if

  oRateID.close
  set oRateID = nothing

  getRateID = lcl_return

end function

'------------------------------------------------------------------------------
	Public Function GetFirstMembershipPeriodId( iMembershipId )
		Dim sSql, oPeriods

  lcl_return = 0

		sSql = "SELECT DISTINCT P.periodid, "
  sSql = sSql & " P.period_desc "
  sSql = sSql & " FROM egov_membership_periods P, "
  sSql = sSql &      " egov_poolpassrates R "
		sSql = sSql & " WHERE R.periodid = P.periodid "
  sSql = sSql & " AND R.isEnabled = 1 "
  sSql = sSql & " AND P.orgid = "        & iOrgId
  sSql = sSql & " AND R.membershipid = " & iMembershipId

  if checkDefaultPeriodExists() then
     sSql = sSql & " AND P.isDefault = 1 "
  end if

  sSql = sSql & " ORDER BY P.period_desc DESC"
		'response.write sSql

		set oPeriods = Server.CreateObject("ADODB.Recordset")
		oPeriods.Open sSql, Application("DSN"), 3, 1
		
		if not oPeriods.eof then
			  lcl_return = CLng(oPeriods("periodid"))
		end if

		oPeriods.close
		set oPeriods = nothing

  GetFirstMembershipPeriodId = lcl_return

	End Function

'------------------------------------------------------------------------------
 public function checkDefaultPeriodExists()
   lcl_return = False

   sSql = "SELECT periodid "
   sSql = sSql & " FROM egov_membership_periods "
   sSql = sSql & " WHERE orgid = " & iorgid
   sSql = sSql & " AND isdefault = 1 "

 		set oDefault = Server.CreateObject("ADODB.Recordset")
	 	oDefault.Open sSql, Application("DSN"), 3, 1

   if not oDefault.eof then
      lcl_return = True
   end if

   oDefault.close
   set oDefault = nothing

   checkDefaultPeriodExists = lcl_return

 end function

'------------------------------------------------------------------------------
	Public Function GetMembershipPeriodName( iPeriodId )
		Dim sSql, oPeriods

		If iPeriodId <> "" Then 
			sSql = "SELECT period_desc FROM egov_membership_periods WHERE periodid = " & iPeriodId

			Set oPeriods = Server.CreateObject("ADODB.Recordset")
			oPeriods.Open sSql, Application("DSN"), 3, 1
			
			If Not oPeriods.EOF Then 
				GetMembershipPeriodName = oPeriods("period_desc")
			Else 
				GetMembershipPeriodName = ""
			End If 

			oPeriods.Close
			Set oPeriods = Nothing
		Else
			GetMembershipPeriodName = ""
		End If 
	End Function 


 '-----------------------------------------------------------------------------
	' Public Sub ShowMembershipRatePublicDisplay( sResidentType )
 '-----------------------------------------------------------------------------
'	Public Sub ShowMembershipRatePublicDisplay( sResidentType )
	function ShowMembershipRatePublicDisplay( ByVal sResidentType )
		Dim sSql
  lcl_return = ""

		sSql = "SELECT public_display "
  sSql = sSql & " FROM egov_membership_rate_displays "
  sSql = sSql & " WHERE membershipid = " & iMembershipId
  sSql = sSql & " AND resident_type = '" & sResidentType & "'"

		Set oDisplay = Server.CreateObject("ADODB.Recordset")
		oDisplay.Open sSql, Application("DSN"), 3, 1

		If Not oDisplay.EOF Then 
  			If oDisplay("public_display") Then
		  		  'response.write "checked=""checked"" "
        lcl_return = " checked=""checked"""
  			End If 
		End If 

		oDisplay.close
		Set oDisplay = Nothing
ShowMembershipRatePublicDisplay = lcl_return

end function


	'------------------------------------------------------------------------------
	' Public Sub ShowMembershipResidentRates( sUserType, iMembershipId, iPeriodid )
	'------------------------------------------------------------------------------
	public sub ShowMembershipResidentRates( sUserType, iMembershipId, iPeriodid )
		Dim sSql, oTypes

		sSql = "SELECT DISTINCT T.resident_type, "
  sSql = sSql & " T.description, "
  sSql = sSql & " T.displayorder "
  sSql = sSql & " FROM egov_poolpassresidenttypes T, "
  sSql = sSQl &      " egov_poolpassrates R "
		sSql = sSql & " WHERE T.resident_type = R.residenttype "
  sSql = sSql & " AND R.isEnabled = 1 "
  sSql = sSql & " AND T.orgid = "         & session("OrgID")
		sSql = sSql & " AND R.membershipid = "  & iMembershipId
  sSql = sSql & " AND R.periodid = "      & iPeriodid
  sSql = sSql & " AND R.residenttype = '" & sUserType & "' "
		sSql = sSql & " ORDER BY T.displayorder "

		set oTypes = Server.CreateObject("ADODB.Recordset")
		oTypes.Open sSql, Application("DSN"), 3, 1
		
		do while not oTypes.eof
  			response.write "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" class=""tableright"">" 
		  	response.write "  <tr><th colspan=""2""> &nbsp; " & oTypes("description") & "</th></tr>" 

  		'Get the rates here 
		  	ShowMembershipRates oTypes("resident_type"), sUserType, iMembershipId, iPeriodid

  			response.write "</table>" & vbrlf

  			oTypes.MoveNext
		loop

		oTypes.Close
		set oTypes = nothing

 end sub

	'------------------------------------------------------------------------------
	' Private Sub ShowMembershipRates( sResidenttype, sUserType )
	'------------------------------------------------------------------------------
	private sub ShowMembershipRates( sResidenttype, sUserType, iMembershipId, iPeriodid )
		Dim sDisabled, sSql, oRates, bPreSelect, iRow, iCurrYear
		sDisabled  = ""
		bPreSelect = False
		iRow       = 0 
		iCurrYear  = Year(Now())

		if sUserType <> sResidentType then
  			sDisabled = " disabled"
		end if

		if sDisabled = "" and iRow = 1 then
 				lcl_checked_rates = " checked=""checked"""
  else
     lcl_checked_rates = ""
  end if

		sSql = "SELECT R.rateid, "
  sSql = sSql & " R.description, "
  sSql = sSql & " R.amount , "
  sSql = sSql & " MP.period_desc "
  sSql = sSql & " FROM egov_poolpassrates R, "
  sSql = sSql &      " egov_membership_periods MP "
		sSql = sSql & " WHERE R.periodid = MP.periodid "
		sSql = sSql & " AND R.orgid = MP.orgid "
  sSql = sSql & " AND R.isEnabled = 1 "
		sSql = sSql & " AND R.orgid = "         & session("OrgID")
		sSql = sSql & " AND R.residenttype = '" & sResidenttype & "' "
		sSql = sSql & " AND R.membershipid = "  & iMembershipId
		sSql = sSql & " AND MP.periodid = "     & iPeriodid
  sSql = sSql & " AND R.residenttype = '" & sUserType     & "' "
		sSql = sSql & " ORDER BY R.displayorder"

		set oRates = Server.CreateObject("ADODB.Recordset")
		oRates.Open sSql, Application("DSN"), 3, 1
		
		do while not oRates.eof
  			iRow = iRow + 1

  			response.write "  <tr>" 
     response.write "      <td><input type=""radio"" name=""rateid"" id=""rateid_" & oRates("rateid") & """ value=""" & oRates("rateid") & """" & sDisabled & lcl_checked_rates & " onclick=""enableDisableContinueButton('" & oRates("rateid") & "');"" />" & oRates("description") & "</td>" 
     response.write "      <td class=""amount"">" & FormatCurrency(oRates("amount")) & "</td>" 
     response.write "  </tr>" 

  			oRates.MoveNext
  loop

		oRates.Close
		set oRates = nothing
		
 end sub

	'------------------------------------------------------------------------------
	public sub ShowMembershipRates_AltLayout( sUserType, iMembershipId, iRateDesc)
		Dim sDisabled, sSql, oRates, bPreSelect, iRow, iCurrYear
		iRow = 0 

  lcl_rate_desc = getDistinctRateDescList(iMembershipID, sUserType)

  'sSql = "SELECT distinct R.rateid, R.description, R.amount , MP.period_desc "
		sSql = "SELECT distinct R.description "
  sSql = sSql & " FROM egov_poolpassrates R, "
  sSql = sSql &      " egov_membership_periods MP "
		sSql = sSql & " WHERE R.periodid = MP.periodid "
		sSql = sSql & " AND R.orgid = MP.orgid "
  sSql = sSql & " AND R.isEnabled = 1 "
		sSql = sSql & " AND R.orgid = "         & session("OrgID")
		sSql = sSql & " AND R.membershipid = "  & iMembershipId
  sSql = sSql & " AND R.residenttype = '" & sUserType     & "' "
  sSql = sSql & " AND R.description IN (" & lcl_rate_desc & ") "
  'sSql = sSql & " ORDER BY R.displayorder, R.description "
  sSql = sSql & " ORDER BY R.description "

		set oRates = Server.CreateObject("ADODB.Recordset")
		oRates.Open sSql, Application("DSN"), 3, 1

  if not oRates.eof then
     response.write "<select name=""rateDesc"" id=""rateDesc"" onchange=""submitPoolPassForm();"">" 

     do while not oRates.eof
     			iRow = iRow + 1

     		if UCASE(oRates("description")) = UCASE(iRateDesc) then
      				lcl_selected_rates = " selected=""selected"""
       else
          lcl_selected_rates = ""
       end if

        response.write "  <option value=""" & oRates("description") & """" & lcl_selected_rates & ">" & oRates("description") & "</option>" 

     			oRates.MoveNext
     loop

     response.write "</select>" 

  end if

		oRates.Close
		set oRates = nothing
		
 end sub

'------------------------------------------------------------------------------
	function getFirstMembershipRateOption_AltLayout( sUserType, iMembershipId)
  lcl_return = ""

  lcl_rate_desc = getDistinctRateDescList(iMembershipID, sUserType)

 'Now grab the first rate.
		sSql = "SELECT R.description, "
  sSql = sSql & " R.amount, "
  sSql = sSql & " R.displayorder "
  sSql = sSql & " FROM egov_poolpassrates R, "
  sSql = sSql &      " egov_membership_periods MP "
		sSql = sSql & " WHERE R.periodid = MP.periodid "
		sSql = sSql & " AND R.orgid = MP.orgid "
  sSql = sSql & " AND R.isEnabled = 1 "
		sSql = sSql & " AND R.orgid = "         & session("OrgID")
		sSql = sSql & " AND R.membershipid = "  & iMembershipId
  sSql = sSql & " AND R.residenttype = '" & sUserType     & "' "
  sSql = sSql & " AND R.description IN (" & lcl_rate_desc & ") "
  sSql = sSql & " ORDER BY R.displayorder, R.description "

		set oFirstRate = Server.CreateObject("ADODB.Recordset")
		oFirstRate.Open sSql, Application("DSN"), 3, 1

  if not oFirstRate.eof then
     lcl_return = oFirstRate("description")
  end if

		oFirstRate.Close
		set oFirstRate = nothing

  getFirstMembershipRateOption_AltLayout = lcl_return
		
 end function

'------------------------------------------------------------------------------
function getDistinctRateDescList(iMembershipID, sUserType)
  lcl_return = "0"

 'Get a distinct list of rate descriptions
		sSql = "SELECT distinct R.description "
  sSql = sSql & " FROM egov_poolpassrates R, "
  sSql = sSql &      " egov_membership_periods MP "
		sSql = sSql & " WHERE R.periodid = MP.periodid "
		sSql = sSql & " AND R.orgid = MP.orgid "
  sSql = sSql & " AND R.isEnabled = 1 "
		sSql = sSql & " AND R.orgid = "         & session("OrgID")
		sSql = sSql & " AND R.membershipid = "  & iMembershipId
  sSql = sSql & " AND R.residenttype = '" & sUserType     & "' "

		set oDistinctRates = Server.CreateObject("ADODB.Recordset")
		oDistinctRates.Open sSql, Application("DSN"), 3, 1

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

	'------------------------------------------------------------------------------
	' Private Sub ShowMembershipRates_old( sResidenttype, sUserType )
	'------------------------------------------------------------------------------
	Private Sub ShowMembershipRates_old( sResidenttype, sUserType )
		Dim sDisabled, sSql, oRates, bPreSelect, iRow, iCurrYear
		sDisabled = ""
		bPreSelect = False
		iRow = 0 
		iCurrYear = Year(Now())

		If sUserType <> sResidentType Then
			sDisabled = " disabled "
		End If 

		' Get the Rates for the Orgid and residenttype
		Set oCmd = Server.CreateObject("ADODB.Command")
		With oCmd
			.ActiveConnection = Application("DSN")
			.CommandText = "GetMembershipRatesList"
			.CommandType = 4
			.Parameters.Append oCmd.CreateParameter("@iOrgid", 3, 1, 4, Session("OrgID"))
			.Parameters.Append oCmd.CreateParameter("@sResidentType", 129, 1, 1, sResidentType)
			.Parameters.Append oCmd.CreateParameter("@iMembershipId", 3, 1, 4, iMembershipId)
			Set oRates = .Execute
		End With
		
		Do While Not oRates.EOF
			iRow = iRow + 1
			' Display 
			response.write vbcrlf & "<tr><td><input type=""radio"" name=""rateid"" " 
			response.write " value=""" & oRates("rateid") & """ " & sDisabled 
			If sDisabled = "" And iRow = 1 Then	
				response.write " checked=""checked"" "
			End If 
			response.write " />" & oRates("description") 
			response.write "</td><td class=""amount"">" & FormatCurrency(oRates("amount")) & "</td></tr>"
			oRates.movenext
		Loop 
		oRates.close
		Set oRates = Nothing
		
	End Sub 

'------------------------------------------------------------------------------
public sub MembershipPurchase(ByVal iUserId, ByVal iRateId, ByVal nAmount, ByVal sPaymentType, ByVal sPaymentLocation, _
                              ByVal iMembershipId, ByVal iPeriodId, ByVal sPoolPassID, ByVal lcl_startdate, ByRef iPassID)
 	Dim sResult

		if LCASE(sPaymentType) = "creditcard" then
  			sResult = "Pending"
		else
			  sResult = "Paid"
		end if

  if lcl_startdate = "" then
     lcl_startdate = Date()
  end if

  lcl_expirationdate = getExpirationDate(iPeriodID,lcl_startdate)

  sSqli = "INSERT INTO egov_poolpasspurchases ( "
  sSqli = sSqli & "userid, "
  sSqli = sSqli & "orgid, "
  sSqli = sSqli & "rateid, "
  sSqli = sSqli & "membershipid, "
  sSqli = sSqli & "periodid, "
  sSqli = sSqli & "paymentamount, "
  sSqli = sSqli & "paymenttype, "
  sSqli = sSqli & "paymentlocation, "
  sSqli = sSqli & "paymentresult, "
  sSqli = sSqli & "startdate, "
  sSqli = sSqli & "expirationdate "

  if sPoolPassID <> "" then
     sSqli = sSqli & ", previous_poolpassid "
  end if

  sSqli = sSqli & ")	VALUES ("
  sSqli = sSqli & iUserID                  & ", "
  sSqli = sSqli & session("orgid")         & ", "
  sSqli = sSqli & iRateID                  & ", "
  sSqli = sSqli & iMembershipID            & ", "
  sSqli = sSqli & iPeriodID                & ", "
  sSqli = sSqli & nAmount                  & ", "
  sSqli = sSqli & "'" & sPaymentType       & "', "
  sSqli = sSqli & "'" & sPaymentLocation   & "', "
  sSqli = sSqli & "'" & sResult            & "', "
  sSqli = sSqli & "'" & lcl_startdate      & "', "
  sSqli = sSqli & "'" & lcl_expirationdate & "' "

  if sPoolPassID <> "" then
     sSqli = sSqli & ", " & sPoolPassID
  end if

  sSqli = sSqli & ")"

  set rsi = Server.CreateObject("ADODB.Recordset")
  rsi.Open sSqli, Application("DSN"), 3, 1

 'Retrieve the poolpassid that was just inserted
  sSqlid = "SELECT IDENT_CURRENT('egov_poolpasspurchases') as NewID"
  rsi.Open sSqlid, Application("DSN"), 3, 1
  iPassID = rsi("NewID").value

  set rsi = nothing

end sub

'------------------------------------------------------------------------------
	public sub AddMember(iMembershipId, iFamilymemberId, iPrevPoolPassID, iIsPunchcard, iPunchcardLimit)
   lcl_member_id     = ""
   lcl_card_printed  = "N"
   lcl_printed_count = 0

   if iPrevPoolPassID <> "" then
     'Get the current member data
      getCurrentMemberInfo iPrevPoolPassID,iFamilyMemberID,lcl_member_id,lcl_card_printed,lcl_printed_count
   else
      lcl_member_id = getNextMemberID()
   end if

   if iIsPunchcard then
      lcl_isPunchcard = 1
   else
      lcl_isPunchcard = 0
   end if

	 'Insert new member records
   sSqli = "INSERT INTO egov_poolpassmembers ("
   sSqli = sSqli & "poolpassid, "
   sSqli = sSqli & "familymemberid, "
   sSqli = sSqli & "memberid, "
   sSqli = sSqli & "card_printed, "
   sSqli = sSqli & "printed_count, "
   sSqli = sSqli & "isPunchcard, "
   sSqli = sSqli & "punchcard_limit, "
   sSqli = sSqli & "pcard_remaining_cnt "
   sSqli = sSqli & ") VALUES ("
   sSqli = sSqli &       iMembershipID     & ", "
   sSqli = sSqli &       iFamilyMemberID   & ", "
   sSqli = sSqli &       lcl_member_id     & ", "
   sSqli = sSqli & "'" & lcl_card_printed  & "', "
   sSqli = sSqli &       lcl_printed_count & ", "
   sSqli = sSqli &       lcl_isPunchcard   & ", "
   sSqli = sSqli &       iPunchcardLimit   & ", "
   sSqli = sSqli &       iPunchcardLimit
   sSqli = sSqli & ")"

   set rsi = Server.CreateObject("ADODB.Recordset")
   rsi.Open sSqli, Application("DSN"), 3, 1

   set rsi = nothing

	end sub


	'------------------------------------------------------------------------------
	Public Sub ShowMembershipInfo( ByVal iPoolpassId )
		Dim sSql, oRs

		sSql = "SELECT U.userfname, "
  sSql = sSql & " U.userlname, "
  sSql = sSql & " U.useraddress, "
  sSql = sSql & " U.useraddress2, "
  sSql = sSql & " U.usercity, "
  sSql = sSql & " U.userstate, "
  sSql = sSql & " U.userzip, "
		sSql = sSql & " P.paymentamount, "
  sSql = sSql & " P.paymenttype, "
  sSql = sSql & " P.paymentdate, "
  sSql = sSql & " P.paymentlocation, "
  sSql = sSql & " R.description, "
		sSql = sSql & " T.description as residenttype, "
  sSql = sSql & " M.membershipdesc, "
  sSql = sSql & " MP.period_desc, "
  sSql = sSql & " MP.period_interval, "
		sSql = sSql & " MP.period_qty, "
  sSql = sSql & " MP.period_type, "
  sSql = sSql & " P.previous_poolpassid, "
  sSql = sSql & " P.startdate, "
  sSql = sSql & " P.expirationdate, "
		sSql = sSql & " R.isPunchcard, "
  sSql = sSql & " R.punchcard_limit, "
  sSql = sSql & " ISNULL(P.processingfee,0.00) AS processingfee, "
  sSql = sSql & " ISNULL(P.sva,'') AS sva,P.note,P.adminid,au.firstname,au.lastname "

		sSql = sSql & " FROM egov_poolpasspurchases P "
		sSql = sSql & " INNER JOIN egov_users U ON U.userid = p.userid "
		sSql = sSql & " INNER JOIN egov_poolpassrates R ON P.rateid = r.rateid "
		sSql = sSql & " INNER JOIN egov_poolpassresidenttypes T ON r.residenttype = t.resident_type AND T.orgid = P.orgid "
		sSql = sSql & " INNER JOIN egov_memberships M ON M.membershipid = P.membershipid "
		sSql = sSql & " INNER JOIN egov_membership_periods MP ON mp.periodid = p.periodid "
		sSql = sSql & " LEFT JOIN users au ON au.userid = P.adminid"
		sSql = sSql & " WHERE P.orgid = " & session("orgid")
		sSql = sSql & " AND P.poolpassid = " & iPoolpassId  
  sSql = sSql & " AND R.isEnabled = 1 "

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then 
			response.write "<tr>" 
			response.write "<td width=""250px"" align=""right"">Pass ID: </td>" 
			response.write "<td>" & iPoolpassId & "</td>" 
			response.write "</tr>" 
			response.write "<tr>" 
			response.write "<td align=""right"">Membership Type: </td>" 
			response.write "<td>" & oRs("membershipdesc") & "</td>" 
			response.write "</tr>" 
			response.write "<tr>" 
			response.write "<td align=""right"">Pass Type: </td>" 
			response.write "<td>" & oRs("residenttype") & " &mdash; " & oRs("description") & "</td>" 
			response.write "</tr>" 

			If oRs("isPunchcard") Then 
				response.write "<tr>" 
				response.write "<td align=""right"">Punchcard: </td>" 
				response.write "<td>Yes</td>" 
				response.write "</tr>" 
				response.write "<tr>" 
				response.write "<td align=""right"">Punchcard Limit: </td>" 
				response.write "<td>" & oRs("punchcard_limit") & "</td>" 
				response.write "</tr>" 
			End If 

			response.write "<tr>" 
			response.write "<td align=""right"">Membership Period: </td>" 
			response.write "<td>" & oRs("period_desc") & "</td>" 
			response.write "</tr>" 
			response.write "<tr>" 
			response.write "<td align=""right"">Purchase Date: </td>" 
			response.write "<td>" & DateValue(oRs("paymentdate")) & "</td>" 
			response.write "</tr>" 
			response.write "<tr>" 
			response.write "<td align=""right"">Membership Start Date: </td>" 
			response.write "<td>" & DateValue(oRs("startdate")) & "</td>" 
			response.write "</tr>" 


			response.write "<tr>" 
			response.write "<td align=""right"">Expiration Date: </td>" 
			response.write "<td>" & DateValue(oRs("expirationdate")) & "</td>" 
			response.write "</tr>" 
			response.write "<tr>" 
			response.write "<td align=""right"">Payment Method: </td>" 
			response.write "<td>" & MakeProper(oRs("paymentlocation")) & " &mdash; " & MakeProper(oRs("paymenttype")) & "</td>" 
			response.write "</tr>" 
			response.write "<tr>" 
			response.write "<td align=""right"">Amount: </td>" 
			response.write "<td>" & FormatCurrency(oRs("paymentamount"),2) & "</td>" 
			response.write "</tr>" 
			If oRs("sva") <> "" Then
				response.write "<tr>" 
				response.write "<td align=""right"">Processing Fee: </td>" 
				response.write "<td>" & FormatCurrency(oRs("processingfee"),2) & "</td>" 
				response.write "</tr>" 
				response.write "<tr>" 
				response.write "<td align=""right"">Amount Charged: </td>" 
				response.write "<td>" & FormatCurrency((CDbl(oRs("processingfee")) + CDbl(oRs("paymentamount"))),2) & "</td>" 
				response.write "</tr>" 
			End If 
			response.write "<tr>" 
			response.write "<td valign=""top""align=""right"">Purchaser: </td>" 
			response.write "<td>" & oRs("userfname") & " " & oRs("userlname") & "<br />" 
			response.write oRs("useraddress") & "<br />" 

			If oRs("useraddress2") <> "" Or IsNull(oRs("useraddress2")) = False Then 
				response.write oRs("useraddress2") & "<br />" 
			End If 

			response.write oRs("usercity") & ", " & oRs("userstate") & " " & oRs("userzip") 
			response.write "</td>" 
			response.write "</tr>" 

			If oRs("previous_poolpassid") <> "" Then 
				response.write "<tr>" 
				response.write "<td align=""right"">Renewal of Pass ID: </td>" 
				response.write "<td><a href=""poolpass_receipt.asp?iPoolPassId=" & oRs("previous_poolpassid") & """>" & oRs("previous_poolpassid") & "</a></td>" 
				response.write "</tr>" 
			End If 

			if oRs("note") <> "" then
				response.write "<tr>" 
				response.write "<td align=""right"">Note: </td>" 
				response.write "<td>" & oRs("note") & "</td>" 
				response.write "</tr>" 
			end if

			if oRs("adminid") <> "" then
				response.write "<tr>" 
				response.write "<td align=""right"">Admin Processor: </td>" 
				response.write "<td>" & oRs("FirstName") & " " & oRs("LastName") & "</td>" 
				response.write "</tr>" 
			end if
		End If 

		oRs.Close
		set oRs = Nothing 

	End Sub 

'------------------------------------------------------------------------------
 public sub ShowMembers( iMembershipId )
		Dim sSql, oMembers

		sSql = "SELECT F.firstname, F.lastname, F.relationship, P.memberid "
  sSql = sSql & " FROM egov_familymembers F, egov_poolpassmembers P "
  sSql = sSql & " WHERE P.poolpassid = " & iMembershipId
  sSql = sSql & " AND F.familymemberid = P.familymemberid "
  sSql = sSql & " ORDER BY birthdate, lastname, firstname "

		set oMembers = Server.CreateObject("ADODB.Recordset")
		oMembers.Open sSql, Application("DSN"), 3, 1

		while not oMembers.eof
  	  response.write "  <tr>" 
     response.write "      <td width=""300px"" align=""right"">" & oMembers("firstname") & " " & oMembers("lastname") & " (" & oMembers("memberid") & ")</td>" 
     response.write "      <td>" & oMembers("relationship") &"</td>" 
     response.write "  </tr>" 
		 	 oMembers.movenext
  wend
			
		oMembers.close
		set oMembers = nothing
	end sub

'------------------------------------------------------------------------------
 sub ShowFamilyMembers(iPoolPassID)
   dim sSQL, sSQL2
   dim lcl_userid, lcl_orgid, lcl_bgcolor
   dim lcl_orghasfeature_memberships_usekeycards

   lcl_userid  = 0
   lcl_bgcolor = "#eeeeee"
   lcl_orghasfeature_memberships_usekeycards = OrgHasFeature("memberships_usekeycards")

   sSQL = "SELECT userid, "
   sSQL = sSQL & " orgid "
   sSQL = sSQL & " FROM egov_poolpasspurchases "
   sSQL = sSQL & " WHERE poolpassid = " & iPoolPassID

  	set rs = Server.CreateObject("ADODB.Recordset")
	  rs.Open sSQL, Application("DSN"), 3, 1

   if not rs.eof then
      lcl_userid = rs("userid")
      lcl_orgid  = rs("orgid")
   end if

   sSQL2 = "SELECT f.familymemberid, "
   sSQL2 = sSQL2 & " f.firstname, "
   sSQL2 = sSQL2 & " f.lastname, "
   sSQL2 = sSQL2 & " f.relationship, "
   sSQL2 = sSQL2 & " isnull(m.memberid,0) as memberid "
   sSQL2 = sSQL2 & " FROM egov_familymembers f "
   sSQL2 = sSQL2 &      " LEFT OUTER JOIN egov_poolpassmembers m "
   sSQL2 = sSQL2 &                   " ON f.familymemberid = m.familymemberid "
   sSQL2 = sSQL2 &                   " AND m.poolpassid = " & iPoolPassID
   sSQL2 = sSQL2 & " WHERE f.isdeleted = 0 "
   sSQL2 = sSQL2 & " AND f.belongstouserid = " & lcl_userid

  	set oFamily = Server.CreateObject("ADODB.Recordset")
	  oFamily.Open sSQL2, Application("DSN"), 3, 1

   if not oFamily.eof then
      do while not oFamily.eof
         lcl_checked                  = ""
         lcl_showMemberIDRelationship = ""

       		if IsOnPoolPass(iPoolPassId, oFamily("familymemberid")) then
         			lcl_checked = " checked=""checked"""
        	end if

         if CLng(oFamily("memberid")) > CLng(0) then
            lcl_showMemberIDRelationship = " (" & oFamily("memberid") & ") "
         end if

       		if UCASE(oFamily("relationship")) <> "SPOUSE" AND UCASE(oFamily("relationship")) <> "PARTNER" then
            if lcl_showMemberIDRelationship <> "" then
               lcl_showMemberIDRelationship = lcl_showMemberIDRelationship & " - "
            end if

            lcl_showMemberIDRelationship = lcl_showMemberIDRelationship & oFamily("relationship")
         end if

        'If the feature is enabled, check to see if any barcodes are associated to the member.
        'If "yes" then build the list and display the barcode icon.
        'If "no" then do not display anything.
         'sBarcodeList = ""
         sDisplayActiveBarcode = ""

         if lcl_orghasfeature_memberships_usekeycards then
            'sBarcodeList = displayBarcodeList(iorgid, _
            '                                  oFamily("memberid"))
            sDisplayActiveBarcode = getActiveBarcode(iorgid, _
                                                     oFamily("memberid"))
         end if

         if sDisplayActiveBarcode = "" then
            sDisplayActiveBarcode = "&nbsp;"
         end if

         response.write "  <tr valign=""top"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
         response.write "      <td align=""right"">" & vbcrlf
         response.write "          <input type=""checkbox"" name=""familymemberid"" value=""" & oFamily("familymemberid") &  """" & lcl_checked & " />" & vbcrlf
         response.write "      </td>" & vbcrlf
         response.write "      <td>" & oFamily("firstname") & " " & oFamily("lastname") & lcl_showMemberIDRelationship & "</td>" & vbcrlf
         response.write "      <td>" & sDisplayActiveBarcode & "</td>" & vbcrlf
         response.write "  </tr>" & vbcrlf

         if lcl_bgcolor = "#eeeeee" then
            lcl_bgcolor = "#ffffff"
         else
            lcl_bgcolor = "#eeeeee"
         end if

       		oFamily.movenext
      loop
   end if

   set rs      = nothing
   set oFamily = nothing

 end sub

'------------------------------------------------------------------------------
'ORIGINAL
'------------------------------------------------------------------------------
' Sub ShowFamilyMembers(iPoolPassID)

'   sSql = "SELECT userid FROM egov_poolpasspurchases WHERE poolpassid = " & iPoolPassID
'  	set rs = Server.CreateObject("ADODB.Recordset")
'	  rs.Open sSql, Application("DSN"), 3, 1

'   if not rs.eof then
'      lcl_userid = rs("userid")
'   else
'      lcl_userid = 0
'   end if

'  	sSql2 = "SELECT f.familymemberid, f.firstname, f.lastname, f.relationship, isnull(m.memberid,0) as memberid "
'   sSql2 = sSql2 & " FROM egov_familymembers f "
'   sSql2 = sSql2 &      " LEFT OUTER JOIN egov_poolpassmembers m "
' 		sSql2 = sSql2 &      " ON f.familymemberid = m.familymemberid AND m.poolpassid = " & iPoolPassID
'  	sSql2 = sSql2 & " WHERE f.isdeleted = 0 "
'   sSql2 = sSql2 & " AND f.belongstouserid = " & lcl_userid

'  	set oFamily = Server.CreateObject("ADODB.Recordset")
'	  oFamily.Open sSql2, Application("DSN"), 3, 1

'   if not oFamily.eof then
'      while not oFamily.eof

'       		if IsOnPoolPass(iPoolPassId, oFamily("familymemberid")) then
'         			lcl_checked = " checked=""checked"""
'         else
'            lcl_checked = ""
'        	end if

'         response.write "  <tr>" 
'         response.write "      <td align=""right"">" 
'         response.write "          <input type=""checkbox"" name=""familymemberid"" value=""" & oFamily("familymemberid") &  """" & lcl_checked & " />" 
'         response.write "      </td>" 
'         response.write "      <td>" 
'         response.write            oFamily("firstname") & " " & oFamily("lastname") 

'         if CLng(oFamily("memberid")) > CLng(0) then
'            response.write " (" & oFamily("memberid") & ") "
'         end if

'       		if UCASE(oFamily("relationship")) <> "SPOUSE" AND UCASE(oFamily("relationship")) <> "PARTNER" then
'            response.write " - " & oFamily("relationship")
'         end if

'         response.write "      </td>" 
'         response.write "  </tr>" 
'       		oFamily.movenext
'      wend
'   end if

'   set rs      = nothing
'   set oFamily = nothing

' end sub

'------------------------------------------------------------------------------
function displayBarcodeList(iOrgID, _
                            iMemberID)

  dim lcl_return, sSQL

  lcl_return = ""

  sSQL = "SELECT mtb.barcode, "
  sSQL = sSQL & " bs.statusname, "
  sSQL = sSQL & " bs.isActiveStatus "
  sSQL = sSQL & " FROM egov_poolpassmembers_to_barcodes mtb  "
  sSQL = sSQL &      " INNER JOIN egov_poolpassmembers_barcode_statuses bs ON bs.statusid = mtb.barcode_statusid "
  sSQL = sSQL & " WHERE mtb.orgid = " & iOrgID
  sSQL = sSQL & " AND mtb.memberid = " & iMemberID
  sSQL = sSQL & " ORDER BY bs.isActiveStatus DESC, bs.statusname "

  set oDisplayBarcodes = Server.CreateObject("ADODB.Recordset")
  oDisplayBarcodes.Open sSQL, Application("DSN"), 3, 1

  if not oDisplayBarcodes.eof then
     lcl_return = "<table border=""0"">"
     lcl_return = lcl_return & "  <tr>"
     lcl_return = lcl_return & "      <th>Barcode</th>"
     lcl_return = lcl_return & "      <th>Status</th>"
     lcl_return = lcl_return & "  </tr>"
      
     do while not oDisplayBarcodes.eof
        sDisplayBarcode    = oDisplayBarcodes("barcode")
        sDisplayStatusName = oDisplayBarcodes("statusname")

        'if oDisplayBarcodes("isActiveStatus") then
        '   sDisplayBarcode    = "<span class=\'isActiveBarcode\'>" & sDisplayBarcode    & "</span>"
        '   sDisplayStatusName = "<span class=\'isActiveBarcode\'>" & sDisplayStatusName & "</span>"
        'end if

        lcl_return = lcl_return & "  <tr>"
        lcl_return = lcl_return & "      <td>" & sDisplayBarcode     & "</td>"
        lcl_return = lcl_return & "      <td>" & sDisplayStatusName  & "</td>"
        lcl_return = lcl_return & "  </tr>"

        oDisplayBarcodes.movenext
     loop

     lcl_return = lcl_return & "</table>"
  end if

  oDisplayBarcodes.close
  set oDisplayBarcodes = nothing

  displayBarcodeList = lcl_return

end function

'------------------------------------------------------------------------------
function getActiveBarcode(iOrgID, _
                          iMemberID)

  dim lcl_return, sSQL

  lcl_return = ""

  sSQL = "SELECT TOP 1 "
  sSQL = sSQL & " mtb.barcode, "
  sSQL = sSQL & " bs.statusname "
  sSQL = sSQL & " FROM egov_poolpassmembers_to_barcodes mtb  "
  sSQL = sSQL &      " INNER JOIN egov_poolpassmembers_barcode_statuses bs ON bs.statusid = mtb.barcode_statusid "
  sSQL = sSQL & " WHERE mtb.orgid = " & iOrgID
  sSQL = sSQL & " AND mtb.memberid = " & iMemberID
  sSQL = sSQL & " AND bs.isActiveStatus = 1 "

  set oGetActiveBarcode = Server.CreateObject("ADODB.Recordset")
  oGetActiveBarcode.Open sSQL, Application("DSN"), 3, 1

  if not oGetActiveBarcode.eof then
     sDisplayBarcode    = oGetActiveBarcode("barcode")
     sDisplayStatusName = oGetActiveBarcode("statusname")

     lcl_return = "<strong>Barcode: </strong>" & sDisplayBarcode & " (" & sDisplayStatusName  & ")"
  end if

  oGetActiveBarcode.close
  set oGetActiveBarcode = nothing

  getActiveBarcode = lcl_return

end function

'------------------------------------------------------------------------------
 function getExpirationDate(p_periodid, p_startdate)
  lcl_expirationdate = ""

  if p_startdate <> "" then
    'Calculate the expiration date
     sSqle = "SELECT CAST(dbo.fn_getMembershipExpirationdate(MP.is_seasonal,MP.period_interval,MP.period_qty,'" & p_startdate & "') AS datetime) AS expirationdate "
     sSqle = sSqle & " FROM egov_membership_periods MP "
     sSqle = sSqle & " WHERE MP.orgid = " & session("orgid")
     sSqle = sSqle & " AND MP.periodid = '" & p_periodid & "'"

     set rse = Server.CreateObject("ADODB.Recordset")
     rse.Open sSqle, Application("DSN"), 3, 1

     if not rse.eof then
        lcl_expirationdate = rse("expirationdate")
     'else
        'lcl_expirationdate = DATEADD(yy,1,p_startdate)
     end if

     set rse = nothing

   end if

   getExpirationDate = lcl_expirationdate

 end function

'------------------------------------------------------------------------------
 function getMembershipStartDate(p_poolpassid)
   lcl_return = Date()

   if p_poolpassid <> "" then
      sSql = "SELECT expirationdate FROM egov_poolpasspurchases WHERE poolpassid = " & p_poolpassid

      set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open sSql, Application("DSN"), 3, 1

      if not rs.eof then
        'Add one day to the renewal start date since we do not want two memberships active on the same date.
         lcl_return = DATEADD("d",1,rs("expirationdate"))
      end if

      set rs = nothing
   end if

   getMembershipStartDate = lcl_return

 end function

'------------------------------------------------------------------------------
end class
%>
