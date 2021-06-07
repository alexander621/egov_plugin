<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: update_class.asp
' AUTHOR: Steve Loar
' CREATED: 04/19/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This script saves changes to classes and events
'
' MODIFICATION HISTORY
' 1.0	04/19/06	Steve Loar - INITIAL VERSION
' 2.0	03/12/06	Steve Loar - Overhauled for Menlo Park project
' 2.1  12/30/08 David Boyer - Added "DisplayRosterPublic" checkbox for Craig, CO custom registration fields.
' 2.2  11/20/09 David Boyer - Added new "team registration" fields
' 2.7	12/02/2009	Steve Loar - Option to only allow purchases on admin but only display on public
' 2.8	10/10/2011	Steve Loar - Added Gender Restriction
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
	Dim sSql, oCmd, iClassId, iMinAge, iMaxAge, dStartdate, dEnddate, dPublishStartdate, dPublishEnddate
	Dim dEvaluationDate, dAlternateDate, sImgUrl, sExternalurl, Item, x, sMin, sMax, sWaitlistmax, sExternallinktext
	Dim dRegistrationStartdate, dRegistrationEnddate, sImgAltTag, Cat, iMembershipId, iPriceDiscountId, iTimeId
	Dim iMinGrade, iMaxGrade, iSu, iMo, iTu, iWe, iTh, iFr, iSa, iMaxTimeID, iClassSeasonId, iSupervisorid
	Dim iMinAgePrecisionId, iMaxAgePrecisionId, dAgeComparedate, sAllowEarlyRegistration, sEarlyRegistrationDate
	Dim sEarlyRegistrationClassSeasonId, sEarlyRegistrationClassId, sShowTerms, sPublicCanOnlyView, iGenderRestrictionId,sNoRefunds

	iClassId           = request("classid")
	iClassSeasonId     = CLng(request("classseasonid"))
	iSupervisorid      = CLng(request("supervisorid"))
	iMinAgePrecisionId = "NULL"
	iMaxAgePrecisionId = "NULL"

	If Trim(request("minage")) = "" Then 
		iMinAge = "NULL"
	Else 
		iMinAge = CDbl(request("minage"))
		iMinAgePrecisionId = request("minageprecisionid")
	End If 

	If Trim(request("maxage")) = "" Then 
		iMaxAge = "NULL"
	Else 
		iMaxAge = CDbl(request("maxage"))
		iMaxAgePrecisionId = request("maxageprecisionid")
	End If 

	If Trim(request("agecomparedate")) = "" Then 
		dAgeComparedate =  " NULL " 
	Else 
		dAgeComparedate = " '" & Trim(request("agecomparedate")) & "' "
	End If 

	iGenderRestrictionId = request("genderrestrictionid")

	iMinGrade = ""
	iMaxGrade = ""

	If request("notes") = "" Then 
 	 	sNotes = " NULL " 
	Else 
 		 sNotes = "'" & dbsafe(Trim(request("notes"))) & "'"
	End If 

	If Trim(request("startdate")) = "" Then 
		dStartdate =  " NULL " 
	Else 
		dStartdate = " '" & Trim(request("startdate")) & "' "
	End If 

	If Trim(request("enddate")) = "" Then 
		dEnddate =  " NULL " 
	Else 
		dEnddate = " '" & Trim(request("enddate")) & "' "
	End If 

	If Trim(request("registrationstartdate")) = "" Then 
		dRegistrationStartdate =  " NULL " 
	Else 
		dRegistrationStartdate = " '" & Trim(request("registrationstartdate")) & "' "
	End If 

	If Trim(request("registrationenddate")) = "" Then 
		dRegistrationEnddate =  " NULL " 
	Else 
		dRegistrationEnddate = " '" & Trim(request("registrationenddate")) & "' "
	End If 

	If Trim(request("publishstartdate")) = "" Then 
		dPublishStartdate =  " NULL " 
	Else 
		dPublishStartdate = " '" & Trim(request("publishstartdate")) & "' "
	End If 

	If Trim(request("publishenddate")) = "" Then 
		dPublishEnddate =  " NULL " 
	Else 
		dPublishEnddate = " '" & Trim(request("publishenddate")) & "' "
	End If 

	If Trim(request("evaluationdate")) = "" Then 
		dEvaluationDate =  " NULL " 
	Else 
		dEvaluationDate = " '" & Trim(request("evaluationdate")) & "' "
	End If 

	If Trim(request("alternatedate")) = "" Then 
		dAlternateDate =  " NULL " 
	Else 
		dAlternateDate = " '" & Trim(request("alternatedate")) & "' "
	End If 

	If Trim(request("imgurl")) = "" Then
		sImgUrl = " NULL "
	Else
		sImgUrl = " '" & dbsafe(Trim(request("imgurl"))) & "' "
	End If 
	If Trim(request("imgalttag")) = "" Then
		sImgAltTag = " NULL "
	Else
		sImgAltTag = " '" & dbsafe(Trim(request("imgalttag"))) & "' "
	End If 
	If Trim(request("externalurl")) = "" Then
		sExternalurl = " NULL "
	Else
		sExternalurl = " '" & dbsafe(Trim(request("externalurl"))) & "' "
	End If 
	If Trim(request("externallinktext")) = "" Then
		sExternallinktext = " NULL "
	Else
		sExternallinktext = " '" & dbsafe(Trim(request("externallinktext"))) & "' "
	End If 

	iMembershipId = " NULL "
	' There should only be one membershipid. This finds it on the pricing row
	For Each Item In request("pricetypeid")
		If CLng(request("membershipid" & Item)) <> CLng(0) Then
			iMembershipId = CLng(request("membershipid" & Item))
		End If 
	Next 

	If CLng(request("pricediscountid")) = CLng(0) Then
		iPriceDiscountId = " NULL "
	Else
		iPriceDiscountId = CLng(request("pricediscountid"))
	End If

	If request("allowearlyregistration") = "on" Then
		  sAllowEarlyRegistration         = "1"
		  sEarlyRegistrationDate          = "'" & request("earlyregistrationdate") & "'"
		  sEarlyRegistrationClassSeasonId = request("earlyregistrationclassseasonid")
		  sEarlyRegistrationClassId       = "NULL"    ' request("earlyregistrationclassid")
	Else
		  sAllowEarlyRegistration         = "0"
		  sEarlyRegistrationDate          = "NULL"
		  sEarlyRegistrationClassSeasonId = "NULL"
		  sEarlyRegistrationClassId       = "NULL"
	End If 

	If request("publiccanonlyview") = "on" Then 
		sPublicCanOnlyView = "1"
	Else
		sPublicCanOnlyView = "0"
	End If 
	If request("norefunds") = "on" Then 
		sNoRefunds = "1"
	Else
		sNoRefunds = "0"
	End If 

	If request("showTerms") = "on" Then 
		sShowTerms = "1"
	Else 
		sShowTerms = "0"
	End If 

	'Team Roster/Registration fields
	if request("displayrosterpublic") = "on" then
		lcl_displayrosterpublic = "1"
	else
		lcl_displayrosterpublic = "0"
	end if

' if request("teamreg_tshirt_accessorytype") = "" then
'    lcl_teamreg_tshirt_accessorytype = "NULL"
' else
'    lcl_teamreg_tshirt_accessorytype = "'" & dbsafe(request("teamreg_tshirt_accessorytype")) & "'"
' end if

	'if request("teamreg_tshirt_enabled") = "on" then
	'	lcl_teamreg_tshirt_enabled = "1"
	'else
	'	lcl_teamreg_tshirt_enabled = "0"
	'end if

	' if request("teamreg_pants_accessorytype") = "" then
	'    lcl_teamreg_pants_accessorytype = "NULL"
	' else
	'    lcl_teamreg_pants_accessorytype = "'" & dbsafe(request("teamreg_pants_accessorytype")) & "'"
	' end if

	'if request("teamreg_pants_enabled") = "on" then
	'	lcl_teamreg_pants_enabled = "1"
	'else
	'	lcl_teamreg_pants_enabled = "0"
	'end if

 lcl_teamreg_tshirt_enabled   = "NULL"
 lcl_teamreg_pants_enabled    = "NULL"
 lcl_teamreg_grade_enabled    = "NULL"
 lcl_teamreg_coach_enabled    = "NULL"
 lcl_teamreg_tshirt_inputtype = "NULL"
 lcl_teamreg_pants_inputtype  = "NULL"
 lcl_teamreg_grade_inputtype  = "NULL"

 if request("teamreg_tshirt_enabled") <> "" then
    lcl_teamreg_tshirt_enabled = ucase(request("teamreg_tshirt_enabled"))
    lcl_teamreg_tshirt_enabled = dbsafe(lcl_teamreg_tshirt_enabled)
    lcl_teamreg_tshirt_enabled = "'" & lcl_teamreg_tshirt_enabled & "'"
 end if

 if request("teamreg_pants_enabled") <> "" then
    lcl_teamreg_pants_enabled = ucase(request("teamreg_pants_enabled"))
    lcl_teamreg_pants_enabled = dbsafe(lcl_teamreg_pants_enabled)
    lcl_teamreg_pants_enabled = "'" & lcl_teamreg_pants_enabled & "'"
 end if

 if request("teamreg_grade_enabled") <> "" then
    lcl_teamreg_grade_enabled = ucase(request("teamreg_grade_enabled"))
    lcl_teamreg_grade_enabled = dbsafe(lcl_teamreg_grade_enabled)
    lcl_teamreg_grade_enabled = "'" & lcl_teamreg_grade_enabled & "'"
 end if

 if request("teamreg_coach_enabled") <> "" then
    lcl_teamreg_coach_enabled = ucase(request("teamreg_coach_enabled"))
    lcl_teamreg_coach_enabled = dbsafe(lcl_teamreg_coach_enabled)
    lcl_teamreg_coach_enabled = "'" & lcl_teamreg_coach_enabled & "'"
 end if

	if request("teamreg_tshirt_inputtype") <> "" then
		  lcl_teamreg_tshirt_inputtype = ucase(request("teamreg_tshirt_inputtype"))
    lcl_teamreg_tshirt_inputtype = dbsafe(lcl_teamreg_tshirt_inputtype)
    lcl_teamreg_tshirt_inputtype = "'" & lcl_teamreg_tshirt_inputtype & "'"
	end if

	if request("teamreg_pants_inputtype") <> "" then
		  lcl_teamreg_pants_inputtype = ucase(request("teamreg_pants_inputtype"))
    lcl_teamreg_pants_inputtype = dbsafe(lcl_teamreg_pants_inputtype)
    lcl_teamreg_pants_inputtype = "'" & lcl_teamreg_pants_inputtype & "'"
	end if

	if request("teamreg_grade_inputtype") <> "" then
		  lcl_teamreg_grade_inputtype = ucase(request("teamreg_grade_inputtype"))
    lcl_teamreg_grade_inputtype = dbsafe(lcl_teamreg_grade_inputtype)
    lcl_teamreg_grade_inputtype = "'" & lcl_teamreg_grade_inputtype & "'"
	end if

'Update the class table
	sSql = "UPDATE egov_class SET "
	sSql = sSql & " classname = '"                     & DBsafe(request("classname"))        & "', "
	sSql = sSql & " classdescription = '"              & DBsafe(request("classdescription")) & "', "
	sSql = sSql & " searchkeywords = '"                & DBsafe(request("searchkeywords"))   & "', "
	sSql = sSql & " minage = "                         & iMinAge                             & ", "
	sSql = sSql & " maxage = "                         & iMaxAge                             & ", "
	sSql = sSql & " genderrestrictionid = "            & iGenderRestrictionId                & ", "
	sSql = sSql & " classseasonid = "                  & iClassSeasonId                      & ", "
	sSql = sSql & " startdate = "                      & dStartdate                          & ", "
	sSql = sSql & " enddate = "                        & dEnddate                            & ", "
	sSql = sSql & " registrationstartdate = "          & dRegistrationStartdate              & ", "
	sSql = sSql & " registrationenddate = "            & dRegistrationEnddate                & ", "
	sSql = sSql & " publishstartdate = "               & dPublishStartdate                   & ", "
	sSql = sSql & " publishenddate = "                 & dPublishEnddate                     & ", "
	sSql = sSql & " evaluationdate = "                 & dEvaluationDate                     & ", "
	sSql = sSql & " alternatedate = "                  & dAlternateDate                      & ", "
	sSql = sSql & " imgurl = "                         & sImgUrl                             & ", "
	sSql = sSql & " locationid = "                     & request("locationid")               & ", "
	sSql = sSql & " pocid = "                          & CLng(request("pocid"))              & ", "
	sSql = sSql & " imgalttag = "                      & sImgAltTag                          & ", "
	sSql = sSql & " externalurl = "                    & sExternalurl                        & ", "
	sSql = sSql & " externallinktext = "               & sExternallinktext                   & ", "
	sSql = sSql & " optionid = "                       & request("optionid")                 & ", "
	sSql = sSql & " membershipid = "                   & iMembershipId                       & ", "
	sSql = sSql & " pricediscountid = "                & iPriceDiscountId                    & ", "
	sSql = sSql & " notes = "                          & sNotes                              & ", "
	sSql = sSql & " supervisorid = "                   & iSupervisorid                       & ", "
	sSql = sSql & " minageprecisionid = "              & iMinAgePrecisionId                  & ", "
	sSql = sSql & " maxageprecisionid = "              & iMaxAgePrecisionId                  & ", "
	sSql = sSql & " agecomparedate = "                 & dAgeComparedate                     & ", "
	sSql = sSql & " allowearlyregistration = "         & sAllowEarlyRegistration             & ", "
	sSql = sSql & " earlyregistrationdate = "          & sEarlyRegistrationDate              & ", "
	sSql = sSql & " earlyregistrationclassseasonid = " & sEarlyRegistrationClassSeasonId     & ", "
	sSql = sSql & " displayrosterpublic = "            & lcl_displayrosterpublic             & ", "
	sSql = sSql & " teamreg_tshirt_enabled = "         & lcl_teamreg_tshirt_enabled          & ", "
	sSql = sSql & " teamreg_pants_enabled = "          & lcl_teamreg_pants_enabled           & ", "
	sSql = sSql & " teamreg_grade_enabled = "          & lcl_teamreg_grade_enabled           & ", "
	sSql = sSql & " teamreg_coach_enabled = "          & lcl_teamreg_coach_enabled           & ", "
	sSql = sSql & " teamreg_tshirt_inputtype = "       & lcl_teamreg_tshirt_inputtype        & ", "
	sSql = sSql & " teamreg_pants_inputtype = "        & lcl_teamreg_pants_inputtype         & ", "
	sSql = sSql & " teamreg_grade_inputtype = "        & lcl_teamreg_grade_inputtype         & ", "
	sSql = sSql & " showTerms = "                      & sShowTerms							 & ", "
	sSql = sSql & " publiccanonlyview = "              & sPublicCanOnlyView         & ", "
	sSql = sSql & " norefunds = "              & sNoRefunds
	'sSql = sSql & ", earlyregistrationclassid = " & sEarlyRegistrationClassId
	sSql = sSql & " WHERE classid = "  & iClassId

	RunSQL sSql

'	Set oCmd = Server.CreateObject("ADODB.Command")
'	With oCmd
'		.ActiveConnection = Application("DSN")

		' Update the class table
'		.CommandText = sSql
		'response.write sSql
		'session("updclassSql") = sSql
'		.Execute
		'session("updclassSql") = ""

'************************************************************************************************

		' Update the Catetories
		sSql = "Delete from egov_class_category_to_class where classid = " & iClassId
		RunSQL sSql

		For Each Cat In Request("categoryid")
			sSql = "Insert INTO egov_class_category_to_class (classid, categoryid) Values ( " & iClassId & ", " & Cat & " )"
			RunSQL sSql
		Next

		' Update the Waiver Rows
		sSql = "Delete from egov_class_to_waivers where classid = " & iClassId
		sSql = sSql
		RunSQL sSql
		For Each Item In request("waiverId")
			Add_ClassWaiver iClassId, Item 
		Next 

		' Update the Instructors
		sSql = "Delete from egov_class_to_instructor where classid = " & iClassId
		RunSQL sSql
'		.CommandText = sSql
'		.execute
		For Each Item In request("instructorid")
			Add_ClassInstructor iClassId, Item 
		Next 

		' Update the Early Registration
		sSql = "Delete from egov_class_earlyregistrations where classid = " & iClassId
		RunSQL sSql
'		.CommandText = sSql
'		.execute
		If request("allowearlyregistration") = "on" Then
			For Each Item In request("earlyregistrationclassid")
				AddEarlyRegistrationClass iClassId, sEarlyRegistrationClassSeasonId, Item 
			Next 
		End If 

		' update the egov_class_pricetype_price - Pricing by citizen type/membership
		' clear out the old ones
		sSql = "Delete from egov_class_pricetype_price where classid = " & iClassId
		RunSQL sSql
'		.execute
		' add in the current set of Prices
		For Each Item In request("pricetypeid")
			If CLng(request("accountid" & Item)) = CLng(0) Then
				iAccountId = "NULL"
			Else
				iAccountId = CLng(request("accountid" & Item))
			End If 
			If Trim(request("registrationstartdate" & Item)) = "" Then 
				dRegistrationStartDate = Null 
			Else
				dRegistrationStartDate = request("registrationstartdate" & Item)
			End If 
			If CLng(request("membershipid" & Item)) = CLng(0) then
				iMembershipId = Null 
			Else
				iMembershipId = CLng(request("membershipid" & Item))
			End If 
			Add_ClassPrice iClassId, Item, request("amount" & Item), iAccountId, clng(request("instructorpercent" & Item)), dRegistrationStartDate, iMembershipId
		Next 

		' Loop through each time row
		For x = 0 To clng(request("maxtimeid"))
			' If marked for deletion 
			If request("delete" & x) = "on" Then 
				' If not the last timeday row
				If Not IsLastTimeDay( request("timeid" & x) ) Then 
					' delete timeday
					sSql = "Delete From egov_class_time_days where timedayid = " & CLng(request("timedayid" & x))
					RunSQL sSql
'					.CommandText = sSql
'					.execute
				Else ' is Last timeday row
					' if no one enrolled for this time slot
					If getEnrolledCount( request("timeid" & x) ) = CLng(0) Then 
						' delete timeday
						sSql = "Delete From egov_class_time_days where timedayid = " & CLng(request("timedayid" & x))
						RunSQL sSql
'						.CommandText = sSql
'						.execute
						' delete time
						sSql = "Delete From egov_class_time where timeid = " & CLng(request("timeid" & x))
						RunSQL sSql
'						.CommandText = sSql
'						.execute
					End If 
				End If 
			Else ' Update or add new
				If request("su" & x) = "on" Then 
					iSu = 1
				Else
					iSu = 0
				End If 
				If request("mo" & x) = "on" Then 
					iMo = 1
				Else
					iMo = 0
				End If 
				If request("tu" & x) = "on" Then 
					iTu = 1
				Else
					iTu = 0
				End If 
				If request("we" & x) = "on" Then 
					iWe = 1
				Else
					iWe = 0
				End If 
				If request("th" & x) = "on" Then 
					iTh = 1
				Else
					iTh = 0
				End If 
				If request("fr" & x) = "on" Then 
					iFr = 1
				Else
					iFr = 0
				End If 
				If request("sa" & x) = "on" Then 
					iSa = 1
				Else
					iSa = 0
				End If 
				If request("min" & x) = "" Then
					iMin = " NULL "
				Else 
					If clng(request("min" & x)) = clng(0) Then
						iMin = " NULL "
					Else
						iMin = clng(request("min" & x))
					End If 
				End If 
				If request("max" & x) = "" Then
					iMax = " NULL "
				Else 
					If clng(request("max" & x)) = clng(0) Then 
						iMax = " NULL "
					Else
						iMax = clng(request("max" & x))
					End If 
				End If 
				If request("waitlistmax" & x) = "" Then
					iWaitlistmax = " NULL "
				Else 
'					If clng(request("waitlistmax" & x)) = clng(0) Then
'						iWaitlistmax = " NULL "
'					Else
						iWaitlistmax = clng(request("waitlistmax" & x))
'					End If 
				End If 
				If CLng(request("instructorid" & x )) = CLng(0) Then
					iInstructorId = " NULL "
				Else
					iInstructorId = CLng(request("instructorid" & x ))
				End If 
				If request("iscanceled" & x ) = "on" Then
					iIsCanceled = " 1 "
				Else
					iIsCanceled = " 0 "
				End If 

				' if timeid <> 0 it already exists, so update it
				If CLng(request("timeid" & x)) > CLng(0) Then
					If Trim(request("activity" & x)) <> "skip" Then 
						' update the time
						sSql = "UPDATE egov_class_time SET activityno = '" & dbsafe(Trim(request("activity" & x))) & "', instructorid = " & iInstructorId 
						sSql = sSql & ", min = " & iMin & ", max = " & iMax & ", waitlistmax = " & iWaitlistmax & ", iscanceled = " & iIsCanceled
						sSql = sSql & " WHERE timeid = " & CLng(request("timeid" & x))
						RunSQL sSql
'						.CommandText = sSql
'						response.write "<br />" & sSql
'						.execute
					End If 

					' update the timeday
					sSql = "UPDATE egov_class_time_days SET starttime = '" & UCase(request("starttime" & x)) & "', endtime = '" & UCase(request("endtime" & x)) 
					sSql = sSql & "', sunday = " & iSu & ", monday = " & iMo & ", tuesday = " & iTu & ", wednesday = " & iWe 
					sSql = sSql & ", thursday = " & iTh & ", friday = " & iFr & ", saturday = " & iSa
					sSql = sSql & " WHERE timedayid = " & CLng(request("timedayid" & x))
					RunSQL sSql
'					.CommandText = sSql
'					response.write "<br />" & sSql
'					.execute

				Else  ' New ones
					response.write "<br /> New One [" & request("activity" & x) & "]"
					' If activityid in not empty
					If request("activity" & x) <> "" Then 
						' if activityid is used
						If timeIdExists( iClassId, request("activity" & x) ) Then 
							' Get the time id
							iTimeId = getTimeId( iClassId, request("activity" & x) ) 
							response.write "<br /> Existing Time ID [" & iTimeId & "]"
						Else ' New activity
							' insert the time - in class_global_functions.asp
							' Add_ClassTime( iClassId, iMin, iMax, iWaitlistmax, sActivityNo, iInstructorId, iEnrollmentsize )
							iTimeId = Add_ClassTime( iClassId, request("min" & x), request("max" & x), request("waitlistmax" & x), Trim(request("activity" & x)), request("instructorid" & x ), 0, 0, 0, 0, "NULL" )
							response.write "<br /> New Time ID [" & iTimeId & "]"
						End If 
						' insert the timeday - in class_global_functions.asp
						Add_ClassTimeDays iTimeId, UCase(request("starttime" & x)), UCase(request("endtime" & x)), iSu, iMo, iTu, iWe, iTh, iFr, iSa
					End If 
				End If 

			End If 
		Next 
		' end loop
		'response.End 

'	End With
'	Set oCmd = Nothing

	' Pull the class times for updating the total hours and meeting counts
	sSql = "SELECT timeid FROM egov_class_time WHERE classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	' loop through the set and get the values then update egov_class_time
	Do While Not oRs.EOF
		iMeetingCount = 0
		dHours = 0.0
		iMeetingCount = GetActivityMeetingCount( iClassId, oRs("timeid"), dHours )
		RunSQL( "UPDATE egov_class_time SET meetingcount = " & iMeetingCount & ", totalhours = " & FormatNumber( dHours,2,,,0 ) & " WHERE timeid = " & oRs("timeid") )
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 
	
	'response.end
	' Return to the edit page
	response.redirect "edit_class.asp?classid=" & iClassId & "&s=u"

%>

<!--#Include file="class_global_functions.asp"--> 

<%
'-------------------------------------------------------------------------------------------------
'  void RunSQL sSql
'-------------------------------------------------------------------------------------------------
Sub RunSQL( ByVal sSql )
	Dim oCmd

	'response.write sSql & "<br /><br />"
	'response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub 


'-------------------------------------------------------------------------------------------------
' boolean IsLastTimeDay( iTimeId )
'-------------------------------------------------------------------------------------------------
Function IsLastTimeDay( ByVal iTimeId )
	Dim sSql, oRs

	sSql = "Select Count(timedayid) as hits from egov_class_time_days where timeid = " & iTimeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(1) Then 
			IsLastTimeDay = False 
		Else
			IsLastTimeDay = True 
		End If 
	Else
		IsLastTimeDay = True  
	End If 

	oRs.close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer getEnrolledCount( iTimeId )
'-------------------------------------------------------------------------------------------------
Function getEnrolledCount( ByVal iTimeId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(classlistid) AS attendee_count FROM egov_class_list WHERE classtimeid = " & iTimeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		getEnrolledCount = CLng(oRs("attendee_count"))
	Else
		getEnrolledCount = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' boolean timeIdExists( iClassId, sActivityNo )
'-------------------------------------------------------------------------------------------------
Function timeIdExists( ByVal iClassId, ByVal sActivityNo )
	Dim sSql, oRs

	sSql = "SELECT COUNT(timeid) AS hits FROM egov_class_time "
	sSql = sSql & "WHERE classid = " & iClassId & " AND activityno = '" & dbsafe(Trim(sActivityNo)) & "'"
	'response.write "<br />" & sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then 
			timeIdExists = True  
		Else
			timeIdExists = False 
		End If 
	Else
		timeIdExists = False   
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer getTimeId( iClassId, sActivityNo )
'-------------------------------------------------------------------------------------------------
Function getTimeId( ByVal iClassId, ByVal sActivityNo )
	Dim sSql, oRs

	sSql = "SELECT timeid FROM egov_class_time "
	sSql = sSql & "WHERE classid = " & iClassId & " AND activityno = '" & dbsafe(Trim(sActivityNo)) & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		getTimeId = Clng(oRs("timeid"))
	Else
		getTimeId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 



%>
