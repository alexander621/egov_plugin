<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: class_createclass.asp
' AUTHOR: Steve Loar
' CREATED: 04/25/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This script creates new classes and events
'
' MODIFICATION HISTORY
' 1.0	04/25/06	Steve Loar - INITIAL VERSION
' 1.2	02/28/07	Steve Loar - Added classSeasonId
' 2.0	03/08/07	Steve Loar - Totaly redid how this works for Menlo Park project
' 2.1	02/15/08	Steve Loar - Early Registration added
' 2.2 12/30/08  David Boyer - Added "DisplayRosterPublic" checkbox for Craig, CO custom registration fields.
' 2.3 06/17/09	 David Boyer - Added "Show Terms" checkbox
' 2.4	12/2/2009	Steve Loar - Option to only allow purchases on admin but display on public
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sSql, oCmd, iClassId, iMinAge, iMaxAge, dStartdate, dEnddate, dPublishStartdate, dPublishEnddate
Dim dEvaluationDate, dAlternateDate, sImgUrl, sExternalurl, Item, x, sMin, sMax, sWaitlistmax, sExternallinktext
Dim bIsParent, iClassTypeId, sClassName, sClassdescription, iClassSeasonId, iMinGrade, iMaxGrade
Dim sActivityNumber, sNotes, iMembershipId, dRegistrationStartDate, iAccountId, iTimeId
Dim iSu, iMo, iTu, iWe, iTh, iFr, iSa, iMaxTimeID, iMinAgePrecisionId, iMaxAgePrecisionId
Dim sAllowEarlyRegistration, sEarlyRegistrationDate, sEarlyRegistrationClassSeasonId, sEarlyRegistrationClassId
Dim sShowTerms, sPublicCanOnlyYiew, iNewClassId, iGenderRestrictionId

iClassTypeId   = CLng(request("classtypeid"))
iClassSeasonId = CLng(request("classseasonid"))

If CLng(iClassTypeId) = CLng(1) Then
	bIsParent = True
Else
	bIsparent = False 
End If 

iMinAgePrecisionId = "NULL"
iMaxAgePrecisionId = "NULL"

If request("minage") = "" Then
	iMinage = 0
Else
	iMinage = CDbl(request("minage"))
	iMinAgePrecisionId = request("minageprecisionid")
End If 

If request("maxage") = "" Then
	iMaxAge = 0
Else
	iMaxAge = CDbl(request("maxage"))
	iMaxAgePrecisionId = request("maxageprecisionid")
End If 

iGenderRestrictionId = request("genderrestrictionid")

'	iMinGrade = request("mingrade")
iMinGrade = ""

'	iMaxGrade = request("maxgrade")
iMaxGrade = ""

If request("notes") = "" Then
	sNotes = ""
Else
	sNotes = dbsafe(Trim(request("notes")))
End If 

'	iMembershipId = "NULL"
iMembershipId = 0
' There should only be one membershipid. This finds it on the pricing row
For Each Item In request("pricetypeid")
	If CLng(request("membershipid" & Item)) <> CLng(0) Then
		iMembershipId = CLng(request("membershipid" & Item))
	End If 
Next 

If request("allowearlyregistration") = "on" Then
	sAllowEarlyRegistration = "1"
	sEarlyRegistrationDate = "'" & request("earlyregistrationdate") & "'"
	sEarlyRegistrationClassSeasonId = request("earlyregistrationclassseasonid")
	sEarlyRegistrationClassId = "NULL"     ' request("earlyregistrationclassid")
Else
	sAllowEarlyRegistration = "0"
	sEarlyRegistrationDate = "NULL"
	sEarlyRegistrationClassSeasonId = "NULL"
	sEarlyRegistrationClassId = "NULL"
End If

if request("displayrosterpublic") = "on" then
	sDisplayRosterPublic = "1"
else
	sDisplayRosterPublic = "0"
end if

if request("showTerms") = "on" then
	sShowTerms = "1"
else
	sShowTerms = "0"
end If
blnNoRefunds = 0

'Create New Class - In class_global_functions.asp
iNewClassId = Add_Class( request("classname"), request("classdescription"), 0, 0, bIsparent, 1, request("imgurl"), _
				request("registrationstartdate"), request("registrationenddate"), request("evaluationdate"), _
				request("alternatedate"), iMinage, iMaxAge, iGenderRestrictionId, request("locationid"), CLng(request("pocid")), _
				request("searchkeywords"), request("externalurl"), request("externallinktext"), iClasstypeid, _
				request("optionid"), 0, 1, "", request("startdate"), request("enddate"), request("publishstartdate"), _
				request("publishenddate"), request("imgalttag"), iMembershipId, request("pricediscountid"), iClassSeasonId, _
				iMinGrade, iMaxGrade, CLng(request("supervisorid")), sNotes, iMinAgePrecisionId, iMaxAgePrecisionId, _
				request("agecomparedate"), sAllowEarlyRegistration, sEarlyRegistrationDate, sEarlyRegistrationClassSeasonId, _
				sEarlyRegistrationClassId, sDisplayRosterPublic, sShowTerms, blnNoRefunds )

' Add in the public can only view flag
If request("publiccanonlyview") = "on" Then 
	sPublicCanOnlyYiew = "1"
Else
	sPublicCanOnlyYiew = "0"
End If 
sSql = "UPDATE egov_class SET publiccanonlyview = " & sPublicCanOnlyYiew & " WHERE classid = " & iNewClassId
RunSQLCommand sSql

' Create categories
For Each Item In Request("categorycheckid")
	Add_ClassCategory iNewClassId, Item 
Next

' Create Waiver Rows
For Each Item In request("waiverId")
	Add_ClassWaiver iNewClassId, Item 
Next 

' Create Instructors
For Each Item In request("instructorid")
	Add_ClassInstructor iNewClassId, Item 
Next 

' Update the Early Registration
If request("allowearlyregistration") = "on" Then
	For Each Item In request("earlyregistrationclassid")
		AddEarlyRegistrationClass iNewClassId, sEarlyRegistrationClassSeasonId, Item 
	Next 
End If 

' Add Prices 
For Each Item In request("pricetypeid")
	If CLng(request("accountid" & Item)) = CLng(0) Then
		iAccountId = "NULL"
	Else
		iAccountId = CLng(request("accountid" & Item))
	End If 
	If Trim(request("registrationstartdate" & Item)) = "" Then 
		dRegistrationStartDate = null
	Else
		dRegistrationStartDate = request("registrationstartdate" & Item)
	End If 
	If CLng(request("membershipid" & Item)) = CLng(0) then
		iMembershipId = Null 
	Else
		iMembershipId = CLng(request("membershipid" & Item))
	End If 
	Add_ClassPrice iNewClassId, Item, request("amount" & Item), iAccountId, clng(request("instructorpercent" & Item)), dRegistrationStartDate, iMembershipId
Next 

' Create the time Rows
' Shoud start a 0 and increment by 1
cOldActivity = ":qwerty:"
x = clng(0)
'response.write "<br />maxtimeid=" & clng(request("maxtimeid"))
iMaxTimeID = CLng(request("maxtimeid")) + 1
Do While x < iMaxTimeID
	If request("min" & x) = "" Then
		iMin = 0
	Else 
		If clng(request("min" & x)) = clng(0) Then
			iMin = 0
		Else
			iMin = clng(request("min" & x))
		End If 
	End If 
	If request("max" & x) = "" Then
		iMax = 0
	Else 
		If clng(request("max" & x)) = clng(0) Then 
			iMax = 0
		Else
			iMax = clng(request("max" & x))
		End If 
	End If 
	If request("waitlistmax" & x) = "" Then
		iWaitlistmax = ""
	Else 
'			If clng(request("waitlistmax" & x)) = clng(0) Then
'				iWaitlistmax = 0
'			Else
			iWaitlistmax = clng(request("waitlistmax" & x))
'			End If 
	End If 
	If CLng(request("instructorid" & x )) = CLng(0) Then
		iInstructorId = 0
	Else
		iInstructorId = CLng(request("instructorid" & x ))
	End If 
	If request("activity" & x) <> "" Then 
		If Trim(request("activity" & x)) <> cOldActivity Then 
			'New activity, so add it
			response.write "<br />Add_ClassTime( " & iNewClassId & ", " & iMin & ", " & iMax & ", " & iWaitlistmax & ", " & Trim(request("activity" & x)) & ", " & iInstructorId & ", 0, ""NULL"" )"
			iTimeId = Add_ClassTime( iNewClassId, iMin, iMax, iWaitlistmax, Trim(request("activity" & x)), iInstructorId, 0, 0, 0, 0, "NULL" )
			cOldActivity = Trim(request("activity" & x))
		End If 
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

		Add_ClassTimeDays iTimeId, request("starttime" & x), request("endtime" & x), iSu, iMo, iTu, iWe, iTh, iFr, iSa
	End If 
	x = x + 1
Loop 

' Pull the class times for updating the total hours and meeting counts
sSql = "SELECT timeid FROM egov_class_time WHERE classid = " & iNewClassId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 0, 1

' loop through the set and get the values then update egov_class_time
Do While Not oRs.EOF
	iMeetingCount = 0
	dHours = 0.0
	iMeetingCount = GetActivityMeetingCount( iNewClassId, oRs("timeid"), dHours )
	RunSQLCommand( "UPDATE egov_class_time SET meetingcount = " & iMeetingCount & ", totalhours = " & FormatNumber( dHours,2,,,0 ) & " WHERE timeid = " & oRs("timeid") )
	oRs.MoveNext 
Loop

oRs.Close
Set oRs = Nothing 

' go to the edit page
response.redirect "edit_class.asp?classid=" & iNewClassId & "&s=n"


%>

<!--#Include file="class_global_functions.asp"-->
