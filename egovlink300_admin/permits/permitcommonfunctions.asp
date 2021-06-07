<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitcommonfunctions.asp
' AUTHOR: Steve Loar
' CREATED: 04/15/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is a collection of shared functions for permits. Try to keep in alphabetical order.
'
' MODIFICATION HISTORY
' 1.0   04/15/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' boolean AllFeesPaidOrWaived( iPermitId )
'-------------------------------------------------------------------------------------------------
Function AllFeesPaidOrWaived( ByVal iPermitId )
	Dim sSql, oRs, dPaidFees, dWaivedTotal, unpaid_balance

	dPaidFees = GetPaidTotal( iPermitId ) 	' in permitcommonfunctions.asp
	'response.write "dPaidFees = " & dPaidFees & "<br />"
	dWaivedTotal = GetWaivedTotal( iPermitId ) 	' in permitcommonfunctions.asp
	'response.write "dWaivedTotal = " & dWaivedTotal & "<br />"

	sSql = "SELECT ISNULL(feetotal,0.00) AS feetotal "
	sSql = sSql & " FROM egov_permits WHERE permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		dFeeTotal = FormatNumber(oRS("feetotal"),2,,,0)
		
		'response.write "feetotal = " & dFeeTotal & "<br />"
		If CDbl(dFeeTotal) > CDbl(0.00) Then
			unpaid_balance = CDbl(dFeeTotal) - CDbl(dPaidFees) - CDbl(dWaivedTotal)
			unpaid_balance = CDbl(FormatNumber(unpaid_balance,2,,,0))
			response.write "<!-- fee total: " & dFeeTotal & " | paid fees: " & dPaidFees & " | waived total: " & dWaivedTotal & " | balance value: " & unpaid_balance & " -->"
			If unpaid_balance = CDbl(0) Then 
				AllFeesPaidOrWaived = True  
				'response.write "AllFeesPaidOrWaived = " & AllFeesPaidOrWaived & "<br />"
			Else
				AllFeesPaidOrWaived = False 
				'response.write "AllFeesPaidOrWaived = " & AllFeesPaidOrWaived & "<br />"
			End If 
		Else
			AllFeesPaidOrWaived = True   
		End If 
	Else
		AllFeesPaidOrWaived = True 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'-------------------------------------------------------------------------------------------------
' boolean AllOtherInspectionsAreDone( iPermitId, iPermitInspectionId )
'-------------------------------------------------------------------------------------------------
Function AllOtherInspectionsAreDone( ByVal iPermitId, ByVal iPermitInspectionId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(I.permitinspectionid) AS hits "
	sSql = sSql & " FROM egov_permitinspections I, egov_inspectionstatuses S "
	sSql = sSql & " WHERE I.inspectionstatusid = S.inspectionstatusid AND S.isdone = 0 AND I.permitid = " & iPermitId
	sSql = sSql & " AND I.permitinspectionid != " & iPermitInspectionId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then
			AllOtherInspectionsAreDone = False 
		Else
			AllOtherInspectionsAreDone = True 
		End If 
	Else
		AllOtherInspectionsAreDone = True 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string BuildPermitNoSearch( sPermitNo )
'--------------------------------------------------------------------------------------------------
Function BuildPermitNoSearch( ByVal sPermitNo )
	Dim sSql, oRs, iStartno, iCharacters, sSearchStr, sCharactersToMatch, sCharactersLength
	' This is tough. You need to deconstruct the passed permit number into the parts that make up a permit number for this org

	sSearchStr = ""
	iStartno = CLng(1)
	sPermitNo = sPermitNo & "0000000000"

	sSql = "SELECT position, characters, permitfield, ISNULL(datatype,'') AS datatype FROM egov_permitnumberformat "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " ORDER BY position"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While NOT oRs.EOF
		sCharactersToMatch = ""
		sCharactersLength = 1
		Select Case oRs("datatype")
			Case "varchar"
				' we need to search for mixed length prefixes, like 1 or 2 for Milford
				sCharactersToMatch = Trim(Mid(sPermitNo, iStartno, CLng(oRs("characters"))))
				sCharactersLength = CLng(Len(sCharactersToMatch))
				If CLng(sCharactersLength) > CLng(0) Then 
					sSearchStr = sSearchStr & " AND RIGHT( P." & oRs("permitfield") & ", " & sCharactersLength & " ) = '" & sCharactersToMatch & "' "
				End If 
			Case "int"
				If IsNumeric(Mid(sPermitNo, iStartno, CLng(oRs("characters")))) Then 
					sCharactersLength = CLng(oRs("characters"))
					sSearchStr = sSearchStr & " AND P." & oRs("permitfield") & " = " & CLng(Mid(sPermitNo, iStartno, sCharactersLength))
				End If 
		End Select 

		iStartno = iStartno + sCharactersLength
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	BuildPermitNoSearch = sSearchStr

End Function 


'-------------------------------------------------------------------------------------------------
' integer Ceiling( n )
'-------------------------------------------------------------------------------------------------
Function Ceiling( ByVal n )
	Dim f

	On Error Resume Next 
	n = CDbl(n)

	f = Floor(n)
	If f = n Then 
		Ceiling = n
		Exit Function
	End If

	Ceiling = CLng(f + 1)

End Function


'-------------------------------------------------------------------------------------------------
' boolean CheckIfStatusCanChange( iPermitId, iPermitStatusId )
'-------------------------------------------------------------------------------------------------
Function CheckIfStatusCanChange( ByVal iPermitId, ByVal iPermitStatusId )
	Dim sSql, oRs, bNeedsReviews, bNeedsFeesPaid, bNeedsLicenses

	' Check if the permit status needs reviews completed to move on.
	sSql = "SELECT needsreviewstochange, needsfeespaidtochange, needslicensestochange FROM egov_permitstatuses "
	sSql = sSql & " WHERE permitstatusid = " & iPermitStatusId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	bHasReqLicenses = true
	If Not oRs.EOF Then
		bNeedsReviews = oRs("needsreviewstochange")
		bNeedsFeesPaid = oRs("needsfeespaidtochange")
		bNeedsLicenses = oRs("needslicensestochange")
	Else
		bNeedsReviews = False 
		bNeedsFeesPaid = False 
		bNeedsLicenses = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

	bSomeFeesSetToZero = SomeFeesSetToZero( iPermitId )
	bAllFeesPaidOrWaived = AllFeesPaidOrWaived( iPermitId )
	bAllReviewsComplete = AllReviewsComplete( iPermitId )

	If bNeedsReviews Then
		CheckIfStatusCanChange = bAllReviewsComplete
	Else
		If bNeedsFeesPaid Then
			' See if all fees are either waived or paid (total fees = waived + paid) and that no fees are set to $0.00
			' This should be a check to clear before moving from approved to issued
			If bSomeFeesSetToZero Then
				CheckIfStatusCanChange = False 
			Else
				If bAllFeesPaidOrWaived Then
					CheckIfStatusCanChange = True 
				Else
					CheckIfStatusCanChange = False 
				End If 
			End If 
		Else
			If PermitStatusAllowsButton( iPermitId ) Then 
				CheckIfStatusCanChange = True 
			Else
				CheckIfStatusCanChange = False 
			End If 
		End If 
		If bNeedsLicenses And CheckIfStatusCanChange = True Then
			' See if the permit requires licenses and if so 
			'	check if any contact has a currently valid license of that type for each type required.
			' This should be a check to clear before moving from approved to issued
			' Only want to flag them for false (failed to pass)
			If PermitHasLicenseRequirement( iPermitId ) Then
				' Pull the required licenses as a string that can be used to check the contacts
				sRequiredTypes = GetRequiredLicenseTypeIdsAsString( iPermitId )
				If sRequiredTypes <> "" Then 
					' Check if any contacts have these licenses
					sSql = "SELECT COUNT(C.licenseid) AS hits "
					sSql = sSql & " FROM egov_permitcontacts_licenses C, egov_permitlicensetypes L, egov_permitcontacts P "
					sSql = sSql & " WHERE P.permitid = " & iPermitId & " AND L.licensetypeid =  C.licensetypeid AND P.ispriorcontact = 0 "
					sSql = sSql & " AND isbillingcontact = 0 AND P.permitid = C.permitid AND P.permitcontactid = C.permitcontactid "
					sSql = sSql & " AND L.licensetypeid IN ( " & sRequiredTypes & " ) AND licenseenddate >= getdate() "

					Set oRs = Server.CreateObject("ADODB.Recordset")
					oRs.Open sSql, Application("DSN"), 3, 1

					If Not oRs.EOF Then
						If CLng(oRs("hits")) = CLng(0) Then
							CheckIfStatusCanChange = False
							bHasReqLicenses = false
						End If 
					Else 
						CheckIfStatusCanChange = False
						bHasReqLicenses = false
					End If 
					oRs.Close
					Set oRs = Nothing 
				End If 
			End If 
		End If 
	End If 

End Function 

Function AllReviewsComplete( ByVal iPermitId )
	retVal = true

	' Check the review statuses for the permit
	' This should be a check to clear before moving from released to approved
	sSql = "SELECT COUNT(R.permitreviewid) AS hits "
	sSql = sSql & " FROM egov_permitreviews R, egov_reviewstatuses S "
	sSql = sSql & " WHERE S.allowpermitissue = 0 AND R.reviewstatusid = S.reviewstatusid AND R.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then 
			' At least one review is holding up the permit
			retVal = False 
		Else
			retVal = True 
		End If 
	Else
		retVal = True 
	End If 

	oRs.Close
	Set oRs = Nothing

	AllReviewsComplete = retVal
End Function

Function AllInspectionsPassed( ByVal iPermitId )
	retVal = false
	sSQL = "SELECT i.permitinspectionid  " _
		& " FROM egov_permitinspections i " _
		& " INNER JOIN egov_inspectionstatuses s ON i.inspectionstatusid = s.inspectionstatusid " _
		& " WHERE permitid = '" & iPermitId & "' and s.ispassed = 0"
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	if oRs.EOF then retVal = true
	oRs.Close
	Set oRs = Nothing
	AllInspectionsPassed = retVal
End Function


'-------------------------------------------------------------------------------------------------
' boolean CheckIfStatusCanIssueCOorTempCO( iPermitStatusId, sCOType )
'-------------------------------------------------------------------------------------------------
Function CheckIfStatusCanIssueCOorTempCO( ByVal iPermitStatusId, ByVal sCOType )
	Dim sSql, oRs

	sSql = "SELECT canissue" & sCOType & "co AS canissue FROM egov_permitstatuses "
	sSql = sSql & " WHERE permitstatusid = " & iPermitStatusId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("canissue") Then 
			CheckIfStatusCanIssueCOorTempCO = True 
		Else
			CheckIfStatusCanIssueCOorTempCO = False 
		End If 
	Else
		CheckIfStatusCanIssueCOorTempCO = False 
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' void CreateFixtureStepFees iPermitFixtureId, iPermitId, iPermitFeeId, iPermitFixtureTypeId 
'-------------------------------------------------------------------------------------------------
Sub CreateFixtureStepFees( ByVal iPermitFixtureId, ByVal iPermitId, ByVal iPermitFeeId, ByVal iPermitFixtureTypeId )
	Dim sSql, oRs

	sSql = "SELECT fixturetypestepfeeid, ISNULL(atleastqty, 0) AS atleastqty, ISNULL(notmorethanqty, 999999999) AS notmorethanqty, "
	sSql = sSql & " ISNULL(baseamount,0.00) AS baseamount, ISNULL(unitqty,1) AS unitqty, ISNULL(unitamount,0.00) AS unitamount "
	sSql = sSql & " FROM egov_permitfixturetypestepfees WHERE permitfixturetypeid = " & iPermitFixtureTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		sSql = "INSERT INTO egov_permitfixturestepfees ( permitid, permitfeeid, permitfixtureid, fixturetypestepfeeid, "
		sSql = sSql & " atleastqty, notmorethanqty, baseamount, unitqty, unitamount ) VALUES ( "
		sSql = sSql & iPermitId & ", " & iPermitFeeId & ", " & iPermitFixtureId & ", " & oRS("fixturetypestepfeeid") & ", "
		sSql = sSql & oRs("atleastqty") & ", " & oRs("notmorethanqty") & ", " & oRs("baseamount") & ", "
		sSql = sSql & oRs("unitqty") & ", " & oRs("unitamount") & " )"
		RunSQL sSql
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void CreateNewLicenseRecords( iContactTypeId, iContactId, iPermitId )
'-------------------------------------------------------------------------------------------------
Sub CreateNewLicenseRecords( ByVal iContactTypeId, ByVal iContactId, ByVal iPermitId )
	Dim sSql, oRs, sLicenseEndDate

	sSql = "SELECT licensetypeid, licensenumber, licenseenddate, licensee FROM egov_permitcontacttype_licenses "
	sSql = sSql & "WHERE permitcontacttypeid = " & iContactTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		If IsNull(oRs("licenseenddate")) Then
			sLicenseEndDate = "NULL"
		Else
			sLicenseEndDate = "'" & dbsafe(oRs("licenseenddate")) & "'"
		End If 
		sSql = "INSERT INTO egov_permitcontacts_licenses ( permitcontactid, permitid, licensetypeid, licensenumber, licenseenddate, licensee ) VALUES ( "
		sSql = sSql & iContactId & ", " & iPermitId &  ", " & dbsafe(oRs("licensetypeid")) & ", '" & dbsafe(oRs("licensenumber")) & "', "
		sSql = sSql & sLicenseEndDate & ",'" & dbsafe(oRs("licensee")) & "' )"
		RunSQL sSql
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void  CreatePermitFeeMultipliers iPermitFeeTypeId, iPermitFeeId, iPermitId 
'-------------------------------------------------------------------------------------------------
Sub CreatePermitFeeMultipliers( ByVal iPermitFeeTypeId, ByVal iPermitFeeId, ByVal iPermitId )
	Dim sSql, oRs, iRowCount

	iRowCount = 0
	sSql = "SELECT F.feemultipliertypeid, F.feemultiplier, F.feemultiplierrate "
	sSql = sSql & " FROM egov_feemultipliertypes F, egov_permitfeetypes_to_feemultipliertypes T "
	sSql = sSql & " WHERE F.feemultipliertypeid = T.feemultipliertypeid AND T.permitfeetypeid = " & iPermitFeeTypeId
	sSql = sSql & " ORDER BY T.displayorder, F.feemultipliertypeid"

	'response.write "<p> Multiplier: " & sSql & "</p><br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		iRowCount = iRowCount + 1
		sSql = "INSERT INTO egov_permitfeemultipliers ( permitid, permitfeeid, permitfeetypeid, feemultipliertypeid, "
		sSql = sSql & " orgid, feemultiplier, feemultiplierrate, displayorder ) VALUES ( " & iPermitId & ", " 
		sSql = sSql & iPermitFeeId & ", " & iPermitFeeTypeId & ", " & oRs("feemultipliertypeid") & ", " 
		sSql = sSql & session("orgid") & ", '" & dbsafe(oRs("feemultiplier")) & "', " & oRs("feemultiplierrate")
		sSql = sSql & ", " & iRowCount & " )"
		RunSQL sSql
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void CreatePermitFixtures( iPermitId, iPermitFeeTypeId, iPermitFeeId )
'-------------------------------------------------------------------------------------------------
Sub CreatePermitFixtures( ByVal iPermitId, ByVal iPermitFeeTypeId, ByVal iPermitFeeId )
	Dim sSql, oRs, iPermitFixtureId, iFixtureCount

	iFixtureCount = 0

	sSql = "SELECT F.permitfixturetypeid, F.permitfixture, ISNULL(F.displayorder,9999) AS displayorder "
	sSql = sSql & " FROM egov_permitfixturetypes F, egov_permitfeetypes_to_permitfixturetypes T "
	sSql = sSql & " WHERE T.permitfixturetypeid = F.permitfixturetypeid AND T.permitfeetypeid = " & iPermitFeeTypeId
	sSql = sSql & " ORDER BY displayorder, F.permitfixturetypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		iFixtureCount = iFixtureCount + 1

		sSql = "INSERT INTO egov_permitfixtures ( permitid, permitfeeid, permitfixturetypeid, orgid, permitfixture, displayorder ) VALUES ( "
		sSql = sSql & iPermitId & ", " & iPermitFeeId & ", " & oRs("permitfixturetypeid") & ", " & session("orgid") & ", '"
		sSql = sSql & dbsafe(oRs("permitfixture")) & "', " & oRs("displayorder") & " )"
		iPermitFixtureId = RunIdentityInsert( sSql )

		' Input the step table entries for each fixture
		CreateFixtureStepFees iPermitFixtureId, iPermitId, iPermitFeeId, oRs("permitfixturetypeid")
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void CreatePermitResidentialUnitStepFees( iPermitId, iPermitFeeTypeId, iPermitFeeId )
'-------------------------------------------------------------------------------------------------
Sub CreatePermitResidentialUnitStepFees( ByVal iPermitId, ByVal iPermitFeeTypeId, ByVal iPermitFeeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(atleastqty, 0) AS atleastqty, ISNULL(notmorethanqty, 999999999) AS notmorethanqty, "
	sSql = sSql & " ISNULL(baseamount,0.00) AS baseamount, ISNULL(unitqty,1) AS unitqty, ISNULL(unitamount,0.00) AS unitamount "
	sSql = sSql & " FROM egov_permitresidentialunittypestepfees WHERE permitfeetypeid = " & iPermitFeeTypeid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		sSql = "INSERT INTO egov_permitresidentialunitstepfees ( permitid, permitfeeid, atleastqty, notmorethanqty, baseamount, unitqty, unitamount ) VALUES ( "
		sSql = sSql & iPermitId & ", " & iPermitFeeId & ", " & oRs("atleastqty") & ", " & oRs("notmorethanqty") & ", "
		sSql = sSql & oRs("baseamount") & ", " & oRs("unitqty") & ", " & oRs("unitamount") & " )"
		'response.write sSql & "<br />"
		RunSQL sSql
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' Sub CreatePermitValuations( iPermitId, iPermitFeeTypeId, iPermitFeeId )
'-------------------------------------------------------------------------------------------------
'Sub CreatePermitValuations( iPermitId, iPermitFeeTypeId, iPermitFeeId )
'	Dim sSql, oRs, iPermitValuationId, iValuationCount
'
'	iValuationCount = 0
'
'	sSql = "SELECT F.permitvaluationtypeid, F.permitvaluation "
'	sSql = sSql & " FROM egov_permitvaluationtypes F, egov_permitfeetypes_to_permitvaluationtypes T "
'	sSql = sSql & " WHERE T.permitvaluationtypeid = F.permitvaluationtypeid AND T.permitfeetypeid = " & iPermitFeeTypeId
'	sSql = sSql & " ORDER BY displayorder, F.permitvaluationtypeid"
'
'	Set oRs = Server.CreateObject("ADODB.Recordset")
'	oRs.Open sSql, Application("DSN"), 3, 1
'
'	Do While Not oRs.EOF 
'		iValuationCount = iValuationCount + 1
'
'		sSql = "INSERT INTO egov_permitvaluations ( permitid, permitfeeid, permitvaluationtypeid, orgid, permitvaluation, displayorder ) VALUES ( "
'		sSql = sSql & iPermitId & ", " & iPermitFeeId & ", " & oRs("permitvaluationtypeid") & ", " & session("orgid") & ", '"
'		sSql = sSql & dbsafe(oRs("permitvaluation")) & "', " & iValuationCount & " )"
'		iPermitValuationId = RunIdentityInsert( sSql )
'		' Input the step table entries for each valuation
'		CreateValuationStepFees iPermitValuationId, iPermitId, iPermitFeeId, oRs("permitvaluationtypeid")
'		oRs.MoveNext
'	Loop 
'
'	oRs.Close
'	Set oRs = Nothing 
'End Sub 


'-------------------------------------------------------------------------------------------------
' void  CreatePermitValuationStepFees( iPermitId, iPermitFeeTypeId, iPermitFeeId )
'-------------------------------------------------------------------------------------------------
Sub CreatePermitValuationStepFees( ByVal iPermitId, ByVal iPermitFeeTypeId, ByVal iPermitFeeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(atleastvalue, 0.00) AS atleastvalue, ISNULL(notmorethanvalue, 999999999.99) AS notmorethanvalue, "
	sSql = sSql & " ISNULL(baseamount,0.00) AS baseamount, ISNULL(unitqty,1) AS unitqty, ISNULL(unitamount,0.00) AS unitamount "
	sSql = sSql & " FROM egov_permitvaluationtypestepfees WHERE permitfeetypeid = " & iPermitFeeTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		sSql = "INSERT INTO egov_permitvaluationstepfees ( permitid, permitfeeid, "
		sSql = sSql & " atleastvalue, notmorethanvalue, baseamount, unitqty, unitamount ) VALUES ( "
		sSql = sSql & iPermitId & ", " & iPermitFeeId & ", "
		sSql = sSql & oRs("atleastvalue") & ", " & oRs("notmorethanvalue") & ", " & oRs("baseamount") & ", "
		sSql = sSql & oRs("unitqty") & ", " & oRs("unitamount") & " )"
		RunSQL sSql
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------------------------------------
' void DRAWDATECHOICES SNAME
'------------------------------------------------------------------------------------------------------------
Sub DrawDateChoices( sName )

	response.write vbcrlf & "<select onChange=""getDates(this.value, '" & sName & "');"" class=""calendarinput"" name=""" & sName & """>"
	response.write vbcrlf & "<option value=""0"">Or Select Date Range from Dropdown...</option>"
	response.write vbcrlf & "<option value=""16"">Today</option>"
	response.write vbcrlf & "<option value=""17"">Yesterday</option>"
	response.write vbcrlf & "<option value=""18"">Tomorrow</option>"
	response.write vbcrlf & "<option value=""11"">This Week</option>"
	response.write vbcrlf & "<option value=""12"">Last Week</option>"
	response.write vbcrlf & "<option value=""14"">Next Week</option>"
	response.write vbcrlf & "<option value=""1"">This Month</option>"
	response.write vbcrlf & "<option value=""2"">Last Month</option>"
	response.write vbcrlf & "<option value=""13"">Next Month</option>"
	response.write vbcrlf & "<option value=""3"">This Quarter</option>"
	response.write vbcrlf & "<option value=""4"">Last Quarter</option>"
	response.write vbcrlf & "<option value=""15"">Next Quarter</option>"
	response.write vbcrlf & "<option value=""6"">Year to Date</option>"
	response.write vbcrlf & "<option value=""5"">Last Year</option>"
	response.write vbcrlf & "<option value=""7"">All Dates to Date</option>"
	response.write vbcrlf & "</select>"

End Sub 


'------------------------------------------------------------------------------------------------------------
' void DrawPriorDateChoices SNAME
'------------------------------------------------------------------------------------------------------------
Sub DrawPriorDateChoices( sName )

	response.write vbcrlf & "<select onChange=""getDates(this.value, '" & sName & "');"" class=""calendarinput"" name=""" & sName & """>"
	response.write vbcrlf & "<option value=""0"">Or Select Date Range from Dropdown...</option>"
	response.write vbcrlf & "<option value=""16"">Today</option>"
	response.write vbcrlf & "<option value=""17"">Yesterday</option>"
	response.write vbcrlf & "<option value=""11"">This Week</option>"
	response.write vbcrlf & "<option value=""12"">Last Week</option>"
	response.write vbcrlf & "<option value=""1"">This Month</option>"
	response.write vbcrlf & "<option value=""2"">Last Month</option>"
	response.write vbcrlf & "<option value=""3"">This Quarter</option>"
	response.write vbcrlf & "<option value=""4"">Last Quarter</option>"
	response.write vbcrlf & "<option value=""6"">Year to Date</option>"
	response.write vbcrlf & "<option value=""5"">Last Year</option>"
	response.write vbcrlf & "<option value=""7"">All Dates to Date</option>"
	response.write vbcrlf & "</select>"

End Sub 


'-------------------------------------------------------------------------------------------------
' integer Floor(  n )
'-------------------------------------------------------------------------------------------------
Function Floor( ByVal n )
	Dim iTmp

	On Error Resume Next 
	n = CDbl(n)

	'Round() rounds up
	iTmp = Round(n)

	'test rounded value against the non rounded value
	'if greater, subtract 1
	If iTmp > n Then 
		iTmp = iTmp - 1
	End If 

	Floor = CLng(iTmp)

End Function


'-------------------------------------------------------------------------------------------------
' string FormatHTML( sHTMLBody )
'-------------------------------------------------------------------------------------------------
Function FormatHTML( ByVal sHTMLBody )
	Dim lcl_return

	lcl_return = "<html>" & vbcrlf
	lcl_return = lcl_return & "<head>" & vbcrlf
	lcl_return = lcl_return & "</head>" & vbcrlf
	lcl_return = lcl_return & "<body bgcolor=""#efefef"">" & vbcrlf
	lcl_return = lcl_return & "<font face=""helvetica, arial"">" & vbcrlf
	lcl_return = lcl_return & "<p style=""margin:0px""></p>" & vbcrlf
	lcl_return = lcl_return & "<table bordercolor=""#4A9E9F"" bgcolor=""#ffffff"" cellspacing=""0"" cellpadding=""5"" width=""95%"" align=""center"" border=""2"" valign=""top"">" & vbcrlf
	lcl_return = lcl_return & "<tr>" & vbcrlf
	lcl_return = lcl_return & "<td style=""font-family:arial,tahoma; font-size:12px; color:#000000;"">" & vbcrlf
	lcl_return = lcl_return & sHTMLBody
	lcl_return = lcl_return & "<center>" & vbcrlf
	lcl_return = lcl_return & "<br />" & vbcrlf
	lcl_return = lcl_return & "<hr color=""black"" size=""1"" width=""95%"">" & vbcrlf
	lcl_return = lcl_return & "<font size=""-2"">Copyright 2004 - " & year(now) & ". <i>Electronic Commerce</i> Link, Inc. dba <i>EC</i> Link.</font>" & vbcrlf
	lcl_return = lcl_return & "</center>" & vbcrlf
	lcl_return = lcl_return & "</td>" & vbcrlf
	lcl_return = lcl_return & "</tr>" & vbcrlf
	lcl_return = lcl_return & "</table>" & vbcrlf
	lcl_return = lcl_return & "</font>" & vbcrlf
	lcl_return = lcl_return & "</body>" & vbcrlf
	lcl_return = lcl_return & "</html>" & vbcrlf

	FormatHTML = lcl_return

End Function 


'-------------------------------------------------------------------------------------------------
' double GetAllFeesTotalForPercentage( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetAllFeesTotalForPercentage( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(feeamount),0.00) AS feeamount FROM egov_permitfees "
	sSql = sSql & " WHERE ispercentagetypefee = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If IsNull(oRs("feeamount")) Then
			GetAllFeesTotalForPercentage = CDbl(0.00)
		Else
			GetAllFeesTotalForPercentage = CDbl(oRs("feeamount"))
		End If 
	Else
		GetAllFeesTotalForPercentage = CDbl(0.00)
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double GetAllPermitFees( iPermitId, bIsOld, dSubTotal, sEndDate, sStartDate )
'--------------------------------------------------------------------------------------------------
Function GetAllPermitFees( ByVal iPermitId, ByVal bIsOld, ByRef dSubTotal, ByVal sEndDate, ByVal sStartDate )
	Dim sSql, oRs, sCompare, sPermitFees

	If bIsOld Then
		sSql = "SELECT ISNULL(SUM(II.invoicedamount),0.00) AS invoicedamount "
		sSql = sSql & " FROM egov_permitinvoiceitems II, egov_permitfeecategorytypes C, egov_permitinvoices I, egov_permits P "
		sSql = sSql & " WHERE II.permitid = " & iPermitId
		sSql = sSql & " AND I.permitid = P.permitid AND I.invoiceid = II.invoiceid "
		sSql = sSql & " AND II.permitfeecategorytypeid = C.permitfeecategorytypeid AND C.isgeneralbuildingtype = 1 AND "
		sSql = sSql & " I.invoicedate > '" & sStartDate & "' AND I.invoicedate < '" & sEndDate & "' "
		sSql = sSql & " AND I.isvoided = 0 AND I.allfeeswaived = 0"
	Else
		sSql = "SELECT ISNULL(SUM(II.invoicedamount),0.00) AS invoicedamount "
		sSql = sSql & " FROM egov_permitinvoiceitems II, egov_permitfeecategorytypes C, egov_permitinvoices I, egov_permits P "
		sSql = sSql & " WHERE II.permitid = " & iPermitId
		sSql = sSql & " AND I.permitid = P.permitid AND I.invoiceid = II.invoiceid "
		sSql = sSql & " AND II.permitfeecategorytypeid = C.permitfeecategorytypeid AND C.isgeneralbuildingtype = 1 AND "
		sSql = sSql & " I.invoicedate < '" & sEndDate & "' "
		sSql = sSql & " AND I.isvoided = 0 AND I.allfeeswaived = 0"
	End If 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sPermitFees = FormatNumber(oRs("invoicedamount"),2)
	Else
		sPermitFees = FormatNumber(0.00,2)
	End If 

	dSubTotal = dSubTotal + CDbl(sPermitFees)

	oRs.Close
	Set oRs = Nothing 

	GetAllPermitFees = sPermitFees

End Function 


'--------------------------------------------------------------------------------------------------
' double GetAllYTDPermitFees( sYearStart, sYearEnd, iInclude )
'--------------------------------------------------------------------------------------------------
Function GetAllYTDPermitFees( ByVal sYearStart, ByVal sYearEnd, ByVal iInclude )
	Dim sSql, oRs, dFees, sIsVoided

	If clng(iInclude ) < clng(2) Then
		sIsVoided = " AND P.isvoided = " & iInclude
	Else
		sIsVoided = ""
	End If 

	sSql = "SELECT ISNULL(SUM(II.invoicedamount),0.00) AS invoicedamount "
	sSql = sSql & " FROM egov_permitinvoiceitems II, egov_permitfeecategorytypes C, egov_permitinvoices I, egov_permits P "
	sSql = sSql & " WHERE I.orgid = " & session("orgid")
	sSql = sSql & " AND I.permitid = P.permitid AND I.invoiceid = II.invoiceid "
	sSql = sSql & " AND II.permitfeecategorytypeid = C.permitfeecategorytypeid AND C.isgeneralbuildingtype = 1 "
	sSql = sSql & " AND P.issueddate < '" & sYearEnd & "' AND P.issueddate > '" & sYearStart & "' AND I.invoicedate < '" & sYearEnd & "' "
	sSql = sSql & sIsVoided & " AND I.isvoided = 0 AND I.allfeeswaived = 0"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		dFees = FormatNumber(oRs("invoicedamount"),2)
	Else
		dFees = 0.00
	End If 

	If CDbl(dFees) = CDbl(0.00) Then
		dFees = 0.00
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetAllYTDPermitFees = FormatNumber(dFees,2)

End Function 


'-------------------------------------------------------------------------------------------------
' double GetAppliedFeeAmount( iPermitFeeId )
'-------------------------------------------------------------------------------------------------
Function GetAppliedFeeAmount( ByVal iPermitFeeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(feeamount,0.00) AS feeamount FROM egov_permitfees WHERE permitfeeid = " & iPermitFeeId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetAppliedFeeAmount = FormatNumber(oRs("feeamount"),2,,,0)
	Else
		GetAppliedFeeAmount = 0.00
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetArchiveContractor( iPermitid )
'--------------------------------------------------------------------------------------------------
Function GetArchiveContractor( ByVal iPermitid )
	Dim sSql, oRs, sContractor

	sContractor = ""

	sSql = "SELECT ISNULL(contractorcompany,'') AS contractorcompany, ISNULL(contractorname,'') AS contractorname "
	sSql = sSql & "FROM egov_permitarchives "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " AND permitid = " & iPermitid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If oRs("contractorcompany") <> "" Then 
			sContractor = oRs("contractorcompany")
		End If 

		If oRs("contractorname") <> "" Then 
			If sContractor <> "" Then 
				sContractor = sContractor & " &ndash; "
			End If 
			sContractor = sContractor & oRs("contractorname")
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetArchiveContractor = sContractor

End Function 


'-------------------------------------------------------------------------------------------------
' double GetCategoryFeesTotalForPercentage( iPermitId, iPermitFeeCategoryTypeId )
'-------------------------------------------------------------------------------------------------
Function GetCategoryFeesTotalForPercentage( ByVal iPermitId, ByVal iPermitFeeCategoryTypeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(feeamount),0.00) AS feeamount FROM egov_permitfees "
	sSql = sSql & " WHERE ispercentagetypefee = 0 AND permitid = " & iPermitId 
	sSql = sSql & " AND permitfeecategorytypeid = " & iPermitFeeCategoryTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If IsNull(oRs("feeamount")) Then
			GetCategoryFeesTotalForPercentage = CDbl(0.00)
		Else
			GetCategoryFeesTotalForPercentage = CDbl(oRs("feeamount"))
		End If 
	Else
		GetCategoryFeesTotalForPercentage = CDbl(0.00)
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetCheckNo( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function GetCheckNo( ByVal iPaymentId )
	Dim sSql, oRs

	sSql = "SELECT checkno FROM egov_verisign_payment_information "
	sSql = sSql & " WHERE checkno IS NOT NULL AND paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetCheckNo = oRs("checkno")
	Else
		GetCheckNo = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetContractorType( iContractorTypeId ) 
'-------------------------------------------------------------------------------------------------
Function GetContractorType( ByVal iContractorTypeId ) 
	Dim sSql, oRs, sReturn

	sSql = sSql & "SELECT contractortype, isother FROM egov_permitcontractortypes WHERE contractortypeid = " & iContractorTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(iContractorTypeId) > CLng(0) Then ' And Not oRs("isother") Then 
			GetContractorType = oRs("contractortype") ' & " Contractor"
		Else
			GetContractorType = "Contractor"
		End If 
	Else
		GetContractorType = "Contractor"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'-------------------------------------------------------------------------------------------------
' boolean GetContactTypeIsOrganization( iContactTypeid )
'-------------------------------------------------------------------------------------------------
Function GetContactTypeIsOrganization( ByVal iPermitContactTypeid )
	Dim sSql, oRs, sReturn

	sSql = sSql & "SELECT isorganization FROM egov_permitcontacttypes WHERE permitcontacttypeid = " & iPermitContactTypeid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("isorganization") Then 
			GetContactTypeIsOrganization = True 
		Else
			GetContactTypeIsOrganization = False 
		End If 
	Else
		GetContactTypeIsOrganization = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double GetCostEstimate( iPermitId, bIsOld )
'--------------------------------------------------------------------------------------------------
Function GetCostEstimate( ByVal iPermitId, ByVal bIsOld, ByRef dCostEstimateSubTotal, ByVal sEndDate, ByVal sStartDate )
	Dim sSql, oRs, sCompare, sCostEstimate

	If bIsOld Then
		sSql = "SELECT ISNULL(SUM(I.netjobvalue),0.00) AS costestimate "
		sSql = sSql & " FROM egov_permitinvoices I, egov_permits P "
		sSql = sSql & " WHERE I.permitid = P.permitid AND I.invoicedate > '" & sStartDate & "' AND I.invoicedate < '" & sEndDate & "' AND P.permitid = " & iPermitId
		sSql = sSql & " AND I.isvoided = 0 GROUP BY I.permitid"
		response.write "<!--" & sSql & "-->"
	Else
		sSql = "SELECT ISNULL(SUM(I.netjobvalue),0.00) AS costestimate "
		sSql = sSql & " FROM egov_permitinvoices I, egov_permits P "
		sSql = sSql & " WHERE I.permitid = P.permitid AND I.invoicedate < '" & sEndDate & "' AND P.permitid = " & iPermitId
		sSql = sSql & " AND I.isvoided = 0 GROUP BY I.permitid"
		response.write "<!--TWF:" & sSql & "-->"
	End If 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sCostEstimate = FormatNumber(oRs("costestimate"),2,,,0)
	Else
		sCostEstimate = FormatNumber(0.00,2,,,0)
	End If 

	dCostEstimateSubTotal = dCostEstimateSubTotal + CDbl(sCostEstimate)

	GetCostEstimate = sCostEstimate

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' double GetCurrentJobValue( iPermitId ) 
'-------------------------------------------------------------------------------------------------
Function GetCurrentJobValue( ByVal iPermitId ) 
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(jobvalue),0.00) AS currentjobvalue FROM egov_permits "
	sSql = sSql & " WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetCurrentJobValue = CDbl(FormatNumber(oRs("currentjobvalue"),2,,,0))
	Else
		GetCurrentJobValue = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetDescriptionOfWork( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetDescriptionOfWork( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(descriptionofwork,'') AS descriptionofwork FROM egov_permits WHERE permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetDescriptionOfWork = oRs("descriptionofwork")
	Else
		GetDescriptionOfWork = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' double GetFeeMultipliers( iPermitFeeId )
'-------------------------------------------------------------------------------------------------
Function GetFeeMultipliers( ByVal iPermitFeeId )
	Dim sSql, oRs, sRate

	sRate = 1.00
	sSql = "SELECT feemultiplierrate FROM egov_permitfeemultipliers WHERE permitfeeid = " & iPermitFeeId
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		' multiply the rates together
		sRate = CDbl(sRate) * CDbl(oRs("feemultiplierrate"))
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	'response.write "sRate = " & sRate & "<br />"
	GetFeeMultipliers = sRate

End Function 


'-------------------------------------------------------------------------------------------------
' double GetFeeMultipliersForDisplay( iPermitFeeId )
'-------------------------------------------------------------------------------------------------
Function GetFeeMultipliersForDisplay( ByVal iPermitFeeId )
	Dim sSql, oRs, sRate

	sRate = ""
	sSql = "SELECT feemultiplierrate FROM egov_permitfeemultipliers WHERE permitfeeid = " & iPermitFeeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		sRate = sRate & " * " & FormatNumber(oRs("feemultiplierrate"),2,,,0)
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	GetFeeMultipliersForDisplay = sRate

End Function 


'--------------------------------------------------------------------------------------------------
' string GetFeeReportingType( iFeeReportingTypeId )
'--------------------------------------------------------------------------------------------------
Function GetFeeReportingType( ByVal iFeeReportingTypeId )
	Dim sSql, oRs

	If CLng(0) = iFeeReportingTypeId Then
		GetFeeReportingType = "Building Permit Fees"
	Else
		sSql = "SELECT feereportingtype FROM egov_permitfeereportingtypes WHERE feereportingtypeid = " & iFeeReportingTypeId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then
			GetFeeReportingType = oRs("feereportingtype")
		Else
			GetFeeReportingType = "Fee Type"
		End If 
		oRs.Close
		Set oRs = Nothing 
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetFirstPaymentMethod( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function GetFirstPaymentMethod( ByVal iPaymentId )
	Dim sSql, oRs, sResult

	sSql = "SELECT ISNULL(L.amount,0.00) AS amount, P.paymenttypename, P.requirescheckno "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment J, egov_paymenttypes P "
	sSql = sSql & " WHERE L.paymentid = J.paymentid AND J.paymentid = " & iPaymentId
	sSql = sSql & " AND L.entrytype = 'debit' AND L.paymenttypeid = P.paymenttypeid"
	sSql = sSql & " ORDER BY requirescheckno DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sResult = oRs("paymenttypename") 
		If oRs("requirescheckno") Then 
			sResult = sResult & " #: " & GetCheckNo( iPaymentId )
		End If 
	Else
		sResult = ""
	End If 
	
	oRs.Close
	Set oRs = Nothing 

	GetFirstPaymentMethod = sResult

End Function 


'-------------------------------------------------------------------------------------------------
' double GetFixtureFeeAmount( iPermitFixtureId, iQty )
'-------------------------------------------------------------------------------------------------
Function GetFixtureFeeAmount( ByVal iPermitFixtureId, ByVal iQty )
	Dim sSql, oRs, sFeeAmount, iUnitQty

	If CLng(iQty) > CLng(0) Then 
		sSql = "SELECT atleastqty, notmorethanqty, baseamount, unitqty, unitamount "
		sSql = sSql & " FROM egov_permitfixturestepfees WHERE "
		sSql = sSql & iQty & " >= atleastqty AND " & iQty & " < notmorethanqty "
		sSql = sSql & " AND permitfixtureid = " & iPermitFixtureId 
		'response.write sSql & "<br /><br />"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then
			If CLng(request("unitqty")) <= CLng(0) Then
				iUnitQty = 1
			Else
				iUnitQty = CLng(request("unitqty"))
			End If 
			sFeeAmount = CLng(iQty) - CLng(oRs("atleastqty"))
			sFeeAmount = sFeeAmount / iUnitQty
			sFeeAmount = CDbl(sFeeAmount) * CDbl(oRs("unitamount"))
			sFeeAmount = sFeeAmount + CDbl(oRs("baseamount"))
			sFeeAmount = FormatNumber(sFeeAmount,2,,,0)
		Else
			sFeeAmount = 0.00
		End If 
		'response.write "sFeeAmount: " & sFeeAmount & "<br /><br />"

		oRs.Close
		Set oRs = Nothing
	Else
		sFeeAmount = 0.00
	End If 
	
	GetFixtureFeeAmount = sFeeAmount

End Function 


'-------------------------------------------------------------------------------------------------
' string GetFixtureName( iPermitFixtureId )
'-------------------------------------------------------------------------------------------------
Function GetFixtureName( ByVal iPermitFixtureId )
	Dim sSql, oRs

	sSql = "SELECT permitfixture FROM egov_permitfixtures WHERE permitfixtureid = " & iPermitFixtureId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetFixtureName = oRS("permitfixture")
	Else
		GetFixtureName = ""
	End If 

	oRs.Close
	Set oRs = Nothing
		
End Function 


'-------------------------------------------------------------------------------------------------
' integer GetFixtureTypeDisplayOrder( iPermitFixtureTypeid )
'-------------------------------------------------------------------------------------------------
Function GetFixtureTypeDisplayOrder( ByVal iPermitFixtureTypeid )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(displayorder,9999) AS displayorder FROM egov_permitfixturetypes WHERE permitfixturetypeid = " & iPermitFixtureTypeid
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetFixtureTypeDisplayOrder = CLng(oRs("displayorder"))
	Else 
		GetFixtureTypeDisplayOrder = CLng(9999)
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetInspectionStatusById( iInspectionStatusId )
'-------------------------------------------------------------------------------------------------
Function GetInspectionStatusById( ByVal iInspectionStatusId )
	Dim sSql, oRs

	sSql = "SELECT inspectionstatus FROM egov_inspectionstatuses WHERE inspectionstatusid = " & iInspectionStatusId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetInspectionStatusById = oRs("inspectionstatus")
	Else
		GetInspectionStatusById = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetInspectionStatusId( sStatusFlag )
'-------------------------------------------------------------------------------------------------
Function GetInspectionStatusId( ByVal sStatusFlag )
	Dim sSql, oRs

	sSql = "SELECT inspectionstatusid FROM egov_inspectionstatuses WHERE isforpermits = 1 AND orgid = " & session("orgid")
	sSql = sSql & " AND " & sStatusFlag & " = 1"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetInspectionStatusId = CLng(oRs("inspectionstatusid"))
	Else
		GetInspectionStatusId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetInvoiceContact( iPermitContactId )
'--------------------------------------------------------------------------------------------------
Function GetInvoiceContact( ByVal iPermitContactId )
	Dim sSql, oRs, sReturn

	sSql = "SELECT ISNULL(company,'') AS company, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname "
	sSql = sSql & " FROM egov_permitcontacts WHERE permitcontactid = " & iPermitContactId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If oRs("firstname") <> "" Then 
			sReturn = oRs("firstname") & " " & oRs("lastname")
		Else
			sReturn = oRs("company")
		End If 
	Else
		sReturn = ""
	End If
	
	oRs.Close
	Set oRs = Nothing 

	GetInvoiceContact = sReturn

End Function 


'-------------------------------------------------------------------------------------------------
' double GetInvoicedTotal( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetInvoicedTotal( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(totalamount),0.00) AS totalamount FROM egov_permitinvoices "
	sSql = sSql & "WHERE isvoided = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetInvoicedTotal =  FormatNumber(oRS("totalamount"),2,,,0)
	Else
		GetInvoicedTotal = "0.00"
	End If 

	oRs.Close
	Set oRs = Nothing
		
End Function 


'-------------------------------------------------------------------------------------------------
' string GetInvoicePaymentDate( iPaymentId ) 
'-------------------------------------------------------------------------------------------------
Function GetInvoicePaymentDate( ByVal iPaymentId )
	Dim sSql, oRs

	sSql = "SELECT paymentdate FROM egov_class_payment WHERE paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetInvoicePaymentDate =  DateValue(oRs("paymentdate"))
	Else
		GetInvoicePaymentDate = ""
	End If 

	oRs.Close
	Set oRs = Nothing
		
End Function 


'-------------------------------------------------------------------------------------------------
' double GetInvoicePaymentTotal( iInvoiceId ) 
'-------------------------------------------------------------------------------------------------
Function GetInvoicePaymentTotal( ByVal iInvoiceId ) 
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(amount),0.00) AS totalamount FROM egov_accounts_ledger "
	sSql = sSql & " WHERE invoiceid = " & iInvoiceId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetInvoicePaymentTotal =  FormatNumber(oRS("totalamount"),2,,,0)
	Else
		GetInvoicePaymentTotal = "0.00"
	End If 

	oRs.Close
	Set oRs = Nothing
		
End Function 


'-------------------------------------------------------------------------------------------------
' integer GetInvoiceStatusId( sStatusFlag )
'-------------------------------------------------------------------------------------------------
Function GetInvoiceStatusId( ByVal sStatusFlag )
	Dim sSql, oRs

	sSql = "SELECT invoicestatusid FROM egov_invoicestatuses WHERE isforpermits = 1 AND orgid = " & session("orgid")
	sSql = sSql & " AND " & sStatusFlag & " = 1"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetInvoiceStatusId =  CLng(oRS("invoicestatusid"))
	Else
		GetInvoiceStatusId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing
		
End Function 


'--------------------------------------------------------------------------------------------------
' boolean GetIsGroupByInvoiceCategories( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetIsGroupByInvoiceCategories( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT groupbyinvoicecategories FROM egov_permitpermittypes WHERE permitid = " & iPermitid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("groupbyinvoicecategories") Then
			GetIsGroupByInvoiceCategories = True 
		Else
			GetIsGroupByInvoiceCategories = False 
		End If 
	Else
		GetIsGroupByInvoiceCategories = False 
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetIsOrganizationFlag( iPermitContactTypeid )
'--------------------------------------------------------------------------------------------------
Function GetIsOrganizationFlag( ByVal iPermitContactTypeid )
	Dim sSql, oRs

	sSql = "SELECT isorganization FROM egov_permitcontacttypes WHERE permitcontacttypeid = " & iPermitContactTypeid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("isorganization") Then
			GetIsOrganizationFlag = clng(1)
		Else
			GetIsOrganizationFlag = clng(0)
		End If 
	Else
		GetIsOrganizationFlag = clng(0)
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetJournalEntryTypeID( sType )
'--------------------------------------------------------------------------------------------------
Function GetJournalEntryTypeID( ByVal sType )
	Dim sSql, oRs, sTypeId

	sSql = "SELECT journalentrytypeid FROM egov_journal_entry_types WHERE journalentrytype = '" & sType & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sTypeId = CLng(oRs("journalentrytypeid") )
	Else 
		sTypeId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing

	GetJournalEntryTypeID = sTypeId

End Function


'--------------------------------------------------------------------------------------------------
' string GetLastLogDate( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetLastLogDate( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT MAX(entrydate) AS entrydate FROM egov_permitlog WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetLastLogDate = DateValue(oRs("entrydate"))
	Else 
		GetLastLogDate = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string GetLastPermitInspectionNote( iPermitInspectionId )
'--------------------------------------------------------------------------------------------------
Function GetLastPermitInspectionNote( ByVal iPermitInspectionId )
	Dim sSql, oRs

	sSql = "SELECT externalcomment, entrydate FROM egov_permitlog "
	sSql = sSql & " WHERE externalcomment IS NOT NULL AND permitinspectionid = " & iPermitInspectionId
	sSql = sSql & " ORDER BY entrydate DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		' Possibly more than one, but we just want the latest public note.
		GetLastPermitInspectionNote = Trim(oRs("externalcomment"))
	Else 
		GetLastPermitInspectionNote = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetLastPermitReviewNote( iPermitReviewId )
'--------------------------------------------------------------------------------------------------
Function GetLastPermitReviewNote( ByVal iPermitReviewId )
	Dim sSql, oRs

	sSql = "SELECT externalcomment, entrydate FROM egov_permitlog "
	sSql = sSql & " WHERE externalcomment IS NOT NULL AND permitreviewid = " & iPermitReviewId
	sSql = sSql & " ORDER BY entrydate DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		' Possibly more than one, but we just want the latest public note.
		GetLastPermitReviewNote = Trim(oRs("externalcomment"))
	Else 
		GetLastPermitReviewNote = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' void GetLocationRequirements iPermitLocationRequirementId, bNeedsAddress, bNeedsLocation
'-------------------------------------------------------------------------------------------------
Sub GetLocationRequirements( ByVal iPermitLocationRequirementId, ByRef bNeedsAddress, ByRef bNeedsLocation )
	Dim sSql, oRs

	sSql = "SELECT needsaddress, needslocation FROM egov_permitlocationrequirements "
	sSql = sSql & " WHERE permitlocationrequirementid = " & iPermitLocationRequirementId
	'session("PermitListSql") = sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	'session("PermitListSql") = ""

	If Not oRs.EOF Then 
		If oRs("needsaddress") Then
			bNeedsAddress = True 
		Else
			bNeedsAddress = False 
		End If 
		If oRs("needslocation") Then
			bNeedsLocation = True 
		Else
			bNeedsLocation = False 
		End If 
	Else 
		bNeedsAddress = False 
		bNeedsLocation = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' double GetMiniumFeeAmount( iPermitFeeId )
'-------------------------------------------------------------------------------------------------
Function GetMiniumFeeAmount( ByVal iPermitFeeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(minimumamount,0.00) AS minimumamount FROM egov_permitfees WHERE permitfeeid = " & iPermitFeeId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetMiniumFeeAmount = FormatNumber(oRs("minimumamount"),2,,,0)
	Else
		GetMiniumFeeAmount = 0.00
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetNextInspectionOrder( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetNextInspectionOrder( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(MAX(inspectionorder),0) AS inspectionorder FROM egov_permitinspections WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetNextInspectionOrder = CLng(oRs("inspectionorder")) + CLng(1)
	Else
		GetNextInspectionOrder = CLng(1)
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' money GetPaidTotal( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPaidTotal( ByVal iPermitId )
	Dim sSql, oRs

	' Get the total paid from the Journal Table
	sSql = "SELECT ISNULL(SUM(L.amount),0.00) AS paymenttotal FROM egov_accounts_ledger L, egov_permitinvoices I "
	sSql = sSql & " WHERE I.isvoided = 0 AND L.invoiceid = I.invoiceid AND L.ispaymentaccount = 0 "
	sSql = sSql & " AND L.permitid = " & iPermitId
	'response.write sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPaidTotal =  FormatNumber(oRS("paymenttotal"),2,,,0)
	Else
		GetPaidTotal = "0.00"
	End If 

	oRs.Close
	Set oRs = Nothing
		
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetPaymentAccountId( iOrgId, iPaymentTypeId )
'--------------------------------------------------------------------------------------------------
Function GetPaymentAccountId( ByVal iOrgId, ByVal iPaymentTypeId )
	Dim sSql, oAccount

	sSql = "SELECT ISNULL(accountid,0) AS accountid FROM egov_organizations_to_paymenttypes "
	sSql = sSql & " WHERE orgid = " & iOrgId & " AND paymenttypeid = " & iPaymentTypeId

	Set oAccount = Server.CreateObject("ADODB.Recordset")
	oAccount.Open sSql, Application("DSN"), 3, 1

	If Not oAccount.EOF Then 
		GetPaymentAccountId = CLng(oAccount("accountid"))
	Else
		GetPaymentAccountId = CLng(0) 
	End If 

	oAccount.Close
	Set oAccount = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Sub GetPermitAlertDetails( ByVal iPermitId, ByRef sAlertMsg, ByRef sAlertSetByUser, ByRef dAlertDate )
'--------------------------------------------------------------------------------------------------
Sub GetPermitAlertDetails( ByVal iPermitId, ByRef sAlertMsg, ByRef sAlertSetByUser, ByRef dAlertDate )
	Dim sSql, oRs

	sSql = "SELECT alertmsg, alertsetbyuserid, alertdate FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If Not IsNull(oRs("alertmsg")) Then 
			sAlertMsg = oRs("alertmsg")
			sAlertSetByUser = GetAdminName( oRs("alertsetbyuserid") )
			dAlertDate = FormatDateTime(oRs("alertdate"), 2)
		Else 
			sAlertMsg = ""
			sAlertSetByUser = ""
			dAlertDate = ""
		End If 
	Else 
		sAlertMsg = ""
		sAlertSetByUser = ""
		dAlertDate = ""
	End If
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetPermitApplicantAddressLabel( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitApplicantAddressLabel( ByVal iPermitId )
	Dim sSql, oRs, sAddressLabel

	sSql = "SELECT permitcontactid, userid, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(address,'') AS address, ISNULL(city,'') AS city, "
	sSql = sSql & " ISNULL(state,'') AS state, ISNULL(zip,'') AS zip, ISNULL(phone,'') AS phone, contacttype " 
	sSql = sSql & " FROM egov_permitcontacts WHERE isapplicant = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("firstname") <> "" Then 
			sAddressLabel = oRs("firstname") & " " & oRs("lastname") & "<br />"
		End If 
		If Not IsNull(oRs("company")) And oRs("company") <> "" Then 
			If sAddressLabel = "" Then 
				sAddressLabel = oRs("company") & "<br />"
			Else 
				sAddressLabel = sAddressLabel & oRs("company") & "<br />"
			End If 
		End If 
		If Not IsNull(oRs("address")) And oRs("address") <> "" Then 
			sAddressLabel = sAddressLabel &  oRs("address") & "<br />"
		End If 
		If Not IsNull(oRs("city")) And oRs("city") <> "" Then
			sAddressLabel = sAddressLabel &  oRs("city") & ", " & oRs("state") & " " & oRs("zip") & "<br />" 
		End If 
	Else
		sAddressLabel = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetPermitApplicantAddressLabel = sAddressLabel
End Function 


'--------------------------------------------------------------------------------------------------
' void GetPermitApplicantInfo iPermitId, sName, sEmail, sPhone, sCell, sFax 
'--------------------------------------------------------------------------------------------------
Sub GetPermitApplicantInfo( ByVal iPermitId, ByRef sName, ByRef sEmail, ByRef sPhone, ByRef sCell, ByRef sFax, ByRef sApplicantAddress, ByRef sApplicantCity, ByRef sApplicantState, ByRef sApplicantZip )
	Dim sSql, oRs, sApplicant

	sName = ""
	sSql = "SELECT ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & "ISNULL(company,'') AS company, ISNULL(email,'') AS email, ISNULL(phone,'') AS phone, "
	sSql = sSql & "ISNULL(cell,'') AS cell, ISNULL(fax,'') AS fax, ISNULL(address,'') AS address,  " 
	sSql = sSql & "ISNULL(city,'') AS city, ISNULL(state,'') AS state, ISNULL(zip,'') AS zip "
	sSql = sSql & "FROM egov_permitcontacts WHERE isapplicant = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("firstname") <> "" Then 
			sName = oRs("firstname") & " " & oRs("lastname")
		End If 
		If oRs("company") <> "" And sName = "" Then 
			sName = oRs("company")
		End If 
		sEmail = oRs("email")
		sPhone = FormatPhoneNumber(oRs("phone"))
		sCell = FormatPhoneNumber(oRs("cell"))
		sFax = FormatPhoneNumber(oRs("fax"))
		sApplicantAddress = oRs("address")
		sApplicantCity = oRs("city")
		sApplicantState = oRs("state")
		sApplicantZip = oRs("zip")
	Else
		sName = ""
		sEmail = ""
		sPhone = ""
		sCell = ""
		sFax = ""
		sApplicantAddress = ""
		sApplicantCity = ""
		sApplicantState = ""
		sApplicantZip = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub  


'--------------------------------------------------------------------------------------------------
' string = GetPermitApplicantName( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitApplicantName( ByVal iPermitId )
	Dim sSql, oRs, sApplicant

	sApplicant = ""
	sSql = "SELECT ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company " 
	sSql = sSql & " FROM egov_permitcontacts WHERE isapplicant = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("firstname") <> "" Then 
			sApplicant = oRs("firstname") & " " & oRs("lastname") & "<br />"
		End If 
		If oRs("company") <> "" And sApplicant = "" Then 
			'sApplicant = sContact & oRs("company") & "<br />" 
			sApplicant = oRs("company") & "<br />" 
		End If 
	Else
		sApplicant = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetPermitApplicantName = sApplicant

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetPermitAttachments( iPermitid, iRow, sAttachmentDate, sAttachmentName, sAttachmentDescription )
'--------------------------------------------------------------------------------------------------
Function GetPermitAttachments( ByVal iPermitid, ByVal iRow, ByRef sAttachmentDate, ByRef sAttachmentName, ByRef sAttachmentDescription )
	Dim sSql, oRs, x
	' This gets information for the first attachment only

	sAttachmentDate = ""
	sAttachmentName = ""
	sAttachmentDescription = ""
	x = 0

	sSql = "SELECT dateadded, attachmentname, description FROM egov_permitattachments WHERE permitid = " & iPermitid
	sSql = sSql & " ORDER BY dateadded, attachmentname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			x = x + 1
			If CLng(x) = CLng(iRow) Then 
				sAttachmentDate = FormatDateTime(oRs("dateadded"),2)
				sAttachmentName = oRs("attachmentname")
				sAttachmentDescription = oRs("description")
				Exit Do  
			End If 
			oRs.MoveNext
		Loop 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetPermitBuildingFees( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitBuildingFees( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(II.invoicedamount),0.00) AS invoicedamount "
	sSql = sSql & " FROM egov_permitinvoiceitems II, egov_permitfeecategorytypes C, egov_permitinvoices I, egov_permits P "
	sSql = sSql & " WHERE II.permitid = " & iPermitId
	sSql = sSql & " AND I.permitid = P.permitid AND I.invoiceid = II.invoiceid "
	sSql = sSql & " AND II.permitfeecategorytypeid = C.permitfeecategorytypeid AND C.isgeneralbuildingtype = 1 "
	sSql = sSql & " AND II.feereportingtypeid IS NULL AND I.isvoided = 0 AND I.allfeeswaived = 0"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitBuildingFees = FormatNumber(oRs("invoicedamount"),2)
	Else
		GetPermitBuildingFees = FormatNumber(0.00,2)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetPermitConstructionType( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitConstructionType( iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(constructiontype,'') AS constructiontype FROM egov_permits P, egov_constructiontypes C "
	sSql = sSql & " WHERE P.constructiontypeid = C.constructiontypeid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitConstructionType = oRs("constructiontype")
	Else
		GetPermitConstructionType = ""
	End If
	
	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetPermitContactAndPhone( iPermitId, sContactType )
'--------------------------------------------------------------------------------------------------
Function GetPermitContactAndPhone( iPermitId, sContactType )
	Dim sSql, oRs, sContact

	sContact = ""
	sSql = " SELECT ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(phone,'') AS phone " 
	sSql = sSql & " FROM egov_permitcontacts WHERE " & sContactType & " = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("firstname") <> "" Then 
			sContact = oRs("firstname") & " " & oRs("lastname") & "<br />"
		End If 
		If oRs("company") <> "" Then 
			sContact = sContact & oRs("company") & "<br />" 
		End If 
		If Trim(oRs("phone")) <> "" Then 
			sContact = sContact & FormatPhoneNumber( oRs("phone") ) 
		End If 
	Else
		sContact = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetPermitContactAndPhone = sContact
End Function 


'-------------------------------------------------------------------------------------------------
' string GetPermitCurrentStatusDate( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitCurrentStatusDate( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT S.statusdatedisplayed, P.applieddate, P.releaseddate, P.approveddate, P.issueddate, P.completeddate FROM egov_permits P, egov_permitstatuses S"
	sSql = sSql & " WHERE P.permitstatusid = S.permitstatusid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		Select Case oRs("statusdatedisplayed") 
			Case "applieddate"
				GetPermitCurrentStatusDate = FormatDateTime(oRs("applieddate"),2)
			Case "releaseddate"
				GetPermitCurrentStatusDate = FormatDateTime(oRs("releaseddate"),2)
			Case "approveddate"
				GetPermitCurrentStatusDate = FormatDateTime(oRs("approveddate"),2)
			Case "issueddate"
				GetPermitCurrentStatusDate = FormatDateTime(oRs("issueddate"),2)
			Case "completeddate"
				GetPermitCurrentStatusDate = FormatDateTime(oRs("completeddate"),2)
		End Select 
	Else
		GetPermitCurrentStatusDate = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' number GetPermitDetailItemAsNumber( iPermitId, sField, sFormat )
'--------------------------------------------------------------------------------------------------
Function GetPermitDetailItemAsNumber( ByVal iPermitId, ByVal sField, ByVal sFormat )
	Dim sSql, oRs

	' Use this to grab numbers from the permit
	sSql = "SELECT ISNULL(" & sField & ",0.00) AS selectedfield FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If IsNull(oRs("selectedfield")) Then
			If sFormat = "integer" Then 
				GetPermitDetailItemAsNumber = "0"
			Else
				' Currency and double
				GetPermitDetailItemAsNumber = "0.00"
			End If 
		Else
			If sFormat = "currency" Then 
				GetPermitDetailItemAsNumber = FormatCurrency(oRs("selectedfield"),2)
			Else
				If sFormat = "integer" Then 
					GetPermitDetailItemAsNumber = FormatNumber(oRs("selectedfield"),0,,,0)
				Else
					' Double
					GetPermitDetailItemAsNumber = FormatNumber(oRs("selectedfield"),2,,,0)
				End If 
			End If 
		End If 
	Else
		If sFormat = "integer" Then 
			GetPermitDetailItemAsNumber = "0"
		Else
			' Currency and double
			GetPermitDetailItemAsNumber = "0.00"
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitDetailItemAsString( iPermitId, sField )
'--------------------------------------------------------------------------------------------------
Function GetPermitDetailItemAsString( ByVal iPermitId, ByVal sField )
	Dim sSql, oRs, sDetailItem
	' Use this to grab strings from the permit

	sSql = "SELECT ISNULL(" & sField & ",'') AS selectedfield FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If IsNull(oRs("selectedfield")) Then
			GetPermitDetailItemAsString = ""
		Else
			sDetailItem = oRs("selectedfield")
			' Replace m-dash with a dash so XML will not crash
			sDetailItem = Replace(sDetailItem,"–","-")
			GetPermitDetailItemAsString = Trim(sDetailItem)
		End If 
	Else
		GetPermitDetailItemAsString = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' date GetPermitDate( iPermitId, sDateField )
'--------------------------------------------------------------------------------------------------
Function GetPermitDate( ByVal iPermitId, ByVal sDateField )
	Dim sSql, oRs

	sSql = "SELECT " & sDateField & " AS keydate FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If IsNull(oRs("keydate")) Then 
			GetPermitDate = "" 
		Else
			GetPermitDate = FormatDateTime(oRs("keydate"),2) 
		End If 
	Else
		GetPermitDate = "" 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' Sub GetPermitDocumentShowFlags( iPermitId, bListFixtures, bShowConstructionType, bShowFeeTotal, bShowOccupancyType, bShowJobValue, bShowWorkDesc, bShowFootages, bShowProposedUse, bShowTotalSqFt, bShowApprovedAs, bShowFeeTypeTotals, bShowOccupancyUse )
'-------------------------------------------------------------------------------------------------
Sub GetPermitDocumentShowFlags( ByVal iPermitId, ByRef bListFixtures, ByRef bShowConstructionType, ByRef bShowFeeTotal, ByRef bShowOccupancyType, ByRef bShowJobValue, ByRef bShowWorkDesc, ByRef bShowFootages, ByRef bShowProposedUse, ByRef bShowOtherContacts, ByRef sShowElectricalContractor, ByRef sShowMechanicalContractor, ByRef sShowPlumbingContractor, ByRef sShowApplicantLicense, ByRef bShowCounty, ByRef bShowParcelid, ByRef bShowPlansBy, ByRef bShowPrimaryContact, ByRef bShowTotalSqFt, ByRef bShowApprovedAs, ByRef bShowFeeTypeTotals, ByRef bShowOccupancyUse, ByRef bShowPayments )
	Dim sSql, oRs

	sSql = "SELECT listfixtures, showconstructiontype, showfeetotal, showoccupancytype, "
	sSql = sSql & " showjobvalue, showworkdesc, showfootages, showproposeduse, showothercontacts, "
	sSql = sSql & " showelectricalcontractor, showmechanicalcontractor, showplumbingcontractor, showapplicantlicense, "
	sSql = sSql & " showcounty, showparcelid, showplansby, showprimarycontact, showtotalsqft, showapprovedas, "
	sSql = sSql & " showfeetypetotals, showoccupancyuse, showpayments "
	sSql = sSql & " FROM egov_permitpermittypes WHERE permitid = " & iPermitid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("listfixtures") Then
			bListFixtures = True 
		Else
			bListFixtures = False 
		End If 
		If oRs("showconstructiontype") Then
			bShowConstructionType = True 
		Else
			bShowConstructionType = False 
		End If
		If oRs("showfeetotal") Then
			bShowFeeTotal = True 
		Else
			bShowFeeTotal = False 
		End If
		If oRs("showoccupancytype") Then
			bShowOccupancyType = True 
		Else
			bShowOccupancyType = False 
		End If
		If oRs("showjobvalue") Then
			bShowJobValue = True 
		Else
			bShowJobValue = False 
		End If
		If oRs("showworkdesc") Then
			bShowWorkDesc = True 
		Else
			bShowWorkDesc = False 
		End If
		If oRs("showfootages") Then
			bShowFootages = True 
		Else
			bShowFootages = False 
		End If
		If oRs("showproposeduse") Then
			bShowProposedUse = True 
		Else
			bShowProposedUse = False 
		End If
		If oRs("showothercontacts") Then
			bShowOtherContacts = True 
		Else
			bShowOtherContacts = False 
		End If
		If oRs("showelectricalcontractor") Then 
			sShowElectricalContractor = True 
		Else
			sShowElectricalContractor = False 
		End If 
		If oRs("showmechanicalcontractor") Then 
			sShowMechanicalContractor = True 
		Else
			sShowMechanicalContractor = False 
		End If 
		If oRs("showplumbingcontractor") Then 
			sShowPlumbingContractor = True 
		Else
			sShowPlumbingContractor = False 
		End If 
		If oRs("showapplicantlicense") Then 
			sShowApplicantLicense = True 
		Else
			sShowApplicantLicense = False 
		End If 
		If oRs("showcounty") Then 
			bShowCounty = True 
		Else
			bShowCounty = False 
		End If 
		If oRs("showparcelid") Then 
			bShowParcelid = True 
		Else
			bShowParcelid = False 
		End If 
		If oRs("showplansby") Then 
			bShowPlansBy = True 
		Else
			bShowPlansBy = False 
		End If 
		If oRs("showprimarycontact") Then 
			bShowPrimaryContact = True 
		Else
			bShowPrimaryContact = False 
		End If 
		If oRs("showtotalsqft") Then
			bShowTotalSqFt = True 
		Else
			bShowTotalSqFt = False
		End If 
		If oRs("showapprovedas") Then
			bShowApprovedAs = True
		Else
			bShowApprovedAs = False
		End If 
		If oRs("showfeetypetotals") Then
			bShowFeeTypeTotals = True
		Else
			bShowFeeTypeTotals = False
		End If
		If oRs("showoccupancyuse") Then
			bShowOccupancyUse = True
		Else
			bShowOccupancyUse = False
		End If
		If oRs("showpayments") Then
			bShowPayments = True
		Else
			bShowPayments = False 
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' string/number GetPermitDocumentValue( iPermitId, sValueColumn )
'-------------------------------------------------------------------------------------------------
Function GetPermitDocumentValue( ByVal iPermitId, ByVal sValueColumn )
	Dim sSql, oRs

	sSql = "SELECT ISNULL( " & sValueColumn & ",'') AS valuecolumn " 
	sSql = sSql & " FROM egov_permitpermittypes T, egov_permits P "
	sSql = sSql & " WHERE T.permittypeid = P.permittypeid AND P.permitid = T.permitid AND P.permitid = " & iPermitId 
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitDocumentValue = oRs("valuecolumn")
	Else
		GetPermitDocumentValue = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' double GetPermitExamHours( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitExamHours( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(examinationhours,0.00) AS examinationhours FROM egov_permits WHERE permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitExamHours = FormatNumber(oRs("examinationhours"),2,,,0)
	Else
		GetPermitExamHours = 0.00
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'-------------------------------------------------------------------------------------------------
' double GetPermitFeeInvoicedAmount( iPermitFeeId )
'-------------------------------------------------------------------------------------------------
Function GetPermitFeeInvoicedAmount( ByVal iPermitFeeId )
	Dim sSql, oRs, sFeeName

	sSql = "SELECT ISNULL(invoicedamount,0.00) AS invoicedamount FROM egov_permitfees "
	sSql = sSql & " WHERE permitfeeid = " & iPermitFeeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitFeeInvoicedAmount = CDbl(oRs("invoicedamount"))
	Else 
		GetPermitFeeInvoicedAmount = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetPermitFeeName( iPermitFeeId )
'-------------------------------------------------------------------------------------------------
Function GetPermitFeeName( ByVal iPermitFeeId )
	Dim sSql, oRs, sFeeName

	sSql = "SELECT permitfeeprefix, permitfee "
	sSql = sSql & " FROM egov_permitfees "
	sSql = sSql & " WHERE permitfeeid = " & iPermitFeeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("permitfeeprefix") <> "" Then
			sFeeName = oRs("permitfeeprefix") & " "
		End If 
		sFeeName = sFeeName & oRs("permitfee")
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetPermitFeeName = sFeeName

End Function 


'-------------------------------------------------------------------------------------------------
' Function GetPermitFeeMethodById( iPermitFeeMethodId )
'-------------------------------------------------------------------------------------------------
Function GetPermitFeeMethodById( iPermitFeeMethodId )
	Dim sSql, oRs

	sSql = "SELECT permitfeemethod FROM egov_permitfeemethods WHERE permitfeemethodid = " & iPermitFeeMethodId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitFeeMethodById = oRs("permitfeemethod")
	Else
		GetPermitFeeMethodById = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' Function GetPermitFeeMethodIdByType( sTypeFlag )
'-------------------------------------------------------------------------------------------------
Function GetPermitFeeMethodIdByType( sTypeFlag )
	Dim sSql, oRs

	sSql = "SELECT permitfeemethodid FROM egov_permitfeemethods WHERE " & sTypeFlag & " = 1 AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitFeeMethodIdByType = oRs("permitfeemethodid")
	Else
		GetPermitFeeMethodIdByType = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double GetPermitFees( iPermitId, bIsOld, dSubTotal, sEndDate, sStartDate )
'--------------------------------------------------------------------------------------------------
Function GetPermitFees( ByVal iPermitId, ByVal bIsOld, ByRef dSubTotal, ByVal sEndDate, ByVal sStartDate )
	Dim sSql, oRs, sCompare, sPermitFees

	If bIsOld Then
		sSql = "SELECT ISNULL(SUM(II.invoicedamount),0.00) AS invoicedamount "
		sSql = sSql & " FROM egov_permitinvoiceitems II, egov_permitfeecategorytypes C, egov_permitinvoices I, egov_permits P "
		sSql = sSql & " WHERE II.permitid = " & iPermitId
		sSql = sSql & " AND I.permitid = P.permitid AND I.invoiceid = II.invoiceid "
		sSql = sSql & " AND II.permitfeecategorytypeid = C.permitfeecategorytypeid AND C.isgeneralbuildingtype = 1 AND "
		sSql = sSql & " I.invoicedate > '" & sStartDate & "' AND I.invoicedate < '" & sEndDate & "' AND II.feereportingtypeid IS NULL "
		sSql = sSql & " AND I.paymentid IS NOT NULL AND I.isvoided = 0 AND I.allfeeswaived = 0"
	Else
		sSql = "SELECT ISNULL(SUM(II.invoicedamount),0.00) AS invoicedamount "
		sSql = sSql & " FROM egov_permitinvoiceitems II, egov_permitfeecategorytypes C, egov_permitinvoices I, egov_permits P "
		sSql = sSql & " WHERE II.permitid = " & iPermitId
		sSql = sSql & " AND I.permitid = P.permitid AND I.invoiceid = II.invoiceid "
		sSql = sSql & " AND II.permitfeecategorytypeid = C.permitfeecategorytypeid AND C.isgeneralbuildingtype = 1 AND "
		sSql = sSql & " I.invoicedate < '" & sEndDate & "' AND II.feereportingtypeid IS NULL "
		sSql = sSql & " AND I.paymentid IS NOT NULL AND I.isvoided = 0 AND I.allfeeswaived = 0"
	End If 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sPermitFees = FormatNumber(oRs("invoicedamount"),2)
	Else
		sPermitFees = FormatNumber(0.00,2)
	End If 

	dSubTotal = dSubTotal + CDbl(sPermitFees)

	oRs.Close
	Set oRs = Nothing 

	GetPermitFees = sPermitFees

End Function 


'--------------------------------------------------------------------------------------------------
' double GetPermitFeeTypeTotal( iPermitId, sReportingType )
'--------------------------------------------------------------------------------------------------
Function GetPermitFeeTypeTotal( ByVal iPermitId, ByVal sReportingType )
	Dim sSql, oRs

'	sSql = "SELECT ISNULL(SUM(II.invoicedamount),0.00) AS feetotal "
'	sSql = sSql & " FROM egov_permitinvoiceitems II, egov_permitfeereportingtypes R, egov_permitinvoices I, egov_permits P "
'	sSql = sSql & " WHERE II.permitid = " & iPermitId
'	sSql = sSql & " AND I.permitid = P.permitid AND I.invoiceid = II.invoiceid "
'	sSql = sSql & " AND R.feereportingtypeid = II.feereportingtypeid AND R." & sReportingType & " = 1 "
'	sSql = sSql & " AND I.isvoided = 0 AND I.allfeeswaived = 0"

	' New way to get the fee total by type as of 1/21/2009 - SJL
	sSql = "SELECT ISNULL(SUM(F.feeamount),0.00) AS feetotal "
	sSql = sSql & " FROM egov_permitfees F, egov_permitfeereportingtypes R, egov_permits P "
	sSql = sSql & " WHERE P.permitid = " & iPermitId & " AND P.permitid = F.permitid AND "
	sSql = sSql & " R.feereportingtypeid = F.feereportingtypeid AND R." & sReportingType & " = 1 "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitFeeTypeTotal = FormatNumber(oRs("feetotal"),2)
	Else
		GetPermitFeeTypeTotal = FormatNumber(0.00,2)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' date GetPermitFinalInspectionDate( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitFinalInspectionDate( ByVal iPermitId )
	Dim sSql, oRs, dFinalInspectionDate

	sSql = "SELECT I.inspecteddate FROM egov_permitinspections I, egov_inspectionstatuses S "
	sSql = sSql & "WHERE I.inspectionstatusid = S.inspectionstatusid AND I.isfinal = 1 AND "
	sSql = sSql & "S.ispassed = 1 AND I.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If IsNull(oRs("inspecteddate")) Then 
			dFinalInspectionDate = ""
		Else 
			dFinalInspectionDate = FormatDateTime(oRs("inspecteddate"),2)
		End If 
	Else
		' No final inspection date so pull the complete date if there is one, otherwise it returns ""
		dFinalInspectionDate = GetPermitDate( "completeddate", iPermitId )
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetPermitFinalInspectionDate = dFinalInspectionDate

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetPermitIdByInvoiceId( iInvoiceId )
'-------------------------------------------------------------------------------------------------
Function GetPermitIdByInvoiceId( ByVal iInvoiceId )
	Dim sSql, oRs

	sSql = "SELECT permitid FROM egov_permitinvoices WHERE invoiceid = " & iInvoiceId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitIdByInvoiceId = oRs("permitid")
	Else
		GetPermitIdByInvoiceId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetPermitIdByPermitFeeId( iPermitFeeId )
'-------------------------------------------------------------------------------------------------
Function GetPermitIdByPermitFeeId( ByVal iPermitFeeId )
	Dim sSql, oRs

	sSql = "SELECT permitid FROM egov_permitfees WHERE permitfeeid = " & iPermitFeeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitIdByPermitFeeId = oRs("permitid")
	Else
		GetPermitIdByPermitFeeId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetPermitIdByPermitInspectionId( iPermitInspectionId )
'-------------------------------------------------------------------------------------------------
Function GetPermitIdByPermitInspectionId( ByVal iPermitInspectionId )
	Dim sSql, oRs

	sSql = "SELECT permitid FROM egov_permitinspections WHERE permitinspectionid = " & iPermitInspectionId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitIdByPermitInspectionId = CLng(oRs("permitid"))
	Else
		GetPermitIdByPermitInspectionId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetPermitIdByPermitReviewId( iPermitReviewId )
'-------------------------------------------------------------------------------------------------
Function GetPermitIdByPermitReviewId( ByVal iPermitReviewId )
	Dim sSql, oRs

	sSql = "SELECT permitid FROM egov_permitreviews WHERE permitreviewid = " & iPermitReviewId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitIdByPermitReviewId = CLng(oRs("permitid"))
	Else
		GetPermitIdByPermitReviewId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitInspectionNotes( iPermitInspectionId )
'--------------------------------------------------------------------------------------------------
Function GetPermitInspectionNotes( ByVal iPermitInspectionId )
	Dim sSql, oRs, sNotes 

	sNotes = ""

	sSql = "SELECT externalcomment, entrydate FROM egov_permitlog "
	sSql = sSql & " WHERE externalcomment IS NOT NULL AND permitinspectionid = " & iPermitInspectionId
	sSql = sSql & " ORDER BY entrydate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		If sNotes <> "" Then
			sNotes = sNotes & "<br /> <br />"
		End If 
		sNotes = sNotes & Trim(oRs("externalcomment"))
		oRs.MoveNext
	Loop
	
	oRs.Close 
	Set oRs = Nothing 

	GetPermitInspectionNotes = sNotes 

End Function 


'--------------------------------------------------------------------------------------------------
' void GetPermitInspectionsList( iPermitId, sInspectionsDue, sInspectorList )
'--------------------------------------------------------------------------------------------------
Sub GetPermitInspectionsList( ByVal iPermitId, ByRef sInspectionsDue, ByRef sInspectorList )
	Dim sSql, oRs, sPhone, iLines

	sInspectionsDue = ""
	sInspectorList = ""

	sSql = "SELECT I.permitinspectionid, I.permitinspectiontype, I.inspectiondescription, I.isrequired, S.inspectionstatus, "
	sSql = sSql & " I.inspecteddate, I.scheduleddate, I.isreinspection, ISNULL(I.inspectoruserid,0) AS inspectoruserid, isfinal "
	sSql = sSql & " FROM egov_permitinspections I, egov_inspectionstatuses S "
	sSql = sSql & " WHERE I.inspectionstatusid = S.inspectionstatusid AND I.permitid = " & iPermitId
	sSql = sSql & " ORDER BY I.inspectionorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		iLines = 0
		If sInspectionsDue <> "" Then 
			sInspectionsDue = sInspectionsDue & "<br /><br />"
		End If 
		sInspectionsDue = sInspectionsDue & oRs("permitinspectiontype") & " - " & oRs("inspectiondescription") 

		iLines = clng((Len(sInspectionsDue) / 89) + .5)

		If sInspectorList <> "" Then 
			sInspectorList = sInspectorList & "<br />"
		End If 
		If CLng(oRs("inspectoruserid")) > CLng(0) Then 
			sPhone = Trim(GetAdminPhone( CLng(oRs("inspectoruserid")) ))
			sInspectorList = sInspectorList & GetAdminName( CLng(oRs("inspectoruserid")) )
			If sPhone <> "" Then
				sInspectorList = sInspectorList & "   " & sPhone
			End If 
		End If 
		For x = 1 To iLines
			sInspectorList = sInspectorList & "<br />"
		Next 
		oRs.MoveNext
	Loop

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void GetPermitInspectionsPDF iPermitid, sInspectionList, sInspectionStatus, sInspectionDate, sInspector, sInspectionNotes 
'--------------------------------------------------------------------------------------------------
Sub GetPermitInspectionsPDF( ByVal iPermitid, ByRef sInspectionList, ByRef sInspectionStatus, ByRef sInspectionDate, ByRef sInspector, ByRef sInspectionNotes )
	Dim sSql, oRs, sPhone, iLines, sLastNote, iLoopCount

	sInspectionList = ""
	sInspectionStatus = ""
	sInspectionDate = ""
	sInspector = ""
	sInspectionNotes = ""
	iLoopCount = 0

	sSql = "SELECT I.permitinspectionid, I.permitinspectiontype, I.inspectiondescription, I.isrequired, S.inspectionstatus, "
	sSql = sSql & " I.inspecteddate, I.scheduleddate, I.isreinspection, ISNULL(I.inspectoruserid,0) AS inspectoruserid, S.shownotes "
	sSql = sSql & " FROM egov_permitinspections I, egov_inspectionstatuses S "
	sSql = sSql & " WHERE I.inspectionstatusid = S.inspectionstatusid AND I.permitid = " & iPermitId
	sSql = sSql & " ORDER BY I.inspectionorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		iLines = 0
		iLoopCount = iLoopCount + 1
		sLastNote = ""
		If sInspectionList <> "" Then 
			' Add a blank line between inspections
			sInspectionList = sInspectionList & "<br /><br />"
		End If 
		sNewInspection = oRs("permitinspectiontype") & " - " & Trim(oRs("inspectiondescription"))
		sInspectionList = sInspectionList &  sNewInspection

		If oRs("shownotes") Then
			sInspectionList = sInspectionList & "<br />      Notes:"
		Else 
			If Len(sNewInspection) < 68 Then
				' If the inspection is less than 1 line force to 2 lines (2 is max??)
				sInspectionList = sInspectionList & "<br />"
			End If 
		End If 
		If Len(sNewInspection) < 68 Then
			iListLine = 1 
		Else
			iListLine = 2 
		End If 

		If sInspectionStatus <> "" Then 
			' Add a blank line between inspections
			sInspectionStatus = sInspectionStatus & "<br /><br />"
		End If 
		sInspectionStatus = sInspectionStatus & oRs("inspectionstatus") & "<br />"

		If sInspectionDate <> "" Then 
			' Add a blank line between inspections
			sInspectionDate = sInspectionDate & "<br /><br />"
		End If 
		If IsNull(oRs("inspecteddate")) Then
			sInspectionDate = sInspectionDate & "<br />"
		Else
			sInspectionDate = sInspectionDate & FormatDateTime(oRs("inspecteddate"),2) & "<br />"
		End If 

		If sInspector <> "" Then 
			' add the blank line between inspections
			sInspector = sInspector & "<br /><br />"
		End If 
		If CLng(oRs("inspectoruserid")) > CLng(0) Then 
			sPhone = Trim(GetAdminPhone( CLng(oRs("inspectoruserid")) ))
			sInspector = sInspector & GetAdminName( CLng(oRs("inspectoruserid")) )
			sInspector = sInspector & "<br />"
			If sPhone <> "" Then
				sInspector = sInspector & sPhone
			Else 
				sInspector = sInspector & "No Phone"
			End If 
		Else
			' No assigned inspector so put lines for the name and phone
			sInspector = sInspector & "Unassigned<br />No Phone"
		End If 

		' Notes
		If oRs("shownotes") Then 
			sLastNote = GetLastPermitInspectionNote( oRs("permitinspectionid") )
			If sLastNote <> "" Then 
				iLines = clng((Len(sLastNote) / 55) + .5)
				' blank lines to preceed the notes
				For x = 1 To iListLine
					sInspectionNotes = sInspectionNotes & "<br />"
				Next 
				sInspectionNotes = sInspectionNotes & sLastNote & "<br /><br />"
			End If 
		Else 
			sInspectionNotes = sInspectionNotes & "<br /><br /><br />"
			iLines = 0
		End If 

		If iLines > 1 Then
			For x = 2 To iLines 
				sInspectionList = sInspectionList & "<br />"
			Next 
		End If 
		If iListLine = 1 Then 
			If iLines > 1 Then
				For x = 2 To iLines 
					sInspector = sInspector & "<br />" '& iLines
					sInspectionStatus = sInspectionStatus & "<br />" '& iLines
					sInspectionDate = sInspectionDate & "<br />"
				Next 
			End If 
		Else
			If iLines > 0 Then
				For x = 1 To iLines 
					sInspector = sInspector & "<br />" '& iLines
					sInspectionStatus = sInspectionStatus & "<br />" '& iLines
					sInspectionDate = sInspectionDate & "<br />"
				Next 
			End If 
		End If 

		oRs.MoveNext
	Loop

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void GetPermitInspectionsXML iPermitid, sInspection1, sInspectionStatus1, sInspectionDate1, sInspector1, sInspection2, sInspectionStatus2, sInspectionDate2, sInspector2, sInspection3, sInspectionStatus3, sInspectionDate3, sInspector3, sInspection4, sInspectionStatus4, sInspectionDate4, sInspector4, sInspection5, sInspectionStatus5, sInspectionDate5, sInspector5, sInspection6, sInspectionStatus6, sInspectionDate6, sInspector6, sInspectionNotes
'--------------------------------------------------------------------------------------------------
Sub GetPermitInspectionsXML( ByVal iPermitid, ByRef sInspection1, ByRef sInspectionStatus1, ByRef sInspectionDate1, ByRef sInspector1, ByRef sInspection2, ByRef sInspectionStatus2, ByRef sInspectionDate2, ByRef sInspector2, ByRef sInspection3, ByRef sInspectionStatus3, ByRef sInspectionDate3, ByRef sInspector3, ByRef sInspection4, ByRef sInspectionStatus4, ByRef sInspectionDate4, ByRef sInspector4, ByRef sInspection5, ByRef sInspectionStatus5, ByRef sInspectionDate5, ByRef sInspector5, ByRef sInspection6, ByRef sInspectionStatus6, ByRef sInspectionDate6, ByRef sInspector6, ByRef sInspectionNotes )
	Dim sSql, oRs, sLastNote, iLoopCount

	sInspection1 = ""
	sInspectionStatus1 = ""
	sInspectionDate1 = ""
	sInspector1 = ""
	sInspection2 = ""
	sInspectionStatus2 = ""
	sInspectionDate2 = ""
	sInspector2 = ""
	sInspection3 = ""
	sInspectionStatus3 = ""
	sInspectionDate3 = ""
	sInspector3 = ""
	sInspection4 = ""
	sInspectionStatus4 = ""
	sInspectionDate4 = ""
	sInspector4 = ""
	sInspection5 = ""
	sInspectionStatus5 = ""
	sInspectionDate5 = ""
	sInspector5 = ""
	sInspection6 = ""
	sInspectionStatus6 = ""
	sInspectionDate6 = ""
	sInspector6 = ""
	sInspectionNotes = ""
	iLoopCount = 0
	sLastNote = ""

	sSql = "SELECT I.permitinspectionid, I.permitinspectiontype, I.inspectiondescription, I.isrequired, S.inspectionstatus, "
	sSql = sSql & " I.inspecteddate, I.scheduleddate, I.isreinspection, ISNULL(I.inspectoruserid,0) AS inspectoruserid, S.shownotes "
	sSql = sSql & " FROM egov_permitinspections I, egov_inspectionstatuses S "
	sSql = sSql & " WHERE I.inspectionstatusid = S.inspectionstatusid AND I.permitid = " & iPermitId
	sSql = sSql & " ORDER BY I.inspectionorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		iLoopCount = iLoopCount + 1

		Select Case iLoopCount

			Case 1
				SetInspectionValues sInspection1, sInspectionStatus1, sInspectionDate1, sInspector1, oRs("permitinspectiontype"), oRs("inspectiondescription"), oRs("inspectionstatus"), oRs("inspecteddate"), oRs("inspectoruserid")

			Case 2
				SetInspectionValues sInspection2, sInspectionStatus2, sInspectionDate2, sInspector2, oRs("permitinspectiontype"), oRs("inspectiondescription"), oRs("inspectionstatus"), oRs("inspecteddate"), oRs("inspectoruserid")
			
			Case 3
				SetInspectionValues sInspection3, sInspectionStatus3, sInspectionDate3, sInspector3, oRs("permitinspectiontype"), oRs("inspectiondescription"), oRs("inspectionstatus"), oRs("inspecteddate"), oRs("inspectoruserid")
			
			Case 4
				SetInspectionValues sInspection4, sInspectionStatus4, sInspectionDate4, sInspector4, oRs("permitinspectiontype"), oRs("inspectiondescription"), oRs("inspectionstatus"), oRs("inspecteddate"), oRs("inspectoruserid")
			
			Case 5
				SetInspectionValues sInspection5, sInspectionStatus5, sInspectionDate5, sInspector5, oRs("permitinspectiontype"), oRs("inspectiondescription"), oRs("inspectionstatus"), oRs("inspecteddate"), oRs("inspectoruserid")
			
			Case 6
				SetInspectionValues sInspection6, sInspectionStatus6, sInspectionDate6, sInspector6, oRs("permitinspectiontype"), oRs("inspectiondescription"), oRs("inspectionstatus"), oRs("inspecteddate"), oRs("inspectoruserid")

		End Select  

		sLastNote = GetPermitInspectionNotes( oRs("permitinspectionid") )
		If sLastNote <> "" Then 
			If sInspectionNotes <> "" Then 
				sInspectionNotes = sInspectionNotes & "<br /><br />"
			End If 
			sInspectionNotes = sInspectionNotes &  oRs("permitinspectiontype") & " Notes:<br />" & sLastNote & " <br />"
		End If 

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' string GetPermitInvoicesStatus( iPermitId, bIsDue )
'--------------------------------------------------------------------------------------------------
Function GetPermitInvoicesStatus( ByVal iPermitId, ByRef bIsDue )
	Dim sSql, oRs

	sSql = "SELECT COUNT(S.invoicestatus) AS hits "
	sSql = sSql & " FROM egov_permitinvoices I, egov_invoicestatuses S "
	sSql = sSql & " WHERE I.invoicestatusid = S.invoicestatusid AND I.orgid = " & session("orgid")
	sSql = sSql & " AND S.isdue = 1 AND I.permitid = " & iPermitId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			GetPermitInvoicesStatus = "DUE INVOICE"
			bIsDue = True 
		Else 
			GetPermitInvoicesStatus = GetPermitPaidStatus( iPermitId )
			bIsDue = False 
		End If 
	Else
		GetPermitInvoicesStatus = ""
		bIsDue = False  
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' boolean GetPermitIsCompleted( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitIsCompleted( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT S.iscompletedstatus FROM egov_permits P, egov_permitstatuses S "
	sSql = sSql & " WHERE P.permitstatusid = S.permitstatusid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("iscompletedstatus") Then 
			GetPermitIsCompleted = True 
		Else
			GetPermitIsCompleted = False 
		End If 
	Else
		GetPermitIsCompleted = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean  GetPermitIsExpired( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitIsExpired( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT isexpired FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("isexpired") Then 
			GetPermitIsExpired = True 
		Else
			GetPermitIsExpired = False 
		End If 
	Else
		GetPermitIsExpired = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean GetPermitIsIssued( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitIsIssued( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT S.isissued, S.iscompletedstatus FROM egov_permits P, egov_permitstatuses S "
	sSql = sSql & " WHERE P.permitstatusid = S.permitstatusid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("isissued") Or oRs("iscompletedstatus") Then 
			GetPermitIsIssued = True 
		Else
			GetPermitIsIssued = False 
		End If 
	Else
		GetPermitIsIssued = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean GetPermitIsOnHold( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitIsOnHold( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT isonhold FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("isonhold") Then 
			GetPermitIsOnHold = True 
		Else
			GetPermitIsOnHold = False 
		End If 
	Else
		GetPermitIsOnHold = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' string GetPermitIssuedBy( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitIssuedBy( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT LEFT(U.firstname,1) + '. ' + U.lastname AS username "
	sSql = sSql & "FROM egov_permitlog L, users U "
	sSql = sSql & "WHERE L.permitid = " & iPermitId
	sSql = sSql & " AND L.adminuserid = U.userid AND L.orgid = U.orgid AND "
	sSql = sSql & "L.activitycomment LIKE 'permit status changed from approved to issued'"
	

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitIssuedBy = oRs("username")
	Else
		GetPermitIssuedBy = ""
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitIssuedDate( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitIssuedDate( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT issueddate FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If IsNull(oRs("issueddate")) Then 
			GetPermitIssuedDate = "" 
		Else
			GetPermitIssuedDate = FormatDateTime(oRs("issueddate"),2) 
		End If 
	Else
		GetPermitIssuedDate = "" 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean GetPermitIsVoided( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitIsVoided( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT isvoided FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("isvoided") Then 
			GetPermitIsVoided = True 
		Else
			GetPermitIsVoided = False 
		End If 
	Else
		GetPermitIsVoided = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitJobSite( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitJobSite( ByVal iPermitId )
	Dim sSql, oRs, sJobSite

	sSql = "SELECT permitaddressid, residentstreetnumber, ISNULL(residentstreetprefix,'') AS residentstreetprefix, "
	sSql = sSql & " residentstreetname, ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection, "
	sSql = sSql & " ISNULL(residentunit,'') AS residentunit, ISNULL(residentcity,'') AS residentcity, "
	sSql = sSql & " ISNULL(residentstate,'') AS residentstate "
	sSql = sSql & " FROM egov_permitaddress WHERE permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sJobSite = oRs("residentstreetnumber")
		If oRs("residentstreetprefix") <> "" Then
			sJobSite = sJobSite & " " & oRs("residentstreetprefix")
		End If
		sJobSite = sJobSite & " " & oRs("residentstreetname")
		If oRs("streetsuffix") <> "" Then
			sJobSite = sJobSite & " " & oRs("streetsuffix")
		End If
		If oRs("streetdirection") <> "" Then
			sJobSite = sJobSite & " " & oRs("streetdirection")
		End If
		If oRs("residentunit") <> "" Then
			sJobSite = sJobSite & ", " & oRs("residentunit")
		End If
		
		If oRs("residentcity") <> "" Then
			sJobSite = sJobSite & ", " & oRs("residentcity")
		End If 
		If oRs("residentstate") <> "" Then
			sJobSite = sJobSite & ", " & oRs("residentstate")
		End If 
	Else 
		sJobSite = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetPermitJobSite = sJobSite

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitJobSitePIN( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitJobSitePIN( ByVal iPermitId )
	Dim sSql, oRs, sJobSite

	sSql = "SELECT ISNULL(parcelidnumber,'') AS parcelidnumber FROM egov_permitaddress WHERE permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitJobSitePIN = oRs("parcelidnumber")
	Else
		GetPermitJobSitePIN = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitJobSitePIN( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitListedOwner( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(listedowner,'') AS listedowner FROM egov_permitaddress WHERE permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitListedOwner = oRs("listedowner")
	Else
		GetPermitListedOwner = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitLocation( iPermitId, sLegalDescription, sListedOwner, iPermitAddressId, sCounty, sParcelid, bBreakOnCity )
'--------------------------------------------------------------------------------------------------
Function GetPermitLocation( ByVal iPermitId, ByRef sLegalDescription, ByRef sListedOwner, ByRef iPermitAddressId, ByRef sCounty, ByRef sParcelid, ByVal bBreakOnCity )
	Dim sSql, oRs

	sSql = "SELECT permitaddressid, residentstreetnumber, ISNULL(residentstreetprefix,'') AS residentstreetprefix, "
	sSql = sSql & " residentstreetname, ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection, "
	sSql = sSql & " ISNULL(residentunit,'') AS residentunit, ISNULL(residentcity,'') AS residentcity, "
	sSql = sSql & " ISNULL(residentstate,'') AS residentstate, ISNULL(legaldescription,'') AS legaldescription, "
	sSql = sSql & " ISNULL(listedowner,'') AS listedowner, ISNULL(residentzip,'') AS residentzip, ISNULL(county,'') AS county, "
	sSql = sSql & " ISNULL(parcelidnumber,'') AS parcelid "
	sSql = sSql & " FROM egov_permitaddress WHERE permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		strStreetNumber = oRs("residentstreetnumber")
		GetPermitLocation = oRs("residentstreetnumber")
		If oRs("residentstreetprefix") <> "" Then
			GetPermitLocation = GetPermitLocation & " " & oRs("residentstreetprefix")
		End If
		strStreetName = oRs("residentstreetname")
		GetPermitLocation = GetPermitLocation & " " & oRs("residentstreetname")
		If oRs("streetsuffix") <> "" Then
			GetPermitLocation = GetPermitLocation & " " & oRs("streetsuffix")
		End If
		If oRs("streetdirection") <> "" Then
			GetPermitLocation = GetPermitLocation & " " & oRs("streetdirection")
		End If
		If oRs("residentunit") <> "" Then
			GetPermitLocation = GetPermitLocation & ", " & oRs("residentunit")
		End If
		If bBreakOnCity Then
			GetPermitLocation = GetPermitLocation & "<br />"
		End If 
		
		If oRs("residentcity") <> "" Then
			If bBreakOnCity Then
				GetPermitLocation = GetPermitLocation & oRs("residentcity")
			Else 
				GetPermitLocation = GetPermitLocation & ", " & oRs("residentcity")
			End If 
		End If 
		If oRs("residentstate") <> "" Then
			GetPermitLocation = GetPermitLocation & ", " & oRs("residentstate")
		End If 
		If oRs("residentzip") <> "" Then 
			GetPermitLocation = GetPermitLocation & " " & oRs("residentzip")
		End If 
		sLegalDescription = Trim(oRs("legaldescription"))
		sListedOwner = Trim(oRs("listedowner"))
		iPermitAddressId = oRs("permitaddressid")
		sCounty =  oRs("county")
		sParcelid =  oRs("parcelid")
	Else 
		GetPermitLocation = ""
		sLegalDescription = ""
		sListedOwner = ""
		sCounty = ""
		sParcelid = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitLocationType( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitLocationType( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT L.locationtype FROM egov_permits P, egov_permitlocationrequirements L "
	sSql = sSql & "WHERE P.permitlocationrequirementid = L.permitlocationrequirementid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitLocationType = oRs("locationtype")
	Else 
		GetPermitLocationType = "none"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitNotes( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitNotes( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT permitnotes FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitNotes = oRs("permitnotes") 
	Else
		GetPermitNotes = ""
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitNumber( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitNumber( ByVal iPermitId )
	Dim sSql, oRs, sPermitNumberYear, sPermitNumberPrefix, sPermitNumber, sFormatedNumber

	sPermitNumberYear = ""
	sPermitNumberPrefix = ""
	sPermitNumber = "0"
	sFormatedNumber = ""

	sSql = "SELECT ISNULL(permitnumber,0) AS permitnumber, permitnumberyear, permitnumberprefix FROM egov_permits WHERE permitid = " & iPermitId
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		'GetPermitNumber = oRs("permitnumberyear") & oRs("permitnumberprefix") & oRs("permitnumber")
		sPermitNumberYear = oRs("permitnumberyear")
		sPermitNumberPrefix = oRs("permitnumberprefix")
		sPermitNumber = Trim(oRs("permitnumber"))
	End If 

	oRs.CLose
	Set oRs = Nothing 

	If CLng(sPermitNumber) > CLng(0) Then 
		' Now get the permit number format
		sSql = "SELECT element, characters FROM egov_permitnumberformat WHERE isforbuildingpermits = 1 AND orgid = " & session("orgid") & " ORDER BY position"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then 
			Do While Not oRs.EOF 
				Select Case oRs("element")
					Case "year"
						sFormatedNumber = sFormatedNumber & Right(sPermitNumberYear, clng(oRs("characters")))
					Case "dash"
						sFormatedNumber = sFormatedNumber & "-"
					Case "prefix"
						If PermitNumberPrefixIsNotNone( sPermitNumberPrefix ) Then 
							sFormatedNumber = sFormatedNumber & sPermitNumberPrefix
						End If 
					Case "space"
						sFormatedNumber = sFormatedNumber & Space(clng(oRs("characters")))
					Case "sequence"
						If clng(Len(sPermitNumber)) < clng(oRs("characters")) Then 
							sFormatedNumber = sFormatedNumber & Replace(Space(clng(oRs("characters")) - Len(sPermitNumber))," ","0") & sPermitNumber
						Else
							sFormatedNumber = sFormatedNumber & sPermitNumber
						End If 
				End Select 
				oRs.MoveNext 
			Loop 
		End If 

		oRs.CLose
		Set oRs = Nothing 
	End If 

	GetPermitNumber = sFormatedNumber

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitOccupancyType( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitOccupancyType( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(occupancytype,'') AS occupancytype FROM egov_occupancytypes O, egov_permits P "
	sSql = sSql & " WHERE P.occupancytypeid = O.occupancytypeid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitOccupancyType = oRs("occupancytype")
	Else
		GetPermitOccupancyType = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitOccupancyTypeGroup( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitOccupancyTypeGroup( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(usegroupcode,'') AS usegroupcode FROM egov_occupancytypes O, egov_permits P "
	sSql = sSql & " WHERE P.occupancytypeid = O.occupancytypeid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitOccupancyTypeGroup = oRs("usegroupcode")
	Else
		GetPermitOccupancyTypeGroup = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitPaidStatus( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitPaidStatus( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(S.invoicestatus) AS hits "
	sSql = sSql & " FROM egov_permitinvoices I, egov_invoicestatuses S "
	sSql = sSql & " WHERE I.invoicestatusid = S.invoicestatusid AND I.orgid = " & session("orgid")
	sSql = sSql & " AND S.ispaid = 1 AND I.permitid = " & iPermitId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			GetPermitPaidStatus = "PAID INVOICE"
		Else 
			GetPermitPaidStatus = ""
		End If 
	Else
		GetPermitPaidStatus = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' double GetPermitPaymentTotal( iPaymentid )
'--------------------------------------------------------------------------------------------------
Function GetPermitPaymentTotal( ByVal iPaymentid )
	Dim sSql, oRs

	sSql = "SELECT SUM(paymenttotal) AS paymenttotal FROM egov_class_payment WHERE paymentid = " & iPaymentid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitPaymentTotal = FormatNumber(oRs("paymenttotal"),2,,,0)
	Else
		GetPermitPaymentTotal = 0.00
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitPermitLocation( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitPermitLocation( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(permitlocation,'') AS permitlocation FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitPermitLocation = oRs("permitlocation")
	Else
		GetPermitPermitLocation = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean GetPermitPermitTypeFlag( iPermitid, sFlag )
'--------------------------------------------------------------------------------------------------
Function GetPermitPermitTypeFlag( ByVal iPermitid, ByVal sFlag )
	Dim sSql, oRs

	sSql = "SELECT " & sFlag & " AS isflagged FROM egov_permitpermittypes WHERE permitid = " & iPermitid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If oRs("isflagged") Then 
			GetPermitPermitTypeFlag = True 
		Else
			GetPermitPermitTypeFlag = False  
		End If 
	Else
		GetPermitPermitTypeFlag = False  
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetPermitPermitTypeValue( iPermitid, sField )
'--------------------------------------------------------------------------------------------------
Function GetPermitPermitTypeValue( ByVal iPermitid, ByVal sField )
	Dim sSql, oRs

	sSql = "SELECT " & sField & " AS fieldvalue FROM egov_permitpermittypes WHERE permitid = " & iPermitid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If oRs("fieldvalue") <> "" And Not IsNull(oRs("fieldvalue")) Then 
			GetPermitPermitTypeValue = oRs("fieldvalue") 
		Else
			GetPermitPermitTypeValue = Null
		End If 
	Else
		GetPermitPermitTypeValue = Null
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitPlansBy( ByVal iPermitid )
'--------------------------------------------------------------------------------------------------
Function GetPermitPlansBy( ByVal iPermitid )
	Dim sSql, oRs, sResults

	sResults = ""

	sSql = "SELECT C.company, C.firstname, C.lastname, C.phone, ISNULL(C.address,'') AS address, "
	sSql = sSql & " ISNULL(C.city,'') AS city, ISNULL(C.state,'') AS state, ISNULL(C.zip,'') AS zip "
	sSql = sSql & " FROM egov_permitcontacts C, egov_permits P "
	sSql = sSql & " WHERE C.permitcontactid = P.plansbycontactid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("firstname") <> "" Then
			sResults = oRs("firstname") & " " & oRs("lastname")
		End If 
		If oRs("company") <> "" Then
			If oRs("firstname") <> "" Then 
				sResults = sResults &  "<br />" & oRs("company")
			Else
				sResults = oRs("company")
			End If 
		End If 

		If oRs("address") <> "" Then 
			sResults = sResults &  "<br />" & oRs("address")
		End If 
		If oRs("city") <> "" Then 
			sResults = sResults &  "<br />" & oRs("city")
			If oRs("state") <> "" Then 
				sResults = sResults & ", " & oRs("state") & " " & oRs("zip")
			End If 
		End If 

		If Not IsNull(oRs("phone")) And oRs("phone") <> "" Then
			sResults = sResults & "<br />" & FormatPhoneNumber( oRs("phone") )
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetPermitPlansBy = sResults

End Function 


'--------------------------------------------------------------------------------------------------
' double GetPermitReportingFees( iPermitId, bIsOld, sReportingType, dSubTotal )
'--------------------------------------------------------------------------------------------------
Function GetPermitReportingFees( ByVal iPermitId, ByVal bIsOld, ByVal sReportingType, ByRef dSubTotal, ByVal sEndDate, ByVal sStartDate )
	Dim sSql, oRs, sCompare, sReportingFees

	If bIsOld Then
		sSql = "SELECT ISNULL(SUM(II.invoicedamount),0.00) AS invoicedamount "
		sSql = sSql & " FROM egov_permitinvoiceitems II, egov_permitfeereportingtypes R, egov_permitinvoices I, egov_permits P "
		sSql = sSql & " WHERE II.permitid = " & iPermitId
		sSql = sSql & " AND I.permitid = P.permitid AND I.invoiceid = II.invoiceid "
		sSql = sSql & " AND R.feereportingtypeid = II.feereportingtypeid AND R." & sReportingType & " = 1 "
		sSql = sSql & " AND I.invoicedate > '" & sStartDate & "' AND I.invoicedate < '" & sEndDate & "' "
		sSql = sSql & " AND I.paymentid IS NOT NULL AND I.isvoided = 0 AND I.allfeeswaived = 0"
	Else
		sSql = "SELECT ISNULL(SUM(II.invoicedamount),0.00) AS invoicedamount "
		sSql = sSql & " FROM egov_permitinvoiceitems II, egov_permitfeereportingtypes R, egov_permitinvoices I, egov_permits P "
		sSql = sSql & " WHERE II.permitid = " & iPermitId
		sSql = sSql & " AND I.permitid = P.permitid AND I.invoiceid = II.invoiceid "
		sSql = sSql & " AND R.feereportingtypeid = II.feereportingtypeid AND R." & sReportingType & " = 1 "
		sSql = sSql & " AND I.invoicedate < '" & sEndDate & "' "
		sSql = sSql & " AND I.paymentid IS NOT NULL AND I.isvoided = 0 AND I.allfeeswaived = 0"
	End If 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sReportingFees = FormatNumber(oRs("invoicedamount"),2)
	Else
		sReportingFees = FormatNumber(0.00,2)
	End If 

	dSubTotal = dSubTotal + CDbl(sReportingFees)

	oRs.Close
	Set oRs = Nothing 

	GetPermitReportingFees = sReportingFees

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitReviewerName( iReviewerUserId )
'--------------------------------------------------------------------------------------------------
Function GetPermitReviewerName( ByVal iReviewerUserId )
	Dim sSql, oRs

	sSql = "SELECT firstname, lastname FROM users WHERE userid = " & iReviewerUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitReviewerName = oRs("firstname") & " " & oRs("lastname")
	Else
		GetPermitReviewerName = ""
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' void GetPermitReviewList iPermitid, sReviewList, sReviewStatus, sReviewDate, sReviewer, sReviewNotes
'-------------------------------------------------------------------------------------------------
Sub GetPermitReviewList( ByVal iPermitid, ByRef sReviewList, ByRef sReviewStatus, ByRef sReviewDate, ByRef sReviewer, ByRef sReviewNotes )
	Dim sSql, oRs, iNoteRows, x, sPhone

	sReviewList = ""
	sReviewStatus = ""
	sReviewDate = ""
	sReviewer = ""
	sReviewNotes = ""

	sSql = "SELECT R.permitreviewid, R.permitreviewtype, S.reviewstatus, S.shownotes, "
	sSql = sSql & " ISNULL(R.revieweruserid,0) AS revieweruserid, R.reviewed "
	sSql = sSql & " FROM egov_permitreviews R, egov_reviewstatuses S "
	sSql = sSql & " WHERE R.reviewstatusid = S.reviewstatusid AND R.permitid = " & iPermitId
	sSql = sSql & " ORDER BY R.reviewed"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		iNoteRows = 0
		iLines = 0
		iRows = 0
		sReviewList = sReviewList & oRs("permitreviewtype")
		If Not IsNull(oRs("reviewed")) Then 
			sReviewDate = sReviewDate & FormatDateTime(oRs("reviewed"),2)
		Else
			sReviewDate = sReviewDate & Space(10)
		End If 
		If CLng(oRs("revieweruserid")) > CLng(0) Then
			sReviewer = sReviewer & GetPermitReviewerName( CLng(oRs("revieweruserid")) )
			sPhone = Trim(GetAdminPhone( CLng(oRs("revieweruserid")) ))
			If sPhone <> "" Then
				sReviewer = sReviewer & "   " & sPhone
			End If 
		Else
			sReviewer = sReviewer & "Unassigned"
		End If 

		sReviewStatus = sReviewStatus & oRs("reviewstatus")
		sReviewStatus = sReviewStatus & "<br />"
		iNoteRows = iNoteRows + 1

		If oRs("shownotes") Then
			'sLastNote = GetLastPermitReviewNote( oRs("permitreviewid") )
			' Changed to get all notes 2/23/2010, Steve Loar
			sLastNote = GetPermitReviewNotes( oRs("permitreviewid") )
			If sLastNote <> "" Then 
				sReviewList = sReviewList & "<br />              Notes:"
				sReviewNotes = sReviewNotes & sLastNote & " <br />"
				iLines = clng((Len(sLastNote) / 47) + .5)
				iNoteRows = iNoteRows + iLines
			End If 
		Else
			sReviewList = sReviewList & "<br />"
		End If 

		If iNoteRows > 0 Then 
			' Add blank rows
			If iNoteRows > 1 Then 
				For x = 1 To iNoteRows - 1
					sReviewList = sReviewList & "<br />"
				Next 
			End If 
			For x = 1 To iNoteRows
				sReviewDate = sReviewDate & "<br />"
				sReviewer = sReviewer & "<br />"
			Next 
		End If 

		' One more to seperate the reviews
		sReviewList = sReviewList & "<br />"
		sReviewDate = sReviewDate & "<br />"
		sReviewer = sReviewer & "<br />"
		sReviewStatus = sReviewStatus & "<br />"
		sReviewNotes = sReviewNotes 
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void GetPermitReviewListForXML iPermitid, sReviewList1, sReviewStatus1, sReviewDate1, sReviewer1, sReviewList2, sReviewStatus2, sReviewDate2, sReviewer2, sReviewList3, sReviewStatus3, sReviewDate3, sReviewer3, sReviewNotes
'-------------------------------------------------------------------------------------------------
Sub GetPermitReviewListForXML( ByVal iPermitid, ByRef sReviewList1, ByRef sReviewStatus1, ByRef sReviewDate1, ByRef sReviewer1, _
	ByRef sReviewList2, ByRef sReviewStatus2, ByRef sReviewDate2, ByRef sReviewer2, ByRef sReviewList3, ByRef sReviewStatus3, _
	ByRef sReviewDate3, ByRef sReviewer3, ByRef sReviewList4, ByRef sReviewStatus4, ByRef sReviewDate4, ByRef sReviewer4, _
	ByRef sReviewList5, ByRef sReviewStatus5, ByRef sReviewDate5, ByRef sReviewer5, ByRef sReviewList6, ByRef sReviewStatus6, _
	ByRef sReviewDate6, ByRef sReviewer6, ByRef sReviewNotes )
	Dim sSql, oRs, iNoteRows, x, sPhone, iRecCount

	sReviewList1 = ""
	sReviewStatus1 = ""
	sReviewDate1 = ""
	sReviewer1 = ""
	sReviewList2 = ""
	sReviewStatus2 = ""
	sReviewDate2 = ""
	sReviewer2 = ""
	sReviewList3 = ""
	sReviewStatus3 = ""
	sReviewDate3 = ""
	sReviewer3 = ""
	sReviewList4 = ""
	sReviewStatus4 = ""
	sReviewDate4 = ""
	sReviewer4 = ""
	sReviewList5 = ""
	sReviewStatus5 = ""
	sReviewDate5 = ""
	sReviewer5 = ""
	sReviewList6 = ""
	sReviewStatus6 = ""
	sReviewDate6 = ""
	sReviewer6 = ""
	sReviewNotes = ""
	iRecCount = clng(0) 

	sSql = "SELECT R.permitreviewid, R.permitreviewtype, S.reviewstatus, S.shownotes, "
	sSql = sSql & " ISNULL(R.revieweruserid,0) AS revieweruserid, R.reviewed "
	sSql = sSql & " FROM egov_permitreviews R, egov_reviewstatuses S "
	sSql = sSql & " WHERE R.reviewstatusid = S.reviewstatusid AND R.permitid = " & iPermitId
	sSql = sSql & " ORDER BY R.reviewed"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		iRecCount = iRecCount + clng(1)
		If iRecCount > clng(6) Then
			Exit Do 
		End If 
		iNoteRows = 0
		iLines = 0
		iRows = 0
		Select Case iRecCount
			Case 1
				SetReviewValues sReviewList1, sReviewDate1, sReviewer1, sReviewStatus1, oRs("permitreviewtype"), oRs("reviewed"), oRs("revieweruserid"), oRs("reviewstatus")
				
			Case 2
				SetReviewValues sReviewList2, sReviewDate2, sReviewer2, sReviewStatus2, oRs("permitreviewtype"), oRs("reviewed"), oRs("revieweruserid"), oRs("reviewstatus")

			Case 3
				SetReviewValues sReviewList3, sReviewDate3, sReviewer3, sReviewStatus3, oRs("permitreviewtype"), oRs("reviewed"), oRs("revieweruserid"), oRs("reviewstatus")

			Case 4
				SetReviewValues sReviewList4, sReviewDate4, sReviewer4, sReviewStatus4, oRs("permitreviewtype"), oRs("reviewed"), oRs("revieweruserid"), oRs("reviewstatus")

			Case 5
				SetReviewValues sReviewList5, sReviewDate5, sReviewer5, sReviewStatus5, oRs("permitreviewtype"), oRs("reviewed"), oRs("revieweruserid"), oRs("reviewstatus")

			Case 6
				SetReviewValues sReviewList6, sReviewDate6, sReviewer6, sReviewStatus6, oRs("permitreviewtype"), oRs("reviewed"), oRs("revieweruserid"), oRs("reviewstatus")

		End Select  

		iNoteRows = iNoteRows + 1
'		If oRs("shownotes") Then
			'sLastNote = GetLastPermitReviewNote( oRs("permitreviewid") )
			' Changed to get all notes 2/23/2010, Steve Loar
			sLastNote = GetPermitReviewNotes( oRs("permitreviewid") )
			If sLastNote <> "" Then 
				sReviewNotes = sReviewNotes & "<br /><br />" & oRs("permitreviewtype") & " Notes and Conditions:<br />" & sLastNote & " <br />"
			End If 
'		End If 

		sReviewNotes = sReviewNotes 
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' string GetPermitReviewNotes( iPermitReviewId )
'--------------------------------------------------------------------------------------------------
Function GetPermitReviewNotes( ByVal iPermitReviewId )
	Dim sSql, oRs, sNotes 

	sNotes = ""

	sSql = "SELECT externalcomment, entrydate FROM egov_permitlog "
	sSql = sSql & " WHERE externalcomment IS NOT NULL AND permitreviewid = " & iPermitReviewId
	sSql = sSql & " ORDER BY entrydate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		If sNotes <> "" Then
			sNotes = sNotes & "<br /> <br />"
		End If 
		sNotes = sNotes & Trim(oRs("externalcomment"))
		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

	GetPermitReviewNotes = sNotes

End Function 


'-------------------------------------------------------------------------------------------------
' boolean GetPermitStatusBlockReview( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitStatusBlockReview( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT blockreviews FROM egov_permitstatuses S, egov_permits P "
	sSql = sSql & " WHERE S.permitstatusid = P.permitstatusid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("blockreviews") Then 
			GetPermitStatusBlockReview = True 
		Else
			GetPermitStatusBlockReview = False 
		End If 
	Else
		GetPermitStatusBlockReview = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetPermitStatusByPermitId( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitStatusByPermitId( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT S.permitstatus FROM egov_permitstatuses S, egov_permits P "
	sSql = sSql & " WHERE P.permitstatusid = S.permitstatusid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitStatusByPermitId = oRs("permitstatus")
	Else
		GetPermitStatusByPermitId = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetPermitStatusByStatusId( iPermitStatusId )
'-------------------------------------------------------------------------------------------------
Function GetPermitStatusByStatusId( ByVal iPermitStatusId )
	Dim sSql, oRs

	sSql = "SELECT permitstatus FROM egov_permitstatuses WHERE permitstatusid = " & iPermitStatusId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitStatusByStatusId = oRs("permitstatus")
	Else
		GetPermitStatusByStatusId = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetPermitStatusIdByStatusType( sType )
'-------------------------------------------------------------------------------------------------
Function GetPermitStatusIdByStatusType( ByVal sType )
	Dim sSql, oRs

	sSql = "SELECT permitstatusid FROM egov_permitstatuses WHERE orgid = " & session("orgid")
	sSql = sSql & " AND isforbuildingpermits = 1 AND " & sType & " = 1"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitStatusIdByStatusType = CLng(oRs("permitstatusid"))
	Else
		GetPermitStatusIdByStatusType = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 

'-------------------------------------------------------------------------------------------------
' integer GetPermitStatusIdByStatusName( sType )
'-------------------------------------------------------------------------------------------------
Function GetPermitStatusIdByStatusName( ByVal sType )
	Dim sSql, oRs

	sSql = "SELECT permitstatusid FROM egov_permitstatuses WHERE orgid = " & session("orgid")
	sSql = sSql & " AND permitstatus = '" & dbsafe(sType) & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitStatusIdByStatusName = CLng(oRs("permitstatusid"))
	Else
		GetPermitStatusIdByStatusName = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetPermitStatusId( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitStatusId( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT permitstatusid FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitStatusId = oRs("permitstatusid")
	Else
		GetPermitStatusId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetPermitStreetAddress( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitStreetAddress( ByVal iPermitId )
	Dim sSql, oRs, sAddress

	'sSql = "SELECT dbo.fn_buildAddress(residentstreetnumber, residentstreetprefix, residentstreetname, "
	'sSql = sSql & "streetsuffix, streetdirection ) AS permitaddress "

	' changed to pull in the unit/suite 7/15/2013, SL
	sSql = "SELECT ISNULL(residentstreetnumber,'') AS residentstreetnumber, ISNULL(residentstreetprefix,'') AS residentstreetprefix, "
	sSql = sSql & "ISNULL(residentstreetname,'') AS residentstreetname, ISNULL(streetsuffix,'') AS streetsuffix, "
	sSql = sSql & "ISNULL(streetdirection,'') AS streetdirection, ISNULL(residentunit,'') AS residentunit "
	sSql = sSql & "FROM egov_permitaddress WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		'GetPermitStreetAddress = oRs("permitaddress")

		sAddress = oRs("residentstreetnumber")
		If oRs("residentstreetprefix") <> "" Then 
			sAddress = sAddress & " " & oRs("residentstreetprefix")
		End If 
		If oRs("residentstreetname") <> "" Then 
			sAddress = sAddress & " " & oRs("residentstreetname")
		End If 
		If oRs("streetsuffix") <> "" Then 
			sAddress = sAddress & " " & oRs("streetsuffix")
		End If 
		If oRs("streetdirection") <> "" Then 
			sAddress = sAddress & " " & oRs("streetdirection")
		End If 
		If oRs("residentunit") <> "" Then 
			sAddress = sAddress & ", " & oRs("residentunit")
		End If 
		GetPermitStreetAddress = sAddress
	Else 
		GetPermitStreetAddress = ""
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitTitle( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitTitle( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT T.permittitle FROM egov_permittypes T, egov_permits P "
	sSql = sSql & " WHERE T.permittypeid = P.permittypeid AND P.permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitTitle = oRs("permittitle")
	Else 
		GetPermitTitle = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitTypeDesc( iPermitId, bIncludePrefix )
'--------------------------------------------------------------------------------------------------
Function GetPermitTypeDesc( ByVal iPermitId, ByVal bIncludePrefix )
	Dim sSql, oRs, sType

	sType = ""
	sSql = "SELECT ISNULL(permittypedesc,'') AS permittypedesc, ISNULL(permittype,'') AS permittype "
	sSql = sSql & "FROM egov_permitpermittypes WHERE permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If bIncludePrefix Then 
			sType = oRs("permittype")
		End If 
		If sType <> "" And oRs("permittypedesc") <> "" Then
			sType = sType & " &ndash; "
		End If 
		sType = sType & oRs("permittypedesc")
	Else 
		sType = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetPermitTypeDesc = sType

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitTypeDescForPDF( iPermitId, bIncludePrefix )
'--------------------------------------------------------------------------------------------------
Function GetPermitTypeDescForPDF( ByVal iPermitId, ByVal bIncludePrefix )
	Dim sSql, oRs, sType

	sType = ""
	sSql = "SELECT permittypedesc, permittype FROM egov_permitpermittypes "
	sSql = sSql & " WHERE permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If bIncludePrefix Then 
			sType = oRs("permittype") 
		End If 
		If sType <> "" And oRs("permittypedesc") <> "" Then
			sType = sType & " - "
		End If 
		sType = sType & oRs("permittypedesc")
	Else 
		sType = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetPermitTypeDescForPDF = sType

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetPermitTypeId( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitTypeId( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT permittypeid FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitTypeId = CLng(oRs("permittypeid"))
	Else 
		GetPermitTypeId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Function 


'-------------------------------------------------------------------------------------------------
' void GetPermitTypeLicenseDetails iLicenseTypeId, sLicenseType, iDisplayOrder 
'-------------------------------------------------------------------------------------------------
Sub GetPermitTypeLicenseDetails( ByVal iLicenseTypeId, ByRef sLicenseType, ByRef iDisplayOrder )
	Dim sSql, oRs

	sSql = "SELECT licensetype, displayorder FROM egov_permitlicensetypes WHERE licensetypeid = " & iLicenseTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sLicenseType = oRs("licensetype")
		iDisplayOrder = oRs("displayorder")
	Else
		sLicenseType = ""
		iDisplayOrder = 1
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' integer GetPermitTypeUseType( iPermitTypeId )
'-------------------------------------------------------------------------------------------------
Function GetPermitTypeUseType( ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(usetypeid,0) AS usetypeid FROM egov_permittypes WHERE permittypeid = " & iPermitTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitTypeUseType = CLng(oRs("usetypeid"))
	Else 
		GetPermitTypeUseType = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetPermitUseClass( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitUseClass( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT useclass FROM egov_permituseclasses U, egov_permits P "
	sSql = sSql & "WHERE P.useclassid = U.useclassid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitUseClass = oRs("useclass")
	Else 
		GetPermitUseClass = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetPermitUseType( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitUseType( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT usetype FROM egov_permitusetypes U, egov_permits P "
	sSql = sSql & "WHERE P.usetypeid = U.usetypeid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitUseType = oRs("usetype")
	Else 
		GetPermitUseType = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' void GetPermitValuesForInvoiceItems( iPermitFeeId, sPermitFeePrefix, sPermitFee, iPermitFeeCategoryTypeId, iFeeReportingTypeId, iIsPercentageTypeFee, iDisplayOrder )
'-------------------------------------------------------------------------------------------------
Sub GetPermitValuesForInvoiceItems( ByVal iPermitFeeId, ByRef sPermitFeePrefix, ByRef sPermitFee, ByRef iPermitFeeCategoryTypeId, ByRef iFeeReportingTypeId, ByRef iIsPercentageTypeFee, ByRef iDisplayOrder )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(permitfeeprefix,'') AS permitfeeprefix,  ISNULL(permitfee,'') AS permitfee, ispercentagetypefee, ISNULL(displayorder,0) AS displayorder, "
	sSql = sSql & " ISNULL(permitfeecategorytypeid, 0) AS permitfeecategorytypeid, ISNULL(feereportingtypeid,0) AS feereportingtypeid "
	sSql = sSql & " FROM egov_permitfees WHERE permitfeeid = " & iPermitFeeId
'	response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sPermitFeePrefix = oRs("permitfeeprefix")
		sPermitFee = oRs("permitfee")
		iPermitFeeCategoryTypeId = oRs("permitfeecategorytypeid")
		iFeeReportingTypeId = oRs("feereportingtypeid")
		If oRs("ispercentagetypefee") Then 
			iIsPercentageTypeFee = 1
		Else
			iIsPercentageTypeFee = 0
		End If 
		iDisplayOrder = oRs("displayorder")
	Else 
		sPermitFeePrefix = ""
		sPermitFee = ""
		iPermitFeeCategoryTypeId = 0
		iFeeReportingTypeId = 0
		iIsPercentageTypeFee = 0
		iDisplayOrder = 1
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' string GetPermitWorkClass( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitWorkClass( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT workclass FROM egov_permitworkclasses W, egov_permits P "
	sSql = sSql & "WHERE P.workclassid = W.workclassid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitWorkClass = oRs("workclass")
	Else 
		GetPermitWorkClass = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetPermitWorkScope( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitWorkScope( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT workscope FROM egov_permitworkscope W, egov_permits P "
	sSql = sSql & "WHERE P.workscopeid = W.workscopeid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitWorkScope = oRs("workscope")
	Else 
		GetPermitWorkScope = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetPrimaryContactIdForPermit( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPrimaryContactIdForPermit( ByVal iPermitId )
	Dim sSql, oRs, sContactType


	If PermitHasAPrimaryContractor( iPermitId ) Then 
		sContactType = "isprimarycontractor"
	Else
		sContactType = "isapplicant"
	End If 

	sSql = "SELECT permitcontactid FROM egov_permitcontacts WHERE " & sContactType & " = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPrimaryContactIdForPermit = CLng(oRs("permitcontactid"))
	Else
		GetPrimaryContactIdForPermit = CLng(0) 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' double GetPriorJobValue( iPermitId ) 
'-------------------------------------------------------------------------------------------------
Function GetPriorJobValue( ByVal iPermitId ) 
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(netjobvalue),0.00) AS priorjobvalue FROM egov_permitinvoices "
	sSql = sSql & " WHERE isvoided = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPriorJobValue = CDbl(FormatNumber(oRs("priorjobvalue"),2,,,0))
	Else
		GetPriorJobValue = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetPriorPermitStatus( iPermitStatusId )
'-------------------------------------------------------------------------------------------------
Function GetPriorPermitStatus( ByVal iPermitStatusId )
	Dim sSql, oRs

	sSql = "SELECT permitstatus FROM egov_permitstatuses WHERE nextpermitstatusid = " & iPermitStatusId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPriorPermitStatus = oRs("permitstatus")
	Else
		GetPriorPermitStatus = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetPriorPermitStatusId( iPermitStatusId )
'-------------------------------------------------------------------------------------------------
Function GetPriorPermitStatusId( ByVal iPermitStatusId )
	Dim sSql, oRs

	sSql = "SELECT permitstatusid FROM egov_permitstatuses WHERE nextpermitstatusid = " & iPermitStatusId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPriorPermitStatusId = oRs("permitstatusid")
	Else
		GetPriorPermitStatusId = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetProposedUse( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetProposedUse( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(proposeduse,'') AS proposeduse FROM egov_permits WHERE permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetProposedUse = oRs("proposeduse")
	Else
		GetProposedUse = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' string GetRequiredLicenseTypeIdsAsString( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetRequiredLicenseTypeIdsAsString( ByVal iPermitId )
	Dim sSql, oRs, sString, iCount

	sString = ""
	iCount = 0
	sSql = "SELECT DISTINCT licensetypeid FROM egov_permits_to_permitlicensetypes WHERE permitid = " & iPermitId 
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		If iCount > 0 Then
			sString = sString & ","
		End If 
		iCount = iCount + 1
		sString = sString & oRs("licensetypeid")
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	'response.write "sString = " & sString
	GetRequiredLicenseTypeIdsAsString = sString

End Function 


'-------------------------------------------------------------------------------------------------
' double GetResidentialUnitFeeAmount( iPermitFeeId, iResidentialUnits ) 
'-------------------------------------------------------------------------------------------------
Function GetResidentialUnitFeeAmount( ByVal iPermitFeeId, ByVal iResidentialUnits ) 
	Dim sSql, oRs, sFeeAmount, iUnitQty

	If CLng(iResidentialUnits) > CLng(0) Then 
		sSql = "SELECT atleastqty, notmorethanqty, baseamount, unitamount "
		sSql = sSql & " FROM egov_permitresidentialunitstepfees WHERE "
		sSql = sSql & iResidentialUnits & " >= atleastqty AND " & iResidentialUnits & " < notmorethanqty "
		sSql = sSql & " AND permitfeeid = " & iPermitFeeId 
		'response.write sSql & "<br /><br />"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then
			sFeeAmount = CLng(iResidentialUnits) - CLng(oRs("atleastqty"))
			sFeeAmount = CDbl(sFeeAmount) * CDbl(oRs("unitamount"))
			sFeeAmount = sFeeAmount + CDbl(oRs("baseamount"))
			sFeeAmount = FormatNumber(sFeeAmount,2,,,0)
		Else
			sFeeAmount = 0.00
		End If 
		'response.write "sFeeAmount: " & sFeeAmount & "<br /><br />"

		oRs.Close
		Set oRs = Nothing
	Else
		sFeeAmount = 0.00
	End If 
	
	GetResidentialUnitFeeAmount = sFeeAmount

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetResidentialUnitFeeMethodId( iOrgId )
'-------------------------------------------------------------------------------------------------
Function GetResidentialUnitFeeMethodId( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT permitfeemethodid FROM egov_permitfeemethods "
	sSql = sSql & " WHERE isresidentialunit = 1 AND orgid = " & iOrgid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetResidentialUnitFeeMethodId = CLng(oRs("permitfeemethodid"))
	Else
		GetResidentialUnitFeeMethodId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' void GetReviewDetailsForStatusEmail( iPermitReviewId, sReview, sReviewDesc, iReviewerId )
'-------------------------------------------------------------------------------------------------
Sub GetReviewDetailsForStatusEmail( ByVal iPermitReviewId, ByRef sReview, ByRef sReviewDesc, ByRef iReviewerId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(permitreviewtype,'') AS permitreviewtype, ISNULL(reviewdescription,'') AS reviewdescription, "
	sSql = sSql & " ISNULL(revieweruserid,0) AS revieweruserid "
	sSql = sSql & " FROM egov_permitreviews WHERE permitreviewid = " & iPermitReviewId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sReview = oRs("permitreviewtype")
		sReviewDesc = oRs("reviewdescription")
		iReviewerId = oRs("revieweruserid")
	Else
		sReview = ""
		sReviewDesc = ""
		iReviewerId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' string GetReviewStatusById( iReviewStatusId )
'-------------------------------------------------------------------------------------------------
Function GetReviewStatusById( ByVal iReviewStatusId )
	Dim sSql, oRs

	sSql = "SELECT reviewstatus FROM egov_reviewstatuses WHERE reviewstatusid = " & iReviewStatusId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetReviewStatusById = oRs("reviewstatus")
	Else
		GetReviewStatusById = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetReviewStatusId( sStatusFlag )
'-------------------------------------------------------------------------------------------------
Function GetReviewStatusId( ByVal sStatusFlag )
	Dim sSql, oRs

	sSql = "SELECT reviewstatusid FROM egov_reviewstatuses WHERE isforpermits = 1 AND orgid = " & session("orgid")
	sSql = sSql & " AND " & sStatusFlag & " = 1"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetReviewStatusId = CLng(oRs("reviewstatusid"))
	Else
		GetReviewStatusId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing
	
End Function 


'--------------------------------------------------------------------------------------------------
' double GetWaivedTotal( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetWaivedTotal( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(totalamount),0.00) AS totalamount FROM egov_permitinvoices "
	sSql = sSql & " WHERE isvoided = 0 AND allfeeswaived = 1 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetWaivedTotal =  FormatNumber(oRS("totalamount"),2,,,0)
	Else
		GetWaivedTotal = "0.00"
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' double GetYTDCostEstimate( sYearStart, sYearEnd, iInclude )
'--------------------------------------------------------------------------------------------------
Function GetYTDCostEstimate( ByVal sYearStart, ByVal sYearEnd, ByVal iInclude )
	Dim sSql, oRs, sIsVoided

	If clng(iInclude ) < clng(2) Then
		sIsVoided = " AND P.isvoided = " & iInclude
	Else
		sIsVoided = ""
	End If 

	sSql = "SELECT ISNULL(SUM(I.netjobvalue),0.00) AS costestimate "
	sSql = sSql & " FROM egov_permitinvoices I, egov_permits P "
	sSql = sSql & " WHERE P.orgid = " & session("orgid") & " AND I.permitid = P.permitid "
	sSql = sSql & " AND P.issueddate < '" & sYearEnd & "' AND P.issueddate > '" & sYearStart & "' "
	sSql = sSql & sIsVoided & " AND I.isvoided = 0"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetYTDCostEstimate = FormatNumber(oRs("costestimate"),2)
	Else
		GetYTDCostEstimate = FormatNumber(0.00,2)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double GetYTDJobValues( sYearStart, sYearEnd, iInclude )
'--------------------------------------------------------------------------------------------------
Function GetYTDJobValues( ByVal sYearStart, ByVal sYearEnd, ByVal iInclude )
	Dim sSql, oRs, sIsVoided

	If clng(iInclude ) < clng(2) Then
		sIsVoided = " AND P.isvoided = " & iInclude
	Else
		sIsVoided = ""
	End If 

	sSql = "SELECT ISNULL(SUM(P.jobvalue),0.00) AS jobvalue "
	sSql = sSql & " FROM egov_permits P "
	sSql = sSql & " WHERE P.orgid = " & session("orgid") 
	sSql = sSql & " AND P.issueddate < '" & sYearEnd & "' AND P.issueddate > '" & sYearStart & "' "
	'sSql = sSql & " AND ((P.issueddate < '" & sYearEnd & "' AND P.issueddate > '" & sYearStart & "') OR EXISTS(SELECT invoiceid FROM egov_permitinvoices I WHERE I.permitid = p.permitid and invoicedate < '" & sYearEnd & "' AND invoicedate >= '" & sYearStart & "' AND paymentid IS NOT NULL AND isvoided = 0 AND allfeeswaived = 0 )) "
	sSql = sSql & sIsVoided 
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetYTDJobValues = FormatNumber(oRs("jobvalue"),2)
	Else
		GetYTDJobValues = FormatNumber(0.00,2)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double GetYTDPermitFees( sYearStart, sYearEnd, iInclude )
'--------------------------------------------------------------------------------------------------
Function GetYTDPermitFees( ByVal sYearStart, ByVal sYearEnd, ByVal iInclude )
	Dim sSql, oRs, dFees, sIsVoided

	If clng(iInclude ) < clng(2) Then
		sIsVoided = " AND P.isvoided = " & iInclude
	Else
		sIsVoided = ""
	End If 

	sSql = "SELECT ISNULL(SUM(II.invoicedamount),0.00) AS invoicedamount "
	sSql = sSql & " FROM egov_permitinvoiceitems II, egov_permitfeecategorytypes C, egov_permitinvoices I, egov_permits P "
	sSql = sSql & " WHERE I.orgid = " & session("orgid")
	sSql = sSql & " AND I.permitid = P.permitid AND I.invoiceid = II.invoiceid "
	sSql = sSql & " AND II.permitfeecategorytypeid = C.permitfeecategorytypeid AND C.isgeneralbuildingtype = 1 "
	sSql = sSql & " AND ((P.issueddate < '" & sYearEnd & "' AND P.issueddate > '" & sYearStart & "') OR (I.invoicedate >= '" & sYearStart & "' AND I.invoicedate < '" & sYearEnd & "')) "
	sSql = sSql & " AND I.paymentid IS NOT NULL AND II.feereportingtypeid IS NULL " & sIsVoided & " AND I.isvoided = 0 AND I.allfeeswaived = 0"
	'response.write "<!--" & sSql & "-->" '"<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		dFees = FormatNumber(oRs("invoicedamount"),2)
	Else
		dFees = 0.00
	End If 

	If CDbl(dFees) = CDbl(0.00) Then
		dFees = 0.00
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetYTDPermitFees = FormatNumber(dFees,2)

End Function 


'--------------------------------------------------------------------------------------------------
' double GetYTDPermitReportingFees( sYearStart, sYearEnd, sReportingType, iInclude )
'--------------------------------------------------------------------------------------------------
Function GetYTDPermitReportingFees( ByVal sYearStart, ByVal sYearEnd, ByVal sReportingType, ByVal iInclude )
	Dim sSql, oRs, sIsVoided

	If clng(iInclude ) < clng(2) Then
		sIsVoided = " AND P.isvoided = " & iInclude
	Else
		sIsVoided = ""
	End If 

	sSql = "SELECT ISNULL(SUM(II.invoicedamount),0.00) AS invoicedamount "
	sSql = sSql & " FROM egov_permitinvoiceitems II, egov_permitfeereportingtypes R, egov_permitinvoices I, egov_permits P "
	sSql = sSql & " WHERE I.orgid = " & session("orgid")
	sSql = sSql & " AND I.permitid = P.permitid AND I.invoiceid = II.invoiceid "
	sSql = sSql & " AND R.feereportingtypeid = II.feereportingtypeid AND R." & sReportingType & " = 1 "
	sSql = sSql & " AND ((P.issueddate < '" & sYearEnd & "' AND P.issueddate > '" & sYearStart & "') OR (I.invoicedate >= '" & sYearStart & "' AND I.invoicedate < '" & sYearEnd & "')) "
	'sSql = sSql & " AND I.paymentid IS NOT NULL AND P.issueddate < '" & sYearEnd & "' AND P.issueddate > '" & sYearStart & "' AND I.invoicedate < '" & sYearEnd & "' "
	sSql = sSql & sIsVoided & " AND I.paymentid IS NOT NULL AND I.isvoided = 0 AND I.allfeeswaived = 0"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		dFees = FormatNumber(oRs("invoicedamount"),2)
	Else
		dFees = 0.00
	End If 

	If CDbl(dFees) = CDbl(0.00) Then
		dFees = 0.00
	End If 

	oRs.Close
	Set oRs = Nothing
	
	GetYTDPermitReportingFees = FormatNumber(dFees,2)

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetYTDResidentialUnits( sYearStart, sYearEnd, iInclude )
'--------------------------------------------------------------------------------------------------
Function GetYTDResidentialUnits( ByVal sYearStart, ByVal sYearEnd, ByVal iInclude )
	Dim sSql, oRs, sIsVoided

	If clng(iInclude ) < clng(2) Then
		sIsVoided = " AND isvoided = " & iInclude
	Else
		sIsVoided = ""
	End If 

	sSql = "SELECT ISNULL(SUM(residentialunits),0) AS residentialunits "
	sSql = sSql & " FROM egov_permits WHERE orgid = " & session("orgid")
	sSql = sSql & " AND issueddate < '" & sYearEnd & "' AND issueddate > '" & sYearStart & "' "
	sSql = sSql & sIsVoided
	'sSql = sSql & " AND isvoided = 0"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetYTDResidentialUnits = oRs("residentialunits")
	Else
		GetYTDResidentialUnits = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void InsertPaymentInformation( iPaymentId, iPaymentTypeId, sAmount, sStatus, sCheckNo )
'--------------------------------------------------------------------------------------------------
Sub InsertPaymentInformation( ByVal iPaymentId, ByVal iLedgerId, ByVal iPaymentTypeId, ByVal sAmount, ByVal sStatus, ByVal sCheckNo, ByVal iAccountId )
	Dim sSql 

	sSql = "Insert Into egov_verisign_payment_information "
	sSql = sSql & " (paymentid, ledgerid, paymenttypeid, amount, paymentstatus, checkno, citizenuserid) Values ("
	sSql = sSql & iPaymentid & ", " & iLedgerId & ", " & iPaymentTypeId & ", " & sAmount & ", '" & sStatus & "', "
	sSql = sSql & sCheckNo & ", " & iAccountId & " )"
	'response.write sSql & "<br /><br />"

	RunSQL sSql

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean InspectionCanSaveChanges( iPermitId, iPermitInspectionId )
'--------------------------------------------------------------------------------------------------
Function InspectionCanSaveChanges( ByVal iPermitId, ByVal iPermitInspectionId )
	Dim sSql, oRs, bPermitIsCompleted

	bPermitIsCompleted = GetPermitIsCompleted( iPermitId ) '	in permitcommonfunctions.asp

	If bPermitIsCompleted Then 
		InspectionCanSaveChanges = False 
	Else 
		sSql = "SELECT I.isfinal, S.isdone FROM egov_permitinspections I, egov_inspectionstatuses S "
		sSql = sSql & " WHERE I.inspectionstatusid = S.inspectionstatusid AND I.permitinspectionid = " & iPermitInspectionId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then
			If oRs("isfinal") Then
				' if the status is in "done" status allow changes, otherwise check for others all done
				If oRs("isdone") Then
					' this inspection is in done status but they may need to update something on it
					InspectionCanSaveChanges = True
				Else 
					' Final inspections can only save if all other inspections for the permit are "done"
					InspectionCanSaveChanges = AllOtherInspectionsAreDone( iPermitId, iPermitInspectionId )  ' in permitcommonfunctions.asp
				End If 
			Else
				InspectionCanSaveChanges = True 
			End If 
		Else
			InspectionCanSaveChanges = False  
		End If 

		oRs.Close
		Set oRs = Nothing 
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' void MakeAPermitLogEntry iPermitid, sActivity, sActivityComment, sInternalComment, sExternalComment, iPermitStatusId, iIsInspectionEntry, iIsReviewEntry, iIsActivityEntry, iPermitReviewId, iPermitInspectionId, iReviewStatusId, iInspectionStatusId 
'--------------------------------------------------------------------------------------------------
Sub MakeAPermitLogEntry( ByVal iPermitid, ByVal sActivity, ByVal sActivityComment, ByVal sInternalComment, ByVal sExternalComment, ByVal iPermitStatusId, ByVal iIsInspectionEntry, _
		ByVal iIsReviewEntry, ByVal iIsActivityEntry, ByVal iPermitReviewId, ByVal iPermitInspectionId, ByVal iReviewStatusId, ByVal iInspectionStatusId )
	Dim sSql, oCmd

	sSql = "INSERT INTO egov_permitlog ( orgid, permitid, entrydate, adminuserid, activity, activitycomment, "
	sSql = sSql & " internalcomment, externalcomment, permitstatusid, isinspectionentry, isreviewentry, "
	sSql = sSql & " isactivityentry, permitreviewid, permitinspectionid, reviewstatusid, inspectionstatusid ) VALUES ( "
	sSql = sSql & session("orgid") & ", " & iPermitid & ", dbo.GetLocalDate(" & Session("OrgID") & ",getdate()), " 
	sSql = sSql & session("userid") & ", " & sActivity & ", " & sActivityComment & ", " & sInternalComment & ", " & sExternalComment 
	sSql = sSql & ", " & iPermitStatusId & ", " & iIsInspectionEntry & ", " & iIsReviewEntry & ", " & iIsActivityEntry
	sSql = sSql & ", " & iPermitReviewId & ", " & iPermitInspectionId & ", " & iReviewStatusId & ", " & iInspectionStatusId & " )"
	
	'response.write "<p>" & sSql & "</p><br /><br />"

	RunSQL sSql

End Sub 


'--------------------------------------------------------------------------------------------------
' integer MakeJournalEntry( iPaymentLocationId, iAdminLocationId, iCitizenId, iAdminUserId, sAmount, iJournalEntryTypeID, sNotes )
'--------------------------------------------------------------------------------------------------
Function MakeJournalEntry( ByVal iPaymentLocationId, ByVal iAdminLocationId, ByVal iCitizenId, ByVal iAdminUserId, ByVal sAmount, ByVal iJournalEntryTypeID, ByVal sNotes )
	Dim sSql

	sSql = "Insert into egov_class_payment (paymentdate, paymentlocationid, orgid, adminlocationid, "
	sSql = sSql & " userid, adminuserid, paymenttotal, journalentrytypeid, notes, isforpermits ) Values (dbo.GetLocalDate(" & Session("orgid") & ",GetDate()), " 
	sSql = sSql & iPaymentLocationId & ", " & Session("orgid") & ", " & iAdminLocationId & ", "
	sSql = sSql & iCitizenId & ", " & iAdminUserId & ", " & sAmount & ", " & iJournalEntryTypeID & ", '" & sNotes & "', 1 )"
	'response.write sSql & "<br /><br />"

	MakeJournalEntry = RunIdentityInsert( sSql )

End Function 


'--------------------------------------------------------------------------------------------------
' integer MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, cPriorBalance, iPriceTypeid, iPermitId, iInvoiceId, iPermitFeeId )
'--------------------------------------------------------------------------------------------------
Function MakeLedgerEntry( ByVal iOrgID, ByVal iAccountId, ByVal iJournalId, ByVal cAmount, ByVal iItemTypeId, ByVal sEntryType, ByVal sPlusMinus, ByVal iItemId, ByVal iIsPaymentAccount, ByVal iPaymentTypeId, ByVal cPriorBalance, ByVal iPriceTypeid, ByVal iPermitId, ByVal iInvoiceId, ByVal iPermitFeeId )
	Dim sSql

	sSql = "Insert Into egov_accounts_ledger ( paymentid,orgid,entrytype,accountid,amount,itemtypeid,plusminus, "
	sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, pricetypeid, permitid, invoiceid, permitfeeid ) Values ( "
	sSql = sSql & iJournalId & ", " & iOrgID & ", '" & sEntryType & "', " & iAccountId & ", " & cAmount & ", " & iItemTypeId & ", '" & sPlusMinus & "', " 
	sSql = sSql & iItemId & ", " & iIsPaymentAccount & ", " & iPaymentTypeId & ", " & cPriorBalance & ", " & iPriceTypeid  & ", "
	sSql = sSql & iPermitId & ", " & iInvoiceId & ", " & iPermitFeeId & " )"

	'response.write sSql & "<br /><br />"

	MakeLedgerEntry = RunIdentityInsert( sSql )

End Function 


'-------------------------------------------------------------------------------------------------
' boolean OrgHasFeeType( sFeeType )
'-------------------------------------------------------------------------------------------------
Function OrgHasFeeType( ByVal sFeeType )
	Dim sSql, oRs

	sSql = "SELECT COUNT(feereportingtypeid) AS hits "
	sSql = sSql & " FROM egov_permitfeereportingtypes WHERE " & sFeeType & " = 1 AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then
			OrgHasFeeType = True 
		Else
			OrgHasFeeType = False 
		End If 
	Else
		OrgHasFeeType = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' boolean PermitFeesArePaid( iPermitId )
'-------------------------------------------------------------------------------------------------
Function PermitFeesArePaid( ByVal iPermitId )

	If SomeFeesSetToZero( iPermitId ) Then
		PermitFeesArePaid = False 
	Else
		If AllFeesPaidOrWaived( iPermitId ) Then
			PermitFeesArePaid = True 
		Else
			PermitFeesArePaid = False 
		End If 
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean PermitHasAPrimaryContractor( iPermitId )
'--------------------------------------------------------------------------------------------------
Function PermitHasAPrimaryContractor( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitcontactid) AS hits "
	sSql = sSql & " FROM egov_permitcontacts WHERE isprimarycontractor = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then 
			PermitHasAPrimaryContractor = True 
		Else
			PermitHasAPrimaryContractor = False 
		End If 
	Else
		PermitHasAPrimaryContractor = False 
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean PermitHasBeenApproved( iPermitId )
'--------------------------------------------------------------------------------------------------
Function PermitHasBeenApproved( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT approveddate FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If IsNull(oRs("approveddate")) Then 
			PermitHasBeenApproved = False
		Else 
			PermitHasBeenApproved = True 
		End If 
	Else
		PermitHasBeenApproved = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' boolean PermitHasDetail( iPermitid, sDetailField )
'--------------------------------------------------------------------------------------------------
Function PermitHasDetail( ByVal iPermitid, ByVal sDetailField )
	Dim sSql, oRs

	sSql = "SELECT COUNT(F.detailfieldid) AS hits "
	sSql = sSql & " FROM egov_permits P, egov_permittypes_to_permitdetailfields F, egov_permitdetailfields D "
	sSql = sSql & " WHERE P.permitid = " & iPermitId & " AND D.detailfield = '" & sDetailField & "' AND "
	sSql = sSql & " P.permittypeid = F.permittypeid AND F.detailfieldid = D.detailfieldid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then 
			PermitHasDetail = True 
		Else 
			PermitHasDetail = False  
		End If 
	Else
		PermitHasDetail = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' boolean PermitHasLicenseRequirement( iPermitid )
'--------------------------------------------------------------------------------------------------
Function PermitHasLicenseRequirement( ByVal iPermitid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(licensetypeid) AS hits FROM egov_permits_to_permitlicensetypes "
	sSql = sSql & "WHERE isrequired = 1 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then 
			PermitHasLicenseRequirement = True 
		Else 
			PermitHasLicenseRequirement = False  
		End If 
	Else
		PermitHasLicenseRequirement = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' boolean PermitIsInBuildingPermitCategory( iPermitId )
'--------------------------------------------------------------------------------------------------
Function PermitIsInBuildingPermitCategory( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT C.isbuildingpermitcategory FROM egov_permitcategories C, egov_permits P "
	sSql = sSql & "WHERE C.permitcategoryid = P.permitcategoryid AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("isbuildingpermitcategory") Then 
			PermitIsInBuildingPermitCategory = True 
		Else 
			PermitIsInBuildingPermitCategory = False 
		End If 
	Else
		PermitIsInBuildingPermitCategory = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' boolean PermitNumberPrefixIsNotNone( sPermitNumberPrefix )
'--------------------------------------------------------------------------------------------------
Function PermitNumberPrefixIsNotNone( ByVal sPermitNumberPrefix )
	Dim sSql, oRs

	sSql = "SELECT isnone FROM egov_permitnumberprefixes WHERE orgid = " & session("orgid")
	sSql = sSql & " AND permitnumberprefix = '" & sPermitNumberPrefix & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("isnone") Then 
			PermitNumberPrefixIsNotNone = False
		Else 
			PermitNumberPrefixIsNotNone = True
		End If 
	Else
		PermitNumberPrefixIsNotNone = True
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' boolean PermitStatusAllowsButton( iPermitId )
'--------------------------------------------------------------------------------------------------
Function PermitStatusAllowsButton( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT S.canusebutton FROM egov_permits P, egov_permitstatuses S "
	sSql = sSql & " WHERE P.permitstatusid = S.permitstatusid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("canusebutton") Then 
			PermitStatusAllowsButton = True   	
		Else 
			PermitStatusAllowsButton = False    
		End If 
	Else
		PermitStatusAllowsButton = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' boolean PermitStatusAllowsDeletes( iPermitId )
'--------------------------------------------------------------------------------------------------
Function PermitStatusAllowsDeletes( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT S.allowdeletes FROM egov_permits P, egov_permitstatuses S "
	sSql = sSql & " WHERE P.permitstatusid = S.permitstatusid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("allowdeletes") Then 
			PermitStatusAllowsDeletes = True   	
		Else 
			PermitStatusAllowsDeletes = False    
		End If 
	Else
		PermitStatusAllowsDeletes = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' boolean PermitStatusAllowsInspections( iPermitId )
'--------------------------------------------------------------------------------------------------
Function PermitStatusAllowsInspections( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT S.blockinspections FROM egov_permits P, egov_permitstatuses S "
	sSql = sSql & " WHERE P.permitstatusid = S.permitstatusid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("blockinspections") Then 
			PermitStatusAllowsInspections = False  	
		Else 
			PermitStatusAllowsInspections = True   
		End If 
	Else
		PermitStatusAllowsInspections = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'-------------------------------------------------------------------------------------------------
' void PushOutPermitExpirationDate iPermitId 
'-------------------------------------------------------------------------------------------------
Sub PushOutPermitExpirationDate( ByVal iPermitId )
	Dim sSql, oRs, sNewDate

	' Get the days to push out the expiration 
	sSql = "SELECT ISNULL(expirationdays,0) AS expirationdays FROM egov_permitpermittypes "
	sSql = sSql & " WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		' update the permit expiration date
		If CLng(oRs("expirationdays")) > CLng(0) Then 
			sNewDate = DateAdd("d", oRs("expirationdays"), Date())
			sSql = "UPDATE egov_permits SET expirationdate = '" & sNewDate & "', isexpired = 0 WHERE permitid = " & iPermitId
			RunSQL sSql
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

	' Set the last activity date
	SetLastActivityDate iPermitId   ' in permitcommonfunctions.asp

End Sub 


'-------------------------------------------------------------------------------------------------
' void RecalcConstTypeTotalSqFt iPermitId, sSqFt, iConstructiontyperate, sType 
'-------------------------------------------------------------------------------------------------
Sub RecalcConstTypeFee( ByVal iPermitId, ByVal sSqFt, ByVal iConstructiontyperate, ByVal sType )
	Dim sSql, oRs, sFeeAmount
		
	sSql = "SELECT F.permitfeeid, ISNULL(F.baseamount,0.00) AS baseamount, F.atleastqty, F.notmorethanqty, "
	sSql = sSql & " ISNULL(F.unitqty,0) AS unitqty, ISNULL(F.unitamount,0.0000) AS unitamount, ISNULL(F.minimumamount,0.00) AS minimumamount "
	sSql = sSql & " FROM egov_permitfees F, egov_permitfeemethods M "
	sSql = sSql & " WHERE M." & sType & " = 1 AND M.permitfeemethodid = F.permitfeemethodid "
	sSql = sSql & " AND F.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		If iConstructiontyperate <> "NULL" Then 
			sFeeAmount = CDbl(sSqFt) * CDbl(iConstructiontyperate)
			sFeeAmount = sFeeAmount * GetFeeMultipliers( oRs("permitfeeid") )
			If CDbl(sFeeAmount) < CDbl(oRs("minimumamount")) Then
				sFeeAmount = CDbl(oRs("minimumamount"))
			End If 
		Else
			' No Construction Rate
			If CDbl(oRs("minimumamount")) > CDbl(0.00) Then
				sFeeAmount = CDbl(oRs("minimumamount"))
			Else 
				sFeeAmount = CDbl(0.00)
			End If 
		End If 
		sFeeAmount = FormatNumber(sFeeAmount,2,,,0)
		sSql = "UPDATE egov_permitfees SET feeamount = " & sFeeAmount & " WHERE permitfeeid = " & oRs("permitfeeid")
		'response.write sSql & "<br />"
		RunSQL sSql

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void RecalcCuFtFee( iPermitId, dblCuFt, sType )
'-------------------------------------------------------------------------------------------------
Sub RecalcCuFtFee( ByVal iPermitId, ByVal dblCuFt, ByVal sType )
	Dim sSql, oRs, sFeeAmount
		
	sSql = "SELECT F.permitfeeid, ISNULL(F.baseamount,0.00) AS baseamount, "
	sSql = sSql & " ISNULL(F.unitqty,0) AS unitqty, ISNULL(F.unitamount,0.0000) AS unitamount, ISNULL(F.minimumamount,0.00) AS minimumamount "
	sSql = sSql & " FROM egov_permitfees F, egov_permitfeemethods M "
	sSql = sSql & " WHERE M." & sType & " = 1 AND M.permitfeemethodid = F.permitfeemethodid "
	sSql = sSql & " AND F.permitid = " & iPermitId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		sFeeAmount = CDbl(dblCuFt)
		If CLng(oRs("unitqty")) > CLng(0) Then 
			sFeeAmount = Ceiling(sFeeAmount / CLng(oRs("unitqty"))) * CDbl(oRs("unitamount"))
		End If 

		sFeeAmount = sFeeAmount * GetFeeMultipliers( oRs("permitfeeid") )
		'response.write "sFeeAmount = " & sFeeAmount & "<br />"

		If CDbl(oRs("baseamount")) > CDbl(0.00) Then
			sFeeAmount = sFeeAmount + CDbl(oRs("baseamount"))
		End If 

		If CDbl(sFeeAmount) < CDbl(oRs("minimumamount")) Then
			sFeeAmount = CDbl(oRs("minimumamount"))
		End If 
		sFeeAmount = FormatNumber(sFeeAmount,2,,,0)
		sSql = "UPDATE egov_permitfees SET feeamount = " & sFeeAmount & " WHERE permitfeeid = " & oRs("permitfeeid")
		'response.write sSql & "<br /><br />"
		RunSQL sSql

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void ReCalcExamHours iPermitId, sHours 
'-------------------------------------------------------------------------------------------------
Sub ReCalcExamHours( ByVal iPermitId, ByVal sHours )
	Dim sSql, oRs, sFeeAmount
		
	sSql = "SELECT F.permitfeeid, ISNULL(F.baseamount,0.00) AS baseamount,  "
	sSql = sSql & " ISNULL(F.unitqty,0) AS unitqty, ISNULL(F.unitamount,0.0000) AS unitamount, ISNULL(F.minimumamount,0.00) AS minimumamount "
	sSql = sSql & " FROM egov_permitfees F, egov_permitfeemethods M "
	sSql = sSql & " WHERE M.ishourly = 1 AND M.permitfeemethodid = F.permitfeemethodid "
	sSql = sSql & " AND F.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		sFeeAmount = CDbl(sHours)
		sFeeAmount = sFeeAmount * GetFeeMultipliers( oRs("permitfeeid") )
		If CLng(oRs("unitqty")) > CLng(1) Then 
			sFeeAmount = sFeeAmount / oRs("unitqty")
		End If 
		If CDbl(oRs("unitamount")) > CDbl(0) Then 
			sFeeAmount = sFeeAmount * CDbl(oRs("unitamount"))
		End If 
		sFeeAmount = sFeeAmount + CDbl(oRs("baseamount"))
		If CDbl(sFeeAmount) < CDbl(oRs("minimumamount")) Then
			sFeeAmount = CDbl(oRs("minimumamount"))
		End If 
		sFeeAmount = FormatNumber(sFeeAmount,2,,,0)
		sSql = "UPDATE egov_permitfees SET feeamount = " & sFeeAmount & " WHERE permitfeeid = " & oRs("permitfeeid")
		'response.write sSql & "<br />"
		RunSQL sSql

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void RecalcPercentageFees iPermitId 
'-------------------------------------------------------------------------------------------------
Sub RecalcPercentageFees( ByVal iPermitId )
	Dim sSql, oRs, sFeeAmount

	sSql = "SELECT permitfeeid, percentage, permitfeecategorytypeid, ISNULL(minimumamount,0.00) AS minimumamount "
	sSql = sSql & "FROM egov_permitfees "
	sSql = sSql & "WHERE ispercentagetypefee = 1 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		If CLng(oRs("permitfeecategorytypeid")) <> CLng(-1) Then 
			sFeeAmount = GetCategoryFeesTotalForPercentage( iPermitId, oRs("permitfeecategorytypeid") )	 ' in permitcommonfunctions.asp
		Else
			sFeeAmount = GetAllFeesTotalForPercentage( iPermitId )	 ' in permitcommonfunctions.asp
		End If 

		sFeeAmount = sFeeAmount * CDbl(oRs("percentage"))

		If CDbl(oRs("minimumamount")) > sFeeAmount Then
			sFeeAmount = CDbl(oRs("minimumamount"))
		End If 
		sFeeAmount = FormatNumber(sFeeAmount,2,,,0)
		sSql = "UPDATE egov_permitfees SET feeamount = " & sFeeAmount & " WHERE permitfeeid = " & oRs("permitfeeid")
		RunSQL sSql
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void RecalcResidentialUnitFees iPermitId, iResidentialUnits 
'-------------------------------------------------------------------------------------------------
Sub RecalcResidentialUnitFees( ByVal iPermitId, ByVal iResidentialUnits )
	Dim sSql, oRs, sFeeAmount

	sSql = "SELECT permitfeeid, ISNULL(minimumamount,0.00) AS minimumamount FROM egov_permitfees "
	sSql = sSql & " WHERE isresidentialunittypefee = 1 AND permitid = " & iPermitId
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		sFeeAmount = GetResidentialUnitFeeAmount( oRs("permitfeeid"), iResidentialUnits )  ' In permitcommonfunctions.asp
		If CDbl(oRs("minimumamount")) > CDbl(sFeeAmount) Then
			sFeeAmount = CDbl(oRs("minimumamount"))
		End If 
		sFeeAmount = FormatNumber(sFeeAmount,2,,,0)
		sSql = "UPDATE egov_permitfees SET feeamount = " & sFeeAmount & " WHERE permitfeeid = " & oRs("permitfeeid")
		'response.write sSql & "<br />"
		RunSQL sSql
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void RecalcSqFtFee( iPermitId, dblSqFt, sType )
'-------------------------------------------------------------------------------------------------
Sub RecalcSqFtFee( ByVal iPermitId, ByVal dblSqFt, ByVal sType )
	Dim sSql, oRs, sFeeAmount
		
	sSql = "SELECT F.permitfeeid, ISNULL(F.baseamount,0.00) AS baseamount, "
	sSql = sSql & " ISNULL(F.unitqty,0) AS unitqty, ISNULL(F.unitamount,0.0000) AS unitamount, ISNULL(F.minimumamount,0.00) AS minimumamount "
	sSql = sSql & " FROM egov_permitfees F, egov_permitfeemethods M "
	sSql = sSql & " WHERE M." & sType & " = 1 AND M.permitfeemethodid = F.permitfeemethodid "
	sSql = sSql & " AND F.permitid = " & iPermitId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		sFeeAmount = CDbl(dblSqFt)
		If CLng(oRs("unitqty")) > CLng(0) Then 
			sFeeAmount = ( sFeeAmount / CLng(oRs("unitqty")) ) * CDbl(oRs("unitamount"))
		End If 

		sFeeAmount = sFeeAmount * GetFeeMultipliers( oRs("permitfeeid") )
		'response.write "sFeeAmount = " & sFeeAmount & "<br />"

		If CDbl(oRs("baseamount")) > CDbl(0.00) Then
			sFeeAmount = sFeeAmount + CDbl(oRs("baseamount"))
		End If 

		If CDbl(sFeeAmount) < CDbl(oRs("minimumamount")) Then
			sFeeAmount = CDbl(oRs("minimumamount"))
		End If 
		sFeeAmount = FormatNumber(sFeeAmount,2,,,0)
		sSql = "UPDATE egov_permitfees SET feeamount = " & sFeeAmount & " WHERE permitfeeid = " & oRs("permitfeeid")
		'response.write sSql & "<br /><br />"
		RunSQL sSql

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void RecalcValuationFees iPermitId, dJobValue 
'-------------------------------------------------------------------------------------------------
Sub RecalcValuationFees( ByVal iPermitId, ByVal dJobValue )
	Dim sSql, oRs, sFeeAmount
	' Get the valuation based fees and recalculate the fee amounts then update the fee records

	' should pull one row per permitfeeid
	sSql = "SELECT F.permitfeeid, ISNULL(S.baseamount,0.00) AS baseamount, S.atleastvalue, S.notmorethanvalue, "
	sSql = sSql & " ISNULL(S.unitqty,0) AS unitqty, ISNULL(S.unitamount,0.00) AS unitamount, ISNULL(F.minimumamount,0.00) AS minimumamount "
	sSql = sSql & " FROM egov_permitfees F, egov_permitvaluationstepfees S "
	sSql = sSql & " WHERE F.isvaluationtypefee = 1 AND F.permitfeeid = S.permitfeeid "
	sSql = sSql & " AND " & dJobValue & " >= S.atleastvalue AND " & dJobValue & " < S.notmorethanvalue "
	sSql = sSql & " AND F.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		' calculate the new fee
		If CLng(oRs("unitqty")) > CLng(0) Then 
			sFeeAmount = Ceiling(CDbl(dJobValue / oRs("unitqty"))) * oRs("unitqty")
			sFeeAmount = CDbl(sFeeAmount) - CDbl(oRs("atleastvalue"))
			sFeeAmount = sFeeAmount / oRs("unitqty")
			sFeeAmount = sFeeAmount * oRs("unitamount")
			sFeeAmount = sFeeAmount + oRs("baseamount")
			If CDbl(sFeeAmount) < CDbl(oRs("minimumamount")) Then
				sFeeAmount = CDbl(oRs("minimumamount"))
			End If 
			sFeeAmount = FormatNumber(sFeeAmount,2,,,0)
		Else
			sFeeAmount = FormatNumber(oRs("baseamount"),2,,,0)
		End If 
		sFeeAmount = FormatNumber(sFeeAmount,2,,,0)
		sSql = "UPDATE egov_permitfees SET feeamount = " & sFeeAmount & " WHERE permitfeeid = " & oRs("permitfeeid")
		'response.write sSql & "<br />"
		RunSQL sSql

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' boolean RescheduleInspectionForStatus( iInspectionStatusId )
'-------------------------------------------------------------------------------------------------
Function RescheduleInspectionForStatus( ByVal iInspectionStatusId )
	Dim sSql, oRs

	sSql = "SELECT rescheduleinspection FROM egov_inspectionstatuses WHERE inspectionstatusid = " &  iInspectionStatusId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("rescheduleinspection") Then
			RescheduleInspectionForStatus = True 
		Else
			RescheduleInspectionForStatus = False 
		End If 
	Else
		RescheduleInspectionForStatus = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer ReschedulePermitInspection( iPermitInspectionId )
'-------------------------------------------------------------------------------------------------
Function ReschedulePermitInspection( ByVal iPermitInspectionId )
	Dim sSql, oRs, iIsFinal, iInitialStatusid, iInspectionOrder, iNewPermitInspectionId, iIsRequired
	Dim sInspectorUserId

	iNewPermitInspectionId = 0
	iInitialStatusid = GetInspectionStatusId( "isinitialstatus" )	' in permitcommonfunctions.asp

	sSql = "SELECT permitid, permittypeid, permitinspectiontypeid, permitinspectiontype, inspectiondescription, "
	sSql = sSql & " ISNULL(inspectoruserid,0) AS inspectoruserid, isrequired, isfinal, contact, contactphone "
	sSql = sSql & " FROM egov_permitinspections WHERE permitinspectionid = " & iPermitInspectionId

	response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If oRs("isfinal") Then
			iIsFinal = 1
		Else
			iIsFinal = 0
		End If 
		If oRs("isrequired") Then
			iIsRequired = 1
		Else
			iIsRequired = 0
		End If 
		If CLng(oRs("inspectoruserid")) = CLng(0) Then 
			sInspectorUserId = "NULL"
		Else
			sInspectorUserId = oRs("inspectoruserid")
		End If 
	
		' Get the next Inspection Order for this permit
		iInspectionOrder = GetNextInspectionOrder( oRs("permitid") )  	' in permitcommonfunctions.asp

		sSql = "INSERT INTO egov_permitinspections ( orgid, permitid, permittypeid, permitinspectiontypeid, permitinspectiontype,  "
		sSql = sSql & " inspectiondescription, inspectoruserid, inspectionorder, isfinal, isrequired, isincluded, inspectionstatusid, "
		sSql = sSql & " contact, contactphone, routeorder, isreinspection ) VALUES ( " & session("orgid")
		sSql = sSql & ", " & oRs("permitid") & ", " & oRs("permittypeid") & ", " & iPermitInspectionTypeId &  ", '" & dbsafe(oRs("permitinspectiontype"))
		sSql = sSql & "', '" & dbsafe(oRs("inspectiondescription")) & "', " & sInspectorUserId & ", " & iInspectionOrder
		sSql = sSql & ", " & iIsFinal & ", " & iIsRequired & ", 1, " & iInitialStatusid & ", '" & dbsafe(oRs("contact")) & "', '"
		sSql = sSql & dbsafe(oRs("contactphone")) & "', 1, 1 )"
		
		response.write sSql & "<br /><br />"

		iNewPermitInspectionId = RunIdentityInsert( sSql )
	End If  

	oRs.Close
	Set oRs = Nothing 

	ReschedulePermitInspection = iNewPermitInspectionId

End Function 


'-------------------------------------------------------------------------------------------------
' integer RunIdentityInsert( sInsertStatement )
'-------------------------------------------------------------------------------------------------
Function RunIdentityInsert( ByVal sInsertStatement )
	Dim sSql, iReturnValue, oInsert

	iReturnValue = 0

'	response.write "<p>" & sInsertStatement & "</p><br /><br />"
'	response.flush
	session("RunIdentityInsertSql") = sInsertStatement

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSql = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.Open sSql, Application("DSN"), 3, 3
	iReturnValue = oInsert("ROWID")
	oInsert.Close
	Set oInsert = Nothing
	session("RunIdentityInsertSql") = ""

	RunIdentityInsert = iReturnValue

End Function


'-------------------------------------------------------------------------------------------------
' void RunSQL sSql 
'-------------------------------------------------------------------------------------------------
Sub RunSQL( ByVal sSql )
	Dim oCmd

'	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush
	session("runsql") = sSql

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

	session("runsql") = ""

End Sub 


'--------------------------------------------------------------------------------------------------
' void SendEmailPermits sToName, sToEmail, sFromName, sFromEmail, sSubject, sHTMLBody 
'--------------------------------------------------------------------------------------------------
Sub SendEmailPermits( ByVal sToName, ByVal sToEmail, ByVal sFromName, ByVal sFromEmail, ByVal sSubject, ByVal sHTMLBody )

	sendEmail "", sToEmail & "(" & sToName & ")", "", sSubject, FormatHTML( sHTMLBody ), clearHTMLTags( sHTMLBody ), "N"

End Sub


'-------------------------------------------------------------------------------------------------
' void SendPermitApprovedAlert iPermitId 
'-------------------------------------------------------------------------------------------------
Sub SendPermitApprovedAlert( ByVal iPermitId )
	Dim sSql, oRs, iPermitTypeId, sToName, sSubject, sHTMLBody, sFromName, sPermitNo, sDesc
	Dim sJobSite, sOrgName, sStatus, sLocation , sLocationType

	' Pull the permit details needed
	sPermitNo = GetPermitNumber( iPermitId )
	sDesc = GetPermitTypeDesc( iPermitId, True )
	sJobSite = GetPermitJobSite( iPermitId )
	iPermitTypeId = GetPermitTypeId( iPermitId )
	sStatus = GetPermitStatusByPermitId( iPermitId ) 
	sLocation = Replace(GetPermitPermitLocation( iPermitId ), Chr(10), Chr(10) & "<br />")
	sLocationType = GetPermitLocationType( iPermitId )

	sSubject = "Permit Reviews Approved for " & sPermitNo
	sOrgName = GetOrgName( session("orgid") )
	sFromName = sOrgName & " E-GOV WEBSITE"

	' Build the email body
	sHTMLBody = "<p>This automated message was sent by the " & sOrgName & " E-Gov web site. Do not reply to this message.</p>" & vbcrlf  & vbcrlf 
	sHTMLBody = sHTMLBody & "<p>All reviews for this permit have been approved.</p>" & vbcrlf  & vbcrlf 
	sHTMLBody = sHTMLBody & "<p>Permit #: " & sPermitNo & "<br />" & vbcrlf
	sHTMLBody = sHTMLBody & "Permit Type: " & sDesc & "<br />" & vbcrlf
	sHTMLBody = sHTMLBody & "Permit Status: " & sStatus & "<br />" & vbcrlf
	If sLocationType = "address" Then 
		sHTMLBody = sHTMLBody & vbcrlf & "Job Site: " & sJobSite
	End If 
	If sLocationType = "location" Then 
		sHTMLBody = sHTMLBody & vbcrlf & "Location: " & sLocation
	End If 
	sHTMLBody = sHTMLBody & "</p>" & vbcrlf & vbcrlf
	sHTMLBody = sHTMLBody & "<p><a href=""" & session("egovclientwebsiteurl") & "/admin/permits/permitedit.asp?permitid=" & iPermitId & """ title=""click to view"">Click here to view this permit.</a></p>"

	' Pull any reviewers that are marked to get the approved alerts and send out
	sSql = "SELECT U.firstname, U.lastname, ISNULL(U.email,'') AS email "
	sSql = sSql & " FROM users U, egov_permittypes_to_permitalerttypes A, egov_permitalerttypes T "
	sSql = sSql & " WHERE u.isdeleted = 0 AND T.isforallapproved = 1 AND T.permitalerttypeid = A.permitalerttypeid "
	sSql = sSql & " AND U.userid = A.notifyuserid AND A.permittypeid = " & iPermitTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		If oRs("email") <> "" Then 
			sToName = oRs("firstname") & " " & oRs("lastname")
			'SendEmailPermits sToName, oRs("email"), sFromName, "webmaster@eclink.com", sSubject, sHTMLBody 
			'sendEmail "", oRs("email") & "(" & sToName & ")", "", sSubject, sHTMLBody, "", ""
			sendEmail "", sToName & " <" & oRs("email") & ">", "", sSubject, sHTMLBody, "", "" 
		End If 
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void SendPermitPassedFinalInspectionAlert iPermitId 
'-------------------------------------------------------------------------------------------------
Sub SendPermitPassedFinalInspectionAlert( ByVal iPermitId )
	Dim sSql, oRs, iPermitTypeId, sToName, sSubject, sHTMLBody, sFromName, sPermitNo, sDesc
	Dim sJobSite, sOrgName, sStatus, sLocation, sLocationType

	' Pull the permit details needed
	sPermitNo = GetPermitNumber( iPermitId )
	sDesc = GetPermitTypeDesc( iPermitId, True )
	sJobSite = GetPermitJobSite( iPermitId )
	iPermitTypeId = GetPermitTypeId( iPermitId )
	sStatus = GetPermitStatusByPermitId( iPermitId ) 
	sLocation = Replace(GetPermitPermitLocation( iPermitId ), Chr(10), Chr(10) & "<br />")
	sLocationType = GetPermitLocationType( iPermitId )

	sSubject = "Permit Has Passed Final Inspection"
	sOrgName = GetOrgName( session("orgid") )
	sFromName = sOrgName & " E-GOV WEBSITE"

	' Build the email body
	sHTMLBody = "<p>This automated message was sent by the " & sOrgName & " E-Gov web site. Do not reply to this message.</p>" & vbcrlf  & vbcrlf 
	sHTMLBody = sHTMLBody & "<p>All inspections for this permit have passed.</p>" & vbcrlf  & vbcrlf 
	sHTMLBody = sHTMLBody & "<p>Permit #: " & sPermitNo & "<br />" & vbcrlf
	sHTMLBody = sHTMLBody & "Permit Type: " & sDesc & "<br />" & vbcrlf
	sHTMLBody = sHTMLBody & "Permit Status: " & sStatus & "<br />" & vbcrlf
	If sLocationType = "address" Then 
		sHTMLBody = sHTMLBody & vbcrlf & "Job Site: " & sJobSite
	End If 
	If sLocationType = "location" Then 
		sHTMLBody = sHTMLBody & vbcrlf & "Location: " & sLocation
	End If 
	sHTMLBody = sHTMLBody & "</p>" & vbcrlf & vbcrlf
	sHTMLBody = sHTMLBody & "<p><a href=""" & session("egovclientwebsiteurl") & "/admin/permits/permitedit.asp?permitid=" & iPermitId & """ title=""click to view"">Click here to view the details for this permit.</a></p>"

	' Pull any inspectors that are marked to get the final passed alerts and send out
	sSql = "SELECT U.firstname, U.lastname, ISNULL(U.email,'') AS email "
	sSql = sSql & " FROM users U, egov_permittypes_to_permitalerttypes A, egov_permitalerttypes T "
	sSql = sSql & " WHERE u.isdeleted = 0 AND T.isforfinalpassed = 1 AND T.permitalerttypeid = A.permitalerttypeid "
	sSql = sSql & " AND U.userid = A.notifyuserid AND A.permittypeid = " & iPermitTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		If oRs("email") <> "" Then 
			sToName = oRs("firstname") & " " & oRs("lastname")
			'SendEmailPermits sToName, oRs("email"), sFromName, "webmaster@eclink.com", sSubject, sHTMLBody 
			'sendEmail "", oRs("email") & "(" & sToName & ")", "", sSubject, sHTMLBody, "", ""
			sendEmail "", sToName & " <" & oRs("email") & ">", "", sSubject, sHTMLBody, "", "" 
		End If 
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' void SendReviewStatusChangeAlert iPermitId, iPermitReviewId, sCurrentStatus, sNewStatus 
'-------------------------------------------------------------------------------------------------
Sub SendReviewStatusChangeAlert( ByVal iPermitId, ByVal iPermitReviewId, ByVal sCurrentStatus, ByVal sNewStatus )
	Dim sSql, oRs, iPermitTypeId, sToName, sSubject, sHTMLBody, sFromName, sPermitNo, sDesc
	Dim sJobSite, sReview, sReviewDesc, sReviewer, iReviewerId, sOrgName, sStatus, sLocation
	Dim sLocationType

	' Pull the permit details needed
	sPermitNo = GetPermitNumber( iPermitId )
	sDesc = GetPermitTypeDesc( iPermitId, True )
	sJobSite = GetPermitJobSite( iPermitId )
	iPermitTypeId = GetPermitTypeId( iPermitId )
	sStatus = GetPermitStatusByPermitId( iPermitId )
	sLocation = Replace(GetPermitPermitLocation( iPermitId ), Chr(10), Chr(10) & "<br />")
	sLocationType = GetPermitLocationType( iPermitId )

	' Pull the review details needed
	sReview = ""
	sReviewDesc = ""
	iReviewerId = 0
	GetReviewDetailsForStatusEmail iPermitReviewId, sReview, sReviewDesc, iReviewerId
	sReviewer = GetAdminName( iReviewerId )

	sSubject = "Permit Review Status Change"
	sOrgName = GetOrgName( session("orgid") )
	sFromName = sOrgName & " E-GOV WEBSITE"

	' Build the email body
	sHTMLBody = "<p>This automated message was sent by the " & sOrgName & " E-Gov web site. Do not reply to this message.</p>" & vbcrlf  & vbcrlf 
	sHTMLBody = sHTMLBody & "<p>The status of this review has changed from " & sCurrentStatus & " to " & sNewStatus & ".</p>" & vbcrlf  & vbcrlf 
	sHTMLBody = sHTMLBody & "<p>Permit #: " & sPermitNo & "<br />" & vbcrlf
	sHTMLBody = sHTMLBody & "Permit Type: " & sDesc & "<br />" & vbcrlf
	sHTMLBody = sHTMLBody & "Permit Status: " & sStatus & "<br />" & vbcrlf
	If sLocationType = "address" Then 
		sHTMLBody = sHTMLBody & vbcrlf & "Job Site: " & sJobSite
	End If 
	If sLocationType = "location" Then 
		sHTMLBody = sHTMLBody & vbcrlf & "Location: " & sLocation 
	End If 
	sHTMLBody = sHTMLBody & "</p>" & vbcrlf & vbcrlf
	sHTMLBody = sHTMLBody & "<p>Review: " & sReview & "<br />" & vbcrlf
	sHTMLBody = sHTMLBody & "Review Desc: " & sReviewDesc & "<br />" & vbcrlf
	sHTMLBody = sHTMLBody & "Reviewer: " & sReviewer & "</p>" & vbcrlf & vbcrlf
	sHTMLBody = sHTMLBody & "<p><a href=""" & session("egovclientwebsiteurl") & "/admin/permits/permitedit.asp?permitid=" & iPermitId & """ title=""click to view"">Click here to view the details for this permit.</a></p>"

	' Pull any reviewers that are marked to get the status change alerts and send out
	sSql = "SELECT U.firstname, U.lastname, ISNULL(U.email,'') AS email "
	sSql = sSql & " FROM users U, egov_permittypes_to_permitalerttypes A, egov_permitalerttypes T "
	sSql = sSql & " WHERE u.isdeleted = 0 AND T.isforstatuschanges = 1 AND T.permitalerttypeid = A.permitalerttypeid "
	sSql = sSql & " AND U.userid = A.notifyuserid AND A.permittypeid = " & iPermitTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		If oRs("email") <> "" Then 
			sToName = oRs("firstname") & " " & oRs("lastname")
			'SendEmailPermits sToName, oRs("email"), sFromName, "webmaster@eclink.com", sSubject, sHTMLBody 
			'sendEmail "", oRs("email") & "(" & sToName & ")", "", sSubject, sHTMLBody, "", ""
			sendEmail "", sToName & " <" & oRs("email") & ">", "", sSubject, sHTMLBody, "", "" 
		End If 
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void SetInspectionValues sInspection, sInspectionStatus, sInspectionDate, sInspector, sPermitInspectionType, sInspectionDescription, sStatus, dInspecteddate, iInspectorUserId )
'--------------------------------------------------------------------------------------------------
Sub SetInspectionValues( ByRef sInspection, ByRef sInspectionStatus, ByRef sInspectionDate, ByRef sInspector, ByVal sPermitInspectionType, ByVal sInspectionDescription, ByVal sStatus, ByVal dInspecteddate, ByVal iInspectorUserId )
	Dim sPhone 
			
	sInspection = sPermitInspectionType & " - " & Trim(sInspectionDescription)

	sInspectionStatus = sStatus

	If dInspecteddate <> "" Then 
		sInspectionDate = FormatDateTime(dInspecteddate,2)
	Else
		sInspectionDate = ""
	End If 

	If CLng(iInspectorUserId) > CLng(0) Then 
		sInspector = GetAdminName( CLng(iInspectorUserId) )
		sInspector = sInspector & "<br />"
		sPhone = Trim(GetAdminPhone( CLng(iInspectorUserId) ))
		If sPhone <> "" Then
			sInspector = sInspector & sPhone
		Else 
			sInspector = sInspector & "No Phone"
		End If 
	Else
		' No assigned inspector so put lines for the name and phone
		sInspector = sInspector & "Unassigned<br />No Phone"
	End If 

End Sub 


'-------------------------------------------------------------------------------------------------
' void SetLastActivityDate iPermitId 
'-------------------------------------------------------------------------------------------------
Sub SetLastActivityDate( ByVal iPermitId )
	Dim sSql

	sSql = "UPDATE egov_permits SET lastactivitydate = dbo.GetLocalDate(" & Session("OrgID") & ",getdate()) WHERE permitid = " & iPermitId
	RunSQL sSql

End Sub 


'-------------------------------------------------------------------------------------------------
' string SetPermitFeeTotal( iPermitId )
'-------------------------------------------------------------------------------------------------
Function SetPermitFeeTotal( ByVal iPermitId )
	Dim sSql, oRs, sResponse

	sSql = "SELECT ISNULL(SUM(feeamount),0.00) AS totalfee FROM egov_permitfees WHERE includefee = 1 AND permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sResponse = FormatNumber(oRs("totalfee"),2,,,0) 
		sSql = "UPDATE egov_permits SET feetotal = " & oRs("totalfee") & " WHERE permitid = " & iPermitId 
		RunSQL sSql
	Else
		sResponse = "Failed"
	End If 

	oRs.Close
	Set oRs = Nothing 

	SetPermitFeeTotal = sResponse

End Function 


'-------------------------------------------------------------------------------------------------
' void SetPermitNumber iPermitid, sPermitNumberPrefix 
'-------------------------------------------------------------------------------------------------
Sub SetPermitNumber( iPermitid, sPermitNumberPrefix )
	Dim sSql, oRs, sPermitNumber, sPermitYear

	sPermitYear = Year(Date())
	sSql = "SELECT ISNULL(max(permitnumber),0) + 1 AS nextpermitnumber FROM egov_permits "
	sSql = sSql & " WHERE permitnumberyear = '" & sPermitYear & "' AND orgid = " & Session("orgid")
	sSql = sSql & " AND permitnumberprefix = '" & sPermitNumberPrefix & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sPermitNumber = oRs("nextpermitnumber")
		sSql = "UPDATE egov_permits set permitnumber = " & sPermitNumber & ", permitnumberyear = '" & sPermitYear & "' "
		sSql = sSql & " WHERE permitid = " & iPermitid 
		RunSQL sSql
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub


'-------------------------------------------------------------------------------------------------
' void SetReviewValues sReviewList, sReviewDate, sReviewer, sReviewStatus, sPermitReviewType, sReviewed, sReviewerUserid, sReviewStatusValue
'-------------------------------------------------------------------------------------------------
Sub SetReviewValues( ByRef sReviewList, ByRef sReviewDate, ByRef sReviewer, ByRef sReviewStatus, ByVal sPermitReviewType, ByVal sReviewed, ByVal sReviewerUserid, ByVal sReviewStatusValue )
	Dim sPhone

	sReviewList = sPermitReviewType
	If Not IsNull(sReviewed) Then 
		sReviewDate = FormatDateTime(sReviewed,2)
	Else
		sReviewDate = Space(10)
	End If 
	If CLng(sReviewerUserid) > CLng(0) Then
		sReviewer = GetPermitReviewerName( CLng(sReviewerUserid) )
		sPhone = Trim(GetAdminPhone( CLng(sReviewerUserid) ))
		If sPhone <> "" Then
			sReviewer = sReviewer & "  " & sPhone
		End If 
	Else
		sReviewer = "Unassigned"
	End If 

	sReviewStatus = sReviewStatusValue

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowAmPmPicks sShowName, sAmPmValue 
'--------------------------------------------------------------------------------------------------
Sub ShowAmPmPicks( ByVal sShowName, ByVal sAmPmValue )

	response.write vbcrlf & "<select name=""" & sSHowName & """>"

	response.write vbcrlf & vbtab & "<option value=""AM"""
	If sAmPmValue = "AM" Then
		response.write " selected=""selected"" "
	End If 
	response.write ">AM</option>"

	response.write vbcrlf & vbtab & "<option value=""PM"""
	If sAmPmValue = "PM" Then
		response.write " selected=""selected"" "
	End If 
	response.write ">PM</option>"

	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowBusinessTypes iBusinessTypeId 
'--------------------------------------------------------------------------------------------------
Sub ShowBusinessTypes( ByVal iBusinessTypeId )
	Dim sSql, oRs

	sSql = "SELECT businesstypeid, businesstype FROM egov_permitbusinesstypes "
	sSql = sSql & " WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write "<select id=""businesstypeid"" name=""businesstypeid"">"
	response.write vbcrlf & "<option value=""0"">Select a Business Type</option>"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("businesstypeid") & """"
		If CLng(iBusinessTypeId) = CLng(oRs("businesstypeid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("businesstype") & "</option>"
		oRs.MoveNext
	Loop
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowContractorTypes( iContractorTypeId )
'--------------------------------------------------------------------------------------------------
Sub ShowContractorTypes( ByVal iContractorTypeId )
	Dim sSql, oRs

	sSql = "SELECT contractortypeid, contractortype FROM egov_permitcontractortypes "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select id=""contractortypeid"" name=""contractortypeid"">"
		response.write vbcrlf & "<option value=""0"">Select a Contractor Type</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("contractortypeid") & """"
			If CLng(oRs("contractortypeid")) = CLng(iContractorTypeId) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("contractortype") & "</option>"
			oRs.MoveNext 
		Loop
		
		response.write vbcrlf & "</select>"
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer ShowContractorUsers( iPermitContactTypeId )
'--------------------------------------------------------------------------------------------------
Function ShowContractorUsers( ByVal iPermitContactTypeId )
	Dim sSql, oRs, iRowCount

	iRowCount = CLng(0)
	sSql = "SELECT U.userid, U.userfname, U.userlname, U.userworkphone, C.canaddothers, C.isprimarycontact FROM egov_users U, egov_permitcontacttypes_to_users C "
	sSql = sSql & " WHERE C.userid = U.userid AND C.permitcontacttypeid = " & iPermitContactTypeId
	sSql = sSql & " ORDER BY userlname, userfname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			iRowCount = iRowCount + CLng(1)
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write "><td align=""center""><input type=""checkbox"" id=""removeuser" & iRowCount & """ name=""removeuser" & iRowCount & """ value=""" & oRs("userid") & """ />"
			response.write "<input type=""hidden"" id=""userid" & iRowCount & """ name=""userid" & iRowCount & """ value=""" & oRs("userid") & """ />"
			response.write "</td>"
			response.write "<td>" & oRs("userfname") & " " & oRs("userlname") & "</td>"
			response.write "<td align=""center""><input type=""checkbox"""
			If oRs("canaddothers") Then
				response.write " checked=""checked"" "
			End If 
			response.write " name=""canaddothers" & iRowCount & """ id=""canaddothers" & iRowCount & """ value=""" & oRs("userid") & """ /></td>"
			response.write "<td align=""center""><input type=""radio"""
			If oRs("isprimarycontact") Then
				response.write " checked=""checked"" "
			End If 
			response.write " name=""isprimarycontact"" value=""" & oRs("userid") & """ /></td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop 
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowContractorUsers = iRowCount

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowFeeReportingTypes iFeeReportingTypeId 
'--------------------------------------------------------------------------------------------------
Sub ShowFeeReportingTypes( ByVal iFeeReportingTypeId )
	Dim sSql, oRs

	sSql = "SELECT feereportingtypeid, feereportingtype FROM egov_permitfeereportingtypes "
	sSql = sSql & " WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY feereportingtype"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write "<select id=""feereportingtypeid"" name=""feereportingtypeid"">"
	response.write vbcrlf & "<option value=""0"">Select a Reporting Type</option>"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("feereportingtypeid") & """"
		If CLng(iFeeReportingTypeId) = CLng(oRs("feereportingtypeid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("feereportingtype") & "</option>"
		oRs.MoveNext
	Loop
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowFontFamily sElementName, sMatch 
'--------------------------------------------------------------------------------------------------
Sub ShowFontFamily( ByVal sElementName, ByVal sMatch )

	response.write vbcrlf & "<select id=""" & sElementName & """ name=""" & sElementName & """>"
	response.write vbcrlf & "<option value=""arial"""
	If sMatch = "arial" Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">arial</option>"
	response.write vbcrlf & "<option value=""courier new"""
	If sMatch = "courier new" Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">courier new</option>"
	response.write vbcrlf & "<option value=""georgia"""
	If sMatch = "georgia" Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">georgia</option>"
	response.write vbcrlf & "<option value=""times new roman"""
	If sMatch = "times new roman" Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">times new roman</option>"
	response.write vbcrlf & "<option value=""verdana"""
	If sMatch = "verdana" Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">verdana</option>"
	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowFontSize sElementName, sMatch 
'--------------------------------------------------------------------------------------------------
Sub ShowFontSize( ByVal sElementName, ByVal sMatch )
	Dim sSql, oRs

	sSql = "SELECT fontsize FROM fontsizes ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select id=""" & sElementName & """ name=""" & sElementName & """>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("fontsize") & """"
			If sMatch = oRs("fontsize") Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("fontsize") & "</option>"
			oRs.MoveNext 
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowFontStyles sElementName, sMatch 
'--------------------------------------------------------------------------------------------------
Sub ShowFontStyles( ByVal sElementName, ByVal sMatch )

	response.write vbcrlf & "<select id=""" & sElementName & """ name=""" & sElementName & """>"
	response.write vbcrlf & "<option value=""normal"""
	If sMatch = "normal" Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">normal</option>"
	response.write vbcrlf & "<option value=""italic"""
	If sMatch = "italic" Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">italic</option>"
	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowFontWeight sElementName, sMatch 
'--------------------------------------------------------------------------------------------------
Sub ShowFontWeight( ByVal sElementName, ByVal sMatch )

	response.write vbcrlf & "<select id=""" & sElementName & """ name=""" & sElementName & """>"
	response.write vbcrlf & "<option value=""normal"""
	If sMatch = "normal" Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">normal</option>"
	response.write vbcrlf & "<option value=""bold"""
	If sMatch = "bold" Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">bold</option>"
	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowInvoiceHeader iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowInvoiceHeader( ByVal iPermitId )
	Dim sInvoiceLogo, sInvoiceHeader

	response.write vbcrlf & "<div id=""invoiceheader"">"

	sInvoiceLogo = GetPermitDocumentValue( iPermitId, "invoicelogo" )
	If sInvoiceLogo <> "" Then
		response.write "<img src=""" & sInvoiceLogo & """ border=""0"" />"
	End If 

	sInvoiceHeader = GetPermitDocumentValue( iPermitId, "invoiceheader" )
	If sInvoiceHeader <> "" Then
		response.write "<div id=""invoiceheadertext"">" 
		response.write "<h3>Permit Invoice</h3><p>"
		response.write sInvoiceHeader
		response.write "</p><br /><br />"
		response.write "</div>"
	Else  
		response.write "<h3>" & Session("sOrgName") & " Permit Invoice</h3><br /><br />"
	End If 

	response.write vbcrlf & "</div>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowLatestContactLicense iPermitId, iPermitContactId 
'--------------------------------------------------------------------------------------------------
Sub ShowLatestContactLicense( ByVal iPermitId, ByVal iPermitContactId )
	Dim sSql, oRs, sRequiredTypes, iCount, sPermitCheckDate
	' 12/18/2008 Modified to show only those that are required by the permit

	sRequiredTypes = GetRequiredLicenseTypeIdsAsString( iPermitId )
	' We want to display for the issued date if we have one
	sPermitCheckDate = GetPermitDate( iPermitId, "issueddate" )
	If sPermitCheckDate = "" Then
		sPermitCheckDate = "getdate()"
	Else
		sPermitCheckDate = "'" & sPermitCheckDate & "'"
	End If 

	' Only show the licenses if they are required.
	If sRequiredTypes <> "" Then 

		iCount = clng(0)
		sSql = "SELECT ISNULL(L.licensetype,'') AS licensetype, C.licenseenddate "
		sSql = sSql & " FROM egov_permitcontacts_licenses C, egov_permitlicensetypes L "
		sSql = sSql & " WHERE permitid = " & iPermitID & " AND L.licensetypeid =  C.licensetypeid "
		sSql = sSql & " AND permitcontactid = " & iPermitContactId & " AND L.licensetypeid IN (" & sRequiredTypes & ") "
		sSql = sSql & " AND licenseenddate >= " & sPermitCheckDate
		sSql = sSql & " ORDER BY licenseenddate DESC"
'		response.write sSql & "<br />"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then 
			response.write vbcrlf & "<div class=""contactlicenses"">"
			Do While Not oRs.EOF 
				If iCount > 0 Then
					response.write "<br />"
				End If 
				iCount = iCount + 1
				response.write " &nbsp; License: <strong>" & oRs("licensetype") & "</strong> Expires: "
				If Not IsNull(oRs("licenseenddate")) Then
					response.write "<strong>" & FormatDateTime(oRs("licenseenddate"),2) & "</strong>"
				End If 
				oRs.MoveNext
			Loop 
			response.write vbcrlf & "</div>"
		End If 

		oRs.Close
		Set oRs = Nothing 
	End If 

	If iCount = 0 Then
		response.write "&nbsp;"
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowLicenseTypePicks iLicenseTypeId, iRowCount
'--------------------------------------------------------------------------------------------------
Sub ShowLicenseTypePicks( ByVal iLicenseTypeId, ByVal iRowCount )
	Dim sSql, oRs

	sSql = "SELECT licensetypeid, licensetype FROM egov_permitlicensetypes WHERE orgid = " & session("orgid") & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select id=""licensetypeid" & iRowCount & """ name=""licensetypeid" & iRowCount & """>"
		'response.write vbcrlf & "<option value=""0"">Select a License Type</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("licensetypeid") & """"
			If CLng(oRs("licensetypeid")) = CLng(iLicenseTypeId) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("licensetype") & "</option>"
			oRs.MoveNext 
		Loop
		
		response.write vbcrlf & "</select>"
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPaymentLocations
'--------------------------------------------------------------------------------------------------
Sub ShowPaymentLocations()
	Dim sSql, oRs

	sSql = "SELECT paymentlocationid, paymentlocationname FROM egov_paymentlocations "
	sSql = sSql & "WHERE isadminmethod = 1 ORDER BY paymentlocationid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select name=""paymentlocationid"">"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("paymentlocationid") & """>" & oRs("paymentlocationname") & "</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' ShowPermitCategoryPicks iPermitCategoryId	
'--------------------------------------------------------------------------------------------------
Sub ShowPermitCategoryPicks( ByVal iPermitCategoryId )
	Dim sSql, oRs

	sSQL = "SELECT permitcategoryid, permitcategory FROM egov_permitcategories "
	sSql = sSql & " WHERE orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY permitcategory"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select id=""permitcategoryid"" name=""permitcategoryid"">"
	response.write vbcrlf & "<option value=""0"">View All Categories</option>"

	Do While NOT oRs.EOF 
		response.write vbcrlf & "<option value=""" & oRs("permitcategoryid") & """ "  
		If CLng(iPermitCategoryId) = CLng(oRs("permitcategoryid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("permitcategory")
		response.write "</option>"
		oRs.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPermitTypes iPermitTypeId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitTypes( ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT permittypeid, ISNULL(permittype,'') AS permittype, ISNULL(permittypedesc,'') AS permittypedesc "
	sSql = sSql & "FROM egov_permittypes "
	'sSql = sSql & " WHERE isbuildingpermittype = 1 AND orgid = " & session("orgid")
	sSql = sSql & "WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY permittype, permittypedesc, permittypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select id=""permittypeid"" name=""permittypeid"">"
		If CLng(iPermitTypeId) = CLng(0) Then
			response.write vbcrlf & "<option value=""0"">Please select a permit type...</option>"
		End If 
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value="""  & oRs("permittypeid") & """"
			If CLng(iPermitTypeId) = CLng(oRs("permittypeid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("permittype") 
			If oRs("permittype") <> "" And oRs("permittypedesc") <> "" Then 
				response.write " &ndash; "
			End If 
			response.write oRs("permittypedesc")
			response.write "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	Else
		response.write vbcrlf & "There are No Permit Types to select."
		response.write vbcrlf & "<input type=""hidden"" id=""permittypeid"" name=""permittypeid"" value=""0"" />"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' string ShowPrimaryContactForPermit( iPermitid )
'--------------------------------------------------------------------------------------------------
Function ShowPrimaryContactForPermit( ByVal iPermitId )
	Dim sSql, oRs, sContact

	sContact = ""

	sSql = " SELECT permitcontactid, permitcontacttypeid, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(address,'') AS address, ISNULL(city,'') AS city, "
	sSql = sSql & " ISNULL(state,'') AS state, ISNULL(zip,'') AS zip, ISNULL(phone,'') AS phone " 
	sSql = sSql & " FROM egov_permitcontacts WHERE isprimarycontact = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("firstname") <> "" Then
			response.write oRs("firstname") & " " & oRs("lastname")
		End If 
		If oRs("company") <> "" Then
			If oRs("firstname") <> "" Then 
				response.write "&nbsp; ( " & oRs("company") & " ) "
			Else
				response.write oRs("company")
			End If 
		End If 
		response.write "</td>"
		response.write "<td colspan=""3"" valign=""top"">"
		If Not IsNull(oRs("phone")) And oRs("phone") <> "" Then
			response.write FormatPhoneNumber( oRs("phone") )
		Else
			response.write "&nbsp;"
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowPrimaryContactForPermit = sContact

End Function 


'--------------------------------------------------------------------------------------------------
' string ShowPrimaryContractorForPermit( iPermitid )
'--------------------------------------------------------------------------------------------------
Function ShowPrimaryContractorForPermit( ByVal iPermitId )
	Dim sSql, oRs, sContact, sContactType

	sContact = ""
	If PermitHasAPrimaryContractor( iPermitId ) Then 
		sContactType = "isprimarycontractor"
	Else
		sContactType = "isapplicant"
	End If 

	sSql = " SELECT permitcontactid, permitcontacttypeid, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(address,'') AS address, ISNULL(city,'') AS city, "
	sSql = sSql & " ISNULL(state,'') AS state, ISNULL(zip,'') AS zip, ISNULL(phone,'') AS phone " 
	sSql = sSql & " FROM egov_permitcontacts WHERE " & sContactType & " = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("firstname") <> "" Then 
			sContact = oRs("firstname") & " " & oRs("lastname") & "<br />"
		End If 
		If oRs("company") <> "" Then 
			If sContact = "" Then 
				sContact = oRs("company") & "<br />" 
			Else 
				sContact = sContact & oRs("company") & "<br />" 
			End If 
		End If 
		If Trim(oRs("address")) <> "" Then 
			sContact = sContact & oRs("address") & "<br />" 
		End If 
		If Trim(oRs("city")) <> "" Then 
			sContact = sContact & oRs("city") & ", " & oRs("state") & " " & oRs("zip") & "<br />"
		End If 
		If Not IsNull(oRs("phone")) And Trim(oRs("phone")) <> "" Then 
			sContact = sContact & FormatPhoneNumber( oRs("phone") ) 
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowPrimaryContractorForPermit = sContact

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowUseClasses iUseClassId 
'--------------------------------------------------------------------------------------------------
Sub ShowUseClasses( ByVal iUseClassId )
	Dim sSql, oRs

	sSql = "SELECT useclassid, useclass FROM egov_permituseclasses "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " ORDER BY useclass" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select id=""useclassid"" name=""useclassid"">"
	response.write vbcrlf & "<option value=""0"">Select a Use Class...</option>"

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("useclassid") & """"
		If CLng(iUseClassId) = CLng(oRs("useclassid")) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("useclass") & "</option>"
		oRs.MoveNext 
	Loop 
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowUseTypes iUseTypeId 
'--------------------------------------------------------------------------------------------------
Sub ShowUseTypes( ByVal iUseTypeId )
	Dim sSql, oRs

	sSql = "SELECT usetypeid, usetype FROM egov_permitusetypes "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " ORDER BY usetype" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select id=""usetypeid"" name=""usetypeid"">"
		response.write vbcrlf & "<option value=""0"">Select a Use Type...</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("usetypeid") & """"
			If CLng(iUseTypeId) = CLng(oRs("usetypeid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("usetype") & "</option>"
			oRs.MoveNext 
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean SomeFeesSetToZero( iPermitid )
'--------------------------------------------------------------------------------------------------
Function SomeFeesSetToZero( ByVal iPermitid )
Dim sSql, oRs

	sSql = "SELECT ISNULL(COUNT(permitfeeid),0) AS hits FROM egov_permitfees "
	sSql = sSql & " WHERE feeamount = 0.00 AND permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then 
			' some are still set to 0.00 and have not been removed
			SomeFeesSetToZero = True 
		Else
			SomeFeesSetToZero = False 
		End If 
	Else
		SomeFeesSetToZero = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' boolean StatusAllowsNewInvoices( iPermitStatusId )
'-------------------------------------------------------------------------------------------------
Function StatusAllowsNewInvoices( ByVal iPermitStatusId )
	Dim sSql, oRs

	sSql = "SELECT allownewinvoices FROM egov_permitstatuses "
	sSql = sSql & " WHERE permitstatusid = " & iPermitStatusId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("allownewinvoices") Then 
			StatusAllowsNewInvoices = True 
		Else
			StatusAllowsNewInvoices = False 
		End If 
	Else
		StatusAllowsNewInvoices = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' boolean StatusAllowsChangesToPropagate( iPermitStatusId )
'-------------------------------------------------------------------------------------------------
Function StatusAllowsChangesToPropagate( ByVal iPermitStatusId )
	Dim sSql, oRs

	sSql = "SELECT changespropagate FROM egov_permitstatuses "
	sSql = sSql & " WHERE permitstatusid = " & iPermitStatusId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("changespropagate") Then 
			StatusAllowsChangesToPropagate = True 
		Else
			StatusAllowsChangesToPropagate = False
		End If 
	Else
		StatusAllowsChangesToPropagate = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' boolean StatusAllowsPermitCancel( iPermitStatusId )
'-------------------------------------------------------------------------------------------------
Function StatusAllowsPermitCancel( ByVal iPermitStatusId )
	Dim sSql, oRs

	sSql = "SELECT cancancelpermit FROM egov_permitstatuses "
	sSql = sSql & " WHERE permitstatusid = " & iPermitStatusId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("cancancelpermit") Then 
			StatusAllowsPermitCancel = True 
		Else
			StatusAllowsPermitCancel = False
		End If 
	Else
		StatusAllowsPermitCancel = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' boolean StatusAllowsSaveChanges( iPermitStatusId )
'-------------------------------------------------------------------------------------------------
Function StatusAllowsSaveChanges( ByVal iPermitStatusId )
	Dim sSql, oRs

	sSql = "SELECT cansavechanges FROM egov_permitstatuses "
	sSql = sSql & " WHERE permitstatusid = " & iPermitStatusId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("cansavechanges") Then 
			StatusAllowsSaveChanges = True 
		Else
			StatusAllowsSaveChanges = False
		End If 
	Else
		StatusAllowsSaveChanges = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' void  UpdateFeeTotal iPermitId, iFeeTotal 
'-------------------------------------------------------------------------------------------------
Sub UpdateFeeTotal( ByVal iPermitId, ByVal iFeeTotal )
	Dim sSql

	sSql = "UPDATE egov_permits SET feetotal = " & iFeeTotal & " WHERE permitid = " & iPermitId

	RunSQL sSql

End Sub 
'-------------------------------------------------------------------------------------------------
' boolean = PermitHasNoPendingInspections( iPermitid, iPassedStatusId )
'-------------------------------------------------------------------------------------------------
Function PermitHasNoPendingInspections( ByVal iPermitid, ByVal iPassedStatusId )
	Dim sSql, oRs

	'sSql = "SELECT COUNT(permitinspectionid) AS hits FROM egov_permitinspections WHERE permitid = " & iPermitId
	'sSql = sSql & " AND inspectionstatusid != " & iPassedStatusId

	sSql = "SELECT COUNT(permitinspectionid) AS hits  "
	sSql = sSql & "FROM egov_permitinspections epi "
	sSql = sSql & "INNER JOIN egov_inspectionstatuses eis ON eis.inspectionstatusid = epi.inspectionstatusid "
	sSql = sSql & "WHERE permitid = " & iPermitId & " and eis.isdone <> 1 "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then 
			PermitHasNoPendingInspections = False  
		Else
			PermitHasNoPendingInspections = True  
		End If 
	Else
		PermitHasNoPendingInspections = True  
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------





%>
