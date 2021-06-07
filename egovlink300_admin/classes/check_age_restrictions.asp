<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: check_age_restrictions.asp
' AUTHOR: Steve Loar
' CREATED: 04/04/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This checks that the selected family member meets a classes age restrictions, is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   04/04/07	Steve Loar - INITIAL VERSION
' 1.1	08/16/2007	Steve Loar - Change to work from age compare date of class not start date
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim Age, sResult, dBirthDate, dAgeCompareDate, iMinAge, iMaxAge, iMinAgePrecisionId, iMaxAgePrecisionId

' flag as passed
sResult = "PASSED"
dAgeCompareDate = Now()
iMinAge = CDbl(0.0)
iMaxAge = CDbl(0.0)
iMinAgePrecisionId = 1
iMaxAgePrecisionId = 1

' Get Class Values agecomparedate, minage, maxage, minageprecisionid, maxageprecisionid
getClassValues CLng(request("classid")), dAgeCompareDate, iMinAge, iMaxAge, iMinAgePrecisionId, iMaxAgePrecisionId

' Get birthdate 
dBirthDate = GetBirthDate( CLng(request("familymemberid")) )

' Get age in decimals
Age = GetAgeOnDate( dBirthDate, dAgeCompareDate )  ' In class_global_functions.asp

If CDbl(iMinAge) > CDbl(0) Then 
	' if precisionid iswholeyear  --  In class_global_functions.asp
	If HasWholeYearPrecision( iMinAgePrecisionId ) Then 
		' age = int(age)
		Age = Int(Age)
	End If 
	' if age < minage
	If CDbl(Age) < CDbl(iMinAge) Then 
		' flag as failed
		sResult = "FAILED"
	End If 
End If 

If CDbl(iMaxAge) > CDbl(0) Then 
	' if precisionid iswholeyear  --  In class_global_functions.asp
	If HasWholeYearPrecision( iMaxAgePrecisionId ) Then 
		' age = int(age)
		Age = Int(Age)
	End If 
	' if age > maxage
	If CDbl(Age) > CDbl(iMaxAge) Then 
		' flag as failed
		sResult = "FAILED"
	End If 
End If 

' For testing
'sResult = sResult & " " & dBirthDate
'sResult = sResult & " " & dAgeCompareDate
'sResult = sResult & " " & Age

' return flag
response.write sResult

%>

<!--#Include file="class_global_functions.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub getClassValues( iClassId, ByRef dStartDate, ByRef iMinAge, ByRef iMaxAge, ByRef iMinAgePrecisionId, ByRef iMaxAgePrecisionId )
'--------------------------------------------------------------------------------------------------
Sub getClassValues( iClassId, ByRef dAgeCompareDate, ByRef iMinAge, ByRef iMaxAge, ByRef iMinAgePrecisionId, ByRef iMaxAgePrecisionId )
	Dim sSql, oClass

	sSql = "Select agecomparedate, isnull(minage,0.0) as minage, isnull(maxage,0.0) as maxage, isnull(minageprecisionid,0) as minageprecisionid, "
	sSql = sSql & " isnull(maxageprecisionid,0) as maxageprecisionid from egov_class where classid = " & iClassId

	Set oClass = Server.CreateObject("ADODB.Recordset")
	oClass.Open sSQL, Application("DSN"), 3, 1

	If Not oClass.EOF Then
		dStartDate = oClass("agecomparedate")
		iMinAge = CDbl(oClass("minage"))
		iMaxAge = CDbl(oClass("maxage"))
		iMinAgePrecisionId = oClass("minageprecisionid")
		iMaxAgePrecisionId = oClass("maxageprecisionid")
	End If 

	oClass.Close
	Set oClass = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetBirthDate( iFamilyMemberId )
'--------------------------------------------------------------------------------------------------
Function GetBirthDate( iFamilyMemberId )
	Dim sSql, oBirthdate

	sSql = "Select birthdate from egov_familymembers where familymemberid = " & iFamilyMemberId

	Set oBirthdate = Server.CreateObject("ADODB.Recordset")
	oBirthdate.Open sSQL, Application("DSN"), 3, 1

	If Not oBirthdate.EOF Then
		If IsNull(oBirthdate("birthdate")) Then
			GetBirthDate = DateAdd("yyyy", -22, Date())
		Else
			GetBirthDate = oBirthdate("birthdate")
		End If 
	Else
		GetBirthDate = DateAdd("yyyy", -22, Date())
	End If 

	oBirthdate.Close
	Set oBirthdate = Nothing 

End Function 


%>

